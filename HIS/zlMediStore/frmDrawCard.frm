VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{876E3FF4-6E21-11D5-AF7D-0080C8EC27A9}#1.5#0"; "ZL9BillEdit.ocx"
Begin VB.Form frmDrawCard 
   Caption         =   "ҩƷ���õ�"
   ClientHeight    =   8025
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   14535
   Icon            =   "frmDrawCard.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8025
   ScaleWidth      =   14535
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton cmdExpend 
      Caption         =   "�Զ��ֽ�(&A)"
      Height          =   350
      Left            =   4800
      TabIndex        =   37
      Top             =   5835
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.CommandButton cmdAllCls 
      Caption         =   "ȫ��(&L)"
      Height          =   350
      Left            =   7560
      TabIndex        =   35
      Top             =   5820
      Visible         =   0   'False
      Width           =   1100
   End
   Begin VB.CommandButton cmdAllSel 
      Caption         =   "ȫ��(&A)"
      Height          =   350
      Left            =   6240
      TabIndex        =   34
      Top             =   5820
      Visible         =   0   'False
      Width           =   1100
   End
   Begin VB.TextBox txtCode 
      Height          =   300
      Left            =   3720
      TabIndex        =   15
      Top             =   5490
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "����(&F)"
      Height          =   350
      Left            =   2040
      TabIndex        =   14
      Top             =   5400
      Width           =   1100
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      Height          =   350
      Left            =   240
      TabIndex        =   13
      Top             =   5400
      Width           =   1100
   End
   Begin VB.CommandButton CmdSave 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   6240
      TabIndex        =   11
      Top             =   5400
      Width           =   1100
   End
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   7560
      TabIndex        =   12
      Top             =   5400
      Width           =   1100
   End
   Begin VB.PictureBox Pic���� 
      BackColor       =   &H80000004&
      Height          =   5325
      Left            =   0
      ScaleHeight     =   5265
      ScaleWidth      =   14415
      TabIndex        =   16
      Top             =   0
      Width           =   14475
      Begin VB.TextBox txt����ԭ�� 
         Height          =   300
         Left            =   900
         TabIndex        =   8
         Top             =   4170
         Visible         =   0   'False
         Width           =   10410
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
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   135
         Visible         =   0   'False
         Width           =   1515
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
         Left            =   840
         Style           =   1  'Graphical
         TabIndex        =   38
         TabStop         =   0   'False
         ToolTipText     =   "������ʵ�:F3"
         Top             =   150
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.ComboBox cboDrawPerson 
         Height          =   300
         Left            =   9645
         TabIndex        =   4
         Top             =   615
         Width           =   1515
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshProvider 
         Height          =   1815
         Left            =   7320
         TabIndex        =   33
         Top             =   960
         Visible         =   0   'False
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   3201
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
      Begin VB.TextBox txtDraw 
         Height          =   300
         Left            =   5355
         TabIndex        =   3
         Top             =   615
         Width           =   2415
      End
      Begin VB.CommandButton cmdDraw 
         Caption         =   "��"
         Height          =   300
         Left            =   7755
         TabIndex        =   5
         Top             =   615
         Width           =   300
      End
      Begin ZL9BillEdit.BillEdit mshBill 
         Height          =   2805
         Left            =   195
         TabIndex        =   6
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
         TabIndex        =   10
         Top             =   4560
         Width           =   10410
      End
      Begin VB.ComboBox cboStock 
         Height          =   300
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   600
         Width           =   1515
      End
      Begin VB.Label lbl����ԭ�� 
         AutoSize        =   -1  'True
         Caption         =   "����ԭ��"
         Height          =   180
         Left            =   120
         TabIndex        =   7
         Top             =   4260
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.Label lbl�޸����� 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�޸�����"
         Height          =   180
         Left            =   6900
         TabIndex        =   43
         Top             =   4980
         Width           =   720
      End
      Begin VB.Label lbl�޸��� 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�޸���"
         Height          =   180
         Left            =   5040
         TabIndex        =   42
         Top             =   4980
         Width           =   540
      End
      Begin VB.Label Txt�޸��� 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   5640
         TabIndex        =   41
         Top             =   4920
         Width           =   915
      End
      Begin VB.Label Txt�޸����� 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   7680
         TabIndex        =   40
         Top             =   4920
         Width           =   1875
      End
      Begin VB.Label lblDrawPerson 
         AutoSize        =   -1  'True
         Caption         =   "������(&P)"
         Height          =   180
         Left            =   8730
         TabIndex        =   36
         Top             =   675
         Width           =   810
      End
      Begin VB.Label lblDifference 
         AutoSize        =   -1  'True
         Caption         =   "��ۺϼ�:"
         Height          =   180
         Left            =   4920
         TabIndex        =   31
         Top             =   3840
         Width           =   810
      End
      Begin VB.Label lblSalePrice 
         AutoSize        =   -1  'True
         Caption         =   "�ۼ۽��ϼ�:"
         Height          =   180
         Left            =   2040
         TabIndex        =   30
         Top             =   3840
         Width           =   1170
      End
      Begin VB.Label lblPurchasePrice 
         AutoSize        =   -1  'True
         Caption         =   "�ɱ����ϼ�:"
         Height          =   180
         Left            =   240
         TabIndex        =   29
         Top             =   3840
         Width           =   1170
      End
      Begin VB.Label Txt����� 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   10350
         TabIndex        =   27
         Top             =   4920
         Width           =   1005
      End
      Begin VB.Label Txt������� 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   12450
         TabIndex        =   26
         Top             =   4920
         Width           =   1875
      End
      Begin VB.Label Txt�������� 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   2940
         TabIndex        =   25
         Top             =   4920
         Width           =   1875
      End
      Begin VB.Label Txt������ 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   900
         TabIndex        =   24
         Top             =   4920
         Width           =   915
      End
      Begin VB.Label txtNo 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   9960
         TabIndex        =   23
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
         TabIndex        =   22
         Top             =   195
         Width           =   480
      End
      Begin VB.Label lblժҪ 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ժ Ҫ(&M)"
         Height          =   180
         Left            =   165
         TabIndex        =   9
         Top             =   4635
         Width           =   720
      End
      Begin VB.Label LblTitle 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "ҩƷ���õ�"
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
         TabIndex        =   21
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
         TabIndex        =   20
         Top             =   4980
         Width           =   540
      End
      Begin VB.Label Lbl�������� 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��������"
         Height          =   180
         Left            =   2160
         TabIndex        =   19
         Top             =   4980
         Width           =   720
      End
      Begin VB.Label Lbl����� 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�����"
         Height          =   180
         Left            =   9765
         TabIndex        =   18
         Top             =   4980
         Width           =   540
      End
      Begin VB.Label Lbl������� 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�������"
         Height          =   180
         Left            =   11640
         TabIndex        =   17
         Top             =   4980
         Width           =   720
      End
      Begin VB.Label LblEnterStock 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ҩ����(&D)"
         Height          =   180
         Left            =   4275
         TabIndex        =   2
         Top             =   675
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
            Picture         =   "frmDrawCard.frx":014A
            Key             =   "PreView"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrawCard.frx":0364
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrawCard.frx":057E
            Key             =   "Add"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrawCard.frx":0798
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrawCard.frx":09B2
            Key             =   "Check"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrawCard.frx":0BCC
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrawCard.frx":0DE6
            Key             =   "Exit"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrawCard.frx":1000
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
            Picture         =   "frmDrawCard.frx":121A
            Key             =   "PreView"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrawCard.frx":1434
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrawCard.frx":164E
            Key             =   "Add"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrawCard.frx":1868
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrawCard.frx":1A82
            Key             =   "Check"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrawCard.frx":1C9C
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrawCard.frx":1EB6
            Key             =   "Exit"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrawCard.frx":20D0
            Key             =   "Find"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar staThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   32
      Top             =   7665
      Width           =   14535
      _ExtentX        =   25638
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmDrawCard.frx":22EA
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ��������Ϣ��ҵ��˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   19288
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   617
            MinWidth        =   617
            Picture         =   "frmDrawCard.frx":2B7E
            Key             =   "PY"
            Object.ToolTipText     =   "ƴ��(F7)"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   2
            Object.Width           =   617
            MinWidth        =   617
            Picture         =   "frmDrawCard.frx":3080
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
      TabIndex        =   28
      Top             =   5520
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
Attribute VB_Name = "frmDrawCard"
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
Private mblnStock As Boolean                '��ǰ����Ա�Ƿ��ǿⷿ��Ա
Private mfrmMain As Form
Private mintcboIndex As Integer
Private mblnEnterCell As Boolean            '�Ƿ�������ENTERCELL()�¼�
Private mblnEdit As Boolean                 '�Ƿ�����޸�
Private mblnChange As Boolean               '�Ƿ���й��༭
Private mintParallelRecord As Integer       '���������󵥾ݲ���ִ�еĴ��� 1���������������2���Ѿ�ɾ���ļ�¼��3���Ѿ���˵ļ�¼
Private mblnAutoExp As Boolean                           '���ݷ������Զ��ֽ�
Private mbln��ʾ As Boolean                 '��ҩƷѡ������ѡ���ҩƷ��������������ݵıȽϿ��Ƿ��ظ��������ظ�������ֻ��ʾһ�Σ�true �Ѿ���ʾ�ˣ�false��û����ʾ

Private mint����� As Integer             '��ʾҩƷ����ʱ�Ƿ���п���飺0-�����;1-��飬�������ѣ�2-��飬�����ֹ
Private mcolUsedCount As Collection         '��ʹ�õ���������
Private mstrPrivs As String                 'Ȩ��
Private mblnViewCost As Boolean             '�鿴�ɱ��� true-���Բ鿴�ɱ��� false-�����Բ鿴�ɱ���
Private mint�����γ��� As Integer           '0-�������γ���,1-�����γ���
Private mint����ʽ As Integer             '����ʱ��0������������1�������������뵥�ݣ�2������Ѳ����ĳ������뵥��

Private Const mlng��ɫ As Long = &HC000C0

Private mint���÷�ʽ As Integer              '0-��ⷿ��ҩ;1-�����������ҩ
Private str���治����ʾ As String
Private mint���淽ʽ As Integer             '0-�������� 1-��������
Private recSort As ADODB.Recordset          '��ҩƷID�����ר�ü�¼��

Private mblnUpdate As Boolean               '��ʾ�Ƿ��Ѹ������¼۸���µ�������

Private mlng����ⷿ As Long
Private mintUnit As Integer                 '��λϵ����1-�ۼ�;2-����;3-סԺ;4-ҩ��

Private mintDrugNameShow As Integer         'ҩƷ��ʾ��0����ʾ��������ƣ�1������ʾ���룻2������ʾ����
Private Const MStrCaption As String = "ҩƷ���ù���"
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
Private Const mconIntCol��д���� As Integer = 20
Private Const mconIntColʵ������ As Integer = 21
Private Const mconIntCol�ɹ��� As Integer = 22
Private Const mconIntCol�ɹ���� As Integer = 23
Private Const mconIntCol�ۼ� As Integer = 24
Private Const mconIntCol�ۼ۽�� As Integer = 25
Private Const mconintCol��� As Integer = 26
Private Const mconintCol��ʵ���� As Integer = 27
Private Const mconIntColҩƷ��������� As Integer = 28
Private Const mconIntColҩƷ���� As Integer = 29
Private Const mconIntColҩƷ���� As Integer = 30
Private Const mconintColԭʼ���� As Integer = 31
Private Const mconIntColS  As Integer = 32            '������
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
Private Function Check����() As Boolean
    Dim rs As New ADODB.Recordset
    Dim rsTmp As New ADODB.Recordset
    Dim n As Integer
    Dim strSQL As String
    
    '������������������Ƿ��㹻
    On Error GoTo errHandle
    With rs
        .Fields.Append "ҩƷID", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "ʵ������", adDouble, 18, adFldIsNullable
        .Fields.Append "ҩƷ����", adLongVarChar, 40, adFldIsNullable
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
        
        For n = 1 To mshBill.rows - 1
            If Val(mshBill.TextMatrix(n, 0)) > 0 Then
                If .RecordCount = 0 Then
                    .AddNew
                    !ҩƷID = Val(mshBill.TextMatrix(n, 0))
                    !ʵ������ = Val(mshBill.TextMatrix(n, mconIntCol��д����)) * Val(mshBill.TextMatrix(n, mconIntCol����ϵ��))
                    !ҩƷ���� = mshBill.TextMatrix(n, 2)
                    .Update
                Else
                    .MoveFirst
                    .Find "ҩƷID=" & Val(mshBill.TextMatrix(n, 0)) & " "
                    If .EOF Then
                        .AddNew
                        !ҩƷID = Val(mshBill.TextMatrix(n, 0))
                        !ʵ������ = Val(mshBill.TextMatrix(n, mconIntCol��д����)) * Val(mshBill.TextMatrix(n, mconIntCol����ϵ��))
                        !ҩƷ���� = mshBill.TextMatrix(n, 2)
                        .Update
                    Else
                        !ʵ������ = !ʵ������ + Val(mshBill.TextMatrix(n, mconIntCol��д����)) * Val(mshBill.TextMatrix(n, mconIntCol����ϵ��))
                        .Update
                    End If
                End If
            End If
        Next
    End With
    
    rs.MoveFirst
    For n = 1 To rs.RecordCount
        strSQL = "select ʵ������ from ҩƷ���� where ����id=[2] and �ⷿid=[1] " & _
        " and ҩƷid=[3] and �ڼ� = [4]"
        Set rsTmp = zlDataBase.OpenSQLRecord(strSQL, MStrCaption, cboStock.ItemData(cboStock.ListIndex), txtDraw.Tag, CLng(rs!ҩƷID), Format(Sys.Currentdate(), IIf(mint���淽ʽ = 0, "yyyy", "yyyymm")))
        
        If rsTmp.RecordCount = 0 Then
            Check���� = False
            str���治����ʾ = "���ڼ䡾" & rs!ҩƷ���� & "��û�����������������ã����޸ĵ��ݣ�"
            Exit Function
        ElseIf rsTmp!ʵ������ < rs!ʵ������ Then
            Check���� = False
            str���治����ʾ = rs!ҩƷ���� & "��������[" & rs!ʵ������ & "]��������������[" & rsTmp!ʵ������ & "]�������ã����޸ĵ��ݣ�"
            Exit Function
            End If
        rs.MoveNext
    Next
    
    Check���� = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub GetDrawPerson(ByVal strDeptId As String)
    Dim rs As New ADODB.Recordset
    Dim n As Integer
    
    On Error GoTo errHandle
    cboDrawPerson.Clear
    
    If strDeptId = "" Then Exit Sub

    gstrSQL = "Select ���,����,���� From ��Ա�� Where (վ�� = [2] Or վ�� is Null) And Id In (Select ��Աid From ������Ա Where ����id=[1]) " & _
              " And (����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or ����ʱ�� Is Null) "
    Set rs = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption, strDeptId, gstrNodeNo)
    
    If rs.RecordCount = 0 Then Exit Sub
    
    For n = 1 To rs.RecordCount
        cboDrawPerson.AddItem (rs!����)
        rs.MoveNext
    Next
    rs.Close
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

'�������������
Private Function GetDepend() As Boolean
    Dim rsDepend As New Recordset
    Dim strSQL As String
    
    On Error GoTo errHandle
    GetDepend = False
    strSQL = "SELECT B.Id " _
            & "FROM ҩƷ�������� A, ҩƷ������ B " _
            & "Where A.���id = B.ID AND A.���� = 7 "
    Set rsDepend = zlDataBase.OpenSQLRecord(strSQL, MStrCaption)
    If rsDepend.EOF Then
        MsgBox "û������ҩƷ���õĳ����������ҩƷ������࣡", vbInformation + vbOKOnly, gstrSysName
        rsDepend.Close
        Exit Function
    End If
    rsDepend.Close

    If mblnStock Then
        strSQL = "SELECT DISTINCT a.id, a.���� " _
               & "FROM ��������˵�� c, �������ʷ��� b, ���ű� a " _
               & "Where (a.վ�� = [2] Or a.վ�� is Null) And c.�������� = b.���� " _
               & "  AND b.���� ='O' AND a.id = c.����id " _
               & "  AND TO_CHAR (a.����ʱ��, 'yyyy-MM-dd') = '3000-01-01'"
    Else
        strSQL = " Select C.ID " & _
                 " From ��������˵�� A,�������ʷ��� B,���ű� C " & _
                 " Where (c.վ�� = [2] Or c.վ�� is Null) And A.��������=B.���� And A.����ID=C.ID " & _
                 "   AND TO_CHAR(C.����ʱ��, 'yyyy-MM-dd')='3000-01-01' And B.����='O'" & _
                 "   And C.ID IN (Select ����ID From ������Ա Where ��ԱID=[1])"
    End If
    Set rsDepend = zlDataBase.OpenSQLRecord(strSQL, "ҩƷ���õ�", UserInfo.�û�ID, gstrNodeNo)
    
    If rsDepend.EOF Then
        If mblnStock Then
            MsgBox "��ҩ����������Ϣ��ȫ,��鿴���Ź���", vbInformation, gstrSysName
        Else
            MsgBox "�㲻�����κ����ò���,������дҩƷ���õ�,��鿴���Ź���", vbInformation, gstrSysName
        End If
        rsDepend.Close
        Exit Function
    End If
    
    GetDepend = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Sub ShowCard(FrmMain As Form, ByVal str���ݺ� As String, ByVal int�༭״̬ As Integer, ByVal blnҩ����Ա As Boolean, _
    Optional int��¼״̬ As Integer = 1, Optional int���÷�ʽ As Integer = 0, Optional BlnSuccess As Boolean = False, Optional int����ʽ As Integer = 0)
    mblnSave = False
    mblnSuccess = False
    mstr���ݺ� = str���ݺ�
    mint�༭״̬ = int�༭״̬
    mint��¼״̬ = int��¼״̬
    mblnSuccess = BlnSuccess
    mblnChange = False
    mblnStock = blnҩ����Ա
    mintParallelRecord = 1
    mint���÷�ʽ = int���÷�ʽ
    mint����ʽ = int����ʽ
    mstrPrivs = GetPrivFunc(glngSys, 1305)
    
    Set mfrmMain = FrmMain
    If Not GetDepend Then Exit Sub
    
    If mint�༭״̬ = 1 Then
        mblnEdit = True
    ElseIf mint�༭״̬ = 2 Then
        mblnEdit = True
    ElseIf mint�༭״̬ = 3 Then
        mblnEdit = False
        CmdSave.Caption = "���(&V)"
        '�������÷�ʽ�����Զ��ֽ�
'        If mint���÷�ʽ = 1 Then
'            cmdExpend.Visible = False
'        Else
            cmdExpend.Visible = mint�����γ��� <> 1 '�������β���Ҫ�Զ��ֽ�
'        End If
    ElseIf mint�༭״̬ = 4 Then
        mblnEdit = False
        CmdSave.Caption = "��ӡ(&P)"
        If Not zlStr.IsHavePrivs(mstrPrivs, "���ݴ�ӡ") Then
            CmdSave.Visible = False
        Else
            CmdSave.Visible = True
        End If
        
        '�鿴�������ݣ�����ԭ��ɼ����ɱ༭
        If mint��¼״̬ Mod 3 = 2 Then
            lbl����ԭ��.Visible = True
            txt����ԭ��.Visible = True
            
            txt����ԭ��.Enabled = False
        End If
        
    ElseIf mint�༭״̬ = 6 Then
        mblnEdit = False
        
        cmdAllSel.Visible = True
        cmdAllCls.Visible = True
        
        '����ʱ����ԭ��ſɼ�
        lbl����ԭ��.Visible = True
        txt����ԭ��.Visible = True
        
        If mint����ʽ = 1 Then
            CmdSave.Caption = "�������(&O)"
            CmdSave.Width = CmdSave.Width + 200
        ElseIf mint����ʽ = 2 Then
            CmdSave.Caption = "��˳���(&V)"
            CmdSave.Width = CmdSave.Width + 200
            
            cmdAllSel.Visible = False
            cmdAllCls.Visible = False
            
            txt����ԭ��.Enabled = False '��˳������ɱ༭
        Else
            CmdSave.Caption = "����(&O)"
            CmdSave.Width = CmdCancel.Width
        End If
    End If
    LblTitle.Caption = GetUnitName & "ҩƷ���õ�" & IIf(mint���÷�ʽ = 0, "(�ⷿ����)", "(��������)")
    Me.Show vbModal, FrmMain
    BlnSuccess = mblnSuccess
    str���ݺ� = mstr���ݺ�
End Sub

Private Sub cboDrawPerson_Click()
    'mshBill.SetFocus
    mshBill.Col = 1
    mshBill.Row = 1
End Sub

Private Sub cboDrawPerson_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim strText As String
    Dim rs As New ADODB.Recordset
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    On Error GoTo errHandle
    With cboDrawPerson
        If Trim(.Text) = "" Then Exit Sub
        strText = UCase(.Text)
        
        mshProvider.Tag = 1
        
        gstrSQL = "Select ���,����,���� From ��Ա�� " & _
                  "Where (վ�� = [3] Or վ�� is Null) And Id In(Select ��Աid From ������Ա Where ����id=[1]) " & _
                  "  And (���� like [2] Or ��� like [2] or ���� like [2]) " & _
                  "  And (����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or ����ʱ�� Is Null) "
        Set rs = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption, _
            Val(Me.txtDraw.Tag), _
            IIf(gstrMatchMethod = "0", "%", "") & strText & "%", _
            gstrNodeNo)
        
        If rs.EOF Then
            KeyCode = 0
            .SelStart = 0
            .SelLength = Len(.Text)
            Exit Sub
        End If
        
        If rs.RecordCount > 1 Then
            Set mshProvider.Recordset = rs
            Dim intCol As Integer
            Dim intRow As Integer
            
            With mshProvider
                If .Visible = False Then .Visible = True
                .Redraw = False
                .SetFocus
                
                For intRow = 0 To .rows - 1
                    .Row = intRow
                    For intCol = 0 To .Cols - 1
                        .Col = intCol
                        If .Row = 0 Then
                            .CellFontBold = True
                        Else
                            .CellFontBold = False
                        End If
                    Next
                Next
                .Font.Bold = False
                .FontFixed.Bold = True
                .ColWidth(0) = 800
                .ColWidth(1) = 800
                .ColWidth(2) = 800
                .Row = 1
                .TopRow = 1
                .Col = 0
                .ColSel = .Cols - 1
                
                .Width = lblDrawPerson.Width + cboDrawPerson.Width + 80
                .Top = txtDraw.Top + txtDraw.Height
                .Left = lblDrawPerson.Left
                .Redraw = True
                Exit Sub
            End With
        Else
            .Text = rs!����
            mshBill.SetFocus
            mshBill.Col = 1
            mshBill.Row = 1
        End If
        rs.Close
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cboDrawPerson_KeyPress(KeyAscii As Integer)
    If InStr(" ';", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub cboStock_Change()
    mblnChange = True
End Sub

Private Sub cboStock_Click()
    Dim rsDetail As New Recordset
    Dim bln��ҩ�ⷿ As Boolean
    Dim str�ⷿ���� As String
    
    On Error GoTo errHandle

    mint����� = MediWork_GetCheckStockRule(cboStock.ItemData(cboStock.ListIndex))
    If mint�༭״̬ = 1 Or mint�༭״̬ = 2 Then
        str�ⷿ���� = ""
        gstrSQL = "Select a.�������� From ��������˵�� A Where a.����id =[1]"
        Set rsDetail = zlDataBase.OpenSQLRecord(gstrSQL, "�жϿⷿ����", cboStock.ItemData(cboStock.ListIndex))
        Do While Not rsDetail.EOF
            str�ⷿ���� = str�ⷿ���� & "," & rsDetail!��������
            rsDetail.MoveNext
        Loop
        If str�ⷿ���� Like "*��ҩ*" Or str�ⷿ���� Like "*�Ƽ���*" Then bln��ҩ�ⷿ = True
        mshBill.ColWidth(mconIntColԭ����) = IIf(bln��ҩ�ⷿ, 800, 0)
    
        If mblnLoad = True Then Call SetSelectorRS(2, MStrCaption, cboStock.ItemData(cboStock.ListIndex), cboStock.ItemData(cboStock.ListIndex), Val(txtDraw.Tag))
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
                    
                    If Not mblnStock Then
                    MsgBox "������������ҩ���ź���ҩ�ˣ�", vbOKOnly, gstrSysName
                        txtDraw.Text = ""
                        txtDraw.Tag = ""
                        cboDrawPerson.Clear
                    End If
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
        '2010-5-7 ������޸�
        mblnChange = True
    End With
    Call ��ʾ�ϼƽ��
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdDraw_Click()
    Dim rsProvider As New Recordset
    Dim strվ������ As String
    
    On Error GoTo errHandle
    strվ������ = GetDeptStationNode(cboStock.ItemData(cboStock.ListIndex))
    If mblnStock Then
        gstrSQL = "SELECT DISTINCT a.id,null �ϼ�id,1 as ĩ��, a.����,a.����,a.���� " _
                & "FROM ��������˵�� c, �������ʷ��� b, ���ű� a " _
                & "Where " & IIf(strվ������ <> "", " (a.վ�� = [3] or a.վ�� is null) AND ", "") & " c.�������� = b.���� " _
                & "  AND b.���� ='O' AND a.id = c.����id " _
                & "  AND (TO_CHAR(a.����ʱ��, 'yyyy-MM-dd') = '3000-01-01' Or a.����ʱ�� Is NULL) " _
                & "Order By a.���� "
    Else
        gstrSQL = " Select C.ID " & _
                  " From ��������˵�� A,�������ʷ��� B,���ű� C " & _
                  " Where " & IIf(strվ������ <> "", " (C.վ�� = [3] or C.վ�� is null) And ", "") & " A.��������=B.���� And A.����ID=C.ID " & _
                  "   AND TO_CHAR(C.����ʱ��, 'yyyy-MM-dd')='3000-01-01' And B.����='O'" & _
                  "   And C.ID IN (Select ����ID From ������Ա Where ��ԱID=[1])"
        'ֻ��ȡ��������������Ŀ���
        gstrSQL = " SELECT DISTINCT C.id,null �ϼ�id,1 as ĩ��, C.����,C.����,C.����" & _
                  " From ��������˵�� A,�������ʷ��� B,���ű� C " & _
                  " Where " & IIf(strվ������ <> "", " (C.վ�� = [3] or C.վ�� is null) And ", "") & " A.��������=B.���� And A.����ID=C.ID " & _
                  "   AND TO_CHAR(C.����ʱ��, 'yyyy-MM-dd')='3000-01-01' And B.����='O'" & _
                  "   And C.ID IN (Select Distinct ���ò���ID From ҩƷ���ÿ��� Where �Է��ⷿid=[2] And ���ò���ID IN (" & gstrSQL & ")) " & _
                  " Order By C.���� "
    End If
    Set rsProvider = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "[��ȡ������ҩ����]", _
        UserInfo.�û�ID, _
        cboStock.ItemData(cboStock.ListIndex), _
        strվ������)
    
    If rsProvider.EOF Then
        rsProvider.Close
        Exit Sub
    End If
    With FrmSelect
        Set .TreeRec = rsProvider
        .StrNode = "������ҩ����"
        .lngMode = 0
        .Show 1, Me
        If .BlnSuccess = False Then
            Unload FrmSelect
            Exit Sub
        End If
        
        Me.txtDraw.Tag = .CurrentID
        Me.txtDraw = .CurrentName
    End With
    Unload FrmSelect
    
    Call GetDrawPerson(Me.txtDraw.Tag)
    cboDrawPerson.SetFocus
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cmdExpend_Click()
    mblnAutoExp = AutoExpend
    cmdExpend.Enabled = False
End Sub

Private Function AutoExpend(Optional blnCheck As Boolean = False) As Boolean
    '���ܣ��Զ��ֽ�
    Dim lng�ⷿID As Long, lngҩƷID As Long, lngҩƷID_Last As Long, lng���� As Long
    Dim bln�ⷿ As Boolean, bln���� As Boolean, blnʱ�� As Boolean, blnAddRow As Boolean
    Dim dbl��д���� As Double, dbl�������� As Double, Dbl���� As Double, dbl����ϵ�� As Double
    Dim dbl�ּ� As Currency, dbl�ּ�_ʱ�� As Double, dbl�ɱ��� As Double
    Dim lngCol As Long, lngCols As Long, lngRow As Long, lngRows As Long
    Dim rsTemp As New ADODB.Recordset
    Dim rsCheck As New ADODB.Recordset
    Dim dblʵ������ As Double
    Dim intCount As Integer
            
    '��ҩƷ��¼�����Զ��ֽ⣬����������ҩƷ
    On Error GoTo ErrHand
    Debug.Print "��ʼ�ֽ⣺" & Now
    Screen.MousePointer = 11
    lngRow = 1: lngCols = mshBill.Cols - 1
    lng�ⷿID = Val(cboStock.ItemData(cboStock.ListIndex))
    bln�ⷿ = CheckStockProperty(lng�ⷿID)
    
    Do While True
        lngҩƷID = Val(mshBill.TextMatrix(lngRow, 0))
        dbl�������� = Val(mshBill.TextMatrix(lngRow, mconIntCol��д����))
        dbl��д���� = dbl��������
        dbl����ϵ�� = Val(mshBill.TextMatrix(lngRow, mconIntCol����ϵ��))
        lng���� = Val(mshBill.TextMatrix(lngRow, mconIntCol����))
        
        If lngҩƷID = 0 Then Exit Do
        
        '��ȡ��ҩƷ���ڳ���ⷿ�Ƿ������ʱ�۵�����
        If lngҩƷID <> lngҩƷID_Last Then
            lngҩƷID_Last = lngҩƷID
            gstrSQL = " Select Nvl(A.ҩ�����,0) ҩ�����,Nvl(A.ҩ������,0) ҩ������," & _
                      " Nvl(B.�Ƿ���,0) ʱ��,Nvl(P.�ּ�,0) �ּ�,Nvl(A.�ɱ���,0) �ɱ���" & _
                      " From ҩƷ��� A,�շ���ĿĿ¼ B,�շѼ�Ŀ P" & _
                      " Where A.ҩƷID = B.ID And B.ID=P.�շ�ϸĿID And A.ҩƷID =[1] " & _
                      " And Sysdate between P.ִ������ And Nvl(P.��ֹ����,Sysdate)" & _
                      GetPriceClassString("P")
                      
            Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "[��ȡ��ҩƷ���ڳ���ⷿ�Ƿ������ʱ�۵�����]", lngҩƷID)
            
            blnʱ�� = (rsTemp!ʱ�� = 1)
            dbl�ּ� = rsTemp!�ּ� * dbl����ϵ��
            bln���� = IIf(bln�ⷿ, (rsTemp!ҩ����� = 1), (rsTemp!ҩ������ = 1))
        End If
        
        '��ȡ�������
        blnAddRow = False
        
        If bln���� = True And lng���� = 0 Then
           gstrSQL = " Select Nvl(A.��������,0)/" & dbl����ϵ�� & " As ��������,Nvl(A.ʵ������,0)/" & dbl����ϵ�� & " As ʵ������," & _
                      " Nvl(A.ʵ�ʽ��,0) ʵ�ʽ��,Nvl(A.ʵ�ʲ��,0) ʵ�ʲ��, nvl(A.ƽ���ɱ���,0) ƽ���ɱ���," & _
                      " Nvl(A.����,0) ����,A.�ϴ����� As ����,to_char(A.Ч��,'yyyy-MM-dd') Ч��,A.�ϴβ��� As ����,A.ԭ����,NVL(A.�ϴι�Ӧ��ID,0) �ϴι�Ӧ��ID," & _
                      " A.��׼�ĺ�,Decode(Nvl(a.���ۼ�, 0), 0, Decode(Nvl(a.ʵ������, 0), 0, b.�ּ�, a.ʵ�ʽ�� / a.ʵ������), a.���ۼ�)*" & dbl����ϵ�� & " As ���ۼ� " & _
                      " From ҩƷ��� A, �շѼ�Ŀ B Where a.ҩƷid = b.�շ�ϸĿid And a.�ⷿID=[1] And a.ҩƷID=[2] And a.����=1 " & _
                      " And Nvl(a.��������,0)>0 And ((Sysdate Between b.ִ������ And b.��ֹ����) Or b.��ֹ���� Is Null) " & _
                      GetPriceClassString("B") & _
                      " Order by " & IIf(gtype_UserSysParms.P150_ҩƷ���������㷨 = 0, " Nvl(A.����,0)", " A.Ч��,Nvl(A.����,0)")
            Set rsCheck = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "[��ȡ��ҩƷ��ָ���������п���¼]", lng�ⷿID, lngҩƷID)
            With rsCheck
                intCount = 0
                Do While Not .EOF
                    intCount = intCount + 1
                    mshBill.Redraw = False
                    '����д��¼
                    blnAddRow = False
                    If .AbsolutePosition <> 1 Then
                        mshBill.MsfObj.AddItem "", lngRow
                        For lngCol = 0 To lngCols
                            mshBill.TextMatrix(lngRow, lngCol) = mshBill.TextMatrix(lngRow - 1, lngCol)
                        Next
                        mshBill.TextMatrix(lngRow, mconIntCol��д����) = "0"
                        mshBill.RowData(lngRow) = mshBill.RowData(lngRow - 1)
                    End If
                    
                    If intCount = 1 Then
                        dblʵ������ = Val(mshBill.TextMatrix(lngRow, mconIntColʵ������))
                    End If
                    
                    '��д���������Ϣ
                    mshBill.TextMatrix(lngRow, mconIntCol�к�) = lngRow
                    mshBill.TextMatrix(lngRow, mconIntCol���) = (lngRow - 1) * 2 + 1
                    mshBill.TextMatrix(lngRow, mconIntCol����) = rsCheck!����
                    mshBill.TextMatrix(lngRow, mconIntCol����) = IIf(IsNull(rsCheck!����), "", rsCheck!����)
                    mshBill.TextMatrix(lngRow, mconIntCol����) = IIf(IsNull(rsCheck!����), "", rsCheck!����)
                    mshBill.TextMatrix(lngRow, mconIntColԭ����) = IIf(IsNull(rsCheck!ԭ����), "", rsCheck!ԭ����)
                    mshBill.TextMatrix(lngRow, mconIntColЧ��) = IIf(IsNull(rsCheck!Ч��), "", rsCheck!Ч��)
                    If gtype_UserSysParms.P149_Ч����ʾ��ʽ = 1 And mshBill.TextMatrix(lngRow, mconIntColЧ��) <> "" Then
                        '����Ϊ��Ч��
                        mshBill.TextMatrix(lngRow, mconIntColЧ��) = Format(DateAdd("D", -1, mshBill.TextMatrix(lngRow, mconIntColЧ��)), "yyyy-mm-dd")
                    End If
                    
                    mshBill.TextMatrix(lngRow, mconIntCol��׼�ĺ�) = IIf(IsNull(rsCheck!��׼�ĺ�), "", rsCheck!��׼�ĺ�)
                    
                    dbl�ּ�_ʱ�� = rsCheck!���ۼ�
                    
                    If dbl��д���� <= rsCheck!�������� Then
                        Dbl���� = dbl��д����
                    Else
                        Dbl���� = rsCheck!��������
                    End If
                    If Dbl���� > dbl��д���� Then Dbl���� = dbl��д����
                    
                    If Dbl���� <> mshBill.TextMatrix(lngRow, mconIntColʵ������) Then
                        mshBill.TextMatrix(lngRow, mconintCol��ʵ����) = zlStr.FormatEx(Dbl���� * dbl����ϵ��, mintNumberDigit, , True)
                    End If
                    
                    mshBill.TextMatrix(lngRow, mconIntCol��д����) = zlStr.FormatEx(Dbl����, mintNumberDigit, , True)
                    mshBill.TextMatrix(lngRow, mconIntColʵ������) = zlStr.FormatEx(Dbl����, mintNumberDigit, , True)
                                        
                    If Trim(mshBill.TextMatrix(lngRow, mconIntColʵ������)) = "" Then mshBill.TextMatrix(lngRow, mconIntColʵ������) = 0
                    
                    mshBill.TextMatrix(lngRow, mconIntColʵ�ʲ��) = zlStr.FormatEx(rsCheck!ʵ�ʲ��, mintMoneyDigit, , True)
                    mshBill.TextMatrix(lngRow, mconIntColʵ�ʽ��) = zlStr.FormatEx(rsCheck!ʵ�ʽ��, mintMoneyDigit, , True)
                    mshBill.TextMatrix(lngRow, mconIntCol��������) = zlStr.FormatEx(rsCheck!��������, mintMoneyDigit, , True)
                    mshBill.TextMatrix(lngRow, mconIntCol�ۼ�) = zlStr.FormatEx(IIf(blnʱ��, dbl�ּ�_ʱ��, dbl�ּ�), mintPriceDigit, , True)
                    mshBill.TextMatrix(lngRow, mconIntCol�ۼ۽��) = zlStr.FormatEx(Val(mshBill.TextMatrix(lngRow, mconIntCol�ۼ�)) * Dbl����, mintMoneyDigit, , True)
                    If Dbl���� <> 0 Then
                        mshBill.TextMatrix(lngRow, mconIntCol�ɹ���) = zlStr.FormatEx(rsCheck!ƽ���ɱ��� * dbl����ϵ��, mintCostDigit, , True)
                    End If
                    mshBill.TextMatrix(lngRow, mconIntCol�ɹ����) = zlStr.FormatEx(Val(mshBill.TextMatrix(lngRow, mconIntCol�ɹ���)) * Dbl����, mintMoneyDigit, , True)
                    mshBill.TextMatrix(lngRow, mconintCol���) = zlStr.FormatEx(Val(mshBill.TextMatrix(lngRow, mconIntCol�ۼ۽��)) - Val(mshBill.TextMatrix(lngRow, mconIntCol�ɹ����)), mintMoneyDigit, , True)
                    
                    dbl��д���� = dbl��д���� - Dbl����
                    dbl�������� = dbl�������� - Dbl����
                    If dbl��д���� = 0 Then Exit Do
                    lngRow = lngRow + 1
                    blnAddRow = True
                    .MoveNext
                Loop
                If dbl�������� <> 0 And rsCheck.RecordCount <> 0 Then
                    If blnAddRow Then
                        mshBill.TextMatrix(lngRow - 1, mconIntCol��д����) = zlStr.FormatEx(dbl�������� + Dbl����, mintNumberDigit, , True)
                    Else
                        mshBill.TextMatrix(lngRow, mconIntCol��д����) = zlStr.FormatEx(dbl�������� + Dbl����, mintNumberDigit, , True)
                    End If
                End If
            End With
            
            '�������¼Ϊ�㣬��˵��δ���зֽ⣬��Ҫ������������ʵ��������Ϊ��
            If dbl��д���� <> 0 And rsCheck.RecordCount = 0 Then
                mshBill.TextMatrix(lngRow, mconIntCol�к�) = lngRow
                mshBill.TextMatrix(lngRow, mconIntCol���) = (lngRow - 1) * 2 + 1
                mshBill.TextMatrix(lngRow, mconIntColʵ������) = 0
                mshBill.TextMatrix(lngRow, mconIntCol�ۼ۽��) = ""
                mshBill.TextMatrix(lngRow, mconIntCol�ɹ����) = ""
                mshBill.TextMatrix(lngRow, mconintCol���) = ""
            End If
        Else
            mshBill.TextMatrix(lngRow, mconIntCol�к�) = lngRow
            mshBill.TextMatrix(lngRow, mconIntCol���) = (lngRow - 1) * 2 + 1
        End If
            
        If blnAddRow = False Then lngRow = lngRow + 1
    Loop
    
    mblnChange = True
    AutoExpend = True
    mshBill.Redraw = True
    Call ShowColor
    Screen.MousePointer = 0
    Debug.Print "�����ֽ⣺" & Now
    Exit Function
ErrHand:
    Screen.MousePointer = 0
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub ShowColor(Optional ByVal lngCurRow As Long = 0)
    '�ڲ��Ļ����ʱ������治��ļ�¼�԰���ɫ��ʾ����
    Dim lngSelect_Row  As Long, lngSelect_Col As Long
    Dim lngҩƷID As Long
    Dim lngColor As Long, lngNewColor As Long '������ڵ���ɫ��Ҫ�ϵ���ɫһ�����򲻴���
    Dim dbl��д���� As Double, dbl�������� As Double
    Dim lngRow As Long, BlnDO As Boolean
    Dim i As Long, j As Long
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo ErrHand
    mshBill.Redraw = False
    mblnEnterCell = False
    lngSelect_Row = mshBill.Row: lngSelect_Col = mshBill.Col
    lngRow = IIf(lngCurRow > 0, lngCurRow, 1)
    
    Do While True
        If lngRow > mshBill.rows - 1 Then Exit Do
        mshBill.Row = lngRow: mshBill.Col = mconIntColҩ��
        lngColor = mshBill.MsfObj.CellForeColor
        
        lngҩƷID = Val(mshBill.TextMatrix(lngRow, 0))
        dbl��д���� = Val(mshBill.TextMatrix(lngRow, mconIntCol��д����))
        dbl�������� = Val(mshBill.TextMatrix(lngRow, mconIntCol��������))
        If lngҩƷID = 0 Then Exit Do
        
        gstrSQL = "select decode(ҩ�����,Null,0,ҩ�����) ҩ�����,decode(ҩ������,Null,0,ҩ������) ҩ������ from ҩƷ��� where ҩƷid=[1]"
        Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, "��ѯ����", lngҩƷID)
        
        If rsTemp Is Nothing Then
            Exit Sub
        Else
            If rsTemp!ҩ����� = 1 Or rsTemp!ҩ������ = 1 Then
                '��治���ҩƷ������ɫ
                BlnDO = False
                If dbl�������� < dbl��д���� Then BlnDO = True
                lngNewColor = IIf(BlnDO, &HC0, &H0)
                If lngColor <> lngNewColor Then
                    'ֻ��ҩ���н�����ɫ����
                    j = mshBill.ColData(mconIntColҩ��)
                    If j = 5 Then mshBill.ColData(mconIntColҩ��) = 0
                    mshBill.Col = mconIntColҩ��
                    mshBill.MsfObj.CellForeColor = lngNewColor
                    mshBill.ColData(mconIntColҩ��) = j
                End If
                
                If lngCurRow > 0 Then Exit Do
                lngRow = lngRow + 1
            Else
                Exit Do
            End If
        End If
    Loop
    mshBill.Row = lngSelect_Row: mshBill.Col = lngSelect_Col
    mshBill.Redraw = True
    mblnEnterCell = True
    Exit Sub
ErrHand:
    mshBill.Redraw = True
    mblnEnterCell = True
    If ErrCenter = 1 Then Resume
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

'
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
    Dim bln�ⷿ As Boolean
    Dim bln���� As Boolean
    Dim intRow As Integer
    Dim lngҩƷID As Long
    Dim rsTemp As ADODB.Recordset
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
    For intLop = 1 To mshBill.rows - 1
        If mshBill.TextMatrix(intLop, 0) <> "" Then '��ҩƷ
            Call AutoAdjustPrice_ByID(Val(mshBill.TextMatrix(intLop, 0)))
        End If
    Next
    
    If mint�༭״̬ = 3 Then        '���
        mstrTime_End = GetBillInfo(7, mstr���ݺ�)
        If mstrTime_End = "" Then
            MsgBox "�õ����Ѿ�����������Աɾ����", vbInformation, gstrSysName
            Exit Sub
        End If
        
        If mstrTime_End > mstrTime_Start Then
            MsgBox "�õ����Ѿ�����������Ա�༭�����˳������ԣ�", vbInformation, gstrSysName
            Exit Sub
        End If
        
        '�������ʱ��Ҫ���ж��Ƿ�������û�����εļ�¼
        If cmdExpend.Enabled = True Then
            bln�ⷿ = CheckStockProperty(cboStock.ItemData(cboStock.ListIndex))
            With mshBill
                For intRow = 1 To .rows - 1
                    lngҩƷID = Val(.TextMatrix(intRow, 0))
                    If lngҩƷID <> 0 Then
                        gstrSQL = " Select Nvl(A.ҩ�����,0) ҩ�����,Nvl(A.ҩ������,0) ҩ������" & _
                                          " From ҩƷ��� A" & _
                                          " Where A.ҩƷID =[1] "
                        Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "[ȡ��������]", lngҩƷID)
                        bln���� = IIf(bln�ⷿ, (rsTemp!ҩ����� = 1), (rsTemp!ҩ������ = 1))
                        If bln���� = True And Val(.TextMatrix(intRow, mconIntCol����)) = 0 Then
                            MsgBox .TextMatrix(intRow, mconIntColҩƷ����) & "�ǲ�����������ҩƷ�������Զ��ֽ������ˣ�", vbInformation, gstrSysName
                            Exit Sub
                        End If
                    End If
                Next
            End With
        End If
        
'        If Not ��鵥��(7, txtNo, False) And Not mblnUpdate Then
'            '�����µļ۸���µ����壬�˳���Ŀ�������û���һ�����յĵ���
'            MsgBox "�м�¼δʹ�����¼۸񣬳����Զ���ɸ��£��ۼۡ��ɱ��ۡ��ۼ۽��ɱ�����ۣ������º����飡", vbInformation, gstrSysName
'            Call RefreshBill
'            mblnUpdate = True
'            mblnChange = True
'            Exit Sub
'        End If
        
        If ���۸� Then
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
        
        '������
        If mint���÷�ʽ = 1 Then
            If Check���� = False Then
                MsgBox str���治����ʾ
                Exit Sub
            End If
        End If
        
        blnTrans = True
        gcnOracle.BeginTrans
        '������ʱ�޸��˵��ݣ����������ɵ��ݱ���
        If mblnChange Or mblnAutoExp = True Then
            If Not SaveCard(True) Then
                gcnOracle.RollbackTrans: Exit Sub
            End If
        End If
        
        If Not SaveCheck Then
            gcnOracle.RollbackTrans: Exit Sub
        End If
        
        If Val(zlDataBase.GetPara("��˴�ӡ", glngSys, ģ���.ҩƷ����)) = 1 Then
            '��ӡ
            If zlStr.IsHavePrivs(mstrPrivs, "���ݴ�ӡ") Then
                printbill
                
                If Val(zlDataBase.GetPara("��ӡҩƷ����", glngSys, ģ���.ҩƷ����)) = 1 And zlStr.IsHavePrivs(mstrPrivs, "ҩƷ�����ӡ") Then
                    '��ҩƷID˳���������
                    recSort.Sort = "ҩƷid"
                    recSort.MoveFirst
                    '��ӡҩƷ����
                    Do While Not recSort.EOF
                        If lng�ϴ�ҩƷID <> Val(recSort!ҩƷID) Then
                            ReportOpen gcnOracle, glngSys, "ZL1_INSIDE_1305_2", Me, "ҩƷ=" & Val(recSort!ҩƷID), 2
                            lng�ϴ�ҩƷID = recSort!ҩƷID
                        End If
                        recSort.MoveNext
                    Loop
                End If
                
            End If
        End If
        
        gcnOracle.CommitTrans
        Unload Me
        Exit Sub
    End If
    
    If mint�༭״̬ = 6 Then '����
        If mblnChange = False And mint����ʽ <> 2 Then
            MsgBox "��¼�����������", vbInformation + vbOKOnly, gstrSysName
            Exit Sub
        End If
        
        If LenB(StrConv(txt����ԭ��.Text, vbFromUnicode)) > txt����ԭ��.MaxLength Then
            MsgBox "����ԭ�򳬳�,���������" & CInt(txt����ԭ��.MaxLength / 2) & "�����ֻ�" & txt����ԭ��.MaxLength & "���ַ�!", vbInformation + vbOKOnly, gstrSysName
            txt����ԭ��.SetFocus
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
'        If Not ��鵥��(7, txtNo, False) And Not mblnUpdate Then
'            '�����µļ۸���µ����壬�˳���Ŀ�������û���һ�����յĵ���
'            MsgBox "�м�¼δʹ�����¼۸񣬳����Զ���ɸ��£��ۼۡ��ɱ��ۡ��ۼ۽��ɱ�����ۣ������º����飡", vbInformation, gstrSysName
'            Call RefreshBill
'            mblnUpdate = True
'            mblnChange = True
'            Exit Sub
'        End If
        
        If ���۸� Then
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
            
        If Val(zlDataBase.GetPara("���̴�ӡ", glngSys, ģ���.ҩƷ����)) = 1 Then
            '��ӡ
            If zlStr.IsHavePrivs(mstrPrivs, "���ݴ�ӡ") Then
                printbill
                
                If Val(zlDataBase.GetPara("��ӡҩƷ����", glngSys, ģ���.ҩƷ����)) = 1 And zlStr.IsHavePrivs(mstrPrivs, "ҩƷ�����ӡ") Then
                    '��ҩƷID˳���������
                    recSort.Sort = "ҩƷid"
                    recSort.MoveFirst
                    '��ӡҩƷ����
                    Do While Not recSort.EOF
                        If lng�ϴ�ҩƷID <> Val(recSort!ҩƷID) Then
                            ReportOpen gcnOracle, glngSys, "ZL1_INSIDE_1305_2", Me, "ҩƷ=" & Val(recSort!ҩƷID), 2
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

    txtDraw.Text = ""
    txtDraw.Tag = "0"
    txtժҪ.Text = ""
    txtDraw.SetFocus
    txtDraw.SelStart = 0
    txtDraw.SelLength = Len(txtDraw.Text)
    
    mblnChange = False
    If txtNo.Tag <> "" Then Me.staThis.Panels(2).Text = "��һ�ŵ��ݵ�NO�ţ�" & txtNo.Tag
    
    Exit Sub
ErrHand:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter = 1 Then Resume
End Sub

Private Function CheckStock() As Boolean
    Dim dbl����ϵ�� As Double, dblʵ������ As Double, dbl��д���� As Double
    Dim lngRow As Long, lngRows As Long, int����� As Integer
    Dim lngҩƷID As Long, lng�ⷿID As Long, lng���� As Long
    Dim bln�ⷿ As Boolean, bln��ҩ As Boolean
    Dim strҩƷID As String, strMsg As String
    Dim rsTemp As ADODB.Recordset
    Dim rsProperty As ADODB.Recordset           'ҩƷ���
    Dim rsCheck As ADODB.Recordset              'ҩƷ���
    Dim arrDrugID As Variant
    Dim i As Integer
    
    On Error GoTo errHandle

    Set rsProperty = New ADODB.Recordset
    With rsProperty
        If .State = 1 Then .Close
        .Fields.Append "ҩƷID", adDouble, 18, adFldIsNullable
        .Fields.Append "ͨ����", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "ҩ�����", adDouble, 18, adFldIsNullable
        .Fields.Append "ҩ������", adDouble, 18, adFldIsNullable
        .Fields.Append "�Ƿ���", adDouble, 18, adFldIsNullable

        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With

    Set rsCheck = New ADODB.Recordset
    With rsCheck
        If .State = 1 Then .Close
        .Fields.Append "ҩƷID", adDouble, 18, adFldIsNullable
        .Fields.Append "����", adDouble, 18, adFldIsNullable
        .Fields.Append "ʵ������", adDouble, 18, adFldIsNullable

        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With

    arrDrugID = Array()

    '��鵥���и�ҩƷ�Ŀ��
    'mint�����:0-�����;1-��飬�������ѣ�2-��飬�����ֹ
    '������ʱ��ҩƷ���ܴ���
    Debug.Print "��ʼ�����:" & Now
    lngRows = mshBill.rows - 1
    lng�ⷿID = Val(cboStock.ItemData(cboStock.ListIndex))
    bln�ⷿ = CheckStockProperty(lng�ⷿID)
    For lngRow = 1 To lngRows
        lngҩƷID = Val(mshBill.TextMatrix(lngRow, 0))
        If lngҩƷID <> 0 Then
            If InStr(1, "," & strҩƷID & ",", "," & lngҩƷID & ",") = 0 Then
                If Len(IIf(strҩƷID = "", "", strҩƷID & ",") & lngҩƷID) > 4000 Then
                    ReDim Preserve arrDrugID(UBound(arrDrugID) + 1)
                    arrDrugID(UBound(arrDrugID)) = strҩƷID
                    strҩƷID = lngҩƷID
                Else
                    strҩƷID = IIf(strҩƷID = "", "", strҩƷID & ",") & lngҩƷID
                End If
            End If
        End If
    Next

    If strҩƷID = "" And UBound(arrDrugID) < 0 Then
        CheckStock = True
        Exit Function
    ElseIf strҩƷID <> "" Then
        ReDim Preserve arrDrugID(UBound(arrDrugID) + 1)
        arrDrugID(UBound(arrDrugID)) = strҩƷID
    End If

    '��ȡ������������ҩƷ������
    gstrSQL = " Select A.ҩƷID,'['||B.����||']'||B.���� ͨ����,A.ҩ�����,A.ҩ������,B.�Ƿ���" & _
              " From ҩƷ��� A,�շ���ĿĿ¼ B " & _
              " Where A.ҩƷID=B.ID And A.ҩƷID in(select * from Table(Cast(f_Num2List([1]) As zlTools.t_NumList))) "

    For i = 0 To UBound(arrDrugID)
        Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, "��ȡ������������ҩƷ������", CStr(arrDrugID(i)))

        If Not rsTemp.EOF Then
            Do While Not rsTemp.EOF
                With rsProperty
                    .AddNew
                    !ҩƷID = rsTemp!ҩƷID
                    !ͨ���� = rsTemp!ͨ����
                    !ҩ����� = rsTemp!ҩ�����
                    !ҩ������ = rsTemp!ҩ������
                    !�Ƿ��� = rsTemp!�Ƿ���

                    .Update
                End With
                rsTemp.MoveNext
            Loop
        End If
    Next

    gstrSQL = "Select a.ҩƷid, Nvl(a.����, 0) As ����, Sum(Nvl(ʵ������, 0)) As ʵ������ " & _
        " From ҩƷ��� A, ҩƷ��� C " & _
        " Where a.�ⷿid = [1] And a.ҩƷid = c.ҩƷid And a.���� = 1 And c.ҩƷid in (select * from Table(Cast(f_Num2list([2]) As Zltools.t_Numlist))) " & _
        " Group By a.ҩƷid, Nvl(a.����, 0) "
    For i = 0 To UBound(arrDrugID)
        Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "[��ȡ������������ҩƷ�ĵ�ǰ���]", lng�ⷿID, CStr(arrDrugID(i)))

        If Not rsTemp.EOF Then
            Do While Not rsTemp.EOF
                With rsCheck
                    .AddNew
                    !ҩƷID = rsTemp!ҩƷID
                    !���� = rsTemp!����
                    !ʵ������ = rsTemp!ʵ������

                    .Update
                End With
                rsTemp.MoveNext
            Loop
        End If
    Next

    '���ÿ��ҩƷ
    For lngRow = 1 To lngRows
        lngҩƷID = Val(mshBill.TextMatrix(lngRow, 0))
        If lngҩƷID <> 0 Then
            lng���� = Val(mshBill.TextMatrix(lngRow, mconIntCol����))
            dbl����ϵ�� = Val(mshBill.TextMatrix(lngRow, mconIntCol����ϵ��))
            dbl��д���� = Val(mshBill.TextMatrix(lngRow, mconIntColʵ������))

            dblʵ������ = 0
            '���Ҹ�ҩƷ�Ŀ���¼
            rsCheck.Filter = "ҩƷID=" & lngҩƷID & " And ����=" & lng����
            If rsCheck.RecordCount <> 0 Then
                dblʵ������ = Val(zlStr.FormatEx(nvl(rsCheck!ʵ������, 0) / dbl����ϵ��, mintNumberDigit, , True))
            End If

            '�������ʵ����������
            If Not (dblʵ������ >= dbl��д����) Then
                int����� = mint�����
                '�����ҩƷ��ʱ�ۻ��������治�㲻������⣬�൱�ڽ�ֹ����
                rsProperty.Filter = "ҩƷID=" & lngҩƷID
                bln��ҩ = (IIf(bln�ⷿ, (rsProperty!ҩ����� = 1), (rsProperty!ҩ������ = 1)) Or (rsProperty!�Ƿ��� = 1))
                strMsg = ""
                If bln��ҩ Then
                    int����� = 2
                    '���������ҩƷ��������С�ڵ����㣬˵��δִ�зֽ⹦��
                    If lng���� <= 0 And IIf(bln�ⷿ, (rsProperty!ҩ����� = 1), (rsProperty!ҩ������ = 1)) Then
                        strMsg = "������ִ�зֽ⹦����ȷ����ҩƷ�ĳ������Σ�"
                    End If
                End If

                '���������̽�����ʾ���ֹ
                Select Case int�����
                Case 1  '����ʾ
                    Debug.Print "�޿���˳�:" & Now
                    If MsgBox(rsProperty!ͨ���� & "�Ŀ�治�㣬�Ƿ������" & strMsg, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                        Exit Function
                    End If
                Case 2
                    Debug.Print "�޿���˳�:" & Now
                    MsgBox rsProperty!ͨ���� & "�Ŀ�治�㣡" & strMsg, vbInformation, gstrSysName
                    Exit Function
                End Select
            End If
        End If
    Next

    rsCheck.Filter = 0
    rsProperty.Filter = 0
    CheckStock = True
    Debug.Print "��ɼ����:" & Now
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function CheckStockProperty(ByVal lng�ⷿID As Long) As Boolean
    Dim rsCheck As New ADODB.Recordset
    
    On Error GoTo errHandle

    '���ָ���ⷿ��ҩ�⡢ҩ�������Ƽ���(����Ŀⷿ�϶���ҩ�⡢ҩ�����Ƽ����е�һ��)
    gstrSQL = " Select ����ID From ��������˵�� " & _
              " Where (�������� like '%ҩ��' Or �������� like '%�Ƽ���') And ����id=[1] "
    Set rsCheck = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "[�ж��ǲ���ҩ�����Ƽ���]", lng�ⷿID)
              
    If rsCheck.EOF Then
        CheckStockProperty = True
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Dim rsEnterStock As New Recordset
    Dim rsDetail As New Recordset
    Dim bln��ҩ�ⷿ As Boolean
    Dim str�ⷿ���� As String
    
    mblnLoad = False
    mblnUpdate = False
    mblnEnterCell = False
    mint���淽ʽ = Val(zlDataBase.GetPara("������������", glngSys, ģ���.ҩƷ����))
    mblnViewCost = zlStr.IsHavePrivs(mstrPrivs, "�鿴�ɱ���")
    txtNo = mstr���ݺ�
    txtNo.Tag = txtNo
    
    On Error GoTo errHandle

    mintDrugNameShow = Int(Val(GetSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\" & "ҩƷ���ù���", "ҩƷ������ʾ��ʽ", 0)))
    If mintDrugNameShow > 2 Or mintDrugNameShow < 0 Then mintDrugNameShow = 0
    mnuColDrug.Item(mintDrugNameShow).Checked = True
    
    mlng����ⷿ = mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex)
    Call GetDrugDigit(mlng����ⷿ, MStrCaption, mintUnit, mintCostDigit, mintPriceDigit, mintNumberDigit, mintMoneyDigit)
    
    mint�����γ��� = Val(zlDataBase.GetPara("ҩƷ�����γ���", glngSys, 1305, 0))
        
    initCard
    
    mstrTime_Start = GetBillInfo(7, mstr���ݺ�)
    RestoreWinState Me, App.ProductName, MStrCaption
    
    'ֻ����ҩ��ⷿ����ʾ"ԭ����"��
    str�ⷿ���� = ""
    gstrSQL = "Select a.�������� From ��������˵�� A Where a.����id =[1]"
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
    
    '��Ʒ���д���
    If gintҩƷ������ʾ = 2 Then
        '��ʾ��Ʒ����
        mshBill.ColWidth(mconIntCol��Ʒ��) = IIf(mshBill.ColWidth(mconIntCol��Ʒ��) = 0, 2000, mshBill.ColWidth(mconIntCol��Ʒ��))
    Else
        '��������ʾ��Ʒ����
        mshBill.ColWidth(mconIntCol��Ʒ��) = 0
    End If
    
    mblnEnterCell = True
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
    Dim numUseAbleCount As Double
    Dim vardrug As Variant
    Dim strOrder As String, strCompare As String
    Dim str���� As String, strArray As String
    Dim intCostDigit As Integer        '�ɱ���С��λ��
    Dim intPriceDigit As Integer       '�ۼ�С��λ��
    Dim intNumberDigit As Integer      '����С��λ��
    Dim intMoneyDigit As Integer       '���С��λ��
    Dim strҩ�� As String
    Dim strSqlOrder As String
    
    '�ⷿ
    On Error GoTo errHandle
    strOrder = zlDataBase.GetPara("����", glngSys, ģ���.ҩƷ����)
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
            
            If Not mblnStock Then
                Me.txtDraw.Tag = UserInfo.����ID
                Me.txtDraw.Text = UserInfo.��������
                Call GetDrawPerson(UserInfo.����ID)
            End If
            
            initGrid
        Case 2, 3, 4, 6
            initGrid
            
            If mint�༭״̬ = 4 Then
                gstrSQL = "select b.id,b.���� from ҩƷ�շ���¼ a,���ű� b where a.�ⷿid=b.id and A.���� = 7 and a.no=[1]"
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
                    strUnitQuantity = "F.���㵥λ AS ��λ, A.��д����,a.ʵ������,a.�ɱ���,a.���ۼ�,'1' as ����ϵ��,"
                Case mconint���ﵥλ
                    strUnitQuantity = "B.���ﵥλ AS ��λ,(A.��д���� / B.�����װ) AS ��д����,(A.ʵ������ / B.�����װ) AS ʵ������,a.�ɱ���*B.�����װ as �ɱ���,a.���ۼ�*B.�����װ as ���ۼ�,B.�����װ as ����ϵ��,"
                Case mconintסԺ��λ
                    strUnitQuantity = "B.סԺ��λ AS ��λ,(A.��д���� / B.סԺ��װ) AS ��д����,(A.ʵ������ / B.סԺ��װ) AS ʵ������,a.�ɱ���*B.סԺ��װ as �ɱ���,a.���ۼ�*B.סԺ��װ as ���ۼ�,B.סԺ��װ as ����ϵ��,"
                Case mconintҩ�ⵥλ
                    strUnitQuantity = "B.ҩ�ⵥλ AS ��λ,(A.��д���� / B.ҩ���װ) AS ��д����,(A.ʵ������ / B.ҩ���װ) AS ʵ������,a.�ɱ���*B.ҩ���װ as �ɱ���,a.���ۼ�*B.ҩ���װ as ���ۼ�,B.ҩ���װ as ����ϵ��,"
            End Select
            
            If mint�༭״̬ <> 6 Then
                gstrSQL = "SELECT W.*,Z.��������/W.����ϵ�� ��������,Z.ʵ�ʽ��,Z.ʵ�ʲ�� " & _
                    " FROM " & _
                    "     (SELECT DISTINCT A.ҩƷID,A.���,'[' || F.���� || ']' As ҩƷ����, F.���� As ͨ����, E.���� As ��Ʒ��," & _
                    "     NVL(E.����,F.����) ����,B.ҩƷ��Դ,B.����ҩ��,F.���,F.���� AS ԭ������,A.����, A.ԭ����, A.����,Nvl(A.����,0) As ����,B.�ӳ���,A.Ч��," & _
                    strUnitQuantity & _
                    "     A.�ɱ����,A.���۽��, A.���,A.ժҪ,A.����ԭ��,������,��������,�޸���,�޸�����,�����,�������," & _
                    "     A.�ⷿID,A.�Է�����ID,C.���� AS ���ò���,F.�Ƿ���,B.ҩ������ As ҩ����������,NVL(A.������,'') As ������,A.��׼�ĺ�,A.��ҩ��ʽ,A.ʵ������ ԭʼ���� " & _
                    "     FROM ҩƷ�շ���¼ A, ҩƷ��� B,�շ���Ŀ���� E ,�շ���ĿĿ¼ F,���ű� C " & _
                    "     WHERE A.ҩƷID = B.ҩƷID AND B.ҩƷID=F.ID " & _
                    "     AND B.ҩƷID=E.�շ�ϸĿID(+) And E.����(+)=3 " & _
                    "     AND A.�Է�����ID=C.ID AND A.��¼״̬ =[3] " & _
                    "     AND A.���� = 7 AND A.NO = [1]) W,ҩƷ��� Z" & _
                    " WHERE W.ҩƷID=Z.ҩƷID(+) AND NVL(W.����,0)=Nvl(Z.����(+),0) " & _
                    " And Z.�ⷿID(+)=[2] And Z.����(+)=1" & _
                    " ORDER BY " & strSqlOrder
            Else
                '��������
                If mint����ʽ <> 2 Then
                    gstrSQL = "SELECT W.*,Z.��������/W.����ϵ�� ��������,Z.ʵ�ʽ��,Z.ʵ�ʲ�� " & _
                        " FROM " & _
                        "     (SELECT DISTINCT A.ҩƷID,A.���,'[' || F.���� || ']' As ҩƷ����, F.���� As ͨ����, E.���� As ��Ʒ��," & _
                        "     NVL(E.����,F.����) ����,B.ҩƷ��Դ,B.����ҩ��,F.���,F.���� AS ԭ������,A.����, A.ԭ����, A.����,NVL(A.����,0) ����,B.�ӳ���,A.Ч��," & _
                        strUnitQuantity & _
                        "     A.�ɱ����,0 ���۽��,0 ���,A.ժҪ,A.�ⷿID,A.�Է�����ID,C.���� AS ���ò���,F.�Ƿ���,B.ҩ������ AS ҩ����������,A.������,A.��׼�ĺ�,A.��ҩ��ʽ,A.��д���� ԭʼ���� " & _
                        "     FROM " & _
                        "         (SELECT MIN(ID) AS ID, SUM(ʵ������) AS ��д����,0 ʵ������,SUM(�ɱ����) AS �ɱ����,ҩƷID,���,����, ԭ����,����,Ч��,NVL(����,0) ����,����,�ɱ���,���ۼ�,ժҪ,�ⷿID,�Է�����ID,������ID,NVL(X.������,'') As ������,X.��׼�ĺ�,X.��ҩ��ʽ " & _
                        "         FROM ҩƷ�շ���¼ X " & _
                        "         WHERE NO=[1] AND ����=7  " & _
                        "         GROUP BY ҩƷID,���,����,ԭ����,����,Ч��,NVL(����,0),����,�ɱ���,���ۼ�,ժҪ,�ⷿID,�Է�����ID,������ID,������,��׼�ĺ�,��ҩ��ʽ" & _
                        "         HAVING SUM(ʵ������)<>0 ) A," & _
                        "         ҩƷ��� B,�շ���Ŀ���� E ,�շ���ĿĿ¼ F,���ű� C " & _
                        "     WHERE A.ҩƷID = B.ҩƷID AND b.ҩƷID=F.ID AND A.�Է�����ID=C.ID " & _
                        "     AND B.ҩƷID=E.�շ�ϸĿID(+) AND E.����(+)=3 ) W,ҩƷ��� Z" & _
                        " WHERE W.ҩƷID=Z.ҩƷID(+) AND NVL(W.����,0)=Nvl(Z.����(+),0) " & _
                        " And Z.�ⷿID(+)=[2] And Z.����(+)=1" & _
                        " ORDER BY " & strSqlOrder
                Else
                    gstrSQL = "Select w.*, z.�������� / w.����ϵ�� ��������, z.ʵ�ʽ��, z.ʵ�ʲ��" & vbNewLine & _
                            "From (Select Distinct a.ҩƷid, a.���, '[' || f.���� || ']' As ҩƷ����, f.���� As ͨ����, e.���� As ��Ʒ��, Nvl(e.����, f.����) ����, b.ҩƷ��Դ," & vbNewLine & _
                            "                       b.����ҩ��, f.���, f.���� As ԭ������, a.����, a.ԭ����,a.����, Nvl(a.����, 0) ����, b.�ӳ���, a.Ч��, " & vbNewLine & _
                            strUnitQuantity & _
                            "                        a.�ɱ����, a.���۽��, a.���, a.ժҪ,a.����ԭ��, a.�ⷿid, a.�Է�����id, c.���� As ���ò���, f.�Ƿ���," & vbNewLine & _
                            "                       b.ҩ������ As ҩ����������, a.������, a.��׼�ĺ�, a.��ҩ��ʽ, a.��д���� ԭʼ����" & vbNewLine & _
                            "       From ҩƷ�շ���¼ a, ҩƷ��� b, �շ���Ŀ���� e, �շ���ĿĿ¼ f, ���ű� c" & vbNewLine & _
                            "       Where a.ҩƷid = b.ҩƷid And b.ҩƷid = f.Id And a.�Է�����id = c.Id And b.ҩƷid = e.�շ�ϸĿid(+) And e.����(+) = 3 And" & vbNewLine & _
                            "             a.��¼״̬ = [3] And a.���� = 7 And a.���ϵ�� = -1 And a.No = [1]) w, ҩƷ��� z" & vbNewLine & _
                            "Where w.ҩƷid = z.ҩƷid(+) And Nvl(w.����, 0) = Nvl(z.����(+), 0) And z.�ⷿid(+) = [2] And z.����(+) = 1" & vbNewLine & _
                            "Order By " & strSqlOrder
                End If
            End If
            
            Set rsInitCard = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption, mstr���ݺ�, cboStock.ItemData(cboStock.ListIndex), mint��¼״̬)
            
            If rsInitCard.EOF Then
                mintParallelRecord = 2
                Exit Sub
            End If
            
            Select Case mint�༭״̬
            Case 2, 6 '�޸ġ�����
                If mint�༭״̬ = 2 Then
                    Txt������ = rsInitCard!������ 'UserInfo.�û�����
                    Txt�������� = Format(rsInitCard!��������, "yyyy-mm-dd hh:mm:ss") 'Format(Sys.Currentdate, "yyyy-mm-dd hh:mm:ss")
                    Txt�޸��� = IIf(IsNull(rsInitCard!�޸���), "", rsInitCard!�޸���)
                    Txt�޸����� = IIf(IsNull(rsInitCard!�޸�����), "", Format(rsInitCard!�޸�����, "yyyy-mm-dd hh:mm:ss"))
                End If
                If mint�༭״̬ = 6 Then
                    Txt������ = UserInfo.�û�����
                    Txt�������� = Format(Sys.Currentdate, "yyyy-mm-dd hh:mm:ss")
                    Txt�޸��� = UserInfo.�û�����
                    Txt�޸����� = Format(Sys.Currentdate, "yyyy-mm-dd hh:mm:ss")
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
            
            If Not (mint�༭״̬ = 6 And mint����ʽ <> 2) Then txt����ԭ��.Text = IIf(IsNull(rsInitCard!����ԭ��), "", rsInitCard!����ԭ��)
            
            If (mint�༭״̬ = 2 Or mint�༭״̬ = 3) And Txt����� <> "" Then
                mintParallelRecord = 3
                Exit Sub
            End If
            
            txtDraw.Tag = rsInitCard!�Է�����id
            txtDraw.Text = rsInitCard!���ò���
            
            mint���÷�ʽ = IIf(IsNull(rsInitCard!��ҩ��ʽ), 0, rsInitCard!��ҩ��ʽ)
            LblTitle.Caption = GetUnitName & "ҩƷ���õ�" & IIf(mint���÷�ʽ = 0, "(�ⷿ����)", "(��������)")
            
            Call GetDrawPerson(txtDraw.Tag)
            cboDrawPerson.Text = IIf(IsNull(rsInitCard!������), "", rsInitCard!������)
            
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
                    .TextMatrix(intRow, mconIntColЧ��) = IIf(IsNull(rsInitCard!Ч��), "", Format(rsInitCard!Ч��, "yyyy-mm-dd"))
                    If gtype_UserSysParms.P149_Ч����ʾ��ʽ = 1 And .TextMatrix(intRow, mconIntColЧ��) <> "" Then
                        '����Ϊ��Ч��
                        .TextMatrix(intRow, mconIntColЧ��) = Format(DateAdd("D", -1, .TextMatrix(intRow, mconIntColЧ��)), "yyyy-mm-dd")
                    End If
                    
                    .TextMatrix(intRow, mconIntCol��д����) = zlStr.FormatEx(IIf(mint�༭״̬ = 6 And mint����ʽ = 2, -1, 1) * rsInitCard!��д����, intNumberDigit, , True)
                    .TextMatrix(intRow, mconIntColʵ������) = zlStr.FormatEx(IIf(mint�༭״̬ = 6 And mint����ʽ = 2, -1, 1) * rsInitCard!ʵ������, intNumberDigit, , True)
                    .TextMatrix(intRow, mconIntCol�ɹ���) = zlStr.FormatEx(rsInitCard!�ɱ���, intCostDigit, , True)
                    .TextMatrix(intRow, mconIntCol�ɹ����) = zlStr.FormatEx(IIf(mint�༭״̬ = 6 And mint����ʽ <> 2, 0, IIf(mint�༭״̬ = 6 And mint����ʽ = 2, -1, 1) * rsInitCard!�ɱ����), intMoneyDigit, , True)
                    .TextMatrix(intRow, mconIntCol�ۼ�) = zlStr.FormatEx(rsInitCard!���ۼ�, intPriceDigit, , True)
                    .TextMatrix(intRow, mconIntCol�ۼ۽��) = zlStr.FormatEx(IIf(mint�༭״̬ = 6 And mint����ʽ = 2, -1, 1) * rsInitCard!���۽��, intMoneyDigit, , True)
                    .TextMatrix(intRow, mconintCol���) = zlStr.FormatEx(IIf(mint�༭״̬ = 6 And mint����ʽ = 2, -1, 1) * rsInitCard!���, intMoneyDigit, , True)
                    .TextMatrix(intRow, mconIntCol����) = IIf(IsNull(rsInitCard!����), "0", rsInitCard!����)
                    .TextMatrix(intRow, mconIntCol����ϵ��) = rsInitCard!����ϵ��
                    .TextMatrix(intRow, mconIntcol�ӳ���) = rsInitCard!�ӳ��� / 100 & "||" & rsInitCard!�Ƿ��� & "||" & rsInitCard!ҩ����������
                    .TextMatrix(intRow, mconIntCol��������) = IIf(IsNull(rsInitCard!��������), "0", rsInitCard!��������)
                    .TextMatrix(intRow, mconIntColʵ�ʲ��) = IIf(IsNull(rsInitCard!ʵ�ʲ��), "0", rsInitCard!ʵ�ʲ��)
                    .TextMatrix(intRow, mconIntColʵ�ʽ��) = IIf(IsNull(rsInitCard!ʵ�ʽ��), "0", rsInitCard!ʵ�ʽ��)
                    .TextMatrix(intRow, mconIntCol��׼�ĺ�) = IIf(IsNull(rsInitCard!��׼�ĺ�), "", rsInitCard!��׼�ĺ�)
                    .TextMatrix(intRow, mconintCol��ʵ����) = IIf(IsNull(rsInitCard!ԭʼ����), "0", rsInitCard!ԭʼ����)
                    .TextMatrix(intRow, mconintColԭʼ����) = .TextMatrix(intRow, mconIntColʵ������)
                    
                    If mint�༭״̬ = 2 Or mint�༭״̬ = 3 Then
                        numUseAbleCount = 0
                        For Each vardrug In mcolUsedCount
                            If vardrug(0) = CStr(rsInitCard!ҩƷID & IIf(IsNull(rsInitCard!����), "0", rsInitCard!����)) Then
                                numUseAbleCount = vardrug(1)
                                mcolUsedCount.Remove vardrug(0)
                                Exit For
                            End If
                        Next
                        str���� = rsInitCard!ҩƷID & IIf(IsNull(rsInitCard!����), "0", rsInitCard!����)
                        If mint�༭״̬ = 2 Then
                            strArray = numUseAbleCount + IIf(IsNull(rsInitCard!��д����), "0", rsInitCard!��д����)
                        Else
                            strArray = numUseAbleCount + IIf(IsNull(rsInitCard!ʵ������), "0", rsInitCard!ʵ������)
                        End If
                        mcolUsedCount.Add Array(str����, strArray), str����
                    End If
                    
                    rsInitCard.MoveNext
                Loop
                .rows = intRow + 2
            End With
            rsInitCard.Close
            
            If mint�༭״̬ = 3 Then    '��˵������
                Call ShowColor
            End If
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
        .TextMatrix(0, mconIntCol��д����) = IIf(mint�༭״̬ = 6, "����", "��д����")
        .TextMatrix(0, mconIntColʵ������) = IIf(mint�༭״̬ = 6, "��������", "ʵ������")
        .TextMatrix(0, mconIntCol�ɹ���) = "�ɱ���"
        .TextMatrix(0, mconIntCol�ɹ����) = "�ɱ����"
        .TextMatrix(0, mconIntCol�ۼ�) = "�ۼ�"
        .TextMatrix(0, mconIntCol�ۼ۽��) = "�ۼ۽��"
        .TextMatrix(0, mconintCol���) = "���"
        .TextMatrix(0, mconIntCol��������) = "��������"
        .TextMatrix(0, mconIntColʵ�ʲ��) = "ʵ�ʲ��"
        .TextMatrix(0, mconIntColʵ�ʽ��) = "ʵ�ʽ��"
        .TextMatrix(0, mconIntcol�ӳ���) = "�ӳ���"
        .TextMatrix(0, mconIntCol����ϵ��) = "����ϵ��"
        .TextMatrix(0, mconIntCol����) = "����"
        .TextMatrix(0, mconintCol��ʵ����) = "��ʵ����"
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
        .ColWidth(mconIntCol��д����) = 1000
        .ColWidth(mconIntColʵ������) = 1000
        .ColWidth(mconIntCol�ɹ���) = 900
        .ColWidth(mconIntCol�ɹ����) = 900
        .ColWidth(mconIntCol�ۼ�) = 900
        .ColWidth(mconIntCol�ۼ۽��) = 900
        .ColWidth(mconintCol���) = 800
        .ColWidth(mconIntCol��������) = 0
        .ColWidth(mconIntColʵ�ʲ��) = 0
        .ColWidth(mconIntColʵ�ʽ��) = 0
        .ColWidth(mconIntcol�ӳ���) = 0
        .ColWidth(mconIntCol����ϵ��) = 0
        .ColWidth(mconIntCol����) = 0
        .ColWidth(mconintCol��ʵ����) = 0
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
        .ColData(mconintCol��ʵ����) = 5
        .ColData(mconintColԭʼ����) = 5
        
        .ColData(mconIntColҩƷ���������) = 5
        .ColData(mconIntColҩƷ����) = 5
        .ColData(mconIntColҩƷ����) = 5
        
        If mint�༭״̬ = 1 Or mint�༭״̬ = 2 Then
            txtDraw.Enabled = True
            cmdDraw.Enabled = True
            txtժҪ.Enabled = True
            
            cboStock.Enabled = True

            .ColData(mconIntColҩ��) = 1
            .ColData(mconIntCol��д����) = 4
            .ColData(mconIntColʵ������) = 5
        ElseIf mint�༭״̬ = 3 Then
            cboDrawPerson.Enabled = False
            
            cboStock.Enabled = False
            
            txtDraw.Enabled = False
            cmdDraw.Enabled = False
            txtժҪ.Enabled = False
            
            .ColData(mconIntCol��д����) = 5
            .ColData(mconIntColʵ������) = 4
        ElseIf mint�༭״̬ = 4 Then
            cboDrawPerson.Enabled = False
        
            cboStock.Enabled = False
            
            txtDraw.Enabled = False
            cmdDraw.Enabled = False
            txtժҪ.Enabled = False
            
            .ColData(mconIntCol��д����) = 5
            .ColData(mconIntColʵ������) = 5
        ElseIf mint�༭״̬ = 6 Then
            cboDrawPerson.Enabled = False
            
            cboStock.Enabled = False
            
            txtDraw.Enabled = False
            cmdDraw.Enabled = False
            txtժҪ.Enabled = True
            
            .ColData(mconIntColҩ��) = 5
            .ColData(mconIntCol��д����) = 5
            .ColData(mconIntColʵ������) = 5
                
            If mint����ʽ <> 2 Then
                .ColData(mconIntColʵ������) = 4
            End If
        End If
        
        .ColData(mconIntCol�ɹ���) = 5
        .ColData(mconIntCol�ɹ����) = 5
        .ColData(mconIntCol�ۼ�) = 5
        .ColData(mconIntCol�ۼ۽��) = 5
        .ColData(mconintCol���) = 5
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
        .ColAlignment(mconIntCol��д����) = flexAlignRightCenter
        .ColAlignment(mconIntColʵ������) = flexAlignRightCenter
        .ColAlignment(mconIntCol�ɹ���) = flexAlignRightCenter
        .ColAlignment(mconIntCol�ɹ����) = flexAlignRightCenter
        .ColAlignment(mconIntCol�ۼ�) = flexAlignRightCenter
        .ColAlignment(mconIntCol�ۼ۽��) = flexAlignRightCenter
        .ColAlignment(mconintCol���) = flexAlignRightCenter
        .ColAlignment(mconintCol��ʵ����) = flexAlignRightCenter
        
        .PrimaryCol = mconIntColҩ��
        .LocateCol = mconIntColҩ��
        If InStr(1, "34", mint�༭״̬) <> 0 Then .ColData(mconIntColҩ��) = 0
    End With
    txtժҪ.MaxLength = Sys.FieldsLength("ҩƷ�շ���¼", "ժҪ")
    txt����ԭ��.MaxLength = Sys.FieldsLength("ҩƷ�շ���¼", "����ԭ��")
    chkIn.Visible = (mint�༭״̬ = 1)
    txtIn.Visible = (mint�༭״̬ = 1)
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    If Me.WindowState = vbMinimized Then Exit Sub
    
    If Me.Width < 14775 Then
        Me.Width = 14775
        Exit Sub
    End If
    
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
    cboDrawPerson.Left = mshBill.Left + mshBill.Width - cboDrawPerson.Width
    lblDrawPerson.Left = cboDrawPerson.Left - lblDrawPerson.Width - 100
    
    LblEnterStock.Left = cboStock.Left + cboStock.Width + (lblDrawPerson.Left - cboStock.Left - cboStock.Width - LblEnterStock.Width - txtDraw.Width - cmdDraw.Width - 100) / 2
    txtDraw.Left = LblEnterStock.Left + LblEnterStock.Width + 100
    cmdDraw.Left = txtDraw.Left + txtDraw.Width
    
    With Lbl������
        .Top = Pic����.Height - 200 - .Height
        .Left = mshBill.Left + 180
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
    
    If txt����ԭ��.Visible Then '����ԭ��ɼ�������λ��
        With txt����ԭ��
            .Top = lblժҪ.Top - 140 - .Height
            .Left = txtժҪ.Left
            .Width = mshBill.Left + mshBill.Width - .Left
        End With
        
        With lbl����ԭ��
            .Top = txt����ԭ��.Top + 50
            .Left = txt����ԭ��.Left - .Width - 100
        End With
    End If
    
    With lblPurchasePrice
        .Left = mshBill.Left
        .Top = IIf(txt����ԭ��.Visible, txt����ԭ��.Top - 140, txtժҪ.Top - 60) - .Height
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
    
    With cmdExpend
        .Left = CmdSave.Left - CmdSave.Width - 500
        .Top = CmdCancel.Top
    End With
    
    With lblCode
        .Top = CmdCancel.Top + 50
    End With
    With txtCode
        .Top = CmdCancel.Top + 30
    End With
    
    With mshProvider
        If .Visible = True Then
            If .Tag = 0 Then
                .Width = LblEnterStock.Width + txtDraw.Width + cmdDraw.Width + 80
                .Top = txtDraw.Top + txtDraw.Height
                .Left = cmdDraw.Left + cmdDraw.Width - .Width
                .Redraw = True
            Else
                .Width = lblDrawPerson.Width + cboDrawPerson.Width + 80
                .Top = txtDraw.Top + txtDraw.Height
                .Left = lblDrawPerson.Left
            End If
        End If
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\ҩƷ���ù���", "ҩƷ������ʾ��ʽ", mintDrugNameShow)
    
    mblnAutoExp = False
    If mshProvider.Visible = True Then
        mshProvider.Visible = False
        txtDraw.SetFocus
        txtDraw.SelLength = Len(txtDraw.Text)
        txtDraw.SelStart = 0
        Cancel = True
        Exit Sub
    End If
    
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
    Dim rs��� As New Recordset
    Dim intRow As Integer
    
    Dim strNo As String
    Dim lng�ⷿID As Long
    Dim lng�Է�����id As Long
    Dim str����� As String
    Dim dat������� As String
    Dim int��� As Integer
    Dim lngҩƷID As Long
    Dim str���� As String
    Dim lng���� As Long
    Dim num��д���� As Double
    Dim numʵ������ As Double
    Dim num�ɱ��� As Double
    Dim num�ɱ���� As Double
    Dim num���۽�� As Double
    Dim num��� As Double
    Dim lng������id As Long
    Dim str���� As String
    Dim datЧ�� As String
    Dim arrSql As Variant
    Dim str��׼�ĺ� As String
    Dim n As Integer
    Dim strҩƷ As String
    
    mblnSave = False
    SaveCheck = False
    arrSql = Array()
    
    On Error GoTo errHandle
    lng�ⷿID = cboStock.ItemData(cboStock.ListIndex)
    lng�Է�����id = txtDraw.Tag
    str����� = UserInfo.�û�����
    strNo = txtNo.Tag
    gstrSQL = "SELECT b.id " _
            & " FROM ҩƷ�������� a, ҩƷ������ b " _
            & "Where a.���id = b.ID " _
            & "  AND a.���� = 7 "
    Call SQLTest(App.Title, "ҩƷ���õ�", gstrSQL)
    Set rs��� = zlDataBase.OpenSQLRecord(gstrSQL, "SaveCheck")
    Call SQLTest
    
    If rs���.EOF Then
        MsgBox "�Բ���û������ҩƷ���õ�����������ҩƷ�������!", vbExclamation + vbOKOnly, gstrSysName
        Exit Function
    End If
    lng������id = rs���!id
    rs���.Close
    
    With mshBill
        If .rows <= 1 Then Exit Function
        For intRow = 1 To .rows - 1
            If .TextMatrix(intRow, 0) <> "" Then
                If Val(.TextMatrix(intRow, mconIntColʵ������)) = 0 Then
                    MsgBox "��" & intRow & "����������0����������⣬���������ݣ�", vbInformation, gstrSysName
                    Exit Function
                End If
            End If
        Next
    End With
    
    '�����
    strҩƷ = CheckNumStock(mshBill, cboStock.ItemData(cboStock.ListIndex), 0, mconIntCol����, mconIntColʵ������, mconIntCol����ϵ��, 1, 1, mintNumberDigit)
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
    
    dat������� = Format(Sys.Currentdate, "yyyy-mm-dd hh:mm:ss")
    With mshBill
        
        
        '��ҩƷID˳���������
        recSort.Sort = "ҩƷid,����,���"
        recSort.MoveFirst
        
        For n = 1 To recSort.RecordCount
            intRow = recSort!�к�
            If .TextMatrix(intRow, 0) <> "" Then
                
                lngҩƷID = .TextMatrix(intRow, 0)
                str���� = .TextMatrix(intRow, mconIntCol����)
                lng���� = .TextMatrix(intRow, mconIntCol����)
                
                If Val(.TextMatrix(intRow, mconIntCol��д����)) = Val(.TextMatrix(intRow, mconIntColʵ������)) Then
                    num��д���� = Val(.TextMatrix(intRow, mconintCol��ʵ����))
                    numʵ������ = Val(.TextMatrix(intRow, mconintCol��ʵ����))
                Else
                    num��д���� = .TextMatrix(intRow, mconIntCol��д����) * .TextMatrix(intRow, mconIntCol����ϵ��)
                    numʵ������ = .TextMatrix(intRow, mconIntColʵ������) * .TextMatrix(intRow, mconIntCol����ϵ��)
                End If
                
                num�ɱ��� = zlStr.FormatEx(.TextMatrix(intRow, mconIntCol�ɹ���) / .TextMatrix(intRow, mconIntCol����ϵ��), gtype_UserDrugDigits.Digit_�ɱ���)
                num�ɱ��� = Get�ɱ���(lngҩƷID, lng�ⷿID, lng����)
                
                num�ɱ���� = .TextMatrix(intRow, mconIntCol�ɹ����)
                num���۽�� = .TextMatrix(intRow, mconIntCol�ۼ۽��)
                num��� = .TextMatrix(intRow, mconintCol���)
                str���� = .TextMatrix(intRow, mconIntCol����)
                datЧ�� = IIf(.TextMatrix(intRow, mconIntColЧ��) = "", "", .TextMatrix(intRow, mconIntColЧ��))
                If gtype_UserSysParms.P149_Ч����ʾ��ʽ = 1 And datЧ�� <> "" Then
                    '����ΪʧЧ��������
                    datЧ�� = Format(DateAdd("D", 1, datЧ��), "yyyy-mm-dd")
                End If
                
                int��� = Val(.TextMatrix(intRow, mconIntCol���))
                str��׼�ĺ� = IIf(Trim(.TextMatrix(intRow, mconIntCol��׼�ĺ�)) = "", "", .TextMatrix(intRow, mconIntCol��׼�ĺ�))
                
                gstrSQL = "zl_ҩƷ����_Verify("
                '���
                gstrSQL = gstrSQL & int���
                'NO
                gstrSQL = gstrSQL & ",'" & strNo & "'"
                '�ⷿID
                gstrSQL = gstrSQL & "," & lng�ⷿID
                '�Է�����ID
                gstrSQL = gstrSQL & "," & lng�Է�����id
                'ҩƷID
                gstrSQL = gstrSQL & "," & lngҩƷID
                '����
                gstrSQL = gstrSQL & ",'" & str���� & "'"
                '����
                gstrSQL = gstrSQL & "," & lng����
                '��д����
                gstrSQL = gstrSQL & "," & num��д����
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
                '�����
                gstrSQL = gstrSQL & ",'" & str����� & "'"
                '�������
                gstrSQL = gstrSQL & ",to_date('" & dat������� & "','yyyy-mm-dd HH24:MI:SS')"
                '����
                gstrSQL = gstrSQL & ",'" & str���� & "'"
                'Ч��
                gstrSQL = gstrSQL & "," & IIf(datЧ�� = "", "Null", "to_date('" & Format(datЧ��, "yyyy-mm-dd") & "','yyyy-mm-dd')")
                '��׼�ĺ�
                gstrSQL = gstrSQL & ",'" & str��׼�ĺ� & "'"
                '���÷�ʽ
                gstrSQL = gstrSQL & "," & mint���÷�ʽ
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
    Dim ����ԭ��_IN  As String
    
    SaveStrike = False
    arrSql = Array()
    With mshBill
        '����������������С����
        For intRow = 1 To .rows - 1
            If Val(.TextMatrix(intRow, mconIntColʵ������)) <> 0 Then
                If Not ��ͬ����(Val(.TextMatrix(intRow, mconIntCol��д����)), Val(.TextMatrix(intRow, mconIntColʵ������))) Then
                    MsgBox "������Ϸ��ĳ�����������" & intRow & "�У���", vbInformation, gstrSysName
                    Exit Function
                End If
            End If
        Next
        
        '�����
        strҩƷ = CheckNumStock(mshBill, cboStock.ItemData(cboStock.ListIndex), 0, mconIntCol����, mconIntColʵ������, mconIntCol����ϵ��, 2, 1, mintNumberDigit)
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
        
        ����ԭ��_IN = Trim(txt����ԭ��.Text)
        
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
                strҩƷID = IIf(strҩƷID = "", "", strҩƷID & ",") & ҩƷID_IN
                If Val(.TextMatrix(intRow, mconIntCol��д����)) = Val(.TextMatrix(intRow, mconIntColʵ������)) Then
                    ��������_IN = Val(.TextMatrix(intRow, mconintCol��ʵ����))
                Else
                    ��������_IN = zlStr.FormatEx(.TextMatrix(intRow, mconIntColʵ������) * .TextMatrix(intRow, mconIntCol����ϵ��), gtype_UserDrugDigits.Digit_����, , True)
                End If
                ���_IN = .TextMatrix(intRow, mconIntCol���)
                
                gstrSQL = "ZL_ҩƷ����_STRIKE("
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
                '������ʽ
                gstrSQL = gstrSQL & "," & mint����ʽ
                '����ԭ��
                gstrSQL = gstrSQL & ",'" & ����ԭ��_IN & "'"
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
    Dim strҩƷID As String
    Dim i As Integer
    Dim intRow As Integer
    Dim intOldRow  As Integer
    
    intOldRow = mshBill.Row
    mshBill.CmdEnable = False
    mblnChange = True
'    Set RecReturn = FrmҩƷѡ����.ShowME(Me, 2,cboStock.ItemData(cboStock.ListIndex), , Val(txtDraw.Tag), True, True, False, False, True, 0, False, mint���÷�ʽ)
    
    If grsMaster.State = adStateClosed Then
        Call SetSelectorRS(2, MStrCaption, cboStock.ItemData(cboStock.ListIndex), cboStock.ItemData(cboStock.ListIndex), Val(txtDraw.Tag))
    End If
    
    Set RecReturn = frmSelector.ShowME(Me, 0, 2, , , , cboStock.ItemData(cboStock.ListIndex), , Val(txtDraw.Tag), 0, True, True, True, 0, , mstrPrivs)
    If RecReturn.RecordCount > 0 Then
        Set RecReturn = CheckData(RecReturn)  '����ظ���¼ �����ظ���¼��ҩƷid���ػ���
    End If
    
    mshBill.CmdEnable = True
    
    If RecReturn.RecordCount > 0 Then
        RecReturn.MoveFirst
        With mshBill
            For i = 1 To RecReturn.RecordCount
                intRow = .Row
'            If RecReturn.RecordCount = 1 Then
                .TextMatrix(intRow, mconIntCol�к�) = .Row
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
                    Choose(mintUnit, 1, RecReturn!�����װ, RecReturn!סԺ��װ, RecReturn!ҩ���װ), IIf(IsNull(RecReturn!����), 0, RecReturn!����), RecReturn!ʱ��, RecReturn!ҩ������, _
                    IIf(IsNull(RecReturn!��׼�ĺ�), "", RecReturn!��׼�ĺ�), nvl(RecReturn!ԭ����)
                .Col = mconIntCol��д����
'            End If
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
        
        If .Col = mconIntCol��д���� Or .Col = mconIntColʵ������ Or .Col = mconIntCol�ɹ��� Or .Col = mconIntCol�ۼ� Or .Col = mconIntCol�ɹ���� Or .Col = mconIntCol�ۼ۽�� Then
            Select Case .Col
                Case mconIntCol��д����, mconIntColʵ������
                    intDigit = mintNumberDigit
                Case mconIntCol�ɹ���
                   intDigit = mintCostDigit
                Case mconIntCol�ۼ�
                    intDigit = mintPriceDigit
                Case mconIntCol�ɹ����, mconIntCol�ۼ۽��
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
            Case mconIntCol��д����, mconIntColʵ������
                .txtCheck = True
                .MaxLength = 16
                .TextMask = ".1234567890"
                Call ��ʾ�����
        End Select
    End With
End Sub

Private Sub mshbill_KeyDown(KeyCode As Integer, Shift As Integer, Cancel As Boolean)
    Dim strKey As String
    Dim rsDrug As New Recordset
    Dim strUnit As String
    Dim strUnitQuantity As String
    Dim rsTemp As ADODB.Recordset
    Dim i As Integer
    Dim intCurRow As Integer
    Dim intOldRow  As Integer
    
    intOldRow = mshBill.Row
    
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
                    
'                    Set RecReturn = FrmҩƷ��ѡѡ����.ShowME(Me, 2, cboStock.ItemData(cboStock.ListIndex), , Val(txtDraw.Tag), strkey, sngLeft, sngTop, True, True, False, False, True, 0, False, mint���÷�ʽ)
                    If grsMaster.State = adStateClosed Then
                        Call SetSelectorRS(2, MStrCaption, cboStock.ItemData(cboStock.ListIndex), cboStock.ItemData(cboStock.ListIndex), Val(txtDraw.Tag))
                    End If
                    
                    Set RecReturn = frmSelector.ShowME(Me, 1, 2, strKey, sngLeft, sngTop, cboStock.ItemData(cboStock.ListIndex), , Val(txtDraw.Tag), 0, True, True, True, 0, , mstrPrivs)
                    If RecReturn.RecordCount > 0 Then
                        Set RecReturn = CheckData(RecReturn) '���ظ��ļ�¼��ʱ���޿���ҩƷ���˵�
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
                                    IIf(IsNull(RecReturn!�ӳ���), "0", RecReturn!�ӳ��� / 100), _
                                    Choose(mintUnit, 1, RecReturn!�����װ, RecReturn!סԺ��װ, RecReturn!ҩ���װ), IIf(IsNull(RecReturn!����), 0, RecReturn!����), RecReturn!ʱ��, RecReturn!ҩ������, IIf(IsNull(RecReturn!��׼�ĺ�), "", RecReturn!��׼�ĺ�), nvl(RecReturn!ԭ����)) = False Then
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
                    If Val(strKey) = 0 And mint�༭״̬ <> 3 And mint�༭״̬ <> 6 Then '������������Ϊ0
                        MsgBox "�Բ�����������Ϊ��,�����䣡", vbInformation + vbOKOnly, gstrSysName
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
                    
                    If mint�༭״̬ = 6 Then
                        If Not ��ͬ����(Val(strKey), Val(.TextMatrix(.Row, mconIntCol��д����))) Then
                            MsgBox "�Բ��𣬳��������ķ���Ӧ����ԭ������һ�£�", vbInformation, gstrSysName
                            Cancel = True
                            .TxtSetFocus
                            Exit Sub
                        End If
                        
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
                    '10.35.40,�������γ���ʱ��������������򲻼��(������Զ��ֽ����ټ��)
                    ElseIf (mint�༭״̬ = 1 Or mint�༭״̬ = 2 Or mint�༭״̬ = 3) And mint�����γ��� = 1 Then
                        If Not CheckUsableNum(cboStock.ItemData(cboStock.ListIndex), Val(.TextMatrix(.Row, 0)), Val(.TextMatrix(.Row, mconIntCol����)), strKey, Val(mshBill.TextMatrix(.Row, mconIntCol����ϵ��)), txtNo.Caption, 7, mint�����, mintNumberDigit) Then
                            Cancel = True
                            .TxtSetFocus
                            Exit Sub
                        End If
                    End If
                    
                    If .TextMatrix(.Row, 0) = "" Then Exit Sub

                    strKey = zlStr.FormatEx(strKey, mintNumberDigit, , True)
                    .Text = strKey
                    
                    If .TextMatrix(.Row, mconIntCol�ۼ�) <> "" Then
                        .TextMatrix(.Row, mconIntCol�ۼ۽��) = zlStr.FormatEx(.TextMatrix(.Row, mconIntCol�ۼ�) * strKey, mintMoneyDigit, , True)
                    End If
                    
'                    .TextMatrix(.Row, mconintCol���) =Str.FormatEx(Get������(Val(cboStock.ItemData(cboStock.ListIndex)), Val(.TextMatrix(.Row, 0)), Val(.TextMatrix(.Row, mconIntCol����)), Val(.TextMatrix(.Row, mconIntColʵ�ʽ��)), Val(.TextMatrix(.Row, mconIntColʵ�ʲ��)), Val(.TextMatrix(.Row, mconIntCol�ۼ۽��)), Val(strkey) * Val(mshBill.TextMatrix(.Row, mconIntCol����ϵ��))), mintMoneyDigit)
                    
                    If strKey <> 0 And (mint�༭״̬ = 1 Or mint�༭״̬ = 2) Then
'                        .TextMatrix(.Row, mconIntCol�ɹ���) =Str.FormatEx((.TextMatrix(.Row, mconIntCol�ۼ۽��) - .TextMatrix(.Row, mconintCol���)) / strkey, mintCostDigit)
                        .TextMatrix(.Row, mconIntCol�ɹ���) = zlStr.FormatEx(Get�ɱ���(Val(.TextMatrix(.Row, 0)), Val(cboStock.ItemData(cboStock.ListIndex)), Val(.TextMatrix(.Row, mconIntCol����))) * Val(Val(mshBill.TextMatrix(.Row, mconIntCol����ϵ��))), mintCostDigit, , True)
                    End If
                    .TextMatrix(.Row, mconIntCol�ɹ����) = zlStr.FormatEx(Val(.TextMatrix(.Row, mconIntCol�ɹ���)) * strKey, mintMoneyDigit, , True)
                    .TextMatrix(.Row, mconintCol���) = zlStr.FormatEx(Val(.TextMatrix(.Row, mconIntCol�ۼ۽��)) - Val(.TextMatrix(.Row, mconIntCol�ɹ����)), mintMoneyDigit, , True)
                    
                    If .Col = mconIntCol��д���� Then
                        .TextMatrix(.Row, mconIntColʵ������) = strKey
                    End If
                End If
                ��ʾ�ϼƽ��
            
        End Select
    End With
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
    Dim strҩ�� As String
    
    SetColValue = False
    
    '����Ƿ��ظ�
'    If Not CheckRepeatMedicine(mshBill, lngҩƷID & "," & "0" & "|" & lng���� & "," & mconIntCol����, intRow) Then
'        Exit Function
'    End If
    
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
        
        If lng���� > 0 Then
            .TextMatrix(intRow, mconIntCol����) = lng����
        Else
            .TextMatrix(intRow, mconIntCol����) = 0
        End If
        
        .TextMatrix(intRow, mconIntCol����) = str����
        .TextMatrix(intRow, mconIntColЧ��) = Format(strЧ��, "yyyy-mm-dd")
        .TextMatrix(intRow, mconIntCol��׼�ĺ�) = str��׼�ĺ�
        
        .TextMatrix(intRow, mconIntCol�ۼ�) = zlStr.FormatEx(num�ۼ� * num����ϵ��, mintPriceDigit, , True)
        If int�Ƿ��� = 1 Then
            .TextMatrix(intRow, mconIntCol�ۼ�) = zlStr.FormatEx(Get���ۼ�(lngҩƷID, cboStock.ItemData(cboStock.ListIndex), lng����, num����ϵ��), mintPriceDigit, , True)
        End If
        .TextMatrix(intRow, mconIntCol��������) = zlStr.FormatEx(num��������, mintNumberDigit)
        .TextMatrix(intRow, mconIntColʵ�ʲ��) = numʵ�ʲ��
        .TextMatrix(intRow, mconIntColʵ�ʽ��) = numʵ�ʽ��
        .TextMatrix(intRow, mconIntcol�ӳ���) = dbl�ӳ��� & "||" & int�Ƿ��� & "||" & intҩ������
        .TextMatrix(intRow, mconIntCol����ϵ��) = num����ϵ��
        
'        If int�Ƿ��� = 1 Then
'            dblPrice = GetPrice(lngҩƷid, lng����, num����ϵ��)
'            .TextMatrix(intRow, mconIntCol�ۼ�) = zlStr.FormatEx(dblPrice, mintPriceDigit)
'        End If
        If IsLowerLimit(cboStock.ItemData(cboStock.ListIndex), lngҩƷID) Then Call SetForeColor_ROW(mlng��ɫ)
        Call CheckLapse(strЧ��)
        
    End With
    SetColValue = True
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

Private Sub mshProvider_DblClick()
    mshProvider_KeyDown vbKeyReturn, 0
End Sub

Private Sub mshProvider_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        mshProvider.Visible = False
        If mshProvider.Tag = 0 Then
            txtDraw.SetFocus
            txtDraw.SelStart = 0
            txtDraw.SelLength = Len(txtDraw.Text)
        Else
            cboDrawPerson.SetFocus
            cboDrawPerson.SelStart = 0
            cboDrawPerson.SelLength = Len(cboDrawPerson.Text)
        End If
    End If
    
    If KeyCode = vbKeyReturn Then
        If mshProvider.Tag = 0 Then
            txtDraw.Text = mshProvider.TextMatrix(mshProvider.Row, 3)
            txtDraw.Tag = mshProvider.TextMatrix(mshProvider.Row, 0)
            mshProvider.Visible = False
            Call GetDrawPerson(txtDraw.Tag)
            cboDrawPerson.SetFocus
        Else
            cboDrawPerson.Text = mshProvider.TextMatrix(mshProvider.Row, 1)
            mshBill.SetFocus
            mshBill.Col = 1
            mshBill.Row = 1
        End If
    End If
    
End Sub

Private Sub mshProvider_LostFocus()
    If mshProvider.Visible Then
        mshProvider.Visible = False
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
            If Val(txtDraw.Tag) = 0 Then
                If Trim(txtDraw.Text) = "" Then
                    MsgBox "�Բ�����ҩ���Ų���Ϊ�գ�", vbOKOnly + vbInformation, gstrSysName
                    txtDraw.SetFocus
                    Exit Function
                Else
                    MsgBox "�Բ���û�����������ҩ���ţ�", vbOKOnly + vbInformation, gstrSysName
                    txtDraw.SetFocus
                    Exit Function
                End If
            End If
            
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
                    
                    If Trim(Trim(.TextMatrix(intLop, mconIntColʵ������))) = "" Then
                        MsgBox "��" & intLop & "��ҩƷ������Ϊ���ˣ����飡", vbInformation, gstrSysName
                        mshBill.SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        .Col = mconIntColʵ������
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
                                        Val(mshBill.TextMatrix(intLop, mconIntCol��д����)), Val(.TextMatrix(intLop, mconIntCol����ϵ��)), _
                                        Trim(txtNo.Caption), 7, mint�����, mintNumberDigit) Then
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

Private Function SaveCard(Optional ByVal blnǿ�Ʊ��� As Boolean = False) As Boolean
    Dim lng������id As Long
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
    Dim dblActual As Double
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
    Dim rs������ As New Recordset
    Dim arrSql As Variant
    Dim intRow As Integer
    Dim str��׼�ĺ� As String
    Dim n As Integer
    Dim blnTran As Boolean
    
    SaveCard = False
    arrSql = Array()
    
    '����������������ID����Ҫ������ҩƷ��Ҫ����
    gstrSQL = "SELECT b.id " _
             & "FROM ҩƷ�������� a, ҩƷ������ b " _
            & "Where a.���id = b.ID " _
              & "AND a.���� = 7 " _
              & "AND b.ϵ�� = -1 " _
              & "AND ROWNUM < 2"
    Call zlDataBase.OpenRecordset(rs������, gstrSQL, "ȡ������")
    If rs������.EOF Then
        MsgBox "�Բ���û������ҩƷ���õĳ����������ҩƷ������࣡", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    lng������id = rs������.Fields(0)
    rs������.Close
    
    With mshBill
        chrNo = Trim(txtNo)
        If chrNo = "" Then chrNo = Sys.GetNextNo(27, Me.cboStock.ItemData(Me.cboStock.ListIndex))
        If IsNull(chrNo) Then Exit Function
        Me.txtNo.Tag = chrNo
        lngStockid = cboStock.ItemData(cboStock.ListIndex)
        lngEnterStockID = txtDraw.Tag
        strBrief = Trim(txtժҪ.Text)
        strBooker = Txt������
        datBookDate = Format(Sys.Currentdate, "yyyy-mm-dd hh:mm:ss")
        
        strAssessor = Txt�����
        On Error GoTo errHandle
        
        If blnǿ�Ʊ��� Then blnTran = True
        
        If mint�༭״̬ = 2 Or blnǿ�Ʊ��� Then        '�޸�
            gstrSQL = "zl_ҩƷ����_Delete('" & mstr���ݺ� & "')"
            ReDim Preserve arrSql(UBound(arrSql) + 1)
            arrSql(UBound(arrSql)) = "0;" & gstrSQL
            
            strBooker = Txt������
            datBookDate = Format(Txt��������, "yyyy-mm-dd hh:mm:ss")
            strModifier = UserInfo.�û����� '�޸���
            datModifyDate = Format(Sys.Currentdate, "yyyy-mm-dd hh:mm:ss") '�޸�����
        End If
            
        '��ҩƷID˳���������
        recSort.Sort = "ҩƷid,����,���"
        recSort.MoveFirst
        
        For n = 1 To recSort.RecordCount
            intRow = recSort!�к�
            If .TextMatrix(intRow, 0) <> "" Then
                If Val(.TextMatrix(intRow, mconIntColʵ������)) = 0 Then
                    MsgBox "������������Ϊ�㣬����������������", vbInformation, gstrSysName
                    Exit Function
                End If
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

                If Val(.TextMatrix(intRow, mconintColԭʼ����)) = Val(.TextMatrix(intRow, mconIntCol��д����)) Then
                    dblQuantity = Val(.TextMatrix(intRow, mconintCol��ʵ����))
                Else
                    dblQuantity = zlStr.FormatEx(.TextMatrix(intRow, mconIntCol��д����) * .TextMatrix(intRow, mconIntCol����ϵ��), gtype_UserDrugDigits.Digit_����)
                End If
                
                If Val(.TextMatrix(intRow, mconIntCol��д����)) = Val(.TextMatrix(intRow, mconIntColʵ������)) Then
                    dblActual = dblQuantity
                Else
                    dblActual = zlStr.FormatEx(.TextMatrix(intRow, mconIntColʵ������) * .TextMatrix(intRow, mconIntCol����ϵ��), gtype_UserDrugDigits.Digit_����)
                End If
                
'                dblPurchasePrice = zlStr.FormatEx(Val(.TextMatrix(intRow, mconIntCol�ɹ���)) / .TextMatrix(intRow, mconIntCol����ϵ��), gtype_UserDrugDigits.Digit_�ɱ���)
                dblPurchasePrice = Get�ɱ���(lngDrugID, lngStockid, lngBatchID)
                
                dblPurchaseMoney = Val(zlStr.FormatEx(Val(FormatEx(dblPurchasePrice * Val(.TextMatrix(intRow, mconIntCol����ϵ��)), mintCostDigit)) * Val(.TextMatrix(intRow, mconIntColʵ������)), mintMoneyDigit, , True)) '.TextMatrix(intRow, mconIntCol�ɹ����)
                
'                dblSalePrice = zlStr.FormatEx(.TextMatrix(intRow, mconIntCol�ۼ�) / .TextMatrix(intRow, mconIntCol����ϵ��), gtype_UserDrugDigits.Digit_���ۼ�)
                dblSalePrice = Get�ۼ�(Split(.TextMatrix(intRow, mconIntcol�ӳ���), "||")(1) = 1, lngDrugID, lngStockid, lngBatchID)
                
                dblSaleMoney = Val(zlStr.FormatEx(Val(FormatEx(dblSalePrice * Val(.TextMatrix(intRow, mconIntCol����ϵ��)), mintPriceDigit)) * Val(.TextMatrix(intRow, mconIntColʵ������)), mintMoneyDigit, , True)) '.TextMatrix(intRow, mconIntCol�ۼ۽��)
                dblMistakePrice = Val(zlStr.FormatEx(dblSaleMoney - dblPurchaseMoney, mintMoneyDigit, , True)) '.TextMatrix(intRow, mconintCol���)
                
'                If Val(.TextMatrix(intRow, mconIntCol���)) = 0 Then
'                    lngSerial = intRow
'                Else
'                    lngSerial = .TextMatrix(intRow, mconIntCol���)
'                End If
                If mint�༭״̬ = 3 Then
                    lngSerial = .TextMatrix(intRow, mconIntCol���)
                Else
                    lngSerial = intRow
                End If
                
                str��׼�ĺ� = IIf(Trim(.TextMatrix(intRow, mconIntCol��׼�ĺ�)) = "", "", .TextMatrix(intRow, mconIntCol��׼�ĺ�))
                
                gstrSQL = "zl_ҩƷ����_INSERT("
                '������ID
                gstrSQL = gstrSQL & lng������id
                'NO
                gstrSQL = gstrSQL & ",'" & chrNo & "'"
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
                gstrSQL = gstrSQL & "," & dblActual
                '�ɱ���
                gstrSQL = gstrSQL & "," & dblPurchasePrice
                '�ɱ����
                gstrSQL = gstrSQL & "," & dblPurchaseMoney
                '�ۼ�
                gstrSQL = gstrSQL & "," & dblSalePrice
                '�ۼ۽��
                gstrSQL = gstrSQL & "," & dblSaleMoney
                '���
                gstrSQL = gstrSQL & "," & dblMistakePrice
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
                '������
                gstrSQL = gstrSQL & ",'" & cboDrawPerson.Text & "'"
                '��׼�ĺ�
                gstrSQL = gstrSQL & ",'" & str��׼�ĺ� & "'"
                '���÷�ʽ
                gstrSQL = gstrSQL & "," & mint���÷�ʽ
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
            " Where a.���� = 7 And a.No = [1] And a.ҩƷid = b.�շ�ϸĿid And c.Id = b.�շ�ϸĿid And Round(a.���ۼ�," & intPriceDigit & ") <> Round(b.�ּ�, " & intPriceDigit & ") And" & _
              "    NVL(c.�Ƿ���, 0) = 0 " & _
            " Union All" & _
            " Select '�ۼ�' As ����, a.���, a.ҩƷid ,nvl(a.����,0) as ����, 0 ԭ��, decode(x.�ּ�,null,decode(nvl(b.���ۼ�,0),0,b.ʵ�ʽ�� / b.ʵ������,b.���ۼ�),x.�ּ�) As �ּ�" & _
            " From ҩƷ�շ���¼ A, ҩƷ��� B, �շ���ĿĿ¼ C , " & _
            "      (Select x.ҩƷid,x.�ⷿid,x.����,x.�ּ� from ҩƷ�۸��¼ x where x.�۸����� = 1 and (x.��ֹ���� Is Null Or Sysdate Between x.ִ������ And Nvl(x.��ֹ����, To_Date('3000-01-01', 'yyyy-MM-dd')))) X" & _
            " Where a.���� = 7 And a.No = [1] And c.Id = a.ҩƷid And Round(a.���ۼ�," & intPriceDigit & ") <> Round(decode(x.�ּ�,null,decode(nvl(b.���ۼ�,0),0,b.ʵ�ʽ�� / b.ʵ������,b.���ۼ�),x.�ּ�), " & intPriceDigit & ") And Nvl(c.�Ƿ���, 0) = 1 And" & _
                  " b.���� = 1 And b.�ⷿid = a.�ⷿid And b.ҩƷid = a.ҩƷid And NVL(b.����, 0) = NVL(a.����, 0) And NVL(b.ʵ������, 0) <> 0 And a.���ϵ�� = -1" & _
                  " AND a.ҩƷid = x.ҩƷid(+) And a.�ⷿid = x.�ⷿid(+) And Nvl(a.����, 0) = Nvl(x.����(+), 0) " & _
            " Union All" & _
            " Select '�ɱ���' As ����, a.���, a.ҩƷid ,nvl(a.����,0) as ����,  0 ԭ��,decode(x.�ּ�,null,b.ƽ���ɱ���,x.�ּ�) As �ּ�" & _
            " From ҩƷ�շ���¼ A, ҩƷ��� B , " & _
            "      (Select x.ҩƷid,x.�ⷿid,x.����,x.�ּ� from ҩƷ�۸��¼ x where x.�۸����� = 2 and (x.��ֹ���� Is Null Or Sysdate Between x.ִ������ And Nvl(x.��ֹ����, To_Date('3000-01-01', 'yyyy-MM-dd')))) X" & _
            " Where a.���� = 7 And a.No = [1] And a.ҩƷid = b.ҩƷid And Nvl(a.����, 0) = Nvl(b.����, 0) and round(a.�ɱ���," & intCostDigit & ")<>round(decode(x.�ּ�,null,b.ƽ���ɱ���,x.�ּ�)," & intCostDigit & ") And a.�ⷿid = b.�ⷿid and a.���ϵ��=-1  and b.����=1" & _
            " AND a.ҩƷid = x.ҩƷid(+) And a.�ⷿid = x.�ⷿid(+) And Nvl(a.����, 0) = Nvl(x.����(+), 0) " & _
            " Order By ����, ҩƷid, ���"
    Set rsPrice = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "[ȡ��ǰ�۸�]", CStr(Me.txtNo.Tag))
    
    If rsPrice.EOF Then Exit Sub
    
    lngRows = mshBill.rows - 1
    For lngRow = 1 To lngRows
        blnAdj = False
        lngҩƷID = Val(mshBill.TextMatrix(lngRow, 0))
        Dbl���� = Val(mshBill.TextMatrix(lngRow, mconIntColʵ������))
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
    Dim lng���� As Long
    Dim dbl�������� As Double, dblʵ������ As Double
    
    On Error GoTo errHandle
    With mshBill
        If .TextMatrix(.Row, mconIntColҩ��) = "" Then
            staThis.Panels(2).Text = ""
            Exit Sub
        End If
        If .TextMatrix(mshBill.Row, 0) = "" Then Exit Sub
        lng���� = Val(.TextMatrix(.Row, mconIntCol����))
        
        If mint���÷�ʽ = 0 Then
            If lng���� > 0 Then
                gstrSQL = "select ��������/" & .TextMatrix(.Row, mconIntCol����ϵ��) & " as  ��������,ʵ������/" & .TextMatrix(.Row, mconIntCol����ϵ��) & " as  ʵ������ " & _
                    " from ҩƷ��� where �ⷿid=[1] " _
                    & " and ҩƷid=[2] " _
                    & " and ����=1 and " _
                    & " nvl(����,0)=[3]"
            Else
                gstrSQL = " Select Sum(��������)/" & .TextMatrix(.Row, mconIntCol����ϵ��) & " as ��������,Sum(ʵ������)/" & .TextMatrix(.Row, mconIntCol����ϵ��) & " as ʵ������ " & _
                    " from ҩƷ��� " & _
                              " Where �ⷿid=[1] " & _
                              " And ҩƷid=[2] And ����=1 "
            End If
            Set rsUseCount = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "[��ʾ�����]", cboStock.ItemData(cboStock.ListIndex), Val(.TextMatrix(.Row, 0)), Val(.TextMatrix(.Row, mconIntCol����)))
        Else
            gstrSQL = "select ��������/" & .TextMatrix(.Row, mconIntCol����ϵ��) & " as  ��������,ʵ������/" & .TextMatrix(.Row, mconIntCol����ϵ��) & " as  ʵ������ from ҩƷ���� where �ڼ�=[1] and �ⷿid=[2] " _
                & " and ҩƷid=[3] And ����ID=[4] "
            Set rsUseCount = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "[��ʾ�����]", Format(Sys.Currentdate(), IIf(mint���淽ʽ = 0, "yyyy", "yyyymm")), cboStock.ItemData(cboStock.ListIndex), Val(.TextMatrix(.Row, 0)), Val(txtDraw.Tag))
        End If
        
        If rsUseCount.EOF Then
            .TextMatrix(.Row, mconIntCol��������) = 0
            dbl�������� = 0
            dblʵ������ = 0
        Else
            .TextMatrix(.Row, mconIntCol��������) = IIf(IsNull(rsUseCount.Fields(0)), 0, rsUseCount.Fields(0))
            dbl�������� = IIf(IsNull(rsUseCount.Fields(0)), 0, rsUseCount.Fields(0))
            dblʵ������ = IIf(IsNull(rsUseCount.Fields(1)), 0, rsUseCount.Fields(1))
        End If
                
        If mint�༭״̬ = 3 Or mint�༭״̬ = 6 Then
            '��ˣ�����ʱ��ʾʵ������
            staThis.Panels(2).Text = "��ҩƷ��ǰ���ʵ������Ϊ[" & FormatEx(dblʵ������, mintNumberDigit) & "]" & .TextMatrix(.Row, mconIntCol��λ)
          Else
            '����������޸�ʱ��ʾ��������
            staThis.Panels(2).Text = "��ҩƷ��ǰ����������Ϊ[" & FormatEx(dbl��������, mintNumberDigit) & "]" & .TextMatrix(.Row, mconIntCol��λ)
        End If
        
        
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub txtDraw_LostFocus()
    If txtDraw.Text = "" Then
        txtDraw.Tag = "0"
        Exit Sub
    End If
End Sub

Private Sub txtDraw_Validate(Cancel As Boolean)
    If txtDraw.Text = "" Then
        txtDraw.Tag = "0"
        Exit Sub
    End If
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
    intNO = 28
    lng�ⷿID = Me.cboStock.ItemData(Me.cboStock.ListIndex)
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    If Trim(txtIn.Text) = "" Then Exit Sub
    
    If Len(txtIn.Text) < 8 Then
        txtIn.Text = zlCommFun.GetFullNO(txtIn.Text, intNO, lng�ⷿID)
    End If
    
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
    
    gstrSQL = "select �շ�ϸĿid,ִ�п���id from �շ�ִ�п���"
    Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, "��ѯ�洢�ⷿ")
    
    '��ȡ�õ��ݲ���ձ��ֻ������ȡ�������ݣ��ҷ��˻�����
    gstrSQL = "SELECT A.ҩƷID,'['||C.����||']' As ����,'['||C.����||']'|| Nvl(F.����,C.����) As ҩƷ����, C.���� As ͨ����,F.���� As ��Ʒ��,C.���,a.����,a.ԭ����," & _
             "        C.���㵥λ AS ���۵�λ,1 AS ����ϵ��,B.���ﵥλ,B.�����װ,B.סԺ��λ,B.סԺ��װ,B.ҩ�ⵥλ,B.ҩ���װ, " & _
             "        NVL(A.����,0) AS ����,Nvl(C.�Ƿ���,0) AS ʱ��,Nvl(B.ҩ������,0) AS ҩ������,Nvl(B.ҩ�����,0) AS ҩ�����,b.���Ч��,A.����,A.Ч��," & _
             "        B.����ѱ���,B.�ӳ���,A.ʵ������,D.��������,D.ʵ�ʽ��,D.ʵ�ʲ��,E.�ּ�,A.��׼�ĺ�,B.ҩƷ��Դ,B.����ҩ��,nvl(d.ƽ���ɱ���,0) as ƽ���ɱ���,a.��ҩ��λid " & _
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
            lngҩƷID = !ҩƷID
            rsTemp.Filter = " �շ�ϸĿid=" & lngҩƷID & " and ִ�п���id=" & lng�ⷿID
            If rsTemp.RecordCount = 0 Then
                MsgBox "ҩƷ[" & !ҩƷ���� & "]δ��" & cboStock.Text & "�����ô洢���ԣ����������ã�"
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
                        MsgBox !ҩƷ���� & "��治�㣬���������ã�", vbInformation, gstrSysName
                        blnInput = True
                    End Select
                End If
            End If
            
            'װ������(SetColValue)
            If blnInput = False Then
                int��װϵ�� = Choose(mintUnit, 1, !�����װ, !סԺ��װ, !ҩ���װ)
                If Not SetColValue(intRow, !ҩƷID, !����, !ͨ����, IIf(IsNull(!��Ʒ��), "", !��Ʒ��), _
                    nvl(!ҩƷ��Դ), nvl(!����ҩ��), nvl(!���), nvl(!����), _
                    Choose(mintUnit, !���۵�λ, !���ﵥλ, !סԺ��λ, !ҩ�ⵥλ), nvl(!�ּ�, 0), _
                    nvl(!����), nvl(!Ч��), nvl(!��������, 0), nvl(!ʵ�ʽ��, 0), nvl(!ʵ�ʲ��, 0), _
                    nvl(!�ӳ��� / 100, 0), int��װϵ��, nvl(!����, 0), !ʱ��, _
                    !ҩ������, IIf(IsNull(!��׼�ĺ�), "", !��׼�ĺ�), nvl(!ԭ����)) Then
                    mshBill.ClearBill
                    Exit Sub
                End If
    
                '��д�������ɹ��ۡ��ۼ۵���
                mshBill.TextMatrix(intRow, mconIntCol�к�) = intRow
                mshBill.TextMatrix(intRow, mconIntColʵ������) = zlStr.FormatEx(!ʵ������ / int��װϵ��, mintNumberDigit, , True)
                mshBill.TextMatrix(intRow, mconIntCol��д����) = zlStr.FormatEx(!ʵ������ / int��װϵ��, mintNumberDigit, , True)
                mshBill.TextMatrix(intRow, mconIntCol�ɹ���) = zlStr.FormatEx(!ƽ���ɱ��� * int��װϵ��, mintCostDigit, , True)
                mshBill.TextMatrix(intRow, mconIntCol�ɹ����) = zlStr.FormatEx(Val(mshBill.TextMatrix(intRow, mconIntCol�ɹ���)) * Val(mshBill.TextMatrix(intRow, mconIntColʵ������)), mintMoneyDigit, , True)
                mshBill.TextMatrix(intRow, mconIntCol�ۼ۽��) = zlStr.FormatEx(Val(mshBill.TextMatrix(intRow, mconIntCol�ۼ�)) * Val(mshBill.TextMatrix(intRow, mconIntColʵ������)), mintMoneyDigit, , True)
                mshBill.TextMatrix(intRow, mconintCol���) = zlStr.FormatEx(Val(mshBill.TextMatrix(intRow, mconIntCol�ۼ۽��)) - mshBill.TextMatrix(intRow, mconIntCol�ɹ����), mintMoneyDigit, , True)
    
                intRow = intRow + 1
                mshBill.rows = mshBill.rows + 1
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

Private Sub txt����ԭ��_GotFocus()
    OS.OpenIme True
    With txt����ԭ��
        .SelStart = 0
        .SelLength = Len(txt����ԭ��.Text)
    End With
End Sub

Private Sub txt����ԭ��_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        OS.PressKey (vbKeyTab)
        KeyCode = 0
    End If
End Sub

Private Sub txt����ԭ��_LostFocus()
    OS.OpenIme
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
    FrmBillPrint.ShowME Me, glngSys, IIf(glngSys \ 100 = 1, "zl1_bill_1305", "zl8_bill_1305"), mint��¼״̬, int��λϵ��, 1305, "ҩƷ���õ�", strNo
End Sub


Private Sub txtDraw_Change()
    With txtDraw
        .Text = UCase(.Text)
        .SelStart = Len(.Text)
    End With
    mblnChange = True
End Sub

Private Sub txtDraw_GotFocus()
    txtDraw.SelStart = 0
    txtDraw.SelLength = Len(txtDraw.Text)
End Sub

Private Sub txtDraw_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim strProviderText As String, strվ������ As String
    Dim adoProvider As New Recordset
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    If mint�༭״̬ = 3 Or mint�༭״̬ = 4 Then Exit Sub
    
    On Error GoTo errHandle
    With txtDraw
        If Trim(.Text) = "" Then Exit Sub
        strProviderText = UCase(.Text)
        strվ������ = GetDeptStationNode(cboStock.ItemData(cboStock.ListIndex))
        
        If mblnStock Then
            gstrSQL = "SELECT DISTINCT a.id,a.����,a.����,a.���� " _
                    & "FROM ��������˵�� c, �������ʷ��� b, ���ű� a " _
                    & "Where " & IIf(strվ������ <> "", "(a.վ�� = [4] or a.վ�� is null) And ", "") & "c.�������� = b.���� " _
                    & "  AND b.���� = 'O' AND a.id = c.����id " _
                    & "  AND TO_CHAR (a.����ʱ��, 'yyyy-MM-dd') = '3000-01-01'" _
                    & "  And (a.���� like [1] Or a.���� like [1] or a.���� like [1]) " _
                    & "Order By a.����"
        Else
            gstrSQL = " Select C.ID " & _
                " From ��������˵�� A,�������ʷ��� B,���ű� C " & _
                " Where " & IIf(strվ������ <> "", "(C.վ�� = [4] or C.վ�� is null) And ", "") & "A.��������=B.���� And A.����ID=C.ID " & _
                "   AND TO_CHAR(C.����ʱ��, 'yyyy-MM-dd')='3000-01-01' And B.����='O'" & _
                "   And C.ID IN (Select ����ID From ������Ա Where ��ԱID=[2])"
                
            'ֻ��ȡ��������������Ŀ���
            gstrSQL = "SELECT DISTINCT a.id,a.����,a.����,a.���� " _
                 & " FROM ��������˵�� c, �������ʷ��� b, ���ű� a " _
                 & " Where " & IIf(strվ������ <> "", "(a.վ�� = [4] or a.վ�� is null) And ", "") & " c.�������� = b.���� " _
                 & "   AND b.���� ='O' AND a.id = c.����id " _
                 & "   AND TO_CHAR (a.����ʱ��, 'yyyy-MM-dd') = '3000-01-01'" _
                 & "   And (a.���� like [1] Or a.���� like [1] or a.���� like [1])" _
                 & "   And a.ID IN (Select Distinct ���ò���ID From ҩƷ���ÿ��� Where �Է��ⷿid=[3] And ���ò���ID IN (" & gstrSQL & "))" _
                 & " Order By a.���� "
        End If
            
        Set adoProvider = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption, _
            IIf(gstrMatchMethod = "0", "%", "") & strProviderText & "%", _
            UserInfo.�û�ID, _
            cboStock.ItemData(cboStock.ListIndex), _
            strվ������)
        
        mshProvider.Tag = 0
        
        If adoProvider.EOF Then
            MsgBox "û�����������ҩ���ţ������䣡", vbOKOnly + vbInformation, gstrSysName
            KeyCode = 0
            .SelStart = 0
            .SelLength = Len(.Text)
            .Tag = 0
            Exit Sub
        End If
        If adoProvider.RecordCount > 1 Then
            Set mshProvider.Recordset = adoProvider
            Dim intCol As Integer
            Dim intRow As Integer
            
            With mshProvider
                If .Visible = False Then .Visible = True
                .Redraw = False
                .SetFocus
                
                For intRow = 0 To .rows - 1
                    .Row = intRow
                    For intCol = 0 To .Cols - 1
                        .Col = intCol
                        If .Row = 0 Then
                            .CellFontBold = True
                        Else
                            .CellFontBold = False
                        End If
                    Next
                Next
                .Font.Bold = False
                .FontFixed.Bold = True
                .ColWidth(0) = 0
                .ColWidth(1) = 800
                .ColWidth(2) = 1000
                .ColWidth(3) = 2500
                .Row = 1
                .TopRow = 1
                .Col = 0
                .ColSel = .Cols - 1
                
                .Width = LblEnterStock.Width + txtDraw.Width + cmdDraw.Width + 80
                .Top = txtDraw.Top + txtDraw.Height
                .Left = cmdDraw.Left + cmdDraw.Width - .Width
                .Redraw = True
                Exit Sub
            End With
        Else
            .Text = adoProvider!����
            .Tag = adoProvider!id
        End If
        adoProvider.Close
        Call GetDrawPerson(.Tag)
        cboDrawPerson.SetFocus
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
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
    
    rsTemp.MoveFirst
    
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
        
        '����Ƿ�ͬʱ��������Ϊ0�����β�Ϊ0������
        rsTemp.MoveFirst
        Do While Not rsTemp.EOF
            For i = 1 To .rows - 2
                '���صļ�¼���ķ������Ժͽ������еķ������Բ�һ��ʱ�������������ȡ���ݵ�����
                If rsTemp!ҩƷID = Val(.TextMatrix(i, 0)) And _
                    ((nvl(rsTemp!����, 0) = 0 And Val(.TextMatrix(i, mconIntCol����)) > 0) Or _
                    (nvl(rsTemp!����, 0) > 0 And Val(.TextMatrix(i, mconIntCol����)) = 0)) Then
                    
                    '���뵽��Ҫ�ų����嵥��
                    If InStr(1, strInfo�������, rsTemp!ҩƷID & "," & nvl(rsTemp!����, 0)) = 0 Then
                         strInfo������� = strInfo������� & .TextMatrix(i, 0) & "," & .TextMatrix(i, mconIntColҩ��) & "|"
                    End If
                    
                    '���뵽�������ѵ��嵥��
                    If InStr(1, "," & str�������Լ�� & ",", "," & .TextMatrix(i, mconIntColҩ��) & ",") = 0 Then
                        str�������Լ�� = IIf(str�������Լ�� = "", "", str�������Լ�� & ",") & .TextMatrix(i, mconIntColҩ��)
                    End If
                End If
            Next
            rsTemp.MoveNext
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
            rsTemp.Filter = strDub
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
        GetPrice = 0
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
    Dim bln���� As Boolean
    Dim lngStockid As Long
    
    On Error GoTo errHandle
    
    ���۸� = False
    
    lngStockid = cboStock.ItemData(cboStock.ListIndex)
    
    With mshBill
        For i = 1 To .rows - 1
            If mshBill.TextMatrix(i, 0) <> "" And Trim(.TextMatrix(i, mconIntCol��д����)) <> "" Then
                bln���� = Get��������(lngStockid, Val(.TextMatrix(i, 0))) '�Ƿ����
                bln�Ƿ�ʱ�� = Val(Split(.TextMatrix(i, mconIntcol�ӳ���), "||")(1)) = 1
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

