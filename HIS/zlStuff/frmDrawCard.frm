VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{876E3FF4-6E21-11D5-AF7D-0080C8EC27A9}#1.5#0"; "ZL9BillEdit.ocx"
Begin VB.Form frmDrawCard 
   Caption         =   "�������õ�"
   ClientHeight    =   6975
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   11400
   Icon            =   "frmDrawCard.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6975
   ScaleWidth      =   11400
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton cmdExpend 
      Caption         =   "�Զ��ֽ�(&A)"
      Height          =   350
      Left            =   1680
      TabIndex        =   44
      Top             =   6000
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.CommandButton cmdRequestDraw 
      Caption         =   "���깺������(&R)"
      Height          =   350
      Left            =   1800
      TabIndex        =   41
      Top             =   5520
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton cmdSel 
      Caption         =   "����(&G)"
      Height          =   360
      Left            =   270
      TabIndex        =   28
      Top             =   5940
      Width           =   810
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshProvider 
      Height          =   1812
      Left            =   1092
      TabIndex        =   27
      Top             =   6840
      Visible         =   0   'False
      Width           =   4092
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
   Begin VB.CommandButton cmdAllCls 
      Caption         =   "ȫ��(&L)"
      Height          =   350
      Left            =   9720
      TabIndex        =   25
      Top             =   5940
      Visible         =   0   'False
      Width           =   1100
   End
   Begin VB.CommandButton cmdAllSel 
      Caption         =   "ȫ��(&A)"
      Height          =   350
      Left            =   8400
      TabIndex        =   24
      Top             =   5940
      Visible         =   0   'False
      Width           =   1100
   End
   Begin VB.TextBox txtCode 
      Height          =   300
      Left            =   5520
      TabIndex        =   15
      Top             =   5610
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "����(&F)"
      Height          =   350
      Left            =   3840
      TabIndex        =   14
      Top             =   5520
      Width           =   1100
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      Height          =   350
      Left            =   240
      TabIndex        =   13
      Top             =   5520
      Width           =   1100
   End
   Begin VB.CommandButton CmdSave 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   8400
      TabIndex        =   11
      Top             =   5520
      Width           =   1100
   End
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   9720
      TabIndex        =   12
      Top             =   5520
      Width           =   1100
   End
   Begin VB.PictureBox Pic���� 
      BackColor       =   &H80000004&
      Height          =   5205
      Left            =   0
      ScaleHeight     =   5145
      ScaleWidth      =   11655
      TabIndex        =   16
      Top             =   0
      Width           =   11715
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
         Left            =   960
         MaxLength       =   8
         TabIndex        =   43
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
         Left            =   480
         Style           =   1  'Graphical
         TabIndex        =   42
         TabStop         =   0   'False
         ToolTipText     =   "������ʵ�:F3"
         Top             =   150
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox txtDrawPerson 
         Height          =   300
         Left            =   9660
         TabIndex        =   6
         Top             =   585
         Width           =   1425
      End
      Begin VB.CommandButton cmdDrawPerson 
         Caption         =   "��"
         Height          =   300
         Left            =   11100
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   555
         Width           =   300
      End
      Begin VB.TextBox txtNO 
         Height          =   300
         IMEMode         =   2  'OFF
         Left            =   9945
         Locked          =   -1  'True
         MaxLength       =   8
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   165
         Width           =   1425
      End
      Begin VB.TextBox txtDraw 
         Height          =   300
         Left            =   5715
         TabIndex        =   3
         Top             =   600
         Width           =   2415
      End
      Begin VB.CommandButton cmdDraw 
         Caption         =   "��"
         Height          =   300
         Left            =   8100
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   600
         Width           =   300
      End
      Begin ZL9BillEdit.BillEdit mshBill 
         Height          =   2790
         Left            =   195
         TabIndex        =   8
         Top             =   945
         Width           =   11235
         _ExtentX        =   19817
         _ExtentY        =   4921
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
         Top             =   4080
         Width           =   10410
      End
      Begin VB.ComboBox cboStock 
         Height          =   300
         Left            =   945
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   600
         Width           =   1905
      End
      Begin VB.Label Lbl������� 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�������"
         Height          =   180
         Left            =   8640
         TabIndex        =   40
         Top             =   4860
         Width           =   720
      End
      Begin VB.Label lbl����� 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "  �����"
         Height          =   180
         Left            =   8640
         TabIndex        =   39
         Top             =   4500
         Width           =   720
      End
      Begin VB.Label Lbl�������� 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��������"
         Height          =   180
         Left            =   120
         TabIndex        =   38
         Top             =   4860
         Width           =   720
      End
      Begin VB.Label Lbl������ 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "  ������"
         Height          =   180
         Left            =   120
         TabIndex        =   37
         Top             =   4500
         Width           =   720
      End
      Begin VB.Label Txt������ 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   900
         TabIndex        =   36
         Top             =   4440
         Width           =   1890
      End
      Begin VB.Label Txt�������� 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   900
         TabIndex        =   35
         Top             =   4800
         Width           =   1890
      End
      Begin VB.Label Txt������� 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   9480
         TabIndex        =   34
         Top             =   4800
         Width           =   1890
      End
      Begin VB.Label Txt����� 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   9480
         TabIndex        =   33
         Top             =   4455
         Width           =   1890
      End
      Begin VB.Label txt�˲��� 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   5085
         TabIndex        =   32
         Top             =   4455
         Width           =   1890
      End
      Begin VB.Label txt�˲����� 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   5085
         TabIndex        =   31
         Top             =   4800
         Width           =   1890
      End
      Begin VB.Label lbl�˲��� 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "  �˲���"
         Height          =   180
         Left            =   4320
         TabIndex        =   30
         Top             =   4500
         Width           =   720
      End
      Begin VB.Label lbl�˲����� 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�˲�����"
         Height          =   180
         Left            =   4275
         TabIndex        =   29
         Top             =   4860
         Width           =   720
      End
      Begin VB.Label lbl������ 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "������(&L)"
         Height          =   180
         Left            =   8790
         TabIndex        =   5
         Top             =   645
         Width           =   825
      End
      Begin VB.Label lblDifference 
         AutoSize        =   -1  'True
         Caption         =   "��ۺϼ�:"
         Height          =   180
         Left            =   4920
         TabIndex        =   22
         Top             =   3840
         Width           =   810
      End
      Begin VB.Label lblSalePrice 
         AutoSize        =   -1  'True
         Caption         =   "�ۼ۽��ϼ�:"
         Height          =   180
         Left            =   2040
         TabIndex        =   21
         Top             =   3840
         Width           =   1170
      End
      Begin VB.Label lblPurchasePrice 
         AutoSize        =   -1  'True
         Caption         =   "�ɱ����ϼ�:"
         Height          =   180
         Left            =   240
         TabIndex        =   20
         Top             =   3840
         Width           =   1170
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
         TabIndex        =   9
         Top             =   4155
         Width           =   650
      End
      Begin VB.Label LblTitle 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "�����������õ�"
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
         Top             =   135
         Width           =   11535
      End
      Begin VB.Label LblStock 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�ⷿ(&S)"
         Height          =   180
         Left            =   225
         TabIndex        =   0
         Top             =   660
         Width           =   630
      End
      Begin VB.Label LblEnterStock 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���ϲ���(&D)"
         Height          =   180
         Left            =   4635
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
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   23
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
            Picture         =   "frmDrawCard.frx":22EA
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
      Left            =   5040
      TabIndex        =   19
      Top             =   5640
      Visible         =   0   'False
      Width           =   495
   End
End
Attribute VB_Name = "frmDrawCard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mbln��������    As Boolean          '����ʱ���ݺ��ۼ�1
Private mstr��ⵥ�� As String              '��ⵥ��

Private mintUnit  As Integer                '��ʾ��λ:0-ɢװ��λ,1-��װ��λ
Private mblnFirst As Boolean
Private mint�༭״̬ As Integer             '1��������2���޸ģ�3����ˣ�4���鿴��5��������ˣ�6��������7-����ⵥ��ȡ����

Private mstr���ݺ� As String                '����ĵ��ݺ�;
Private mint��¼״̬ As Integer             '1:������¼;2-������¼;3-�Ѿ�������ԭ��¼
Private mblnSuccess As Boolean              'ֻҪ��һ�ųɹ�����ΪTrue������ΪFalse
Private mblnSave As Boolean                 '�Ƿ���̺����   TURE���ɹ���
Private mfrmMain As Form
Private mintcboIndex As Integer
Private mblnEdit As Boolean                 '�Ƿ�����޸�
Private mblnChange As Boolean               '�Ƿ���й��༭
Private mintParallelRecord As Integer       '���������󵥾ݲ���ִ�еĴ��� 1���������������2���Ѿ�ɾ���ļ�¼��3���Ѿ���˵ļ�¼
Private mint����� As Integer             '��ʾ�������ϳ���ʱ�Ƿ���п���飺0-�����;1-��飬�������ѣ�2-��飬�����ֹ
Private mcolUsedCount As Collection         '��ʹ�õ���������
Private mstrPrivs As String                     'Ȩ��
Private Const mstrCaption As String = "�������õ�"
Private mint������ȷ���� As Integer         '0-���ò���ȷ���� 1-������ȷ����

Private mlng����ID As Long          '����ⵥ��ȡ����ʱ��Ч
Private mstr������ As String        '����ⵥ��ȡ����ʱ��Ч
 '���˺�:2007/06/10:����10813
Private mstrTime_Start As String            '���뵥�ݱ༭�ĵ���ʱ�� ,��Ҫ�ж��Ƿ񵥾ݱ����˸��Ĺ�,����༭��,���ܽ������
Private mstrTime_End As String
Private mblnEnter As Boolean        '���ƶ���
Private Const mlngModule = 1717
Private mblnCostView As Boolean                 '�鿴�ɱ��� true-����鿴 false-������鿴
 
'���˺�:20060803�������ò�������
Private mbln��ͨ���� As Boolean
Private mblnHave������; As Boolean 'ȷ���Ƿ��ʼ���˲���������;��,�����ʼ,���ṩѡ����ѡ��,��������¼��.
Private mblnUpdate As Boolean               '��ʾ�Ƿ��Ѹ������¼۸���µ�������
Private mstrRequestNO As String     '���깺������NO ���մ��������깺����ʽ���ã��������깺������
Private mstr�ظ����� As String '��¼�ظ�������

Private recSort As ADODB.Recordset          '��ҩƷID�����ר�ü�¼��

'----------------------------------------------------------------------------------------------------------
'���˺�:����С��λ���ĸ�ʽ��
'�޸�:2007/03/06
Private mFMT As g_FmtString
'----------------------------------------------------------------------------------------------------------
'=========================================================================================
Private Type POINTAPI
     x As Long
     y As Long
End Type
Private Declare Function ClientToScreen Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long

Private Enum mBillCol
        C_����ID = 0
        C_�к� = 1
        C_���� = 2
        C_��� = 3
        c_��� = 4
        C_�������� = 5
        C_ָ������� = 6
        C_ʵ�ʽ�� = 7
        C_ʵ�ʲ�� = 8
        c_����ϵ�� = 9
        c_���� = 10
        C_���� = 11
        C_��׼�ĺ� = 12
        c_��λ = 13
        c_���� = 14
        C_Ч�� = 15
        C_���ʧЧ�� = 16
        C_�깺���� = 17
        C_��д���� = 18
        C_ʵ������ = 19
        c_ԭʼ���� = 20
        C_�ɹ��� = 21
        C_�ɹ���� = 22
        C_�ۼ� = 23
        C_�ۼ۽�� = 24
        C_��� = 25
        C_���ٱ�־ = 26
        C_������Ϣ = 27 '����ID|ʹ��ʱ��|����
        C_���ٲ��� = 28
        C_�������� = 29
        C_������� = 30 '����ʱ���ã���Ҫ��Ը�������ĳ���
End Enum
Private mstrĬ�ϲ�����; As String
Private Const mBillCols As Integer = 31              '������


'=========================================================================================


'�������������
Private Function GetDepend() As Boolean
    Dim rsTemp As New Recordset
    Dim strSQL As String
    
    On Error GoTo ErrHandle
    GetDepend = False
    strSQL = "" & _
        "   SELECT B.Id " & _
        "   FROM ҩƷ�������� A, ҩƷ������ B " & _
        "   Where A.���id = B.ID AND A.���� = 35"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "�������õ�")
    If rsTemp.EOF Then
        ShowMsgBox "û�������������õĳ����������������������ã�"
        rsTemp.Close
        Exit Function
    End If
    
    strSQL = "" & _
        "   SELECT DISTINCT a.id, a.���� " & _
        "   FROM ��������˵�� c, �������ʷ��� b, ���ű� a " & _
        "   Where c.�������� = b.����   AND a.id = c.����id and (a.վ��=[1] or a.վ�� is null) " & _
        "       AND ( TO_CHAR (a.����ʱ��, 'yyyy-MM-dd') = '3000-01-01' or a.����ʱ�� is null)"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, mstrCaption, gstrNodeNo)
    If rsTemp.EOF Then
        MsgBox "������ϵ��ȫ,���ڲ��Ź��������ã�", vbInformation, gstrSysName
        rsTemp.Close
        Exit Function
    End If
    
    strSQL = "Select ���� From ����������; where rownum<=1 "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "����������;-����")
    mblnHave������; = rsTemp.EOF = False
    
    strSQL = "Select ���� From ����������; where nvl(ȱʡ��־,0)=1"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "����������;-����")
    If rsTemp.EOF = False Then
        mstrĬ�ϲ�����; = zlStr.Nvl(rsTemp!����)
    Else
        mstrĬ�ϲ�����; = ""
    End If
    
    GetDepend = True
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Public Sub ShowCard(frmMain As Form, ByVal str���ݺ� As String, ByVal int�༭״̬ As Integer, _
    Optional int��¼״̬ As Integer = 1, Optional strPrivs As String, Optional blnSuccess As Boolean = False, _
    Optional lng���ò���id As Long = 0, Optional str������ As String = "")
    '-----------------------------------------------------------------------------------------------------------
    '����:�������
    '���:frmMain-���õ�������
    '    str���ݺ�-���ݺ�(���ڱ༭����Ϊ7<����ⵥ��ȡ����>,��ʾ��ⵥ��,����������õ��ݺ�
    '    int�༭״̬-�༭����:1.������2���޸ģ�3�����գ�4���鿴�� 6-��������,7-����ⵥ��ȡ����
    '    int��¼״̬-��¼״̬
    '    strPrivs-Ȩ�޴�
    '    lng����id-��������ò���ID(�༭����=7��Ч)
    '    str������-�����������(�༭����=7��Ч)
    '����:blnSuccess-���سɹ����,true,��ʾ������һ�ŵ��ݱ���ɹ�,�����ʾ��һ�ŵ��ݱ���ɹ�
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2008-10-27 11:45:51
    '-----------------------------------------------------------------------------------------------------------
    
    Dim strReg As String
    mlng����ID = lng���ò���id: mstr������ = str������
    
    mblnSave = False: mblnSuccess = False
    mstr��ⵥ�� = "": mstr���ݺ� = ""
    
    If int�༭״̬ = 7 Then
        mstr��ⵥ�� = str���ݺ�
    Else
        mstr���ݺ� = str���ݺ�
    End If
    
    mint�༭״̬ = int�༭״̬
    mint��¼״̬ = int��¼״̬
    mblnSuccess = blnSuccess
    mblnChange = False
    mintParallelRecord = 1
    mstrPrivs = strPrivs
    mblnCostView = zlStr.IsHavePrivs(mstrPrivs, "�鿴�ɱ���")
    
    Set mfrmMain = frmMain
    If Not GetDepend Then Exit Sub
    
    Call GetRegInFor(g˽��ģ��, "�������ù���", "���ݺ��ۼ�", strReg)
    mbln�������� = IIf(strReg = "", True, Val(strReg) = 1)
    
   
     
    If mint�༭״̬ = 1 Or mint�༭״̬ = 7 Then
'        If mbln�������� Then
'            mstr���ݺ� = NextNo(73)
'        End If
        mblnEdit = True

        txtNo.Locked = True
        txtNo.TabStop = True

        txtNo = mstr���ݺ�
        txtNo.Tag = txtNo.Text
    ElseIf mint�༭״̬ = 2 Then
        mblnEdit = True
    ElseIf mint�༭״̬ = 3 Then
        mblnEdit = False
        CmdSave.Caption = "���(&V)"
    ElseIf mint�༭״̬ = 5 Then
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
        CmdSave.Caption = "����(&O)"
        cmdAllCls.Visible = True
        cmdAllSel.Visible = True
    End If
      
    LblTitle.Caption = GetUnitName & LblTitle.Caption
    Me.Show vbModal, frmMain
    blnSuccess = mblnSuccess
    str���ݺ� = mstr���ݺ�
End Sub

Private Sub cboStock_Change()
    mblnChange = True
End Sub
Private Sub cboStock_Click()
    mint����� = Get������(cboStock.ItemData(cboStock.ListIndex))
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
            For i = 1 To mshBill.Rows - 1
                If mshBill.TextMatrix(i, 0) <> "" Then
                    Exit For
                End If
            Next
            If i <> mshBill.Rows Then
                If MsgBox("����ı�ⷿ���п���Ҫ�ı���Ӧ���ĵĵ�λ��" & vbCrLf & "��Ҫ������е������ݣ����Ƿ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
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
                .TextMatrix(intRow, mBillCol.C_�ۼ۽��) = Format(0, mFMT.FM_���)
                .TextMatrix(intRow, mBillCol.C_���) = Format(0, mFMT.FM_���)
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
                .TextMatrix(intRow, mBillCol.C_�ۼ۽��) = Format(.TextMatrix(intRow, mBillCol.C_��д����) * .TextMatrix(intRow, mBillCol.C_�ۼ�), mFMT.FM_���)
                .TextMatrix(intRow, mBillCol.C_���) = Format(.TextMatrix(intRow, mBillCol.C_�ۼ۽��) - .TextMatrix(intRow, mBillCol.C_�ɹ����), mFMT.FM_���)
            End If
        Next
    End With
    Call ��ʾ�ϼƽ��
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdDraw_Click()
    Dim rsTemp As New Recordset
    Dim blnClear As Boolean, blnCancel As Boolean
    Dim i As Long
    Dim vRect As RECT
    Dim strվ������ As String
    
    On Error GoTo ErrHandle
    vRect = zlControl.GetControlRect(txtDraw.hwnd)
    strվ������ = GetDeptStationNode(cboStock.ItemData(cboStock.ListIndex))
    
    If mbln��ͨ���� Then
        '��ͨ�������죬ֻ��ѡ���Լ������Ŀ���
        '���˺�:20060803
        '����:8468
        gstrSQL = "" & _
            " SELECT a.id, null as �ϼ�id, ĩ��, a.����,a.����,a.���� " & _
            " FROM ���ű� a " & _
            " Where (TO_CHAR (a.����ʱ��, 'yyyy-MM-dd') = '3000-01-01' Or a.����ʱ�� Is NULL) " & _
            IIf(strվ������ <> "", " And (a.վ�� = [2] or a.վ�� is null) ", "")
        gstrSQL = gstrSQL & " And a.ID in (Select ����ID From ������Ա where  ��Աid = [1] ) "
        Set rsTemp = zlDatabase.ShowSQLSelect(Me, gstrSQL, 0, "�������ò���ѡ��", False, "", "ѡ����ص����ò���", _
                     False, False, True, vRect.Left - 15, vRect.Top, txtDraw.Height, blnCancel, False, False, UserInfo.Id, strվ������)
    Else
        If gstrNodeNo = "-" Then
            'û��վ���,��������ʾ
            gstrSQL = "" & _
                " SELECT  a.id, �ϼ�id, ĩ��, a.����,a.����,a.���� " & _
                " FROM  ���ű� a " & _
                " Where (TO_CHAR (a.����ʱ��, 'yyyy-MM-dd') = '3000-01-01' Or a.����ʱ�� Is NULL) and (a.վ��=[1] or a.վ�� is null) "
            gstrSQL = gstrSQL & " start with �ϼ�id is null connect by prior id=�ϼ�id "
            Set rsTemp = zlDatabase.ShowSQLSelect(Me, gstrSQL, 1, "�������ò���ѡ��", False, "", "ѡ����ص����ò���", _
                         False, False, True, vRect.Left - 15, vRect.Top, txtDraw.Height, blnCancel, False, False, gstrNodeNo)
        Else
            '����վ�㣬��Ҫ�ǿ����ϼ�������վ���ţ����¼�δ���õ���������ֻ�����б�ʽ���д���
            gstrSQL = "" & _
                " SELECT  a.id, null as �ϼ�id, ĩ��, a.����,a.����,a.���� " & _
                " FROM  ���ű� a " & _
                " Where (TO_CHAR (a.����ʱ��, 'yyyy-MM-dd') = '3000-01-01' Or a.����ʱ�� Is NULL) " & _
                IIf(strվ������ <> "", " And (a.վ�� = [1] or a.վ�� is null) ", "")
            Set rsTemp = zlDatabase.ShowSQLSelect(Me, gstrSQL, 0, "�������ò���ѡ��", False, "", "ѡ����ص����ò���", _
                         False, False, True, vRect.Left - 15, vRect.Top, txtDraw.Height, blnCancel, False, False, strվ������)
        End If
    
    End If
       
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
    If rsTemp Is Nothing Then
        If txtDraw.Enabled Then txtDraw.SetFocus
        Exit Sub
    End If
    If rsTemp.State <> 1 Then
        If txtDraw.Enabled Then txtDraw.SetFocus
        Exit Sub
    End If
    blnClear = False
    If Val(txtDraw.Tag) <> Val(zlStr.Nvl(rsTemp!Id)) And Val(txtDraw.Tag) <> 0 Then
            '��Ҫ����Ƿ��Ѿ������úõĸ��ٲ�����Ϣ
            With mshBill
                For i = 1 To .Rows - 1
                    If Trim(.TextMatrix(.Row, mBillCol.C_������Ϣ)) <> "" And Trim(.TextMatrix(.Row, mBillCol.C_������Ϣ)) <> "||" Then
                        If MsgBox("�ڵ�" & i & "����,�Ѿ������˸��ٲ�����Ϣ, " & vbCrLf & "�Ƿ���Ҫ����Ѿ����úõĸ��ٲ�����Ϣ?", vbQuestion + vbDefaultButton2 + vbYesNo) = vbNo Then
                            If txtDraw.Enabled Then txtDraw.SetFocus
                            Exit Sub
                        Else
                            blnClear = True
                            Exit For
                        End If
                        
                    End If
                Next
                If blnClear Then
                    For i = 1 To .Rows - 1
                        .TextMatrix(i, mBillCol.C_������Ϣ) = ""
                        .TextMatrix(i, mBillCol.C_���ٲ���) = ""
                    Next
                End If
                
            End With
    End If
    
    Me.txtDraw = zlStr.Nvl(rsTemp!����) & "-" & zlStr.Nvl(rsTemp!����)
    Me.txtDraw.Tag = zlStr.Nvl(rsTemp!Id)
    
    gstrSQL = "Select ��������, ����id, ������� From ��������˵�� Where ����id=[1] And ��������='����'"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, Val(zlStr.Nvl(rsTemp!Id)))
    If rsTemp.EOF Then
        gstrSQL = "Select ��������, ����id, ������� From ��������˵�� Where ����id=[1] And ��������='�ٴ�'"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, Val(Me.txtDraw.Tag))
        If rsTemp.EOF = False Then
            cmdDraw.Tag = "�ٴ�"
        Else
            cmdDraw.Tag = ""
        End If
    Else
        cmdDraw.Tag = "����"
    End If
    If txtDrawPerson.Enabled Then txtDrawPerson.SetFocus
    Local���ٲ�����Ϣ
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cmdDrawPerson_Click()
    If ShowSelect("") = False Then Exit Sub
    mshBill.SetFocus
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
    Else
        FindRownew mshBill, mBillCol.C_����, txtCode.Text, True
        lblCode.Visible = False
        txtCode.Visible = False
    End If
End Sub

Private Sub CmdHelp_Click()
    Call ShowHelp(App.ProductName, Me.hwnd, Me.Name, Int(glngSys / 100))
End Sub

Private Sub cmdRequestDraw_Click()
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
    
    If Val(txtDraw.Tag) = 0 Then
        MsgBox "���ϲ��Ų���Ϊ�գ�", vbInformation, gstrSysName
        txtDraw.SetFocus
        Exit Sub
    End If
    
    mstrRequestNO = frmDrawCondition.ShowMe(Me, mintUnit, cboStock.List(cboStock.ListIndex), cboStock.ItemData(cboStock.ListIndex), txtDraw.Text, txtDraw.Tag)
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
        
        gstrSQL = "Select a.Id as ����id, d.���� as �ƻ�����,a.����,a.���� ,a.���,c.�ּ� as �ۼ�,a.���㵥λ as ɢװ��λ,a.�Ƿ��� as ʱ��,b.��װ��λ,b.����ϵ��,b.ָ�������" & vbNewLine & _
                    ",e.�ϴβ��� as ����,e.�ϴ����� as ����,nvl(e.����,0) as ����,e.Ч��,e.���Ч��,e.��������,nvl(e.ʵ������,0) as ʵ������,e.ʵ�ʽ��,e.ʵ�ʲ��,e.���ۼ�,e.ƽ���ɱ���,e.��׼�ĺ�,b.�ⷿ����,b.���÷���, nvl(b.���ٲ���,0) as ���ٲ���" & vbNewLine & _
                    "From �շ���ĿĿ¼ A, �������� B, �շѼ�Ŀ C," & vbNewLine & _
                    "     (Select  b.����id, Sum(b.�ƻ�����) As ����" & vbNewLine & _
                    "       From ���ϲɹ��ƻ� A, ���ϼƻ����� B" & vbNewLine & _
                    "       Where a.Id = b.�ƻ�id and a.����=1 And a.No In (Select * From Table(Cast(f_Str2list([1]) As Zltools.t_Strlist)))" & vbNewLine & _
                    "       Group By b.����id) D,ҩƷ��� e" & vbNewLine & _
                    "Where a.Id = b.����id And b.����id = c.�շ�ϸĿid And a.Id = d.����id and b.����id=e.ҩƷid(+)  and e.�ⷿid=[2] and e.ʵ������>0 and e.����=1 And Sysdate Between c.ִ������ And c.��ֹ����" & _
                    GetPriceClassString("C")
        
        If gSystem_Para.P156_�����㷨 = 0 Then
            gstrSQL = gstrSQL & " Order by a.id,Nvl(e.����, 0)"
        Else
            gstrSQL = gstrSQL & " Order by a.id,e.Ч��,Nvl(e.����, 0)"
        End If
        
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "cmdRequestDraw_Click", mstrRequestNO, cboStock.ItemData(cboStock.ListIndex))
                
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
                If Format(str���Ч��, "yyyy-mm-dd") < Format(sys.Currentdate, "yyyy-mm-dd") And Trim(str���Ч��) <> "" Then
                   If MsgBox("[" & rsTemp!���� & "-" & rsTemp!���� & "]" & "�����Ѿ��������Ч��,�Ƿ�Ҫ���ã�", vbYesNo + vbDefaultButton2 + vbQuestion, gstrSysName) <> vbYes Then
                        blnDo = True
                   End If
                End If
                
                strЧ�� = IIf(IsNull(rsTemp!Ч��), "", Format(rsTemp!Ч��, "yyyy-MM-dd"))
                If IsDate(strЧ��) Then
                    If Format(strЧ��, "yyyy-MM-dd") < Format(sys.Currentdate, "yyyy-MM-dd") Then
                        MsgBox "[" & rsTemp!���� & "-" & rsTemp!���� & "]" & "���������Ѿ�ʧЧ�ˣ�", vbInformation, gstrSysName
                    End If
                End If
                
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
                
                'ֻ�в��ظ��Ĳ���ӵ������ȥ
                If blnDo = False Then
                    SetRequestColValue lngRow, rsTemp!����ID, "[" & rsTemp!���� & "]" & rsTemp!����, _
                                IIf(IsNull(rsTemp!���), "", rsTemp!���), IIf(IsNull(rsTemp!����), "", rsTemp!����), _
                                IIf(mintUnit = 0, rsTemp!ɢװ��λ, rsTemp!��װ��λ), _
                                dblPrice, rsTemp!ƽ���ɱ���, IIf(IsNull(rsTemp!����), "", rsTemp!����), _
                                IIf(IsNull(rsTemp!Ч��), "", Format(rsTemp!Ч��, "yyyy-MM-dd")), _
                                IIf(IsNull(rsTemp!���Ч��), "", Format(rsTemp!���Ч��, "yyyy-MM-dd")), _
                                rsTemp!�ƻ�����, _
                                IIf(IsNull(rsTemp!��������), "0", rsTemp!��������), _
                                dbl����, _
                                IIf(IsNull(rsTemp!ָ�������), "0", rsTemp!ָ�������), _
                                IIf(mintUnit = 0, 1, rsTemp!����ϵ��), IIf(IsNull(rsTemp!����), 0, rsTemp!����), rsTemp!ʱ��, rsTemp!���÷���, IIf(IsNull(rsTemp!��׼�ĺ�), "", rsTemp!��׼�ĺ�), rsTemp!���ٲ���, rsTemp!�ⷿ����
                End If
                blnDo = False
                rsTemp.MoveNext
            End With
        Loop
    End If
End Sub

Private Function SetRequestColValue(ByVal intRow As Integer, ByVal lng����ID As Long, _
        ByVal str���� As String, ByVal str��� As String, ByVal str���� As String, _
        ByVal str��λ As String, ByVal num�ۼ� As Double, ByVal num�ɱ��� As Double, ByVal str���� As String, _
        ByVal strЧ�� As String, ByVal str���ʧЧ�� As String, ByVal num�깺���� As Double, ByVal num�������� As Double, ByVal numʵ������ As Double, _
        ByVal numָ������� As Double, _
        ByVal num����ϵ�� As Double, ByVal lng���� As Long, _
        ByVal int�Ƿ��� As Integer, ByVal int���÷��� As Integer, ByVal str��׼�ĺ� As String, ByVal int���ٲ��� As Integer, ByVal int�ⷿ���� As Integer) As Boolean
    
        Dim intCount As Integer
        Dim intCol As Integer
        Dim dblPrice As Double
        Dim rsTemp As New Recordset
        Dim bln���� As Boolean
        Dim lngRow As Long
        
    On Error GoTo ErrHandle
    SetRequestColValue = False
    
    With mshBill
        .TextMatrix(intRow, 0) = lng����ID
        .TextMatrix(intRow, mBillCol.C_����) = str����
        .TextMatrix(intRow, mBillCol.c_���) = str���
        .TextMatrix(intRow, mBillCol.C_����) = str����
        .TextMatrix(intRow, mBillCol.C_��׼�ĺ�) = str��׼�ĺ�
        .TextMatrix(intRow, mBillCol.c_��λ) = str��λ
        .TextMatrix(intRow, mBillCol.c_����) = str����
        .TextMatrix(intRow, mBillCol.C_Ч��) = Format(strЧ��, "yyyy-mm-dd")
        .TextMatrix(intRow, mBillCol.C_���ʧЧ��) = Format(str���ʧЧ��, "yyyy-mm-dd")
    
        .TextMatrix(intRow, mBillCol.C_�ۼ�) = Format(num�ۼ� * num����ϵ��, mFMT.FM_���ۼ�)
        .TextMatrix(intRow, mBillCol.C_�ɹ���) = Format(num�ɱ��� * num����ϵ��, mFMT.FM_�ɱ���)
        .TextMatrix(intRow, mBillCol.C_�깺����) = Format(num�깺���� / num����ϵ��, mFMT.FM_����)
        .TextMatrix(intRow, mBillCol.C_��������) = Format(num�������� / num����ϵ��, mFMT.FM_����)
        .TextMatrix(intRow, mBillCol.C_��д����) = Format(numʵ������ / num����ϵ��, mFMT.FM_����)
        .TextMatrix(intRow, mBillCol.C_ʵ������) = Format(numʵ������ / num����ϵ��, mFMT.FM_����)
        .TextMatrix(intRow, mBillCol.C_�ۼ۽��) = Format(Val(.TextMatrix(intRow, mBillCol.C_�ۼ�)) * Val(.TextMatrix(intRow, mBillCol.C_��д����)), mFMT.FM_���)
        .TextMatrix(intRow, mBillCol.C_�ɹ����) = Format(Val(.TextMatrix(intRow, mBillCol.C_�ɹ���)) * Val(.TextMatrix(intRow, mBillCol.C_��д����)), mFMT.FM_���)
        .TextMatrix(intRow, mBillCol.C_���) = Format(Val(.TextMatrix(intRow, mBillCol.C_�ۼ۽��)) - Val(.TextMatrix(intRow, mBillCol.C_�ɹ����)), mFMT.FM_���)
        .TextMatrix(intRow, mBillCol.C_ָ�������) = numָ������� & "||" & int�Ƿ��� & "||" & int���÷���
        .TextMatrix(intRow, mBillCol.c_����ϵ��) = num����ϵ��
        .TextMatrix(intRow, mBillCol.c_����) = lng����
        .TextMatrix(intRow, mBillCol.C_��������) = Check��������(intRow, int���÷���, int�ⷿ����)
        
        .TextMatrix(intRow, mBillCol.C_���ٱ�־) = int���ٲ���
    End With
'    Call ��ʾ�����
    SetRequestColValue = True
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub cmdSel_Click()
    Dim lng����ID As Long, lng�շ�ID As Long, lng����id As Long, strʹ��ʱ�� As String, str���� As String, blnEdit As Boolean
    Dim strTemp As String, arrtemp As Variant
    Dim str���� As String
    
    If Val(txtDraw.Tag) = 0 Then
        ShowMsgBox "���ò���δѡ��,����ѡ�����ò��ź���ѡ����!"
        Exit Sub
    End If
    
    lng�շ�ID = Get�շ�ID()
    With mshBill
        lng����ID = Val(.TextMatrix(.Row, 0))
        strTemp = .TextMatrix(.Row, C_������Ϣ)
        If Trim(strTemp) <> "" Then
            arrtemp = Split(strTemp, "|")
            lng����id = Val(arrtemp(0))
            strʹ��ʱ�� = arrtemp(1)
            str���� = arrtemp(2)
        Else
            lng����id = 0
            strʹ��ʱ�� = ""
            str���� = ""
        End If
    End With
    blnEdit = IIf(mint�༭״̬ = 1 Or mint�༭״̬ = 2 Or mint�༭״̬ = 7, True, False)
    If frmDrawPatiInfor.ShowEdit(Me, lng�շ�ID, Val(txtDraw.Tag), cmdDraw.Tag, lng����ID, blnEdit, lng����id, str����, strʹ��ʱ��, str����) = False Then
        mshBill.SetFocus
        Exit Sub
    End If
    With mshBill
        .TextMatrix(.Row, mBillCol.C_������Ϣ) = lng����id & "|" & strʹ��ʱ�� & "|" & str����
        .TextMatrix(.Row, mBillCol.C_���ٲ���) = str����
        .SetFocus
    End With
End Sub

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
        dbl����ϵ�� = Val(mshBill.TextMatrix(lngRow, mBillCol.c_����ϵ��))
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
        
        '�������������Ϊ�㣬��Ҫ�Զ��ֽ⣻���β�Ϊ0�������ʱ����˶�Ӧ���ε�����
        blnAddRow = False
        If bln���� = True And lng���� = 0 Then
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
                    mblnChange = True
                    
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
                    mshBill.TextMatrix(lngRow, mBillCol.C_���) = lngRow
                    mshBill.TextMatrix(lngRow, mBillCol.c_����) = rsCheck!����
                    mshBill.TextMatrix(lngRow, mBillCol.c_����) = IIf(IsNull(rsCheck!����), "", rsCheck!����)
                    mshBill.TextMatrix(lngRow, mBillCol.C_Ч��) = IIf(IsNull(rsCheck!Ч��), "", rsCheck!Ч��)
                    mshBill.TextMatrix(lngRow, mBillCol.C_����) = IIf(IsNull(rsCheck!����), "", rsCheck!����)
                    mshBill.TextMatrix(lngRow, mBillCol.C_��׼�ĺ�) = IIf(IsNull(rsCheck!��׼�ĺ�), "", rsCheck!��׼�ĺ�)
                    mshBill.TextMatrix(lngRow, mBillCol.C_��������) = IIf(bln���� = True, 1, 0)
                    
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
                        mshBill.TextMatrix(lngRow, mBillCol.c_ԭʼ����) = Val(mshBill.TextMatrix(lngRow, mBillCol.C_ʵ������)) * Val(mshBill.TextMatrix(lngRow, mBillCol.c_����ϵ��))
                    End If
                    
                    mshBill.TextMatrix(lngRow, mBillCol.C_��д����) = Format(dbl����, mFMT.FM_����)
                    mshBill.TextMatrix(lngRow, mBillCol.C_ʵ������) = Format(dbl����, mFMT.FM_����)
                    
                    If Trim(mshBill.TextMatrix(lngRow, mBillCol.C_ʵ������)) = "" Then mshBill.TextMatrix(lngRow, mBillCol.C_ʵ������) = 0
                    mshBill.TextMatrix(lngRow, mBillCol.C_ʵ�ʲ��) = Format(rsCheck!ʵ�ʲ��, mFMT.FM_���)
                    mshBill.TextMatrix(lngRow, mBillCol.C_ʵ�ʽ��) = Format(rsCheck!ʵ�ʽ��, mFMT.FM_���)
                    mshBill.TextMatrix(lngRow, mBillCol.C_��������) = Format(rsCheck!��������, mFMT.FM_���)
                    mshBill.TextMatrix(lngRow, mBillCol.C_�ۼ�) = Format(IIf(blnʱ��, dbl�ּ�_ʱ��, dbl�ּ�), mFMT.FM_���ۼ�)
                    mshBill.TextMatrix(lngRow, mBillCol.C_�ۼ۽��) = Format(Val(mshBill.TextMatrix(lngRow, mBillCol.C_�ۼ�)) * dbl����, mFMT.FM_���)
                    
                    '�����µķ�ʽ����ɱ��� �ɱ���=ҩƷ���.ƽ���ɱ���
                    mshBill.TextMatrix(lngRow, mBillCol.C_�ɹ���) = Format(Get�ɱ���(lng����ID, lng�ⷿID, Val(mshBill.TextMatrix(lngRow, mBillCol.c_����))) * dbl����ϵ��, mFMT.FM_�ɱ���)
                    mshBill.TextMatrix(lngRow, mBillCol.C_�ɹ����) = Format(Val(mshBill.TextMatrix(lngRow, mBillCol.C_�ɹ���)) * dbl����, mFMT.FM_���)
                    mshBill.TextMatrix(lngRow, mBillCol.C_���) = Format(Val(mshBill.TextMatrix(lngRow, mBillCol.C_�ۼ۽��)) - Val(mshBill.TextMatrix(lngRow, mBillCol.C_�ɹ����)), mFMT.FM_���)
                    
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
                mshBill.TextMatrix(lngRow, mBillCol.C_���) = lngRow
                mshBill.TextMatrix(lngRow, mBillCol.C_ʵ������) = ""
                mshBill.TextMatrix(lngRow, mBillCol.C_�ۼ۽��) = ""
                mshBill.TextMatrix(lngRow, mBillCol.C_�ɹ����) = ""
                mshBill.TextMatrix(lngRow, mBillCol.C_���) = ""
            End If
        Else
            mshBill.TextMatrix(lngRow, mBillCol.C_�к�) = lngRow
            mshBill.TextMatrix(lngRow, mBillCol.C_���) = lngRow
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

Private Sub Form_Activate()
    If mblnFirst = False Then Exit Sub
    mblnFirst = False
    
'    mblnChange = False
    Select Case mintParallelRecord
        Case 1
            '����
        Case 2
            If mint�༭״̬ = 6 Then
                ShowMsgBox "�õ�����û�п��Գ��������ģ����飡"
            Else
                '�����ѱ�ɾ��
                ShowMsgBox "�õ����ѱ�ɾ�������飡"
            End If
            Unload Me
            Exit Sub
        Case 3
            '�޸ĵĵ����ѱ����
            ShowMsgBox "�õ����ѱ���������ˣ����飡"
            Unload Me
            Exit Sub
    End Select
    If mint�༭״̬ = 7 Then
        If IsCtrlSetFocus(CmdSave) Then
            zlControl.ControlSetFocus CmdSave
        End If
    End If
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
            dbl����ϵ�� = Val(mshBill.TextMatrix(lngRow, mBillCol.c_����ϵ��))
            dbl��д���� = Val(mshBill.TextMatrix(lngRow, mBillCol.C_ʵ������))
            
            dbl�������� = 0
            '���Ҹò��ϵĿ���¼
            rsCheck.Filter = "����ID=" & lng����ID & " And ����=" & lng����
            If rsCheck.RecordCount <> 0 Then
                dbl�������� = zlStr.Nvl(rsCheck!ʵ������, 0) / dbl����ϵ��
            End If
            
            '������Ŀ�����������
            If Not (dbl�������� >= dbl��д����) Then
                int����� = mint�����
                
                rsProperty.Filter = "����ID=" & lng����ID
                
                If Not (Val(mshBill.TextMatrix(lngRow, mBillCol.C_��������)) = 0 And Split(mshBill.TextMatrix(lngRow, mBillCol.C_ָ�������), "||")(1) = 0) Then
                    '����ò�����ʱ�ۻ��������治�㲻������⣬�൱�ڽ�ֹ���⣻���۲���������Ҫ�жϷֽ⣬ֻ��Ҫ���ݲ������Ƽ���
                    bln���� = (IIf(bln�ⷿ, (rsProperty!�ⷿ���� = 1), (rsProperty!���÷��� = 1)) Or (rsProperty!�Ƿ��� = 1))
                    strMsg = ""
                    If bln���� Then
                        int����� = 2
                        '��������β��ϣ�������С�ڵ����㣬˵��δִ�зֽ⹦��
                        If lng���� <= 0 And IIf(bln�ⷿ, (rsProperty!�ⷿ���� = 1), (rsProperty!���÷��� = 1)) Then
                            strMsg = "������ִ�зֽ⹦����ȷ���β��ϵĳ������Σ�"
                        End If
                    End If
                End If
                '���������̽�����ʾ���ֹ
                Select Case int�����
                Case 1  '����ʾ
                    If MsgBox(rsProperty!ͨ���� & "�Ŀ�治��" & "(���ʵ������Ϊ" & dbl�������� & ")���Ƿ������" & strMsg, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                        mshBill.Row = lngRow
                        mshBill.MsfObj.TopRow = lngRow
                        Exit Function
                    End If
                Case 2
                    MsgBox rsProperty!ͨ���� & "�Ŀ�治��" & "(���ʵ������Ϊ" & dbl�������� & ")��" & strMsg, vbInformation, gstrSysName
                    mshBill.Row = lngRow
                    mshBill.MsfObj.TopRow = lngRow
                    Exit Function
                End Select
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

Private Sub CmdSave_Click()
    Dim blnSuccess As Boolean
    Dim strReg As String
    Dim intRow As Integer
    
    '�����������ݼ�
    Call SetSortRecord
    
    If mint�༭״̬ = 4 Then    '�鿴
        '��ӡ
        printbill
        '�˳�
        Unload Me
        Exit Sub
    End If
    
    If mint�༭״̬ = 5 Then        '�������
        
        If Not ���ϵ������(Txt������.Caption) Then Exit Sub
        
        mstrTime_End = GetBillInfo(20, txtNo.Tag)
        If mstrTime_End = "" Then
            MsgBox "ע��:" & vbCrLf & "  �õ����Ѿ�����������Աɾ��,���ܼ�����", vbInformation, gstrSysName
            Exit Sub
        End If
        If mstrTime_End <> mstrTime_Start Then
            If MsgBox("ע��:" & vbCrLf & "  �õ����Ѿ�����������Ա�༭�����ܼ���!" & vbCrLf & "  �Ƿ�����ˢ�µ���?", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
                Call initCard
            End If
            Exit Sub
        End If
        '����Ƿ�ֽ�
        If CheckStock = False Then Exit Sub
        
        If Not ��鵥��(20, txtNo.Tag, False) And Not mblnUpdate Then
            MsgBox "�м�¼δʹ�����¼۸񣬳����Զ���ɸ��£��ۼۡ��ɱ��ۡ��ۼ۽��ɱ�����ۣ������º����飡", vbInformation, gstrSysName
            Call RefreshBill
            mblnUpdate = True
            mblnChange = True
            Exit Sub
        End If
        
        '������ʱ�޸��˵��ݣ����������ɵ��ݱ���
        If mblnChange Then
            If Not SaveCard() Then
                gcnOracle.RollbackTrans: Exit Sub
            End If
        End If
                
        If SaveCheck = True Then
            Unload Me
        End If
        Exit Sub
    End If
    
    If mint�༭״̬ = 3 Then        '���
        
        If Not ���ϵ������(Txt������.Caption) Then Exit Sub
        
        '���˺�:2007/06/10:����10813
        mstrTime_End = GetBillInfo(20, txtNo.Tag)
        If mstrTime_End = "" Then
            MsgBox "ע��:" & vbCrLf & "  �õ����Ѿ�����������Աɾ��,���ܼ�����", vbInformation, gstrSysName
            Exit Sub
        End If
        If mstrTime_End <> mstrTime_Start Then
            If MsgBox("ע��:" & vbCrLf & "  �õ����Ѿ�����������Ա�༭�����ܼ���!" & vbCrLf & "  �Ƿ�����ˢ�µ���?", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
                Call initCard
            End If
            Exit Sub
        End If
        
        '����Ƿ�ֽ�
        If CheckStock = False Then Exit Sub
        
        For intRow = 1 To mshBill.Rows - 1
            If Val(mshBill.TextMatrix(intRow, 0)) <> 0 Then
                If Val(mshBill.TextMatrix(intRow, mBillCol.C_��������)) = 1 And Val(mshBill.TextMatrix(intRow, mBillCol.C_ʵ������)) = 0 Then
                    MsgBox "��" & intRow & "�е������������������޿�棬������0�������ã�", vbInformation, gstrSysName
                    Exit Sub
                End If
            End If
        Next
        
        If Not ��鵥��(20, txtNo.Tag, False) And Not mblnUpdate Then
            MsgBox "�м�¼δʹ�����¼۸񣬳����Զ���ɸ��£��ۼۡ��ɱ��ۡ��ۼ۽��ɱ�����ۣ������º����飡", vbInformation, gstrSysName
            Call RefreshBill
            mblnUpdate = True
            mblnChange = True
            Exit Sub
        End If
        
        '������ʱ�޸��˵��ݣ����������ɵ��ݱ���
        If mblnChange Then
            If Not SaveCard(True) Then
                gcnOracle.RollbackTrans: Exit Sub
            End If
        End If
                
        If SaveCheck = True Then
            strReg = IIf(Val(zlDatabase.GetPara("��˴�ӡ", glngSys, mlngModule, "0")) = 1, 1, 0)
            If Val(strReg) = 1 Then
                '��ӡ
                If InStr(mstrPrivs, "���ݴ�ӡ") <> 0 Then
                    printbill
                End If
            End If
            Unload Me
        End If
        Exit Sub
    End If
    
    If mint�༭״̬ = 6 Then '����
        For intRow = 1 To mshBill.Rows - 1
            If Val(mshBill.TextMatrix(intRow, mBillCol.C_ʵ������)) < 0 Then '�������ó����ż��
                If CompareUsableQuantity(intRow, Val(mshBill.TextMatrix(intRow, mBillCol.C_ʵ������))) = False Then
                    mshBill.SetFocus
                    mshBill.Row = intRow
                    Exit Sub
                End If
            End If
        Next
        
        If SaveStrike Then Unload Me
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
   
    If mint�༭״̬ = 7 Then    '����ⵥ��ȡ
        Unload Me
        Exit Sub
    End If
    
'    If mbln�������� Then
'        mstr���ݺ� = NextNo(73)
'        txtNO = mstr���ݺ�
'    End If
    
    mblnSave = False
    mblnEdit = True
    mshBill.ClearBill
    Call RefreshRowNO(mshBill, mBillCol.C_�к�, 1)

    txtDraw.Text = ""
    txtDraw.Tag = "0"
    txtժҪ.Text = ""
    If txtDraw.Enabled = True Then
        txtDraw.SetFocus
        txtDraw.SelStart = 0
        txtDraw.SelLength = Len(txtDraw.Text)
    End If
    mblnChange = False
    If txtNo.Tag <> "" Then Me.stbThis.Panels(2).Text = "��һ�ŵ��ݵ�NO�ţ�" & txtNo.Tag
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
            " Where a.���� = 20 And a.No = [1] And a.ҩƷid = b.�շ�ϸĿid And c.Id = b.�շ�ϸĿid And Round(a.���ۼ�," & g_С��λ��.obj_ɢװС��.���ۼ�С�� & ") <> Round(b.�ּ�, " & g_С��λ��.obj_ɢװС��.���ۼ�С�� & ") And" & _
              "    NVL(c.�Ƿ���, 0) = 0" & _
            " Union All" & _
            " Select '�ۼ�' As ����, a.���, a.ҩƷid As ����id, Nvl(a.����, 0) As ����, decode(nvl(b.����,0),0,b.ʵ�ʽ�� / b.ʵ������,b.���ۼ�) As �ּ�" & _
            " From ҩƷ�շ���¼ A, ҩƷ��� B, �շ���ĿĿ¼ C" & _
            " Where a.���� = 20 And a.No = [1] And c.Id = a.ҩƷid And Round(a.���ۼ�," & g_С��λ��.obj_ɢװС��.���ۼ�С�� & ") <> Round(decode(nvl(b.����,0),0,b.ʵ�ʽ�� / b.ʵ������,b.���ۼ�), " & g_С��λ��.obj_ɢװС��.���ۼ�С�� & ") And Nvl(c.�Ƿ���, 0) = 1 And" & _
                  " b.���� = 1 And b.�ⷿid = a.�ⷿid And b.ҩƷid = a.ҩƷid And NVL(b.����, 0) = NVL(a.����, 0) And NVL(b.ʵ������, 0) <> 0 And a.���ϵ�� = -1" & _
            " Union All" & _
            " Select '�ɱ���' As ����, a.���, a.ҩƷid As ����id, Nvl(a.����, 0) As ����, b.ƽ���ɱ��� As �ּ�" & _
            " From ҩƷ�շ���¼ A, ҩƷ��� B" & _
            " Where a.���� = 20 And a.No = [1] And a.ҩƷid = b.ҩƷid And Nvl(a.����, 0) = Nvl(b.����, 0) and round(a.�ɱ���," & g_С��λ��.obj_ɢװС��.�ɱ���С�� & ")<>round(b.ƽ���ɱ���," & g_С��λ��.obj_ɢװС��.�ɱ���С�� & ") And a.�ⷿid = b.�ⷿid and a.���ϵ��=-1 and b.����=1" & _
            " Order By ����, ����id, ���"

    Set rsprice = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption & "[ȡ��ǰ�۸�]", CStr(Me.txtNo.Text))
    
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
                dbl���ۼ� = Val(Format(rsprice!�ּ� * Val(mshBill.TextMatrix(lngRow, mBillCol.c_����ϵ��)), mFMT.FM_���ۼ�))
                dbl���۽�� = Val(Format(dbl���ۼ� * dbl����, mFMT.FM_���))
                dbl��� = Val(Format(dbl���۽�� - dbl�ɱ����, mFMT.FM_���))
            End If

            rsprice.Filter = "����='�ɱ���' And ����id=" & lng����ID & " And ����=" & Val(mshBill.TextMatrix(lngRow, mBillCol.c_����))
            If rsprice.RecordCount > 0 Then
                blnAdj = True
                dbl���۽�� = Val(Format(dbl���ۼ� * dbl����, mFMT.FM_���))
                dbl�ɱ��� = Val(Format(rsprice!�ּ� * Val(mshBill.TextMatrix(lngRow, mBillCol.c_����ϵ��)), mFMT.FM_���))
                dbl�ɱ���� = Val(Format(dbl�ɱ��� * dbl����, mFMT.FM_���))
                dbl��� = Val(Format(dbl���۽�� - dbl�ɱ����, mFMT.FM_���))
            End If

            If blnAdj = True Then
                '�Ե�ǰ���¼۸����µ���������ݣ��ۼۡ��ɱ��ۡ����۽��ɱ�����ۣ�
                mshBill.TextMatrix(lngRow, mBillCol.C_�ۼ�) = Format(dbl���ۼ�, mFMT.FM_���ۼ�)
                mshBill.TextMatrix(lngRow, mBillCol.C_�ۼ۽��) = Format(dbl���۽��, mFMT.FM_���)
                mshBill.TextMatrix(lngRow, mBillCol.C_�ɹ���) = Format(dbl�ɱ���, mFMT.FM_�ɱ���)
                mshBill.TextMatrix(lngRow, mBillCol.C_�ɹ����) = Format(dbl�ɱ����, mFMT.FM_���)
                mshBill.TextMatrix(lngRow, mBillCol.C_���) = Format(dbl���, mFMT.FM_���)
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

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Dim strReg As String
    
    mblnUpdate = False
    strReg = Val(zlDatabase.GetPara("���ĵ�λ", glngSys, mlngModule, "0"))
    mintUnit = Val(strReg)
    mint������ȷ���� = Val(zlDatabase.GetPara(258, glngSys, 0))
    
    mblnFirst = True
  
    '���˺�:����С����ʽ����
    With mFMT
        .FM_�ɱ��� = GetFmtString(mintUnit, g_�ɱ���)
        .FM_��� = GetFmtString(mintUnit, g_���)
        .FM_���ۼ� = GetFmtString(mintUnit, g_�ۼ�)
        .FM_���� = GetFmtString(mintUnit, g_����)
    End With
    
       
    txtNo = mstr���ݺ�
    txtNo.Tag = txtNo.Text
    
    '------------------------------------------------------------------------------------------------------------------
    '���˺�:20060803:��������
    '����:8468
    mbln��ͨ���� = Check��ͨ����
    '------------------------------------------------------------------------------------------------------------------
    Call initCard
    
    '�ָ����Ի���������
    RestoreWinState Me, App.ProductName, mstrCaption
    '�ָ����Ի��������ú󣬻���Ҫ��Ȩ�޿��Ƶ��н�һ������
    With mshBill
        .ColWidth(mBillCol.C_�ɹ���) = IIf(mblnCostView = True, 900, 0)
        .ColWidth(mBillCol.C_�ɹ����) = IIf(mblnCostView = True, 900, 0)
        .ColWidth(mBillCol.C_���) = IIf(mblnCostView = True, 900, 0)
    End With
    mblnChange = False
End Sub

Private Sub initCard()
    Dim i As Integer
    Dim rsTemp As New Recordset
    Dim strUnit As String
    Dim strUnitQuantity As String
    Dim intRow As Integer
    Dim numUseAbleCount As Double
    Dim vardrug As Variant
    Dim strOrder As String, strCompare As String
    Dim str���� As String, strArray As String
    
    '�������޸İ����쵥���ð�ť�ɼ�
    If mint�༭״̬ = 1 Or mint�༭״̬ = 2 Then
        cmdRequestDraw.Visible = True
    Else
        cmdRequestDraw.Visible = False
    End If
    
    '�ⷿ
    On Error GoTo ErrHandle
    strOrder = zlDatabase.GetPara("��������", glngSys, mlngModule, "00")
    strOrder = IIf(strOrder = "", "00", strOrder)
    
    strCompare = Mid(strOrder, 1, 1)
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
            Txt������ = UserInfo.�û���
            Txt�������� = Format(sys.Currentdate, "yyyy-mm-dd hh:mm:ss")
            initGrid
            '�������ͨ��������,���Ƿ�ֻ�߱�һ������,�����ǰ��Աֻ����һ�����ң�����������ò�����
            If mbln��ͨ���� Then
                gstrSQL = "" & _
                   "   SELECT DISTINCT a.id, a.����,a.����,a.���� " & _
                   "   FROM ��������˵�� c, �������ʷ��� b, ���ű� a " & _
                   "   Where c.�������� = b.���� " & _
                   "           AND a.id = c.����id " & _
                   "           AND (TO_CHAR (a.����ʱ��, 'yyyy-MM-dd') = '3000-01-01' Or a.����ʱ�� Is NULL)" & _
                   "           And a.ID in (Select ����ID From ������Ա where ȱʡ=1 and ��Աid =[1])"
                Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, UserInfo.Id)
                
                If Not rsTemp.EOF Then
                    Me.txtDraw = zlStr.Nvl(rsTemp!����) & "-" & zlStr.Nvl(rsTemp!����)
                    Me.txtDraw.Tag = zlStr.Nvl(rsTemp!Id)
                End If
                txtDrawPerson.Text = gstrUserName
                txtDrawPerson.Tag = gstrUserName
            End If
            txtժҪ = mstrĬ�ϲ�����;
        Case 2, 3, 4, 5, 6, 7
            initGrid
            
            If mint�༭״̬ = 4 Then
                gstrSQL = "" & _
                    "   Select b.id,b.���� " & _
                    "   From ҩƷ�շ���¼ a,���ű� b " & _
                    "   Where a.�ⷿid=b.id and A.���� = 20 and a.no=[1] "
                Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, mstr���ݺ�)
                
                If rsTemp.EOF Then
                    mintParallelRecord = 2
                    Exit Sub
                End If

                With cboStock
                    .AddItem rsTemp!����
                    .ItemData(.NewIndex) = rsTemp!Id
                    .ListIndex = 0
                End With
                rsTemp.Close
            End If
            Select Case mintUnit
                Case 0
                    strUnitQuantity = "d.���㵥λ AS ��λ,a.���� as �깺����, A.��д����,a.ʵ������,a.�ɱ���,a.���ۼ�,'1' as ����ϵ��,"
                Case Else
                    strUnitQuantity = "B.��װ��λ AS ��λ,(A.���� / B.����ϵ��) AS �깺����,(A.��д���� / B.����ϵ��) AS ��д����,(A.ʵ������ / B.����ϵ��) AS ʵ������,a.�ɱ���*B.����ϵ�� as �ɱ���,a.���ۼ�*B.����ϵ�� as ���ۼ�,B.����ϵ�� as ����ϵ��,"
            End Select
            
            Select Case mint�༭״̬
            Case 7
                    gstrSQL = "" & _
                    "   Select w.����ID,w.���,w.������Ϣ,W.����,w.���,w.ԭ����,w.����,w.��׼�ĺ� ,w.����,w.����,w.ָ�������, " & _
                    "           w.�ⷿ����,w.���Ч��,w.Ч��,w.�������,w.���ʧЧ��, " & _
                    "           w.һ���Բ���,w.���Ч��,w.��λ,w.ԭʼ���� ԭʼ����,w.��д����,w.ʵ������,w.�깺����,w.���ۼ�,w.���۽��,w.����ϵ��, " & _
                    "           (w.���۽�� - Decode(Sign(nvl(z.ʵ�ʽ��,0)),1,w.���۽�� * (nvl(z.ʵ�ʲ��,0) / z.ʵ�ʽ��),w.���۽�� * w.ָ������� / 100)) / decode(w.ʵ������,0,1,w.ʵ������)  �ɱ���, " & _
                    "           (w.���۽�� - Decode(Sign(z.ʵ�ʽ��),1,w.���۽�� * (z.ʵ�ʲ�� / z.ʵ�ʽ��),w.���۽�� * w.ָ������� / 100)) �ɱ����, " & _
                    "           Decode(Sign(z.ʵ�ʽ��),1,w.���۽�� * (z.ʵ�ʲ�� / z.ʵ�ʽ��),w.���۽�� * w.ָ������� / 100) ���, " & _
                    "            w.ժҪ,w.������,w.��������,w.��ҩ��, w.�����, w.�������,w.�ⷿid,w.�Է�����id,W.���ò���,W.������,w.�Ƿ���,w.���÷���,z.��������/w.����ϵ�� as  ��������,z.ʵ�ʽ��,z.ʵ�ʲ��,W.���ٲ���   " & _
                    "    From (  SELECT distinct a.ҩƷid ����id,A.���,('[' || D.���� || ']' || D.����) AS ������Ϣ,  " & _
                    "                    zlSpellCode(D.����) ����,D.���,D.���� as ԭ����,A.����,A.��׼�ĺ�, A.����,a.����,b.ָ�������,b.�ⷿ����,  " & _
                    "                    b.���Ч��,A.Ч��,A.�������,A.���Ч�� as ���ʧЧ��,B.һ���Բ���,b.���Ч��,A.��д���� ԭʼ����, " & strUnitQuantity & _
                    "                    A.�ɱ����,A.���۽��, A.���," & _
                    "                    a.ժҪ,a.������,A.��������,A.��ҩ��, A.�����, A.�������,a.�ⷿid ,D.�Ƿ���,b.���÷���," & _
                    "                    M.ID as �Է�����ID,M.���� as ���ò���,[5] as ������,b.���ٲ��� " & _
                    "            FROM ҩƷ�շ���¼ A, �������� B,�շ���ĿĿ¼ D,(Select ID,���� From ���ű� where id=[4] ) M  " & _
                    "            Where A.ҩƷid = B.����id and a.ҩƷid=D.id   " & _
                    "                    AND A.��¼״̬ =[3]  " & _
                    "                    AND A.���� = 15 AND A.No = [1]  " & _
                    "           ) w  , (  Select ҩƷid ����id,Nvl(����,0) ����,��������,ʵ�ʽ��,ʵ�ʲ��   " & _
                    "                    From ҩƷ��� where �ⷿid=[2]  and ����=1)  z " & _
                    "    Where w.����id=z.����id(+)  and nvl(w.����,0)=nvl(z.����(+),0)   " & _
                    "   ORDER BY " & IIf(strCompare = "0", "���", IIf(strCompare = "1", "������Ϣ", "����")) & IIf(Right(strOrder, 1) = "0", " Asc", " Desc")
                    
            Case 6
                    gstrSQL = "" & _
                    "   Select w.*,z.��������/w.����ϵ�� ��������,nvl(z.ʵ������,0) / w.����ϵ�� As �������,z.ʵ�ʽ��,z.ʵ�ʲ�� " & _
                    "   From (  SELECT distinct a.����id,A.���,('[' || d.���� || ']' || d.����) AS ������Ϣ," & _
                    "                   zlSpellCode(d.����) ����,d.���,d.���� as ԭ����,A.����,A.��׼�ĺ�, A.����,a.����,b.ָ�������,a.Ч��," & _
                                        strUnitQuantity & _
                    "                   a.��д���� ԭʼ����,A.�ɱ����,0 ���۽��,0 ���, " & _
                    "                   a.ժҪ,a.�ⷿid,a.�Է�����id,c.���� as ���ò���,a.������,d.�Ƿ���,b.�ⷿ����,b.���÷���," & _
                    "                   b.���ٲ���,a.����ID,a.��ҳID,a.����,a.�Ա�,a.����,a.����,a.ҽ�Ƹ��ʽ,a.��ǰ����ID,a.��ǰ����ID,a.ʹ��ʱ��,a.����,a.���Ч�� " & _
                    "           FROM (  Select min(x.id) as id, sum(x.ʵ������) as ��д����,0 ʵ������,sum(x.�ɱ����) as �ɱ����,x.������,x.ҩƷid ����ID," & _
                    "                           x.���,x.����,x.��׼�ĺ�, x.����,x.Ч��,x.���Ч��,0 as ����,Nvl(x.����,0) ����,x.����,x.�ɱ���,x.���ۼ�,x.ժҪ,x.�ⷿID,x.�Է�����ID,x.������ID," & _
                    "                           max(M.����ID) as ����ID,max(M.��ҳID) as ��ҳID,max(M.����) as ����,max(M.�Ա�) �Ա�,max(M.����) ����,max(M.����) as ����,max(M.ҽ�Ƹ��ʽ) ҽ�Ƹ��ʽ,max(M.��ǰ����ID) ��ǰ����ID,max(M.��ǰ����ID) ��ǰ����ID,max(M.ʹ��ʱ��) ʹ��ʱ��,max(M.���� ) ���� " & _
                    "                   From ҩƷ�շ���¼ x,����������Ϣ M  " & _
                    "                   WHERE x.NO=[1] AND x.����=20 and x.id=M.�շ�ID(+) " & _
                    "                   Group by x.ҩƷID,x.���,x.����,x.��׼�ĺ�,x.����,x.Ч��,x.���Ч��,Nvl(x.����,0),x.����,x.�ɱ���,x.���ۼ�,x.ժҪ,x.�ⷿID,x.�Է�����ID,x.������ID,x.������" & _
                    "                   having sum(x.��д����)<>0 " & _
                    "               ) A, �������� B,�շ���ĿĿ¼ D,���ű� C " & _
                    "           Where A.����id = B.����id  and a.����id=d.id AND a.�Է�����id=c.id " & _
                    "       ) w,(Select  ҩƷid ����id,Nvl(����,0) ����,��������,ʵ������,ʵ�ʽ��,ʵ�ʲ�� " & _
                    "            From ҩƷ��� " & _
                    "            Where �ⷿid=[2] and ����=1)  z " & _
                    "   Where w.����id=z.����id(+) and nvl(w.����,0)=nvl(z.����(+),0) " & _
                    "   ORDER BY " & IIf(strCompare = "0", "���", IIf(strCompare = "1", "������Ϣ", "����")) & IIf(Right(strOrder, 1) = "0", " Asc", " Desc")
            Case Else
                    gstrSQL = "" & _
                    "   Select w.*,z.��������/w.����ϵ�� ��������,z.ʵ�ʽ��,z.ʵ�ʲ�� " & _
                    "   From (  SELECT distinct a.ҩƷid ����id,A.���,('[' || d.���� || ']' ||d.����) AS ������Ϣ," & _
                    "                   zlSpellCode(d.����) ����,d.���,d.���� as ԭ����,A.����,A.��׼�ĺ�, A.����,a.����,b.ָ�������,a.Ч��," & _
                                        strUnitQuantity & _
                    "                   a.��д���� ԭʼ����,A.�ɱ����,A.���۽��, A.���, " & _
                    "                   a.ժҪ,a.������,a.������,a.��������,a.��ҩ�� as �˲���,a.��ҩ���� as �˲�����,a.�����,a.�������,a.���Ч��,a.�ⷿid,a.�Է�����id,c.���� as ���ò���,d.�Ƿ���,b.�ⷿ����,b.���÷��� ," & _
                    "                   b.���ٲ���,M.����ID,M.��ҳID,M.����,M.�Ա�,M.����,M.����,M.ҽ�Ƹ��ʽ,M.��ǰ����ID,M.��ǰ����ID,M.ʹ��ʱ��,M.���� " & _
                    "           FROM ҩƷ�շ���¼ A, �������� B,�շ���ĿĿ¼ D,���ű� C,����������Ϣ M" & _
                    "           Where A.ҩƷid = B.����id and a.ҩƷid=d.id and A.id=M.�շ�ID(+)  " & _
                    "                   AND a.�Է�����id=c.id and A.��¼״̬ =[3]" & _
                    "                   AND A.���� = 20 AND A.No = [1] " & _
                    "           ) w,(   Select  ҩƷid ����id,Nvl(����,0) ����,��������,ʵ�ʽ��,ʵ�ʲ�� " & _
                    "                   From ҩƷ��� where �ⷿid=[2]   and ����=1)  z " & _
                    "   Where w.����id=z.����id(+) and nvl(w.����,0)=nvl(z.����(+),0)" & _
                    " ORDER BY " & IIf(strCompare = "0", "���", IIf(strCompare = "1", "������Ϣ", "����")) & IIf(Right(strOrder, 1) = "0", " Asc", " Desc")
            End Select
            
            
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "�������õ�", IIf(mint�༭״̬ = 7, mstr��ⵥ��, mstr���ݺ�), cboStock.ItemData(cboStock.ListIndex), mint��¼״̬, mlng����ID, mstr������)
            '���˺�:2007/06/10:����10813
            mstrTime_Start = GetBillInfo(20, mstr���ݺ�)
             
            If rsTemp.EOF Then
                mintParallelRecord = 2
                Exit Sub
            End If
            
            Select Case mint�༭״̬
            Case 2, 6, 7
                Txt������ = UserInfo.�û���
                Txt�������� = Format(sys.Currentdate, "yyyy-mm-dd hh:mm:ss")
                If mint�༭״̬ = 6 Then
                    Txt����� = UserInfo.�û���
                    Txt������� = Format(sys.Currentdate, "yyyy-mm-dd hh:mm:ss")
                End If
            Case Else
                Txt������ = rsTemp!������
                Txt�������� = Format(rsTemp!��������, "yyyy-mm-dd hh:mm:ss")
                Txt����� = IIf(IsNull(rsTemp!�����), "", rsTemp!�����)
                Txt������� = IIf(IsNull(rsTemp!�������), "", Format(rsTemp!�������, "yyyy-mm-dd hh:mm:ss"))
                txt�˲��� = IIf(IsNull(rsTemp!�˲���), "", rsTemp!�˲���)
                txt�˲����� = IIf(IsNull(rsTemp!�˲�����), "", Format(rsTemp!�˲�����, "yyyy-mm-dd hh:mm:ss"))
            End Select
            txtժҪ.Text = IIf(IsNull(rsTemp!ժҪ), "", rsTemp!ժҪ)
            
            If (mint�༭״̬ = 2 Or mint�༭״̬ = 3 Or mint�༭״̬ = 5) And Txt����� <> "" Then
                mintParallelRecord = 3
                Exit Sub
            End If
            
            txtDraw.Text = rsTemp!���ò���
            txtDraw.Tag = rsTemp!�Է�����id
            
            txtDrawPerson.Text = zlStr.Nvl(rsTemp!������)
            txtDrawPerson.Tag = zlStr.Nvl(rsTemp!������)
            
            If mint�༭״̬ = 2 Or mint�༭״̬ = 3 Or mint�༭״̬ = 5 Then
                Set mcolUsedCount = New Collection
            End If
            
            intRow = 0
            With mshBill
                Do While Not rsTemp.EOF
                    
                    intRow = intRow + 1
                    .Rows = intRow + 1
                    .TextMatrix(intRow, 0) = rsTemp.Fields(0)
                    .TextMatrix(intRow, mBillCol.C_����) = rsTemp!������Ϣ
                    .TextMatrix(intRow, mBillCol.C_���) = rsTemp!���
                    .TextMatrix(intRow, mBillCol.c_���) = IIf(IsNull(rsTemp!���), "", rsTemp!���)
                    .TextMatrix(intRow, mBillCol.C_����) = IIf(IsNull(rsTemp!����), "", rsTemp!����)
                    .TextMatrix(intRow, mBillCol.C_��׼�ĺ�) = IIf(IsNull(rsTemp!��׼�ĺ�), "", rsTemp!��׼�ĺ�)
                    .TextMatrix(intRow, mBillCol.c_��λ) = rsTemp!��λ
                    .TextMatrix(intRow, mBillCol.c_����) = IIf(IsNull(rsTemp!����), "", rsTemp!����)
                    .TextMatrix(intRow, mBillCol.C_Ч��) = IIf(IsNull(rsTemp!Ч��), "", Format(rsTemp!Ч��, "yyyy-mm-dd"))
                    
                    .TextMatrix(intRow, mBillCol.C_�깺����) = Format(rsTemp!�깺����, mFMT.FM_����)
                    .TextMatrix(intRow, mBillCol.C_��д����) = Format(rsTemp!��д����, mFMT.FM_����)
                    .TextMatrix(intRow, mBillCol.C_ʵ������) = Format(rsTemp!ʵ������, mFMT.FM_����)
                    
                    .TextMatrix(intRow, mBillCol.c_ԭʼ����) = Val(zlStr.Nvl(rsTemp!ԭʼ����))
                    
                    If mint�༭״̬ = 6 Then
                        .TextMatrix(intRow, mBillCol.C_�������) = Format(rsTemp!�������, mFMT.FM_����) 'ֻ�г���ʱ�ż���
                    End If
                    
                    .TextMatrix(intRow, mBillCol.C_�ɹ���) = Format(rsTemp!�ɱ���, mFMT.FM_�ɱ���)
                    .TextMatrix(intRow, mBillCol.C_�ɹ����) = Format(IIf(mint�༭״̬ = 6, 0, rsTemp!�ɱ����), mFMT.FM_���)
                    .TextMatrix(intRow, mBillCol.C_�ۼ�) = Format(rsTemp!���ۼ�, mFMT.FM_���ۼ�)
                    .TextMatrix(intRow, mBillCol.C_�ۼ۽��) = Format(rsTemp!���۽��, mFMT.FM_���)
                    .TextMatrix(intRow, mBillCol.C_���) = Format(rsTemp!���, mFMT.FM_���)
                    .TextMatrix(intRow, mBillCol.c_����) = IIf(IsNull(rsTemp!����), "0", rsTemp!����)
                    .TextMatrix(intRow, mBillCol.c_����ϵ��) = rsTemp!����ϵ��
                    .TextMatrix(intRow, mBillCol.C_ָ�������) = rsTemp!ָ������� & "||" & rsTemp!�Ƿ��� & "||" & rsTemp!���÷���
                    .TextMatrix(intRow, mBillCol.C_��������) = IIf(IsNull(rsTemp!��������), "0", rsTemp!��������)
                    .TextMatrix(intRow, mBillCol.C_ʵ�ʲ��) = IIf(IsNull(rsTemp!ʵ�ʲ��), "0", rsTemp!ʵ�ʲ��)
                    .TextMatrix(intRow, mBillCol.C_ʵ�ʽ��) = IIf(IsNull(rsTemp!ʵ�ʽ��), "0", rsTemp!ʵ�ʽ��)
                    .TextMatrix(intRow, mBillCol.C_���ʧЧ��) = IIf(IsNull(rsTemp!���Ч��), "", Format(rsTemp!���Ч��, "yyyy-mm-dd"))
                    
                    .TextMatrix(intRow, mBillCol.C_���ٱ�־) = zlStr.Nvl(rsTemp!���ٲ���)
                    .TextMatrix(intRow, mBillCol.C_��������) = Check��������(intRow, rsTemp!���÷���, rsTemp!�ⷿ����)
                    If mint�༭״̬ <> 7 Then
                        '����ID|ʹ��ʱ��|����
                        .TextMatrix(intRow, mBillCol.C_������Ϣ) = zlStr.Nvl(rsTemp!����ID) & "|" & IIf(IsNull(rsTemp!ʹ��ʱ��), "", Format(rsTemp!ʹ��ʱ��, "yyyy-mm-dd")) & "|" & zlStr.Nvl(rsTemp!����)
                        .TextMatrix(intRow, mBillCol.C_���ٲ���) = zlStr.Nvl(rsTemp!����)
                    End If
                    
                    If mint�༭״̬ = 2 Or mint�༭״̬ = 3 Or mint�༭״̬ = 5 Then
                        numUseAbleCount = 0
                        For Each vardrug In mcolUsedCount
                            If vardrug(0) = CStr(rsTemp!����ID & IIf(IsNull(rsTemp!����), "0", rsTemp!����)) Then
                                numUseAbleCount = vardrug(1)
                                mcolUsedCount.Remove vardrug(0)
                                Exit For
                            End If
                        Next
                        str���� = rsTemp!����ID & IIf(IsNull(rsTemp!����), "0", rsTemp!����)
                        If mint�༭״̬ = 2 Or mint�༭״̬ = 5 Then
                            strArray = numUseAbleCount + IIf(IsNull(rsTemp!��д����), "0", rsTemp!��д����)
                        Else
                            strArray = numUseAbleCount + IIf(IsNull(rsTemp!ʵ������), "0", rsTemp!ʵ������)
                        End If
                        mcolUsedCount.Add Array(str����, strArray), str����
                    End If
                    rsTemp.MoveNext
                Loop
                .Rows = intRow + 2
            End With
            rsTemp.Close
    End Select
    
    gstrSQL = "Select ��������, ����id, ������� From ��������˵�� Where ����id=[1] And ��������='����'"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, Val(Me.txtDraw.Tag))
    If rsTemp.EOF Then
        gstrSQL = "Select ��������, ����id, ������� From ��������˵�� Where ����id=[1] And ��������='�ٴ�'"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, Val(Me.txtDraw.Tag))
        If rsTemp.EOF = False Then
            cmdDraw.Tag = "�ٴ�"
        Else
            cmdDraw.Tag = ""
        End If
    Else
        cmdDraw.Tag = "����"
    End If
    


    rsTemp.Close
    Call RefreshRowNO(mshBill, mBillCol.C_�к�, 1)
    Call ��ʾ�ϼƽ��
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub initGrid()
    With mshBill
        .Active = True
        .Cols = mBillCols
        
        .MsfObj.FixedCols = 1
        .TextMatrix(0, mBillCol.C_�к�) = ""
        .TextMatrix(0, mBillCol.C_����) = "���������"
        .TextMatrix(0, mBillCol.C_���) = "���"
        .TextMatrix(0, mBillCol.c_���) = "���"
        .TextMatrix(0, mBillCol.C_����) = "����"
        .TextMatrix(0, mBillCol.C_��׼�ĺ�) = "��׼�ĺ�"
        .TextMatrix(0, mBillCol.c_��λ) = "��λ"
        .TextMatrix(0, mBillCol.c_����) = "����"
        .TextMatrix(0, mBillCol.C_Ч��) = "ʧЧ��"
        .TextMatrix(0, mBillCol.C_���ʧЧ��) = "���ʧЧ��"
                
        .TextMatrix(0, mBillCol.C_�깺����) = "�깺����"
        .TextMatrix(0, mBillCol.C_��д����) = IIf(mint�༭״̬ = 6, "����", "��д����")
        .TextMatrix(0, mBillCol.C_ʵ������) = IIf(mint�༭״̬ = 6, "��������", "ʵ������")
        .TextMatrix(0, mBillCol.c_ԭʼ����) = "ԭʼ����"
        
        .TextMatrix(0, mBillCol.C_�ɹ���) = "�ɱ���"
        .TextMatrix(0, mBillCol.C_�ɹ����) = "�ɱ����"
        .TextMatrix(0, mBillCol.C_�ۼ�) = "�ۼ�"
        .TextMatrix(0, mBillCol.C_�ۼ۽��) = "�ۼ۽��"
        .TextMatrix(0, mBillCol.C_���) = "���"
        .TextMatrix(0, mBillCol.C_��������) = "��������"
        .TextMatrix(0, mBillCol.C_ʵ�ʲ��) = "ʵ�ʲ��"
        .TextMatrix(0, mBillCol.C_ʵ�ʽ��) = "ʵ�ʽ��"
        .TextMatrix(0, mBillCol.C_ָ�������) = "ָ�������"
        .TextMatrix(0, mBillCol.c_����ϵ��) = "����ϵ��"
        .TextMatrix(0, mBillCol.c_����) = "����"
         
        .TextMatrix(0, mBillCol.C_���ٱ�־) = "���ٱ�־"
        .TextMatrix(0, mBillCol.C_������Ϣ) = "������Ϣ"
        .TextMatrix(0, mBillCol.C_���ٲ���) = "���ٲ���"
        .TextMatrix(0, mBillCol.C_��������) = "��������"
        .TextMatrix(0, mBillCol.C_�������) = "�������" '���ø�������ʱ����
        .TextMatrix(1, 0) = ""
        .TextMatrix(1, mBillCol.C_�к�) = "1"
        
        .ColWidth(0) = 0
        .ColWidth(mBillCol.C_�к�) = 300
        .ColWidth(mBillCol.C_����) = 2000
        .ColWidth(mBillCol.C_���) = 0
        .ColWidth(mBillCol.c_���) = 900
        .ColWidth(mBillCol.C_����) = 800
        .ColWidth(mBillCol.C_��׼�ĺ�) = 1000
        .ColWidth(mBillCol.c_��λ) = 500
        .ColWidth(mBillCol.c_����) = 800
        .ColWidth(mBillCol.C_Ч��) = 1000
        .ColWidth(mBillCol.C_���ʧЧ��) = 1000
        .ColWidth(mBillCol.C_�깺����) = IIf(mint�༭״̬ = 6, 0, 800)
        .ColWidth(mBillCol.C_��д����) = 800
        .ColWidth(mBillCol.C_ʵ������) = 800
        .ColWidth(mBillCol.c_ԭʼ����) = 0
        .ColWidth(mBillCol.C_�������) = 0
        
        .ColWidth(mBillCol.C_�ɹ���) = IIf(mblnCostView = False, 0, 800)
        .ColWidth(mBillCol.C_�ɹ����) = IIf(mblnCostView = False, 0, 900)
        .ColWidth(mBillCol.C_�ۼ�) = 800
        .ColWidth(mBillCol.C_�ۼ۽��) = 900
        .ColWidth(mBillCol.C_���) = IIf(mblnCostView = False, 0, 800)
        
        .ColWidth(mBillCol.C_��������) = 0
        
        .ColWidth(mBillCol.C_ʵ�ʲ��) = 0
        .ColWidth(mBillCol.C_ʵ�ʽ��) = 0
        .ColWidth(mBillCol.C_ָ�������) = 0
        .ColWidth(mBillCol.c_����ϵ��) = 0
        .ColWidth(mBillCol.c_����) = 0
        .ColWidth(mBillCol.C_������Ϣ) = 0
        .ColWidth(mBillCol.C_���ٱ�־) = 0
        .ColWidth(mBillCol.C_���ٲ���) = 1000
        .ColWidth(mBillCol.C_��������) = 0

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
        .ColData(mBillCol.C_���) = 5
        .ColData(mBillCol.C_����) = 5
        .ColData(mBillCol.C_��׼�ĺ�) = 5
        .ColData(mBillCol.c_��λ) = 5
        .ColData(mBillCol.c_����) = 5
        .ColData(mBillCol.C_Ч��) = 5
        .ColData(mBillCol.C_���ʧЧ��) = 5
        .ColData(mBillCol.C_�깺����) = 5
        .ColData(mBillCol.c_ԭʼ����) = 5

        .ColData(mBillCol.C_������Ϣ) = 0
        .ColData(mBillCol.C_���ٲ���) = 5
        .ColData(mBillCol.C_���ٱ�־) = 5
        .ColData(mBillCol.C_��������) = 5
        .ColData(mBillCol.C_�������) = 5

        If mint�༭״̬ = 1 Or mint�༭״̬ = 2 Then
            txtDraw.Enabled = True
            cmdDraw.Enabled = True
            txtժҪ.Enabled = True
            txtDrawPerson.Enabled = True
            cmdDrawPerson.Enabled = True
            cboStock.Enabled = True

            .ColData(mBillCol.C_����) = 1
            .ColData(mBillCol.C_��д����) = 4
            .ColData(mBillCol.C_ʵ������) = 5
        ElseIf mint�༭״̬ = 3 Or mint�༭״̬ = 5 Or mint�༭״̬ = 6 Then
            cboStock.Enabled = False
            
            txtDraw.Enabled = False
            cmdDraw.Enabled = False
            txtDrawPerson.Enabled = False
            cmdDrawPerson.Enabled = False
            txtժҪ.Enabled = False
            
            .ColData(mBillCol.C_��д����) = 5
            .ColData(mBillCol.C_ʵ������) = 4
        ElseIf mint�༭״̬ = 4 Then
            cboStock.Enabled = False
            
            txtDraw.Enabled = False
            cmdDraw.Enabled = False
            txtDrawPerson.Enabled = False
            cmdDrawPerson.Enabled = False
            txtժҪ.Enabled = False
            
            .ColData(mBillCol.C_��д����) = 5
            .ColData(mBillCol.C_ʵ������) = 5
            
        End If
        
        .ColData(mBillCol.C_�ɹ���) = 5
        .ColData(mBillCol.C_�ɹ����) = 5
        .ColData(mBillCol.C_�ۼ�) = 5
        .ColData(mBillCol.C_�ۼ۽��) = 5
        .ColData(mBillCol.C_���) = 5
        
        
        .ColData(mBillCol.C_��������) = 5
        
        .ColData(mBillCol.C_ʵ�ʲ��) = 5
        .ColData(mBillCol.C_ʵ�ʽ��) = 5
        .ColData(mBillCol.C_ָ�������) = 5
        .ColData(mBillCol.c_����ϵ��) = 5
        .ColData(mBillCol.c_����) = 5
        
        .ColAlignment(mBillCol.C_����) = flexAlignLeftCenter
        .ColAlignment(mBillCol.c_���) = flexAlignLeftCenter
        .ColAlignment(mBillCol.C_����) = flexAlignLeftCenter
        .ColAlignment(mBillCol.C_��׼�ĺ�) = flexAlignLeftCenter
        .ColAlignment(mBillCol.c_��λ) = flexAlignCenterCenter
        .ColAlignment(mBillCol.c_����) = flexAlignLeftCenter
        .ColAlignment(mBillCol.C_Ч��) = flexAlignLeftCenter
        .ColAlignment(mBillCol.C_�깺����) = flexAlignRightCenter
        .ColAlignment(mBillCol.C_��д����) = flexAlignRightCenter
        .ColAlignment(mBillCol.C_ʵ������) = flexAlignRightCenter
        
        .ColAlignment(mBillCol.C_�ɹ���) = flexAlignRightCenter
        .ColAlignment(mBillCol.C_�ɹ����) = flexAlignRightCenter
        .ColAlignment(mBillCol.C_�ۼ�) = flexAlignRightCenter
        .ColAlignment(mBillCol.C_�ۼ۽��) = flexAlignRightCenter
        .ColAlignment(mBillCol.C_���) = flexAlignRightCenter
        .ColAlignment(mBillCol.C_���ʧЧ��) = flexAlignCenterCenter
        .ColAlignment(mBillCol.C_���ٲ���) = flexAlignLeftCenter
        .PrimaryCol = mBillCol.C_����
        .LocateCol = mBillCol.C_����
        If InStr(1, "345", mint�༭״̬) <> 0 Then .ColData(mBillCol.C_����) = 0
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
    With txtNo
        .Left = mshBill.Left + mshBill.Width - .Width
        LblNO.Left = .Left - LblNO.Width - 100
        .Top = LblTitle.Top
        LblNO.Top = .Top
    End With
    
    
    LblStock.Left = mshBill.Left
    cboStock.Left = LblStock.Left + LblStock.Width + 100
    
    cmdDrawPerson.Left = mshBill.Left + mshBill.Width - cmdDraw.Width
    txtDrawPerson.Left = cmdDrawPerson.Left - txtDrawPerson.Width
    lbl������.Left = txtDrawPerson.Left - lbl������.Width '
    
    cmdDraw.Left = lbl������.Left - cmdDraw.Width * 2
    txtDraw.Left = cmdDraw.Left - txtDraw.Width
    LblEnterStock.Left = txtDraw.Left - LblEnterStock.Width - 100
    
    
    With Lbl��������
        .Top = Pic����.Height - 200 - .Height
        .Left = mshBill.Left + 100
    End With
    
    With Txt��������
        .Top = Lbl��������.Top - 80
        .Left = Lbl��������.Left + Lbl��������.Width + 100
    End With
    
    
    With Lbl������
        .Top = Lbl��������.Top - .Height - 140
        .Left = mshBill.Left + 100
    End With
    
    With Txt������
        .Top = Lbl������.Top - 80
        .Left = Lbl������.Left + Lbl������.Width + 100
    End With
    
    With lbl�˲���
        .Top = Lbl������.Top
        .Left = Abs(mshBill.Width - .Width - txt�˲���.Width - 100) / 2
    End With
    With txt�˲���
        .Top = lbl�˲���.Top - 80
        .Left = lbl�˲���.Left + lbl�˲���.Width + 100
    End With
    
    With lbl�˲�����
        .Top = Lbl��������.Top
        .Left = lbl�˲���.Left
    End With
    With txt�˲�����
        .Top = Txt��������.Top
        .Left = txt�˲���.Left
    End With
    
    
    With Txt�������
        .Top = Lbl��������.Top - 80
        .Left = mshBill.Left + mshBill.Width - .Width
    End With
    
    With Lbl�������
        .Top = Lbl��������.Top
        .Left = Txt�������.Left - 100 - .Width
    End With
    
    With Txt�����
        .Top = Lbl������.Top - 80
        .Left = mshBill.Left + mshBill.Width - .Width
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
    
    With cmdRequestDraw
        .Top = cmdHelp.Top
        .Left = cmdHelp.Left + cmdHelp.Width + 100
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
    
    If mint�༭״̬ = 5 Or mint�༭״̬ = 3 Then
        cmdExpend.Visible = True
        cmdExpend.Move CmdSave.Left - cmdExpend.Width - 100, CmdSave.Top
    End If
    
    Call Local���ٲ�����Ϣ
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mshProvider.Visible = True Then
        mshProvider.Visible = False
        txtDraw.SetFocus
        txtDraw.SelLength = Len(txtDraw.Text)
        txtDraw.SelStart = 0
        Cancel = True
        Exit Sub
    End If
    
    If mblnChange = False Or mint�༭״̬ = 4 Or mint�༭״̬ = 3 Or mint�༭״̬ = 5 Then
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

Private Function SaveCheck() As Boolean
    Dim rsTemp As New Recordset
    Dim intRow As Integer
    
    Dim strNo As String
    Dim lng�ⷿID As Long
    Dim lng�Է�����id As Long
    Dim str����� As String
    Dim dat������� As String
    
    Dim int��� As Integer
    Dim lng����ID As Long
    Dim str���� As String
    Dim lng���� As Long
    Dim dbl��д���� As Double
    Dim dblʵ������ As Double
    Dim dbl�ɱ��� As Double
    Dim dbl�ɱ���� As Double
    Dim dbl���۽�� As Double
    Dim dbl��� As Double
    Dim lng������ID As Long
    Dim str���� As String
    Dim strЧ�� As String
    Dim arrSQL As Variant
    Dim n As Long
    
    mblnSave = False
    SaveCheck = False
    arrSQL = Array()
    
    On Error GoTo ErrHandle
    lng�ⷿID = cboStock.ItemData(cboStock.ListIndex)
    lng�Է�����id = txtDraw.Tag
    str����� = UserInfo.�û���
    strNo = txtNo.Tag
    gstrSQL = "" & _
        "   SELECT b.id " & _
        "   FROM ҩƷ�������� a, ҩƷ������ b " & _
        "   Where a.���id = b.ID  AND a.���� = 35 "
    zlDatabase.OpenRecordset rsTemp, gstrSQL, mstrCaption
   
    If rsTemp.EOF Then
        MsgBox "û�������������õ������������������������!", vbExclamation + vbOKOnly, gstrSysName
        Exit Function
    End If
    lng������ID = rsTemp!Id
    rsTemp.Close
    
    dat������� = Format(sys.Currentdate, "yyyy-mm-dd hh:mm:ss")
    With mshBill
        '��ҩƷID˳���������
        recSort.Sort = "ҩƷid,����,���"
        recSort.MoveFirst
        
        For n = 1 To recSort.RecordCount
            intRow = recSort!�к�
'        For intRow = 1 To .Rows - 1
            If .TextMatrix(intRow, 0) <> "" Then
                
                lng����ID = Val(.TextMatrix(intRow, 0))
                str���� = .TextMatrix(intRow, mBillCol.C_����)
                lng���� = Val(.TextMatrix(intRow, mBillCol.c_����))
                dbl��д���� = Round(Val(.TextMatrix(intRow, mBillCol.C_��д����)) * Val(.TextMatrix(intRow, mBillCol.c_����ϵ��)), g_С��λ��.obj_���С��.����С��)
                dblʵ������ = Round(Val(.TextMatrix(intRow, mBillCol.C_ʵ������)) * Val(.TextMatrix(intRow, mBillCol.c_����ϵ��)), g_С��λ��.obj_���С��.����С��)
                
'                If dbl��д���� = dblʵ������ Then
'                    dbl��д���� = Round(Val(.TextMatrix(intRow, mBillCol.c_ԭʼ����)), g_С��λ��.obj_���С��.����С��)
'                    dblʵ������ = dbl��д����
'                End If
                
                dbl�ɱ��� = Round(Val(.TextMatrix(intRow, mBillCol.C_�ɹ���)) / Val(.TextMatrix(intRow, mBillCol.c_����ϵ��)), g_С��λ��.obj_���С��.�ɱ���С��)
                dbl�ɱ���� = Round(Val(.TextMatrix(intRow, mBillCol.C_�ɹ����)), g_С��λ��.obj_���С��.���С��)
                dbl���۽�� = Round(Val(.TextMatrix(intRow, mBillCol.C_�ۼ۽��)), g_С��λ��.obj_���С��.���С��)
                
                dbl��� = Round(Val(.TextMatrix(intRow, mBillCol.C_���)), g_С��λ��.obj_���С��.���С��)
                str���� = .TextMatrix(intRow, mBillCol.c_����)
                strЧ�� = IIf(.TextMatrix(intRow, mBillCol.C_Ч��) = "", "Null", "to_date('" & .TextMatrix(intRow, mBillCol.C_Ч��) & "','yyyy-mm-dd')")
                int��� = Val(.TextMatrix(intRow, mBillCol.C_���))
                         
                'zl_��������_VERIFY( /*NO_IN*/, /*�ⷿID_IN*/, /*�Է�����ID_IN*/,
                    '/*����ID_IN*/, /*����_IN*/, /*����_IN*/, /*��д����_IN*/,
                    '/*ʵ������_IN*/, /*�ɱ���_IN*/, /*�ɱ����_IN*/, /*���۽��_IN*/,
                    '/*���_IN*/, /*������ID_IN*/, /*�����_IN*/, /*�������_IN*/,
                    '/*����_IN*/, /*Ч��_IN*/, /*��˷�ʽ_In*/ );
                    
                gstrSQL = "zl_��������_Verify(" & int��� & ",'" & strNo & "'," & lng�ⷿID & "," & lng�Է�����id & "," & _
                     lng����ID & ",'" & str���� & "'," & lng���� & "," & dbl��д���� & "," & _
                     dblʵ������ & "," & dbl�ɱ��� & "," & dbl�ɱ���� & "," & dbl���۽�� & "," & _
                     dbl��� & "," & lng������ID & ",'" & str����� & "',to_date('" & dat������� & "','yyyy-mm-dd HH24:MI:SS'),'" & _
                     str���� & "'," & strЧ�� & "," & IIf(mint�༭״̬ = 3, 0, 1) & ")"
                     
                
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = CStr(lng����ID) & ";" & vbCrLf & gstrSQL
            End If
            
            recSort.MoveNext
        Next
    End With
    
    If Not ExecuteSql(arrSQL, mstrCaption, False) Then Exit Function
'    If Not ��鵥��(20, txtNO.Tag) Then
'        gcnOracle.RollbackTrans
'        Exit Function
'    End If
    gcnOracle.CommitTrans
    
    SaveCheck = True
    mblnSave = True
    mblnSuccess = True
    mblnChange = False
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function SaveStrike() As Boolean
    Dim �д�_IN As Integer
    Dim ԭ��¼״̬_IN As Integer
    Dim NO_IN As String
    Dim ���_IN As Integer
    Dim ����ID_IN As Long
    Dim ��������_IN As Double
    Dim ������_IN As String
    Dim ��������_IN  As String
    Dim intRow As Integer
    Dim n As Long
    
    SaveStrike = False
    With mshBill
        '����������������С����
        For intRow = 1 To .Rows - 1
            If Val(.TextMatrix(intRow, mBillCol.C_ʵ������)) <> 0 Then
                If Not ��ͬ����(Val(.TextMatrix(intRow, mBillCol.C_��д����)), Val(.TextMatrix(intRow, mBillCol.C_ʵ������))) Then
                    MsgBox "������Ϸ��ĳ�����������" & intRow & "�У���", vbInformation, gstrSysName
                    Exit Function
                End If
            End If
        Next
    
        NO_IN = Trim(txtNo.Tag)
        ������_IN = UserInfo.�û���
        ��������_IN = Format(sys.Currentdate, "yyyy-mm-dd hh:mm:ss")
        ԭ��¼״̬_IN = mint��¼״̬
        
        On Error GoTo ErrHandle
        gcnOracle.BeginTrans
        
        �д�_IN = 0
        
        '��ҩƷID˳���������
        recSort.Sort = "ҩƷid,����,���"
        recSort.MoveFirst
        
        For n = 1 To recSort.RecordCount
            intRow = recSort!�к�
'        For intRow = 1 To .Rows - 1
            If .TextMatrix(intRow, 0) <> "" And Val(.TextMatrix(intRow, mBillCol.C_ʵ������)) <> 0 Then
                �д�_IN = �д�_IN + 1
                
                ����ID_IN = Val(.TextMatrix(intRow, 0))
                ��������_IN = Round(Val(.TextMatrix(intRow, mBillCol.C_ʵ������)) * Val(.TextMatrix(intRow, mBillCol.c_����ϵ��)), g_С��λ��.obj_ɢװС��.����С��)
                If Val(.TextMatrix(intRow, mBillCol.C_ʵ������)) = Val(.TextMatrix(intRow, mBillCol.C_��д����)) Then
                    ��������_IN = Val(.TextMatrix(intRow, mBillCol.c_ԭʼ����))
                End If
                
                
                ���_IN = Val(.TextMatrix(intRow, mBillCol.C_���))
                
                'ZL_��������_STRIKE(/*�д�_IN*/,/*ԭ��¼״̬_IN*/,/*NO_IN*/,/*���_IN*/, /*����ID_IN*/,
                '/*��������_IN*/,/*������_IN*/, /*��������_IN*/);
                gstrSQL = "ZL_��������_STRIKE(" & _
                    �д�_IN & "," & _
                    ԭ��¼״̬_IN & ",'" & _
                    NO_IN & "'," & _
                    ���_IN & "," & _
                    ����ID_IN & "," & _
                    ��������_IN & ",'" & _
                     ������_IN & "',to_date('" & _
                     Format(��������_IN, "yyyy-mm-dd hh:mm:ss") & "','yyyy-mm-dd HH24:MI:SS'))"
                
                Call zlDatabase.ExecuteProcedure(gstrSQL, mstrCaption)
            End If
            
            recSort.MoveNext
        Next
        gcnOracle.CommitTrans
        
        If �д�_IN = 0 Then
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
    Call ErrCenter
    Call SaveErrLog
End Function

Private Sub mshBill_AfterAddRow(Row As Long)
    Call RefreshRowNO(mshBill, mBillCol.C_�к�, Row)
End Sub

Private Sub mshBill_AfterDeleteRow()
    Call RefreshRowNO(mshBill, mBillCol.C_�к�, mshBill.Row)
End Sub

Private Sub mshBill_BeforeAddRow(Row As Long)
    If mshBill.ColData(mBillCol.C_����) = 0 Then
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
            If MsgBox("��ȷʵҪɾ���������ģ�", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                Cancel = True
            End If
        End If
    End With
End Sub
Private Function Get�շ�ID() As Long
    '------------------------------------------------------------------------------------------
    '����:��ȡ��ǰ�е��շ�ID
    '����:
    '����:�շ�ID
    '------------------------------------------------------------------------------------------
    Dim lng����ID As Long, lng��� As Long
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo ErrHandle
    If mint�༭״̬ = 1 Then Exit Function
    With mshBill
        lng����ID = Val(.TextMatrix(.Row, 0))
        lng��� = Val(.TextMatrix(.Row, C_���))
    End With
    gstrSQL = "Select ID From ҩƷ�շ���¼ where ����=20 and  NO=[1] AND (��¼״̬=1 or mod(��¼״̬,3)=0) and ҩƷid=[2] and ���=[3]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, mstr���ݺ�, lng����ID, lng���)
    If rsTemp.EOF = False Then
        Get�շ�ID = Val(zlStr.Nvl(rsTemp!Id))
        Exit Function
    End If
    Get�շ�ID = 0
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

 
Private Sub mshbill_CommandClick()
    Dim RecReturn As Recordset
    Dim lng����ID As Long, lng�շ�ID As Long, lng����id As Long, strʹ��ʱ�� As String, str���� As String, blnEdit As Boolean
    Dim strTemp As String, arrtemp As Variant
    Dim i As Integer
    Dim int����� As Integer
    
    int����� = mshBill.Row
    
    Select Case mshBill.Col
    Case C_���ٲ���
    Case Else
            Set RecReturn = Frm����ѡ����.ShowMe(Me, 2, cboStock.ItemData(cboStock.ListIndex), _
                                cboStock.ItemData(cboStock.ListIndex), cboStock.ItemData(cboStock.ListIndex), _
                                IIf(mint������ȷ���� = 1, True, False), , , , , , , , , , mlngModule, , mstrPrivs, IIf(mint������ȷ���� = 1, True, False), False)
            If RecReturn.RecordCount > 0 Then
                mblnChange = True
                With mshBill
                    RecReturn.MoveFirst
                    For i = 1 To RecReturn.RecordCount
                        If SetColValue(.Row, RecReturn!����ID, "[" & RecReturn!���� & "]" & RecReturn!����, _
                            IIf(IsNull(RecReturn!���), "", RecReturn!���), IIf(IsNull(RecReturn!����), "", RecReturn!����), _
                            IIf(mintUnit = 0, RecReturn!ɢװ��λ, RecReturn!��װ��λ), _
                            IIf(IsNull(RecReturn!�ۼ�), 0, RecReturn!�ۼ�), IIf(IsNull(RecReturn!����), "", RecReturn!����), _
                            IIf(IsNull(RecReturn!Ч��), "", Format(RecReturn!Ч��, "yyyy-MM-dd")), _
                            IIf(IsNull(RecReturn!���ʧЧ��), "", Format(RecReturn!���ʧЧ��, "yyyy-MM-dd")), _
                            IIf(IsNull(RecReturn!��������), "0", RecReturn!��������), _
                            IIf(IsNull(RecReturn!ʵ�ʽ��), "0", RecReturn!ʵ�ʽ��), _
                            IIf(IsNull(RecReturn!ʵ�ʲ��), "0", RecReturn!ʵ�ʲ��), _
                            IIf(IsNull(RecReturn!ָ�������), "0", RecReturn!ָ�������), _
                            IIf(mintUnit = 0, 1, RecReturn!����ϵ��), IIf(IsNull(RecReturn!����), 0, RecReturn!����), RecReturn!ʱ��, RecReturn!�ⷿ����, RecReturn!���÷���, IIf(IsNull(RecReturn!��׼�ĺ�), "", RecReturn!��׼�ĺ�)) Then

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
                
'                    If RecReturn.RecordCount = 1 Then
'
'                        SetColValue .Row, RecReturn!����ID, "[" & RecReturn!���� & "]" & RecReturn!����, _
'                            IIf(IsNull(RecReturn!���), "", RecReturn!���), IIf(IsNull(RecReturn!����), "", RecReturn!����), _
'                            IIf(mintUnit = 0, RecReturn!ɢװ��λ, RecReturn!��װ��λ), _
'                            IIf(IsNull(RecReturn!�ۼ�), 0, RecReturn!�ۼ�), IIf(IsNull(RecReturn!����), "", RecReturn!����), _
'                            IIf(IsNull(RecReturn!Ч��), "", Format(RecReturn!Ч��, "yyyy-MM-dd")), _
'                            IIf(IsNull(RecReturn!���ʧЧ��), "", Format(RecReturn!���ʧЧ��, "yyyy-MM-dd")), _
'                            IIf(IsNull(RecReturn!��������), "0", RecReturn!��������), _
'                            IIf(IsNull(RecReturn!ʵ�ʽ��), "0", RecReturn!ʵ�ʽ��), _
'                            IIf(IsNull(RecReturn!ʵ�ʲ��), "0", RecReturn!ʵ�ʲ��), _
'                            IIf(IsNull(RecReturn!ָ�������), "0", RecReturn!ָ�������), _
'                            IIf(mintUnit = 0, 1, RecReturn!����ϵ��), IIf(IsNull(RecReturn!����), 0, RecReturn!����), RecReturn!ʱ��, RecReturn!�ⷿ����, RecReturn!���÷���, IIf(IsNull(RecReturn!��׼�ĺ�), "", RecReturn!��׼�ĺ�)
'                        .Col = mBillCol.C_��д����
'                    End If
                End With
                RecReturn.Close
            End If
    End Select
End Sub

Private Sub mshbill_EditChange(curText As String)
    mblnChange = True
End Sub

Private Sub mshBill_EditKeyPress(KeyAscii As Integer)
    Dim strKey As String
    Dim intDigit As Integer
    
    With mshBill
        If .Col = mBillCol.C_��д���� Or .Col = mBillCol.C_ʵ������ Then
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
    If Row > 0 Then
        mshBill.SetRowColor CLng(Row), &HFFCECE, True
    End If
    If mblnEnter Then Exit Sub
    
    Call Local���ٲ�����Ϣ
    
    With mshBill
        Call SetInputFormat(.Row)
        If .Row <> .LastRow Then
        End If
        
        Select Case .Col
            Case mBillCol.C_����
                .TxtCheck = False
                .MaxLength = 80
                'ֻ��ҩ���в���ʾ�ϼ���Ϣ�Ϳ����
                Call ��ʾ�ϼƽ��
                Call ��ʾ�����
            Case mBillCol.C_��д����, mBillCol.C_ʵ������
                .TxtCheck = True
                .MaxLength = 16
                .TextMask = "-.1234567890"
        End Select
    End With
End Sub

Private Sub mshbill_KeyDown(KeyCode As Integer, Shift As Integer, Cancel As Boolean)
    Dim strKey As String
    Dim rsDrug As New Recordset
    Dim strUnit As String
    Dim strUnitQuantity As String
    Dim i As Integer
    Dim int����� As Integer
    
    int����� = mshBill.Row
    
    
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
                    
                    Set RecReturn = FrmMulitSel.ShowSelect(Me, 2, cboStock.ItemData(cboStock.ListIndex), _
                                        cboStock.ItemData(cboStock.ListIndex), cboStock.ItemData(cboStock.ListIndex), _
                                        strKey, sngLeft, sngTop, mshBill.MsfObj.CellWidth, mshBill.MsfObj.CellHeight, _
                                        IIf(mint������ȷ���� = 1, True, False), , , , , , , , , mlngModule, , mstrPrivs, IIf(mint������ȷ���� = 1, True, False), False)
                    
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
                                IIf(IsNull(RecReturn!��������), "0", RecReturn!��������), _
                                IIf(IsNull(RecReturn!ʵ�ʽ��), "0", RecReturn!ʵ�ʽ��), _
                                IIf(IsNull(RecReturn!ʵ�ʲ��), "0", RecReturn!ʵ�ʲ��), _
                                IIf(IsNull(RecReturn!ָ�������), "0", RecReturn!ָ�������), _
                                IIf(mintUnit = 0, 1, RecReturn!����ϵ��), IIf(IsNull(RecReturn!����), 0, RecReturn!����), RecReturn!ʱ��, RecReturn!�ⷿ����, RecReturn!���÷���, IIf(IsNull(RecReturn!��׼�ĺ�), "", RecReturn!��׼�ĺ�)) Then
                            
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
                    
'                    If RecReturn.RecordCount = 1 Then
'                        If SetColValue(.Row, RecReturn!����ID, "[" & RecReturn!���� & "]" & RecReturn!����, _
'                                IIf(IsNull(RecReturn!���), "", RecReturn!���), IIf(IsNull(RecReturn!����), "", RecReturn!����), _
'                                IIf(mintUnit = 0, RecReturn!ɢװ��λ, RecReturn!��װ��λ), _
'                                IIf(IsNull(RecReturn!�ۼ�), 0, RecReturn!�ۼ�), IIf(IsNull(RecReturn!����), "", RecReturn!����), _
'                                IIf(IsNull(RecReturn!Ч��), "", Format(RecReturn!Ч��, "yyyy-MM-dd")), _
'                                IIf(IsNull(RecReturn!���ʧЧ��), "", Format(RecReturn!���ʧЧ��, "yyyy-MM-dd")), _
'                                IIf(IsNull(RecReturn!��������), "0", RecReturn!��������), _
'                                IIf(IsNull(RecReturn!ʵ�ʽ��), "0", RecReturn!ʵ�ʽ��), _
'                                IIf(IsNull(RecReturn!ʵ�ʲ��), "0", RecReturn!ʵ�ʲ��), _
'                                IIf(IsNull(RecReturn!ָ�������), "0", RecReturn!ָ�������), _
'                                IIf(mintUnit = 0, 1, RecReturn!����ϵ��), IIf(IsNull(RecReturn!����), 0, RecReturn!����), RecReturn!ʱ��, RecReturn!�ⷿ����, RecReturn!���÷���, IIf(IsNull(RecReturn!��׼�ĺ�), "", RecReturn!��׼�ĺ�)) = False Then
'                            Cancel = True
'                            Exit Sub
'                        End If
'                        .Text = .TextMatrix(.Row, .Col)
'                    Else
'                        Cancel = True
'                    End If
                    Call ��ʾ�����
                End If
            Case mBillCol.C_���ٲ���
                
 
            Case mBillCol.C_��д����, mBillCol.C_ʵ������
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
                        MsgBox "��������Ϊ��,�����䣡", vbInformation + vbOKOnly, gstrSysName
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    End If
                    
                    If mint�༭״̬ = 6 Then
                        If Abs(Val(strKey)) > Abs(Val(.TextMatrix(.Row, mBillCol.C_��д����))) Then
                            MsgBox "�����������ܴ�������������", vbInformation + vbOKOnly, gstrSysName
                            Cancel = True
                            .TxtSetFocus
                            Exit Sub
                        End If
                    End If
                    
                    If .TextMatrix(.Row, 0) = "" Then Exit Sub
                    If Not CompareUsableQuantity(.Row, strKey) Then
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    End If
                    
                    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                    '�ɱ��۵Ĺ�ʽ��     ������=����*�ۼ�
                    '                  ������=������*��ʵ�ʲ��/ʵ�ʽ�
                    '                  if ʵ�ʽ��=0 then  ������=������*ָ�������
                    '                  ���ۣ��ɱ��ۣ�=��������-�����ۣ�/����
                    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                    
                    strKey = Format(strKey, mFMT.FM_����)
                    .Text = strKey
                    
                    If .TextMatrix(.Row, mBillCol.C_�ۼ�) <> "" Then
                        .TextMatrix(.Row, mBillCol.C_�ۼ۽��) = Format(.TextMatrix(.Row, mBillCol.C_�ۼ�) * strKey, mFMT.FM_���)
                    End If
                    
                    If mint�༭״̬ <> 6 Then
'                        Dim dbl��� As Double, dbl���� As Double, dbl�ɱ���� As Double
'                        Call ��֤�����ۼ���(cboStock.ItemData(cboStock.ListIndex), Val(.TextMatrix(.row, 0)), Val(.TextMatrix(.row, mBillCol.c_����)), Val(.TextMatrix(.row, mBillCol.C_����ϵ��)), Val(.TextMatrix(.row, mBillCol.C_ʵ�ʲ��)), Val(.TextMatrix(.row, mBillCol.C_ʵ�ʽ��)), Val(Split(.TextMatrix(.row, mBillCol.C_ָ�������), "||")(0)) / 100, Val(strKey), Val(.TextMatrix(.row, mBillCol.C_�ۼ۽��)), dbl���, dbl����, dbl�ɱ����)
'                        .TextMatrix(.row, mBillCol.C_���) = Format(dbl���, mFMT.FM_���)
                        .TextMatrix(.Row, mBillCol.C_�ɹ���) = Format(Get�ɱ���(Val(.TextMatrix(.Row, 0)), cboStock.ItemData(cboStock.ListIndex), Val(.TextMatrix(.Row, mBillCol.c_����))) * Val(.TextMatrix(.Row, mBillCol.c_����ϵ��)), mFMT.FM_�ɱ���)
'                        .TextMatrix(.row, mBillCol.C_�ɹ����) = Format(dbl�ɱ����, mFMT.FM_���)
'                    Else
'                        .TextMatrix(.row, mBillCol.C_�ɹ����) = Format(Val(.TextMatrix(.row, mBillCol.C_�ɹ���)) * strKey, mFMT.FM_���)
'                        .TextMatrix(.row, mBillCol.C_���) = Format(Val(.TextMatrix(.row, mBillCol.C_�ۼ۽��)) - Val(.TextMatrix(.row, mBillCol.C_�ɹ����)), mFMT.FM_���)
                    End If
                    .TextMatrix(.Row, mBillCol.C_�ɹ����) = Format(Val(.TextMatrix(.Row, mBillCol.C_�ɹ���)) * strKey, mFMT.FM_���)
                    .TextMatrix(.Row, mBillCol.C_���) = Format(Val(.TextMatrix(.Row, mBillCol.C_�ۼ۽��)) - Val(.TextMatrix(.Row, mBillCol.C_�ɹ����)), mFMT.FM_���)
                    
                    If .Col = mBillCol.C_��д���� Then
                        .TextMatrix(.Row, mBillCol.C_ʵ������) = strKey
                    End If
                End If
                ��ʾ�ϼƽ��
        End Select
    End With
End Sub

Private Function Check��������(ByVal intRow As Integer, ByVal int���÷��� As Integer, ByVal int�ⷿ����) As Integer
    '���ܣ������������ڵ�ǰ�ⷿ�Ƿ����
    '����ֵ��1-������0-������
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo ErrHandle
    
    gstrSQL = "Select Distinct 0 " & _
            "From ��������˵�� " & _
            "Where ((�������� Like '���ϲ���') Or (�������� Like '�Ƽ���')) And ����id = [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, cboStock.ItemData(cboStock.ListIndex))
    If rsTemp.RecordCount = 0 Then
        Check�������� = IIf(int�ⷿ���� = 1, 1, 0)
    Else
        Check�������� = IIf(int���÷��� = 1, 1, 0)
    End If

    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

'�Ӳ���������ȡֵ��������Ӧ����
Private Function SetColValue(ByVal intRow As Integer, ByVal lng����ID As Long, _
        ByVal str���� As String, ByVal str��� As String, ByVal str���� As String, _
        ByVal str��λ As String, ByVal num�ۼ� As Double, ByVal str���� As String, _
        ByVal strЧ�� As String, ByVal str���ʧЧ�� As String, ByVal num�������� As Double, ByVal numʵ�ʽ�� As Double, _
        ByVal numʵ�ʲ�� As Double, ByVal numָ������� As Double, _
        ByVal num����ϵ�� As Double, ByVal lng���� As Long, _
        ByVal int�Ƿ��� As Integer, ByVal int�ⷿ���� As Integer, ByVal int���÷��� As Integer, ByVal str��׼�ĺ� As String) As Boolean
    
        Dim intCount As Integer
        Dim intCol As Integer
        Dim dblPrice As Double
        Dim rsTemp As New Recordset
        Dim bln���� As Boolean
        
    On Error GoTo ErrHandle
    SetColValue = False
    If Format(str���ʧЧ��, "yyyy-mm-dd") < Format(sys.Currentdate, "yyyy-mm-dd") And Trim(str���ʧЧ��) <> "" Then
       If MsgBox("���ġ�" & str���� & "(" & lng���� & ")���Ѿ��������ʧЧ��,�Ƿ�Ҫ���ã�", vbYesNo + vbDefaultButton2 + vbQuestion, gstrSysName) <> vbYes Then
            Exit Function
       End If
    End If
    
    With mshBill
        .TextMatrix(intRow, mBillCol.C_��������) = Check��������(intRow, int���÷���, int�ⷿ����)
        
        If int�Ƿ��� = 1 Then
            gstrSQL = "" & _
                "   Select nvl(���ۼ�,0)*" & num����ϵ�� & " as  �����ۼ�,ʵ�ʽ��/ʵ������* " & num����ϵ�� & " as ƽ�����ۼ�" & _
                "   From ҩƷ��� " & _
                "   Where �ⷿid=[1]" & _
                "           and ҩƷid=[2]" & _
                "           and ����=1 and ʵ������>0 and " & _
                "           nvl(����,0)=[3]"
            
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, cboStock.ItemData(cboStock.ListIndex), lng����ID, lng����)
                        
            If rsTemp.EOF Then
                If mint������ȷ���� = 1 Then
                    MsgBox "ʱ������û�п�棬���ܳ��⣬���飡", vbOKOnly, gstrSysName
                    Exit Function
                Else
                    dblPrice = num�ۼ� * num����ϵ��
                End If
            Else
                If Val(.TextMatrix(intRow, mBillCol.C_��������)) = 1 Then
                    dblPrice = rsTemp!�����ۼ�
                Else
                    dblPrice = rsTemp!ƽ�����ۼ�
                End If
            End If
        End If
        
        Dim lngRow As Long
        For lngRow = 1 To .Rows - 1
            If lngRow <> intRow And .TextMatrix(lngRow, 0) <> "" Then
                If .TextMatrix(lngRow, 0) = lng����ID And Val(.TextMatrix(lngRow, mBillCol.c_����)) = lng���� Then
                    If UBound(Split(mstr�ظ�����, "��")) < 3 Then mstr�ظ����� = mstr�ظ����� & .TextMatrix(lngRow, mBillCol.C_����) & "��"  '����¼�����ظ�������
                    'Call MsgBox("�������ϡ�" & .TextMatrix(lngRow, mBillCol.C_����) & "( " & lng���� & ")���Ѿ����ڣ�������ӣ�", vbOKOnly + vbInformation + vbDefaultButton2, gstrSysName)
                    Exit Function
                End If
            End If
        Next
        
        .TextMatrix(intRow, mBillCol.C_�к�) = intRow
        .TextMatrix(intRow, 0) = lng����ID
        .TextMatrix(intRow, mBillCol.C_����) = str����
        .TextMatrix(intRow, mBillCol.c_���) = str���
        .TextMatrix(intRow, mBillCol.C_����) = str����
        .TextMatrix(intRow, mBillCol.C_��׼�ĺ�) = str��׼�ĺ�
        .TextMatrix(intRow, mBillCol.c_��λ) = str��λ
        .TextMatrix(intRow, mBillCol.c_����) = str����
        .TextMatrix(intRow, mBillCol.C_Ч��) = Format(strЧ��, "yyyy-mm-dd")
        .TextMatrix(intRow, mBillCol.C_���ʧЧ��) = Format(str���ʧЧ��, "yyyy-mm-dd")
    
        .TextMatrix(intRow, mBillCol.C_�ۼ�) = Format(num�ۼ� * num����ϵ��, mFMT.FM_���ۼ�)
        .TextMatrix(intRow, mBillCol.C_��������) = Format(num��������, mFMT.FM_����)
        .TextMatrix(intRow, mBillCol.C_ʵ�ʲ��) = numʵ�ʲ��
        .TextMatrix(intRow, mBillCol.C_ʵ�ʽ��) = numʵ�ʽ��
        .TextMatrix(intRow, mBillCol.C_ָ�������) = numָ������� & "||" & int�Ƿ��� & "||" & int���÷���
        .TextMatrix(intRow, mBillCol.c_����ϵ��) = num����ϵ��
        .TextMatrix(intRow, mBillCol.c_����) = lng����
        If int�Ƿ��� = 1 Then .TextMatrix(intRow, mBillCol.C_�ۼ�) = Format(dblPrice, mFMT.FM_���ۼ�)
        
        gstrSQL = "Select ���ٲ��� From �������� where ����id=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, lng����ID)
        If rsTemp.EOF = False Then
            .TextMatrix(intRow, mBillCol.C_���ٱ�־) = zlStr.Nvl(rsTemp!���ٲ���)
        End If
        Call CheckLapse(strЧ��)
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

Private Sub mshBill_KeyPress(KeyAscii As Integer)
    If mshBill.Col = C_���ٲ��� Then
         If KeyAscii = vbKeySpace Or KeyAscii = vbKeyBack Then
             With mshBill
                .TextMatrix(.Row, C_������Ϣ) = ""
                .TextMatrix(.Row, C_���ٲ���) = ""
             End With
         End If
    End If
End Sub
 

Private Sub mshBill_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
        'If button = 1 Then
'            If mshBill.MouseRow = 0 Then
'              Call Local���ٲ�����Ϣ
'            End If
       ' End If
End Sub

Private Sub mshProvider_DblClick()
    mshProvider_KeyDown vbKeyReturn, 0
End Sub

Private Sub mshProvider_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo ErrHandle
    If KeyCode = vbKeyEscape Then
        mshProvider.Visible = False
        txtDraw.SetFocus
        txtDraw.SelStart = 0
        txtDraw.SelLength = Len(txtDraw.Text)
    End If
    
    If KeyCode = vbKeyReturn Then
        txtDraw.Text = mshProvider.TextMatrix(mshProvider.Row, 3)
        txtDraw.Tag = mshProvider.TextMatrix(mshProvider.Row, 0)
        mshProvider.Visible = False
        
        gstrSQL = "Select ��������, ����id, ������� From ��������˵�� Where ����id=[1] And ��������='����'"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, Val(Me.txtDraw.Tag))
        If rsTemp.EOF Then
            gstrSQL = "Select ��������, ����id, ������� From ��������˵�� Where ����id=[1] And ��������='�ٴ�'"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, Val(Me.txtDraw.Tag))
            If rsTemp.EOF = False Then
                cmdDraw.Tag = "�ٴ�"
            Else
                cmdDraw.Tag = ""
            End If
        Else
            cmdDraw.Tag = "����"
        End If
        
        If txtDrawPerson.Enabled Then txtDrawPerson.SetFocus
    End If
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub mshProvider_LostFocus()
    If mshProvider.Visible Then
        mshProvider.Visible = False
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
    ValidData = False
    Dim intLop As Integer
    
    If txtNo.Locked = False Then
        If Trim(txtNo.Text) = "" Then
            ShowMsgBox "���ݺŲ���Ϊ��"
            Exit Function
        End If
        
        If InStr(1, txtNo.Text, "'") <> 0 Then
            ShowMsgBox "���ݺ��в��ܺ��зǷ��ַ�"
            Exit Function
        End If
        If InStr(1, txtNo.Text, ";") <> 0 Then
            ShowMsgBox "���ݺ��в��ܺ��зǷ��ַ�"
            Exit Function
        End If
    End If
    
    If LenB(StrConv(txtNo.Text, vbFromUnicode)) > txtNo.MaxLength Then
        ShowMsgBox "���ݺų���,���������" & CInt(txtNo.MaxLength / 2) & "�����֣���ò�Ҫ���֣���" & txtNo.MaxLength & "���ַ�!"
        txtNo.SetFocus
        Exit Function
    End If
    If InStr(1, txtժҪ.Text, ";") <> 0 Then
        ShowMsgBox "ժҪ�в�������ֺ�"
        If txtժҪ.Enabled Then txtժҪ.SetFocus
        Exit Function
    End If
    
    With mshBill
        If .TextMatrix(1, 0) <> "" Then         '�����з�����
            If Val(txtDraw.Tag) = 0 Then
                If Trim(txtDraw.Text) = "" Then
                    ShowMsgBox "���ò��Ų���Ϊ�գ�"
                    txtDraw.SetFocus
                    Exit Function
                Else
                    ShowMsgBox "û������������ò��ţ�"
                    txtDraw.SetFocus
                    Exit Function
                End If
            End If
            
            If Trim(txtDrawPerson.Tag) = "" Then
                If MsgBox("��δѡ����ص�������,�Ƿ����?", vbYesNo + vbQuestion + vbDefaultButton2, gstrSysName) = vbNo Then
                    If txtDrawPerson.Enabled Then txtDrawPerson.SetFocus
                    Exit Function
                End If
            End If
            
            If LenB(StrConv(txtժҪ.Text, vbFromUnicode)) > txtժҪ.MaxLength Then
                ShowMsgBox "ժҪ����,���������" & CInt(txtժҪ.MaxLength / 2) & "�����ֻ�" & txtժҪ.MaxLength & "���ַ�!"
                txtժҪ.SetFocus
                Exit Function
            End If
        
            For intLop = 1 To .Rows - 1
                If Val(.TextMatrix(intLop, 0)) > 0 And Trim(.TextMatrix(intLop, mBillCol.C_����)) = "" Then
                    MsgBox "��" & intLop & "�����ĵ�����Ϊ���ˣ����飡", vbInformation, gstrSysName
                    mshBill.SetFocus
                    .Row = intLop
                    .MsfObj.TopRow = intLop
                    .Col = mBillCol.C_����
                    Exit Function
                End If
                
                If Trim(.TextMatrix(intLop, mBillCol.C_����)) <> "" Then
                    If Trim(Trim(.TextMatrix(intLop, mBillCol.C_��д����))) = "" Then
                        ShowMsgBox "��" & intLop & "�����ĵ�����Ϊ���ˣ����飡"
                        mshBill.SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        .Col = mBillCol.C_��д����
                        Exit Function
                    End If
                    
                    If Trim(Trim(.TextMatrix(intLop, mBillCol.C_ʵ������))) = "" Then
                        ShowMsgBox "��" & intLop & "�����ĵ�����Ϊ���ˣ����飡"
                        mshBill.SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        .Col = mBillCol.C_ʵ������
                        Exit Function
                    End If
                    If Val(.TextMatrix(intLop, mBillCol.C_��д����)) > 9999999999# Then
                        ShowMsgBox "��" & intLop & "�����ĵ���д�������������ݿ��ܹ������" & vbCrLf & "���Χ9999999999�����飡"
                        mshBill.SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        .Col = mBillCol.C_��д����
                        Exit Function
                    End If
                    
                    If Val(.TextMatrix(intLop, mBillCol.C_ʵ������)) > 9999999999# Then
                        ShowMsgBox "��" & intLop & "�����ĵ�ʵ���������������ݿ��ܹ������" & vbCrLf & "���Χ9999999999�����飡"
                        mshBill.SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        .Col = mBillCol.C_ʵ������
                        Exit Function
                    End If
                    
                    If Val(.TextMatrix(intLop, mBillCol.C_�ɹ����)) > 9999999999999# Then
                        ShowMsgBox "��" & intLop & "�����ĵĳɱ������������ݿ��ܹ������" & vbCrLf & "���Χ9999999999999�����飡"
                        mshBill.SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        .Col = IIf(.ColData(mBillCol.C_��д����) = 4, mBillCol.C_��д����, mBillCol.C_ʵ������)
                        Exit Function
                    End If
                    If Val(.TextMatrix(intLop, mBillCol.C_�ۼ۽��)) > 9999999999999# Then
                        ShowMsgBox "��" & intLop & "�����ĵ��ۼ۽����������ݿ��ܹ������" & vbCrLf & "���Χ9999999999999�����飡"
                        mshBill.SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        .Col = IIf(.ColData(mBillCol.C_��д����) = 4, mBillCol.C_��д����, mBillCol.C_ʵ������)
                        Exit Function
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
    Dim lng������ID As Long
    Dim chrNo As Variant
    Dim lng��� As Long
    Dim lng�ⷿID As Long
    Dim lng���ò���id As Long
    Dim lng����ID As Long
    Dim str���� As String
    Dim lng���� As Long
    Dim str���� As String
    Dim strЧ�� As String
    Dim dbl��д���� As Double
    Dim dbl�ɱ���  As Double
    Dim dbl�ɱ����  As Double
    Dim dbl��� As Double
    Dim dbl���۽�� As Double
    Dim dbl���  As Double
    Dim strժҪ As String
    Dim str������ As String
    Dim str�������� As String
    Dim str����� As String
    Dim datAssessDate As String
    Dim str���Ч�� As String
    Dim rsTemp As New Recordset
    Dim intRow As Integer
    Dim arrtemp As Variant
    Dim dbl�깺���� As Double
    Dim cllProc As Collection
    Dim n As Long
    
    SaveCard = False
    Set cllProc = New Collection
    
    On Error GoTo ErrHandle
     
    '����������������ID����Ҫ���������Ķ�Ҫ����
    gstrSQL = "" & _
        "   SELECT b.id " & _
        "   FROM ҩƷ�������� a, ҩƷ������ b " & _
        "   Where a.���id = b.ID " & _
        "           AND a.���� = 35 " & _
        "           AND b.ϵ�� = -1 " & _
        "           AND ROWNUM < 2"
    
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "ȡ������"
    
    'Call OpenRecordset(rsTemp, "ȡ������")
    If rsTemp.EOF Then
        MsgBox "û�������������õĳ����������������������ã�", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    lng������ID = rsTemp.Fields(0)
    rsTemp.Close
    
    With mshBill
        chrNo = Trim(txtNo)
        lng�ⷿID = cboStock.ItemData(cboStock.ListIndex)
        
        If mint�༭״̬ = 1 Or mint�༭״̬ = 7 Then 'mbln�������� Or
            If chrNo <> "" Then
                If CheckNOExists(73, chrNo) Then Exit Function
            End If
        
            If chrNo = "" Then chrNo = sys.GetNextNo(73, lng�ⷿID)
            If IsNull(chrNo) Then Exit Function
        End If
        txtNo.Tag = chrNo
        
        lng���ò���id = txtDraw.Tag
        strժҪ = Trim(txtժҪ.Text)
        str������ = Txt������
        str�������� = Format(sys.Currentdate, "yyyy-mm-dd hh:mm:ss")
        str����� = Txt�����
        If mint�༭״̬ = 2 Or mint�༭״̬ = 5 Or blnǿ�Ʊ��� = True Then        '�޸�
            gstrSQL = "zl_��������_Delete('" & mstr���ݺ� & "')"
            AddArray cllProc, gstrSQL
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
                lng���� = .TextMatrix(intRow, mBillCol.c_����)
                strЧ�� = IIf(.TextMatrix(intRow, mBillCol.C_Ч��) = "", "", .TextMatrix(intRow, mBillCol.C_Ч��))
                str���Ч�� = IIf(.TextMatrix(intRow, mBillCol.C_���ʧЧ��) = "", "", .TextMatrix(intRow, mBillCol.C_���ʧЧ��))
                dbl��д���� = Round(Val(.TextMatrix(intRow, mBillCol.C_��д����)) * Val(.TextMatrix(intRow, mBillCol.c_����ϵ��)), g_С��λ��.obj_���С��.����С��)
                
                If mint�༭״̬ = 3 Or mint�༭״̬ = 5 Then '������˺����ʱ����������Զ��ֽ����Ҫɾ��ԭʼ���ٲ����µģ��������µ�ʱ�ֽ���ܳ�����д��������ʵ��������������ʱ��Ӧ����ʵ�������жϿ���Ƿ��㹻
                    If Val(.TextMatrix(intRow, mBillCol.C_��д����)) <> Val(.TextMatrix(intRow, mBillCol.C_ʵ������)) Then
                        dbl��д���� = Round(Val(.TextMatrix(intRow, mBillCol.C_ʵ������)) * Val(.TextMatrix(intRow, mBillCol.c_����ϵ��)), g_С��λ��.obj_���С��.����С��)
                    End If
                End If
                
                dbl�깺���� = Round(Val(.TextMatrix(intRow, mBillCol.C_�깺����)) * Val(.TextMatrix(intRow, mBillCol.c_����ϵ��)), g_С��λ��.obj_���С��.����С��)
                
                dbl�ɱ��� = Round(Val(.TextMatrix(intRow, mBillCol.C_�ɹ���)) / Val(.TextMatrix(intRow, mBillCol.c_����ϵ��)), g_С��λ��.obj_���С��.�ɱ���С��)
                dbl�ɱ���� = Round(Val(.TextMatrix(intRow, mBillCol.C_�ɹ����)), g_С��λ��.obj_���С��.���С��)
                dbl��� = Round(Val(.TextMatrix(intRow, mBillCol.C_�ۼ�)) / Val(.TextMatrix(intRow, mBillCol.c_����ϵ��)), g_С��λ��.obj_���С��.���ۼ�С��)
                dbl���۽�� = Round(Val(.TextMatrix(intRow, mBillCol.C_�ۼ۽��)), g_С��λ��.obj_���С��.���С��)
                dbl��� = Round(Val(.TextMatrix(intRow, mBillCol.C_���)), g_С��λ��.obj_���С��.���С��)
                arrtemp = Split(.TextMatrix(intRow, mBillCol.C_������Ϣ) & "||", "|")
                
                lng��� = intRow
                
                'Zl_��������_Insert
                gstrSQL = "zl_��������_INSERT("
                '  ������id_In In ҩƷ�շ���¼.������id%Type,
                gstrSQL = gstrSQL & "" & lng������ID & ","
                '  No_In         In ҩƷ�շ���¼.NO%Type,
                gstrSQL = gstrSQL & "'" & chrNo & "',"
                '  ���_In       In ҩƷ�շ���¼.���%Type,
                gstrSQL = gstrSQL & "" & lng��� & ","
                '  �ⷿid_In     In ҩƷ�շ���¼.�ⷿid%Type,
                gstrSQL = gstrSQL & "" & lng�ⷿID & ","
                '  �Է�����id_In In ҩƷ�շ���¼.�Է�����id%Type,
                gstrSQL = gstrSQL & "" & lng���ò���id & ","
                '  ����id_In     In ҩƷ�շ���¼.ҩƷid%Type,
                gstrSQL = gstrSQL & "" & lng����ID & ","
                '  ����_In       In ҩƷ�շ���¼.����%Type,
                gstrSQL = gstrSQL & "" & lng���� & ","
                '  ��д����_In   In ҩƷ�շ���¼.��д����%Type,
                gstrSQL = gstrSQL & "" & dbl��д���� & ","
                '  �ɱ���_In     In ҩƷ�շ���¼.�ɱ���%Type,
                gstrSQL = gstrSQL & "" & dbl�ɱ��� & ","
                '  �ɱ����_In   In ҩƷ�շ���¼.�ɱ����%Type,
                gstrSQL = gstrSQL & "" & dbl�ɱ���� & ","
                '  ���ۼ�_In     In ҩƷ�շ���¼.���ۼ�%Type,
                gstrSQL = gstrSQL & "" & dbl��� & ","
                '  ���۽��_In   In ҩƷ�շ���¼.���۽��%Type,
                gstrSQL = gstrSQL & "" & dbl���۽�� & ","
                '  ���_In       In ҩƷ�շ���¼.���%Type,
                gstrSQL = gstrSQL & "" & dbl��� & ","
                '  ������_In     In ҩƷ�շ���¼.������%Type,
                gstrSQL = gstrSQL & "" & IIf(txtDrawPerson.Text = "", "NULL", "'" & txtDrawPerson.Text & "'") & ","
                '  ������_In     In ҩƷ�շ���¼.������%Type,
                gstrSQL = gstrSQL & "'" & str������ & "',"
                '  ��������_In   In ҩƷ�շ���¼.��������%Type,
                gstrSQL = gstrSQL & "to_date('" & str�������� & "','yyyy-mm-dd HH24:MI:SS'),"
                '  ����_In       In ҩƷ�շ���¼.����%Type := Null,
                gstrSQL = gstrSQL & "'" & str���� & "',"
                '  ����_In       In ҩƷ�շ���¼.����%Type := Null,
                gstrSQL = gstrSQL & "'" & str���� & "',"
                '  Ч��_In       In ҩƷ�շ���¼.Ч��%Type := Null,
                gstrSQL = gstrSQL & "" & IIf(strЧ�� = "", "Null", "to_date('" & Format(strЧ��, "yyyy-mm-dd") & "','yyyy-mm-dd')") & ","
                '  ���Ч��_In   In ҩƷ�շ���¼.���Ч��%Type := Null,
                gstrSQL = gstrSQL & "" & IIf(str���Ч�� = "", "Null", "to_date('" & Format(str���Ч��, "yyyy-mm-dd") & "','yyyy-mm-dd')") & ","
                '  ժҪ_In       In ҩƷ�շ���¼.ժҪ%Type := Null
                gstrSQL = gstrSQL & "'" & strժҪ & "',"
                If Val(arrtemp(0)) = 0 Then
                    '  ����id_In     In ����������Ϣ.����id%Type := Null,
                    gstrSQL = gstrSQL & "NULL,"
                    '  ʹ��ʱ��_In   In ����������Ϣ.ʹ��ʱ��%Type := Null,
                    gstrSQL = gstrSQL & "NULL,"
                    '  ����_In       In ����������Ϣ.����%Type := Null
                    gstrSQL = gstrSQL & "NULL,"
                    '   �깺����_in
                    gstrSQL = gstrSQL & dbl�깺���� & ")"
                Else
                    '  ����id_In     In ����������Ϣ.����id%Type := Null,
                    gstrSQL = gstrSQL & "" & Val(arrtemp(0)) & ","
                    '  ʹ��ʱ��_In   In ����������Ϣ.ʹ��ʱ��%Type := Null,
                    gstrSQL = gstrSQL & "" & IIf(Trim(arrtemp(1)) = "", "NULL", "to_date('" & Trim(arrtemp(1)) & "','yyyy-mm-dd')") & ","
                    '  ����_In       In ����������Ϣ.����%Type := Null
                    gstrSQL = gstrSQL & "'" & Trim(arrtemp(2)) & "',"
                    '   �깺����_in
                    gstrSQL = gstrSQL & dbl�깺���� & ")"
                End If
                AddArray cllProc, gstrSQL
            End If
            
            recSort.MoveNext
        Next
    End With
        
                
    Call ExecuteProcedureArrAy(cllProc, mstrCaption, True)
'    If Not ��鵥��(20, txtNO.Tag) Then
'        gcnOracle.RollbackTrans
'        Exit Function
'    End If
    gcnOracle.CommitTrans
    
    mblnSave = True
    mblnSuccess = True
    mblnChange = False
    SaveCard = True
    Exit Function
ErrHandle:
    gcnOracle.RollbackTrans
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function


Private Sub ��ʾ�ϼƽ��()
    Dim curTotal As Double, Cur���ʽ�� As Double, Cur���ʲ�� As Double
    Dim intLop As Integer
    
    curTotal = 0: Cur���ʽ�� = 0: Cur���ʲ�� = 0:
    
    With mshBill
        For intLop = 1 To .Rows - 1
            curTotal = curTotal + Val(.TextMatrix(intLop, mBillCol.C_�ɹ����))
            Cur���ʽ�� = Cur���ʽ�� + Val(.TextMatrix(intLop, mBillCol.C_�ۼ۽��))
        Next
    End With
    
    Cur���ʲ�� = Cur���ʽ�� - curTotal
    lblPurchasePrice.Caption = "�ɱ����ϼƣ�" & Format(curTotal, mFMT.FM_���)
    lblSalePrice.Caption = "�ۼ۽��ϼƣ�" & Format(Cur���ʽ��, mFMT.FM_���)
    lblDifference.Caption = "��ۺϼƣ�" & Format(Cur���ʲ��, mFMT.FM_���)
End Sub

Private Sub ��ʾ�����()
    Dim rsTemp As New Recordset
    
    On Error GoTo ErrHandle
    With mshBill
        If .TextMatrix(.Row, mBillCol.C_����) = "" Then
            stbThis.Panels(2).Text = ""
            Exit Sub
        End If
        If .TextMatrix(mshBill.Row, 0) = "" Then Exit Sub
        
        If Val(.TextMatrix(.Row, mBillCol.c_����)) > 0 Then
            gstrSQL = "" & _
                "   Select ��������/" & .TextMatrix(.Row, mBillCol.c_����ϵ��) & " as  �������� " & _
                "   From ҩƷ��� " & _
                "   Where �ⷿid=[1]" & _
                "           and ҩƷid=[2]" & _
                "           and ����=1 and " & _
                "           nvl(����,0)=[3]"
        Else
            gstrSQL = "Select Sum(Nvl(��������, 0)) / " & .TextMatrix(.Row, mBillCol.c_����ϵ��) & " As �������� " & _
                " From ҩƷ��� Where �ⷿid = [1] And ҩƷid = [2] And ���� = 1 "
        End If
        
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, cboStock.ItemData(cboStock.ListIndex), Val(.TextMatrix(.Row, 0)), Val(.TextMatrix(.Row, mBillCol.c_����)))
        
        If rsTemp.EOF Then
            .TextMatrix(.Row, mBillCol.C_��������) = 0
        Else
            .TextMatrix(.Row, mBillCol.C_��������) = IIf(IsNull(rsTemp.Fields(0)), 0, rsTemp.Fields(0))
        End If
        rsTemp.Close
        
        stbThis.Panels(2).Text = "�����ĵ�ǰ�����Ϊ[" & Format(.TextMatrix(.Row, mBillCol.C_��������), mFMT.FM_����) & "]" & .TextMatrix(.Row, mBillCol.c_��λ)
    End With
    Exit Sub
ErrHandle:
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

Private Sub txtDrawPerson_Change()
    txtDrawPerson.Tag = ""
End Sub

Private Sub txtDrawPerson_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    If Trim(txtDrawPerson.Tag) <> "" Then
        OS.PressKey vbKeyTab
        Exit Sub
    End If
    If Trim(txtDrawPerson.Text) = "" Then
        OS.PressKey vbKeyTab
        Exit Sub
    End If
    If ShowSelect(Trim(txtDrawPerson.Text)) = False Then
        Exit Sub
    End If
    OS.PressKey vbKeyTab
End Sub

Private Sub txtDrawPerson_LostFocus()
    If txtDrawPerson.Tag = "" Then
        If Trim(txtDrawPerson.Text) <> "" Then
            If ShowSelect(Trim(txtDrawPerson.Text)) = False Then
                txtDrawPerson.Text = ""
                Exit Sub
            End If
        End If
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
                "      a.������� Is Not Null And a.No = [1] And a.�ⷿid + 0 = [2]" & GetPriceClassString("E") & vbNewLine & _
                "Order By a.���"

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
                MsgBox "����[" & !ҩƷ���� & "]δ��" & cboStock.Text & "�����ô洢���ԣ����������ã�"
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
                int��װϵ�� = IIf(mintUnit = 0, 1, !����ϵ��)
                If Not SetColValue(intRow, !����ID, "[" & !���� & "]" & !ͨ����, _
                   Nvl(!���), Nvl(!����), _
                   IIf(mintUnit = 0, !���۵�λ, !��װ��λ), _
                    Nvl(!�ּ�, 0), Nvl(!����), _
                    Nvl(!Ч��), IIf(IsNull(!���Ч��), "", Format(!���Ч��, "yyyy-MM-dd")), _
                    Nvl(!��������, 0), Nvl(!ʵ�ʽ��, 0), Nvl(!ʵ�ʲ��, 0), _
                    IIf(IsNull(!ָ�������), "0", !ָ�������), int��װϵ��, Nvl(!����, 0), !ʱ��, _
                    !�ⷿ����, !���÷���, IIf(IsNull(!��׼�ĺ�), "", !��׼�ĺ�)) Then
                    mshBill.ClearBill
                    Exit Sub
                End If

                '��д�������ɹ��ۡ��ۼ۵���
                mshBill.TextMatrix(intRow, mBillCol.C_�к�) = intRow
                mshBill.TextMatrix(intRow, mBillCol.C_��д����) = Format(!ʵ������ / int��װϵ��, mFMT.FM_����)
                mshBill.TextMatrix(intRow, mBillCol.C_ʵ������) = Format(!ʵ������ / int��װϵ��, mFMT.FM_����)
                mshBill.TextMatrix(intRow, mBillCol.C_�ɹ���) = Format(!ƽ���ɱ��� * int��װϵ��, mFMT.FM_�ɱ���)
                mshBill.TextMatrix(intRow, mBillCol.C_�ɹ����) = Format(Val(mshBill.TextMatrix(intRow, mBillCol.C_�ɹ���)) * Val(mshBill.TextMatrix(intRow, mBillCol.C_ʵ������)), mFMT.FM_���)
                mshBill.TextMatrix(intRow, mBillCol.C_�ۼ۽��) = Format(Val(mshBill.TextMatrix(intRow, mBillCol.C_�ۼ�)) * Val(mshBill.TextMatrix(intRow, mBillCol.C_ʵ������)), mFMT.FM_���)
                mshBill.TextMatrix(intRow, mBillCol.C_���) = Format(Val(mshBill.TextMatrix(intRow, mBillCol.C_�ۼ۽��)) - mshBill.TextMatrix(intRow, mBillCol.C_�ɹ����), mFMT.FM_���)

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
    txtժҪ.Tag = ""
End Sub

Private Sub txtժҪ_GotFocus()
    ImeLanguage True
    
    With txtժҪ
        .SelStart = 0
        .SelLength = Len(txtժҪ.Text)
    End With
End Sub

Private Sub txtժҪ_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim strKey As String
    If KeyCode = vbKeyReturn Then
        If mblnHave������; = False Then
            OS.PressKey vbKeyTab: Exit Sub
        End If
        strKey = Trim(txtժҪ)
        If txtժҪ.Tag <> "" Then OS.PressKey vbKeyTab: Exit Sub
        If SelectItem(Me, txtժҪ, strKey, "����������;", "����������;ѡ��", True) = False Then
            OS.PressKey vbKeyTab
            Exit Sub
        End If
        KeyCode = 0
    End If
End Sub

Private Sub txtժҪ_KeyPress(KeyAscii As Integer)
    zlControl.TxtCheckKeyPress txtժҪ, KeyAscii, m�ı�ʽ
    If KeyAscii = Asc(";") Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtժҪ_LostFocus()
    ImeLanguage False
End Sub

'������������бȽ�
Private Function CompareUsableQuantity(ByVal intRow As Integer, ByVal dbl��д���� As Double) As Boolean
    Dim dblUsableQuantity As Double      'ʵ��������Ӧ���������
    Dim dbltotal As Double
    Dim vardrug As Variant, intLop As Integer
    
    'mint�����: 0-�����;1-��飬�������ѣ�2-��飬�����ֹ
    
    CompareUsableQuantity = False
    
    If Not (mint�༭״̬ = 5 Or mint�༭״̬ = 3 Or mint�༭״̬ = 6) Then '����Ǻ˲飬��˻��߳�������������
        If dbl��д���� < 0 And mint������ȷ���� = 0 And Val(mshBill.TextMatrix(intRow, mBillCol.C_��������)) = 1 Then '�������ñ�����ȷ����
            MsgBox "�������ϸ������ñ�����ȷ���Σ��뵽����ϵͳ����->���������������������ã�", vbInformation, gstrSysName
            Exit Function
        End If
        If mint������ȷ���� = 0 Then
            CompareUsableQuantity = True
            Exit Function
        End If
    ElseIf mint�༭״̬ = 6 And dbl��д���� > 0 Then    '���������൱������� ����Ҫ�����
        CompareUsableQuantity = True
        Exit Function
    End If
    
    With mshBill
        If .TextMatrix(intRow, 0) = "" Then Exit Function
        
        If mint�༭״̬ = 6 Then '����ʱֱ�Ӽ��ʵ�ʿ�棬�����ÿ��ô洢
            dblUsableQuantity = Format(.TextMatrix(intRow, mBillCol.C_�������), mFMT.FM_����)
        ElseIf mint�༭״̬ = 2 Then
            If gSystem_Para.para_������¿��ÿ�� = False Then
                '���û��Ԥ��������������������ԭʼ����
                dblUsableQuantity = Val(.TextMatrix(intRow, mBillCol.C_��������))
            Else
                dblUsableQuantity = Val(.TextMatrix(intRow, mBillCol.C_��������)) + Val(.TextMatrix(intRow, mBillCol.c_ԭʼ����)) / Val(.TextMatrix(intRow, mBillCol.c_����ϵ��))
            End If
        ElseIf mint�༭״̬ = 3 Then
            dblUsableQuantity = Val(.TextMatrix(intRow, mBillCol.C_��������)) + Val(.TextMatrix(intRow, mBillCol.c_ԭʼ����)) / Val(.TextMatrix(intRow, mBillCol.c_����ϵ��))
        Else
            dblUsableQuantity = Val(Format(.TextMatrix(intRow, mBillCol.C_��������), mFMT.FM_����))
        End If

        '��ABS�ǿ��ǿ��Ը����������
        If mint����� = 0 Then
            '0-�����
        ElseIf mint����� = 1 Then
            '1-��飬��������
            If IIf(mint�༭״̬ = 6, Abs(dbl��д����), dbl��д����) > dblUsableQuantity Then
                If MsgBox("�������������" & IIf(mint�༭״̬ = 6, Abs(dbl��д����), dbl��д����) & "�������˸����ĵ�" & IIf(mint�༭״̬ = 6, "ʵ��", "����") & "���������" & dblUsableQuantity & "�����Ƿ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
            End If
        ElseIf mint����� = 2 Then
            '2-��飬�����ֹ
            If IIf(mint�༭״̬ = 6, Abs(dbl��д����), dbl��д����) > dblUsableQuantity Then
                MsgBox "�������������" & IIf(mint�༭״̬ = 6, Abs(dbl��д����), dbl��д����) & "�������˸����ĵ�" & IIf(mint�༭״̬ = 6, "ʵ��", "����") & "���������" & dblUsableQuantity & "���������䣡", vbExclamation + vbOKOnly, gstrSysName
                Exit Function
            End If
        End If
    End With
    CompareUsableQuantity = True
End Function

'��ӡ����
Private Sub printbill()
    Dim strNo As String
    strNo = txtNo.Tag
    FrmBillPrint.ShowMe Me, glngSys, "zl1_bill_1717", mint��¼״̬, mintUnit, 1717, "�������õ�", strNo
End Sub
Private Sub txtDraw_Change()
    With txtDraw
        .Text = UCase(.Text)
        .SelStart = Len(.Text)
    End With
    txtDraw.Tag = ""
    txtDrawPerson.Text = ""
    mblnChange = True
End Sub

Private Sub txtDraw_GotFocus()
    txtDraw.SelStart = 0
    txtDraw.SelLength = Len(txtDraw.Text)
End Sub

Private Sub txtDraw_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim strProviderText As String, strվ������ As String
    Dim rsTemp As New Recordset
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    If mint�༭״̬ = 3 Or mint�༭״̬ = 5 Or mint�༭״̬ = 4 Then Exit Sub
    If txtDraw.Tag <> "" Then OS.PressKey vbKeyTab: Exit Sub
    
    strվ������ = GetDeptStationNode(cboStock.ItemData(cboStock.ListIndex))
    On Error GoTo ErrHandle
    With txtDraw
        If Trim(.Text) = "" Then Exit Sub
        strProviderText = GetMatchingSting(UCase(.Text))
        
        gstrSQL = "" & _
            " SELECT a.id,a.����,a.����,a.���� " & _
            " FROM ���ű� a " & _
            " Where ( TO_CHAR (a.����ʱ��, 'yyyy-MM-dd') = '3000-01-01' or a.����ʱ�� is null ) " & _
            IIf(strվ������ <> "", " And a.վ�� = [3] ", "") & _
            "   And (a.���� like [1] Or a.���� like [1] or a.���� like [1])"
        If mbln��ͨ���� Then
            '��ͨ�������죬ֻ��ѡ���Լ������Ŀ���
            '���˺�:20060803
            '����:8468
            gstrSQL = gstrSQL & " And a.ID in (Select ����ID From ������Ա where ��Աid =[2]) "
        End If
            
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, strProviderText, UserInfo.Id, strվ������)
     
        If rsTemp.EOF Then
            MsgBox "û������������ò��ţ������䣡", vbOKOnly + vbInformation, gstrSysName
            KeyCode = 0
            .SelStart = 0
            .SelLength = Len(.Text)
            .Tag = 0
            Exit Sub
        End If
        If rsTemp.RecordCount > 1 Then
            Set mshProvider.Recordset = rsTemp
            Dim intCol As Integer
            Dim intRow As Integer
            
            With mshProvider
                If .Visible = False Then .Visible = True
                .Redraw = False
                .SetFocus
                
                For intRow = 0 To .Rows - 1
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
                
                .Top = txtDraw.Top + txtDraw.Height + 25
                .Left = cmdDraw.Left + cmdDraw.Width - .Width
                .Redraw = True
            End With
            SetObjMuchSelectHeigth Me, mshProvider, txtDraw
            mshProvider.TopRow = 1
            mshProvider.Row = 1
            mshProvider.ColSel = mshProvider.Cols - 1
            mshProvider.SelectionMode = flexSelectionByRow
            Exit Sub
        Else
            .Text = rsTemp!���� & "-" & rsTemp!����
            .Tag = rsTemp!Id
        End If
        
        gstrSQL = "Select ��������, ����id, ������� From ��������˵�� Where ����id=[1] And ��������='����'"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, Val(Me.txtDraw.Tag))
        If rsTemp.EOF Then
            gstrSQL = "Select ��������, ����id, ������� From ��������˵�� Where ����id=[1] And ��������='�ٴ�'"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, Val(Me.txtDraw.Tag))
            If rsTemp.EOF = False Then
                cmdDraw.Tag = "�ٴ�"
            Else
                cmdDraw.Tag = ""
            End If
        Else
            cmdDraw.Tag = "����"
        End If
        
        If txtDrawPerson.Enabled Then txtDrawPerson.SetFocus
    End With
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Sub SetObjMuchSelectHeigth(ByVal frmMain As Object, _
    ByVal mshSel As MSHFlexGrid, _
    ByVal objCtl As Object)
    
    
    '���ö�ѡ�ĸ߶ȺͶ���
    Dim sngHeight As Single
    Dim sngminHeight As Single
    Dim intRow As Long
    Dim intMinRow As Long
    Dim sngTop As Single
    Dim sngFrmMinHeight As Single
    
   
    sngTop = objCtl.Top + objCtl.Height + 25
    intRow = mshSel.Row
    
    mshSel.Row = mshSel.Rows - 1
    sngHeight = ((mshSel.RowHeight(1) + 5) * (mshSel.Rows + 1))
    mshSel.Row = IIf(mshSel.Rows - 1 < 6, mshSel.Row, 6)
    sngminHeight = mshSel.CellTop + mshSel.CellHeight
    sngFrmMinHeight = IIf(frmMain.ScaleHeight - (sngTop) > 0, frmMain.ScaleHeight - sngTop, 0)
       
    If sngHeight > sngFrmMinHeight Then
        If sngFrmMinHeight - sngminHeight < 0 Then
            sngHeight = IIf(sngFrmMinHeight < 2000, 2000, sngFrmMinHeight)
        Else
            sngHeight = sngFrmMinHeight
        End If
        
    ElseIf sngHeight < sngminHeight Then
            sngHeight = sngminHeight
    End If
    mshSel.Height = sngHeight
End Sub
Private Function ShowSelect(ByVal strSeach As String) As Boolean
    '����:�ṩ��������ѡ��
    '����:intSelect:0-������
    
    Dim blnCancel As Boolean, strKey As String, strTittle As String, lngH As Long
    Dim objCtl As Object: Dim vRect As RECT
    Dim rsTemp  As ADODB.Recordset
    'zlDatabase.ShowSelect
    '���ܣ��๦��ѡ����
    '������
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
    '���أ�ȡ��=Nothing,ѡ��=SQLԴ�ĵ��м�¼��
    '˵����
    '     1.ID���ϼ�ID����Ϊ�ַ�������
    '     2.ĩ�����ֶβ�Ҫ����ֵ
    'Ӧ�ã������ڸ������������������Ǻܴ��ѡ����,����ƥ���б�ȡ�
    
    
    Set objCtl = txtDrawPerson
      
    strTittle = "��Աѡ��"
    If strSeach = "" Then
        gstrSQL = "" & _
                "   Select ID, ���,����,���� From ��Ա�� a " & _
                "   Where   exists(select 1 from ������Ա where ��Աid=a.id and ����id=[1]) " & _
                "           and (a.����ʱ�� >= To_Date('3000-01-01', 'YYYY-MM-DD') Or a.����ʱ�� Is Null) and (a.վ��=[3] or a.վ�� is null) " & _
                "   order by ���"
    Else
        gstrSQL = "" & _
                "   Select ID, ���,����,���� From ��Ա�� a " & _
                "   Where ((����) like [2] or  ���  like [2] or  ����  like  [2]) and (a.վ��=[3] or a.վ�� is null) " & _
                "           and exists(select 1 from ������Ա where ��Աid=a.id and ����id=[1]) " & _
                "       and (a.����ʱ�� >= To_Date('3000-01-01', 'YYYY-MM-DD') Or a.����ʱ�� Is Null)" & _
                "   order by ���"
    End If
    
    vRect = zlControl.GetControlRect(objCtl.hwnd)
    lngH = objCtl.Height
    strKey = GetMatchingSting(strSeach)
    Set rsTemp = zlDatabase.ShowSQLSelect(Me, gstrSQL, 0, strTittle, False, "", "", False, False, True, vRect.Left - 15, vRect.Top, lngH, blnCancel, False, False, Val(txtDraw.Tag), strKey, gstrNodeNo)
        
    If blnCancel = True Then
        If objCtl.Enabled Then objCtl.SetFocus
        Exit Function
    End If
    If rsTemp Is Nothing Then
        ShowMsgBox "û������������������,����!"
        If objCtl.Enabled Then objCtl.SetFocus
        Exit Function
    End If
    objCtl.Text = zlStr.Nvl(rsTemp!����)
    objCtl.Tag = zlStr.Nvl(rsTemp!����)
    
    ShowSelect = True
End Function

Private Function Local���ٲ�����Ϣ()
    '--------------------------------------------------------------------------------------
    '����:��λ���ٲ�����Ϣ
    '--------------------------------------------------------------------------------------
    Dim lngTemp As Long, lngPreCol As Long
    Dim i As Long
    
    With mshBill
        If Val(.TextMatrix(.Row, mBillCol.C_���ٱ�־)) = 0 Then
            cmdSel.Visible = False
            Exit Function
        End If
                
        If cmdDraw.Tag <> "�ٴ�" And cmdDraw.Tag <> "����" Then
            If mint�༭״̬ = 1 Or mint�༭״̬ = 2 Then
                cmdSel.Visible = False
                Exit Function
            End If
        End If
        
        If mint�༭״̬ > 2 And mint�༭״̬ <> 7 Then
            If Trim(.TextMatrix(.Row, mBillCol.C_������Ϣ)) = "" Or Trim(.TextMatrix(.Row, mBillCol.C_������Ϣ)) = "||" Then
                cmdSel.Visible = False
                Exit Function
            End If
        End If
        lngPreCol = .Col
        mblnEnter = True
        
        .Redraw = False
        .ColData(mBillCol.C_���ٲ���) = 0
        .Col = mBillCol.C_���ٲ���
        
        lngTemp = .Left
        cmdSel.Left = .Left + .MsfObj.CellLeft + .MsfObj.CellWidth - cmdSel.Width + 30   ' lngTemp - cmdSel.Width + 30 '
        cmdSel.Top = .CellTop + .Top + 15
        cmdSel.Height = .RowHeight(.Row) ' .MsfObj.CellHeight
        .Col = lngPreCol
        If .MsfObj.ColIsVisible(mBillCol.C_���ٲ���) = True And .MsfObj.RowIsVisible(.Row) = True Then
            cmdSel.Visible = True
        Else
            cmdSel.Visible = False
        End If
        
        .Redraw = True
        mblnEnter = False
    End With
End Function

Private Sub SetInputFormat(ByVal intRow As Integer)
    '--------------------------------------------------------------------------------------------------------
    '����:���õ�ǰ�еı༭��ʽ
    '����:introw-��ǰ��
    '����:
    '����:���˺�
    '����:2007/08/21
    '--------------------------------------------------------------------------------------------------------
    
    With mshBill
    
        '1.������2���޸ģ�3�����գ�4���鿴��5���޸ķ�Ʊ��6��������
        If mint�༭״̬ = 1 Or mint�༭״̬ = 2 Then
            If Val(.TextMatrix(intRow, mBillCol.C_���ٱ�־)) = 1 And cmdDraw.Tag = "�ٴ�" Then
                .ColData(mBillCol.C_���ٲ���) = 0
            Else
                .ColData(mBillCol.C_���ٲ���) = 0
            End If
        Else
             .ColData(mBillCol.C_���ٲ���) = 0
        End If
    End With
End Sub

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
                !��� = IIf(Val(mshBill.TextMatrix(n, mBillCol.C_���)) = 0, n, Val(mshBill.TextMatrix(n, mBillCol.C_���)))
                !ҩƷid = Val(mshBill.TextMatrix(n, 0))
                !���� = Val(mshBill.TextMatrix(n, mBillCol.c_����))
                
                .Update
            End If
        Next
        
    End With
End Sub
