VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{AF9744ED-CAFC-4877-8437-2C20C14CEA4E}#9.0#0"; "zlIDKind.ocx"
Begin VB.Form frmҩƷ������ҩNew 
   Caption         =   "ҩƷ������ҩ"
   ClientHeight    =   5310
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   7635
   Icon            =   "frmҩƷ������ҩnew.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5310
   ScaleWidth      =   7635
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.Timer tmrMsgRefresh 
      Interval        =   60000
      Left            =   4200
      Top             =   3840
   End
   Begin VB.Timer tmrCall 
      Interval        =   5000
      Left            =   6240
      Top             =   2880
   End
   Begin VB.Timer TimePrintCancelBill 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   5520
      Top             =   2880
   End
   Begin VB.Timer TimeRefresh 
      Enabled         =   0   'False
      Left            =   4920
      Top             =   2880
   End
   Begin VB.Timer TimePrint 
      Enabled         =   0   'False
      Left            =   4320
      Top             =   2880
   End
   Begin VB.PictureBox picCondition 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   4455
      Left            =   120
      ScaleHeight     =   4455
      ScaleWidth      =   3615
      TabIndex        =   3
      Top             =   240
      Width           =   3615
      Begin VB.PictureBox picList 
         BorderStyle     =   0  'None
         Height          =   1335
         Left            =   120
         ScaleHeight     =   1335
         ScaleWidth      =   2535
         TabIndex        =   15
         Top             =   2760
         Width           =   2535
         Begin XtremeSuiteControls.TabControl tbcList 
            Height          =   975
            Left            =   120
            TabIndex        =   16
            Top             =   120
            Width           =   1455
            _Version        =   589884
            _ExtentX        =   2566
            _ExtentY        =   1720
            _StockProps     =   64
            Enabled         =   -1  'True
         End
      End
      Begin VB.PictureBox picConMain 
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   0  'None
         Height          =   2175
         Left            =   0
         ScaleHeight     =   2175
         ScaleWidth      =   3495
         TabIndex        =   4
         Top             =   120
         Width           =   3495
         Begin VB.TextBox txtPati 
            Appearance      =   0  'Flat
            Height          =   270
            Left            =   960
            TabIndex        =   21
            Top             =   1080
            Width           =   1245
         End
         Begin VB.CheckBox chk��ʾ��ȷ�ϵ��� 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0FF&
            Caption         =   "��ʾ��ȷ�ϵ���"
            Height          =   195
            Left            =   120
            TabIndex        =   20
            Top             =   1920
            Width           =   1695
         End
         Begin VB.CheckBox Chk��ʾ��ҩ�������� 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0FF&
            Caption         =   "��ʾ��ҩ��������"
            Height          =   195
            Left            =   120
            TabIndex        =   17
            Top             =   1920
            Value           =   1  'Checked
            Width           =   3015
         End
         Begin VB.CommandButton cmdFind 
            Height          =   300
            Left            =   2880
            Picture         =   "frmҩƷ������ҩnew.frx":030A
            Style           =   1  'Graphical
            TabIndex        =   10
            ToolTipText     =   "������λ(F2)"
            Top             =   1080
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.ComboBox cbo���� 
            Enabled         =   0   'False
            Height          =   276
            Left            =   960
            TabIndex        =   7
            Top             =   1440
            Width           =   2415
         End
         Begin VB.ComboBox cboʱ�䷶Χ 
            Height          =   300
            Left            =   960
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   40
            Width           =   2415
         End
         Begin VB.CommandButton cmdIC 
            Caption         =   "����"
            Height          =   300
            Left            =   2760
            TabIndex        =   5
            Top             =   1080
            Visible         =   0   'False
            Width           =   615
         End
         Begin MSComCtl2.DTPicker Dtp����ʱ�� 
            Height          =   315
            Left            =   960
            TabIndex        =   8
            Top             =   720
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "yyyy��MM��dd�� HH:mm:ss"
            Format          =   123207683
            CurrentDate     =   36985
         End
         Begin MSComCtl2.DTPicker Dtp��ʼʱ�� 
            Height          =   315
            Left            =   960
            TabIndex        =   9
            Top             =   375
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "yyyy��MM��dd�� HH:mm:ss"
            Format          =   123207683
            CurrentDate     =   36985
         End
         Begin VB.CheckBox Chk��ʾ���̵��� 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0FF&
            Caption         =   "��ʾ���й��̵���"
            Height          =   195
            Left            =   120
            TabIndex        =   18
            TabStop         =   0   'False
            Top             =   1920
            Width           =   3000
         End
         Begin zlIDKind.IDKindNew IDKNType 
            Height          =   375
            Left            =   120
            TabIndex        =   22
            Top             =   1028
            Width           =   795
            _ExtentX        =   1402
            _ExtentY        =   661
            ShowSortName    =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            FontSize        =   9
            FontName        =   "����"
            IDKind          =   -1
            ShowPropertySet =   -1  'True
            AllowAutoICCard =   -1  'True
            BackColor       =   12632319
         End
         Begin VB.Image imgFilter 
            Height          =   240
            Left            =   2400
            Picture         =   "frmҩƷ������ҩnew.frx":0454
            Top             =   1080
            Width           =   240
         End
         Begin VB.Image img���� 
            Height          =   240
            Left            =   600
            Picture         =   "frmҩƷ������ҩnew.frx":6CA6
            ToolTipText     =   "ѡ����"
            Top             =   1470
            Width           =   240
         End
         Begin VB.Label lbl���� 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0FF&
            Caption         =   "����"
            Height          =   180
            Left            =   120
            TabIndex        =   14
            Top             =   1500
            Width           =   360
         End
         Begin VB.Label lblTimeEnd 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0FF&
            Caption         =   "����ʱ��"
            Height          =   180
            Left            =   120
            TabIndex        =   13
            Top             =   787
            Width           =   720
         End
         Begin VB.Label lblTimeBegin 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0FF&
            Caption         =   "��ʼʱ��"
            Height          =   180
            Left            =   120
            TabIndex        =   12
            Top             =   469
            Width           =   720
         End
         Begin VB.Label lblʱ�䷶Χ 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0FF&
            Caption         =   "ʱ�䷶Χ"
            Height          =   180
            Left            =   120
            TabIndex        =   11
            Top             =   110
            Width           =   720
         End
      End
   End
   Begin VB.PictureBox picDetail 
      BorderStyle     =   0  'None
      Height          =   1575
      Left            =   4200
      ScaleHeight     =   1575
      ScaleWidth      =   3015
      TabIndex        =   0
      Top             =   840
      Width           =   3015
      Begin VB.Frame fraLine 
         Height          =   2085
         Left            =   120
         TabIndex        =   1
         Top             =   -120
         Width           =   45
      End
      Begin XtremeSuiteControls.TabControl tbcDetail 
         Height          =   975
         Left            =   360
         TabIndex        =   2
         Top             =   120
         Width           =   1455
         _Version        =   589884
         _ExtentX        =   2566
         _ExtentY        =   1720
         _StockProps     =   64
         Enabled         =   -1  'True
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   19
      Top             =   4950
      Width           =   7635
      _ExtentX        =   13467
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmҩƷ������ҩnew.frx":6DF0
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ��������Ϣ��ҵ��˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   8387
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
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
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   4920
      Top             =   240
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Bindings        =   "frmҩƷ������ҩnew.frx":7684
      Left            =   5520
      Top             =   240
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmҩƷ������ҩNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'''''''''''''''''''��������
Private mlngMode As Long
Private mstrPrivs As String                              'Ȩ�޴�

Private gstrProductName As String
Private mint�ֺ� As Integer
Private mlngIC����id As Long                           'ͨ��IC����ȡ����id

Private Const cstLocate As Integer = 0
Private Const cstFilter As Integer = 1

Private mfrmList As New frm������ҩ�б�
Private mfrmDetail As New frm������ҩ��ϸ
Private mfrmRecipe As New frm����

'���ѿ�
Private mstrCardType As String   '���ѿ�/���п���𣬸�ʽ������|ȫ��|ˢ����־|�����ID|���ų���|ȱʡ��־(1-��ǰȱʡ;0-��ȱʡ)|�Ƿ�����ʻ�(1-�����ʻ�;0-�������ʻ�)|��������(�ڼ�λ���ڼ�λ����,��Ϊ������);��
Private mintCardCount As Integer  '������
Private mint���￨���� As Integer
Private mobjcard As Card

Private mobjSquareCard As Object             'һ��ͨ�ӿ�
Private mobjPlugIn As Object             '��ҽӿڶ���
Private mobjCISJOB As Object  '���Ӳ������Ķ���

Private mstrStockName As String

Private mint���˲�ѯ As Integer                             '�Ƿ�ͨ���˹�������в�ѯ

Private WithEvents mobjIDCard As clsIDCard
Attribute mobjIDCard.VB_VarHelpID = -1
Private WithEvents mobjICCard As clsICCard
Attribute mobjICCard.VB_VarHelpID = -1

'��Ϣ��ض������
Private WithEvents mobjMipModule As zl9ComLib.clsMipModule
Attribute mobjMipModule.VB_VarHelpID = -1
Private mdteMsgRefresh As Date              '�ϴ�ˢ��ʱ��
Private mblnExistMsg As Boolean             '��һ��ʱ������Ƿ��յ���Ϣ

Private mblnCard As Boolean                             '�Ƿ�ˢ���￨
Private mblnScaner As Boolean                           '�Ƿ�ɨ��������
Private mstrScanerLastNo As String                      'ɨ����������ϴ�NO
Private mblnScaned As Boolean                           '�Ѿ�ɨ���һ��
Private mblnFinding As Boolean                          '���Ҷ�λģʽʱ�Ƿ��ҵ�����
Private mintOld����ģʽ As Integer                      'ɨ�����֤֮ǰ�����
Private mblnBrushCard As Boolean                        '�Ƿ�ˢ��
Private mstrLastBrushCardNo As String                   '�ϴ�ˢ��NO

Private mblnStart As Boolean
Private mblnInput As Boolean

Private mdate�ϴ�У��ʱ�� As Date
Private mstr�Զ���ҩ�� As String                        '�����Զ���ҩ������
Private mblnδȡҩ��ҩ As Boolean
Private mstr���� As String

Private mint��ҩ��ʽ As Integer                         '�����û���������ҩģʽ���ǵ�����ҩģʽ������סԺ���ݵ����ʼ�顣

Private mstrChargePrivs As String                        '���ﻮ��Ȩ�޴�
Private mstrStuffPrivs As String                         '���ķ��Ź���Ȩ�޴�

Public RecPart As New ADODB.Recordset                   'ҩ��
Private mrsDrugStock As ADODB.Recordset                 '�洢�ⷿ
Private mrsIsDosage As ADODB.Recordset                  '��ҩ����
Private mrsApplyforcredit As Recordset                  '���ڼ�¼������������ĵ���

Private mblnLoadDrug As Boolean
Private mblnPackerConnect As Boolean    '��ҩ���Ƿ��Ѿ�����
Private mstrOpr As String               '��ҩ����
Private mintAutoSendFlow As Integer     '��ҩ���̿��ƣ�0-���п�ʼ��ҩ���̣�1-�п�ʼ��ҩ��������ҩ����
Private Enum mSendOper                  '��ҩ�������̣�0-��ʼ��ҩ,1-������ҩ
    StartSend = 0
    EndSend = 1
End Enum
Private mblnCompatible As Boolean       '�����Լ�飺true-�������½ӿ�,false-������

Public BlnSetParaSuccess As Boolean                     '���óɹ����
Private BlnRefresh As Boolean
Private IntTimes As Integer                             '���ӳ�
Private BlnInRefresh As Boolean                         '�Ƿ���ˢ��״̬
Private mblnIsFirst As Boolean                          'δУ��

Private mstrDeptNode As String          '��ǰҩ����վ��

Private Type Type_Queue
    blnCallOver As Boolean             '��ǰ�����Ƿ������
    strPCName As String                '����������
    strSendWin As String                '��ǰҩ���ķ�ҩ����
    blnRemoteCall As Boolean             '������Զ�̺��л���
    blnWin As Boolean
End Type
Private mQueue As Type_Queue  '�Ŷӽк�ʹ�õ�һЩ����

Private mbln��������ˢ�� As Boolean

Private mblnStateTimeRefresh As Boolean
Private mblnStateTimePrint As Boolean
Private mblnStateTimeCall As Boolean

Private mrsList As ADODB.Recordset
Private mrsDetail As ADODB.Recordset

Private mstr����Ա As String
Private mstr��ҩ�� As String

Private mstr��������ʾ As String
Private mstr�۸�ʧЧ��ʾ As String

Private mstrPrintRecipe As String                       '���ڷ�ҩ���ӡ����¼���ݺš��������ͣ����ݺ�1,��������1,��¼����1,�����־1,��������1,|���ݺ�2,��������2......

Private mbln���￨ As Boolean                           '�Ƿ��Զ���λ�����￨

Private str��λ�� As String                             '��λ��

Private mstr��� As String

Private mstrBill As String                               '��¼�����ѷ�ҩ�����ż���������

Private mintUnit As Integer                             '��λϵ����1-�ۼ�;2-����;3-סԺ;4-ҩ��

Public intģʽ  As Integer

Private mintTab As Integer

'�Ӳ�������ȡҩƷ�۸����������С��λ��
'Private mintCostDigit As Integer            '�ɱ���С��λ��
'Private mintPriceDigit As Integer           '�ۼ�С��λ��
'Private mintNumberDigit As Integer          '����С��λ��
Private mintMoneyDigit As Integer           '���С��λ��

Private Const mconint�ۼ۵�λ As Integer = 1
Private Const mconint���ﵥλ As Integer = 2
Private Const mconintסԺ��λ As Integer = 3
Private Const mconintҩ�ⵥλ As Integer = 4

Private mstrOracleMoneyForamt As String                 'ORACLE�н���ʽ
Private mstrVBMoneyForamt As String                     'VB�н���ʽ

Private mstrRPTDefaultScheme_Recipt As String           '����ǩ�����Ĭ�ϸ�ʽ
Private mstrRPTScheme_��ҩ�� As String
Private mstrRPTScheme_������ʽ As String

Private mblnSendIsOver As Boolean         '��ҩ�Ƿ������false-δ������true-����

'Ĭ�ϵĴ����С
Private Const mcstlngWinNormalWidth As Long = 13500
Private Const mcstlngWinNormalHeight As Long = 9000

Private mstr�������� As String  '��ʽ����������,��ɫ|��������,��ɫ...
Private mclsComLib As Object
Private mobjDrugMAC As Object

'�б�����
Private Enum mListType
    ��ҩȷ�� = 0
    ����ҩ = 1
    ����ҩ = 2
    ����ҩ = 3
    ��ʱδ�� = 4
    ��ҩ = 5
End Enum

'ʱ�䷶Χ
Private Enum mTimeRange
    ���� = 0
    
    ������ = 1
    ������ = 2
    ָ��ʱ�䷶Χ = 3
End Enum

Private Enum mFindType
    ���ݺ� = 1
    ����� = 2
    ���� = 3
    ���֤ = 4
    IC�� = 5
    ҽ���� = 6
    סԺ�� = 7
End Enum

'Ȩ��
Private Type Type_Privs
    bln����ҩ�� As Boolean
    bln��ҩ As Boolean
    bln��ҩ As Boolean
    bln������ҩ���Ĵ��� As Boolean
    bln���ѽ��ʴ��� As Boolean
    bln���ѽ��ʴ��� As Boolean
    bln���˳�Ժ���˴��� As Boolean
    blnУ�鴦�� As Boolean
    blnҽ����ѯ As Boolean
    bln������ҩ��� As Boolean
    bln���˸������� As Boolean
    bln�޸Ĺ������� As Boolean
    bln�������� As Boolean
    bln������ҩ���Ĵ��� As Boolean
    bln���������� As Boolean
    bln��ҩ As Boolean
    blnֹͣ��ҩ As Boolean
    bln�ָ���ҩ As Boolean
    blnȡҩȷ�� As Boolean
    blnҩƷ�Զ����ӿ� As Boolean
    bln���Ӳ������� As Boolean
    bln�����ѯ����ʱ�䷶Χ���� As Boolean
End Type
Private mPrives As Type_Privs

'ʹ�õ��Ĳ���������ϵͳ�����������������򱾻�ע���
Private Type Type_Params
    '�������е�ϵͳ����
    bln����δ��˴�����ҩ As Boolean
    bln����δ�շѴ�����ҩ As Boolean
    blnҽ������ As Boolean
    int����λ�� As Integer
    bln��˻��۵� As Boolean
    blnˢ����֤ As Boolean
    bln�����������۷��� As Boolean
    intҩƷ������ʾ As Integer          '0-��������ƣ�1-�����룬2-������
    bln��ҩǰ�շѻ���� As Boolean
    bln������     As Boolean          '�Ƿ����ô������ϵͳ

    '�������е���������
    intShowBill�շ� As Integer
    intShowBill���� As Integer
    intShowBill��ҩ As Integer          '0-��ʾ������ҩ��,1-ֻ��ʾδ��ӡ�Ĵ���ҩ����,2-ֻ��ʾ�Ѵ�ӡ�Ĵ���ҩ����
    bln���ʵ� As Boolean
    lngPrintBackInterval As Long
    lngPrintDelay As Long
    int��ʾ�������� As Integer
    lngRefreshInterval As Long
    lngPrintInterval As Long
    intУ�鷢ҩ�� As Integer
    intУ����ҩ�� As Integer
    int�Զ����� As Integer
    bln��ʾ��С��λ As Boolean
    IntShowCol As Integer
    IntAutoPrint As Integer
    intPrint As Integer
    intPrintDrugLable As Integer
    int��ӡ���ķ����嵥 As Integer
    lngҩ��ID As Long
    Str���� As String
    str��ҩ�� As String
    strPrintWindow As String
    bln�Զ���ҩ As Boolean
    int�Զ���ҩʱ�� As Integer
    strSourceDep As String
    int��ҩ���Զ���ӡ As Integer
    int��ѯδ��ҩ�������� As Integer                '���ݲ������õĲ�ѯ��ʹ���ڷ�ҩʱ��ѯ��ǰ����[��ǰҩ����������]��[����ҩ��]��δ��ҩ����
    int��ҩ���Զ���ӡҩƷ��ǩ As Integer
    bln��ҩ��ˢ����֤ As Boolean
    bln��ҩɨ�� As Boolean
    bln��ҩ�շ� As Boolean
    bln��ӡ���и�ʽ As Boolean
    blnPreview As Boolean
    
    intOverTime As Integer
    intType As Integer              '0-��ʾ�����סԺ������1-��ʾ���ﴦ����2-��ʾסԺ����
    str����ˢ����ҩ As String       '����ˢ����ҩ�Ŀ����ID
    bln����ʱ����� As Boolean      'ҩƷҽ��������ʱ��(�״�ʱ��)���ˣ�0-����������ʱ����ˣ�1-������ʱ�����
    int�����ʾ As Integer          '�����ʾ��ʽ��0-��ʾӦ�ս��,1-��ʾʵ�ս��,2-��ʾӦ�պ�ʵ�ս��
    blnȡҩȷ�� As Boolean          '�Ƿ����ò���ʵ��ȡҩȷ��ģʽ��0-�����ã�1-����
    bln��ҩ���� As Boolean        '�Ƿ���ò����ڵ�ǰҩ���Ƿ���δ���Ĳ��ϵ���
    blnɨ������ As Boolean        '0-���Զ�����,1-ɨ����Զ���������'
    int�س���ʽ As Integer          'ͨ��¼���ˢ������ʱϵͳ�Զ���ӻس�����ķ�ʽ��0-ϵͳ���Զ��س�,1-��¼��ﵽ��Ŀ�򿨺ų���ʱ�Զ��س�
    
    '�Ŷӽк��漰����
    blnStartQueue As Boolean        '�����Ŷӽк�
    intSoundType As Integer         '�������ͣ�0-ϵͳ������1-΢������
    blnShowQueue As Boolean         '��ʾ�ŶӶ���
    blnStartCall As Boolean         '������������
    intCallType As Integer          '�кŷ�ʽ��0-���ؽкţ�1-Զ�˽к�
    strRemoteCall As String         'Զ�˺���վ��
    intSoundSpeed As Integer        '�����㲥����
    intSoundTimes As Integer        '�������Ŵ���
    lngShowComponent As Long        '��ʾ�豸���
    intCircleTime As Integer        '������ѯʱ��
    blnSign As Boolean              'ǩ������ҩһ�����
    
    
    'ע������
    int���涨λ As Integer
    int�������� As Integer
    int���̵��� As Integer
    int����ģʽ���� As Integer
    int��ȷ�ϵ��� As Integer
    strDefaultPrinter As String   '��ҩ����ǩĬ�ϵĴ�ӡ��
    
    int����ģʽ As Integer
    
    'ҩ���Ƿ���Ҫ��ҩ
    blnMustDosageProcess As Boolean
    
    'ҩ���Ƿ���Ҫ��ҩȷ��
    blnMustDosageOkProcess As Boolean
    
    '�����
    IntCheckStock As Integer
    
    '�û�����Ĵ�����ɫ����ע���ȡ���ַ�������;�ָ�
    strUserRecipeColor As String
    
    '������ɫ����ҩ����ǩ��Ӧ�Ĵ�ӡ���б���;�ָ�
    strPrinters As String
    
    '��ҩ���ʹ���ǩָ���Ĵ�ӡ��ʽ
    str��ҩ��ʽ As String
    str������ʽ As String
    
    '���ú�����ҩPASS
    blnStarPass As Boolean
    
    '�ⷿ��λ
    strUnit As String
    
    intShowName As Integer         'ҩƷ������ʾ��ʽ
    intFont As Integer             '�����
    
    blnDispensing As Boolean        '�����Ƿ�ͬʱ֪ͨ�ӿ�׼����ҩ
End Type
Private mParams As Type_Params

Private Type Type_Condition
    intListType As Integer
    bln���������� As Boolean
    int������� As Integer                  'ҩ���ķ������1-���ﲡ��;2-סԺ����;3-�����סԺ
    int��Ժ��ҩ As Integer
    bln��ʾ���̵��� As Boolean
    bln��ʾ��ҩ�������� As Boolean
    bln��ʾ��ȷ�ϵ��� As Boolean
End Type
Private mcondition As Type_Condition

Private Type Type_mSQLCondition
    lngҩ��ID As Long
    date��ʼ���� As Date
    date�������� As Date
    str��ʼNO As String
    str����NO As String
    str���� As String
    str���￨ As String
    str��ʶ�� As String
    lng����ID As Long
    str������ As String
    str����� As String
    lngҩƷid As Long
    str��ǰNO As String
    str����� As String
    str���֤ As String
    strҽ���� As String
    lngסԺ�� As Long
    intOverTime As Integer
    lng����ID As Long
End Type
Private mSQLCondition As Type_mSQLCondition
Private Function CardConfirm(ByVal rsData As ADODB.Recordset) As Boolean
    '���ѿ�����ȷ�Ͻӿ�
    '�����������ҩ�����Ұ���������ˣ������˶�ε���ˢ�����ѽӿ�
    'ʵ����֮ǰ�ѽ���У�飬����������������Ҫˢ�����ѣ����ֹ��ҩ����������Ӧ�ò������������ˢ������
    '��ʱ�������ִ���ʽ�������Ժ��䶯
    Dim lngCard����ID As Long
    Dim strCardNo As String
        
    On Error GoTo ErrHand
    
    If mParams.bln��ҩǰ�շѻ���� = False Then
        CardConfirm = True
        Exit Function
    End If
    
    If mobjSquareCard Is Nothing Then
        MsgBox "һ��ͨ�������ϣ����ܽ���ˢ�����ѡ�", vbInformation, gstrSysName
        Exit Function
    End If
    
    'ע�⴫��ļ�¼���Ǵ�����ϸ
    '�շѵ���
    rsData.Filter = "��־=1 And ��¼����=1 And ���շ�=0 And ����ID>0 "
    rsData.Sort = "����ID,NO"
     
    Do While Not rsData.EOF
        If lngCard����ID <> rsData!����ID Then
            If strCardNo <> "" Then
                'ˢ������
                If zlfuncCard_Confirm(mobjSquareCard, Me, mlngMode, mstrPrivs, lngCard����ID, mobjcard.�ӿ����, 1, strCardNo) = False Then
                    Exit Function
                End If
            End If
             
            lngCard����ID = rsData!����ID
            strCardNo = rsData!NO
        Else
            If strCardNo = "" Then
                strCardNo = rsData!NO
            ElseIf InStr(1, strCardNo, rsData!NO) = 0 Then
                strCardNo = strCardNo & "," & rsData!NO
            End If
        End If
        rsData.MoveNext
     Loop
     
     If strCardNo <> "" Then
         'ˢ������
         If zlfuncCard_Confirm(mobjSquareCard, Me, mlngMode, mstrPrivs, lngCard����ID, mobjcard.�ӿ����, 1, strCardNo) = False Then
             Exit Function
         End If
     End If
    
    lngCard����ID = 0
    strCardNo = ""
    
    '���˵��ݣ�ֻ�����ﲡ�˽��д���
    rsData.Filter = "��־=1 And ��¼����=2 And ���շ�=0 And ����ID>0 "
    rsData.Sort = "����ID,NO"
    Do While Not rsData.EOF
        If rsData!�����־ = 1 Or rsData!�����־ = 4 Then
            If lngCard����ID <> rsData!����ID Then
                If strCardNo <> "" Then
                    'ˢ������
                    If zlfuncCard_Confirm(mobjSquareCard, Me, mlngMode, mstrPrivs, lngCard����ID, mobjcard.�ӿ����, 2, strCardNo) = False Then
                        Exit Function
                    End If
                    strCardNo = ""
                End If
                
                lngCard����ID = rsData!����ID
                strCardNo = rsData!NO
            Else
                If strCardNo = "" Then
                    strCardNo = rsData!NO
                ElseIf InStr(1, strCardNo, rsData!NO) = 0 Then
                    strCardNo = strCardNo & "," & rsData!NO
                End If
            End If
        End If
        rsData.MoveNext
    Loop
    If strCardNo <> "" Then
        'ˢ������
        If zlfuncCard_Confirm(mobjSquareCard, Me, mlngMode, mstrPrivs, lngCard����ID, mobjcard.�ӿ����, 2, strCardNo) = False Then
            Exit Function
        End If
    End If
    
    CardConfirm = True
    Exit Function
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    CardConfirm = False
End Function



Private Function CheckPati(ByVal rsData As ADODB.Recordset) As Boolean
    '�����ˢ���ѿ�ȷ��ģʽ�����ܶಡ��������ҩ
    Dim lng�շѲ���ID As Long
    Dim lng���˲���ID As Long
    Dim blnSend As Boolean
    Const cstMsg As String = "���ܶ������ͬʱ����ˢ������ȷ�ϣ���ȷ����ѡ������ͬһ�����ˣ�"
    
    If mParams.bln��ҩǰ�շѻ���� = False Then
        CheckPati = True
        Exit Function
    End If
    
    blnSend = True
    
    'ˢ��ģʽ����鲡���Ƿ��в�����Ϣ��¼
    rsData.Filter = "���շ�=0 And ����ID=0"
    If Not rsData.EOF Then
        blnSend = False
        CheckPati = False
        MsgBox "δ�շѵĻ��۵�������ҩ�������շѣ�", vbInformation, gstrSysName
        Exit Function
    End If
        
    '����շѵ��Ƿ���ڲ�ͬ�Ĳ���
    rsData.Filter = "��־=1 And ��¼����=1 And ���շ�=0 And ����ID>0"
    rsData.Sort = "����ID,NO"
    Do While Not rsData.EOF
        If lng�շѲ���ID = 0 Then
            lng�շѲ���ID = rsData!����ID
        ElseIf lng�շѲ���ID <> rsData!����ID Then
            blnSend = False
            Exit Do
        End If
        rsData.MoveNext
    Loop
    
    If blnSend = False Then
        MsgBox cstMsg, vbInformation, gstrSysName
        Exit Function
    End If
    
    blnSend = True
    
    '�����˵��Ƿ���ڲ�ͬ�Ĳ���
    rsData.Filter = "��־=1 And ��¼����=2 And ���շ�=0 And ����ID>0"
    rsData.Sort = "����ID,NO"
    Do While Not rsData.EOF
        If rsData!�����־ = 1 Or rsData!�����־ = 4 Then
            If lng���˲���ID = 0 Then
                lng���˲���ID = rsData!����ID
            ElseIf lng���˲���ID <> rsData!����ID Then
                blnSend = False
                Exit Do
            End If
        End If
        rsData.MoveNext
    Loop
        
    If blnSend = False Then
        MsgBox cstMsg, vbInformation, gstrSysName
        Exit Function
    End If
    
    '����շѵ��ͼ��˵��Ƿ�ͬһ������
    If lng�շѲ���ID <> 0 And lng���˲���ID <> 0 And lng�շѲ���ID <> lng���˲���ID Then
        MsgBox cstMsg, vbInformation, gstrSysName
        Exit Function
    End If
    
    CheckPati = True
End Function

Private Sub CloseQueue()
    '�ر�LCD����
    If Not gobjLEDShow Is Nothing Then
        Call gobjLEDShow.zlDrugShowClose
        Set gobjLEDShow = Nothing
    End If
End Sub

Private Sub GetPatiType(ByVal rsList As ADODB.Recordset)
    '���ܵ�ǰ�����еĲ������ͼ���Ӧ��ɫ������״̬����ʾ
    
    If rsList Is Nothing Then Exit Sub
    If rsList.RecordCount = 0 Then Exit Sub
    
    With rsList
        .MoveFirst
        
        Do While Not .EOF
            If InStr(1, "|" & mstr��������, "|" & !�������� & ",") = 0 Then
                mstr�������� = IIf(mstr�������� = "", "", mstr�������� & "|") & !�������� & "," & zldatabase.GetPatiColor(IIf(IsNull(!��������), "", !��������))
            End If
            .MoveNext
        Loop
    End With
    
    
End Sub

Private Sub GetSendWindows(ByVal lngҩ��ID As Long)
    'ȡ��ǰҩ���ķ�ҩ����
    Dim rstemp As ADODB.Recordset
    On Error GoTo errHandle
    gstrSQL = "select ����,����,�ϰ�� from ��ҩ���� where ҩ��id=[1]"
    Set rstemp = zldatabase.OpenSQLRecord(gstrSQL, "GetSendWindows", lngҩ��ID)
    
    mQueue.blnWin = False
    mQueue.strSendWin = ""
    Do While Not rstemp.EOF
        mQueue.strSendWin = IIf(mQueue.strSendWin = "", "", mQueue.strSendWin & ",") & rstemp!����
        If rstemp!�ϰ�� = 1 Then mQueue.blnWin = True
        rstemp.MoveNext
    Loop
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function IsValidMsg(strMsgCode As String, strMsgXml As String) As Boolean
    '����ҵ����������ж��Ƿ�����Ч��Ϣ
    Dim objXML As New zl9ComLib.clsXML
    Dim strCodeNode As String
    Dim strҩ��id As String
    Dim int�������� As Integer
    Dim int�շ�״̬ As Integer
    Dim str��ҩ���� As String
    Dim blnValid As Boolean
    Dim rsMsg As New ADODB.Recordset
    Dim lngParentID As Long
    
'    'ZLHIS_CHARGE_003
'    patient_info ������Ϣ
'    patient_id ����id
'    patient_name ����
'    identity_card ���֤��
'    in_number סԺ��
'    out_number �����
'    charge_bill
'       bill_no ���ݺ���
'       bill_kind �������� 1-�շѵ�;2-���ʵ�
'       drug_window ��ҩ����
'       charge_state �շ�״̬ 1-δ�շ�;2-���շ�
'       charge_time �շ�ʱ��
'       charge_person �շ���Ա
'    bill_item
'       charge_item_id �շ���Ŀid
'       charge_item_kind �շ����
'       execute_dept_id ִ�в���id
'       execute_dept_title ִ�в���
    
'    'ZLHIS_CIS_006
'    patient_info ������Ϣ
'    patient_id ����id
'    patient_name ����
'    in_number סԺ��
'    out_number �����
'    patient_clinic ������Ϣ
'    patient_source ������Դ
'    clinic_id ����id
'    clinic_dept_id �������id
'    clinic_dept_title ��������
'    clinic_room ���ﲡ��
'    clinic_bed ���ﲡ��
'    charge_bill
'        send_serial ��������
'        send_time ����ʱ��
'        send_person ������Ա
'        bill_no ���ݺ���
'        bill_kind �������� 1-�շѵ�;2-���ʵ�
'        charge_state �շ�״̬ 1-δ�շ�;2-���շ�
'        send_order ����ҽ��
'        order_id ҽ��id
'        order_relevant_id ���ID
'        order_info ҽ������
'        order_rate ִ��Ƶ��
'        order_route_id ��ҩ;��id
'        order_route ��ҩ;��
'        order_starttime ��ʼʱ��
'        order_single ����
'        order_total ����
'        order_entrust ҽ������
'        order_item_id Ʒ��id
'        charge_item_kind ҩƷ���
'        charge_item_id ҩƷid
'        execute_dept_id ִ�в���id
    
    On Error GoTo ErrHand
    
    If objXML Is Nothing Then Exit Function

    '��XML�ļ�
    objXML.OpenXMLDocument strMsgXml
    
    '�Ӵ򿪵�XML�ļ���ȡָ���ڵ��ֵ�͵�ǰ�ͻ����������ñȽϣ�����Ϣ���ܰ������NO�������ֻҪ��һ��NO���������ͱ�ʾ��Ч
    '1.�ж�ҩ��id
    If strMsgCode = "ZLHIS_CHARGE_003" Then
        strCodeNode = "bill_item"
    ElseIf strMsgCode = "ZLHIS_CIS_006" Then
        strCodeNode = "charge_bill"
    End If
    If objXML.GetMultiNodeRecord(strCodeNode, rsMsg) = False Then Exit Function
    If rsMsg Is Nothing Then Exit Function
    If rsMsg.RecordCount = 0 Then Exit Function
    
    blnValid = False
    Do While Not rsMsg.EOF
        If rsMsg("node_name").Value = "execute_dept_id" Then
            If Val(rsMsg("node_value").Value) = mSQLCondition.lngҩ��ID Then
                blnValid = True
                Exit Do
            End If
        End If
        rsMsg.MoveNext
    Loop
    If blnValid = False Then Exit Function
     
     
    '    Select Case mParams.intShowBill�շ�
'        Case 0  '����ʾ����
'        Case 1  '��ʾδ�շ�
'        Case 2  '��ʾ���շ�
'        Case 3  '��ʾ���д���
'    End Select
        
'        Select Case mParams.intShowBill����
'        Case 0  '����ʾ����
'        Case 1  '��ʾδ���
'        Case 2  '��ʾ�����
'        Case 3  '��ʾ���д���

    '2.�жϵ������ʺ��շ�/���״̬
    If mParams.intShowBill�շ� = 0 And mParams.intShowBill���� = 0 Then Exit Function
    If mParams.intShowBill�շ� <> 3 Or mParams.intShowBill���� <> 3 Then
        If objXML.GetMultiNodeRecord("charge_bill", rsMsg) = False Then Exit Function
        If rsMsg Is Nothing Then Exit Function
        If rsMsg.RecordCount = 0 Then Exit Function
                
        blnValid = False
        Do While Not rsMsg.EOF
            If zlStr.NVL(rsMsg("parent_id").Value) <> "" Then
                If lngParentID = 0 Then
                    lngParentID = Val(rsMsg("parent_id").Value)
                ElseIf lngParentID <> Val(rsMsg("parent_id").Value) Then
                    'ֻҪ��һ��NO���㵥�����ʺ��շ�/���״̬�ͱ�ʾ��Ч���˳�ѭ��
                    If (mParams.intShowBill�շ� = 1 And int�������� = 1 And int�շ�״̬ = 1) Or _
                        (mParams.intShowBill�շ� = 2 And int�������� = 1 And int�շ�״̬ = 2) Or _
                        (mParams.intShowBill���� = 1 And int�������� = 2 And int�շ�״̬ = 1) Or _
                        (mParams.intShowBill���� = 2 And int�������� = 2 And int�շ�״̬ = 2) Then
                        blnValid = True
                        Exit Do
                    End If
                    
                    lngParentID = Val(rsMsg("parent_id").Value)
                End If
                 
                If rsMsg("node_name").Value = "bill_kind" Then
                    int�������� = Val(rsMsg("node_value").Value)
                ElseIf rsMsg("node_name").Value = "charge_state" Then
                    int�շ�״̬ = Val(rsMsg("node_value").Value)
                End If
            End If
          
            rsMsg.MoveNext
            
            If rsMsg.EOF Then
                'ֻҪ��һ��NO���㵥�����ʺ��շ�/���״̬�ͱ�ʾ��Ч���˳�ѭ��
                If (mParams.intShowBill�շ� = 1 And int�������� = 1 And int�շ�״̬ = 1) Or _
                    (mParams.intShowBill�շ� = 2 And int�������� = 1 And int�շ�״̬ = 2) Or _
                    (mParams.intShowBill���� = 1 And int�������� = 2 And int�շ�״̬ = 1) Or _
                    (mParams.intShowBill���� = 2 And int�������� = 2 And int�շ�״̬ = 2) Then
                    blnValid = True
                    Exit Do
                End If
            End If
        Loop
        If blnValid = False Then Exit Function
    End If
    
    '3.�жϷ�ҩ����
    If strMsgCode = "ZLHIS_CHARGE_003" Then
        If objXML.GetMultiNodeRecord("charge_bill", rsMsg) = False Then Exit Function
        If rsMsg Is Nothing Then Exit Function
        If rsMsg.RecordCount = 0 Then Exit Function
        
        blnValid = False
        Do While Not rsMsg.EOF
            If rsMsg("node_name").Value = "drug_window" Then
                If InStr(mParams.Str����, rsMsg("node_value").Value) > 0 Or mParams.Str���� = "" Or rsMsg("node_value").Value = "" Then
                    blnValid = True
                    Exit Do
                End If
            End If
            rsMsg.MoveNext
        Loop
        If blnValid = False Then Exit Function
    End If
    
    IsValidMsg = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
End Function

Private Sub ShowQueue()
    On Error GoTo errHandle
    
    '��ʾ�ŶӶ���
    If mParams.blnShowQueue = True And mParams.blnStartQueue = True Then
        If gobjLEDShow Is Nothing Then
            If Not CreateObject_LED(mParams.lngShowComponent) Then Exit Sub
        End If
        
        If Not gobjLEDShow Is Nothing Then
            '�ر�LCD����
            Call gobjLEDShow.zlDrugShowClose
            Call gobjLEDShow.zlDrugShow(mParams.lngҩ��ID, mParams.Str����, mParams.blnMustDosageProcess, mParams.blnMustDosageOkProcess)
        End If
    Else
        If Not gobjLEDShow Is Nothing Then
            '�ر�LCD����
            Call gobjLEDShow.zlDrugShowClose
            Set gobjLEDShow = Nothing
        End If
    End If
    
    Exit Sub
errHandle:
    Set gobjLEDShow = Nothing
'    If ErrCenter = 1 Then
'        Resume
'    End If
End Sub

Private Sub ShowMedicalRecord(ByVal rsData As ADODB.Recordset)
    '�����ܡ������ĵ�ǰ���˵ĵ��Ӳ���
    
    Dim int���� As Integer
    Dim lng��ҳID As Long
    Dim rstemp As ADODB.Recordset
    
    On Error GoTo errHandle
    
    With rsData
        If Not .EOF Then
            '�жϵ�ǰ����ʱ���ﻹ��סԺ
            If !��¼���� = 1 Or (!��¼���� = 2 And (!�����־ = 1 Or !�����־ = 4)) Then
                int���� = 1
            Else
                int���� = 2
            End If

            '����סԺҽ������վ��ҽ�����͵������շѵ����
            If int���� = 1 And !��Ժ = 1 Then
                int���� = 2
            End If

            '���õ��Ӳ������Ľӿ�
            If Not mobjCISJOB Is Nothing Then
                If int���� = 2 Then
                    lng��ҳID = Val(!��ҳid)
                    
                    '����סԺ����ֱ��ͨ�������շѵķ�ʽ��ҩ(δ����ҽ������),������ҳIDΪ�յ����
                    If lng��ҳID = 0 Then
                        gstrSQL = "Select ��ҳid From ��Ժ���� Where ����id = [1]"
                        
                        Set rstemp = zldatabase.OpenSQLRecord(gstrSQL, "��ѯסԺ������ҳID", !����ID)
                        
                        If Not rstemp.EOF Then lng��ҳID = rstemp!��ҳid
                    End If
                    
                    On Error Resume Next
                    Call mobjCISJOB.ShowArchive(Me, !����ID, lng��ҳID)
                    err.Clear: On Error GoTo 0
                Else
                    '��Ϊ���ﲡ�ˣ���ѯ��Ӧ�ĹҺ�id
                    gstrSQL = "Select ID As �Һ�id From ���˹Һż�¼ Where ����id = [1]"
                    
                    Set rstemp = zldatabase.OpenSQLRecord(gstrSQL, "��ȡ���ﲡ�˹Һ�ID", !����ID)
                    
                    If Not rstemp.EOF Then
                        On Error Resume Next
                        Call mobjCISJOB.ShowArchive(Me, !����ID, rstemp!�Һ�id)
                        err.Clear: On Error GoTo 0
                    End If
                End If
            End If
        End If
    End With
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub zlCallMain()
    Dim intCount As Integer
    Dim dateStart As Date
    Dim strsql As String
    Dim rstemp As ADODB.Recordset
    Dim strCall As String
    Dim strCallName As String
    Dim strCallWindows As String
    Dim blnCallTime As Boolean

    '���û�����ú��й���,���˳�
    If mParams.blnStartCall = False Then Exit Sub
    
    '�����ȫ��Զ�����������˳�
    If mParams.intCallType = 1 Then Exit Sub
    
    '����ϴκ���δ��ɣ����˳�
    If mQueue.blnCallOver = False Then Exit Sub
    
    '������Ҫһ��ʱ�䣬�ȹر�Timer�ؼ�
    blnCallTime = tmrCall.Enabled
    If blnCallTime = True Then
        tmrCall.Enabled = False
    End If
    
    mQueue.blnCallOver = False
    
    On Error GoTo errHandle
        
    '��ȡ��ǰ���еĵ���
    If mQueue.blnRemoteCall = True Then
        'Զ�˺���ģʽʱ�������кŴ���
        strsql = "Select /*+ Rule*/ Distinct a.����, a.NO, a.�ⷿid, a.��ҩ����, a.�������� " & _
            " From δ��ҩƷ��¼ A, Table(Cast(f_Str2list([2]) As Zltools.t_Strlist)) B " & _
            " Where a.��ҩ���� = b.Column_Value And (a.���� = 8 or a.����=9) And a.�ⷿid = [1] And a.�Ŷ�״̬ = 3 And a.�������� Is Not Null "
        strCallWindows = mQueue.strSendWin
    Else
        '���ؽк�ģʽʱֻ����һ���кŴ���
        strsql = "Select /*+ Rule*/ Distinct a.����, a.NO, a.�ⷿid, a.��ҩ����, a.�������� " & _
            " From δ��ҩƷ��¼ A, Table(Cast(f_Str2list([2]) As Zltools.t_Strlist)) B " & _
            " Where a.��ҩ���� = b.Column_Value And (a.���� = 8 or a.����=9) And a.�ⷿid = [1] And a.�Ŷ�״̬ = 3 And a.�������� Is Not Null "
        strCallWindows = mParams.Str����
    End If
    Set rstemp = zldatabase.OpenSQLRecord(strsql, "zlCallMain", mParams.lngҩ��ID, strCallWindows)

    While rstemp.EOF = False
        DoEvents
        FS.ShowFlash "���ں���...", Me
        
        strCall = rstemp!��������
        
        '�����д�������1ʱѭ������
        intCount = 0
        While intCount < mParams.intSoundTimes
            If mParams.intSoundType = CALLSOUND_MS Then
                '΢������
                Call zlCall_MsSoundPlay(strCall, mParams.intSoundSpeed)
            Else
                'ϵͳ����
                Call zlCall_SystemSoundPlay(strCall, mParams.intSoundSpeed)
            End If

            intCount = intCount + 1
                                         
            If mParams.intSoundTimes > 1 Then
                DoEvents
                Call Sleep(3)
            End If
        Wend

        '�����������������ݣ�����ˢ����ʾ�Ĵ������
        gstrSQL = "Zl_δ��ҩƷ��¼_����("
            'NO
            gstrSQL = gstrSQL & "'" & rstemp!NO & "'"
            '����
            gstrSQL = gstrSQL & "," & rstemp!����
            'ҩ��id
            gstrSQL = gstrSQL & "," & rstemp!�ⷿid
            '��ҩ����
            gstrSQL = gstrSQL & ",'" & rstemp!��ҩ���� & "'"
            '��������
            gstrSQL = gstrSQL & ",Null"
            gstrSQL = gstrSQL & ")"
        Call zldatabase.ExecuteProcedure(gstrSQL, "tmrCall_Timer")
        
        rstemp.MoveNext
    Wend
    
    DoEvents
    FS.StopFlash
    DoEvents

    '���¿���ʱ��ؼ�
    If blnCallTime = True Then
        tmrCall.Interval = mParams.intCircleTime * 1000
        tmrCall.Enabled = True
    End If
    
    mQueue.blnCallOver = True
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    
    '���¿���ʱ��ؼ�
    If blnCallTime = True Then
        tmrCall.Interval = mParams.intCircleTime * 1000
        tmrCall.Enabled = True
    End If
    
    mQueue.blnCallOver = True
End Sub

Private Sub cbo����_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim str�������� As String
    
    str�������� = "F"
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    If cbo����.ListCount = 0 Then Exit Sub
    
    If cbo����.ListIndex >= 0 Then
        If Val(cbo����.Tag) = cbo����.ItemData(cbo����.ListIndex) Then
            Exit Sub
        End If
    End If
    
    If Select����ѡ����(Me, cbo����, Trim(cbo����.Text), str��������, , "2,3") = False Then
        Exit Sub
    End If
    If cbo����.ListIndex >= 0 Then
        cbo����.Tag = cbo����.ItemData(cbo����.ListIndex)
    End If
End Sub

Private Sub cbo����_KeyPress(KeyAscii As Integer)
    '�������뵥����
    If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub cbo����_LostFocus()
    If cbo����.ListIndex = -1 Then
        cbo����.Text = ""
    End If
End Sub

Private Sub cbo����_Validate(Cancel As Boolean)
    If cbo����.ListIndex = -1 Then
        cbo����.Text = ""
    End If
End Sub

Private Sub chk��ʾ��ȷ�ϵ���_Click()
    Call SaveSetting("ZLSOFT", "����ģ��\����\" & App.ProductName & "\" & "ҩƷ������ҩ", "��ʾ��ȷ�ϵ���", chk��ʾ��ȷ�ϵ���.Value)
    RefreshList mcondition.intListType
End Sub

Private Sub mobjICCard_ShowICCardInfo(ByVal strNo As String)
    If Not txtPati.Locked And txtPati.Text = "" And Me.ActiveControl Is txtPati And strNo <> "" Then
        txtPati.Text = strNo
        
        If txtPati.Text = "" Then
            Call mobjICCard.SetEnabled(False)
        Else
            mParams.int����ģʽ = mFindType.IC��
'            Call SetInputState(mParams.int����ģʽ)
            
            DoEvents
            
            Call txtPati_KeyPress(vbKeyReturn)
        End If
    End If
End Sub
Private Sub BillPrint_Back()
    '��ӡ�˷ѵ���
    
    Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1341_8", Me, "ҩ��=" & mParams.lngҩ��ID)
End Sub

Private Sub BillPrint_Custom(ByVal Control As XtremeCommandBars.ICommandBarControl)
    '��ӡ�Զ��屨��
    
    'Ĭ�ϲ�����ҩƷ=ҩƷid��ҩ��=ҩ��id��NO=����NO����������=ҩƷ�շ���¼.���ݣ�����ID=����ID
    Dim lng����ID As Long
    Dim rsTmp As ADODB.Recordset
    Dim strName As String
    Dim str��ǰ���� As String
    Dim Int���� As Integer, strNo As String
    
    strName = Split(Control.Parameter, ",")(1)
    
    If strName = "ZL" & glngSys \ 100 & "_INSIDE_1341" Then
        Call ReportOpen(gcnOracle, glngSys, strName, Me)
    Else
        str��ǰ���� = mfrmList.GetCurrentRecipe
    
        If str��ǰ���� <> "" Then
            Int���� = Val(Split(str��ǰ����, "|")(0))
            strNo = Split(str��ǰ����, "|")(1)
            lng����ID = Val(Split(str��ǰ����, "|")(3))
        End If
        
        Call ReportOpen(gcnOracle, Split(Control.Parameter, ",")(0), strName, Me, _
            "ҩƷ=" & IIf(mSQLCondition.lngҩƷid = 0, "", mSQLCondition.lngҩƷid), _
            "ҩ��=" & IIf(mParams.lngҩ��ID = 0, "", mParams.lngҩ��ID), _
            "NO=" & strNo, _
            "��������=" & IIf(Int���� = 0, "", Int����), _
            "����ID=" & IIf(lng����ID = 0, "", lng����ID))
    End If
End Sub
Private Sub BillPrint_Dosage()
    '��ӡ��ҩ��
    Dim str��ǰ���� As String
    Dim Int���� As Integer, strNo As String
    Dim strUnit As String
    Dim int�����־ As Integer
    Dim int�������� As Integer
    Dim str�շ���� As String
    Dim lngRow As Long
    
    On Error GoTo errHandle
    
    str��ǰ���� = mfrmList.GetCurrentRecipe
    
    If str��ǰ���� = "" Then Exit Sub
    
    Int���� = Val(Split(str��ǰ����, "|")(0))
    strNo = Split(str��ǰ����, "|")(1)
    int�����־ = Val(Split(str��ǰ����, "|")(5))
    int�������� = Val(Split(str��ǰ����, "|")(6))
    str�շ���� = Split(str��ǰ����, "|")(7)
    lngRow = Val(Split(str��ǰ����, "|")(12))
    
    '��鵥���Ƿ����
    If Not CheckBillExist(Int����, strNo) Then
        MsgBox "δ�ҵ�ָ������,�����Ѿ�����������Ա����,����������ֹ��", vbInformation, gstrSysName
        RefreshList mcondition.intListType
        Exit Sub
    End If
    
    strUnit = GetUnit(mParams.lngҩ��ID, Int����, strNo, int�����־)

    If str�շ���� = "1" Then
        SetLocatePrinter int��������, Val(Split(mParams.str��ҩ��ʽ, ";")(0)) - 1
        Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1341_3", Me, _
            "NO=" & strNo, "����=" & IIf(Int���� = 8, 1, 2), "ҩ��=" & mParams.lngҩ��ID, "��װϵ��=" & IIf(strUnit = "���ﵥλ", "C.�����װ", "C.סԺ��װ"), "ReportFormat=" & Val(Split(mParams.str��ҩ��ʽ, ";")(0)), "PrintEmpty=0", 2)
        
        '�ָ�����ǩ�ı��ش�ӡ������
        Call SavePrinterSet("ZL1_BILL_1341_3", mstrRPTDefaultScheme_Recipt, mParams.strDefaultPrinter)
    ElseIf str�շ���� = "2" Then
        Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1341_4", Me, _
            "NO=" & strNo, "����=" & IIf(Int���� = 8, 1, 2), "ReportFormat=" & Val(Split(mParams.str��ҩ��ʽ, ";")(1)), "PrintEmpty=0", 2)
    Else
        'ͬʱ��ӡ��ҩ����ҩ�Ĵ���ǩ
        Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1341_4", Me, _
            "NO=" & strNo, "����=" & IIf(Int���� = 8, 1, 2), "ReportFormat=" & Val(Split(mParams.str��ҩ��ʽ, ";")(1)), "PrintEmpty=0", 2)
            
        SetLocatePrinter int��������, Val(Split(mParams.str��ҩ��ʽ, ";")(0)) - 1
        Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1341_3", Me, _
            "NO=" & strNo, "����=" & IIf(Int���� = 8, 1, 2), "ҩ��=" & mParams.lngҩ��ID, "��װϵ��=" & IIf(strUnit = "���ﵥλ", "C.�����װ", "C.סԺ��װ"), "ReportFormat=" & Val(Split(mParams.str��ҩ��ʽ, ";")(0)), "PrintEmpty=0", 2)
        
        '�ָ�����ǩ�ı��ش�ӡ������
        Call SavePrinterSet("ZL1_BILL_1341_3", mstrRPTDefaultScheme_Recipt, mParams.strDefaultPrinter)
    End If
    
    gstrSQL = "Zl_δ��ҩƷ��¼_���´�ӡ״̬("
    '����
    gstrSQL = gstrSQL & Int����
    'NO
    gstrSQL = gstrSQL & ",'" & strNo & "'"
    '�ⷿID
    gstrSQL = gstrSQL & "," & mParams.lngҩ��ID
    '��Դ����
    gstrSQL = gstrSQL & ",Null"
    '��ӡ����
    gstrSQL = gstrSQL & ",3"
    gstrSQL = gstrSQL & ")"
    Call zldatabase.ExecuteProcedure(gstrSQL, Me.Caption & "-���µ����Ѵ�ӡ")
    
    '�����б��ӡ��ʶ
    Call mfrmList.SetPrintFlag(lngRow)
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub BillPrint_Lable()
    '��ӡҩƷ��ǩ
    Dim str��ǰ���� As String
    Dim Int���� As Integer, strNo As String
    Dim strUnit As String
    Dim int�����־ As Integer
    Dim str�շ���� As String
    
    str��ǰ���� = mfrmList.GetCurrentRecipe
    
    If str��ǰ���� = "" Then Exit Sub
    
    Int���� = Val(Split(str��ǰ����, "|")(0))
    strNo = Split(str��ǰ����, "|")(1)
    int�����־ = Val(Split(str��ǰ����, "|")(5))
    str�շ���� = Split(str��ǰ����, "|")(7)
    
    strUnit = GetUnit(mParams.lngҩ��ID, Int����, strNo, int�����־)
    
    '��鵥���Ƿ����
    If Not CheckBillExist(Int����, strNo) Then
        MsgBox "δ�ҵ�ָ������,�����Ѿ�����������Ա����,����������ֹ��", vbInformation, gstrSysName
        RefreshList mcondition.intListType
        Exit Sub
    End If
    
    If str�շ���� = "1" Then
        Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1341_6", Me, _
            "NO=" & strNo, "����=" & IIf(Int���� = 8, 1, 2), "ҩ��=" & mParams.lngҩ��ID, "��װϵ��=" & IIf(strUnit = "���ﵥλ", "D.�����װ", "D.סԺ��װ"), "PrintEmpty=0", 2)
    ElseIf str�շ���� = "2" Then
        Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1341_7", Me, _
            "NO=" & strNo, "ҩ��=" & mParams.lngҩ��ID, "����=" & IIf(Int���� = 8, 1, 2), "PrintEmpty=0", 2)
    Else
        'ͬʱ��ӡ��ҩ����ҩ��ҩƷ��ǩ
        Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1341_6", Me, _
            "NO=" & strNo, "����=" & IIf(Int���� = 8, 1, 2), "ҩ��=" & mParams.lngҩ��ID, "��װϵ��=" & IIf(strUnit = "���ﵥλ", "D.�����װ", "D.סԺ��װ"), "PrintEmpty=0", 2)
            
        Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1341_7", Me, _
            "NO=" & strNo, "ҩ��=" & mParams.lngҩ��ID, "����=" & IIf(Int���� = 8, 1, 2), "PrintEmpty=0", 2)
    End If
    
End Sub

Private Sub BillPrint_Recipe()
    '��ӡ����ǩ
    Dim str��ǰ���� As String
    Dim Int���� As Integer, strNo As String
    Dim strUnit As String
    Dim int�����־ As Integer
    Dim int�������� As Integer
    Dim str�շ���� As String
    
    str��ǰ���� = mfrmList.GetCurrentRecipe
    
    If str��ǰ���� = "" Then Exit Sub
    
    Int���� = Val(Split(str��ǰ����, "|")(0))
    strNo = Split(str��ǰ����, "|")(1)
    int�����־ = Val(Split(str��ǰ����, "|")(5))
    int�������� = Val(Split(str��ǰ����, "|")(6))
    str�շ���� = Split(str��ǰ����, "|")(7)
    
    strUnit = GetUnit(mParams.lngҩ��ID, Int����, strNo, int�����־)
    
    If str�շ���� = "1" Then
        SetLocatePrinter int��������, Val(Split(mParams.str������ʽ, ";")(0)) - 1
        Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1341_3", Me, _
            "NO=" & strNo, _
            "����=" & IIf(Int���� = 8, 1, 2), _
            "ҩ��=" & mParams.lngҩ��ID, "��װϵ��=" & IIf(strUnit = "���ﵥλ", "C.�����װ", "C.סԺ��װ"), _
            "ReportFormat=" & Val(Split(mParams.str������ʽ, ";")(0)), "PrintEmpty=0", 2)
        
        '�ָ�����ǩ�ı��ش�ӡ������
        Call SavePrinterSet("ZL1_BILL_1341_3", mstrRPTDefaultScheme_Recipt, mParams.strDefaultPrinter)
    ElseIf str�շ���� = "2" Then
        Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1341_4", Me, _
            "NO=" & strNo, _
            "����=" & IIf(Int���� = 8, 1, 2), _
            "ReportFormat=" & Val(Split(mParams.str������ʽ, ";")(1)), "PrintEmpty=0", 2)
    Else
        'ͬʱ��ӡ��ҩ����ҩ�Ĵ���ǩ
        Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1341_4", Me, _
            "NO=" & strNo, _
            "����=" & IIf(Int���� = 8, 1, 2), _
            "ReportFormat=" & Val(Split(mParams.str������ʽ, ";")(1)), "PrintEmpty=0", 2)
        
        SetLocatePrinter int��������, Val(Split(mParams.str������ʽ, ";")(0)) - 1
        Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1341_3", Me, _
            "NO=" & strNo, _
            "����=" & IIf(Int���� = 8, 1, 2), _
            "ҩ��=" & mParams.lngҩ��ID, "��װϵ��=" & IIf(strUnit = "���ﵥλ", "C.�����װ", "C.סԺ��װ"), _
            "ReportFormat=" & Val(Split(mParams.str������ʽ, ";")(0)), "PrintEmpty=0", 2)
        
        '�ָ�����ǩ�ı��ش�ӡ������
        Call SavePrinterSet("ZL1_BILL_1341_3", mstrRPTDefaultScheme_Recipt, mParams.strDefaultPrinter)
    End If
    
End Sub

Private Sub BillPrint_Report()
    '��ӡ��ҩ�嵥
    
    Call ReportOpen(gcnOracle, glngSys, IIf(glngSys \ 100 = 1, "ZL1_BILL_1341_2", "ZL8_BILL_1341_2"), Me, _
        "�ⷿ=" & mstrStockName & "|" & mParams.lngҩ��ID, _
        "��װϵ��=" & IIf(mintUnit = mconint���ﵥλ, "D.�����װ", "D.סԺ��װ"))
End Sub

Private Sub BillPrint_Return()
    '��ӡ��ҩ֪ͨ��
    Dim str��ǰ���� As String
    Dim Int���� As Integer, strNo As String, Str��ҩʱ�� As String
    Dim strUnit As String
    Dim int�����־ As Integer
    
    str��ǰ���� = mfrmList.GetCurrentRecipe
    
    If str��ǰ���� = "" Then Exit Sub
    
    Int���� = Val(Split(str��ǰ����, "|")(0))
    strNo = Split(str��ǰ����, "|")(1)
    Str��ҩʱ�� = Split(str��ǰ����, "|")(2)
    int�����־ = Split(str��ǰ����, "|")(5)
    
    strUnit = GetUnit(mParams.lngҩ��ID, Int����, strNo, int�����־)
    
    Call ReportOpen(gcnOracle, glngSys, IIf(glngSys \ 100 = 1, "ZL1_BILL_1341_1", "ZL8_BILL_1341_1"), _
    Me, "No=" & strNo, "����=" & Int����, "��װϵ��=" & IIf(strUnit = "���ﵥλ", "D.�����װ", "D.סԺ��װ"), "��ҩʱ��=" & Str��ҩʱ��, 2)
End Sub

Private Sub BillPrint_Change()
    '��ӡҽ������֪ͨ��
    Dim str��ǰ���� As String
    Dim Int���� As Integer, strNo As String, Str��ҩʱ�� As String
    Dim strUnit As String
    Dim int�����־ As Integer
    
    str��ǰ���� = mfrmList.GetCurrentRecipe
    
    If str��ǰ���� = "" Then Exit Sub
    
    Int���� = Val(Split(str��ǰ����, "|")(0))
    strNo = Split(str��ǰ����, "|")(1)
    
    strUnit = GetUnit(mParams.lngҩ��ID, Int����, strNo, 1)
    
    Call ReportOpen(gcnOracle, glngSys, IIf(glngSys \ 100 = 1, "ZL1_BILL_1341_9", "ZL1_BILL_1341_9"), _
    Me, "No=" & strNo, "��װϵ��=" & IIf(strUnit = "���ﵥλ", "C.�����װ", "c.סԺ��װ"), "��ҩ�ⷿ=" & mParams.lngҩ��ID, 2)
End Sub

Private Sub ChangeDosagePeople()
    '�л���ҩ��
    Dim strName As String
    
    SetTimerState False
    
    strName = zldatabase.UserIdentify(Me, "У����ҩ��", glngSys, 1341, "��ҩ")
    
    SetTimerState True
    
    If Trim(strName) = "" Then Exit Sub
    
    mstr�Զ���ҩ�� = strName
    
    mdate�ϴ�У��ʱ�� = Sys.Currentdate
End Sub

Private Function CheckCard(ByVal rsData As ADODB.Recordset) As Boolean
    'һ��ͨ����ˢ����֤
    Dim dblSumMoney As Double
    Dim lng����ID As Long
    Dim blnCheck As Boolean
    Dim bytType As Byte '����סԺ��־
    
    If mParams.bln��ҩǰ�շѻ���� = True Then
        CheckCard = True
        Exit Function
    End If
    
    If mParams.bln��ҩ��ˢ����֤ = True Then
        blnCheck = True
        rsData.Filter = "��־=1 And ����ID>0 And ���￨��<>''"
    ElseIf mParams.bln��˻��۵� = True And mParams.blnˢ����֤ = True Then
        blnCheck = True
        rsData.Filter = "�����־=1 And ��־=1 And ��¼����=2 And ��¼״̬=0 And ����ID>0 And ���￨��<>''"
    End If
    
    If blnCheck = True Then
        rsData.Sort = "����ID"

        With rsData
            Do While Not .EOF
                If Val(!��¼����) = 1 Or (Val(!��¼����) = 2 And (Val(!�����־)) = 1 Or (Val(!�����־)) = 4) Then
                    bytType = 1
                Else
                    bytType = 2
                End If
            
                If lng����ID <> !����ID Then
                    If lng����ID <> 0 Then
                        If zldatabase.PatiIdentify(Me, glngSys, lng����ID, dblSumMoney, mlngMode, bytType) = False Then Exit Function
                    End If
                    
                    dblSumMoney = !ʵ�ս��
                    lng����ID = !����ID
                Else
                    dblSumMoney = dblSumMoney + !ʵ�ս��
                End If

                .MoveNext

                If .EOF Then
                    If zldatabase.PatiIdentify(Me, glngSys, lng����ID, dblSumMoney, mlngMode, bytType) = False Then Exit Function
                End If
            Loop
        End With
    End If

    CheckCard = True
End Function

Public Sub ClearForm_Detail()
    If Not mfrmDetail Is Nothing Then mfrmDetail.FormClear
End Sub

Public Sub ClearForm_Recipe()
    If Not mfrmDetail Is Nothing Then mfrmRecipe.FormClear
End Sub
Public Sub FindListRow(ByVal intFindType As Integer, ByVal strFind As String, ByVal str���� As String)
    If Not mfrmList Is Nothing Then
        mfrmList.FindSpecialRow "���ݺ�", strFind, "", mobjSquareCard, str����
    End If
End Sub

Private Sub GetDosage(ByVal lngҩ��ID As Long)
    'ȡҩ����ҩ������Ϣ
    On Error GoTo errHandle
    gstrSQL = "Select ��ҩ, Nvl(����, 1) As ���� From ҩ����ҩ���� Where ҩ��id = [1]"
    Set mrsIsDosage = zldatabase.OpenSQLRecord(gstrSQL, "ȡҩ����ҩ������Ϣ", lngҩ��ID)
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub GetDrugStock(ByVal lng�ⷿID As Long)
    On Error GoTo errHandle
    gstrSQL = "Select �շ�ϸĿid As ҩƷID From �շ�ִ�п��� Where ִ�п���id = [1]"
    Set mrsDrugStock = zldatabase.OpenSQLRecord(gstrSQL, "ȡ�洢�ⷿ", lng�ⷿID)
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Function GetRecipeByNO(ByVal strNo As String, Optional ByVal int��ѯ As Integer) As ADODB.Recordset
'���ܣ���ȡָ�������ŵ�ҩƷ����
'������
'  strNO��������
'  int��ѯ:1-[��ҩ]��ǩ�¿��Բ�ѯ[����ҩ]��[����ҩ]�еĵ���
'���أ�ҩƷ��¼������

    Dim rsData As ADODB.Recordset
    Dim strMsg As String
    Dim strsql As String
    
    If strNo = "" Then Exit Function
    On Error GoTo errHandle
    If mcondition.intListType <> mListType.��ҩ Or int��ѯ = 1 Then
        '���жϵ����Ƿ����
        gstrSQL = "Select Distinct A.��¼״̬,A.NO, A.����, B.����, Decode(A.����,8,'�շ�',9,'����') ����, A.�ⷿid As ҩ��ID, C.���� As ҩ��, " & _
                  "    B.��¼����, A.��������, B.�����־, a.��ҩ����, a.������� " & vbNewLine & _
                  "From ҩƷ�շ���¼ A, ������ü�¼ B, ���ű� C, ���ű� D " & vbNewLine & _
                  "Where A.����id = B.ID And A.�ⷿid = C.ID And A.�Է�����id = D.ID And Nvl(B.����״̬,0)<>1 " & _
                  "    And mod(A.��¼״̬,3)=1 And A.NO = [1] "
            
        If mPrives.bln������ҩ���Ĵ��� = False Or mcondition.intListType = mListType.����ҩ Then
            gstrSQL = gstrSQL & " And (Nvl(A.�ⷿid, 0) = 0 Or A.�ⷿid + 0 = [2]) "
        End If
        
        If mstrDeptNode <> "" Then
            gstrSQL = gstrSQL & " And (D.վ�� = [3] Or D.վ�� Is Null) "
        End If
        
        If mcondition.int������� = 3 Then
            gstrSQL = gstrSQL & " And A.���� In (8,9)" '���ＰסԺ���е���
            strsql = Replace(gstrSQL, "������ü�¼", "סԺ���ü�¼")
            strsql = Replace(strsql, "And Nvl(B.����״̬,0)<>1", "")
            gstrSQL = gstrSQL & " Union All " & strsql
        ElseIf mcondition.int������� = 1 Then
            gstrSQL = gstrSQL & " And A.���� In (8,9) " '���ﻮ�ۼ��������
        Else
            gstrSQL = gstrSQL & " And A.���� = 9 " 'סԺ����
            gstrSQL = Replace(gstrSQL, "������ü�¼", "סԺ���ü�¼")
            gstrSQL = Replace(gstrSQL, "And Nvl(B.����״̬,0)<>1", "")
        End If
        
        Set rsData = zldatabase.OpenSQLRecord(gstrSQL, "�жϵ����Ƿ����", strNo, mSQLCondition.lngҩ��ID, mstrDeptNode)
        
        If rsData.EOF Then
            Set GetRecipeByNO = Nothing
            Exit Function
        End If
        
        Set GetRecipeByNO = rsData
    Else
        gstrSQL = " Select Distinct A.��¼״̬,P.���� As ҩ��,Decode(A.����,8,'�շ�',9,'����') ����,Decode(A.����,8,'�շ�',9,'����') ����,A.No,A.����,H.����,A.�ⷿid as ҩ��id, '' ��ҩ��,'' �����, '' �������,H.�����־,H.��¼����,A.�������� " & _
                 " From " & _
                 "     (SELECT A.ID,A.No,A.����,A.ҩƷID,A.����ID,A.����,A.����,A.Ч��," & _
                 "          DECODE(SIGN((A.ʵ������*NVL(A.����,1))-B.�ѷ�����),0,A.����,1) ����," & _
                 "          DECODE(SIGN((A.ʵ������*NVL(A.����,1))-B.�ѷ�����),0,A.ʵ������,B.�ѷ�����) ʵ������,A.��¼״̬," & _
                 "          A.���ۼ�,A.���۽��,A.����,A.Ƶ��,A.�÷�,A.ժҪ,A.�����,A.�������,A.������,A.��������,A.��ҩ��,A.�Է�����ID,A.�ⷿID" & _
                 "      From" & _
                 "          (SELECT A.ID,A.No,A.����,A.ҩƷID,A.���,A.����ID,A.����,A.����,A.Ч��,A.ʵ������,A.����,A.��¼״̬,A.���ۼ�,A.���۽��,A.����,A.Ƶ��,A.�÷�,A.ժҪ,A.�����,A.�������,A.������,A.��������,A.��ҩ��,A.�Է�����ID,A.�ⷿID " & _
                 "          From ҩƷ�շ���¼ A" & _
                 "          WHERE A.����� IS NOT NULL AND (A.��¼״̬=1 OR MOD(A.��¼״̬,3)=0)" & _
                 "          And A.�ⷿID+0=[2] And A.No =[1] ) A," & _
                 "          (SELECT A.NO,A.����,A.ҩƷID,A.���,SUM(NVL(A.����,1)*A.ʵ������) �ѷ�����" & _
                 "          From ҩƷ�շ���¼ A" & _
                 "          Where A.����� Is Not Null" & _
                 "          And A.�ⷿID+0=[2] And A.No =[1] " & _
                 "          GROUP BY A.no,A.����,A.ҩƷID,A.���) B" & _
                 "      Where A.no = B.no And A.���� = B.���� And A.ҩƷID+0 = B.ҩƷID And A.��� = B.���" & _
                 "     ) A,������ü�¼ H,���ű� P " & _
                 " Where A.�ⷿID=P.id And A.�ⷿID+0=[2] " & _
                 " And A.No =[1] " & _
                 " And A.����ID=H.ID And (Mod(A.��¼״̬,3)=0 Or A.��¼״̬=1) And A.ʵ������<>0 "
        
        If mcondition.int������� = 3 Then
            gstrSQL = gstrSQL & " And A.���� In (8,9)" '���ＰסԺ���е���
            gstrSQL = gstrSQL & " Union All " & Replace(gstrSQL, "������ü�¼", "סԺ���ü�¼")
        ElseIf mcondition.int������� = 1 Then
            gstrSQL = gstrSQL & " And A.���� In (8,9) " '���ﻮ�ۼ��������
        Else
            gstrSQL = gstrSQL & " And A.���� = 9 " 'סԺ����
            gstrSQL = Replace(gstrSQL, "������ü�¼", "סԺ���ü�¼")
        End If
        
        'һ�Ŵ���������ͬʱ������������󱸱��У���ˣ���������Ƴ�����ֱ�ӴӺ󱸱�����ȡ������ԭSQL����
        'ҩƷ������ҩ��ͬʱ�Ե��� IN (8,9)�ĵ��ݣ���˲��ų�����8���߶�9���е����
        Dim blnMoved As Boolean
        
        blnMoved = Sys.IsMovedByNO("ҩƷ�շ���¼", strNo, " ���� IN ", " (8,9)")
        
        '�����������ת��������Ҫͬʱ�Ӻ󱸱�����ȡ���ݣ����ܴ��ڲ�ͬ���͵ĵ��ݷֱ�������󱸱��У�
        If blnMoved Then
            strsql = gstrSQL
            strsql = Replace(strsql, "ҩƷ�շ���¼", "HҩƷ�շ���¼")
            strsql = Replace(strsql, "������ü�¼", "H������ü�¼")
            strsql = Replace(strsql, "סԺ���ü�¼", "HסԺ���ü�¼")
            gstrSQL = gstrSQL & " UNION ALL " & strsql
        End If
        
        Set GetRecipeByNO = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, strNo, mSQLCondition.lngҩ��ID)
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub GetStockName(ByVal lng�ⷿID As Long)
    Dim rsTmp As ADODB.Recordset
    On Error GoTo errHandle
    gstrSQL = "Select ���� From ���ű� Where ID = [1]"
    Set rsTmp = zldatabase.OpenSQLRecord(gstrSQL, "ȡ�ⷿ����", lng�ⷿID)
    
    If Not rsTmp.EOF Then
        mstrStockName = rsTmp!����
    Else
        mstrStockName = ""
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub PrintRecipe()
    '��ӡ����ǩ
    Dim blnPrint As Boolean
    Dim arrRecipe
    Dim n As Integer
    Dim intNum As Integer
    Dim strRecipeNo As String
    Dim intBillType As Integer
    Dim int�������� As Integer
    Dim str�շ���� As String
    
    If mstrPrintRecipe = "" Then Exit Sub
    
    mstrPrintRecipe = mstrPrintRecipe & "|"
    
    If mParams.IntAutoPrint < 2 Then
        blnPrint = IIf(mParams.IntAutoPrint = 1, True, False)
        If mParams.IntAutoPrint = 0 Then
            If MsgBox("��ӡ�ô���������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then blnPrint = True
        End If
        
        If blnPrint Then
            arrRecipe = Split(mstrPrintRecipe, "|")
            intNum = UBound(arrRecipe)
            
            For n = 0 To intNum
                If arrRecipe(n) <> "" Then
                    strRecipeNo = Split(arrRecipe(n), ",")(0)
                    intBillType = Val(Split(arrRecipe(n), ",")(1))
                    int�������� = Val(Split(arrRecipe(n), ",")(4))
                    str�շ���� = Split(arrRecipe(n), ",")(5)
    
                    If str�շ���� = "1" Then
                        SetLocatePrinter int��������, Val(Split(mParams.str������ʽ, ";")(0)) - 1
                        Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1341_3", Me, _
                            "NO=" & strRecipeNo, _
                            "����=" & IIf(intBillType = 8, 1, 2), _
                            "ҩ��=" & mParams.lngҩ��ID, "��װϵ��=" & IIf(mParams.strUnit = "���ﵥλ", "C.�����װ", "C.סԺ��װ"), _
                            "ReportFormat=" & Val(Split(mParams.str������ʽ, ";")(0)), "PrintEmpty=0", IIf(mParams.blnPreview And mParams.IntAutoPrint = 0, 0, 2))
                        
                        '�ָ�����ǩ�ı��ش�ӡ������
                        Call SavePrinterSet("ZL1_BILL_1341_3", mstrRPTDefaultScheme_Recipt, mParams.strDefaultPrinter)
                    ElseIf str�շ���� = "2" Then
                        Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1341_4", Me, _
                            "NO=" & strRecipeNo, _
                            "����=" & IIf(intBillType = 8, 1, 2), _
                            "ReportFormat=" & Val(Split(mParams.str������ʽ, ";")(1)), "PrintEmpty=0", IIf(mParams.blnPreview And mParams.IntAutoPrint = 0, 0, 2))
                    Else
                        Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1341_4", Me, _
                            "NO=" & strRecipeNo, _
                            "����=" & IIf(intBillType = 8, 1, 2), _
                            "ReportFormat=" & Val(Split(mParams.str������ʽ, ";")(1)), "PrintEmpty=0", IIf(mParams.blnPreview And mParams.IntAutoPrint = 0, 0, 2))
                            
                        SetLocatePrinter int��������, Val(Split(mParams.str������ʽ, ";")(0)) - 1
                        Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1341_3", Me, _
                            "NO=" & strRecipeNo, _
                            "����=" & IIf(intBillType = 8, 1, 2), _
                            "ҩ��=" & mParams.lngҩ��ID, "��װϵ��=" & IIf(mParams.strUnit = "���ﵥλ", "C.�����װ", "C.סԺ��װ"), _
                            "ReportFormat=" & Val(Split(mParams.str������ʽ, ";")(0)), "PrintEmpty=0", IIf(mParams.blnPreview And mParams.IntAutoPrint = 0, 0, 2))
                        
                        '�ָ�����ǩ�ı��ش�ӡ������
                        Call SavePrinterSet("ZL1_BILL_1341_3", mstrRPTDefaultScheme_Recipt, mParams.strDefaultPrinter)
                    End If
                End If
            Next
        End If
    End If
    
    blnPrint = False
    If mParams.int��ҩ���Զ���ӡҩƷ��ǩ < 2 Then
        blnPrint = IIf(mParams.int��ҩ���Զ���ӡҩƷ��ǩ = 1, True, False)
        If mParams.int��ҩ���Զ���ӡҩƷ��ǩ = 0 Then
            If MsgBox("��ӡҩƷ��ǩ��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then blnPrint = True
        End If
    
        If blnPrint Then
            arrRecipe = Split(mstrPrintRecipe, "|")
            intNum = UBound(arrRecipe)
            
            For n = 0 To intNum
                If arrRecipe(n) <> "" Then
                    strRecipeNo = Split(arrRecipe(n), ",")(0)
                    intBillType = Val(Split(arrRecipe(n), ",")(1))
                    int�������� = Val(Split(arrRecipe(n), ",")(4))
                    str�շ���� = Split(arrRecipe(n), ",")(5)
                    
                    If str�շ���� = "1" Then
                        Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1341_6", Me, _
                            "NO=" & strRecipeNo, "����=" & IIf(intBillType = 8, 1, 2), "ҩ��=" & mParams.lngҩ��ID, "��װϵ��=" & IIf(mParams.strUnit = "���ﵥλ", "D.�����װ", "D.סԺ��װ"), "PrintEmpty=0", 2)
                    ElseIf str�շ���� = "2" Then
                        Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1341_7", Me, _
                            "NO=" & strRecipeNo, "ҩ��=" & mParams.lngҩ��ID, "����=" & IIf(intBillType = 8, 1, 2), "PrintEmpty=0", 2)
                    Else
                        Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1341_6", Me, _
                            "NO=" & strRecipeNo, "����=" & IIf(intBillType = 8, 1, 2), "ҩ��=" & mParams.lngҩ��ID, "��װϵ��=" & IIf(mParams.strUnit = "���ﵥλ", "D.�����װ", "D.סԺ��װ"), "PrintEmpty=0", 2)
                        
                        Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1341_7", Me, _
                            "NO=" & strRecipeNo, "ҩ��=" & mParams.lngҩ��ID, "����=" & IIf(intBillType = 8, 1, 2), "PrintEmpty=0", 2)
                    End If
                End If
            Next
        End If
    End If
    
    mstrPrintRecipe = ""
End Sub

Private Sub PrintDosage()
    '��ӡ��ҩ��
    Dim blnPrint As Boolean
    Dim arrRecipe
    Dim n As Integer
    Dim intNum As Integer
    Dim strRecipeNo As String
    Dim intBillType As Integer
    Dim int�����־ As Integer
    Dim int�������� As Integer
    Dim str�շ���� As String
    Dim strUnit As String
    Dim blnIsPrintForPrintDosage As Boolean
    
    On Error GoTo errHandle
     
    If mstrPrintRecipe = "" Then Exit Sub
    
    If mParams.int��ҩ���Զ���ӡ = 2 Then Exit Sub
    
    If mParams.int��ҩ���Զ���ӡ < 2 Then
        blnIsPrintForPrintDosage = IIf(mParams.int��ҩ���Զ���ӡ = 1, True, False)
    
        If mParams.int��ҩ���Զ���ӡ = 0 Then
            blnIsPrintForPrintDosage = IIf(MsgBox("��ӡ����ҩ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes, True, False)
        End If
        
        If blnIsPrintForPrintDosage Then
            arrRecipe = Split(mstrPrintRecipe, "|")
            intNum = UBound(arrRecipe)
            
            For n = 0 To intNum
                If arrRecipe(n) <> "" Then
                    strRecipeNo = Split(arrRecipe(n), ",")(0)
                    intBillType = Val(Split(arrRecipe(n), ",")(1))
                    int�����־ = Val(Split(arrRecipe(n), ",")(3))
                    int�������� = Val(Split(arrRecipe(n), ",")(4))
                    str�շ���� = Split(arrRecipe(n), ",")(5)
                    
                    strUnit = GetUnit(mParams.lngҩ��ID, intBillType, strRecipeNo, int�����־)
                    
                    If str�շ���� = "1" Then
                        SetLocatePrinter int��������, Val(Split(mParams.str��ҩ��ʽ, ";")(0)) - 1
                        Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1341_3", Me, _
                            "NO=" & strRecipeNo, "����=" & IIf(intBillType = 8, 1, 2), "ҩ��=" & mParams.lngҩ��ID, "��װϵ��=" & IIf(strUnit = "���ﵥλ", "C.�����װ", "C.סԺ��װ"), "ReportFormat=" & Val(Split(mParams.str��ҩ��ʽ, ";")(0)), "PrintEmpty=0", 2)
                        
                        '�ָ�����ǩ�ı��ش�ӡ������
                        Call SavePrinterSet("ZL1_BILL_1341_3", mstrRPTDefaultScheme_Recipt, mParams.strDefaultPrinter)
                    ElseIf str�շ���� = "2" Then
                        Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1341_4", Me, _
                            "NO=" & strRecipeNo, "����=" & IIf(intBillType = 8, 1, 2), "ReportFormat=" & Val(Split(mParams.str��ҩ��ʽ, ";")(1)), "PrintEmpty=0", 2)
                    Else
                        Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1341_4", Me, _
                            "NO=" & strRecipeNo, "����=" & IIf(intBillType = 8, 1, 2), "ReportFormat=" & Val(Split(mParams.str��ҩ��ʽ, ";")(1)), "PrintEmpty=0", 2)
                            
                        SetLocatePrinter int��������, Val(Split(mParams.str��ҩ��ʽ, ";")(0)) - 1
                        Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1341_3", Me, _
                            "NO=" & strRecipeNo, "����=" & IIf(intBillType = 8, 1, 2), "ҩ��=" & mParams.lngҩ��ID, "��װϵ��=" & IIf(strUnit = "���ﵥλ", "C.�����װ", "C.סԺ��װ"), "ReportFormat=" & Val(Split(mParams.str��ҩ��ʽ, ";")(0)), "PrintEmpty=0", 2)
                        
                        '�ָ�����ǩ�ı��ش�ӡ������
                        Call SavePrinterSet("ZL1_BILL_1341_3", mstrRPTDefaultScheme_Recipt, mParams.strDefaultPrinter)
                    End If
                End If
            Next
        End If
        
        '����ѭ�����´�ӡ״̬�����ŵ���ӡ�����ѭ��������
        If blnIsPrintForPrintDosage Then
            arrRecipe = Split(mstrPrintRecipe, "|")
            intNum = UBound(arrRecipe)
                
            For n = 0 To intNum
                If arrRecipe(n) <> "" Then
                    strRecipeNo = Split(arrRecipe(n), ",")(0)
                    intBillType = Val(Split(arrRecipe(n), ",")(1))
                                           
                    gstrSQL = "Zl_δ��ҩƷ��¼_���´�ӡ״̬("
                    '����
                    gstrSQL = gstrSQL & intBillType
                    'NO
                    gstrSQL = gstrSQL & ",'" & strRecipeNo & "'"
                    '�ⷿID
                    gstrSQL = gstrSQL & "," & mParams.lngҩ��ID
                    '��Դ����
                    gstrSQL = gstrSQL & ",Null"
                    '��ӡ����
                    gstrSQL = gstrSQL & ",3"
                    gstrSQL = gstrSQL & ")"
                    Call zldatabase.ExecuteProcedure(gstrSQL, Me.Caption & "-���µ����Ѵ�ӡ")
                End If
            Next
        End If
    End If
    mstrPrintRecipe = ""
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function BillHaveHerial(ByVal strNo As String, ByVal Int���� As Integer, ByVal int���� As Integer, Optional ByRef str�շ�ϸĿid As String, Optional ByRef str�շ���� As String) As String
'--------------------------------------------
'����Ƿ�����ҩ����
'-------------------------------------------
    Dim rsTmp As New ADODB.Recordset
    
    On Error GoTo errH
    
    If int���� = 1 Then
        gstrSQL = "Select NO,�շ����,�շ�ϸĿid From ������ü�¼ Where NO=[1] And ��¼״̬ IN(0,1,3)" & _
            " And ��¼����=[3] And ִ�в���ID+0=[2]"
    Else
        gstrSQL = "Select NO,�շ����,�շ�ϸĿid From סԺ���ü�¼ Where NO=[1] And ��¼״̬ IN(0,1,3)" & _
            " And ��¼����=[3] And ִ�в���ID+0=[2]"
    End If
    Set rsTmp = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, strNo, mParams.lngҩ��ID, IIf(Int���� = 8, 1, 2))
    
    Do While Not rsTmp.EOF
        str�շ���� = str�շ���� & rsTmp!�շ���� & ","
        BillHaveHerial = BillHaveHerial & rsTmp!�շ���� & ";"
        If InStr(1, "," & str�շ�ϸĿid, "," & rsTmp!�շ�ϸĿid & ",") < 1 Then str�շ�ϸĿid = str�շ�ϸĿid & rsTmp!�շ�ϸĿid & ","
        rsTmp.MoveNext
    Loop
'    BillHaveHerial = Not rsTmp.EOF
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function CheckAnother() As Boolean
'------------------------------------------
'����Ƿ����ù�ҩ������ҩ��,�����Ǹ��ݲ�����ҩҩ������
'------------------------------------------
    Dim BlnInҩ�� As Boolean, blnסԺ As Boolean, Bln���� As Boolean
    Dim BlnSetPeople As Boolean
    Dim RecTestPeople As New ADODB.Recordset
    Dim LngOldҩ��ID As Long, StrOld��ҩ�� As String
    
    CheckAnother = False
    On Error GoTo errHandle
    If mParams.lngҩ��ID <> 0 Then
        With RecPart
            .MoveFirst
            .Find "ID=" & mParams.lngҩ��ID
            BlnInҩ�� = (RecPart.EOF <> True)
            
            If BlnInҩ�� Then   '˵���ò�������ҩ��
                'ȡ��λ
                blnסԺ = False

                gstrSQL = "Select nvl(�������,1) ������� From ��������˵�� Where ����ID+0=[1]"
                Set RecTestPeople = zldatabase.OpenSQLRecord(gstrSQL, "ȡ���ŷ������", mParams.lngҩ��ID)
                
                With RecTestPeople
                    Do While Not .EOF
                        If !������� = 2 Or !������� = 3 Then blnסԺ = True: Exit Do
                        .MoveNext
                    Loop
                    Bln���� = False
                    If blnסԺ Then
                        If .RecordCount <> 0 Then .MoveFirst
                        Do While Not .EOF
                            If !������� = 3 Then Bln���� = True: Exit Do
                            .MoveNext
                        Loop
                    End If
                End With
                If blnסԺ = False Then
                    mcondition.int������� = 1
                Else
                    mcondition.int������� = IIf(Bln����, 3, 2)
                End If
            End If
        End With
    End If
    
    '���ö�Ӧ��ҩ������������Զ�����ҩ�ӿ������Ҫ���÷�ҩ����
    If mParams.lngҩ��ID = 0 Or BlnInҩ�� = False Or (Not mobjDrugMAC Is Nothing And InStr(1, mParams.Str����, ",") > 0) Then
        '�����ô���
        With Frm��ҩ��������
            If mParams.lngҩ��ID = 0 Or BlnInҩ�� = False Then
                MsgBox IIf(mParams.str��ҩ�� = "", "������ҩ������ҩ�ˣ�", "������ҩ����"), vbInformation, gstrSysName
                Set .RecPart = RecPart.Clone
                .strShow = IIf(mParams.str��ҩ�� = "", "������ҩ������ҩ�ˣ�", "������ҩ����")
            Else
                MsgBox "����ҩ���Զ���ҩֻ������һ����ҩ���ڣ�", vbInformation, gstrSysName
                Set .RecPart = RecPart.Clone
                .strShow = "����ҩ���Զ���ҩֻ������һ����ҩ����!"
            End If
            .mstrPrivs = mstrPrivs
            .In_���÷�ҩ = mblnLoadDrug
            .Show 1, Me
        End With
        Call GetParams

        '��δ����ҩ�����˳�
        If mParams.lngҩ��ID = 0 Then Exit Function
        
        mfrmList.SetParams
        mfrmDetail.SetParams
        mfrmRecipe.SetParams
        
        '���»�ȡ��ҩ����ʹ�õ�λ
        With RecPart
            .MoveFirst
            .Find "ID=" & mParams.lngҩ��ID
            BlnInҩ�� = (RecPart.EOF <> True)
            
            If BlnInҩ�� Then   '˵���ò�������ҩ��
                'ȡ��λ
                blnסԺ = False

                gstrSQL = "Select nvl(�������,1) ������� From ��������˵�� Where ����ID+0=[1]"
                Set RecTestPeople = zldatabase.OpenSQLRecord(gstrSQL, "ȡ���ŷ������", mParams.lngҩ��ID)
                
                With RecTestPeople
                    Do While Not .EOF
                        If !������� = 2 Or !������� = 3 Then blnסԺ = True: Exit Do
                        .MoveNext
                    Loop
                    Bln���� = False
                    If blnסԺ Then
                        If .RecordCount <> 0 Then .MoveFirst
                        Do While Not .EOF
                            If !������� = 3 Then Bln���� = True: Exit Do
                            .MoveNext
                        Loop
                    End If
                End With
                If blnסԺ = False Then
                    mcondition.int������� = 1
                Else
                    mcondition.int������� = IIf(Bln����, 3, 2)
                End If
            Else
                Exit Function    '��ҩ�����˳�
            End If
        End With
    End If
    
    If mParams.blnMustDosageProcess = True And mParams.str��ҩ�� <> "|��ǰ����Ա|" Then
        LngOldҩ��ID = mParams.lngҩ��ID
        StrOld��ҩ�� = mParams.str��ҩ��
        
        '������ҩ��
        BlnSetPeople = False
        If mParams.str��ҩ�� = "" Then
            MsgBox "��������ҩ�ˣ�", vbInformation, gstrSysName
            With Frm��ҩ��������
                Set .RecPart = RecPart.Clone
                .strShow = "��������ҩ�ˣ�"
                .mstrPrivs = mstrPrivs
                .In_���÷�ҩ = mblnLoadDrug
                .Show 1, Me
            End With
            Call GetParams
            mfrmList.SetParams
            mfrmDetail.SetParams
            mfrmRecipe.SetParams

            If mParams.str��ҩ�� = "" Then
                MsgBox "������������ҩ�ˣ�����ϵͳ����Ա��ϵ��", vbInformation, gstrSysName
                Exit Function
            End If
        End If
        
        '�����ҩ�˷Ǳ�����,�������������
        gstrSQL = " Select Count(*) Records From ������Ա Where ��ԱID=" & _
                 " (Select Distinct ID From ��Ա�� Where ����=[2]) And " & _
                 " ����ID+0 =[1]"
        Set RecTestPeople = zldatabase.OpenSQLRecord(gstrSQL, "ȡ������Ա", mParams.lngҩ��ID, mParams.str��ҩ��)
        
        With RecTestPeople
            If .EOF Then
                BlnSetPeople = True
            Else
                If IsNull(!Records) Then
                    BlnSetPeople = True
                Else
                    If !Records = 0 Then
                        BlnSetPeople = True
                    End If
                End If
            End If
        End With
        If BlnSetPeople Then
            MsgBox "��������ҩ�ˣ�ԭ��ҩ���Ѳ����ڱ�ҩ������", vbInformation, gstrSysName
            With Frm��ҩ��������
                Set .RecPart = RecPart.Clone
                .strShow = "��������ҩ�ˣ�ԭ��ҩ���Ѳ����ڱ�ҩ������"
                .mstrPrivs = mstrPrivs
                .In_���÷�ҩ = mblnLoadDrug
                .Show 1, Me
            End With
            Call GetParams
            mfrmList.SetParams
            mfrmDetail.SetParams
            mfrmRecipe.SetParams
        
            If mParams.str��ҩ�� = "" Then
                MsgBox "������������ҩ�ˣ�ԭ��ҩ���Ѳ����ڱ�ҩ����������ϵͳ����Ա��ϵ��", vbInformation, gstrSysName
                Exit Function
            End If
            If StrOld��ҩ�� = mParams.str��ҩ�� And LngOldҩ��ID = mParams.lngҩ��ID Then Exit Function
        End If
    End If
    
    '��������Ŷӽкţ������������˶����ҩ���ڣ���Ҫ�������ò���
    If mParams.blnStartQueue = True And InStr(mstr����, ",") > 0 Then
        MsgBox "�������Ŷӽкţ��������ö����ҩ���ڣ����������ã�", vbInformation, gstrSysName
        
        With Frm��ҩ��������
            Set .RecPart = RecPart.Clone
            .mstrPrivs = mstrPrivs
            .In_���÷�ҩ = mblnLoadDrug
            .Show 1, Me
        End With
        
        Call GetParams
        mfrmList.SetParams
        mfrmDetail.SetParams
        mfrmRecipe.SetParams
        
        '��δ������ȷ�ģ��˳�
        If mParams.blnStartQueue = True And InStr(mParams.Str����, ",") > 0 Then
            MsgBox "��ҩ�������ò���ȷ�����ܽ��з�ҩ������", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    CheckAnother = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function DependOnCheck() As Boolean
    Dim strsql As String
    '�������ݼ��
    DependOnCheck = False
    On Error GoTo errHandle
    With RecPart
        gstrSQL = " Select A.����||'-'||A.���� ҽ�� From ��Ա�� A,��Ա����˵�� B" & _
                 " Where B.��Ա����='ҽ��' And A.ID=B.��ԱID" & _
                 " And (A.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.����ʱ�� Is Null) " & _
                 " Order by A.����"
        Call zldatabase.OpenRecordset(RecPart, gstrSQL, "�������ݼ��")
        
        If .EOF Then
            MsgBox "���ʼ����Ա��ҽ����", vbInformation, gstrSysName
            Exit Function
        End If
    End With
    
    If zlStr.IsHavePrivs(mstrPrivs, "����ҩ��") Then
        strsql = "(Select Distinct ����ID From ��������˵�� Where �������� Like '%ҩ��')"
    Else
        strsql = "(Select distinct A.����ID From ������Ա A,��������˵�� B " & _
                 " Where A.��ԱID=[1] And A.����ID=B.����ID And B.�������� Like '%ҩ��')"
    End If
    gstrSQL = " Select Distinct P.ID,P.���� From ���ű� P " & _
             " Where (P.վ�� = '" & gstrNodeNo & "' Or P.վ�� is Null) And P.ID In " & strsql & _
             " And (P.����ʱ�� Is Null Or P.����ʱ��=To_Date('3000-01-01','yyyy-MM-dd'))"
    Set RecPart = zldatabase.OpenSQLRecord(gstrSQL, "ȡҩ��", glngUserId)
    
    With RecPart
        If .EOF Then
            If zlStr.IsHavePrivs(mstrPrivs, "����ҩ��") Then
                strsql = "���ʼ��ҩ���������Ź���"
            Else
                strsql = "�㲻��ҩ����Ա������ʹ�ñ�ģ�飡"
            End If
            MsgBox strsql, vbInformation, gstrSysName
            Exit Function
        End If
    End With
    
    DependOnCheck = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Public Sub GetCondition()
    Dim dteTime As Date
    Dim strName As String
    
    
    
    dteTime = Sys.Currentdate
    
    mSQLCondition.lngҩ��ID = mParams.lngҩ��ID
    
    'ʱ�䷶Χ
    Select Case cboʱ�䷶Χ.ListIndex
        Case mTimeRange.����
            mSQLCondition.date��ʼ���� = CDate(Format(dteTime, "yyyy-mm-dd") & " 00:00:00")
            mSQLCondition.date�������� = CDate(Format(dteTime, "yyyy-mm-dd") & " 23:59:59")
        Case mTimeRange.������
            mSQLCondition.date��ʼ���� = CDate(Format(DateAdd("d", -1, dteTime), "yyyy-mm-dd") & " 00:00:00")
            mSQLCondition.date�������� = CDate(Format(dteTime, "yyyy-mm-dd") & " 23:59:59")
        Case mTimeRange.������
            mSQLCondition.date��ʼ���� = CDate(Format(DateAdd("d", -2, dteTime), "yyyy-mm-dd") & " 00:00:00")
            mSQLCondition.date�������� = CDate(Format(dteTime, "yyyy-mm-dd") & " 23:59:59")
        Case mTimeRange.ָ��ʱ�䷶Χ
            mSQLCondition.date��ʼ���� = CDate(Format(Dtp��ʼʱ��.Value, "yyyy-mm-dd hh:mm:ss"))
            mSQLCondition.date�������� = CDate(Format(Dtp����ʱ��.Value, "yyyy-mm-dd hh:mm:ss"))
        Case Else
            mSQLCondition.date��ʼ���� = CDate(Format(dteTime, "yyyy-mm-dd") & " 00:00:00")
            mSQLCondition.date�������� = CDate(Format(dteTime, "yyyy-mm-dd") & " 23:59:59")
    End Select
    
    mcondition.bln��ʾ��ҩ�������� = (Chk��ʾ��ҩ��������.Value = 1)
    mcondition.bln��ʾ���̵��� = (Chk��ʾ���̵���.Value = 1)
    mcondition.bln��ʾ��ȷ�ϵ��� = (chk��ʾ��ȷ�ϵ���.Value = 1)
    
    If mint���˲�ѯ = 1 Then
        If mSQLCondition.str���￨ <> "" Then
            If Split(Split(mSQLCondition.str���￨, "|")(1), ",")(0) = "�������֤" Then
                '���֤
                If UBound(Split(mSQLCondition.str���￨, "|")) > 1 Then
                    mSQLCondition.lng����ID = Split(mSQLCondition.str���￨, "|")(2)
                Else
                    mSQLCondition.str���֤ = Split(mSQLCondition.str���￨, "|")(0)
                End If
            ElseIf Split(Split(mSQLCondition.str���￨, "|")(1), ",")(0) = "IC��" Then
                'IC��
                If Not mobjSquareCard Is Nothing Then Call mobjSquareCard.zlGetPatiID("IC��", UCase(Trim(Split(mSQLCondition.str���￨, "|")(0))), False, mlngIC����id)
                mSQLCondition.lng����ID = mlngIC����id
            Else
                '�������ѿ���ȡ����ID
                mSQLCondition.lng����ID = zlfuncCard_GetPatiID(mobjSquareCard, Split(Split(mSQLCondition.str���￨, "|")(1), ",")(1), Split(mSQLCondition.str���￨, "|")(0))
            End If
            
            mSQLCondition.str���￨ = ""
        End If
    End If
    
    If imgFilter.BorderStyle = cstFilter Then
        '�������
        mSQLCondition.str���￨ = ""
        mSQLCondition.str��ǰNO = ""
        mSQLCondition.str����� = ""
        mSQLCondition.str���� = ""
        mSQLCondition.str���֤ = ""
        mSQLCondition.lng����ID = 0
        mSQLCondition.strҽ���� = ""
        mSQLCondition.lngסԺ�� = 0
    
        Select Case IDKNType.GetCurCard.����
            Case "���ݺ�"
                mSQLCondition.str��ǰNO = txtPati.Text
            Case "�����"
                mSQLCondition.str����� = txtPati.Text
            Case "����"
                mSQLCondition.str���� = txtPati.Text & "%"
            Case "���֤"
                mSQLCondition.str���֤ = txtPati.Text
            Case "IC��"
                mSQLCondition.lng����ID = mlngIC����id
            Case "ҽ����"
                mSQLCondition.strҽ���� = txtPati.Text
            Case "סԺ��"
                mSQLCondition.lngסԺ�� = Val(txtPati.Text)
            Case Else
                '�������ѿ���ȡ����ID
                mSQLCondition.lng����ID = zlfuncCard_GetPatiID(mobjSquareCard, mobjcard.�ӿ����, txtPati.Text)
        End Select
    End If
    
    mSQLCondition.intOverTime = mParams.intOverTime
End Sub

Private Sub GetParams()
    Dim arrColumn
    Dim strTmp As String
    Dim bln�Ƿ���ҩȷ�� As Boolean
    Dim intParaType As Integer
    Dim rsData As ADODB.Recordset
    Dim i As Integer
    
    With mParams
        .bln����δ��˴�����ҩ = (gtype_UserSysParms.P6_δ��˼��ʴ�����ҩ = 1)
        .bln����δ�շѴ�����ҩ = (gtype_UserSysParms.P148_δ�շѴ�����ҩ = 1)
        .blnҽ������ = (gtype_UserSysParms.P68_����ҩ�������Ϻ���ҩ = 0)
        .int����λ�� = GetDigit(0, 1, 4)
        .bln��˻��۵� = True
        .blnˢ����֤ = (Val(Left(gtype_UserSysParms.P28_���ﲡ������ʱ��Ҫˢ����֤, 1)) = 1)
        .bln�����������۷��� = (gtype_UserSysParms.P98_���ʱ����������۷��� <> 0)
        .bln��ҩǰ�շѻ���� = (gtype_UserSysParms.P163_��Ŀִ��ǰ�������շѻ��ȼ������ = 1)
        .bln������ = ((gtype_UserSysParms.P240_ҩ��������� = 1 Or gtype_UserSysParms.P240_ҩ��������� = 3) And gtype_UserSysParms.P241_�������ʱ�� = 2)
        
        '����ͬʱ֪ͨ�豸׼����ҩ
        .blnDispensing = Val(zldatabase.GetPara("����ʱ֪ͨ��ʼ��ҩ", glngSys, 1341)) = 1
        
        '�������ã�����
        .lngҩ��ID = Val(zldatabase.GetPara("��ҩҩ��", glngSys, 1341))
        .Str���� = Replace(zldatabase.GetPara("��ҩ����", glngSys, 1341), "'", "")
        .str��ҩ�� = zldatabase.GetPara("��ҩ��", glngSys, 1341)
        .bln�Զ���ҩ = (Val(zldatabase.GetPara("�Զ���ҩ", glngSys, 1341)) = 1)
        .int�Զ���ҩʱ�� = Val(zldatabase.GetPara("�Զ���ҩʱ��", glngSys, 1341))
        .IntAutoPrint = Val(zldatabase.GetPara("��ҩ���Զ���ӡ", glngSys, 1341))
        .int��ҩ���Զ���ӡ = Val(zldatabase.GetPara("��ҩ���Զ���ӡ", glngSys, 1341, 2))
        .int��ҩ���Զ���ӡҩƷ��ǩ = Val(zldatabase.GetPara("��ҩ���ӡҩƷ��ǩ", glngSys, 1341, 2))
        .intShowBill�շ� = Val(zldatabase.GetPara("�շѴ�����ʾ��ʽ", glngSys, 1341, 3))
        .intShowBill���� = Val(zldatabase.GetPara("���ʴ�����ʾ��ʽ", glngSys, 1341, 3))
        .intShowBill��ҩ = Val(zldatabase.GetPara("����ҩ���ݴ�ӡ��ʾ��ʽ", glngSys, 1341, 0))
        .int��ѯδ��ҩ�������� = Val(zldatabase.GetPara("��ѯδ��ҩ��������", glngSys, 1341, 0))
        
        '�������ã�����
        .intУ�鷢ҩ�� = Val(zldatabase.GetPara("У�鷢ҩ��", glngSys, 1341))
        .intУ����ҩ�� = Val(zldatabase.GetPara("У����ҩ��", glngSys, 1341))
        .IntShowCol = Val(zldatabase.GetPara("��ʾ����", glngSys, 1341))
        .bln��ʾ��С��λ = (Val(zldatabase.GetPara("��ʾ��С��λ", glngSys, 1341)) = 1)
        .int�Զ����� = Val(zldatabase.GetPara("�Զ�����", glngSys, 1341))
        .bln��ҩ��ˢ����֤ = (Val(zldatabase.GetPara("��ҩ��ˢ����֤", glngSys, 1341)) = 1)
        .bln��ҩɨ�� = (Val(zldatabase.GetPara("��ҩģʽɨ����ȷ��", glngSys, 1341)) = 1)
        .intOverTime = Val(zldatabase.GetPara("��ʱδ��ҩƷ��ʾʱ����", glngSys, 1341, 0))
        .intType = Val(zldatabase.GetPara("������סԺ����", glngSys, 1341, 0))
        .str����ˢ����ҩ = zldatabase.GetPara("����ˢ����ҩ", glngSys, 1341, "")
        .bln����ʱ����� = (Val(zldatabase.GetPara("ҩƷҽ��������ʱ�����", glngSys, 1341, 0)) = 1)
        .int�����ʾ = Val(zldatabase.GetPara("�����ʾ��ʽ", glngSys, 1341, 0))
        .blnȡҩȷ�� = (Val(zldatabase.GetPara("���ò���ʵ��ȡҩȷ��ģʽ", glngSys, 1341, 0)) = 1)
        .bln��ҩ���� = (Val(zldatabase.GetPara("��ҩ�������ķ������", glngSys, 1341, 0)) = 1)
        .blnɨ������ = (Val(zldatabase.GetPara("����ҩ����ɨ����Զ�����", glngSys, 1341, 0)) = 1)
        .bln��ҩ�շ� = (Val(zldatabase.GetPara("��ҩʱ��δ�շѵĵ��ݽ����շ�", glngSys, 1341, 0)) = 1)
        .int�س���ʽ = Val(zldatabase.GetPara("����ʱϵͳ�Զ��س���ʽ", glngSys, 1341, 0))
         
        '�������ã���ӡ
        .intPrint = Val(zldatabase.GetPara("�����µ����Ƿ��ӡ", glngSys, 1341))
        .intPrintDrugLable = Val(zldatabase.GetPara("��ӡҩƷ��ǩ", glngSys, 1341))
        .int��ӡ���ķ����嵥 = Val(zldatabase.GetPara("��ӡ���ķ��ϵ�", glngSys, 1341))
        .bln���ʵ� = (Val(zldatabase.GetPara("��ӡ�������ʵ�", glngSys, 1341)) = 1)
        .strPrintWindow = Replace(zldatabase.GetPara("��ӡָ����ҩ����", glngSys, 1341), "'", "")
        
        .lngPrintInterval = Val(zldatabase.GetPara("��ӡ���", glngSys, 1341))
        .lngRefreshInterval = Val(zldatabase.GetPara("ˢ�¼��", glngSys, 1341))
        .lngPrintDelay = Val(zldatabase.GetPara("��ӡ�ӳ�", glngSys, 1341, 60))
        .lngPrintBackInterval = Val(zldatabase.GetPara("��ӡ�˷ѵ��ݼ��", glngSys, 1341))
        .blnSign = Val(zldatabase.GetPara("ǩ��ʱ������ҩ", glngSys, 1341))
        .bln��ӡ���и�ʽ = (Val(zldatabase.GetPara("��ӡƱ�ݵ����и�ʽ", glngSys, 1341, 0)) = 1)
        .blnPreview = (Val(zldatabase.GetPara("��ӡ����ǩʱ��Ԥ���ٴ�ӡ", glngSys, 1341, 0)) = 1)
        
        '�������ã���Դ����
        .strSourceDep = zldatabase.GetPara("��Դ����", glngSys, 1341)
        
        '�������ã�������ɫ
        .strUserRecipeColor = zldatabase.GetPara("������ɫ", glngSys, 1341)
        If .strUserRecipeColor = "" Then .strUserRecipeColor = GetDefaultRecipeColor
        
        '��ӡ���б�
        .strPrinters = zldatabase.GetPara("������Ӧ�Ĵ�ӡ��", glngSys, 1341)
        
        '��ҩ���ʹ���ǩָ���Ĵ�ӡ��ʽ
        .str��ҩ��ʽ = zldatabase.GetPara("��ҩ����ӡ��ʽ", glngSys, 1341, "2;2")
        .str������ʽ = zldatabase.GetPara("����ǩ��ӡ��ʽ", glngSys, 1341, "1;1")
        
        '��������
        .int��ʾ�������� = Val(zldatabase.GetPara("��ʾ��������", glngSys, 1341))
        strTmp = zldatabase.GetPara("������", glngSys, 1341, "0")
        .intFont = Val(zldatabase.GetPara("����", glngSys, 1341))
        
        'ȡ��ҩƷ���Ƶĸ�ʽ��ʽ
        If strTmp = "" Then strTmp = "0"
'        .str������ = "0|ҩƷ����,0|������,0|Ӣ����,0|���,0|����,0|��λ,0|����,0|����,0|���,0|����,0|�÷�,0|Ƶ��,0|����,0|�����,0|�ⷿ��λ,0|������,0|׼����,0|��ҩ��,0|��ע"
        If InStr(1, strTmp, "|") > 0 Then
            .intҩƷ������ʾ = Val(Mid(strTmp, 1, 1))
        Else
            .intҩƷ������ʾ = Val(strTmp)
        End If
        
        '�Ŷӽк���ز���
        .blnStartQueue = (Val(zldatabase.GetPara("�����Ŷӽк�", glngSys, 1341, 0, Null, True, intParaType, .lngҩ��ID)) = 1)
        .intSoundType = Val(zldatabase.GetPara("��������", glngSys, 1341, 0, Null, True, intParaType, .lngҩ��ID))
        .blnShowQueue = (Val(zldatabase.GetPara("��ʾ�ŶӶ���", glngSys, 1341, 1, Null, True, intParaType, .lngҩ��ID)) = 1)
        .blnStartCall = (Val(zldatabase.GetPara("������������", glngSys, 1341, 1, Null, True, intParaType, .lngҩ��ID)) = 1)
        .intCallType = Val(zldatabase.GetPara("�кŷ�ʽ", glngSys, 1341, 0, Null, True, intParaType, .lngҩ��ID))
        .strRemoteCall = zldatabase.GetPara("Զ�˺���վ��", glngSys, 1341, "", Null, True, intParaType, .lngҩ��ID)
        .intSoundSpeed = Val(zldatabase.GetPara("�����㲥����", glngSys, 1341, 65, Null, True, intParaType, .lngҩ��ID))
        .intSoundTimes = Val(zldatabase.GetPara("�������Ŵ���", glngSys, 1341, 1, Null, True, intParaType, .lngҩ��ID))
        .lngShowComponent = Val(zldatabase.GetPara("��ʾ�豸���", glngSys, 1341, 101, Null, True, intParaType, .lngҩ��ID))
        .intCircleTime = Val(zldatabase.GetPara("������ѯʱ��", glngSys, 1341, 5, Null, True, intParaType, .lngҩ��ID))
        
        'ȡ����ĸ�ʽ���ƣ�Ĭ��ȡ��һ����ʽ��
        If mstrRPTDefaultScheme_Recipt = "" Then
            Set rsData = DeptSendWork_Get��ҩ����ʽ("ZL1_BILL_1341_3")
            If Not rsData.EOF Then
                mstrRPTDefaultScheme_Recipt = rsData!��ʽ
                rsData.MoveNext
            End If
            If Not rsData.EOF Then mstrRPTScheme_��ҩ�� = rsData!��ʽ
            If rsData.RecordCount >= 3 Then
                rsData.MoveNext
                For i = 3 To rsData.RecordCount
                    mstrRPTScheme_������ʽ = mstrRPTScheme_������ʽ & IIf(mstrRPTScheme_������ʽ = "", "", ";") & rsData!��ʽ
                    rsData.MoveNext
                Next
            End If
        End If
        'Ĭ�ϵ���ҩ����ǩ��ӡ����������ǰ�İ汾�����δӲ�ͬ��λ��ȡֵ
        If mstrRPTDefaultScheme_Recipt <> "" Then .strDefaultPrinter = GetSetting("ZLSOFT", "˽��ģ��\zl9Report\LocalSet\ZL1_BILL_1341_3\" & mstrRPTDefaultScheme_Recipt, "Printer")
        If .strDefaultPrinter = "" Then .strDefaultPrinter = GetSetting("ZLSOFT", "˽��ģ��\zl9Report\LocalSet\ZL1_BILL_1341_3\���и�ʽ", "Printer")
        If .strDefaultPrinter = "" Then .strDefaultPrinter = GetSetting("ZLSOFT", "˽��ģ��\zl9Report\LocalSet\ZL1_BILL_1341_3", "Printer")
        If .strDefaultPrinter = "" Then .strDefaultPrinter = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\zl9Report\LocalSet\ZL1_BILL_1341_3", "Printer")
        
        '��������
        .IntCheckStock = MediWork_GetCheckStockRule(.lngҩ��ID)
        
        '�Ƿ���Ҫ��ҩ����
        .blnMustDosageProcess = RecipeSendWork_DispensingMedi(.lngҩ��ID, bln�Ƿ���ҩȷ��)
        '�Ƿ���Ҫ��ҩȷ�Ϲ���
        .blnMustDosageOkProcess = bln�Ƿ���ҩȷ��
        
        'PASS
        If gintPass <> 0 And zlStr.IsHavePrivs(mstrPrivs, "������ҩ���") Then
            .blnStarPass = True
        End If
        
        mstr���� = .Str����
        If .blnStartQueue = True And .blnStartCall = True And .Str���� <> "" Then
            GetChildWin
        End If
        
        'վ��
        mstrDeptNode = GetDeptStationNode(.lngҩ��ID)
        
        Load����
    End With
End Sub
Private Sub GetPrivs()
    Dim strPrivs As String
    
    With mPrives
        .bln����ҩ�� = IsInString(mstrPrivs, "����ҩ��", ";")
        .bln��ҩ = IsInString(mstrPrivs, "��ҩ", ";")
        .bln��ҩ = IsInString(mstrPrivs, "��ҩ", ";")
        .bln������ҩ���Ĵ��� = IsInString(mstrPrivs, "������ҩ���Ĵ���", ";")
        .bln���ѽ��ʴ��� = IsInString(mstrPrivs, "���ѽ��ʴ���", ";")
        .bln���ѽ��ʴ��� = IsInString(mstrPrivs, "���ѽ��ʴ���", ";")
        .bln���˳�Ժ���˴��� = IsInString(mstrPrivs, "���˳�Ժ���˴���", ";")
        .blnУ�鴦�� = IsInString(mstrPrivs, "У�鴦��", ";")
        .blnҽ����ѯ = IsInString(mstrPrivs, "ҽ����ѯ", ";")
        .bln������ҩ��� = IsInString(mstrPrivs, "������ҩ���", ";")
        .bln���˸������� = IsInString(mstrPrivs, "���˸�������", ";")
        .bln�޸Ĺ������� = IsInString(mstrPrivs, "�޸Ĺ�������", ";")
        .bln�������� = IsInString(mstrPrivs, "��������", ";")
        .bln������ҩ���Ĵ��� = IsInString(mstrPrivs, "������ҩ���Ĵ���", ";")
        .bln���������� = IsInString(mstrPrivs, "����������", ";")
        .bln��ҩ = IsInString(mstrPrivs, "��ҩ", ";")
        .blnֹͣ��ҩ = IsInString(mstrPrivs, "ֹͣ��ҩ", ";")
        .bln�ָ���ҩ = IsInString(mstrPrivs, "�ָ���ҩ", ";")
        .blnȡҩȷ�� = IsInString(mstrPrivs, "ȡҩȷ��", ";")
        .bln���Ӳ������� = IsInString(mstrPrivs, "���Ӳ�������", ";")
        .bln�����ѯ����ʱ�䷶Χ���� = IsInString(mstrPrivs, "�����ѯ����ʱ�䷶Χ����", ";")
        
        'ҩƷ�Զ����豸�ӿڣ�����ģ�飩
        strPrivs = ";" & zl9ComLib.GetPrivFunc(glngSys, 9010) & ";"
        .blnҩƷ�Զ����ӿ� = IsInString(strPrivs, "����", ";")
    End With

End Sub


Private Sub Loadʱ�䷶Χ()
    With cboʱ�䷶Χ
        .Clear
        .AddItem "0-����"
        .AddItem "1-������"
        .AddItem "2-������"
        .AddItem "3-ָ��ʱ�䷶Χ"
        
        .ListIndex = 0
        .Tag = 0
    End With
End Sub

Private Sub InitApplyforcredit()
    '������������ļ�¼��
    Set mrsApplyforcredit = New ADODB.Recordset
    With mrsApplyforcredit
        If .State = 1 Then .Close
        
        .Fields.Append "����ID", adDouble, 18, adFldIsNullable
        .Fields.Append "�շ�ID", adDouble, 18, adFldIsNullable              'ҩƷ�շ�ID
        .Fields.Append "��־", adDouble, 1, adFldIsNullable      '0-������õ��ݷ�ҩ��1-����õ��ݷ�ҩ
        .Fields.Append "NO", adLongVarChar, 8, adFldIsNullable
        .Fields.Append "ҩƷ����", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "����", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "����", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "������������", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "����", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "�Ա�", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "����", adLongVarChar, 10, adFldIsNullable
        
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
End Sub

Private Sub InitPanes()
    Dim lngHeight As Long
    
    '��ʼ�������ؼ�
    'DockingPane
    '-----------------------------------------------------
    Me.dkpMain.SetCommandBars Me.cbsMain
    Me.dkpMain.Options.UseSplitterTracker = False 'ʵʱ�϶�
    Me.dkpMain.Options.ThemedFloatingFrames = True
    Me.dkpMain.Options.AlphaDockingContext = True
'    Me.dkpMain.Options.DefaultPaneOptions = PaneNoCloseable + PaneNoFloatable + PaneNoHideable + PaneNoCaption
    
    Dim objPaneCon As Pane
    Dim objPaneList As Pane
    Dim objPaneDetail As Pane
    
    lngHeight = 145
    
    If cboʱ�䷶Χ.ListIndex <> 3 Then
        lngHeight = lngHeight - 55
    End If
    
    If lbl����.Visible = False Then
        lngHeight = lngHeight - 25
    End If
    
    Set objPaneCon = Me.dkpMain.CreatePane(mconPane_Recipe_Condition, 230, lngHeight, DockLeftOf, Nothing)
    objPaneCon.Title = "��������"
    objPaneCon.Options = PaneNoCloseable Or PaneNoFloatable
    If Val(zldatabase.GetPara("ʹ�ø��Ի����")) = 0 Then objPaneCon.Hidden = False
End Sub
Private Sub InitComandBars()
    '��ʼ���˵�������ȫ���˵����������������˵���
    Dim cbrControlMain As CommandBarControl
    Dim cbrControl As CommandBarControl
    Dim cbrMenuBar As CommandBarPopup
    Dim objPopup As CommandBarPopup
    Dim cbrToolBar As CommandBar
    Dim objPane As Pane
    Dim blnGroup As Boolean
    Dim intCount As Integer
    Dim strCardName As String
    
    'CommandBars
    '-----------------------------------------------------
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    
    Me.cbsMain.VisualTheme = xtpThemeOffice2003

    With Me.cbsMain.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '����VisualTheme����Ч
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16

    End With
    
    Me.cbsMain.EnableCustomization False
    Me.cbsMain.Icons = frmPublic.imgPublic.Icons
    
    '-----------------------------------------------------
    '�˵�����
    Me.cbsMain.ActiveMenuBar.Title = "�˵�"
    Me.cbsMain.ActiveMenuBar.EnableDocking (xtpFlagStretched)
    
    Set cbrMenuBar = Me.cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, mconMenu_FilePopup, "�ļ�(&F)", -1, False)
    cbrMenuBar.Id = mconMenu_FilePopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_File_PrintSet, "��ӡ����(&S)��")
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_File_Preview, "Ԥ��(&V)")
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_File_Print, "��ӡ(&P)")
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_File_Excel, "�����&Excel��")
        
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_File_Recipe_BillPrintDosage, "��ӡ��ҩ��(&B)")
        cbrControlMain.BeginGroup = True
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_File_Recipe_BillPrintRecipe, "��ӡ����ǩ(&D)")
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_File_Recipe_BillPrintReport, "��ӡ��ҩ�嵥(&W)")
        If InStr(1, mstrPrivs, "��ӡ�ѷ�ҩ�嵥") > 0 Then
            cbrControlMain.Visible = True
        Else
            cbrControlMain.Visible = False
        End If
        
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_File_Recipe_BillPrintReturn, "��ӡ��ҩ֪ͨ��(&R)")
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_File_Recipe_BillPrintLable, "��ӡҩƷ��ǩ(&L)")
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_File_Recipe_BillPrintBack, "��ӡ�˷ѵ���(T)")
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_File_Recipe_BillPrintChange, "��ӡҽ������֪ͨ��(C)")
        
        If InStr(1, mstrPrivs, "��ӡ���˷ѵ���") > 0 Then
            cbrControlMain.Visible = True
        Else
            cbrControlMain.Visible = False
        End If
        
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_File_Parameter, "��������(&T)")
        cbrControlMain.BeginGroup = True
        
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_File_Exit, "�˳�(&X)")
        cbrControlMain.BeginGroup = True
    End With
    
    Set cbrMenuBar = Me.cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, mconMenu_EditPopup, "�༭(&E)", -1, False)
    cbrMenuBar.Id = mconMenu_EditPopup
    With cbrMenuBar.CommandBar.Controls
'        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Edit_Recipe_Dosage, "��ҩģʽ(&D)")
'        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Edit_Recipe_Abolish, "ȡ��ģʽ(&A)")
'        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Edit_Recipe_Send, "��ҩģʽ(&C)")
'        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Edit_Recipe_Return, "��ҩģʽ(&H)")
        
        
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Edit_Recipe_Batch, "������ҩ(&B)")
'        cbrControlMain.BeginGroup = True
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Edit_Recipe_SendOther, "������ҩ���Ĵ���(&F)")
        If InStr(1, mstrPrivs, "������ҩ���Ĵ���") > 0 Then
            cbrControlMain.Visible = True
        Else
            cbrControlMain.Visible = False
        End If
        
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Edit_Recipe_ReturnBatch, "������ҩ���Ĵ���(&T)")
        If InStr(1, mstrPrivs, "������ҩ���Ĵ���") > 0 Then
            cbrControlMain.Visible = True
        Else
            cbrControlMain.Visible = False
        End If
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Edit_Recipe_SendByBill, "��Ʊ�ݺŷ�ҩ(&I)")
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Edit_Recipe_ReturnByBill, "��Ʊ�ݺ���ҩ(&R)")
        
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Edit_Recipe_Flag, "ֹͣ��ҩ���(&S)")
        cbrControlMain.Visible = (mPrives.blnֹͣ��ҩ = True Or mPrives.bln�ָ���ҩ = True)
        blnGroup = (mPrives.blnֹͣ��ҩ = True Or mPrives.bln�ָ���ҩ = True)
        cbrControlMain.BeginGroup = blnGroup
        
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Edit_Recipe_Charge, "���ﻮ��(&M)")
        cbrControlMain.Visible = IsHavePrivs(mstrChargePrivs, "����")
        cbrControlMain.BeginGroup = Not blnGroup
        
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Edit_Recipe_Stuff, "���ķ���(&W)")
        cbrControlMain.Visible = IsHavePrivs(mstrStuffPrivs, "�������Ϸ���")
        
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Edit_Recipe_TakeDrug, "ȡҩȷ��(&T)")
        cbrControlMain.Visible = (mParams.blnȡҩȷ�� And mPrives.blnȡҩȷ��)
        
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Edit_Recipe_Change, "�л���ҩ��(&E)")
        cbrControlMain.BeginGroup = True
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Edit_Recipe_Windows, "������ҩ����(&N)")
        
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Edit_Recipe_Call, "����(&G)")
        
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Edit_Recipe_Cancle, "ȡ��ȷ��(&G)")
        cbrControlMain.BeginGroup = True
        
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Edit_Recipe_SendHot, "��ҩ")
        cbrControlMain.BeginGroup = True
        cbrControlMain.Visible = False
        
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Edit_Recipe_VerifySign, "��֤ǩ��(&S)")
        cbrControlMain.Visible = gblnESign������ҩ
                        
'        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Edit_Recipe_EMR, "������ѯ(&L)")
'        cbrControlMain.BeginGroup = True
        
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Edit_Recipe_Hot_IC, "��IC��(&I)")
        cbrControlMain.Visible = False
        
        If TypeName(mobjDrugMAC) = "clsDrugPacker" Then
            Set cbrControlMain = .Add(xtpControlPopup, mconMenu_Edit_Recipe_AutoSend, "�����Զ���ҩ������")
            
            cbrControlMain.CommandBar.Controls.Add(xtpControlButton, mconMenu_Edit_Recipe_AutoSend_Open, "���ô����ϴ�").Checked = True
            cbrControlMain.CommandBar.Controls.Add xtpControlButton, mconMenu_Edit_Recipe_AutoSend_Set, "����WebService����ĵ�ַ"
            cbrControlMain.CommandBar.Controls.Add xtpControlButton, mconMenu_Edit_Recipe_AutoSend_LoadDrug, "�ϴ�ҩƷ��������"
            cbrControlMain.CommandBar.Controls.Add xtpControlButton, mconMenu_Edit_Recipe_AutoSend_LoadStock, "�ϴ�ҩƷ�������"
            mblnLoadDrug = True
        End If
        
        '��Ҳ�������չ����
        Call zlPlugIn_SetMenu(glngSys, glngModul, mobjPlugIn, cbrMenuBar.CommandBar.Controls, mconMenu_Edit_PlugIn)
    End With
    
'    '�Զ�����ҩ���ò˵�
'    If Not gobjPackerMZ Is Nothing Then
'        Set cbrMenuBar = Me.cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, mconMenu_AutoSend, "ҩ���Զ����ӿ�(&V)", -1, False)
'        cbrMenuBar.Id = mconMenu_AutoSend
'    End If
        
    Set cbrMenuBar = Me.cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, mconMenu_ViewPopup, "�鿴(&V)", -1, False)
    cbrMenuBar.Id = mconMenu_ViewPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControlMain = .Add(xtpControlPopup, mconMenu_View_ToolBar, "������(&T)")
        Set cbrControl = cbrControlMain.CommandBar.Controls.Add(xtpControlButton, mconMenu_View_ToolBar_Button, "��׼��ť(&S)", -1, False)
        cbrControl.Checked = True
        Set cbrControl = cbrControlMain.CommandBar.Controls.Add(xtpControlButton, mconMenu_View_ToolBar_Text, "�ı���ǩ(&T)", -1, False)
        cbrControl.Checked = True
        Set cbrControl = cbrControlMain.CommandBar.Controls.Add(xtpControlButton, mconMenu_View_ToolBar_Size, "��ͼ��(&B)", -1, False)
        cbrControl.Checked = True
        
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_View_StatusBar, "״̬��(&S)")
        cbrControlMain.Checked = True
        
        Set cbrControlMain = .Add(xtpControlPopup, mconMenu_View_FontSize, "����(&F)")
        cbrControlMain.BeginGroup = True
        Set cbrControl = cbrControlMain.CommandBar.Controls.Add(xtpControlButton, mconMenu_View_FontSize_1, "С����(&S)", -1, False)
        If mParams.intFont = 0 Then cbrControl.Checked = True
        cbrControl.Parameter = 0
        Set cbrControl = cbrControlMain.CommandBar.Controls.Add(xtpControlButton, mconMenu_View_FontSize_2, "������(&M)", -1, False)
        If mParams.intFont = 1 Then cbrControl.Checked = True
        cbrControl.Parameter = 1
        Set cbrControl = cbrControlMain.CommandBar.Controls.Add(xtpControlButton, mconMenu_View_FontSize_3, "������(&B)", -1, False)
        If mParams.intFont = 2 Then cbrControl.Checked = True
        cbrControl.Parameter = 2
        
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_View_Filter, "����(&F)")
        cbrControlMain.BeginGroup = True
        
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_View_Refresh, "ˢ��(&R)")
        cbrControlMain.BeginGroup = True
    End With
    
    Set cbrMenuBar = Me.cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, mconMenu_HelpPopup, "����(&H)", -1, False)
    cbrMenuBar.Id = mconMenu_HelpPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Help_Help, "��������(&H)")
        Set cbrControlMain = .Add(xtpControlPopup, mconMenu_Help_Web, "&WEB�ϵ�" & gstrProductName)
        cbrControlMain.CommandBar.Controls.Add xtpControlButton, mconMenu_Help_Web_Home, gstrProductName & "��ҳ(&H)", -1, False
        cbrControlMain.CommandBar.Controls.Add xtpControlButton, mconMenu_Help_Web_Forum, "���ͷ���(&M)", -1, False
        cbrControlMain.CommandBar.Controls.Add xtpControlButton, mconMenu_Help_Web_Mail, "���ͷ���(&M)", -1, False
        
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Help_About, "����(&A)��")
        cbrControlMain.BeginGroup = True
    End With
    
    '�����
    With Me.cbsMain.KeyBindings
        .Add FCONTROL, Asc("P"), mconMenu_File_Print
        .Add FCONTROL, Asc("D"), mconMenu_Edit_Recipe_Dosage
        .Add FCONTROL, Asc("A"), mconMenu_Edit_Recipe_Abolish
        .Add FCONTROL, Asc("C"), mconMenu_Edit_Recipe_Send
        .Add FCONTROL, Asc("H"), mconMenu_Edit_Recipe_Return
        .Add FCONTROL, Asc("Q"), mconMenu_Edit_Recipe_Cancel
    
        .Add FCONTROL, VK_F4, mconMenu_Edit_Recipe_Hot_IC
        
        .Add 0, VK_F12, mconMenu_File_Parameter
        .Add 0, VK_F5, mconMenu_View_Refresh
        .Add 0, VK_F1, mconMenu_Help_Help
        
        .Add 0, VK_F2, mconMenu_Edit_Recipe_SendHot
        
        .Add 0, VK_F6, mconMenu_File_Recipe_BillPrintDosage
        .Add 0, VK_F4, mconMenu_File_Recipe_BillPrintRecipe
        .Add 0, VK_F11, mconMenu_File_Recipe_BillPrintLable
        .Add 0, VK_F8, mconMenu_Edit_Recipe_Charge
        .Add 0, VK_F9, mconMenu_Edit_Recipe_Stuff
        .Add 0, VK_F12, mconMenu_File_Parameter
        .Add 0, VK_F7, mconMenu_View_Filter
        .Add 0, VK_F3, mconMenu_Edit_Recipe_Call
    End With

'    '���ò����ò˵�
'    With Me.cbsMain.Options
'        .AddHiddenCommand conMenu_File_PrintSet
'        .AddHiddenCommand conMenu_File_Excel
'        .AddHiddenCommand conMenu_View_Refresh
'    End With
    
    '���õ����˵�
    Set cbrMenuBar = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, mconMenu_InputPopup, "¼��(&I)", -1, False)
    cbrMenuBar.Id = mconMenu_InputPopup
    cbrMenuBar.Visible = False
    With cbrMenuBar.CommandBar.Controls
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Input_Recipe_NO, "���ݺ�(&0)")
        cbrControlMain.Parameter = "��|���ݺ�|0||||||"
        cbrControlMain.Checked = True
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Input_Recipe_OPNO, "�����(&1)")
        cbrControlMain.Parameter = "��|�����|0|||||"
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Input_Recipe_Name, "����(&2)")
        cbrControlMain.Parameter = "��|����|0|||||"
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Input_Recipe_IDCard, "���֤(&3)")
        cbrControlMain.Parameter = "��|���֤|0|||||"
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Input_Recipe_ICCard, "IC��(&4)")
        cbrControlMain.Parameter = "IC|IC����|1|||||"
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Input_Recipe_MINo, "ҽ����(&5)")
        cbrControlMain.Parameter = "ҽ|ҽ����|0|||||"
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Input_Recipe_HosNumber, "סԺ��(&6)")
        cbrControlMain.Parameter = "ס|סԺ��|0|||||"
        
        '��̬ȡ����ҽ�ƿ�����Ҫ�����ѿ���
        If mstrCardType <> "" Then
            mintCardCount = UBound(Split(mstrCardType, ";")) + 1
            For intCount = 0 To UBound(Split(mstrCardType, ";"))
                'ȡ���ѿ�����
                strCardName = Split(Split(mstrCardType, ";")(intCount), "|")(1)
                
                Set cbrControlMain = .Add(xtpControlButton, mconMenu_Input_Recipe_HosNumber + intCount + 1, strCardName & "(&" & intCount + 7 & ")")
                
                '���濨��Ϣ
                cbrControlMain.Parameter = Split(mstrCardType, ";")(intCount)
                
                If intCount = 0 Then
                    cbrControlMain.BeginGroup = True
                End If
                
                If Split(cbrControlMain.Parameter, "|")(gCardFormat.����) = "��" Then
                    mint���￨���� = Val(Split(cbrControlMain.Parameter, "|")(gCardFormat.���ų���))
                End If
            Next
        End If
    End With
    
    '-----------------------------------------------------
    '����������
    Set cbrToolBar = Me.cbsMain.Add("������", xtpBarTop)
    cbrToolBar.ShowTextBelowIcons = False
    cbrToolBar.EnableDocking xtpFlagStretched
    With cbrToolBar.Controls
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_File_Preview, "Ԥ��")
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_File_Print, "��ӡ")
        
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_View_Filter, "����")
        cbrControlMain.BeginGroup = True
        
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Edit_Recipe_Charge, "����")
        cbrControlMain.Visible = IsHavePrivs(mstrChargePrivs, "����")
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Edit_Recipe_Stuff, "����")
        cbrControlMain.Visible = IsHavePrivs(mstrStuffPrivs, "�������Ϸ���")
        
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Edit_Recipe_TakeDrug, "ȡҩ")
        cbrControlMain.Visible = (mParams.blnȡҩȷ�� And mPrives.blnȡҩȷ��)
                
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Edit_Recipe_Call, "����")
        
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Edit_Recipe_Cancle, "ȡ��ȷ��")
        
'        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Edit_Recipe_EMR, "������ѯ")
'        cbrControlMain.BeginGroup = True
        
'        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Edit_Recipe_AddSign, "ǩ��")
'        cbrControlMain.Visible = gblnҩƷʹ�õ���ǩ��
        
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Edit_Recipe_VerifySign, "��֤ǩ��")
        cbrControlMain.Visible = gblnESign������ҩ
        
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_View_Refresh, "ˢ��")
        cbrControlMain.BeginGroup = True
        
        '���Ӳ�������
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Edit_Recipe_MedicalRecord, "���Ӳ�������")
        cbrControlMain.BeginGroup = True
        cbrControlMain.Visible = mPrives.bln���Ӳ�������
        
        '��Ҳ�������չ����
        Call zlPlugIn_SetToolbar(glngSys, glngModul, mobjPlugIn, cbrToolBar.Controls, mconMenu_Edit_PlugIn)
        
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Help_Help, "����")
        cbrControlMain.BeginGroup = True
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_File_Exit, "�˳�")
        
    End With
    For Each cbrControlMain In cbrToolBar.Controls
        cbrControlMain.Style = xtpButtonIconAndCaption
    Next
End Sub


Private Function RecipeWork_SendByBatch(ByVal intListType As Integer) As Boolean
    Dim rsBatchData As ADODB.Recordset
    Dim rsData As ADODB.Recordset
    Dim strRecipeString As String
    Dim intCount As Integer
    Dim n As Integer
    Dim arrRecipe
    Dim intBillType As Integer
    Dim strNo As String
    Dim int��¼���� As Integer
    Dim int�����־ As Integer
    
    strRecipeString = mfrmList.GetCurrentBatchRecipe
    
    If strRecipeString = "" Then Exit Function
    
    Set rsBatchData = New ADODB.Recordset
    With rsBatchData
        If .State = 1 Then .Close
        .Fields.Append "��־", adLongVarChar, 1, adFldIsNullable
        .Fields.Append "����", adLongVarChar, 8, adFldIsNullable
        .Fields.Append "NO", adLongVarChar, 8, adFldIsNullable
        .Fields.Append "�շ�ID", adDouble, 18, adFldIsNullable
        .Fields.Append "ҩƷID", adDouble, 18, adFldIsNullable
        .Fields.Append "����", adDouble, 18, adFldIsNullable
        .Fields.Append "���", adLongVarChar, 8, adFldIsNullable
        .Fields.Append "Ʒ��", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "ʵ�ս��", adDouble, 18, adFldIsNullable
        .Fields.Append "��¼����", adLongVarChar, 8, adFldIsNullable
        .Fields.Append "��¼״̬", adLongVarChar, 8, adFldIsNullable
        .Fields.Append "����ID", adDouble, 18, adFldIsNullable
        .Fields.Append "���￨��", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "�����־", adLongVarChar, 8, adFldIsNullable
        .Fields.Append "ҩ��ID", adDouble, 18, adFldIsNullable
        .Fields.Append "��������", adDouble, 1, adFldIsNullable
        .Fields.Append "�շ����", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "����", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "���շ�", adDouble, 1, adFldIsNullable
        .Fields.Append "��������", adDate, , adFldIsNullable
        .Fields.Append "����ģʽ", adDouble, 1, adFldIsNullable
        
        .Fields.Append "ҩ��", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "����", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "����", adDouble, 50, adFldIsNullable
        .Fields.Append "��װ", adDouble, 50, adFldIsNullable
        .Fields.Append "��λ", adLongVarChar, 1, adFldIsNullable
        .Fields.Append "�Ա�", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "����", adLongVarChar, 10, adFldIsNullable
        
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
    
    arrRecipe = Split(strRecipeString, "|")
    intCount = UBound(arrRecipe)
    
    For n = 0 To intCount
        intBillType = Val(Split(arrRecipe(n), ",")(0))
        strNo = Split(arrRecipe(n), ",")(1)
        int��¼���� = Split(arrRecipe(n), ",")(5)
        int�����־ = Split(arrRecipe(n), ",")(6)
       
        Set rsData = GetRecipeRecord(intBillType, strNo, int�����־, int��¼����)
        
        If Not rsData Is Nothing Then
            With rsBatchData
                Do While Not rsData.EOF
                    .AddNew
                    !��־ = 1
                    !���� = rsData!����
                    !NO = rsData!NO
                    !�շ�ID = rsData!�շ�ID
                    !ҩƷID = rsData!ҩƷID
                    !���� = rsData!����
                    !��� = rsData!���
                    !Ʒ�� = rsData!Ʒ��
                    !ʵ�ս�� = rsData!ʵ�ս��
                    !��¼���� = rsData!��¼����
                    !��¼״̬ = rsData!��¼״̬
                    !����ID = rsData!����ID
                    !���￨�� = zlStr.NVL(rsData!���￨��, "")
                    !�����־ = rsData!�����־
                    !ҩ��ID = rsData!ҩ��ID
                    !�������� = rsData!��������
                    !�շ���� = rsData!�շ����
                    !���� = rsData!����
                    !���շ� = rsData!���շ�
                    !�������� = rsData!��������
                    !����ģʽ = rsData!����ģʽ
                    
                    !ҩ�� = rsData!ҩƷ����
                    !���� = rsData!����
                    !���� = rsData!���� * rsData!��װ
                    !��װ = rsData!��װ
                    !��λ = rsData!��λ
                    !�Ա� = rsData!�Ա�
                    !���� = rsData!����

                    .Update
                    
                    rsData.MoveNext
                Loop
            End With
        End If
    Next
    
    If intListType = mListType.����ҩ Then
        If RecipeWork_Send(rsBatchData) = False Then
            RecipeWork_SendByBatch = False
        Else
            If imgFilter.BorderStyle = cstFilter Then
                txtPati.Text = ""
                txtPati.SetFocus
            End If
        End If
    ElseIf intListType = mListType.����ҩ Then
        If RecipeWork_Dosage(rsBatchData) = False Then
            RecipeWork_SendByBatch = False
        Else
            If imgFilter.BorderStyle = cstFilter Then
                txtPati.Text = ""
                txtPati.SetFocus
            End If
        End If
    End If
End Function


Private Sub SetComandBars()
    '������Ĳ˵�״̬
    Dim cbrControl As CommandBarControl
    Dim cbrMenu As CommandBarControl
    
    On Error GoTo errHandle
    
    If mParams.blnȡҩȷ�� And mPrives.blnȡҩȷ�� Then
        Set cbrMenu = Me.cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_Edit_Recipe_TakeDrug, , True)
        Set cbrControl = Me.cbsMain(2).Controls.Find(xtpControlButton, mconMenu_Edit_Recipe_TakeDrug, , True)
        
        If Not cbrMenu Is Nothing Then cbrMenu.Visible = (mcondition.intListType = mListType.��ҩ)
        If Not cbrControl Is Nothing Then cbrControl.Visible = (mcondition.intListType = mListType.��ҩ)
    End If
    
    If gblnESign������ҩ = True Then
'        Set cbrMenu = Me.cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_Edit_Recipe_AddSign, , True)
'        Set cbrControl = Me.cbsMain(2).Controls.Find(xtpControlButton, mconMenu_Edit_Recipe_AddSign, , True)
'
'        If mcondition.intListType <> mListType.��ҩ Then
'            If Not cbrMenu Is Nothing Then cbrMenu.Enabled = False
'            If Not cbrControl Is Nothing Then cbrControl.Enabled = False
'        End If
        
        Set cbrMenu = Me.cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_Edit_Recipe_VerifySign, , True)
        Set cbrControl = Me.cbsMain(2).Controls.Find(xtpControlButton, mconMenu_Edit_Recipe_VerifySign, , True)
        
        If Not cbrMenu Is Nothing Then cbrMenu.Visible = True
        If Not cbrControl Is Nothing Then cbrControl.Visible = True
        
        If Not cbrMenu Is Nothing Then cbrMenu.Enabled = False
        If Not cbrControl Is Nothing Then cbrControl.Enabled = False
    Else
        Set cbrMenu = Me.cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_Edit_Recipe_VerifySign, , True)
        Set cbrControl = Me.cbsMain(2).Controls.Find(xtpControlButton, mconMenu_Edit_Recipe_VerifySign, , True)
        
        If Not cbrMenu Is Nothing Then cbrMenu.Visible = False
        If Not cbrControl Is Nothing Then cbrControl.Visible = False
    End If
    
    Set cbrMenu = Me.cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_Edit_Recipe_Windows, , True)
    If Not cbrMenu Is Nothing Then
        If InStr(1, ";" & mstrPrivs & ";", ";������ҩ����;") < 1 Or mParams.Str���� = "" Then
            cbrMenu.Visible = False
        Else
            cbrMenu.Visible = True
        End If
    End If
    
    Set cbrMenu = Me.cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_Edit_Recipe_Call, , True)
    Set cbrControl = Me.cbsMain(2).Controls.Find(xtpControlButton, mconMenu_Edit_Recipe_Call, , True)
    If Not cbrMenu Is Nothing Then
        If mParams.blnStartQueue And mParams.blnStartCall And InStr(1, ";" & mstrPrivs & ";", ";�к�;") > 0 Then
            cbrMenu.Visible = True
        Else
            cbrMenu.Visible = False
        End If
        
        If tbcList.Selected.index = mListType.����ҩ Then
            cbrMenu.Enabled = True
        Else
            cbrMenu.Enabled = False
        End If
    End If
    
    If Not cbrControl Is Nothing Then
        If mParams.blnStartQueue And mParams.blnStartCall And InStr(1, ";" & mstrPrivs & ";", ";�к�;") > 0 Then
            cbrControl.Visible = True
        Else
            cbrControl.Visible = False
        End If
    
        If tbcList.Selected.index = mListType.����ҩ Then
            cbrControl.Enabled = True
        Else
            cbrControl.Enabled = False
        End If
        
    End If
    
'    Set cbrMenu = Me.cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_Edit_Recipe_SendOther, , True)
'    If Not cbrMenu Is Nothing Then
'        If InStr(1, mstrPrivs, "������ҩ���Ĵ���") > 0 Then
'            cbrMenu.Visible = True
'        Else
'            cbrMenu.Visible = False
'        End If
'    End If
'
'    Set cbrMenu = Me.cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_Edit_Recipe_ReturnBatch, , True)
'    If Not cbrMenu Is Nothing Then
'        If InStr(1, mstrPrivs, "������ҩ���Ĵ���") > 0 Then
'            cbrMenu.Visible = True
'        Else
'            cbrMenu.Visible = False
'        End If
'    End If
    
    Set cbrMenu = Me.cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_Edit_Recipe_Cancle, , True)
    Set cbrControl = Me.cbsMain(2).Controls.Find(xtpControlButton, mconMenu_Edit_Recipe_Cancle, , True)
    If tbcList.Selected.index = mListType.����ҩ And mParams.blnMustDosageProcess Then
        cbrControl.Visible = True
        cbrMenu.Visible = True
        cbrMenu.Caption = "ȡ����ҩ"
        cbrControl.Caption = "ȡ����ҩ"
    ElseIf tbcList.Selected.index = mListType.����ҩ And mParams.blnMustDosageOkProcess And InStr(1, ";" & mstrPrivs & ";", ";��ҩȷ��;") > 0 Then
        cbrControl.Visible = True
        cbrMenu.Visible = True
        cbrMenu.Caption = "ȡ��ȷ��"
        cbrControl.Caption = "ȡ��ȷ��"
    Else
        cbrControl.Visible = False
    End If
    
    If TypeName(mobjDrugMAC) = "clsDrugPacker" Then
        If mblnPackerConnect Then
            Set cbrMenu = Me.cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_Edit_Recipe_AutoSend_Open, , True)
            cbrMenu.Checked = True
            cbrMenu.Enabled = True
            mblnLoadDrug = True
            
            Set cbrMenu = Me.cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_Edit_Recipe_AutoSend_LoadDrug, , True)
            cbrMenu.Enabled = True
            
            Set cbrMenu = Me.cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_Edit_Recipe_AutoSend_LoadStock, , True)
            cbrMenu.Enabled = True
        Else
            Set cbrMenu = Me.cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_Edit_Recipe_AutoSend_Open, , True)
            cbrMenu.Checked = False
            cbrMenu.Enabled = False
            mblnLoadDrug = False
            
            Set cbrMenu = Me.cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_Edit_Recipe_AutoSend_LoadDrug, , True)
            cbrMenu.Enabled = False
            
            Set cbrMenu = Me.cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_Edit_Recipe_AutoSend_LoadStock, , True)
            cbrMenu.Enabled = False
        End If
    End If
    
    Exit Sub
errHandle:
    If ErrCenter <> 1 Then Resume
    Call SaveErrLog
End Sub

Public Function RecipeWork(ByVal intType As Integer, ByVal blnByNo As Boolean, Optional vsfDetail As VSFlexGrid, Optional blnδȡҩ��ҩ As Boolean = False) As Boolean
    Select Case intType
        Case mListType.��ҩȷ��
            If RecipeWork_DosageOk = False Then RecipeWork = False
        Case mListType.����ҩ
            If imgFilter.BorderStyle = cstFilter Then
                '������ҩ
                If RecipeWork_SendByBatch(mListType.����ҩ) = False Then RecipeWork = False
            Else
                If RecipeWork_Dosage(mfrmDetail.GetRecord) = False Then RecipeWork = False
            End If
        Case mListType.����ҩ
            If RecipeWork_Abolish = False Then RecipeWork = False
        Case mListType.����ҩ, mListType.��ʱδ��
            mblnδȡҩ��ҩ = blnδȡҩ��ҩ
            If imgFilter.BorderStyle = cstFilter And blnByNo = False Then
                '������ҩ
                mint��ҩ��ʽ = 0
                If RecipeWork_SendByBatch(mListType.����ҩ) = False Then RecipeWork = False
            Else
                mint��ҩ��ʽ = 1
                If RecipeWork_Send(mfrmDetail.GetRecord) = False Then RecipeWork = False
                
                
            End If
        Case mListType.��ҩ
            If RecipeWork_Return(vsfDetail) = False Then RecipeWork = False
    End Select
    
    RefreshList intType
    
    RecipeWork = True
    
    txtPati.SetFocus
    mstrScanerLastNo = ""
End Function

Private Function RecipeWork_TakeDrug() As Boolean
    '����ȡҩȷ��
    Dim blnInTrans As Boolean
    Dim str��ǰ���� As String
    Dim Int���� As Integer, strNo As String
    Dim strUnit As String
    Dim int���� As Integer
    Dim lngǩ��id As Long
    Dim inδȡҩ As Integer
    Dim date��ҩʱ�� As Date
    
    On Error GoTo errHandle
    
    If mcondition.intListType <> mListType.��ҩ Then Exit Function
    
    str��ǰ���� = mfrmList.GetCurrentRecipe
    
    If str��ǰ���� = "" Then Exit Function
    
    Int���� = Val(Split(str��ǰ����, "|")(0))
    strNo = Split(str��ǰ����, "|")(1)
    inδȡҩ = Val(Split(str��ǰ����, "|")(11))
    date��ҩʱ�� = Sys.Currentdate
    
    If inδȡҩ = 1 Then
        If MsgBox("�Ƿ񽫴���[" & strNo & "]���Ϊ������ȡҩ��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
    Else
        If MsgBox("�Ƿ񽫴���[" & strNo & "]���Ϊ����δȡҩ��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
    End If
    
    gstrSQL = "Zl_ҩƷ�շ���¼_ȷ��ȡҩ("
    '�ⷿID
    gstrSQL = gstrSQL & mParams.lngҩ��ID
    '����
    gstrSQL = gstrSQL & "," & Int����
    'NO
    gstrSQL = gstrSQL & ",'" & strNo & "'"
    '�Ƿ�δȡҩ
    gstrSQL = gstrSQL & "," & IIf(inδȡҩ = 1, "Null", 1)
    'ȡҩȷ����Ա
    gstrSQL = gstrSQL & ",'" & gstrUserName & "'"
    'ȡҩʱ��
    gstrSQL = gstrSQL & ",to_date('" & date��ҩʱ�� & "','yyyy-MM-dd hh24:mi:ss') "
    gstrSQL = gstrSQL & ")"

    Call zldatabase.ExecuteProcedure(gstrSQL, "RecipeWork_TakeDrug")

    RefreshList mcondition.intListType
    
    Exit Function
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Private Function RecipeWork_DosageOk() As Boolean
    '��ҩȷ��
    Dim str����Ա As String
    Dim str��ǰ���� As String
    Dim int�Ƿ�ȷ�� As Integer
    Dim rsData As ADODB.Recordset
    Dim arrSql As Variant
    Dim i As Integer
    
    On Error GoTo ErrHand
    
    If mfrmDetail.CmdSend.Caption = "��ҩȷ��(&O)" Then
        str����Ա = gstrUserName
'        If mParams.blnMustDosageProcess Then
        int�Ƿ�ȷ�� = 1
'        Else
'            int�Ƿ�ȷ�� = 2
'        End If
        
    End If
    Set rsData = mfrmDetail.GetRecord
    
    rsData.Filter = "��־=1"
    rsData.Sort = "����,No"
    arrSql = Array()
    
    
    Do While Not rsData.EOF
        If str��ǰ���� <> rsData!���� & "|" & rsData!NO Then
            str��ǰ���� = rsData!���� & "|" & rsData!NO
            
            '��鵥���Ƿ����
            If Not CheckBillExist(rsData!����, rsData!NO) Then
                MsgBox "δ�ҵ�ָ������,�����Ѿ�����������Ա����,����������ֹ��", vbInformation, gstrSysName
                Exit Function
            End If
        End If
        
        rsData.MoveNext
    Loop
    
    
    rsData.MoveFirst
    Do While Not rsData.EOF
        gstrSQL = "Zl_δ��ҩƷ��¼_��ҩȷ��("
            'NO
            gstrSQL = gstrSQL & "'" & rsData!NO & "'"
            '����
            gstrSQL = gstrSQL & "," & rsData!����
            '�ⷿID
            gstrSQL = gstrSQL & "," & mParams.lngҩ��ID
            '��ҩȷ��
            gstrSQL = gstrSQL & "," & int�Ƿ�ȷ��
            '����Ա
            gstrSQL = gstrSQL & ",'" & str����Ա & "')"
            
        ReDim Preserve arrSql(UBound(arrSql) + 1)
        arrSql(UBound(arrSql)) = gstrSQL
        rsData.MoveNext
    Loop
    gcnOracle.BeginTrans
    For i = 0 To UBound(arrSql)
        Call zldatabase.ExecuteProcedure(CStr(arrSql(i)), "RecipeWork_DosageOk")
    Next
    gcnOracle.CommitTrans
    RecipeWork_DosageOk = True
    Exit Function
ErrHand:
    gcnOracle.RollbackTrans
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function


Private Function RecipeWork_Dosage(ByVal rsData As ADODB.Recordset) As Boolean
    '��ҩ
    Dim blnInTrans As Boolean
    Dim str����Ա As String
    Dim str��ǰ���� As String
    Dim strDosUser As String
    Dim int���� As Integer
    Dim arrSql As Variant
    Dim i As Long
    Dim blnBeginTrans As Boolean
    Dim strǩ����¼ As String
    Dim date��ҩ���� As Date
    Dim strNosToPlugIn As String
    Dim strReserve As String
    
    On Error GoTo ErrHand
    
    mstrPrintRecipe = ""
    
    arrSql = Array()
    
    date��ҩ���� = Sys.Currentdate
    
    rsData.Filter = "��־=1"
    rsData.Sort = "����,No"
    
    Do While Not rsData.EOF
        If str��ǰ���� <> rsData!���� & "|" & rsData!NO Then
            str��ǰ���� = rsData!���� & "|" & rsData!NO
            
            '��鵥���Ƿ����
            If Not CheckBillExist(rsData!����, rsData!NO) Then
                MsgBox "δ�ҵ�ָ������,�����Ѿ�����������Ա����,����������ֹ��", vbInformation, gstrSysName
                Exit Function
            End If
        
            '����Ƿ�����
            If CheckBill(rsData!ҩ��ID, 1, rsData!����, rsData!NO, rsData!��¼����, rsData!�����־) <> 0 Then Exit Function
        End If
        
        rsData.MoveNext
    Loop
    
    '���õ���ǩ��ʱ����û��Ƿ�ע��
    If gblnESign������ҩ = True Then
        If Not gobjESign.CheckCertificate(gstrDbUser) Then
            Exit Function
        End If
    End If
        
    'У����ҩ�ˣ�������õ���ǩ����ʹ��
    If gblnESign������ҩ = False Then
        If mParams.intУ����ҩ�� = 1 Then
            str����Ա = zldatabase.UserIdentify(Me, "У����ҩ��", glngSys, 1341, "��ҩ")
        Else
            str����Ա = mParams.str��ҩ��
        End If
        If str����Ա = "" Then Exit Function
    End If
    
    If mParams.bln��ҩ�շ� And mParams.bln��ҩǰ�շѻ���� Then
        '�ϵ�һ��ͨ����ˢ��
        If CheckCard(rsData) = False Then Exit Function
        
        '�µ����ѿ�ˢ�����ѽӿ�
        If Not CardConfirm(rsData) Then Exit Function
    End If
    
    '�ȸ�������
    rsData.Filter = "��־=1"
    rsData.Sort = "����,No"
    
    If mParams.IntCheckStock = 2 Then
        Do While Not rsData.EOF
            gstrSQL = "zl_ҩƷ�շ���¼_��������("
            '�շ�ID
            gstrSQL = gstrSQL & rsData!�շ�ID
            'ҩƷID
            gstrSQL = gstrSQL & "," & rsData!ҩƷID
            '����
            gstrSQL = gstrSQL & "," & rsData!����
            gstrSQL = gstrSQL & ")"
            
            ReDim Preserve arrSql(UBound(arrSql) + 1)
            arrSql(UBound(arrSql)) = gstrSQL
            
            rsData.MoveNext
        Loop
    End If
    
    '��������ҩ��
    rsData.Filter = "��־=1"
    rsData.Sort = "����,No"
    str��ǰ���� = ""
    strDosUser = mfrmDetail.Get��ҩ��
    If strDosUser = "" Then strDosUser = IIf(mParams.str��ҩ�� = "|��ǰ����Ա|", gstrUserName, str����Ա)
    
    Do While Not rsData.EOF
        If str��ǰ���� <> rsData!���� & "|" & rsData!NO Then
            If Val(rsData!��¼����) = 1 Or (Val(rsData!��¼����) = 2 And (Val(rsData!�����־)) = 1 Or (Val(rsData!�����־)) = 4) Then
                int���� = 1
            Else
                int���� = 2
            End If
            
            If mPrives.bln������ҩ���Ĵ��� = True And mParams.lngҩ��ID <> Val(rsData!ҩ��ID) Then
                gstrSQL = "Zl_ҩƷ�շ���¼_���Ŀⷿ("
                '�ֿⷿID
                gstrSQL = gstrSQL & mParams.lngҩ��ID
                '����
                gstrSQL = gstrSQL & "," & rsData!����
                'NO
                gstrSQL = gstrSQL & ",'" & rsData!NO & "'"
                'ԭ�ⷿID
                gstrSQL = gstrSQL & "," & Val(rsData!ҩ��ID)
                '����
                gstrSQL = gstrSQL & "," & int����
                '��������
                gstrSQL = gstrSQL & ",to_date('" & rsData!�������� & "','yyyy-MM-dd')"
                gstrSQL = gstrSQL & ")"
                
                ReDim Preserve arrSql(UBound(arrSql) + 1)
                arrSql(UBound(arrSql)) = gstrSQL
            End If
        
            str��ǰ���� = rsData!���� & "|" & rsData!NO
            
            gstrSQL = "zl_ҩƷ�շ���¼_������ҩ��("
            '�ⷿID
            gstrSQL = gstrSQL & mParams.lngҩ��ID
            '����
            gstrSQL = gstrSQL & "," & rsData!����
            'NO
            gstrSQL = gstrSQL & ",'" & rsData!NO & "'"
            '����
            gstrSQL = gstrSQL & "," & int����
            '��ҩ��
            gstrSQL = gstrSQL & ",'" & IIf(gblnESign������ҩ = True, gstrUserName, IIf(mParams.intУ����ҩ�� = 1, str����Ա, strDosUser)) & "'"
            '��ҩ����
            gstrSQL = gstrSQL & ",to_date('" & date��ҩ���� & "','yyyy-MM-dd hh24:mi:ss') "
            gstrSQL = gstrSQL & ")"
            
            ReDim Preserve arrSql(UBound(arrSql) + 1)
            arrSql(UBound(arrSql)) = gstrSQL
            
            '����������˵���ǩ��������Ҫ����ҩ�˽��е���ǩ������
            If gblnESign������ҩ = True And gblnESignUserStoped = False Then
                strǩ����¼ = ""
                If GetSignatureRecored(EsignTache.Dosage, rsData!����, rsData!NO, mParams.lngҩ��ID, strǩ����¼, 0, date��ҩ����, gstrUserName) = False Then
                    Exit Function
                End If
                
                If strǩ����¼ <> "" Then
                    gstrSQL = "Zl_ҩƷǩ����¼_Insert(" & strǩ����¼ & ")"
                    
                    ReDim Preserve arrSql(UBound(arrSql) + 1)
                    arrSql(UBound(arrSql)) = gstrSQL
                End If
            End If

            mstrPrintRecipe = IIf(mstrPrintRecipe = "", "", mstrPrintRecipe & "|") & rsData!NO & "," & rsData!���� & "," & rsData!��¼���� & "," & rsData!�����־ & "," & rsData!�������� & "," & rsData!�շ����
            
            strNosToPlugIn = strNosToPlugIn & rsData!���� & "," & rsData!NO & "|"
        End If
        
        rsData.MoveNext
    Loop
    
    gcnOracle.BeginTrans
    blnInTrans = True
    For i = 0 To UBound(arrSql)
        Call zldatabase.ExecuteProcedure(CStr(arrSql(i)), "RecipeWork_Abolish")
    Next
    gcnOracle.CommitTrans
   
    blnInTrans = False
    
    PrintDosage
        
    '���÷�ҩ�����ҽӿ�
    If Not mobjPlugIn Is Nothing Then
        If Right(strNosToPlugIn, 1) = "|" Then strNosToPlugIn = Left(strNosToPlugIn, Len(strNosToPlugIn) - 1)
        On Error Resume Next
        mobjPlugIn.DrugDosageByRecipe mParams.lngҩ��ID, strNosToPlugIn, date��ҩ����, strReserve
        err.Clear: On Error GoTo 0
    End If
    
    RecipeWork_Dosage = True
    Exit Function
ErrHand:
    If blnInTrans Then gcnOracle.RollbackTrans
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub VerifySign()
    Dim rsData As Recordset
    Dim int�ɲ��� As Integer
    
    If gblnESign������ҩ = False Then Exit Sub
    Set rsData = mfrmDetail.GetRecord(int�ɲ���)
    
'    rsData.Filter = "��־=1"
    rsData.Sort = "����,No"
    
    If Not rsData.EOF Then
        '����������˵���ǩ��������Ҫ����ҩ�˽��е���ǩ������
        If VerifySignatureRecored_bak(IIf(Me.tbcList.Item(mListType.����ҩ).Selected = True, EsignTache.Dosage, IIf(int�ɲ��� = 1, EsignTache.send, EsignTache.returnStep)), rsData!����, rsData!NO, mParams.lngҩ��ID, 0, IIf(Me.tbcList.Item(mListType.����ҩ).Selected = True, rsData!��ҩ����, rsData!�������)) = False Then
            Exit Sub
        End If
    End If
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function RecipeWork_Abolish() As Boolean
    'ȡ����ҩ
    Dim blnInTrans As Boolean
    Dim str����Ա As String
    Dim str��ǰ���� As String
    Dim rsData As ADODB.Recordset
    Dim int���� As Integer
    Dim arrSql As Variant
    Dim i As Long
    Dim blnBeginTrans As Boolean
    Dim lngǩ��id As Long
    
    On Error GoTo ErrHand
    
    '���õ���ǩ��ʱ����û��Ƿ�ע��
    If gblnESign������ҩ = True Then
        If Not gobjESign.CheckCertificate(gstrDbUser) Then
            Exit Function
        End If
    End If
    
    arrSql = Array()
    
    Set rsData = mfrmDetail.GetRecord
    
    rsData.Filter = "��־=1"
    rsData.Sort = "����,No"
    
    Do While Not rsData.EOF
        If str��ǰ���� <> rsData!���� & "|" & rsData!NO Then
            str��ǰ���� = rsData!���� & "|" & rsData!NO
        
            '����Ƿ�����
            If CheckBill(rsData!ҩ��ID, 2, rsData!����, rsData!NO, rsData!��¼����, rsData!�����־) <> 0 Then Exit Function
        End If
        
        rsData.MoveNext
    Loop
    
    rsData.Filter = "��־=1"
    rsData.Sort = "����,No"
    str��ǰ���� = ""
    
    Do While Not rsData.EOF
        If str��ǰ���� <> rsData!���� & "|" & rsData!NO Then
            str��ǰ���� = rsData!���� & "|" & rsData!NO
        
            '����������˵���ǩ������ȡ����ҩ�˵���ǩ��
            If gblnESign������ҩ = True And gblnESignUserStoped = False Then
                lngǩ��id = 0
                If DelSignatureRecored_Check(EsignTache.Dosage, rsData!����, rsData!NO, mParams.lngҩ��ID, lngǩ��id, 0, CDate(rsData!��ҩ����)) = False Then
                    Exit Function
                End If
                
                If lngǩ��id > 0 Then
                    gstrSQL = "zl_ҩƷǩ����¼_Delete(" & lngǩ��id & ")"
                    
                    ReDim Preserve arrSql(UBound(arrSql) + 1)
                    arrSql(UBound(arrSql)) = gstrSQL
                End If
            End If
                
            If Val(rsData!��¼����) = 1 Or (Val(rsData!��¼����) = 2 And (Val(rsData!�����־)) = 1 Or (Val(rsData!�����־)) = 4) Then
                int���� = 1
            Else
                int���� = 2
            End If
            
            gstrSQL = "zl_ҩƷ�շ���¼_������ҩ��("
            '�ⷿID
            gstrSQL = gstrSQL & mParams.lngҩ��ID
            '����
            gstrSQL = gstrSQL & "," & rsData!����
            'NO
            gstrSQL = gstrSQL & ",'" & rsData!NO & "'"
            '����
            gstrSQL = gstrSQL & "," & int����
            '��ҩ��
            gstrSQL = gstrSQL & ",Null"
            '��ҩ����
            gstrSQL = gstrSQL & ",Null"
            gstrSQL = gstrSQL & ")"

            ReDim Preserve arrSql(UBound(arrSql) + 1)
            arrSql(UBound(arrSql)) = gstrSQL
        End If
        
        rsData.MoveNext
    Loop
    
    gcnOracle.BeginTrans
    blnInTrans = True
    For i = 0 To UBound(arrSql)
        Call zldatabase.ExecuteProcedure(CStr(arrSql(i)), "RecipeWork_Abolish")
    Next
    gcnOracle.CommitTrans
    blnInTrans = False
    
    RecipeWork_Abolish = True
    Exit Function
ErrHand:
    If blnInTrans Then gcnOracle.RollbackTrans
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function RecipeWork_Return(ByVal vsfDetail As VSFlexGrid) As Boolean
    Dim str���� As String, dbl��ҩ�� As Double, strSubSql As String
    Dim Int���� As Integer
    Dim strNo As String
    Dim dblSumMoney  As Double
    Dim bln�Ƿ�����ҩ  As Boolean
    Dim lngRow As Integer
    Dim rstemp As ADODB.Recordset
    Dim str��Ŵ� As String
    Dim blnInTrans As Boolean
    Dim blnIsReturn As Boolean
    Dim int���� As Integer
    Dim arrSql As Variant
    Dim i As Integer
    Dim strǩ����¼ As String
    Dim Int��ҩ As Integer
    Dim strReturnInfo As String
    Dim strReserve As String
    
    On Error GoTo ErrHand
    
    '���õ���ǩ��ʱ����û��Ƿ�ע��
    If gblnESign������ҩ = True Then
        If Not gobjESign.CheckCertificate(gstrDbUser) Then
            Exit Function
        End If
    End If
    
    arrSql = Array()
    
    Int���� = Val(vsfDetail.TextMatrix(1, vsfDetail.ColIndex("����")))
    strNo = vsfDetail.TextMatrix(1, vsfDetail.ColIndex("NO"))
    
    '��ת�������ݲ��������
    If Sys.IsMovedByNO("ҩƷ�շ���¼", strNo, "���� = ", Int����) Then
        MsgBox "�ô����ѱ�ת���������������ҩ������", vbInformation, gstrSysName
        Exit Function
    End If
    '����Ƿ�����
    If CheckBill(0, 4, Int����, strNo, Val(vsfDetail.TextMatrix(1, vsfDetail.ColIndex("��¼����"))), Val(vsfDetail.TextMatrix(1, vsfDetail.ColIndex("�����־"))), True) <> 0 Then Exit Function
    Call GetBillSequence(vsfDetail)
    
    mrsList.Filter = "����=" & Int���� & " And NO='" & strNo & "' "
    If Not mrsList.EOF Then dblSumMoney = Val(mrsList!���)
    
    If mstr��� = "" Then Exit Function
    If Not IsReceiptBalance_Charge(1, mstrPrivs, Int����, strNo, mstr���, Val(vsfDetail.TextMatrix(1, vsfDetail.ColIndex("��¼����"))), Val(vsfDetail.TextMatrix(1, vsfDetail.ColIndex("�����־")))) Then Exit Function
    If Not IsOutPatient(mstrPrivs, Int����, strNo, Val(vsfDetail.TextMatrix(1, vsfDetail.ColIndex("��¼����"))), Val(vsfDetail.TextMatrix(1, vsfDetail.ColIndex("�����־")))) Then Exit Function
    If Not CheckBillControl(mcondition.intListType + 1, Int����, strNo, dblSumMoney) Then Exit Function

    If MsgBox("��ȷ������Ϊ[" & strNo & "]" & "�Ĵ�����ҩ��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
    
    str���� = Format(Sys.Currentdate, "yyyy-MM-dd HH:mm:ss")
    
    Select Case mParams.strUnit
        Case "�ۼ۵�λ"
            strSubSql = "*1"
        Case "���ﵥλ"
            strSubSql = "*Decode(�����װ,Null,1,0,1,�����װ)"
        Case "סԺ��λ"
            strSubSql = "*Decode(סԺ��װ,Null,1,0,1,סԺ��װ)"
        Case "ҩ�ⵥλ"
            strSubSql = "*Decode(ҩ���װ,Null,1,0,1,ҩ���װ)"
        End Select
    
    bln�Ƿ�����ҩ = False
    For lngRow = 1 To vsfDetail.rows - 2
        dbl��ҩ�� = Val(vsfDetail.TextMatrix(lngRow, vsfDetail.ColIndex("��ҩ��")))

        gstrSQL = " Select round(" & dbl��ҩ�� & strSubSql & ",5) ���� From ҩƷ���" & _
                     " Where ҩƷID=[1]"
        Set rstemp = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(vsfDetail.TextMatrix(lngRow, vsfDetail.ColIndex("ҩƷID"))))
                     
        With rstemp
            dbl��ҩ�� = !����
        End With
        
        If mParams.bln��ʾ��С��λ = True Then
            If (Val(vsfDetail.TextMatrix(lngRow, vsfDetail.ColIndex("��ҩ��(���װ)"))) = Val(vsfDetail.TextMatrix(lngRow, vsfDetail.ColIndex("׼������"))) And _
                Val(vsfDetail.TextMatrix(lngRow, vsfDetail.ColIndex("��ҩ��(С��װ)"))) = Val(vsfDetail.TextMatrix(lngRow, vsfDetail.ColIndex("׼����С")))) Or _
                (Val(vsfDetail.TextMatrix(lngRow, vsfDetail.ColIndex("��ҩ��"))) = Val(vsfDetail.TextMatrix(lngRow, vsfDetail.ColIndex("׼������"))) * Val(vsfDetail.TextMatrix(lngRow, vsfDetail.ColIndex("��װ"))) + Val(vsfDetail.TextMatrix(lngRow, vsfDetail.ColIndex("׼����С")))) Then
                
                dbl��ҩ�� = Val(vsfDetail.TextMatrix(lngRow, vsfDetail.ColIndex("ʵ������")))
            End If
        Else
            If Val(vsfDetail.TextMatrix(lngRow, vsfDetail.ColIndex("��ҩ��"))) = Val(vsfDetail.TextMatrix(lngRow, vsfDetail.ColIndex("׼����"))) Then
                dbl��ҩ�� = Val(vsfDetail.TextMatrix(lngRow, vsfDetail.ColIndex("ʵ������")))
            End If
        End If
        
        If dbl��ҩ�� <> 0 Then
            blnIsReturn = False
            
            '�ȼ���ִ��Ԥ����
            Call AutoAdjustPrice_ByID(Val(vsfDetail.TextMatrix(lngRow, vsfDetail.ColIndex("ҩƷID"))))
        
            '���۸�
            If CheckPrice(Val(vsfDetail.TextMatrix(lngRow, vsfDetail.ColIndex("id"))), mstr�۸�ʧЧ��ʾ) = False Then
                If MsgBox("ҩƷ[" & vsfDetail.TextMatrix(lngRow, vsfDetail.ColIndex("ҩƷ����")) & "]" & mstr�۸�ʧЧ��ʾ, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                    blnIsReturn = True
                End If
            Else
                blnIsReturn = True
            End If
            
            If blnIsReturn = True Then
                If Val(vsfDetail.TextMatrix(lngRow, vsfDetail.ColIndex("��¼����"))) = 1 Or (Val(vsfDetail.TextMatrix(lngRow, vsfDetail.ColIndex("��¼����"))) = 2 And (Val(vsfDetail.TextMatrix(lngRow, vsfDetail.ColIndex("�����־")))) = 1 Or (Val(vsfDetail.TextMatrix(lngRow, vsfDetail.ColIndex("�����־")))) = 4) Then
                    int���� = 1
                Else
                    int���� = 2
                End If
                
                gstrSQL = "zl_ҩƷ�շ���¼_������ҩ("
                '�շ�ID
                gstrSQL = gstrSQL & Val(vsfDetail.TextMatrix(lngRow, vsfDetail.ColIndex("id")))
                '�����
                gstrSQL = gstrSQL & ",'" & gstrUserName & "'"
                '�������
                gstrSQL = gstrSQL & ",to_date('" & str���� & "','yyyy-MM-dd hh24:mi:ss') "
                '����
                gstrSQL = gstrSQL & "," & IIf(Trim(vsfDetail.TextMatrix(lngRow, vsfDetail.ColIndex("������"))) = "", "NULL", "'" & vsfDetail.TextMatrix(lngRow, vsfDetail.ColIndex("������")) & "'")
                'Ч��
                gstrSQL = gstrSQL & "," & IIf(Trim(vsfDetail.TextMatrix(lngRow, vsfDetail.ColIndex("��Ч��"))) = "", "NULL", "to_date('" & vsfDetail.TextMatrix(lngRow, vsfDetail.ColIndex("��Ч��")) & "','yyyy-MM-dd')")
                '����
                gstrSQL = gstrSQL & "," & IIf(Trim(vsfDetail.TextMatrix(lngRow, vsfDetail.ColIndex("�²���"))) = "", "NULL", "'" & Trim(vsfDetail.TextMatrix(lngRow, vsfDetail.ColIndex("�²���"))) & "'")
                '��ҩ��
                gstrSQL = gstrSQL & "," & dbl��ҩ��
                '��ҩ�ⷿ
                gstrSQL = gstrSQL & ",NULL"
                '��ҩ��
                gstrSQL = gstrSQL & ",NULL"
                '����λ��
                gstrSQL = gstrSQL & "," & mParams.int����λ��
                '����
                gstrSQL = gstrSQL & "," & int����
                gstrSQL = gstrSQL & ")"
                    
'                Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption & "-ҩƷ��ҩ")
                
                ReDim Preserve arrSql(UBound(arrSql) + 1)
                arrSql(UBound(arrSql)) = gstrSQL
            
                bln�Ƿ�����ҩ = True
                
                strReturnInfo = IIf(strReturnInfo = "", "", strReturnInfo & "|") & Val(vsfDetail.TextMatrix(lngRow, vsfDetail.ColIndex("id"))) & "," & dbl��ҩ��
            End If
        End If
    Next
    
    '������ز����������Զ����ʣ����ҵ�ǰ�˷ѵ����Ǽ��ʵ�����ôִ������/סԺ����
    If mParams.int�Զ����� = 1 And Val(vsfDetail.TextMatrix(1, vsfDetail.ColIndex("��¼����"))) = 2 And bln�Ƿ�����ҩ = True Then
        For lngRow = 1 To vsfDetail.rows - 2
            If Val(vsfDetail.TextMatrix(lngRow, vsfDetail.ColIndex("��ҩ��"))) <> 0 Then
                str��Ŵ� = str��Ŵ� & IIf(str��Ŵ� = "", vsfDetail.TextMatrix(lngRow, vsfDetail.ColIndex("���")), "," & vsfDetail.TextMatrix(lngRow, vsfDetail.ColIndex("���")))
            End If
        Next
        If Val(vsfDetail.TextMatrix(1, vsfDetail.ColIndex("�����־"))) = 1 Or Val(vsfDetail.TextMatrix(1, vsfDetail.ColIndex("�����־"))) = 4 Then
            gstrSQL = "Zl_������ʼ�¼_Delete("
            'NO
            gstrSQL = gstrSQL & "'" & strNo & "'"
            '��Ŵ�
            gstrSQL = gstrSQL & ",'" & str��Ŵ� & "'"
            '����Ա���
            gstrSQL = gstrSQL & ",'" & gstrUserCode & "'"
            '����Ա����
            gstrSQL = gstrSQL & ",'" & gstrUserName & "'"
            gstrSQL = gstrSQL & ")"
        Else
            gstrSQL = "Zl_סԺ���ʼ�¼_Delete("
            'NO
            gstrSQL = gstrSQL & "'" & strNo & "'"
            '��Ŵ�
            gstrSQL = gstrSQL & ",'" & str��Ŵ� & "'"
            '����Ա���
            gstrSQL = gstrSQL & ",'" & gstrUserCode & "'"
            '����Ա����
            gstrSQL = gstrSQL & ",'" & gstrUserName & "'"
            '��¼����
            gstrSQL = gstrSQL & "," & Val(vsfDetail.TextMatrix(1, vsfDetail.ColIndex("��¼����")))
            gstrSQL = gstrSQL & ")"
        End If
'        Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption & "-��ҩ����")

        ReDim Preserve arrSql(UBound(arrSql) + 1)
        arrSql(UBound(arrSql)) = gstrSQL
    End If
    
    '��ʾͣ��ҩƷ
    Int��ҩ = 1
    Call CheckStopMedi(Int���� & "|" & strNo, Int��ҩ)
    If Int��ҩ = 2 Then Exit Function
    
    
    '���д�����ҩ����
    gcnOracle.BeginTrans
    blnInTrans = True
    For i = 0 To UBound(arrSql)
        Call zldatabase.ExecuteProcedure(CStr(arrSql(i)), Me.Caption & "-ҩƷ��ҩ")
    Next
    
    If gblnESign������ҩ = True And gblnESignUserStoped = False Then
        strǩ����¼ = ""
        If GetSignatureRecored(EsignTache.returnStep, Int����, strNo, mParams.lngҩ��ID, strǩ����¼, 0, CDate(str����)) = False Then
            gcnOracle.RollbackTrans
            blnInTrans = False
            Exit Function
        End If
        
        If strǩ����¼ = "" Then
            gcnOracle.RollbackTrans
            blnInTrans = False
            MsgBox "����ҩ�˵���ǩ��ʧ�ܣ�", vbInformation, gstrSysName
            Exit Function
        End If
        
        If strǩ����¼ <> "" Then
            gstrSQL = "Zl_ҩƷǩ����¼_Insert(" & strǩ����¼ & ")"
            Call zldatabase.ExecuteProcedure(gstrSQL, "ǩ������")
        End If
    End If
    gcnOracle.CommitTrans
    blnInTrans = False
    
    '��ӡ�˷�֪ͨ��
    Dim Str��ҩʱ�� As String, int��װϵ�� As Integer
    
    If bln�Ƿ�����ҩ Then
        Str��ҩʱ�� = str����
        int��װϵ�� = IIf(Int���� = 8, 1, 2)
        
        If MsgBox("����Ҫ��ӡ��ҩ֪ͨ����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
            Call ReportOpen(gcnOracle, glngSys, IIf(glngSys \ 100 = 1, "ZL1_BILL_1341_1", "ZL8_BILL_1341_1"), _
            Me, "No=" & strNo, "����=" & Int����, "��װϵ��=" & IIf(int��װϵ�� = 1, "D.�����װ", "D.סԺ��װ"), "��ҩʱ��=" & Str��ҩʱ��, 2)
        End If
    Else
        MsgBox "����û����ҩ��"
    End If
    
    '������ҩ�����ҽӿ�
    If Not mobjPlugIn Is Nothing And bln�Ƿ�����ҩ Then
        On Error Resume Next
        mobjPlugIn.DrugReturnByID mParams.lngҩ��ID, strReturnInfo, CDate(str����), strReserve
        err.Clear: On Error GoTo 0
    End If
    
    Exit Function
ErrHand:
    If blnInTrans Then gcnOracle.RollbackTrans
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Private Sub ResetFilter()
    '�������ù�������
    Dim strReturn As String, IntOper As Integer
    
    IntOper = mcondition.intListType + 1
    
    With FrmҩƷ��ҩ����
        strReturn = .ShowMe(Me, mParams.lngҩ��ID, IntOper, mstrPrivs, mbln���￨, _
            mSQLCondition.date��ʼ����, _
            mSQLCondition.date��������, _
            mSQLCondition.str��ʼNO, _
            mSQLCondition.str����NO, _
            mSQLCondition.str����, _
            mSQLCondition.str���￨, _
            mSQLCondition.str��ʶ��, _
            mSQLCondition.lng����ID, _
            mSQLCondition.str������, _
            mSQLCondition.str�����, _
            mSQLCondition.lngҩƷid, _
            mSQLCondition.strҽ����, _
            mcondition.int��Ժ��ҩ)
        If strReturn = "" Then Exit Sub
    End With
    
    mint���˲�ѯ = 1

    If mPrives.bln�����ѯ����ʱ�䷶Χ���� Then
        cboʱ�䷶Χ.ListIndex = 3
    Else
        cboʱ�䷶Χ.ListIndex = 0
    End If
    
    Call picConMain_Resize
    Call picCondition_Resize
    Dtp��ʼʱ��.Value = mSQLCondition.date��ʼ����
    Dtp����ʱ��.Value = mSQLCondition.date��������
    
    If imgFilter.BorderStyle = cstFilter Then
        Call txtPati_KeyPress(13)
    Else
        Call RefreshList(mcondition.intListType)
    End If
    
    mint���˲�ѯ = 0
End Sub

Private Sub ResetParams()
    Dim strTmp As String
    Dim intCurrTab As Integer
    
    BlnSetParaSuccess = False
    BlnRefresh = False
    
    '�ر�Timer
    Call SetTimerState(False)
    
    With Frm��ҩ��������
        Set .RecPart = RecPart.Clone
        .mstrPrivs = mstrPrivs
'        .In_���÷�ҩ = (Not gobjPackerMZ Is Nothing)
        .In_���÷�ҩ = mblnLoadDrug
        If Not mobjMipModule Is Nothing Then
            If mobjMipModule.IsConnect = True Then
                .In_������Ϣ = True
            Else
                .In_������Ϣ = False
            End If
        Else
            .In_������Ϣ = False
        End If
        .Show 1, Me
    End With
    
    If Not BlnSetParaSuccess Then
        '�����ޱ仯ʱ
    
        '����Timer
        Call SetTimerState(True)
    Else
        '�����б仯ʱ�����¸����ڵĲ���
        Call GetParams
        mfrmList.SetParams
        mfrmDetail.SetParams
        mfrmRecipe.SetParams
        
        '����ʱ��ؼ�
        If mParams.lngRefreshInterval > 0 Then
            If mParams.lngRefreshInterval > 60 Then
                mParams.lngRefreshInterval = 60
            End If
            With TimeRefresh
                .Enabled = True
                .Interval = mParams.lngRefreshInterval * 1000
            End With
        Else
            TimeRefresh.Enabled = False
        End If
        
        If mParams.lngPrintInterval > 0 Then
            If mParams.lngPrintInterval > 60 Then
                mParams.lngPrintInterval = 60
            End If
            With TimePrint
                .Enabled = True
                .Interval = mParams.lngPrintInterval * 1000
            End With
        Else
            TimePrint.Enabled = False
        End If
        
        IntTimes = 0
        
        If mParams.lngPrintBackInterval <> 0 Then
            With TimePrintCancelBill
                .Enabled = False
                .Enabled = True
            End With
        Else
            TimePrintCancelBill.Enabled = False
        End If
        
        '���ýк���ѯʱ����������������������ȫ�ֽк�Զ�˻�����ʱ
        tmrCall.Enabled = False
        If mParams.blnStartQueue = True And mParams.blnStartCall = True And (mParams.intCallType = 0 And mQueue.strPCName = mParams.strRemoteCall And mQueue.strPCName <> "") Then
            tmrCall.Enabled = True
            tmrCall.Interval = mParams.intCircleTime * 1000
        End If
        
        GetDrugStock mParams.lngҩ��ID
        GetDosage mParams.lngҩ��ID
        GetStockName mParams.lngҩ��ID
        GetSendWindows mParams.lngҩ��ID
        
        If Not gobjESign Is Nothing Then
            gblnESign������ҩ = EsignIsOpen(mParams.lngҩ��ID)
        End If
        
        strTmp = Me.dkpMain.FindPane(mconPane_Recipe_Condition).Title
        strTmp = mstrStockName & Mid(strTmp, InStr(strTmp, ":"))
        Me.dkpMain.FindPane(mconPane_Recipe_Condition).Title = strTmp
        
        intCurrTab = mcondition.intListType
        
        If mParams.blnMustDosageProcess = True Then
'            tbcList.Item(mconTab_Recipe_Abolish).Visible = True
            tbcList.Item(mconTab_Recipe_Dosage).Visible = True
'            tbcList.Item(mconTab_Recipe_Abolish).Selected = True
'            tbcList.Item(mconTab_Recipe_Dosage).Selected = True
        Else
            tbcList.Item(mconTab_Recipe_Dosage).Visible = False
'            tbcList.Item(mconTab_Recipe_Abolish).Visible = False
'            tbcList.Item(mconTab_Recipe_Return).Selected = True
'            tbcList.Item(mconTab_Recipe_Send).Selected = True
        End If
        
'        If mParams.blnMustDosageOkProcess = True Then
'            tbcList.Item(mconTab_Recipe_DosageOk).Visible = True
'        Else
'            tbcList.Item(mconTab_Recipe_DosageOk).Visible = False
'        End If
        
        tbcList.Item(mconTab_Recipe_OverTime).Visible = (mParams.intOverTime > 0)
    
        
        If CheckAnother = False Then Exit Sub
        
        If tbcList.Item(mcondition.intListType).Visible = True Then
            tbcList.Item(mcondition.intListType).Selected = True
        Else
            If mParams.blnMustDosageProcess = True Then
                tbcList.Item(mconTab_Recipe_Dosage).Selected = True
            Else
                tbcList.Item(mconTab_Recipe_Send).Selected = True
            End If
        End If
        
        Call tbcList_SelectedChanged(tbcList.Item(mcondition.intListType))
        If intCurrTab = mcondition.intListType Then
            RefreshList mcondition.intListType
        End If
        
        '������ʾ�ŶӴ���
        If mParams.blnShowQueue And mParams.blnStartQueue Then
            Call ShowQueue
        Else
            CloseQueue
        End If
        
        Call GetOpr
        
        mbln��������ˢ�� = False
        If mParams.str����ˢ����ҩ <> "" Then
            mbln��������ˢ�� = InStr(1, "," & mParams.str����ˢ����ҩ & ",", "," & mobjcard.�ӿ���� & ",") > 0
        End If
    End If
End Sub

'Private Sub SetInputState(ByVal intType As Integer)
'    Dim cbrControl As CommandBarControl
'
'    Set cbrControl = Me.cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_Input_Recipe_NO + intType, , True)
'    If Not cbrControl Is Nothing Then
'        SetInputPopupCheck cbrControl
'    End If
'End Sub
Private Sub SetLocatePrinter(ByVal intRecipeType As Integer, Optional ByVal int��ʽ As Integer)
    '��ӡ��ҩ����ǩʱ��������ɫ��ָ����Ӧ�Ĵ�ӡ��
    'int��ʽ:   ��������Ӧ�ĸ�ʽ
    Dim strPrinter As String
    Dim i As Integer
    
    If mParams.strPrinters = "" Then Exit Sub
    
    If intRecipeType < 0 Or intRecipeType > 5 Then intRecipeType = 0
    
    On Error GoTo errHandle
    
    If InStr(mParams.strPrinters, "?") = 0 Then
        '������ǰ�Ĵ洢����
        strPrinter = Split(mParams.strPrinters, ";")(intRecipeType)
    Else
        strPrinter = Mid(Split(Split(mParams.strPrinters, ";")(intRecipeType), ",")(int��ʽ), InStr(Split(Split(mParams.strPrinters, ";")(intRecipeType), ",")(int��ʽ), "?") + 1)
    End If
    
    If strPrinter <> "" Then
        '���洦������ָ���Ĵ�ӡ��������ע���
        Call SavePrinterSet("ZL1_BILL_1341_3", mstrRPTDefaultScheme_Recipt, strPrinter)
        Call SavePrinterSet("ZL1_BILL_1341_3", mstrRPTScheme_��ҩ��, strPrinter)
        For i = 0 To UBound(Split(mstrRPTScheme_������ʽ, ";"))
            Call SavePrinterSet("ZL1_BILL_1341_3", Split(mstrRPTScheme_������ʽ, ";")(i), strPrinter)
        Next
    End If
    
    'ͬʱ��ӡ���и�ʽ
    If int��ʽ = -1 Then
        If InStr(mParams.strPrinters, "?") = 0 Then
            Exit Sub
        Else
            For i = 0 To UBound(Split(Split(mParams.strPrinters, ";")(intRecipeType), ","))
                strPrinter = Mid(Split(Split(mParams.strPrinters, ";")(intRecipeType), ",")(i), InStr(Split(Split(mParams.strPrinters, ";")(intRecipeType), ",")(i), "?") + 1)
                If strPrinter <> "" Then
                    If i = 0 Then
                        Call SavePrinterSet("ZL1_BILL_1341_3", mstrRPTDefaultScheme_Recipt, strPrinter)
                    ElseIf i = 1 Then
                        Call SavePrinterSet("ZL1_BILL_1341_3", mstrRPTScheme_��ҩ��, strPrinter)
                    Else
                        Call SavePrinterSet("ZL1_BILL_1341_3", Split(mstrRPTScheme_������ʽ, ";")(i - 2), strPrinter)
                    End If
                End If
            Next
        End If
    End If
    
    Exit Sub
errHandle:
    Resume Next
End Sub

Private Sub SavePrinterSet(ByVal strRPTCode As String, ByVal strRPTScheme As String, ByVal strPrinter As String)
    '�����ӡ����Ϣ���������ڱ����ӡʱ��ʱ���Ĵ�ӡ�����ƣ��÷�Ϊ�������Σ���ӡǰ�������δ�ӡ��Ҫ�Ĵ�ӡ������ӡ��ָ�ΪĬ�ϵĴ�ӡ��
    '���Σ�strRPTCode-������룻strRPTScheme-�����ʽ��strPrinter-��ӡ������
    SaveSetting "ZLSOFT", "˽��ģ��\zl9Report\LocalSet\" & strRPTCode & "\" & strRPTScheme, "Printer", strPrinter
End Sub
Private Sub SetPaneTitle(ByVal intType As Integer)
    Dim strTitleCon As String
    Dim strTitleList As String
    
    Select Case intType
        Case mListType.��ҩȷ��
            strTitleCon = "��ҩȷ��"
        Case mListType.����ҩ
            strTitleCon = "����ҩ"
        Case mListType.����ҩ
            strTitleCon = "����ҩ"
        Case mListType.����ҩ
            strTitleCon = "����ҩ"
        Case mListType.��ʱδ��
            strTitleCon = "��ʱδ��"
        Case mListType.��ҩ
            strTitleCon = "��ҩ"
    End Select
    
    Me.dkpMain.FindPane(mconPane_Recipe_Condition).Title = mstrStockName & ":" & strTitleCon
End Sub

Private Sub SetTimerState(ByVal BlnSet As Boolean)
    '�رպ�����Timer�ؼ����е�������ʱ����
    'blnSet��True-������False-�ر�
    
    If BlnSet Then
        '����ʱ�ָ�ԭ����״̬
        TimeRefresh.Enabled = mblnStateTimeRefresh
        TimePrint.Enabled = mblnStateTimePrint
        tmrCall.Enabled = mblnStateTimeCall
    Else
        '�ر�ʱ�ȼ�¼ԭ����״̬
        mblnStateTimeRefresh = TimeRefresh.Enabled
        mblnStateTimePrint = TimePrint.Enabled
        mblnStateTimeCall = tmrCall.Enabled
        
        If mblnStateTimeRefresh Then TimeRefresh.Enabled = False
        If mblnStateTimePrint Then TimePrint.Enabled = False
        If mblnStateTimeCall Then tmrCall.Enabled = False
    End If
End Sub
Private Sub GetBillSequence(ByVal vsfDetail As VSFlexGrid)
    Dim intRow As Integer, intRows As Integer
    Dim int��� As Integer
    '��ȡ��ǰ����ҩ������ҩ��������Ч���
    mstr��� = ""
    intRows = vsfDetail.rows - 2
    
    If mcondition.intListType = mListType.��ҩ Then
        '��ҩ����Ϊ���ʾ����Ҫ�˵���ϸ����ͳ�Ƴ�������ϸ�����
        For intRow = 1 To intRows
            If Val(vsfDetail.TextMatrix(intRow, vsfDetail.ColIndex("��ҩ��"))) <> 0 Then
                int��� = Val(vsfDetail.TextMatrix(intRow, vsfDetail.ColIndex("���")))
                If InStr(1, mstr��� & ",", "," & int��� & ",") = 0 Then
                    mstr��� = mstr��� & "," & int���
                End If
            End If
        Next
    Else
        For intRow = 1 To intRows
            int��� = Val(vsfDetail.TextMatrix(intRow, vsfDetail.ColIndex("���")))
            If InStr(1, mstr��� & ",", "," & int��� & ",") = 0 Then
                mstr��� = mstr��� & "," & int���
            End If
        Next
    End If
    If mstr��� <> "" Then mstr��� = Mid(mstr���, 2)
End Sub
Private Function RecipeWork_Send(ByVal rsData As ADODB.Recordset) As Boolean
    '��ҩ
    Dim str����Ա As String
    Dim str��ǰ���� As String
    
    On Error GoTo ErrHand
    
    mblnSendIsOver = False
    
    If rsData Is Nothing Then Exit Function
    
    rsData.Filter = "��־=1"
    rsData.Sort = "����,No"
    
    '���õ���ǩ��ʱ����û��Ƿ�ע��
    If gblnESign������ҩ = True Then
        If Not gobjESign.CheckCertificate(gstrDbUser) Then
            Exit Function
        End If
    End If
    
    '������鴦��
    If Not CheckBatchRecipe(rsData) Then Exit Function
    
    '�ϵ�һ��ͨ����ˢ��
    If CheckCard(rsData) = False Then Exit Function
    
    '�µ����ѿ�ˢ�����ѽӿ�
    If Not CardConfirm(rsData) Then Exit Function
    
    '����������ҩ
    If Not SendBatchRecipe(rsData) Then
        Exit Function
    End If
    
    '����֧����֮�󣬷�����ʾ��Ϣ
    Call msg_upload(rsData)
    
    PrintRecipe
    
    '������ҽӿڹ��ܣ��緢ҩ�����������ܣ�ÿ�η�ҩֻ����һ�Σ�
    If Not mobjPlugIn Is Nothing Then
        rsData.MoveFirst
        On Error Resume Next
        Call mobjPlugIn.OutPatiMedicineAfter(rsData!����ID, rsData!NO, rsData!����, mParams.lngҩ��ID)
        err.Clear: On Error GoTo 0
    End If
    
    '���ò����ڸ�ҩ���Ƿ���δ������������
    If mParams.bln��ҩ���� Then
        Call checkStuff(rsData!����ID)
    End If
    
    RecipeWork_Send = True
    mblnSendIsOver = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
    mblnSendIsOver = True
End Function

Private Sub checkStuff(ByVal lng����ID As Long)
    Dim strsql As String
    Dim rstemp As Recordset
    
    On Error GoTo ErrHand
    strsql = " select count(A.id) ���� from ҩƷ�շ���¼ A,������ü�¼ B where A.NO=B.NO and A.����id=B.id and B.����id=[2] and A.���� in (24,25) and A.�ⷿid=[1] and A.����� is null and (A.��¼״̬=1 or MOD(A.��¼״̬,3)=1)"
    Set rstemp = zldatabase.OpenSQLRecord(strsql, "checkStuff", mParams.lngҩ��ID, lng����ID)
    
    If rstemp!���� > 0 Then
        MsgBox "�ò��˻���δ�����������ϣ���ע�ⷢ�ţ�", vbInformation, gstrSysName
    End If
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
Private Sub msg_upload(ByVal rsData As Recordset)
    '����֧������ʾ��Ϣ
    Dim strMsg As String
    Dim strsql As String
    Dim cmdTmp As New ADODB.Command
    Dim cmdPara As New ADODB.Parameter
    
    On Error GoTo ErrHand
    rsData.MoveFirst
    
    Set cmdTmp = New ADODB.Command
    Set cmdPara = cmdTmp.CreateParameter("����ID", adVarNumeric, adParamInput, 18, rsData!����ID)
    cmdTmp.Parameters.Append cmdPara
    Set cmdPara = cmdTmp.CreateParameter("NO", adVarChar, adParamInput, 100, rsData!NO)
    cmdTmp.Parameters.Append cmdPara
    Set cmdPara = cmdTmp.CreateParameter("˵��", adLongVarChar, adParamOutput, 4000)
    cmdTmp.Parameters.Append cmdPara
    
    cmdTmp.ActiveConnection = gcnOracle
    cmdTmp.CommandType = adCmdStoredProc
    cmdTmp.CommandText = "Zl_MSG_PointOut"
    cmdTmp.Execute
    strMsg = Trim(zlStr.NVL(cmdTmp.Parameters("˵��"), ""))
    
    If strMsg <> "" Then
        MsgBox strMsg, vbInformation, gstrSysName
    End If
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Function RefreshDetail_Return(ByVal BillStyle As Integer, ByVal BillNo As String, ByVal str������� As String, ByVal int�ɲ��� As Integer, ByVal int�����־ As Integer, ByVal int��¼���� As Integer, Optional blnByNo As Boolean = False, Optional lng��¼״̬ As Long) As Boolean
    Dim IntStyle As Integer, intUnit As Integer
    Dim strSubSql As String
    Dim strName As String
    Dim blnMoved As Boolean
    Dim lng����ID As Long
    Dim int��ҳid As Integer
    Dim strWeight As String
    
    Dim rstemp As New ADODB.Recordset
    Dim RecBill As New ADODB.Recordset
    '--��ȡ��������--
    'BillStyle-��������;BIllNO-���ݺ�
    '��λ��ʾ���ݷ����������������ﵥλ��סԺ��סԺ���סԺ��λ���������ۼ۵�λ��
'    On Error Resume Next
'    err = 0
    On Error GoTo errHandle
    RefreshDetail_Return = False
  
    mParams.strUnit = GetUnit(mSQLCondition.lngҩ��ID, BillStyle, BillNo, int�����־)
    Select Case mParams.strUnit
    Case "�ۼ۵�λ"
        strSubSql = "1"
    Case "���ﵥλ"
        strSubSql = "Decode(�����װ,Null,1,0,1,�����װ)"
    Case "סԺ��λ"
        strSubSql = "Decode(סԺ��װ,Null,1,0,1,סԺ��װ)"
    Case "ҩ�ⵥλ"
        strSubSql = "Decode(ҩ���װ,Null,1,0,1,ҩ���װ)"
    End Select
    Call Get��λ��
    
    '�õ�ҩƷ���ƴ�
    Select Case mParams.intҩƷ������ʾ
    Case 0  'ҩƷ����������
        strName = "'['||C.����||']'||" & IIf(gintҩƷ������ʾ = 1, "NVL(E.����,C.����)", "C.����") & " As Ʒ��,"
    Case 1  'ҩƷ����
        strName = "C.���� As Ʒ��,"
    Case 2  'ҩƷ����
        strName = IIf(gintҩƷ������ʾ = 1, "NVL(E.����,C.����)", "C.����") & " As Ʒ��,"
    End Select
    
    strName = strName & IIf(gintҩƷ������ʾ <> 1, "NVL(E.����,'')", "Decode(E.����,Null,'',C.����)") & " As ������, "
    
    '������ʾ��������
    '�����ܴ���һ�Ŵ���ͬʱ������󱸱��ж�����
    blnMoved = Sys.IsMovedByNO("ҩƷ�շ���¼", BillNo, " ���� = ", BillStyle)
    gstrSQL = " SELECT DISTINCT B.��ҩ����,B.�������,B.�˲���,S.���� As ҩ��,B.��¼״̬ ״̬,B.����,B.��������,B.NO,H.���,T.���� ����,H.����,H.�Ա�,H.����,H.��ʶ�� סԺ��,H.����,H.������,B.ID As �շ�ID,B.ҩƷID,nvl(n.����,'') �䷽����," & _
             " DECODE(B.����,NULL,'',B.����)||DECODE(B.����,NULL,'',0,'','('||B.����||')') ����,DECODE(D.��ΣҩƷ,null,0,0,0,1) ��ΣҩƷ,to_char(B.Ч��,'yyyy-mm-dd') Ч��,X.�����,X.��������,X.���￨��,decode(X.��ϵ�˵绰,null,decode(X.�ֻ���,null,X.��ͥ�绰,X.�ֻ���),X.��ϵ�˵绰) ��ϵ�˵绰," & _
             " NVL(B.����,0) ����,NVL(D.ҩ������,0) ����," & strName & _
             IIf(gintҩƷ������ʾ = 1, "NVL(E.����,C.����)", "C.����") & " As ҩƷ����, " & _
             " DECODE(C.���,NULL,B.����,DECODE(B.����,NULL,C.���,C.���||'|'||B.����)) ���,Nvl(b.����, Nvl(c.����, '')) ����, b.ԭ����," & str��λ�� & "," & _
             " NVL(B.����,1) ����,NVL(H.����,1) ԭʼ����," & _
             " B.��������/" & strSubSql & " ��������, B.�������� С��λ������," & _
             " B.�ѷ�����/" & strSubSql & " ׼����,B.�ѷ����� С��λ׼����,B.�ѷ����� ʵ������,B.ʵ������ С��λ����," & _
             " B.����,B.�÷�,B.Ƶ��,B.������,B.��������,H.����Ա����,B.��ҩ��,B.����� ��ҩ��,I.���㵥λ," & _
             " round(B.���۽��," & mintMoneyDigit & " ) ���۽��,round(Nvl(B.����, 1) * B.ʵ������ / (Nvl(H.����, 1) * H.����) * Nvl(H.ʵ�ս��,0)," & mintMoneyDigit & " ) ʵ�ս��,H.�ѱ�,I.���� As ҩ�� ," & _
             " P.�������,Nvl(P.������,0) ������,Nvl(P.�Ƿ�Ƥ��,0) As �Ƿ�Ƥ��, H.�����־, H.��¼����,B.�ⷿid As ҩ��id,Nvl(M.���ID,0) As ���ID,M.����ҽ��,M.Ƶ�ʼ��,M.����˵��,M.�����λ,Nvl(M.����ʱ��,H.�Ǽ�ʱ��) As ����ʱ��,M.ҽ����Ч ҽ����־,M.��ʼִ��ʱ�� ��ʼʱ��,M.ִ����ֹʱ�� ����ʱ��,M.Ƶ�ʴ���,Nvl(Nvl(M.���ID,M.id),0) As ҽ��id,Nvl(M.��ҳid,0) as ��ҳid," & _
             " M.Ƥ�Խ��,M.����ҩƷ˵��,D.ҩ��ID, f.���� As ����,H.���� As ��ҩ��̬,C.��� As ҩƷ���,M.ҽ������,decode(m.��ҩĿ��,1,'Ԥ��',2,'����',3,'Ԥ��������','') ��ҩĿ��,m.��ҩ����,D.����ϵ��,"
             
    If int�ɲ��� = 1 Then  '�����������ǽ�ȥ
        gstrSQL = gstrSQL & " B.�ѷ�����*D.����ϵ�� ����,Decode(Sign(Nvl(K.�������, 0) - Nvl(L.����, 0)), -1, 0, 1) �������,Z.���� As Ӣ����,Nvl(H.����ID,0) As ����ID,Nvl(x.��Ժ, 0) As ��Ժ FROM "
        gstrSQL = gstrSQL & "   (SELECT A.ID,A.NO,A.����,A.ҩƷID,A.����ID,A.����,A.����,A.����,A.ԭ����,A.Ч��," & _
                 "          NVL(A.����,1) ����,A.ʵ������,NVL(A.����,1)*A.ʵ������-B.�ѷ����� ��������,B.�ѷ�����,A.��¼״̬," & _
                 "          A.���ۼ�,A.���۽��,A.����,A.Ƶ��,A.�÷�,A.ժҪ,A.��ҩ����,A.������,A.��������,A.�˲���,A.��ҩ��,A.�����,A.�������,A.�Է�����ID,A.�ⷿID, A.�������� " & _
                 "      FROM" & _
                 "          (SELECT A.ID,A.NO,A.����,A.ҩƷID,A.���,A.����ID,A.����,A.����,A.����,A.ԭ����,A.Ч��,A.����,A.ʵ������,A.��¼״̬,A.���ۼ�,A.���۽��,A.����,A.Ƶ��,A.�÷�,A.ժҪ,A.��ҩ����,A.������,A.��ҩ��,A.��������,A.�˲���,A.�����,A.�������,A.�Է�����ID,A.�ⷿID, Nvl(A.ע��֤��, 0) As �������� " & _
                 "          FROM ҩƷ�շ���¼ A" & _
                 "          WHERE A.����� IS NOT NULL AND (A.��¼״̬=1 OR MOD(A.��¼״̬,3)=0)" & _
                 "          AND A.�ⷿID+0=[1] "
        If blnByNo = False Then
            gstrSQL = gstrSQL & " AND A.������� Between [2] And [3] "
        Else
            gstrSQL = gstrSQL & " And A.����=[4] And A.NO=[5] "
        End If
        
        gstrSQL = gstrSQL & "          ) A," & _
                 "          (SELECT A.NO,A.����,A.ҩƷID,A.���,SUM(NVL(A.����,1)*A.ʵ������) �ѷ�����" & _
                 "          FROM ҩƷ�շ���¼ A" & _
                 "          WHERE A.����� IS NOT NULL" & _
                 "          AND A.�ⷿID+0=[1] "
        
        If blnByNo = False Then
            gstrSQL = gstrSQL & " AND A.������� Between [2] And [3] "
        Else
            gstrSQL = gstrSQL & " And A.����=[4] And A.NO=[5] "
        End If
                
        gstrSQL = gstrSQL & "          GROUP BY A.NO,A.����,A.ҩƷID,A.���) B" & _
                 "      WHERE A.NO = B.NO AND A.���� = B.���� AND A.ҩƷID+0 = B.ҩƷID AND A.��� = B.��� " & _
                 "      )"
    Else
        gstrSQL = gstrSQL & " B.ʵ������*D.����ϵ�� ����,Decode(Sign(Nvl(K.�������, 0) - Nvl(L.����, 0)), -1, 0, 1) �������,Z.���� As Ӣ����,Nvl(H.����ID,0) As ����ID,Nvl(x.��Ժ, 0) As ��Ժ FROM "
        gstrSQL = gstrSQL & "(Select 0 �ѷ�����,0 ��������,0 ׼������,Nvl(A.ע��֤��, 0) As ��������,A.* From ҩƷ�շ���¼ A where FLOOR(A.��¼״̬/3)+1=[10])"
    End If
    gstrSQL = gstrSQL & _
            "       B,ҩƷ��� D,ҩƷ���� P,�շ���ĿĿ¼ C,�շ���Ŀ���� E,������ü�¼ H,����ҽ����¼ M,����ҽ����¼ G,���ű� S,���ű� T,������ĿĿ¼ I,������Ŀ���� Z ,����֧������ F,������ĿĿ¼ N,������Ϣ X, " & _
            "(Select b.�ⷿid, b.ҩƷid, Nvl(Sum(b.ʵ������), 0) ������� " & _
            " From ҩƷ�շ���¼ A, ҩƷ��� B " & _
            " Where a.ҩƷid = b.ҩƷid And b.���� = 1 And b.�ⷿid + 0 = [1] And a.���� = [4] And a.No = [5] " & _
            " Group By b.�ⷿid, b.ҩƷid) K, ҩƷ�����޶� L " & _
            " Where H.��������ID=T.ID(+) And B.ҩƷID=D.ҩƷID And D.ҩ��ID=P.ҩ��ID And C.ID=D.ҩƷID And H.ҽ�����=M.ID(+) And Nvl(M.���id, M.ID) = G.ID(+) and G.�䷽id=N.id(+) " & _
            " And D.ҩƷID=E.�շ�ϸĿID(+) and E.����(+)=3 And D.ҩ��id = Z.������Ŀid(+) And Z.����(+) = 2 And h.���մ���id = f.Id(+) " & _
            " And S.ID=B.�ⷿID And B.����ID=H.ID And B.NO=[5] And B.����=[4] And B.�ⷿID+0=[1] and H.����id=X.����id(+) "
    
    If mSQLCondition.str������ <> "" Then gstrSQL = gstrSQL & " And B.������=[7] "
    If mSQLCondition.str����� <> "" Then gstrSQL = gstrSQL & " And B.�����=[8] "
    If mSQLCondition.lngҩƷid > 0 Then gstrSQL = gstrSQL & " And B.ҩƷID=[9] "
    
    If IsDate(str�������) Then
             gstrSQL = gstrSQL & " And B.�������=To_Date([6],'yyyy-MM-dd hh24:mi:ss')"
    End If
    gstrSQL = gstrSQL & " And B.����� Is Not Null And D.ҩ��id=I.id " & _
                        " And B.ҩƷid = L.ҩƷid(+) And Nvl(B.�ⷿid, 24) = L.�ⷿid(+) And" & _
                        " D.ҩ��id = I.ID And Nvl(B.�ⷿid, 24) + 0 = K.�ⷿid(+) And B.ҩƷid = K.ҩƷid(+) "
    
    gstrSQL = gstrSQL & " Order by H.���,B.ҩƷID,Nvl(B.����,0)"
    
    If int��¼���� = 1 Or (int��¼���� = 2 And (int�����־ = 1 Or int�����־ = 4)) Then
        '����
        gstrSQL = Replace(gstrSQL, "H.����", "'' ����")
    Else
        'סԺ
        gstrSQL = Replace(gstrSQL, "������ü�¼", "סԺ���ü�¼")
    End If
    
    '�������ת������ֱ�ӴӺ󱸱�����ȡ����
    If blnMoved Then
        gstrSQL = Replace(gstrSQL, "ҩƷ�շ���¼", "HҩƷ�շ���¼")
        gstrSQL = Replace(gstrSQL, "������ü�¼", "H������ü�¼")
        gstrSQL = Replace(gstrSQL, "סԺ���ü�¼", "HסԺ���ü�¼")
    End If
    
    Set RecBill = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, _
            mSQLCondition.lngҩ��ID, _
            mSQLCondition.date��ʼ����, _
            mSQLCondition.date��������, _
            BillStyle, _
            BillNo, _
            str�������, _
            mSQLCondition.str������, _
            mSQLCondition.str�����, _
            mSQLCondition.lngҩƷid, _
            lng��¼״̬)
    
    If err <> 0 Then
        MsgBox "��ȡ����ʱ����������Ԥ֪�Ĵ���", vbInformation, gstrSysName
        Exit Function
    End If
    
    If Not RecBill.EOF Then
        If NVL(RecBill!����ID) <> 0 Then
            lng����ID = RecBill!����ID
            int��ҳid = NVL(RecBill!��ҳid)
            If int��¼���� = 1 Or (int��¼���� = 2 And (int�����־ = 1 Or int�����־ = 4)) Then
                '����
                gstrSQL = "select A.id,B.��¼���� ���� from ���˻����¼ A,���˻������� B where A.id=B.��¼id and B.��Ŀ����='����' and ����id=[1] order by A.Id desc"
            Else
                'סԺ
                 gstrSQL = "select ���� from ������ҳ where ����id=[1] and ��ҳid=[2]"
                 
            End If
            Set rstemp = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, lng����ID, int��ҳid)
            
            If Not rstemp.EOF Then
                strWeight = NVL(rstemp!����)
            End If
        End If
    End If

    mfrmDetail.RefreshList RecBill, strWeight, int�ɲ���
    mfrmRecipe.RefreshRecipe RecBill, strWeight, int�ɲ���
    
    RefreshDetail_Return = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function SendBatchRecipe(ByVal rsData As ADODB.Recordset) As Boolean
    Dim n As Integer
    Dim lngRow As Long, lngҩƷid As Long, LngID As Long, lng���� As Long, lng���� As Long
    Dim rsSendRecipeByNo As ADODB.Recordset
    Dim rsSendRecipeDetail As ADODB.Recordset
    Dim int���� As Integer
    Dim strNO�� As String
    Dim arrSql As Variant
    Dim blnInTrans As Boolean
    Dim i As Integer
    Dim strǩ����¼ As String
    Dim date��ҩʱ�� As Date
    Dim strNo As String  '���ڴ�����ҩ��
    Dim strReturn As String, strMessage As String
    Dim strReserve As String
    
    On Error GoTo ErrHand
    
    date��ҩʱ�� = Sys.Currentdate
    
    Set rsSendRecipeByNo = New ADODB.Recordset
    With rsSendRecipeByNo
        If .State = 1 Then .Close
        .Fields.Append "NO", adLongVarChar, 8, adFldIsNullable
        .Fields.Append "����", adDouble, 18, adFldIsNullable
        .Fields.Append "��ҩ��", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "������", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "ҩ��ID", adDouble, 18, adFldIsNullable
        .Fields.Append "��¼����", adDouble, 1, adFldIsNullable
        .Fields.Append "�����־", adDouble, 1, adFldIsNullable
        .Fields.Append "��������", adDouble, 1, adFldIsNullable
        .Fields.Append "�շ����", adDouble, 1, adFldIsNullable
        .Fields.Append "����", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "�˲���", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "��������", adDate, , adFldIsNullable
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
    
    Set rsSendRecipeDetail = New ADODB.Recordset
    With rsSendRecipeDetail
        If .State = 1 Then .Close
        .Fields.Append "NO", adLongVarChar, 8, adFldIsNullable
        .Fields.Append "�շ�ID", adDouble, 18, adFldIsNullable
        .Fields.Append "ҩƷID", adDouble, 18, adFldIsNullable
        .Fields.Append "����", adDouble, 18, adFldIsNullable
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
    
    rsData.Filter = "��־=1"
    rsData.Sort = "����,NO"
    Do While Not rsData.EOF
        With rsSendRecipeByNo
            If strNO�� <> rsData!���� & "|" & rsData!NO Then
                strNO�� = rsData!���� & "|" & rsData!NO
                .AddNew
                !ҩ��ID = rsData!ҩ��ID
                !NO = rsData!NO
                !���� = rsData!����
                !��ҩ�� = IIf(tbcDetail.Selected.index = 1, mfrmRecipe.Get��ҩ��, mfrmDetail.Get��ҩ��)
                !������ = IIf(tbcDetail.Selected.index = 1, mfrmRecipe.Get����ҽ��, mfrmDetail.Get����ҽ��)
                !��¼���� = rsData!��¼����
                !�����־ = rsData!�����־
                !�������� = rsData!��������
                !�շ���� = rsData!�շ����
                !���� = IIf(IsNull(rsData!����), "", rsData!����)
                !�˲��� = IIf(tbcDetail.Selected.index = 1, mfrmRecipe.Get�˲���, mfrmDetail.Get�˲���)
                !�������� = rsData!��������
                .Update
            End If
        End With
        rsData.MoveNext
    Loop
    
    rsData.Filter = "��־=1"
    rsData.Sort = "����,NO,ҩƷID"
    Do While Not rsData.EOF
        With rsSendRecipeDetail
            .AddNew
            !NO = rsData!NO
            !�շ�ID = rsData!�շ�ID
            !ҩƷID = rsData!ҩƷID
            !���� = rsData!����
            .Update
        End With
        
        rsData.MoveNext
    Loop
    
    arrSql = Array()
    
    mstrPrintRecipe = ""
    
    '�������������������ҩ
    rsSendRecipeByNo.Sort = "NO"
    rsSendRecipeByNo.MoveFirst
    For n = 1 To rsSendRecipeByNo.RecordCount
        '�ȼ���ִ��Ԥ����
        Call AutoAdjustPrice_ByNO(rsSendRecipeByNo!����, rsSendRecipeByNo!NO)

        rsSendRecipeDetail.Filter = "NO='" & rsSendRecipeByNo!NO & "'"
        rsSendRecipeDetail.MoveFirst
        
        If Val(rsSendRecipeByNo!��¼����) = 1 Or (Val(rsSendRecipeByNo!��¼����) = 2 And (Val(rsSendRecipeByNo!�����־) = 1 Or Val(rsSendRecipeByNo!�����־) = 4)) Then
            int���� = 1
        Else
            int���� = 2
        End If
        
        If mPrives.bln������ҩ���Ĵ��� = True And mParams.lngҩ��ID <> Val(rsSendRecipeByNo!ҩ��ID) Then
            gstrSQL = "Zl_ҩƷ�շ���¼_���Ŀⷿ("
            '�ֿⷿID
            gstrSQL = gstrSQL & mParams.lngҩ��ID
            '����
            gstrSQL = gstrSQL & "," & rsSendRecipeByNo!����
            'NO
            gstrSQL = gstrSQL & ",'" & rsSendRecipeByNo!NO & "'"
            'ԭ�ⷿID
            gstrSQL = gstrSQL & "," & Val(rsSendRecipeByNo!ҩ��ID)
            '����
            gstrSQL = gstrSQL & "," & int����
            '��������
            gstrSQL = gstrSQL & ",to_date('" & rsSendRecipeByNo!�������� & "','yyyy-MM-dd')"
            gstrSQL = gstrSQL & ")"
            
            ReDim Preserve arrSql(UBound(arrSql) + 1)
            arrSql(UBound(arrSql)) = gstrSQL
            
'            Call zldatabase.ExecuteProcedure(gstrSQL, Me.Caption & "-���Ŀⷿ")
        End If
        
        If mParams.IntCheckStock = 2 Then
            For lngRow = 1 To rsSendRecipeDetail.RecordCount
                gstrSQL = "zl_ҩƷ�շ���¼_��������("
                '�շ�ID
                gstrSQL = gstrSQL & rsSendRecipeDetail!�շ�ID
                'ҩƷID
                gstrSQL = gstrSQL & "," & rsSendRecipeDetail!ҩƷID
                '����
                gstrSQL = gstrSQL & "," & rsSendRecipeDetail!����
                gstrSQL = gstrSQL & ")"
                
                ReDim Preserve arrSql(UBound(arrSql) + 1)
                arrSql(UBound(arrSql)) = gstrSQL
'                Call zldatabase.ExecuteProcedure(gstrSQL, Me.Caption & "-��������")
                
                rsSendRecipeDetail.MoveNext
            Next
        End If
        
        gstrSQL = "zl_ҩƷ�շ���¼_������ҩ("
        '�ⷿID
        gstrSQL = gstrSQL & mParams.lngҩ��ID
        '����
        gstrSQL = gstrSQL & "," & rsSendRecipeByNo!����
        'NO
        gstrSQL = gstrSQL & ",'" & rsSendRecipeByNo!NO & "'"
        '��ҩ��(�����)
        gstrSQL = gstrSQL & ",'" & mstr����Ա & "'"
        '��ҩ��(���뾭����ҩ����ʱ������ҩ�˲���)
        gstrSQL = gstrSQL & "," & IIf(mParams.blnMustDosageProcess = True, "Null", IIf(rsSendRecipeByNo!��ҩ�� = "", "NULL", "'" & rsSendRecipeByNo!��ҩ�� & "'")) & ""
        'У���ˣ�����ҽ����
        gstrSQL = gstrSQL & "," & IIf(rsSendRecipeByNo!������ = "", "NULL", "'" & rsSendRecipeByNo!������ & "'") & ""
        '��ҩ��ʽ
        gstrSQL = gstrSQL & ",1"
        '��ҩʱ��
        gstrSQL = gstrSQL & ",to_date('" & date��ҩʱ�� & "','yyyy-MM-dd hh24:mi:ss') "
        '����Ա����
        gstrSQL = gstrSQL & ",'" & gstrUserCode & "'"
        '����Ա����
        gstrSQL = gstrSQL & ",'" & gstrUserName & "'"
        '����λ��
        gstrSQL = gstrSQL & "," & mParams.int����λ��
        '�Զ���˼��˵�
        gstrSQL = gstrSQL & "," & IIf(mParams.bln��˻��۵�, 1, 0)
        '�Ƿ�����
        gstrSQL = gstrSQL & "," & int����
        '�˲���
        gstrSQL = gstrSQL & ",'" & rsSendRecipeByNo!�˲��� & "'"
        '�����Ƿ�ʵ��ȡҩ
        gstrSQL = gstrSQL & "," & IIf(mblnδȡҩ��ҩ, 1, "Null")
        
        gstrSQL = gstrSQL & ")"
        
        ReDim Preserve arrSql(UBound(arrSql) + 1)
        arrSql(UBound(arrSql)) = gstrSQL
            
        '��¼�ô����ż���������
        mstrBill = rsSendRecipeByNo!NO & "|" & rsSendRecipeByNo!����
        mstrPrintRecipe = IIf(mstrPrintRecipe = "", "", mstrPrintRecipe & "|") & rsSendRecipeByNo!NO & "," & rsSendRecipeByNo!���� & "," & rsSendRecipeByNo!��¼���� & "," & rsSendRecipeByNo!�����־ & "," & rsSendRecipeByNo!�������� & "," & rsSendRecipeByNo!�շ����
        mfrmList.mstrLastName = rsSendRecipeByNo!����

        strNo = strNo & rsSendRecipeByNo!���� & "," & rsSendRecipeByNo!NO & "|"
        rsSendRecipeByNo.MoveNext
    Next
    
    mstr����Ա = ""
    mstr��ҩ�� = ""
    
'    '�ȴ���ҩ�����񣬷�ҩϵͳδ׼��������ʾ�ӿڷ�����Ϣ������Ա����ѡ���Ƿ�ҩ
'    If Not gobjPackerMZ Is Nothing And strNo <> "" Then
'        If gobjPackerMZ.HisUpload(mlngMode, 2, strNo, mParams.lngҩ��ID) = False Then
'            If MsgBox("�Զ���ҩϵͳδ׼���ã��Ƿ������ҩ��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
'                Exit Function
'            End If
'        End If
'    End If

    '�ȴ���ҩ�����񣬷�ҩϵͳδ׼��������ʾ�ӿڷ�����Ϣ������Ա����ѡ���Ƿ�ҩ
    If TypeName(mobjDrugMAC) = "clsDrugPacker" Then
        If mblnPackerConnect And strNo <> "" And mblnLoadDrug Then
            If mblnCompatible = False Then
                '�������½ӿڲ���
                If mobjDrugMAC.DYEY_MZ_TransRecipeList(mstrOpr, UserInfo.�û�����, UserInfo.�û�����, mParams.lngҩ��ID, Mid(strNo, 1, Len(strNo) - 1), strReturn) = False Then
                    If MsgBox("�Զ���ҩϵͳδ׼���ã��Ƿ������ҩ��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                        Exit Function
                    End If
                End If
            Else
                If mobjDrugMAC.DYEY_MZ_TransRecipeList(mstrOpr, UserInfo.�û�����, UserInfo.�û�����, mParams.lngҩ��ID, Mid(strNo, 1, Len(strNo) - 1), strReturn, IIf(mintAutoSendFlow = 0, mSendOper.StartSend, mSendOper.EndSend)) = False Then
                    If MsgBox("�Զ���ҩϵͳδ׼���ã��Ƿ������ҩ��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                        Exit Function
                    End If
                End If
            End If
        End If
    ElseIf TypeName(mobjDrugMAC) = "clsDrugMachine" Then
        If mblnPackerConnect Then
            If Right(strNo, 1) = "|" Then strNo = Left(strNo, Len(strNo) - 1)
            mobjDrugMAC.Operation gstrDbUser, Val("23-��ɷ�ҩ"), "1|" & Replace(strNo, "|", ";"), strMessage
'           If strMessage <> "" Then MsgBox strMessage, vbInformation, gstrSysName
        End If
    End If
    
    '���÷�ҩǰ����ҽӿ�
    err.Clear
    If Not mobjPlugIn Is Nothing Then
        If Right(strNo, 1) = "|" Then strNo = Left(strNo, Len(strNo) - 1)
        On Error Resume Next
        If mobjPlugIn.DrugBeforeSendByRecipe(mParams.lngҩ��ID, strNo, strReserve) = False Then
            If err.Number <> 0 Then
                err.Clear: On Error GoTo 0
            Else
                Exit Function
            End If
        End If
        err.Clear: On Error GoTo 0
    End If
    
    On Error GoTo ErrHand
    
    '���д���ҩ����
    gcnOracle.BeginTrans
    blnInTrans = True
    
    '�����������ǩ��
    '����������˵���ǩ��������Ҫ�Է�ҩ�˽��е���ǩ������
    If gblnESign������ҩ = True And gblnESignUserStoped = False Then
        rsSendRecipeByNo.MoveFirst
        For n = 1 To rsSendRecipeByNo.RecordCount
            strǩ����¼ = ""
            If GetSignatureRecored(EsignTache.send, rsSendRecipeByNo!����, rsSendRecipeByNo!NO, mParams.lngҩ��ID, strǩ����¼, 0, date��ҩʱ��, IIf(mstr����Ա = "", gstrUserName, mstr����Ա)) = False Then
                gcnOracle.RollbackTrans
                blnInTrans = False
                Exit Function
            End If
            
            If strǩ����¼ <> "" Then
                gstrSQL = "Zl_ҩƷǩ����¼_Insert(" & strǩ����¼ & ")"
                
                ReDim Preserve arrSql(UBound(arrSql) + 1)
                arrSql(UBound(arrSql)) = gstrSQL
            Else
                gcnOracle.RollbackTrans
                blnInTrans = False
                MsgBox "�Է�ҩ�˵���ǩ��ʧ�ܣ�", vbInformation, gstrSysName
                Exit Function
            End If
            
            rsSendRecipeByNo.MoveNext
        Next
    End If
    
    For i = 0 To UBound(arrSql)
        Call zldatabase.ExecuteProcedure(CStr(arrSql(i)), Me.Caption & "-��ҩ")
    Next
        
    gcnOracle.CommitTrans
    blnInTrans = False
    SendBatchRecipe = True
    
    '���÷�ҩ�����ҽӿ�
    If Not mobjPlugIn Is Nothing Then
        If Right(strNo, 1) = "|" Then strNo = Left(strNo, Len(strNo) - 1)
        On Error Resume Next
        mobjPlugIn.DrugSendByRecipe mParams.lngҩ��ID, strNo, date��ҩʱ��, strReserve
        err.Clear: On Error GoTo 0
    End If
    
    Exit Function
ErrHand:
    SendBatchRecipe = False
    If blnInTrans Then gcnOracle.RollbackTrans
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function CheckNotAudited(ByRef rsData As ADODB.Recordset) As Boolean
    Dim rsTmp As New ADODB.Recordset
    Dim bln�������� As Boolean
    Dim bln������ As Boolean
    
    On Error GoTo errHandle
    
    Call InitApplyforcredit
    
    CheckNotAudited = True
    bln�������� = True
    
    '��⵱ǰҩ���Ƿ�ΪסԺҩ�����������˳�������
    gstrSQL = "Select *" & vbNewLine & _
            "From ���ű� A, ��������˵�� B" & vbNewLine & _
            "Where a.Id = b.����id And a.Id = [1] And (b.�������� Like '%ҩ��' Or b.�������� Like '%ҩ��') And b.������� In (2, 3)"

    Set rsTmp = zldatabase.OpenSQLRecord(gstrSQL, "��⵱ǰҩ���Ƿ�ΪסԺҩ��", mParams.lngҩ��ID)
    If rsTmp.EOF Then Exit Function
    

    gstrSQL = "Select c.���� As ������������, c.����id" & vbNewLine & _
            "From ҩƷ�շ���¼ A, סԺ���ü�¼ B, ���˷������� C" & vbNewLine & _
            "Where a.����id = b.Id And b.Id = c.����id And a.ҩƷid = c.�շ�ϸĿid And a.Id = [1] And c.״̬ = 0"

    
    With rsData
        rsData.Filter = "��־=1"
        rsData.Sort = "����,NO"
    
        Do While Not .EOF
            Set rsTmp = zldatabase.OpenSQLRecord(gstrSQL, "����Ƿ������������δ��˵ĵ���", rsData!�շ�ID)

            If rsTmp.RecordCount > 0 Then
                bln�������� = False

                With mrsApplyforcredit
                    .AddNew
                    
                    !��־ = 1
                    !NO = rsData!NO
                    !ҩƷ���� = rsData!ҩ��
                    !���� = rsData!����
                    !���� = zlStr.FormatEx(rsData!���� / rsData!��װ, 4) & rsData!��λ
                    !������������ = zlStr.FormatEx(rsTmp!������������ / rsData!��װ, 4) & rsData!��λ
                    !���� = rsData!����
                    !�Ա� = rsData!�Ա�
                    !���� = rsData!����
                    !����ID = rsTmp!����ID
                    !�շ�ID = rsData!�շ�ID
                End With

            End If

            .MoveNext
        Loop
    End With

    '�Ժ�����������ĵ��ݽ��д���
    If bln�������� = False Then
        Call frm���ŷ�ҩ���������嵥.ShowCard(Me, mrsApplyforcredit, bln������, 1)

        '���Ӵ��巵���û��Ƿ����ִ�в���������ȡ�������ֹ��������
        CheckNotAudited = bln������
        If CheckNotAudited = False Then Exit Function
        
        '����ȡ�����͵ĵ��ݵ�ִ��״̬
        mrsApplyforcredit.Filter = "��־ = 0"
        
        If mrsApplyforcredit.RecordCount > 0 Then
            If mint��ҩ��ʽ = 1 Then
                CheckNotAudited = False
                Exit Function
            End If
            
            Do While Not mrsApplyforcredit.EOF
                rsData.Filter = "NO = '" & mrsApplyforcredit!NO & "'"
                If rsData.RecordCount > 0 Then
                    Do While Not rsData.EOF
                        rsData!��־ = 0
                        rsData.Update
                        rsData.MoveNext
                    Loop
                End If
                mrsApplyforcredit.MoveNext
            Loop
        End If

        rsData.Filter = ""
    End If
    
    Exit Function
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function CheckBatchRecipe(ByVal rsData As ADODB.Recordset) As Boolean
    Dim n As Integer
    Dim rstemp As ADODB.Recordset
    Dim blnFirst As Boolean
    Dim lngRow As Long, lngҩƷid As Long, LngID As Long, lng���� As Long, lng���� As Long
    Dim blnBatchSend As Boolean
    Dim i As Integer
    Dim dblSumMoney As Double
    Dim strRecipeString As String
    Dim rsCheck As ADODB.Recordset
    Dim arrRecipe
    Dim intCount As Integer
    Dim int��ǰ���� As Integer
    Dim str��ǰNO As String
    Dim str�շ����� As String
    Dim str�շ�ϸĿid As String
    Dim strTemp As String
    Dim str�˲��� As String
    
    On Error GoTo ErrHand
    If rsData!����ģʽ = 1 Then
        If gobjCharge Is Nothing Then
            Set gobjCharge = CreateObject("zl9OutExse.clsOutExse")
            If gobjCharge Is Nothing Then Exit Function
        End If
        
        If Not gobjCharge Is Nothing Then
            strTemp = BillHaveHerial(rsData!NO, rsData!����, 1, str�շ�ϸĿid, str�շ�����)
            If str�շ����� <> "" Then
                If Not gobjCharge.zlCheckExcuteItemValied(Me, gcnOracle, UserInfo.�û�����, glngSys, mlngMode, rsData!����ID, str�շ�����, rsData!NO, str�շ�ϸĿid) Then
                    CheckBatchRecipe = False
                    Exit Function
                End If
            End If
        End If
    End If
       
    '��鲡�˷������
    Set rsCheck = New ADODB.Recordset
    With rsCheck
        If .State = 1 Then .Close
        .Fields.Append "����", adLongVarChar, 8, adFldIsNullable
        .Fields.Append "NO", adLongVarChar, 8, adFldIsNullable
        .Fields.Append "����ID", adDouble, 18, adFldIsNullable
        .Fields.Append "ʵ�ս��", adDouble, 18, adFldIsNullable
        .Fields.Append "���շ�", adDouble, 1, adFldIsNullable
        .Fields.Append "��¼����", adDouble, 1, adFldIsNullable
        .Fields.Append "�����־", adDouble, 1, adFldIsNullable
        
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
    
    strRecipeString = mfrmList.GetCurrentBatchRecipe
    
    arrRecipe = Split(strRecipeString, "|")
    intCount = UBound(arrRecipe)
    
    For n = 0 To intCount
        rsCheck.AddNew
        rsCheck!���� = Val(Split(arrRecipe(n), ",")(0))
        rsCheck!NO = Split(arrRecipe(n), ",")(1)
        rsCheck!����ID = Val(Split(arrRecipe(n), ",")(2))
        rsCheck!ʵ�ս�� = Val(Split(arrRecipe(n), ",")(3))
        rsCheck!���շ� = Val(Split(arrRecipe(n), ",")(4))
        rsCheck!��¼���� = Val(Split(arrRecipe(n), ",")(5))
        rsCheck!�����־ = Val(Split(arrRecipe(n), ",")(6))
        rsCheck.Update
    Next
    If Not CheckSendBillMoney(rsCheck) Then Exit Function
    
    '���ҩƷ�洢�ⷿ
    If CheckDrugStock(rsData) = False Then Exit Function
    
    '���[סԺ����]�Ƿ������������δ��˵ĵ���
    If CheckNotAudited(rsData) = False Then Exit Function
    
    rsData.Filter = "��־=1"
    rsData.Sort = "����,NO"
    
    '��鵱ǰ�����������ڻ�����ҩ����δ��ҩ����
    Call CheckOtherUndeliveredDocuments(rsData!����ID)
    
    '���[�˲���]�Ƿ�Ϊ��
    str�˲��� = IIf(tbcDetail.Selected.index = 1, mfrmRecipe.Get�˲���, mfrmDetail.Get�˲���)
    If str�˲��� = "" Then
        MsgBox "�˲���Ϊ�գ�����ִ�з�ҩ������", vbInformation, gstrSysName
        Exit Function
    End If
    
    Do While Not rsData.EOF
        '����Ƿ�����
        If CheckBill(rsData!ҩ��ID, 3, rsData!����, rsData!NO, rsData!��¼����, rsData!�����־) <> 0 Then Exit Function
        
        '����Ƿ��շ�(��ҩ����)
        gstrSQL = " Select Decode(��ҩ��,Null,'','���ŷ�ҩ','',��ҩ��) ��ҩ��,���շ� From δ��ҩƷ��¼" & _
                 " Where No=[1] And (�ⷿID=[3] Or �ⷿID Is NULL) And ����=[2]"
        Set rstemp = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, rsData!NO, Val(rsData!����), Val(rsData!ҩ��ID))
        
        With rstemp
            If .EOF Then
                MsgBox "�ô����Ѿ�����������Ա����", vbInformation, gstrSysName
                CheckBatchRecipe = False
                Exit Function
            End If
            
            If mParams.blnMustDosageProcess = True Then
                If IsNull(!��ҩ��) Then
                    MsgBox "�ô�����δ��ҩ������ִ�з�ҩ������", vbInformation, gstrSysName
                    Exit Function
                End If
                If Trim(!��ҩ��) = "" Then
                    MsgBox "�ô�����δ��ҩ������ִ�з�ҩ������", vbInformation, gstrSysName
                    Exit Function
                End If
            End If

            mstr��ҩ�� = zlStr.NVL(!��ҩ��)
            
            If mParams.bln��ҩǰ�շѻ���� = False Then
                'δ�շѵĻ��۵�
                If rsData!���� = 8 And !���շ� = 0 And mParams.bln����δ�շѴ�����ҩ = False Then
                    MsgBox "�ô�����δ�շѣ�����ִ�з�ҩ������", vbInformation, gstrSysName
                    Exit Function
                End If
            
                'δ��˵ļ��˻��۵�
                If rsData!���� = 9 And !���շ� = 0 And mParams.bln����δ��˴�����ҩ = False Then
                    MsgBox "�ô�����δ��ˣ�����ִ�з�ҩ������", vbInformation, gstrSysName
                    Exit Function
                End If
            End If
            
            If Not IsReceiptBalance_Charge(0, mstrPrivs, rsData!����, rsData!NO, rsData!���, rsData!��¼����, rsData!�����־) Then Exit Function
            If Not IsOutPatient(mstrPrivs, rsData!����, rsData!NO, rsData!��¼����, rsData!�����־) Then Exit Function
            
            mrsList.Filter = "����=" & rsData!���� & " And NO='" & rsData!NO & "' "
            If Not mrsList.EOF Then dblSumMoney = Val(mrsList!���)
            
            If Not CheckBillControl(mcondition.intListType + 1, rsData!����, rsData!NO, dblSumMoney) Then Exit Function
            
            'У�鷢ҩ��
            If mParams.intУ�鷢ҩ�� = 1 And Not blnFirst Then
                mstr����Ա = zldatabase.UserIdentify(Me, "У�鷢ҩ��", glngSys, 1341, "��ҩ")
                blnFirst = True
            Else
                mstr����Ա = gstrUserName
            End If
            If mstr����Ա = "" Then Exit Function
        End With
        
        rsData.MoveNext
    Loop
        
    '�������
    rsData.Sort = "����,NO"
    rsData.MoveFirst
    Do While Not rsData.EOF
        If int��ǰ���� <> rsData!���� And str��ǰNO <> rsData!NO Then
            int��ǰ���� = rsData!����
            str��ǰNO = rsData!NO
                                    
            '���۹���
            If CheckPriceAdjustByNO(Val(rsData!����), Val(rsData!ҩ��ID), rsData!NO) = False Then
                Exit Function
            End If
            
            '����ҩƷ��ʾ
            If Not CheckSpec(rsData!ҩ��ID, rsData!NO, rsData!����) Then Exit Function
            
            If mstr��������ʾ <> "" Then
                If MsgBox("����Ϊ[" & rsData!NO & "]" & "�Ĵ����к������¶�����ҩƷ��ȷ����ҩ��" & mstr��������ʾ, vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then Exit Function
            End If
            
            '�����
            If Not CheckStock(rsData!ҩ��ID, rsData!NO, rsData!����) Then Exit Function
        End If
        
        rsData.MoveNext
    Loop
    
    '��ҩʱ�������ѿ�ȷ��ֻ֧��һ������ģʽ
    If CheckPati(rsData) = False Then Exit Function
    
    CheckBatchRecipe = True
    Exit Function
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    CheckBatchRecipe = False
End Function

Private Function CheckStock(ByVal lngNOҩ��id As Long, ByVal strNo As String, ByVal IntBillStyle As Integer) As Boolean
    Dim RecCheckStock As New ADODB.Recordset, RecBillData As New ADODB.Recordset
    Dim dblStock As Double, intCheck As Integer
    Dim dblUsableStock As Double
    '--�����--
    '0-�����;1-���,��������;2-���,�����ֹ
'    On Error Resume Next
'    err = 0
    On Error GoTo errHandle
    CheckStock = False
    intCheck = mParams.IntCheckStock
    
    '���м��
    If intCheck <> 0 Then
        gstrSQL = " SELECT A.ҩƷID,SUM(NVL(A.ʵ������,0)*NVL(A.����,1)) ����," & _
                " '['||C.����||']'||" & IIf(gintҩƷ������ʾ = 1, "NVL(L.����,C.����)", "C.����") & " Ʒ��,NVL(A.����,0) ����, Nvl(A.����,'') ���� " & _
                " FROM ҩƷ�շ���¼ A,ҩƷ��� B,�շ���ĿĿ¼ C,�շ���Ŀ���� L " & _
                " WHERE A.ҩƷID=B.ҩƷID AND B.ҩƷID=C.ID" & _
                " AND B.ҩƷID=L.�շ�ϸĿID(+) AND L.����(+)=3 AND L.����(+)=1 " & _
                " AND A.����� IS NULL AND MOD(A.��¼״̬,3)=1 AND NVL(A.ժҪ,'С��')<>'�ܷ�'" & _
                " AND A.NO=[1] AND A.����=[2] AND (A.�ⷿID+0=[3] OR A.�ⷿID IS NULL) " & _
                " GROUP BY A.ҩƷID,'['||C.����||']'||" & IIf(gintҩƷ������ʾ = 1, "NVL(L.����,C.����)", "C.����") & ",���� ,A.����"
        Set RecBillData = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, strNo, IntBillStyle, lngNOҩ��id)
        
        With RecBillData
            Do While Not .EOF
                gstrSQL = " Select nvl(��������,0) AS ��������, nvl(ʵ������,0) AS ʵ������ " & _
                         " From ҩƷ��� " & _
                         " Where �ⷿID+0=[1] And ҩƷID=[2] " & _
                         " And ����=1 And Nvl(����,0)=[3]"
                Set RecCheckStock = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, mParams.lngҩ��ID, CLng(RecBillData!ҩƷID), CLng(RecBillData!����))
                
                With RecCheckStock
                    If .EOF Then
                        dblStock = 0
                        dblUsableStock = 0
                    Else
                        dblStock = !ʵ������
                        dblUsableStock = !��������
                    End If
                    
                    '����Ǵ�������ҩ������(����ҩ���͵�ǰҩ����һ��ʱ)�����Ҫ���ʵ��������ҲҪ����������
                    If dblStock < RecBillData!���� Or (lngNOҩ��id <> mParams.lngҩ��ID And dblUsableStock < RecBillData!����) Then
                        If RecBillData!���� > 0 And NVL(RecBillData!����, "") <> "" Then
                            Select Case intCheck
                                Case 1
                                    If MsgBox(RecBillData!Ʒ�� & "����Ϊ[" & RecBillData!���� & "]�Ŀ�����������Ƿ������ҩ��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
                                Case 2
                                    MsgBox RecBillData!Ʒ�� & "����Ϊ[" & RecBillData!���� & "]�Ŀ�������������ܼ�����ҩ��", vbInformation, gstrSysName: Exit Function
                            End Select
                        Else
                            Select Case intCheck
                                Case 1
                                    If MsgBox(RecBillData!Ʒ�� & "�Ŀ�����������Ƿ������ҩ��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
                                Case 2
                                    MsgBox RecBillData!Ʒ�� & "�Ŀ�������������ܼ�����ҩ��", vbInformation, gstrSysName: Exit Function
                            End Select
                        End If
                    End If
                End With
                .MoveNext
            Loop
        End With
    End If
    
    If err <> 0 Then
        MsgBox "�����ʱ����������Ԥ֪�Ĵ���", vbInformation, gstrSysName
        Exit Function
    End If
    CheckStock = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Private Function CheckSpec(ByVal lngNOҩ��id As Long, ByVal strNo As String, ByVal IntBillStyle As Integer) As Boolean
    Dim strNote As String
    Dim rstemp As New ADODB.Recordset
    
    mstr��������ʾ = ""
    
    '�Զ�����ҩƷ���м��
    On Error GoTo errHandle
    gstrSQL = " SELECT Distinct " & _
        " '['||C.����||']'||" & IIf(gintҩƷ������ʾ = 1, "NVL(L.����,C.����)", "C.����") & " Ʒ��,X.�������" & _
        " FROM ҩƷ�շ���¼ A,ҩƷ��� B,�շ���ĿĿ¼ C,�շ���Ŀ���� L,ҩƷ���� X " & _
        " WHERE A.ҩƷID=B.ҩƷID AND B.ҩ��ID=X.ҩ��ID And B.ҩƷID=C.ID " & _
        " AND B.ҩƷID=L.�շ�ϸĿID(+) AND L.����(+)=3 AND L.����(+)=1 " & _
        " AND A.����� IS NULL AND MOD(A.��¼״̬,3)=1 AND NVL(A.ժҪ,'С��')<>'�ܷ�'" & _
        " AND A.NO=[1] AND A.����=[2] AND (A.�ⷿID+0=[3] OR A.�ⷿID IS NULL) " & _
        " And X.�������<>'��ͨҩ'" & _
        " Order by X.�������"
    Set rstemp = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption & "[�Զ�����ҩƷ���м��]", strNo, IntBillStyle, lngNOҩ��id)
    
    If rstemp.RecordCount = 0 Then
        CheckSpec = True
        Exit Function
    End If
    
    With rstemp
        Do While Not .EOF
            strNote = strNote & vbCrLf & Space(4) & !������� & "-" & !Ʒ��
            .MoveNext
        Loop
    End With
'    If MsgBox("�Ƿ�����¶����顢������ҩƷ���з�ҩ��" & strNote, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
    mstr��������ʾ = strNote
    CheckSpec = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Private Function CheckDrugStock(ByVal rsData As ADODB.Recordset) As Boolean
    Dim lngҩƷid As Long
    
    If mrsDrugStock Is Nothing Then
        GetDrugStock mParams.lngҩ��ID
        If mrsDrugStock Is Nothing Then
            MsgBox "δ���ô洢�ⷿ�����ܷ�ҩ��", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    rsData.Filter = "��־=1"
    rsData.Sort = "ҩƷID"
    
    Do While Not rsData.EOF
        If lngҩƷid <> rsData!ҩƷID Then
            lngҩƷid = rsData!ҩƷID
            
            mrsDrugStock.Filter = "ҩƷID=" & lngҩƷid
            If mrsDrugStock.EOF Then
                MsgBox rsData!Ʒ�� & "δ���ô洢�ⷿ�����ܷ�ҩ��", vbInformation, gstrSysName
                Exit Function
            End If
        End If
        rsData.MoveNext
    Loop

    CheckDrugStock = True
End Function

Private Function CheckSendBillMoney(ByVal rsData As ADODB.Recordset) As Boolean
    '��ҩ��飭��鲡�˷����������ݼ��ʱ�����������Ӧ����
    'blnBatch��True-������ҩ;False-��������ҩ
    '��Ҫ�㷨��
    '1��ϵͳ����"ִ�к��Զ����"��Чʱ�ż��
    '2��ֻ�Լ��ʻ��۵�
    '3��������ID���㵥�ݻ��ܽ��
    '4�����ݼ��ʱ�����������Ӧ����
    Dim n As Integer
    Dim rsTmp As ADODB.Recordset
    Dim rs������� As ADODB.Recordset
    Dim strNo As String
    Dim lng����ID As Long
    Dim cur������� As Currency
    
    Dim strFirstNo As String
    Dim str������� As String
    Dim str��������� As String
    
    On Error GoTo errH
    
    'ϵͳ����"ִ�к��Զ����"��Чʱ�ż��
    If mParams.bln��˻��۵� = False Then
        CheckSendBillMoney = True
        Exit Function
    End If
    
    If rsData Is Nothing Then
        CheckSendBillMoney = True
        Exit Function
    End If
    
    With rsData
        'ֻ�Լ��ʻ��۵��ż��
        .Filter = "����=9 And ���շ�=0"
        
        '������ID���㵥�ݻ��ܽ��
        .Sort = "����ID"
        
        If .RecordCount = 0 Then
            CheckSendBillMoney = True
            Exit Function
        End If
        
        .MoveFirst
        
        '���ݼ��ʱ�����������Ӧ����
        Do While Not .EOF
            If lng����ID <> Val(!����ID) Then
                If lng����ID <> 0 Then
                    '�ж���סԺ�������ﲡ��
                    If !��¼���� = 1 Or (!��¼���� = 2 And (!�����־ = 1 Or !�����־ = 4)) Then
                        gstrSQL = "Select Distinct '����' As ��Դ, " & _
                            " B.����id,0 ��ҳid,0 ���˲���id, C.���� " & _
                            " From ҩƷ�շ���¼ A,������ü�¼ B,������Ϣ C " & _
                            " Where A.����id=B.Id And b.����id = c.����id " & _
                            " And A.����=9 And A.no=[1] "
                    Else
                        gstrSQL = " Select Distinct 'סԺ' As ��Դ, " & _
                            " B.����id,nvl(B.��ҳid,0) ��ҳid,B.���˲���id, C.���� " & _
                            " From ҩƷ�շ���¼ A,סԺ���ü�¼ B,������Ϣ C " & _
                            " Where A.����id=B.Id And b.����id = c.����id " & _
                            " And A.����=9 And A.no=[1] "
                    End If
                    Set rsTmp = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, strFirstNo)
                    
                    'ȡ�������
                    gstrSQL = " Select /*+ Rule*/ Distinct b.����, b.���� " & _
                    " From ������ü�¼ a, �շ���Ŀ��� b, ҩƷ�շ���¼ c,Table(Cast(f_Str2list([1]) As zlTools.t_Strlist)) d " & _
                    " Where a.�շ���� = b.���� And a.Id = c.����id And c.���� = 9 And c.No=d.Column_Value "
                    If !��¼���� = 1 Or (!��¼���� = 2 And (!�����־ = 1 Or !�����־ = 4)) Then
                    Else
                        gstrSQL = Replace(gstrSQL, "������ü�¼", "סԺ���ü�¼")
                    End If
                    Set rs������� = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, strNo)
                    
                    Do While Not rs�������.EOF
                        str������� = str������� & rs�������!����
                        str��������� = str��������� & "," & rs�������!����
                        rs�������.MoveNext
                    Loop
                                        
                    '���������
                    If Not FinishBillingWarn(rsTmp, cur�������, str�������, str���������) Then
                        CheckSendBillMoney = False
                        Exit Function
                    End If
                End If
                
                strNo = !NO
                cur������� = Val(Getʵ�ս��(Val(!����), !NO, !�����־))
                strFirstNo = !NO
                lng����ID = Val(!����ID)
            Else
                strNo = strNo & "," & !NO
                cur������� = cur������� + Val(Getʵ�ս��(Val(!����), !NO, !�����־))
            End If
            
            .MoveNext
            
            If .EOF Then
                .MovePrevious
                '�ж���סԺ�������ﲡ��
                If !��¼���� = 1 Or (!��¼���� = 2 And (!�����־ = 1 Or !�����־ = 4)) Then
                    gstrSQL = "Select Distinct '����' As ��Դ, " & _
                        " B.����id,0 ��ҳid,0 ���˲���id, C.���� " & _
                        " From ҩƷ�շ���¼ A,������ü�¼ B,������Ϣ C " & _
                        " Where A.����id=B.Id And b.����id = c.����id " & _
                        " And A.����=9 And A.no=[1] "
                Else
                    gstrSQL = "Select Distinct 'סԺ' As ��Դ, " & _
                        " B.����id,nvl(B.��ҳid,0) ��ҳid,B.���˲���id, C.���� " & _
                        " From ҩƷ�շ���¼ A,סԺ���ü�¼ B,������Ϣ C " & _
                        " Where A.����id=B.Id And b.����id = c.����id " & _
                        " And A.����=9 And A.no=[1] "
                End If
                Set rsTmp = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, strFirstNo)
                
                'ȡ�������
                gstrSQL = " Select /*+ Rule*/ Distinct b.����, b.���� " & _
                    " From ������ü�¼ a, �շ���Ŀ��� b, ҩƷ�շ���¼ c,Table(Cast(f_Str2list([1]) As zlTools.t_Strlist)) d " & _
                    " Where a.�շ���� = b.���� And a.Id = c.����id And c.���� = 9 And c.No=d.Column_Value "
                If !��¼���� = 1 Or (!��¼���� = 2 And (!�����־ = 1 Or !�����־ = 4)) Then
                Else
                    gstrSQL = Replace(gstrSQL, "������ü�¼", "סԺ���ü�¼")
                End If
                Set rs������� = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, strNo)
                
                Do While Not rs�������.EOF
                    str������� = str������� & rs�������!����
                    str��������� = str��������� & "," & rs�������!����
                    rs�������.MoveNext
                Loop
                                    
                '���������
                If Not FinishBillingWarn(rsTmp, cur�������, str�������, str���������) Then
                    CheckSendBillMoney = False
                    Exit Function
                End If
                
                .MoveNext
            End If
        Loop
    End With
    
    CheckSendBillMoney = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function FinishBillingWarn(ByVal rsTmp As ADODB.Recordset, ByVal cur��� As Currency, ByVal str��� As String, ByVal str����� As String) As Boolean
'���ܣ���ִ��������Զ���˵ķ���ʱ���Բ��˷��ý��м��ʱ�����
'������objRecord=����Ҫ���ִ�еĲ�����Ϣ��������
'      str���="CDE..."����������漰�����շ����
'      str�����="���,����,..."����Ӧ�������������ʾ
    Dim rsPati As ADODB.Recordset
    Dim rsWarn As ADODB.Recordset
    Dim strWarn As String, intWarn As Integer
    Dim strsql As String, intR As Integer, i As Long
    Dim cur���� As Currency
    
    On Error GoTo errH
    
    If rsTmp!��Դ.Value = "סԺ" Then
        'סԺ���˱���
        strsql = _
            " Select ����ID,Ԥ�����,�������,0 as Ԥ����� From ������� Where ����=1 And ����=2 And ����ID=[1]" & _
            " Union ALL" & _
            " Select A.����ID,0,0,Sum(���) From ����ģ����� A,������ҳ B" & _
            " Where A.����ID=B.����ID And A.��ҳID=B.��ҳID And B.���� Is Not Null And A.����ID=[1] And A.��ҳID=[2] Group by A.����ID"
        strsql = "Select ����ID,Nvl(Sum(Ԥ�����),0)-Nvl(Sum(�������),0)+Nvl(Sum(Ԥ�����),0) as ʣ��� From (" & strsql & ") Group by ����ID"
        
        strsql = "Select zl_PatiWarnScheme(A.����ID,B.��ҳID) As ���ò���,C.ʣ���," & _
            " Decode(A.������,Null,Null,zl_PatientSurety(A.����ID,B.��ҳID)) as ������" & _
            " From ������Ϣ A,������ҳ B,(" & strsql & ") C" & _
            " Where A.����ID=B.����ID And A.��ҳid=B.��ҳid And A.����ID=C.����ID(+)" & _
            " And A.����ID=[1] And B.��ҳID=[2]"
        Set rsPati = zldatabase.OpenSQLRecord(strsql, Me.Caption, Val(rsTmp!����ID), Val(rsTmp!��ҳid))
    Else
        '���������ﱨ��
        strsql = "Select ����ID,Ԥ�����,������� From ������� Where ����=1 And ����=1 And ����ID=[1]"
        strsql = "Select zl_PatiWarnScheme(A.����ID) As ���ò���,A.������," & _
            " Nvl(B.Ԥ�����,0)-Nvl(B.�������,0)+Nvl(E.�ʻ����,0) as ʣ���" & _
            " From ������Ϣ A,(" & strsql & ") B,ҽ�����˹����� D,ҽ�����˵��� E" & _
            " Where A.����ID=B.����ID(+) " & _
            " And A.����id = D.����id(+) And A.����=D.����(+) And D.����=E.����(+) And D.ҽ����=E.ҽ����(+) And D.��־(+)=1" & _
            " And A.����ID=[1]"
        Set rsPati = zldatabase.OpenSQLRecord(strsql, Me.Caption, Val(rsTmp!����ID))
    End If
    
    intWarn = -1 '���ʱ���ʱȱʡҪ��ʾ
    'ִ�б���:���ﲡ�˲���ID=0
    strsql = "Select Nvl(��������,1) as ��������," & _
        " ����ֵ,������־1,������־2,������־3 From ���ʱ�����" & _
        " Where Nvl(����ID,0)=[1] And ���ò���=[2]"
    Set rsWarn = zldatabase.OpenSQLRecord(strsql, Me.Caption, Val(rsTmp!���˲���ID), CStr(zlStr.NVL(rsPati!���ò���)))
    If Not rsWarn.EOF Then
        If rsWarn!�������� = 2 Then cur���� = GetPatiDayMoney(Val(rsTmp!����ID))
        str����� = Mid(str�����, 2)
        For i = 1 To Len(str���)
            intR = BillingWarn(Me, mstrPrivs, rsWarn, rsTmp!����, zlStr.NVL(rsPati!ʣ���, 0), cur����, cur���, zlStr.NVL(rsPati!������, 0), Mid(str���, i, 1), Split(str�����, ",")(i - 1), strWarn, intWarn)
            If InStr(",2,3,", intR) > 0 Then Exit Function
        Next
    End If
    
    FinishBillingWarn = True
    Exit Function
errH:
    
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function BillingWarn(frmParent As Object, ByVal strPrivs As String, _
    rsWarn As ADODB.Recordset, ByVal str���� As String, ByVal curʣ���� As Currency, _
    ByVal cur���ս�� As Currency, ByVal Cur���ʽ�� As Currency, ByVal cur������� As Currency, _
    ByVal str�շ���� As String, ByVal str������� As String, str�ѱ���� As String, _
    intWarn As Integer) As Integer
'����:�Բ��˼��ʽ��б�����ʾ
'����:rsWarn=���������������õļ�¼��(�ò��˲���,�����ֺ���ҽ��)
'     str�շ����=��ǰҪ�������,���ڷ��౨��
'     str�������=�������,������ʾ
'     intWarn=�Ƿ���ʾѯ���Ե���ʾ,-1=Ҫ��ʾ,0=ȱʡΪ��,1-ȱʡΪ��
'����:str�ѱ����="CDE":�����ڱ��α�����һ�����,"-"Ϊ������𡣸÷������ڴ����ظ�����
'     intWarn=����ѯ������ʾ�е�ѡ����,0=Ϊ��,1-Ϊ��
'     0;û�б���,����
'     1:������ʾ���û�ѡ�����
'     2:������ʾ���û�ѡ���ж�
'     3:������ʾ�����ж�
'     4:ǿ�Ƽ��ʱ���,����
    Dim bln�ѱ��� As Boolean, byt��־ As Byte
    Dim byt��ʽ As Byte, byt�ѱ���ʽ As Byte
    Dim ArrTmp As Variant, vMsg As VbMsgBoxResult
    Dim str���� As String, i As Long
    
    BillingWarn = 0
    
    '�����������:NULL��û������,0�������˵�
    If rsWarn.State = 0 Then Exit Function
    If rsWarn.EOF Then Exit Function
    If IsNull(rsWarn!����ֵ) Then Exit Function
    
    '��Ӧ���λ��Ч��������
    If Not IsNull(rsWarn!������־1) Then
        If rsWarn!������־1 = "-" Or InStr(rsWarn!������־1, str�շ����) > 0 Then byt��־ = 1
        If rsWarn!������־1 = "-" Then str������� = "" '�������ʱ,������ʾ��������
    End If
    If byt��־ = 0 And Not IsNull(rsWarn!������־2) Then
        If rsWarn!������־2 = "-" Or InStr(rsWarn!������־2, str�շ����) > 0 Then byt��־ = 2
        If rsWarn!������־2 = "-" Then str������� = "" '�������ʱ,������ʾ��������
    End If
    If byt��־ = 0 And Not IsNull(rsWarn!������־3) Then
        If rsWarn!������־3 = "-" Or InStr(rsWarn!������־3, str�շ����) > 0 Then byt��־ = 3
        If rsWarn!������־3 = "-" Then str������� = "" '�������ʱ,������ʾ��������
    End If
    If byt��־ = 0 Then Exit Function '����Ч����
    
    '������־2ʵ�����������жϢ٢�,����ֻ��һ���жϢ�
    '���ִ����ǰ����һ�����ֻ������һ�ֱ�����ʽ(������������ʱ)
    'ʾ����"-" �� ",ABC,567,DEF"
    '������־2ʾ����"-��" �� ",ABC��,567��,DEF��"
    bln�ѱ��� = InStr(str�ѱ����, str�շ����) > 0 Or str�ѱ���� Like "-*"
    
    If bln�ѱ��� Then '��intWarn = -1ʱ,Ҳ��ǿ���ٱ���
        If byt��־ = 2 Then
            If str�ѱ���� Like "-*" Then
                byt�ѱ���ʽ = IIf(Right(str�ѱ����, 1) = "��", 2, 1)
            Else
                ArrTmp = Split(str�ѱ����, ",")
                For i = 0 To UBound(ArrTmp)
                    If InStr(ArrTmp(i), str�շ����) > 0 Then
                        byt�ѱ���ʽ = IIf(Right(ArrTmp(i), 1) = "��", 2, 1)
                        'Exit For 'ȡ��˵����סԺ����ģ��
                    End If
                Next
            End If
        Else
            Exit Function
        End If
    End If
    
    If str������� <> "" Then str������� = """" & str������� & """����"
    str���� = IIf(cur������� = 0, "", "(��������:" & Format(cur�������, "0.00") & ")")
    curʣ���� = curʣ���� + cur������� - Cur���ʽ��
    cur���ս�� = cur���ս�� + Cur���ʽ��
        
    '---------------------------------------------------------------------
    If rsWarn!�������� = 1 Then  '�ۼƷ��ñ���(����)
        Select Case byt��־
            Case 1 '���ڱ���ֵ(����Ԥ����ľ�)��ʾѯ�ʼ���
                If curʣ���� < rsWarn!����ֵ Then
                    If InStr(";" & strPrivs & ";", ";ǿ�Ƽ���;") = 0 Then
                        If intWarn = -1 Then
                            vMsg = frmMsgBox.ShowMsgBox(str���� & " ��ǰʣ���" & str���� & ":" & Format(curʣ����, "0.00") & ",����" & str������� & "����ֵ:" & Format(rsWarn!����ֵ, "0.00") & ",����ò��˼�����", frmParent)
                            If vMsg = vbNo Or vMsg = vbCancel Then
                                If vMsg = vbCancel Then intWarn = 0
                                BillingWarn = 2
                            ElseIf vMsg = vbYes Or vMsg = vbIgnore Then
                                If vMsg = vbIgnore Then intWarn = 1
                                BillingWarn = 1
                            End If
                        Else
                            If intWarn = 0 Then
                                BillingWarn = 2
                            ElseIf intWarn = 1 Then
                                BillingWarn = 1
                            End If
                        End If
                    Else
                        If intWarn = -1 Then
                            vMsg = frmMsgBox.ShowMsgBox("ǿ�Ƽ�������:" & vbCrLf & vbCrLf & str���� & " ��ǰʣ���" & str���� & ":" & Format(curʣ����, "0.00") & " ����" & str������� & "����ֵ:" & Format(rsWarn!����ֵ, "0.00") & "��", frmParent, True)
                            If vMsg = vbIgnore Then intWarn = 1
                        End If
                        BillingWarn = 4
                    End If
                End If
            Case 2 '���ڱ���ֵ��ʾѯ�ʼ���,Ԥ����ľ�ʱ��ֹ����
                If Not bln�ѱ��� Then
                    If curʣ���� < 0 Then
                        byt��ʽ = 2
                        If InStr(";" & strPrivs & ";", ";ǿ�Ƽ���;") = 0 Then
                            If intWarn = -1 Then
                                vMsg = frmMsgBox.ShowMsgBox(str���� & " ��ǰʣ���" & str���� & "�Ѿ��ľ�," & str������� & "��ֹ���ʡ�", frmParent, True)
                                If vMsg = vbIgnore Then intWarn = 1
                            End If
                            BillingWarn = 3
                        Else
                            If intWarn = -1 Then
                                vMsg = frmMsgBox.ShowMsgBox(str������� & "ǿ�Ƽ�������:" & vbCrLf & vbCrLf & str���� & " ��ǰʣ���" & str���� & "�Ѿ��ľ���", frmParent, True)
                                If vMsg = vbIgnore Then intWarn = 1
                            End If
                            BillingWarn = 4
                        End If
                    ElseIf curʣ���� < rsWarn!����ֵ Then
                        byt��ʽ = 1
                        If InStr(";" & strPrivs & ";", ";ǿ�Ƽ���;") = 0 Then
                            If intWarn = -1 Then
                                vMsg = frmMsgBox.ShowMsgBox(str���� & " ��ǰʣ���" & str���� & ":" & Format(curʣ����, "0.00") & ",����" & str������� & "����ֵ:" & Format(rsWarn!����ֵ, "0.00") & ",����ò��˼�����", frmParent)
                                If vMsg = vbNo Or vMsg = vbCancel Then
                                    If vMsg = vbCancel Then intWarn = 0
                                    BillingWarn = 2
                                ElseIf vMsg = vbYes Or vMsg = vbIgnore Then
                                    If vMsg = vbIgnore Then intWarn = 1
                                    BillingWarn = 1
                                End If
                            Else
                                If intWarn = 0 Then
                                    BillingWarn = 2
                                ElseIf intWarn = 1 Then
                                    BillingWarn = 1
                                End If
                            End If
                        Else
                            If intWarn = -1 Then
                                vMsg = frmMsgBox.ShowMsgBox("ǿ�Ƽ�������:" & vbCrLf & vbCrLf & str���� & " ��ǰʣ���" & str���� & ":" & Format(curʣ����, "0.00") & ",����" & str������� & "����ֵ:" & Format(rsWarn!����ֵ, "0.00") & "��", frmParent, True)
                                If vMsg = vbIgnore Then intWarn = 1
                            End If
                            BillingWarn = 4
                        End If
                    End If
                Else
                    '�ϴ��ѱ�����ѡ�������ǿ�Ƽ���
                    If byt�ѱ���ʽ = 1 Then
                        '�ϴε��ڱ���ֵ��ѡ�������ǿ�Ƽ���,���ٴ�����ڵ����,������Ҫ�ж�Ԥ�����Ƿ�ľ�
                        If curʣ���� < 0 Then
                            byt��ʽ = 2
                            If InStr(";" & strPrivs & ";", ";ǿ�Ƽ���;") = 0 Then
                                If intWarn = -1 Then
                                    vMsg = frmMsgBox.ShowMsgBox(str���� & " ��ǰʣ���" & str���� & "�Ѿ��ľ�," & str������� & "��ֹ���ʡ�", frmParent, True)
                                    If vMsg = vbIgnore Then intWarn = 1
                                End If
                                BillingWarn = 3
                            Else
                                If intWarn = -1 Then
                                    vMsg = frmMsgBox.ShowMsgBox(str������� & "ǿ�Ƽ�������:" & vbCrLf & vbCrLf & str���� & " ��ǰʣ���" & str���� & "�Ѿ��ľ���", frmParent, True)
                                    If vMsg = vbIgnore Then intWarn = 1
                                End If
                                BillingWarn = 4
                            End If
                        End If
                    ElseIf byt�ѱ���ʽ = 2 Then
                        '�ϴ�Ԥ�����Ѿ��ľ���ǿ�Ƽ���,���ٴ���
                        Exit Function
                    End If
                End If
            Case 3 '���ڱ���ֵ��ֹ����
                If curʣ���� < rsWarn!����ֵ Then
                    If InStr(";" & strPrivs & ";", ";ǿ�Ƽ���;") = 0 Then
                        If intWarn = -1 Then
                            vMsg = frmMsgBox.ShowMsgBox(str���� & " ��ǰʣ���" & str���� & ":" & Format(curʣ����, "0.00") & ",����" & str������� & "����ֵ:" & Format(rsWarn!����ֵ, "0.00") & ",��ֹ���ʡ�", frmParent, True)
                            If vMsg = vbIgnore Then intWarn = 1
                        End If
                        BillingWarn = 3
                    Else
                        If intWarn = -1 Then
                            vMsg = frmMsgBox.ShowMsgBox("ǿ�Ƽ�������:" & vbCrLf & vbCrLf & str���� & " ��ǰʣ���" & str���� & ":" & Format(curʣ����, "0.00") & ",����" & str������� & "����ֵ:" & Format(rsWarn!����ֵ, "0.00") & "��", frmParent, True)
                            If vMsg = vbIgnore Then intWarn = 1
                        End If
                        BillingWarn = 4
                    End If
                End If
        End Select
    ElseIf rsWarn!�������� = 2 Then  'ÿ�շ��ñ���(����)
        Select Case byt��־
            Case 1 '���ڱ���ֵ��ʾѯ�ʼ���
                If cur���ս�� > rsWarn!����ֵ Then
                    If InStr(";" & strPrivs & ";", ";ǿ�Ƽ���;") = 0 Then
                        If intWarn = -1 Then
                            vMsg = frmMsgBox.ShowMsgBox(str���� & " ���շ���:" & Format(cur���ս��, gtype_UserSysParms.P9_���ý���λ��) & ",����" & str������� & "����ֵ:" & Format(rsWarn!����ֵ, "0.00") & ",����ò��˼�����", frmParent)
                            If vMsg = vbNo Or vMsg = vbCancel Then
                                If vMsg = vbCancel Then intWarn = 0
                                BillingWarn = 2
                            ElseIf vMsg = vbYes Or vMsg = vbIgnore Then
                                If vMsg = vbIgnore Then intWarn = 1
                                BillingWarn = 1
                            End If
                        Else
                            If intWarn = 0 Then
                                BillingWarn = 2
                            ElseIf intWarn = 1 Then
                                BillingWarn = 1
                            End If
                        End If
                    Else
                        If intWarn = -1 Then
                            vMsg = frmMsgBox.ShowMsgBox("ǿ�Ƽ�������:" & vbCrLf & vbCrLf & str���� & " ���շ���:" & Format(cur���ս��, gtype_UserSysParms.P9_���ý���λ��) & ",����" & str������� & "����ֵ:" & Format(rsWarn!����ֵ, "0.00") & "��", frmParent, True)
                            If vMsg = vbIgnore Then intWarn = 1
                        End If
                        BillingWarn = 4
                    End If
                End If
            Case 3 '���ڱ���ֵ��ֹ����
                If cur���ս�� > rsWarn!����ֵ Then
                    If InStr(";" & strPrivs & ";", ";ǿ�Ƽ���;") = 0 Then
                        If intWarn = -1 Then
                            vMsg = frmMsgBox.ShowMsgBox(str���� & " ���շ���:" & Format(cur���ս��, gtype_UserSysParms.P9_���ý���λ��) & ",����" & str������� & "����ֵ:" & Format(rsWarn!����ֵ, "0.00") & ",��ֹ���ʡ�", frmParent, True)
                            If vMsg = vbIgnore Then intWarn = 1
                        End If
                        BillingWarn = 3
                    Else
                        If intWarn = -1 Then
                            vMsg = frmMsgBox.ShowMsgBox("ǿ�Ƽ�������:" & vbCrLf & vbCrLf & str���� & " ���շ���:" & Format(cur���ս��, gtype_UserSysParms.P9_���ý���λ��) & ",����" & str������� & "����ֵ:" & Format(rsWarn!����ֵ, "0.00") & "��", frmParent, True)
                            If vMsg = vbIgnore Then intWarn = 1
                        End If
                        BillingWarn = 4
                    End If
                End If
        End Select
    End If
    
    '���ڼ�����Ĳ���,�����ѱ������
    If BillingWarn = 1 Or BillingWarn = 4 Then
        If byt��־ = 1 Then
            If rsWarn!������־1 = "-" Then
                str�ѱ���� = "-"
            Else
                str�ѱ���� = str�ѱ���� & "," & rsWarn!������־1
            End If
        ElseIf byt��־ = 2 Then
            If rsWarn!������־2 = "-" Then
                str�ѱ���� = "-"
            Else
                str�ѱ���� = str�ѱ���� & "," & rsWarn!������־2
            End If
            '���ӱ�ע���ж��ѱ����ľ��巽ʽ
            str�ѱ���� = str�ѱ���� & IIf(byt��ʽ = 2, "��", "��")
        ElseIf byt��־ = 3 Then
            If rsWarn!������־3 = "-" Then
                str�ѱ���� = "-"
            Else
                str�ѱ���� = str�ѱ���� & "," & rsWarn!������־3
            End If
        End If
    End If
End Function
Public Function GetPatiDayMoney(lng����ID As Long) As Currency
'���ܣ���ȡָ�����˵��췢���ķ����ܶ�
    Dim rsTmp As New ADODB.Recordset
    Dim strsql As String
    
    On Error GoTo errH
    
    strsql = "Select zl_PatiDayCharge([1]) as ��� From Dual"
    Set rsTmp = zldatabase.OpenSQLRecord(strsql, "mdlCISKernel", lng����ID)
    If Not rsTmp.EOF Then GetPatiDayMoney = zlStr.NVL(rsTmp!���, 0)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
Private Function CheckBill(ByVal lngNOҩ��id As Long, ByVal IntOper As Integer, ByVal IntBillStyle As Integer, ByVal strNo As String, ByVal int��¼���� As Integer, ByVal int�����־ As Integer, Optional ByVal bln��ʾ As Boolean = False) As Integer
    Dim dblCount As Double
    Dim intRow As Integer, intRows As Integer
    Dim rstemp As New ADODB.Recordset
    Dim RecCheck As New ADODB.Recordset
    Dim vsfDetail As VSFlexGrid
    
    '--���ݽ�Ҫִ�еĲ������ж��Ƿ�����--
    'IntOper:1-��ҩ;2-ȡ����ҩ;3-��ҩ;4-��ҩ;5-ȡ����ҩ
    '����:
    '0-�������
    '1-δ��ҩ
    '2-����ҩ
    '3-�ѷ�ҩ
    '4-��ɾ��
    '5-δ��ҩ
    On Error GoTo errHandle
    If lngNOҩ��id = 0 Then lngNOҩ��id = mParams.lngҩ��ID
    
    '��������ȡ����ҩʱ�ļ��
    If IntOper = 5 Then
        gstrSQL = "Select ����� From ҩƷ�շ���¼ Where No=[1] And ����=[2] And �ⷿID+0=[3] And ��¼״̬=1 And ����� IS Not Null And Rownum=1 "
        Set RecCheck = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, strNo, IntBillStyle, lngNOҩ��id)
        If RecCheck.EOF Then
            CheckBill = 4
            MsgBox "δ�ҵ�ָ�����ݣ����ѱ���������Ա����,����������ֹ��", vbInformation, gstrSysName
        End If
        Exit Function
    End If
     
    gstrSQL = " Select A.��ҩ��,A.����� From ҩƷ�շ���¼ A" & _
        " Where A.No=[1] And A.����=[2] " & _
        " " & IIf(IntOper <> 4, " And mod(A.��¼״̬,3)=1", "") & " And Rownum=1 " & _
        " And Nvl(Ltrim(Rtrim(A.ժҪ)),'С��')<>'�ܷ�' And (A.�ⷿID+0=[3] Or A.�ⷿID Is NULL)"
    
    If IntOper = 4 Then
        gstrSQL = gstrSQL & " And ����� IS Not Null"
    Else
        gstrSQL = gstrSQL & " And ����� IS Null"
    End If

    Set RecCheck = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, strNo, IntBillStyle, lngNOҩ��id)
    
    With RecCheck
        If .EOF Then CheckBill = 4: MsgBox "δ�ҵ�ָ������,�����Ѿ�����������Ա����,����������ֹ��", vbInformation, gstrSysName: Exit Function
        If Not IsNull(!�����) Then
            If InStr(1, "123", IntOper) <> 0 Then CheckBill = 3: MsgBox "�ô����ѱ���������Ա��ҩ��" & IIf(IntOper = 1, "��ҩ", IIf(IntOper = 2, "ȡ����ҩ", IIf(IntOper = 3, "��ҩ", "��ҩ"))) & "������ֹ��", vbInformation, gstrSysName: Exit Function
        Else
            If InStr(1, "4", IntOper) <> 0 Then CheckBill = 5: MsgBox "�ô�����δ��ҩ������������ֹ��", vbInformation, gstrSysName: Exit Function
            If Not IsNull(!��ҩ��) Then
                If InStr(1, "1", IntOper) <> 0 Then CheckBill = 2: MsgBox "�ô�������ҩ������������ֹ��", vbInformation, gstrSysName: Exit Function
            Else
                If InStr(1, "2", IntOper) <> 0 Then CheckBill = 1: MsgBox "�ô���δ��ҩ������������ֹ��", vbInformation, gstrSysName: Exit Function
            End If
        End If
    End With
    
    '�������ҩ������Ƿ�����δ����ҽ����ҩ
    If mParams.blnҽ������ = False And bln��ʾ Then
        Set vsfDetail = mfrmDetail.GetDetailList
        intRows = vsfDetail.rows - 2
        For intRow = 1 To intRows
            dblCount = Val(vsfDetail.TextMatrix(intRow, vsfDetail.ColIndex("��ҩ��")))
            If dblCount <> 0 Then
                gstrSQL = "select ���� From ҩƷ�շ���¼ Where ID=[1] "
                Set rstemp = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption & "[����Ƿ�������]", Val(vsfDetail.TextMatrix(intRow, vsfDetail.ColIndex("Id"))))

                If (rstemp!���� Like "1*") Then       '����
                    gstrSQL = "select B.ִ��״̬ from ����ҽ����¼ A,����ҽ������ B,������ü�¼ C where A.���id=B.ҽ��ID and A.id=C.ҽ����� and  C.ID=(Select ����ID From ҩƷ�շ���¼ Where ID=[1])"
                    If int��¼���� = 1 Or (int��¼���� = 2 And (int�����־ = 1 Or int�����־ = 4)) Then
                    Else
                        gstrSQL = Replace(gstrSQL, "������ü�¼", "סԺ���ü�¼")
                    End If
                    Set rstemp = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption & "[���ҽ���ĸ�ҩ;���Ƿ��Ѿ�ִ��]", Val(vsfDetail.TextMatrix(intRow, vsfDetail.ColIndex("Id"))))
                
                    If Not rstemp.EOF Then
                        If rstemp!ִ��״̬ = 0 Then
                            gstrSQL = "Select Nvl(ҽ�����,0) ҽ�����,Nvl(�����־,1) �����־ From ������ü�¼ Where ID=(Select ����ID From ҩƷ�շ���¼ Where ID=[1])"
                            If int��¼���� = 1 Or (int��¼���� = 2 And (int�����־ = 1 Or int�����־ = 4)) Then
                            Else
                                gstrSQL = Replace(gstrSQL, "������ü�¼", "סԺ���ü�¼")
                            End If
                            
                            Set rstemp = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption & "[����Ƿ���ҽ��]", Val(vsfDetail.TextMatrix(intRow, vsfDetail.ColIndex("Id"))))
        
                            If Not rstemp.EOF Then
                                If (rstemp!�����־ = 1 Or rstemp!�����־ = 4) And rstemp!ҽ����� <> 0 Then
                                    gstrSQL = "Select Nvl(��ҳid, 0) As ��ҳid, �Һŵ�, decode(ҽ��״̬,4,1,0) ���� From ����ҽ����¼ Where ������Դ=1  And ID=[1]"
                                    Set rstemp = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption & "[�жϸ�ҽ���Ƿ�����]", CLng(rstemp!ҽ�����))
                                    
                                    If Not rstemp.EOF Then
                                        If rstemp!��ҳid > 0 And IsNull(rstemp!�Һŵ�) Then
                                            '������ҳID����û�йҺŵ��Ĳ���ҽ���Ƿ����ϵ�����
                                        Else
                                            If rstemp!���� = 0 Then
                                                CheckBill = 1
                                                MsgBox "��" & intRow & "�е�ҩƷ��¼��Ӧ��ҽ����δ���ϣ���������ҩ��", vbInformation, gstrSysName
                                                Exit Function
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    Else
                        gstrSQL = "Select Nvl(ҽ�����,0) ҽ�����,Nvl(�����־,1) �����־ From ������ü�¼ Where ID=(Select ����ID From ҩƷ�շ���¼ Where ID=[1])"
                        If int��¼���� = 1 Or (int��¼���� = 2 And (int�����־ = 1 Or int�����־ = 4)) Then
                        Else
                            gstrSQL = Replace(gstrSQL, "������ü�¼", "סԺ���ü�¼")
                        End If
                        
                        Set rstemp = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption & "[����Ƿ���ҽ��]", Val(vsfDetail.TextMatrix(intRow, vsfDetail.ColIndex("Id"))))
    
                        If Not rstemp.EOF Then
                            If (rstemp!�����־ = 1 Or rstemp!�����־ = 4) And rstemp!ҽ����� <> 0 Then
                                gstrSQL = "Select Nvl(��ҳid, 0) As ��ҳid, �Һŵ�, decode(ҽ��״̬,4,1,0) ���� From ����ҽ����¼ Where ������Դ=1  And ID=[1]"
                                Set rstemp = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption & "[�жϸ�ҽ���Ƿ�����]", CLng(rstemp!ҽ�����))
                                
                                If Not rstemp.EOF Then
                                    If rstemp!��ҳid > 0 And IsNull(rstemp!�Һŵ�) Then
                                        '������ҳID����û�йҺŵ��Ĳ���ҽ���Ƿ����ϵ�����
                                    Else
                                        If rstemp!���� = 0 Then
                                            CheckBill = 1
                                            MsgBox "��" & intRow & "�е�ҩƷ��¼��Ӧ��ҽ����δ���ϣ���������ҩ��", vbInformation, gstrSysName
                                            Exit Function
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        Next
    End If
    
    CheckBill = 0
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Private Function CheckBillExist(ByVal Int���� As Integer, ByVal strNo As String) As Boolean
    Dim rstemp As ADODB.Recordset
    On Error GoTo errHandle
    gstrSQL = "Select ID From ҩƷ�շ���¼ " & _
             " Where ����=[1] And NO=[2] And Rownum<2"
    Set rstemp = zldatabase.OpenSQLRecord(gstrSQL, "��鵥���Ƿ����", Int����, strNo)
    CheckBillExist = Not rstemp.EOF
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Private Sub RefreshList(ByVal intType As Integer)
    If mblnStart = False Then Exit Sub

    Call AviShow(Me)
    
    Call GetCondition
    
    Select Case intType
        Case mListType.��ҩȷ��
            RefreshList_DosageOk
        Case mListType.����ҩ
            Call RefreshList_Dosage
        Case mListType.����ҩ
            Call RefreshList_Abolish
        Case mListType.����ҩ
            Call RefreshList_Send
        Case mListType.��ʱδ��
            Call RefreshList_OverTime
        Case mListType.��ҩ
            Call RefreshList_Return
    End Select
    
    Call AviShow(Me, False)
    
    If mblnInput = False Then
        With mfrmList.vsfList
            If .Visible And .Enabled Then .SetFocus
        End With
    Else
        If txtPati.Enabled = True Then txtPati.SetFocus
    End If
End Sub

Private Sub CheckOtherUndeliveredDocuments(ByVal lng����ID As Long)
    '����:���ݲ������ĵ�ǰ������[��ǰҩ����������]��[����ҩ��]�Ƿ����δ��ҩ����
    Dim rstemp As New ADODB.Recordset
    Dim date��ʼ���� As Date
    Dim date�������� As Date
    Dim dteTime As Date
    Dim strMsg As String
    Dim i As Integer
    
    On Error GoTo errHandle
    
    If mParams.int��ѯδ��ҩ�������� = 0 Then Exit Sub
        
    dteTime = Sys.Currentdate
    date��ʼ���� = CDate(Format(DateAdd("d", -mParams.int��ѯδ��ҩ�������� + 1, dteTime), "yyyy-mm-dd") & " 00:00:00")
    date�������� = CDate(DateAdd("s", -1, Format(DateAdd("d", 1, dteTime), "yyyy-mm-dd") & " 00:00:00"))
    
    gstrSQL = "Select Distinct a.No, a.��ҩ����, b.���� As ҩ������" & vbNewLine & _
        "From δ��ҩƷ��¼ A, ���ű� B" & vbNewLine & _
        "Where a.�ⷿid = b.Id And a.����id = [1] And a.�ⷿid = [2] And a.��ҩ���� Is Not Null And a.��ҩ���� Not In (Select b.Column_Value From Table(Cast(f_Str2list([3]) As Zltools.t_Strlist)) B) And" & vbNewLine & _
        "      a.�������� Between [4] And [5]" & vbNewLine & _
        "Union All" & vbNewLine & _
        "Select a.No, a.��ҩ����, b.���� As ҩ������" & vbNewLine & _
        "From δ��ҩƷ��¼ A, ���ű� B" & vbNewLine & _
        "Where a.�ⷿid = b.Id And a.����id = [1] And a.�ⷿid <> [2] And a.�������� Between [4] And [5]"
    
    Set rstemp = zldatabase.OpenSQLRecord(gstrSQL, "����������ҩ�����������ڵ�δ��ҩ����", _
            lng����ID, _
            mParams.lngҩ��ID, _
            mParams.Str����, _
            date��ʼ����, _
            date��������)
    
    '�������򵯳���ʾ��
    If Not rstemp.EOF Then
        If rstemp.RecordCount > 3 Then
            For i = 1 To 3
                strMsg = strMsg & vbCrLf & "���ݺ�:" & rstemp!NO & "   ҩ��:" & rstemp!ҩ������ & "   ��ҩ����:" & IIf(IsNull(rstemp!��ҩ����), "��", rstemp!��ҩ����)
                rstemp.MoveNext
            Next
            strMsg = strMsg & vbCrLf & "��һ�� " & rstemp.RecordCount & " ������"
        Else
            Do While Not rstemp.EOF
                strMsg = strMsg & vbCrLf & "���ݺ�:" & rstemp!NO & "   ҩ��:" & rstemp!ҩ������ & "   ��ҩ����:" & IIf(IsNull(rstemp!��ҩ����), "��", rstemp!��ҩ����)
                rstemp.MoveNext
            Loop
        End If
        
        MsgBox "�ò��˻�������δ��ҩ��������" & strMsg
    End If
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub RefreshList_DosageOk()
    'ˢ����ҩ�б�
    Dim blnҽ���� As Boolean
    Dim rsData As ADODB.Recordset
    Dim strSub1 As String
    Dim strSub2 As String
    Dim strסԺ As String
    Dim str���� As String
    Dim lng����ID As Long
    Dim strSqlTmp As String
    Dim str����ʱ�� As String
    
    On Error GoTo errHandle
    If mSQLCondition.str���֤ <> "" And Not mobjSquareCard Is Nothing Then
        If mobjSquareCard.zlGetPatiID("���֤", UCase(mSQLCondition.str���֤), False, lng����ID) = False Then lng����ID = 0
    End If
    
    gstrSQL = "Select '' As ��ɫ, �������� ,'' As ѡ�� ,'0' As ��־,����,����,���շ�,��ҩ��,NO,����,ǩ��ʱ��," & _
            " to_Char(Sum(Round(���۽��," & mintMoneyDigit & ")),'" & mstrOracleMoneyForamt & "') AS ���,����,�ɲ���," & _
            " ˵��,���￨��,�����,���֤��,IC����,����ID,ҽ����,סԺ��,�Ŷ�״̬,��ҩ����," & _
            " Sum(Round(ʵ�ս��," & mintMoneyDigit & ")) ʵ�ս��,�����־,��¼����,Zl_Get�շ����(����,NO,[1]) As �շ����,�������� " & _
            " From ("
            
    strSqlTmp = " Select A.���ȼ�,A.����,A.����,A.���շ�,A.��ҩ��,A.NO,A.����,C.���۽��,A.����,A.�ɲ���,A.˵��,A.���￨��,A.�����,A.���֤��," & _
            " A.IC����,A.����ID,A.ҽ����,A.סԺ��,A.�Ŷ�״̬,d.ʵ�ս��, Nvl(A.��������,Nvl(C.ע��֤��,0)) As ��������,D.�����־,D.��¼����,D.�շ����,A.��ҩ����,c.ǩ��ʱ��,a.�������� " & _
            " From ("
    
    str���� = "Select distinct B.���￨��,B.�����,B.���֤��,B.IC����,B.ҽ����,B.סԺ��,A.���ȼ�,A.��ҩ����,A.��������,Decode(Nvl(A.���շ�,0),1,'','(δ)')||Decode(A.����,8,'�շ�',9,'����') ����," & _
            " A.����,A.���շ�,'' ��ҩ��,A.No,A.����,To_Char(A.��������,'yyyy-MM-dd hh24:mi:ss') ����,1 �ɲ���,' ' ˵��,B.����ID, A.��������,A.�Ŷ�״̬,a.�Է�����id, b.�������� " & _
            " From δ��ҩƷ��¼ A,������Ϣ B,������ü�¼ C " & _
            " Where A.��ҩ�� is null "
    
    '�Ƿ���ʾ��ȷ�ϵ���
    If mcondition.bln��ʾ��ȷ�ϵ��� = False Then
        str���� = str���� & " And (A.�Ŷ�״̬=0 or A.�Ŷ�״̬ is null) "
    Else
        str���� = str���� & " And (A.�Ŷ�״̬=0 or A.�Ŷ�״̬=1 or A.�Ŷ�״̬ is null) "
    End If
    
    '��Ҫ����
    str���� = str���� & " And (A.�ⷿID=[1] Or A.�ⷿID Is NULL) And A.�������� Between [2] And [3] "
    
    If mSQLCondition.str��ʼNO <> "" Or mSQLCondition.str����NO <> "" Then
        If mSQLCondition.str��ʼNO <> "" And mSQLCondition.str����NO <> "" Then
            str���� = str���� & " And A.NO Between [4] And [5] "
        Else
            If mSQLCondition.str��ʼNO <> "" Then
                str���� = str���� & " And A.NO = [4] "
            Else
                str���� = str���� & " And A.NO = [5] "
            End If
        End If
    End If
    
    If mSQLCondition.str���� <> "" Then str���� = str���� & " And Upper(A.����) Like [6] "
    
    If mSQLCondition.str���￨ <> "" Then str���� = str���� & " And Upper(B.���￨��) = [7] "
    
    If mSQLCondition.str��ʶ�� <> "" Then str���� = str���� & " And Upper(DECODE(A.����,8,B.�����,B.סԺ��)) Like [8] "
    
    If mSQLCondition.lng����ID > 0 Then str���� = str���� & " And A.�Է�����ID+0=[9] "
    
    If mSQLCondition.str��ǰNO <> "" Then str���� = str���� & " And A.NO=[13] "
    
    If mSQLCondition.str����� <> "" Then str���� = str���� & " And B.�����=[14] "
    
'    If mSQLCondition.str���֤ <> "" Then str���� = str���� & " And B.���֤��=[15] "

    If mSQLCondition.str���֤ <> "" Then str���� = str���� & " And B.����ID=[15] "

    If mSQLCondition.lng����ID <> 0 Then str���� = str���� & " And B.����ID=[16] "
    
    If mSQLCondition.strҽ���� <> "" Then str���� = str���� & " And B.ҽ����=[17] "
    
    If mSQLCondition.lngסԺ�� <> 0 Then str���� = str���� & " And B.סԺ��=[18] "
    
            
    blnҽ���� = (mSQLCondition.strҽ���� <> "")
    str���� = str���� & " And A.����ID=B.����ID" & IIf(blnҽ���� = True, "", "(+)") & ""
    
    str���� = str���� & IIf(mParams.Str���� = "", "", " And (A.��ҩ���� In (Select * From Table(Cast(f_Str2list([19]) As Zltools.t_Strlist))) Or A.��ҩ���� Is Null) ")
    
    '������ʾ���Զ���ӡ������:ע��"δ��ҩƷ��¼"�ı���ΪA
    Select Case mParams.intShowBill�շ�
        Case 0  '����ʾ����
            strSub1 = "1=2"
        Case 1  '��ʾδ�շ�
            strSub1 = "A.����<>9 And Nvl(A.���շ�,0)=0 And A.����=8"
        Case 2  '��ʾ���շ�
            strSub1 = "A.����<>9 And A.���շ�=1 And A.����=8"
        Case 3  '��ʾ���д���
            strSub1 = "A.����<>9 And A.����=8"
    End Select
    Select Case mParams.intShowBill����
        Case 0  '����ʾ����
            strSub2 = "1=2"
        Case 1  '��ʾδ���
            strSub2 = "A.����<>8 And Nvl(A.���շ�,0)=0 And A.����=9"
        Case 2  '��ʾ�����
            strSub2 = "A.����<>8 And A.���շ�=1 And A.����=9"
        Case 3  '��ʾ���д���
            strSub2 = "A.����<>8 And A.����=9"
    End Select
    
    str���� = str���� & " And A.���� IN(8,9) And (" & strSub1 & " Or " & strSub2 & ")"
    
    str���� = str���� & " And Mod(C.��¼����, 10) = Decode(A.����, 8, 1, 2) And A.No = C.No And A.�ⷿid = C.ִ�в���id "
            
    If mParams.bln����ʱ����� = False Then
'        str���� = Replace(str����, ",������ü�¼ C", "")
    Else
        str����ʱ�� = Replace(str����, "And A.�������� Between [2] And [3]", "")
        str���� = str���� & " And C.ҽ����� Is Null "
        
        str����ʱ�� = str����ʱ�� & " And C.ҽ����� Is Not Null And C.����ʱ�� Between [2] And [3] "
        str���� = str���� & " Union All " & str����ʱ��
    End If
    
    str���� = strSqlTmp & str���� & ") A,ҩƷ�շ���¼ C, ������ü�¼ D, ���ű� B " & IIf(mParams.strSourceDep = "", "", ",Table(Cast(f_Num2List([20]) As zlTools.t_NumList)) E ") & _
              " Where C.����id = D.ID And nvl(c.��ҩ��ʽ,-999)<>-1 and A.����=C.���� And A.NO=C.NO And C.����� Is NULL " & _
              " And Nvl(D.����״̬,0)<>1 And (C.�ⷿid=[1] Or C.�ⷿid Is null)  And a.�Է�����id = b.Id "
    
    If mstrDeptNode <> "" Then
        str���� = str���� & " And (b.վ�� = [21] Or b.վ�� Is Null) "
    End If
    
    '�ų��Ѿ�����ֹͣ��ҩ��No
    str���� = str���� & " and Not Exists(select 1 from ҩƷ�շ���¼ F where F.����=C.���� and F.�ⷿid=C.�ⷿid and F.no=C.no and ��ҩ��ʽ=-1) "
    
    '�ų�������Һ�������Ĺ����в����ĵ���
    str���� = str���� & " And Not Exists (Select 1 From ��Һ��ҩ���� Y Where Y.�շ�id = C.ID) "
    
    '��Ժ��ҩ:'0-����,1-������Ժ��ҩ,2-������Ժ��ҩ
    If mcondition.int��Ժ��ҩ = 0 Then
    ElseIf mcondition.int��Ժ��ҩ = 1 Then
        str���� = str���� & " And Not Ltrim(To_Char(Nvl(C.����,0),'00')) Like '_3'"
    ElseIf mcondition.int��Ժ��ҩ = 2 Then
        str���� = str���� & " And Ltrim(To_Char(Nvl(C.����,0),'00')) Like '_3'"
    End If
    
    str���� = str���� & IIf(mParams.strSourceDep = "", "", " And C.�Է�����id+0=E.Column_Value ")
    
    If imgFilter.BorderStyle = cstFilter Then
        If Val(mParams.int����ģʽ) <= 7 Then
            If Trim(txtPati.Text) = "" Then
                str���� = str���� & " And 1 = 2 "
            End If
        Else
            If mSQLCondition.lng����ID = 0 Then
                str���� = str���� & " And 1 = 2 "
            End If
        End If
    End If
    
    If mParams.intType = 1 Then
        str���� = str���� & " And (D.�����־=1 or D.�����־=4)"
    ElseIf mParams.intType = 2 Then
        str���� = str���� & " And (D.�����־<>1 and D.�����־<>4)"
    End If
    
    If mcondition.int������� = 1 Then
        '���ﻮ�ۼ��������
        gstrSQL = gstrSQL & str����
    Else
        If mcondition.int������� = 3 Then
            '���ＰסԺ���е���
            strסԺ = Replace(str����, "������ü�¼", "סԺ���ü�¼")
            strסԺ = Replace(strסԺ, "And Nvl(D.����״̬,0)<>1", "")
        Else
            'סԺ����
            strסԺ = Replace(str����, "������ü�¼", "סԺ���ü�¼")
            strסԺ = Replace(strסԺ, "And Nvl(D.����״̬,0)<>1", "")
            str���� = ""
        End If
    
        If mPrives.bln���������� Then
            If img����.BorderStyle = 0 Then
                '����ʾ��������
                strסԺ = strסԺ & " And (D.�����־ <> 2 Or (D.�����־ = 2 And D.���˲���id <> D.��������id)) "
            End If
            If img����.BorderStyle = 1 And cbo����.ListIndex <> -1 Then
                'Ҫ��ʾ�������������Ҳ��˲������ڵ�ǰѡ��Ĳ���
                strסԺ = strסԺ & " And D.���˲���id = " & cbo����.ItemData(cbo����.ListIndex)
                str���� = ""
            End If
        End If
        
        If str���� = "" Then
            gstrSQL = gstrSQL & strסԺ
        Else
            gstrSQL = gstrSQL & str���� & " Union All " & strסԺ
        End If
    End If
    
    '''''Group By
    gstrSQL = gstrSQL & ") A GROUP BY A.���ȼ�,A.����,A.����,A.���շ�,A.��ҩ��,A.No,A.����,A.����,A.�ɲ���,A.˵��," & _
        " A.���￨��,A.�����,A.���֤��,A.IC����,A.����ID,A.ҽ����,A.סԺ��,A.�Ŷ�״̬,A.��������,A.�����־,A.��¼����, a.��ҩ����,A.ǩ��ʱ��,a.�������� "
    
    '''''Order By
    gstrSQL = gstrSQL & " Order By A.ǩ��ʱ��,A.����,A.����,A.No"
    
    Set rsData = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, _
            mSQLCondition.lngҩ��ID, _
            mSQLCondition.date��ʼ����, _
            mSQLCondition.date��������, _
            mSQLCondition.str��ʼNO, _
            mSQLCondition.str����NO, _
            UCase(mSQLCondition.str����), _
            mSQLCondition.str���￨, _
            mSQLCondition.str��ʶ��, _
            mSQLCondition.lng����ID, _
            mSQLCondition.str������, _
            mSQLCondition.str�����, _
            mSQLCondition.lngҩƷid, _
            mSQLCondition.str��ǰNO, _
            mSQLCondition.str�����, _
            lng����ID, _
            mSQLCondition.lng����ID, _
            mSQLCondition.strҽ����, _
            mSQLCondition.lngסԺ��, _
            mParams.Str����, _
            mParams.strSourceDep, _
            mstrDeptNode)
    
    stbThis.Panels(2) = ""
    If Not rsData.EOF Then
        stbThis.Panels(2) = "����" & rsData.RecordCount & "�Ŵ�����" & GetSumMoney(rsData)
    End If
    
    Set mrsList = rsData
    
    If Not mfrmList Is Nothing Then mfrmList.RefreshList mListType.��ҩȷ��, mrsList
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub




Public Sub RefreshList_Dosage(Optional ByVal strNo As String, Optional ByVal blnNoRefreshDetail As Boolean)
    'ˢ����ҩ�б�
    Dim blnҽ���� As Boolean
    Dim rsData As ADODB.Recordset
    Dim strSub1 As String
    Dim strSub2 As String
    Dim strסԺ As String
    Dim str���� As String
    Dim cbrControl As CommandBarControl
    Dim cbrMenu As CommandBarControl
    Dim lng����ID As Long
    Dim strSqlTmp As String
    Dim str����ʱ�� As String
    
    On Error GoTo errHandle
    If mSQLCondition.str���֤ <> "" And Not mobjSquareCard Is Nothing Then
        If mobjSquareCard.zlGetPatiID("���֤", UCase(mSQLCondition.str���֤), False, lng����ID) = False Then lng����ID = 0
    End If
    
    gstrSQL = "Select '' As ��ɫ, �������� ,'' As ѡ�� ,'0' As ��־,����,����,���շ�,��ҩ��,NO,����,ǩ��ʱ��," & _
            " to_Char(Sum(Round(���۽��," & mintMoneyDigit & ")),'" & mstrOracleMoneyForamt & "') AS ���,����,�ɲ���," & _
            " ˵��,���￨��,�����,���֤��,IC����,����ID,ҽ����,סԺ��,��ҩ����," & _
            " Sum(Round(ʵ�ս��," & mintMoneyDigit & ")) ʵ�ս��,�����־,��¼����,Zl_Get�շ����(����,NO,[1]) As �շ����,��������, ��ӡ״̬ " & IIf(mParams.bln������, ",�����,���id", "") & _
            " From ("
            
    strSqlTmp = " Select A.���ȼ�,A.����,A.����,A.���շ�,A.��ҩ��,A.NO,A.����,C.���۽��,A.����,A.�ɲ���,A.˵��,A.���￨��,A.�����,A.���֤��," & _
            " A.IC����,A.����ID,A.ҽ����,A.סԺ��,d.ʵ�ս��*(Nvl(c.����,1)*c.ʵ������/(Nvl(d.����,1)*d.����)) ʵ�ս��, Nvl(A.��������,Nvl(C.ע��֤��,0)) As ��������,D.�����־,D.��¼����,D.�շ����,A.��ҩ����,c.ǩ��ʱ��,a.��������, a.��ӡ״̬ " & IIf(mParams.bln������, ",a.�����,a.���id", "") & _
            " From ("
            
    str���� = "Select distinct B.���￨��,B.�����,B.���֤��,B.IC����,B.ҽ����,B.סԺ��,A.���ȼ�,A.��ҩ����,A.��������,Decode(Nvl(A.���շ�,0),1,'','(δ)')||Decode(A.����,8,'�շ�',9,'����') ����," & _
            " A.����,A.���շ�,'' ��ҩ��,A.No,A.����,To_Char(A.��������,'yyyy-MM-dd hh24:mi:ss') ����,1 �ɲ���,' ' ˵��,B.����ID, A.��������,a.�Է�����id, b.��������, Decode(a.��ӡ״̬,1,1,3,1,0) ��ӡ״̬ " & IIf(mParams.bln������, ",Q.�����,Q.id  ���id", "") & _
            " From δ��ҩƷ��¼ A,������Ϣ B,������ü�¼ C " & IIf(mParams.bln������, ",��������¼ Q,���������ϸ K ", "") & _
            " Where A.��ҩ�� Is Null "
            
    str���� = str���� & IIf(mParams.bln������, " and c.ҽ�����=k.ҽ��id(+) and Q.id(+)=K.��id and K.����ύ(+)=1 ", "")
    
    '�Ƿ�����ҩȷ�ϻ���
    If mParams.blnMustDosageOkProcess = True Then
        str���� = str���� & " and A.�Ŷ�״̬=1"
    End If
    
    '��Ҫ����
    str���� = str���� & " And (A.�ⷿID=[1] Or A.�ⷿID Is NULL) And A.�������� Between [2] And [3] "
    
    If mSQLCondition.str��ʼNO <> "" Or mSQLCondition.str����NO <> "" Then
        If mSQLCondition.str��ʼNO <> "" And mSQLCondition.str����NO <> "" Then
            str���� = str���� & " And A.NO Between [4] And [5] "
        Else
            If mSQLCondition.str��ʼNO <> "" Then
                str���� = str���� & " And A.NO = [4] "
            Else
                str���� = str���� & " And A.NO = [5] "
            End If
        End If
    End If
    
    If mSQLCondition.str���� <> "" Then str���� = str���� & " And Upper(A.����) Like [6] "
    
    If mSQLCondition.str���￨ <> "" Then str���� = str���� & " And Upper(B.���￨��) = [7] "
    
    If mSQLCondition.str��ʶ�� <> "" Then str���� = str���� & " And Upper(DECODE(A.����,8,B.�����,B.סԺ��)) Like [8] "
    
    If mSQLCondition.lng����ID > 0 Then str���� = str���� & " And A.�Է�����ID+0=[9] "
    
    If mSQLCondition.str��ǰNO <> "" Then str���� = str���� & " And A.NO=[13] "
    
    If mSQLCondition.str����� <> "" Then str���� = str���� & " And B.�����=[14] "
    
'    If mSQLCondition.str���֤ <> "" Then str���� = str���� & " And B.���֤��=[15] "
    
    If mSQLCondition.str���֤ <> "" Then str���� = str���� & " And B.����ID=[15] "
    
    If mSQLCondition.lng����ID <> 0 Then str���� = str���� & " And B.����ID=[16] "
    
    If mSQLCondition.strҽ���� <> "" Then str���� = str���� & " And B.ҽ����=[17] "
    
    If mSQLCondition.lngסԺ�� <> 0 Then str���� = str���� & " And B.סԺ��=[18] "
    
            
    blnҽ���� = (mSQLCondition.strҽ���� <> "")
    str���� = str���� & " And A.����ID=B.����ID" & IIf(blnҽ���� = True, "", "(+)") & ""
    
    str���� = str���� & IIf(mParams.Str���� = "", "", " And (A.��ҩ���� In (Select * From Table(Cast(f_Str2list([19]) As Zltools.t_Strlist))) Or A.��ҩ���� Is Null) ")
    
    '������ʾ���Զ���ӡ������:ע��"δ��ҩƷ��¼"�ı���ΪA
    Select Case mParams.intShowBill�շ�
        Case 0  '����ʾ����
            strSub1 = "1=2"
        Case 1  '��ʾδ�շ�
            strSub1 = "A.����<>9 And Nvl(A.���շ�,0)=0 And A.����=8"
        Case 2  '��ʾ���շ�
            strSub1 = "A.����<>9 And A.���շ�=1 And A.����=8"
        Case 3  '��ʾ���д���
            strSub1 = "A.����<>9 And A.����=8"
    End Select
    Select Case mParams.intShowBill����
        Case 0  '����ʾ����
            strSub2 = "1=2"
        Case 1  '��ʾδ���
            strSub2 = "A.����<>8 And Nvl(A.���շ�,0)=0 And A.����=9"
        Case 2  '��ʾ�����
            strSub2 = "A.����<>8 And A.���շ�=1 And A.����=9"
        Case 3  '��ʾ���д���
            strSub2 = "A.����<>8 And A.����=9"
    End Select
    
    str���� = str���� & " And A.���� IN(8,9) And (" & strSub1 & " Or " & strSub2 & ")"
    
    str���� = str���� & " And Mod(C.��¼����, 10) = Decode(A.����, 8, 1, 2) And A.No = C.No And A.�ⷿid = C.ִ�в���id "
    
    '��ҩ��ӡ״̬��0-��ʾ������ҩ��,1-ֻ��ʾδ��ӡ�Ĵ���ҩ����,2-ֻ��ʾ�Ѵ�ӡ�Ĵ���ҩ����
    If mParams.intShowBill��ҩ = 1 Then
        str���� = str���� & " And Nvl(A.��ӡ״̬,0) Not In(1,3)"
    ElseIf mParams.intShowBill��ҩ = 2 Then
        str���� = str���� & " And Nvl(A.��ӡ״̬,0) In(1,3)"
    End If
    
    If mParams.bln����ʱ����� = False Then
'        str���� = Replace(str����, ",������ü�¼ C", "")
    Else
        
        
        str����ʱ�� = Replace(str����, "And A.�������� Between [2] And [3]", "")
        str���� = str���� & " And C.ҽ����� Is Null "
        
        str����ʱ�� = str����ʱ�� & " And C.ҽ����� Is Not Null And C.����ʱ�� Between [2] And [3] "
        str���� = str���� & " Union All " & str����ʱ��
    End If
    
    str���� = strSqlTmp & str���� & ") A,ҩƷ�շ���¼ C, ������ü�¼ D, ���ű� B " & IIf(mParams.strSourceDep = "", "", ",Table(Cast(f_Num2List([20]) As zlTools.t_NumList)) E ") & _
              " Where C.����id = D.ID And nvl(c.��ҩ��ʽ,-999)<>-1 and A.����=C.���� And A.NO=C.NO And C.����� Is NULL " & _
              " And Nvl(D.����״̬,0)<>1 And (C.�ⷿid=[1] Or C.�ⷿid Is null)  And a.�Է�����id = b.Id "
    
    If mstrDeptNode <> "" Then
        str���� = str���� & " And (b.վ�� = [21] Or b.վ�� Is Null) "
    End If
    
    '�ų��Ѿ�����ֹͣ��ҩ��No
    str���� = str���� & " and Not Exists(select 1 from ҩƷ�շ���¼ F where F.����=C.���� and F.�ⷿid=C.�ⷿid and F.no=C.no and ��ҩ��ʽ=-1) "
    
    '�ų�������Һ�������Ĺ����в����ĵ���
    str���� = str���� & " And Not Exists (Select 1 From ��Һ��ҩ���� Y Where Y.�շ�id = C.ID) "
    
    '�Ƿ���ʾ��ҩ��������
    If mcondition.bln��ʾ��ҩ�������� = False Then
        str���� = str���� & " And C.��¼״̬=1 "
    Else
        str���� = str���� & " And MOD(C.��¼״̬,3)=1 "
    End If
    
    '��Ժ��ҩ:'0-����,1-������Ժ��ҩ,2-������Ժ��ҩ
    If mcondition.int��Ժ��ҩ = 0 Then
    ElseIf mcondition.int��Ժ��ҩ = 1 Then
        str���� = str���� & " And Not Ltrim(To_Char(Nvl(C.����,0),'00')) Like '_3'"
    ElseIf mcondition.int��Ժ��ҩ = 2 Then
        str���� = str���� & " And Ltrim(To_Char(Nvl(C.����,0),'00')) Like '_3'"
    End If
    
    str���� = str���� & IIf(mParams.strSourceDep = "", "", " And C.�Է�����id+0=E.Column_Value ")
    
    If imgFilter.BorderStyle = cstFilter Then
        If Val(mParams.int����ģʽ) <= 7 Then
            If Trim(txtPati.Text) = "" Then
                str���� = str���� & " And 1 = 2 "
            End If
        Else
            If mSQLCondition.lng����ID = 0 Then
                str���� = str���� & " And 1 = 2 "
            End If
        End If
    End If
    
    If mParams.intType = 1 Then
        str���� = str���� & " And (D.�����־=1 or D.�����־=4)"
    ElseIf mParams.intType = 2 Then
        str���� = str���� & " And (D.�����־<>1 and D.�����־<>4)"
    End If
    
    If mcondition.int������� = 1 Then
        '���ﻮ�ۼ��������
        gstrSQL = gstrSQL & str����
    Else
        If mcondition.int������� = 3 Then
            '���ＰסԺ���е���
            strסԺ = Replace(str����, "������ü�¼", "סԺ���ü�¼")
            strסԺ = Replace(strסԺ, "And Nvl(D.����״̬,0)<>1", "")
        Else
            'סԺ����
            strסԺ = Replace(str����, "������ü�¼", "סԺ���ü�¼")
            strסԺ = Replace(strסԺ, "And Nvl(D.����״̬,0)<>1", "")
            str���� = ""
        End If
    
        If mPrives.bln���������� Then
            If img����.BorderStyle = 0 Then
                '����ʾ��������
                strסԺ = strסԺ & " And (D.�����־ <> 2 Or (D.�����־ = 2 And D.���˲���id <> D.��������id)) "
            End If
            If img����.BorderStyle = 1 And cbo����.ListIndex <> -1 Then
                'Ҫ��ʾ�������������Ҳ��˲������ڵ�ǰѡ��Ĳ���
                strסԺ = strסԺ & " And D.���˲���id = " & cbo����.ItemData(cbo����.ListIndex)
                str���� = ""
            End If
        End If
        
        If str���� = "" Then
            gstrSQL = gstrSQL & strסԺ
        Else
            gstrSQL = gstrSQL & str���� & " Union All " & strסԺ
        End If
    End If
    
    '''''Group By
    gstrSQL = gstrSQL & ") A GROUP BY A.���ȼ�,A.����,A.����,A.���շ�,A.��ҩ��,A.No,A.����,A.����,A.�ɲ���,A.˵��," & _
        " A.���￨��,A.�����,A.���֤��,A.IC����,A.����ID,A.ҽ����,A.סԺ��,A.��������,A.�����־,A.��¼����, a.��ҩ����,A.ǩ��ʱ��,a.��������, a.��ӡ״̬ " & IIf(mParams.bln������, ",a.�����,a.���id", "")
    
    '''''Order By
    gstrSQL = gstrSQL & " Order By A.ǩ��ʱ��,A.����,A.����,A.No"
    
    Set rsData = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, _
            mSQLCondition.lngҩ��ID, _
            mSQLCondition.date��ʼ����, _
            mSQLCondition.date��������, _
            mSQLCondition.str��ʼNO, _
            mSQLCondition.str����NO, _
            UCase(mSQLCondition.str����), _
            mSQLCondition.str���￨, _
            mSQLCondition.str��ʶ��, _
            mSQLCondition.lng����ID, _
            mSQLCondition.str������, _
            mSQLCondition.str�����, _
            mSQLCondition.lngҩƷid, _
            mSQLCondition.str��ǰNO, _
            mSQLCondition.str�����, _
            lng����ID, _
            mSQLCondition.lng����ID, _
            mSQLCondition.strҽ����, _
            mSQLCondition.lngסԺ��, _
            mstr����, _
            mParams.strSourceDep, _
            mstrDeptNode)
    
    stbThis.Panels(2) = ""
    
    Set cbrMenu = Me.cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_Edit_Recipe_Cancle, , True)
    Set cbrControl = Me.cbsMain(2).Controls.Find(xtpControlButton, mconMenu_Edit_Recipe_Cancle, , True)
        
    If Not rsData.EOF Then
        cbrMenu.Enabled = True
        cbrControl.Enabled = True
        stbThis.Panels(2) = "����" & rsData.RecordCount & "�Ŵ�����" & GetSumMoney(rsData)
    Else
        cbrMenu.Enabled = False
        cbrControl.Enabled = False
    End If
    
    Set mrsList = rsData
    
    If Not mfrmList Is Nothing Then mfrmList.RefreshList mListType.����ҩ, mrsList, strNo, blnNoRefreshDetail
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
Private Function GetSumMoney(ByVal rsRecipt As ADODB.Recordset) As String
    Dim rstemp As ADODB.Recordset
    Dim dblӦ�ս�� As Double
    Dim dlbʵ�ս�� As Double
    Set rstemp = rsRecipt.Clone
    
    With rstemp
        .MoveFirst
        Do While Not .EOF
            dblӦ�ս�� = dblӦ�ս�� + Val(.Fields("���").Value)
            dlbʵ�ս�� = dlbʵ�ս�� + Val(.Fields("ʵ�ս��").Value)
            .MoveNext
        Loop
    End With
    
    If mParams.int�����ʾ = 1 Then
        GetSumMoney = "ʵ�ս�" & FormatEx(dlbʵ�ս��, mintMoneyDigit) & "Ԫ"
    ElseIf mParams.int�����ʾ = 2 Then
        GetSumMoney = "Ӧ�ս�" & FormatEx(dblӦ�ս��, mintMoneyDigit) & "Ԫ" & "  ʵ�ս�" & FormatEx(dlbʵ�ս��, mintMoneyDigit) & "Ԫ"
    Else
        GetSumMoney = "Ӧ�ս�" & FormatEx(dblӦ�ս��, mintMoneyDigit) & "Ԫ"
    End If
End Function
Public Sub RefreshList_Send(Optional ByVal strNo As String, Optional ByVal blnNoRefreshDetail As Boolean)
    'ˢ�´���ҩ�б�
    Dim blnҽ���� As Boolean
    Dim rsData As ADODB.Recordset
    Dim strSub1 As String
    Dim strSub2 As String
    Dim strסԺ As String
    Dim str���� As String
    Dim strInput As String
    Dim cbrControl As CommandBarControl
    Dim cbrMenu As CommandBarControl
    Dim lng����ID As Long
    Dim strSqlTmp As String
    Dim str����ʱ�� As String
    
    On Error GoTo errHandle
    If mSQLCondition.str���֤ <> "" And Not mobjSquareCard Is Nothing Then
        If mobjSquareCard.zlGetPatiID("���֤", UCase(mSQLCondition.str���֤), False, lng����ID) = False Then lng����ID = 0
    End If
    
    gstrSQL = "Select '' As ��ɫ, ��������,'' As ѡ�� ,'0' As ��־,����,����,���շ�,��ҩ��,NO,����,����ʱ��,ǩ��ʱ��," & _
            " to_Char(Sum(Round(���۽��," & mintMoneyDigit & ")),'" & mstrOracleMoneyForamt & "') AS ���,����,�ɲ���," & _
            " ˵��,���￨��,�����,���֤��,IC����,����ID,ҽ����,סԺ��,��ҩ����," & _
            " Sum(Round(ʵ�ս��," & mintMoneyDigit & ")) ʵ�ս��,�����־,��¼����,Zl_Get�շ����(����,NO,[1]) As �շ����,��������" & IIf(mParams.bln������, ",�����,���id", "") & _
            " From ("
            
    strSqlTmp = " Select A.����ʱ��,A.���ȼ�,A.����,A.����,A.���շ�,A.��ҩ��,A.NO,A.����,C.���۽��,A.����,A.�ɲ���,A.˵��,A.���￨��,A.�����,A.���֤��," & _
            " A.IC����,A.����ID,A.ҽ����,A.סԺ��,d.ʵ�ս��*(Nvl(c.����,1)*c.ʵ������/(Nvl(d.����,1)*d.����)) ʵ�ս��, Nvl(A.��������,Nvl(C.ע��֤��,0)) As ��������,D.�����־,D.��¼����,D.�շ����,A.��ҩ����,c.ǩ��ʱ��,a.��������" & IIf(mParams.bln������, ",a.�����,a.���id", "") & _
            " From ("
    str���� = "Select distinct A.����ʱ��,B.���￨��,B.�����,B.���֤��,B.IC����,B.ҽ����,B.סԺ��,A.���ȼ�,A.��ҩ����,A.��������,Decode(Nvl(A.���շ�,0),1,'','(δ)')||Decode(A.����,8,'�շ�',9,'����') ����, " & _
            " A.����,A.���շ�,'' ��ҩ��,A.No,A.����,To_Char(A.��������,'yyyy-MM-dd hh24:mi:ss') ����,1 �ɲ���,' ' ˵��,B.����ID, A.��������,a.�Է�����id, b.��������" & IIf(mParams.bln������, ",Q.�����,Q.id  ���id", "") & _
            " From δ��ҩƷ��¼ A,������Ϣ B,������ü�¼ C " & IIf(mParams.bln������, ",��������¼ Q,���������ϸ K ", "") & _
            " Where 1=1 "
    
    str���� = str���� & IIf(mParams.bln������, " and c.ҽ�����=k.ҽ��id(+) and Q.id(+)=K.��id and K.����ύ(+)=1 ", "")
    
    '�Ƿ�����ҩȷ�ϻ���
    If mParams.blnMustDosageOkProcess = True And mParams.blnMustDosageProcess = True Then
        str���� = str���� & " and A.�Ŷ�״̬ in (2,3,4)"
    ElseIf mParams.blnMustDosageOkProcess = True And mParams.blnMustDosageProcess = False Then
        str���� = str���� & " and A.�Ŷ�״̬ in (1,2,3,4)"
    End If
    
    '��Ҫ����
    If mParams.blnMustDosageProcess = True Then str���� = str���� & " And A.��ҩ�� Is Not Null "
    
    str���� = str���� & " And (A.�ⷿID=[1] Or A.�ⷿID Is NULL) And A.�������� Between [2] And [3] "

    If mSQLCondition.str��ʼNO <> "" Or mSQLCondition.str����NO <> "" Then
        If mSQLCondition.str��ʼNO <> "" And mSQLCondition.str����NO <> "" Then
            str���� = str���� & " And A.NO Between [4] And [5] "
        Else
            If mSQLCondition.str��ʼNO <> "" Then
                str���� = str���� & " And A.NO = [4] "
            Else
                str���� = str���� & " And A.NO = [5] "
            End If
        End If
    End If
    
    If mSQLCondition.str���� <> "" Then str���� = str���� & " And Upper(A.����) Like Upper([6]) "
    
    If mSQLCondition.str���￨ <> "" Then str���� = str���� & " And Upper(B.���￨��) = [7] "
    
    If mSQLCondition.str��ʶ�� <> "" Then str���� = str���� & " And Upper(DECODE(A.����,8,B.�����,B.סԺ��)) Like [8] "
    
    If mSQLCondition.lng����ID > 0 Then str���� = str���� & " And A.�Է�����ID+0=[9] "
    
    If mSQLCondition.str��ǰNO <> "" Then str���� = str���� & " And A.NO=[13] "
    
    If mSQLCondition.str����� <> "" Then str���� = str���� & " And B.�����=[14] "
    
'    If mSQLCondition.str���֤ <> "" Then str���� = str���� & " And B.���֤��=[15] "

    If mSQLCondition.str���֤ <> "" Then str���� = str���� & " And B.����ID=[15] "
    
    If mSQLCondition.lng����ID <> 0 Then str���� = str���� & " And B.����ID=[16] "
    
    If mSQLCondition.strҽ���� <> "" Then str���� = str���� & " And B.ҽ����=[17] "
    
    If mSQLCondition.lngסԺ�� <> 0 Then str���� = str���� & " And B.סԺ��=[18] "
    
    blnҽ���� = (mSQLCondition.strҽ���� <> "")
    
    str���� = str���� & " And A.����ID=B.����ID" & IIf(blnҽ���� = True, "", "(+)") & ""
    
    str���� = str���� & IIf(mParams.Str���� = "", "", " And (A.��ҩ���� In (Select * From Table(Cast(f_Str2list([19]) As Zltools.t_Strlist))) Or A.��ҩ���� Is Null) ")
    
    '�շѵ���
    Select Case mParams.intShowBill�շ�
        Case 0  '����ʾ����
            strSub1 = "1=2"
        Case 1  '��ʾδ�շ�
            strSub1 = "(Nvl(A.���շ�,0)=0 And A.����=8)"
        Case 2  '��ʾ���շ�
            strSub1 = "(A.���շ�=1 And A.����=8)"
        Case 3  '��ʾ���д���
            strSub1 = "A.����=8"
    End Select
    '���ʵ���
    Select Case mParams.intShowBill����
        Case 0  '����ʾ����
            strSub2 = "1=2"
        Case 1  '��ʾδ���
            strSub2 = "(Nvl(A.���շ�,0)=0 And A.����=9)"
        Case 2  '��ʾ�����
            strSub2 = "(A.���շ�=1 And A.����=9)"
        Case 3  '��ʾ���д���
            strSub2 = "A.����=9"
    End Select
    
    str���� = str���� & " And (" & strSub1 & " Or " & strSub2 & ")"
    
    str���� = str���� & " And Mod(C.��¼����, 10) = Decode(A.����, 8, 1, 2) And A.No = C.No And A.�ⷿid = C.ִ�в���id "
    
    If mParams.bln����ʱ����� = False Then
'        str���� = Replace(str����, ",������ü�¼ C", "")
    Else
        str����ʱ�� = Replace(str����, "And A.�������� Between [2] And [3]", "")
        str���� = str���� & " And C.ҽ����� Is Null "
        
        str����ʱ�� = str����ʱ�� & " And C.ҽ����� Is Not Null And C.����ʱ�� Between [2] And [3] "
        str���� = str���� & " Union All " & str����ʱ��
    End If
    
    str���� = strSqlTmp & str���� & ") A,ҩƷ�շ���¼ C, ������ü�¼ D, ���ű� B " & IIf(mParams.strSourceDep = "", "", ",Table(Cast(f_Num2List([20]) As zlTools.t_NumList)) E ") & _
              " Where C.����id = D.ID And nvl(c.��ҩ��ʽ,-999)<>-1 and A.����=C.���� And A.NO=C.NO And C.����� Is NULL " & _
              " And Nvl(D.����״̬,0)<>1 And (C.�ⷿid=[1] Or C.�ⷿid Is null) And a.�Է�����id = b.Id "
    
    If mstrDeptNode <> "" Then
        str���� = str���� & " And (b.վ�� = [21] Or b.վ�� Is Null) "
    End If
    
    '�ų��Ѿ�����ֹͣ��ҩ��No
    str���� = str���� & " and Not Exists(select 1 from ҩƷ�շ���¼ F where F.����=C.���� and F.�ⷿid=C.�ⷿid and F.no=C.no and ��ҩ��ʽ=-1) "
    
    '�ų�������Һ�������Ĺ����в����ĵ���
    str���� = str���� & " And Not Exists (Select 1 From ��Һ��ҩ���� Y Where Y.�շ�id = C.ID) "
    
    '�Ƿ���ʾ��ҩ��������
    If mcondition.bln��ʾ��ҩ�������� = False Then
        str���� = str���� & " And C.��¼״̬=1 "
    Else
        str���� = str���� & " And MOD(C.��¼״̬,3)=1 "
    End If
    
    '��Ժ��ҩ:'0-����,1-������Ժ��ҩ,2-������Ժ��ҩ
    If mcondition.int��Ժ��ҩ = 0 Then
    ElseIf mcondition.int��Ժ��ҩ = 1 Then
        str���� = str���� & " And Not Ltrim(To_Char(Nvl(C.����,0),'00')) Like '_3'"
    ElseIf mcondition.int��Ժ��ҩ = 2 Then
        str���� = str���� & " And Ltrim(To_Char(Nvl(C.����,0),'00')) Like '_3'"
    End If
    
    str���� = str���� & IIf(mParams.strSourceDep = "", "", " And C.�Է�����id+0=E.Column_Value ")
    
    If imgFilter.BorderStyle = cstFilter Then
        If Val(mParams.int����ģʽ) <= 7 Then
            If Trim(txtPati.Text) = "" Then
                str���� = str���� & " And 1 = 2 "
            End If
        Else
            If mSQLCondition.lng����ID = 0 Then
                str���� = str���� & " And 1 = 2 "
            End If
        End If
    End If
    
    If mParams.intType = 1 Then
        str���� = str���� & " And (D.�����־=1 or D.�����־=4)"
    ElseIf mParams.intType = 2 Then
        str���� = str���� & " And (D.�����־<>1 and D.�����־<>4)"
    End If
    
    
    If mcondition.int������� = 1 Then
        '���ﻮ�ۼ��������
        gstrSQL = gstrSQL & str����
    Else
        If mcondition.int������� = 3 Then
            '���ＰסԺ���е���
            strסԺ = Replace(str����, "������ü�¼", "סԺ���ü�¼")
            strסԺ = Replace(strסԺ, "And Nvl(D.����״̬,0)<>1", "")
        Else
            'סԺ����
            strסԺ = Replace(str����, "������ü�¼", "סԺ���ü�¼")
            strסԺ = Replace(strסԺ, "And Nvl(D.����״̬,0)<>1", "")
            str���� = ""
        End If
    
        If mPrives.bln���������� Then
            If img����.BorderStyle = 0 Then
                '����ʾ��������
                strסԺ = strסԺ & " And (D.�����־ <> 2 Or (D.�����־ = 2 And D.���˲���id <> D.��������id)) "
            End If
            If img����.BorderStyle = 1 And cbo����.ListIndex <> -1 Then
                'Ҫ��ʾ�������������Ҳ��˲������ڵ�ǰѡ��Ĳ���
                strסԺ = strסԺ & " And D.���˲���id = " & cbo����.ItemData(cbo����.ListIndex)
                str���� = ""
            End If
        End If
        
        If str���� = "" Then
            gstrSQL = gstrSQL & strסԺ
        Else
            gstrSQL = gstrSQL & str���� & " Union All " & strסԺ
        End If
    End If
    
    '''''Group By
    gstrSQL = gstrSQL & ") A GROUP BY A.����ʱ��,A.���ȼ�,A.����,A.����,A.���շ�,A.��ҩ��,A.No,A.����,A.����,A.�ɲ���,A.˵��," & _
        " A.���￨��,A.�����,A.���֤��,A.IC����,A.����ID,A.ҽ����,A.סԺ��, A.��������,A.�����־,A.��¼����,a.��ҩ����,A.ǩ��ʱ��,a.��������" & IIf(mParams.bln������, ",a.�����,a.���id", "")
    
    '''''Order By
    gstrSQL = gstrSQL & " Order By A.ǩ��ʱ��,A.����,A.����,A.No"
    
    Set rsData = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, _
            mSQLCondition.lngҩ��ID, _
            mSQLCondition.date��ʼ����, _
            mSQLCondition.date��������, _
            mSQLCondition.str��ʼNO, _
            mSQLCondition.str����NO, _
            UCase(mSQLCondition.str����), _
            mSQLCondition.str���￨, _
            mSQLCondition.str��ʶ��, _
            mSQLCondition.lng����ID, _
            mSQLCondition.str������, _
            mSQLCondition.str�����, _
            mSQLCondition.lngҩƷid, _
            mSQLCondition.str��ǰNO, _
            mSQLCondition.str�����, _
            lng����ID, _
            mSQLCondition.lng����ID, _
            mSQLCondition.strҽ����, _
            mSQLCondition.lngסԺ��, _
            mParams.Str����, _
            mParams.strSourceDep, _
            mstrDeptNode)
    
    stbThis.Panels(2) = ""
    
    Set cbrMenu = Me.cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_Edit_Recipe_Cancle, , True)
    Set cbrControl = Me.cbsMain(2).Controls.Find(xtpControlButton, mconMenu_Edit_Recipe_Cancle, , True)
        
    If Not rsData.EOF Then
        cbrMenu.Enabled = True
        cbrControl.Enabled = True
        stbThis.Panels(2) = "����" & rsData.RecordCount & "�Ŵ�����" & GetSumMoney(rsData)
    Else
        cbrMenu.Enabled = False
        cbrControl.Enabled = False
    End If

    Set mrsList = rsData
    If Not mfrmList Is Nothing Then
        '���˳��м�¼�ͱ����
        mblnFinding = True
        
        If Val(mParams.int����ģʽ) <= 7 Then
            strInput = txtPati.Text
        Else
            '���ѿ����ʱ����Ϊ��ID+����
            strInput = mobjcard.�ӿ���� & "|" & txtPati.Text
        End If
                
        mfrmList.ShowList mListType.����ҩ, imgFilter.BorderStyle, (mParams.blnStartCall And mParams.blnStartQueue), mParams.blnMustDosageOkProcess, mParams.blnMustDosageProcess, mParams.bln������, IDKNType.GetCurCard.����, strInput
        mfrmList.RefreshList mListType.����ҩ, mrsList, strNo, blnNoRefreshDetail
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Sub RefreshList_OverTime(Optional ByVal strNo As String, Optional ByVal blnNoRefreshDetail As Boolean)
    'ˢ�³�ʱ����ҩ�б�
    Dim blnҽ���� As Boolean
    Dim rsData As ADODB.Recordset
    Dim strSub1 As String
    Dim strSub2 As String
    Dim strסԺ As String
    Dim str���� As String
    Dim strInput As String
    Dim lng����ID As Long
    Dim strSqlTmp As String
    Dim str����ʱ�� As String
    
    On Error GoTo errHandle
    If mSQLCondition.str���֤ <> "" And Not mobjSquareCard Is Nothing Then
        If mobjSquareCard.zlGetPatiID("���֤", UCase(mSQLCondition.str���֤), False, lng����ID) = False Then lng����ID = 0
    End If
    
    gstrSQL = "Select '' As ��ɫ, ��������,'' As ѡ�� ,'0' As ��־,����,����,���շ�,��ҩ��,NO,����,ǩ��ʱ��," & _
            " to_Char(Sum(Round(���۽��," & mintMoneyDigit & ")),'" & mstrOracleMoneyForamt & "') AS ���,����,�ɲ���," & _
            " ˵��,���￨��,�����,���֤��,IC����,����ID,ҽ����,סԺ��,��ҩ����," & _
            " Sum(Round(ʵ�ս��," & mintMoneyDigit & ")) ʵ�ս��,�����־,��¼����,Zl_Get�շ����(����,NO,[1]) As �շ����,�������� " & _
            " From ("
            
    strSqlTmp = " Select A.���ȼ�,A.����,A.����,A.���շ�,A.��ҩ��,A.NO,A.����,C.���۽��,A.����,A.�ɲ���,A.˵��,A.���￨��,A.�����,A.���֤��," & _
            " A.IC����,A.����ID,A.ҽ����,A.סԺ��,d.ʵ�ս��*(Nvl(c.����,1)*c.ʵ������/(Nvl(d.����,1)*d.����)) ʵ�ս��, Nvl(A.��������,Nvl(C.ע��֤��,0)) As ��������,D.�����־,D.��¼����,D.�շ����,A.��ҩ����,c.ǩ��ʱ��,a.�������� " & _
            " From ( "
            
    str���� = "Select distinct B.���￨��,B.�����,B.���֤��,B.IC����,B.ҽ����,B.סԺ��,A.���ȼ�,A.��ҩ����,A.��������,Decode(Nvl(A.���շ�,0),1,'','(δ)')||Decode(A.����,8,'�շ�',9,'����') ����," & _
            " A.����,A.���շ�,'' ��ҩ��,A.No,A.����,To_Char(A.��������,'yyyy-MM-dd hh24:mi:ss') ����,1 �ɲ���,' ' ˵��,B.����ID, A.��������,a.�Է�����id, b.�������� " & _
            " From δ��ҩƷ��¼ A,������Ϣ B,������ü�¼ C " & _
            " Where 1=1 "
    
    '��Ҫ����
    If mParams.blnMustDosageProcess = True Then str���� = str���� & " And A.��ҩ�� Is Not Null "
    
    str���� = str���� & " And (A.�ⷿID=[1] Or A.�ⷿID Is NULL) And A.�������� Between [2] And [3] "
    
    str���� = str���� & " And A.�������� < Sysdate - (1 / 24 / 60) * [22] "

    If mSQLCondition.str��ʼNO <> "" Or mSQLCondition.str����NO <> "" Then
        If mSQLCondition.str��ʼNO <> "" And mSQLCondition.str����NO <> "" Then
            str���� = str���� & " And A.NO Between [4] And [5] "
        Else
            If mSQLCondition.str��ʼNO <> "" Then
                str���� = str���� & " And A.NO = [4] "
            Else
                str���� = str���� & " And A.NO = [5] "
            End If
        End If
    End If
    
    If mSQLCondition.str���� <> "" Then str���� = str���� & " And Upper(A.����) Like Upper([6]) "
    
    If mSQLCondition.str���￨ <> "" Then str���� = str���� & " And Upper(B.���￨��) = [7] "
    
    If mSQLCondition.str��ʶ�� <> "" Then str���� = str���� & " And Upper(DECODE(A.����,8,B.�����,B.סԺ��)) Like [8] "
    
    If mSQLCondition.lng����ID > 0 Then str���� = str���� & " And A.�Է�����ID+0=[9] "
    
    If mSQLCondition.str��ǰNO <> "" Then str���� = str���� & " And A.NO=[13] "
    
    If mSQLCondition.str����� <> "" Then str���� = str���� & " And B.�����=[14] "
    
'    If mSQLCondition.str���֤ <> "" Then str���� = str���� & " And B.���֤��=[15] "

    If mSQLCondition.str���֤ <> "" Then str���� = str���� & " And B.����ID=[15] "
    
    If mSQLCondition.lng����ID <> 0 Then str���� = str���� & " And B.����ID=[16] "
    
    If mSQLCondition.strҽ���� <> "" Then str���� = str���� & " And B.ҽ����=[17] "
    
    If mSQLCondition.lngסԺ�� <> 0 Then str���� = str���� & " And B.סԺ��=[18] "
    
    blnҽ���� = (mSQLCondition.strҽ���� <> "")
    str���� = str���� & " And A.����ID=B.����ID" & IIf(blnҽ���� = True, "", "(+)") & ""
    
    str���� = str���� & IIf(mParams.Str���� = "", "", " And (A.��ҩ���� In (Select * From Table(Cast(f_Str2list([19]) As Zltools.t_Strlist))) Or A.��ҩ���� Is Null) ")
    
    '�շѵ���
    Select Case mParams.intShowBill�շ�
        Case 0  '����ʾ����
            strSub1 = "1=2"
        Case 1  '��ʾδ�շ�
            strSub1 = "(Nvl(A.���շ�,0)=0 And A.����=8)"
        Case 2  '��ʾ���շ�
            strSub1 = "(A.���շ�=1 And A.����=8)"
        Case 3  '��ʾ���д���
            strSub1 = "A.����=8"
    End Select
    '���ʵ���
    Select Case mParams.intShowBill����
        Case 0  '����ʾ����
            strSub2 = "1=2"
        Case 1  '��ʾδ���
            strSub2 = "(Nvl(A.���շ�,0)=0 And A.����=9)"
        Case 2  '��ʾ�����
            strSub2 = "(A.���շ�=1 And A.����=9)"
        Case 3  '��ʾ���д���
            strSub2 = "A.����=9"
    End Select
    
    str���� = str���� & " And (" & strSub1 & " Or " & strSub2 & ")"
    
    str���� = str���� & " And Mod(C.��¼����, 10) = Decode(A.����, 8, 1, 2) And A.No = C.No And A.�ⷿid = C.ִ�в���id "
              
    If mParams.bln����ʱ����� = False Then
'        str���� = Replace(str����, ",������ü�¼ C", "")
    Else
        str����ʱ�� = Replace(str����, "And A.�������� Between [2] And [3]", "")
        str���� = str���� & " And C.ҽ����� Is Null "
        
        str����ʱ�� = str����ʱ�� & " And C.ҽ����� Is Not Null And C.����ʱ�� Between [2] And [3] "
        str���� = str���� & " Union All " & str����ʱ��
    End If
    
    str���� = strSqlTmp & str���� & ") A,ҩƷ�շ���¼ C, ������ü�¼ D, ���ű� B " & IIf(mParams.strSourceDep = "", "", ",Table(Cast(f_Num2List([20]) As zlTools.t_NumList)) E ") & _
              " Where C.����id = D.ID And nvl(c.��ҩ��ʽ,-999)<>-1 and A.����=C.���� And A.NO=C.NO And C.����� Is NULL " & _
              " And Nvl(D.����״̬,0)<>1 And (C.�ⷿid=[1] Or C.�ⷿid Is null) And a.�Է�����id = b.Id "
    
    If mstrDeptNode <> "" Then
        str���� = str���� & " And (b.վ�� = [21] Or b.վ�� Is Null) "
    End If
    
    '�ų��Ѿ�����ֹͣ��ҩ��No
    str���� = str���� & " and Not Exists(select 1 from ҩƷ�շ���¼ F where F.����=C.���� and F.�ⷿid=C.�ⷿid and F.no=C.no and ��ҩ��ʽ=-1) "
    
    
    '�ų�������Һ�������Ĺ����в����ĵ���
    str���� = str���� & " And Not Exists (Select 1 From ��Һ��ҩ���� Y Where Y.�շ�id = C.ID) "
    
    '�Ƿ���ʾ��ҩ��������
    If mcondition.bln��ʾ��ҩ�������� = False Then
        str���� = str���� & " And C.��¼״̬=1 "
    Else
        str���� = str���� & " And MOD(C.��¼״̬,3)=1 "
    End If
    
    '��Ժ��ҩ:'0-����,1-������Ժ��ҩ,2-������Ժ��ҩ
    If mcondition.int��Ժ��ҩ = 0 Then
    ElseIf mcondition.int��Ժ��ҩ = 1 Then
        str���� = str���� & " And Not Ltrim(To_Char(Nvl(C.����,0),'00')) Like '_3'"
    ElseIf mcondition.int��Ժ��ҩ = 2 Then
        str���� = str���� & " And Ltrim(To_Char(Nvl(C.����,0),'00')) Like '_3'"
    End If
    
    str���� = str���� & IIf(mParams.strSourceDep = "", "", " And C.�Է�����id+0=E.Column_Value ")
    
    If imgFilter.BorderStyle = cstFilter Then
        If Val(mParams.int����ģʽ) <= 7 Then
            If Trim(txtPati.Text) = "" Then
                str���� = str���� & " And 1 = 2 "
            End If
        Else
            If mSQLCondition.lng����ID = 0 Then
                str���� = str���� & " And 1 = 2 "
            End If
        End If
    End If
    
    If mParams.intType = 1 Then
        str���� = str���� & " And (D.�����־=1 or D.�����־=4)"
    ElseIf mParams.intType = 2 Then
        str���� = str���� & " And (D.�����־<>1 and D.�����־<>4)"
    End If
    
    If mcondition.int������� = 1 Then
        '���ﻮ�ۼ��������
        gstrSQL = gstrSQL & str����
    Else
        If mcondition.int������� = 3 Then
            '���ＰסԺ���е���
            strסԺ = Replace(str����, "������ü�¼", "סԺ���ü�¼")
            strסԺ = Replace(strסԺ, "And Nvl(D.����״̬,0)<>1", "")
        Else
            'סԺ����
            strסԺ = Replace(str����, "������ü�¼", "סԺ���ü�¼")
            strסԺ = Replace(strסԺ, "And Nvl(D.����״̬,0)<>1", "")
            str���� = ""
        End If
    
        If mPrives.bln���������� Then
            If img����.BorderStyle = 0 Then
                '����ʾ��������
                strסԺ = strסԺ & " And (D.�����־ <> 2 Or (D.�����־ = 2 And D.���˲���id <> D.��������id)) "
            End If
            If img����.BorderStyle = 1 And cbo����.ListIndex <> -1 Then
                'Ҫ��ʾ�������������Ҳ��˲������ڵ�ǰѡ��Ĳ���
                strסԺ = strסԺ & " And D.���˲���id = " & cbo����.ItemData(cbo����.ListIndex)
                str���� = ""
            End If
        End If
        
        If str���� = "" Then
            gstrSQL = gstrSQL & strסԺ
        Else
            gstrSQL = gstrSQL & str���� & " Union All " & strסԺ
        End If
    End If
    
    '''''Group By
    gstrSQL = gstrSQL & ") A GROUP BY A.���ȼ�,A.����,A.����,A.���շ�,A.��ҩ��,A.No,A.����,A.����,A.�ɲ���,A.˵��," & _
        " A.���￨��,A.�����,A.���֤��,A.IC����,A.����ID,A.ҽ����,A.סԺ��,A.��������,A.�����־,A.��¼����,a.��ҩ����,A.ǩ��ʱ��,a.�������� "
    
    '''''Order By
    gstrSQL = gstrSQL & " Order By A.ǩ��ʱ��,A.����,A.����,A.No"
    
    Set rsData = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, _
            mSQLCondition.lngҩ��ID, _
            mSQLCondition.date��ʼ����, _
            mSQLCondition.date��������, _
            mSQLCondition.str��ʼNO, _
            mSQLCondition.str����NO, _
            UCase(mSQLCondition.str����), _
            mSQLCondition.str���￨, _
            mSQLCondition.str��ʶ��, _
            mSQLCondition.lng����ID, _
            mSQLCondition.str������, _
            mSQLCondition.str�����, _
            mSQLCondition.lngҩƷid, _
            mSQLCondition.str��ǰNO, _
            mSQLCondition.str�����, _
            lng����ID, _
            mSQLCondition.lng����ID, _
            mSQLCondition.strҽ����, _
            mSQLCondition.lngסԺ��, _
            mParams.Str����, _
            mParams.strSourceDep, _
            mstrDeptNode, _
            mSQLCondition.intOverTime)
    
    stbThis.Panels(2) = ""
    If Not rsData.EOF Then
        stbThis.Panels(2) = "��������ʱ��" & mParams.intOverTime & "����δ��ҩ����������" & rsData.RecordCount & "�ţ�" & GetSumMoney(rsData)
    Else
        stbThis.Panels(2) = "��������ʱ��" & mParams.intOverTime & "����δ��ҩ����������0��"
    End If

    Set mrsList = rsData
    If Not mfrmList Is Nothing Then
        If Val(mParams.int����ģʽ) <= 7 Then
            strInput = txtPati.Text
        Else
            '���ѿ����ʱ����Ϊ��ID+����
            strInput = mobjcard.�ӿ���� & "|" & txtPati.Text
        End If
        
        mfrmList.ShowList mListType.��ʱδ��, imgFilter.BorderStyle, (mParams.blnStartCall And mParams.blnStartQueue), mParams.blnMustDosageOkProcess, mParams.blnMustDosageProcess, mParams.bln������, IDKNType.GetCurCard.����, strInput
        mfrmList.RefreshList mListType.��ʱδ��, mrsList, strNo, blnNoRefreshDetail
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub




Private Sub Load����()
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo errHandle
    
    If cbo����.ListCount > 0 And mstrDeptNode = cbo����.Tag Then Exit Sub
    
    '����
    gstrSQL = " Select ����||'-'||���� ����,ID From ���ű� " & _
             " Where ID in (Select ����ID From ��������˵�� Where ��������='����' And ������� IN(2,3))" & _
             " And (����ʱ�� Is Null Or ����ʱ��=To_Date('3000-01-01','yyyy-MM-dd')) "
    
    If mstrDeptNode <> "" Then
        gstrSQL = gstrSQL & " And (վ�� = [1] Or վ�� Is Null) "
    End If
    
    gstrSQL = gstrSQL & " Order By ����||'-'||���� "
    
    Set rsTmp = zldatabase.OpenSQLRecord(gstrSQL, "ȡ���в���", mstrDeptNode)
    
    With cbo����
        .Clear
        .Tag = mstrDeptNode
        Do While Not rsTmp.EOF
            .AddItem rsTmp!����
            .ItemData(.NewIndex) = rsTmp!Id
            rsTmp.MoveNext
        Loop
        If .ListIndex <> -1 Then
            .ListIndex = 0
        End If
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
Public Sub RefreshList_Return(Optional ByVal strNo As String, Optional ByVal blnNoRefreshDetail As Boolean)
    'ˢ����ҩ�б�
    Dim rsData As ADODB.Recordset
    Dim strSqlSendType As String
    Dim strSqlSourceDep As String
    Dim strSql�������� As String
    Dim strSqlFilter As String
    Dim strSqlSub As String
    Dim strSqlҽ���� As String
    Dim strSub1 As String
    Dim strSub2 As String
    Dim blnҽ���� As Boolean
    Dim strGroup As String
    Dim str���� As String
    Dim strסԺ As String
    Dim strSql���� As String
    Dim bln����ʾ���� As Boolean
    Dim lng����ID As Long
    
    On Error GoTo errHandle
    
    If mSQLCondition.str���֤ <> "" And Not mobjSquareCard Is Nothing Then
        If mobjSquareCard.zlGetPatiID("���֤", UCase(mSQLCondition.str���֤), False, lng����ID) = False Then lng����ID = 0
    End If
    
    ''strCond1
    If mSQLCondition.str��ʼNO <> "" Or mSQLCondition.str����NO <> "" Then
        If mSQLCondition.str��ʼNO <> "" And mSQLCondition.str����NO <> "" Then
            strSqlSub = strSqlSub & " And A.NO Between [4] And [5] "
        Else
            If mSQLCondition.str��ʼNO <> "" Then
                strSqlSub = strSqlSub & " And A.NO = [4] "
            Else
                strSqlSub = strSqlSub & " And A.NO = [5] "
            End If
        End If
    End If

    If mcondition.int������� = 2 Then
        strSql�������� = " And A.���� = 9 "
    Else
        strSql�������� = " And A.���� In (8,9)"
    End If
    
    If mSQLCondition.str���� <> "" Then strSqlSub = strSqlSub & " And Upper(H.����) Like [6] "
    
    If mSQLCondition.str���￨ <> "" Then strSqlSub = strSqlSub & " And Upper(B.���￨��) = [7] "
    
    If mSQLCondition.str��ʶ�� <> "" Then strSqlSub = strSqlSub & " And Upper(DECODE(A.����,8,B.�����,B.סԺ��)) Like [8] "
    
    If mSQLCondition.lng����ID > 0 Then strSqlSub = strSqlSub & " And A.�Է�����ID+0=[9] "
    
    If mSQLCondition.str������ <> "" Then strSqlSub = strSqlSub & " And A.������=[10] "
    
    If mSQLCondition.str����� <> "" Then strSqlSub = strSqlSub & " And A.�����=[11] "
    
    If mSQLCondition.lngҩƷid > 0 Then strSqlSub = strSqlSub & " And A.ҩƷID+0=[12] "
    
    If mSQLCondition.str��ǰNO <> "" Then strSqlSub = strSqlSub & " And A.NO=[13] "
    
    If mSQLCondition.str����� <> "" Then strSqlSub = strSqlSub & " And B.�����=[14] "
    
'    If mSQLCondition.str���֤ <> "" Then strSqlSub = strSqlSub & " And B.���֤��=[15] "
    
    If mSQLCondition.str���֤ <> "" Then strSqlSub = strSqlSub & " And B.����ID=[15] "
    
    If mSQLCondition.lng����ID <> 0 Then strSqlSub = strSqlSub & " And B.����ID=[16] "
    
    If mSQLCondition.strҽ���� <> "" Then strSqlSub = strSqlSub & " And B.ҽ����=[17] "
    
    If mSQLCondition.lngסԺ�� <> 0 Then strSqlSub = strSqlSub & " And B.סԺ��=[18] "
    
    ''strSqlҽ����
    blnҽ���� = (mSQLCondition.strҽ���� <> "")
    strSqlҽ���� = " AND H.����ID=B.����ID" & IIf(blnҽ���� = True, "", "(+)") & ""
    
    ''strSqlSendType
    '��Ժ��ҩ:'0-����,1-������Ժ��ҩ,2-������Ժ��ҩ
    If mcondition.int��Ժ��ҩ = 0 Then
    ElseIf mcondition.int��Ժ��ҩ = 1 Then
        strSqlSendType = " And Not Ltrim(To_Char(Nvl(A.����,0),'00')) Like '_3'"
    ElseIf mcondition.int��Ժ��ҩ = 2 Then
        strSqlSendType = " And Ltrim(To_Char(Nvl(A.����,0),'00')) Like '_3'"
    End If
    
    ''strSqlSourceDep
    strSqlSourceDep = IIf(mParams.strSourceDep = "", "", " And A.�Է�����id In (Select * From Table(Cast(f_Num2list([19]) As Zltools.t_Numlist))) ")

    ''strSqlFilter
    If imgFilter.BorderStyle = cstFilter Then
        If Val(mParams.int����ģʽ) <= 7 Then
            If Trim(txtPati.Text) = "" Then
                strSqlFilter = " And 1 = 2 "
            End If
        Else
            If mSQLCondition.lng����ID = 0 Then
                strSqlFilter = " And 1 = 2 "
            End If
        End If
    End If
    
    ''������ҩ
    If mPrives.bln���������� Then
        If img����.BorderStyle = 0 Then
            '����ʾ��������
            strSql���� = " And (H.�����־ <> 2 Or (H.�����־ = 2 And H.���˲���id <> H.��������id)) "
        End If
        If img����.BorderStyle = 1 And cbo����.ListIndex <> -1 Then
            'Ҫ��ʾ�������������Ҳ��˲������ڵ�ǰѡ��Ĳ���
            strSql���� = " And H.���˲���id = " & cbo����.ItemData(cbo����.ListIndex)
            bln����ʾ���� = True
        End If
    End If
    
    '�շѵ���
    Select Case mParams.intShowBill�շ�
        Case 0  '����ʾ����
            strSub1 = "1=2"
        Case 3  '��ʾ���д���
            strSub1 = "A.����=8"
    End Select
    '���ʵ���
    Select Case mParams.intShowBill����
        Case 0  '����ʾ����
            strSub2 = "1=2"
        Case 3  '��ʾ���д���
            strSub2 = "A.����=9"
    End Select
    
    '�շѵ���
    Select Case mParams.intShowBill�շ�
        Case 0  '����ʾ����
            strSub1 = "1=2"
        Case 1  '��ʾδ�շ�
            strSub1 = "(Nvl(H.��¼״̬,0)=0 And A.����=8)"
        Case 2  '��ʾ���շ�
            strSub1 = "(H.��¼״̬>=1 And A.����=8)"
        Case 3  '��ʾ���д���
            strSub1 = "A.����=8"
    End Select
    '���ʵ���
    Select Case mParams.intShowBill����
        Case 0  '����ʾ����
            strSub2 = "1=2"
        Case 1  '��ʾδ���
            strSub2 = "(Nvl(H.��¼״̬,0)=0 And A.����=9)"
        Case 2  '��ʾ�����
            strSub2 = "(H.��¼״̬>=1 And A.����=9)"
        Case 3  '��ʾ���д���
            strSub2 = "A.����=9"
    End Select
    
    strSqlSub = strSqlSub & " And (" & strSub1 & " Or " & strSub2 & ")"
    
    '����κ�һ��ҩƷ�������������һ������ϸ�ֱ���������д��ڵ��������ˣ���ֱ��ͨ������UNION�󱸵ķ�ʽ���
    '���ڷ��ü�¼������㣬��������Ҫ������ͨ�����ü�¼������������Ӻ���Ч����ȫ��ɨ�裬��ˣ�ֻ��ͨ����������SQL UNION ������SQL�ķ�ʽ���
    If mcondition.bln��ʾ���̵��� = False Then
        gstrSQL = " SELECT DISTINCT '' As ��ɫ, A.��������,'' As ѡ��,'0' As ��־,Decode(Nvl(h.��¼״̬, 0),  0,'(δ)','') || Decode(a.����, 8, '�շ�', 9, '����') ����," & _
                 "      A.����,1 ���շ�,A.����� ��ҩ��,A.NO,H.����,trim(to_char(sum(A.���۽��),'" & mstrOracleMoneyForamt & "')) AS ���,trim(to_char(Sum((Nvl(a.����, 1) * a.ʵ������)/(Nvl(H.����,1)*H.����) * H.ʵ�ս��),'" & mstrOracleMoneyForamt & "')) As ʵ�ս��," & _
                 "      TO_CHAR(A.�������,'YYYY-MM-DD HH24:MI:SS') ����,1 �ɲ���,' ' ˵��,B.���￨��,B.�����,B.���֤��,B.IC����,B.����ID,B.ҽ����,B.סԺ��,H.�����־, H.��¼����,Zl_Get�շ����(A.����,A.NO,[1]) As �շ����,B.��������,A.δȡҩ " & _
                 " FROM " & _
                 "      (SELECT A.ID,A.NO,A.����,A.ҩƷID,A.����ID,A.����,A.����,A.Ч��," & _
                 "          NVL(A.����,1) ����,A.ʵ������,NVL(A.����,1)*A.ʵ������-B.�ѷ����� ��������,B.�ѷ�����,A.��¼״̬,A.��ҩ����," & _
                 "          A.���ۼ�,round(a.���ۼ�*Nvl(a.����, 1)*a.ʵ������," & mintMoneyDigit & ") ���۽��,A.����,A.Ƶ��,A.�÷�,A.ժҪ,A.�����,A.�������,A.�Է�����ID,A.�ⷿID, A.������, A.��������,A.δȡҩ " & _
                 "      FROM" & _
                 "          (SELECT A.ID,A.NO,A.����,A.ҩƷID,A.���,A.����ID,A.����,A.����,A.Ч��,A.����,A.ʵ������,A.��¼״̬,A.��ҩ����,A.���ۼ�,A.����,A.Ƶ��,A.�÷�,A.ժҪ,A.�����,A.�������,A.�Է�����ID,A.�ⷿID, A.������, Nvl(A.ע��֤��, 0) As ��������,Nvl(A.�Ƿ�δȡҩ,0) As δȡҩ " & _
                 "          FROM ҩƷ�շ���¼ A" & _
                 "          WHERE nvl(A.��ҩ��ʽ,-999)<>-1 and A.����� IS NOT NULL AND (A.��¼״̬=1 OR MOD(A.��¼״̬,3)=0)" & _
                 "          AND A.�ⷿID+0=[1] And A.������� Between [2] And [3]  " & strSql�������� & strSqlSendType & _
                 "          And Not Exists (Select 1 From ��Һ��ҩ���� Y,ҩƷ�շ���¼ Z Where y.�շ�id=Z.ID AND Z.NO= A.NO And z.����=a.���� And z.�ⷿid = a.�ⷿid) " & _
                 "          ) A," & _
                 "          (SELECT A.NO,A.����,A.ҩƷID,A.���,SUM(NVL(A.����,1)*A.ʵ������) �ѷ�����,SUM(A.���۽��) ���۽��" & _
                 "          FROM ҩƷ�շ���¼ A" & _
                 "          WHERE nvl(A.��ҩ��ʽ,-999)<>-1 and A.����� IS NOT NULL" & strSql�������� & strSqlSendType & _
                 "          AND A.�ⷿID+0=[1] And A.������� Between [2] And [3]  " & strSqlSourceDep & _
                 "          GROUP BY A.NO,A.����,A.ҩƷID,A.���) B"
        gstrSQL = gstrSQL & _
                 "      WHERE A.NO = B.NO AND A.���� = B.���� AND A.ҩƷID+0 = B.ҩƷID AND A.��� = B.��� AND B.�ѷ�����<>0" & _
                 "     ) A,������ü�¼ H,������Ϣ B" & _
                 " WHERE A.�ⷿID+0=[1] " & _
                 " " & strSqlSub & strSqlFilter & strSqlҽ���� & _
                 " AND (A.��¼״̬=1 OR MOD(A.��¼״̬,3)=0) AND A.����� IS NOT NULL AND A.����ID=H.ID AND A.ʵ������<>0 "
    Else
        gstrSQL = " SELECT DISTINCT '' As ��ɫ, A.��������,'' As ѡ��,'0' As ��־,Decode(Nvl(h.��¼״̬, 0),  0,'(δ)','') || Decode(a.����, 8, '�շ�', 9, '����') ����,A.����,1 ���շ�,A.����� ��ҩ��," & _
                  "      A.NO,H.����,trim(to_char(sum(A.���۽��),'" & mstrOracleMoneyForamt & "')) AS ���,trim(to_char(Sum((Nvl(a.����, 1) * a.ʵ������)/(Nvl(H.����,1)*H.����) * H.ʵ�ս��),'" & mstrOracleMoneyForamt & "')) As ʵ�ս��,TO_CHAR(A.�������,'YYYY-MM-DD HH24:MI:SS') ����,A.�ɲ���," & _
                  "      DECODE(A.��¼״̬,1,'��1�η�ҩ',DECODE(MOD(A.��¼״̬,3),0,'��1�η�ҩ',1,'��'||(FLOOR(A.��¼״̬/3)+1)||'�η�ҩ',2,'��'||(FLOOR(A.��¼״̬/3)+1)||'����ҩ')) ˵��,B.���￨��,B.�����,B.���֤��,B.IC����,B.����ID,B.ҽ����,B.סԺ��,H.�����־, H.��¼����,Zl_Get�շ����(A.����,A.NO,[1]) As �շ����,B.��������,A.δȡҩ " & _
                  " FROM " & _
                  "      (SELECT * FROM" & _
                  "          (SELECT A.ID,A.NO,A.����,A.ҩƷID,A.����ID,A.����,A.����,A.Ч��," & _
                  "              NVL(A.����,1) ����,A.ʵ������,NVL(A.����,1)*A.ʵ������-B.�ѷ����� ��������,B.�ѷ�����,A.��¼״̬,A.��ҩ����," & _
                  "              A.���ۼ� , round(a.���ۼ�*Nvl(a.����, 1)*a.ʵ������," & mintMoneyDigit & ") ���۽��, A.����, A.Ƶ��, A.�÷�, A.ժҪ, A.�����, A.�������, A.�Է�����ID, A.�ⷿID,1 �ɲ���, A.������, A.��������,A.δȡҩ " & _
                  "          FROM" & _
                  "              (SELECT A.ID,A.NO,A.����,A.ҩƷID,A.���,A.����ID,A.����,A.����,A.Ч��,A.����,A.ʵ������,A.��¼״̬,A.��ҩ����,A.���ۼ�,A.���۽��, A.����, A.Ƶ��, A.�÷�, A.ժҪ, A.�����, A.�������, A.�Է�����ID, A.�ⷿID, A.������, Nvl(A.ע��֤��, 0) As ��������,Nvl(A.�Ƿ�δȡҩ,0) As δȡҩ " & _
                  "              FROM ҩƷ�շ���¼ A" & _
                  "              WHERE nvl(a.��ҩ��ʽ,-999)<>-1 and A.����� IS NOT NULL AND (A.��¼״̬=1 OR MOD(A.��¼״̬,3)=0)" & _
                  "              AND A.�ⷿID+0=[1] And A.������� Between [2] And [3]  " & strSql�������� & strSqlSendType & _
                  "              And Not Exists (Select 1 From ��Һ��ҩ���� Y,ҩƷ�շ���¼ Z Where y.�շ�id=Z.ID AND Z.NO=A.NO And z.����=a.���� And z.�ⷿid = a.�ⷿid)  " & _
                  "              ) A," & _
                  "              (SELECT A.NO,A.����,A.ҩƷID,A.���,SUM(NVL(A.����,1)*A.ʵ������) �ѷ�����" & _
                  "              FROM ҩƷ�շ���¼ A" & _
                  "              WHERE nvl(a.��ҩ��ʽ,-999)<>-1 and A.����� IS NOT NULL " & strSql�������� & strSqlSendType & _
                  "              AND A.�ⷿID+0=[1] And A.������� Between [2] And [3]  " & strSqlSourceDep & _
                  "              GROUP BY A.NO,A.����,A.ҩƷID,A.���) B"
         gstrSQL = gstrSQL & _
                  "          WHERE A.NO = B.NO AND A.���� = B.���� AND A.ҩƷID+0 = B.ҩƷID AND A.��� = B.���)" & _
                  "          UNION" & _
                  "          SELECT A.ID,A.NO,A.����,A.ҩƷID,A.����ID,A.����,A.����,A.Ч��," & _
                  "          NVL(A.����,1) ����,A.ʵ������,0 ������,0 �ѷ�����,A.��¼״̬,A.��ҩ����," & _
                  "          A.���ۼ� , round(A.���۽��," & mintMoneyDigit & ") ���۽��, A.����, A.Ƶ��, A.�÷�, A.ժҪ, A.�����, A.�������, A.�Է�����ID, A.�ⷿID," & _
                  "          DECODE(��¼״̬,1,1,DECODE(MOD(��¼״̬,3),0,1,MOD(��¼״̬,3)+1)) �ɲ���, A.������, Nvl(A.ע��֤��, 0) As ��������,Nvl(A.�Ƿ�δȡҩ,0) As δȡҩ " & _
                  "          FROM ҩƷ�շ���¼ A" & _
                  "          WHERE nvl(a.��ҩ��ʽ,-999)<>-1 and Not Exists (Select 1 From ��Һ��ҩ���� Y,ҩƷ�շ���¼ Z Where y.�շ�id=Z.ID AND  Z.NO= A.NO And z.����=a.���� And z.�ⷿid = a.�ⷿid) and NOT (��¼״̬=1 OR MOD(��¼״̬,3)=0) And A.������� Between [2] And [3]  " & strSql�������� & strSqlSendType & strSqlSourceDep
         gstrSQL = gstrSQL & _
                  "     ) A,������ü�¼ H,������Ϣ B" & _
                  " WHERE A.�ⷿID+0=[1] " & _
                  " " & strSqlSub & strSqlFilter & strSqlҽ���� & _
                  " AND A.����� IS NOT NULL AND A.����ID=H.ID "
    End If
    
    'Group
    If mcondition.bln��ʾ���̵��� = False Then
        strGroup = " GROUP BY A.��������,Decode(Nvl(h.��¼״̬, 0),  0,'(δ)','') || Decode(a.����, 8, '�շ�', 9, '����'),A.����,1,A.�����,A.NO,H.����," & _
            " TO_CHAR(A.�������,'YYYY-MM-DD HH24:MI:SS'),B.���￨��,B.�����,B.���֤��,B.IC����,B.����ID,B.ҽ����,B.סԺ��,H.�����־, H.��¼����,B.��������,A.δȡҩ "
    Else
        strGroup = " GROUP BY A.��������,Decode(Nvl(h.��¼״̬, 0),  0,'(δ)','') || Decode(a.����, 8, '�շ�', 9, '����') ,A.����,1,A.�����," & _
            " A.NO,H.����,TO_CHAR(A.�������,'YYYY-MM-DD HH24:MI:SS'),A.�ɲ���," & _
            " DECODE(A.��¼״̬,1,'��1�η�ҩ',DECODE(MOD(A.��¼״̬,3),0,'��1�η�ҩ',1,'��'||(FLOOR(A.��¼״̬/3)+1)||'�η�ҩ',2,'��'||(FLOOR(A.��¼״̬/3)+1)||'����ҩ')),B.���￨��,B.�����,B.���֤��,B.IC����,B.����ID,B.ҽ����,B.סԺ��,H.�����־, H.��¼����,B.��������,A.δȡҩ "
    End If
    
    
    
    If mParams.intType = 1 Then
        gstrSQL = gstrSQL & " And (H.�����־=1 or H.�����־=4)"
    ElseIf mParams.intType = 2 Then
        gstrSQL = gstrSQL & " And (H.�����־<>1 and H.�����־<>4)"
    End If
    
    '�������סԺ
    If mcondition.int������� = 1 Then
        '���ﻮ�ۼ��������
        gstrSQL = gstrSQL & strGroup
    Else
        If mcondition.int������� = 3 Then
            '���ＰסԺ���е���
            str���� = gstrSQL
            strסԺ = Replace(str����, "������ü�¼", "סԺ���ü�¼")
            
            str���� = str���� & strGroup
            strסԺ = strסԺ & strSql���� & strGroup
            
            If bln����ʾ���� = True Then
                gstrSQL = strסԺ
            Else
                gstrSQL = str���� & " Union All " & strסԺ
            End If
        Else
            'סԺ����
            strסԺ = Replace(gstrSQL, "������ü�¼", "סԺ���ü�¼")
            strסԺ = strסԺ & strSql���� & strGroup
            gstrSQL = strסԺ
        End If
    End If
     
    'order by
    gstrSQL = gstrSQL & " order by ����,����,NO "
     
    Dim blnMoved As Boolean
    Dim str��ʼ���� As String, strsql As String
    
    str��ʼ���� = Format(mSQLCondition.date��ʼ����, "yyyy-mm-dd hh:mm:ss")
    
    '�жϴӿ�ʼ���ں��Ƿ����ת���Ĵ�������
    blnMoved = Sys.IsMovedByDate(str��ʼ����)
    
    '�����������ת��������Ҫͬʱ�Ӻ󱸱�����ȡ����
    If blnMoved Then
        strsql = gstrSQL
        strsql = Replace(strsql, "ҩƷ�շ���¼", "HҩƷ�շ���¼")
        strsql = Replace(strsql, "������ü�¼", "H������ü�¼")
        strsql = Replace(strsql, "סԺ���ü�¼", "HסԺ���ü�¼")
        gstrSQL = gstrSQL & " UNION ALL " & strsql
    End If
     
    Set rsData = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, _
            mSQLCondition.lngҩ��ID, _
            mSQLCondition.date��ʼ����, _
            mSQLCondition.date��������, _
            mSQLCondition.str��ʼNO, _
            mSQLCondition.str����NO, _
            UCase(mSQLCondition.str����), _
            mSQLCondition.str���￨, _
            mSQLCondition.str��ʶ��, _
            mSQLCondition.lng����ID, _
            mSQLCondition.str������, _
            mSQLCondition.str�����, _
            mSQLCondition.lngҩƷid, _
            mSQLCondition.str��ǰNO, _
            mSQLCondition.str�����, _
            lng����ID, _
            mSQLCondition.lng����ID, _
            mSQLCondition.strҽ����, _
            mSQLCondition.lngסԺ��, _
            mParams.strSourceDep)

    stbThis.Panels(2) = ""
    If Not rsData.EOF Then
        stbThis.Panels(2) = "����" & rsData.RecordCount & "�Ŵ�����" & GetSumMoney(rsData)
    End If
    
    Set mrsList = rsData
    If Not mrsList Is Nothing Then mfrmList.RefreshList mListType.��ҩ, mrsList, strNo, blnNoRefreshDetail
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub RefreshList_Abolish()
    'ˢ��ȡ����ҩ�б�
    Dim blnҽ���� As Boolean
    Dim rsData As ADODB.Recordset
    Dim strSub1 As String
    Dim strSub2 As String
    Dim strסԺ As String
    Dim str���� As String
    Dim lng����ID As Long
    Dim strSqlTmp As String
    Dim str����ʱ�� As String
    
    On Error GoTo errHandle
    If mSQLCondition.str���֤ <> "" And Not mobjSquareCard Is Nothing Then
        If mobjSquareCard.zlGetPatiID("���֤", UCase(mSQLCondition.str���֤), False, lng����ID) = False Then lng����ID = 0
    End If
    
    gstrSQL = "Select '' As ��ɫ, ��������,'' As ѡ�� ,'0' As ��־,����,����,���շ�,��ҩ��,NO,����,ǩ��ʱ��," & _
            " to_Char(Sum(Round(���۽��," & mintMoneyDigit & ")),'" & mstrOracleMoneyForamt & "') AS ���,����," & _
            " �ɲ���,˵��,���￨��,�����,���֤��,IC����,����ID,ҽ����,סԺ��,��ҩ����," & _
            " Sum(Round(ʵ�ս��," & mintMoneyDigit & ")) ʵ�ս��,�����־,��¼����,Zl_Get�շ����(����,NO,[1]) As �շ����,�������� " & _
            " From ("
            
    strSqlTmp = " Select A.���ȼ�,A.����,A.����,A.���շ�,A.��ҩ��,A.NO,A.����,C.���۽��,A.����,A.�ɲ���,A.˵��,A.���￨��,A.�����,A.���֤��," & _
            " A.IC����,A.����ID,A.ҽ����,A.סԺ��,d.ʵ�ս��, Nvl(A.��������,Nvl(C.ע��֤��,0)) As ��������,D.�����־,D.��¼����,D.�շ����,A.��ҩ����,c.ǩ��ʱ��,a.�������� " & _
            " From ("
            
    str���� = "Select distinct B.���￨��,B.�����,B.���֤��,B.IC����,B.ҽ����,B.סԺ��,A.���ȼ�,A.��ҩ����,A.��������,Decode(Nvl(A.���շ�,0),1,'','(δ)')||Decode(A.����,8,'�շ�',9,'����') ����," & _
            " A.����,A.���շ�,'' ��ҩ��,A.No,A.����,To_Char(A.��������,'yyyy-MM-dd hh24:mi:ss') ����,1 �ɲ���,' ' ˵��,B.����ID, A.��������,a.�Է�����id, b.�������� " & _
            "  From δ��ҩƷ��¼ A,������Ϣ B,������ü�¼ C " & _
            "  Where A.��ҩ�� Is Not Null "
    
    '��Ҫ����
    str���� = str���� & " And (A.�ⷿID=[1] Or A.�ⷿID Is NULL) And A.�������� Between [2] And [3] "
    
    If mSQLCondition.str��ʼNO <> "" Or mSQLCondition.str����NO <> "" Then
        If mSQLCondition.str��ʼNO <> "" And mSQLCondition.str����NO <> "" Then
            str���� = str���� & " And A.NO Between [4] And [5] "
        Else
            If mSQLCondition.str��ʼNO <> "" Then
                str���� = str���� & " And A.NO = [4] "
            Else
                str���� = str���� & " And A.NO = [5] "
            End If
        End If
    End If
    
    If mSQLCondition.str���� <> "" Then str���� = str���� & " And Upper(A.����) Like [6] "
    
    If mSQLCondition.str���￨ <> "" Then str���� = str���� & " And Upper(B.���￨��) = [7] "
    
    If mSQLCondition.str��ʶ�� <> "" Then str���� = str���� & " And Upper(DECODE(A.����,8,B.�����,B.סԺ��)) Like [8] "
    
    If mSQLCondition.lng����ID > 0 Then str���� = str���� & " And A.�Է�����ID+0=[9] "
    
    If mSQLCondition.str��ǰNO <> "" Then str���� = str���� & " And A.NO=[13] "
    
    If mSQLCondition.str����� <> "" Then str���� = str���� & " And B.�����=[14] "
    
'    If mSQLCondition.str���֤ <> "" Then str���� = str���� & " And B.���֤��=[15] "
    
    If mSQLCondition.str���֤ <> "" Then str���� = str���� & " And B.����ID=[15] "
    
    If mSQLCondition.lng����ID <> 0 Then str���� = str���� & " And B.����ID=[16] "
    
    If mSQLCondition.strҽ���� <> "" Then str���� = str���� & " And B.ҽ����=[17] "
    
    If mSQLCondition.lngסԺ�� <> 0 Then str���� = str���� & " And B.סԺ��=[18] "
    
            
    blnҽ���� = (mSQLCondition.strҽ���� <> "")
    str���� = str���� & " And A.����ID=B.����ID" & IIf(blnҽ���� = True, "", "(+)") & ""
    
    str���� = str���� & IIf(mParams.Str���� = "", "", " And (A.��ҩ���� In (Select * From Table(Cast(f_Str2list([19]) As Zltools.t_Strlist))) Or A.��ҩ���� Is Null) ")
    
    '������ʾ���Զ���ӡ������:ע��"δ��ҩƷ��¼"�ı���ΪA
    Select Case mParams.intShowBill�շ�
        Case 0  '����ʾ����
            strSub1 = "1=2"
        Case 1  '��ʾδ�շ�
            strSub1 = "A.����<>9 And Nvl(A.���շ�,0)=0 And A.����=8"
        Case 2  '��ʾ���շ�
            strSub1 = "A.����<>9 And A.���շ�=1 And A.����=8"
        Case 3  '��ʾ���д���
            strSub1 = "A.����<>9 And A.����=8"
    End Select
    Select Case mParams.intShowBill����
        Case 0  '����ʾ����
            strSub2 = "1=2"
        Case 1  '��ʾδ���
            strSub2 = "A.����<>8 And Nvl(A.���շ�,0)=0 And A.����=9"
        Case 2  '��ʾ�����
            strSub2 = "A.����<>8 And A.���շ�=1 And A.����=9"
        Case 3  '��ʾ���д���
            strSub2 = "A.����<>8 And A.����=9"
    End Select
    
    str���� = str���� & " And A.���� IN(8,9) And (" & strSub1 & " Or " & strSub2 & ")"
    
    str���� = str���� & " And Mod(C.��¼����, 10) = Decode(A.����, 8, 1, 2) And A.No = C.No And A.�ⷿid = C.ִ�в���id "
            
    If mParams.bln����ʱ����� = False Then
'        str���� = Replace(str����, ",������ü�¼ C", "")
    Else
        str����ʱ�� = Replace(str����, "And A.�������� Between [2] And [3]", "")
        str���� = str���� & " And C.ҽ����� Is Null "
        
        str����ʱ�� = str����ʱ�� & " And C.ҽ����� Is Not Null And C.����ʱ�� Between [2] And [3] "
        str���� = str���� & " Union All " & str����ʱ��
    End If
    
    str���� = strSqlTmp & str���� & ") A,ҩƷ�շ���¼ C, ������ü�¼ D, ���ű� B " & IIf(mParams.strSourceDep = "", "", ",Table(Cast(f_Num2List([20]) As zlTools.t_NumList)) E ") & _
              " Where C.����id = D.ID And nvl(c.��ҩ��ʽ,-999)<>-1 and A.����=C.���� And A.NO=C.NO And C.����� Is NULL " & _
              " And Nvl(D.����״̬,0)<>1 And (C.�ⷿid=[1] Or C.�ⷿid Is null) And a.�Է�����id = b.Id "
    
    If mstrDeptNode <> "" Then
        str���� = str���� & " And (b.վ�� = [21] Or b.վ�� Is Null) "
    End If
    
    '�ų�������Һ�������Ĺ����в����ĵ���
    str���� = str���� & " And Not Exists (Select 1 From ��Һ��ҩ���� Y Where Y.�շ�id = C.ID) "
    
    '�Ƿ���ʾ��ҩ��������
    If mcondition.bln��ʾ��ҩ�������� = False Then
        str���� = str���� & " And C.��¼״̬=1 "
    Else
        str���� = str���� & " And MOD(C.��¼״̬,3)=1 "
    End If
    
    '��Ժ��ҩ:'0-����,1-������Ժ��ҩ,2-������Ժ��ҩ
    If mcondition.int��Ժ��ҩ = 0 Then
    ElseIf mcondition.int��Ժ��ҩ = 1 Then
        str���� = str���� & " And Not Ltrim(To_Char(Nvl(C.����,0),'00')) Like '_3'"
    ElseIf mcondition.int��Ժ��ҩ = 2 Then
        str���� = str���� & " And Ltrim(To_Char(Nvl(C.����,0),'00')) Like '_3'"
    End If
    
    str���� = str���� & IIf(mParams.strSourceDep = "", "", " And C.�Է�����id+0=E.Column_Value ")
    
    If imgFilter.BorderStyle = cstFilter Then
        If Val(mParams.int����ģʽ) <= 7 Then
            If Trim(txtPati.Text) = "" Then
                str���� = str���� & " And 1 = 2 "
            End If
        Else
            If mSQLCondition.lng����ID = 0 Then
                str���� = str���� & " And 1 = 2 "
            End If
        End If
    End If
    
    If mParams.intType = 1 Then
        str���� = str���� & " And (D.�����־=1 or D.�����־=4)"
    ElseIf mParams.intType = 2 Then
        str���� = str���� & " And (D.�����־<>1 and D.�����־<>4)"
    End If
    
    
    If mcondition.int������� = 1 Then
        '���ﻮ�ۼ��������
        gstrSQL = gstrSQL & str����
    Else
        If mcondition.int������� = 3 Then
            '���ＰסԺ���е���
            strסԺ = Replace(str����, "������ü�¼", "סԺ���ü�¼")
            strסԺ = Replace(strסԺ, "And Nvl(D.����״̬,0)<>1", "")
        Else
            'סԺ����
            strסԺ = Replace(str����, "������ü�¼", "סԺ���ü�¼")
            strסԺ = Replace(strסԺ, "And Nvl(D.����״̬,0)<>1", "")
            str���� = ""
        End If
    
        If mPrives.bln���������� Then
            If img����.BorderStyle = 0 Then
                '����ʾ��������
                strסԺ = strסԺ & " And (D.�����־ <> 2 Or (D.�����־ = 2 And D.���˲���id <> D.��������id)) "
            End If
            If img����.BorderStyle = 1 And cbo����.ListIndex <> -1 Then
                'Ҫ��ʾ�������������Ҳ��˲������ڵ�ǰѡ��Ĳ���
                strסԺ = strסԺ & " And D.���˲���id = " & cbo����.ItemData(cbo����.ListIndex)
                str���� = ""
            End If
        End If
        
        If str���� = "" Then
            gstrSQL = gstrSQL & strסԺ
        Else
            gstrSQL = gstrSQL & str���� & " Union All " & strסԺ
        End If
    End If
    
    '''''Group By
    gstrSQL = gstrSQL & ") A GROUP BY A.���ȼ�,A.����,A.����,A.���շ�,A.��ҩ��,A.No,A.����,A.����,A.�ɲ���,A.˵��," & _
        " A.���￨��,A.�����,A.���֤��,A.IC����,A.����ID,A.ҽ����,A.סԺ��,A.��������,A.�����־,A.��¼����, a.��ҩ����,A.ǩ��ʱ��,a.�������� "
    
    '''''Order By
    gstrSQL = gstrSQL & " Order By A.ǩ��ʱ��,A.����,A.����,A.No"
    
    Set rsData = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, _
            mSQLCondition.lngҩ��ID, _
            mSQLCondition.date��ʼ����, _
            mSQLCondition.date��������, _
            mSQLCondition.str��ʼNO, _
            mSQLCondition.str����NO, _
            UCase(mSQLCondition.str����), _
            mSQLCondition.str���￨, _
            mSQLCondition.str��ʶ��, _
            mSQLCondition.lng����ID, _
            mSQLCondition.str������, _
            mSQLCondition.str�����, _
            mSQLCondition.lngҩƷid, _
            mSQLCondition.str��ǰNO, _
            mSQLCondition.str�����, _
            lng����ID, _
            mSQLCondition.lng����ID, _
            mSQLCondition.strҽ����, _
            mSQLCondition.lngסԺ��, _
            mParams.Str����, _
            mParams.strSourceDep, _
            mstrDeptNode)
    
    stbThis.Panels(2) = ""
    If Not rsData.EOF Then
        stbThis.Panels(2) = "����" & rsData.RecordCount & "�Ŵ�����" & GetSumMoney(rsData)
    End If
    
    Set mrsList = rsData
    If Not mfrmList Is Nothing Then mfrmList.RefreshList mListType.����ҩ, mrsList
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub ShowWindow_Batch()
    '����������ҩ����
    
    With FrmҩƷ������ҩ
        .In_������� = mcondition.int�������
        .In_��ҩ���� = mParams.Str����
        .In_ҩ��ID = mParams.lngҩ��ID
        .In_����� = mParams.IntCheckStock
        .In_У�鴦�� = IIf(mPrives.blnУ�鴦��, 1, 0)
        .In_����δ��ҩ��ҩ = IIf(mParams.blnMustDosageProcess = False, 1, 0)
        .IN_����δ��˷�ҩ = IIf(mParams.bln����δ��˴�����ҩ, 1, 0)
        .IN_����δ�շѷ�ҩ = IIf(mParams.bln����δ�շѴ�����ҩ, 1, 0)
        .In_Ȩ�� = mstrPrivs
        .str��ҩ�� = IIf(tbcDetail.Selected.index = 1, mfrmRecipe.Get��ҩ��, mfrmDetail.Get��ҩ��)
        .str�˲��� = IIf(tbcDetail.Selected.index = 1, mfrmRecipe.Get�˲���, mfrmDetail.Get�˲���)
        .In_����λ�� = mParams.int����λ��
        .IN_��˻��۵� = IIf(mParams.bln��˻��۵�, 1, 0)
        .In_������ҩ������ = False
        .In_���� = mstrOpr
        .In_�Զ���ҩ = mblnPackerConnect
        .In_���÷�ҩ = mblnLoadDrug
        Set .In_DrugMAC = mobjDrugMAC
        Set .In_PlugIn = mobjPlugIn
        .Show 1, Me
    End With
    
    RefreshList mcondition.intListType
End Sub

Private Sub ShowWindow_Charge()
    '���ﻮ��
    Dim strCommon As String, intAtom As Integer, blnOK As Boolean
            
    On Error Resume Next
    If gobjCharge Is Nothing Then
        Set gobjCharge = CreateObject("zl9OutExse.clsOutExse")
        If gobjCharge Is Nothing Then Exit Sub
    End If
    
    err.Clear: On Error GoTo 0
    
    '�������úϷ�������
    strCommon = Format(Now, "yyyyMMddHHmm")
    strCommon = TranPasswd(strCommon) & "||" & AnalyseComputer
    intAtom = GlobalAddAtom(strCommon)
    Call SaveSetting("ZLSOFT", "����ȫ��", "����", intAtom)
    blnOK = gobjCharge.Charge(Me, gcnOracle, glngSys, gstrDbUser, 1, 0)
    Call GlobalDeleteAtom(intAtom)
    
    '��ɻ���
    'ˢ��δ��ҩ����
    RefreshList mcondition.intListType
End Sub

Private Sub ShowWindow_EMR()
    '������ѯ
End Sub

Private Sub ShowWindow_Flag()
    'ֹͣ��ҩ���
    Dim frmFlag As New Frm���ٷ�ҩ������־
    
    frmFlag.In_����� = mParams.IntCheckStock
    frmFlag.gstrParentName = Me.Name
    frmFlag.Show vbModal
    
    RefreshList mcondition.intListType
End Sub

Private Sub ShowWindow_ReturnBatch()
    '������ҩ���Ĵ���
    
    frm������ҩ.In_Ȩ�� = mstrPrivs
    Set frm������ҩ.In_PlugIn = mobjPlugIn
    If Not frm������ҩ.ShowEditor(Me, mParams.lngҩ��ID, True, mParams.int����λ��) Then Exit Sub
    
    RefreshList mcondition.intListType
End Sub

Private Sub ShowWindow_ReturnByBill()
    '��Ʊ�ݺ���ҩ
    
    frm��Ʊ�ݺ�������ҩ.In_Ȩ�� = mstrPrivs
    Set frm��Ʊ�ݺ�������ҩ.In_PlugIn = mobjPlugIn
    If Not frm��Ʊ�ݺ�������ҩ.ShowEditor(Me, mParams.lngҩ��ID, mParams.int����λ��) Then Exit Sub
    
    RefreshList mcondition.intListType
End Sub

Private Sub ShowWindow_SendByBill()
    '��Ʊ�ݺŷ�ҩ
    
    With Frm��Ʊ�ݺ�������ҩ
        .In_������� = mcondition.int�������
        .In_��ҩ���� = mParams.Str����
        .In_ҩ��ID = mParams.lngҩ��ID
        .In_����� = mParams.IntCheckStock
        .In_У�鴦�� = IIf(mPrives.blnУ�鴦��, 1, 0)
        .In_����δ��ҩ��ҩ = IIf(mParams.blnMustDosageProcess = False, 1, 0)
        .IN_����δ��˷�ҩ = IIf(mParams.bln����δ��˴�����ҩ, 1, 0)
        .IN_����δ�շѷ�ҩ = IIf(mParams.bln����δ�շѴ�����ҩ, 1, 0)
        .In_Ȩ�� = mstrPrivs
        .str��ҩ�� = IIf(mParams.str��ҩ�� = "|��ǰ����Ա|", gstrUserName, mParams.str��ҩ��)
        .In_����λ�� = mParams.int����λ��
        .IN_��˻��۵� = IIf(mParams.bln��˻��۵�, 1, 0)
        .In_���� = mstrOpr
        Set .In_DrugMAC = mobjDrugMAC
        Set .In_PlugIn = mobjPlugIn
        .Show 1, Me
    End With
    RefreshList mcondition.intListType
End Sub

Private Sub ShowWindow_SendOther()
    '������ҩ���Ĵ���
    
    With FrmҩƷ������ҩ
        .In_������� = mcondition.int�������
        .In_��ҩ���� = mParams.Str����
        .In_ҩ��ID = mParams.lngҩ��ID
        .In_����� = mParams.IntCheckStock
        .In_У�鴦�� = IIf(mPrives.blnУ�鴦��, 1, 0)
        .In_����δ��ҩ��ҩ = IIf(mParams.blnMustDosageProcess = False, 1, 0)
        .IN_����δ��˷�ҩ = IIf(mParams.bln����δ��˴�����ҩ, 1, 0)
        .IN_����δ�շѷ�ҩ = IIf(mParams.bln����δ�շѴ�����ҩ, 1, 0)
        .In_Ȩ�� = mstrPrivs
        .str��ҩ�� = IIf(tbcDetail.Selected.index = 1, mfrmRecipe.Get��ҩ��, mfrmDetail.Get��ҩ��)
        .str�˲��� = IIf(tbcDetail.Selected.index = 1, mfrmRecipe.Get�˲���, mfrmDetail.Get�˲���)
        .In_����λ�� = mParams.int����λ��
        .IN_��˻��۵� = IIf(mParams.bln��˻��۵�, 1, 0)
        .In_������ҩ������ = True
        .In_�Զ���ҩ = mblnPackerConnect
        .In_���÷�ҩ = mblnLoadDrug
        .Show 1, Me
    End With
    
    RefreshList mcondition.intListType
End Sub

Private Sub ShowWindow_Stuff()
    '���ķ���
    Dim strCommon As String, intAtom As Integer, blnOK As Boolean
    Dim str��ǰ���� As String
    Dim strNo As String
    Dim lng����ID As Long
    Dim rsTmp As ADODB.Recordset
    
    str��ǰ���� = mfrmList.GetCurrentRecipe
    
    If str��ǰ���� <> "" Then
        strNo = Split(str��ǰ����, "|")(1)
        lng����ID = Val(Split(str��ǰ����, "|")(3))
    End If
    
    On Error Resume Next
    If gobjStuff Is Nothing Then
        Set gobjStuff = CreateObject("zl9Stuff.clsStuff")
        If gobjStuff Is Nothing Then Exit Sub
    End If

    err.Clear: On Error GoTo 0

    '�������úϷ�������
    strCommon = Format(Now, "yyyyMMddHHmm")
    strCommon = TranPasswd(strCommon) & "||" & AnalyseComputer
    intAtom = GlobalAddAtom(strCommon)
    Call SaveSetting("ZLSOFT", "����ȫ��", "����", intAtom)
    Call gobjStuff.TransStuff(Me, gcnOracle, glngSys, gstrDbUser, lng����ID, strNo, mParams.lngҩ��ID, Format(mSQLCondition.date��ʼ����, "yyyy-mm-dd hh:mm:ss"), Format(mSQLCondition.date��������, "yyyy-mm-dd hh:mm:ss"))
    Call GlobalDeleteAtom(intAtom)
End Sub

Private Sub cbo����_Click()
    If cbo����.ListIndex = -1 Then Exit Sub
    If cbo����.Enabled = False Then Exit Sub
    
    If cbo����.ItemData(cbo����.ListIndex) <> Val(cbo����.Tag) Then
        cbo����.Tag = cbo����.ItemData(cbo����.ListIndex)
        Call RefreshList(mcondition.intListType)
    End If
End Sub


Private Sub cboʱ�䷶Χ_Click()
    With cboʱ�䷶Χ
        If .ListIndex <> Val(.Tag) Then
            If (Val(.Tag) = 3 And .ListIndex < 3) Or (Val(.Tag) < 3 And .ListIndex = 3) Then
                Call picConMain_Resize
                Call picCondition_Resize
            End If
            .Tag = .ListIndex
        End If
        
        If .ListIndex < mTimeRange.ָ��ʱ�䷶Χ And mblnStart = True Then
            RefreshList mcondition.intListType
        End If
    End With
End Sub


Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim cbrControl As CommandBarControl
    Dim strReturn As String
    
    Select Case Control.Id
        '''''�ļ�
        Case mconMenu_File_PrintSet     '��ӡ����
            zlPrintSet
        Case mconMenu_File_Preview      '��ӡԤ��
            zlSubPrint 2
        Case mconMenu_File_Print        '��ӡ
            zlSubPrint 1
        Case mconMenu_File_Excel        '�����Excel
            zlSubPrint 3
        
        Case mconMenu_File_Recipe_BillPrintDosage       '��ӡ��ҩ��
            Call BillPrint_Dosage
        Case mconMenu_File_Recipe_BillPrintRecipe       '��ӡ����ǩ
            Call BillPrint_Recipe
        Case mconMenu_File_Recipe_BillPrintReport       '��ӡ��ҩ�嵥
            Call BillPrint_Report
        Case mconMenu_File_Recipe_BillPrintReturn       '��ӡ��ҩ֪ͨ��
            Call BillPrint_Return
        Case mconMenu_File_Recipe_BillPrintLable        '��ӡҩƷ��ǩ
            Call BillPrint_Lable
        Case mconMenu_File_Recipe_BillPrintBack         '��ӡ�˷ѵ���
            Call BillPrint_Back
        Case mconMenu_File_Recipe_BillPrintChange        '��ӡҽ������֪ͨ��
            Call BillPrint_Change
        Case mconMenu_File_Parameter                    '��������
            ResetParams
            
        Case mconMenu_File_Exit                         '�˳�
            Unload Me
        
        '''''�༭
        Case mconMenu_Edit_Recipe_Batch                 '������ҩ(&B)
            Call ShowWindow_Batch
        Case mconMenu_Edit_Recipe_SendOther             '������ҩ���Ĵ���(&F)
            Call ShowWindow_SendOther
        Case mconMenu_Edit_Recipe_ReturnBatch           '������ҩ���Ĵ���(&T)
            Call ShowWindow_ReturnBatch
        Case mconMenu_Edit_Recipe_SendByBill            '��Ʊ�ݺŷ�ҩ(&I)
            Call ShowWindow_SendByBill
        Case mconMenu_Edit_Recipe_ReturnByBill          '��Ʊ�ݺ���ҩ(&R)
            Call ShowWindow_ReturnByBill
        Case mconMenu_Edit_Recipe_Flag                  'ֹͣ��ҩ���(&S)
            Call ShowWindow_Flag
        Case mconMenu_Edit_Recipe_Charge                '���ﻮ��(&M)-F8
            Call ShowWindow_Charge
        Case mconMenu_Edit_Recipe_Stuff                 '���ķ���(@W)-F9
            Call ShowWindow_Stuff
        Case mconMenu_Edit_Recipe_TakeDrug              'ȡҩȷ��(&T)
            Call RecipeWork_TakeDrug
        Case mconMenu_Edit_Recipe_Call                  '����
            Call RecipeWork_Call
        Case mconMenu_Edit_Recipe_Cancle                'ȡ��ȷ��
            If Control.Caption = "ȡ��ȷ��" Then
                Call RecipeWork_DosageOk
            Else
                Call RecipeWork_Abolish
            End If
            
            Call RefreshList(mcondition.intListType)
        Case mconMenu_Edit_PlugIn + 1 To mconMenu_Edit_PlugIn + 99 '��ҷ�ҩҵ���ܵ���
            DrugSendRecipeNormal Control.Parameter
        Case mconMenu_Edit_Recipe_Change                '�л���ҩ��(&E)
            Call ChangeDosagePeople
        Case mconMenu_Edit_Recipe_Windows               '��������
            Call ChangWin
        Case mconMenu_Edit_Recipe_EMR                   '������ѯ
            Call ShowWindow_EMR
        Case mconMenu_Edit_Recipe_SendHot               '��ҩ��ݼ�����-F2
            If tbcDetail.Selected.index = 0 Then
                mfrmDetail.CmdProcess
            ElseIf tbcDetail.Item(1).Visible = True Then
                mfrmRecipe.CmdProcess
            End If
        
        '''''�鿴
        Case mconMenu_View_ToolBar_Button               '��׼��ť
            Control.Checked = Not Control.Checked
            Me.cbsMain(2).Visible = Control.Checked
            Me.cbsMain.RecalcLayout
        Case mconMenu_View_ToolBar_Text                 '�ı���ǩ
            Control.Checked = Not Control.Checked
            For Each cbrControl In Me.cbsMain(2).Controls
                cbrControl.Style = IIf(cbrControl.Style = xtpButtonIcon, xtpButtonIconAndCaption, xtpButtonIcon)
            Next
            Me.cbsMain.RecalcLayout
        Case mconMenu_View_ToolBar_Size                 '��ͼ��
            Control.Checked = Not Control.Checked
            Me.cbsMain.Options.LargeIcons = Not Me.cbsMain.Options.LargeIcons
            Me.cbsMain.RecalcLayout
        Case mconMenu_View_StatusBar                    '״̬��
            Me.stbThis.Visible = Not Me.stbThis.Visible
            Me.cbsMain.RecalcLayout
        Case mconMenu_View_FontSize_1, mconMenu_View_FontSize_2, mconMenu_View_FontSize_3                   '�ֺ�����
            mParams.intFont = Val(Control.Parameter)
            Call SetFontSize
            Call zldatabase.SetPara("����", mParams.intFont, glngSys, 1341)
        
        Case mconMenu_View_Filter                       '���ݹ���
            Call ResetFilter
        Case mconMenu_View_Refresh                      'ˢ��
            Call RefreshList(mcondition.intListType)
        Case mconMenu_Edit_Recipe_VerifySign            '��֤����ǩ��
            VerifySign
        Case mconMenu_Edit_Recipe_AutoSend_Open
            '���ô����ϴ�
            Control.Checked = Not Control.Checked
            mblnLoadDrug = Control.Checked
        Case mconMenu_Edit_Recipe_AutoSend_Set
            mblnPackerConnect = mobjDrugMAC.DYEY_MZ_SetServer
            SetComandBars
        Case mconMenu_Edit_Recipe_AutoSend_LoadDrug
            Call mobjDrugMAC.DYEY_MZ_TransDrug(1, UserInfo.�û�����, UserInfo.�û�����, strReturn)
        Case mconMenu_Edit_Recipe_AutoSend_LoadStock
            Call mobjDrugMAC.DYEY_MZ_TransStock(Val(mstrOpr), UserInfo.�û�����, UserInfo.�û�����, mParams.lngҩ��ID, strReturn)
            
        '''''����
        Case mconMenu_Help_Help                         '����
'            Call ShowHelp(App.ProductName, Me.hWnd, Me.Name)
            Call ShowHelp(App.ProductName, Me.hWnd, "FrmҩƷ��ҩ����")
        Case mconMenu_Help_Web                          'WEB�ϵ�����
        Case mconMenu_Help_Web_Home                     '������ҳ
            Call zlHomePage(Me.hWnd)
        Case mconMenu_Help_Web_Forum                    '������̳
            Call zlWebForum(Me.hWnd)
        Case mconMenu_Help_Web_Mail                     '���ͷ���
            Call zlMailTo(Me.hWnd)
        Case mconMenu_Help_About                        '����
            Call ShowAbout(Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
        Case mconMenu_Edit_Recipe_MedicalRecord         '���Ӳ�������
            Call ShowMedicalRecord(mfrmDetail.GetRecord)
     
        ''''�����ȼ�
'        Case mconMenu_Edit_Recipe_Hot_IC
'            If mParams.int����ģʽ = mFindType.IC�� Then
'                Call cmdIC_Click
'            End If
            
        Case Else
            If Control.Id > 401 And Control.Id < 499 Then
                'ִ���Զ��屨��
                Call BillPrint_Custom(Control)
            End If
            
            '�����˵�
'            If Control.Id >= mconMenu_Input_Recipe_NO And Control.Id <= mconMenu_Input_Recipe_NO + 6 + mintCardCount Then
'                Call SetInputPopupCheck(Control)            '������Ŀ�����˵�
'            End If
            
'            'ҩ���Զ���ҩ�ӿڲ˵�
'            If Control.Id > mconMenu_AutoSend And Control.Id < mconMenu_AutoSend + 10 Then
'                gobjPackerMZ.SetInterface Control.Id - mconMenu_AutoSend - 1, mParams.lngҩ��ID
'            End If
    End Select
End Sub

Private Sub DrugSendRecipeNormal(ByVal strFunName As String)
    Dim str��ǰ���� As String, Int���� As Integer, strNo As String
    
    If Not mobjPlugIn Is Nothing Then
        str��ǰ���� = mfrmList.GetCurrentRecipe
        
        If str��ǰ���� <> "" Then
            Int���� = Val(Split(str��ǰ����, "|")(0))
            strNo = Split(str��ǰ����, "|")(1)
        End If
        
        On Error Resume Next
        Call mobjPlugIn.DrugSendWorkNormal(glngModul, strFunName, mParams.lngҩ��ID, strNo, Int����)
        err.Clear: On Error GoTo 0
    End If
    
End Sub

Private Function RecipeWork_Call() As Boolean
    '����
    Dim str��ǰ���� As String
    Dim Int���� As Integer
    Dim strNo As String
    Dim Str���� As String
    Dim strName As String
    Dim strCall As String
    Dim strMsg As String
    
    On Error GoTo ErrHand
    
    str��ǰ���� = mfrmList.GetCurrentRecipe
    
    If str��ǰ���� <> "" Then
        Int���� = Val(Split(str��ǰ����, "|")(0))
        strNo = Split(str��ǰ����, "|")(1)
        strName = Split(str��ǰ����, "|")(8)
        Str���� = Split(str��ǰ����, "|")(9)
        strMsg = Split(str��ǰ����, "|")(10)
    End If
     
    Call mfrmList.SetCalling
    
    strCall = "�롢" & strName & "��" & strName & "��" & "����" & mstr����
        
    gstrSQL = "Zl_δ��ҩƷ��¼_����("
            'NO
            gstrSQL = gstrSQL & "'" & strNo & "'"
            '����
            gstrSQL = gstrSQL & "," & Int����
            'ҩ��id
            gstrSQL = gstrSQL & "," & mParams.lngҩ��ID
            '��ҩ����
            gstrSQL = gstrSQL & ",'" & Str���� & "'"
            '��������
            gstrSQL = gstrSQL & ",'" & strCall & "'"
            gstrSQL = gstrSQL & ")"
            
    Call zldatabase.ExecuteProcedure(gstrSQL, "RecipeWork_Call")
    
    'ˢ����ʾ�������
    If mParams.blnShowQueue = True Then
        If Not gobjLEDShow Is Nothing Then
            Call gobjLEDShow.zlDrugShow(mParams.lngҩ��ID, mParams.Str����, mParams.blnMustDosageProcess, mParams.blnMustDosageOkProcess, strName)
        End If
    End If
    
    '��������˱�������ϵͳ����������
    If mParams.blnStartQueue = True Then
        If mParams.blnStartCall = True And mParams.intCallType = 0 And mQueue.blnCallOver = True Then
            Call zlCallMain
        End If
    End If
    
    RecipeWork_Call = True
    
    '����ͬʱ֪ͨ�豸׼����ҩ
    If mParams.blnDispensing Then
        Call DrugDispensing("" & Int���� & "," & strNo)
    End If
    
    Exit Function
    
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub DrugDispensing(ByVal strNos As String)
'���ܣ�֪ͨ�豸׼����ҩ

    Dim strReturn As String
    
    On Error GoTo hErr
    
    If UCase(TypeName(mobjDrugMAC)) = UCase("clsDrugPacker") Then
        If mcondition.intListType = mListType.����ҩ And mblnPackerConnect And mintAutoSendFlow = 1 _
            And strNos <> "" And mblnCompatible Then
            If mobjDrugMAC.DYEY_MZ_TransRecipeList(mstrOpr, UserInfo.�û�����, UserInfo.�û�����, _
                mParams.lngҩ��ID, strNos, strReturn, mSendOper.StartSend) = False Then
                Call MsgBox("ҩƷ�Զ����豸ϵͳδ׼���ã�֪ͨ��ҩ��ʼʧ�ܣ�", vbInformation, gstrSysName)
            End If
        End If
    ElseIf UCase(TypeName(mobjDrugMAC)) = UCase("clsDrugMachine") Then
        If mcondition.intListType = mListType.����ҩ And mblnPackerConnect Then
            Call mobjDrugMAC.Operation(gstrDbUser, Val("22-��ʼ��ҩ"), "1|" & Replace(strNos, "|", ";"), strReturn)
        End If
    End If
    
    Exit Sub
    
hErr:
    Call ErrCenter
End Sub

Private Sub SetFontSize()
    Dim intFont As Integer
    Dim stdfnt As StdFont
    
    Select Case mParams.intFont
        Case 0
            intFont = 9
        Case 1
            intFont = 11
        Case 2
            intFont = 15
        Case Else
            intFont = 9
    End Select
    
    mfrmList.SetFontSize intFont
    mfrmDetail.SetFontSize intFont
    
    If Not tbcList.PaintManager.Font Is Nothing Then
        With tbcList
            Set stdfnt = .PaintManager.Font
            stdfnt.Size = intFont
             Set .PaintManager.Font = stdfnt
              .PaintManager.Layout = xtpTabLayoutAutoSize
        End With
    End If
    
    If Not tbcDetail.PaintManager.Font Is Nothing Then
        With tbcDetail
            Set stdfnt = .PaintManager.Font
            stdfnt.Size = intFont
             Set .PaintManager.Font = stdfnt
              .PaintManager.Layout = xtpTabLayoutAutoSize
        End With
    End If
    Me.FontSize = intFont
End Sub
Private Sub zlSubPrint(ByVal bytMode As Byte)
    'bytMode��1-��ӡ��2-Ԥ����3-�����Excel
    Dim ObjThis As Object
    Dim objPrint As New zlPrint1Grd
    Dim ObjAppRow As New zlTabAppRow
    Dim strTitle As String
    
    'ȡ��ӡ�б����
    Set ObjThis = mfrmList.GetPrintObject(True)
    
    If ObjThis Is Nothing Then
        mfrmList.GetPrintObject False
        Exit Sub
    End If
    
    Select Case tbcList.Selected.index
        Case mListType.����ҩ
            strTitle = "ҩƷ����ҩ�嵥"
        Case mListType.����ҩ
            strTitle = "ҩƷ����ҩ�嵥"
        Case mListType.����ҩ, mListType.��ʱδ��
            strTitle = "ҩƷ����ҩ�嵥"
        Case mListType.��ҩ
            strTitle = "ҩƷ��ҩ�嵥"
    End Select
    
    Set ObjAppRow = New zlTabAppRow
    ObjAppRow.Add "��ӡ��:" & gstrUserName
    ObjAppRow.Add "��ӡ����:" & Format(Sys.Currentdate, "yyyy-MM-dd")
    objPrint.BelowAppRows.Add ObjAppRow
    
    Set ObjAppRow = New zlTabAppRow
    ObjAppRow.Add "��ʼʱ��:" & Format(Dtp��ʼʱ��.Value, "yyyy-MM-dd HH:mm:ss")
    ObjAppRow.Add "����ʱ��:" & Format(Dtp����ʱ��.Value, "yyyy-MM-dd HH:mm:ss")
    objPrint.UnderAppRows.Add ObjAppRow
    
    objPrint.Title.Text = strTitle
    Set objPrint.Body = ObjThis
    
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
    
    mfrmList.GetPrintObject False
End Sub
Private Sub cbsMain_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If Me.stbThis.Visible Then Bottom = Me.stbThis.Height
End Sub


Private Sub cbsMain_Resize()
    Dim lngLeft As Long, lngTop  As Long, lngRight  As Long, lngBottom  As Long
    
    On Error Resume Next
    
    Call Me.cbsMain.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)

    Me.picDetail.Move lngLeft, lngTop, lngRight - lngLeft, lngBottom - lngTop
End Sub


Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.Id
        Case mconMenu_View_StatusBar '״̬��
            Control.Checked = Me.stbThis.Visible
        Case mconMenu_View_FontSize_1, mconMenu_View_FontSize_2, mconMenu_View_FontSize_3       '����
            Control.Checked = Val(Control.Parameter) = mParams.intFont
        Case mconMenu_Edit_Recipe_MedicalRecord
            Control.Enabled = mfrmDetail.CmdSend.Enabled
     End Select
End Sub
Private Sub Chk��ʾ���̵���_Click()
    Call SaveSetting("ZLSOFT", "����ģ��\����\" & App.ProductName & "\" & "ҩƷ������ҩ", "��ʾ��ҩ���̵���", Chk��ʾ���̵���.Value)
    RefreshList mcondition.intListType
End Sub

Private Sub Chk��ʾ��ҩ��������_Click()
    RefreshList mcondition.intListType
End Sub

Private Sub cmdFind_Click()
    Call Form_KeyDown(vbKeyF3, 0)
End Sub

'Private Sub cmdIC_Click()
'    Dim strOutXML As String
'    Dim strText As String
'
'    If Val(lblPati.Tag) = mFindType.IC�� Then
'        If mobjICCard Is Nothing Then
'            Set mobjICCard = CreateObject("zlICCard.clsICCard")
'            Set mobjICCard.gcnOracle = gcnOracle
'        End If
'        If Not mobjICCard Is Nothing Then
'            txtPati.Text = mobjICCard.Read_Card()
'            If txtPati.Text <> "" Then Call txtPati_KeyPress(vbKeyReturn)
'        End If
'    Else
'        If Not mobjSquareCard Is Nothing Then
'            Call mobjSquareCard.zlReadCard(Me, mlngMode, Val(Split(txtPati.Tag, "|")(gCardFormat.�����ID)), True, "", strText, strOutXML)
'            txtPati.Text = strText
'            If txtPati.Text <> "" Then Call txtPati_KeyPress(vbKeyReturn)
'        End If
'    End If
'End Sub

Private Sub dkpMain_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.Id
        Case 1
            Item.Handle = picCondition.hWnd
        Case 2
            Item.Handle = picList.hWnd
    End Select
End Sub

Private Sub Form_Activate()
'    If mblnStart = False Then
'        Unload Me
'        Exit Sub
'    End If

    Call picConMain_Resize
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim blnFirst As Boolean
    Dim strInput As String
    Dim strNos As String
    Dim strReturn As String
    Dim strCard As String
    
    If KeyCode = vbKeyF3 Then
        If imgFilter.BorderStyle = cstLocate Then
            If txtPati.Text = "" Then
                txtPati.SetFocus
            Else
                Call txtPati_Validate(False)
                Call zlControl.TxtSelAll(txtPati)
                strCard = IDKNType.GetCurCard.����
                If strCard = "IC��" Then
                    If Not mobjSquareCard Is Nothing Then Call mobjSquareCard.zlGetPatiID("IC��", UCase(Trim(txtPati.Text)), False, mlngIC����id)
                    strInput = mlngIC����id
                
                ElseIf strCard = "����" Or strCard = "���ݺ�" Or strCard = "סԺ��" Or strCard = "ҽ����" Or strCard = "���֤" Or strCard = "�����" Then
                    
                    strInput = txtPati.Text
                Else
                    '���ѿ����ʱ����Ϊ��ID+����
                    strInput = mobjcard.�ӿ���� & "|" & txtPati.Text
                End If
                If mfrmList.FindSpecialRow(IDKNType.GetCurCard.����, strInput, strNos, mobjSquareCard) = True Then
                    mblnFinding = True
                    If mcondition.intListType = mListType.����ҩ And mParams.int����ģʽ = mFindType.���ݺ� And mParams.bln��ҩɨ�� = True Then
                        '��ҩģʽ����ɨ����ʱ����
                        If mblnScaned = False Then
                            '��һ��ɨ��
                            mblnScaned = True
                        Else
                            '�ڶ���ɨ�裬ȷ����ҩ
                            mblnScaned = False
                            mstrScanerLastNo = ""
                            If tbcDetail.Selected.index = 0 Then
                                mfrmDetail.CmdProcess
                            ElseIf tbcDetail.Item(1).Visible = True Then
                                mfrmRecipe.CmdProcess
                            End If
                        End If
                        txtPati.SetFocus
                        txtPati.Text = ""
                    ElseIf mcondition.intListType = mListType.����ҩ And mbln��������ˢ�� = True Then
                        '����ˢ����ҩģʽ
                         
                        If mblnBrushCard = False Then
                            '��һ��ˢ��
                            mblnBrushCard = True
                        Else
                            If txtPati.Text = mstrLastBrushCardNo Then
                                '�ڶ���ˢ����ȷ�Ϸ�ҩ
                                mblnBrushCard = False
                                mstrLastBrushCardNo = ""
                                If tbcDetail.Selected.index = 0 Then
                                    mfrmDetail.CmdProcess
                                ElseIf tbcDetail.Item(1).Visible = True Then
                                    mfrmRecipe.CmdProcess
                                End If
                            Else
                                '����ˢ��ͬһ�ſ�
                                mblnBrushCard = True
                                mstrLastBrushCardNo = txtPati.Text
                            End If
                        End If
                        txtPati.SetFocus
                        txtPati.Text = ""
                    Else
                        If tbcDetail.Selected.index = 0 Then
                            If mfrmDetail.CmdSend.Enabled Then mfrmDetail.CmdSend.SetFocus
                        ElseIf tbcDetail.Item(1).Visible = True Then
                            If mfrmRecipe.CmdSend.Enabled Then mfrmRecipe.CmdSend.SetFocus
                        End If
                    End If
                    
                    '����Զ���ҩ�п�ʼ��ҩ���̣�����ýӿ��ϴ�����
                    '����б����и��в��˵Ķ��������һ���ϴ�
                    '�����ݽӿ�ʱû���������
                    If TypeName(mobjDrugMAC) = "clsDrugPacker" Then
                        If mcondition.intListType = mListType.����ҩ And mblnPackerConnect And mblnLoadDrug And mintAutoSendFlow = 1 And strNos <> "" Then
                            If mblnCompatible = True Then
                                If mobjDrugMAC.DYEY_MZ_TransRecipeList(mstrOpr, UserInfo.�û�����, UserInfo.�û�����, mParams.lngҩ��ID, strNos, strReturn, mSendOper.StartSend) = False Then
                                    If MsgBox("�Զ���ҩϵͳδ׼���ã��Ƿ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                                        Exit Sub
                                    End If
                                End If
                            End If
                        End If
                    ElseIf TypeName(mobjDrugMAC) = "clsDrugMachine" Then
                        If mcondition.intListType = mListType.����ҩ And mblnPackerConnect Then
                            mobjDrugMAC.Operation gstrDbUser, Val("22-��ʼ��ҩ"), "1|" & Replace(strNos, "|", ";"), strReturn
'                           If strReturn <> "" Then MsgBox strReturn, vbInformation, gstrSysName
                        End If
                    End If
                 Else
                    'û���ҵ�ʱ
                    If mcondition.intListType = mListType.����ҩ And mParams.int����ģʽ = mFindType.���ݺ� And mParams.bln��ҩɨ�� = True Then
                        mblnScaned = False
                        mstrScanerLastNo = ""
                        txtPati.SetFocus
                        txtPati.Text = ""
                    End If
                    mblnBrushCard = False
                    mstrLastBrushCardNo = ""
                End If
            End If
        Else
'            Call SetFilter(MnuEditHandback.Checked)
            Me.IDKNType.ActiveFastKey
            RefreshList mcondition.intListType
        End If
    End If
    
    If KeyCode = 70 Or KeyCode = 102 Then
        If Shift = vbCtrlMask Then   'Ctrl+F
            txtPati.SetFocus
        End If
    End If
    
    'Ctrl+F4  ��IC��
'    If KeyCode = vbKeyF4 Or KeyCode = 102 Then
'        If Shift = vbCtrlMask Then
'            If cmdIC.Visible = True Then
'                Call cmdIC_Click
'            End If
'        End If
'    End If
End Sub

Private Sub Form_Load()
    Dim dteTime As Date
    Dim strMessage As String, strPrivs As String
   
    mblnStart = False
    mblnSendIsOver = True
    mlngMode = glngModul
    mstrPrivs = gstrprivs
    
    mQueue.strPCName = AnalyseComputer
    
    Me.Width = mcstlngWinNormalWidth
    Me.Height = mcstlngWinNormalHeight
    
    picConMain.BackColor = &H80000005
    lblʱ�䷶Χ.BackColor = picConMain.BackColor
    lblTimeBegin.BackColor = picConMain.BackColor
    lblTimeEnd.BackColor = picConMain.BackColor
    IDKNType.BackColor = picConMain.BackColor
    lbl����.BackColor = picConMain.BackColor
    Chk��ʾ���̵���.BackColor = picConMain.BackColor
    chk��ʾ��ȷ�ϵ���.BackColor = picConMain.BackColor
    Chk��ʾ��ҩ��������.BackColor = picConMain.BackColor
    
    mdate�ϴ�У��ʱ�� = Sys.Currentdate
    mstr�Զ���ҩ�� = ""
    
    intģʽ = 1
    Set mclsComLib = New zl9ComLib.clsComLib
    
    mstrChargePrivs = GetPrivFunc(glngSys, 1120)
    mstrStuffPrivs = GetPrivFunc(glngSys, 1723)
    
    If gstrUserName = "" Then
        MsgBox "��Ϊ��ǰ�û����ö�Ӧ�Ĳ���Ա����ʹ�ñ�ģ�飡", vbInformation, gstrSysName
        Exit Sub
    End If
     
    'ȡ���ý��λ�������ڽ�����ʾ
    mintMoneyDigit = gtype_UserSysParms.P9_���ý���λ��
    '���ý���ʽ
    Call GetMoneyFormat
    
    'ȡȨ��
    Call GetPrivs
    
    '�������ݼ��
    If DependOnCheck = False Then Exit Sub
    
    'ȡ����
    Call GetParams
    
    With mParams
        'ע������
        .int���涨λ = Val(GetSetting("ZLSOFT", "����ģ��\����\" & App.ProductName & "\" & "ҩƷ������ҩ", "���涨λ", cstLocate))
        .int�������� = GetSetting("ZLSOFT", "����ģ��\����\" & App.ProductName & "\" & "ҩƷ������ҩ", "��ʾ��ҩ��������", 1)
        .int���̵��� = GetSetting("ZLSOFT", "����ģ��\����\" & App.ProductName & "\" & "ҩƷ������ҩ", "��ʾ��ҩ���̵���", 1)
        .int��ȷ�ϵ��� = GetSetting("ZLSOFT", "����ģ��\����\" & App.ProductName & "\" & "ҩƷ������ҩ", "��ʾ��ȷ�ϵ���", 1)
        .int����ģʽ���� = Val(GetSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\" & "ҩƷ������ҩ", "����ģʽ", "1"))
        If .int����ģʽ���� < 1 Then
            .int����ģʽ���� = 1
        End If
    End With
    
    Call GetOpr
    
    Call SetFontSize
    
    '����������
    If CheckAnother = False Then Exit Sub
    
    Call Loadʱ�䷶Χ
    
    If Not mPrives.bln�����ѯ����ʱ�䷶Χ���� Then
        cboʱ�䷶Χ.ListIndex = 0
        cboʱ�䷶Χ.Tag = 0
        cboʱ�䷶Χ.Enabled = False
    End If
    
    Call Load����
    
    Call GetDrugStock(mParams.lngҩ��ID)
    Call GetDosage(mParams.lngҩ��ID)
    Call GetSendWindows(mParams.lngҩ��ID)
    
    '�������Ӳ������Ķ���
    If mobjCISJOB Is Nothing Then
        On Error Resume Next
        Set mobjCISJOB = CreateObject("zl9CISJob.clsCISJob")
        
        If Not mobjCISJOB Is Nothing Then
            Call mobjCISJOB.InitCISJob(gcnOracle, Me, glngSys, mstrPrivs, gobjBrower.mobjEmr)
        End If
        err.Clear: On Error GoTo 0
    End If
    
    '��ʼ������
    dteTime = Sys.Currentdate
    Dtp��ʼʱ��.Value = Format(dteTime, "yyyy-MM-dd 00:00:00")
    Dtp����ʱ��.Value = Format(dteTime, "yyyy-MM-dd") & " 23:59:59"
    
    GetStockName mParams.lngҩ��ID
    
    '���˿���
    imgFilter.BorderStyle = mParams.int���涨λ
    If imgFilter.BorderStyle = 0 Then
        imgFilter.ToolTipText = "����л�������ģʽ"
    Else
        imgFilter.ToolTipText = "����л�����λģʽ"
    End If
    
    '�������˿��أ�Ĭ����0-����ʾ
    img����.BorderStyle = mParams.int��ʾ��������
    
    cbo����.Enabled = (img����.BorderStyle = 1)
    
    If mPrives.bln���������� = False Then
        lbl����.Visible = False
        img����.Visible = False
        cbo����.Visible = False
    End If
    
    '����ʱ��ؼ�
    TimeRefresh.Enabled = False
    TimePrint.Enabled = False
    If mParams.lngRefreshInterval > 0 Then
        If mParams.lngRefreshInterval > 60 Then
            mParams.lngRefreshInterval = 60
        End If
        With TimeRefresh
            .Enabled = True
            .Interval = mParams.lngRefreshInterval * 1000
        End With
    End If
    
    If mParams.lngPrintInterval > 0 Then
        If mParams.lngPrintInterval > 60 Then
            mParams.lngPrintInterval = 60
        End If
        With TimePrint
            .Enabled = True
            .Interval = mParams.lngPrintInterval * 1000
        End With
    End If
    IntTimes = 0
    If mParams.lngPrintBackInterval <> 0 Then
        With TimePrintCancelBill
            .Enabled = False
            .Enabled = True
        End With
    Else
        TimePrintCancelBill.Enabled = False
    End If
    
    '�жϱ����Ƿ���Զ�̺��л���
    mQueue.blnRemoteCall = False
    If mParams.intCallType = 0 And mParams.strRemoteCall = mQueue.strPCName And mQueue.strPCName <> "" Then
        mQueue.blnRemoteCall = True
    End If
    
    '���ýк���ѯʱ����������������������ȫ��Զ�˻�����ʱ
    tmrCall.Enabled = False
    If mParams.blnStartQueue = True And mParams.blnStartCall = True And mQueue.blnRemoteCall = True Then
        tmrCall.Enabled = True
        tmrCall.Interval = mParams.intCircleTime * 1000
    End If
    
    Set mobjIDCard = New clsIDCard
    Set mobjICCard = New clsICCard
    Call mobjIDCard.SetParent(Me.hWnd)
    Call mobjICCard.SetParent(Me.hWnd)
    Set mobjICCard.gcnOracle = gcnOracle
    
    '����ǩ���ӿڿ���
    gblnESign������ҩ = EsignIsOpen(mParams.lngҩ��ID)
    gblnESignUserStoped = False
    If gblnESign������ҩ = True Then
        On Error Resume Next
        Set gobjESign = CreateObject("zl9ESign.clsESign")
        err.Clear: On Error GoTo 0
        If Not gobjESign Is Nothing Then
            If Not gobjESign.Initialize(gcnOracle, glngSys) Then
                Set gobjESign = Nothing
                gblnESign������ҩ = False
            Else
                gblnESign������ҩ = True
                gblnESignUserStoped = gobjESign.CertificateStoped(gstrUserName)
            End If
        Else
            gblnESign������ҩ = False
        End If
    End If
    
    'һ��ͨ�ӿ�
    mstrCardType = zlfuncCard_Ini(mobjSquareCard, Me, mlngMode)
    
    '�Զ���ҩ���ӿ�
    mblnPackerConnect = False
    On Error Resume Next
    
    '���ҩƷ�Զ����ӿ�Ȩ�޺Ͳ���
    If Val(zldatabase.GetPara("����ҩƷ�Զ����豸�ӿ�", glngSys, Val("9010-ҩƷ�Զ����豸�ӿ�"))) = 1 _
        And mPrives.blnҩƷ�Զ����ӿ� = True Then
        
        Set mobjDrugMAC = Nothing
        '�����½ӿ�
        Set mobjDrugMAC = CreateObject("zlDrugMachine.clsDrugMachine")
        If err.Number <> 0 Then
            '��ξɽӿ�
            Set mobjDrugMAC = CreateObject("zlDrugPacker.clsDrugPacker")
        End If
    Else
        Set mobjDrugMAC = CreateObject("zlDrugPacker.clsDrugPacker")
    End If
    On Error GoTo 0
    
    If TypeName(mobjDrugMAC) = "clsDrugMachine" Then
        '�½ӿ�
        ''��ȡ�ӿڵ�Ȩ��
        strPrivs = ";" & zl9ComLib.GetPrivFunc(glngSys, Val("9010-ҩƷ�Զ����豸�ӿ�")) & ";"
        If strPrivs Like "*;����;*" Then
            mblnPackerConnect = mobjDrugMAC.Init(1, mclsComLib, strMessage)
        Else
            mblnPackerConnect = False
        End If
    ElseIf TypeName(mobjDrugMAC) = "clsDrugPacker" Then
        '�ɽӿ�
        mblnPackerConnect = mobjDrugMAC.DYEY_MZ_IniSoap(, , gstrUnitName)
        
        On Error Resume Next
        mintAutoSendFlow = mobjDrugMAC.DYEY_MZ_GetSendType      '���־ɽӿ��޸÷���
        mblnCompatible = (err.Number = 0)
        On Error GoTo 0
    Else
        mblnPackerConnect = False
        mintAutoSendFlow = False
    End If
    
    '��ҩҵ����Ҳ���
    Call zlPlugIn_Ini(glngSys, glngModul, mobjPlugIn)
    
    '���ò˵�
    Call InitComandBars
    Call InitPanes
    Call InitTabControl
    Call InitIDKindNew
        
    Chk��ʾ��ҩ��������.Value = IIf(mParams.int�������� = 1, 1, 0)
    Chk��ʾ���̵���.Value = IIf(mParams.int���̵��� = 1, 1, 0)
    chk��ʾ��ȷ�ϵ���.Value = IIf(mParams.int��ȷ�ϵ��� = 1, 1, 0)
    
'    ����Զ��屨��
    Call zldatabase.ShowReportMenu(Me, glngSys, glngModul, gstrprivs)
    
    '�ָ�¼��״̬
'    Call SetInputState(mParams.int����ģʽ)
    
    '�ָ�����
    If Val(zldatabase.GetPara("ʹ�ø��Ի����")) = 1 Then
        On Error Resume Next
        
        dkpMain.LoadStateFromString GetSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\��������\" & App.ProductName & "\" & Me.Name & "\" & TypeName(dkpMain), dkpMain.Name & dkpMain.PanesCount, "")
        SetPaneTitle tbcList.Selected.index
    End If
    Call RestoreWinState(Me, App.ProductName)
    
    '���Ŷ���ʾ����
    Call ShowQueue

    Call zlCall_SystemSoundPlay("", 65)
    
    mQueue.blnCallOver = True
    
'    '����ҩ���Զ���ҩ
'    If gtype_UserSysParms.P222_ҩ���Զ�����ҩ�ӿ� = 1 Then
'        err = 0
'        On Error Resume Next
'
'        If gobjPackerMZ Is Nothing Then
'            Set gobjPackerMZ = CreateObject("zlDrugPacker.clsDrugPacker")
'            err.Clear
'
'            If Not gobjPackerMZ Is Nothing Then
'                gobjPackerMZ.InitCommon gcnOracle, Me, glngSys, mlngMode, mParams.lngҩ��ID
'            End If
'        End If
'    End If

 
    '��ʼ����Ϣ����
    err = 0
    On Error Resume Next
    Set mobjMipModule = New zl9ComLib.clsMipModule
    Call mobjMipModule.InitMessage(glngSys, mlngMode, mstrPrivs)
    Call AddMipModule(mobjMipModule)
       
    mblnStart = True

'    '���ش���ʱ����Ϣ������Ч��ǰ������Ҫˢ��һ��
'    If Not mobjMipModule Is Nothing Then
'        If mobjMipModule.IsConnect = True Then
            RefreshList IIf(mParams.blnMustDosageProcess = True, mListType.����ҩ, mListType.����ҩ)
'        End If
'    End If

    mdteMsgRefresh = Now
End Sub

Private Sub GetMoneyFormat()
    Dim n As Integer
    Dim strOracleTmp As String
    Dim strVbTmp As String
    
    strOracleTmp = "999999990."
    strVbTmp = "########0."
    For n = 1 To mintMoneyDigit
        strOracleTmp = strOracleTmp & "0"
        strVbTmp = strVbTmp & "0"
    Next
    
    mstrOracleMoneyForamt = strOracleTmp
    mstrVBMoneyForamt = strVbTmp
    
End Sub
Private Sub Form_Resize()
    On Error Resume Next
    
    If Me.Width < mcstlngWinNormalWidth Then Me.Width = mcstlngWinNormalWidth
    If Me.Height < mcstlngWinNormalHeight Then Me.Height = mcstlngWinNormalHeight
End Sub



Public Function RefreshDetail_Send(ByVal lngNO�ⷿID As Long, ByVal BillStyle As Integer, ByVal BillNo As String, ByVal int�����־ As Integer, ByVal int��¼���� As Integer, Optional ByVal int�Ŷ����� As Integer, Optional ByVal int����� As Integer) As Boolean
    Dim IntStyle As Integer, intUnit As Integer
    Dim strSubSql As String
    Dim strName As String
    Dim blnMoved As Boolean
    Dim lng�ⷿID As Long
    Dim lng����ID As Long
    Dim int��ҳid As Integer
    Dim strWeight As String
    
    Dim rstemp As New ADODB.Recordset
    Dim RecBill As New ADODB.Recordset
    '--��ȡ��������--
    'BillStyle-��������;BIllNO-���ݺ�
    '��λ��ʾ���ݷ����������������ﵥλ��סԺ��סԺ���סԺ��λ���������ۼ۵�λ��
'    On Error Resume Next
'    err = 0
    On Error GoTo errHandle
    RefreshDetail_Send = False
    
    If mPrives.bln������ҩ���Ĵ��� = False Then
        lng�ⷿID = mSQLCondition.lngҩ��ID
    Else
        lng�ⷿID = lngNO�ⷿID
    End If
    
    If lng�ⷿID = 0 Then lng�ⷿID = mSQLCondition.lngҩ��ID
    
    mParams.strUnit = GetUnit(lng�ⷿID, BillStyle, BillNo, int�����־)
    Select Case mParams.strUnit
    Case "�ۼ۵�λ"
        strSubSql = "1"
    Case "���ﵥλ"
        strSubSql = "Decode(�����װ,Null,1,0,1,�����װ)"
    Case "סԺ��λ"
        strSubSql = "Decode(סԺ��װ,Null,1,0,1,סԺ��װ)"
    Case "ҩ�ⵥλ"
        strSubSql = "Decode(ҩ���װ,Null,1,0,1,ҩ���װ)"
    End Select
    Call Get��λ��
    
    '�õ�ҩƷ���ƴ�
    Select Case mParams.intҩƷ������ʾ
    Case 0  'ҩƷ����������
        strName = "'['||C.����||']'||" & IIf(gintҩƷ������ʾ = 1, "NVL(E.����,C.����)", "C.����") & " As Ʒ��,"
    Case 1  'ҩƷ����
        strName = "C.���� As Ʒ��,"
    Case 2  'ҩƷ����
        strName = IIf(gintҩƷ������ʾ = 1, "NVL(E.����,C.����)", "C.����") & " As Ʒ��,"
    End Select
    
    strName = strName & IIf(gintҩƷ������ʾ <> 1, "NVL(E.����,'')", "Decode(E.����,Null,'',C.����)") & " As ������, "
    
    gstrSQL = " SELECT DISTINCT B.��¼״̬ ״̬,S.���� As ҩ��,Nvl(B.�ⷿID,0) as ҩ��ID,B.����,B.NO,Nvl(A.��������,Nvl(B.ע��֤��,0)) As ��������,nvl(n.����,'') �䷽����,H.���,T.���� ����,H.����,H.�Ա�,H.����,H.��ʶ�� סԺ��,H.����,H.������,B.ID As �շ�ID," & _
        " B.ҩƷID,D.ҩ��id,DECODE(B.����,NULL,'',B.����)||DECODE(B.����,NULL,'',0,'','('||B.����||')') ����,A.���շ�,DECODE(D.��ΣҩƷ,null,0,0,0,1) ��ΣҩƷ,to_char(B.Ч��,'yyyy-mm-dd') Ч��,X.�����," & _
        " NVL(B.����,0) ����,NVL(D.ҩ������,0) ����,F.���� As ����,B.��� As �շ����,C.��� As ҩƷ���,H.�����־, " & strName & _
        " DECODE(C.���,NULL,B.����,DECODE(B.����,NULL,C.���,C.���||'|'||B.����)) ���,Nvl(b.����, Nvl(c.����, '')) ����,b.ԭ����," & str��λ�� & ",Nvl(K.ʵ������,0)/" & strSubSql & " �����,Nvl(K.ʵ������,0) ���ʵ������," & _
        IIf(gintҩƷ������ʾ = 1, "NVL(E.����,C.����)", "C.����") & " As ҩƷ����,decode(m.��ҩĿ��,1,'Ԥ��',2,'����',3,'Ԥ��������','') ��ҩĿ��,m.��ҩ����, " & _
        " NVL(B.����,1) ����,NVL(H.����,1) ԭʼ����,B.����,B.�÷�,B.Ƶ��,B.������,B.��������,H.����Ա����," & IIf(mcondition.intListType <> mListType.��ҩ, "B.��ҩ��", "B.�����") & " ��ҩ��,B.��ҩ����,B.�������, " & _
        " L.�ⷿ��λ,Nvl(M.���ID,0) As ���ID,M.ҽ������,M.����ҩƷ˵��,M.����ҽ��,M.Ƶ�ʼ��,M.�����λ,Nvl(M.����ʱ��,H.�Ǽ�ʱ��) As ����ʱ��,M.ҽ����Ч ҽ����־,M.��ʼִ��ʱ�� ��ʼʱ��,M.ִ����ֹʱ�� ����ʱ��," & _
        " M.Ƶ�ʴ���,Nvl(Nvl(M.���ID,M.id),0) As ҽ��id,nvl(M.�����,-1) �����,M.Ƥ�Խ��,M.����˵��,D.ҩ��ID,I.���㵥λ,D.����ϵ��," & _
        " round(B.���۽��," & mintMoneyDigit & ") ���۽��,Nvl(B.����, 1) * B.ʵ������ / (Nvl(H.����, 1) * H.����) * Nvl(H.ʵ�ս��,0) As ʵ�ս��,H.�ѱ�,P.�������,Nvl(p.������,0) ������, " & _
        " B.ʵ������*D.����ϵ��* Nvl(B.����, 1) ����,B.ʵ������,Decode(Sign(Nvl(J.�������, 0) - Nvl(L.����, 0)), -1, 0, 1) �������,Z.���� As Ӣ����,1 As ��־,Nvl(H.����ID,0) As ����ID,H.��¼����, H.��¼״̬,Zl_Get�շ����([15],[14],[13]) �շ����,X.���￨��,X.����ģʽ,decode(X.��ϵ�˵绰,null,decode(X.�ֻ���,null,X.��ͥ�绰,X.�ֻ���),X.��ϵ�˵绰) ��ϵ�˵绰,Nvl(m.��ҳid,0) As ��ҳid, Nvl(G.��鷽��, H.����) As ��ҩ��̬, I.���� As ҩ��,X.��������,Nvl(P.�Ƿ�Ƥ��,0) As �Ƿ�Ƥ��, Nvl(x.��Ժ, 0) As ��Ժ "
    gstrSQL = gstrSQL & _
        " FROM ҩƷ�շ���¼ B,ҩƷ��� D,ҩƷ���� P,�շ���ĿĿ¼ C,�շ���Ŀ���� E," & _
        " ������ü�¼ H,����ҽ����¼ M,����ҽ����¼ G,������Ϣ X,���ű� S,���ű� T,ҩƷ��� K,ҩƷ�����޶� L,������ĿĿ¼ I,������Ŀ���� Z ,δ��ҩƷ��¼ A,����֧������ F,������ĿĿ¼ N, " & _
        " (Select b.�ⷿid, b.ҩƷid, Nvl(Sum(b.ʵ������), 0) ������� " & _
        " From ҩƷ�շ���¼ A, ҩƷ��� B " & _
        " Where a.ҩƷid = b.ҩƷid And b.���� = 1 And b.�ⷿid + 0 = [13] And a.���� = [15] And a.No = [14] " & _
        " Group By b.�ⷿid, b.ҩƷid) J " & _
        " WHERE A.����=B.���� And A.NO=B.No And D.ҩƷID=C.ID And D.ҩ��ID=P.ҩ��ID And H.ҽ�����=M.ID(+) And Nvl(M.���id, M.ID) = G.ID(+) AND C.ID=E.�շ�ϸĿID(+) AND E.����(+)=3 " & _
        " And B.ҩƷID=L.ҩƷID(+) And Nvl(B.�ⷿID,[13])=L.�ⷿID(+) And H.���մ���ID=F.ID(+) and G.�䷽id=N.id(+) " & _
        " AND H.��������ID=T.ID(+) AND B.ҩƷID=D.ҩƷID AND MOD(B.��¼״̬,3)=1" & _
        " AND S.ID=NVL(B.�ⷿID,[13]) AND B.����ID=H.ID AND B.NO=[14] AND B.����=[15] AND NVL(B.�ⷿID,[13])+0=[13] AND LTRIM(RTRIM(NVL(B.ժҪ,'С��')))<>'�ܷ�'" & _
        " AND B.ҩƷID=K.ҩƷID(+) AND K.����(+)=1 AND NVL(B.�ⷿID,[13])=K.�ⷿID(+) AND NVL(B.����,0)=NVL(K.����(+),0) AND B.����� IS NULL And D.ҩ��id=I.id " & _
        " And Nvl(B.�ⷿid, [13]) + 0 = J.�ⷿid(+) And B.ҩƷid = J.ҩƷid(+) And D.ҩ��id = Z.������Ŀid(+) And Z.����(+) = 2 And H.����id = X.����id(+) "
        
    gstrSQL = gstrSQL & " Order by H.���,B.ҩƷID,Nvl(B.����,0)"
    
    If int��¼���� = 1 Or (int��¼���� = 2 And (int�����־ = 1 Or int�����־ = 4)) Then
        '����
        gstrSQL = Replace(gstrSQL, "H.����", "'' ����")
    Else
        'סԺ
        gstrSQL = Replace(gstrSQL, "������ü�¼", "סԺ���ü�¼")
    End If
     
    Set RecBill = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, _
            mSQLCondition.date��ʼ����, _
            mSQLCondition.date��������, _
            mSQLCondition.str��ʼNO, _
            mSQLCondition.str����NO, _
            mSQLCondition.str����, _
            mSQLCondition.str���￨, _
            mSQLCondition.str��ʶ��, _
            mSQLCondition.lng����ID, _
            mSQLCondition.str������, _
            mSQLCondition.str�����, _
            mSQLCondition.lngҩƷid, _
            mSQLCondition.str��ǰNO, _
            lng�ⷿID, BillNo, BillStyle)
    
    If err <> 0 Then
        MsgBox "��ȡ����ʱ����������Ԥ֪�Ĵ���", vbInformation, gstrSysName
        Exit Function
    End If
    
'    If RecBill!ҩ��id <> mParams.lngҩ��ID Then RecBill.Filter = "״̬=1"
    
    If Not RecBill.EOF Then
        If NVL(RecBill!����ID) <> 0 Then
            lng����ID = RecBill!����ID
            int��ҳid = NVL(RecBill!��ҳid)
            If int��¼���� = 1 Or (int��¼���� = 2 And (int�����־ = 1 Or int�����־ = 4)) Then
                '����
                gstrSQL = "select A.id,B.��¼���� ���� from ���˻����¼ A,���˻������� B where A.id=B.��¼id and B.��Ŀ����='����' and ����id=[1] order by A.Id desc"
            Else
                'סԺ
                 gstrSQL = "select ���� from ������ҳ where ����id=[1] and ��ҳid=[2]"
                 
            End If
            Set rstemp = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, lng����ID, int��ҳid)
            
            If Not rstemp.EOF Then
                strWeight = NVL(rstemp!����)
            End If
        End If
    End If
    
    mfrmDetail.RefreshList RecBill, strWeight, 0, int�Ŷ�����, int�����
    
    mfrmRecipe.RefreshRecipe RecBill, strWeight, 0, int�Ŷ�����, int�����
    
    RefreshDetail_Send = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function GetRecipeRecord(ByVal BillStyle As Integer, ByVal BillNo As String, ByVal int�����־ As Integer, ByVal int��¼���� As Integer) As ADODB.Recordset
    Dim IntStyle As Integer, intUnit As Integer
    Dim strSubSql As String
    Dim strName As String
    Dim blnMoved As Boolean
    
    Dim rstemp As New ADODB.Recordset
    Dim RecBill As New ADODB.Recordset
    '--��ȡ��������--
    'BillStyle-��������;BIllNO-���ݺ�
    '��λ��ʾ���ݷ����������������ﵥλ��סԺ��סԺ���סԺ��λ���������ۼ۵�λ��
'    On Error Resume Next
    On Error GoTo errHandle
    mParams.strUnit = GetUnit(mSQLCondition.lngҩ��ID, BillStyle, BillNo, int�����־)
    Select Case mParams.strUnit
    Case "�ۼ۵�λ"
        strSubSql = "1"
    Case "���ﵥλ"
        strSubSql = "Decode(�����װ,Null,1,0,1,�����װ)"
    Case "סԺ��λ"
        strSubSql = "Decode(סԺ��װ,Null,1,0,1,סԺ��װ)"
    Case "ҩ�ⵥλ"
        strSubSql = "Decode(ҩ���װ,Null,1,0,1,ҩ���װ)"
    End Select
    Call Get��λ��
    
    '�õ�ҩƷ���ƴ�
    Select Case mParams.intҩƷ������ʾ
    Case 0  'ҩƷ����������
        strName = "'['||C.����||']'||" & IIf(gintҩƷ������ʾ = 1, "NVL(E.����,C.����)", "C.����") & " As Ʒ��,"
    Case 1  'ҩƷ����
        strName = "C.���� As Ʒ��,"
    Case 2  'ҩƷ����
        strName = IIf(gintҩƷ������ʾ = 1, "NVL(E.����,C.����)", "C.����") & " As Ʒ��,"
    End Select
    
    strName = strName & IIf(gintҩƷ������ʾ <> 1, "NVL(E.����,'')", "Decode(E.����,Null,'',C.����)") & " As ������, "
    
    gstrSQL = " SELECT DISTINCT B.����,B.NO,Nvl(A.��������,0) As ��������,H.���,T.���� ����,H.����,H.�Ա�,H.����,H.��ʶ�� סԺ��,H.������,B.ID As �շ�ID," & _
        " B.ҩƷID,DECODE(B.����,NULL,'',B.����)||DECODE(B.����,NULL,'',0,'','('||B.����||')') ����,A.���շ�," & _
        " NVL(B.����,0) ����,NVL(D.ҩ������,0) ����,F.���� As ����,B.��� As �շ����,C.��� As ҩƷ���,H.�����־, " & strName & _
        " DECODE(C.���,NULL,B.����,DECODE(B.����,NULL,C.���,C.���||'|'||B.����)) ���," & str��λ�� & ",K.ʵ������/" & strSubSql & " �����," & _
        IIf(gintҩƷ������ʾ = 1, "NVL(E.����,C.����)", "C.����") & " As ҩƷ����, " & _
        " NVL(H.����,1) ����,B.����,B.�÷�,B.Ƶ��,B.������,B.��������,H.����Ա����," & IIf(mcondition.intListType <> mListType.��ҩ, "B.��ҩ��", "B.�����") & " ��ҩ��," & _
        " L.�ⷿ��λ,Nvl(M.���ID,0) As ���ID,M.ҽ������,Nvl(Nvl(M.���ID,M.id),0) As ҽ��id,nvl(M.�����,-1) �����,I.���㵥λ," & _
        " round(B.���۽��," & mintMoneyDigit & ") ���۽��,Nvl(B.����, 1) * B.ʵ������ / (Nvl(H.����, 1) * H.����) * Nvl(H.ʵ�ս��,0) As ʵ�ս��,H.�ѱ�,P.�������, " & _
        " B.ʵ������*D.����ϵ��* Nvl(B.����, 1) ����,Decode(Sign(Nvl(J.�������, 0) - Nvl(L.����, 0)), -1, 0, 1) �������,Z.���� As Ӣ����,1 As ��־,Nvl(H.����ID,0) As ����ID,H.��¼����, H.��¼״̬,Zl_Get�շ����(b.����,b.NO,[13]) As �շ����,X.����ģʽ, X.���￨��,X.��������,B.�ⷿid As ҩ��ID " & _
        " FROM ҩƷ�շ���¼ B,ҩƷ��� D,ҩƷ���� P,�շ���ĿĿ¼ C,�շ���Ŀ���� E,"
            
    If int��¼���� = 1 Or (int��¼���� = 2 And (int�����־ = 1 Or int�����־ = 4)) Then
        gstrSQL = gstrSQL & " ������ü�¼ H,"
    Else
        gstrSQL = gstrSQL & " סԺ���ü�¼ H,"
    End If
            
    gstrSQL = gstrSQL & " ����ҽ����¼ M,������Ϣ X,���ű� S,���ű� T,ҩƷ��� K,ҩƷ�����޶� L,������ĿĿ¼ I,������Ŀ���� Z ,δ��ҩƷ��¼ A,����֧������ F," & _
        " (Select �ⷿid, ҩƷid, Nvl(Sum(ʵ������), 0) ������� From ҩƷ��� Where ���� = 1 And �ⷿid = [13] Group By �ⷿid, ҩƷid) J " & _
        " WHERE A.����=B.���� And A.NO=B.No And D.ҩƷID=C.ID And D.ҩ��ID=P.ҩ��ID And H.ҽ�����=M.ID(+) AND C.ID=E.�շ�ϸĿID(+) AND E.����(+)=3 " & _
        " And B.ҩƷID=L.ҩƷID(+) And Nvl(B.�ⷿID,[13])=L.�ⷿID(+) And H.���մ���ID=F.ID(+) " & _
        " AND H.��������ID=T.ID(+) AND B.ҩƷID=D.ҩƷID AND MOD(B.��¼״̬,3)=1" & _
        " AND S.ID=NVL(B.�ⷿID,[13]) AND B.����ID=H.ID AND B.NO=[14] AND B.����=[15] AND NVL(B.�ⷿID,[13])+0=[13] AND LTRIM(RTRIM(NVL(B.ժҪ,'С��')))<>'�ܷ�'" & _
        " AND B.ҩƷID=K.ҩƷID(+) AND K.����(+)=1 AND NVL(B.�ⷿID,[13])=K.�ⷿID(+) AND NVL(B.����,0)=NVL(K.����(+),0) AND B.����� IS NULL And D.ҩ��id=I.id " & _
        " And Nvl(B.�ⷿid, [13]) + 0 = J.�ⷿid(+) And B.ҩƷid = J.ҩƷid(+) And D.ҩ��id = Z.������Ŀid(+) And Z.����(+) = 2 And H.����id = X.����id(+) " & _
        " Order by H.���,B.ҩƷID,Nvl(B.����,0)"
     
    Set GetRecipeRecord = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, _
            mSQLCondition.date��ʼ����, _
            mSQLCondition.date��������, _
            mSQLCondition.str��ʼNO, _
            mSQLCondition.str����NO, _
            mSQLCondition.str����, _
            mSQLCondition.str���￨, _
            mSQLCondition.str��ʶ��, _
            mSQLCondition.lng����ID, _
            mSQLCondition.str������, _
            mSQLCondition.str�����, _
            mSQLCondition.lngҩƷid, _
            mSQLCondition.str��ǰNO, _
            mSQLCondition.lngҩ��ID, BillNo, BillStyle)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function



Private Sub Get��λ��()
    Const str�ۼ� As String = "C.���㵥λ As �ۼ۵�λ,C.���㵥λ As ��λ,1 As ��װ,ltrim(to_char(B.���ۼ�,'999990.00000')) ����,ltrim(to_char(B.ʵ������,'999990.00000')) ����"
    Const str���� As String = "C.���㵥λ As �ۼ۵�λ,D.���ﵥλ As ��λ,D.�����װ As ��װ,ltrim(to_char(B.���ۼ�*Decode(D.�����װ,Null,1,0,1,D.�����װ),'999990.00000')) ����,ltrim(to_char(B.ʵ������/Decode(D.�����װ,Null,1,0,1,D.�����װ),'999990.00000')) ����"
    Const strסԺ As String = "C.���㵥λ As �ۼ۵�λ,D.סԺ��λ As ��λ,D.סԺ��װ As ��װ,ltrim(to_char(B.���ۼ�*Decode(D.סԺ��װ,Null,1,0,1,D.סԺ��װ),'999990.00000')) ����,ltrim(to_char(B.ʵ������/Decode(D.סԺ��װ,Null,1,0,1,D.סԺ��װ),'999990.00000')) ����"
    Const strҩ�� As String = "C.���㵥λ As �ۼ۵�λ,D.ҩ�ⵥλ As ��λ,D.ҩ���װ As ��װ,ltrim(to_char(B.���ۼ�*Decode(D.ҩ���װ,Null,1,0,1,D.ҩ���װ),'999990.00000')) ����,ltrim(to_char(B.ʵ������/Decode(D.ҩ���װ,Null,1,0,1,D.ҩ���װ),'999990.00000')) ����"
    
    Select Case mParams.strUnit
    Case "�ۼ۵�λ"
        str��λ�� = str�ۼ�
    Case "���ﵥλ"
        str��λ�� = str����
    Case "סԺ��λ"
        str��λ�� = strסԺ
    Case "ҩ�ⵥλ"
        str��λ�� = strҩ��
    End Select
End Sub
Private Sub Form_Unload(Cancel As Integer)
    mblnBrushCard = False
    mstrLastBrushCardNo = ""
    mQueue.strSendWin = ""
    mstr�������� = ""
    
    '�������������ϵͳ�����˳�����ʱ�ر����ڲ��ŵ�����
    If mParams.blnStartQueue = True And mParams.blnStartCall = True And mQueue.blnCallOver = False Then
        Call StopPlayStr
    End If
    tmrCall.Enabled = False
    mQueue.blnCallOver = True

    Set mobjDrugMAC = Nothing
    Set mclsComLib = Nothing
    
    TimeRefresh.Enabled = False
    TimePrint.Enabled = False
    TimePrintCancelBill.Enabled = False
    
    zldatabase.SetPara "��ʾ��������", img����.BorderStyle, glngSys, 1341, IsInString(mstrPrivs, "��������", ";")
    
    Call SaveSetting("ZLSOFT", "����ģ��\����\" & App.ProductName & "\" & "ҩƷ������ҩ", "���涨λ", imgFilter.BorderStyle)
    Call SaveSetting("ZLSOFT", "����ģ��\����\" & App.ProductName & "\" & "ҩƷ������ҩ", "��ʾ��ҩ��������", Chk��ʾ��ҩ��������.Value)
    Call SaveSetting("ZLSOFT", "����ģ��\����\" & App.ProductName & "\" & "ҩƷ������ҩ", "��ʾ��ҩ���̵���", Chk��ʾ���̵���.Value)
    Call SaveSetting("ZLSOFT", "����ģ��\����\" & App.ProductName & "\" & "ҩƷ������ҩ", "��ʾ��ȷ�ϵ���", chk��ʾ��ȷ�ϵ���.Value)
    
'    '��������
'    Call SaveSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\" & "ҩƷ������ҩ", "δ��ҩ��������", strOrder_1)
'    Call SaveSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\" & "ҩƷ������ҩ", "����ҩ��������", strOrder_2)
'    Call SaveSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\" & "ҩƷ������ҩ", "δ��ҩ��������", strOrder_3)
'    Call SaveSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\" & "ҩƷ������ҩ", "�ѷ�ҩ��������", strOrder_4)
    
    '��������ģʽ
    Call SaveSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\" & "ҩƷ������ҩ", "����ģʽ", mParams.int����ģʽ����)
    
    Call SaveWinState(Me, App.ProductName)
    
    '���洰��
    If Val(zldatabase.GetPara("ʹ�ø��Ի����")) = 1 Then
        Call SaveSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\��������\" & App.ProductName & "\" & Me.Name & "\" & TypeName(dkpMain), dkpMain.Name & dkpMain.PanesCount, dkpMain.SaveStateToString)
    End If
    
    'ж�ص��Ӳ������Ľӿ�
    Set mobjCISJOB = Nothing
    
    'ж�����֤ˢ���ӿ�
    If Not mobjIDCard Is Nothing Then
        Call mobjIDCard.SetEnabled(False)
        Set mobjIDCard = Nothing
    End If
    
    'ж��IC��ˢ���ӿ�
    If Not mobjICCard Is Nothing Then
        Call mobjICCard.SetEnabled(False)
        Set mobjICCard = Nothing
    End If
    
    'ж��CARD����
    If Not mobjcard Is Nothing Then
        Set mobjcard = Nothing
    End If
    
    'ж�ص���ǩ���ӿ�
    Set gobjESign = Nothing
    
    'ж��һ��ͨ�ӿ�
    mstrCardType = ""
    Call zlfuncCard_Unload(mobjSquareCard)
    
    'ж�����õĴ���
    If Not mfrmList Is Nothing Then
        Unload mfrmList
        Set mfrmList = Nothing
    End If
    
    If Not mfrmDetail Is Nothing Then
        Unload mfrmDetail
        Set mfrmDetail = Nothing
    End If
    
    If Not mfrmRecipe Is Nothing Then
        Unload mfrmRecipe
        Set mfrmRecipe = Nothing
    End If
    
    '�ر���ʾ�������
    CloseQueue
    
    'ж�ع�������
    mSQLCondition.str���￨ = ""
    mSQLCondition.str��ǰNO = ""
    mSQLCondition.str����� = ""
    mSQLCondition.str���� = ""
    mSQLCondition.str���֤ = ""
    mSQLCondition.lng����ID = 0
    mSQLCondition.strҽ���� = ""
    mSQLCondition.lngסԺ�� = 0
    
    
    'ж����Ϣ����
    If Not mobjMipModule Is Nothing Then
        Call mobjMipModule.CloseMessage
        Call DelMipModule(mobjMipModule)
        Set mobjMipModule = Nothing
    End If
    mblnExistMsg = False
    
    'ж����ҽӿ�
    Call zlPlugIn_Unload(mobjPlugIn)
End Sub

Private Sub imgFilter_Click()
    imgFilter.BorderStyle = Abs(imgFilter.BorderStyle - 1)
    
    If imgFilter.BorderStyle = 0 Then
        imgFilter.ToolTipText = "����л�������ģʽ"
    Else
        imgFilter.ToolTipText = "����л�����λģʽ"
    End If
    
    mParams.int���涨λ = imgFilter.BorderStyle
    '������涨λ��ʽ
    Call SaveSetting("ZLSOFT", "����ģ��\����\" & App.ProductName & "\" & "ҩƷ������ҩ", "���涨λ", mParams.int���涨λ)
    
    '����ˢ��
    mSQLCondition.str���￨ = ""
    mSQLCondition.str��ǰNO = ""
    mSQLCondition.str����� = ""
    mSQLCondition.str���� = ""
    mSQLCondition.str���֤ = ""
    mSQLCondition.lng����ID = 0
    mSQLCondition.strҽ���� = ""
    mSQLCondition.lngסԺ�� = 0
    mlngIC����id = 0
    
    txtPati.Text = ""
    RefreshList mcondition.intListType
End Sub

Private Sub img����_Click()
    With img����
        .BorderStyle = Abs(.BorderStyle - 1)
        
        cbo����.Enabled = (.BorderStyle = 1)
        
        If cbo����.Enabled = True Then
            Load����
        Else
            cbo����.ListIndex = -1
        End If
        
        Call RefreshList(mcondition.intListType)
    End With
End Sub


Private Sub mobjIDCard_ShowIDCardInfo(ByVal strID As String, ByVal strName As String, ByVal strSex As String, _
                            ByVal strNation As String, ByVal datBirthDay As Date, ByVal strAddress As String)
    
    If Not txtPati.Locked And txtPati.Text = "" And Me.ActiveControl Is txtPati Then
        txtPati.Text = strID
        
        If txtPati.Text <> "" Then
            mParams.int����ģʽ = mFindType.���֤
'            Call SetInputState(mParams.int����ģʽ)
            
            DoEvents
            
            Call txtPati_KeyPress(vbKeyReturn)
            
            DoEvents
            
            If mintOld����ģʽ <> mParams.int����ģʽ Then
                mParams.int����ģʽ = mintOld����ģʽ
'                Call SetInputState(mParams.int����ģʽ)
            End If
        End If
    End If
End Sub

Private Sub mobjMipModule_ReceiveMessage(ByVal strMsgItemIdentity As String, ByVal strMsgContent As String)
    '1.������Ϣ����Ϣ���ͺ�����ҵ��Լ��
    '2.���ݿͻ������������ж��Ƿ�����Ч��Ϣ
    '3.Ĭ��1����ˢ��һ��
    Const CST_INT_MSGREFRESHINTERVAL As Integer = 1
    Const CST_STR_MSGCODE As String = "ZLHIS_CHARGE_003,ZLHIS_CIS_006"
    
    '��Ϣ����Ϊ��ʱ�˳�
    If mobjMipModule Is Nothing Then Exit Sub
    
    '��Ϣ��������ʧ��ʱ��������Ϣ
    If mobjMipModule.IsConnect = False Then Exit Sub
        
    '������ҩ���յ���Ϣ���ͣ�����/�շѺ�����ҽ��վ������ҩƷ���������಻����
    If InStr("," & CST_STR_MSGCODE & ",", "," & strMsgItemIdentity & ",") = 0 Then Exit Sub

    '���ݿͻ������������ж��Ƿ�����Ч��Ϣ
    If IsValidMsg(strMsgItemIdentity, strMsgContent) = False Then Exit Sub
    
    'ִ�е������ʾ�ѽ��յ���Ч��Ϣ
    mblnExistMsg = True
    
    '��ǰ������Ǵ���ҩ�����ҩ�����򲻼���
    If (mParams.blnMustDosageProcess = True And mcondition.intListType <> mListType.����ҩ) Or _
        (mParams.blnMustDosageProcess = False And mcondition.intListType <> mListType.����ҩ) Then
        Exit Sub
    End If
    
    '������յ���Ч��Ϣʱ���ϴ�ˢ�³���1����������ˢ��
    If DateDiff("n", mdteMsgRefresh, Now) > CST_INT_MSGREFRESHINTERVAL Then
        'ˢ��ǰ�رռ�ʱ����ˢ�º��ٿ���
        tmrMsgRefresh.Enabled = False
        DoEvents
        Call RefreshList(mcondition.intListType)
        DoEvents
        tmrMsgRefresh.Enabled = True
        
        'ˢ�º��¼��ǰˢ��ʱ��
        mdteMsgRefresh = Now
        
        '��Ϣ������ΪFalse
        mblnExistMsg = False
    End If
End Sub


Private Sub picCondition_Resize()
    On Error Resume Next
    
    With picConMain
        .Top = 0
        .Left = 0
        .Width = picCondition.Width
        
    End With
    
    With picList
        .Top = picConMain.Top + picConMain.Height
        .Left = 0
        .Width = picCondition.Width
        .Height = picCondition.Height - .Top
    End With
End Sub


Private Sub picConMain_Resize()
    On Error Resume Next
    
    With cboʱ�䷶Χ
        .Width = picCondition.Width - .Left - 50
    End With

    If cboʱ�䷶Χ.ListIndex <> 3 Then
        lblTimeBegin.Visible = False
        Dtp��ʼʱ��.Visible = False
        lblTimeEnd.Visible = False
        Dtp����ʱ��.Visible = False
        
        With Me.IDKNType
            .Top = lblʱ�䷶Χ.Top + lblʱ�䷶Χ.Height + 180
        End With
        
        With txtPati
            .Top = cboʱ�䷶Χ.Top + cboʱ�䷶Χ.Height + 180
        End With
    Else
        lblTimeBegin.Visible = True
        Dtp��ʼʱ��.Visible = True
        lblTimeEnd.Visible = True
        Dtp����ʱ��.Visible = True
        
        With lblTimeBegin
            .Top = lblʱ�䷶Χ.Top + lblʱ�䷶Χ.Height + 180
        End With
        
        With Dtp��ʼʱ��
            .Top = lblTimeBegin.Top + lblTimeBegin.Height / 2 - .Height / 2
            .Width = cboʱ�䷶Χ.Width
        End With
        
        With lblTimeEnd
            .Top = lblTimeBegin.Top + lblTimeBegin.Height + 180
        End With
        
        With Dtp����ʱ��
            .Top = Dtp��ʼʱ��.Top + Dtp��ʼʱ��.Height + 60
            .Width = cboʱ�䷶Χ.Width
        End With
        
        With IDKNType
            .Top = lblTimeEnd.Top + lblTimeEnd.Height + 180
        End With
        
        With txtPati
            .Top = IDKNType.Top + IDKNType.Height / 2 - .Height / 2
        End With
    End If
    
    With cmdIC
        .Visible = (mobjcard.�Ƿ�ˢ�� = 1)
        .Top = txtPati.Top
        .Left = picCondition.Width - .Width - 80
    End With

    With imgFilter
        .Top = txtPati.Top
        .Left = IIf(mobjcard.�Ƿ�ˢ�� = 1, cmdIC.Left, picCondition.Width) - imgFilter.Width - 120
    End With

    With cmdFind
        .Top = cmdIC.Top
        .Left = imgFilter.Left + 120
    End With

    With txtPati
        .Width = imgFilter.Left - .Left - 200
    End With
    
    If lbl����.Visible = True Then
        With lbl����
            .Top = IDKNType.Top + IDKNType.Height + 180
        End With
        
        With img����
            .Top = lbl����.Top - 30
        End With
        
        With cbo����
            .Top = img����.Top - 30
            .Width = cboʱ�䷶Χ.Width
        End With
        
        With Chk��ʾ��ҩ��������
            .Left = lbl����.Left
            .Top = lbl����.Top + 350
        End With
    Else
        With Chk��ʾ��ҩ��������
            .Left = IDKNType.Left
            .Top = IDKNType.Top + 350
        End With
    End If
    
    With Chk��ʾ���̵���
        .Left = Chk��ʾ��ҩ��������.Left
        .Top = Chk��ʾ��ҩ��������.Top
    End With
    
    With chk��ʾ��ȷ�ϵ���
        .Left = Chk��ʾ��ҩ��������.Left
        .Top = Chk��ʾ��ҩ��������.Top
    End With
    
    With picConMain
        .Height = Chk��ʾ��ҩ��������.Top + Chk��ʾ��ҩ��������.Height + 50
    End With
End Sub


Private Sub picDetail_Resize()
    On Error Resume Next
    
    With fraLine
'        .Top = 0
        .Left = 0
        .Height = picDetail.Height + 100
    End With
    
    With tbcDetail
        .Top = 0
        .Left = fraLine.Left + 50
        .Width = picDetail.Width - fraLine.Width
        .Height = picDetail.Height - 50
    End With
End Sub

Private Sub picList_Resize()
    On Error Resume Next
    
    With tbcList
        .Move 0, 0, picList.Width, picList.Height
    End With
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
        
        .InsertItem(0, "������ϸ�嵥", mfrmDetail.hWnd, 0).Tag = "������ϸ�嵥_"
        .InsertItem(1, "����ǩ", mfrmRecipe.hWnd, 0).Tag = "����ǩ_"
        
        .Item(0).Selected = True
    End With
    
    With Me.tbcList
        With .PaintManager
            .Appearance = xtpTabAppearancePropertyPage2003
            .ClientFrame = xtpTabFrameSingleLine
            .BoldSelected = True
            .OneNoteColors = True
            .ShowIcons = True
        End With
        
        .InsertItem(mconTab_Recipe_DosageOk, "��ҩȷ��", mfrmList.hWnd, 0).Tag = "��ҩȷ��"
        .InsertItem(mconTab_Recipe_Dosage, "����ҩ", mfrmList.hWnd, 0).Tag = "����ҩ_"
        .InsertItem(mconTab_Recipe_Abolish, "����ҩ", mfrmList.hWnd, 0).Tag = "����ҩ_"
        .InsertItem(mconTab_Recipe_Send, "����ҩ", mfrmList.hWnd, 0).Tag = "����ҩ_"
        .InsertItem(mconTab_Recipe_OverTime, "��ʱδ��", mfrmList.hWnd, 0).Tag = "��ʱδ��_"
        .InsertItem(mconTab_Recipe_Return, "��ҩ", mfrmList.hWnd, 0).Tag = "��ҩ_"
        
        .Item(mconTab_Recipe_Send).Selected = True
        .Item(mconTab_Recipe_DosageOk).Visible = False
        .Item(mconTab_Recipe_Abolish).Visible = False
        
        If mParams.blnMustDosageProcess = True Then
'            .Item(mconTab_Recipe_Abolish).Selected = True
            .Item(mconTab_Recipe_Dosage).Selected = True
        Else
            .Item(mconTab_Recipe_Dosage).Visible = False
'            .Item(mconTab_Recipe_Abolish).Visible = False
        End If
        
'        If mParams.blnMustDosageOkProcess = True And InStr(1, mstrPrivs, "��ҩȷ��") > 0 Then
'            .Item(mconTab_Recipe_DosageOk).Selected = True
'        Else
'            .Item(mconTab_Recipe_DosageOk).Visible = False
'        End If
'
        If mParams.blnMustDosageOkProcess = False And mParams.blnMustDosageProcess = False Then
            .Item(mconTab_Recipe_Send).Selected = True
        End If
        
        If Not .Item(mconTab_Recipe_OverTime) Is Nothing Then
            .Item(mconTab_Recipe_OverTime).Visible = (mParams.intOverTime > 0)
        End If
    End With
End Sub

Private Sub tbcDetail_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
'    If Item.Index = 0 Then
'        If Not mfrmDetail Is Nothing Then mfrmDetail.ShowList mcondition.intListType, mcondition.bln��ʾ���̵���
'    Else
'        If Not mfrmRecipe Is Nothing Then mfrmRecipe.ShowRecipe mcondition.intListType
'    End If
    
    mintTab = Item.index
End Sub

Private Sub tbcList_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    Dim strTitleCon As String
    Dim strTitleList As String
    Dim objPaneCon As Pane
    '������Ĳ˵�״̬
    Dim cbrControl As CommandBarControl
    Dim cbrMenu As CommandBarControl
    
    If Item.Tag = "" Then Exit Sub
    
    Select Case Item.index
        Case mListType.��ҩȷ��
            strTitleCon = "��ҩȷ��"
            strTitleList = "�����б�(��ҩȷ��)"
        Case mListType.����ҩ
            strTitleCon = "����ҩ"
            strTitleList = "�����б�(����ҩ)"
        Case mListType.����ҩ
            strTitleCon = "����ҩ"
            strTitleList = "�����б�(����ҩ)"
        Case mListType.����ҩ
            strTitleCon = "����ҩ"
            strTitleList = "�����б�(����ҩ)"
        Case mListType.��ʱδ��
            strTitleCon = "����ҩ"
            strTitleList = "�����б�(��ʱδ��)"
        Case mListType.��ҩ
            strTitleCon = "��ҩ"
            strTitleList = "�����б�(��ҩ)"
    End Select
    
    If mPrives.bln�����ѯ����ʱ�䷶Χ���� Then
        If mPrives.bln�޸Ĺ������� = False Then
            '��Ȩ��ʱ��ҩ����ָ��ʱ�䷶Χ��ѯ
            If Item.index = mListType.��ҩ Then
                With cboʱ�䷶Χ
                    .Clear
                    .AddItem "0-����"
                    .AddItem "1-������"
                    .AddItem "2-������"
                    
                    If Val(.Tag) = 3 Then
                        .ListIndex = 0
                        Call picConMain_Resize
                        Call picCondition_Resize
                    Else
                        .ListIndex = Val(.Tag)
                    End If
                     .Tag = .ListIndex
                End With
            
                Me.Dtp��ʼʱ��.Enabled = False
                Me.Dtp����ʱ��.Enabled = False
            Else
                With cboʱ�䷶Χ
                    .Clear
                    .AddItem "0-����"
                    .AddItem "1-������"
                    .AddItem "2-������"
                    .AddItem "3-ָ��ʱ�䷶Χ"
                    
                    .ListIndex = Val(.Tag)
                 End With
        
                Me.Dtp��ʼʱ��.Enabled = True
                Me.Dtp����ʱ��.Enabled = True
            End If
        End If
    End If
             
    Chk��ʾ��ҩ��������.Visible = (Item.index <> mListType.��ҩ And Item.index <> mListType.��ҩȷ��)
    Chk��ʾ���̵���.Visible = (Item.index = mListType.��ҩ)
    Me.chk��ʾ��ȷ�ϵ���.Visible = (Item.index = mListType.��ҩȷ��)
    
    Me.dkpMain.FindPane(mconPane_Recipe_Condition).Title = mstrStockName & ":" & strTitleCon
    
    If Not mfrmList Is Nothing Then
        mfrmList.ShowList Item.index, imgFilter.BorderStyle, (mParams.blnStartCall And mParams.blnStartQueue), mParams.blnMustDosageOkProcess, mParams.blnMustDosageProcess, mParams.bln������
    End If
    
    If Not mfrmDetail Is Nothing Then
        mfrmDetail.ShowList Item.index, mcondition.bln��ʾ���̵���
    End If
    
    If Not mfrmRecipe Is Nothing Then
        mfrmRecipe.ShowRecipe Item.index
    End If
    
    mcondition.intListType = Item.index
    
    SetComandBars
    
    DoEvents
    Call RefreshList(mcondition.intListType)
    
    If Me.dkpMain.FindPane(mconPane_Recipe_Condition).Hidden = False And Visible Then mfrmList.SetFocus
End Sub

Private Sub TimePrint_Timer()
    If Me.WindowState = vbMinimized Then Exit Sub
    
    '�������ڲ��ǵ�ǰ����ʱ�˳�
    If InStr(1, "frmҩƷ������ҩNew;frm������ҩ��ϸ;frm������ҩ�б�;frm����", Screen.ActiveForm.Name) = 0 Then Exit Sub
'    If Screen.ActiveForm.hWnd <> Me.hWnd Then Exit Sub
    
    If mcondition.intListType = mListType.��ҩ Then
        Exit Sub
    End If
    
    '�����Ϣ������Ч��ͨ����ѯ��ʽ�Զ���ӡ
    If Not mobjMipModule Is Nothing Then
        If mobjMipModule.IsConnect = True Then
            Exit Sub
        End If
    End If
     
    TimePrint.Enabled = False
    DoEvents
    '���ô�ӡ����
    Call AutoPrint
    DoEvents
    TimePrint.Enabled = True
End Sub
Private Function AutoPrint()
'���ܣ��Զ���ӡ����
    Dim recAutoPrint As New ADODB.Recordset, strErr As String
    Dim datCurr As Date, strRefresh As String, strCond As String
    Dim strUnit As String
    Dim str����Ա As String
    Dim blnInTrans As Boolean
    Dim blnIgnore As Boolean
    Dim strName As String
    Dim strSub1 As String
    Dim strSub2 As String
    Dim bln��ҩ���� As Boolean
    Dim strǩ����¼ As String
    Dim str�շ����� As String
    Dim lng����ID As Long
    Dim strNo As String
    Dim strReturn As String
    Dim strסԺ As String
    
    '���ݴ�ӡ�����������
    '0-����ӡδ��ҩ����
    '1-��ӡ����������δ��ҩ����
    '2-��ӡ����������δ��ҩ����
    '3-ѡ���ӡ(��ҩ����)
    If BlnInRefresh Then Exit Function
    
    On Error GoTo ErrHand
    
    If mblnIsFirst = False And mParams.bln�Զ���ҩ Then
        If mParams.int�Զ���ҩʱ�� > 0 Then
            If DateDiff("s", mdate�ϴ�У��ʱ��, Sys.Currentdate) > mParams.int�Զ���ҩʱ�� * 60 Then
                If mParams.intУ����ҩ�� = 1 Then
                    strName = zldatabase.UserIdentify(Me, "У����ҩ��", glngSys, 1341, "��ҩ")
                
                    If Trim(strName) = "" Then Exit Function
                    mstr�Զ���ҩ�� = strName
                End If
                
                mdate�ϴ�У��ʱ�� = Sys.Currentdate
            End If
        End If
    End If
    
    '��ӡ���ķ����嵥����ǰ�棬������������Ӱ�죬Ҫ�з���ģ�鵥�ݴ�ӡ��Ȩ��
    If mParams.int��ӡ���ķ����嵥 = 1 And IsHavePrivs(mstrStuffPrivs, "���ݴ�ӡ") Then
        gstrSQL = "Select NO, ����, ��������, 1 As ����, Nvl(��������, 0) As ��������, Nvl(��ӡ״̬, 0) As ��ӡ״̬, a.����id, a.���ȼ�, a.���� " & vbNewLine & _
            "From δ��ҩƷ��¼ A, ������Ϣ B" & vbNewLine & _
            "Where a.����id = b.����id And a.�ⷿid + 0 = [1] And a.�������� Between [2] And [3] And a.��ӡ״̬ = 0 And" & vbNewLine & _
            "      a.���� In (24, 25)"
            
        '�ų��쳣����
        gstrSQL = gstrSQL & " And Not Exists (Select 1 From ������ü�¼ C Where c.No = a.No And c.ִ�в���id = a.�ⷿid And Decode(a.����, 24, 1, 25, 2) = c.��¼���� And c.ִ��״̬ = 9) "
        
        '�������ü�¼���ж������סԺ
        gstrSQL = gstrSQL & " And Exists (Select 1 From ������ü�¼ C Where a.No = c.No And a.�ⷿid = c.ִ�в���id And Decode(a.����, 24, 1, 25, 2) = c.��¼���� ) "
        
        '�����סԺ
        strסԺ = gstrSQL
        strסԺ = Replace(strסԺ, "1 As ����", "2 As ����")
        strסԺ = Replace(strסԺ, "������ü�¼", "סԺ���ü�¼")
        gstrSQL = gstrSQL & " Union All " & strסԺ
    
        gstrSQL = gstrSQL & " Order by ���ȼ�,����,No"
        
        Set recAutoPrint = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, _
            mSQLCondition.lngҩ��ID, _
            mSQLCondition.date��ʼ����, _
            mSQLCondition.date��������)
        
        datCurr = Sys.Currentdate()
        
        With recAutoPrint
            If .RecordCount > 0 Then
                If DateDiff("s", !��������, datCurr) > mParams.lngPrintDelay Then
                    Do While Not .EOF
                        '���´�ӡ״̬
                        gstrSQL = "Zl_δ��ҩƷ��¼_���´�ӡ״̬("
                        '����
                        gstrSQL = gstrSQL & Val(!����)
                        'NO
                        gstrSQL = gstrSQL & ",'" & !NO & "'"
                        '�ⷿID
                        gstrSQL = gstrSQL & "," & mParams.lngҩ��ID
                        '��Դ����
                        gstrSQL = gstrSQL & ",Null"
                        '��ӡ����
                        gstrSQL = gstrSQL & ",1"
                        gstrSQL = gstrSQL & ")"
                        
                        Call zldatabase.ExecuteProcedure(gstrSQL, Me.Caption & "-���µ����Ѵ�ӡ")
                        
                        '��ӡ����
                        Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1723", Me, "�ⷿ==" & mParams.lngҩ��ID, "NO=" & !NO, "����=" & Val(!����), "�����=����� is null", 2)
                        
                        .MoveNext
                    Loop
                End If
            End If
        End With
    End If
    '���ϴ�ӡ���ķ����嵥
    
    gstrSQL = " Select  NO,����,��������,1 As ����,Nvl(��������,0) As ��������,Nvl(��ӡ״̬,0) As ��ӡ״̬,A.����ID, a.���ȼ�, a.���� " & _
               " From δ��ҩƷ��¼ A, ������Ϣ B "
    gstrSQL = gstrSQL & " Where �ⷿID+0=[1] " & _
               " And �������� Between [2] And [3] " & _
               " And ��ӡ״̬ Not In (1,2) "
    
    gstrSQL = gstrSQL & IIf(mParams.blnMustDosageOkProcess, " And A.�Ŷ�״̬=1 ", "")
    gstrSQL = gstrSQL & IIf(mParams.bln�Զ���ҩ, " And ��ҩ�� Is Null ", "")
    gstrSQL = gstrSQL & " And A.����ID=B.����ID" & IIf(mSQLCondition.strҽ���� <> "", "", "(+)") & ""
    
    Select Case mParams.intPrint
        Case 0
            If mParams.intPrintDrugLable = 0 Then Exit Function
        Case 1
            If Not mParams.bln���ʵ� Then gstrSQL = gstrSQL & " And ����=8"
        Case 2
            If mParams.bln���ʵ� Then
                If mParams.Str���� <> "" Then
                    gstrSQL = gstrSQL & " And A.���� In (8,9) And A.��ҩ���� In (Select * From Table(Cast(f_Str2list([19]) As Zltools.t_Strlist))) "
                End If
            Else
                If mParams.Str���� <> "" Then
                    gstrSQL = gstrSQL & " And A.����=8 And A.��ҩ���� In (Select * From Table(Cast(f_Str2list([19]) As Zltools.t_Strlist))) "
                Else
                    gstrSQL = gstrSQL & " And A.����=8"
                End If
            End If
        Case 3
            If mParams.bln���ʵ� Then
                If mParams.strPrintWindow <> "" Then
                    gstrSQL = gstrSQL & " And A.���� In (8,9) And A.��ҩ���� In (Select * From Table(Cast(f_Str2list([20]) As Zltools.t_Strlist))) "
                End If
            Else
                If mParams.strPrintWindow <> "" Then
                    gstrSQL = gstrSQL & " And A.����=8 And A.��ҩ���� In (Select * From Table(Cast(f_Str2list([20]) As Zltools.t_Strlist))) "
                Else
                    gstrSQL = gstrSQL & " And A.����=8"
                End If
            End If
    End Select
    
    If mcondition.int������� = 2 Then
        gstrSQL = gstrSQL & " And A.���� = 9 And A.��ҳID Is Not NULL " '��סԺ����
    Else
        gstrSQL = gstrSQL & " And A.���� In (8,9)" '���ＰסԺ���е���
    End If
    
    If mSQLCondition.str��ʼNO <> "" Or mSQLCondition.str����NO <> "" Then
        If mSQLCondition.str��ʼNO <> "" And mSQLCondition.str����NO <> "" Then
            gstrSQL = gstrSQL & " And A.NO Between [4] And [5] "
        Else
            If mSQLCondition.str��ʼNO <> "" Then
                gstrSQL = gstrSQL & " And A.NO = [4] "
            Else
                gstrSQL = gstrSQL & " And A.NO = [5] "
            End If
        End If
    End If
    
    If mSQLCondition.str���� <> "" Then gstrSQL = gstrSQL & " And Upper(A.����) Like [6] "
    
    If mSQLCondition.str���￨ <> "" Then gstrSQL = gstrSQL & " And Upper(B.���￨��) = [7] "
    
    If mSQLCondition.str��ʶ�� <> "" Then gstrSQL = gstrSQL & " And Upper(DECODE(A.����,8,B.�����,B.סԺ��)) Like [8] "
    
    If mSQLCondition.lng����ID > 0 Then gstrSQL = gstrSQL & " And A.�Է�����ID+0=[9] "
    
    If mSQLCondition.str��ǰNO <> "" Then gstrSQL = gstrSQL & " And A.NO=[13] "
    
    If mSQLCondition.str����� <> "" Then gstrSQL = gstrSQL & " And B.�����=[14] "
    
'    If mSQLCondition.str���֤ <> "" Then gstrSQL = gstrSQL & " And B.���֤��=[15] "

    If mSQLCondition.str���֤ <> "" Then gstrSQL = gstrSQL & " And B.����ID=[15] "
    
    If mSQLCondition.lng����ID <> 0 Or (Me.txtPati.Text <> "" And mParams.int����ģʽ = mFindType.IC��) Then gstrSQL = gstrSQL & " And B.����id=[16] "
    
    If mSQLCondition.strҽ���� <> "" Then gstrSQL = gstrSQL & " And B.ҽ����=[17] "
    
    If mSQLCondition.lngסԺ�� <> 0 Then gstrSQL = gstrSQL & " And B.סԺ��=[18] "
    
    Select Case mParams.intShowBill�շ�
        Case 0  '����ʾ����
            strSub1 = "1=2"
        Case 1  '��ʾδ�շ�
            strSub1 = "A.����<>9 And Nvl(A.���շ�,0)=0 And A.����=8"
        Case 2  '��ʾ���շ�
            strSub1 = "A.����<>9 And A.���շ�=1 And A.����=8"
        Case 3  '��ʾ���д���
            strSub1 = "A.����<>9 And A.����=8"
    End Select
    Select Case mParams.intShowBill����
        Case 0  '����ʾ����
            strSub2 = "1=2"
        Case 1  '��ʾδ���
            strSub2 = "A.����<>8 And Nvl(A.���շ�,0)=0 And A.����=9"
        Case 2  '��ʾ�����
            strSub2 = "A.����<>8 And A.���շ�=1 And A.����=9"
        Case 3  '��ʾ���д���
            strSub2 = "A.����<>8 And A.����=9"
    End Select
    
    gstrSQL = gstrSQL & " And A.���� IN(8,9) And (" & strSub1 & " Or " & strSub2 & ")"
    
    gstrSQL = gstrSQL & IIf(mParams.strSourceDep = "", "", " And A.�Է�����id+0 In (Select * From Table(Cast(f_Num2list([21]) As Zltools.t_Numlist)))")
    
    '�ų��쳣����
    gstrSQL = gstrSQL & " And Not Exists (Select 1 From ������ü�¼ C Where c.No = a.No And c.ִ�в���id = a.�ⷿid And c.ִ��״̬ = 9) "
    
    '�������ü�¼���ж������סԺ
    gstrSQL = gstrSQL & " And Exists (Select 1 From ������ü�¼ C Where a.No = c.No And a.�ⷿid = c.ִ�в���id And Decode(a.����, 8, 1, 9, 2) = c.��¼����) "
    
    '���������סԺ���ü�¼��ѯ
    If mcondition.int������� = 2 Then
        '��סԺ
        gstrSQL = Replace(gstrSQL, "1 As ����", "2 As ����")
        gstrSQL = Replace(gstrSQL, "������ü�¼", "סԺ���ü�¼")
    ElseIf mcondition.int������� = 3 Then
        '�����סԺ
        strסԺ = gstrSQL
        strסԺ = Replace(strסԺ, "1 As ����", "2 As ����")
        strסԺ = Replace(strסԺ, "������ü�¼", "סԺ���ü�¼")
        gstrSQL = gstrSQL & " Union All " & strסԺ
    End If
    
    gstrSQL = gstrSQL & " Order by ���ȼ�,����,No"
    
    On Error GoTo ErrHand
    BlnInRefresh = True
    
    recAutoPrint.Sort = "����id"
    
    If mSQLCondition.str���֤ <> "" And Not mobjSquareCard Is Nothing Then
        If mobjSquareCard.zlGetPatiID("���֤", UCase(mSQLCondition.str���֤), False, lng����ID) = False Then lng����ID = 0
        mSQLCondition.lng����ID = lng����ID
    End If
    
    Set recAutoPrint = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, _
            mSQLCondition.lngҩ��ID, _
            mSQLCondition.date��ʼ����, _
            mSQLCondition.date��������, _
            mSQLCondition.str��ʼNO, _
            mSQLCondition.str����NO, _
            mSQLCondition.str����, _
            mSQLCondition.str���￨, _
            mSQLCondition.str��ʶ��, _
            mSQLCondition.lng����ID, _
            mSQLCondition.str������, _
            mSQLCondition.str�����, _
            mSQLCondition.lngҩƷid, _
            mSQLCondition.str��ǰNO, _
            mSQLCondition.str�����, _
            lng����ID, _
            mSQLCondition.lng����ID, _
            mSQLCondition.strҽ����, _
            mSQLCondition.lngסԺ��, _
            mParams.Str����, _
            mParams.strPrintWindow, _
            mParams.strSourceDep)

    datCurr = Sys.Currentdate()
        
    With recAutoPrint
        Do While Not .EOF
            '��ӡ����
            If DateDiff("s", !��������, datCurr) > mParams.lngPrintDelay Then
                If mParams.intPrint > 0 Then
                    If mParams.bln�Զ���ҩ = True And IsDosage(Val(!����), !NO, Val(!����)) Then
                        '�����Զ���ҩ���ڴ�ӡǰ���
                        blnIgnore = False

'                        '����Ƿ���Ҫ��ҩ
'                        If Not IsDosage(Val(!����), !NO, Val(!����)) Then
'                            blnIgnore = True
'                        End If

                        '����Ƿ�����
                        If CheckBill(mSQLCondition.lngҩ��ID, 1, Val(!����), !NO, Val(!����), Val(!����)) <> 0 Then
                            blnIgnore = True
                        End If

                        If blnIgnore = False Then
                            gcnOracle.BeginTrans
                            blnInTrans = True

                            '��������ҩ��
                            str����Ա = IIf(mstr�Զ���ҩ�� <> "", mstr�Զ���ҩ��, IIf(mParams.str��ҩ�� = "|��ǰ����Ա|", gstrUserName, mParams.str��ҩ��))

                            gstrSQL = "zl_ҩƷ�շ���¼_������ҩ��("
                            '�ⷿID
                            gstrSQL = gstrSQL & mParams.lngҩ��ID
                            '����
                            gstrSQL = gstrSQL & "," & Val(!����)
                            'NO
                            gstrSQL = gstrSQL & ",'" & !NO & "'"
                            '����
                            gstrSQL = gstrSQL & "," & Val(!����)
                            '��ҩ��
                            gstrSQL = gstrSQL & ",'" & str����Ա & "'"
                            '��ҩ����
                            gstrSQL = gstrSQL & ",to_date('" & datCurr & "','yyyy-MM-dd hh24:mi:ss') "
                            gstrSQL = gstrSQL & ")"

                            Call zldatabase.ExecuteProcedure(gstrSQL, Me.Caption & "-������ҩ��")
                            
                            If Val(!��ӡ״̬) <> 3 Then
                                gstrSQL = "Zl_δ��ҩƷ��¼_���´�ӡ״̬("
                                '����
                                gstrSQL = gstrSQL & Val(!����)
                                'NO
                                gstrSQL = gstrSQL & ",'" & !NO & "'"
                                '�ⷿID
                                gstrSQL = gstrSQL & "," & mParams.lngҩ��ID
                                '��Դ����
                                gstrSQL = gstrSQL & "," & IIf(mParams.strSourceDep = "", "Null", "'" & mParams.strSourceDep & "'")
                                '��ӡ����
                                gstrSQL = gstrSQL & ",3"
                                gstrSQL = gstrSQL & ")"
                                
                                Call zldatabase.ExecuteProcedure(gstrSQL, Me.Caption & "-���µ����Ѵ�ӡ")
                                
                                strUnit = GetUnit(mParams.lngҩ��ID, !����, !NO, Val(!����))
                                str�շ����� = BillHaveHerial(!NO, !����, Val(!����))
                                
                                If mParams.bln��ӡ���и�ʽ Then
                                    If InStr(1, str�շ�����, "7") <> 0 And (InStr(1, str�շ�����, "5") <> 0 Or InStr(1, str�շ�����, "6") <> 0) Then
                                        SetLocatePrinter !��������
                                        Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1341_3", Me, _
                                            "NO=" & !NO, "����=" & IIf(!���� = 8, 1, 2), "ҩ��=" & mParams.lngҩ��ID, "��װϵ��=" & IIf(strUnit = "���ﵥλ", "C.�����װ", "C.סԺ��װ"), "PrintEmpty=0", 2)
                                        
                                        '�ָ�����ǩ�ı��ش�ӡ������
                                        Call SavePrinterSet("ZL1_BILL_1341_3", mstrRPTDefaultScheme_Recipt, mParams.strDefaultPrinter)
                                                                        
                                        Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1341_4", Me, _
                                            "NO=" & !NO, "����=" & IIf(!���� = 8, 1, 2), "PrintEmpty=0", 2)
                                    ElseIf (InStr(1, str�շ�����, "5") <> 0 Or InStr(1, str�շ�����, "6") <> 0) Then
                                        SetLocatePrinter !��������
                                        Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1341_3", Me, _
                                            "NO=" & !NO, "����=" & IIf(!���� = 8, 1, 2), "ҩ��=" & mParams.lngҩ��ID, "��װϵ��=" & IIf(strUnit = "���ﵥλ", "C.�����װ", "C.סԺ��װ"), "PrintEmpty=0", 2)
                                        
                                        '�ָ�����ǩ�ı��ش�ӡ������
                                        Call SavePrinterSet("ZL1_BILL_1341_3", mstrRPTDefaultScheme_Recipt, mParams.strDefaultPrinter)
                                    Else
                                        Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1341_4", Me, _
                                            "NO=" & !NO, "����=" & IIf(!���� = 8, 1, 2), "PrintEmpty=0", 2)
                                    End If

                                Else
                                    If InStr(1, str�շ�����, "7") <> 0 And (InStr(1, str�շ�����, "5") <> 0 Or InStr(1, str�շ�����, "6") <> 0) Then
                                        SetLocatePrinter !��������, Val(Split(mParams.str��ҩ��ʽ, ";")(0)) - 1
                                        Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1341_3", Me, _
                                            "NO=" & !NO, "����=" & IIf(!���� = 8, 1, 2), "ҩ��=" & mParams.lngҩ��ID, "��װϵ��=" & IIf(strUnit = "���ﵥλ", "C.�����װ", "C.סԺ��װ"), "ReportFormat=" & Val(Split(mParams.str��ҩ��ʽ, ";")(0)), "PrintEmpty=0", 2)
                                        
                                        '�ָ�����ǩ�ı��ش�ӡ������
                                        Call SavePrinterSet("ZL1_BILL_1341_3", mstrRPTDefaultScheme_Recipt, mParams.strDefaultPrinter)
                                                                        
                                        Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1341_4", Me, _
                                            "NO=" & !NO, "����=" & IIf(!���� = 8, 1, 2), "ReportFormat=" & Val(Split(mParams.str��ҩ��ʽ, ";")(1)), "PrintEmpty=0", 2)
                                    ElseIf (InStr(1, str�շ�����, "5") <> 0 Or InStr(1, str�շ�����, "6") <> 0) Then
                                        SetLocatePrinter !��������, Val(Split(mParams.str��ҩ��ʽ, ";")(0)) - 1
                                        Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1341_3", Me, _
                                            "NO=" & !NO, "����=" & IIf(!���� = 8, 1, 2), "ҩ��=" & mParams.lngҩ��ID, "��װϵ��=" & IIf(strUnit = "���ﵥλ", "C.�����װ", "C.סԺ��װ"), "ReportFormat=" & Val(Split(mParams.str��ҩ��ʽ, ";")(0)), "PrintEmpty=0", 2)
                                        
                                        '�ָ�����ǩ�ı��ش�ӡ������
                                        Call SavePrinterSet("ZL1_BILL_1341_3", mstrRPTDefaultScheme_Recipt, mParams.strDefaultPrinter)
                                    Else
                                        Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1341_4", Me, _
                                            "NO=" & !NO, "����=" & IIf(!���� = 8, 1, 2), "ReportFormat=" & Val(Split(mParams.str��ҩ��ʽ, ";")(1)), "PrintEmpty=0", 2)
                                    End If
                                End If
                            End If

                            '����������˵���ǩ��������Ҫ����ҩ�˽��е���ǩ������
                            If gblnESign������ҩ = True And gblnESignUserStoped = False Then
                                strǩ����¼ = ""
                                If GetSignatureRecored(EsignTache.Dosage, Val(!����), !NO, mParams.lngҩ��ID, strǩ����¼) = False Then
                                    If blnInTrans = True Then gcnOracle.RollbackTrans
                                    Exit Function
                                End If
                                
                                If strǩ����¼ <> "" Then
                                    gstrSQL = "Zl_ҩƷǩ����¼_Insert(" & strǩ����¼ & ")"
                                    
                                    Call zldatabase.ExecuteProcedure(gstrSQL, "ǩ����¼")
                                End If
                            End If

                            gcnOracle.CommitTrans
                            blnInTrans = False

                            mblnIsFirst = False
                        End If
                    ElseIf Val(!��ӡ״̬) <> 3 Then
                        gstrSQL = "Zl_δ��ҩƷ��¼_���´�ӡ״̬("
                        '����
                        gstrSQL = gstrSQL & Val(!����)
                        'NO
                        gstrSQL = gstrSQL & ",'" & !NO & "'"
                        '�ⷿID
                        gstrSQL = gstrSQL & "," & mParams.lngҩ��ID
                        '��Դ����
                        gstrSQL = gstrSQL & "," & IIf(mParams.strSourceDep = "", "Null", "'" & mParams.strSourceDep & "'")
                        '��ӡ����
                        gstrSQL = gstrSQL & ",3"
                        gstrSQL = gstrSQL & ")"
                        
                        Call zldatabase.ExecuteProcedure(gstrSQL, Me.Caption & "-���µ����Ѵ�ӡ")
                        
                        strUnit = GetUnit(mParams.lngҩ��ID, !����, !NO, Val(!����))
                        str�շ����� = BillHaveHerial(!NO, !����, Val(!����))
                        
                        If mParams.bln��ӡ���и�ʽ Then
                            If InStr(1, str�շ�����, "7") <> 0 And (InStr(1, str�շ�����, "5") <> 0 Or InStr(1, str�շ�����, "6") <> 0) Then
                                SetLocatePrinter !��������
                                Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1341_3", Me, _
                                    "NO=" & !NO, "����=" & IIf(!���� = 8, 1, 2), "ҩ��=" & mParams.lngҩ��ID, "��װϵ��=" & IIf(strUnit = "���ﵥλ", "C.�����װ", "C.סԺ��װ"), "PrintEmpty=0", 2)
                                
                                '�ָ�����ǩ�ı��ش�ӡ������
                                Call SavePrinterSet("ZL1_BILL_1341_3", mstrRPTDefaultScheme_Recipt, mParams.strDefaultPrinter)
                                                        
                                Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1341_4", Me, _
                                    "NO=" & !NO, "����=" & IIf(!���� = 8, 1, 2), "PrintEmpty=0", 2)
                            ElseIf (InStr(1, str�շ�����, "5") <> 0 Or InStr(1, str�շ�����, "6") <> 0) Then
                                SetLocatePrinter !��������, -1
                                Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1341_3", Me, _
                                    "NO=" & !NO, "����=" & IIf(!���� = 8, 1, 2), "ҩ��=" & mParams.lngҩ��ID, "��װϵ��=" & IIf(strUnit = "���ﵥλ", "C.�����װ", "C.סԺ��װ"), "PrintEmpty=0", 2)
                                
                                '�ָ�����ǩ�ı��ش�ӡ������
                                Call SavePrinterSet("ZL1_BILL_1341_3", mstrRPTDefaultScheme_Recipt, mParams.strDefaultPrinter)
                            Else
                                Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1341_4", Me, _
                                    "NO=" & !NO, "����=" & IIf(!���� = 8, 1, 2), "PrintEmpty=0", 2)
                            End If

                        Else
                        
                            If InStr(1, str�շ�����, "7") <> 0 And (InStr(1, str�շ�����, "5") <> 0 Or InStr(1, str�շ�����, "6") <> 0) Then
                                SetLocatePrinter !��������, Val(Split(mParams.str��ҩ��ʽ, ";")(0)) - 1
                                Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1341_3", Me, _
                                    "NO=" & !NO, "����=" & IIf(!���� = 8, 1, 2), "ҩ��=" & mParams.lngҩ��ID, "��װϵ��=" & IIf(strUnit = "���ﵥλ", "C.�����װ", "C.סԺ��װ"), "ReportFormat=" & Val(Split(mParams.str��ҩ��ʽ, ";")(0)), "PrintEmpty=0", 2)
                                
                                '�ָ�����ǩ�ı��ش�ӡ������
                                Call SavePrinterSet("ZL1_BILL_1341_3", mstrRPTDefaultScheme_Recipt, mParams.strDefaultPrinter)
                                                        
                                Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1341_4", Me, _
                                    "NO=" & !NO, "����=" & IIf(!���� = 8, 1, 2), "ReportFormat=" & Val(Split(mParams.str��ҩ��ʽ, ";")(1)), "PrintEmpty=0", 2)
                            ElseIf (InStr(1, str�շ�����, "5") <> 0 Or InStr(1, str�շ�����, "6") <> 0) Then
                                SetLocatePrinter !��������, Val(Split(mParams.str��ҩ��ʽ, ";")(0)) - 1
                                Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1341_3", Me, _
                                    "NO=" & !NO, "����=" & IIf(!���� = 8, 1, 2), "ҩ��=" & mParams.lngҩ��ID, "��װϵ��=" & IIf(strUnit = "���ﵥλ", "C.�����װ", "C.סԺ��װ"), "ReportFormat=" & Val(Split(mParams.str��ҩ��ʽ, ";")(0)), "PrintEmpty=0", 2)
                                
                                '�ָ�����ǩ�ı��ش�ӡ������
                                Call SavePrinterSet("ZL1_BILL_1341_3", mstrRPTDefaultScheme_Recipt, mParams.strDefaultPrinter)
                            Else
                                Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1341_4", Me, _
                                    "NO=" & !NO, "����=" & IIf(!���� = 8, 1, 2), "ReportFormat=" & Val(Split(mParams.str��ҩ��ʽ, ";")(1)), "PrintEmpty=0", 2)
                            End If
                        End If
                    End If
                End If
                
                '���ǰ��û���ж��Ƿ�����ҩ�����������ٴ���
                If mParams.intPrint <= 0 Then
                    str�շ����� = BillHaveHerial(!NO, !����, Val(!����))
                End If
                
                '��ӡҩƷ��ǩ
                If mParams.intPrintDrugLable = 1 And Val(!��ӡ״̬) <> 4 Then
                    gstrSQL = "Zl_δ��ҩƷ��¼_���´�ӡ״̬("
                    '����
                    gstrSQL = gstrSQL & Val(!����)
                    'NO
                    gstrSQL = gstrSQL & ",'" & !NO & "'"
                    '�ⷿID
                    gstrSQL = gstrSQL & "," & mParams.lngҩ��ID
                    '��Դ����
                    gstrSQL = gstrSQL & "," & IIf(mParams.strSourceDep = "", "Null", "'" & mParams.strSourceDep & "'")
                    '��ӡ����
                    gstrSQL = gstrSQL & ",4"
                    gstrSQL = gstrSQL & ")"
                    
                    Call zldatabase.ExecuteProcedure(gstrSQL, Me.Caption & "-���µ����Ѵ�ӡ")
                        
                    If InStr(1, str�շ�����, "7") <> 0 And (InStr(1, str�շ�����, "5") <> 0 Or InStr(1, str�շ�����, "6") <> 0) Then
                        Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1341_6", Me, _
                            "NO=" & !NO, "����=" & IIf(!���� = 8, 1, 2), "ҩ��=" & mParams.lngҩ��ID, "��װϵ��=" & IIf(strUnit = "���ﵥλ", "D.�����װ", "D.סԺ��װ"), "PrintEmpty=0", 2)
                        
                        Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1341_7", Me, _
                            "NO=" & !NO, "ҩ��=" & mParams.lngҩ��ID, "����=" & IIf(!���� = 8, 1, 2), "PrintEmpty=0", 2)
                    ElseIf (InStr(1, str�շ�����, "5") <> 0 Or InStr(1, str�շ�����, "6") <> 0) Then
                        Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1341_6", Me, _
                            "NO=" & !NO, "����=" & IIf(!���� = 8, 1, 2), "ҩ��=" & mParams.lngҩ��ID, "��װϵ��=" & IIf(strUnit = "���ﵥλ", "D.�����װ", "D.סԺ��װ"), "PrintEmpty=0", 2)
                    Else
                        Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1341_7", Me, _
                            "NO=" & !NO, "ҩ��=" & mParams.lngҩ��ID, "����=" & IIf(!���� = 8, 1, 2), "PrintEmpty=0", 2)
                    End If
                End If
            End If

            .MoveNext
        Loop
        If .RecordCount > 0 Then .MoveFirst
    End With
    BlnInRefresh = False
    Exit Function
ErrHand:
    If blnInTrans Then gcnOracle.RollbackTrans
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub GetOpr()
'��ȡ��ҩ���ڵı��
    Dim strsql As String
    Dim rstemp As Recordset
    
    On Error GoTo errHandle
    strsql = "Select ���� From ��ҩ���� Where ҩ��id=[1] And ����=[2]"
    Set rstemp = zldatabase.OpenSQLRecord(strsql, "GetOpr", mParams.lngҩ��ID, mParams.Str����)
    
    If Not rstemp.EOF Then
        mstrOpr = rstemp!����
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function IsDosage(ByVal Int���� As Integer, ByVal strNo As String, ByVal int���� As Integer) As Boolean
    '��鵱ǰ�����Ƿ���Ҫ������ҩ����
    
    On Error GoTo ErrHand
    
    If Int���� = 0 Then Exit Function
    If strNo = "" Then Exit Function
    
    If mrsIsDosage Is Nothing Then
        GetDosage mParams.lngҩ��ID
        If mrsIsDosage Is Nothing Then
            Exit Function
        End If
    End If
    
    mrsIsDosage.Filter = "����=" & int����
    If mrsIsDosage.EOF Then Exit Function

    IsDosage = (mrsIsDosage!��ҩ = 1)
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub TimePrintCancelBill_Timer()
    Dim curDateBegin As Date
    Dim curDateEnd As Date
    
    '���ô�ӡ�˷ѵ�
    IntTimes = IntTimes + 1
    '�����������˳�
    If IntTimes < mParams.lngPrintBackInterval Then Exit Sub
    IntTimes = 0
    
    curDateEnd = Format(Sys.Currentdate, "yyyy-MM-dd hh:mm:ss")
    curDateBegin = DateAdd("n", 0 - mParams.lngPrintBackInterval, curDateEnd)
    
    Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1341_8", Me, "��ʼʱ��=" & Format(curDateBegin, "yyyy-MM-dd hh:mm"), "����ʱ��=" & Format(curDateEnd, "yyyy-MM-dd hh:mm"), "ҩ��=" & mParams.lngҩ��ID, 2)
End Sub

Private Sub TimeRefresh_Timer()
    '�����Զ�ˢ��δ֪����
    Dim thwnd As Long
    
    On Error Resume Next
    
    If Me.WindowState = vbMinimized Then Exit Sub
    
    '�������ڲ��ǵ�ǰ����ʱ�˳�
    If InStr(1, "frmҩƷ������ҩNew;frm������ҩ��ϸ;frm������ҩ�б�;frm����", Screen.ActiveForm.Name) = 0 Then Exit Sub
'    If Screen.ActiveForm.hWnd <> Me.hWnd Then Exit Sub

    thwnd = GetForegroundWindow()
    If thwnd <> Me.hWnd Then Exit Sub
    
    If mcondition.intListType = mListType.��ҩ Then Exit Sub
    
    '�����ҩδ�������˳�
    If mblnSendIsOver = False Then Exit Sub
    
    '�����Ϣ������Ч��ͨ����ѯ��ʽ�Զ�ˢ��
    If Not mobjMipModule Is Nothing Then
        If mobjMipModule.IsConnect = True Then
            '��ǰ�Ǵ���ҩ�����ҩ����ʱ��ִ����ѯˢ��
            If (mParams.blnMustDosageProcess = True And mcondition.intListType = mListType.����ҩ) Or _
                (mParams.blnMustDosageProcess = False And mcondition.intListType = mListType.����ҩ) Then
                Exit Sub
            End If
        End If
    End If
    
    TimeRefresh.Enabled = False
    DoEvents
        RefreshList mcondition.intListType
    DoEvents
    TimeRefresh.Enabled = True
End Sub

Private Sub tmrCall_Timer()
    '������Ϊ����������Զ�˺��л���ʹ��ʱ���ſ���ʱ��ؼ�
    
    '���ú���������
    Call zlCallMain
End Sub

Private Sub tmrMsgRefresh_Timer()
    If Me.WindowState = vbMinimized Then Exit Sub
    
    If Screen.ActiveForm Is Nothing Then Exit Sub
    '�������ڲ��ǵ�ǰ����ʱ�˳�
    If InStr(1, "frmҩƷ������ҩNew;frm������ҩ��ϸ;frm������ҩ�б�;frm����", Screen.ActiveForm.Name) = 0 Then Exit Sub
'    If Screen.ActiveForm.hWnd <> Me.hWnd Then Exit Sub
    
    '��ҩҵ��ʱ�������Զ�ˢ�»��ӡ
    If mcondition.intListType = mListType.��ҩ Then Exit Sub
    
    '��Ϣ������Чʱ������
    If mobjMipModule Is Nothing Then Exit Sub
    If mobjMipModule.IsConnect = False Then Exit Sub
     
    '����Ϣʱˢ�»��ӡ
    If mblnExistMsg = True Then
        'ˢ��ǰ�ȹرռ�ʱ����ˢ������ٿ���
        tmrMsgRefresh.Enabled = False
        DoEvents
        
        '��ǰ�Ǵ���ҩ�����ҩ������ˢ��
        If (mParams.blnMustDosageProcess = True And mcondition.intListType = mListType.����ҩ) Or _
            (mParams.blnMustDosageProcess = False And mcondition.intListType = mListType.����ҩ) Then
            Call RefreshList(mcondition.intListType)
        End If
                
        'ͬʱҲ�����Զ���ӡ
        If mParams.intPrint > 0 Then
            Call AutoPrint
        End If
        
        DoEvents
        tmrMsgRefresh.Enabled = True
                
        'ˢ�º��¼��ǰˢ��ʱ��
        mdteMsgRefresh = Now
        
        '��Ϣ������ΪFalse
        mblnExistMsg = False
    End If
End Sub

Private Sub txtPati_Change()
    If Not mobjIDCard Is Nothing Then mobjIDCard.SetEnabled (txtPati.Text = "" And Me.ActiveControl Is txtPati)
    If Not mobjICCard Is Nothing Then mobjICCard.SetEnabled (txtPati.Text = "" And Me.ActiveControl Is txtPati)
    
End Sub


Private Sub txtPati_GotFocus()
    txtPati.BackColor = &HE1FEDA
    
    If Not mobjIDCard Is Nothing And txtPati.Text = "" Then
        Call mobjIDCard.SetEnabled(True)
    End If
    
    If Not mobjICCard Is Nothing And txtPati.Text = "" Then
        Call mobjICCard.SetEnabled(True)
    End If
    
    txtPati.Text = ""
    Call zlControl.TxtSelAll(txtPati)
    
    mblnInput = True
End Sub

Private Sub txtPati_KeyPress(KeyAscii As Integer)
    Dim blnDoIt As Boolean
    Dim strInput As String
    Dim strCondition As String
    Dim i As Integer
    Dim blnˢ�� As Boolean
    Dim blnSta As Boolean
    Dim lng����ID As Long
    Dim rsData As Recordset
    Dim str���� As String
    Dim str����id As String
    Dim strCard As String
    Dim strRecipeString As String
    Dim arrRecipe
    Dim intCount As Integer, n As Integer
    Dim strNos As String
    Dim strReturn As String
    
    strCard = IDKNType.GetCurCard.����
    If KeyAscii = 13 Then
        KeyAscii = 0
        blnDoIt = True
        
        If strCard = "IC��" Then
            If Not mobjSquareCard Is Nothing Then Call mobjSquareCard.zlGetPatiID("IC��", UCase(Trim(txtPati.Text)), False, mlngIC����id)
            If txtPati.Text <> "" Then blnDoIt = True
        Else
            If Trim(txtPati.Text) <> "" Then blnDoIt = True
        End If
        
        If Not (strCard = "����" Or strCard = "���ݺ�" Or strCard = "סԺ��" Or strCard = "ҽ����" Or strCard = "���֤" Or strCard = "�����") And KeyAscii <> 8 Then blnˢ�� = True
    ElseIf KeyAscii <> 13 Then
        mblnCard = False
        mblnScaner = False
        If strCard = "����" Then
            '�������
            mblnCard = zlCommFun.InputIsCard(txtPati, KeyAscii, glngSys)
        ElseIf mcondition.intListType = mListType.����ҩ And mParams.int����ģʽ = mFindType.���ݺ� And mParams.bln��ҩɨ�� = True Then
            mblnScaner = InputIsScaner(txtPati, KeyAscii)
        ElseIf mcondition.intListType = mListType.����ҩ And mParams.blnɨ������ = True Then
            mblnScaner = InputIsScaner(txtPati, KeyAscii)
        Else
            mblnScaner = InputIsScaner(txtPati, KeyAscii)
        End If
        
        If mblnCard Then
            If strCard = "����" Then
                If Len(txtPati.Text) = mint���￨���� - 1 And KeyAscii <> 8 And txtPati.SelLength <> Len(txtPati.Text) Then
                    txtPati.Text = txtPati.Text & Chr(KeyAscii)
                    txtPati.SelStart = Len(txtPati.Text)
                    KeyAscii = 0: blnDoIt = True
                End If
            End If
        Else
            Select Case strCard
'                Case mFindType.���￨
'                    If InStr(":��;��?��''||" & Chr(22), Chr(KeyAscii)) > 0 Then
'                        KeyAscii = 0
'                    Else
'                        KeyAscii = Asc(UCase(Chr(KeyAscii)))
'                    End If
                Case "�����"
                    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
                Case "���ݺ�"
                    KeyAscii = Asc(UCase(Chr(KeyAscii)))
                    If mcondition.intListType = mListType.����ҩ And mblnScaner And mParams.bln��ҩɨ�� = True Then
                        txtPati.Text = txtPati.Text & UCase(Chr(KeyAscii))
                        txtPati.SelStart = Len(txtPati.Text)
                        KeyAscii = 0
                        
                        If Len(txtPati.Text) = 8 Then
                            blnDoIt = True
                            If mstrScanerLastNo <> txtPati.Text Then
                                mblnScaned = False
                                mstrScanerLastNo = txtPati.Text
                            End If
                        End If
                    Else
                        If Not (InStr("0123456789" & Chr(8), Chr(KeyAscii)) > 0 Or (KeyAscii >= Asc("A") And KeyAscii <= Asc("Z"))) Then
                            KeyAscii = 0
                        End If
                    End If
                Case "����"
                    If InStr(":��;��?��", Chr(KeyAscii)) > 0 Then
                        KeyAscii = 0
                    End If
'                Case "�������֤"
'                Case "IC��"
                Case "סԺ��"
                Case "ҽ����"
                Case Else
                    blnˢ�� = True
'                    Me.txtPati.MaxLength = 100
                    '�����������ѿ�
                    If InStr(":��;��?��''||" & Chr(22) & Chr(32), Chr(KeyAscii)) > 0 Then
                        KeyAscii = 0
                    Else
                        KeyAscii = Asc(UCase(Chr(KeyAscii)))
                    End If
                    
                    If Len(txtPati.Text) = IDKNType.GetCardNoLen - 1 And KeyAscii <> 8 And mParams.int�س���ʽ = 1 Then
                        txtPati.Text = txtPati.Text & Chr(KeyAscii)
                        txtPati.SelStart = Len(txtPati.Text)
                        KeyAscii = 0
                        blnDoIt = True
                    End If
                    mstrLastBrushCardNo = txtPati.Text & IIf(KeyAscii = 0, "", Chr(KeyAscii))
            End Select
        End If
    End If
    
    If blnDoIt Then
        If mParams.blnMustDosageOkProcess And blnˢ�� And InStr(1, mstrPrivs, "��ҩȷ��") > 0 Then
            On Error GoTo errHandle
            
            If strCard = "����" Or strCard = "���ݺ�" Or strCard = "סԺ��" Or strCard = "ҽ����" Or strCard = "���֤" Or strCard = "�����" Then
                strInput = txtPati.Text
            Else
                '���ѿ����ʱ����Ϊ��ID+����
                strInput = mobjcard.�ӿ���� & "|" & txtPati.Text
            End If
            lng����ID = zlfuncCard_GetPatiID(mobjSquareCard, Val(Split(strInput, "|")(0)), Split(strInput, "|")(1))
            
            If lng����ID <> 0 Then
                gstrSQL = "Select distinct A.NO,A.���� From δ��ҩƷ��¼ A,ҩƷ�շ���¼ B Where A.NO=B.NO and A.����=B.���� and A.�ⷿid=B.�ⷿid and A.����id=[1] And A.�ⷿid=[2] and A.�������� between [3] and [4] and nvl(A.�Ŷ�״̬,0)=0"
                Set rsData = zldatabase.OpenSQLRecord(gstrSQL, "", lng����ID, mParams.lngҩ��ID, mSQLCondition.date��ʼ����, mSQLCondition.date��������)
                strCondition = ""
                
                If Not rsData Is Nothing Then
                    
                    If rsData.RecordCount > 0 Then
                        rsData.MoveFirst
                    Else
                        
                        Me.stbThis.Panels(2).Text = "���ţ�" & txtPati.Text & "���޴�����Ϣ��"
                        blnSta = True
                    End If
                    gcnOracle.BeginTrans
                    Do While Not rsData.EOF
                        strCondition = IIf(strCondition = "", strCondition, strCondition & " OR ") & "NO='" & rsData!NO & "'"
                        gstrSQL = "Zl_δ��ҩƷ��¼_��ҩȷ��("
                            'NO
                            gstrSQL = gstrSQL & "'" & rsData!NO & "'"
                            '����
                            gstrSQL = gstrSQL & "," & rsData!����
                            '�ⷿID
                            gstrSQL = gstrSQL & "," & mParams.lngҩ��ID
                            '��ҩȷ��
                            gstrSQL = gstrSQL & "," & 1
                            '����Ա
                            gstrSQL = gstrSQL & ",'" & gstrUserName & "')"
        
                            Call zldatabase.ExecuteProcedure(gstrSQL, "RecipeWork_DosageOk")
                        rsData.MoveNext
                    Loop
                
                End If
    
                gcnOracle.CommitTrans
    
                Call RefreshList(mcondition.intListType)
            Else
                Me.stbThis.Panels(2).Text = "���ţ�" & txtPati.Text & "���޲�����Ϣ��"
                blnSta = True
            End If
        End If
        
        If strCard = "���ݺ�" And mParams.blnMustDosageOkProcess And mParams.blnMustDosageProcess And InStr(1, mstrPrivs, "��ҩȷ��") > 0 Then
            txtPati.Text = GetFullNO(txtPati.Text, 13)
            gstrSQL = _
                "Select distinct nvl(A.����id,'') ����id,nvl(A.����,'') ����,A.NO,A.���� " & _
                "From δ��ҩƷ��¼ A,ҩƷ�շ���¼ B " & _
                "Where A.NO=B.NO and A.����=B.���� and A.�ⷿid=B.�ⷿid and A.NO=[1] And A.�ⷿid=[2] " & _
                "    And A.�������� between [3] and [4] and nvl(A.�Ŷ�״̬,0)=0"
            Set rsData = zldatabase.OpenSQLRecord(gstrSQL, "", Me.txtPati.Text, mParams.lngҩ��ID, mSQLCondition.date��ʼ����, mSQLCondition.date��������)
            
            If rsData.EOF Then
                Me.stbThis.Panels(2).Text = "NOΪ[" & txtPati.Text & "]�ĵ��ݲ����ڣ�"
                blnSta = True
            Else
                
                Do While Not rsData.EOF
                    If zlStr.NVL(rsData!����ID) = "" Then
                        str���� = str���� & rsData!���� & ","
                    Else
                        str����id = str����id & rsData!����ID & ","
                    End If
                    rsData.MoveNext
                Loop
                
                If str����id <> "" Or str���� <> "" Then
                    If str����id <> "" And str���� = "" Then
                        gstrSQL = "Select /*+ Rule*/ distinct A.NO,A.����,nvl(A.����,'') ����,D.�Ա�,D.����,A.�������� From δ��ҩƷ��¼ A,ҩƷ�շ���¼ B,������ü�¼ D,Table(Cast(f_Num2list([1]) As zlTools.t_Numlist)) C " & _
                                    "Where A.����=B.���� And A.NO=B.NO And A.�ⷿid=B.�ⷿid And B.����id=D.id And  A.����id=C.Column_Value And A.�ⷿid=[2] and A.�������� between [3] and [4] and nvl(A.�Ŷ�״̬,0)=0 " & _
                                    " Union All " & _
                                    "Select /*+ Rule*/ distinct A.NO,A.����,nvl(A.����,'') ����,D.�Ա�,D.����,A.�������� From δ��ҩƷ��¼ A,ҩƷ�շ���¼ B,סԺ���ü�¼ D,Table(Cast(f_Num2list([1]) As zlTools.t_Numlist)) C " & _
                                    "Where A.����=B.���� And A.NO=B.NO And A.�ⷿid=B.�ⷿid And B.����id=D.id And  A.����id=C.Column_Value And A.�ⷿid=[2] and A.�������� between [3] and [4] and nvl(A.�Ŷ�״̬,0)=0 "
                        Set rsData = zldatabase.OpenSQLRecord(gstrSQL, "", str����id, mParams.lngҩ��ID, mSQLCondition.date��ʼ����, mSQLCondition.date��������)
                    ElseIf str���� <> "" And str����id = "" Then
                        gstrSQL = "Select /*+ Rule*/ distinct A.NO,A.����,nvl(A.����,'') ����,D.�Ա�,D.����,A.�������� From δ��ҩƷ��¼ A,ҩƷ�շ���¼ B,������ü�¼ D,Table(Cast(f_Str2list([1]) As zlTools.t_Strlist)) C " & _
                                    "Where A.����=B.���� And A.NO=B.NO And A.�ⷿid=B.�ⷿid And B.����id=D.id And  A.����=C.Column_Value And A.�ⷿid=[2] and A.�������� between [3] and [4] and nvl(A.�Ŷ�״̬,0)=0 " & _
                                    " Union All " & _
                                    "Select /*+ Rule*/ distinct A.NO,A.����,nvl(A.����,'') ����,D.�Ա�,D.����,A.�������� From δ��ҩƷ��¼ A,ҩƷ�շ���¼ B,סԺ���ü�¼ D,Table(Cast(f_Str2list([1]) As zlTools.t_Strlist)) C " & _
                                    "Where A.����=B.���� And A.NO=B.NO And A.�ⷿid=B.�ⷿid And B.����id=D.id And  A.����=C.Column_Value And A.�ⷿid=[2] and A.�������� between [3] and [4] and nvl(A.�Ŷ�״̬,0)=0 "

                        Set rsData = zldatabase.OpenSQLRecord(gstrSQL, "", str����, mParams.lngҩ��ID, mSQLCondition.date��ʼ����, mSQLCondition.date��������)
                    Else
                        gstrSQL = "Select /*+ Rule*/ distinct A.NO,A.����,nvl(A.����,'') ����,D.�Ա�,D.����,A.�������� From δ��ҩƷ��¼ A,ҩƷ�շ���¼ B,������ü�¼ D,Table(Cast(f_Num2list([1]) As zlTools.t_Numlist)) C,Table(Cast(f_Str2list([5]) As zlTools.t_Strlist)) E " & _
                                    "Where A.����=B.���� And A.NO=B.NO And A.�ⷿid=B.�ⷿid And B.����id=D.id And (A.����id=C.Column_Value or A.����=E.Column_Value) And A.�ⷿid=[2] and A.�������� between [3] and [4] and nvl(A.�Ŷ�״̬,0)=0 " & _
                                    " Union All " & _
                                    "Select /*+ Rule*/ distinct A.NO,A.����,nvl(A.����,'') ����,D.�Ա�,D.����,A.�������� From δ��ҩƷ��¼ A,ҩƷ�շ���¼ B,סԺ���ü�¼ D,Table(Cast(f_Num2list([1]) As zlTools.t_Numlist)) C,Table(Cast(f_Str2list([5]) As zlTools.t_Strlist)) E " & _
                                    "Where A.����=B.���� And A.NO=B.NO And A.�ⷿid=B.�ⷿid And B.����id=D.id And (A.����id=C.Column_Value or A.����=E.Column_Value) And A.�ⷿid=[2] and A.�������� between [3] and [4] and nvl(A.�Ŷ�״̬,0)=0 "
                                    
                        Set rsData = zldatabase.OpenSQLRecord(gstrSQL, "", str����id, mParams.lngҩ��ID, mSQLCondition.date��ʼ����, mSQLCondition.date��������, str����)
                    End If
                End If
                
                If rsData.RecordCount > 1 Then
                    frmSelectNo.ShowMe rsData, Me, txtPati.Text
                End If
                
                If Not rsData Is Nothing Then
                    If rsData.RecordCount > 0 Then
                        rsData.MoveFirst
                        gcnOracle.BeginTrans
                        Do While Not rsData.EOF
                            strCondition = IIf(strCondition = "", strCondition, strCondition & " OR ") & "NO='" & rsData!NO & "'"
                            gstrSQL = "Zl_δ��ҩƷ��¼_��ҩȷ��("
                                'NO
                                gstrSQL = gstrSQL & "'" & rsData!NO & "'"
                                '����
                                gstrSQL = gstrSQL & "," & rsData!����
                                '�ⷿID
                                gstrSQL = gstrSQL & "," & mParams.lngҩ��ID
                                '��ҩȷ��
                                gstrSQL = gstrSQL & "," & 1
                                '����Ա
                                gstrSQL = gstrSQL & ",'" & gstrUserName & "')"
            
                                Call zldatabase.ExecuteProcedure(gstrSQL, "RecipeWork_DosageOk")
                            rsData.MoveNext
                        Loop
                        
                        gcnOracle.CommitTrans
                        Call RefreshList(mcondition.intListType)
                    End If
                End If
            End If
        End If
        
        DoEvents
        KeyAscii = 0
        mblnFinding = False
        
        If imgFilter.BorderStyle = cstLocate Then
            Call Form_KeyDown(vbKeyF3, 0)
        Else
            If strCard = "���ݺ�" Then
                If IsNumeric(txtPati.Text) Then
                    txtPati.Text = UCase(GetFullNO(txtPati.Text, 13))
                End If
            End If
            
            DoEvents
            RefreshList mcondition.intListType
            
            '��ȡ���˳������д���
            strRecipeString = mfrmList.GetCurrentBatchRecipe
    
            arrRecipe = Split(strRecipeString, "|")
            intCount = UBound(arrRecipe)
            
            For n = 0 To intCount
                strNos = IIf(strNos = "", "", strNos & "|") & Val(Split(arrRecipe(n), ",")(0)) & "," & Split(arrRecipe(n), ",")(1)
            Next
            
            '����Զ���ҩ�п�ʼ��ҩ���̣�����ýӿ��ϴ�����
            '����ģʽ�ϴ����˳������д���
            '�����ݽӿ�ʱû���������
            If TypeName(mobjDrugMAC) = "clsDrugPacker" Then
                If mcondition.intListType = mListType.����ҩ And mblnPackerConnect And mblnLoadDrug And mintAutoSendFlow = 1 And strNos <> "" Then
                    If mblnCompatible = True Then
                        If mobjDrugMAC.DYEY_MZ_TransRecipeList(mstrOpr, UserInfo.�û�����, UserInfo.�û�����, mParams.lngҩ��ID, strNos, strReturn, mSendOper.StartSend) = False Then
                            If MsgBox("�Զ���ҩϵͳδ׼���ã��Ƿ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                                Exit Sub
                            End If
                        End If
                    End If
                End If
            ElseIf TypeName(mobjDrugMAC) = "clsDrugMachine" Then
                If mcondition.intListType = mListType.����ҩ And mblnPackerConnect Then
                    mobjDrugMAC.Operation gstrDbUser, Val("22-��ʼ��ҩ"), "1|" & Replace(strNos, "|", ";"), strReturn
'                           If strReturn <> "" Then MsgBox strReturn, vbInformation, gstrSysName
                End If
            End If
                    
            If mblnScaner Then
                txtPati.Text = ""
                txtPati.SetFocus
            End If
        End If
        
        If mParams.blnSign And mParams.blnMustDosageOkProcess And blnˢ�� And InStr(1, mstrPrivs, "��ҩȷ��") > 0 Then
            For i = 0 To rsData.RecordCount - 1
                mfrmList.SetSign (strCondition)
                If tbcDetail.Selected.index = 0 Then
                    mfrmDetail.CmdProcess
                ElseIf tbcDetail.Item(1).Visible = True Then
                    mfrmRecipe.CmdProcess
                End If
            Next
        End If
        
        If blnˢ�� And blnSta = False Then
            Me.stbThis.Panels(2).Text = "���ţ�" & txtPati.Text
            txtPati.Text = ""
            txtPati.SetFocus
        End If

        Call zlControl.TxtSelAll(txtPati)
        
        '����ҩ״̬��ɨ����Զ�����
        If mParams.blnɨ������ And mcondition.intListType = mListType.����ҩ And mblnScaner And mParams.blnStartCall And mblnFinding And (mblnBrushCard Or Not mbln��������ˢ��) Then
            txtPati.Text = ""
            txtPati.SetFocus
            Call RecipeWork_Call
        End If

    End If
    Exit Sub
errHandle:
    gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub


Private Function InputIsScaner(ByRef txtInput As Object, ByVal KeyAscii As Integer) As Boolean
'���ܣ��ж�ָ���ı����е�ǰ�����Ƿ����������豸���룺��ʱ֧�ֶԡ�ҩƷ�շ���¼.NO������
'������KeyAscii=��KeyPress�¼��е��õĲ���
    Static sngInputBegin As Single
    Dim sngNow As Single, blnScaner As Boolean, strText As String
    
    '����ǰ�������ʾ������(��δ��ʾ����)
    strText = txtInput.Text
    If txtInput.SelLength = Len(txtInput.Text) Then strText = ""
    If KeyAscii = 8 Or KeyAscii = 9 Or KeyAscii = 10 Then
        If strText <> "" Then strText = Mid(strText, 1, Len(strText) - 1)
    Else
        strText = strText & Chr(KeyAscii)
    End If
    
    '�ж��Ƿ��������豸����
    sngNow = Timer
    If txtInput.Text = "" Or strText = "" Then
        sngInputBegin = sngNow
    Else
        If Format((sngNow - sngInputBegin) / Len(strText), "0.000") < 0.04 Then blnScaner = True
    End If
    
    InputIsScaner = blnScaner
End Function

Private Sub txtPati_LostFocus()
    txtPati.BackColor = &H80000005
    
    If Not mobjIDCard Is Nothing Then mobjIDCard.SetEnabled (False)
    If Not mobjICCard Is Nothing Then mobjICCard.SetEnabled (False)
    
    mblnInput = False
End Sub


Private Sub txtPati_Validate(Cancel As Boolean)
    If Val(mParams.int����ģʽ) = mFindType.���ݺ� Then
        If IsNumeric(txtPati.Text) Then
            txtPati.Text = GetFullNO(txtPati.Text, 13)
        End If
    End If
End Sub

Private Function Getʵ�ս��(ByVal Int���� As Integer, ByVal strNo As String, ByVal int�����־ As Integer) As Double
    Dim strsql As String
    Dim rsʵ�ս�� As ADODB.Recordset
    
    If int�����־ = 1 Or int�����־ = 4 Then
        strsql = "Select Nvl(Sum(A.ʵ�ս��), 0) ʵ�ս�� From ������ü�¼ A Where A.��¼״̬ = 0 And A.Id In (Select Distinct B.����id From ҩƷ�շ���¼ B Where B.���� = [1] And B.No = [2]) "
    Else
        strsql = "Select Nvl(Sum(A.ʵ�ս��), 0) ʵ�ս�� From סԺ���ü�¼ A Where A.��¼״̬ = 0 And A.Id In (Select Distinct B.����id From ҩƷ�շ���¼ B Where B.���� = [1] And B.No = [2]) "
    End If
    
    On Error GoTo errRow
    Set rsʵ�ս�� = zldatabase.OpenSQLRecord(strsql, "Getʵ�ս��", Int����, strNo, int�����־)
    Getʵ�ս�� = rsʵ�ս��!ʵ�ս��
    Exit Function
errRow:
    If ErrCenter = 1 Then Resume
End Function

Private Sub ChangWin()
    Dim date��ʼ���� As Date
    Dim date�������� As Date
    Dim dteTime As Date
    
    dteTime = Sys.Currentdate
    'ʱ�䷶Χ
    Select Case cboʱ�䷶Χ.ListIndex
        Case mTimeRange.����
            date��ʼ���� = CDate(Format(dteTime, "yyyy-mm-dd") & " 00:00:00")
            date�������� = CDate(Format(dteTime, "yyyy-mm-dd") & " 23:59:59")
        Case mTimeRange.������
            date��ʼ���� = CDate(Format(DateAdd("d", -1, dteTime), "yyyy-mm-dd") & " 00:00:00")
            date�������� = CDate(Format(dteTime, "yyyy-mm-dd") & " 23:59:59")
        Case mTimeRange.������
            date��ʼ���� = CDate(Format(DateAdd("d", -2, dteTime), "yyyy-mm-dd") & " 00:00:00")
            date�������� = CDate(Format(dteTime, "yyyy-mm-dd") & " 23:59:59")
        Case mTimeRange.ָ��ʱ�䷶Χ
            date��ʼ���� = CDate(Format(Dtp��ʼʱ��.Value, "yyyy-mm-dd hh:mm:ss"))
            date�������� = CDate(Format(Dtp����ʱ��.Value, "yyyy-mm-dd hh:mm:ss"))
        Case Else
            date��ʼ���� = CDate(Format(dteTime, "yyyy-mm-dd") & " 00:00:00")
            date�������� = CDate(Format(dteTime, "yyyy-mm-dd") & " 23:59:59")
    End Select
    
    Call GetSendWindows(mParams.lngҩ��ID)
    If mQueue.blnWin = False Then
        MsgBox "��ҩ�����д������°࣬���ܽ��е�����ҩ���ڲ�����", vbInformation, gstrSysName
    Else
        Call frm������ҩ����.ShowMe(mParams.lngҩ��ID, Me, date��ʼ����, date��������, mstrDeptNode)
    End If
End Sub

Private Sub InitIDKindNew()
    Dim int����ģʽ As Integer
    Dim strTemp As String
    
    int����ģʽ = mParams.int����ģʽ����
    strTemp = "��|���ݺ�|0;��|�����|0;��|����|0;��|���֤|0;IC|IC����|1;ҽ|ҽ����|0;ס|סԺ��|0"
    Me.IDKNType.IDKindStr = strTemp
    Call IDKNType.zlInit(Me, glngSys, mlngMode, gcnOracle, gstrDbUser, mobjSquareCard, strTemp, txtPati)
'    IDKNType.SetAutoReadCard True
    Me.IDKNType.IDKind = int����ģʽ
End Sub

Private Sub IDKNType_ItemClick(index As Integer, objCard As zlIDKind.Card)
    Set mobjcard = objCard
    mParams.int����ģʽ���� = index
    mParams.int����ģʽ = Get����ģʽ(IDKNType.GetCurCard.����)
    mintOld����ģʽ = mParams.int����ģʽ
    
'    txtPati.MaxLength = objCard.���ų���
    If objCard.�������Ĺ��� <> "" Then
        txtPati.PasswordChar = "*"
    Else
        txtPati.PasswordChar = ""
    End If
    
    mbln��������ˢ�� = False
    If mParams.str����ˢ����ҩ <> "" Then
        mbln��������ˢ�� = InStr(1, "," & mParams.str����ˢ����ҩ & ",", "," & objCard.�ӿ���� & ",") > 0
    End If
    
    picConMain_Resize
End Sub

Private Function Get����ģʽ(ByVal str���� As String) As Integer
    '��IDKind�з��ص�ǰ�����ڲ������������
    Dim i As Integer
    Dim str���ʹ� As String
    
    'str���ʹ��봫���IDKindStr�������ơ�˳��һ��
    str���ʹ� = "���ݺ�,�����,����,�������֤,IC��,ҽ����,סԺ��"
    
    For i = 0 To UBound(Split(str���ʹ�, ","))
        If Split(str���ʹ�, ",")(i) = str���� Then
            Get����ģʽ = i + 1
            Exit For
        End If
    Next
    
    '��IDKindf���ص����Ͳ���IDKindStr��������ͣ���ֵһ������IDKindStr���͸���������
    If Get����ģʽ = 0 Then Get����ģʽ = 8
     
End Function

Private Sub IDKNType_ReadCard(ByVal objCard As zlIDKind.Card, objPatiInfor As zlIDKind.PatiInfor, blnCancel As Boolean)
    
    txtPati.Text = objPatiInfor.����
    If txtPati.Text <> "" Then Call txtPati_KeyPress(vbKeyReturn)
End Sub


Private Sub GetChildWin()
    Dim strsql As String
    Dim rstemp As Recordset
    
    On Error GoTo errHandle
    
    strsql = "select ���� from ��ҩ���� A,Table(f_Str2list([1])) B  where A.����=B.Column_Value or A.�кŴ���=B.Column_Value "
    
    Set rstemp = zldatabase.OpenSQLRecord(strsql, "��ȡ�Ӵ���", mParams.Str����)
    
    mParams.Str���� = ""
    Do While Not rstemp.EOF
        mParams.Str���� = mParams.Str���� & rstemp!���� & ","
        rstemp.MoveNext
    Loop
    
    
    Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub









