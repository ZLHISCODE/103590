VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{876E3FF4-6E21-11D5-AF7D-0080C8EC27A9}#1.4#0"; "ZL9BILLEDIT.OCX"
Begin VB.Form frmIdentify���� 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "�����֤"
   ClientHeight    =   6915
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11985
   Icon            =   "frmIdentify����.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6915
   ScaleWidth      =   11985
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdˢ�� 
      Caption         =   "ˢ��(&R)"
      Height          =   350
      Left            =   6120
      TabIndex        =   46
      Top             =   6480
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   10575
      TabIndex        =   45
      Top             =   6480
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Enabled         =   0   'False
      Height          =   350
      Left            =   9330
      TabIndex        =   44
      Top             =   6480
      Width           =   1100
   End
   Begin VB.CommandButton cmd���� 
      Caption         =   "ɾ����ʷ��¼(&D)"
      Enabled         =   0   'False
      Height          =   350
      Index           =   2
      Left            =   3840
      TabIndex        =   49
      ToolTipText     =   "��ݼ���DEL"
      Top             =   6480
      Width           =   1725
   End
   Begin VB.CommandButton cmd���� 
      Caption         =   "������ʷ��¼(&I)"
      Enabled         =   0   'False
      Height          =   350
      Index           =   1
      Left            =   1980
      TabIndex        =   48
      ToolTipText     =   "��ݼ���Ctrl+I"
      Top             =   6480
      Width           =   1725
   End
   Begin VB.CommandButton cmd���� 
      Caption         =   "������ʷ��¼(&A)"
      Enabled         =   0   'False
      Height          =   350
      Index           =   0
      Left            =   120
      TabIndex        =   47
      ToolTipText     =   "��ݼ���Ctrl+A"
      Top             =   6480
      Width           =   1725
   End
   Begin TabDlg.SSTab tabShow 
      Height          =   3735
      Left            =   120
      TabIndex        =   41
      Top             =   2640
      Width           =   11805
      _ExtentX        =   20823
      _ExtentY        =   6588
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabHeight       =   520
      Enabled         =   0   'False
      TabCaption(0)   =   "סԺ��¼(&1)"
      TabPicture(0)   =   "frmIdentify����.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Bill(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "�����¼(&2)"
      TabPicture(1)   =   "frmIdentify����.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Bill(1)"
      Tab(1).ControlCount=   1
      Begin ZL9BillEdit.BillEdit Bill 
         Height          =   3255
         Index           =   0
         Left            =   90
         TabIndex        =   42
         Top             =   390
         Width           =   11625
         _ExtentX        =   20505
         _ExtentY        =   5741
         CellAlignment   =   9
         Text            =   ""
         TextMatrix0     =   ""
         MaxDate         =   2958465
         MinDate         =   -53688
         Value           =   36395
         Cols            =   2
         RowHeight0      =   315
         RowHeightMin    =   315
         ColWidth0       =   1005
         BackColor       =   -2147483643
         BackColorBkg    =   -2147483643
         BackColorSel    =   10249818
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         ForeColorSel    =   -2147483634
         GridColor       =   -2147483630
         ColAlignment0   =   9
         ListIndex       =   -1
         CellBackColor   =   -2147483643
      End
      Begin ZL9BillEdit.BillEdit Bill 
         Height          =   3255
         Index           =   1
         Left            =   -74910
         TabIndex        =   43
         Top             =   390
         Width           =   11625
         _ExtentX        =   20505
         _ExtentY        =   5741
         CellAlignment   =   9
         Text            =   ""
         TextMatrix0     =   ""
         MaxDate         =   2958465
         MinDate         =   -53688
         Value           =   36395
         Cols            =   2
         RowHeight0      =   315
         RowHeightMin    =   315
         ColWidth0       =   1005
         BackColor       =   -2147483643
         BackColorBkg    =   -2147483643
         BackColorSel    =   10249818
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         ForeColorSel    =   -2147483634
         GridColor       =   -2147483630
         ColAlignment0   =   9
         ListIndex       =   -1
         CellBackColor   =   -2147483643
      End
   End
   Begin VB.Frame fra������Ϣ 
      Caption         =   "������Ϣ(&X)"
      Enabled         =   0   'False
      Height          =   1905
      Left            =   90
      TabIndex        =   6
      Top             =   600
      Width           =   11835
      Begin VB.ComboBox cbo������ϵ 
         Height          =   300
         Left            =   7320
         Style           =   2  'Dropdown List
         TabIndex        =   26
         Top             =   690
         Visible         =   0   'False
         Width           =   1395
      End
      Begin VB.ComboBox cbo�Ҵ���ʽ 
         Enabled         =   0   'False
         Height          =   300
         Left            =   10320
         Style           =   2  'Dropdown List
         TabIndex        =   40
         Top             =   1470
         Width           =   1365
      End
      Begin VB.ComboBox cbo��Ժ���� 
         Enabled         =   0   'False
         Height          =   300
         Left            =   3510
         Style           =   2  'Dropdown List
         TabIndex        =   34
         Top             =   1470
         Width           =   1335
      End
      Begin VB.ComboBox cbo��Ժ��ʽ 
         Enabled         =   0   'False
         Height          =   300
         Left            =   1170
         Style           =   2  'Dropdown List
         TabIndex        =   32
         Top             =   1470
         Width           =   1335
      End
      Begin VB.ComboBox cbo�Ҵ����� 
         Enabled         =   0   'False
         Height          =   300
         Left            =   7860
         Style           =   2  'Dropdown List
         TabIndex        =   38
         Top             =   1470
         Width           =   1395
      End
      Begin MSMask.MaskEdBox txt��ֹ���� 
         Height          =   300
         Left            =   5760
         TabIndex        =   30
         Top             =   1080
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   0   'False
         MaxLength       =   10
         Format          =   "yyyy-MM-dd"
         Mask            =   "####-##-##"
         PromptChar      =   "_"
      End
      Begin VB.ComboBox cbo���ⲡ�� 
         Height          =   300
         Left            =   1170
         Style           =   2  'Dropdown List
         TabIndex        =   28
         Top             =   1080
         Width           =   2745
      End
      Begin VB.ComboBox cbo����Ա���� 
         Enabled         =   0   'False
         Height          =   300
         Left            =   10080
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   690
         Width           =   1575
      End
      Begin VB.ComboBox cbo����Ա 
         Height          =   300
         Left            =   7320
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   690
         Width           =   1395
      End
      Begin VB.ComboBox cbo�������� 
         Height          =   300
         Left            =   4500
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   690
         Width           =   1725
      End
      Begin VB.ComboBox cbo�α���� 
         Height          =   300
         Left            =   1170
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   690
         Width           =   1965
      End
      Begin VB.TextBox txt�籣֤�� 
         Height          =   300
         Left            =   10080
         MaxLength       =   16
         TabIndex        =   16
         Top             =   300
         Width           =   1575
      End
      Begin VB.TextBox txt���֤�� 
         Height          =   300
         Left            =   7320
         MaxLength       =   18
         TabIndex        =   14
         Top             =   300
         Width           =   1725
      End
      Begin VB.ComboBox cbo�Ա� 
         Height          =   300
         Left            =   5250
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   300
         Width           =   975
      End
      Begin VB.TextBox txt���� 
         Height          =   300
         Left            =   3240
         MaxLength       =   20
         TabIndex        =   10
         Top             =   300
         Width           =   1275
      End
      Begin VB.ComboBox cboҽ����� 
         Height          =   300
         Left            =   1170
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   300
         Width           =   1335
      End
      Begin MSMask.MaskEdBox txt��Ժ���� 
         Height          =   300
         Left            =   5760
         TabIndex        =   36
         Top             =   1470
         Visible         =   0   'False
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   0   'False
         MaxLength       =   10
         Format          =   "yyyy-MM-dd"
         Mask            =   "####-##-##"
         PromptChar      =   "_"
      End
      Begin VB.Label lbl������ϵ 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "������ϵ*"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   6390
         TabIndex        =   25
         Top             =   750
         Visible         =   0   'False
         Width           =   810
      End
      Begin VB.Label lbl�Ҵ���ʽ 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "�Ҵ���ʽ"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   9480
         TabIndex        =   39
         Top             =   1530
         Width           =   720
      End
      Begin VB.Label lbl��Ժ���� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "��Ժ����"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   4980
         TabIndex        =   35
         Top             =   1530
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.Label lbl��Ժ���� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "��Ժ����"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   2670
         TabIndex        =   33
         Top             =   1530
         Width           =   720
      End
      Begin VB.Label lbl��Ժ��ʽ 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "��Ժ��ʽ"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   330
         TabIndex        =   31
         Top             =   1530
         Width           =   720
      End
      Begin VB.Label lbl�Ҵ����� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "�Ҵ�����"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   7020
         TabIndex        =   37
         Top             =   1530
         Width           =   720
      End
      Begin VB.Label lbl���ⲡ��ֹ���� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "���ⲡ��Ч��ֹ����"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   4080
         TabIndex        =   29
         Top             =   1140
         Width           =   1620
      End
      Begin VB.Label lbl���ⲡ�� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "���ⲡ��"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   330
         TabIndex        =   27
         Top             =   1140
         Width           =   720
      End
      Begin VB.Label lbl����Ա���� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "����Ա����"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   9060
         TabIndex        =   23
         Top             =   750
         Width           =   900
      End
      Begin VB.Label lbl����Ա 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "����Ա*"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   6570
         TabIndex        =   21
         Top             =   750
         Width           =   630
      End
      Begin VB.Label lbl�������� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "������(��)*"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   3390
         TabIndex        =   19
         Top             =   750
         Width           =   990
      End
      Begin VB.Label lbl�α���� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "�α����*"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   240
         TabIndex        =   17
         Top             =   750
         Width           =   810
      End
      Begin VB.Label lbl�籣֤�� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "�籣֤��*"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   9150
         TabIndex        =   15
         Top             =   360
         Width           =   810
      End
      Begin VB.Label lbl���֤�� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "���֤��*"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   6390
         TabIndex        =   13
         Top             =   360
         Width           =   810
      End
      Begin VB.Label lbl�Ա� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "�Ա�*"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   4680
         TabIndex        =   11
         Top             =   360
         Width           =   450
      End
      Begin VB.Label lbl���� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "����*"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   2670
         TabIndex        =   9
         Top             =   360
         Width           =   450
      End
      Begin VB.Label lblҽ����� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "ҽ�����*"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   240
         TabIndex        =   7
         Top             =   360
         Width           =   810
      End
   End
   Begin VB.CommandButton cmd���� 
      Caption         =   "����(&R)"
      Enabled         =   0   'False
      Height          =   345
      Left            =   3150
      TabIndex        =   2
      Top             =   180
      Width           =   1155
   End
   Begin VB.TextBox txtȷ�Ͽ��� 
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   8130
      MaxLength       =   12
      TabIndex        =   5
      Top             =   210
      Width           =   1905
   End
   Begin VB.TextBox txt���� 
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6120
      MaxLength       =   12
      TabIndex        =   4
      Top             =   210
      Width           =   1905
   End
   Begin VB.ComboBox cbo�������� 
      Height          =   300
      Left            =   1470
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   210
      Width           =   1605
   End
   Begin VB.Label lbl���� 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "��/�ֲ��(&S)*"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   4860
      TabIndex        =   3
      Top             =   270
      Width           =   1170
   End
   Begin VB.Label lbl�������� 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "��������(&T)*"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   300
      TabIndex        =   0
      Top             =   270
      Width           =   1080
   End
End
Attribute VB_Name = "frmIdentify����"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Enum ��ʷ��¼
    סԺ = 0
    ����
End Enum
Enum ����
    ���� = 0
    ����
    ɾ��
End Enum

'���ؼ�����
Private Const col_ҽ�ƻ��� As Integer = 0
Private Const col_�������� As Integer = 2
Private Const col_��Ժ���� As Integer = 2
Private Const col_��Ժ���� As Integer = 4
'סԺ��ʷ��¼
Private Const col_��Ժ���� As Integer = 1
Private Const col_��Ժ���� As Integer = 3
'������ʷ��¼
Private Const col_ҽ����� As Integer = 1

Private mbytType As Byte                'ģʽ 0-�����շѣ�1-��Ժ�Ǽǣ�2-������������סԺ,3-�Һ�
Private mlng����ID  As Long
Private mstrReturn As String
Private mdbl�ʻ���� As Double
Private mbln����ҽԺ As Boolean
Private mbln���ⲡ����ҽԺ As Boolean
'1������ҽԺ���жϣ������ϲ��Զ���ҽԺ�����жϣ��ɲ���Ա�жϣ��Ƕ���ҽԺ������ͨ���̣�
'A�� ��ҽ��ҽԺ����02����ר��ҽԺ��ҽԺ����03����Ϊ���вα��˵Ķ���ҽԺ��
'B�� ���������סԺ���ڲα��˵ķǶ���ҽԺ�������ҽ��������
'C�� �α��˳�סԺ�����ڲα��˵ķǶ���ҽԺ�������ҽ��������
'
'2���ֲ���д����:
'A�������ڿ����סԺ�������ֲ���Ӧ��¼�ֶηֽ���Ϣ


'���ݡ���������Ҽ��˲У���Ҫȷ��������ϵ
Private Sub LoadInitData()
    Dim strData As String
    'װ��ȱʡ���ݣ�����ָ����ϵ��������Щ���ݣ����ǵ���Щ���ǻ������ݣ���˴˴�д����
    
    strData = "�ֲ�,0|��,1"
    Call LoadCboData(cbo��������, strData)
    strData = "��,1|Ů,2|δ֪,9"
    Call LoadCboData(cbo�Ա�, strData)
    strData = ",0|��͸,1|���������Ż���,2|��͸+���������Ż���,3|������,4|��͸+������,5|���������Ż���+������,6|��͸+���������Ż���+������,7"
    Call LoadCboData(cbo���ⲡ��, strData)
    strData = "��ְ,11|��ְ����פ��,12|��ְ�����Ҽ��˲о���,13|����,21|������ذ���,22|��ְ�����Ҽ��˲о���,23|���ݶ����Ҽ��˲о���,24|" & _
            "��ְ,25|��ְ��ذ���,26|����,31|�Ϻ��,32|����ȫ����Ա,34|��ְ˾�ּ�ҽ����Ա,35|����˾�ּ�ҽ����Ա,36|��ְ������ҽ����Ա,37|" & _
            "���ݸ�����ҽ����Ա,38|��ͱ�����ְ,40|��ͱ�������,41|��ͱ�����ְ,42|��������Ҽ��˲о���,49|��ԺԺʿ,51|���������Ա,52|" & _
            "���������Ա,61|�����ְ��Ա,63|��ҵ����������Ա,65|��ҵ���������׵ذ�����Ա,66|��ҵ������ְ��Ա,67|��ҵ������ְ�׵ذ�����Ա,68|" & _
            "����������Ա,71|������ְ��Ա,73|֧Ԯ����������Ա,75|֧Ԯ������ְ��Ա,77|�Ʋ�������Ա,81|�Ʋ���ְ��Ա,83|" & _
            "��ҵע������������Ա,85|��ҵע�����������׵ذ�����Ա,86|��ҵע��������ְ��Ա,87|��ҵע��������ְ�׵ذ�����Ա,88|������Ա,91"
    Call LoadCboData(cbo�α����, strData)
    strData = "������,1010|������,1020|������,1030|������,1040|������,1050|������������,1051|����������,1052|����������ɽ,1053|" & _
            "��̨��,1060|��̨�������,1061|ʯ��ɽ��,1070|������,1080|����������·,1081|�������ϵ�,1082|��ͷ����,1090|��ɽ��,1110|" & _
            "��ƽ��,2210|˳����,2220|ͨ����,2230|������,2240|ƽ����,2260|������,2270|������,2280|������,2290|�����о��ü���������,2310|������ҽ������,2320"
    Call LoadCboData(cbo��������, strData)
    strData = "������,1|����,0"
    Call LoadCboData(cbo����Ա, strData)
    strData = ",-1|����,110|�м�,120|������,140|������,150|������,160|������,170|������,180|" & _
            "������,190|��̨��,200|ʯ��ɽ��,210|��ͷ����,220|��ɽ��,230|ͨ����,240|" & _
            "������,250|��ƽ��,260|˳����,270|������,280|������,290|ƽ����,310|������,320|���ÿ�����,330"
    Call LoadCboData(cbo����Ա����, strData)
    strData = ",-1|����,1|ʡ,2|�ƻ�������,3|��,4|��,41|��,5|�ֵ�,51|����,6|����,7|����,9"
    Call LoadCboData(cbo������ϵ, strData)
    
        
    '��������ֻ��ʾ�������ⲡ�������סԺ������ʾסԺ������ȫ����ʾ
    If mbytType = 0 Then
        strData = "�������ⲡ,12"
    ElseIf mbytType = 1 Then
        strData = "�������ⲡ,12|סԺ,21"
    Else
        strData = "�������ⲡ,12|סԺ,21"
    End If
    Call LoadCboData(cboҽ�����, strData)
    
    strData = "��ͨ,0|����,2"
    Call LoadCboData(cbo�Ҵ�����, strData)
    strData = "����,0|��ת��,1"
    Call LoadCboData(cbo�Ҵ���ʽ, strData)
    strData = "����Ժ,0|ת��Ժ,1"
    Call LoadCboData(cbo��Ժ��ʽ, strData)
    strData = "��ͨ,0|���ⲡ����,1|����,3|��ҽҽԺ��Ŀ�,4"   '������ֲ,2
    Call LoadCboData(cbo��Ժ����, strData)
End Sub

Private Sub LoadCboData(ByVal cboObj As ComboBox, ByVal strData As String)
    Dim arrData
    Dim intIndex As Integer, intCOUNT As Integer
    
    arrData = Split(strData, "|")
    intCOUNT = UBound(arrData)
    With cboObj
        .Clear
        For intIndex = 0 To intCOUNT
            .AddItem Split(arrData(intIndex), ",")(0)
            .ItemData(.NewIndex) = Split(arrData(intIndex), ",")(1)
        Next
        .ListIndex = 0
    End With
End Sub

Private Sub InitBill()
    Dim arrCol
    Dim billObj As BillEdit
    Dim intCol As Integer, intCols As Integer
    '��ʽ˵��������,���,������
    Const strסԺ As String = "ҽԺ����,1500,1|��Ժ����,1000,3|��Ժ����,1000,2|��Ժ����,1000,3|��Ժ����,1000,2|" & _
        "ͳ��֧��,1000,4|���/����Ա����,1600,4|�����Ը�,1000,4|�����Է�,1000,4|ͳ��ⶥ��ҽ����,1800,4"
    Const str���� As String = "ҽԺ����,1500,1|ҽ�����,1200,3|��������,1000,2|ͳ��֧��,1000,4|" & _
        "���/����Ա����,1600,4|�����Ը�,1000,4|�����Է�,1000,4|ͳ��ⶥ��ҽ����,1700,4"
    
    '��סԺ�����г�ʼ��
    arrCol = Split(strסԺ, "|")
    intCols = UBound(arrCol)
    Set billObj = Bill(סԺ)
    billObj.ClearBill
    billObj.Active = True
    billObj.Cols = intCols + 1
    For intCol = 0 To intCols
        billObj.TextMatrix(0, intCol) = Split(arrCol(intCol), ",")(0)
        billObj.ColWidth(intCol) = Split(arrCol(intCol), ",")(1)
        billObj.ColData(intCol) = Split(arrCol(intCol), ",")(2)
    Next
    
    '����������г�ʼ��
    arrCol = Split(str����, "|")
    intCols = UBound(arrCol)
    Set billObj = Bill(����)
    billObj.ClearBill
    billObj.Active = True
    billObj.Cols = intCols + 1
    For intCol = 0 To intCols
        billObj.TextMatrix(0, intCol) = Split(arrCol(intCol), ",")(0)
        billObj.ColWidth(intCol) = Split(arrCol(intCol), ",")(1)
        billObj.ColData(intCol) = Split(arrCol(intCol), ",")(2)
    Next
End Sub

Private Sub ReadBill()
    On Error GoTo errHand
    Dim rsHostory As New ADODB.Recordset
    '��ȡָ�����˵���ʷ�����¼
    If Trim(txtȷ�Ͽ���.Text) = "" Then Exit Sub
    Call InitBill
    
    Call DebugTool("��ȡ��ʷ�����¼")
    Call WriteBusinessLOG("��ȡ��ʷ�����¼", "", "")
    '��־Ϊ2��ʾסԺ������������
    gstrSQL = "SELECT DECODE(ҽ�����,21,2,22,2,23,2,1) AS ��־,ҽ�ƻ���,����,B.���� AS ҽ�����," & _
             " TO_CHAR(��Ժ����,'yyyy-MM-dd') AS ��Ժ����,C.���� AS ��Ժ����," & _
             " TO_CHAR(��Ժ����,'yyyy-MM-dd') AS ��Ժ����,D.���� AS ��Ժ����, " & _
             " �����ܶ�,ͳ��֧��,���֧��,�����Ը�,�����Է�,ͳ��ⶥ��ҽ���� " & _
             " FROM �ֲ����Ѽ�¼ A, " & _
             "      (SELECT B.����,B.����" & _
             "       FROM ָ������ A,ָ����ϵ���ձ� B" & _
             "       WHERE A.���=B.��� And A.����='ҽ�����') B," & _
             "      (SELECT B.����,B.����" & _
             "       FROM ָ������ A,ָ����ϵ���ձ� B" & _
             "       WHERE A.���=B.��� And A.����='��Ժ��ʽ') C," & _
             "      (SELECT B.����,B.����" & _
             "       FROM ָ������ A,ָ����ϵ���ձ� B" & _
             "       WHERE A.���=B.��� And A.����='��Ժ���') D" & _
             " WHERE A.ҽ�����=B.���� AND A.��Ժ����=C.����(+) And A.��Ժ����=D.����(+)" & _
             " AND A.����='" & txtȷ�Ͽ���.Text & "'" & _
             " ORDER BY ��־,��Ժ����"
    If rsHostory.State = 1 Then rsHostory.Close
    Call SQLTest(App.Title, "ZL9INSURE\READBILL", gstrSQL): rsHostory.Open gstrSQL, gcnBJYB: Call SQLTest
    Call DebugTool("��ʷ�����¼������" & rsHostory.RecordCount)
    Call WriteBusinessLOG("��ʷ�����¼������" & rsHostory.RecordCount, "", "")
    
    Call WriteBill(rsHostory)
    Exit Sub
errHand:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub ReadPatient()
    On Error GoTo errHand
    Dim RSPATIENT As New ADODB.Recordset
    If Trim(txtȷ�Ͽ���.Text) = "" Then Exit Sub
    
    Call DebugTool("��ȡ�ò��˵Ļ�����Ϣ--��ǰ�ڱ�Ժ������ľ�����Ϣ")
    Call WriteBusinessLOG("��ȡ�ò��˵Ļ�����Ϣ--��ǰ�ڱ�Ժ������ľ�����Ϣ", "", "")
    gstrSQL = "SELECT Nvl(����ID,0) AS ����ID,����,�籣֤��,����,B.���� AS �Ա�,Ѫ��,���֤��, " & _
             "     C.���� AS �α����,D.���� AS �ɷѵ�������,����Ա,����Ա����,���ֱ�ʶ, " & _
             "     TO_CHAR(���ⲡ��ֹ����,'yyyy-MM-dd') AS ���ⲡ��ֹ���� " & _
             " FROM �����ʻ� A, " & _
             "     (SELECT B.����,B.���� " & _
             "      FROM ָ������ A,ָ����ϵ���ձ� B " & _
             "      WHERE A.����='�Ա�' AND A.���=B.���) B, " & _
             "     (SELECT B.����,B.���� " & _
             "      FROM ָ������ A,ָ����ϵ���ձ� B " & _
             "      WHERE A.����='ҽ���α���Ա���' AND A.���=B.���) C, " & _
             "     (SELECT B.����,B.���� " & _
             "      FROM ָ������ A,ָ����ϵ���ձ� B " & _
             "      WHERE A.����='��������' AND A.���=B.���) D" & _
             " WHERE ����='" & Trim(txtȷ�Ͽ���.Text) & "'" & _
             " AND A.�Ա�=B.����(+) AND A.�α����=C.����(+) AND A.�ɷѵ�������=D.����(+)"
    If RSPATIENT.State = 1 Then RSPATIENT.Close
    Call SQLTest(App.Title, "ZL9INSURE\READBILL", gstrSQL): RSPATIENT.Open gstrSQL, gcnBJYB: Call SQLTest
    Call DebugTool("�ɹ���ȡ�ò�����������ʱ�ǼǵĻ�����Ϣ")
    Call WriteBusinessLOG("�ɹ���ȡ�ò�����������ʱ�ǼǵĻ�����Ϣ", "", "")
    
    Call WritePatient(RSPATIENT)
    Exit Sub
errHand:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub WritePatient(ByVal RSPATIENT As ADODB.Recordset)
    On Error GoTo errHand
    '�����˵Ļ�����Ϣд�����
    
    With RSPATIENT
        If .RecordCount = 0 Then Exit Sub
        txt����.Text = Nvl(!����)
        Call zlControl.CboLocate(cbo�Ա�, !�Ա�)
        txt���֤��.Text = Nvl(!���֤��)
        txt�籣֤��.Text = Nvl(!�籣֤��)
        Call zlControl.CboLocate(cbo�α����, !�α����)
        Call zlControl.CboLocate(cbo��������, !�ɷѵ�������)
        Call zlControl.CboLocate(cbo������ϵ, !����Ա����, True)
        Call zlControl.CboLocate(cbo����Ա, !����Ա, True)
        Call zlControl.CboLocate(cbo����Ա����, !����Ա����, True)
        Call zlControl.CboLocate(cbo���ⲡ��, !���ֱ�ʶ, True)
        txt��ֹ����.Text = Nvl(!���ⲡ��ֹ����, "____-__-__")
    End With
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub WriteBill(ByVal rsHostory As ADODB.Recordset)
    Dim lngRow As Long
    Dim objBill As BillEdit
    On Error GoTo errHand
    '����ʷ�����¼��д�������
    Bill(סԺ).Redraw = False
    Bill(����).Redraw = False
    
    With rsHostory
        'סԺ
        .Filter = "��־=2"
        Set objBill = Bill(סԺ)
        Do While Not .EOF
            lngRow = .AbsolutePosition
            objBill.TextMatrix(lngRow, 0) = Nvl(rsHostory!ҽ�ƻ���)
            objBill.TextMatrix(lngRow, 1) = Nvl(rsHostory!��Ժ����)
            objBill.TextMatrix(lngRow, 2) = Nvl(rsHostory!��Ժ����)
            objBill.TextMatrix(lngRow, 3) = Nvl(rsHostory!��Ժ����)
            objBill.TextMatrix(lngRow, 4) = Nvl(rsHostory!��Ժ����)
            'objBill.TextMatrix(lngRow, 5) = Format(Nvl(rsHostory!�����ܶ�, 0), "#####0.00;-#####0.00; ;")
            objBill.TextMatrix(lngRow, 5) = Format(Nvl(rsHostory!ͳ��֧��, 0), "#####0.00;-#####0.00; ;")
            objBill.TextMatrix(lngRow, 6) = Format(Nvl(rsHostory!���֧��, 0), "#####0.00;-#####0.00; ;")
            objBill.TextMatrix(lngRow, 7) = Format(Nvl(rsHostory!�����Ը�, 0), "#####0.00;-#####0.00; ;")
            objBill.TextMatrix(lngRow, 8) = Format(Nvl(rsHostory!�����Է�, 0), "#####0.00;-#####0.00; ;")
            objBill.TextMatrix(lngRow, 9) = Format(Nvl(rsHostory!ͳ��ⶥ��ҽ����, 0), "#####0.00;-#####0.00; ;")
            
            lngRow = lngRow + 1
            objBill.Rows = objBill.Rows + 1
            .MoveNext
        Loop
        '����
        .Filter = "��־=1"
        Set objBill = Bill(����)
        Do While Not .EOF
            lngRow = .AbsolutePosition
            objBill.TextMatrix(lngRow, 0) = Nvl(rsHostory!ҽ�ƻ���)
            objBill.TextMatrix(lngRow, 1) = Nvl(rsHostory!ҽ�����)
            objBill.TextMatrix(lngRow, 2) = Nvl(rsHostory!��Ժ����)
            'objBill.TextMatrix(lngRow, 3) = Format(Nvl(rsHostory!�����ܶ�, 0), "#####0.00;-#####0.00; ;")
            objBill.TextMatrix(lngRow, 3) = Format(Nvl(rsHostory!ͳ��֧��, 0), "#####0.00;-#####0.00; ;")
            objBill.TextMatrix(lngRow, 4) = Format(Nvl(rsHostory!���֧��, 0), "#####0.00;-#####0.00; ;")
            objBill.TextMatrix(lngRow, 5) = Format(Nvl(rsHostory!�����Ը�, 0), "#####0.00;-#####0.00; ;")
            objBill.TextMatrix(lngRow, 6) = Format(Nvl(rsHostory!�����Է�, 0), "#####0.00;-#####0.00; ;")
            objBill.TextMatrix(lngRow, 7) = Format(Nvl(rsHostory!ͳ��ⶥ��ҽ����, 0), "#####0.00;-#####0.00; ;")
            
            lngRow = lngRow + 1
            objBill.Rows = objBill.Rows + 1
            .MoveNext
        Loop
    End With
errHand:
    rsHostory.Filter = 0
    Bill(סԺ).Redraw = True
    Bill(����).Redraw = True
End Sub

Private Sub Bill_cboClick(Index As Integer, ListIndex As Long)
    With Bill(Index)
'        If .LastRow <> .Row Then Exit Sub
'        .TextMatrix(.Row, .Col) = .CboText
    End With
End Sub

Private Sub Bill_CommandClick(Index As Integer)
    Dim blnReturn As Boolean
    Dim rsTmp As New ADODB.Recordset
    With Bill(Index)
        If Bill(Index).COL = col_ҽ�ƻ��� Then
            gstrSQL = "" & _
                " SELECT A.ҽԺ����,A.ҽԺ����,zlSpellcode(A.ҽԺ����) As ����,B.����||'-'||B.���� AS ҽԺ�ȼ�,C.����||'-'||C.���� AS ҽԺ����" & _
                " FROM ҽԺ�ȼ� A," & _
                "     (SELECT B.����,B.����" & _
                "     FROM ָ������ A,ָ����ϵ���ձ� B" & _
                "     WHERE A.���=B.��� AND A.����='ҽԺ�ȼ�') B," & _
                "     (SELECT B.����,B.����" & _
                "     FROM ָ������ A,ָ����ϵ���ձ� B" & _
                "     WHERE A.���=B.��� AND A.����='ҽԺ����') C" & _
                " WHERE A.ҽԺ�ȼ�=B.����(+) AND A.ҽԺ����=C.����(+) AND A.��Ч����<=SYSDATE"
            If rsTmp.State = 1 Then rsTmp.Close
            Call SQLTest(App.Title, "ZL9INSURE\���ղ�������", gstrSQL): rsTmp.Open gstrSQL, gcnBJYB: Call SQLTest
            If rsTmp.RecordCount = 0 Then
                MsgBox "û���ҵ���ҽԺ��Ϣ�������䣡", vbInformation, gstrSysName
                Exit Sub
            Else
                '����ѡ����
                If rsTmp.RecordCount > 1 Then
                    '�����ֶδ���3�ģ���ʹֻ��һ����¼�ѸöԻ�����ʾ�������Ա����û��õ��������Ϣ
                    blnReturn = frmListSel.ShowSelect(TYPE_����, rsTmp, "ҽԺ����", "ҽԺ�ȼ�ѡ��", "��ѡ��ҽԺ�ȼ���")
                Else
                    blnReturn = True
                End If
            End If
            If blnReturn Then
                .Text = rsTmp!ҽԺ����
                .TextMatrix(.Row, .COL) = .Text
            End If
        End If
    End With
End Sub

Private Sub Bill_EnterCell(Index As Integer, Row As Long, COL As Long)
    With Bill(Index)
        If COL = 1 Then     'col_��Ժ���� ,col_ҽ�����
            .Clear
            If Index = סԺ Then
                .AddItem "��ͨ"
                .ItemData(.NewIndex) = 0
                .AddItem "���ⲡ����"
                .ItemData(.NewIndex) = 1
                .AddItem "����"
                .ItemData(.NewIndex) = 3
                .AddItem "��ҽҽԺ��Ŀ�"
                .ItemData(.NewIndex) = 4
                .ListIndex = 0
            Else
                .AddItem "�������ⲡ"
                .ItemData(.NewIndex) = 12
                .AddItem "��ͥ����"
                .ItemData(.NewIndex) = 31
                .ListIndex = 0
            End If
        ElseIf COL = col_��Ժ���� And Index = סԺ Then
            .Clear
            '0-��Ժ,1-ת��Ժ��2-��;����
            .AddItem "��Ժ"
            .ItemData(.NewIndex) = 0
            .AddItem "ת��Ժ"
            .ItemData(.NewIndex) = 1
            .AddItem "��;����"
            .ItemData(.NewIndex) = 2
            .ListIndex = 0
        ElseIf COL = col_ҽ�ƻ��� Then
            .TxtCheck = False
        Else
            If .ColData(.COL) = 4 Then
                'Ĭ��Ϊ�ǽ��������
                .TxtCheck = True
                .TextMask = "-0123456789."
            End If
        End If
    End With
End Sub

Private Sub Bill_GotFocus(Index As Integer)
    tabShow.Tab = Index
End Sub

Private Sub Bill_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer, Cancel As Boolean)
    Dim StrInput As String
    Dim blnReturn As Boolean
    Dim rsTmp As New ADODB.Recordset
    
    With Bill(Index)
        If KeyCode <> vbKeyReturn Then Exit Sub
        If .TxtVisible = False Then
            StrInput = IIf(.TextMatrix(.Row, .COL) = "", " ", .TextMatrix(.Row, .COL))
            .Text = StrInput
            .TextMatrix(.Row, .COL) = StrInput
        Else
            StrInput = UCase(Trim(.Text))
            If .COL = col_ҽ�ƻ��� Then
                If Trim(StrInput) = "" Then Exit Sub
                gstrSQL = "SELECT * FROM (" & _
                    " SELECT A.ҽԺ����,A.ҽԺ����,zlSpellcode(A.ҽԺ����) As ����,B.����||'-'||B.���� AS ҽԺ�ȼ�,C.����||'-'||C.���� AS ҽԺ����" & _
                    " FROM ҽԺ�ȼ� A," & _
                    "     (SELECT B.����,B.����" & _
                    "     FROM ָ������ A,ָ����ϵ���ձ� B" & _
                    "     WHERE A.���=B.��� AND A.����='ҽԺ�ȼ�') B," & _
                    "     (SELECT B.����,B.����" & _
                    "     FROM ָ������ A,ָ����ϵ���ձ� B" & _
                    "     WHERE A.���=B.��� AND A.����='ҽԺ����') C" & _
                    " WHERE A.ҽԺ�ȼ�=B.����(+) AND A.ҽԺ����=C.����(+) AND A.��Ч����<=SYSDATE) A" & _
                    " WHERE (A.ҽԺ���� Like '" & StrInput & "%' Or A.ҽԺ���� Like '" & StrInput & "%' Or A.���� Like '" & StrInput & "%')"
                If rsTmp.State = 1 Then rsTmp.Close
                Call SQLTest(App.Title, "ZL9INSURE\���ղ�������", gstrSQL): rsTmp.Open gstrSQL, gcnBJYB: Call SQLTest
                If rsTmp.RecordCount = 0 Then
                    MsgBox "û���ҵ���ҽԺ��Ϣ�������䣡", vbInformation, gstrSysName
                    Cancel = True
                    .TxtSetFocus
                    Exit Sub
                Else
                    '����ѡ����
                    If rsTmp.RecordCount > 1 Then
                        '�����ֶδ���3�ģ���ʹֻ��һ����¼�ѸöԻ�����ʾ�������Ա����û��õ��������Ϣ
                        blnReturn = frmListSel.ShowSelect(TYPE_����, rsTmp, "ҽԺ����", "ҽԺ�ȼ�ѡ��", "��ѡ��ҽԺ�ȼ���")
                    Else
                        blnReturn = True
                    End If
                End If
                If blnReturn Then
                    .Text = rsTmp!ҽԺ����
                    .TextMatrix(.Row, .COL) = .Text
                End If
            ElseIf .COL = col_��Ժ���� Or (.COL = col_��Ժ���� And Index = סԺ) Then
                If Trim(StrInput) <> "" Then
                    If Not IsDate(StrInput) Then
                        MsgBox "������Ч���������ݣ������䣡", vbInformation, gstrSysName
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    End If
                End If
                
                If Index = סԺ Then
                    If .COL = col_��Ժ���� Then
                        '���ܴ��ڳ�Ժ����
                        If Trim(.TextMatrix(.Row, col_��Ժ����)) <> "" Then
                            If StrInput > .TextMatrix(.Row, col_��Ժ����) Then
                                MsgBox "��Ժ���ڲ��ܴ��ڳ�Ժ���ڣ�", vbInformation, gstrSysName
                                Cancel = True
                                .TxtSetFocus
                                Exit Sub
                            End If
                        End If
                        If .Row > 1 Then
                            '����С��������¼�ĳ�Ժ����
                            If Trim(.TextMatrix(.Row - 1, col_��Ժ����)) <> "" Then
                                If StrInput < .TextMatrix(.Row - 1, col_��Ժ����) Then
                                    MsgBox "��Ժ���ڲ���С���ϴξ���ĳ�Ժ���ڣ�", vbInformation, gstrSysName
                                    Cancel = True
                                    .TxtSetFocus
                                    Exit Sub
                                End If
                            End If
                        End If
                        If .Row + 1 <= .Rows - 1 Then
                            '���ܴ�����һ����¼����Ժ����
                            If Trim(.TextMatrix(.Row + 1, col_��Ժ����)) <> "" Then
                                If StrInput > .TextMatrix(.Row + 1, col_��Ժ����) Then
                                    MsgBox "��Ժ���ڲ��ܴ���" & .Row + 1 & "�еǼǵ���Ժ���ڣ�", vbInformation, gstrSysName
                                    Cancel = True
                                    .TxtSetFocus
                                    Exit Sub
                                End If
                            End If
                        End If
                    ElseIf .COL = col_��Ժ���� Then
                        '��Ժ���ڲ���С����Ժ����
                        If Trim(.TextMatrix(.Row, col_��Ժ����)) <> "" Then
                            If StrInput < .TextMatrix(.Row, col_��Ժ����) Then
                                MsgBox "��Ժ���ڲ���С����Ժ���ڣ�", vbInformation, gstrSysName
                                Cancel = True
                                .TxtSetFocus
                                Exit Sub
                            End If
                        End If
                        If .Row + 1 <= .Rows - 1 Then
                            '��Ժ���ڲ��ܴ�����һ����¼����Ժ����
                            If Trim(.TextMatrix(.Row + 1, col_��Ժ����)) <> "" Then
                                If StrInput > .TextMatrix(.Row + 1, col_��Ժ����) Then
                                    MsgBox "��Ժ���ڲ��ܴ���" & .Row + 1 & "�еǼǵ���Ժ���ڣ�", vbInformation, gstrSysName
                                    Cancel = True
                                    .TxtSetFocus
                                    Exit Sub
                                End If
                            End If
                        End If
                    End If
                Else
                    If .COL = col_��Ժ���� Then
                        If .Row > 1 Then
                            If Trim(.TextMatrix(.Row - 1, col_��Ժ����)) <> "" Then
                                If StrInput < .TextMatrix(.Row - 1, col_��Ժ����) Then
                                    MsgBox "�������ڲ���С���ϴξ������ڣ�", vbInformation, gstrSysName
                                    Cancel = True
                                    .TxtSetFocus
                                    Exit Sub
                                End If
                            End If
                        End If
                        If .Row + 1 <= .Rows - 1 Then
                            '�������ڲ��ܴ�����һ����¼�ľ�������
                            If Trim(.TextMatrix(.Row + 1, col_��Ժ����)) <> "" Then
                                If StrInput > .TextMatrix(.Row + 1, col_��Ժ����) Then
                                    MsgBox "�������ڲ��ܴ���" & .Row + 1 & "�еǼǵľ������ڣ�", vbInformation, gstrSysName
                                    Cancel = True
                                    .TxtSetFocus
                                    Exit Sub
                                End If
                            End If
                        End If
                    End If
                End If
            Else    '���ǽ���У�ֻ��������
                .Text = Format(StrInput, "#0.00")
            End If
        End If
    End With
End Sub

Private Sub cbo��������_Click()
    Dim blnEnable As Boolean
    blnEnable = (cbo��������.ItemData(cbo��������.ListIndex) = 1)
    cmd����.Enabled = blnEnable
End Sub

Private Sub cbo�α����_Click()
    Dim bln������ϵ As Boolean
    Const str������ϵ = ";����;��������Ҽ��˲о���;"
    
    bln������ϵ = (InStr(1, str������ϵ, ";" & cbo�α����.Text & ";") <> 0)
    lbl����Ա.Visible = Not bln������ϵ
    lbl����Ա����.Visible = Not bln������ϵ
    cbo����Ա.Visible = Not bln������ϵ
    cbo����Ա����.Visible = Not bln������ϵ
    lbl������ϵ.Visible = bln������ϵ
    cbo������ϵ.Visible = bln������ϵ
End Sub

Private Sub cbo����Ա_Click()
    Dim objBill As BillEdit
    On Error Resume Next
    
    Me.cbo����Ա����.Enabled = (Me.cbo����Ա.ItemData(Me.cbo����Ա.ListIndex) = 0)
    'ֻ�й���Ա������������ⶥ��ҽ���ڽ��
    Set objBill = Bill(סԺ)
    With objBill
        .ColData(.Cols - 1) = 4
        If Me.cbo����Ա����.Visible And Me.cbo����Ա.ItemData(Me.cbo����Ա.ListIndex) <> 0 Then
            .ColData(.Cols - 1) = 5
        End If
    End With
    Set objBill = Bill(����)
    With objBill
        .ColData(.Cols - 1) = 4
        If Me.cbo����Ա����.Visible And Me.cbo����Ա.ItemData(Me.cbo����Ա.ListIndex) <> 0 Then
            .ColData(.Cols - 1) = 5
        End If
    End With
End Sub

Private Sub cbo���ⲡ��_Click()
    Dim blnEnable As Boolean
    blnEnable = (cbo���ⲡ��.ListIndex <> 0)
    txt��ֹ����.Enabled = blnEnable
End Sub

Private Sub cboҽ�����_Click()
    Dim blnEnable As Boolean
    blnEnable = (cboҽ�����.ItemData(cboҽ�����.ListIndex) = 21)
    cbo��Ժ��ʽ.Enabled = blnEnable
    cbo��Ժ����.Enabled = blnEnable
    txt��Ժ����.Enabled = blnEnable
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim intPage As Integer
    Dim intCurState As Integer  '��ǰ״̬
    Dim lngRow As Long, lngRows As Long
    Dim strBorn As String       '��������
    Dim StrInput As String      '���������
    Dim strInsert As String     '���ֽ�����ϴ�
    Dim blnClear As Boolean
    Dim blnTrans As Boolean
    Dim objBill As BillEdit
    Dim strIdentify As String, strAddition As String
    Dim rsTmp As New ADODB.Recordset
    On Error GoTo errHand
    
    If Not ValidData Then Exit Sub
    StrInput = txtȷ�Ͽ���.Text & "|" & txt�籣֤��.Text
    If Not �������_����(StrInput, True) Then Exit Sub
    
    '��ȡ���ò��˵ĵ�ǰ״̬
    intCurState = 0
    gstrSQL = "Select Nvl(��ǰ״̬,0) ��ǰ״̬ From �����ʻ� Where ����=[1] And ����=[2]"
    
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, CStr(txtȷ�Ͽ���.Text), TYPE_����)
     
    If Not ChkRsState(rsTmp) Then intCurState = rsTmp!��ǰ״̬
    
    '�����Ժ����������������֤
    If mbytType <> 2 And intCurState = 1 Then
        MsgBox "�òα��˵�ǰ��Ժ����������������֤��", vbInformation, gstrSysName
        Exit Sub
    End If
    
    'ת���������ڣ����֤�Ǳ����룬�϶�����19��� ��
    If Len(txt���֤��.Text) = 15 Then
        '15λ
        strBorn = "19" & Mid(txt���֤��.Text, 7, 2) & "-" & Mid(txt���֤��.Text, 9, 2) & "-" & Mid(txt���֤��.Text, 11, 2)
    Else
        '18λ
        strBorn = Mid(txt���֤��.Text, 7, 4) & "-" & Mid(txt���֤��.Text, 11, 2) & "-" & Mid(txt���֤��.Text, 13, 2)
    End If
    
    '����������Ϣ
    '�����ַ���
    '�������˵�����Ϣ�������ʽ��
    '0����;1ҽ����;2����;3����;4�Ա�;5��������;6���֤;7��λ����(����);
    '8.���Ĵ���;9.˳���;10��Ա���;11�ʻ����;12��ǰ״̬;13����ID;14��ְ(0,1);15����֤��;16�����;17�Ҷȼ�
    '18�ʻ������ۼ�,19�ʻ�֧���ۼ�,20����ͳ���ۼ�,21ͳ�ﱨ���ۼ�,22סԺ�����ۼ�
    strIdentify = txtȷ�Ͽ���.Text                              '0����
    strIdentify = strIdentify & ";" & txt�籣֤��.Text          '1ҽ����
    strIdentify = strIdentify & ";"                             '2����
    strIdentify = strIdentify & ";" & txt����.Text              '3����
    strIdentify = strIdentify & ";" & cbo�Ա�.Text              '4�Ա�
    strIdentify = strIdentify & ";" & strBorn                   '5��������
    strIdentify = strIdentify & ";" & txt���֤��.Text          '6���֤
    strIdentify = strIdentify & ";"                             '7.��λ����(����)
    strAddition = ";0"                                          '8.���Ĵ���
    strAddition = strAddition & ";"                             '9.˳���
    strAddition = strAddition & ";" & cbo�α����.Text          '10��Ա���
    strAddition = strAddition & ";" & mdbl�ʻ����              '11�ʻ����
    strAddition = strAddition & ";" & intCurState               '12��ǰ״̬
    strAddition = strAddition & ";0"                            '13����ID
    strAddition = strAddition & ";1"                            '14��ְ(1,2,3)
    strAddition = strAddition & ";"                             '15����֤��
    strAddition = strAddition & ";"                             '16�����
    strAddition = strAddition & ";"                             '17�Ҷȼ�
    strAddition = strAddition & ";0"                            '18�ʻ������ۼ�
    strAddition = strAddition & ";0"                            '19�ʻ�֧���ۼ�
    strAddition = strAddition & ";0"                            '20���깤���ܶ�
    strAddition = strAddition & ";0"                            '21סԺ�����ۼ�

    mlng����ID = BuildPatiInfo(0, strIdentify & strAddition, mlng����ID, TYPE_����)
    '���ظ�ʽ:�м���벡��ID
    If mlng����ID > 0 Then
        mstrReturn = strIdentify & ";" & mlng����ID & strAddition
    Else
        Exit Sub
    End If
    
    gcnBJYB.BeginTrans
    blnTrans = True
    If mbytType <> 2 Then
        '�������ղ��˵Ļ�����Ϣ
    '    ����ID,����,�籣֤��,����,�Ա�,Ѫ��,���֤��,ҵ������,��Ժ���,
    '    ��Ժ��ʽ,��Ժ����,�α����,�ɷѵ�������,�����ʻ����,����Ա,
    '    ����Ա���� , ����ҽԺ, ���ֱ�ʶ, ���ⲡ��ֹ����, ���ⲡ����ҽԺ
        gstrSQL = "zl_�����ʻ�_INSERT(" & mlng����ID & ",'" & txtȷ�Ͽ���.Text & "','" & txt�籣֤��.Text & "'," & _
            "'" & txt����.Text & "','" & Me.cbo�Ա�.ItemData(Me.cbo�Ա�.ListIndex) & "',NULL,'" & txt���֤��.Text & "'," & _
            "'" & Me.cboҽ�����.ItemData(Me.cboҽ�����.ListIndex) & "'," & Me.cbo��Ժ����.ItemData(Me.cbo��Ժ����.ListIndex) & "," & _
            "" & Me.cbo��Ժ��ʽ.ItemData(Me.cbo��Ժ��ʽ.ListIndex) & ",TO_DATE('" & txt��Ժ����.Text & "','yyyy-MM-dd')" & "," & _
            "'" & Me.cbo�α����.ItemData(Me.cbo�α����.ListIndex) & "','" & Me.cbo��������.ItemData(Me.cbo��������.ListIndex) & "'," & _
            "" & mdbl�ʻ���� & "," & Me.cbo����Ա.ItemData(Me.cbo����Ա.ListIndex) & "," & IIf(cbo������ϵ.Visible = False, Me.cbo����Ա����.ItemData(Me.cbo����Ա����.ListIndex), Me.cbo������ϵ.ItemData(Me.cbo������ϵ.ListIndex)) & "," & _
            "" & "1,'" & Me.cbo���ⲡ��.ItemData(Me.cbo���ⲡ��.ListIndex) & "'," & IIf(Me.txt��ֹ����.Text = "____-__-__", "NULL", "TO_DATE('" & Me.txt��ֹ����.Text & "','yyyy-MM-dd')") & ",1)"
        gcnBJYB.Execute gstrSQL, , adCmdStoredProc
    End If
    
    '������ʷ�ֲ����Ѽ�¼
    gcnBJYB.Execute "zl_�ֲ����Ѽ�¼_DELETEALL('" & txtȷ�Ͽ���.Text & "')"
    blnClear = True
    For intPage = 0 To Bill.UBound
        Set objBill = Bill(intPage)
        lngRows = objBill.Rows - 1
        For lngRow = 1 To lngRows
            If Trim(objBill.TextMatrix(lngRow, col_ҽ�ƻ���)) <> "" Then
'               ����,ҽ�ƻ���,ҽ�����,��Ժ����,��Ժ����,��Ժ����,��Ժ����,
'               ͳ��֧��,���֧��,�����Ը�,�����Է�,ͳ��ⶥ��ҽ����,������ˮ��,�����ʷ��¼
                strInsert = GetMoneySQL(intPage, lngRow)
                gstrSQL = "zl_�ֲ����Ѽ�¼_INSERT('" & txtȷ�Ͽ���.Text & "','" & objBill.TextMatrix(lngRow, col_ҽ�ƻ���) & "'," & _
                    Getҽ�����(intPage, lngRow) & "," & GetסԺ����(intPage, lngRow) & ",To_Date('" & objBill.TextMatrix(lngRow, col_��Ժ����) & "','yyyy-MM-dd')," & _
                    GetסԺ����(intPage, lngRow, False) & ",To_Date('" & objBill.TextMatrix(lngRow, IIf(intPage = סԺ, col_��Ժ����, col_��Ժ����)) & "','yyyy-MM-dd')," & _
                    strInsert & ",NULL," & IIf(blnClear, 1, 0) & ")"
                gcnBJYB.Execute gstrSQL, , adCmdStoredProc
                blnClear = False
            End If
        Next
    Next
    
    '����
    gcnBJYB.CommitTrans
    
    gComInfo_����.���� = txtȷ�Ͽ���.Text
    gComInfo_����.ҵ������ = Me.cboҽ�����.ItemData(Me.cboҽ�����.ListIndex)
    Unload Me
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
    If blnTrans Then gcnBJYB.RollbackTrans
End Sub

Private Sub cmd����_Click()
    MsgBox "�Բ���Ŀǰ�ݲ�֧�ֳֿ����˽������ѣ�", vbInformation, gstrSysName
    
End Sub

Private Sub cmd����_Click(Index As Integer)
    Dim lngRow As Long
    
    On Error Resume Next
    Select Case Index
    Case ����.����
        Bill(tabShow.Tab).Rows = Bill(tabShow.Tab).Rows + 1
        Bill(tabShow.Tab).Row = Bill(tabShow.Tab).Rows - 1
        Bill(tabShow.Tab).SetFocus
    Case ����.����
        lngRow = Bill(tabShow.Tab).Row
        Bill(tabShow.Tab).msfObj.AddItem "", Bill(tabShow.Tab).Row
        Bill(tabShow.Tab).Row = lngRow
        Bill(tabShow.Tab).SetFocus
    Case ����.ɾ��
        Bill(tabShow.Tab).SetFocus
        SendKeys "{DELETE}", 1
    End Select
End Sub

Private Sub cmdˢ��_Click()
    Dim objControl As Control
    mdbl�ʻ���� = 0
    Call InitBill
    '�������������
    For Each objControl In Me.Controls
        If UCase(TypeName(objControl)) = "TEXTBOX" Then
            objControl.Text = ""
        ElseIf UCase(TypeName(objControl)) = "COMBOBOX" Then
            objControl.ListIndex = 0
        End If
    Next
    
    fra������Ϣ.Enabled = False
    tabShow.Enabled = False
    cmd����(����).Enabled = False
    cmd����(����).Enabled = False
    cmd����(ɾ��).Enabled = False
    
    '�������뿨�Ż�ѡ��������
    Me.cbo��������.Enabled = True
    Me.txt����.Enabled = True
    Me.txtȷ�Ͽ���.Enabled = True
    cmdOK.Enabled = False
    If Me.txt��Ժ����.Text = "____-__-__" Then Me.txt��Ժ����.Text = Format(zlDatabase.Currentdate(), "yyyy-MM-dd")
    If cbo��������.Enabled Then cbo��������.SetFocus
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Not (Me.ActiveControl.Name = "txtȷ�Ͽ���" Or Me.ActiveControl.Name = "Bill") Then
            Call zlCommFun.PressKey(vbKeyTab)
        End If
    End If
End Sub

Private Sub Form_Load()
    mdbl�ʻ���� = 0
    Call InitBill
    Call LoadInitData
    
    Me.txt��Ժ����.Text = Format(zlDatabase.Currentdate(), "yyyy-MM-dd")
End Sub

Private Sub tabShow_GotFocus()
    If Not tabShow.Enabled Then Exit Sub
    If Bill(tabShow.Tab).Active Then Bill(tabShow.Tab).SetFocus
End Sub

Private Sub txt��ֹ����_GotFocus()
    With txt��ֹ����
        .SelStart = 0
        .SelLength = 10
    End With
End Sub

Private Sub txt����_GotFocus()
    If Trim(txt����.Text) = "" Then
        txt����.Text = "S"
    End If
    txt����.SelStart = 0
    If Not (Trim(txt����.Text) = "" Or Trim(txt����.Text) = "S") Then
        txt����.SelLength = Len(txt����.Text) - 1
    End If
End Sub

Private Sub txt����_KeyPress(KeyAscii As Integer)
    If Not ((KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or KeyAscii = vbKeyBack) Then KeyAscii = 0
End Sub

Private Sub txt�籣֤��_KeyPress(KeyAscii As Integer)
    If Not ((KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or KeyAscii = vbKeyBack) Then KeyAscii = 0
End Sub

Private Sub txt���֤��_KeyPress(KeyAscii As Integer)
    If Not ((KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or KeyAscii = vbKeyBack) Then KeyAscii = 0
End Sub

Private Sub txtȷ�Ͽ���_GotFocus()
    If Trim(txtȷ�Ͽ���.Text) = "" Then
        txtȷ�Ͽ���.Text = "S"
    End If
    txtȷ�Ͽ���.SelStart = 0
    If Not (Trim(txtȷ�Ͽ���.Text) = "" Or Trim(txtȷ�Ͽ���.Text) = "S") Then
        txtȷ�Ͽ���.SelLength = Len(txtȷ�Ͽ���.Text) - 1
    End If
End Sub

Private Sub txtȷ�Ͽ���_KeyPress(KeyAscii As Integer)
    If KeyAscii <> 13 Then
        If Not ((KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or KeyAscii = vbKeyBack) Then KeyAscii = 0
        Exit Sub
    End If
    
    KeyAscii = 0
    If Trim(txt����.Text) <> Trim(txtȷ�Ͽ���.Text) Then
        MsgBox "����������ֲ�Ų���ͬ�����ٴ�ȷ���ֲ�ţ�", vbInformation, gstrSysName
        txt����.SetFocus
        Exit Sub
    End If
    If Len(txt����.Text) < txt����.MaxLength Then
        MsgBox "�������������ֲ�ţ�����Ϊ" & txt����.MaxLength & "λ����", vbInformation, gstrSysName
        txt����.SetFocus
        Exit Sub
    End If
    If Not IsNumeric(Mid(txt����.Text, 1, txt����.MaxLength - 1)) Then
        MsgBox "�ֲ���к��зǷ��ַ�����ȷ�ϣ�", vbInformation, gstrSysName
        txt����.SetFocus
        Exit Sub
    End If
    
    '��ȡ���ò��˵Ļ�����Ϣ
    Call ReadPatient
    '��Ҫ��ȡ�ò����ڱ�Ժ����ʷ�����¼������������
    Call ReadBill
    
    '��ֹ����ѡ�񡢿�������
    Me.cbo��������.Enabled = False
    Me.cmd����.Enabled = False
    Me.txt����.Enabled = False
    Me.txtȷ�Ͽ���.Enabled = False
    
    fra������Ϣ.Enabled = True
    tabShow.Enabled = True
    cmd����(����).Enabled = True
    cmd����(����).Enabled = True
    cmd����(ɾ��).Enabled = True
    cmdOK.Enabled = True
    If cboҽ�����.Enabled Then cboҽ�����.SetFocus
End Sub

Private Sub txt��Ժ����_GotFocus()
    With txt��Ժ����
        .SelStart = 0
        .SelLength = 10
    End With
End Sub

Public Function GetIdentify(ByVal bytType As Byte, Optional lng����ID As Long) As String
    mlng����ID = lng����ID
    mbytType = bytType
    mstrReturn = ""
    Me.Show 1
    GetIdentify = mstrReturn
End Function

Private Function ValidData() As Boolean
    '�ԺϷ��Խ��м��
    Dim strValid As String
    Dim blnValid As Boolean
    Dim objBill As BillEdit
    Dim intPage As Integer, lngRow As Long, lngRows As Long
    On Error GoTo errHand
    '���������Ƿ�����
    '--�ı���ȫ���Ǳ�����
    If Not CheckTEXTBOX Then Exit Function
    '--������֤�ĺϷ���
    If Not (Len(txt���֤��.Text) = 15 Or Len(txt���֤��.Text) = 18) Then
        MsgBox "��������ȷ�����֤��Ϣ����λ��������", vbInformation, gstrSysName
        txt���֤��.SetFocus
        Exit Function
    End If
    '���ܷ�ֽ�������ڣ�����������������������Ϊ�Ƿ�
    If Len(txt���֤��.Text) = 15 Then
        '15λ
        strValid = "19" & Mid(txt���֤��.Text, 7, 2) & "-" & Mid(txt���֤��.Text, 9, 2) & "-" & Mid(txt���֤��.Text, 11, 2)
    Else
        '18λ
        strValid = Mid(txt���֤��.Text, 7, 4) & "-" & Mid(txt���֤��.Text, 11, 2) & "-" & Mid(txt���֤��.Text, 13, 2)
    End If
    blnValid = IsDate(strValid)
    If Not blnValid Then
        MsgBox "��������ȷ�����֤��Ϣ��", vbInformation, gstrSysName
        txt���֤��.SetFocus
        Exit Function
    End If
    
    '--�ټ�鹫��Ա������������������ȱʡ��ֵ�����ؼ��
    If cbo������ϵ.Visible Then
        If cbo������ϵ.ItemData(cbo������ϵ.ListIndex) = -1 Then
            MsgBox "��ѡ��������ϵ��", vbInformation, gstrSysName
            cbo������ϵ.SetFocus
            Exit Function
        End If
    Else
        If cbo����Ա.ItemData(cbo����Ա.ListIndex) = 0 Then
            If cbo����Ա����.ItemData(cbo����Ա����.ListIndex) = -1 Then
                MsgBox "��ѡ����Ա������", vbInformation, gstrSysName
                cbo����Ա����.SetFocus
                Exit Function
            End If
        End If
    End If
    
    '�����ʷ��¼
    For intPage = 0 To Bill.UBound
        Set objBill = Bill(intPage)
        lngRows = objBill.Rows - 1
        For lngRow = 1 To lngRows
            If Trim(objBill.TextMatrix(lngRow, col_ҽ�ƻ���)) <> "" Then
                '�����Ժ����/�������ڡ���Ժ�����Ƿ���д
                If Trim(objBill.TextMatrix(lngRow, col_��Ժ����)) = "" Then
                    MsgBox "�������" & lngRow & "�е�" & IIf(intPage = סԺ, "��Ժ���ڣ�", "�������ڣ�"), vbInformation, gstrSysName
                    tabShow.Tab = intPage
                    tabShow.SetFocus
                    Exit Function
                End If
                'ֻ��סԺ��Ҫ����Ժ����
                If intPage = סԺ Then
                    If Trim(objBill.TextMatrix(lngRow, col_��Ժ����)) = "" Then
                        MsgBox "�������" & lngRow & "�еĳ�Ժ���ڣ�", vbInformation, gstrSysName
                        tabShow.Tab = intPage
                        tabShow.SetFocus
                        Exit Function
                    End If
                End If
                '�����Ժ���͡���Ժ���͡�ҽ������Ƿ�����
                If Trim(objBill.TextMatrix(lngRow, col_��Ժ����)) = "" Then
                    MsgBox "��ѡ���" & lngRow & "�е�" & IIf(intPage = סԺ, "��Ժ���ͣ�", "ҽ�����"), vbInformation, gstrSysName
                    tabShow.Tab = intPage
                    tabShow.SetFocus
                    Exit Function
                End If
                If intPage = סԺ Then
                    If Trim(objBill.TextMatrix(lngRow, col_��Ժ����)) = "" Then
                        MsgBox "��ѡ���" & lngRow & "�еĳ�Ժ���ͣ�", vbInformation, gstrSysName
                        tabShow.Tab = intPage
                        tabShow.SetFocus
                        Exit Function
                    End If
                End If
            Else
                If lngRow < objBill.Rows - 1 Then
                    MsgBox "��ɾ����Ч���У�����" & lngRow & "���ǿ��У�", vbInformation, gstrSysName
                    tabShow.Tab = intPage
                    tabShow.SetFocus
                    Exit Function
                End If
            End If
        Next
    Next
    '�ٴμ�����ںϷ���
    For intPage = 0 To Bill.UBound
        Set objBill = Bill(intPage)
        lngRows = objBill.Rows - 1
        For lngRow = 1 To lngRows
            With objBill
                If Trim(.TextMatrix(lngRow, col_ҽ�ƻ���)) <> "" Then
                    If intPage = סԺ Then
                        '���ܴ��ڳ�Ժ����
                        If Trim(.TextMatrix(lngRow, col_��Ժ����)) <> "" Then
                            If .TextMatrix(lngRow, col_��Ժ����) > .TextMatrix(lngRow, col_��Ժ����) Then
                                MsgBox "��Ժ���ڲ��ܴ��ڳ�Ժ���ڣ�", vbInformation, gstrSysName
                                tabShow.Tab = intPage
                                tabShow.SetFocus
                                Exit Function
                            End If
                        End If
                        If lngRow > 1 Then
                            '����С��������¼�ĳ�Ժ����
                            If Trim(.TextMatrix(lngRow - 1, col_��Ժ����)) <> "" Then
                                If .TextMatrix(lngRow, col_��Ժ����) < .TextMatrix(lngRow - 1, col_��Ժ����) Then
                                    MsgBox "��Ժ���ڲ���С���ϴξ���ĳ�Ժ���ڣ�", vbInformation, gstrSysName
                                    tabShow.Tab = intPage
                                    tabShow.SetFocus
                                    Exit Function
                                End If
                            End If
                        End If
                        If lngRow + 1 <= .Rows - 1 Then
                            '���ܴ�����һ����¼����Ժ����
                            If Trim(.TextMatrix(lngRow + 1, col_��Ժ����)) <> "" Then
                                If .TextMatrix(lngRow, col_��Ժ����) > .TextMatrix(lngRow + 1, col_��Ժ����) Then
                                    MsgBox "��Ժ���ڲ��ܴ���" & lngRow + 1 & "�еǼǵ���Ժ���ڣ�", vbInformation, gstrSysName
                                    tabShow.Tab = intPage
                                    tabShow.SetFocus
                                    Exit Function
                                End If
                            End If
                        End If
                        '��Ժ���ڲ���С����Ժ����
                        If Trim(.TextMatrix(lngRow, col_��Ժ����)) <> "" Then
                            If .TextMatrix(lngRow, col_��Ժ����) < .TextMatrix(lngRow, col_��Ժ����) Then
                                MsgBox "��Ժ���ڲ���С����Ժ���ڣ�", vbInformation, gstrSysName
                                tabShow.Tab = intPage
                                tabShow.SetFocus
                                Exit Function
                            End If
                        End If
                        If lngRow + 1 <= .Rows - 1 Then
                            '��Ժ���ڲ��ܴ�����һ����¼����Ժ����
                            If Trim(.TextMatrix(lngRow + 1, col_��Ժ����)) <> "" Then
                                If .TextMatrix(lngRow, col_��Ժ����) > .TextMatrix(lngRow + 1, col_��Ժ����) Then
                                    MsgBox "��Ժ���ڲ��ܴ���" & lngRow + 1 & "�еǼǵ���Ժ���ڣ�", vbInformation, gstrSysName
                                    tabShow.Tab = intPage
                                    tabShow.SetFocus
                                    Exit Function
                                End If
                            End If
                        End If
                    Else
                        If lngRow > 1 Then
                            If Trim(.TextMatrix(lngRow - 1, col_��Ժ����)) <> "" Then
                                If .TextMatrix(lngRow, col_��Ժ����) < .TextMatrix(lngRow - 1, col_��Ժ����) Then
                                    MsgBox "�������ڲ���С���ϴξ������ڣ�", vbInformation, gstrSysName
                                    tabShow.Tab = intPage
                                    tabShow.SetFocus
                                    Exit Function
                                End If
                            End If
                        End If
                        If lngRow + 1 <= .Rows - 1 Then
                            '�������ڲ��ܴ�����һ����¼�ľ�������
                            If Trim(.TextMatrix(lngRow + 1, col_��Ժ����)) <> "" Then
                                If .TextMatrix(lngRow, col_��Ժ����) > .TextMatrix(lngRow + 1, col_��Ժ����) Then
                                    MsgBox "�������ڲ��ܴ���" & lngRow + 1 & "�еǼǵľ������ڣ�", vbInformation, gstrSysName
                                    tabShow.Tab = intPage
                                    tabShow.SetFocus
                                    Exit Function
                                End If
                            End If
                        End If
                    End If
                End If
            End With
        Next
    Next
    
    ValidData = True
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function

Private Function CheckTEXTBOX() As Boolean
    '����ı�����MaskEdit�����������Ƿ�Ϸ�
    Dim objControl As Control
    For Each objControl In Me.Controls
        Select Case UCase(TypeName(objControl))
        Case "TEXTBOX"
            If objControl.Enabled Then
                If Trim(objControl.Text) = "" Then
                    MsgBox "������" & Mid(objControl.Name, 4) & "��", vbInformation, gstrSysName
                    objControl.SetFocus
                    Exit Function
                End If
                If LenB(StrConv(objControl.Text, vbFromUnicode)) > objControl.MaxLength Then
                    MsgBox Mid(objControl.Name, 4) & "���������" & objControl.MaxLength & "���ַ�����", vbInformation, gstrSysName
                    objControl.SetFocus
                    Exit Function
                End If
            End If
        Case "MASKEDBOX"
            If objControl.Enabled Then
                If Not IsDate(objControl.Text) Then
                    MsgBox "������Ϸ���" & Mid(objControl.Name, 4) & "��", vbInformation, gstrSysName
                    objControl.SetFocus
                    Exit Function
                End If
            End If
        End Select
    Next
    CheckTEXTBOX = True
End Function

Private Function Getҽ�����(ByVal intPage As Integer, ByVal lngRow As Long) As Integer
    Dim strҽ����� As String
    '��ȡ��������õ�ҽ�����
    If intPage = סԺ Then
        'סԺ=21
        Getҽ����� = 21
    Else
        '�������ⲡ=12,��ͥ����=31
        strҽ����� = Bill(intPage).TextMatrix(lngRow, col_ҽ�����)
        Select Case strҽ�����
        Case "�������ⲡ"
            Getҽ����� = 12
        Case "��ͥ����"
            Getҽ����� = 31
        End Select
    End If
End Function

Private Function GetסԺ����(ByVal intPage As Integer, ByVal lngRow As Long, Optional ByVal bln��Ժ���� As Boolean = True) As Integer
    Dim strסԺ���� As String
    'ֻ��סԺ�Ŵ�����Ժ�������Ժ���ͣ�ȱʡȡ��Ժ���ͣ�����ȡ��Ժ����
    If intPage <> סԺ Then
        GetסԺ���� = 0
        Exit Function
    End If
    If bln��Ժ���� Then
        strסԺ���� = Bill(intPage).TextMatrix(lngRow, col_��Ժ����)
        Select Case strסԺ����
        Case "��ͨ"
            GetסԺ���� = 0
        Case "���ⲡ����"
            GetסԺ���� = 1
        Case "������ֲ"
            GetסԺ���� = 2
        Case "����"
            GetסԺ���� = 3
        Case "��ҽҽԺ��Ŀ�"
            GetסԺ���� = 4
        End Select
    Else
        strסԺ���� = Bill(intPage).TextMatrix(lngRow, col_��Ժ����)
        Select Case strסԺ����
        Case "����"
            GetסԺ���� = 0
        Case "ת��Ժ"
            GetסԺ���� = 1
        Case "��;����"
            GetסԺ���� = 2
        End Select
    End If
End Function

Private Function GetMoneySQL(ByVal intPage As Integer, ByVal lngRow As Long) As String
    Dim strReturn As String
    Dim intStart As Integer, intEnd As Integer
    Const intסԺ As Integer = 5
    Const int���� As Integer = 3
    '��ȡ��
    intEnd = Bill(intPage).Cols - 1
    If intPage = סԺ Then
        intStart = intסԺ
    Else
        intStart = int����
    End If
    
    strReturn = ""
    For intStart = intStart To intEnd
        strReturn = strReturn & "," & Val(Bill(intPage).TextMatrix(lngRow, intStart))
    Next
    strReturn = Mid(strReturn, 2)
    GetMoneySQL = strReturn
End Function
