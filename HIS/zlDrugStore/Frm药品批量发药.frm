VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FrmҩƷ������ҩ 
   Caption         =   "������ҩ"
   ClientHeight    =   6900
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10365
   Icon            =   "FrmҩƷ������ҩ.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6900
   ScaleWidth      =   10365
   StartUpPosition =   1  '����������
   Begin VB.Frame fra�ദ�� 
      Height          =   675
      Left            =   120
      TabIndex        =   16
      Top             =   520
      Visible         =   0   'False
      Width           =   9825
      Begin VB.CommandButton cmd���� 
         Caption         =   "����"
         Height          =   375
         Left            =   8520
         TabIndex        =   24
         Top             =   210
         Width           =   1095
      End
      Begin VB.CheckBox chkסԺ 
         Caption         =   "סԺ"
         Height          =   255
         Left            =   4200
         TabIndex        =   22
         Top             =   240
         Width           =   735
      End
      Begin VB.CheckBox chk���� 
         Caption         =   "����"
         Height          =   255
         Left            =   3480
         TabIndex        =   21
         Top             =   240
         Width           =   735
      End
      Begin VB.CheckBox chk���� 
         Caption         =   "����"
         Height          =   255
         Left            =   1680
         TabIndex        =   19
         Top             =   240
         Value           =   1  'Checked
         Width           =   855
      End
      Begin VB.CheckBox chk�շ� 
         Caption         =   "�շ�"
         Height          =   255
         Left            =   960
         TabIndex        =   18
         Top             =   240
         Value           =   1  'Checked
         Width           =   735
      End
      Begin VB.CommandButton cmd���˿��� 
         Height          =   250
         Left            =   8145
         Picture         =   "FrmҩƷ������ҩ.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   250
         Width           =   270
      End
      Begin VB.TextBox txt���˿��� 
         Height          =   300
         Left            =   6240
         TabIndex        =   35
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label lbl���˿��� 
         Caption         =   "���˿���"
         Height          =   255
         Left            =   5400
         TabIndex        =   23
         Top             =   270
         Width           =   735
      End
      Begin VB.Label lbl�����־ 
         AutoSize        =   -1  'True
         Caption         =   "�����־"
         Height          =   180
         Left            =   2640
         TabIndex        =   20
         Top             =   270
         Width           =   720
      End
      Begin VB.Label lbl�շ����� 
         AutoSize        =   -1  'True
         Caption         =   "��������"
         Height          =   180
         Left            =   120
         TabIndex        =   17
         Top             =   270
         Width           =   720
      End
   End
   Begin VB.CheckBox chk���� 
      Caption         =   "����������ѡ��"
      Height          =   255
      Index           =   1
      Left            =   8760
      TabIndex        =   36
      Top             =   300
      Width           =   1575
   End
   Begin VB.CheckBox chk���� 
      Caption         =   "�����˿���ѡ��"
      Height          =   255
      Index           =   0
      Left            =   8760
      TabIndex        =   33
      Top             =   50
      Width           =   1575
   End
   Begin VB.ComboBox Cboҩ�� 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   915
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   180
      Width           =   1560
   End
   Begin VB.CommandButton CmdHelp 
      Caption         =   "����(&H)"
      Height          =   350
      Left            =   90
      TabIndex        =   6
      Top             =   6030
      Width           =   1100
   End
   Begin VB.CommandButton CmdPrint 
      Caption         =   "��ӡ(&P)"
      Height          =   350
      Left            =   2580
      TabIndex        =   5
      Top             =   6030
      Visible         =   0   'False
      Width           =   1100
   End
   Begin VB.CommandButton CmdPrintSet 
      Caption         =   "����(&S)"
      Height          =   350
      Left            =   1350
      TabIndex        =   4
      Top             =   6030
      Visible         =   0   'False
      Width           =   1100
   End
   Begin TabDlg.SSTab TabShow 
      Height          =   2685
      Left            =   100
      TabIndex        =   1
      Top             =   3120
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   4736
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "������ϸ(&D)"
      TabPicture(0)   =   "FrmҩƷ������ҩ.frx":0E44
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Msf������ϸ"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "ҩƷ����(&T)"
      TabPicture(1)   =   "FrmҩƷ������ҩ.frx":0E60
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Msf��������"
      Tab(1).ControlCount=   1
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid Msf������ϸ 
         Height          =   2175
         Left            =   0
         TabIndex        =   7
         Top             =   360
         Width           =   9720
         _ExtentX        =   17145
         _ExtentY        =   3836
         _Version        =   393216
         FixedCols       =   0
         GridColor       =   -2147483625
         GridColorFixed  =   0
         AllowBigSelection=   0   'False
         FocusRect       =   0
         HighLight       =   0
         FillStyle       =   1
         GridLinesFixed  =   1
         AllowUserResizing=   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid Msf�������� 
         Height          =   2265
         Left            =   -75000
         TabIndex        =   8
         Top             =   360
         Width           =   9255
         _ExtentX        =   16325
         _ExtentY        =   3995
         _Version        =   393216
         FixedCols       =   0
         GridColor       =   -2147483625
         GridColorFixed  =   0
         AllowBigSelection=   0   'False
         FocusRect       =   0
         HighLight       =   0
         FillStyle       =   1
         GridLinesFixed  =   1
         AllowUserResizing=   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   7920
      TabIndex        =   3
      Top             =   6030
      Width           =   1100
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Msf�����б� 
      Height          =   1755
      Left            =   100
      TabIndex        =   0
      Top             =   1300
      Width           =   9825
      _ExtentX        =   17330
      _ExtentY        =   3096
      _Version        =   393216
      FixedCols       =   0
      ForeColorSel    =   -2147483640
      GridColor       =   -2147483631
      GridColorFixed  =   8421504
      AllowBigSelection=   0   'False
      FocusRect       =   0
      HighLight       =   0
      FillStyle       =   1
      GridLinesFixed  =   1
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "ȷ��(&O)"
      Enabled         =   0   'False
      Height          =   350
      Left            =   6720
      TabIndex        =   2
      Top             =   6030
      Width           =   1100
   End
   Begin MSComCtl2.DTPicker Dtp��ʼDate 
      Height          =   300
      Left            =   3360
      TabIndex        =   10
      Top             =   180
      Width           =   2085
      _ExtentX        =   3678
      _ExtentY        =   529
      _Version        =   393216
      CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
      Format          =   123207683
      CurrentDate     =   37007
   End
   Begin MSComCtl2.DTPicker Dtp����Date 
      Height          =   300
      Left            =   6480
      TabIndex        =   11
      Top             =   180
      Width           =   2085
      _ExtentX        =   3678
      _ExtentY        =   529
      _Version        =   393216
      CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
      Format          =   123207683
      CurrentDate     =   37007
   End
   Begin VB.CheckBox chkҩ�� 
      Caption         =   "ҩ��"
      Height          =   180
      Left            =   240
      TabIndex        =   14
      Top             =   240
      Value           =   1  'Checked
      Width           =   735
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   25
      Top             =   6540
      Width           =   10365
      _ExtentX        =   18283
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "FrmҩƷ������ҩ.frx":0E7C
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ��������Ϣ��ҵ��˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   13203
            Text            =   "δ�����κδ���"
            TextSave        =   "δ�����κδ���"
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
   Begin VB.Frame fra������ 
      Height          =   675
      Left            =   120
      TabIndex        =   26
      Top             =   520
      Width           =   9825
      Begin VB.TextBox TxtNo 
         Height          =   300
         Left            =   840
         TabIndex        =   31
         Top             =   240
         Width           =   1845
      End
      Begin VB.TextBox Txt���� 
         Height          =   300
         Left            =   6480
         MaxLength       =   12
         TabIndex        =   28
         Top             =   240
         Width           =   1845
      End
      Begin VB.TextBox txtҽ���� 
         Height          =   300
         Left            =   3720
         TabIndex        =   27
         Top             =   240
         Width           =   1845
      End
      Begin VB.Label LblNo 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "������"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   120
         TabIndex        =   32
         Top             =   270
         Width           =   540
      End
      Begin VB.Label Lbl���� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   6000
         TabIndex        =   30
         Top             =   270
         Width           =   360
      End
      Begin VB.Label lblҽ���� 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "ҽ����"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   3060
         TabIndex        =   29
         Top             =   270
         Width           =   540
      End
   End
   Begin VB.Label lblҩ�� 
      AutoSize        =   -1  'True
      Caption         =   "ҩ��"
      Height          =   180
      Left            =   480
      TabIndex        =   15
      Top             =   240
      Width           =   360
   End
   Begin VB.Label Lbl��ʼDate 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "��ʼ����"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   2520
      TabIndex        =   13
      Top             =   240
      Width           =   720
   End
   Begin VB.Label Lbl����Date 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "��������"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   5640
      TabIndex        =   12
      Top             =   240
      Width           =   720
   End
End
Attribute VB_Name = "FrmҩƷ������ҩ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'--�ⲿ���ݲ���--
Private mblnModify As Boolean
Private strPrivs As String
Private strUnit As String
Private mint������� As Integer                     'ҩ���ķ������1-���ﲡ��;2-סԺ����;3-�����סԺ
Private mlngҩ��ID As Long                          'ҩ��
Private IntSendAfterDosage As Integer               '����δ��ҩ��ҩ
Private Int����δ��˴�����ҩ As Integer            '����δ��˴�����ҩ
Private mint����δ�շѴ�����ҩ As Integer           '����δ�շѴ�����ҩ
Private IntCheckStock As Integer                    '�����
Private IntУ�鴦�� As Integer                      'У�鴦��
Private Str���� As String                           '��ҩ����
Private int����λ�� As Integer                  '���ý���λ��
Private int��˻��۵� As Integer                    'ִ�к��Զ���˻��۵�
Private mbln������ҩ������ As Boolean               '������ҩ������
Private mbln��ҩǰ�շѻ���� As Boolean             '��ҩǰ����Ҫ�շѻ������
Private mbln����ʱ����� As Boolean                 'ҩƷҽ��������ʱ��(�״�ʱ��)���ˣ�0-����������ʱ����ˣ�1-������ʱ�����
Private mint�����ʾ As Integer                     '�����ʾ��ʽ��0-��ʾӦ�ս��,1-��ʾʵ�ս��,2-��ʾӦ�պ�ʵ�ս��
Private mstrOpr As String
Private mblnLoadDrug As Boolean
Private mblnConPacker As Boolean                    '�����Զ���ҩ�Ƿ�����
Private mbln������ As Boolean

'--������ʹ�ñ���--
Private RecBill As New ADODB.Recordset              '���ݼ�¼
Private RecTotal As New ADODB.Recordset             '��������
Private BlnStartUp As Boolean
Private LngListRow As Long                          '�����б�
Private LngDetailRow As Long                        '������ϸ
Private LngTotalRow As Long                         '��������
Private StrBillNo As String                         '���ܵ��ݺ�
Private strID As String                             '����ID

Private mrsApplyforcredit As Recordset                  '���ڼ�¼������������ĵ���

Private LngBillCount As Long
Public str��ҩ�� As String
Public str�˲��� As String

Private rs��� As ADODB.Recordset

Private rs������Դ���� As ADODB.Recordset            '��¼���д���ҩ��������Դ����

Private rs����������ϸ As ADODB.Recordset            '��¼�������ܵļ�¼��ʵ���ǰ����ݺŵ���ϸ��¼
Private mstr���ܵ��� As String

Private mstrDeptNode As String          '��ǰҩ����վ��
Private mobjDrugMAC As Object
Private mobjPlugIn As Object             '��ҽӿڶ���

'���ݲ�������
Private Type Type_BillControl
    bln�Ƿ���� As Boolean
    intʱ������ As Integer
    bln���˵��� As Boolean
    dbl������� As Double
End Type
Private myBillControl As Type_BillControl

Public Property Get In_DrugMAC() As Object
    Set In_DrugMAC = mobjDrugMAC
End Property
Public Property Set In_DrugMAC(ByVal objVal As Object)
    Set mobjDrugMAC = objVal
End Property
Public Property Get In_PlugIn() As Object
    Set In_PlugIn = mobjPlugIn
End Property
Public Property Set In_PlugIn(ByVal objVal As Object)
    Set mobjPlugIn = objVal
End Property
Public Property Get In_�Զ���ҩ() As Boolean
    In_�Զ���ҩ = mblnConPacker
End Property

Public Property Let In_�Զ���ҩ(ByVal vNewValue As Boolean)
    mblnConPacker = vNewValue
End Property

Public Property Get In_���÷�ҩ() As Boolean
    In_���÷�ҩ = mblnLoadDrug
End Property

Public Property Let In_���÷�ҩ(ByVal vNewValue As Boolean)
    mblnLoadDrug = vNewValue
End Property
Public Property Get In_����() As String
    In_���� = mstrOpr
End Property

Public Property Let In_����(ByVal vNewValue As String)
    mstrOpr = vNewValue
End Property

Public Property Get In_������ҩ������() As Boolean
    In_������ҩ������ = mbln������ҩ������
End Property

Public Property Let In_������ҩ������(ByVal vNewValue As Boolean)
    mbln������ҩ������ = vNewValue
End Property
Public Property Get In_Ȩ��() As String
    In_Ȩ�� = strPrivs
End Property

Public Property Let In_Ȩ��(ByVal vNewValue As String)
    strPrivs = vNewValue
End Property

Public Property Get In_�������() As Integer
    In_������� = mint�������
End Property
Public Property Let In_�������(ByVal vNewValue As Integer)
    mint������� = vNewValue
End Property

Public Property Get In_У�鴦��() As Integer
    In_У�鴦�� = IntУ�鴦��
End Property

Public Property Let In_У�鴦��(ByVal vNewValue As Integer)
    IntУ�鴦�� = vNewValue
End Property

Public Property Get In_�����() As Integer
    In_����� = IntCheckStock
End Property

Public Property Let In_�����(ByVal vNewValue As Integer)
    IntCheckStock = vNewValue
End Property

Public Property Get In_ҩ��ID() As Long
    In_ҩ��ID = mlngҩ��ID
End Property

Public Property Let In_ҩ��ID(ByVal vNewValue As Long)
    mlngҩ��ID = vNewValue
End Property

Public Property Get In_��ҩ����() As String
    In_��ҩ���� = Str����
End Property

Public Property Let In_��ҩ����(ByVal vNewValue As String)
    Str���� = vNewValue
End Property

Public Property Get In_����δ��ҩ��ҩ() As Integer
    In_����δ��ҩ��ҩ = IntSendAfterDosage
End Property

Public Property Let In_����δ��ҩ��ҩ(ByVal vNewValue As Integer)
    IntSendAfterDosage = vNewValue
End Property

Public Property Get IN_����δ��˷�ҩ() As Integer
    IN_����δ��˷�ҩ = Int����δ��˴�����ҩ
End Property

Public Property Let IN_����δ��˷�ҩ(ByVal vNewValue As Integer)
    Int����δ��˴�����ҩ = vNewValue
End Property

Public Property Get IN_����δ�շѷ�ҩ() As Integer
    IN_����δ�շѷ�ҩ = mint����δ�շѴ�����ҩ
End Property

Public Property Let IN_����δ�շѷ�ҩ(ByVal vNewValue As Integer)
    mint����δ�շѴ�����ҩ = vNewValue
End Property

Public Property Get In_����λ��() As Integer
    In_����λ�� = int����λ��
End Property

Public Property Let In_����λ��(ByVal vNewValue As Integer)
    int����λ�� = vNewValue
End Property

Public Property Get IN_��˻��۵�() As Integer
    IN_��˻��۵� = int��˻��۵�
End Property

Public Property Let IN_��˻��۵�(ByVal vNewValue As Integer)
    int��˻��۵� = vNewValue
End Property

Private Function CheckBillOperate() As Boolean
    Dim n, i As Integer
    Dim Dbl��� As Double
    
    For n = 1 To Msf�����б�.rows - 1
        If Msf�����б�.TextMatrix(n, 2) <> "" Then
            Msf�����б�.Row = n
            Call Msf�����б�_EnterCell
            DoEvents
            
            Dbl��� = 0
            
            For i = 1 To Msf������ϸ.rows - 2
                Dbl��� = Dbl��� + Val(Msf������ϸ.TextMatrix(i, 7))
            Next
            
            If CheckBillControl(3, Val(Msf�����б�.RowData(n)), Msf�����б�.TextMatrix(n, 2), Dbl���) = False Then
                Exit Function
            End If
        End If
    Next
    
    CheckBillOperate = True
End Function

Private Function CheckDrugStock() As Boolean
    Dim rsTmp As ADODB.Recordset
    Dim n As Integer
    Dim lngTemp�ⷿid As Long
    
    If mbln������ҩ������ = False Then
        lngTemp�ⷿid = Cboҩ��.ItemData(Cboҩ��.ListIndex)
    Else
        lngTemp�ⷿid = mlngҩ��ID
    End If
    
    If Msf�����б�.TextMatrix(1, 2) <> "" Then
        For n = 1 To Msf��������.rows - 2
            If MediWork_CheckStorageStock(lngTemp�ⷿid, Val(Msf��������.TextMatrix(n, 9))) = False Then
                MsgBox Msf��������.TextMatrix(n, 1) & "δ���ô洢�ⷿ�����ܷ�ҩ��", vbInformation, gstrSysName
                Exit Function
            End If
        Next
    End If
    
    CheckDrugStock = True
End Function

Private Sub GetRecipe(ByVal intType As Integer, ByVal txtInput As TextBox)
    'intType��1�������ţ�2��ҽ���ţ�3����������
    Dim RecRecord As New ADODB.Recordset
    Dim intYear As Integer, strYear As String
    Dim intRow As Integer
    Dim strNo As String, IntBill As Integer, ArrTmp, strTmp As String
    Dim strCon As String
    Dim strBeginDate As String
    Dim strEndDate As String
    Dim strInput As String
    Dim strSqlFrom As String
    Dim lngҩ��ID As Long
    Dim strsql As String
    Dim int��¼���� As Integer
    Dim int�����־ As Integer
    
    On Error GoTo errHandle
    
    If Trim(txtInput.Text) = "" Then Exit Sub
    If intType <> 3 Then
        strInput = Trim(UCase(txtInput.Text))
    Else
        strInput = Trim(txtInput.Text)
    End If
    
    If Me.Cboҩ��.ListIndex = -1 Then Exit Sub
        
    strBeginDate = Format(Dtp��ʼDate.Value, "yyyy-MM-dd hh:mm:ss")
    strEndDate = Format(Dtp����Date.Value, "yyyy-MM-dd hh:mm:ss")
        
    If intType = 1 Then
        strCon = " And C.No=[1] "
    ElseIf intType = 2 Then
        strCon = " And D.ҽ����=[1] "
    Else
        strCon = " And B.���� Like [2] "
    End If
    
    If mbln������ҩ������ Then
        strCon = strCon & " And C.��¼״̬=1 And C.�������� Between To_Date([4],'yyyy-MM-dd hh24:mi:ss') And To_Date([5] ,'yyyy-MM-dd hh24:mi:ss') "
    Else
        strCon = strCon & " And Mod(C.��¼״̬,3)=1 And C.�������� Between To_Date([4] ,'yyyy-MM-dd hh24:mi:ss') And To_Date([5] ,'yyyy-MM-dd hh24:mi:ss') "
    End If
    
    '������ҩ��֧��ˢ�����ѣ�������ȡ�ı��������շѻ�����˵Ĵ���
    gstrSQL = " Select Distinct S.���� As ҩ��, Decode(C.����,8,'�շ�',9,'����') ����,C.No,C.����,A.���շ�,Decode(A.��ҩ��,Null,'','���ŷ�ҩ','',A.��ҩ��) ��ҩ��,P.���� ����,B.����,B.��ʶ�� סԺ��,'' ����," & _
             " B.������ ����ҽ��,B.����Ա���� ������,To_Char(C.��������,'yyyy-MM-dd') ��������, S.ID As ҩ��id,B.��¼����,B.�����־,A.����ID, d.�������� " & _
             " From δ��ҩƷ��¼ A,������ü�¼ B,ҩƷ�շ���¼ C,���ű� P,���ű� S, ������Ϣ D " & IIf(mbln������, ",��������¼ Q,���������ϸ K ", "") & _
             " Where C.����ID=B.ID And B.��������ID+0=P.ID And Nvl(C.�ⷿID,0)+0=S.ID and Nvl(A.�ⷿID,0)=Nvl(C.�ⷿID,0)  And A.No=C.No " & IIf(mbln������, " and b.ҽ�����=k.ҽ��id(+) and Q.id(+)=K.��id and K.����ύ(+)=1 And ((b.ҽ����� is null or nvl(q.�����,0) = 1) or not Exists(select 1 from ��������¼ Q where q.����id = d.����id and q.�ύ����id = p.id and q.��ҩҩ��id = b.ִ�в���id And q.id = k.��id))", "") & _
             IIf(mbln������ҩ������, "", IIf(Str���� = "", "", " And (C.��ҩ���� In (Select * From Table(Cast(f_Str2list([7]) As Zltools.t_Strlist))) Or C.��ҩ���� Is NULL)")) & _
             " And C.����� Is Null And Nvl(B.����״̬,0)<>1 " & _
             " and Not Exists(select 1 from ҩƷ�շ���¼ F where F.����=C.���� and F.�ⷿid=C.�ⷿid and F.no=C.no and ��ҩ��ʽ=-1) " & _
             " And C.����=A.���� and nvl(C.��ҩ��ʽ,-999)<>-1 And A.����id=D.����id(+) " & strCon
    
    If Me.chk�շ�.Value = 1 And Me.chk����.Value = 1 Then
        If mbln��ҩǰ�շѻ���� = True Then
            gstrSQL = gstrSQL & " And A.���� In(8,9) And A.���շ�=1 "
        ElseIf mint����δ�շѴ�����ҩ = False Then
            gstrSQL = gstrSQL & " And (C.����=8 And A.���շ�=1 Or C.����=9) "
        ElseIf Int����δ��˴�����ҩ = False Then
            gstrSQL = gstrSQL & " And (C.����=9 And A.���շ�=1 Or C.����=8) "
        Else
            gstrSQL = gstrSQL & " And A.���� In(8,9) "
        End If
    ElseIf Me.chk�շ�.Value = 1 Then
        If mbln��ҩǰ�շѻ���� = True Or mint����δ�շѴ�����ҩ = False Then
            gstrSQL = gstrSQL & " And A.����=8 And A.���շ�=1 "
        Else
            gstrSQL = gstrSQL & " And A.����=8 "
        End If
    ElseIf Me.chk����.Value = 1 Then
        If mbln��ҩǰ�շѻ���� = True Or mint����δ�շѴ�����ҩ = False Then
            gstrSQL = gstrSQL & " And A.����=9 And A.���շ�=1 "
        Else
            gstrSQL = gstrSQL & " And A.����=9 "
        End If
    End If
    
    If mstrDeptNode <> "" Then
        gstrSQL = gstrSQL & " And (P.վ�� = [8] Or P.վ�� Is Null) "
    End If
    
    If mbln������ҩ������ = True Then
        If chkҩ��.Value = 1 Then
            gstrSQL = gstrSQL & " And C.�ⷿID+0=[3] "
        Else
            gstrSQL = gstrSQL & " And C.�ⷿID+0<>[6] "
        End If
    Else
        gstrSQL = gstrSQL & " And (C.�ⷿID+0=[3] OR C.�ⷿID IS NULL)"
    End If
    
    If mbln����ʱ����� = True Then
        strsql = gstrSQL & " And B.ҽ����� Is Null"
        gstrSQL = Replace(gstrSQL, "C.��������", "B.����ʱ��") & " And B.ҽ����� Is Not Null"
        gstrSQL = strsql & " Union All " & gstrSQL
    End If
    
    If mint������� = 3 Then
        strsql = Replace(gstrSQL, "'' ����", "B.����")
        strsql = Replace(strsql, "������ü�¼", "סԺ���ü�¼")
        strsql = Replace(strsql, "And Nvl(B.����״̬,0)<>1", "")
        gstrSQL = gstrSQL & " Union All " & strsql
    ElseIf mint������� = 2 Then
        gstrSQL = Replace(gstrSQL, "'' ����", "B.����")
        gstrSQL = Replace(gstrSQL, "������ü�¼", "סԺ���ü�¼")
        gstrSQL = Replace(gstrSQL, "And Nvl(B.����״̬,0)<>1", "")
    End If
    
    Set RecBill = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, strInput, strInput & "%", Me.Cboҩ��.ItemData(Me.Cboҩ��.ListIndex), strBeginDate, strEndDate, mlngҩ��ID, Str����, mstrDeptNode)
    
    If RecBill.EOF Then
        
        If intType = 1 Then
            If CheckBillExist(strInput, mlngҩ��ID) = False Then GoTo ExitSub
        End If
        
        MsgBox "δ�ҵ�ָ����������ָ���Ĵ���δ�շѻ�δ��ˣ����������룡", vbInformation, gstrSysName
        GoTo ExitSub
    End If
    
    If RecBill.RecordCount > 1 Then
        strTmp = Frm����ѡ��.ShowMe(Me, RecBill)
        If strTmp = "" Then GoTo ExitSub
        
        ArrTmp = Split(strTmp, ";")
        strNo = ArrTmp(0)
        IntBill = ArrTmp(1)
        lngҩ��ID = ArrTmp(2)
                
        RecBill.MoveFirst
'        RecBill.Find "����=" & IntBill
        RecBill.Filter = "����=" & IntBill & " And No='" & strNo & "'"
        
        int��¼���� = RecBill!��¼����
        int�����־ = RecBill!�����־
    Else
        strNo = RecBill!NO
        IntBill = RecBill!����
        lngҩ��ID = RecBill!ҩ��ID
        int��¼���� = RecBill!��¼����
        int�����־ = RecBill!�����־
    End If
    
    Me.TxtNo = strNo
    Me.TxtNo.Tag = IntBill
    Me.Txt����.Tag = IntBill
    
    '����Ѵ��ڸõ��ݣ����˳�
    If SetLocateBill(False) Then
        MsgBox "�ô����Ѿ����룬�����䣡", vbInformation, gstrSysName
        GoTo ExitSub
    End If
    
    '�����ǰ���봦���Ŀ�������¼��Ĵ����Ŀ��Ҳ�ͬ���������ʾ
    If CheckSource(IntBill, strNo, lngҩ��ID) = False Then Exit Sub
    If WriteSendListData(0) = False Then GoTo ExitSub
    
    LngBillCount = LngBillCount + 1
    Me.stbThis.Panels(2).Text = IIf(LngBillCount = 0, "δ�����κδ���", "������" & LngBillCount & "�Ŵ���")
    '��λ���ղ�����Ĵ�����
    Call SetLocateBill
    
    With Msf�����б�
        CmdOK.Enabled = (.RowData(IIf(.rows - 1 = 1, 1, .rows - 2)) <> 0)
    End With
    
    mblnModify = True
'    If tabShow.Tab = 1 Then Call RefreshData
    Call RefreshData(lngҩ��ID)
    With TxtNo
        .SelStart = 0
        .SelLength = Len(txtInput)
    End With
    Exit Sub
    
ExitSub:
    With txtInput
        .SelStart = 0
        .SelLength = Len(txtInput)
        .SetFocus
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
Private Function CheckBillExist(ByVal strNo As String, ByVal lngҩ��ID As Long) As Boolean
    Dim rstemp As ADODB.Recordset
    
    On Error GoTo errHandle
    gstrSQL = " Select A.��ҩ��,A.�����,C.����Ա���� ������ " & _
            " From ҩƷ�շ���¼ A, ������ü�¼ C " & _
            " Where A.����id = C.ID And mod(A.��¼״̬,3)=1 And Rownum=1 " & _
            " And A.No=[1] And A.���� in (8,9,10)"
    
    If mbln������ҩ������ = True Then
        If chkҩ��.Value = 1 Then
            gstrSQL = gstrSQL & " And A.�ⷿID+0=[2] "
        Else
            gstrSQL = gstrSQL & " And A.�ⷿID+0<>[2] "
        End If
    Else
        gstrSQL = gstrSQL & " And (A.�ⷿID+0=[2] OR A.�ⷿID IS NULL)"
    End If
    
    gstrSQL = gstrSQL & "Union All" & Replace(gstrSQL, "������ü�¼", "סԺ���ü�¼")
    
    Set rstemp = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, strNo, lngҩ��ID)
        
    With rstemp
        If .EOF Then CheckBillExist = False: MsgBox "�ô���[" & strNo & "]�����ڣ�", vbInformation, gstrSysName: Exit Function
        If Not IsNull(!�����) Then
            CheckBillExist = False
            MsgBox "�ô���[" & strNo & "]�ѱ���������Ա��ҩ��", vbInformation, gstrSysName: Exit Function
        End If
    End With

    CheckBillExist = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Private Sub IniControl()
    LngBillCount = 0
    Me.stbThis.Panels(2).Text = IIf(LngBillCount = 0, "δ�����κδ���", "������" & LngBillCount & "�Ŵ���")
    
    '��ʼ��
    strID = ""
    StrBillNo = ""
    TxtNo = ""
    Txt���� = ""
    
    With Msf��������
        .Clear
        .rows = 2
        .RowData(1) = 0
    End With
    With Msf�����б�
        .Clear
        .rows = 2
        .RowData(1) = 0
    End With
    With Msf������ϸ
        .Clear
        .rows = 2
        .RowData(1) = 0
    End With
    
    Call SetFormat(1)
    Call SetFormat(2)
    Call SetFormat(3)
    
    Call InitRec
    CmdOK.Enabled = False
    If Me.chk����(0).Value = 0 And Me.chk����(1).Value = 0 Then TxtNo.SetFocus
End Sub

Private Sub InitRecSum()
    '��ʼ���������ݼ�
    Set rs����������ϸ = New ADODB.Recordset
    With rs����������ϸ
        If .State = 1 Then .Close
        .Fields.Append "���ݺ�", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "ҩƷ����", adLongVarChar, 40, adFldIsNullable
        .Fields.Append "��Ʒ��", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "����", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "���", adLongVarChar, 40, adFldIsNullable
        .Fields.Append "��λ", adLongVarChar, 10, adFldIsNullable
        .Fields.Append "����", adDouble, 18, adFldIsNullable
        .Fields.Append "����", adDouble, 18, adFldIsNullable
        .Fields.Append "���", adDouble, 18, adFldIsNullable
        .Fields.Append "ʵ�ս��", adDouble, 18, adFldIsNullable
        .Fields.Append "ҩƷid", adDouble, 18, adFldIsNullable
        .Fields.Append "����", adDouble, 18, adFldIsNullable
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
End Sub

Private Sub Iniҩ��()
    Dim rstemp As ADODB.Recordset
    Dim n As Integer
    Dim lngTemp As Long
    
    On Error GoTo errHandle

    Me.Cboҩ��.Enabled = mbln������ҩ������
    
    If zlStr.IsHavePrivs(strPrivs, "����ҩ��") Or mbln������ҩ������ Then
        gstrSQL = "(Select Distinct ����ID From ��������˵�� Where �������� Like '%ҩ��')"
    Else
        gstrSQL = "(Select distinct A.����ID From ������Ա A,��������˵�� B " & _
                 " Where A.��ԱID=[1] And A.����ID=B.����ID And B.�������� Like '%ҩ��')"
    End If
    gstrSQL = " Select Distinct P.ID,P.���� From ���ű� P " & _
             " Where (P.վ�� = '" & gstrNodeNo & "' Or P.վ�� is Null) And P.ID In " & gstrSQL & _
             " And (P.����ʱ�� Is Null Or P.����ʱ��=To_Date('3000-01-01','yyyy-MM-dd'))"
    Set rstemp = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, glngUserId)
    
    With Me.Cboҩ��
        Do While Not rstemp.EOF
            If Not mbln������ҩ������ Then
                .AddItem rstemp!����
                n = .NewIndex
                .ItemData(n) = rstemp!Id
                
                If lngTemp = 0 Then
                    If rstemp!Id = mlngҩ��ID Then
                        lngTemp = n
                    End If
                End If
            Else
                If rstemp!Id <> mlngҩ��ID Then
                    .AddItem rstemp!����
                    n = .NewIndex
                    .ItemData(n) = rstemp!Id
                Else
                    Me.Caption = "������ҩ������(��ǰҩ����" & rstemp!���� & ")"
                End If
            End If
            
            rstemp.MoveNext
        Loop
        
        .ListIndex = lngTemp
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
Private Sub SetFormat(Optional ByVal IntStyle As Integer = 1)
    Dim intCol As Integer
    '���ø��б�ؼ��ĸ�ʽ

    Select Case IntStyle
    Case 1
        With Msf�����б�
            .rows = 2
            .Cols = 15
            
            .TextMatrix(0, 0) = "ҩ��"
            .TextMatrix(0, 1) = "����"
            .TextMatrix(0, 2) = "NO"
            .TextMatrix(0, 3) = "����"
            .TextMatrix(0, 4) = "����"
            .TextMatrix(0, 5) = "סԺ��"
            .TextMatrix(0, 6) = "����"
            .TextMatrix(0, 7) = "�շ�Ա"
            .TextMatrix(0, 8) = "����ҽ��"
            .TextMatrix(0, 9) = "��������"
            .TextMatrix(0, 10) = "ҩ��ID"
            .TextMatrix(0, 11) = "��¼����"
            .TextMatrix(0, 12) = "�����־"
            .TextMatrix(0, 13) = "���շ�"
            .TextMatrix(0, 14) = "����ID"
            
            For intCol = 0 To .Cols - 1
                .ColAlignmentFixed(intCol) = 4
            Next
            
            If BlnStartUp = False Then
                .ColWidth(0) = IIf(mbln������ҩ������ = True And chkҩ��.Value = 0, 1500, 0)
                .ColWidth(1) = 500
                .ColWidth(2) = 1000
                .ColWidth(3) = 1200
                .ColWidth(4) = 1000
                .ColWidth(5) = 1000
                .ColWidth(6) = 800
                .ColWidth(7) = 1000
                .ColWidth(8) = 1000
                .ColWidth(9) = 1200
                .ColWidth(10) = 0
                .ColWidth(11) = 0
                .ColWidth(12) = 0
                .ColWidth(13) = 0
                .ColWidth(14) = 0
                
                .Row = 1
                Call RestoreFlexState(Msf�����б�, Me.Name)
                If glngSys \ 100 <> 1 Then
                    .ColWidth(3) = 0
                    .ColWidth(5) = 0
                    .ColWidth(6) = 0
                End If
                .ColWidth(8) = IIf(IntУ�鴦�� = 1, 0, 1000)
            End If
        End With
    Case 2
        With Msf������ϸ
            .rows = 2
            .Cols = 8
    
            .TextMatrix(0, 0) = "ҩƷ����"
            .TextMatrix(0, 1) = "��Ʒ��"
            .TextMatrix(0, 2) = "���"
            .TextMatrix(0, 3) = "��λ"
            .TextMatrix(0, 4) = "����"
            .TextMatrix(0, 5) = "����"
            .TextMatrix(0, 6) = "Ӧ�ս��"
            .TextMatrix(0, 7) = "ʵ�ս��"
            
            For intCol = 0 To .Cols - 1
                .ColAlignmentFixed(intCol) = 4
            Next
    
            If BlnStartUp = False Then
                .ColWidth(0) = 2000
                .ColWidth(2) = 1500
                .ColWidth(3) = 500
                .ColWidth(4) = 800
                .ColWidth(5) = 800
                .ColWidth(6) = 1000
                .ColWidth(7) = 1000
                
                .Row = 1
                Call RestoreFlexState(Msf������ϸ, Me.Name)
                If gintҩƷ������ʾ = 2 Then
                    If .ColWidth(1) = 0 Then .ColWidth(1) = 2000
                Else
                    .ColWidth(1) = 0
                End If
                
                If mint�����ʾ = 0 Then
                    .ColWidth(7) = 0
                    If .ColWidth(6) <= 0 Then .ColWidth(6) = 1000
                ElseIf mint�����ʾ = 1 Then
                    .ColWidth(6) = 0
                    If .ColWidth(7) <= 0 Then .ColWidth(7) = 1000
                Else
                    If .ColWidth(6) <= 0 Then .ColWidth(6) = 1000
                    If .ColWidth(7) <= 0 Then .ColWidth(7) = 1000
                End If
                
            End If
        End With
    Case 3
        With Msf��������
            .rows = 2
            .Cols = 11
    
            .TextMatrix(0, 0) = "���"
            .TextMatrix(0, 1) = "ҩƷ����"
            .TextMatrix(0, 2) = "��Ʒ��"
            .TextMatrix(0, 3) = "���"
            .TextMatrix(0, 4) = "��λ"
            .TextMatrix(0, 5) = "����"
            .TextMatrix(0, 6) = "����"
            .TextMatrix(0, 7) = "Ӧ�ս��"
'            .TextMatrix(0, 8) = "ҩƷid"
'            .TextMatrix(0, 9) = "����"
            .TextMatrix(0, 8) = "ʵ�ս��"
            .TextMatrix(0, 9) = "ҩƷid"
            .TextMatrix(0, 10) = "����"

            For intCol = 0 To .Cols - 1
                .ColAlignmentFixed(intCol) = 4
            Next
            
            If BlnStartUp = False Then
                .ColWidth(0) = 500
                .ColWidth(1) = 2000
                .ColWidth(3) = 1500
                .ColWidth(4) = 500
                .ColWidth(5) = 800
                .ColWidth(6) = 800
                .ColWidth(7) = 1000
                .ColWidth(8) = 1000
                .ColWidth(9) = 0
                .ColWidth(10) = 0
                .Row = 1
                Call RestoreFlexState(Msf��������, Me.Name)
                
                If gintҩƷ������ʾ = 2 Then
                    If .ColWidth(2) = 0 Then .ColWidth(2) = 2000
                Else
                    .ColWidth(2) = 0
                End If
                
                If mint�����ʾ = 0 Then
                    .ColWidth(8) = 0
                    If .ColWidth(7) <= 0 Then .ColWidth(7) = 1000
                ElseIf mint�����ʾ = 1 Then
                    .ColWidth(7) = 0
                    If .ColWidth(8) <= 0 Then .ColWidth(8) = 1000
                Else
                    If .ColWidth(7) <= 0 Then .ColWidth(7) = 1000
                    If .ColWidth(8) <= 0 Then .ColWidth(8) = 1000
                End If
            End If
        End With
    End Select
End Sub



Private Sub chk����_Click()
    If Me.chk����.Value = 0 Then
        Me.chk�շ�.Value = 1
    End If
End Sub

Private Sub chk����_Click(index As Integer)
    Dim i As Integer
    
    
    If index = 0 Then
        If Me.chk����(0).Value = 1 Then
            If Me.chk����(1).Value = 1 Then
                Me.chk����(1).Value = 0
            End If
            Me.lbl���˿���.Caption = "���˿���"
        End If
        
    Else
        If Me.chk����(1).Value = 1 Then
            If Me.chk����(0).Value = 1 Then
                Me.chk����(0).Value = 0
            End If
            Me.lbl���˿���.Caption = "��������"
        End If
    End If
    
    If Me.chk����(0).Value = 0 And Me.chk����(1).Value = 0 Then
        Me.fra�ദ��.Visible = False
        Me.fra������.Visible = True

    Else
        Me.fra�ദ��.Visible = True
        Me.fra������.Visible = False
    End If



    If rs����������ϸ.RecordCount > 0 Then
        Set RecTotal = Nothing
        With Msf�����б�
            .Clear
            .rows = 2
            Call SetFormat(1)
        End With

        With Msf������ϸ
            .Clear
            .rows = 2
            Call SetFormat(2)
        End With

        With Msf��������
            .Clear
            .rows = 2
            Call SetFormat(3)
        End With

        rs����������ϸ.MoveLast

        For i = 0 To rs����������ϸ.RecordCount - 1
            ''''ɾ����ǰ��
            rs����������ϸ.Delete adAffectCurrent
            ''''��ǰ�ƶ�ָ��
            rs����������ϸ.MovePrevious
        Next

        Me.stbThis.Panels(2).Text = "δ�����κδ���"
    End If
End Sub

Private Sub chk����_Click()
    If Me.chkסԺ.Enabled = False Then
        Me.chk����.Value = 1
    Else
        If Me.chk����.Value = 0 Then
            Me.chkסԺ.Value = 1
        End If
    End If
End Sub

Private Sub chk�շ�_Click()
    If Me.chk�շ�.Value = 0 Then
        Me.chk����.Value = 1
    End If
End Sub

Private Sub chkҩ��_Click()
    IniControl
    
    If chkҩ��.Value = 1 Then
        Cboҩ��.Enabled = True
        Msf�����б�.ColWidth(0) = 0
    Else
        Cboҩ��.Enabled = False
        If Msf�����б�.ColWidth(0) = 0 Then
            Msf�����б�.ColWidth(0) = 1500
        End If
    End If
End Sub

Private Sub chkסԺ_Click()
    If Me.chk����.Enabled = False Then
        Me.chkסԺ.Value = 1
     Else
        If Me.chkסԺ.Value = 0 Then
            Me.chk����.Value = 1
        End If
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
    Exit Sub
End Sub

Private Sub CmdHelp_Click()
    Call ShowHelp(App.ProductName, Me.hWnd, Me.Name)
End Sub

Private Sub cmdOk_Click()
    '���õ���ǩ��ʱ����û��Ƿ�ע��
    If gblnESign������ҩ = True Then
        If Not gobjESign.CheckCertificate(gstrDbUser) Then
            Exit Sub
        End If
    End If
    
'    Call RefreshData
     '��鵥���Ƿ��ǵ���ĵ���
    If mbln������ҩ������ Then
        If CheckDate = False Then Exit Sub
    End If
    
    If CheckDrugStock = False Then Exit Sub
    If CheckStock = False Then Exit Sub
    If Not CheckCorrelation Then Exit Sub
    If Not CheckBillOperate Then Exit Sub
    If SendBill = False Then Exit Sub
    
    IniControl
End Sub

Private Function CheckDate() As Boolean
'���ڷ�����ҩ������ʱ������Ƿ��ǵ���ĵ���
    Dim i As Integer
    Dim dateCur As Date
    
    dateCur = Sys.Currentdate
    With Msf�����б�
        For i = 1 To .rows - 1
            If .TextMatrix(i, 2) <> "" Then
                If Format(.TextMatrix(i, 9), "YYYY-MM-DD") < Format(dateCur, "YYYY-MM-DD") Then
                    If MsgBox("        �����ǵ��쵥�ݣ���ɾ�������������»��ܣ�" & vbCrLf & "����Ѿ����˱���Ŀ�����Ҫ���³������Ƿ����������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                        CheckDate = False
                    Else
                        CheckDate = True
                    End If
                    Exit Function
                End If
            End If
        Next
    End With
    
    CheckDate = True
End Function

Private Sub cmdPrint_Click()
    Dim HisPrint As New zlPrint1Grd
    Dim HisRow As New zlTabAppRow
    Dim ArrayNo, IntArray As Integer
    Dim LngSelectRow As Long, intCol As Integer
    
    On Error Resume Next
    'ȡ������ѡ��״̬
    With Msf��������
        .Redraw = False
        
        LngSelectRow = .Row     '���浱ǰѡ����
        If LngTotalRow > 0 And LngTotalRow < .rows Then
            .Row = LngTotalRow       '����ϴ�ѡ����
            For intCol = 0 To .Cols - 1
                    .Col = intCol
                    .CellBackColor = &H80000005
                    .CellForeColor = &H80000008
            Next
            .Col = 0
        End If
    End With
    
    HisPrint.Title = "ҩƷ����"
    Set HisRow = New zlTabAppRow
    HisRow.Add "����:" & Format(Sys.Currentdate, "yyyy��MM��dd��")
    HisPrint.UnderAppRows.Add HisRow
    
    ArrayNo = Split(StrBillNo, ";")
    
    Set HisRow = New zlTabAppRow
    HisRow.Add "���ݺ�:"
    HisPrint.BelowAppRows.Add HisRow
    For IntArray = 0 To UBound(ArrayNo)
        Set HisRow = New zlTabAppRow
        HisRow.Add Space(10) & ArrayNo(IntArray)
        HisPrint.BelowAppRows.Add HisRow
    Next
    
    Set HisPrint.Body = Msf��������
    Select Case zlPrintAsk(HisPrint)
    Case 1
        zlPrintOrView1Grd HisPrint, 1
    Case 2
        zlPrintOrView1Grd HisPrint, 2
    Case 3
        zlPrintOrView1Grd HisPrint, 3
    End Select
    
    '�ָ�����ѡ��״̬
    With Msf��������
        
        LngTotalRow = LngSelectRow
        .Row = LngTotalRow       '���õ�ǰѡ����
        For intCol = 0 To .Cols - 1
                .Col = intCol
                .CellBackColor = &H8000000D
                .CellForeColor = &H80000005
        Next
        .Col = 0
        .Redraw = True
    End With
End Sub

Private Sub CmdPrintSet_Click()
    zlPrintSet
End Sub

Private Sub cmd���˿���_Click()
'���','����','����','����','Ӫ��'
    If Me.lbl���˿���.Caption = "���˿���" Then
        If Select����(Me, Me.txt���˿���, "", "�ٴ�", False, mint�������) = False Then
            Exit Sub
        End If
    Else
        If Select����(Me, Me.txt���˿���, "", "���,����,����,����,Ӫ��", False, mint�������) = False Then
            Exit Sub
        End If
    End If
    
End Sub

Private Sub cmd����_Click()
    Dim strBeginDate As String
    Dim strEndDate As String
    Dim strCon As String
    Dim strsql As String
    Dim i As Integer
    
    On Error GoTo errHandle
    If Me.chk����.Value = 0 And Me.chk�շ�.Value = 0 Then
        MsgBox "��ѡ���շ����ͣ�", vbInformation + vbOKOnly, gstrSysName
        Exit Sub
    End If
    
    If Me.chk����.Value = 0 And Me.chkסԺ.Value = 0 Then
        MsgBox "��ѡ�������־��", vbInformation + vbOKOnly, gstrSysName
        Exit Sub
    End If
    
    If Trim(Me.txt���˿���.Tag) = "" Then
        MsgBox "��ѡ���˿��ң�", vbInformation + vbOKOnly, gstrSysName
        Exit Sub
    End If
        
    strBeginDate = Format(Dtp��ʼDate.Value, "yyyy-MM-dd hh:mm:ss")
    strEndDate = Format(Dtp����Date.Value, "yyyy-MM-dd hh:mm:ss")
        
    If mbln������ҩ������ Then
        strCon = strCon & " And C.��¼״̬=1 And C.�������� Between To_Date([3],'yyyy-MM-dd hh24:mi:ss') And To_Date([4] ,'yyyy-MM-dd hh24:mi:ss') "
    Else
        strCon = strCon & " And Mod(C.��¼״̬,3)=1 And C.�������� Between To_Date([3] ,'yyyy-MM-dd hh24:mi:ss') And To_Date([4] ,'yyyy-MM-dd hh24:mi:ss') "
    End If
    
    '������ҩ��֧��ˢ�����ѣ�������ȡ�ı��������շѻ�����˵Ĵ���
    gstrSQL = " Select /*+ Rule*/ Distinct S.���� As ҩ��, Decode(C.����,8,'�շ�',9,'����') ����,C.No,C.����,A.���շ�,Decode(A.��ҩ��,Null,'','���ŷ�ҩ','',A.��ҩ��) ��ҩ��,P.���� ����,B.����,B.��ʶ�� סԺ��,'' ����," & _
             " B.������ ����ҽ��,B.����Ա���� ������,To_Char(C.��������,'yyyy-MM-dd') ��������, S.ID As ҩ��id,B.��¼����,B.�����־,A.����ID, d.�������� " & _
             " From δ��ҩƷ��¼ A,������ü�¼ B,ҩƷ�շ���¼ C,���ű� P,���ű� S, ������Ϣ D " & IIf(Str���� = "", "", ",Table(Cast(f_Str2list([6]) As zlTools.t_Strlist)) E ") & IIf(mbln������, ",��������¼ Q,���������ϸ K ", "") & _
             " Where C.����ID=B.ID And B.��������ID+0=P.ID And Nvl(C.�ⷿID,0)+0=S.ID and Nvl(A.�ⷿID,0)=Nvl(C.�ⷿID,0)  And A.No=C.No " & IIf(mbln������, " and b.ҽ�����=k.ҽ��id(+) and Q.id(+)=K.��id and K.����ύ(+)=1 And (b.ҽ����� is null or nvl(q.�����,0) = 1)", "") & _
             IIf(Str���� = "", "", " And (C.��ҩ����=E.Column_Value Or C.��ҩ���� Is NULL)") & _
            IIf(IntSendAfterDosage = 0, " And C.��ҩ�� is not null And C.��ҩ���� is not null", "") & _
             " And C.����� Is Null  And Nvl(B.����״̬,0)<>1 " & _
             " and Not Exists(select 1 from ҩƷ�շ���¼ F where F.����=C.���� and F.�ⷿid=C.�ⷿid and F.no=C.no and ��ҩ��ʽ=-1) " & _
             " And C.����=A.���� and nvl(C.��ҩ��ʽ,-999)<>-1 And A.����id=D.����id(+) " & strCon
    
    If Me.chk�շ�.Value = 1 And Me.chk����.Value = 1 Then
        If mbln��ҩǰ�շѻ���� = True Then
            gstrSQL = gstrSQL & " And A.���� In(8,9) And A.���շ�=1 "
        ElseIf mint����δ�շѴ�����ҩ = False Then
            gstrSQL = gstrSQL & " And (C.����=8 And A.���շ�=1 Or C.����=9) "
        ElseIf Int����δ��˴�����ҩ = False Then
            gstrSQL = gstrSQL & " And (C.����=9 And A.���շ�=1 Or C.����=8) "
        Else
            gstrSQL = gstrSQL & " And A.���� In(8,9) "
        End If
    ElseIf Me.chk�շ�.Value = 1 Then
        If mbln��ҩǰ�շѻ���� = True Or mint����δ�շѴ�����ҩ = False Then
            gstrSQL = gstrSQL & " And A.����=8 And A.���շ�=1 "
        Else
            gstrSQL = gstrSQL & " And A.����=8 "
        End If
    ElseIf Me.chk����.Value = 1 Then
        If mbln��ҩǰ�շѻ���� = True Or mint����δ�շѴ�����ҩ = False Then
            gstrSQL = gstrSQL & " And A.����=9 And A.���շ�=1 "
        Else
            gstrSQL = gstrSQL & " And A.����=9 "
        End If
    End If
    
    If Me.chk����.Value <> 1 Or Me.chkסԺ.Value <> 1 Then
        If Me.chk����.Value = 1 Then
            gstrSQL = gstrSQL & " And (B.��¼����=1 or (b.��¼����=2 and (B.�����־=1 or B.�����־=4)))"
        Else
            gstrSQL = gstrSQL & " And (b.��¼����=2 and (B.�����־<>1 and B.�����־<>4))"
        End If
    End If
    
    If mstrDeptNode <> "" Then
        gstrSQL = gstrSQL & " And (P.վ�� = [7] Or P.վ�� Is Null) "
    End If
    
    If mbln������ҩ������ = True Then
        If chkҩ��.Value = 1 Then
            gstrSQL = gstrSQL & " And C.�ⷿID+0=[2] "
        Else
            gstrSQL = gstrSQL & " And C.�ⷿID+0<>[5] "
        End If
    Else
        gstrSQL = gstrSQL & " And (C.�ⷿID+0=[2] OR C.�ⷿID IS NULL)"
    End If
    
    If Me.lbl���˿���.Caption = "���˿���" Then
        gstrSQL = gstrSQL & " And B.���˿���id in (Select * From Table(Cast(f_Str2list([1]) As Zltools.t_Strlist))) "
    Else
        gstrSQL = gstrSQL & " And B.��������ID in (Select * From Table(Cast(f_Str2list([1]) As Zltools.t_Strlist))) "
    End If
    
    If Me.chk����.Value = 1 And Me.chkסԺ.Value = 1 Then
        strsql = Replace(gstrSQL, "'' ����", "B.����")
        strsql = Replace(strsql, "������ü�¼", "סԺ���ü�¼")
        strsql = Replace(strsql, "And Nvl(B.����״̬,0)<>1", "")
        gstrSQL = gstrSQL & " Union All " & strsql
    ElseIf Me.chkסԺ.Value = 1 Then
        gstrSQL = Replace(gstrSQL, "'' ����", "B.����")
        gstrSQL = Replace(gstrSQL, "������ü�¼", "סԺ���ü�¼")
        gstrSQL = Replace(gstrSQL, "And Nvl(B.����״̬,0)<>1", "")
    End If
    
    Set RecBill = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, Me.txt���˿���.Tag, Me.Cboҩ��.ItemData(Me.Cboҩ��.ListIndex), strBeginDate, strEndDate, mlngҩ��ID, Str����, mstrDeptNode)
    
    '��յ�ǰ�б�
    Set RecTotal = Nothing
    With Msf�����б�
        .Clear
        .rows = 2
        Call SetFormat(1)
    End With
    
    With Msf������ϸ
        .Clear
        .rows = 2
        Call SetFormat(2)
    End With
    
    With Msf��������
        .Clear
        .rows = 2
        Call SetFormat(3)
    End With
    
    If rs����������ϸ.RecordCount > 0 Then
        rs����������ϸ.MoveLast
        
        For i = 0 To rs����������ϸ.RecordCount - 1
            ''''ɾ����ǰ��
            rs����������ϸ.Delete adAffectCurrent
            ''''��ǰ�ƶ�ָ��
            rs����������ϸ.MovePrevious
        Next
        
        Me.stbThis.Panels(2).Text = "δ�����κδ���"
    End If
    
    If Not RecBill.EOF Then
        Call InitRec
        '��������Ϣ����ϸ���д���ڲ�ӳ���¼����
'        With rs���
'            If RecBill.RecordCount <> 0 Then
'                Do While Not RecBill.EOF
'                    .AddNew
'                    !���ݱ�ʶ = RecBill!NO & "|" & RecBill!����
'                    !��� = RecBill!���
'                    !��¼���� = RecBill!��¼����
'                    !�����־ = RecBill!�����־
'                    .Update
'                    RecBill.MoveNext
'                Loop
'            End If
'            RecBill.MoveFirst
'        End With
        
        If WriteSendListData(1) = True Then
            
            Me.stbThis.Panels(2).Text = "������" & RecBill.RecordCount & "�Ŵ���"
            
            Call Msf�����б�_EnterCell
            
            
            
'            '��λ���ղ�����Ĵ�����
'            Call SetLocateBill
            
            With Msf�����б�
                CmdOK.Enabled = (.RowData(IIf(.rows - 1 = 1, 1, .rows - 2)) <> 0)
            End With
            
            mblnModify = True
            Call RefreshData(Me.Cboҩ��.ItemData(Me.Cboҩ��.ListIndex))
        End If
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_Activate()
    If BlnStartUp = False Then
        Unload Me
        Exit Sub
    End If
    
    If mbln������ҩ������ = True Then
        chkҩ��.Visible = True
        lblҩ��.Visible = False
       
        chkҩ��.Value = Val(GetSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\" & "ҩƷ������ҩ", "����ҩҩ��ѡ��", "1"))
    
    Else
        chkҩ��.Visible = False
        lblҩ��.Visible = True
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Dim dateCurDate As Date
    
    BlnStartUp = False
    LngBillCount = 0
    
    strID = ""
    StrBillNo = ""
    
    dateCurDate = Sys.Currentdate()
    Me.Dtp��ʼDate.Value = Format(dateCurDate, "yyyy-MM-dd 00:00:00")
    Me.Dtp����Date.Value = Format(dateCurDate, "yyyy-MM-dd 23:59:59")
    
    If mbln������ҩ������ Then
        Dtp��ʼDate.MinDate = dateCurDate - 30
        Dtp����Date.MinDate = dateCurDate - 30
    End If
    
    mbln��ҩǰ�շѻ���� = (gtype_UserSysParms.P163_��Ŀִ��ǰ�������շѻ��ȼ������ = 1)
    mbln����ʱ����� = (Val(zldatabase.GetPara("ҩƷҽ��������ʱ�����", glngSys, 1341, 0)) = 1)
    mint�����ʾ = Val(zldatabase.GetPara("�����ʾ��ʽ", glngSys, 1341, 0))
    mbln������ = ((gtype_UserSysParms.P240_ҩ��������� = 1 Or gtype_UserSysParms.P240_ҩ��������� = 3) And gtype_UserSysParms.P241_�������ʱ�� = 2)
    
    
    Call SetFormat(1)
    Call SetFormat(2)
    Call SetFormat(3)
    
    Call InitRec
    
    Call Iniҩ��
   
    If Me.Cboҩ��.ListCount = 0 Then
        MsgBox "û����������ҩ����", vbInformation, gstrSysName
        Unload Me
    End If
     
    If mint������� = 3 Then
        Me.chk����.Value = 1
        Me.chkסԺ.Value = 1
    End If
    
    If mint������� <> 1 And mint������� <> 3 Then
        Me.chk����.Enabled = False
        Me.chkסԺ.Value = 1
    End If
    
    If mint������� <> 2 And mint������� <> 3 Then
        Me.chkסԺ.Enabled = False
        Me.chk����.Value = 1
    End If
    
    BlnStartUp = True
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    If Me.Width < 9495 Then Me.Width = 9495
    If Me.Height < 6705 Then Me.Height = 6705
    
    With CmdHelp
        .Top = Me.ScaleHeight - .Height - 100 - Me.stbThis.Height
    End With
    
    With CmdPrintSet
        .Top = CmdHelp.Top
        .Left = CmdHelp.Left + CmdHelp.Width + 100
    End With
    
    With CmdPrint
        .Top = CmdHelp.Top
        .Left = CmdPrintSet.Left + CmdPrintSet.Width + 100
    End With
    
    With CmdCancel
        .Top = CmdHelp.Top
        .Left = Me.ScaleWidth - .Width - 100
    End With
    
    With CmdOK
        .Top = CmdHelp.Top
        .Left = CmdCancel.Left - .Width - 100
    End With
    
    With Msf�����б�
        .Height = (CmdOK.Top - 200 - .Top) / 2
        .Width = Me.ScaleWidth - .Left - 50
    End With
    
    With TabShow
        .Top = Msf�����б�.Top + Msf�����б�.Height + 100
        .Height = CmdOK.Top - 100 - .Top
        .Width = Msf�����б�.Width
    End With
    
    With Msf��������
        .Left = 50
        .Height = TabShow.Height - .Top - 80
        .Width = TabShow.Width - .Left - 50
    End With
    
    With Msf������ϸ
        .Left = 50
        .Height = TabShow.Height - .Top - 80
        .Width = TabShow.Width - .Left - 50
    End With
    
    With fra������
        .Left = 100
        .Width = Me.ScaleWidth - .Left - 50
    End With
    
    With fra�ദ��
        .Left = 100
        .Width = Me.ScaleWidth - .Left - 50
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveFlexState(Msf��������, Me.Name)
    Call SaveFlexState(Msf�����б�, Me.Name)
    Call SaveFlexState(Msf������ϸ, Me.Name)
    
    If mbln������ҩ������ = True Then
        Call SaveSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\" & "ҩƷ������ҩ", "����ҩҩ��ѡ��", chkҩ��.Value)
    End If
End Sub

Private Sub Msf��������_EnterCell()
    Dim LngSelectRow As Long
    Dim intCol As Integer
    With Msf��������
        .Redraw = False
        
        
        LngSelectRow = .Row     '���浱ǰѡ����
        If LngTotalRow > 0 And LngTotalRow < .rows Then
            .Row = LngTotalRow       '����ϴ�ѡ����
            For intCol = 0 To .Cols - 1
                    .Col = intCol
                    .CellBackColor = &H80000005
                    .CellForeColor = &H80000008
            Next
            .Col = 0
        End If
        
        LngTotalRow = LngSelectRow
        .Row = LngTotalRow       '���õ�ǰѡ����
        For intCol = 0 To .Cols - 1
                .Col = intCol
                .CellBackColor = &H8000000D
                .CellForeColor = &H80000005
        Next
        .Col = 0
        .Redraw = True
    End With
End Sub

Private Sub Msf��������_GotFocus()
    With Msf��������
        .GridColorFixed = &H80000008
        .GridColor = &H80000008
    End With
End Sub

Private Sub Msf��������_LostFocus()
    With Msf��������
        .GridColorFixed = &H80000011
        .GridColor = &H80000011
    End With
End Sub

Private Sub Msf�����б�_EnterCell()
    Dim LngSelectRow As Long
    Dim intCol As Integer
    With Msf�����б�
        .Redraw = False


        LngSelectRow = .Row     '���浱ǰѡ����
        If LngListRow > 0 And LngListRow < .rows Then
            .Row = LngListRow       '����ϴ�ѡ����
            For intCol = 0 To .Cols - 1
                .Col = intCol
                .CellBackColor = &H80000005
                If intCol <> 4 Then
                    .CellForeColor = &H80000008
                End If
            Next
            .Col = 0
        End If

        LngListRow = LngSelectRow
'        If LngSelectRow = 0 Then
'            LngListRow = 1
'        End If
        .Row = LngListRow       '���õ�ǰѡ����
        For intCol = 0 To .Cols - 1
            .Col = intCol
            .CellBackColor = &H8000000D
            If intCol <> 4 Then
                .CellForeColor = &H80000005
            End If
        Next
        .Col = 0
        .Redraw = True
        
        If Trim(.TextMatrix(.Row, 2)) = "" Then
            With Msf������ϸ
                .Clear
                .rows = 2
                Call SetFormat(2)
            End With
            Exit Sub
        End If
        
        '��ʾ������ϸ
        Call ReadBillData(.RowData(.Row), .TextMatrix(.Row, 2), Val(.TextMatrix(.Row, 10)), Val(.TextMatrix(.Row, 11)), Val(.TextMatrix(.Row, 12)))
    End With
End Sub

Private Sub Msf�����б�_GotFocus()
    With Msf�����б�
        .GridColorFixed = &H80000008
        .GridColor = &H80000008
    End With
End Sub

Private Sub Msf�����б�_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim lng���� As Long, strNo As String
    Dim int��¼���� As Integer
    Dim int�����־ As Integer
    If KeyCode = vbKeyDelete Then
        If Msf�����б�.TextMatrix(Msf�����б�.Row, 2) = "" Then Exit Sub
        With rs����������ϸ
            If .RecordCount > 0 Then
                .MoveFirst
                Do While Not .EOF
                    .Find "���ݺ�='" & Msf�����б�.TextMatrix(Msf�����б�.Row, 2) & "'"
                    If Not .EOF Then .Delete
                    If Not .EOF Then .MoveNext
                Loop
            End If
        End With
        With rs������Դ����
            If .RecordCount > 0 Then
                .MoveFirst
                .Find "��Դ����='" & Msf�����б�.TextMatrix(Msf�����б�.Row, 3) & "'"
                If Not .EOF Then .Delete
            End If
        End With
        With Msf�����б�
            lng���� = Val(.RowData(.Row))
            strNo = .TextMatrix(.Row, 2)
            If .rows - 1 = 1 Then
                .TextMatrix(1, 0) = ""
                .TextMatrix(1, 1) = ""
                .TextMatrix(1, 2) = ""
                .TextMatrix(1, 3) = ""
                .TextMatrix(1, 4) = ""
                .TextMatrix(1, 5) = ""
                .TextMatrix(1, 6) = ""
                .TextMatrix(1, 7) = ""
                .TextMatrix(1, 8) = ""
                .TextMatrix(1, 9) = ""
                .TextMatrix(1, 10) = ""
                .TextMatrix(1, 11) = ""
                .TextMatrix(1, 12) = ""
                .RowData(1) = 0
            Else
                If Trim(.TextMatrix(.Row, 2)) <> "" Then .RemoveItem .Row: LngBillCount = LngBillCount - 1
            End If
            
            CmdOK.Enabled = (.RowData(IIf(.rows - 1 = 1, 1, .rows - 2)) <> 0)
            Me.stbThis.Panels(2).Text = IIf(LngBillCount = 0, "δ�����κδ���", "������" & LngBillCount & "�Ŵ���")
            'Call RefreshData
        
            'ɾ���õ���
            With rs���
                If .RecordCount <> 0 Then .MoveFirst
                .Find "���ݱ�ʶ='" & strNo & "|" & lng���� & "'"
                If Not .EOF Then .Delete
            End With
            
            If rs���.RecordCount = 0 Then InitRec
        End With
        
        Msf�����б�_EnterCell
        mblnModify = True
'        If tabShow.Tab = 1 Then Call RefreshData
        Call WriteTotalDataToBill
    End If
End Sub

Private Sub Msf�����б�_LostFocus()
    With Msf�����б�
        .GridColorFixed = &H80000011
        .GridColor = &H80000011
    End With
End Sub

Private Sub Msf������ϸ_EnterCell()
    Dim LngSelectRow As Long
    Dim intCol As Integer
    With Msf������ϸ
        .Redraw = False
        
        LngSelectRow = .Row     '���浱ǰѡ����
        If LngDetailRow > 0 And LngDetailRow < .rows Then
            .Row = LngDetailRow       '����ϴ�ѡ����
            For intCol = 0 To .Cols - 1
                    .Col = intCol
                    .CellBackColor = &H80000005
                    .CellForeColor = &H80000008
            Next
            .Col = 0
        End If
        
        LngDetailRow = LngSelectRow
        .Row = LngDetailRow       '���õ�ǰѡ����
        For intCol = 0 To .Cols - 1
                .Col = intCol
                .CellBackColor = &H8000000D
                .CellForeColor = &H80000005
        Next
        .Col = 0
        .Redraw = True
    End With
End Sub

Private Sub Msf������ϸ_GotFocus()
    With Msf������ϸ
        .GridColorFixed = &H80000008
        .GridColor = &H80000008
    End With
End Sub

Private Sub Msf������ϸ_LostFocus()
    With Msf������ϸ
        .GridColorFixed = &H80000011
        .GridColor = &H80000011
    End With
End Sub


Private Sub tabShow_Click(PreviousTab As Integer)
    Select Case TabShow.Tab
    Case 0
        Msf������ϸ.ZOrder
        Msf������ϸ_EnterCell
    Case 1
'        Call RefreshData
        WriteTotalDataToBill
        Msf��������.ZOrder
        Msf��������_EnterCell
    End Select
End Sub

Private Sub TxtNo_GotFocus()
    GetFocus TxtNo
End Sub

Private Sub TxtNo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
       
    '--���������λ,�򰴹������--
    Me.TxtNo = UCase(LTrim(Me.TxtNo))
    Me.TxtNo.Text = GetFullNO(Me.TxtNo.Text, 13)
    
    If mstrDeptNode = "" Or Cboҩ��.Tag <> mstrDeptNode Then
        mstrDeptNode = GetDeptStationNode(Val(Cboҩ��.ItemData(Cboҩ��.ListIndex)))
        Cboҩ��.Tag = mstrDeptNode
    End If
    
    Call GetRecipe(1, TxtNo)
End Sub

Private Function CheckSource(ByVal Int���� As Integer, ByVal strNo As String, ByVal lngҩ��ID As Long) As Boolean
    Dim rs As New ADODB.Recordset
    Dim n As Integer
    Dim bln�ظ����� As Boolean
    
    On Error GoTo errHandle
    gstrSQL = "Select B.���� as ����,B.���� as ��Դ���� From ҩƷ�շ���¼ A,���ű� B Where A.�Է�����id=B.id and No=[1] And ����=[2] " & _
          " And Mod(��¼״̬,3)=1 And ����� Is Null And (�ⷿID+0=[3] Or �ⷿID Is NULL) And Rownum<2"
    Set rs = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, strNo, Int����, lngҩ��ID)
    
    If rs.RecordCount = 0 Then
        CheckSource = False
        Exit Function
    End If
    
    With rs������Դ����
        If .RecordCount = 0 Then
            .AddNew
            !���� = rs!����
            !��Դ���� = rs!��Դ����
            CheckSource = True
        Else
            .MoveFirst
            For n = 1 To .RecordCount
                If !���� = rs!���� Then
                    bln�ظ����� = True
                    Exit For
                End If
                .MoveNext
            Next
            If Not bln�ظ����� Then
                If MsgBox("��ǰ�����Ŀ���������[" & rs!���� & "]" & rs!��Դ���� & "����ȷ��Ҫ����ô�����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    Exit Function
                Else
                    .AddNew
                    !���� = rs!����
                    !��Դ���� = rs!��Դ����
                    CheckSource = True
                End If
            Else
                CheckSource = True
            End If
        End If
    End With
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Private Function ReadBillData(ByVal BillStyle As Integer, ByVal BillNo As String, ByVal lngҩ��ID As Long, ByVal int��¼���� As Integer, ByVal int�����־ As Integer) As Boolean
    Dim IntStyle As Integer
    Dim str��� As String
    Dim str��ϸ��λ�� As String
    '--��ȡ��������--
    'BillStyle-��������;BIllNO-���ݺ�
    '��λ��ʾ���ݷ����������������ﵥλ��סԺ��סԺ���סԺ��λ���������ۼ۵�λ��
'    On Error Resume Next
'    err = 0
    On Error GoTo errHandle
    ReadBillData = False
    
    strUnit = GetUnit(lngҩ��ID, BillStyle, BillNo, int�����־)
    Select Case strUnit
    Case "�ۼ۵�λ"
        str��ϸ��λ�� = "C.���㵥λ ��λ,B.���ۼ� ����,B.ʵ������*Nvl(B.����,1) ����"
    Case "���ﵥλ"
        str��ϸ��λ�� = "D.���ﵥλ ��λ,B.���ۼ�*Decode(D.�����װ,Null,1,0,1,D.�����װ) ����,B.ʵ������/Decode(D.�����װ,Null,1,0,1,D.�����װ)*Nvl(B.����,1) ����"
    Case "סԺ��λ"
        str��ϸ��λ�� = "D.סԺ��λ ��λ,B.���ۼ�*Decode(D.סԺ��װ,Null,1,0,1,D.סԺ��װ) ����,B.ʵ������/Decode(D.סԺ��װ,Null,1,0,1,D.סԺ��װ)*Nvl(B.����,1) ����"
    Case "ҩ�ⵥλ"
        str��ϸ��λ�� = "D.ҩ�ⵥλ ��λ,B.���ۼ�*Decode(D.ҩ���װ,Null,1,0,1,D.ҩ���װ) ����,B.ʵ������/Decode(D.ҩ���װ,Null,1,0,1,D.ҩ���װ)*Nvl(B.����,1) ����"
    End Select
    str��ϸ��λ�� = str��ϸ��λ�� & ",B.���۽�� ���,Nvl(B.����, 1) * B.ʵ������ / (Nvl(F.����, 1) * F.����) * F.ʵ�ս�� As ʵ�ս�� "
    
    gstrSQL = " SELECT DISTINCT F.���,F.����ID,'['||C.����||']'|| " & IIf(gintҩƷ������ʾ = 1, "Nvl(A.����,C.����)", "C.����") & " As Ʒ��,A.���� AS ��Ʒ��, " & _
        " DECODE(C.���,NULL,B.����,DECODE(B.����,NULL,C.���,C.���||'|'||B.����)) ���," & _
        str��ϸ��λ�� & _
        " FROM ҩƷ�շ���¼ B,ҩƷ��� D,�շ���ĿĿ¼ C,�շ���Ŀ���� A,������ü�¼ F" & _
        " WHERE B.ҩƷID=D.ҩƷID AND D.ҩƷID=C.ID And B.����ID=F.ID" & _
        " AND D.ҩƷID=A.�շ�ϸĿID(+) AND A.����(+)=3 " & _
        " AND MOD(B.��¼״̬,3)=1 AND B.NO=[1] AND B.����=[2] " & _
        " AND (B.�ⷿID+0=[3] OR B.�ⷿID IS NULL) " & _
        " And ����� Is Null" & _
        " Order by F.���"
    If int��¼���� = 1 Or (int��¼���� = 2 And (int�����־ = 1 Or int�����־ = 4)) Then
    Else
        gstrSQL = Replace(gstrSQL, "������ü�¼", "סԺ���ü�¼")
    End If
    
    Set RecBill = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, BillNo, BillStyle, lngҩ��ID)
    
    With RecBill
        str��� = ""
        Do While Not .EOF
            str��� = str��� & "," & !���
            .MoveNext
        Loop
        If str��� <> "" Then str��� = Mid(str���, 2)
        .MoveFirst
    End With
    
    '��������Ϣ����ϸ���д���ڲ�ӳ���¼����
    With rs���
        If .RecordCount <> 0 Then .MoveFirst
        .Find "���ݱ�ʶ='" & BillNo & "|" & BillStyle & "'"
        If str��� <> "" Then
            If .EOF Then
                .AddNew
                !���ݱ�ʶ = BillNo & "|" & BillStyle
                !��� = str���
                !��¼���� = int��¼����
                !�����־ = int�����־
                .Update
            End If
        End If
    End With
    
    If WriteDataToBill() = False Then Exit Function

    If err <> 0 Then
        MsgBox "��ȡ����ʱ����������Ԥ֪�Ĵ���", vbInformation, gstrSysName
        Exit Function
    End If
    ReadBillData = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function CheckBill(ByVal intRow As Integer, ByVal IntBillStyle As Integer, ByVal strNo As String, ByVal lngҩ��ID As Long, ByVal int��¼���� As Integer, ByVal int�����־ As Integer) As Integer
    Dim RecCheck As New ADODB.Recordset

    '--���ݽ�Ҫִ�еĲ������ж��Ƿ�����--
    '����:
    '0-�������
    '1-δ��ҩ
    '2-����ҩ
    '3-�ѷ�ҩ
    '4-��ɾ��
    '5-δ��ҩ
    On Error GoTo errHandle
    gstrSQL = " Select A.��ҩ��,A.�����,nvl(B.���շ�,0) ���շ�, C.����Ա���� ������ " & _
            " From ҩƷ�շ���¼ A,δ��ҩƷ��¼ B, ������ü�¼ C " & _
            " Where A.No=B.No And A.����=B.���� And A.����id = C.ID And mod(A.��¼״̬,3)=1 And A.����� IS Null And Rownum=1 " & _
            " And A.No=[1] And A.����=[2] And (A.�ⷿID+0=[3] Or A.�ⷿID Is NULL)"
    If int��¼���� = 1 Or (int��¼���� = 2 And (int�����־ = 1 Or int�����־ = 4)) Then
    Else
        gstrSQL = Replace(gstrSQL, "������ü�¼", "סԺ���ü�¼")
    End If
    
    Set RecCheck = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, strNo, IntBillStyle, lngҩ��ID)
        
    With RecCheck
        If .EOF Then CheckBill = 4: MsgBox "δ�ҵ�����[" & strNo & "],�����Ѿ�����������Ա����,����������ֹ��", vbInformation, gstrSysName: Exit Function
        If Not IsNull(!�����) Then
            CheckBill = 3: MsgBox "�ô���[" & strNo & "]�ѱ���������Ա��ҩ����ҩ������ֹ��", vbInformation, gstrSysName: Exit Function
        End If
        
        '�������շѱ�־
        Msf�����б�.TextMatrix(intRow, 13) = !���շ�
    End With

    CheckBill = 0
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function WriteSendListData(ByVal intType As Integer) As Boolean
    'intType��0-¼�뵥������ʽ��1-������ȡ��ʽ
    Dim RecCheck As New ADODB.Recordset
    Dim i As Integer
    Dim blnContinue As Boolean
    
    WriteSendListData = False
    
    mstr���ܵ��� = ""
    Do While Not (RecBill.EOF)
        blnContinue = True
        
        '��������ʽ�������������ľ���ʾ�ͽ�ֹ��������ʽʱ����ȡ��Ӧ�Ĵ���Ҳ����ʾ
        If mbln������ҩ������ = False And IntSendAfterDosage = 0 Then
            If IsNull(RecBill!��ҩ��) Then
                If intType = 0 Then
                    MsgBox "�ô�����δ��ҩ������ִ�з�ҩ������", vbInformation, gstrSysName
                    Exit Function
                Else
                    blnContinue = False
                End If
            End If
            If Trim(RecBill!��ҩ��) = "" Then
                If intType = 0 Then
                    MsgBox "�ô�����δ��ҩ������ִ�з�ҩ������", vbInformation, gstrSysName
                    Exit Function
                Else
                    blnContinue = False
                End If
            End If
        End If
        
        If blnContinue = True Then
            With Msf�����б�
                .Redraw = False
                .TextMatrix(.rows - 1, 0) = RecBill!ҩ��
                .TextMatrix(.rows - 1, 1) = RecBill!����
                .TextMatrix(.rows - 1, 2) = RecBill!NO
                .TextMatrix(.rows - 1, 3) = IIf(IsNull(RecBill!����), "", RecBill!����)
                .TextMatrix(.rows - 1, 4) = IIf(IsNull(RecBill!����), "", RecBill!����)
                .TextMatrix(.rows - 1, 5) = IIf(IsNull(RecBill!סԺ��), "", RecBill!סԺ��)
                .TextMatrix(.rows - 1, 6) = IIf(IsNull(RecBill!����), "", RecBill!����)
                .TextMatrix(.rows - 1, 7) = IIf(IsNull(RecBill!������), "", RecBill!������)
                .TextMatrix(.rows - 1, 8) = IIf(IsNull(RecBill!����ҽ��), "", RecBill!����ҽ��)
                .TextMatrix(.rows - 1, 9) = IIf(IsNull(RecBill!��������), "", RecBill!��������)
                .TextMatrix(.rows - 1, 10) = RecBill!ҩ��ID
                .TextMatrix(.rows - 1, 11) = RecBill!��¼����
                .TextMatrix(.rows - 1, 12) = RecBill!�����־
                .TextMatrix(.rows - 1, 13) = RecBill!���շ�
                .TextMatrix(.rows - 1, 14) = IIf(IsNull(RecBill!����ID), "", RecBill!����ID)
                .RowData(.rows - 1) = RecBill!����
    '            str���ݺ� = RecBill!NO
                If chk����(0).Value = 0 And chk����(1).Value = 0 Then
                    mstr���ܵ��� = RecBill!���� & "," & RecBill!NO
                Else
                    mstr���ܵ��� = IIf(mstr���ܵ��� = "", "", mstr���ܵ��� & "|") & RecBill!���� & "," & RecBill!NO
                End If
                
                .Row = .rows - 1
                .Col = 4
                .CellForeColor = zldatabase.GetPatiColor(IIf(IsNull(RecBill!��������), "", RecBill!��������))
    
                .rows = .rows + 1
                .RowData(.rows - 1) = 0
                .Redraw = True
            End With
            WriteSendListData = True
        End If
        
        RecBill.MoveNext
    Loop
    
End Function

Private Function RefreshData(ByVal lngҩ��ID As Long) As Boolean
    Dim intRow As Integer, intRows As Integer
    Dim arrID
    Dim StrNoThis As String, IntBillThis As Integer
    Dim str���ܵ�λ�� As String
    Dim strTemp As String
    
    If mblnModify = False Then Exit Function
    RefreshData = False
    On Error GoTo errHandle
    '��ջ��ܱ��
    With Msf��������
        .Clear
        .rows = 2
        SetFormat (3)
    End With
 
   
    '��ʾ��������
    Dim intUnit As Integer
    intUnit = Val(zldatabase.GetPara("ҩ������", glngSys, 1341, 0))
    If intUnit = 0 Then
        strUnit = GetDrugUnit(lngҩ��ID, "", True)
    ElseIf intUnit = 1 Then
        strUnit = GetSpecUnit(lngҩ��ID, gint����ҩ��)
    Else
        strUnit = GetSpecUnit(lngҩ��ID, gintסԺҩ��)
    End If
    Select Case strUnit
    Case "�ۼ۵�λ"
        str���ܵ�λ�� = "C.���㵥λ ��λ,B.���ۼ� ����,Sum(B.ʵ������*Nvl(B.����,1)) ����"
    Case "���ﵥλ"
        str���ܵ�λ�� = "D.���ﵥλ ��λ,B.���ۼ�*Decode(D.�����װ,Null,1,0,1,D.�����װ) ����,Sum(B.ʵ������/Decode(D.�����װ,Null,1,0,1,D.�����װ)*Nvl(B.����,1)) ����"
    Case "סԺ��λ"
        str���ܵ�λ�� = "D.סԺ��λ ��λ,B.���ۼ�*Decode(D.סԺ��װ,Null,1,0,1,D.סԺ��װ) ����,Sum(B.ʵ������/Decode(D.סԺ��װ,Null,1,0,1,D.סԺ��װ)*Nvl(B.����,1)) ����"
    Case "ҩ�ⵥλ"
        str���ܵ�λ�� = "D.ҩ�ⵥλ ��λ,B.���ۼ�*Decode(D.ҩ���װ,Null,1,0,1,D.ҩ���װ) ����,Sum(B.ʵ������/Decode(D.ҩ���װ,Null,1,0,1,D.ҩ���װ)*Nvl(B.����,1)) ����"
    End Select
    
    str���ܵ�λ�� = str���ܵ�λ�� & ",Sum(B.���۽��) ���,Sum(Nvl(B.����, 1) * B.ʵ������ / (Nvl(B.���ø���, 1) * B.����) * B.ʵ�ս��) As ʵ�ս�� "

    gstrSQL = "Select A.No,A.ҩƷid,A.����,A.���ۼ�,A.ʵ������,A.����,A.���۽��,C.���� As ���ø���,C.����,C.ʵ�ս��,A.���� From ҩƷ�շ���¼ A,������ü�¼ C,Table(f_Str2list2([1], '|', ',')) B " & _
        " Where A.NO=B.C2 And A.����=B.C1 And Mod(A.��¼״̬,3)=1 And A.����� Is Null And (A.�ⷿID+0=[2] Or A.�ⷿID Is NULL) And A.����id = C.Id "
    gstrSQL = gstrSQL & " Union All " & Replace(gstrSQL, "������ü�¼", "סԺ���ü�¼")
    gstrSQL = "Select Distinct D.*,'['||D.����||']'|| " & IIf(gintҩƷ������ʾ = 1, "Nvl(A.����,D.ͨ������)", "D.ͨ������") & " As Ʒ��,A.���� AS ��Ʒ�� " & _
             " From " & _
             "     (SELECT B.No,D.ҩƷID,C.����,C.���� ͨ������,NVL(B.����,0) ����," & _
             "     DECODE(C.���,NULL,B.����,DECODE(B.����,NULL,C.���,C.���||'|'||B.����)) ���," & str���ܵ�λ�� & _
             "     FROM (" & gstrSQL & ") B," & _
             "           ҩƷ��� D,�շ���ĿĿ¼ C " & _
             "     WHERE B.ҩƷID+0=D.ҩƷID AND D.ҩƷID=C.ID" & _
             "     GROUP BY B.No,D.ҩƷID,C.����,C.����,NVL(B.����,0)," & _
             "     DECODE(C.���,NULL,B.����,DECODE(B.����,NULL,C.���,C.���||'|'||B.����)),"
    Select Case strUnit
    Case "�ۼ۵�λ"
        gstrSQL = gstrSQL & "C.���㵥λ,B.���ۼ�"
    Case "���ﵥλ"
        gstrSQL = gstrSQL & "D.���ﵥλ,B.���ۼ�*Decode(D.�����װ,Null,1,0,1,D.�����װ)"
    Case "סԺ��λ"
        gstrSQL = gstrSQL & "D.סԺ��λ,B.���ۼ�*Decode(D.סԺ��װ,Null,1,0,1,D.סԺ��װ)"
    Case "ҩ�ⵥλ"
        gstrSQL = gstrSQL & "D.ҩ�ⵥλ,B.���ۼ�*Decode(D.ҩ���װ,Null,1,0,1,D.ҩ���װ)"
    End Select
    gstrSQL = gstrSQL & ") D,�շ���Ŀ���� A" & _
            " Where D.ҩƷID=A.�շ�ϸĿID(+) AND A.����(+)=3"
    gstrSQL = gstrSQL & " Order By D.����"
    
    
    If Len(mstr���ܵ���) > 4000 Then
        For intRow = 0 To UBound(GetArrayByStr(mstr���ܵ���, 3900, "|"))
            
            Set RecTotal = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, CStr(GetArrayByStr(mstr���ܵ���, 3900, "|")(intRow)), lngҩ��ID)
            Call WriteTotalDataToBill(intRow > 0, Not (intRow = UBound(GetArrayByStr(mstr���ܵ���, 3900, "|"))))
        Next
    Else
        Set RecTotal = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, mstr���ܵ���, lngҩ��ID)
        Call WriteTotalDataToBill
    End If
    
    If err <> 0 Then
        MsgBox "��ʾ��������ʱ����������Ԥ֪�Ĵ���", vbInformation, gstrSysName
        Exit Function
    End If
    
    mblnModify = False
    RefreshData = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function WriteTotalDataToBill(Optional ByVal blnFirst As Boolean, Optional ByVal blnLast As Boolean) As Boolean
    Dim dblӦ�ս�� As Double
    Dim dblʵ�ս�� As Double
    Dim str�����ʾ As String
    
    On Error GoTo errHandle
    
    '����������װ��
    If Not blnFirst Then
        With Msf��������
            .Redraw = False
            .Clear
            .rows = 2
            Call SetFormat(3)
            .Redraw = True
        End With
    End If
    
    '��䵥������
    
    If RecTotal.State = 0 Then Exit Function
    
    If RecTotal.RecordCount > 0 Then
'        If Not RecTotal.EOF Then Call InitRecSum
    
        Do While Not RecTotal.EOF
            With rs����������ϸ
                .AddNew
                !���ݺ� = RecTotal!NO
                !ҩƷ���� = RecTotal!Ʒ��
                !��Ʒ�� = IIf(IsNull(RecTotal!��Ʒ��), "", RecTotal!��Ʒ��)
                !���� = RecTotal!����
                !��� = IIf(IsNull(RecTotal!���), "", RecTotal!���)
                !��λ = IIf(IsNull(RecTotal!��λ), "", RecTotal!��λ)
                !���� = RecTotal!����
                !���� = RecTotal!����
                !��� = RecTotal!���
                !ʵ�ս�� = RecTotal!ʵ�ս��
                !ҩƷid = RecTotal!ҩƷid
                !���� = RecTotal!����
            End With
            RecTotal.MoveNext
        Loop
    End If
    
    If blnLast Then Exit Function
    
    With rs����������ϸ
        If .RecordCount <> 0 Then
            .Sort = "����,����"
            .MoveFirst
        End If
        Do While Not .EOF
            If Msf��������.rows = 2 And Msf��������.TextMatrix(1, 1) = "" Then
                Msf��������.TextMatrix(Msf��������.rows - 1, 0) = Msf��������.rows - 1
                Msf��������.TextMatrix(Msf��������.rows - 1, 1) = !ҩƷ����
                Msf��������.TextMatrix(Msf��������.rows - 1, 2) = !��Ʒ��
                Msf��������.TextMatrix(Msf��������.rows - 1, 3) = IIf(IsNull(!���), "", !���)
                Msf��������.TextMatrix(Msf��������.rows - 1, 4) = IIf(IsNull(!��λ), "", !��λ)
                Msf��������.TextMatrix(Msf��������.rows - 1, 5) = Format(!����, "#####0.00000;-#####0.00000; ;")
                Msf��������.TextMatrix(Msf��������.rows - 1, 6) = Format(!����, "#####0.00000;-#####0.00000; ;")
                Msf��������.TextMatrix(Msf��������.rows - 1, 7) = Format(!���, "#####0.00;-#####0.00; ;")
                Msf��������.TextMatrix(Msf��������.rows - 1, 8) = Format(!ʵ�ս��, "#####0.00;-#####0.00; ;")
                Msf��������.TextMatrix(Msf��������.rows - 1, 9) = !ҩƷid
                Msf��������.TextMatrix(Msf��������.rows - 1, 10) = !����
                Msf��������.MergeRow(Msf��������.rows - 1) = False
            ElseIf Msf��������.TextMatrix(Msf��������.rows - 1, 9) <> !ҩƷid Then
                Msf��������.rows = Msf��������.rows + 1
                Msf��������.TextMatrix(Msf��������.rows - 1, 0) = Msf��������.rows - 1
                Msf��������.TextMatrix(Msf��������.rows - 1, 1) = !ҩƷ����
                Msf��������.TextMatrix(Msf��������.rows - 1, 2) = !��Ʒ��
                Msf��������.TextMatrix(Msf��������.rows - 1, 3) = IIf(IsNull(!���), "", !���)
                Msf��������.TextMatrix(Msf��������.rows - 1, 4) = IIf(IsNull(!��λ), "", !��λ)
                Msf��������.TextMatrix(Msf��������.rows - 1, 5) = Format(!����, "#####0.00000;-#####0.00000; ;")
                Msf��������.TextMatrix(Msf��������.rows - 1, 6) = Format(!����, "#####0.00000;-#####0.00000; ;")
                Msf��������.TextMatrix(Msf��������.rows - 1, 7) = Format(!���, "#####0.00;-#####0.00; ;")
                Msf��������.TextMatrix(Msf��������.rows - 1, 8) = Format(!ʵ�ս��, "#####0.00;-#####0.00; ;")
                Msf��������.TextMatrix(Msf��������.rows - 1, 9) = !ҩƷid
                Msf��������.TextMatrix(Msf��������.rows - 1, 10) = !����
                Msf��������.MergeRow(Msf��������.rows - 1) = False
            ElseIf Msf��������.TextMatrix(Msf��������.rows - 1, 10) <> !���� Then
                Msf��������.rows = Msf��������.rows + 1
                Msf��������.TextMatrix(Msf��������.rows - 1, 0) = Msf��������.rows - 1
                Msf��������.TextMatrix(Msf��������.rows - 1, 1) = !ҩƷ����
                Msf��������.TextMatrix(Msf��������.rows - 1, 2) = !��Ʒ��
                Msf��������.TextMatrix(Msf��������.rows - 1, 3) = IIf(IsNull(!���), "", !���)
                Msf��������.TextMatrix(Msf��������.rows - 1, 4) = IIf(IsNull(!��λ), "", !��λ)
                Msf��������.TextMatrix(Msf��������.rows - 1, 5) = Format(!����, "#####0.00000;-#####0.00000; ;")
                Msf��������.TextMatrix(Msf��������.rows - 1, 6) = Format(!����, "#####0.00000;-#####0.00000; ;")
                Msf��������.TextMatrix(Msf��������.rows - 1, 7) = Format(!���, "#####0.00;-#####0.00; ;")
                Msf��������.TextMatrix(Msf��������.rows - 1, 8) = Format(!ʵ�ս��, "#####0.00;-#####0.00; ;")
                Msf��������.TextMatrix(Msf��������.rows - 1, 9) = !ҩƷid
                Msf��������.TextMatrix(Msf��������.rows - 1, 10) = !����
                Msf��������.MergeRow(Msf��������.rows - 1) = False
            Else
                Msf��������.TextMatrix(Msf��������.rows - 1, 6) = Format(CDbl(Val(Msf��������.TextMatrix(Msf��������.rows - 1, 6))) + !����, "#####0.00000;-#####0.00000; ;")
                Msf��������.TextMatrix(Msf��������.rows - 1, 7) = Format(CDbl(Val(Msf��������.TextMatrix(Msf��������.rows - 1, 7))) + !���, "#####0.00000;-#####0.00000; ;")
                Msf��������.TextMatrix(Msf��������.rows - 1, 8) = Format(CDbl(Val(Msf��������.TextMatrix(Msf��������.rows - 1, 8))) + !ʵ�ս��, "#####0.00000;-#####0.00000; ;")
            End If
            dblӦ�ս�� = dblӦ�ս�� + !���
            dblʵ�ս�� = dblʵ�ս�� + !ʵ�ս��
            .MoveNext
        Loop
        
        '��ʾ�ϼ�
        Msf��������.rows = Msf��������.rows + 1
        Msf��������.TextMatrix(Msf��������.rows - 1, 0) = "�ϼ�"
        Msf��������.TextMatrix(Msf��������.rows - 1, 1) = "�ϼ�"
        Msf��������.TextMatrix(Msf��������.rows - 1, 2) = "�ϼ�"
        Msf��������.TextMatrix(Msf��������.rows - 1, 3) = "�ϼ�"
        Msf��������.TextMatrix(Msf��������.rows - 1, 4) = "�ϼ�"
        
        If mint�����ʾ = 1 Then
            str�����ʾ = "ʵ�ս�" & Format(dblʵ�ս��, "#####0.00;-#####0.00; ;")
        ElseIf mint�����ʾ = 2 Then
            str�����ʾ = "Ӧ�ս�" & Format(dblӦ�ս��, "#####0.00;-#####0.00; ;") & "    ʵ�ս�" & Format(dblʵ�ս��, "#####0.00;-#####0.00; ;")
        Else
            str�����ʾ = "Ӧ�ս�" & Format(dblӦ�ս��, "#####0.00;-#####0.00; ;")
        End If
        
        Msf��������.TextMatrix(Msf��������.rows - 1, 5) = str�����ʾ
        Msf��������.TextMatrix(Msf��������.rows - 1, 6) = str�����ʾ
        Msf��������.TextMatrix(Msf��������.rows - 1, 7) = str�����ʾ
        Msf��������.TextMatrix(Msf��������.rows - 1, 8) = str�����ʾ
        
        Msf��������.MergeCells = flexMergeFree
        Msf��������.MergeRow(Msf��������.rows - 1) = True
    End With

    WriteTotalDataToBill = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function WriteDataToBill() As Boolean
    Dim dblӦ�ս�� As Double
    Dim dblʵ�ս�� As Double
    Dim str�����ʾ As String
    
    '--��ʾָ����������ϸ--
    On Error Resume Next
    err = 0
    
    WriteDataToBill = False
    With Msf������ϸ
        .Clear
        .rows = 2
        Call SetFormat(2)
    End With
    
    '��䵥������
    With RecBill
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            Msf������ϸ.MergeRow(.AbsolutePosition) = False
            Msf������ϸ.TextMatrix(.AbsolutePosition, 0) = !Ʒ��
            Msf������ϸ.TextMatrix(.AbsolutePosition, 1) = IIf(IsNull(!��Ʒ��), "", !��Ʒ��)
            Msf������ϸ.TextMatrix(.AbsolutePosition, 2) = IIf(IsNull(!���), "", !���)
            Msf������ϸ.TextMatrix(.AbsolutePosition, 3) = IIf(IsNull(!��λ), "", !��λ)
            Msf������ϸ.TextMatrix(.AbsolutePosition, 4) = Format(!����, "#####0.00000;-#####0.00000; ;")
            Msf������ϸ.TextMatrix(.AbsolutePosition, 5) = Format(!����, "#####0.00000;-#####0.00000; ;")
            Msf������ϸ.TextMatrix(.AbsolutePosition, 6) = Format(!���, "#####0.00;-#####0.00; ;")
            Msf������ϸ.TextMatrix(.AbsolutePosition, 7) = Format(!ʵ�ս��, "#####0.00;-#####0.00; ;")
            dblӦ�ս�� = dblӦ�ս�� + Val(!���)
            dblʵ�ս�� = dblʵ�ս�� + Val(!ʵ�ս��)
            
            If .AbsolutePosition >= Msf������ϸ.rows - 1 Then Msf������ϸ.rows = Msf������ϸ.rows + 1
            .MoveNext
        Loop
    End With
    With Msf������ϸ
        .TextMatrix(.rows - 1, 0) = "�ϼ�"
        .TextMatrix(.rows - 1, 1) = "�ϼ�"
        .TextMatrix(.rows - 1, 2) = "�ϼ�"
        .TextMatrix(.rows - 1, 3) = "�ϼ�"
        
        If mint�����ʾ = 1 Then
            str�����ʾ = "ʵ�ս�" & Format(dblʵ�ս��, "#####0.00;-#####0.00; ;")
        ElseIf mint�����ʾ = 2 Then
            str�����ʾ = "Ӧ�ս�" & Format(dblӦ�ս��, "#####0.00;-#####0.00; ;") & "    ʵ�ս�" & Format(dblʵ�ս��, "#####0.00;-#####0.00; ;")
        Else
            str�����ʾ = "Ӧ�ս�" & Format(dblӦ�ս��, "#####0.00;-#####0.00; ;")
        End If
        
        .TextMatrix(.rows - 1, 4) = str�����ʾ
        .TextMatrix(.rows - 1, 5) = str�����ʾ
        .TextMatrix(.rows - 1, 6) = str�����ʾ
        .TextMatrix(.rows - 1, 7) = str�����ʾ
        .MergeCells = flexMergeFree
        .MergeRow(.rows - 1) = True
    End With
    
    If err <> 0 Then
        MsgBox "��ʾ����ʱ����������Ԥ֪�Ĵ���", vbInformation, gstrSysName
        Exit Function
    End If
    WriteDataToBill = True
End Function

Private Function SetLocateBill(Optional ByVal BlnEnterCell As Boolean = True) As Boolean
    Dim intRow As Integer
    
    SetLocateBill = False
    With Msf�����б�
        For intRow = 1 To .rows - 1
            If .TextMatrix(intRow, 2) = TxtNo And TxtNo.Tag = .RowData(intRow) Then
                .Row = intRow
                .TopRow = intRow
                SetLocateBill = True
                Exit For
            End If
        Next
    End With
    
    If BlnEnterCell Then Msf�����б�_EnterCell
End Function

Private Function CheckStock() As Boolean
    Dim RecCheckStock As New ADODB.Recordset
    Dim dblStock As Double
    Dim strSubSql As String
    Dim n As Integer
    Dim lngTemp�ⷿid As Long
    Dim dblUsableStock As Double
    
    On Error GoTo errHandle
    If mbln������ҩ������ = False Then
        lngTemp�ⷿid = Cboҩ��.ItemData(Cboҩ��.ListIndex)
    Else
        lngTemp�ⷿid = mlngҩ��ID
    End If
    
    IntCheckStock = MediWork_GetCheckStockRule(lngTemp�ⷿid)
    
    '�����
    If IntCheckStock = 0 Then CheckStock = True: Exit Function
    
    '���������ת��Ϊ��Ӧ��λ��ʵ������
    Dim intUnit As Integer
    intUnit = Val(zldatabase.GetPara("ҩ������", glngSys, 1341, 0))
    If intUnit = 0 Then
        strUnit = GetDrugUnit(lngTemp�ⷿid, "", True)
    ElseIf intUnit = 1 Then
        strUnit = GetSpecUnit(lngTemp�ⷿid, gint����ҩ��)
    Else
        strUnit = GetSpecUnit(lngTemp�ⷿid, gintסԺҩ��)
    End If
    Select Case strUnit
    Case "�ۼ۵�λ"
        strSubSql = "/1"
    Case "���ﵥλ"
        strSubSql = "/Decode(B.�����װ,Null,1,0,1,B.�����װ)"
    Case "סԺ��λ"
        strSubSql = "/Decode(B.סԺ��װ,Null,1,0,1,B.סԺ��װ)"
    Case "ҩ�ⵥλ"
        strSubSql = "/Decode(B.ҩ���װ,Null,1,0,1,B.ҩ���װ)"
    End Select
    
    CheckStock = False
    If Msf�����б�.TextMatrix(1, 2) <> "" Then
        For n = 1 To Msf��������.rows - 2
            gstrSQL = " Select nvl(��������,0)" & strSubSql & " AS ��������, nvl(ʵ������,0)" & strSubSql & " AS ʵ������ " & _
                         " From ҩƷ��� A,ҩƷ��� B" & _
                         " Where B.ҩƷID=A.ҩƷID And A.����=1 And A.�ⷿID=[1] And A.ҩƷID=[2] And Nvl(A.����,0)=[3]"
            Set RecCheckStock = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngTemp�ⷿid, Val(Msf��������.TextMatrix(n, 9)), Val(Msf��������.TextMatrix(n, 10)))
                         
            With RecCheckStock
                If .EOF Then
                    dblStock = 0
                    dblUsableStock = 0
                Else
                    dblStock = !ʵ������
                    dblUsableStock = !��������
                End If
                
                '����Ǵ�������ҩ�����������Ҫ���ʵ��������ҲҪ����������
                If dblStock < Val(Msf��������.TextMatrix(n, 6)) Or (mbln������ҩ������ = True And dblUsableStock < Val(Msf��������.TextMatrix(n, 6))) Then
                    If Msf��������.TextMatrix(n, 10) <> 0 Then
                        Select Case IntCheckStock
                        Case 1
                            If MsgBox(Msf��������.TextMatrix(n, 1) & "�����ο�����������Ƿ������ҩ��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
                        Case 2
                            MsgBox Msf��������.TextMatrix(n, 1) & "�����ο�������������ܼ�����ҩ��", vbInformation, gstrSysName: Exit Function
                        End Select
                    Else
                        Select Case IntCheckStock
                        Case 1
                            If MsgBox(Msf��������.TextMatrix(n, 1) & "�Ŀ�����������Ƿ������ҩ��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
                        Case 2
                            MsgBox Msf��������.TextMatrix(n, 1) & "�Ŀ�������������ܼ�����ҩ��", vbInformation, gstrSysName: Exit Function
                        End Select
                    End If
                End If
            End With
        Next
    End If
    
    CheckStock = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

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

Private Function CheckNotAudited(ByRef rsData As ADODB.Recordset) As Boolean
    Dim rsTmp As New ADODB.Recordset
    Dim bln�������� As Boolean
    Dim bln������ As Boolean
    Dim str��ϸ��λ�� As String
    
    On Error GoTo errHandle
    
    Call InitApplyforcredit
    
    CheckNotAudited = True
    bln�������� = True
    
    '��⵱ǰҩ���Ƿ�ΪסԺҩ�����������˳�������
    gstrSQL = "Select *" & vbNewLine & _
            "From ���ű� A, ��������˵�� B" & vbNewLine & _
            "Where a.Id = b.����id And a.Id = [1] And (b.�������� Like '%ҩ��' Or b.�������� Like '%ҩ��') And b.������� In (2, 3)"

    Set rsTmp = zldatabase.OpenSQLRecord(gstrSQL, "��⵱ǰҩ���Ƿ�ΪסԺҩ��", Cboҩ��.ItemData(Cboҩ��.ListIndex))
    If rsTmp.EOF Then Exit Function
    
    Select Case strUnit
    Case "���ﵥλ"
        str��ϸ��λ�� = "e.�����װ as ��װ,e.���ﵥλ as ��λ"
    Case "סԺ��λ"
        str��ϸ��λ�� = "e.סԺ��װ as ��װ,e.סԺ��λ as ��λ"
    Case "ҩ�ⵥλ"
        str��ϸ��λ�� = "e.ҩ���װ as ��װ,e.ҩ�ⵥλ as ��λ"
    Case Else
        str��ϸ��λ�� = "e.�����װ as ��װ,e.���ﵥλ as ��λ"
    End Select
    
    gstrSQL = " select A.*,b.����,b.�Ա�,b.����,c.���� As ������������,decode(d.����,null,f.����,d.����) as ҩ��,e.�����װ as ��װ,e.���ﵥλ as ��λ" & vbNewLine & _
            "  from (select distinct id as �շ�id,����id, ҩƷid, ����, ʵ������ from ҩƷ�շ���¼ where No =  [1] and mod(��¼״̬, 3) = 1 and ������� is null) A," & vbNewLine & _
            "סԺ���ü�¼ B,���˷������� C,ҩƷ���� D, ҩƷ��� E,�շ���ĿĿ¼ F" & vbNewLine & _
            " where a.����id = b.Id And b.Id = c.����id and a.ҩƷid = d.ҩƷid(+) and e.ҩƷid = a.ҩƷid and e.ҩƷid = f.id"

    
    With rsData
        rsData.Sort = "����,NO"
    
        Do While Not .EOF
            Set rsTmp = zldatabase.OpenSQLRecord(gstrSQL, "����Ƿ������������δ��˵ĵ���", rsData!NO)

            If rsTmp.RecordCount > 0 Then
                bln�������� = False

                With mrsApplyforcredit
                    Do While Not rsTmp.EOF
                        .AddNew
                        
                        !��־ = 1
                        !NO = rsData!NO
                        !ҩƷ���� = rsTmp!ҩ��
                        !���� = rsTmp!����
                        !���� = Format(rsTmp!ʵ������ / rsTmp!��װ, "#####0.0000;-#####0.0000; ;") & rsTmp!��λ
                        !������������ = Format(rsTmp!������������ / rsTmp!��װ, "#####0.0000;-#####0.0000; ;") & rsTmp!��λ
                        !���� = rsTmp!����
                        !�Ա� = rsTmp!�Ա�
                        !���� = rsTmp!����
                        !����ID = rsTmp!����ID
                        !�շ�ID = rsTmp!�շ�ID
                        
                        rsTmp.MoveNext
                    Loop
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
            Do While Not mrsApplyforcredit.EOF
                rsData.Filter = "No = '" & mrsApplyforcredit!NO & "'"
                If rsData.RecordCount > 0 Then
                    rsData.Delete
                    rsData.Update
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

Private Function SendBill() As Boolean
    Dim intRow As Integer
    Dim StrDate As String
    Dim rsSendRecipeByNo As ADODB.Recordset
    Dim int���� As Integer
    Dim arrSql As Variant
    Dim i As Long
    Dim blnInTrans As Boolean
    Dim strǩ����¼ As String
    Dim strReturn As String
    Dim strNo As String
    Dim cur���ܷ�ҩ�� As Currency
    Dim strReserve As String
    Dim lngҩ��ID As Long
    
    On Error GoTo ErrHand
    
    arrSql = Array()

    SendBill = False
    
    Set rsSendRecipeByNo = New ADODB.Recordset
    With rsSendRecipeByNo
        If .State = 1 Then .Close
        .Fields.Append "NO", adLongVarChar, 8, adFldIsNullable
        .Fields.Append "����", adDouble, 18, adFldIsNullable
        .Fields.Append "ҩ��ID", adDouble, 18, adFldIsNullable
        .Fields.Append "��¼����", adDouble, 18, adFldIsNullable
        .Fields.Append "�����־", adDouble, 18, adFldIsNullable
        .Fields.Append "���շ�", adDouble, 18, adFldIsNullable
        .Fields.Append "����ID", adDouble, 18, adFldIsNullable
        .Fields.Append "��������", adDate, , adFldIsNullable
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
    
    StrDate = Format(Sys.Currentdate, "yyyy-MM-dd HH:mm:ss")
    cur���ܷ�ҩ�� = Val(zldatabase.GetNextNo(20))
    
    With Msf�����б�
        For intRow = 1 To .rows - 1
            If .RowData(intRow) <> 0 Then
                '��鴦������ֹ����
                If CheckBill(intRow, .RowData(intRow), .TextMatrix(intRow, 2), Val(.TextMatrix(intRow, 10)), Val(.TextMatrix(intRow, 11)), Val(.TextMatrix(intRow, 12))) <> 0 Then
                    Exit Function
                End If
                
                '���۹���
                If CheckPriceAdjustByNO(Val(.RowData(intRow)), Val(.TextMatrix(intRow, 10)), .TextMatrix(intRow, 2), IIf(mbln������ҩ������, mlngҩ��ID, 0)) = False Then
                    Exit Function
                End If
                
                With rsSendRecipeByNo
                    .AddNew
                    !NO = Msf�����б�.TextMatrix(intRow, 2)
                    !���� = Msf�����б�.RowData(intRow)
                    !ҩ��ID = Val(Msf�����б�.TextMatrix(intRow, 10))
                    !��¼���� = Val(Msf�����б�.TextMatrix(intRow, 11))
                    !�����־ = Val(Msf�����б�.TextMatrix(intRow, 12))
                    !���շ� = Val(Msf�����б�.TextMatrix(intRow, 13))
                    !����ID = Val(Msf�����б�.TextMatrix(intRow, 14))
                    !�������� = Msf�����б�.TextMatrix(intRow, 9)
                    .Update
                End With
            End If
        Next
    End With
    
    '���[סԺ����]�Ƿ������������δ��˵ĵ���
    If CheckNotAudited(rsSendRecipeByNo) = False Then Exit Function

    '�������������������ҩ
    rsSendRecipeByNo.Sort = "NO"
    rsSendRecipeByNo.MoveFirst
    For intRow = 1 To rsSendRecipeByNo.RecordCount
        '�ȼ���ִ��Ԥ����
        Call AutoAdjustPrice_ByNO(rsSendRecipeByNo!����, rsSendRecipeByNo!NO)
        
        If Val(rsSendRecipeByNo!��¼����) = 1 Or (Val(rsSendRecipeByNo!��¼����) = 2 And (Val(rsSendRecipeByNo!�����־) = 1 Or Val(rsSendRecipeByNo!�����־) = 4)) Then
            int���� = 1
        Else
            int���� = 2
        End If
        
        '������ҩ��ʱ��֧��ˢ������
'        'δ�շѵĻ��۵�
'        If rsSendRecipeByNo!���� = 8 And rsSendRecipeByNo!���շ� = 0 And mint����δ�շѴ�����ҩ = 0 Then
'            If Not gobjSquareCard Is Nothing And mbln��ҩǰ�շѻ���� = True Then
'                'ˢ���շ�
'                If gobjSquareCard.zlSquareAffirm(Me, 1341, strPrivs, rsSendRecipeByNo!����id, 0, False, 1, rsSendRecipeByNo!NO) = False Then
'                    Exit Function
'                End If
'            Else
'                MsgBox "�ô���[" & rsSendRecipeByNo!NO & "]��δ�շѻ��շѳ�����󣬷�ҩ������ֹ��", vbInformation, gstrSysName
'                Exit Function
'            End If
'        End If
'
'        'δ��˵ļ��˻��۵�
'        If rsSendRecipeByNo!���� = 9 And rsSendRecipeByNo!���շ� = 0 And Int����δ��˴�����ҩ = 0 Then
'            If Not gobjSquareCard Is Nothing And mbln��ҩǰ�շѻ���� = True Then
'                'ˢ���շ�
'                If gobjSquareCard.zlSquareAffirm(Me, 1341, strPrivs, rsSendRecipeByNo!����id, 0, False, 2, rsSendRecipeByNo!NO) = False Then
'                    Exit Function
'                End If
'            Else
'                MsgBox "�ô���[" & rsSendRecipeByNo!NO & "]��δ��˻���˳�����󣬷�ҩ������ֹ��", vbInformation, gstrSysName
'                Exit Function
'            End If
'        End If
        
        If mbln������ҩ������ Then
            gstrSQL = "Zl_ҩƷ�շ���¼_���Ŀⷿ("
            '�ֿⷿID
            gstrSQL = gstrSQL & mlngҩ��ID
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
        End If
        
        gstrSQL = "zl_ҩƷ�շ���¼_������ҩ("
        '�ⷿID
        gstrSQL = gstrSQL & mlngҩ��ID
        '����
        gstrSQL = gstrSQL & "," & rsSendRecipeByNo!����
        'NO
        gstrSQL = gstrSQL & ",'" & rsSendRecipeByNo!NO & "'"
        '�����
        gstrSQL = gstrSQL & ",'" & gstrUserName & "'"
        '��ҩ��
        gstrSQL = gstrSQL & "," & IIf(IntSendAfterDosage = 0, "NULL", "'" & str��ҩ�� & "'")
        'У����
        gstrSQL = gstrSQL & ",NULL"
        '��ҩ��ʽ
        gstrSQL = gstrSQL & ",2"
        '��ҩʱ��
        gstrSQL = gstrSQL & ",to_date('" & StrDate & "','yyyy-MM-dd hh24:mi:ss')"
        '����Ա���
        gstrSQL = gstrSQL & ",'" & gstrUserCode & "'"
        '����Ա����
        gstrSQL = gstrSQL & ",'" & gstrUserName & "'"
        '����λ��
        gstrSQL = gstrSQL & "," & int����λ��
        '��˻��۵�
        gstrSQL = gstrSQL & "," & int��˻��۵�
        '�Ƿ�����
        gstrSQL = gstrSQL & "," & int����
        '�˲���
        gstrSQL = gstrSQL & ",'" & str�˲��� & "'"
        '�Ƿ�δȡҩ
        gstrSQL = gstrSQL & ",NULL"
        '���ܷ�ҩ��
        gstrSQL = gstrSQL & "," & cur���ܷ�ҩ��
        
        gstrSQL = gstrSQL & ")"

        ReDim Preserve arrSql(UBound(arrSql) + 1)
        arrSql(UBound(arrSql)) = gstrSQL
        
        strNo = strNo & rsSendRecipeByNo!���� & "," & rsSendRecipeByNo!NO & "|"
        rsSendRecipeByNo.MoveNext
    Next
    
    '���÷�ҩǰ����ҽӿ�
    err.Clear
    If Not mobjPlugIn Is Nothing Then
        If Right(strNo, 1) = "|" Then strNo = Left(strNo, Len(strNo) - 1)
        On Error Resume Next
        If mobjPlugIn.DrugBeforeSendByRecipe(mlngҩ��ID, strNo, strReserve) = False Then
            If err.Number <> 0 Then
                err.Clear: On Error GoTo 0
            Else
                Exit Function
            End If
        End If
        err.Clear: On Error GoTo 0
    End If
    
    On Error GoTo ErrHand
    
    '�ȴ���ҩ����
    gcnOracle.BeginTrans
    blnInTrans = True
    
    '�����������ǩ�����ŵ�ҵ����ǰ�棬��ֹ���������������
    '����������˵���ǩ��������Ҫ����ҩ�˽��е���ǩ������
    If gblnESign������ҩ = True And gblnESignUserStoped = False Then
        rsSendRecipeByNo.MoveFirst
        For intRow = 1 To rsSendRecipeByNo.RecordCount
            '��ΪҪ�Ȳ�ѯ�����ٸ���ҩ�������Դ���ʱ�õ���ԭ����ҩ������ѯ��ͬʱ��Ҫ������ҩ��������ǩ��ԭ��
            lngҩ��ID = IIf(mbln������ҩ������ = True, Val(rsSendRecipeByNo!ҩ��ID), mlngҩ��ID)
            strǩ����¼ = ""
            If GetSignatureRecored(EsignTache.send, rsSendRecipeByNo!����, rsSendRecipeByNo!NO, _
                    lngҩ��ID, strǩ����¼, 0, CDate(StrDate), gstrUserName, mlngҩ��ID) = False Then
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
        Call zldatabase.ExecuteProcedure(CStr(arrSql(i)), "RecipeWork_Abolish")
    Next
    
    gcnOracle.CommitTrans
    blnInTrans = False
    
    If TypeName(mobjDrugMAC) = "clsDrugPacker" Then
        If mblnConPacker And strNo <> "" And mblnLoadDrug And Not mbln������ҩ������ Then
            Call mobjDrugMAC.DYEY_MZ_TransRecipeList(mstrOpr, UserInfo.�û�����, UserInfo.�û�����, mlngҩ��ID, Mid(strNo, 1, Len(strNo) - 1), strReturn)
        End If
    ElseIf TypeName(mobjDrugMAC) = "clsDrugMachine" Then
        If mblnConPacker Then
            If Right(strNo, 1) = "|" Then strNo = Left(strNo, Len(strNo) - 1)
            mobjDrugMAC.Operation gstrDbUser, Val("22-��ʼ��ҩ"), "1|" & Replace(strNo, "|", ";"), strReturn
        End If
    End If
        
    If MsgBox("����Ҫ��ӡ�����嵥��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
        Call ReportOpen(gcnOracle, glngSys, IIf(glngSys \ 100 = 1, "ZL1_BILL_1341_2", "ZL8_BILL_1341_2"), Me, "�ⷿ=" & IIf(mbln������ҩ������ = True, mlngҩ��ID, Cboҩ��.ItemData(Cboҩ��.ListIndex)), "��ҩ��ʽ=������ҩ|2", "��װϵ��=" & IIf(strUnit = "���ﵥλ", "D.�����װ", "D.סԺ��װ"), "��ҩʱ��=" & StrDate, 2)
    End If
    
    '���÷�ҩ�����ҽӿ�
    If Not mobjPlugIn Is Nothing Then
        If Right(strNo, 1) = "|" Then strNo = Left(strNo, Len(strNo) - 1)
        On Error Resume Next
        mobjPlugIn.DrugSendByRecipe mlngҩ��ID, strNo, CDate(StrDate), strReserve
        err.Clear: On Error GoTo 0
    End If
    
    SendBill = True
    Exit Function
ErrHand:
    If blnInTrans = True Then gcnOracle.RollbackTrans
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function

Private Function CheckCorrelation() As Boolean
    Dim strNo As String, lng���� As Long, str��� As String
    '��鴦���Ƿ��ѽ��ʡ����ò����Ƿ��ѳ�Ժ������Ȩ�޽��м��
    With rs���
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            strNo = !���ݱ�ʶ
            lng���� = Split(strNo, "|")(1)
            strNo = Split(strNo, "|")(0)
            str��� = nvl(!���)
            If Not IsReceiptBalance_Charge(0, strPrivs, lng����, strNo, str���, Val(!��¼����), Val(!�����־)) Then Exit Function
            If Not IsOutPatient(strPrivs, lng����, strNo, Val(!��¼����), Val(!�����־)) Then Exit Function
            .MoveNext
        Loop
    End With
    
    CheckCorrelation = True
End Function

Private Sub InitRec()
    Set rs��� = New ADODB.Recordset
    With rs���
        If .State = 1 Then .Close
        .Fields.Append "���ݱ�ʶ", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "���", adLongVarChar, 500, adFldIsNullable
        .Fields.Append "��¼����", adDouble, 18, adFldIsNullable
        .Fields.Append "�����־", adDouble, 18, adFldIsNullable
        
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
    
    Set rs������Դ���� = New ADODB.Recordset
    With rs������Դ����
        If .State = 1 Then .Close
        .Fields.Append "����", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "��Դ����", adLongVarChar, 100, adFldIsNullable
        
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
    
    Call InitRecSum
    
End Sub

Private Sub txt���˿���_GotFocus()
    GetFocus txt���˿���
End Sub

Private Sub txt���˿���_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    If Trim(txt���˿���.Text) = "" Then Exit Sub
    
    If Select����(Me, Me.txt���˿���, Trim(txt���˿���.Text), "�ٴ�", False, mint�������) = False Then
        Exit Sub
    End If
End Sub

Private Function Select����(ByVal FrmMain As Form, ByVal objCtl As Control, ByVal strSearch As String, _
    Optional str�������� As String = "", _
    Optional bln����Ա As Boolean = False, _
    Optional ByVal int������� As Integer, _
    Optional strsql As String = "") As Boolean
    '------------------------------------------------------------------------------
    '����:����ѡ����
    '����:objCtl-ָ���ؼ�
    '     strSearch-Ҫ����������
    '     str��������-��������:��"V,W,K"
    '     bln����Ա-�Ƿ�Ӳ���Ա����
    '     strSQL-ֱ�Ӹ���SQL��ȡ����(�����ű�ı���һ��Ҫ��A)
    '����:�ɹ�,����true,���򷵻�False
    '------------------------------------------------------------------------------
    Dim i As Long
    Dim blnCancel As Boolean, strKey As String, strTittle As String, lngH As Long, strFind As String
    Dim vRect As RECT
    Dim rstemp  As ADODB.Recordset
    Dim strComment As String
    
    On Error GoTo errHandle
    
    strTittle = "����ѡ����"
    vRect = zlControl.GetControlRect(objCtl.hWnd)
    lngH = objCtl.Height
    
    strKey = GetMatchingSting(strSearch, False)
    
    If strsql <> "" Then
    
        gstrSQL = strsql
    Else
        gstrSQL = "" & _
        "   Select /*+ Rule*/ distinct a.Id,a.�ϼ�id,a.����,a.����,a.����,a.λ�� ,To_Char(a.����ʱ��, 'yyyy-mm-dd') As ����ʱ��, " & _
        "          decode(To_Char(a.����ʱ��, 'yyyy-mm-dd'),'3000-01-01','',To_Char(a.����ʱ��, 'yyyy-mm-dd')) ����ʱ��"
    
        If str�������� = "" And bln����Ա = False Then
            gstrSQL = gstrSQL & vbCrLf & _
            "   From ���ű� a" & _
            "   Where 1=1"
        Else
            gstrSQL = gstrSQL & vbCrLf & _
            "   From ���ű� a, �������ʷ��� b,��������˵�� c," & _
            IIf(str�������� = "", "", "       (Select Column_Value From Table(Cast(f_Str2list([2]) As zlTools.t_Strlist))) J") & _
            "   Where c.�������� = b.����" & IIf(str�������� = "", "(+)", " and B.����=J.column_value ") & _
            "         AND a.id = c.����id and" & IIf(int������� <> 3, " c.�������=[4] ", " (c.�������=1 or c.�������=2 or c.�������=[4])") & _
            IIf(bln����Ա = False, "", " And a.ID IN (Select ����ID From ������Ա Where ��ԱID=[1])")
        End If
        gstrSQL = gstrSQL & vbCrLf & _
            "   and  (a.����ʱ��>=to_date('3000-01-01','yyyy-mm-dd') or a.����ʱ�� is null ) And (a.վ��=[5] or a.վ�� is null) "
    End If
    
    strFind = ""
    If strSearch <> "" Then
        strFind = "   and  (a.���� like upper([3]) or a.���� like upper([3]) or a.���� like [3] )"
        If IsNumeric(strSearch) Then                         '���������,��ֻȡ����
            If Mid(gtype_UserSysParms.Para_���뷽ʽ, 1, 1) = "1" Then strFind = " And (A.���� Like Upper([3]))"
        ElseIf zlCommFun.IsCharAlpha(strSearch) Then           '01,11.����ȫ����ĸʱֻƥ�����
            '0-ƴ����,1-�����,2-����
            '.int���뷽ʽ = Val(zlDatabase.GetPara("���뷽ʽ" ))
            If Mid(gtype_UserSysParms.Para_���뷽ʽ, 2, 1) = "1" Then strFind = " And  (a.���� Like Upper([3]))"
        ElseIf zlCommFun.IsCharChinese(strSearch) Then  'ȫ����
            strFind = " And a.���� Like [3] "
        End If
    End If
    
    If strSearch = "" And str�������� = "" And bln����Ա = False And strsql = "" Then
        gstrSQL = gstrSQL & _
        "   Start With A.�ϼ�id Is Null Connect By Prior A.ID = A.�ϼ�id "
    Else
        gstrSQL = gstrSQL & vbCrLf & strFind & vbCrLf & " Order by A.����"
    End If
    
    If strSearch = "" And str�������� = "" And bln����Ա = False And strsql = "" Then
        '�����¼�
        Set rstemp = zldatabase.ShowSQLMultiSelect(FrmMain, gstrSQL, 1, strTittle, False, "", "", False, False, True, vRect.Left - 15, vRect.Top, lngH, blnCancel, False, False, strKey, int�������)
    Else
        Set rstemp = zldatabase.ShowSQLMultiSelect(FrmMain, gstrSQL, 0, strTittle, False, "", "", False, False, True, vRect.Left - 15, vRect.Top, lngH, blnCancel, False, False, glngUserId, str��������, strKey, int�������, gstrNodeNo)
    End If
    
    If blnCancel = True Then
        Call zlCtlSetFocus(objCtl, True)
        Exit Function
    End If
    
    If rstemp Is Nothing Then
        MsgBox "û�����������Ĳ���,����!"
        If objCtl.Enabled Then objCtl.SetFocus
        Exit Function
    End If
    
    Call zlCtlSetFocus(objCtl, True)
    
    objCtl.Text = ""
    objCtl.Tag = ""
    
    For i = 1 To rstemp.RecordCount
        If i = 1 Then
            objCtl.Text = zlStr.nvl(rstemp!����) & "-" & zlStr.nvl(rstemp!����)
        ElseIf i = 2 Then
            objCtl.Text = objCtl.Text & "..."
        End If
        
        strComment = IIf(strComment = "", "", strComment & ",") & zlStr.nvl(rstemp!����) & "-" & zlStr.nvl(rstemp!����)
        
        '����ID���浽Tag����
        objCtl.Tag = IIf(objCtl.Tag = "", "", objCtl.Tag & ",") & Val(rstemp!Id)
        
        rstemp.MoveNext
    Next
    
    objCtl.ToolTipText = strComment
        
    zlCommFun.PressKey vbKeyTab
    
    Select���� = True
    
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub Txt����_GotFocus()
    GetFocus Txt����
End Sub


Private Sub Txt����_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    Call GetRecipe(3, Txt����)
End Sub


Private Sub txtҽ����_GotFocus()
    GetFocus txtҽ����
End Sub

Private Sub txtҽ����_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    Call GetRecipe(2, txtҽ����)
End Sub


