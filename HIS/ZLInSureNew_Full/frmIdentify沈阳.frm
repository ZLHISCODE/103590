VERSION 5.00
Begin VB.Form frmIdentify���� 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ҽ���������ʶ��"
   ClientHeight    =   6780
   ClientLeft      =   1665
   ClientTop       =   2985
   ClientWidth     =   9345
   Icon            =   "frmIdentify����.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6780
   ScaleWidth      =   9345
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      Caption         =   "������Ϣ"
      Height          =   2115
      Left            =   240
      TabIndex        =   38
      Top             =   4110
      Width           =   8925
      Begin VB.Frame Frame3 
         Caption         =   "�涨����Ϣ(&F)"
         Height          =   1575
         Left            =   4650
         TabIndex        =   41
         Top             =   390
         Width           =   4095
         Begin VB.TextBox txt������� 
            Enabled         =   0   'False
            Height          =   300
            Left            =   1410
            MaxLength       =   12
            TabIndex        =   43
            Top             =   300
            Width           =   2415
         End
         Begin VB.TextBox txt�������� 
            Enabled         =   0   'False
            Height          =   300
            Left            =   1410
            MaxLength       =   20
            TabIndex        =   45
            Top             =   690
            Width           =   2415
         End
         Begin VB.TextBox txt�������� 
            Enabled         =   0   'False
            Height          =   300
            Left            =   1410
            MaxLength       =   60
            TabIndex        =   47
            Top             =   1080
            Width           =   2415
         End
         Begin VB.Label lbl������� 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "�������(&D)"
            Enabled         =   0   'False
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   360
            TabIndex        =   42
            Top             =   360
            Width           =   990
         End
         Begin VB.Label lbl�������� 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "��������(&B)"
            Enabled         =   0   'False
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   360
            TabIndex        =   44
            Top             =   750
            Width           =   990
         End
         Begin VB.Label lbl�������� 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "��������(&X)"
            Enabled         =   0   'False
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   360
            TabIndex        =   46
            Top             =   1140
            Width           =   990
         End
      End
      Begin VB.TextBox txt���������Ϣ 
         Enabled         =   0   'False
         Height          =   1410
         Left            =   300
         MaxLength       =   12
         MultiLine       =   -1  'True
         TabIndex        =   40
         Tag             =   "persfundcon"
         Top             =   540
         Width           =   3885
      End
      Begin VB.Label lbl���������Ϣ 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "���������Ϣ(&U)"
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   300
         TabIndex        =   39
         Top             =   300
         Width           =   1350
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   6510
      TabIndex        =   48
      Top             =   6360
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   7860
      TabIndex        =   49
      Top             =   6360
      Width           =   1100
   End
   Begin VB.CommandButton cmdChange 
      Caption         =   "�޸�����(&M)"
      Height          =   405
      Left            =   240
      TabIndex        =   50
      Top             =   6330
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "���˻�����Ϣ"
      Height          =   3885
      Left            =   180
      TabIndex        =   0
      Top             =   150
      Width           =   8985
      Begin VB.CommandButton cmd������Ϣ 
         Caption         =   "��"
         Height          =   300
         Left            =   8460
         TabIndex        =   51
         TabStop         =   0   'False
         Top             =   3450
         Width           =   285
      End
      Begin VB.TextBox txt������Ϣ 
         Height          =   300
         Left            =   1530
         TabIndex        =   37
         Top             =   3450
         Width           =   6915
      End
      Begin VB.ComboBox cboҵ������ 
         Height          =   300
         Left            =   1530
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   720
         Width           =   1485
      End
      Begin VB.TextBox txt���� 
         Height          =   300
         Left            =   1530
         MaxLength       =   10
         TabIndex        =   2
         Top             =   330
         Width           =   1485
      End
      Begin VB.TextBox txtסԺ���� 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   300
         Left            =   1530
         MaxLength       =   2
         TabIndex        =   17
         Top             =   2670
         Width           =   525
      End
      Begin VB.TextBox txt�ʻ���� 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   300
         Left            =   1530
         MaxLength       =   18
         TabIndex        =   14
         Top             =   2280
         Width           =   1455
      End
      Begin VB.TextBox txt�α��˵�λ 
         Enabled         =   0   'False
         Height          =   300
         Left            =   5970
         MaxLength       =   100
         TabIndex        =   35
         Tag             =   "corp_id|corp_name"
         Top             =   3060
         Width           =   2775
      End
      Begin VB.TextBox txt��ذ��ó��� 
         Enabled         =   0   'False
         Height          =   300
         Left            =   5970
         MaxLength       =   100
         TabIndex        =   33
         Tag             =   "city_code|city_name"
         Top             =   2670
         Width           =   2775
      End
      Begin VB.TextBox txtְ�� 
         Enabled         =   0   'False
         Height          =   300
         Left            =   5970
         MaxLength       =   10
         TabIndex        =   23
         Tag             =   "position_name"
         Top             =   720
         Width           =   2775
      End
      Begin VB.TextBox txt���⹤�� 
         Enabled         =   0   'False
         Height          =   300
         Left            =   5970
         MaxLength       =   50
         TabIndex        =   31
         Tag             =   "work_type|work_type_name"
         Top             =   2280
         Width           =   2775
      End
      Begin VB.TextBox txt�����չ���Ⱥ 
         Enabled         =   0   'False
         Height          =   300
         Left            =   5970
         MaxLength       =   50
         TabIndex        =   29
         Tag             =   "special_code|special_name"
         Top             =   1890
         Width           =   2775
      End
      Begin VB.TextBox txt����Ա���� 
         Enabled         =   0   'False
         Height          =   300
         Left            =   5970
         MaxLength       =   50
         TabIndex        =   27
         Tag             =   "official_code|official_name"
         Top             =   1500
         Width           =   2775
      End
      Begin VB.TextBox txt���� 
         Enabled         =   0   'False
         Height          =   300
         Left            =   5970
         MaxLength       =   50
         TabIndex        =   25
         Tag             =   "folk_code|folk_name"
         Top             =   1110
         Width           =   2775
      End
      Begin VB.TextBox txt��Ա��� 
         Enabled         =   0   'False
         Height          =   300
         Left            =   5970
         MaxLength       =   50
         TabIndex        =   21
         Tag             =   "pers_type|pers_name"
         Top             =   330
         Width           =   2775
      End
      Begin VB.TextBox txt���֤�� 
         Enabled         =   0   'False
         Height          =   300
         Left            =   1530
         MaxLength       =   18
         TabIndex        =   19
         Tag             =   "idcard"
         Top             =   3060
         Width           =   2715
      End
      Begin VB.ComboBox cbo�Ա� 
         Enabled         =   0   'False
         Height          =   300
         Left            =   3420
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   1890
         Width           =   825
      End
      Begin VB.TextBox txt���� 
         Enabled         =   0   'False
         Height          =   300
         Left            =   1530
         MaxLength       =   10
         TabIndex        =   10
         Tag             =   "name"
         Top             =   1890
         Width           =   1035
      End
      Begin VB.TextBox txtҽ���� 
         Enabled         =   0   'False
         Height          =   300
         Left            =   1530
         MaxLength       =   30
         TabIndex        =   8
         Tag             =   "insr_code"
         Top             =   1500
         Width           =   2715
      End
      Begin VB.TextBox txt���˱�� 
         Enabled         =   0   'False
         Height          =   300
         Left            =   1530
         MaxLength       =   8
         TabIndex        =   6
         Tag             =   "indi_id"
         Top             =   1110
         Width           =   1485
      End
      Begin VB.Label lbl������Ϣ 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "������Ϣ(&J)"
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   480
         TabIndex        =   36
         Top             =   3510
         Width           =   990
      End
      Begin VB.Label lblҵ������ 
         Caption         =   "ҵ������(&E)"
         Height          =   225
         Left            =   465
         TabIndex        =   3
         Top             =   758
         Width           =   1005
      End
      Begin VB.Label lbl���� 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "����(&R)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   840
         TabIndex        =   1
         Top             =   390
         Width           =   630
      End
      Begin VB.Label lbl��λ 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Ԫ"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   3060
         TabIndex        =   15
         Top             =   2340
         Width           =   180
      End
      Begin VB.Label lblסԺ���� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "סԺ����(&Y)"
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   480
         TabIndex        =   16
         Top             =   2730
         Width           =   990
      End
      Begin VB.Label lbl�ʻ���� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "�ʻ����(&Q)"
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   480
         TabIndex        =   13
         Top             =   2340
         Width           =   990
      End
      Begin VB.Label lbl�α��˵�λ 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "�α��˵�λ(&I)"
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   4740
         TabIndex        =   34
         Top             =   3120
         Width           =   1170
      End
      Begin VB.Label lbl��ذ��ó��� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "��ذ��ó���(&A)"
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   4560
         TabIndex        =   32
         Top             =   2730
         Width           =   1350
      End
      Begin VB.Label lblְ�� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "ְ��(&Z)"
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   5280
         TabIndex        =   22
         Top             =   780
         Width           =   630
      End
      Begin VB.Label lbl���⹤�� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "���⹤��(&G)"
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   4920
         TabIndex        =   30
         Top             =   2340
         Width           =   990
      End
      Begin VB.Label lbl�����չ���Ⱥ 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "�����չ���Ⱥ(&T)"
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   4560
         TabIndex        =   28
         Top             =   1950
         Width           =   1350
      End
      Begin VB.Label lbl����Ա���� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "����Ա����(&S)"
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   4740
         TabIndex        =   26
         Top             =   1560
         Width           =   1170
      End
      Begin VB.Label lbl���� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "����(&R)"
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   5280
         TabIndex        =   24
         Top             =   1170
         Width           =   630
      End
      Begin VB.Label lbl��Ա��� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "��Ա���(&L)"
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   4920
         TabIndex        =   20
         Top             =   390
         Width           =   990
      End
      Begin VB.Label lbl���֤�� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "���֤��(&K)"
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   480
         TabIndex        =   18
         Top             =   3120
         Width           =   990
      End
      Begin VB.Label lbl�Ա� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "�Ա�(&S)"
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   2730
         TabIndex        =   11
         Top             =   1950
         Width           =   630
      End
      Begin VB.Label lbl���� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "����(&N)"
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   840
         TabIndex        =   9
         Top             =   1950
         Width           =   630
      End
      Begin VB.Label lblҽ���� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "ҽ����(&Y)"
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   660
         TabIndex        =   7
         Top             =   1560
         Width           =   810
      End
      Begin VB.Label lbl���˱�� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "���˱��(&P)"
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   480
         TabIndex        =   5
         Top             =   1170
         Width           =   990
      End
   End
End
Attribute VB_Name = "frmIdentify����"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'���볣�����ܶ���ɹ����ģ�������ʹ�õ��ĵط��������壬�ڱ���ʱͳһ�޸�
#Const gverControl = 99  ' 0-��֧�ֶ�̬ҽ��(9.19��ǰ),1-֧�ֶ�̬ҽ���޸��Ӳ���(9.22��ǰ) , _
'    2-����������������ʽ��������һ��;����������ԭʼ��������һ��;�����շ�����������;3-������������GetNextNO();
'    99-���н������Ӹ��Ӳ���(���°�)

Private mstrReturn As String
Private mstr������ As String
Private mlng����ID As Long
Private mbytType As Byte
Private mblnStart As Boolean
Private mstr���� As String
Private mbln�������� As Boolean
Private mbln�ಡ�� As Boolean
Private mrs���� As New ADODB.Recordset
Private Const madLongVarCharDefault As Integer = 10          '�ַ����ֶ�ȱʡ����
Private Const madDoubleDefault As Integer = 18               '�������ֶ�ȱʡ����
Private Const madDbDateDefault As Integer = 20               '�������ֶ�ȱʡ����

'--------------------------------------------------
'����涨���Ĳ���ֻ�ܴӽӿڷ��ص��������Ĳ�����ѡ��
'�������˿��Դ����в�����ѡ��

Public Function GetPatient(ByVal bytType As Byte, Optional ByVal lng����ID As Long = 0) As String
    mstrReturn = ""
    mstr������ = ""
    mlng����ID = lng����ID
    mbytType = bytType
    Me.Show 1
    
    GetPatient = mstrReturn
End Function

Private Sub cboҵ������_Click()
    If Not mblnStart Then Exit Sub
    Call txt����_KeyDown(vbKeyReturn, 0)
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdChange_Click()
    mstr������ = frm�޸�����.ChangePassword(txt����.Text)
End Sub

Private Sub cmdOK_Click()
    Dim bln��Ժ As Boolean
    Dim str�������� As String
    Dim intDays As Integer
    Dim strBeginDate As String
    Dim strIdentify As String, strAddition As String
    Dim lng����ID As Long
    Dim str˳��� As String
    Dim rsTemp As New ADODB.Recordset
    
    If Trim(txt���˱��.Text) = "" Then
        MsgBox "��δ��ȡ���˵Ļ�����Ϣ������������󰴻س���", vbInformation, gstrSysName
        txt����.SetFocus
        Exit Sub
    End If
    
    '��鲡��״̬
    gstrSQL = "select nvl(��ǰ״̬,0) as ״̬,˳��� from �����ʻ� where ����=[1] and ҽ����=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, TYPE_������, CStr(txtҽ����.Text))
    
    If rsTemp.RecordCount > 0 Then
        If rsTemp("״̬") > 0 Then
            str˳��� = Nvl(rsTemp!˳���)
            bln��Ժ = True
            If Not mbln�������� Then
                MsgBox "�ò����Ѿ���Ժ������ͨ�������֤��", vbInformation, gstrSysName
                Exit Sub
            End If
        End If
    End If
    
    'ת����������
    str�������� = "1980-01-01"
    
    '����ѡ������Ϣ
    If txt������Ϣ.Tag = "" Then
        MsgBox "��Ϊ�òα�����ѡ�񼲲�������Ϣ��", vbInformation, gstrSysName
        txt������Ϣ.SetFocus
        Exit Sub
    End If
    gstrSQL = "Select ID From ���ղ��� Where ����=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ���ղ���", CStr(txt������Ϣ.Tag))
    If Not rsTemp.EOF Then
        lng����ID = rsTemp!ID
    End If
    
    '�����ַ���
    '�������˵�����Ϣ�������ʽ��
    '0����;1ҽ����;2����;3����;4�Ա�;5��������;6���֤;7��λ����(����);
    '8.���Ĵ���;9.˳���;10��Ա���;11�ʻ����;12��ǰ״̬;13����ID;14��ְ(0,1);15����֤��;16�����;17�Ҷȼ�
    '18�ʻ������ۼ�,19�ʻ�֧���ۼ�,20����ͳ���ۼ�,21ͳ�ﱨ���ۼ�,22סԺ�����ۼ�
    strIdentify = txt���˱��.Text                              '0����
    strIdentify = strIdentify & ";" & txtҽ����.Text            '1ҽ����
    strIdentify = strIdentify & ";" & txt����.Text              '2����
    strIdentify = strIdentify & ";" & txt����.Text              '3����
    strIdentify = strIdentify & ";" & cbo�Ա�.Text              '4�Ա�
    strIdentify = strIdentify & ";" & str��������               '5��������
    strIdentify = strIdentify & ";" & txt���֤��.Text          '6���֤
    strIdentify = strIdentify & ";" & txt�α��˵�λ.Text        '7.��λ����(����)
    strAddition = ";0"                                          '8.���Ĵ���
    strAddition = strAddition & ";" & str˳���                 '9.˳���
    strAddition = strAddition & ";" & txt��Ա���.Text          '10��Ա���
    strAddition = strAddition & ";" & Val(txt�ʻ����.Text)     '11�ʻ����
    strAddition = strAddition & ";0"                            '12��ǰ״̬
    strAddition = strAddition & ";" & lng����ID                 '13����ID
    strAddition = strAddition & ";1"                            '14��ְ(1,2,3)
    strAddition = strAddition & ";"                             '15����֤��
    strAddition = strAddition & ";"                             '16�����
    strAddition = strAddition & ";" & cboҵ������.ItemData(cboҵ������.ListIndex)                           '17�Ҷȼ�
    strAddition = strAddition & ";" & Val(txt�ʻ����.Text)     '18�ʻ������ۼ�
    strAddition = strAddition & ";0"                            '19�ʻ�֧���ۼ�
    strAddition = strAddition & ";0"                            '20���깤���ܶ�
    strAddition = strAddition & ";" & Val(txtסԺ����.Text)     '21סԺ�����ۼ�
    
    mlng����ID = BuildPatiInfo(0, strIdentify & strAddition, mlng����ID, TYPE_������)
    '���ظ�ʽ:�м���벡��ID
    If mlng����ID > 0 Then
        mstrReturn = strIdentify & ";" & mlng����ID & strAddition
    End If
    
    '���������ҵ�������ϵͳ�涨�ĹҺ���Ч������û�йҺż�¼���������շ�
    If mbytType = 0 And bln��Ժ = False Then
        #If gverControl >= 4 Then
            intDays = Val(zlDatabase.GetPara("�Һ���Ч����", glngSys, , 0)) - 1
        #Else
            intDays = Val(GetPara("�Һ���Ч����", glngSys, , , 0)) - 1
        #End If
        
        '����Һ���Ч����Ϊ�㣬��ʾ�����շ�ǰ�����Բ��Һ�
        If intDays > -1 Then
            strBeginDate = Format(DateAdd("d", IIf(intDays = -1, 30, intDays) * -1, zlDatabase.Currentdate()), "yyyy-MM-dd")
            
            'ȡ�ö�ʱ���ڣ����޹Һż�¼
            gstrSQL = " Select 1 From ������ü�¼" & _
                      " Where ��¼����=4 And ��¼״̬=1 And ����ID=[1] And ����ʱ��>[2]"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ�涨��ʱ���ڣ��ò������޹Һż�¼", mlng����ID, CDate(strBeginDate))
            If rsTemp.EOF Then
                MsgBox "�ò��˻�û�Һţ����ܽ���������ݵǼǣ�", vbInformation, gstrSysName
                mstrReturn = ""
                Exit Sub
            End If
        End If
    End If
    
    If mbytType = 1 Then
        gstrSQL = "zl_�����ʻ�_������Ϣ(" & mlng����ID & "," & TYPE_������ & ",'ҵ������','''" & cboҵ������.ItemData(cboҵ������.ListIndex) & "''')"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "����סԺҵ������")
    End If
    
    If mstr������ <> "" Then
        '�����޸�����ӿڣ����ɹ�������ʾ������
'        1   card_no    ҽ��������  20  ��
'        2   card_type  ��֤����    12  ��  "01"��ҽ����
'        3   password   ԭ����      6   ��
'        4   newpassword������      6   ��
        If ���ýӿ�_׼��_������(Function_������.����_�޸�����) Then
            'д��ڲ���
            gstrField_������ = "card_no||card_type||password||newpassword"
            gstrValue_������ = mstr���� & "||" & "01" & "||" & txt����.Text & "||" & mstr������
            Call ���ýӿ�_д��ڲ���_������(1)
            If ���ýӿ�_ִ��_������() Then
                '���¸����ʻ��е���Ϣ
                gstrSQL = "zl_�����ʻ�_������Ϣ(" & mlng����ID & "," & TYPE_������ & ",'����','''" & mstr������ & "''')"
                Call zlDatabase.ExecuteProcedure(gstrSQL, "��������")
            Else
                MsgBox "�����޸�ʧ�ܣ��Կɼ���������", vbInformation, gstrSysName
            End If
        End If
    End If
    
    gCominfo_������.ҵ������ = cboҵ������.ItemData(cboҵ������.ListIndex)
    gCominfo_������.�������� = txt������Ϣ.Tag
    gCominfo_������.���˱�� = txt���˱��.Text
    gCominfo_������.�ʻ���� = Val(txt�ʻ����.Text)
    
    Unload Me
End Sub

Private Sub cmd������Ϣ_Click()
    Dim bln���ⲡ As Boolean
    Dim rs���� As ADODB.Recordset
    bln���ⲡ = (Me.cboҵ������.ItemData(Me.cboҵ������.ListIndex) = ҵ�����_������.����涨��)
    
    If Not bln���ⲡ Then
        gstrSQL = " Select A.ID,A.����,A.����,A.���� " & _
                " From ���ղ��� A where A.����=[1]"
        Set rs���� = zlDatabase.OpenSQLRecord(gstrSQL, "�����֤", TYPE_������)
        If rs����.RecordCount > 0 Then
            If frmListSel.ShowSelect(TYPE_������, rs����, "ID", "ҽ������ѡ��", "��ѡ��ҽ�����֣�") = True Then
                txt������Ϣ.Tag = rs����!����
                txt������Ϣ.Text = "(" & rs����!���� & ")" & rs����!����
                lbl������Ϣ.Tag = txt������Ϣ.Text '���ڻָ���ʾ
            End If
        End If
    Else
        If mrs����.RecordCount > 0 Then
            If frmListSel.ShowSelect(TYPE_������, mrs����, "ID", "���ⲡ��ѡ��", "��ѡ���ض���ҽ�����֣�") = True Then
                txt������Ϣ.Tag = mrs����!����
                txt������Ϣ.Text = "(" & mrs����!���� & ")" & mrs����!����
                lbl������Ϣ.Tag = txt������Ϣ.Text '���ڻָ���ʾ
            End If
        End If
    End If
    cmdOK.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Dim rsTemp As New ADODB.Recordset
    mblnStart = False
    
    With cbo�Ա�
        .Clear
        .AddItem "Ů"
        .AddItem "��"
        .ListIndex = 1
    End With
    
    With cboҵ������
        .Clear
        If mbytType = 0 Or mbytType = 2 Then
            .AddItem "��ͨ����"
            .ItemData(.NewIndex) = ҵ�����_������.��ͨ����
            .AddItem "����涨��"
            .ItemData(.NewIndex) = ҵ�����_������.����涨��
            .AddItem "���Ｑ��"
            .ItemData(.NewIndex) = ҵ�����_������.���Ｑ��
            .AddItem "�����ؼ�"
            .ItemData(.NewIndex) = ҵ�����_������.�����ؼ�
            .AddItem "��������"
            .ItemData(.NewIndex) = ҵ�����_������.��������
            .AddItem "��������"
            .ItemData(.NewIndex) = ҵ�����_������.��������
        ElseIf mbytType = 1 Or mbytType = 2 Then
            .AddItem "��ͨסԺ"
            .ItemData(.NewIndex) = ҵ�����_������.��ͨסԺ
            .AddItem "��ͥ����"
            .ItemData(.NewIndex) = ҵ�����_������.��ͥ����
            .AddItem "����סԺ"
            .ItemData(.NewIndex) = ҵ�����_������.����סԺ
            .AddItem "����סԺ"
            .ItemData(.NewIndex) = ҵ�����_������.����סԺ
        ElseIf mbytType = 3 Then
            .AddItem "��ͨ�Һ�"
            .ItemData(.NewIndex) = ҵ�����_������.��ͨ����
        End If
        .ListIndex = 0
        .Enabled = (mbytType <> 2)
    End With
    lbl������Ϣ.Enabled = True
    txt������Ϣ.Enabled = True
    
    'ȡסԺ�����Ƿ���������ҵ��
    gstrSQL = "Select Nvl(����ֵ,0) ����ֵ From ���ղ��� Where ���=7 And ����=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ȡסԺ�����Ƿ���������ҵ��", TYPE_������)
    If Not rsTemp.EOF Then
        mbln�������� = (rsTemp!����ֵ = 1)
    End If
    
    '����ǹҺţ�����ʾȱʡ���0000000-����
    If mbytType = 3 Then
        gstrSQL = "Select ID,����,���� From ���ղ��� Where ����='0000000'"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ����Һ�ȱʡ����")
        If rsTemp.EOF Then
            MsgBox "���ʼ���Һ�ȱʡ���֣���(0000000)������"
        Else
            txt������Ϣ.Tag = rsTemp!����
            txt������Ϣ.Text = "(" & rsTemp!���� & ")" & rsTemp!����
            lbl������Ϣ.Tag = txt������Ϣ.Text '���ڻָ���ʾ
        End If
    End If
    
    mblnStart = True
End Sub

Private Sub ClearCons()
    Dim strFields As String
    Dim objTextBox As Control
    strFields = "ID" & "," & adDouble & "," & "18" & "|" & _
                "����" & "," & adVarChar & "," & "50" & "|" & _
                "����" & "," & adLongVarChar & "," & "100"
    Call Record_Init(mrs����, strFields)
    
    For Each objTextBox In Controls
        If UCase(TypeName(objTextBox)) = "TEXTBOX" And Not (objTextBox.Name = "txt����" Or objTextBox.Name = "txt������Ϣ") Then
            objTextBox.Text = ""
        End If
    Next
End Sub

Private Sub txt������Ϣ_GotFocus()
    OpenIme ""
    Call zlControl.TxtSelAll(txt������Ϣ)
End Sub

Private Sub txt������Ϣ_KeyPress(KeyAscii As Integer)
    Dim rsTemp As New ADODB.Recordset, strText As String
    Dim blnReturn As Boolean, lst As ListItem
    Dim bln���ⲡ As Boolean
    
    If KeyAscii <> vbKeyReturn Then Exit Sub
    If txt������Ϣ.Text = "" And txt������Ϣ.Tag <> "" Then Exit Sub
    bln���ⲡ = (Me.cboҵ������.ItemData(Me.cboҵ������.ListIndex) = ҵ�����_������.����涨��)
    
    On Error GoTo errHandle
    
    strText = txt������Ϣ.Text
    If InStr(1, strText, "(") <> 0 Then
        If InStr(1, strText, ")") <> 0 Then
            strText = Mid(strText, 2, InStr(1, strText, ")") - 2)
        End If
    End If
    If Not bln���ⲡ Then
        gstrSQL = "Select A.ID,A.����,A.����,A.����" & _
                 "   FROM ���ղ��� A WHERE A.����=[1] And (" & _
                 "A.���� like [2] || '%' or A.���� like [2] || '%' or A.���� like [2] || '%')"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, TYPE_������, strText)
        If rsTemp.RecordCount = 0 Then
            MsgBox "�����ڸò��֣����������룡", vbInformation, gstrSysName
            txt������Ϣ.Text = lbl������Ϣ.Tag
            zlControl.TxtSelAll txt������Ϣ
            Exit Sub
        Else
            '����ѡ����
            If rsTemp.RecordCount > 1 Then
                '�����ֶδ���3�ģ���ʹֻ��һ����¼�ѸöԻ�����ʾ�������Ա����û��õ��������Ϣ
                blnReturn = frmListSel.ShowSelect(TYPE_������, rsTemp, "ID", "ҽ������ѡ��", "��ѡ��ҽ�����֣�")
            Else
                blnReturn = True
            End If
        End If
    Else
        If IsNumeric(strText) Then
            mrs����.Filter = "���� Like '" & strText & "*'"
        Else
            mrs����.Filter = "���� Like '" & strText & "*'"
        End If
        If mrs����.RecordCount = 0 Then
            MsgBox "�����ڸ����ⲡ�֣����������룡", vbInformation, gstrSysName
            mrs����.Filter = 0
            txt������Ϣ.Text = lbl������Ϣ.Tag
            zlControl.TxtSelAll txt������Ϣ
            Exit Sub
        Else
            If mrs����.RecordCount > 1 Then
                blnReturn = frmListSel.ShowSelect(TYPE_������, mrs����, "ID", "���ⲡ��ѡ��", "��ѡ���ض���ҽ�����֣�")
            Else
                blnReturn = True
            End If
        End If
    End If
    
    If blnReturn = False Then
        '��¼����û�п�ѡ�������
        txt������Ϣ.Text = lbl������Ϣ.Tag
        zlControl.TxtSelAll txt������Ϣ
        If bln���ⲡ Then mrs����.Filter = 0
        Exit Sub
    Else
        '�϶����м�¼����
        If Not bln���ⲡ Then
            txt������Ϣ.Tag = rsTemp!����
            txt������Ϣ.Text = "(" & rsTemp!���� & ")" & rsTemp!����
            lbl������Ϣ.Tag = txt������Ϣ.Text '���ڻָ���ʾ
        Else
            txt������Ϣ.Tag = mrs����!����
            txt������Ϣ.Text = "(" & mrs����!���� & ")" & mrs����!����
            lbl������Ϣ.Tag = txt������Ϣ.Text '���ڻָ���ʾ
        End If
    End If
    
    If bln���ⲡ Then mrs����.Filter = 0
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    If bln���ⲡ Then mrs����.Filter = 0
End Sub

Private Sub txt����_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim blnTrans As Boolean
    Dim lngҵ������ As Long
    Dim strTemp As String, strData As String
    Dim str���˱�� As String, str�α��˵�λ As String
    Dim rsTemp As New ADODB.Recordset
    Const str�����ʻ� As String = "003"
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    If Not (mbytType = 0 Or mbytType = 1 Or mbytType = 3) Then Exit Sub
    
    On Error GoTo errHand
    '�����������
    Call ClearCons
    lngҵ������ = cboҵ������.ItemData(cboҵ������.ListIndex)
    
    '--��IC��
    '����IC���е���Ϣ
    If Not ���ýӿ�_׼��_������(Function_������.����_����) Then Exit Sub
    If Not ���ýӿ�_ִ��_������ Then Exit Sub
    
    'ȡ���صļ�¼��
    'If Not ���ýӿ�_ָ����¼��_������("ICInfo") Then Exit Sub
    'Modified By ���� ��������ɳ ԭ�򣺽���������ɿ��Ÿ�Ϊ���֤��
    'If Not ���ýӿ�_��ȡ����_������("indi_id", str���˱��) Then Exit Sub
    If Not ���ýӿ�_��ȡ����_������("card_no", str���˱��) Then Exit Sub
    
    '--���ÿ��Ƿ��ں������У����ö����ӿ�ʱ���ӿ����Ѵ�������ݣ�
'    If Not ���ýӿ�_׼��_������(Function_������.����_������У��) Then Exit Sub
'    If Not ���ýӿ�_ִ��_������ Then Exit Sub
    
    '--��ȡ���˻�����Ϣ��ҵ����Ϣ��str���˱��="12010619671003053"��
    Select Case mbytType
    Case 0 '����
'        glngReturn_������ = CZ_Start(glngInterface_������, Function_������.��ͨ����_�����֤)
'        1   idcard         ���֤��        20  ��
'        2   hospital_id    ҽ�ƻ�������    20  ��
'        3   busi_type      ҵ������        2   ��  "11"������
'        4   password       ����            6   ��
        If Not ���ýӿ�_׼��_������(IIf(lngҵ������ = ҵ�����_������.����涨��, Function_������.����涨��_�����֤, Function_������.��ͨ����_�����֤)) Then Exit Sub
    Case 1  'סԺ
        'glngReturn_������ = CZ_Start(glngInterface_������, Function_������.��ͨסԺ_�����֤)
'        1   iccardno       �ſ�����        20  ��
'        2   hospital_id    ҽ�ƻ�������    20  ��
'        3   busi_type      ҵ������        2   ��  "12"��סԺ
'        4   reg_flag       �ǼǱ�־        1   ��  "0"����ͨסԺ�Ǽ�
'        5   password       ����            6   ��
        If Not ���ýӿ�_׼��_������(Function_������.��ͨסԺ_�����֤) Then Exit Sub
    Case 3 '�Һ�
        If Not ���ýӿ�_׼��_������(Function_������.��ͨ����_�����֤) Then Exit Sub
    End Select
    
    '��д��ڲ���
    'Modified By ���� ��������ɳ ԭ�򣺽���������ɿ��Ÿ�Ϊ���֤��
    Call CZ_DataPut(glngInterface_������, 1, "iccardno", str���˱��)
    Call CZ_DataPut(glngInterface_������, 1, "hospital_id", gCominfo_������.ҽԺ����)
    Call CZ_DataPut(glngInterface_������, 1, "busi_type", lngҵ������)
    If mbytType = 1 Then Call CZ_DataPut(glngInterface_������, 1, "reg_flag", "0")
    Call CZ_DataPut(glngInterface_������, 1, "password", txt����.Text)
    '���ýӿ�
    If Not ���ýӿ�_ִ��_������ Then Exit Sub
    '��ȡ���ؼ�¼��������TextBox��Tag��ȡ������Ϣ1-22��
'    1   indi_id        ���˱��            8
'    2   insr_code      ���պ�              30
'    3   name           ����                10
'    4   sex            �Ա�                1   "0"��Ů    "1"����
'    5   idcard         ���֤����          20
'    6   pers_type      ��Ա������        2
'    7   pers_name      ��Ա�������        20
'    8   folk_code      �������            2
'    9   folk_name      ��������            20
'    10  official_code  ����Ա����          2
'    11  official_name  ����Ա��������      20
'    12  special_code   �����չ���Ⱥ����    3
'    13  special_name   �����չ���Ⱥ����    20
'    14  position_name  ְ��                10
'    15  work_type      ���⹤��            3
'    16  work_type_name ���⹤������        20
'    17  city_code      ��ذ��ó��б���    30
'    18  city_name      ��ذ��ó�������    30
'    19  corp_id        �α��˵�λ����      20
'    20  corp_name      �α��˵�λ����      50
'    22  persfundcon    �Ѷ��������Ϣ      1024
'--------------------�������κνӿڶ�Ҫ���صĻ�����Ϣ--------------------
'��סԺ��      21  sum_year       �����ۼ�סԺ����    2
'�����      21  last_balance   �����ʻ����        18  ��λ��Ԫ
'���涨����    22  serial_apply   �������            12
'���涨����    23  icd            ��������            20  ���������涨���Ĳ��ֱ���
'���涨����    24  disease        ��������            60  ���������涨���ļ�������
    If Not ���ýӿ�_ָ����¼��_������("PersonInfo") Then Exit Sub
    
    '����TextBox��Tag��ȡ������Ϣ1-22
    Call ReadFromInterface
    '���������������������סԺ���ˣ�����������Ⱥ��ֵ�ж�
    If mint���õ���_���� = 2 And mbytType = 1 Then
        '��������
        Call ���ýӿ�_��ȡ����_������("special_code", strTemp)
        Select Case strTemp
        Case "1"
            MsgBox "�ò�����������������ס��Ժ��ѡ����ʱ��ע��ѡ��", vbInformation, gstrSysName
        Case "0"
            MsgBox "�ò��������û�ж�������סԺ��¼��ѡ����ʱ��ע��ѡ��", vbInformation, gstrSysName
        End Select
    End If
    
    Call ���ýӿ�_��ȡ����_������("sex", strData)
    Me.cbo�Ա�.ListIndex = Val(strData)
    
    '���ݲ�ͬ��ҵ�����ͣ���ȡ�ֶ�ֵ
    Call ���ýӿ�_��ȡ����_������("corp_id", str�α��˵�λ)
    Select Case cboҵ������.ItemData(cboҵ������.ListIndex)
    Case ҵ�����_������.��ͨ����, ҵ�����_������.���Ｑ��, _
         ҵ�����_������.�����ؼ�, ҵ�����_������.��������, ҵ�����_������.��������
        Call ���ýӿ�_��ȡ����_������("last_balance", strData)
        txt�ʻ���� = Format(strData, "#####0.00;-#####0.00; ;")
    Case ҵ�����_������.����涨��
        'ȡ���صļ�����¼��
        On Error Resume Next
        Dim lngID As Long
        Dim str���� As String, str���� As String
        Dim strColumns As String, strValues As String
        '��ʾ�������ʼ������ʾ�����һ���������Ϣ
        Call ���ýӿ�_ָ����¼��_������("PersonInfo")
        Call ���ýӿ�_��ȡ����_������("serial_apply", strData)
        txt������� = strData
        
        glngReturn_������ = CZ_SetRecordset(glngInterface_������, "spinfo")
        mbln�ಡ�� = (glngReturn_������ > 0)        '�趨��¼��ʱ������ɹ����ؼ�¼�������򷵻�-1
        
        '��������¼���������ڴ���
        On Error GoTo errHand
        strColumns = "ID|����|����"
        blnTrans = True
        
        gcnOracle.BeginTrans
        If mbln�ಡ�� Then
            Call DebugTool("�ಡ��")
            Do While True
                strValues = ""
                Call ���ýӿ�_��ȡ����_������("icd", str����)
                strValues = strValues & "|" & str����
                Call ���ýӿ�_��ȡ����_������("disease", str����)
                strValues = strValues & "|" & str����
                '�ж��Ƿ���ڸò���
                gstrSQL = " Select ID From ���ղ���" & _
                          " Where ����=[1] And ����=[2]"
                Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "�Ƿ���ڴ˲���", TYPE_������, str����)
                If rsTemp.RecordCount = 0 Then
                    lngID = zlDatabase.GetNextID("���ղ���")
                    gstrSQL = "zl_���ղ���_INSERT(" & lngID & "," & TYPE_������ & ",'" & str���� & _
                                "','" & str���� & "',NULL,0,NULL,NULL)"
                Else
                    lngID = rsTemp!ID
                End If
                
                strValues = lngID & strValues
                Call Record_Add(mrs����, strColumns, strValues)
                
                Call DebugTool("�Ѽ���һ�м�¼")
                If Not ���ýӿ�_�ƶ���¼��_������(MoveNext) Then Exit Do
            Loop
        Else
            '���Ƕಡ�ֵĻ����϶�ֻ��һ��
            Call DebugTool("������")
            strValues = ""
            Call ���ýӿ�_ָ����¼��_������("PersonInfo")
            Call ���ýӿ�_��ȡ����_������("icd", str����)
            strValues = strValues & "|" & str����
            Call ���ýӿ�_��ȡ����_������("disease", str����)
            strValues = strValues & "|" & str����
            
            '�ж��Ƿ���ڸò���
            gstrSQL = " Select ID From ���ղ���" & _
                      " Where ����=[1] And ����=[2]"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "�Ƿ���ڴ˲���", TYPE_������, str����)
            If rsTemp.RecordCount = 0 Then
                lngID = zlDatabase.GetNextID("���ղ���")
                gstrSQL = "zl_���ղ���_INSERT(" & lngID & "," & TYPE_������ & ",'" & str���� & _
                            "','" & str���� & "',NULL,0,NULL,NULL)"
            Else
                lngID = rsTemp!ID
            End If
            
            strValues = lngID & strValues
            Call Record_Add(mrs����, strColumns, strValues)
            Call DebugTool("�Ѽ���һ�м�¼")
        End If
        gcnOracle.CommitTrans
        
        '���ֻ��һ��������Ϣ��ֱ����ʾ����
        blnTrans = False
        If mrs����.RecordCount = 1 Then
            txt������Ϣ.Tag = mrs����!����
            txt������Ϣ.Text = "(" & mrs����!���� & ")" & mrs����!����
            lbl������Ϣ.Tag = txt������Ϣ.Text '���ڻָ���ʾ
        ElseIf mrs����.RecordCount > 1 Then
            If frmListSel.ShowSelect(TYPE_������, mrs����, "ID", "����ѡ��", "��ѡ��ҽ�����֣�") = True Then
                txt������Ϣ.Tag = mrs����!����
                txt������Ϣ.Text = "(" & mrs����!���� & ")" & mrs����!����
                lbl������Ϣ.Tag = txt������Ϣ.Text '���ڻָ���ʾ
                
                mrs����.Filter = 0
            End If
        End If
    Case Else   '����סԺ
        Call ���ýӿ�_��ȡ����_������("sum_year", strData)
        txtסԺ���� = strData
    End Select
    
    '--��ȡ�����ʻ�������סԺû�з��أ�ǿ��ȡһ�Σ�
    If Not ���ýӿ�_׼��_������(Function_������.����_�������) Then Exit Sub
    'д��ڲ���
'    1   fund_id    ������    3   ��
'    2   indi_id    ���˱��    8   ��
'    3   corp_ID    ��λ���    3
    'Modified By ���� ��������ɳ ԭ����Ҫ�ഫһ��������corp_id��
    gstrField_������ = "fund_id||indi_id||corp_id"
    gstrValue_������ = str�����ʻ� & "||" & txt���˱�� & "||" & str�α��˵�λ
    Call ���ýӿ�_д��ڲ���_������(1)
    If Not ���ýӿ�_ִ��_������ Then Exit Sub
    If Not ���ýӿ�_ָ����¼��_������("PersonAccount") Then Exit Sub
    Call ���ýӿ�_��ȡ����_������("last_balance", strData)
    txt�ʻ���� = Format(strData, "#####0.00;-#####0.00; ;")
    gCominfo_������.�ʻ���� = Val(strData)
    
    Exit Sub
errHand:
    If ErrCenter = 1 Then Resume
    If blnTrans Then gcnOracle.RollbackTrans
End Sub

Public Function NeedName(strList As String) As String
    NeedName = Mid(strList, InStr(strList, "-") + 1)
End Function

Public Sub CheckInputLen(txt As Object, KeyAscii As Integer)
    If KeyAscii = 13 Then KeyAscii = 0: Exit Sub
    If KeyAscii < 32 And KeyAscii >= 0 Then Exit Sub
    If txt.MaxLength = 0 Then Exit Sub
    If zlCommFun.ActualLen(txt.Text & Chr(KeyAscii)) > txt.MaxLength Then KeyAscii = 0
End Sub

Private Sub ReadFromInterface()
    Dim objTextBox As TextBox, objCons As Object
    Dim arrData, strData As String, strTemp As String
    
    For Each objCons In Controls
        If UCase(TypeName(objCons)) = "TEXTBOX" Then
            Set objTextBox = objCons
            If Trim(objTextBox.Tag) <> "" And objTextBox.Name <> "txt������Ϣ" Then
                arrData = Split(objTextBox.Tag, "|")
                If UBound(arrData) = 0 Then
                    Call ���ýӿ�_��ȡ����_������(arrData(0), strData)
                Else
                    Call ���ýӿ�_��ȡ����_������(arrData(0), strData)
                    Call ���ýӿ�_��ȡ����_������(arrData(1), strTemp)
                    If strData <> "" Then strData = "[" & strData & "]" & strTemp
                End If
                objTextBox.Text = strData
            End If
        End If
    Next
End Sub

Private Sub Record_Add(ByRef rsObj As ADODB.Recordset, ByVal strFields As String, ByVal strValues As String)
    Dim arrFields, arrValues, intField As Integer
    '��Ӽ�¼
    'strFields:�ֶ���|�ֶ���
    'strValues:ֵ|ֵ
    
    '���ӣ�
    'Dim strFields As String, strValues As String
    'strFields = "RecordID|��ĿID|ժҪ"
    'strValues = "5188|6666|��Ŀ����"
    'Call Record_Update(rsVoucher, strFields, strValues)

    arrFields = Split(strFields, "|")
    arrValues = Split(strValues, "|")
    intField = UBound(arrFields)
    If intField = 0 Then Exit Sub

    With rsObj
        .AddNew
        For intField = 0 To intField
            .Fields(arrFields(intField)).Value = IIf(UCase(arrValues(intField)) = "NULL", Null, arrValues(intField))
        Next
        .Update
    End With
End Sub

Private Sub Record_Init(ByRef rsObj As ADODB.Recordset, ByVal strFields As String)
    Dim arrFields, intField As Integer
    Dim strFieldName As String, intType As Integer, lngLength As Long
    '��ʼ��ӳ���¼��
    'strFields:�ֶ���,����,����|�ֶ���,����,����    �������Ϊ��,��ȡĬ�ϳ���
    '�ַ���:adLongVarChar;������:adDouble;������:adDBDate
    
    '���ӣ�
    'Dim rsVoucher As New ADODB.Recordset, strFields As String
    'strFields = "RecordID," & adDouble & ",18|��ĿID," & adDouble & ",18|ժҪ, " & adLongVarChar & ",50|" & _
    '"ɾ��," & adDouble & ",1"
    'Call Record_Init(rsVoucher, strFields)

    arrFields = Split(strFields, "|")
    Set rsObj = New ADODB.Recordset

    With rsObj
        If .State = 1 Then .Close
        For intField = 0 To UBound(arrFields)
            strFieldName = Split(arrFields(intField), ",")(0)
            intType = Split(arrFields(intField), ",")(1)
            lngLength = Split(arrFields(intField), ",")(2)

            '��ȡ�ֶ�ȱʡ����
            If lngLength = 0 Then
                Select Case intType
                Case adDouble
                    lngLength = madDoubleDefault
                Case adVarChar
                    lngLength = madLongVarCharDefault
                Case adLongVarChar
                    lngLength = madLongVarCharDefault
                Case Else
                    lngLength = madDbDateDefault
                End Select
            End If
            .Fields.Append strFieldName, intType, lngLength, adFldIsNullable
        Next
        
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
End Sub
