VERSION 5.00
Begin VB.Form frmIdentify��ɽ 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "�����֤"
   ClientHeight    =   5085
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8235
   Icon            =   "frmIdentify��ɽ.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5085
   ScaleWidth      =   8235
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.ComboBox Cmb���� 
      Height          =   300
      ItemData        =   "frmIdentify��ɽ.frx":000C
      Left            =   3240
      List            =   "frmIdentify��ɽ.frx":0034
      TabIndex        =   44
      Text            =   "�б���"
      Top             =   4650
      Width           =   1500
   End
   Begin VB.CommandButton cmdRead 
      Caption         =   "����(&R)"
      Height          =   350
      Left            =   210
      TabIndex        =   41
      Top             =   4620
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   6720
      TabIndex        =   43
      Top             =   4620
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Enabled         =   0   'False
      Height          =   350
      Left            =   5490
      TabIndex        =   42
      Top             =   4620
      Width           =   1100
   End
   Begin VB.Frame Frame1 
      Caption         =   "������Ϣ"
      Enabled         =   0   'False
      Height          =   4335
      Left            =   150
      TabIndex        =   0
      Top             =   120
      Width           =   7935
      Begin VB.TextBox txtInfo 
         Height          =   300
         Index           =   19
         Left            =   5310
         TabIndex        =   40
         Top             =   3840
         Width           =   2175
      End
      Begin VB.TextBox txtInfo 
         Height          =   300
         Index           =   18
         Left            =   5310
         TabIndex        =   38
         Top             =   3450
         Width           =   2175
      End
      Begin VB.TextBox txtInfo 
         Height          =   300
         Index           =   17
         Left            =   5310
         TabIndex        =   36
         Top             =   3060
         Width           =   2175
      End
      Begin VB.TextBox txtInfo 
         Height          =   300
         Index           =   16
         Left            =   5310
         TabIndex        =   34
         Top             =   2670
         Width           =   2175
      End
      Begin VB.TextBox txtInfo 
         Height          =   300
         Index           =   15
         Left            =   5310
         TabIndex        =   32
         Top             =   2280
         Width           =   2175
      End
      Begin VB.TextBox txtInfo 
         Height          =   300
         Index           =   14
         Left            =   5310
         TabIndex        =   30
         Top             =   1890
         Width           =   2175
      End
      Begin VB.TextBox txtInfo 
         Height          =   300
         Index           =   13
         Left            =   5310
         TabIndex        =   28
         Top             =   1500
         Width           =   2175
      End
      Begin VB.TextBox txtInfo 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   12
         Left            =   5310
         PasswordChar    =   "*"
         TabIndex        =   26
         Top             =   1110
         Width           =   2175
      End
      Begin VB.TextBox txtInfo 
         Height          =   300
         Index           =   11
         Left            =   5310
         TabIndex        =   24
         Top             =   720
         Width           =   2175
      End
      Begin VB.TextBox txtInfo 
         Height          =   300
         Index           =   10
         Left            =   5310
         TabIndex        =   22
         Top             =   330
         Width           =   2175
      End
      Begin VB.TextBox txtInfo 
         Height          =   300
         Index           =   9
         Left            =   1500
         TabIndex        =   20
         Top             =   3840
         Width           =   2175
      End
      Begin VB.TextBox txtInfo 
         Height          =   300
         Index           =   8
         Left            =   1500
         TabIndex        =   18
         Top             =   3450
         Width           =   2175
      End
      Begin VB.TextBox txtInfo 
         Height          =   300
         Index           =   7
         Left            =   1500
         TabIndex        =   16
         Top             =   3060
         Width           =   2175
      End
      Begin VB.TextBox txtInfo 
         Height          =   300
         Index           =   6
         Left            =   1500
         TabIndex        =   14
         Top             =   2670
         Width           =   2175
      End
      Begin VB.TextBox txtInfo 
         Height          =   300
         Index           =   5
         Left            =   1500
         TabIndex        =   12
         Top             =   2280
         Width           =   2175
      End
      Begin VB.TextBox txtInfo 
         Height          =   300
         Index           =   4
         Left            =   1500
         TabIndex        =   10
         Top             =   1890
         Width           =   2175
      End
      Begin VB.TextBox txtInfo 
         Height          =   300
         Index           =   3
         Left            =   1500
         TabIndex        =   8
         Top             =   1500
         Width           =   2175
      End
      Begin VB.TextBox txtInfo 
         Height          =   300
         Index           =   2
         Left            =   1500
         TabIndex        =   6
         Top             =   1110
         Width           =   2175
      End
      Begin VB.TextBox txtInfo 
         Height          =   300
         Index           =   1
         Left            =   1500
         TabIndex        =   4
         Top             =   720
         Width           =   2175
      End
      Begin VB.TextBox txtInfo 
         Height          =   300
         Index           =   0
         Left            =   1500
         TabIndex        =   2
         Top             =   330
         Width           =   2175
      End
      Begin VB.Label lblInfo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "סԺ����"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   19
         Left            =   4530
         TabIndex        =   39
         Top             =   3900
         Width           =   720
      End
      Begin VB.Label lblInfo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "�˻����"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   18
         Left            =   4530
         TabIndex        =   37
         Top             =   3510
         Width           =   720
      End
      Begin VB.Label lblInfo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "����סԺ�ۼ�"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   17
         Left            =   4170
         TabIndex        =   35
         Top             =   3120
         Width           =   1080
      End
      Begin VB.Label lblInfo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "����߶��ۼ�"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   16
         Left            =   4170
         TabIndex        =   33
         Top             =   2730
         Width           =   1080
      End
      Begin VB.Label lblInfo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "ͳ��֧���ۼ�"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   15
         Left            =   4170
         TabIndex        =   31
         Top             =   2340
         Width           =   1080
      End
      Begin VB.Label lblInfo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "����ͳ��֧����"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   14
         Left            =   3990
         TabIndex        =   29
         Top             =   1950
         Width           =   1260
      End
      Begin VB.Label lblInfo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "�����ز�֧����"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   13
         Left            =   3990
         TabIndex        =   27
         Top             =   1560
         Width           =   1260
      End
      Begin VB.Label lblInfo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "�˻�����"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   12
         Left            =   4530
         TabIndex        =   25
         Top             =   1170
         Width           =   720
      End
      Begin VB.Label lblInfo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "�˻�״̬"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   11
         Left            =   4530
         TabIndex        =   23
         Top             =   780
         Width           =   720
      End
      Begin VB.Label lblInfo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "����ҽ������"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   10
         Left            =   4170
         TabIndex        =   21
         Top             =   390
         Width           =   1080
      End
      Begin VB.Label lblInfo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "����Ա���"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   9
         Left            =   540
         TabIndex        =   19
         Top             =   3900
         Width           =   900
      End
      Begin VB.Label lblInfo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "��λ����"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   8
         Left            =   720
         TabIndex        =   17
         Top             =   3510
         Width           =   720
      End
      Begin VB.Label lblInfo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "��λ����"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   7
         Left            =   720
         TabIndex        =   15
         Top             =   3120
         Width           =   720
      End
      Begin VB.Label lblInfo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "��Ա���"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   6
         Left            =   720
         TabIndex        =   13
         Top             =   2730
         Width           =   720
      End
      Begin VB.Label lblInfo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "��Ա״̬"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   5
         Left            =   720
         TabIndex        =   11
         Top             =   2340
         Width           =   720
      End
      Begin VB.Label lblInfo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "���֤��"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   4
         Left            =   720
         TabIndex        =   9
         Top             =   1950
         Width           =   720
      End
      Begin VB.Label lblInfo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "�Ա�"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   3
         Left            =   1080
         TabIndex        =   7
         Top             =   1560
         Width           =   360
      End
      Begin VB.Label lblInfo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   2
         Left            =   1080
         TabIndex        =   5
         Top             =   1170
         Width           =   360
      End
      Begin VB.Label lblInfo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   1
         Left            =   1080
         TabIndex        =   3
         Top             =   780
         Width           =   360
      End
      Begin VB.Label lblInfo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "�α�ID��"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   0
         Left            =   720
         TabIndex        =   1
         Top             =   390
         Width           =   720
      End
   End
   Begin VB.Label Lbl���� 
      AutoSize        =   -1  'True
      Caption         =   "��ѡ�������"
      Height          =   180
      Left            =   2040
      TabIndex        =   45
      Top             =   4725
      Width           =   1080
   End
End
Attribute VB_Name = "frmIdentify��ɽ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrReturn As String
Private mbytType As Byte
Private mlng����ID As Long

Public Function GetPatient(ByVal bytType As Byte, Optional ByVal lng����ID As Long = 0) As String
    mstrReturn = ""
    mbytType = bytType
    mlng����ID = lng����ID
    Me.Show 1
    
    GetPatient = mstrReturn
End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Const intҽ���� As Integer = 0
    Const int���� As Integer = 1
    Const int���� As Integer = 2
    Const int�Ա� As Integer = 3
    Const int���֤�� As Integer = 4
    Const int��λ���� As Integer = 7
    Const int��λ���� As Integer = 8
    Const int���� As Integer = 12
    Const int�ʻ���� As Integer = 18
    Const intסԺ���� As Integer = 19
    Dim str�������� As String
    Dim strIdentify As String, strAddition As String
    Dim rsTemp As New ADODB.Recordset
    Dim str�·� As String, str���� As String
    
    If Trim(txtInfo(intҽ����).Text) = "" Then
        MsgBox "δ�õ��α����˵�ҽ��ID�ţ��޷�������", vbInformation, gstrSysName
        cmdRead.SetFocus
        Exit Sub
    End If
    
    '������(2005-10-08)  ��鲡��״̬,������ɽ��ҽ����ͬ���صĲα�ID������ͬ�����Ը��ݲα�ID����ѡ�����������Ψһ�ж�
    gstrSQL = "select nvl(��ǰ״̬,0) as ״̬ from �����ʻ� where ����=[1] and ҽ����=[2] And ����=[3]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, gintInsure, CStr(txtInfo(intҽ����).Text), CInt(Cmb����.ListIndex))
    
    If rsTemp.RecordCount > 0 Then
        If rsTemp("״̬") > 0 Then
            MsgBox "�ò����Ѿ���Ժ������ͨ�������֤��", vbInformation, gstrSysName
            Exit Sub
        End If
    End If
    
    '��׼������
    If Len(txtInfo(int���֤��).Text) = 15 Then
        str�������� = Mid(txtInfo(int���֤��).Text, 7, 6)
        If Mid(str��������, 1, 2) <= "05" Then
            str�������� = "20" & str��������
        Else
            str�������� = "19" & str��������
        End If
    Else
        str�������� = Mid(txtInfo(int���֤��).Text, 7, 8)
    End If
    
    If Mid(str��������, 5, 2) < 1 Or Mid(str��������, 5, 2) > 12 Then
       str�·� = Frm��ɽ_��ʾ.�������ڸ���_��ɽ(1, Mid(str��������, 5, 2))
    Else
       str�·� = Mid(str��������, 5, 2)
    End If
    If Mid(str��������, 7, 2) < 1 Or Mid(str��������, 7, 2) > 31 Then
       str���� = Frm��ɽ_��ʾ.�������ڸ���_��ɽ(2, Mid(str��������, 7, 2))
    Else
       str���� = Mid(str��������, 7, 2)
    End If
    If Mid(str��������, 5, 2) = 2 And Mid(str��������, 7, 2) > 28 Then
       str�·� = "2"
       str���� = "28"
    End If
    str�������� = Mid(str��������, 1, 4) & "-" & str�·� & "-" & str����
    
    '�����ַ���
    '�������˵�����Ϣ�������ʽ��
    '0����;1ҽ����;2����;3����;4�Ա�;5��������;6���֤;7��λ����(����);
    '8.���Ĵ���;9.˳���;10��Ա���;11�ʻ����;12��ǰ״̬;13����ID;14��ְ(0,1);15����֤��;16�����;17�Ҷȼ�
    '18�ʻ������ۼ�,19�ʻ�֧���ۼ�,20����ͳ���ۼ�,21ͳ�ﱨ���ۼ�,22סԺ�����ۼ�
    strIdentify = txtInfo(int����).Text                         '0����
    strIdentify = strIdentify & ";" & txtInfo(intҽ����).Text   '1ҽ���ţ����˱�ţ�
    strIdentify = strIdentify & ";" & txtInfo(int����).Text     '2����
    strIdentify = strIdentify & ";" & txtInfo(int����).Text     '3����
    strIdentify = strIdentify & ";" & txtInfo(int�Ա�).Text     '4�Ա�
    strIdentify = strIdentify & ";" & str��������               '5��������
    strIdentify = strIdentify & ";" & txtInfo(int���֤��).Text '6���֤
    strIdentify = strIdentify & ";" & txtInfo(int��λ����).Text & "(" & txtInfo(int��λ����).Text & ")"          '7.��λ����(����)
    '������(2005-10-08)�޸�
    strAddition = ";" & Cmb����.ListIndex                       '8.���Ĵ���
    strAddition = strAddition & ";"                             '9.˳���
    strAddition = strAddition & ";"                             '10��Ա���
    strAddition = strAddition & ";" & Val(txtInfo(int�ʻ����).Text)      '11�ʻ����
    strAddition = strAddition & ";0"                            '12��ǰ״̬
    strAddition = strAddition & ";"                             '13����ID
    strAddition = strAddition & ";1"                            '14��ְ(1,2,3)
    strAddition = strAddition & ";"                             '15����֤��
    strAddition = strAddition & ";"                             '16�����
    strAddition = strAddition & ";"                             '17�Ҷȼ�
    strAddition = strAddition & ";" & Val(txtInfo(int�ʻ����).Text)     '18�ʻ������ۼ�
    strAddition = strAddition & ";0"                            '19�ʻ�֧���ۼ�
    strAddition = strAddition & ";0"                            '20���깤���ܶ�
    strAddition = strAddition & ";" & Val(txtInfo(intסԺ����).Text)  '21סԺ�����ۼ�
    
    Call DebugTool(strIdentify & strAddition)
    mlng����ID = BuildPatiInfo(0, strIdentify & strAddition, mlng����ID, TYPE_��ɽ)
    '���ظ�ʽ:�м���벡��ID
    If mlng����ID > 0 Then
        mstrReturn = strIdentify & ";" & mlng����ID & strAddition
    End If
    
    Unload Me
End Sub

Private Sub cmdRead_Click()
    gbytReturn_��ɽ = LS_GetPersonInfo(gPersonInfo_��ɽ)
    If GetErrInfo_��ɽ Then Exit Sub
    
    With gPersonInfo_��ɽ
        txtInfo(0).Text = .PSN_ID              ' As Integer      'ҽ�Ʋα�ID��
'        txtInfo(1).Text = .PSN_No             ' As Integer      '�α��˱���
        txtInfo(2).Text = .PSN_NAME            ' As String * 100 '�α�������
        txtInfo(3).Text = .Sex                 ' As String * 100 '�Ա�
        txtInfo(4).Text = .IDCARD              ' As String * 100 '���֤����
        txtInfo(5).Text = .PSN_STS             ' As String * 100 '�α���״̬
        txtInfo(6).Text = .PSN_TYP             ' As String * 100 '��Ա���
        txtInfo(7).Text = .UNIT_CODE           ' As String * 100 '��λ����
        txtInfo(8).Text = .UNIT_NAME           ' As String * 100 '��λ����
        txtInfo(9).Text = .OFFICAL_TYP         ' As String * 100 '����Ա���
        txtInfo(10).Text = .HAI_TYP            ' As String * 100 '����ҽ������
        txtInfo(11).Text = .ACCT_STS           ' As String * 100 'ҽ���˻�״̬
        txtInfo(12).Text = .HI_ACCT_PWD        ' As String * 100 'ҽ���ʻ�����
        txtInfo(13).Text = .SILL_PAY_AMT_TOTAL ' As Single       '���ڽ����������⼲��֧�����
        txtInfo(14).Text = .SILL_YR_FUND_AMT   ' As Single       '��������ͳ�����֧�����
        txtInfo(15).Text = .YR_FUND_AMT        ' As Single       '����ͳ�����֧�����
        txtInfo(16).Text = .HAI_YR_HIGH_AMT    ' As Single       '���ڲ���߶�֧�����
        txtInfo(17).Text = .HAI_YR_INBED_AMT   ' As Single       '���ڲ���סԺ����֧�����
        txtInfo(18).Text = .GZ_CUR_AMT         ' As Single       '�����˻����
        txtInfo(19).Text = .YR_INBED_CNT       ' As Integer      '����סԺ����
        txtInfo(1).Text = .CARD_NO
    End With
    cmdOK.Enabled = True
    cmdOK.SetFocus
End Sub

Private Sub Form_Load()
    If mlng����ID <> 0 Then Call ReadPatient
End Sub

Private Sub ClearCons()
    Dim intClear As Integer, intCOUNT As Integer
    '�����������
    
    intCOUNT = txtInfo.UBound - 1
    For intClear = 0 To intCOUNT
        txtInfo(intClear).Text = ""
    Next
End Sub

Private Sub ReadPatient()
    '
End Sub

Private Sub WriteFace()
    '
End Sub
