VERSION 5.00
Begin VB.Form frmIdentify�ش�У԰�� 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "�û�����Ϣ"
   ClientHeight    =   4680
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7305
   Icon            =   "frmIdentify�ش�У԰��.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4680
   ScaleWidth      =   7305
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmd�鿨 
      Caption         =   "�ض�(&R)"
      Height          =   350
      Left            =   240
      TabIndex        =   37
      Top             =   4185
      Width           =   1100
   End
   Begin VB.Frame fra 
      Height          =   75
      Index           =   1
      Left            =   -30
      TabIndex        =   41
      Top             =   630
      Width           =   8340
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   6090
      TabIndex        =   39
      Top             =   4185
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   4710
      TabIndex        =   38
      Top             =   4185
      Width           =   1100
   End
   Begin VB.Frame fra 
      Height          =   75
      Index           =   0
      Left            =   -15
      TabIndex        =   40
      Top             =   4020
      Width           =   8340
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "����"
      Height          =   180
      Index           =   10
      Left            =   5370
      TabIndex        =   47
      Top             =   2130
      Width           =   360
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Index           =   10
      Left            =   5790
      TabIndex        =   46
      Top             =   2070
      Width           =   1335
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "��������"
      Height          =   180
      Index           =   8
      Left            =   5010
      TabIndex        =   45
      Top             =   1725
      Width           =   720
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Index           =   8
      Left            =   5790
      TabIndex        =   44
      Top             =   1665
      Width           =   1335
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "�Ա�"
      Height          =   180
      Index           =   7
      Left            =   3045
      TabIndex        =   43
      Top             =   1725
      Width           =   360
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Index           =   7
      Left            =   3450
      TabIndex        =   42
      Top             =   1665
      Width           =   1335
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Index           =   20
      Left            =   990
      TabIndex        =   36
      ToolTipText     =   "�ս����ۼƽ��"
      Top             =   3645
      Width           =   1335
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "�ȴ�ʱ��"
      Height          =   180
      Index           =   20
      Left            =   225
      TabIndex        =   35
      Top             =   3705
      Width           =   720
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Index           =   13
      Left            =   5790
      TabIndex        =   34
      ToolTipText     =   "�ս����ۼƽ��"
      Top             =   2475
      Width           =   1335
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "���ۼƶ�"
      Height          =   180
      Index           =   13
      Left            =   5010
      TabIndex        =   33
      Top             =   2535
      Width           =   720
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Index           =   19
      Left            =   5790
      TabIndex        =   32
      ToolTipText     =   "�ϴν����ն˺�"
      Top             =   3255
      Width           =   1335
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "�ϴ��ն˺�"
      Height          =   180
      Index           =   19
      Left            =   4830
      TabIndex        =   31
      Top             =   3315
      Width           =   900
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Index           =   18
      Left            =   3450
      TabIndex        =   30
      ToolTipText     =   "�ϴν���ʱ��"
      Top             =   3255
      Width           =   1335
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "�ϴ�ʱ��"
      Height          =   180
      Index           =   18
      Left            =   2685
      TabIndex        =   29
      Top             =   3315
      Width           =   720
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Index           =   17
      Left            =   990
      TabIndex        =   28
      ToolTipText     =   "�ϴν��׽��"
      Top             =   3255
      Width           =   1365
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "�ϴ���ˮ��"
      Height          =   180
      Index           =   17
      Left            =   45
      TabIndex        =   27
      Top             =   3315
      Width           =   900
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "Ǯ��1���"
      Height          =   180
      Index           =   14
      Left            =   135
      TabIndex        =   21
      Top             =   2925
      Width           =   810
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Index           =   14
      Left            =   990
      TabIndex        =   22
      Top             =   2865
      Width           =   1365
   End
   Begin VB.Label lblInfor 
      AutoSize        =   -1  'True
      Caption         =   "ͨ��IC����֤��Ա��ݣ�������֤�����Ϣ��ʾ������"
      Height          =   180
      Left            =   600
      TabIndex        =   0
      Top             =   390
      Width           =   4320
   End
   Begin VB.Image img 
      Height          =   480
      Left            =   30
      Picture         =   "frmIdentify�ش�У԰��.frx":000C
      Top             =   120
      Width           =   480
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "�ϴ���ˮ��"
      Height          =   180
      Index           =   16
      Left            =   4830
      TabIndex        =   25
      Top             =   2925
      Width           =   900
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Index           =   16
      Left            =   5790
      TabIndex        =   26
      ToolTipText     =   "�ϴν�����ˮ"
      Top             =   2865
      Width           =   1335
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Index           =   15
      Left            =   3450
      TabIndex        =   24
      Top             =   2865
      Width           =   1335
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "Ǯ��2���"
      Height          =   180
      Index           =   15
      Left            =   2595
      TabIndex        =   23
      Top             =   2925
      Width           =   810
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Index           =   9
      Left            =   990
      TabIndex        =   20
      Top             =   2070
      Width           =   3795
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "���֤��"
      Height          =   180
      Index           =   9
      Left            =   225
      TabIndex        =   19
      Top             =   2130
      Width           =   720
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Index           =   12
      Left            =   3450
      TabIndex        =   18
      Top             =   2475
      Width           =   1335
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "�����к�"
      Height          =   180
      Index           =   12
      Left            =   2685
      TabIndex        =   17
      Top             =   2535
      Width           =   720
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Index           =   11
      Left            =   990
      TabIndex        =   16
      Top             =   2475
      Width           =   1335
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "ע������"
      Height          =   180
      Index           =   11
      Left            =   225
      TabIndex        =   15
      Top             =   2535
      Width           =   720
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Index           =   6
      Left            =   990
      TabIndex        =   14
      Top             =   1665
      Width           =   1365
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Index           =   1
      Left            =   3450
      TabIndex        =   4
      Top             =   900
      Width           =   1335
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "����"
      Height          =   180
      Index           =   6
      Left            =   585
      TabIndex        =   13
      Top             =   1725
      Width           =   360
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Index           =   5
      Left            =   5790
      TabIndex        =   12
      Top             =   1290
      Width           =   1335
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "����Ч��"
      Height          =   180
      Index           =   5
      Left            =   5010
      TabIndex        =   11
      Top             =   1350
      Width           =   720
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Index           =   4
      Left            =   3450
      TabIndex        =   10
      Top             =   1290
      Width           =   1335
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "�����ѷ���"
      Height          =   180
      Index           =   4
      Left            =   2505
      TabIndex        =   9
      Top             =   1350
      Width           =   900
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Index           =   3
      Left            =   990
      TabIndex        =   8
      Top             =   1290
      Width           =   1365
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "У԰����"
      Height          =   180
      Index           =   3
      Left            =   225
      TabIndex        =   7
      Top             =   1350
      Width           =   720
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Index           =   2
      Left            =   5790
      TabIndex        =   6
      Top             =   900
      Width           =   1335
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "����ˮ��"
      Height          =   180
      Index           =   2
      Left            =   5010
      TabIndex        =   5
      Top             =   960
      Width           =   720
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "������"
      Height          =   180
      Index           =   1
      Left            =   2865
      TabIndex        =   3
      Top             =   960
      Width           =   540
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Index           =   0
      Left            =   990
      TabIndex        =   2
      Top             =   900
      Width           =   1365
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "�����豸��"
      Height          =   180
      Index           =   0
      Left            =   45
      TabIndex        =   1
      Top             =   960
      Width           =   900
   End
End
Attribute VB_Name = "frmIdentify�ش�У԰��"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mblnFirst  As Boolean
Dim mstrReturn As String    '������Ϣ��
Dim mlng����ID As Long
'mbytType-ʶ�����ͣ�0-�����շѣ�1-��Ժ�Ǽǣ�2-������������סԺ,3-�Һ�,4-����
Dim mbytType As Byte
Dim mblnOK As Boolean
Dim mbln�Զ��Һ� As Boolean
Private Function IsValid() As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:������֤
    '--�����:
    '--������:
    '--��  ��:��֤�ɹ�����true,���򷵻�False
    '-----------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    
    IsValid = False
    
    '��鲡��״̬
    Dim lng����ID As Long
    gstrSQL = "select ����id,nvl(��ǰ״̬,0) as ״̬ from �����ʻ� where ����=[1] and ҽ����=[2]"
    
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, TYPE_�ش�У԰��, g�������_�ش�У԰��.����)
    If mbytType <> 4 Then   '����סԺ����ʱ������֤��ǰ״̬
        If rsTemp.RecordCount > 0 Then
            If rsTemp("״̬") > 0 Then
                MsgBox "�ò����Ѿ���Ժ������ͨ�������֤��", vbInformation, gstrSysName
                Exit Function
            End If
        End If
    Else
        'סԺ����ʱ,�账���Ƿ�Ϊͬһ����
        If rsTemp.EOF Then
            ShowMsgbox "�ڱ����ʻ��в����ڵ�ǰ����!"
            Exit Function
        Else
            lng����ID = Nvl(rsTemp!����ID, 0)
            If mlng����ID <> lng����ID Then
                ShowMsgbox "��ٽ��ʵĵ�ǰ�����������֤�Ĳ��˲�һ��!"
                Exit Function
            End If
        End If
    End If
    IsValid = True
End Function

Private Sub cmdCancel_Click()
    mstrReturn = ""
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim strTmp As String
    Dim strTmp1 As String
    
    '��֤����
    If IsValid = False Then Exit Sub
    
    'ȷ����ط��ش�
    
    '0����;1ҽ����;2����;3����;4�Ա�;5��������;6���֤;7��λ����(����);
    '8.���Ĵ���;9.˳���;10��Ա���;11�ʻ����;12��ǰ״̬;13����ID;14��ְ(0,1);15����֤��;16�����;17�Ҷȼ�
    '18�ʻ������ۼ�,19�ʻ�֧���ۼ�,20����ͳ���ۼ�,21ͳ�ﱨ���ۼ�,22סԺ�����ۼ�

    '0����;1ҽ����;2����;3����;4�Ա�;5��������;6���֤;7��λ����(����);8������ID��
    '9����;10.˳���;11��Ա���;12�ʻ����;13��ǰ״̬;14����ID;15��ְ(0,1);16����֤��;17�����;18�Ҷȼ�
    '19�ʻ������ۼ�,20�ʻ�֧���ۼ�,21����ͳ���ۼ�,22ͳ�ﱨ���ۼ�,23סԺ�����ۼ�;24�������� (1����������);25��������
    
    
    mstrReturn = ""
    With g�������_�ش�У԰��
        strTmp = .����                         '0����
        strTmp = strTmp & ";" & .����ˮ��    '1ҽ����
        strTmp = strTmp & ";"               '2����
        strTmp = strTmp & ";" & .����       '3����
        strTmp = strTmp & ";" & .�Ա�       '4�Ա�
        strTmp = strTmp & ";" & .��������   '5��������
        strTmp = strTmp & ";" & .���֤��   '6���֤
        strTmp = strTmp & ";" & .������   '7��λ����(����)
        
        strTmp1 = ""
        strTmp1 = strTmp1 & ";"    '8���Ĵ���
        strTmp1 = strTmp1 & ";" & .�����ѷ���    '9˳���
        strTmp1 = strTmp1 & ";"       '10��Ա���,�����ת�ﵥ��
        strTmp1 = strTmp1 & ";" & (.����Ǯ��1��� + .����Ǯ��2���) / 100     '11�ʻ����
        strTmp1 = strTmp1 & ";0"               '12��ǰ״̬
        strTmp1 = strTmp1 & ";"               '13����ID
        strTmp1 = strTmp1 & ";"   '.�������  '14��ְ(0,1)
        strTmp1 = strTmp1 & ";"   '15����֤��,Ŀǰ�Ҵ���ǲ��������ʻ����
        strTmp1 = strTmp1 & ";" & IIf(.���� = 0, "", .����) '16�����
        strTmp1 = strTmp1 & ";"     '17�Ҷȼ�,��ľ���������
        strTmp1 = strTmp1 & ";"         '18�ʻ������ۼ�
        strTmp1 = strTmp1 & ";"        '19�ʻ�֧���ۼ�
        strTmp1 = strTmp1 & ";"  '20����ͳ���ۼ�
        strTmp1 = strTmp1 & ";"  '21ͳ�ﱨ���ۼ�
        strTmp1 = strTmp1 & ";"        '22סԺ�����ۼ�
        
    End With
    
    mlng����ID = BuildPatiInfo(0, strTmp & strTmp1, mlng����ID, TYPE_�ش�У԰��)
    
    '���ظ�ʽ:�м���벡��ID
    If mlng����ID > 0 Then
        mstrReturn = strTmp & ";" & mlng����ID & strTmp1
    Else
        Unload Me
    End If
    '�洢У԰����Ϣ
    '����:zl_У԰����Ϣ_Insert(����ID_IN,����_IN,����_IN,����ˮ��_IN,����_IN,�����ѷ���_IN,����Ч��_IN,����_IN,
    '   ע������_IN,���֤��_IN,����Ǯ��1���_IN,����Ǯ��2���_IN,���������к�_IN,�ϴν�����ˮ��_IN,�ϴν��׽��_IN
    '   �ϴν���ʱ��_IN,�ϴν����ն˺�_IN,���ȴ�ʱ��_IN,�ս����ۼƽ��_IN
    
    
    strTmp = "zl_У԰����Ϣ_Insert("
    With g�������_�ش�У԰��
        strTmp = strTmp & _
        mlng����ID & "," & _
        TYPE_�ش�У԰�� & "," & _
        0 & "," & _
        .����ˮ�� & "," & _
        .���� & "," & _
        .�����ѷ��� & "," & _
        IIf(.����Ч�� = "", "NULL", "'" & .����Ч�� & "'") & "," & _
        IIf(.���� = "", "NULL", "'" & .���� & "'") & "," & _
        IIf(.ע������ = "", "NULL", "'" & .ע������ & "'") & "," & _
        IIf(.���֤�� = "", "NULL", "'" & .���֤�� & "'") & "," & _
        .����Ǯ��1��� & "," & _
        .����Ǯ��2��� & "," & _
        .��������� & "," & _
        .�ϴν�����ˮ�� & "," & _
        .�ϴν��׽�� & "," & _
        .�ϴν���ʱ�� & "," & _
        .�ϴν����ն˺� & "," & _
        .���ȴ�ʱ�� & "," & _
        .�ս����ۼƽ�� & ")"
    End With
    Err = 0
    On Error GoTo errHand:
    zlDatabase.ExecuteProcedure strTmp, Me.Caption
     Unload Me
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub cmd�鿨_Click()
    cmd�鿨.Enabled = False
    SetCtlEn False
    mblnOK = ReadCard
    If mblnOK And mbln�Զ��Һ� Then
        cmdOK_Click
        Unload Me
        Exit Sub
    End If
    SetCtlEn True
    cmd�鿨.Enabled = True
End Sub
Private Sub SetCtlEn(ByVal blnTrue As Boolean)
    cmd�鿨.Enabled = blnTrue
    cmdOK.Enabled = blnTrue And mblnOK
    cmdCancel.Enabled = blnTrue
End Sub
Private Sub Form_Activate()
    
    If mblnFirst = False Then Exit Sub
    mblnFirst = False
    mblnOK = False
    '����ʱ����
    cmd�鿨_Click
End Sub
Private Function ReadCard() As Boolean
    Dim int�Ա� As Integer
    Dim i As Integer
    ReadCard = False
   
   '��֤�û����
    If GetUserCardInfor() = False Then
        For i = 0 To 20
            lblEdit(i) = ""
        Next
        Exit Function
    End If
    
    Err = 0
    On Error Resume Next
    '������������Ϣ��ֵ
    With g�������_�ش�У԰��
        lblEdit(0) = .�����豸
        lblEdit(1) = .������
        lblEdit(2) = .����ˮ��
        lblEdit(3) = .����
        lblEdit(4) = .�����ѷ���
        lblEdit(5) = .����Ч��
        lblEdit(6) = .����
        lblEdit(7) = .�Ա�
        lblEdit(8) = .��������
        lblEdit(9) = .���֤��
        lblEdit(10) = .����
        lblEdit(11) = .ע������
        lblEdit(12) = .���������
        lblEdit(13) = .�ս����ۼƽ�� / 100
        lblEdit(14) = .����Ǯ��1��� / 100
        lblEdit(15) = .����Ǯ��2��� / 100
        lblEdit(16) = .�ϴν�����ˮ��
        lblEdit(17) = .�ϴν��׽�� / 100
        lblEdit(18) = .�ϴν���ʱ��
        lblEdit(19) = .�ϴν����ն˺�
        lblEdit(20) = .���ȴ�ʱ��
    End With
    ReadCard = True
End Function

Public Function GetPatient(ByVal bytType As Byte, Optional ByVal lng����ID As Long = 0, Optional bln�Զ��Һ� As Boolean = False) As String
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:��ȡ���˵������Ϣ
    '--�����:bytType-����(mbytType-ʶ�����ͣ�0-�����շѣ�1-��Ժ�Ǽǣ�2-������������סԺ,3-�Һ�,4-����)
    '         lng����ID-����ID
    '--������:
    '--��  ��:
    '-----------------------------------------------------------------------------------------------------------
    mstrReturn = ""
    mlng����ID = lng����ID
    mbytType = bytType
    mbln�Զ��Һ� = bln�Զ��Һ�
    Me.Show 1
    GetPatient = mstrReturn
End Function

Private Sub Form_Load()
    mblnFirst = True
End Sub



