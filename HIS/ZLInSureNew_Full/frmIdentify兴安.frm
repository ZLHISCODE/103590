VERSION 5.00
Begin VB.Form frmIdentify�˰� 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "�����֤"
   ClientHeight    =   5145
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7080
   Icon            =   "frmIdentify�˰�.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5145
   ScaleWidth      =   7080
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.TextBox TxtEdit 
      Height          =   300
      Index           =   1
      Left            =   4635
      MaxLength       =   20
      TabIndex        =   3
      Top             =   945
      Width           =   2295
   End
   Begin VB.ComboBox cbo�������� 
      Height          =   300
      Left            =   4635
      Style           =   2  'Dropdown List
      TabIndex        =   19
      Top             =   2527
      Width           =   2295
   End
   Begin VB.CommandButton cmdȡ�� 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   5730
      TabIndex        =   33
      Top             =   4605
      Width           =   1100
   End
   Begin VB.CommandButton cmdȷ�� 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   4440
      TabIndex        =   32
      Top             =   4635
      Width           =   1100
   End
   Begin VB.Frame fra 
      Height          =   75
      Index           =   1
      Left            =   -75
      TabIndex        =   35
      Top             =   585
      Width           =   8340
   End
   Begin VB.Frame fra 
      Height          =   75
      Index           =   0
      Left            =   0
      TabIndex        =   34
      Top             =   4380
      Width           =   8340
   End
   Begin VB.TextBox TxtEdit 
      Height          =   285
      Index           =   0
      Left            =   1005
      MaxLength       =   20
      TabIndex        =   1
      Top             =   953
      Width           =   2265
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Index           =   15
      Left            =   4635
      TabIndex        =   31
      Top             =   4050
      Width           =   2325
   End
   Begin VB.Label lbl 
      Caption         =   "����סԺ����"
      Height          =   180
      Index           =   15
      Left            =   3525
      TabIndex        =   30
      Top             =   4110
      Width           =   1080
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   14
      Left            =   1365
      TabIndex        =   29
      Top             =   4065
      Width           =   1905
   End
   Begin VB.Label lbl 
      Caption         =   "��������ͳ��"
      Height          =   180
      Index           =   14
      Left            =   225
      TabIndex        =   28
      Top             =   4110
      Width           =   1080
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Index           =   13
      Left            =   4635
      TabIndex        =   27
      Top             =   3675
      Width           =   2325
   End
   Begin VB.Label lbl 
      Caption         =   "���������Ը���"
      Height          =   180
      Index           =   13
      Left            =   3360
      TabIndex        =   26
      Top             =   3735
      Width           =   1260
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   12
      Left            =   1365
      TabIndex        =   25
      Top             =   3690
      Width           =   1905
   End
   Begin VB.Label lbl 
      Caption         =   "��������ҩƷ"
      Height          =   180
      Index           =   12
      Left            =   240
      TabIndex        =   24
      Top             =   3735
      Width           =   1080
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "��������"
      Height          =   180
      Index           =   11
      Left            =   3885
      TabIndex        =   18
      Top             =   2587
      Width           =   720
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "��������"
      Height          =   180
      Index           =   9
      Left            =   255
      TabIndex        =   20
      Top             =   2962
      Width           =   720
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   9
      Left            =   1005
      TabIndex        =   21
      Top             =   2910
      Width           =   5925
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "���ֱ���"
      Height          =   180
      Index           =   8
      Left            =   255
      TabIndex        =   16
      Top             =   2587
      Width           =   720
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   8
      Left            =   1005
      TabIndex        =   17
      Top             =   2535
      Width           =   2265
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "�ʻ����"
      Height          =   180
      Index           =   7
      Left            =   3885
      TabIndex        =   14
      Top             =   2182
      Width           =   720
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Index           =   7
      Left            =   4635
      TabIndex        =   15
      Top             =   2122
      Width           =   2295
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "�ʻ�״̬"
      Height          =   180
      Index           =   6
      Left            =   255
      TabIndex        =   12
      Top             =   2182
      Width           =   720
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   6
      Left            =   1005
      TabIndex        =   13
      Top             =   2130
      Width           =   2265
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "��Ա���"
      Height          =   180
      Index           =   5
      Left            =   3885
      TabIndex        =   10
      Top             =   1755
      Width           =   720
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Index           =   5
      Left            =   4635
      TabIndex        =   11
      Top             =   1695
      Width           =   2295
   End
   Begin VB.Label lblInfor 
      AutoSize        =   -1  'True
      Caption         =   "ͨ���̿���֤��Ա��ݣ�������֤�����Ϣ��ʾ������"
      Height          =   180
      Left            =   675
      TabIndex        =   36
      Top             =   345
      Width           =   4320
   End
   Begin VB.Image img 
      Height          =   480
      Left            =   105
      Picture         =   "frmIdentify�˰�.frx":030A
      Top             =   75
      Width           =   480
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "ҽ������"
      Height          =   180
      Index           =   0
      Left            =   255
      TabIndex        =   0
      Top             =   1005
      Width           =   720
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "ҽ��֤��"
      Height          =   180
      Index           =   1
      Left            =   3885
      TabIndex        =   2
      Top             =   1005
      Width           =   720
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "����"
      Height          =   180
      Index           =   2
      Left            =   615
      TabIndex        =   4
      Top             =   1387
      Width           =   360
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "�Ա�"
      Height          =   180
      Index           =   3
      Left            =   4245
      TabIndex        =   6
      Top             =   1387
      Width           =   360
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "����"
      Height          =   180
      Index           =   4
      Left            =   615
      TabIndex        =   8
      Top             =   1755
      Width           =   360
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "��λ����"
      Height          =   180
      Index           =   10
      Left            =   255
      TabIndex        =   22
      Top             =   3352
      Width           =   720
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   2
      Left            =   1005
      TabIndex        =   5
      Top             =   1335
      Width           =   2265
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   3
      Left            =   4635
      TabIndex        =   7
      Top             =   1335
      Width           =   975
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   4
      Left            =   1005
      TabIndex        =   9
      Top             =   1703
      Width           =   1455
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   10
      Left            =   1005
      TabIndex        =   23
      Top             =   3300
      Width           =   5940
   End
End
Attribute VB_Name = "frmIdentify�˰�"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mbytType As Byte            '0-�����շѣ�1-��Ժ�Ǽǣ�2-������������סԺ,3-�Һ�,4-����

Private mlng����ID As Long
Private mstrReturn As String
Private mintPreCol As Integer, mintsort As Integer
Dim mblnChange As Boolean

Private Sub cbo��������_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub
Private Sub txtEdit_Change(Index As Integer)
    If Index = 0 And mblnChange = False Then
        g�������_�˰�.���˱�� = ""
        g�������_�˰�.���� = ""
    End If
End Sub

Private Sub txtEdit_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Dim strCurrDate As String
    Dim rsTemp As New ADODB.Recordset
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    g�������_�˰�.byt���� = mbytType
    mblnChange = True
    If Index = 0 Then
        SetOKCtrl False
        mblnChange = True
        
        If txtEdit(0).Text = "" Then
            ShowMsgbox "�����뿨��"
            txtEdit(0).SetFocus
            Exit Sub
        End If
        
        g�������_�˰�.���� = Mid(txtEdit(0).Text, 4, 16)
        If ��ݼ���_�˰� = False Then Exit Sub
        Call LoadCtrlData
        txtEdit(0).Text = g�������_�˰�.����
        SetOKCtrl True
    Else
        If mbytType = 0 Then
        Else
            SetOKCtrl False
            mblnChange = True
            g�������_�˰�.���� = txtEdit(1)
            g�������_�˰�.���˱�� = txtEdit(1)
            
            If ��ݼ���_�˰� = False Then Exit Sub
            Call LoadCtrlData
            SetOKCtrl True
        End If
    End If
    SetCboListIndex
    zlCommFun.PressKey vbKeyTab
End Sub

Private Sub SetOKCtrl(ByVal blnEn As Boolean)
    cmdȷ��.Enabled = blnEn
End Sub

Private Function IsValid() As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:��֤���ݵĺϷ���
    '--�����:
    '--������:
    '--��  ��:
    '-----------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    IsValid = False
    If mbytType = 0 Then
        If Trim(txtEdit(0).Text) = "" Then
            MsgBox "��û������ҽ�����ţ�", vbInformation, gstrSysName
            txtEdit(0).SetFocus
            Exit Function
        End If
        If Trim(txtEdit(1).Text) <> g�������_�˰�.���˱�� Then
            ShowMsgbox "ҽ��֤�Ų���ȷ��,����"
            txtEdit(1).SetFocus
            Exit Function
        End If
        If Trim(g�������_�˰�.����) = "" Then
            MsgBox "��û���������֤��", vbInformation, gstrSysName
            txtEdit(0).SetFocus
            Exit Function
        End If
    
    Else
        If Trim(txtEdit(1).Text) = "" Then
            MsgBox "��û������ҽ��֤�ţ�", vbInformation, gstrSysName
            If txtEdit(1).Enabled Then txtEdit(1).SetFocus
            Exit Function
        End If
        If Trim(g�������_�˰�.����) = "" Then
            MsgBox "��û���������֤��", vbInformation, gstrSysName
            If txtEdit(1).Enabled Then txtEdit(1).SetFocus
            Exit Function
        End If
    
    End If
    
    If cbo��������.Text = "" Then
        ShowMsgbox "�������δѡ��"
        Exit Function
    End If
    
    If mbytType <> 2 Then
        If mbytType = 4 Then
            '����鵱ǰ��̬
        Else
            '��鲡��״̬
            gstrSQL = "select nvl(��ǰ״̬,0) as ״̬ from �����ʻ� where ����=[1] and ҽ����=[2]"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, TYPE_�˰�, g�������_�˰�.���˱��)
            If rsTemp.RecordCount > 0 Then
                If rsTemp("״̬") > 0 Then
                    MsgBox "�ò����Ѿ���Ժ������ͨ�������֤��", vbInformation, gstrSysName
                    Exit Function
                End If
            End If
        End If
        If mbytType = 0 Or mbytType = 3 Then
            '����
        End If
    Else
        '�����������סԺ�ģ�ֻ��ˢ����ʾһ�����ݶ��ѣ�������
        Unload Me
        Exit Function
    End If
    IsValid = True
End Function

Private Sub cmdȡ��_Click()
    Unload Me
End Sub

Private Sub cmdȷ��_Click()
    Dim lng����ID As Long
    
    Dim strIdentify As String, strAddition As String
    Dim rsTemp As New ADODB.Recordset
    Dim str��� As String
    Dim int��ǰ״̬ As Integer
    
    If IsValid = False Then Exit Sub
    If mbytType = 0 Then
        Dim StrInput As String, strOutput As String
        StrInput = InitInfor_�˰�.ҽԺ���� & vbTab
        StrInput = StrInput & UserInfo.��� & vbTab
        StrInput = StrInput & UserInfo.���� & vbTab
        StrInput = StrInput & cbo��������.ItemData(cbo��������.ListIndex)
        If ҵ������_�˰�(����Աע��, StrInput, strOutput) = False Then Exit Sub
        
        g�������_�˰�.������ˮ�� = Split(strOutput, vbTab)(0)
    End If
    
    g�������_�˰�.�������� = cbo��������.Text
    
    int��ǰ״̬ = 0
    
    If mbytType = 4 Then
        '��ȷ����ǰ״̬,��Ϊ��ǰ״̬�ǲ��ܸı��
        gstrSQL = "Select * from �����ʻ� where ����=" & TYPE_�˰� & " and  ҽ����='" & g�������_�˰�.���˱�� & "'"
        
        zlDatabase.OpenRecordset rsTemp, gstrSQL, Me.Caption
        If Not rsTemp.EOF Then
            mlng����ID = Nvl(rsTemp!����ID, 0)
            int��ǰ״̬ = Nvl(rsTemp!��ǰ״̬, 0)
        End If
        rsTemp.Close
    End If
    
    '�������˵�����Ϣ�������ʽ��
    '0����;1ҽ����;2����;3����;4�Ա�;5��������;6���֤;7��λ����(����);
    '8.���Ĵ���;9.˳���;10��Ա���;11�ʻ����;12��ǰ״̬;13����ID;14��ְ(0,1);15����֤��;16�����;17�Ҷȼ�
    '18�ʻ������ۼ�,19�ʻ�֧���ۼ�,20����ͳ���ۼ�,21ͳ�ﱨ���ۼ�,22סԺ�����ۼ�
    With g�������_�˰�
        
        strIdentify = .����                               '0����
        strIdentify = strIdentify & ";" & .���˱��           '1ҽ����
        strIdentify = strIdentify & ";" & ""                 '2����
        strIdentify = strIdentify & ";" & .����               '3����
        strIdentify = strIdentify & ";" & .�Ա�                 '4�Ա�
        strIdentify = strIdentify & ";" & ""                    '5��������
        strIdentify = strIdentify & ";" & ""                      '6���֤
        strIdentify = strIdentify & ";" & .��λ����                 '7.��λ����(����)
        strAddition = ";0"                                          '8.���Ĵ���
        strAddition = strAddition & ";" & .סԺ�ǼǺ�                            '9.˳���
        strAddition = strAddition & ";" & .��Ա���                 '10��Ա���
        strAddition = strAddition & ";" & .�ʻ����                 '11�ʻ����
        
        strAddition = strAddition & ";" & int��ǰ״̬                            '12��ǰ״̬
        strAddition = strAddition & ";"                             '13����ID
        strAddition = strAddition & ";1"                            '14��ְ(1,2,3)
        strAddition = strAddition & ";" & IIf(.���ִ��� = "", "", .���ִ��� & "-" & .��������)                          '15����֤��
        strAddition = strAddition & ";" & .����                     '16�����
        strAddition = strAddition & ";"                             '17�Ҷȼ�
        strAddition = strAddition & ";" & .�ʻ����                 '18�ʻ������ۼ�
        strAddition = strAddition & ";0"                            '19�ʻ�֧���ۼ�
        strAddition = strAddition & ";0"                            '20���깤���ܶ�
        strAddition = strAddition & ";"                             '21סԺ�����ۼ�
    End With
    
    mlng����ID = BuildPatiInfo(0, strIdentify & strAddition, mlng����ID, TYPE_�˰�)
    
    '�����ʻ�:�����ֶ�:��������ҩƷ,���������Ը���,������������,�������û���ͳ��,�������ô�ͳ��,����סԺ����,����סԺ�𸶱�׼
    If mbytType = 0 Then
        '����:
        '���±����ʻ��������Ϣ
        gstrSQL = "zl_�����ʻ�_������Ϣ(" & mlng����ID & "," & TYPE_�˰� & ",'��������ҩƷ','" & g�������_�˰�.��������ҩƷ��� & "')"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "���浱������ҩƷ���")
        gstrSQL = "zl_�����ʻ�_������Ϣ(" & mlng����ID & "," & TYPE_�˰� & ",'���������Ը���','" & g�������_�˰�.���������Ը��� & "')"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "���浱�������Ը���")
        gstrSQL = "zl_�����ʻ�_������Ϣ(" & mlng����ID & "," & TYPE_�˰� & ",'������������','" & g�������_�˰�.������������ͳ�� & "')"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "���浱����������ͳ��")
        gstrSQL = "zl_�����ʻ�_������Ϣ(" & mlng����ID & "," & TYPE_�˰� & ",'������ˮ��','''" & g�������_�˰�.������ˮ�� & "''')"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "����������ˮ��")
    ElseIf mbytType = 1 Then
        gstrSQL = "zl_�����ʻ�_������Ϣ(" & mlng����ID & "," & TYPE_�˰� & ",'�������û���ͳ��','" & g�������_�˰�.�������û���ͳ�� & "')"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "���汾�����û���ͳ��")
        gstrSQL = "zl_�����ʻ�_������Ϣ(" & mlng����ID & "," & TYPE_�˰� & ",'�������ô�ͳ��','" & g�������_�˰�.�������ô�ͳ�� & "')"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "���汾�����ô�ͳ��")
        gstrSQL = "zl_�����ʻ�_������Ϣ(" & mlng����ID & "," & TYPE_�˰� & ",'����סԺ����','" & g�������_�˰�.����ڼ���סԺ & "')"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "���浱��ڼ���סԺ")
        gstrSQL = "zl_�����ʻ�_������Ϣ(" & mlng����ID & "," & TYPE_�˰� & ",'����סԺ�𸶱�׼','" & g�������_�˰�.����סԺ�𸶱�׼ & "')"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "���汾��סԺ�𸶱�׼")
    End If
    '���ظ�ʽ:�м���벡��ID
    If mlng����ID > 0 Then
        mstrReturn = strIdentify & ";" & mlng����ID & strAddition
    End If
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub
Public Function GetPatient(Optional bytType As Byte, Optional lng����ID As Long = 0) As String
    mbytType = bytType
    mlng����ID = lng����ID
    mstrReturn = ""
    DebugTool "���������֤,����ʼ���������Ϣ"
    
    If LoadBaseData = False Then
        DebugTool "����ʧ��(�����֤)"
        Exit Function
    End If
    DebugTool "����ɹ�(�����֤)"
    
    Me.Show 1
    lng����ID = mlng����ID
    GetPatient = mstrReturn
End Function
Private Function LoadBaseData() As Boolean
    '���ػ�������
    Dim rsTemp As New ADODB.Recordset
    LoadBaseData = False
    On Error GoTo errHand:
    
    If mbytType = 0 Then
        cbo��������.AddItem "��ͨҽ������"
        cbo��������.ListIndex = cbo��������.NewIndex
        cbo��������.ItemData(cbo��������.NewIndex) = 1
        cbo��������.AddItem "����ҽ������"
        cbo��������.ItemData(cbo��������.NewIndex) = 2
        txtEdit(0).Enabled = True
        txtEdit(1).Enabled = True
    Else
        cbo��������.AddItem "ҽ��סԺ"
        cbo��������.ListIndex = cbo��������.NewIndex
        txtEdit(0).Enabled = False
        txtEdit(1).Enabled = True
        cbo��������.Enabled = False
        lbl(12).Caption = "�����𸶱�׼"
        lbl(13).Caption = "���û���ͳ��"
        lbl(14).Caption = "���ô�ͳ��"
    End If
    
    LoadBaseData = True
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function
Private Sub LoadCtrlData()
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:�������
    '--�����:
    '--������:
    '--��  ��:
    '-----------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    With g�������_�˰�
       ' txtEdit(1).Text = .���˱��
        lblEdit(2).Caption = .����
        lblEdit(3).Caption = .�Ա�
        lblEdit(4).Caption = .����
        lblEdit(5).Caption = .��Ա���
        lblEdit(6).Caption = .�ʻ�״̬
        lblEdit(7).Caption = .�ʻ����
        lblEdit(8).Caption = .���ִ���
        lblEdit(9).Caption = .��������
        lblEdit(10).Caption = .��λ����
        
        lblEdit(15).Caption = .����ڼ���סԺ
        
        If mbytType = 0 Then
            lblEdit(12).Caption = .��������ҩƷ���
            lblEdit(13).Caption = .���������Ը���
            lblEdit(14).Caption = .������������ͳ��
        ElseIf mbytType = 1 Then
            lblEdit(12).Caption = .����סԺ�𸶱�׼
            lblEdit(13).Caption = .�������û���ͳ��
            lblEdit(14).Caption = .�������ô�ͳ��
        End If
    End With
End Sub
Private Sub SetCboListIndex()
    '���ÿؼ�����
    Dim i As Long
    If mbytType <> 0 Then Exit Sub
    If InStr(1, "���ݶ���", g�������_�˰�.��Ա���) <> 0 Then
       For i = 0 To cbo��������.ListCount - 1
            If cbo��������.ItemData(i) = 2 Then
                cbo��������.ListIndex = i
            End If
       Next
       cbo��������.Enabled = False
    Else
       cbo��������.Enabled = True
    End If
    
End Sub
