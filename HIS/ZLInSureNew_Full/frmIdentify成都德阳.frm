VERSION 5.00
Begin VB.Form frmIdentify�ɶ����� 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "���������֤"
   ClientHeight    =   4770
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6810
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4770
   ScaleWidth      =   6810
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmd���� 
      Caption         =   "��"
      Height          =   285
      Left            =   6360
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   3540
      Width           =   255
   End
   Begin VB.TextBox txtEdit 
      Height          =   300
      Left            =   4380
      MaxLength       =   25
      TabIndex        =   3
      Tag             =   "��ᱣ�Ϻ�"
      Top             =   1005
      Width           =   2265
   End
   Begin VB.CommandButton cmd�鿨 
      Caption         =   "���»�ȡ(&R)"
      Height          =   350
      Left            =   120
      TabIndex        =   22
      Top             =   4245
      Width           =   1305
   End
   Begin VB.ComboBox cbo�籣 
      Height          =   300
      Left            =   855
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1005
      Width           =   2310
   End
   Begin VB.CommandButton cmdȡ�� 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   5580
      TabIndex        =   24
      Top             =   4245
      Width           =   1100
   End
   Begin VB.Frame fra 
      Height          =   75
      Index           =   1
      Left            =   -60
      TabIndex        =   26
      Top             =   615
      Width           =   8340
   End
   Begin VB.Frame fra 
      Height          =   75
      Index           =   0
      Left            =   -510
      TabIndex        =   25
      Top             =   3960
      Width           =   8340
   End
   Begin VB.TextBox txt���� 
      Height          =   315
      Left            =   825
      TabIndex        =   29
      Top             =   3525
      Width           =   5820
   End
   Begin VB.CommandButton cmdȷ�� 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   4290
      TabIndex        =   23
      Top             =   4245
      Width           =   1100
   End
   Begin VB.Label lbl���� 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "����(&F)"
      Height          =   180
      Left            =   165
      TabIndex        =   30
      Top             =   3630
      Width           =   630
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "ҽ������"
      Height          =   180
      Index           =   4
      Left            =   3645
      TabIndex        =   14
      Top             =   2355
      Width           =   720
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   6
      Left            =   855
      TabIndex        =   17
      Top             =   2760
      Width           =   2310
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   0
      Left            =   855
      TabIndex        =   5
      Top             =   1440
      Width           =   2310
   End
   Begin VB.Label lblInfor 
      AutoSize        =   -1  'True
      Caption         =   "ҽ�����˻�����Ϣ��ʾ������ͨ��[���»�ȡ]��ť���½��ж�ȡ���˻�����Ϣ��"
      Height          =   180
      Left            =   630
      TabIndex        =   27
      Top             =   360
      Width           =   6300
   End
   Begin VB.Image img 
      Height          =   480
      Left            =   60
      Picture         =   "frmIdentify�ɶ�����.frx":0000
      Top             =   105
      Width           =   480
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "��¼��"
      Height          =   180
      Index           =   0
      Left            =   300
      TabIndex        =   4
      Top             =   1485
      Width           =   540
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "��ᱣ�Ϻ�"
      Height          =   180
      Index           =   1
      Left            =   3465
      TabIndex        =   2
      Top             =   1065
      Width           =   900
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "����"
      Height          =   180
      Index           =   2
      Left            =   4005
      TabIndex        =   6
      Top             =   1485
      Width           =   360
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "�Ա�"
      Height          =   180
      Index           =   3
      Left            =   480
      TabIndex        =   8
      Top             =   1905
      Width           =   360
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "��������"
      Height          =   180
      Index           =   5
      Left            =   120
      TabIndex        =   12
      Top             =   2355
      Width           =   720
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "���ݹ���"
      Height          =   180
      Index           =   6
      Left            =   120
      TabIndex        =   16
      Top             =   2805
      Width           =   720
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "�籣����"
      Height          =   180
      Index           =   7
      Left            =   120
      TabIndex        =   0
      Top             =   1065
      Width           =   720
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "����"
      Height          =   180
      Index           =   8
      Left            =   4005
      TabIndex        =   10
      Top             =   1905
      Width           =   360
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "��λ����"
      Height          =   180
      Index           =   10
      Left            =   120
      TabIndex        =   20
      Top             =   3210
      Width           =   720
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "ҽ�Ʊ�־"
      Height          =   180
      Index           =   12
      Left            =   3645
      TabIndex        =   18
      Top             =   2805
      Width           =   720
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Index           =   1
      Left            =   4380
      TabIndex        =   7
      Top             =   1425
      Width           =   2265
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   2
      Left            =   855
      TabIndex        =   9
      Top             =   1860
      Width           =   2310
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Index           =   3
      Left            =   4380
      TabIndex        =   11
      Top             =   1845
      Width           =   2265
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   4
      Left            =   855
      TabIndex        =   13
      Top             =   2310
      Width           =   2310
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Index           =   5
      Left            =   4380
      TabIndex        =   15
      Top             =   2295
      Width           =   2265
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   8
      Left            =   855
      TabIndex        =   21
      Top             =   3165
      Width           =   5775
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Index           =   7
      Left            =   4380
      TabIndex        =   19
      Top             =   2745
      Width           =   2265
   End
End
Attribute VB_Name = "frmIdentify�ɶ�����"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mbytType As Byte            '0-�����շѣ�1-��Ժ�Ǽǣ�2-������������סԺ,3-�Һ�,4-����

Private mlng����ID As Long
Private mstrReturn As String
Private mintPreCol As Integer, mintsort As Integer
Private mblnFirst As Boolean        '��һ����ϵͳʱ����
Private mblnChange As Boolean
Private Sub cbo�籣_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub cmd�鿨_Click()
   If ��ȡ�α���Ա��Ϣ = False Then
        cmdȷ��.Enabled = False
        Call ClearData
        Exit Sub
    End If
    Call LoadCtrlData
    cmdȷ��.Enabled = True
End Sub

Private Sub Form_Activate()
    '
    If mblnFirst = False Then Exit Sub
    mblnFirst = False
    
    
'    If ��ȡ�α���Ա��Ϣ = False Then
'        cmdȷ��.Enabled = False
'        Exit Sub
'    End If
'    Call LoadCtrlData
    cmdȷ��.Enabled = False
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
    Dim StrInput As String, strOutput As String
    Dim lng״̬ As Long
    
    IsValid = False
    If Trim(g�������_�ɶ�����.����) = "" Then
        MsgBox "��û���������֤��", vbInformation, gstrSysName
        If cmd�鿨.Enabled Then cmd�鿨.SetFocus
        Exit Function
    End If
    
     If cbo�籣.Text = "" Then
        ShowMsgbox "�籣������δѡ��"
        Exit Function
    End If
    If g�������_�ɶ�����.���Ϻ� = "" Then
        ShowMsgbox "��������ᱣ�Ϻ�!"
        Exit Function
    End If
      
    If mbytType <> 2 Then
        If mbytType = 4 Then
            '�����¼ǰ��̬
        Else
            '��鲡��״̬
            gstrSQL = "select nvl(��ǰ״̬,0) as ״̬ from �����ʻ� where ����=[1] and ҽ����=[2]"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, TYPE_�ɶ�����, g�������_�ɶ�����.���Ϻ�)
            If rsTemp.RecordCount > 0 Then
                If rsTemp("״̬") > 0 Then
                    MsgBox "�ò����Ѿ���Ժ������ͨ�������֤��", vbInformation, gstrSysName
                    Exit Function
                End If
            End If
        End If
    Else
        '�����������סԺ�ģ�ֻ��ˢ����ʾһ�����ݶ��ѣ�������
        '�����������סԺ�ģ�ֻ��ˢ����ʾһ�����ݶ��ѣ�������
        '��ȷ����ǰ״̬,��Ϊ��ǰ״̬�ǲ��ܸı��
        gstrSQL = "Select * from �����ʻ� where ����=[1] and  ҽ����=[2]"
        
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, TYPE_�ɶ�����, g�������_�ɶ�����.���Ϻ�)
        If Not rsTemp.EOF Then
            mlng����ID = Nvl(rsTemp!����ID, 0)
        End If
        rsTemp.Close
        mstrReturn = mlng����ID
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
    Dim StrInput  As String, strOutput As String
    Dim strIdentify As String, strAddition As String
    Dim rsTemp As New ADODB.Recordset
    Dim str�籣 As String
    Dim int��ǰ״̬ As Integer
    Dim lng״̬ As Long
    
    
    lng����ID = IIf(Val(Me.txt����.Tag) = 0, 0, Val(Me.txt����.Tag))
    g�������_�ɶ�����.�������� = Split(cbo�籣.Text, "-")(0)
    
    If IsValid = False Then Exit Sub
    
    If lng����ID <> 0 Then
        gstrSQL = "Select * From ���ղ��� where id=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ����", lng����ID)
        g�������_�ɶ�����.���ֱ��� = Nvl(rsTemp!����)
        g�������_�ɶ�����.�������� = Nvl(rsTemp!����)
    Else
        g�������_�ɶ�����.���ֱ��� = ""
        g�������_�ɶ�����.�������� = ""
    End If
    
    int��ǰ״̬ = 0
    If mbytType = 4 Then
        '��ȷ����ǰ״̬,��Ϊ��ǰ״̬�ǲ��ܸı��
        gstrSQL = "Select * from �����ʻ� where ����=[1] and  ҽ����=[2]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "", TYPE_�ɶ�����, g�������_�ɶ�����.���Ϻ�)
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
    With g�������_�ɶ�����
        
        strIdentify = .��¼��                                '0����
        strIdentify = strIdentify & ";" & .���Ϻ�            '1ҽ����
        strIdentify = strIdentify & ";"                    '2����
        strIdentify = strIdentify & ";" & .����               '3����
        strIdentify = strIdentify & ";" & Decode(.�Ա�, "1", "��", "2", "Ů", .�Ա�)              '4�Ա�
        strIdentify = strIdentify & ";" & .��������                '5��������
        strIdentify = strIdentify & ";" & ""           '6���֤
        strIdentify = strIdentify & ";" & .��λ���� & IIf(.��λ���� = 0, "", "(" & .��λ���� & ")")          '7.��λ����(����)
        strAddition = ";" & cbo�籣.ItemData(cbo�籣.ListIndex)                                           '8.���Ĵ���
        strAddition = strAddition & ";" & .ҽ������                              '9.˳���
        strAddition = strAddition & ";"                                '10��Ա���
        strAddition = strAddition & ";" & ""                 '11�ʻ����
        
        strAddition = strAddition & ";" & int��ǰ״̬                            '12��ǰ״̬
        strAddition = strAddition & ";" & lng����ID             '13����ID
        strAddition = strAddition & ";1"                        '14��ְ(1,2,3)
        strAddition = strAddition & ";" & .���ݹ���           '15����֤��
        strAddition = strAddition & ";" & .����                     '16�����
        strAddition = strAddition & ";"                         '17�Ҷȼ�
        strAddition = strAddition & ";"                         '18�ʻ������ۼ�
        strAddition = strAddition & ";0"                            '19�ʻ�֧���ۼ�
        strAddition = strAddition & ";0"                            '20���깤���ܶ�
        strAddition = strAddition & ";"                             '21סԺ�����ۼ�
    End With
    
    mlng����ID = BuildPatiInfo(0, strIdentify & strAddition, mlng����ID, TYPE_�ɶ�����)
    If mlng����ID = 0 Then Exit Sub
    
    If mbytType = 0 Or mbytType = 3 Then
    Else
        
        gstrSQL = "zl_�����ʻ�_������Ϣ(" & mlng����ID & "," & TYPE_�ɶ����� & ",'ҽ�Ʊ�־','''" & g�������_�ɶ�����.ҽ�Ʊ�־ & "''')"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "ҽ�Ʊ�־")
        
        gstrSQL = "zl_�����ʻ�_������Ϣ(" & mlng����ID & "," & TYPE_�ɶ����� & ",'�籣���','''" & g�������_�ɶ�����.�������� & "''')"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "�籣���")
    End If
    g�������_�ɶ�����.����ID = mlng����ID
    
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
    If Load�籣���� = False Then
        DebugTool "����ʧ��(�����֤)"
        Exit Function
    End If
    DebugTool "����ɹ�(�����֤)"
    
    Me.Show 1
    lng����ID = mlng����ID
    GetPatient = mstrReturn
End Function
Private Sub LoadCtrlData()
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:�������
    '--�����:
    '--������:
    '--��  ��:
    '-----------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    With g�������_�ɶ�����
        lblEdit(0).Caption = .��¼��
        lblEdit(1).Caption = .����
        lblEdit(2).Caption = Decode(.�Ա�, "1", "��", "2", "Ů", .�Ա�)
        lblEdit(3).Caption = .����
        
        lblEdit(4).Caption = .��������
        lblEdit(5).Caption = .ҽ������
        
        lblEdit(6).Caption = .���ݹ���
        
        lblEdit(7).Caption = .ҽ�Ʊ�־
        lblEdit(8).Caption = .��λ���� & IIf(.��λ���� <> "", "(" & .��λ���� & ")", "")
    End With
End Sub
Private Sub Form_Load()
        mblnFirst = True
End Sub

Private Function Load�籣����() As Boolean
    Dim rsTemp As New ADODB.Recordset
On Error GoTo errHand:
    gstrSQL = "Select * From ��������Ŀ¼ where ����=[1] and ���<>0 order by ����"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, TYPE_�ɶ�����)
    If rsTemp.RecordCount = 0 Then
        ShowMsgbox "�����籣����Ŀ¼�����ڲ��������ػ���!"
        Exit Function
    End If
    
    With rsTemp
        cbo�籣.Clear
        Do While Not .EOF
            cbo�籣.AddItem Nvl(!����) & "--" & Nvl(!����)
            cbo�籣.ItemData(cbo�籣.NewIndex) = Nvl(!���, 0)
            .MoveNext
        Loop
    End With
    cbo�籣.ListIndex = 0
    SetDefaultSel
'    cbo�籣.Enabled = False
    Load�籣���� = True
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function
Private Function SetDefaultSel() As Boolean
    Dim strReg As String
    Dim i As Integer
    
    SetDefaultSel = False
    Err = 0: On Error GoTo errHand:
    Call GetRegInFor(g����ģ��, "ҽ��", "�籣��������", strReg)
    If cbo�籣.ListCount = 0 Then Exit Function
    For i = 0 To cbo�籣.ListCount
        If Split(cbo�籣.List(i), "--")(0) = strReg Then
            cbo�籣.ListIndex = i
            Exit For
        End If
    Next
    If cbo�籣.ListIndex < 0 Then
        cbo�籣.ListIndex = 0
    End If
    SetDefaultSel = True
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function

Private Function ��ȡ�α���Ա��Ϣ() As Boolean
    '��ȡ�α���Ա��Ϣ
    Dim StrInput As String
    Dim strOutput As String
    Dim strArr
    
    ��ȡ�α���Ա��Ϣ = False
    
    
    Err = 0
    On Error GoTo errHand:
    g�������_�ɶ�����.�������� = Split(cbo�籣.Text, "-")(0)
    g�������_�ɶ�����.���Ϻ� = txtEdit.Text
   
   If g�������_�ɶ�����.���Ϻ� = "" Then
        ShowMsgbox "��������ᱣ�Ϻ�!"
        Exit Function
    End If
    'ASBBH   PCHAR   �α���Ա����ᱣ�Ϻ�
    'ABXJGBH PCHAR   �α���Ա���ڵı��ջ������
    
    StrInput = g�������_�ɶ�����.���Ϻ�
    StrInput = StrInput & "||" & g�������_�ɶ�����.��������
    
    If ҵ������_�ɶ�����(��òα���Ա����, StrInput, strOutput) = False Then
        Call ClearData
        Exit Function
    End If
    
    strArr = Split(strOutput, "||")
    '����: ���˼�¼��||ҽ������||���ݹ���||��λ����||����||�Ա�||�������ڣ���ʽ��YYYY-MM-DD��||��λ���||�μӻ���ҽ�Ʊ�־
    
    With g�������_�ɶ�����
        .��¼�� = strArr(0)
        .ҽ������ = strArr(1)
        .���ݹ��� = strArr(2)
        .��λ���� = strArr(3)
        
        .���� = strArr(4)
        .�Ա� = strArr(5)
        .�������� = strArr(6)
        .��λ���� = strArr(7)
        .ҽ�Ʊ�־ = strArr(8)
        .���� = Get����(.��������)
    End With
    ��ȡ�α���Ա��Ϣ = True
    Exit Function
errHand:
        If ErrCenter = 1 Then
            Resume
        End If
End Function
Private Function Get����(ByVal strDate As String) As Integer
    Dim rsTemp As New ADODB.Recordset
    Err = 0
    On Error GoTo errHand:
    gstrSQL = "Select (sysdate-to_date('" & strDate & "','yyyy-mm-dd'))/365 as ���� from dual "
    zlDatabase.OpenRecordset rsTemp, gstrSQL, Me.Caption
    If Not rsTemp.EOF Then
        Get���� = Int(Nvl(rsTemp!����, 0))
        Exit Function
    End If
    Exit Function
errHand:
End Function
Private Sub ClearData()
    Dim i As Long
    '��������Ϣ
    With g�������_�ɶ�����
        .���� = ""
        .�Ա� = ""
        .�������� = ""
        .�������� = ""
        .��λ���� = ""
        .��λ���� = ""
        .���Ϻ� = ""
        .��¼�� = ""
    End With
    For i = 0 To lblEdit.UBound
        lblEdit(i).Caption = ""
    Next
End Sub
Private Sub txtEdit_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    cmd�鿨_Click
    zlCommFun.PressKey vbKeyTab
End Sub

Private Sub txtEdit_KeyPress(KeyAscii As Integer)
    zlControl.TxtCheckKeyPress txtEdit, KeyAscii, m�ı�ʽ
End Sub
Private Sub cmd����_Click()
    Dim rsTemp As New ADODB.Recordset
    gstrSQL = " Select A.ID,����,A.����,A.����,decode(A.���,1,'���Բ�',2,'���ֲ�','��ͨ��') as ��� " & _
            " From ���ղ��� A where A.����=" & TYPE_�ɶ�����
    
    Set rsTemp = frmPubSel.ShowSelect(Me, gstrSQL, 0, "ҽ������", , txt����.Text)
    If rsTemp.State = 0 Then Exit Sub
    If Not rsTemp Is Nothing Then
        txt����.Text = rsTemp("����")
        txt����.Tag = rsTemp("ID")
        zlControl.TxtSelAll txt����
    End If
    txt����.SetFocus
End Sub

Private Sub txt����_Change()
    txt����.Tag = ""
    txt����.ForeColor = &HC0&
End Sub

Private Sub txt����_KeyPress(KeyAscii As Integer)
    Dim rsTemp As New ADODB.Recordset, strText As String
    Dim blnReturn As Boolean, lst As ListItem
    
    
    If KeyAscii <> vbKeyReturn Then Exit Sub
    If txt����.Text = "" Or txt����.Tag <> "" Then
        SendKeys "{TAB}"
        Exit Sub
    End If
    
    On Error GoTo errHandle
    
    strText = txt����.Text
    gstrSQL = "Select A.ID,A.����,A.����,A.����,decode(A.���,1,'���Բ�',2,'���ⲡ','��ͨ��') ��� " & _
             "   FROM ���ղ��� A WHERE A.����=[1] And " & _
             "(A.���� like [2] || '%' or A.���� like [2] || '%' or A.���� like [2] || '%')"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, TYPE_�ɶ�����, strText)
    
    If rsTemp.RecordCount > 0 Then
        '����ѡ����
        If rsTemp.RecordCount > 1 Then
            '�����ֶδ���3�ģ���ʹֻ��һ����¼�ѸöԻ�����ʾ�������Ա����û��õ��������Ϣ
            blnReturn = frmListSel.ShowSelect(TYPE_�ɶ�����, rsTemp, "ID", "ҽ������ѡ��", "��ѡ���ض���ҽ�����֣�")
        Else
            blnReturn = True
        End If
    End If
    
    If blnReturn = False Then
        '��¼����û�п�ѡ�������
        zlControl.TxtSelAll txt����
        Exit Sub
    Else
        '�϶����м�¼����
        txt����.Text = rsTemp("����")
        txt����.Tag = rsTemp("ID")
        SendKeys "{TAB}"
    End If
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
Private Sub txt����_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        txt����.Text = ""
        txt����.Tag = ""
    End If
End Sub

