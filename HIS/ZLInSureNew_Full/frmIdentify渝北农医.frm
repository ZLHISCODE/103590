VERSION 5.00
Begin VB.Form frmIdentify�山ũҽ 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "�����֤"
   ClientHeight    =   4635
   ClientLeft      =   6615
   ClientTop       =   5505
   ClientWidth     =   5880
   Icon            =   "frmIdentify�山ũҽ.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4635
   ScaleWidth      =   5880
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.TextBox txt��ͥ�ʻ���� 
      Enabled         =   0   'False
      Height          =   300
      Left            =   1650
      TabIndex        =   15
      Top             =   3030
      Width           =   2355
   End
   Begin VB.ComboBox cboҽ����� 
      Height          =   300
      Left            =   1650
      Style           =   2  'Dropdown List
      TabIndex        =   17
      Top             =   3420
      Width           =   2385
   End
   Begin VB.TextBox txt������Ϣ 
      Height          =   300
      Left            =   1650
      TabIndex        =   19
      Top             =   3810
      Width           =   2355
   End
   Begin VB.TextBox txt����֢ 
      Height          =   300
      Left            =   1650
      TabIndex        =   21
      Top             =   4200
      Width           =   2355
   End
   Begin VB.CommandButton cmdȡ�� 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   4590
      TabIndex        =   24
      Top             =   840
      Width           =   1100
   End
   Begin VB.CommandButton cmdȷ�� 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   4590
      TabIndex        =   23
      Top             =   360
      Width           =   1100
   End
   Begin VB.Frame Frame1 
      Height          =   4515
      Left            =   4290
      TabIndex        =   22
      Top             =   -30
      Width           =   30
   End
   Begin VB.TextBox txt�����ʻ���� 
      Enabled         =   0   'False
      Height          =   300
      Left            =   1650
      TabIndex        =   13
      Top             =   2640
      Width           =   2355
   End
   Begin VB.TextBox txt���֤�� 
      Enabled         =   0   'False
      Height          =   300
      Left            =   1650
      TabIndex        =   11
      Top             =   2250
      Width           =   2355
   End
   Begin VB.TextBox txt�������� 
      Enabled         =   0   'False
      Height          =   300
      Left            =   1650
      TabIndex        =   9
      Top             =   1860
      Width           =   1365
   End
   Begin VB.TextBox txt�Ա� 
      Enabled         =   0   'False
      Height          =   300
      Left            =   1650
      TabIndex        =   7
      Top             =   1470
      Width           =   855
   End
   Begin VB.TextBox txt���� 
      Enabled         =   0   'False
      Height          =   300
      Left            =   1650
      TabIndex        =   5
      Top             =   1080
      Width           =   2355
   End
   Begin VB.TextBox txt���˱��� 
      Enabled         =   0   'False
      Height          =   300
      Left            =   1650
      TabIndex        =   3
      Top             =   690
      Width           =   2355
   End
   Begin VB.TextBox txtҽ��֤�� 
      Height          =   300
      Left            =   1650
      MaxLength       =   25
      TabIndex        =   1
      Top             =   300
      Width           =   2355
   End
   Begin VB.Label lbl��ͥ�ʻ���� 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "��ͥ�ʻ����(&F)"
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   210
      TabIndex        =   14
      Top             =   3090
      Width           =   1350
   End
   Begin VB.Label lblҽ����� 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "ҽ�����(&T)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   570
      TabIndex        =   16
      Top             =   3480
      Width           =   990
   End
   Begin VB.Label lbl������Ϣ 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "������Ϣ(&D)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   570
      TabIndex        =   18
      Top             =   3870
      Width           =   990
   End
   Begin VB.Label lbl����֢ 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "����֢(&Z)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   750
      TabIndex        =   20
      Top             =   4260
      Width           =   810
   End
   Begin VB.Label lbl������� 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "�����ʻ����(&Y)"
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   210
      TabIndex        =   12
      Top             =   2700
      Width           =   1350
   End
   Begin VB.Label lbl���֤�� 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "���֤��(&I)"
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   570
      TabIndex        =   10
      Top             =   2310
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
      Left            =   570
      TabIndex        =   8
      Top             =   1920
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
      Left            =   930
      TabIndex        =   6
      Top             =   1530
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
      Left            =   930
      TabIndex        =   4
      Top             =   1140
      Width           =   630
   End
   Begin VB.Label lbl���˱��� 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "���˱���(&P)"
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   570
      TabIndex        =   2
      Top             =   750
      Width           =   990
   End
   Begin VB.Label lblҽ��֤�� 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "ҽ��֤��(&L)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   570
      TabIndex        =   0
      Top             =   360
      Width           =   990
   End
End
Attribute VB_Name = "frmIdentify�山ũҽ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mbytType As Byte            '0-�����շѣ�1-��Ժ�Ǽǣ�2-������������סԺ,3-�Һ�
Private mlng����ID As Long
Private mstrReturn As String

Private Sub cboҽ�����_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub txt����֢_GotFocus()
    Call zlControl.TxtSelAll(txt����֢)
End Sub

Private Sub txt����֢_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub txt������Ϣ_GotFocus()
    Call zlControl.TxtSelAll(txt������Ϣ)
End Sub

Private Sub txt������Ϣ_KeyPress(KeyAscii As Integer)
    Dim strLike As String
    Dim StrInput As String
    Dim str�Ա� As String
    Dim blnCancel As Boolean
    Dim rsTemp As ADODB.Recordset
    On Error GoTo errHandle
    If KeyAscii <> vbKeyReturn Then Exit Sub
    
    If txt������Ϣ.Text = lbl������Ϣ.Tag And txt������Ϣ.Text <> "" Then
        Call zlCommFun.PressKey(vbKeyTab)
    ElseIf txt������Ϣ.Text = "" Then
        txt������Ϣ.Tag = "": lbl������Ϣ.Tag = ""
        Call zlCommFun.PressKey(vbKeyTab) '��������
    Else
        strLike = IIf(GetSetting("ZLSOFT", "����ģ��\����", "����ƥ��", 0) = 0, "%", "")
        StrInput = UCase(txt������Ϣ.Text)
        str�Ա� = txt�Ա�.Text
        If str�Ա� = "��" Then
            str�Ա� = " And (A.�Ա�����='��' Or A.�Ա����� is NULL)"
        ElseIf str�Ա� = "Ů" Then
            str�Ա� = " And (A.�Ա�����='Ů' Or A.�Ա����� is NULL)"
        Else
            str�Ա� = ""
        End If
        gstrSQL = "Select A.ID,A.����,A.����,A.����,A.����,A.˵��,A.�Ա�����,B.���" & _
            " From ��������Ŀ¼ A,����������� B" & _
            " Where A.���=B.���� And A.��� Not IN('B','Z')" & _
            " And (A.���� Like '" & StrInput & "%'" & _
            " Or Upper(A.����) Like '" & strLike & StrInput & "%'" & _
            " Or Upper(A.����) Like '" & strLike & StrInput & "%'" & _
            " Or Upper(A.����) Like '" & strLike & StrInput & "%')" & _
            " And Rownum<=100" & str�Ա� & _
            " Order by A.���,A.����"
        Set rsTemp = zlDatabase.ShowSelect(Me, gstrSQL, 0, "��������Input", , , , , , True, _
            txt������Ϣ.Left + Me.Left, _
            txt������Ϣ.Top + Me.Top, txt������Ϣ.Height, blnCancel, , True)
        If Not rsTemp Is Nothing Then
            txt������Ϣ.Tag = rsTemp!ID
            txt������Ϣ.Text = "(" & rsTemp!���� & ")" & rsTemp!����
            lbl������Ϣ.Tag = txt������Ϣ.Text '���ڻָ���ʾ
            Call zlCommFun.PressKey(vbKeyTab)
        Else
            If Not blnCancel Then
                MsgBox "û���ҵ�ƥ��ļ������롣", vbInformation, gstrSysName
            End If
            If lbl������Ϣ.Tag <> "" Then txt������Ϣ.Text = lbl������Ϣ.Tag
            txt������Ϣ.SetFocus
        End If
    End If
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub cmdȡ��_Click()
    Unload Me
End Sub

Private Sub cmdȷ��_Click()
    Dim lng����ID As Long
    Dim str�������� As String
    Dim strIdentify As String, strAddition As String
    Dim rsTemp As New ADODB.Recordset
    
    If Trim(txtҽ��֤��.Text) = "" Then
        MsgBox "��û������ҽ��֤�ţ�", vbInformation, gstrSysName
        txtҽ��֤��.SetFocus
        Exit Sub
    End If
    If Trim(txt����.Text) = "" Then
        MsgBox "��û�л�ȡ��ҽ�����˵������Ϣ������ͨ����֤��", vbInformation, gstrSysName
        txtҽ��֤��.SetFocus
        Exit Sub
    End If
    If mbytType <> 3 Then
        If txt������Ϣ.Tag = "" Then
            MsgBox "�����벡�˵ļ�����Ϣ��", vbInformation, gstrSysName
            txt������Ϣ.SetFocus
            Exit Sub
        End If
    
        '��ȡ����ID
        gstrSQL = "Select ID From ��������Ŀ¼ Where ����=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ���ղ���", CStr(txt������Ϣ.Tag))
        If Not rsTemp.EOF Then
            lng����ID = rsTemp!ID
        End If
    End If
    
    If mbytType <> 2 Then
        '��鲡��״̬
        gstrSQL = "select nvl(��ǰ״̬,0) as ״̬ from �����ʻ� where ����=[1] and ҽ����=[2]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, TYPE_�山ũҽ, CStr(txtҽ��֤��.Text))
        If rsTemp.RecordCount > 0 Then
            If rsTemp("״̬") > 0 Then
                MsgBox "�ò����Ѿ���Ժ������ͨ�������֤��", vbInformation, gstrSysName
                Exit Sub
            End If
        End If
    Else
        '�����������סԺ�ģ�ֻ��ˢ����ʾһ�����ݶ��ѣ�������
        Unload Me
        Exit Sub
    End If
    
    '�������˵�����Ϣ�������ʽ��
    '0����;1ҽ����;2����;3����;4�Ա�;5��������;6���֤;7��λ����(����);
    '8.���Ĵ���;9.˳���;10��Ա���;11�ʻ����;12��ǰ״̬;13����ID;14��ְ(0,1);15����֤��;16�����;17�Ҷȼ�
    '18�ʻ������ۼ�,19�ʻ�֧���ۼ�,20����ͳ���ۼ�,21ͳ�ﱨ���ۼ�,22סԺ�����ۼ�
    strIdentify = txtҽ��֤��.Text                              '0����
    strIdentify = strIdentify & ";" & txt���˱���.Text          '1ҽ����
    strIdentify = strIdentify & ";"                             '2����
    strIdentify = strIdentify & ";" & txt����.Text              '3����
    strIdentify = strIdentify & ";" & txt�Ա�.Text              '4�Ա�
    strIdentify = strIdentify & ";" & txt��������.Text          '5��������
    strIdentify = strIdentify & ";" & txt���֤��.Text          '6���֤
    strIdentify = strIdentify & ";"                             '7.��λ����(����)
    strAddition = ";0"                                          '8.���Ĵ���
    strAddition = strAddition & ";"                             '9.˳���
    strAddition = strAddition & ";"                             '10��Ա���
    strAddition = strAddition & ";" & Val(txt�����ʻ����.Text)     '11�ʻ����
    strAddition = strAddition & ";0"                            '12��ǰ״̬
    strAddition = strAddition & ";" & lng����ID                 '13����ID
    strAddition = strAddition & ";1"                            '14��ְ(1,2,3)
    strAddition = strAddition & ";"                             '15����֤��
    strAddition = strAddition & ";"                             '16�����
    strAddition = strAddition & ";"                             '17�Ҷȼ�
    strAddition = strAddition & ";" & Val(txt�����ʻ����.Text)     '18�ʻ������ۼ�
    strAddition = strAddition & ";0"                            '19�ʻ�֧���ۼ�
    strAddition = strAddition & ";0"                            '20���깤���ܶ�
    strAddition = strAddition & ";"                             '21סԺ�����ۼ�
    
    mlng����ID = BuildPatiInfo(0, strIdentify & strAddition, mlng����ID, TYPE_�山ũҽ)
    '���ظ�ʽ:�м���벡��ID
    If mlng����ID > 0 Then
        mstrReturn = strIdentify & ";" & mlng����ID & strAddition
    End If
    
    '���¼�ͥ�ʻ����
    gstrSQL = "zl_�����ʻ�_������Ϣ(" & mlng����ID & "," & TYPE_�山ũҽ & ",'��ͥ�ʻ����','''" & Val(txt��ͥ�ʻ����.Text) & "''')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "���¼�ͥ�ʻ����")
    
    With gComInfo_�山ũҽ
        .ҽ��֤�� = txtҽ��֤��.Text
        .���˱�� = txt���˱���.Text
        .�������� = txt������Ϣ.Tag
        .����֢ = txt����֢.Text
        .ҵ������ = Me.cboҽ�����.ItemData(Me.cboҽ�����.ListIndex)
    End With
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
'      11  ҽ�����    ��ͨ����
'      12  ҽ�����    ��������
'      13  ҽ�����    ��������
'      21  ҽ�����    ��ͨסԺ
'      22  ҽ�����    ת��ҽԺסԺ
    With Me.cboҽ�����
        .Clear
        '��������ҵ��
        If mbytType = 0 Or mbytType = 2 Or mbytType = 3 Then
            .AddItem "��ͨ����"
            .ItemData(.NewIndex) = 11
            .AddItem "��������"
            .ItemData(.NewIndex) = 12
            .AddItem "��������"
            .ItemData(.NewIndex) = 13
        End If
        '����סԺҵ��
        If mbytType = 1 Or mbytType = 2 Then
            .AddItem "��ͨסԺ"
            .ItemData(.NewIndex) = 21
            .AddItem "ת��ҽԺסԺ"
            .ItemData(.NewIndex) = 22
        End If
        .ListIndex = 0
    End With
    
    '�Һſ��Բ������뼲���벢��֢��Ϣ
    If mbytType = 3 Then
        txt������Ϣ.Enabled = False
        txt����֢.Enabled = False
    End If
End Sub

Private Sub txtҽ��֤��_GotFocus()
    Call zlControl.TxtSelAll(txtҽ��֤��)
End Sub

Public Function GetPatient(Optional bytType As Byte, Optional lng����ID As Long = 0) As String
    mbytType = bytType
    mlng����ID = lng����ID
    mstrReturn = ""
    Me.Show 1
    lng����ID = mlng����ID
    GetPatient = mstrReturn
End Function

Private Sub txtҽ��֤��_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim StrInput As String
    Dim arrOutput
    If KeyCode <> vbKeyReturn Then Exit Sub
    StrInput = Trim(txtҽ��֤��.Text)
    If Trim(StrInput) = "" Then Exit Sub
    
    Call ���ýӿ�_׼��_�山ũҽ(��ȡ������Ϣ_�山ũҽ, StrInput)
    If Not ���ýӿ�_�山ũҽ Then Exit Sub
    
    arrOutput = Split(gstrOutput_�山ũҽ, gstrSplit_Col_����������)
    If Val(arrOutput(7)) = 1 Then
        MsgBox "�ÿ��ѱ��������������ٴΰ������Ǽǣ�", vbInformation, gstrSysName
        Exit Sub
    End If
    
    txt���˱���.Text = arrOutput(1)
    txt����.Text = arrOutput(2)
    txt�Ա�.Text = IIf(Val(arrOutput(3)) = 0, "��", "Ů")
    txt��������.Text = Format(arrOutput(4), "yyyy-MM-dd")
    txt���֤��.Text = arrOutput(5)
    txt�����ʻ����.Text = Format(arrOutput(6), "#0.00")
    txt��ͥ�ʻ����.Text = Format(arrOutput(12), "#0.00")
    If KeyCode = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub
