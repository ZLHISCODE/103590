VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmIdentifyüɽ 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ҽ�����˵���"
   ClientHeight    =   3915
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7095
   Icon            =   "frmIdentifyüɽ.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3915
   ScaleWidth      =   7095
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdChange 
      Caption         =   "�޸�����(&M)"
      Height          =   405
      Left            =   330
      TabIndex        =   34
      Top             =   3390
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame fra���� 
      Caption         =   "���˻�����Ϣ"
      Height          =   3135
      Index           =   0
      Left            =   180
      TabIndex        =   0
      Top             =   150
      Width           =   6795
      Begin VB.TextBox txt�ʻ���� 
         Enabled         =   0   'False
         Height          =   300
         Left            =   1320
         MaxLength       =   20
         TabIndex        =   15
         Top             =   2670
         Width           =   1455
      End
      Begin VB.CommandButton cmdSelect 
         Caption         =   "��"
         Height          =   240
         Index           =   0
         Left            =   2490
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   360
         Width           =   255
      End
      Begin VB.TextBox txt���� 
         Height          =   300
         Left            =   1320
         MaxLength       =   20
         TabIndex        =   2
         Top             =   330
         Width           =   1455
      End
      Begin VB.TextBox txt���� 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1320
         MaxLength       =   8
         PasswordChar    =   "*"
         TabIndex        =   5
         Top             =   720
         Width           =   1455
      End
      Begin VB.TextBox txt��Ա��� 
         Height          =   300
         Left            =   1320
         MaxLength       =   8
         TabIndex        =   11
         Top             =   1905
         Width           =   1455
      End
      Begin VB.TextBox txtסԺ���� 
         Height          =   300
         Left            =   4440
         MaxLength       =   5
         TabIndex        =   19
         Top             =   720
         Width           =   855
      End
      Begin VB.CommandButton cmdSelect 
         Caption         =   "��"
         Height          =   240
         Index           =   1
         Left            =   6240
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   2700
         Width           =   255
      End
      Begin VB.TextBox txt���� 
         Height          =   300
         Left            =   1320
         MaxLength       =   20
         TabIndex        =   7
         Top             =   1125
         Width           =   1455
      End
      Begin VB.CommandButton cmdSelect 
         Caption         =   "��"
         Height          =   240
         Index           =   2
         Left            =   6240
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   1935
         Width           =   255
      End
      Begin VB.TextBox txt����֤�� 
         Height          =   300
         Left            =   4440
         MaxLength       =   26
         TabIndex        =   28
         Top             =   2280
         Width           =   2085
      End
      Begin VB.ComboBox cmb���� 
         Height          =   300
         Left            =   4440
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   330
         Width           =   2085
      End
      Begin VB.ComboBox cmb�Ա� 
         Height          =   300
         Left            =   4440
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   1125
         Width           =   2085
      End
      Begin VB.TextBox txt���֤�� 
         Height          =   300
         Left            =   4440
         MaxLength       =   18
         TabIndex        =   23
         Top             =   1515
         Width           =   2085
      End
      Begin VB.ComboBox Cbo��ǰ״̬ 
         Height          =   300
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   2280
         Width           =   1455
      End
      Begin MSComCtl2.DTPicker dtp���� 
         Height          =   300
         Left            =   1320
         TabIndex        =   9
         Top             =   1515
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   87031811
         CurrentDate     =   36526
      End
      Begin VB.TextBox txt���� 
         Height          =   300
         Left            =   4440
         Locked          =   -1  'True
         TabIndex        =   30
         Top             =   2670
         Width           =   2085
      End
      Begin VB.TextBox txt��λ���� 
         Height          =   300
         Left            =   4440
         MaxLength       =   8
         TabIndex        =   25
         Top             =   1905
         Width           =   2085
      End
      Begin VB.Label lbl�ʻ���� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "�ʻ����(&L)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   240
         TabIndex        =   14
         Top             =   2730
         Width           =   990
      End
      Begin VB.Label lbl���� 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "����(&D)"
         Height          =   180
         Left            =   600
         TabIndex        =   1
         Top             =   390
         Width           =   630
      End
      Begin VB.Label lbl���� 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "����(&P)"
         Height          =   180
         Left            =   600
         TabIndex        =   4
         Top             =   780
         Width           =   630
      End
      Begin VB.Label lbl����֤�� 
         AutoSize        =   -1  'True
         Caption         =   "����֤��(&Z)"
         Height          =   180
         Left            =   3360
         TabIndex        =   27
         Top             =   2340
         Width           =   990
      End
      Begin VB.Label lbl��Ա��� 
         AutoSize        =   -1  'True
         Caption         =   "��Ա���(&E)"
         Height          =   180
         Left            =   240
         TabIndex        =   10
         Top             =   1965
         Width           =   990
      End
      Begin VB.Label lbl��λ���� 
         AutoSize        =   -1  'True
         Caption         =   "��λ����(&U)"
         Height          =   180
         Left            =   3360
         TabIndex        =   24
         Top             =   1965
         Width           =   990
      End
      Begin VB.Label lbl���� 
         AutoSize        =   -1  'True
         Caption         =   "����(&F)"
         Height          =   180
         Left            =   3720
         TabIndex        =   29
         Top             =   2730
         Width           =   630
      End
      Begin VB.Label lblסԺ���� 
         AutoSize        =   -1  'True
         Caption         =   "סԺ����(&S)"
         Height          =   180
         Left            =   3360
         TabIndex        =   18
         Top             =   780
         Width           =   990
      End
      Begin VB.Label lbl���� 
         AutoSize        =   -1  'True
         Caption         =   "����(&N)"
         Height          =   180
         Left            =   600
         TabIndex        =   6
         Top             =   1185
         Width           =   630
      End
      Begin VB.Label lbl���֤�� 
         AutoSize        =   -1  'True
         Caption         =   "���֤��(&I)"
         Height          =   180
         Left            =   3360
         TabIndex        =   22
         Top             =   1575
         Width           =   990
      End
      Begin VB.Label lblҽ������ 
         AutoSize        =   -1  'True
         Caption         =   "ҽ������(&R)"
         Height          =   180
         Left            =   3360
         TabIndex        =   16
         Top             =   390
         Width           =   990
      End
      Begin VB.Label lbl�Ա� 
         AutoSize        =   -1  'True
         Caption         =   "�Ա�(&X)"
         Height          =   180
         Left            =   3720
         TabIndex        =   20
         Top             =   1185
         Width           =   630
      End
      Begin VB.Label lbl�������� 
         AutoSize        =   -1  'True
         Caption         =   "��������(&B)"
         Height          =   180
         Left            =   240
         TabIndex        =   8
         Top             =   1575
         Width           =   990
      End
      Begin VB.Label lbl��ǰ״̬ 
         AutoSize        =   -1  'True
         Caption         =   "��Ա���(&K)"
         Height          =   180
         Left            =   240
         TabIndex        =   12
         Top             =   2340
         Width           =   990
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   5760
      TabIndex        =   33
      Top             =   3450
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   4350
      TabIndex        =   32
      Top             =   3450
      Width           =   1100
   End
End
Attribute VB_Name = "frmIdentifyüɽ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
#Const gverControl = 99  ' 0-��֧�ֶ�̬ҽ��(9.19��ǰ),1-֧�ֶ�̬ҽ���޸��Ӳ���(9.22��ǰ) , _

Private Enum ѡ��Enum
    Select���� = 0
    Select���� = 1
    Select��λ = 2
End Enum

Dim mstrIdentify As String
Dim mbytType As Byte        '0-����;1-סԺ;2-����������סԺ;3-����
Dim strNewPass As String
Dim mlng����ID As Long

Public Function ShowCard(Optional bytType As Byte, Optional lng����ID As Long) As String
'���ܣ�����ҽ�����˵������Ϣ
'������bytType-ʶ�����ͣ�0-���1-סԺ
'���أ�
    '0����;1ҽ����;2����;3����;4�Ա�;5��������;6���֤;7��λ����(����);
    '8.���Ĵ���;9.˳���;10��Ա���;11�ʻ����;12��ǰ״̬;13����ID;14��ְ(0,1);15����֤��;16�����;17�Ҷȼ�
    '18�ʻ������ۼ�,19�ʻ�֧���ۼ�,20����ͳ���ۼ�,21ͳ�ﱨ���ۼ�,22סԺ�����ۼ�,23��������
    Dim rsTemp As New ADODB.Recordset
    mbytType = bytType
    mlng����ID = lng����ID
    mstrIdentify = ""
    
    cmb�Ա�.Clear
    gstrSQL = "select ����,���� from �Ա� order by ����"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    
    Do Until rsTemp.EOF
        cmb�Ա�.AddItem rsTemp("����") & "." & rsTemp("����")
        rsTemp.MoveNext
    Loop
    
    cmb����.Clear
    gstrSQL = "select A.��������,B.���,B.����,B.���� from ������� A,��������Ŀ¼ B where A.���=[1] and A.���=b.����"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, TYPE_�Ĵ�üɽ)
    
    If rsTemp("��������") = 0 Then
        lblҽ������.Visible = False
        cmb����.Visible = False
        cmb����.AddItem "1.����" '������
    End If
    Do Until rsTemp.EOF
        cmb����.AddItem rsTemp("����") & "." & rsTemp("����")
        cmb����.ItemData(cmb����.NewIndex) = rsTemp("���")
        rsTemp.MoveNext
    Loop
    cmb����.ListIndex = 0
    
    '1-��ְ;2-����;3-����
    gstrSQL = "Select ���,���� From ������Ⱥ Where ����=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, TYPE_�Ĵ�üɽ)
    Cbo��ǰ״̬.Clear
    Do While Not rsTemp.EOF
        Cbo��ǰ״̬.AddItem rsTemp!����
        Cbo��ǰ״̬.ItemData(Cbo��ǰ״̬.NewIndex) = rsTemp!���
        rsTemp.MoveNext
    Loop
    Cbo��ǰ״̬.ListIndex = 0
    cmb�Ա�.ListIndex = 0
        
    dtp����.MaxDate = zlDatabase.Currentdate
    Call SetEnable(mbytType <> 2)
    txt�ʻ����.Enabled = (mlng����ID = 0 And mbytType = 2)
    If mlng����ID <> 0 Then
        txt���� = "A" & mlng����ID
        Call txt����_KeyPress(vbKeyReturn)
    End If
    frmIdentifyüɽ.Show vbModal
    ShowCard = mstrIdentify
End Function

Private Sub SetEnable(ByVal blnEnable As Boolean)
    cmdChange.Visible = blnEnable
    txt����.Enabled = Not blnEnable
    dtp����.Enabled = Not blnEnable
    txt��Ա���.Enabled = Not blnEnable
    Cbo��ǰ״̬.Enabled = Not blnEnable
    cmb����.Enabled = Not blnEnable
    txtסԺ����.Enabled = Not blnEnable
    cmb�Ա�.Enabled = Not blnEnable
    txt���֤��.Enabled = Not blnEnable
    txt��λ����.Enabled = Not blnEnable
    txt����֤��.Enabled = Not blnEnable
    Me.cmdSelect(Select��λ).Enabled = Not blnEnable
    
    txt����.PasswordChar = IIf(blnEnable = False, "", "*")
End Sub

Private Sub Cbo��ǰ״̬_Click()
    txt����֤��.Enabled = (Cbo��ǰ״̬.ListIndex = 1 Or Cbo��ǰ״̬.ListIndex = 2)
End Sub

Private Sub cmb����_Click()
    Dim lng���ų��� As Long, lng����֤���� As Long
    Dim rsTemp As New ADODB.Recordset
    
    'ȱʡֵ
    lng���ų��� = 20
    lng����֤���� = 26
    
    gstrSQL = "select ������,����ֵ from ���ղ��� where ����=[1] and ����=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, TYPE_�Ĵ�üɽ, CInt(cmb����.ItemData(cmb����.ListIndex)))
    
    Do Until rsTemp.EOF
        Select Case rsTemp("������")
            Case "���ų���"
                If IsNull(rsTemp("����ֵ")) = False Then lng���ų��� = Val(rsTemp("����ֵ"))
            Case "����֤����"
                If IsNull(rsTemp("����ֵ")) = False Then lng����֤���� = Val(rsTemp("����ֵ"))
        End Select
        rsTemp.MoveNext
    Loop
    
    txt����.MaxLength = lng���ų���
    txt����֤��.MaxLength = lng����֤����
End Sub

Private Sub cmdCancel_Click()
    mstrIdentify = ""
    Unload Me
End Sub

Private Sub cmdChange_Click()
    If IsValid() = False Then Exit Sub
    With frm�޸�����
        strNewPass = .ChangePassword(txt����.Text)
    End With
End Sub

Private Sub cmdOK_Click()
    Dim rsTemp As New ADODB.Recordset
    Dim strIdentify As String, strAddition As String
    Dim lng�䶯ID As Long, lng����ID As Long, lng���� As Long
    Dim blnTrans As Boolean
    Dim curͳ���ۼ� As Currency
    Dim lng��� As Long
    Dim str�Ƚϴ� As String
    
    '���������ݵ���ȷ��
    If IsValid() = False Then
        Exit Sub
    End If
    
    '�õ��������
    If cmb����.Visible = False Then
        lng���� = 0
    Else
        If cmb����.ListIndex < 0 Then
            MsgBox "��ѡ��������ҽ�����ġ�", vbInformation, gstrSysName
            cmb����.SetFocus
            Exit Sub
        End If
        lng���� = cmb����.ItemData(cmb����.ListIndex)
    End If
    
    '��鲡��״̬
    gstrSQL = "select nvl(��ǰ״̬,0) as ״̬,�Ҷȼ�,��ע from �����ʻ� where ����=[1] and ����=[2] and ҽ����=[3]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, TYPE_�Ĵ�üɽ, lng����, CStr(Trim(txt����.Text)))
    
    If rsTemp.RecordCount > 0 Then
        If rsTemp("״̬") > 0 Then
            MsgBox "�ò����Ѿ���Ժ������ͨ�������֤��", vbInformation, gstrSysName
            Exit Sub
        End If
        Select Case Nvl(rsTemp!�Ҷȼ�, 0)
        Case 1
            MsgBox "��ҽ�����Ѿ�����������ʹ�ã�" & IIf(Nvl(rsTemp!��ע) <> "", "��" & rsTemp!��ע & "��", ""), vbInformation, gstrSysName
            Exit Sub
        Case 9
            MsgBox "��ҽ�����Ѿ�����������ʹ�ã�", vbInformation, gstrSysName
            Exit Sub
        End Select
    End If
    
    '������ݿ��е������Ƿ���ȷ
    If Not ����ʻ���Ϣ_����(txt����.Text) Then Exit Sub
    
    If strNewPass = "" Then strNewPass = Trim(txt����.Text)
    '��ȡ�����ͳ���ۼ�
    lng��� = Format(zlDatabase.Currentdate, "yyyy")
    gstrSQL = " Select Nvl(����ͳ���ۼ�,0) ͳ���ۼ� From �ʻ������Ϣ " & _
              " Where ���=" & lng��� & " And ����ID =" & _
              "     (Select ����ID From �����ʻ� where ����=[1] and ����=[2] and ҽ����=[3])"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ���ͳ���ۼ�", TYPE_�Ĵ�üɽ, lng����, CStr(Trim(txt����.Text)))
    If Not rsTemp.EOF Then
        curͳ���ۼ� = rsTemp!ͳ���ۼ�
    Else
        curͳ���ۼ� = 0
    End If
    '�����ַ���
    '�������˵�����Ϣ�������ʽ��
    '0����;1ҽ����;2����;3����;4�Ա�;5��������;6���֤;7��λ����(����);
    '8.���Ĵ���;9.˳���;10��Ա���;11�ʻ����;12��ǰ״̬;13����ID;14��ְ(0,1);15����֤��;16�����;17�Ҷȼ�
    '18�ʻ������ۼ�,19�ʻ�֧���ۼ�,20����ͳ���ۼ�,21ͳ�ﱨ���ۼ�,22סԺ�����ۼ�,23��������
    strIdentify = Trim(txt����.Text)                       '0����
    strIdentify = strIdentify & ";" & Trim(txt����.Text)   '1ҽ����
    strIdentify = strIdentify & ";" & strNewPass   '2����
    strIdentify = strIdentify & ";" & Trim(txt����.Text)   '3����
    strIdentify = strIdentify & ";" & Replace(GetTextFromCombo(cmb�Ա�, True), "'", "") '4�Ա�
    strIdentify = strIdentify & ";" & Format(dtp����.Value, "yyyy-MM-dd") '5��������
    strIdentify = strIdentify & ";" & Trim(txt���֤��.Text)    '6���֤
    strIdentify = strIdentify & ";" & Trim(txt��λ����.Text) & "(" & Trim(txt��λ����.Text) & ")"  '7.��λ����(����)
    strAddition = ";" & lng����                                 '8.���Ĵ���
    strAddition = strAddition & ";"                             '9.˳���
    strAddition = strAddition & ";" & Trim(txt��Ա���.Text)       '10��Ա���
    strAddition = strAddition & ";" & IIf(txt�ʻ����.Enabled, Val(txt�ʻ����.Text), "NULL")  '11�ʻ����
    strAddition = strAddition & ";0"                            '12��ǰ״̬
    strAddition = strAddition & ";" & txt����.Tag     '13����ID
    strAddition = strAddition & ";" & Cbo��ǰ״̬.ItemData(Cbo��ǰ״̬.ListIndex) '14��ְ(1,2,3)
    strAddition = strAddition & ";" & Trim(txt����֤��.Text) '15����֤��
    strAddition = strAddition & ";" & DateDiff("yyyy", dtp����.Value, dtp����.MaxDate) '16�����
    strAddition = strAddition & ";"                             '17�Ҷȼ�
    strAddition = strAddition & ";" & Val(txt�ʻ����.Text)       '18�ʻ������ۼ�
    strAddition = strAddition & ";0"      '19�ʻ�֧���ۼ�
    strAddition = strAddition & ";" & curͳ���ۼ�      '20����ͳ���ۼ�
    strAddition = strAddition & ";0"       '21ͳ�ﱨ���ۼ�
    strAddition = strAddition & ";" & txtסԺ����.Text      '22סԺ�����ۼ�
    strAddition = strAddition & ";"                                             '23�������� (1����������)
    
    gcnOracle.BeginTrans
    blnTrans = True
    lng����ID = BuildPatiInfo_üɽ(mbytType, strIdentify & strAddition, mlng����ID)
    '���ظ�ʽ:�м���벡��ID
    If lng����ID > 0 Then
        mstrIdentify = strIdentify & ";" & lng����ID & strAddition
    End If
    
    '������½���������Ҫ�����ʻ��䶯��¼
    If txt�ʻ����.Enabled Then
        Call ����ʻ���Ϣ_����(txt����.Text, True)
        lng�䶯ID = zlDatabase.GetNextID("�ʻ��䶯��¼")
        gstrSQL = "ZL_�ʻ��䶯��¼_INSERT(" & _
                 lng�䶯ID & "," & TYPE_�Ĵ�üɽ & ",1," & lng����ID & "," & _
                 Val(txt�ʻ����.Text) & ",'" & gstrUserName & "','����ҽ�����˵���ʱ¼��ĳ�ʼֵ',1)"
        Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    End If
    '���¸����ʻ��е���Ϣ
    gstrSQL = "zl_�����ʻ�_������Ϣ(" & mlng����ID & "," & TYPE_�Ĵ�üɽ & ",'����','''" & strNewPass & "''')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "��������")
    If mbytType = 4 Then
        Call ����ʻ���Ϣ_����(txt����.Text, True)
    End If
    
    gcnOracle.CommitTrans
    
    '��ӡ��Ƭ
    If txt�ʻ����.Enabled Or mbytType = 4 Then
        Call zl9Report.ReportOpen(gcnOracle, glngSys, "ZL1_INSIDE_1604", Me, "����=" & 25, "ҽ����=" & txt����, 2)
    End If
    
    Unload Me
    Exit Sub
errHand:
    If ErrCenter = 1 Then Resume
    If blnTrans Then gcnOracle.RollbackTrans
End Sub

Private Sub cmdSelect_Click(Index As Integer)
    Dim rsTemp As ADODB.Recordset
    
    Select Case Index
        Case Select����
            gstrSQL = " Select A.����ID as ID,A.����,A.ҽ����,'******' ����,B.����,B.�Ա�,B.��������,B.���֤��,C.��� as ����ID " & _
                    " ,A.��Ա���,A.��λ����,A.����ID,D.���� as ����,A.��ְ as ��ְID,A.����֤��,A.�ʻ����" & _
                    " From �����ʻ� A,������Ϣ B,��������Ŀ¼ C,���ղ��� D" & _
                    "  where A.����ID=B.����ID And Nvl(A.�Ҷȼ�,0)<>9 And A.����=" & TYPE_�Ĵ�üɽ & _
                    "  and A.����=C.���� and A.����=C.��� and A.����ID=D.ID(+)"
            
            Call Get�ʻ����
            zlControl.TxtSelAll txt����
            If txt����.Enabled Then txt����.SetFocus
        Case Select��λ
            Set rsTemp = frmPubSel.ShowSelect(Me, _
                    " Select ID,�ϼ�ID,ĩ��,����,����,��ַ,�绰,��������,�ʺ�,��ϵ�� From ��Լ��λ" & _
                    " Start With �ϼ�ID is NULL Connect by Prior ID=�ϼ�ID", _
                    2, "������λ", , txt��λ����.Text)
            
            If Not rsTemp Is Nothing Then
                txt��λ����.Text = rsTemp("����")
                zlControl.TxtSelAll txt��λ����
            End If
            txt��λ����.SetFocus
        Case Select����
            gstrSQL = " Select A.ID,A.����,A.����,A.����,decode(A.���,1,'���Բ�',2,'���ֲ�','��ͨ��') as ��� " & _
                    " From ���ղ��� A where A.����=" & TYPE_�Ĵ�üɽ
            
            Set rsTemp = frmPubSel.ShowSelect(Me, gstrSQL, 0, "ҽ������", , txt����.Text)
            If Not rsTemp Is Nothing Then
                txt����.Text = rsTemp("����")
                txt����.Tag = rsTemp("ID")
                zlControl.TxtSelAll txt����
            End If
            txt����.SetFocus
    End Select
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey (vbKeyTab)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub txt����_GotFocus()
    zlControl.TxtSelAll txt����
End Sub

Private Sub txt����_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        txt����.Text = ""
        txt����.Tag = ""
    End If
End Sub

Private Sub txt��λ����_GotFocus()
    zlControl.TxtSelAll txt��λ����
End Sub

Private Sub txt����_GotFocus()
    zlCommFun.OpenIme True
    zlControl.TxtSelAll txt����
End Sub

Private Sub txt����_KeyPress(KeyAscii As Integer)
    Dim strCode As String
    Dim str���� As String
    Dim rsTemp As New ADODB.Recordset
    
    If Len(txt����.Text) = txt����.MaxLength Or KeyAscii = vbKeyReturn Then
        strCode = UCase(Replace(Trim(txt����.Text), "'", ""))
        If strCode = "" Then Exit Sub
        
        If IsNumeric(Mid(strCode, 1, Len(strCode) - 1)) Then 'ˢ��
            str���� = " and A.����='" & strCode & "'"
        ElseIf (Left(strCode, 1) = "A" Or Left(strCode, 1) = "-") And IsNumeric(Mid(strCode, 2)) Then '����ID
            str���� = " and A.����ID=" & Mid(strCode, 2)
        ElseIf (Left(strCode, 1) = "B" Or Left(strCode, 1) = "+") And IsNumeric(Mid(strCode, 2)) Then 'סԺ��(��ס(��)Ժ�Ĳ���)
            str���� = " and B.סԺ��=" & Mid(strCode, 2)
        ElseIf (Left(strCode, 1) = "D" Or Left(strCode, 1) = "*") And IsNumeric(Mid(strCode, 2)) Then '�����(�������ﲡ��)
            str���� = " and B.�����=" & Mid(strCode, 2)
        Else '��������
            str���� = " and A.����='" & strCode & "'"
        End If
    
        gstrSQL = " Select A.����ID as ID,A.����,A.ҽ����,'******' ����,B.����,B.�Ա�,B.��������,B.���֤��,C.��� as ����ID " & _
                " ,A.��Ա���,A.��λ����,A.����ID,D.���� as ����,A.��ְ as ��ְID,A.����֤��,A.�ʻ����" & _
                " From �����ʻ� A,������Ϣ B,��������Ŀ¼ C,���ղ��� D" & _
                "  where A.����ID=B.����ID And Nvl(A.�Ҷȼ�,0)<>9 And A.����=" & TYPE_�Ĵ�üɽ & _
                "  and A.����=C.���� and A.����=C.��� and A.����ID=D.ID(+)" & str����
        
        Call Get�ʻ����
    End If
End Sub

Private Sub txt����_LostFocus()
    zlCommFun.OpenIme False
End Sub

Private Sub txt��Ա���_GotFocus()
    zlControl.TxtSelAll txt��Ա���
End Sub

Private Sub txt���֤��_GotFocus()
    zlControl.TxtSelAll txt���֤��
End Sub

Private Sub txt����֤��_GotFocus()
    zlControl.TxtSelAll txt����
End Sub

Private Sub txt����_GotFocus()
    zlControl.TxtSelAll txt����
End Sub

Private Sub txt����_GotFocus()
    zlControl.TxtSelAll txt����
End Sub

Private Sub txt�ʻ����_LostFocus()
    txt�ʻ���� = Format(txt�ʻ����, "#####0.00;-#####0.00; ;")
End Sub

Private Sub txtסԺ����_GotFocus()
    zlControl.TxtSelAll txtסԺ����
End Sub

Private Sub Get�ʻ����()
'���Ѿ����ڵļ�¼�ж����ʻ���Ϣ
    Dim rs�ʻ� As ADODB.Recordset
    Dim rsTemp As New ADODB.Recordset
    Dim lngIndex As Long
    
    Set rs�ʻ� = frmPubSel.ShowSelect(Me, gstrSQL, 0, "�����ʻ�", , txt����.Text, "", False, True)
    If Not rs�ʻ� Is Nothing Then
        If mbytType = 2 Or mbytType = 4 Then txt����.Enabled = False
        txt����.Text = rs�ʻ�("����")
        '�������õ�����
        If mbytType = 2 Or mbytType = 4 Then
            gstrSQL = "Select ���� From �����ʻ� Where ����=[1] And ҽ����=[2]"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, TYPE_�Ĵ�üɽ, CStr(txt����.Text))
            
            txt����.Text = Nvl(rsTemp!����, "")
            txt����.PasswordChar = "*"
            If mbytType = 4 Then
                If IsNumeric(Right(txt����.Text, 1)) Then
                    txt����.Text = txt����.Text & "A"
                Else
                    txt����.Text = Mid(txt����.Text, 1, Len(txt����.Text) - 1) & Chr(asc(Right(txt����.Text, 1)) + 1)
                End If
            End If
        End If
        txt����.Text = IIf(IsNull(rs�ʻ�("����")), "", rs�ʻ�("����"))
        txt���֤��.Text = IIf(IsNull(rs�ʻ�("���֤��")), "", rs�ʻ�("���֤��"))
        txt��Ա���.Text = IIf(IsNull(rs�ʻ�("��Ա���")), "", rs�ʻ�("��Ա���"))
        txt��λ����.Text = IIf(IsNull(rs�ʻ�("��λ����")), "", rs�ʻ�("��λ����"))
        txt����.Text = IIf(IsNull(rs�ʻ�("����")), "", rs�ʻ�("����"))
        txt����.Tag = IIf(IsNull(rs�ʻ�("����ID")), "", rs�ʻ�("����ID"))
        
        Call SetComboByText(cmb�Ա�, IIf(IsNull(rs�ʻ�("�Ա�")), "", rs�ʻ�("�Ա�")), True)
        Cbo��ǰ״̬.ListIndex = rs�ʻ�("��ְID") - 1
        txt����֤��.Text = ""
        txt����֤��.Text = IIf(IsNull(rs�ʻ�("����֤��")), "", rs�ʻ�("����֤��"))
        If IsNull(rs�ʻ�("��������")) = False Then
            dtp����.Value = rs�ʻ�("��������")
        End If
        
        For lngIndex = 0 To cmb����.ListCount - 1
            If cmb����.ItemData(lngIndex) = rs�ʻ�("����ID") Then
                cmb����.ListIndex = lngIndex
                Exit For
            End If
        Next
        txt�ʻ���� = Format(rs�ʻ�!�ʻ����, "#####0.00;-#####0.00; ;")
        txt�ʻ����.Enabled = False
        
        '�ٶ����ʻ������Ϣ
        gstrSQL = "select * from �ʻ������Ϣ where ����=[1] and ����ID=[2] and ���=[3]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, TYPE_�Ĵ�üɽ, CLng(rs�ʻ�("ID")), Year(dtp����.MaxDate))
        
        If rsTemp.EOF = False Then
            '�����ʻ����
            txtסԺ����.Text = Nvl(rsTemp("סԺ�����ۼ�"), 0) & "/" & Nvl(rsTemp("��ԺסԺ����"), 0)
        Else
            txtסԺ����.Text = "0/0"
        End If
    End If
End Sub

Private Function IsValid() As Boolean
'���ܣ�������ݵ���ȷ��
    Dim lngIndex As Long
    Dim rsTemp As New ADODB.Recordset
    
    If Trim(txt����) = "" Then
        MsgBox "���Ų���Ϊ�գ�", vbInformation, gstrSysName
        If txt����.Enabled Then txt����.SetFocus
        Exit Function
    End If
    If Trim(txt����) = "" Then
        MsgBox "��������Ϊ�գ�", vbInformation, gstrSysName
        If txt����.Enabled Then txt����.SetFocus
        Exit Function
    End If
    If Trim(txt�ʻ����) <> "" Then
        If Not IsNumeric(txt�ʻ����) Then
            MsgBox "�ʻ�����к��зǷ��ַ���", vbInformation, gstrSysName
            If txt�ʻ����.Enabled Then txt�ʻ����.SetFocus
            Exit Function
        End If
        If Val(txt�ʻ����.Text) < 0 Then
            MsgBox "�ʻ�����С���㣡", vbInformation, gstrSysName
            If txt�ʻ����.Enabled Then txt�ʻ����.SetFocus
            Exit Function
        End If
    End If
    If cmb����.ListIndex < 0 Then
        MsgBox "����Ҫѡ��һ��ҽ�����ģ�", vbInformation, gstrSysName
        If cmb����.Enabled Then cmb����.SetFocus
        Exit Function
    End If
    If UBound(Split(txtסԺ����.Text, "/")) <> 1 Then
        MsgBox "�����뱾ԺסԺ��������ԺסԺ����������ʽ����Ժ/��Ժ���磺1/1��", vbInformation, gstrSysName
        txtסԺ����.SetFocus
        Exit Function
    End If
    If Trim(txt���֤��.Text) <> "" Then
        If Not IsNumeric(txt���֤��) Then
            MsgBox "���֤���к��зǷ��ַ���", vbInformation, gstrSysName
            If txt���֤��.Enabled Then txt���֤��.SetFocus
            Exit Function
        End If
    End If
    
    'У��������ȷ��
    gstrSQL = "Select ���� From �����ʻ� Where ����=[1] And ҽ����=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, TYPE_�Ĵ�üɽ, Trim(txt����.Text))
    If Not rsTemp.EOF Then
        If Nvl(rsTemp!����, "") <> txt����.Text Then
            MsgBox "����������飡", vbInformation, gstrSysName
            If txt����.Enabled Then txt����.SetFocus
            Exit Function
        End If
    End If
    
    IsValid = True
End Function

Public Function BuildPatiInfo_üɽ(ByVal bytType As Byte, ByVal strInfo As String, ByVal lng����ID As Long) As Long
'���ܣ����������ʻ���Ϣ
'������bytType=0-����,1-סԺ
'      strInfo='0����;1ҽ����;2����;3����;4�Ա�;5��������;6���֤;7��λ����(����);
'      8����;9.˳���;10��Ա���;11�ʻ����;12��ǰ״̬;13����ID;14��ְ(1,2,3);15����֤��;16�����;17�Ҷȼ�
'      18�ʻ������ۼ�;19�ʻ�֧���ۼ�;20����ͳ���ۼ�;21ͳ�ﱨ���ۼ�;22סԺ�����ۼ�;23�������
'      24��������;25�����ۼ�;26����ͳ���޶�
'���أ�����ID
    Const MAX_BOUND = 26 'Ҫ�������Ϣ����
    
    Dim rsPati As ADODB.Recordset, str��λ���� As String, lng���� As Long
    Dim rsTemp As New ADODB.Recordset
    Dim strSQL As String, curDate As Date
    Dim lng���� As Long, array��Ϣ As Variant
    
    On Error GoTo errHandle
    
    If Len(Trim(strInfo)) <> 0 Then
        curDate = zlDatabase.Currentdate
        
        '200308z012:��֤�������Ϣ������
        If UBound(Split(strInfo, ";")) < MAX_BOUND Then
            strInfo = strInfo & String(MAX_BOUND - UBound(Split(strInfo, ";")), ";")
        End If
        array��Ϣ = Split(strInfo, ";")
        
        '�ӵ�7��������ȡ����λ����
        If array��Ϣ(7) Like "*(*" Then
            str��λ���� = Split(array��Ϣ(7), "(")(UBound(Split(array��Ϣ(7), "(")))
            str��λ���� = Mid(str��λ����, 1, Len(str��λ����) - 1)
        End If
        'ȡ����
        If IsDate(array��Ϣ(5)) Then
            lng���� = Int(curDate - CDate(array��Ϣ(5))) / 365
        End If
        
        lng���� = Val(array��Ϣ(8))
        
        If lng����ID > 0 Then
            '�ò����Ѿ�����
            gstrSQL = "Select nvl(����ID,0) ����ID from �����ʻ� where ҽ����=[1] and ����=[2]"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "�����ʻ�", CStr(array��Ϣ(1)), TYPE_�Ĵ�üɽ)
            If rsTemp.EOF = False Then
                If rsTemp("����ID") <> lng����ID Then
                    MsgBox "�Ѿ�������ͬҽ���ŵ�����һλ���ˣ����ڲ��˹�������н���λ�ϲ���", vbInformation, gstrSysName
                    Exit Function
                End If
            End If
        End If
        
        '�ʻ�Ψһ������,����,ҽ����
        #If gverControl < 6 Then
            strSQL = "Select A.*,B.ҽ���� From ������Ϣ A," & _
                " (Select * From �����ʻ�" & _
                " Where ����=" & TYPE_�Ĵ�üɽ & _
                " And ҽ����='" & CStr(array��Ϣ(1)) & "') B" & _
                " Where " & IIf(lng����ID = 0, "A.����ID=B.����ID", "A.����ID=B.����ID(+) and A.����ID=" & lng����ID) '���ܲ���ID�Ѿ�ȷ��
        #Else
            strSQL = "Select A.����id, A.�����, A.סԺ��, A.���￨��, A.����֤��, A.�ѱ�, A.ҽ�Ƹ��ʽ, A.����, A.�Ա�, A.����, A.��������, A.�����ص�, A.���֤��, A.����֤��," & vbNewLine & _
                "A.���, A.ְҵ, A.����, A.����, A.����, A.ѧ��, A.����״��, A.��ͥ��ַ,A.��ͥ�绰, A.��ͥ��ַ�ʱ� As �����ʱ�, A.�໤��, A.��ϵ������, A.��ϵ�˹�ϵ, A.��ϵ�˵�ַ, " & vbNewLine & _
                "A.��ϵ�˵绰, A.��ͬ��λid, A.������λ, A.��λ�绰, A.��λ�ʱ�, A.��λ������, A.��λ�ʺ�, A.������, A.������, A.��������, A.����ʱ��, A.����״̬,A.��������, A.סԺ����," & vbNewLine & _
                "A.��ǰ����id, A.��ǰ����id, A.��ǰ����, A.��Ժʱ��, A.��Ժʱ��, A.��Ժ, A.Ic����, A.������, A.ҽ����, A.����, A.��ѯ����, A.�Ǽ�ʱ��, A.ͣ��ʱ��, A.����,B.ҽ���� From ������Ϣ A," & _
                " (Select * From �����ʻ�" & _
                " Where ����=[1]" & _
                " And ҽ����=[2]) B" & _
                " Where " & IIf(lng����ID = 0, "A.����ID=B.����ID", "A.����ID=B.����ID(+) and A.����ID=[3]")  '���ܲ���ID�Ѿ�ȷ��
        #End If
        Set rsPati = zlDatabase.OpenSQLRecord(strSQL, "ҽ���ӿ�", TYPE_�Ĵ�üɽ, CStr(array��Ϣ(1)), lng����ID)
        
        If rsPati.EOF Then
            '�ޱ����ʻ�����Ϊû�в�����Ϣ
            If lng����ID = 0 Then lng����ID = GetNextNO(1)
            strSQL = "zl_������Ϣ_Insert(" & lng����ID & ",NULL,NULL,'������ҽ�Ʊ���'," & _
                "'" & array��Ϣ(3) & "','" & array��Ϣ(4) & "'," & IIf(Val(array��Ϣ(16)) = 0, lng����, Val(array��Ϣ(16))) & "," & _
                "To_Date('" & Format(array��Ϣ(5), "yyyy-MM-dd") & "','YYYY-MM-DD')," & _
                "NULL,'" & array��Ϣ(6) & "',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL," & _
                "NULL,NULL,NULL,NULL,NULL,NULL,'" & array��Ϣ(7) & "',NULL,NULL,NULL," & _
                "NULL,NULL,NULL," & TYPE_�Ĵ�üɽ & "," & _
                "To_Date('" & Format(curDate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'))"
            Call SQLTest(App.ProductName, "ҽ���ӿ�", strSQL)
            gcnOracle.Execute strSQL, , adCmdStoredProc
            Call SQLTest
        Else
            '�в�����Ϣ�ͱ����ʻ���Ϣ
            If rsPati("����") <> array��Ϣ(3) Then
                If MsgBox("����ԭ�еǼǵ������� " & rsPati("����") & " ����ˢ���õ������� " & array��Ϣ(3) & " ������" & vbCrLf & _
                          "��������²���ԭ�еĵǼ���Ϣ���Ƿ�ȷ����", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
                    Exit Function
                End If
            End If
            If lng����ID = 0 Then lng����ID = rsPati!����ID
            
            strSQL = "zl_������Ϣ_Update(" & _
                lng����ID & "," & IIf(IsNull(rsPati!�����), "NULL", rsPati!�����) & "," & _
                IIf(IsNull(rsPati!סԺ��), "NULL", rsPati!סԺ��) & ",'" & IIf(IsNull(rsPati!�ѱ�), "", rsPati!�ѱ�) & "'," & _
                "'" & IIf(IsNull(rsPati!ҽ�Ƹ��ʽ), "", rsPati!ҽ�Ƹ��ʽ) & "'," & _
                "'" & array��Ϣ(3) & "','" & array��Ϣ(4) & "'," & IIf(Val(array��Ϣ(16)) = 0, lng����, Val(array��Ϣ(16))) & "," & _
                "To_Date('" & Format(array��Ϣ(5), "yyyy-MM-dd") & "','YYYY-MM-DD')," & _
                "'" & IIf(IsNull(rsPati!�����ص�), "", rsPati!�����ص�) & "','" & array��Ϣ(6) & "'," & _
                "'" & IIf(IsNull(rsPati!���), "", rsPati!���) & "','" & IIf(IsNull(rsPati!ְҵ), "", rsPati!ְҵ) & "'," & _
                "'" & IIf(IsNull(rsPati!����), "", rsPati!����) & "','" & IIf(IsNull(rsPati!����), "", rsPati!����) & "'," & _
                "'" & IIf(IsNull(rsPati!ѧ��), "", rsPati!ѧ��) & "','" & IIf(IsNull(rsPati!����״��), "", rsPati!����״��) & "'," & _
                "'" & IIf(IsNull(rsPati!��ͥ��ַ), "", rsPati!��ͥ��ַ) & "','" & IIf(IsNull(rsPati!��ͥ�绰), "", rsPati!��ͥ�绰) & "'," & _
                "'" & IIf(IsNull(rsPati!�����ʱ�), "", rsPati!�����ʱ�) & "','" & IIf(IsNull(rsPati!��ϵ������), "", rsPati!��ϵ������) & "'," & _
                "'" & IIf(IsNull(rsPati!��ϵ�˹�ϵ), "", rsPati!��ϵ�˹�ϵ) & "','" & IIf(IsNull(rsPati!��ϵ�˵�ַ), "", rsPati!��ϵ�˵�ַ) & "'," & _
                "'" & IIf(IsNull(rsPati!��ϵ�˵绰), "", rsPati!��ϵ�˵绰) & "'," & IIf(IsNull(rsPati!��ͬ��λID), "NULL", rsPati!��ͬ��λID) & "," & _
                "'" & array��Ϣ(7) & "','" & IIf(IsNull(rsPati!��λ�绰), "", rsPati!��λ�绰) & "'," & _
                "'" & IIf(IsNull(rsPati!��λ�ʱ�), "", rsPati!��λ�ʱ�) & "','" & IIf(IsNull(rsPati!��λ������), "", rsPati!��λ������) & "'," & _
                "'" & IIf(IsNull(rsPati!��λ�ʺ�), "", rsPati!��λ�ʺ�) & "','" & IIf(IsNull(rsPati!������), "", rsPati!������) & "'," & _
                "" & IIf(IsNull(rsPati!������), "NULL", rsPati!������) & "," & TYPE_�Ĵ�üɽ & ")"
            Call SQLTest(App.ProductName, "ҽ���ӿ�", strSQL)
            gcnOracle.Execute strSQL, , adCmdStoredProc
            Call SQLTest
        End If
        
        '�������±����ʻ���Ϣ(�Զ�)
        strSQL = "zl_�����ʻ�_insert(" & lng����ID & "," & TYPE_�Ĵ�üɽ & "," & _
            lng���� & "," & _
            "'" & IIf(array��Ϣ(0) = "-1", array��Ϣ(1), array��Ϣ(0)) & "'," & _
            "'" & array��Ϣ(1) & "'," & _
            "'" & array��Ϣ(2) & "'," & _
            "'" & array��Ϣ(9) & "'," & _
            "'" & array��Ϣ(15) & "'," & _
            "'" & array��Ϣ(10) & "'," & _
            "'" & str��λ���� & "'," & _
            array��Ϣ(11) & "," & _
            Val(array��Ϣ(12)) & "," & _
            IIf(Val(array��Ϣ(13)) = 0, "NULL", Val(array��Ϣ(13))) & "," & _
            IIf(Val(array��Ϣ(14)) = 0, 1, Val(array��Ϣ(14))) & "," & _
            IIf(Val(array��Ϣ(16)) = 0, lng����, Val(array��Ϣ(16))) & "," & _
            "'" & array��Ϣ(17) & "'," & _
            "To_Date('" & Format(curDate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'))"
        Call SQLTest(App.ProductName, "ҽ���ӿ�", strSQL)
        gcnOracle.Execute strSQL, , adCmdStoredProc
        Call SQLTest
        
        '���������ʻ������Ϣ(�Զ�)
        '200308z012:�ɶ�:����"24��������=zyjs,25�����ۼ�=tcbxbl,26����ͳ���޶�=zyxe"
        strSQL = "zl_�ʻ������Ϣ_Insert(" & lng����ID & "," & TYPE_�Ĵ�üɽ & "," & Year(curDate) & "," & _
            Val(array��Ϣ(18)) & "," & Val(array��Ϣ(19)) & "," & _
            Val(array��Ϣ(20)) & "," & Val(array��Ϣ(21)) & "," & _
            Val(Split(array��Ϣ(22), "/")(0)) & "," & Val(Split(array��Ϣ(22), "/")(1)) & "," & Val(array��Ϣ(24)) & "," & Val(array��Ϣ(25)) & "," & Val(array��Ϣ(26)) & ")"
        Call SQLTest(App.ProductName, "ҽ���ӿ�", strSQL)
        gcnOracle.Execute strSQL, , adCmdStoredProc
        Call SQLTest
    End If
    BuildPatiInfo_üɽ = lng����ID
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
End Function
