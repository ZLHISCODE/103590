VERSION 5.00
Begin VB.Form frm����ѡ��_ɽ�� 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "����ѡ��"
   ClientHeight    =   2220
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5025
   Icon            =   "frm����ѡ��_ɽ��.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2220
   ScaleWidth      =   5025
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��"
      Height          =   375
      Left            =   3465
      TabIndex        =   7
      Top             =   1635
      Width           =   1215
   End
   Begin VB.CommandButton cmd��Ժ��Ϣ 
      Caption         =   "��"
      Height          =   300
      Left            =   4425
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1050
      Width           =   285
   End
   Begin VB.CommandButton cmd��Ժ��Ϣ 
      Caption         =   "��"
      Height          =   300
      Left            =   4410
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   510
      Width           =   285
   End
   Begin VB.TextBox txt��Ժ���� 
      Height          =   300
      Left            =   1035
      TabIndex        =   4
      Top             =   1050
      Width           =   3390
   End
   Begin VB.TextBox txt��Ժ���� 
      Height          =   300
      Left            =   1035
      TabIndex        =   1
      Top             =   510
      Width           =   3390
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "ȷ��"
      Height          =   375
      Left            =   2130
      TabIndex        =   6
      Top             =   1650
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "��Ժ����"
      Height          =   270
      Left            =   240
      TabIndex        =   3
      Top             =   1110
      Width           =   870
   End
   Begin VB.Label Label1 
      Caption         =   "��Ժ����"
      Height          =   270
      Left            =   255
      TabIndex        =   0
      Top             =   570
      Width           =   870
   End
End
Attribute VB_Name = "frm����ѡ��_ɽ��"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Dim mstrSQL As String
Dim mrsTMP As New ADODB.Recordset
Dim mlng����ID As Long
Dim mblnOK As Boolean '�ǰ���ȷ��ť�˳�
Dim mstr��Ժ���ֱ��� As String
Dim mstr��Ժ�������� As String
Dim mstr��Ժ���ֱ��� As String
Dim mstr��Ժ�������� As String


Private Function �������ѡ��(Optional strDisName As String = "") As String
    Dim rsTmp As ADODB.Recordset
    Dim strTmpSQL As String
    
    If strDisName <> "" Then
        strTmpSQL = "select rownum as ID,aka120  ���ֱ���,aka121 ��������,aka066 ������,aae035 ������� from ka06" & _
                    " where aka120 like '%" & Trim(strDisName) & "%' or aka121 like '%" & Trim(strDisName) & "%' or Upper(aka066) like '%" & UCase(Trim(strDisName)) & "%'"
    Else
        strTmpSQL = "select rownum as ID,aka120  ���ֱ���,aka121 ��������,aka066 ������,aae035 ������� from ka06"
    End If
    Set rsTmp = frmPubSel.ShowSelect(Me, strTmpSQL, 0, "����", True, , , , , gcnSxDr)
    If rsTmp Is Nothing Then
        �������ѡ�� = "|"
        Exit Function
    End If
    �������ѡ�� = rsTmp!���ֱ��� & "|" & rsTmp!��������
End Function

Public Function Select����(lng����ID As Long, ByRef str��Ժ���ֱ���, ByRef str��Ժ��������, ByRef str��Ժ���ֱ���, ByRef str��Ժ��������) As Boolean

    mstrSQL = "Select * from ���ղ��� where ID=(select ����ID from �����ʻ� where ����ID=" & lng����ID & ")"
    Call OpenRecordset(mrsTMP, "��Ժ����", mstrSQL)
    If mrsTMP.EOF Then
        mstr��Ժ�������� = ""
    Else
        mstr��Ժ�������� = mrsTMP!����
        mstr��Ժ���ֱ��� = mrsTMP!����
    End If
    
    mstrSQL = "Select * from ���ղ��� where ID=(select ��Ժ����ID from �����ʻ� where ����ID=" & lng����ID & ")"
    Call OpenRecordset(mrsTMP, "��Ժ����", mstrSQL)
    If mrsTMP.EOF Then
        mstr��Ժ�������� = ""
    Else
        mstr��Ժ�������� = mrsTMP!����
        mstr��Ժ���ֱ��� = mrsTMP!����
    End If
    mlng����ID = lng����ID
    Select���� = mblnOK
    
    frm����ѡ��_ɽ��.Show 1
    
    str��Ժ���ֱ��� = mstr��Ժ���ֱ���
    str��Ժ�������� = mstr��Ժ��������
    str��Ժ���ֱ��� = mstr��Ժ���ֱ���
    str��Ժ�������� = mstr��Ժ��������
    Select���� = mblnOK
    
    Unload Me
End Function

Private Sub cmd��Ժ��Ϣ_Click()
  ''������ѡ����
    Dim strReturn As String
    strReturn = "|"

    strReturn = �������ѡ��
    If Trim(strReturn) <> "|" Then
        txt��Ժ����.Text = Split(strReturn, "|")(1)
        txt��Ժ����.Tag = Split(strReturn, "|")(0)
    End If
End Sub

Private Sub cmd��Ժ��Ϣ_Click()
  ''������ѡ����
    Dim strReturn As String
    strReturn = "|"
    strReturn = �������ѡ��
    If Trim(strReturn) <> "|" Then

        txt��Ժ����.Text = Split(strReturn, "|")(1)
        txt��Ժ����.Tag = Split(strReturn, "|")(0)
        
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{Tab}", 1
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    'ȥ��������
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    txt��Ժ����.Text = mstr��Ժ��������
    txt��Ժ����.Tag = mstr��Ժ���ֱ���
    txt��Ժ����.Text = mstr��Ժ��������
    txt��Ժ����.Tag = mstr��Ժ���ֱ���

End Sub

Private Sub OKButton_Click()

    Dim cur����ID  As Currency  '��currency�����׳��ֽ���δ֪����.
    Dim str���ּ��� As String
    
    If Trim(txt��Ժ����.Text) = "" Or Trim(txt��Ժ����.Text) = "" Then
        MsgBox "����ѡ���֣�", vbInformation, gstrSysName
        Exit Sub
    End If
    
        '���没����Ϣ�����ղ��ֱ���
      '�жϿ�����û���������,���У���ֱ��ȡ�ò���ID
    mstrSQL = "select * from ���ղ��� where ����=" & TYPE_ɽ�� & _
                                         " and ����='" & txt��Ժ����.Tag & "'"
    Call OpenRecordset(mrsTMP, "�鲡��ID", mstrSQL)
    If mrsTMP.EOF Then
        mstrSQL = "select ���ղ���_ID.NextVal as ID from Dual "
        Call OpenRecordset(mrsTMP, "ȡ����ID", mstrSQL)
        cur����ID = 1
        If Not mrsTMP.EOF Then cur����ID = mrsTMP!ID
        
        mstrSQL = "select zlspellcode('" & txt��Ժ����.Text & "') as ���� from dual"
        Call OpenRecordset(mrsTMP, "ȡ���ּ���", mstrSQL)
        str���ּ��� = mrsTMP!����
        
        mstrSQL = "zl_���ղ���_insert(" & cur����ID & "," & TYPE_ɽ�� & ",'" & _
                                         txt��Ժ����.Tag & "','" & _
                                         txt��Ժ����.Text & "','" & _
                                         str���ּ��� & "',1,NULL,NULL)"
        gcnOracle.Execute mstrSQL, , adCmdStoredProc
        
        
        mrsTMP.Close
        Set mrsTMP = Nothing
    Else
       cur����ID = mrsTMP!ID
    End If

    gstrSQL = " ZL_�����ʻ�_������Ϣ(" & mlng����ID & "," & TYPE_ɽ�� & ",'����ID','''" & cur����ID & "''')"
    Call zldatabase.ExecuteProcedure(gstrSQL, "����ID")

    mstrSQL = "select * from ���ղ��� where ����=" & TYPE_ɽ�� & _
                                         " and ����='" & txt��Ժ����.Tag & "'"
    Call OpenRecordset(mrsTMP, "�鲡��ID", mstrSQL)
    If mrsTMP.EOF Then
        mstrSQL = "select ���ղ���_ID.NextVal as ID from Dual "
        Call OpenRecordset(mrsTMP, "ȡ����ID", mstrSQL)
        cur����ID = 1
        If Not mrsTMP.EOF Then cur����ID = mrsTMP!ID
        
        mstrSQL = "select zlspellcode('" & txt��Ժ����.Text & "') as ���� from dual"
        Call OpenRecordset(mrsTMP, "ȡ���ּ���", mstrSQL)
        str���ּ��� = mrsTMP!����
        
        mstrSQL = "zl_���ղ���_insert(" & cur����ID & "," & TYPE_ɽ�� & ",'" & _
                                         txt��Ժ����.Tag & "','" & _
                                         txt��Ժ����.Text & "','" & _
                                         str���ּ��� & "',1,NULL,NULL)"
        gcnOracle.Execute mstrSQL, , adCmdStoredProc
        
        mrsTMP.Close
        Set mrsTMP = Nothing
    Else
       cur����ID = mrsTMP!ID
    End If
    gstrSQL = " ZL_�����ʻ�_������Ϣ(" & mlng����ID & "," & TYPE_ɽ�� & ",'��Ժ����ID','''" & cur����ID & "''')"
    Call zldatabase.ExecuteProcedure(gstrSQL, "����ID")
    
    mstr��Ժ�������� = txt��Ժ����.Text
    mstr��Ժ���ֱ��� = txt��Ժ����.Tag
    
    mstr��Ժ�������� = txt��Ժ����.Text
    mstr��Ժ���ֱ��� = txt��Ժ����.Tag
    mblnOK = True
    Unload Me
End Sub


Private Sub txt��Ժ����_KeyPress(KeyAscii As Integer)
  ''������ѡ����
    Dim strReturn As String
    
    If KeyAscii <> vbKeyReturn Then Exit Sub
    strReturn = �������ѡ��(Trim(txt��Ժ����.Text))
    If Trim(strReturn) <> "|" Then

    txt��Ժ����.Text = Split(strReturn, "|")(1)
    txt��Ժ����.Tag = Split(strReturn, "|")(0)
    End If
End Sub

Private Sub txt��Ժ����_KeyPress(KeyAscii As Integer)
  ''������ѡ����
    Dim strReturn As String
    If KeyAscii <> vbKeyReturn Then Exit Sub
    strReturn = �������ѡ��(Trim(txt��Ժ����.Text))
    If Trim(strReturn) <> "|" Then

    txt��Ժ����.Text = Split(strReturn, "|")(1)
    txt��Ժ����.Tag = Split(strReturn, "|")(0)
    End If
End Sub

