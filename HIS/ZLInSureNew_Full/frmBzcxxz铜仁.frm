VERSION 5.00
Begin VB.Form frmBzcxxzͭ�� 
   Caption         =   "��������ѡ��"
   ClientHeight    =   1365
   ClientLeft      =   4200
   ClientTop       =   4155
   ClientWidth     =   4710
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   1365
   ScaleWidth      =   4710
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ��"
      Height          =   495
      Left            =   3240
      TabIndex        =   3
      Top             =   720
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��"
      Height          =   495
      Left            =   3240
      TabIndex        =   2
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton cmd���� 
      Caption         =   "��"
      Height          =   375
      Left            =   2400
      TabIndex        =   1
      Top             =   600
      Width           =   375
   End
   Begin VB.TextBox txt���� 
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Top             =   600
      Width           =   2295
   End
   Begin VB.Label Label1 
      Caption         =   "�ò��˲����Ѷ�ʧ������ѡ��!"
      Height          =   375
      Left            =   480
      TabIndex        =   4
      Top             =   120
      Width           =   1935
   End
End
Attribute VB_Name = "frmBzcxxzͭ��"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private mintTimes As Integer
Private mblnOK As Boolean
Private mint���� As Integer
Private mlng����ID As Long
Private mstr���ֱ��� As String

Private Sub cmdCancel_Click()
txt����.Tag = ""
mlng����ID = 0
Unload Me
End Sub


Private Sub cmdOK_Click()
    mlng����ID = Val(txt����.Tag)
    mblnOK = True
    Unload Me
End Sub

Private Sub cmd����_Click()
    On Error GoTo errHandle
    Dim rs���� As ADODB.Recordset
    gstrSQL = " Select A.ID,A.����,A.����,A.����,decode(A.���,1,'���Բ�',2,'���ֲ�',0,'��ͨ��') as ��� " & _
            " From ���ղ��� A where A.����=[1] And A.��� IN ([2])"
    Set rs���� = zlDatabase.OpenSQLRecord(gstrSQL, "����ѡ����", TYPE_ͭ��, CStr(IIf(mint���� = 0, "0,1,2", "0")))
    
    If rs����.RecordCount > 0 Then
        If frmListSel.ShowSelect(TYPE_ͭ��, rs����, "ID", "ҽ������ѡ��", "��ѡ��ҽ�����֣�") = True Then
            txt����.Text = rs����("����")
            txt����.Tag = rs����("ID")
            mstr���ֱ��� = rs����("����")
        End If
    End If
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Public Function GetPatient(ByVal int���� As Integer, ByVal bln�޸����� As Boolean, ����ID As Long) As Boolean
    Me.Show vbModal
    If mblnOK = True Then
        ����ID = mlng����ID
    End If
    GetPatient = mblnOK
End Function

Private Sub txt����_KeyPress(KeyAscii As Integer)
    Dim rsTmp As New ADODB.Recordset, strText As String
    Dim blnReturn As Boolean, lst As ListItem
    Dim strǰ   As String
    
    On Error GoTo errHandle
    
    If KeyAscii <> vbKeyReturn Then Exit Sub
        txt����.Tag = ""
    If txt����.Text = "" Or txt����.Tag <> "" Then
        SendKeys "{TAB}"
        Exit Sub
    End If
    
    strǰ = IIf(gstrMatchMethod = 0, "%", "")
    strText = txt����.Text
    gstrSQL = "Select A.ID,A.����,A.����,A.����,decode(A.���,1,'���Բ�',2,'���ⲡ',0,'��ͨ��') ��� " & _
             "   FROM ���ղ��� A WHERE A.����=[1] And A.��� IN ([2]) And (" & _
             "   A.���� like '" & strǰ & "' || [3] || '%' or A.���� like '" & strǰ & "' || [3] || '%' or A.���� like '" & strǰ & "' || [3] || '%')"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, TYPE_ͭ��, IIf(mint���� = 0, "0,1,2", "0"), strText)
    
    If rsTmp.RecordCount > 0 Then
        '����ѡ����
        If rsTmp.RecordCount > 1 Then
            '�����ֶδ���3�ģ���ʹֻ��һ����¼�ѸöԻ�����ʾ�������Ա����û��õ��������Ϣ
            blnReturn = frmListSel.ShowSelect(TYPE_ͭ��, rsTmp, "ID", "ҽ������ѡ��", "��ѡ���ض���ҽ�����֣�")
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
        txt����.Text = rsTmp("����")
        txt����.Tag = rsTmp("ID")
        mstr���ֱ��� = rsTmp("����")
       ' SendKeys "{TAB}"
       cmdOK.SetFocus
    End If
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
