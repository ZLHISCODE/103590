VERSION 5.00
Begin VB.Form frmIdentify�������� 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "�����֤"
   ClientHeight    =   3615
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8010
   Icon            =   "frmIdentify��������.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3615
   ScaleWidth      =   8010
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.Frame fra���� 
      Caption         =   "���˻�����Ϣ"
      Height          =   2925
      Index           =   0
      Left            =   150
      TabIndex        =   0
      Top             =   120
      Width           =   7695
      Begin VB.TextBox txt���� 
         Enabled         =   0   'False
         Height          =   300
         Left            =   1200
         MaxLength       =   20
         TabIndex        =   2
         Top             =   285
         Width           =   2265
      End
      Begin VB.TextBox TxtIC��״̬ 
         Enabled         =   0   'False
         Height          =   300
         Left            =   1200
         MaxLength       =   19
         TabIndex        =   12
         Top             =   1980
         Width           =   2265
      End
      Begin VB.TextBox Txt�������� 
         Enabled         =   0   'False
         Height          =   300
         Left            =   4830
         MaxLength       =   19
         TabIndex        =   18
         Top             =   720
         Width           =   2625
      End
      Begin VB.TextBox Txt���������� 
         Enabled         =   0   'False
         Height          =   300
         Left            =   4830
         MaxLength       =   19
         TabIndex        =   20
         Top             =   1140
         Width           =   2625
      End
      Begin VB.TextBox Txt��λ���� 
         Enabled         =   0   'False
         Height          =   300
         Left            =   4830
         MaxLength       =   19
         TabIndex        =   16
         Top             =   285
         Width           =   2625
      End
      Begin VB.TextBox Txt����ۼ� 
         Enabled         =   0   'False
         Height          =   300
         Left            =   4830
         MaxLength       =   19
         TabIndex        =   26
         Top             =   1980
         Width           =   2625
      End
      Begin VB.TextBox Txt�ʻ���� 
         Enabled         =   0   'False
         Height          =   300
         Left            =   6030
         MaxLength       =   19
         TabIndex        =   24
         Top             =   1560
         Width           =   1425
      End
      Begin VB.TextBox TxtסԺ���� 
         Enabled         =   0   'False
         Height          =   300
         Left            =   4830
         MaxLength       =   19
         TabIndex        =   22
         Top             =   1560
         Width           =   405
      End
      Begin VB.TextBox Txt����״̬ 
         Enabled         =   0   'False
         Height          =   300
         Left            =   1200
         MaxLength       =   19
         TabIndex        =   14
         Top             =   2400
         Width           =   2265
      End
      Begin VB.TextBox txt���� 
         Enabled         =   0   'False
         Height          =   300
         Left            =   3030
         MaxLength       =   20
         TabIndex        =   6
         Top             =   720
         Width           =   435
      End
      Begin VB.TextBox txtҽ���� 
         Enabled         =   0   'False
         Height          =   300
         Left            =   1200
         MaxLength       =   19
         TabIndex        =   10
         Top             =   1560
         Width           =   2265
      End
      Begin VB.CommandButton cmd���� 
         Caption         =   "��"
         Height          =   240
         Left            =   7170
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   2430
         Width           =   255
      End
      Begin VB.ComboBox cob�Ա� 
         Enabled         =   0   'False
         Height          =   300
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   720
         Width           =   975
      End
      Begin VB.TextBox txt���� 
         Height          =   300
         Left            =   4830
         Locked          =   -1  'True
         TabIndex        =   28
         Top             =   2400
         Width           =   2625
      End
      Begin VB.TextBox txt���� 
         Enabled         =   0   'False
         Height          =   300
         Left            =   1200
         MaxLength       =   12
         TabIndex        =   8
         Top             =   1140
         Width           =   2265
      End
      Begin VB.Label lbl���� 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "����(&N)"
         Enabled         =   0   'False
         Height          =   180
         Left            =   480
         TabIndex        =   1
         Top             =   345
         Width           =   630
      End
      Begin VB.Label LblIC��״̬ 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "IC��״̬(&S)"
         Enabled         =   0   'False
         Height          =   180
         Left            =   120
         TabIndex        =   11
         Top             =   2040
         Width           =   990
      End
      Begin VB.Label Lbl�������� 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "��������(&A)"
         Enabled         =   0   'False
         Height          =   180
         Left            =   3750
         TabIndex        =   17
         Top             =   780
         Width           =   990
      End
      Begin VB.Label Lbl���������� 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "����������(&F)"
         Enabled         =   0   'False
         Height          =   180
         Left            =   3570
         TabIndex        =   19
         Top             =   1200
         Width           =   1170
      End
      Begin VB.Label Lbl��λ���� 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "��λ����(&W)"
         Enabled         =   0   'False
         Height          =   180
         Left            =   3750
         TabIndex        =   15
         Top             =   345
         Width           =   990
      End
      Begin VB.Label Lbl����ۼ� 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "����ۼ�(&L)"
         Enabled         =   0   'False
         Height          =   180
         Left            =   3750
         TabIndex        =   25
         Top             =   2040
         Width           =   990
      End
      Begin VB.Label Lbl�ʻ���� 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "���(&Q)"
         Enabled         =   0   'False
         Height          =   180
         Left            =   5310
         TabIndex        =   23
         Top             =   1620
         Width           =   630
      End
      Begin VB.Label LblסԺ���� 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "סԺ����(&Z)"
         Enabled         =   0   'False
         Height          =   180
         Left            =   3750
         TabIndex        =   21
         Top             =   1620
         Width           =   990
      End
      Begin VB.Label lbl����״̬ 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "����״̬(&T)"
         Enabled         =   0   'False
         Height          =   180
         Left            =   120
         TabIndex        =   13
         Top             =   2460
         Width           =   990
      End
      Begin VB.Label lbl���� 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "����(&F)"
         Height          =   180
         Left            =   4080
         TabIndex        =   27
         Top             =   2460
         Width           =   630
      End
      Begin VB.Label lbl���� 
         AutoSize        =   -1  'True
         Caption         =   "����(&A)"
         Enabled         =   0   'False
         Height          =   180
         Left            =   2310
         TabIndex        =   5
         Top             =   780
         Width           =   630
      End
      Begin VB.Label lbl���� 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "����(&D)"
         Enabled         =   0   'False
         Height          =   180
         Left            =   480
         TabIndex        =   7
         Top             =   1200
         Width           =   630
      End
      Begin VB.Label lblҽ���� 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "ҽ����(&Y)"
         Enabled         =   0   'False
         Height          =   180
         Left            =   300
         TabIndex        =   9
         Top             =   1620
         Width           =   810
      End
      Begin VB.Label lbl�Ա� 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "�Ա�(&X)"
         Enabled         =   0   'False
         Height          =   180
         Left            =   480
         TabIndex        =   3
         Top             =   780
         Width           =   630
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   6600
      TabIndex        =   31
      Top             =   3150
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   5340
      TabIndex        =   30
      Top             =   3150
      Width           =   1100
   End
End
Attribute VB_Name = "frmIdentify��������"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private blnOK As Boolean
Dim mlng����ID As Long
Dim mint���� As Integer

Public Function ShowCard(Optional lng����ID As Long, Optional ByVal int���� As Integer) As Boolean
'���ܣ�����ҽ�����˵������Ϣ
'������bytType-ʶ�����ͣ�0-���1-סԺ
'���أ�
    '0����;1ҽ����;2����;3����;4�Ա�;5��������;6���֤;7��λ����(����);
    '8.���Ĵ���;9.˳���;10��Ա���;11�ʻ����;12��ǰ״̬;13����ID;14��ְ(0,1);15����֤��;16�����;17�Ҷȼ�
    '18�ʻ������ۼ�,19�ʻ�֧���ۼ�,20����ͳ���ۼ�,21ͳ�ﱨ���ۼ�,22סԺ�����ۼ�
    Dim rsTemp As New ADODB.Recordset
    mlng����ID = lng����ID
    mint���� = int����
    
    cob�Ա�.Clear
    gstrSQL = "select ����,���� from �Ա� order by ����"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    
    Do Until rsTemp.EOF
        cob�Ա�.AddItem rsTemp("����") & "." & rsTemp("����")
        rsTemp.MoveNext
    Loop
    cob�Ա�.ListIndex = 0
    rsTemp.Close
    Call Get�ʻ����

    frmIdentify��������.Show vbModal
    ShowCard = blnOK
End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    '���²���
    
    gstrSQL = "zl_�����ʻ�_������Ϣ(" & mlng����ID & "," & mint���� & ",'����ID','" & IIf(txt����.Tag = "", "NULL", txt����.Tag) & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "���²���")
    blnOK = True
    Unload Me
End Sub

Private Sub Get�ʻ����()
'���Ѿ����ڵļ�¼�ж����ʻ���Ϣ
    Dim strValue As String
    Dim rs�ʻ� As ADODB.Recordset
    Dim rsTemp As New ADODB.Recordset
    Dim lngIndex As Long
    
    gstrSQL = " Select A.����ID as ID,A.����,A.ҽ����,B.����,B.�Ա�,B.����, " & _
              " A.����ID,substr(D.����,instr(D.����,'@@')+2) as ����,��λ���� As ����״̬,Nvl(��Ա���,0) As סԺ����,����֤�� As ���ҽ�������ۼ�,�ʻ����" & _
              " From �����ʻ� A,������Ϣ B,��������Ŀ¼ C,���ղ��� D" & _
              " Where A.����ID=B.����ID and A.����=" & mint���� & _
              " And A.����=C.���� and A.����=C.��� and A.����ID=D.ID(+) And A.����ID=" & mlng����ID
    Set rs�ʻ� = frmPubSel.ShowSelect(Me, gstrSQL, 0, "�����ʻ�", , "", "", False, True)
    If Not rs�ʻ� Is Nothing Then
    
        '�������õ�����
        mlng����ID = rs�ʻ�!ID
        txt����.Text = IIf(IsNull(rs�ʻ�("����")), "", rs�ʻ�("����"))
        txt����.Text = IIf(IsNull(rs�ʻ�("����")), "", rs�ʻ�("����"))
        Call SetComboByText(cob�Ա�, IIf(IsNull(rs�ʻ�("�Ա�")), "", rs�ʻ�("�Ա�")), True)
        txt����.Text = IIf(IsNull(rs�ʻ�("����")), "", rs�ʻ�("����"))
        txtҽ����.Text = IIf(IsNull(rs�ʻ�("ҽ����")), "", rs�ʻ�("ҽ����"))
        txtסԺ����.Text = IIf(IsNull(rs�ʻ�("סԺ����")), "", rs�ʻ�("סԺ����"))
        Txt����״̬.Text = IIf(IsNull(rs�ʻ�("����״̬")), "", rs�ʻ�("����״̬"))
        txt�ʻ����.Text = Format(rs�ʻ�("�ʻ����"), "#####0.00;-#####0.00; ;")
        Txt����ۼ�.Text = Format(rs�ʻ�("���ҽ�������ۼ�"), "#####0.00;-#####0.00; ;")
'        txt����.Text = IIf(IsNull(rs�ʻ�("����")), "", rs�ʻ�("����"))
'        txt����.Tag = IIf(IsNull(rs�ʻ�("����ID")), "", rs�ʻ�("����ID"))
    End If
    
    '��д������Ϣ
    Call Record_Locate(mrsIniItems, "����,Dwmc00")
    txt��λ����.Text = Nvl(mrsIniItems!ֵ, "")
    Call Record_Locate(mrsIniItems, "����,Icztmc")
    TxtIC��״̬.Text = Nvl(mrsIniItems!ֵ, "")
    Call Record_Locate(mrsIniItems, "����,Dqmc00")
    Txt��������.Text = Nvl(mrsIniItems!ֵ, "")
    Call Record_Locate(mrsIniItems, "����,Fzxmc0")
    Txt����������.Text = Nvl(mrsIniItems!ֵ, "")
End Sub

Private Sub cmd����_Click()
    Dim rsTemp As New ADODB.Recordset
    gstrSQL = " Select A.ID,substr(A.����,1,instr(A.����,'@@')-1) ����,substr(A.����,instr(A.����,'@@')+2) ����,A.����,decode(A.���,1,'���Բ�',2,'���ֲ�','��ͨ��') as ��� " & _
            " From ���ղ��� A where A.����=" & mint����
    
    Set rsTemp = frmPubSel.ShowSelect(Me, gstrSQL, 0, "ҽ������", , txt����.Text)
    If rsTemp.State = 0 Then Exit Sub
    If Not rsTemp Is Nothing Then
        txt����.Text = rsTemp("����")
        txt����.Tag = rsTemp("ID")
        zlControl.TxtSelAll txt����
    End If
    txt����.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey (vbKeyTab)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    blnOK = False
End Sub

Private Sub txt����_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        txt����.Text = ""
        txt����.Tag = ""
    End If
End Sub

Private Function Record_Locate(ByRef rsObj As ADODB.Recordset, ByVal strPrimary As String, Optional ByVal blnDelete As Boolean = False) As Boolean
    Dim arrTmp
    '��λ��ָ����¼
    'strPrimary:����,ֵ
    'blnDelete=True,��ü�¼������"ɾ��"�ֶ�
    Record_Locate = False
    
    arrTmp = Split(strPrimary, ",")
    With rsObj
        If .RecordCount = 0 Then Exit Function
        .MoveFirst
        .Find arrTmp(0) & "='" & arrTmp(1) & "'"
        If .EOF Then Exit Function
        If blnDelete Then
            Do While Not .EOF
                If !ɾ�� = 0 Then Record_Locate = True: Exit Do
                .MoveNext
            Loop
        Else
            Record_Locate = True
        End If
    End With
End Function
