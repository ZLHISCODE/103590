VERSION 5.00
Begin VB.Form frmIdentify�Ͼ��� 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "�������������֤"
   ClientHeight    =   4395
   ClientLeft      =   30
   ClientTop       =   270
   ClientWidth     =   4590
   Icon            =   "frmIdentify�Ͼ���.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4395
   ScaleWidth      =   4590
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmd��ҽ�� 
      Caption         =   "��ҽ��(&S)"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   180
      TabIndex        =   8
      Top             =   3840
      Width           =   1365
   End
   Begin VB.Frame fra���� 
      Caption         =   "ҽ�����˻�����Ϣ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3475
      Left            =   120
      TabIndex        =   10
      Top             =   240
      Width           =   4404
      Begin VB.TextBox txt���� 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1920
         TabIndex        =   3
         Top             =   1830
         Width           =   1695
      End
      Begin VB.ComboBox cbo�Ա� 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "frmIdentify�Ͼ���.frx":000C
         Left            =   1920
         List            =   "frmIdentify�Ͼ���.frx":001C
         TabIndex        =   2
         Top             =   1360
         Width           =   1695
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Left            =   1920
         TabIndex        =   6
         Top             =   2820
         Width           =   1692
      End
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Left            =   1920
         TabIndex        =   1
         Top             =   855
         Width           =   1692
      End
      Begin VB.CommandButton cmd������Ϣ 
         Caption         =   "��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Left            =   3624
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   2295
         Width           =   372
      End
      Begin VB.TextBox txt���ﲡ�� 
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Left            =   1920
         TabIndex        =   4
         Top             =   2295
         Width           =   1692
      End
      Begin VB.TextBox txt���� 
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Left            =   1920
         TabIndex        =   0
         Top             =   360
         Width           =   1692
      End
      Begin VB.Label Label4 
         Caption         =   "����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1320
         TabIndex        =   16
         Top             =   1920
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "�Ա�"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1320
         TabIndex        =   15
         Top             =   1440
         Width           =   495
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "��ȷ����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   855
         TabIndex        =   14
         Top             =   2880
         Width           =   960
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1320
         TabIndex        =   12
         Top             =   915
         Width           =   480
      End
      Begin VB.Label lbl���ﲡ�� 
         AutoSize        =   -1  'True
         Caption         =   "���ﲡ��(&F)"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   480
         TabIndex        =   13
         Top             =   2370
         Width           =   1320
      End
      Begin VB.Label lbl���� 
         AutoSize        =   -1  'True
         Caption         =   "��ʶ��(&N)"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   720
         TabIndex        =   11
         Top             =   420
         Width           =   1080
      End
   End
   Begin VB.CommandButton cmdCancle 
      Caption         =   "ȡ��(&C)"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   3432
      TabIndex        =   9
      Top             =   3840
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   2232
      TabIndex        =   7
      Top             =   3840
      Width           =   1100
   End
End
Attribute VB_Name = "frmIdentify�Ͼ���"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mbytType As Byte
Private mstrIdentify As String
Private mlng����ID As Long, mlng����ID As Long
Private mstr�������� As String
Private mstr���ֱ��� As String
Private mstr�������� As String
Private mstrҽ���� As String
Private mstrsubInsure  As String '������ҽ���ķ������ݣ��������|�Ż����|ҽ����|���|ͣ��

Private mintInsure As Integer, mstrReturn As String

Private Sub cbo�Ա�_KeyPress(KeyAscii As Integer)
  If cbo�Ա�.ListIndex = -1 Then cbo�Ա�.ListIndex = 0
End Sub

Private Sub cmdCancle_Click()
    mstrIdentify = ""
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim strSQL As String
    Dim strIdentify As String
    Dim strAddition As String
    Dim lngSequence As String
'    Dim str�Ա� As String, str�������� As String
    On Error GoTo errHandle
    
    '�ж��Ƿ�����ҽ����������
    If Trim(Text1.Text) = "" Then
        MsgBox "δ��ȡ��ҽ����������", vbInformation, gstrSysName
        txt����.SetFocus
        Exit Sub
    End If
    
    mstr�������� = Trim(Text1.Text)
    mstrҽ���� = txt����.Text
    
    If Trim(txt���ﲡ��.Text) = "" Or txt���ﲡ�� <> mstr�������� Then
        MsgBox "���ﲡ��δ¼�������", vbInformation, gstrSysName
        txt���ﲡ��.SetFocus
        Exit Sub
    End If
    
    If InStr(1, Me.Tag, "|") <> 0 Then
'        str�Ա� = Split(Me.Tag, "|")(0)
'        str�������� = Split(Me.Tag, "|")(1)
    End If
    
    '�˴��޷�ȡ�ÿ��ź�ҽ����,������ʱ���뱣�ղ�������,�Ժ�õ����ź��ٽ����޸�
    lngSequence = Right(String(20, "0") & Text1.Tag, 20)
'      strInfo='0����;1ҽ����;2����;3����;4�Ա�;5��������;6���֤;7��λ����(����);
'      8����;9.˳���;10��Ա���;11�ʻ����;12��ǰ״̬;13����ID;14��ְ(1,2,3);15����֤��;16�����;17�Ҷȼ�
'      18�ʻ������ۼ�;19�ʻ�֧���ۼ�;20����ͳ���ۼ�;21ͳ�ﱨ���ۼ�;22סԺ�����ۼ�;23�������
'      24��������;25�����ۼ�;26����ͳ���޶�
    
    strIdentify = strIdentify & ";"                                       '0����
    strIdentify = strIdentify & txt����.Text & ";"                  '1ҽ���ţ����˱�ţ�
    strIdentify = strIdentify & ";"                                 '2����
    strIdentify = strIdentify & Text1.Text & ";"                   '3����
    strIdentify = strIdentify & cbo�Ա�.Text & ";"                                 '4�Ա�
    strIdentify = strIdentify & IIf(Trim(txt����.Text) = "", Format(zlDatabase.Currentdate, "yyyy-mm-dd"), Get��������("", Val(txt����.Text))) & ";"                       '5��������
    strIdentify = strIdentify & ";"                                 '6���֤
    strIdentify = strIdentify & ";"                               '7.��λ����(����)
    strAddition = "0;"                                          '8.���Ĵ���
    strAddition = strAddition & ";"                               '9.˳���
    strAddition = strAddition & ";"                            '10��Ա���
    strAddition = strAddition & "10000;"                              '11�ʻ����
    strAddition = strAddition & "0;"                            '12��ǰ״̬
    strAddition = strAddition & mlng����ID & ";"                 '13����ID
    strAddition = strAddition & "1;"                            '14��ְ(1,2,3)
    strAddition = strAddition & mstrsubInsure & ";"             '15����֤��
    strAddition = strAddition & ";"                             '16�����
    strAddition = strAddition & ";"                             '17�Ҷȼ�
    strAddition = strAddition & ";"                             '18�ʻ������ۼ�
    strAddition = strAddition & ";"                            '19�ʻ�֧���ۼ�
    strAddition = strAddition & "0;"                            '20����ͳ���ۼ�
    strAddition = strAddition & "0;"                            '21ͳ�ﱨ���ۼ�
    strAddition = strAddition & "0;"                             '22סԺ�����ۼ�
    strAddition = strAddition & ";"                             '23��������
    
    mlng����ID = BuildPatiInfo(0, strIdentify & strAddition, mlng����ID, TYPE_�Ͼ���)
    '���ظ�ʽ:�м���벡��ID
    If mlng����ID > 0 Then
        mstrIdentify = strIdentify & mlng����ID & ";" & strAddition
    End If
    If Trim(Text2.Text) <> "" Then
        mstr�������� = Trim(Text2.Text)
    Else
        mstr�������� = Trim(Text1.Text)
    End If
    
    Unload Me
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    MsgBox mlng����ID & "��" & lngSequence & "��" & Text1.Tag & "��" & strIdentify & strAddition, vbInformation, gstrSysName
End Sub


Private Sub cmd������Ϣ_Click()
    Dim rsTemp As New ADODB.Recordset
    
    gstrSQL = "select id,����,����,decode(���,1,'���Բ�',2,'���ⲡ','��ͨ��') as ���� from ���ղ��� where ����=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ѡ����", TYPE_�Ͼ���)
    
    If frmListSel.ShowSelect(TYPE_�Ͼ���, rsTemp, "ID", "ҽ������ѡ��", "��ѡ���ض���ҽ�����֣�") Then
        txt���ﲡ��.Text = rsTemp!����
        mlng����ID = rsTemp!ID
        mstr���ֱ��� = rsTemp!����
        mstr�������� = rsTemp!����
    Else
        txt���ﲡ��.SetFocus
    End If
End Sub

Private Sub cmd��ҽ��_Click()
    '��ʾ����ҽ���������֤����
    '��������:�������|�Ż����|ҽ����|���|ͣ��
    If Trim(txt����.Text) = "" Then
        MsgBox "����ȷ��ҽ��������ݣ�", vbInformation, gstrSysName
        Exit Sub
    End If
    
    mstrsubInsure = frm��ҽ�������֤.ShowME(Me.Text1.Text)
    cmdOK.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyReturn Then zlCommFun.PressKey (vbKeyTab)
End Sub

Private Sub Form_Load()
    Dim str������� As String
    If mbytType = 0 Or mbytType = 3 Then
        txt���ﲡ��.Enabled = True
    Else
        txt���ﲡ��.Enabled = False
    End If
    
    str������� = GetSetting("ZLSOFT", "����ȫ��", "����ҽ���ӿ�", "")
    cmd��ҽ��.Enabled = (str������� <> "")
End Sub

Private Sub txt���ﲡ��_GotFocus()
    OpenIme ("")
    Call zlControl.TxtSelAll(txt���ﲡ��)
End Sub

Private Sub txt���ﲡ��_KeyPress(KeyAscii As Integer)
    Dim rsTemp As New ADODB.Recordset
    Dim strText As String
    Dim blnReturn As Boolean
    
    If KeyAscii <> vbKeyReturn Then Exit Sub
    
    On Error GoTo errorhandle
    '�������ﲡ��

    strText = txt���ﲡ��.Text
    gstrSQL = "select A.id,A.����,A.���� from ���ղ��� A where A.����=[1] and (A.���� like [2] || '%' or A.���� like [2] || '%' or A.���� like [2] || '%')"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "���ﲡ��", TYPE_�Ͼ���, strText)
    
    If rsTemp.RecordCount = 1 Then
        blnReturn = True
    Else
        blnReturn = frmListSel.ShowSelect(TYPE_�Ͼ���, rsTemp, "ID", "ҽ������ѡ��", "��ѡ���ض���ҽ�����֣�")
    End If
    
    If blnReturn Then
        txt���ﲡ��.Text = rsTemp!����
        mlng����ID = rsTemp!ID
        mstr���ֱ��� = rsTemp!����
        mstr�������� = rsTemp!����
        zlCommFun.PressKey (vbKeyTab)
    Else
        txt���ﲡ��_GotFocus
    End If
    Exit Sub
    
errorhandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub txt����_KeyPress(KeyAscii As Integer)
   If InStr("1234567890" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txt����_GotFocus()
    Call zlControl.TxtSelAll(txt����)
End Sub

Public Function Identify(ByVal bytType As Byte, lng����ID As Long) As String
    mbytType = bytType
    mlng����ID = lng����ID
    Me.Show 1
    Identify = mstrIdentify
    With gPatInfo_�Ͼ���
        .ҽ���� = mstrҽ����
        .�������� = mstr��������
        .���ֱ��� = mstr���ֱ���
        .�������� = mstr��������
    End With
    lng����ID = mlng����ID
End Function

'Private Sub Txt����_KeyDown(KeyCode As Integer, Shift As Integer)
'    If Trim(txt����.Text) = "" Then KeyCode = 0
'    If KeyCode <> vbKeyReturn Then Exit Sub
'Dim rsTemp As ADODB.Recordset
'    gstrSQL = "select A.ҽ����,B.����,B.�Ա�,B.����,B.������λ" & _
'                " from �����ʻ� A,������Ϣ B" & _
'                " Where A.ҽ����=[1] AND A.����=[2] and A.����ID=B.����ID "
'    Set rsTemp = OpenSQLRecord(gstrSQL, "��ȡ������Ϣ", CStr(txt����.Text), CInt(mintInsure))
'    If rsTemp.EOF Then 'û�ڱ�ҽԺ�����
'        Text1.Locked = False
'        cbo�Ա�.Locked = False
'        txt����.Locked = False
'
'    Else
'        txt����.Tag = Nvl(rsTemp!ҽ����)
'        Text1.Text = Nvl(rsTemp!����)
'        cbo�Ա�.Text = Nvl(rsTemp!�Ա�)
'        txt����.Text = Nvl(rsTemp!����)
'
'
'    End If
'
'End Sub

Private Sub txt����_KeyPress(KeyAscii As Integer)
    Dim datCurr As Date
    Dim StrInput As String, strSQL As String, rsTemp As New ADODB.Recordset
    
    
    On Error GoTo errHandle
    If KeyAscii <> vbKeyReturn Then Exit Sub
    If txt����.Tag = "." Then Exit Sub
    
        If txt���� <> "" Then
        Me.Tag = ""
            If Left(txt����.Text, 1) = "." Then
        txt����.Tag = "."
        StrInput = txt����.Text
        If Not IsNumeric(Mid(StrInput, 2)) Then Exit Sub
        If Len(Mid(StrInput, 2)) <= 4 Then
            datCurr = zlDatabase.Currentdate()
            strSQL = PreFixNO & Format(CDate(Format(datCurr, "YYYY-MM-dd")) - CDate(Format(datCurr, "YYYY") & "-01-01") + 1, "000") & Format(Mid(StrInput, 2), "0000") '����˳����
        Else
            strSQL = GetFullNO(Mid(StrInput, 2))
        End If
        '�������ʱ����Ҫ�ҺŽ���
        strSQL = "Select A.����id,A.����,A.��ʶ��,A.�Ա�,A.���� From ������ü�¼ A Where A.NO='" & strSQL & "' And A.��¼����=4 And A.��¼״̬=1 "
        Set rsTemp = gcnOracle.Execute(strSQL)
'        If rsTemp.EOF Then
'            MsgBox "����ĹҺŵ���", vbInformation, gstrSysName
'            Exit Sub
'        End If
        mlng����ID = rsTemp!����ID
        strSQL = "Select a.����,a.�����,A.�Ա�,A.����,b.ҽ����,c.ID,c.����,c.���� From ������Ϣ a,�����ʻ� b,���ղ��� c Where a.����id=b.����id(+) and b.����id=c.id(+) and a.����ID=" & rsTemp!����ID
        Set rsTemp = gcnOracle.Execute(strSQL)
        If rsTemp.EOF Then
            MsgBox "��ȡ������Ϣ����", vbInformation, gstrSysName
            Exit Sub
        ElseIf IsNull(rsTemp!�����) Then
            MsgBox "�ò��˵�ҽ������û��¼��", vbInformation, gstrSysName
            Exit Sub
        Else
            If IsNull(rsTemp!ҽ����) Then
                MsgBox "�벹¼�ò���ҽ����", vbInformation, gstrSysName
            End If
            Text1.Text = rsTemp!����
            Text1.Tag = rsTemp!�����
            cbo�Ա�.Text = rsTemp!�Ա�
            txt����.Text = rsTemp!����
            txt���ﲡ��.Text = Nvl(rsTemp!����)
            txt����.Text = Nvl(rsTemp!ҽ����)
            mlng����ID = Nvl(rsTemp!ID, 0)
            mstr���ֱ��� = Nvl(rsTemp!����)
            mstr�������� = Nvl(rsTemp!����)
            '����Ѵ��ڸò��ˣ��������Ա�������ȡ����
'            Me.Tag = Nvl(rsTemp!�Ա�) & "|" & Format(Nvl(rsTemp!��������, zlDatabase.Currentdate), "yyyy-MM-dd")
        End If
        Else
            txt����.Tag = ""
'            If mbytType = 0 Then
'            MsgBox "��������ҽ�����˹Һŵ���", vbInformation, gstrSysName
'            Else
            Dim a As String
            a = txt����.Text
            Dim rsTemp1 As ADODB.Recordset
            gstrSQL = "select A.ҽ����,B.����,B.�Ա�,B.����,c.ID,C.����,c.����" & _
                " from �����ʻ� A,������Ϣ B,���ղ��� c" & _
                " Where A.ҽ����=[1] AND A.����=[2] and A.����ID=B.����ID and a.����id=c.id(+)"
            Set rsTemp1 = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ������Ϣ", CStr(txt����.Text), CInt(mintInsure))
            If rsTemp1.EOF Then 'û�ڱ�ҽԺ�����
            Text1.Text = ""
            cbo�Ա�.Text = ""
            txt����.Text = ""
            txt���ﲡ��.Text = ""
            Text1.Enabled = True
            cbo�Ա�.Enabled = True
            txt����.Enabled = True
            Else
            txt����.Tag = Nvl(rsTemp1!ҽ����)
            Text1.Text = Nvl(rsTemp1!����)
            cbo�Ա�.Text = Nvl(rsTemp1!�Ա�)
            txt����.Text = Nvl(rsTemp1!����)
            txt���ﲡ��.Text = Nvl(rsTemp1!����)
            mlng����ID = Nvl(rsTemp1!ID, 0)
            mstr���ֱ��� = Nvl(rsTemp1!����)
            mstr�������� = Nvl(rsTemp1!����)
            End If
'            End If
'        Else
'          Text1.Locked = False
'          cbo�Ա�.Locked = False
'          txt����.Locked = False
        End If
        Else
        zlCommFun.PressKey (vbKeyTab)
End If
    Exit Sub
errHandle:
    MsgBox "��ҽ������û�н�������", vbInformation, gstrSysName
End Sub


Private Sub txt����_Validate(Cancel As Boolean)
  txt����.Tag = Trim(txt����.Text)
End Sub


