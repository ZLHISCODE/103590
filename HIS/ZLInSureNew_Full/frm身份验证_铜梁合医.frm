VERSION 5.00
Begin VB.Form frm�����֤_ͭ����ҽ 
   Caption         =   "ͭ����ҽ�����֤"
   ClientHeight    =   4965
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8010
   Icon            =   "frm�����֤_ͭ����ҽ.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   331
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   534
   StartUpPosition =   1  '����������
   Begin VB.CommandButton Cmd_ok 
      Caption         =   "ȷ��(&O)"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4920
      TabIndex        =   25
      Top             =   4560
      Width           =   1335
   End
   Begin VB.CommandButton Cmd_cancle 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6600
      TabIndex        =   24
      Top             =   4560
      Width           =   1335
   End
   Begin VB.Frame Frame2 
      Caption         =   "������Ϣ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   120
      TabIndex        =   16
      Top             =   3120
      Width           =   7815
      Begin VB.CheckBox Chk���Բ� 
         Caption         =   "�Ƿ��������Բ�"
         Enabled         =   0   'False
         Height          =   255
         Left            =   120
         TabIndex        =   37
         Top             =   248
         Width           =   1695
      End
      Begin VB.CheckBox Chk��֯���� 
         Caption         =   "�Ƿ�ũ�ϰ���֯����"
         Enabled         =   0   'False
         Height          =   255
         Left            =   120
         TabIndex        =   36
         Top             =   608
         Width           =   1935
      End
      Begin VB.TextBox Txt�������� 
         Height          =   270
         Left            =   5400
         Locked          =   -1  'True
         TabIndex        =   35
         Top             =   600
         Width           =   2295
      End
      Begin VB.TextBox Txt���ִ��� 
         Height          =   270
         Left            =   5400
         Locked          =   -1  'True
         TabIndex        =   34
         Top             =   240
         Width           =   2295
      End
      Begin VB.CheckBox ChkԤ������ 
         Caption         =   "�Ƿ�Ԥ������"
         Enabled         =   0   'False
         Height          =   255
         Left            =   2640
         TabIndex        =   33
         Top             =   608
         Width           =   1575
      End
      Begin VB.ComboBox Cob״̬ 
         Height          =   300
         Left            =   3120
         TabIndex        =   32
         Text            =   "һ��"
         Top             =   225
         Width           =   855
      End
      Begin VB.TextBox Txt��Ͻ�� 
         BackColor       =   &H80000009&
         Height          =   270
         Left            =   1320
         Locked          =   -1  'True
         MaxLength       =   16
         TabIndex        =   26
         Top             =   960
         Width           =   4935
      End
      Begin VB.Label Label16 
         Caption         =   "��������"
         Height          =   270
         Left            =   4560
         TabIndex        =   40
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label15 
         Caption         =   "���ִ���"
         Height          =   270
         Left            =   4560
         TabIndex        =   39
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label17 
         Caption         =   "״̬:"
         Height          =   255
         Left            =   2640
         TabIndex        =   38
         Top             =   248
         Width           =   495
      End
      Begin VB.Label Label14 
         Caption         =   "��Ͻ��"
         Height          =   270
         Left            =   120
         TabIndex        =   27
         Top             =   960
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "������Ϣ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   7815
      Begin VB.TextBox Txt�������� 
         BackColor       =   &H80000009&
         Height          =   270
         Left            =   4920
         Locked          =   -1  'True
         TabIndex        =   29
         Top             =   360
         Width           =   1815
      End
      Begin VB.TextBox Txt��ҽ���� 
         Height          =   270
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   28
         Top             =   360
         Width           =   2055
      End
      Begin VB.TextBox Txtסַ 
         BackColor       =   &H80000009&
         Height          =   270
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   22
         Top             =   2640
         Width           =   4575
      End
      Begin VB.TextBox Txt���� 
         BackColor       =   &H80000009&
         Height          =   270
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   720
         Width           =   2055
      End
      Begin VB.TextBox Txt�Ա� 
         BackColor       =   &H80000009&
         Height          =   270
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   1080
         Width           =   2055
      End
      Begin VB.TextBox Txt�������� 
         BackColor       =   &H80000009&
         Height          =   270
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   2280
         Width           =   2055
      End
      Begin VB.TextBox Txt�������� 
         BackColor       =   &H80000009&
         Height          =   270
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   1440
         Width           =   2055
      End
      Begin VB.TextBox Txt�����Ա� 
         BackColor       =   &H80000009&
         Height          =   270
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   1875
         Width           =   2055
      End
      Begin VB.TextBox Txt�ʻ���� 
         BackColor       =   &H80000009&
         Height          =   270
         Left            =   4920
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   2280
         Width           =   1815
      End
      Begin VB.TextBox Txt���֤�� 
         BackColor       =   &H80000009&
         Height          =   270
         Left            =   4920
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   720
         Width           =   1815
      End
      Begin VB.TextBox Txt��ͥ�ʺ� 
         BackColor       =   &H80000009&
         Height          =   270
         Left            =   4920
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   1080
         Width           =   1815
      End
      Begin VB.TextBox Txt�������֤ 
         BackColor       =   &H80000009&
         Height          =   270
         Left            =   4920
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   1440
         Width           =   1815
      End
      Begin VB.TextBox Txt������ϵ 
         BackColor       =   &H80000009&
         Height          =   270
         Left            =   4920
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   1875
         Width           =   1815
      End
      Begin VB.Label Label13 
         Caption         =   "ҽ�ƻ�������"
         Height          =   255
         Left            =   3480
         TabIndex        =   31
         Top             =   375
         Width           =   1095
      End
      Begin VB.Label Label10 
         Caption         =   "��ҽ����"
         Height          =   270
         Left            =   240
         TabIndex        =   30
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "ס    ַ"
         Height          =   270
         Left            =   240
         TabIndex        =   23
         Top             =   2640
         Width           =   735
      End
      Begin VB.Label Label8 
         Caption         =   "�����Ա�"
         Height          =   270
         Left            =   240
         TabIndex        =   21
         Top             =   1875
         Width           =   735
      End
      Begin VB.Label Label7 
         Caption         =   "��������"
         Height          =   270
         Left            =   240
         TabIndex        =   20
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label Label6 
         Caption         =   "��������"
         Height          =   270
         Left            =   240
         TabIndex        =   19
         Top             =   2280
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "��    ��"
         Height          =   270
         Left            =   240
         TabIndex        =   18
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "��    ��"
         Height          =   270
         Left            =   240
         TabIndex        =   17
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "���֤��"
         Height          =   255
         Left            =   3480
         TabIndex        =   15
         Top             =   735
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "��ͥ�ʺ�"
         Height          =   255
         Left            =   3480
         TabIndex        =   14
         Top             =   1095
         Width           =   1095
      End
      Begin VB.Label Label9 
         Caption         =   "�������֤��"
         Height          =   255
         Left            =   3480
         TabIndex        =   13
         Top             =   1455
         Width           =   1095
      End
      Begin VB.Label Label11 
         Caption         =   "������ϵ"
         Height          =   255
         Left            =   3480
         TabIndex        =   12
         Top             =   1875
         Width           =   1095
      End
      Begin VB.Label Label12 
         Caption         =   "�ʻ����"
         Height          =   255
         Left            =   3480
         TabIndex        =   11
         Top             =   2295
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frm�����֤_ͭ����ҽ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private ���ش�1 As String, ���ش�2 As String, ��ˮ�� As String, Cȷ�� As Boolean, b��ҽ��� As Byte
Private R���� As String, R�Ա� As String, R���֤�� As String, R��ͥ�ʺ� As String, var�������� As String, var�������� As String
Private R�������� As String, R�������� As String, R�����Ա� As String, R��ҽ���� As String
Private R�������֤ As String, Rסַ As String, R������ϵ As String, R�ʻ���� As Double, Str��ҽ��Ϣ As String
Private R���ִ��� As String, R�������� As String

Public Function GetIdentify(Optional bytType As Byte, Optional lng����ID As Long = 0, Optional ByRef intinsure As Integer = 0) As String
Cȷ�� = False
b��ҽ��� = bytType
    Me.Show 1
    If Cȷ�� Then
        lng����ID = BuildPatiInfo(bytype, ���ش�1 & ���ش�2, lng����ID, type_ͭ����ҽ)
        gstrSQL = "zl_�����ʻ�_������Ϣ(" & lng����ID & "," & intinsure & ",'��ҽ��Ϣ','''" & Str��ҽ��Ϣ & "''')"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "����������ϵ�")
        gstrSQL = "zl_�����ʻ�_������Ϣ(" & lng����ID & "," & intinsure & ",'�ʻ����','" & R�ʻ���� & "')"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "�������")
        gstrSQL = "zl_�����ʻ�_������Ϣ(" & lng����ID & "," & intinsure & ",'��λ����','" & Trim(MidUni(R��ͥ�ʺ�, 1, 20)) & "')"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "�����ͥ�ʺ��ڱ����ʻ��ĵ�λ������")
        GetIdentify = ���ش�1 & ";" & lng����ID & ���ش�2
        gstrSQL = "zl_�����ʻ�_������Ϣ(" & lng����ID & "," & intinsure & ",'���ִ���','''" & Trim(MidUni(R���ִ���, 1, 20)) & "''')"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "���没�ִ����ڱ����ʻ���")
        gstrSQL = "zl_�����ʻ�_������Ϣ(" & lng����ID & "," & intinsure & ",'��������','''" & Trim(MidUni(R��������, 1, 200)) & "''')"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "���没�������ڱ����ʻ���")
    Else
        GetIdentify = ""
    End If
End Function

Private Sub Cmd_cancle_Click()
Unload Me
End Sub
Private Sub Cmd_ok_Click()
    If Me.Txt��ҽ����.Text = "" Then
        MsgBox "��ҽ��Ų���Ϊ��", vbInformation, gstrSysName
        Me.Txt��ҽ����.SetFocus
        Exit Sub
    End If
    
    '���췵�ش�
    If b��ҽ��� = 0 Then '����
        If Me.Chk���Բ�.Value = 1 And (Trim(Me.txt���ִ���.Text) = "" Or Trim(Me.txt��������.Text) = "") Then
            MsgBox "���Բ�����ѡ���ִ���", vbInformation, gstrSysName
            Me.txt���ִ���.SetFocus
            Exit Sub
        End If
    ElseIf b��ҽ��� = 1 Then 'סԺ
        If Trim(Me.txt���ִ���.Text) = "" Or Trim(Me.txt��������.Text) = "" Then
            MsgBox "סԺ�����ȡ���ִ���Ͳ�������", vbInformation, gstrSysName
            Me.txt���ִ���.SetFocus
            Exit Sub
        End If
    End If
    ���ش�1 = Me.Txt��ҽ����.Text & _
            ";" & Me.Txt��ҽ����.Text & _
            ";;" & Me.txt����.Text & _
            ";" & Me.txt�Ա�.Text & _
            ";" & Mid(Me.txt��������.Text, 1, 4) & "-" & Mid(Me.txt��������.Text, 6, 2) & "-" & Mid(Me.txt��������.Text, 9, 2) & _
            ";" & Trim(Me.txt���֤��.Text) & _
            ";" & Me.Txtסַ.Text
    ���ش�2 = ";" & gintInsure & _
            ";" & ��ˮ�� & _
            ";" & MidUni(Me.Txt������ϵ.Text, 1, 8) & _
            ";" & Me.txt�ʻ����.Text & _
            ";0;;1;" & Me.Txt�������֤ & _
            ";;;;;;;;;"
    Str��ҽ��Ϣ = Trim(Me.Txt��������.Text) & "|" & IIf(Me.Chk��֯���� = 0, 0, 1) & "|" & IIf(Me.Chk���Բ�.Value = 0, 0, 1) & "|" & IIf(Me.ChkԤ������.Value = 0, 0, 1) & "|" & Trim(Me.Cob״̬.Text) & "|" & MidUni(Trim(Me.Txt��Ͻ��.Text), 1, 16)
    Cȷ�� = True
    Unload Me
End Sub


Private Sub Form_Load()
    Me.Cob״̬.AddItem ("һ��")
    Me.Cob״̬.AddItem ("Σ")
    Me.Cob״̬.AddItem ("��")
    Me.Cob״̬.AddItem ("����")
End Sub

Private Sub Txt���ִ���_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
    R���ִ��� = Space(30)
    R�������� = Space(200)
        If GetBZDM(R���ִ���, R��������, 150, 150) <> 1 Then
            MsgBox "������Ϣ" & GetMyLastError(), vbInformation, "��ҽ������Ϣ"
        Else
            Me.txt���ִ��� = Trim(MidUni(R���ִ���, 1, 30))
            Me.txt�������� = Trim(MidUni(R��������, 1, 200))
            Me.Cmd_ok.SetFocus
        End If
    End If
End Sub

Private Sub Txt��ҽ����_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        var�������� = Space(20)
        var�������� = Space(100)
        R��ҽ���� = Space(20)
        Dim rsҽԺ���� As New ADODB.Recordset
       
        gstrSQL = "select  ҽԺ���� from ������� where ���=[1]"
        Set rsҽԺ���� = zlDatabase.OpenSQLRecord(gstrSQL, "ȡҽԺ����", gintInsure)
        var�������� = rsҽԺ����!ҽԺ����
        rsҽԺ����.Close
        Set rsҽԺ���� = Nothing
        
        If Trim(Me.Txt��ҽ����) = "" Then
            If GetCHRYBM(R��ҽ����, 150, 150) <> 1 Then
                MsgBox "������Ϣ:" & GetMyLastError(), vbInformation, "��ҽ������Ϣ"
                Exit Sub
            End If
        Call OutInfo
        
         Else
            R��ҽ���� = Trim(Me.Txt��ҽ����.Text)
            Call OutInfo
            Txt��Ͻ��.SetFocus
        End If
    End If
End Sub
Private Sub OutInfo()
R���� = Space(10)
R�Ա� = Space(4)
R���֤�� = Space(20)
R��ͥ�ʺ� = Space(20)
R�������� = Space(12)
R�������� = Space(10)
R�����Ա� = Space(4)
R�������֤ = Space(20)
Rסַ = Space(50)
R������ϵ = Space(10)

    If GetRyInfo(R��ҽ����, R����, R�Ա�, R���֤��, R��ͥ�ʺ�, R��������, R��������, R�����Ա�, R�������֤, Rסַ, R������ϵ, R�ʻ����) <> 1 Then
            MsgBox "������Ϣ" & GetMyLastError, vbInformation, "��ҽ������Ϣ"
            Me.Cmd_ok.Enabled = False
            Me.Txt��ҽ����.SetFocus
            Exit Sub
    Else
        '�����Ϣ������
        With Me
            .Txt��ҽ����.Text = Trim(MidUni(R��ҽ����, 1, 20))
            .Txt�������� = Trim(MidUni(var��������, 1, 10))
            .txt����.Text = Trim(MidUni(R����, 1, 10))
            .txt�Ա�.Text = IIf(Trim(MidUni(R�Ա�, 1, 4)) <> "��" And Trim(MidUni(R�Ա�, 1, 4)) <> "Ů", "δ֪", Trim(MidUni(R�Ա�, 1, 4)))
            .txt���֤��.Text = Trim(MidUni(R���֤��, 1, 20))
            .Txt��ͥ�ʺ�.Text = Trim(MidUni(R��ͥ�ʺ�, 1, 20))
            .txt��������.Text = Trim(MidUni(R��������, 1, 12))
            .Txt��������.Text = Trim(MidUni(R��������, 1, 10))
            .Txt�����Ա�.Text = Trim(MidUni(R�����Ա�, 1, 4))
            .Txt�������֤.Text = Trim(MidUni(R�������֤, 1, 20))
            .Txtסַ.Text = Trim(MidUni(Rסַ, 1, 50))
            .Txt������ϵ.Text = Trim(MidUni(R������ϵ, 1, 10))
            .txt�ʻ����.Text = Trim(MidUni(CStr(R�ʻ����), 1, 10))
            .Txt��Ͻ��.Locked = False
            .Txt��Ͻ��.BackColor = &HFFFFFF
            .txt�ʻ����.ForeColor = &HFF0000
            .Chk���Բ�.Enabled = True
            .Chk��֯����.Enabled = True
            .ChkԤ������.Enabled = True
        End With
        If b��ҽ��� = 1 Then
            Me.txt���ִ���.SetFocus
        Else
            Me.Txt��Ͻ��.SetFocus
        End If
        Me.Cmd_ok.Enabled = True
    End If
End Sub


Private Sub Txt��Ͻ��_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
           Cmd_ok.SetFocus
    End If
End Sub

