VERSION 5.00
Begin VB.Form frm��Ӧ�̹��� 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "������������"
   ClientHeight    =   3330
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5340
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3330
   ScaleWidth      =   5340
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.TextBox Txt���ö� 
      Height          =   300
      Index           =   1
      Left            =   2565
      TabIndex        =   12
      Top             =   2235
      Width           =   1215
   End
   Begin VB.TextBox Txt���ö� 
      Height          =   300
      Index           =   0
      Left            =   915
      TabIndex        =   10
      Top             =   2235
      Width           =   1215
   End
   Begin VB.TextBox Txt������ 
      Height          =   300
      Index           =   1
      Left            =   2565
      TabIndex        =   8
      Top             =   1815
      Width           =   1215
   End
   Begin VB.TextBox TxtName 
      Height          =   300
      Left            =   915
      TabIndex        =   5
      Top             =   1380
      Width           =   2865
   End
   Begin VB.TextBox TxtCode 
      Height          =   300
      Index           =   1
      Left            =   2565
      TabIndex        =   3
      Top             =   960
      Width           =   1215
   End
   Begin VB.TextBox TxtCode 
      Height          =   300
      Index           =   0
      Left            =   915
      TabIndex        =   1
      Top             =   960
      Width           =   1215
   End
   Begin VB.Frame fra���� 
      Caption         =   "����"
      Height          =   1770
      Left            =   3870
      TabIndex        =   22
      Top             =   795
      Width           =   1320
      Begin VB.CheckBox chkType 
         Caption         =   "����(&W)"
         Height          =   195
         Index           =   4
         Left            =   165
         TabIndex        =   24
         Tag             =   "4"
         Top             =   1485
         Value           =   1  'Checked
         Width           =   990
      End
      Begin VB.CheckBox chkType 
         Caption         =   "����(&Q)"
         Height          =   195
         Index           =   3
         Left            =   165
         TabIndex        =   16
         Tag             =   "4"
         Top             =   1170
         Value           =   1  'Checked
         Width           =   990
      End
      Begin VB.CheckBox chkType 
         Caption         =   "ҩƷ(&Y)"
         Height          =   195
         Index           =   0
         Left            =   165
         TabIndex        =   13
         Tag             =   "1"
         Top             =   270
         Value           =   1  'Checked
         Width           =   990
      End
      Begin VB.CheckBox chkType 
         Caption         =   "����(&M)"
         Height          =   195
         Index           =   1
         Left            =   165
         TabIndex        =   14
         Tag             =   "2"
         Top             =   570
         Value           =   1  'Checked
         Width           =   990
      End
      Begin VB.CheckBox chkType 
         Caption         =   "�豸(&S)"
         Height          =   195
         Index           =   2
         Left            =   165
         TabIndex        =   15
         Tag             =   "4"
         Top             =   870
         Value           =   1  'Checked
         Width           =   990
      End
   End
   Begin VB.Frame fraTemp 
      Height          =   30
      Index           =   1
      Left            =   -75
      TabIndex        =   20
      Top             =   705
      Width           =   5415
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   4110
      TabIndex        =   18
      Top             =   2955
      Width           =   1100
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   2925
      TabIndex        =   17
      Top             =   2955
      Width           =   1100
   End
   Begin VB.Frame fraTemp 
      Height          =   30
      Index           =   0
      Left            =   -30
      TabIndex        =   19
      Top             =   2805
      Width           =   5415
   End
   Begin VB.TextBox Txt������ 
      Height          =   300
      Index           =   0
      Left            =   915
      TabIndex        =   7
      Top             =   1815
      Width           =   1215
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   6
      Left            =   2250
      TabIndex        =   11
      Top             =   2295
      Width           =   195
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "���ö�(&G)"
      Height          =   180
      Index           =   5
      Left            =   90
      TabIndex        =   9
      Top             =   2295
      Width           =   810
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   4
      Left            =   2250
      TabIndex        =   23
      Top             =   1875
      Width           =   195
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "������(&X)"
      Height          =   180
      Index           =   3
      Left            =   90
      TabIndex        =   6
      Top             =   1875
      Width           =   810
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "����(&N)"
      Height          =   180
      Index           =   2
      Left            =   270
      TabIndex        =   4
      Top             =   1440
      Width           =   630
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   1
      Left            =   2250
      TabIndex        =   2
      Top             =   1020
      Width           =   195
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "����(&D)"
      Height          =   180
      Index           =   0
      Left            =   270
      TabIndex        =   0
      Top             =   1020
      Width           =   630
   End
   Begin VB.Label Label1 
      Caption         =   "���������������������Ҫ�������ݡ�"
      Height          =   285
      Left            =   840
      TabIndex        =   21
      Top             =   375
      Width           =   4110
   End
   Begin VB.Image img���� 
      Height          =   480
      Left            =   195
      Picture         =   "frm��Ӧ�̹���.frx":0000
      Top             =   150
      Width           =   480
   End
End
Attribute VB_Name = "frm��Ӧ�̹���"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mblnCancel As Boolean
Private mstrFilter As String
Dim mstrPrivs As String
Private Const mlngModule = 1025
Private mcllFilter As Collection

Private Sub chkType_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub cmdCancel_Click()
    mblnCancel = True
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim intByte As Integer
    mblnCancel = False
    mstrFilter = ""
    'by lesfeng 2009-12-2 �����Ż� �����ڼ����ö� ���˴������� ���Ѿ��޸�
    Set mcllFilter = New Collection
    mcllFilter.Add Array("", ""), "����"
    mcllFilter.Add "", "����"
    mcllFilter.Add Array("0", "0"), "������"
    mcllFilter.Add Array("0", "0"), "���ö�"
    
    If Trim(TxtCode(0).Text) <> "" And Trim(TxtCode(1).Text) = "" Then
        mstrFilter = mstrFilter & " and ����>=[1]"
    ElseIf Trim(TxtCode(1).Text) = "" And Trim(TxtCode(1).Text) <> "" Then
        mstrFilter = mstrFilter & " and ����<=[2]"
    ElseIf Trim(TxtCode(1).Text) <> "" And Trim(TxtCode(1).Text) <> "" Then
        mstrFilter = mstrFilter & " and ����>=[1] and ����<=[2]"
    End If
    
    mcllFilter.Remove "����"
    mcllFilter.Add Array(Trim(TxtCode(0).Text), Trim(TxtCode(1).Text)), "����"
    
    If Trim(TxtName.Text) <> "" Then
        mstrFilter = mstrFilter & " and ���� like [3]"
        mcllFilter.Remove "����"
        mcllFilter.Add GetMatchingSting(TxtName.Text), "����"
    End If
    
    If Trim(Txt������(0).Text) <> "" And Trim(Txt������(1).Text) = "" Then
        mstrFilter = mstrFilter & " and ������>=[4]"
    ElseIf Trim(Txt������(1).Text) = "" And Trim(Txt������(1).Text) <> "" Then
        mstrFilter = mstrFilter & " and ������<=[5]"
    ElseIf Trim(Txt������(0).Text) <> "" And Trim(Txt������(1).Text) <> "" Then
        mstrFilter = mstrFilter & " and ������>=[4] and ������<=[5]"
    End If
    mcllFilter.Remove "������"
    mcllFilter.Add Array(Val(Txt������(0).Text), Val(Txt������(1).Text)), "������"
    
    If Trim(Txt���ö�(0).Text) <> "" And Trim(Txt���ö�(1).Text) = "" Then
        mstrFilter = mstrFilter & " and ���ö�>=[6]"
    ElseIf Trim(Txt���ö�(1).Text) = "" And Trim(Txt���ö�(1).Text) <> "" Then
        mstrFilter = mstrFilter & " and ���ö�<=[7]"
    ElseIf Trim(Txt���ö�(0).Text) <> "" And Trim(Txt���ö�(1).Text) <> "" Then
        mstrFilter = mstrFilter & " and ���ö�>=[6] and ���ö�<=[7]"
    End If
    mcllFilter.Remove "���ö�"
    mcllFilter.Add Array(Val(Txt���ö�(0).Text), Val(Txt���ö�(1).Text)), "���ö�"
    
    Dim i As Long
    Dim str���� As String
    Dim strTmp As String
    
    str���� = ""
    strTmp = ""
    For i = 0 To 4
        If chkType(i).Value = 1 And chkType(i).Enabled = True Then
            str���� = str���� & " or substr(����," & i + 1 & ",1)=1"
            strTmp = strTmp & "1"
        Else
            strTmp = strTmp & "0"
        End If
    Next
    
    Call zlDatabase.SetPara("��Ӧ������", strTmp, glngSys, mlngModule)
 
    If str���� <> "" And str���� <> "00000" Then
        str���� = " And (" & Mid(str����, 4) & ") "
    Else
        '����
        str���� = ""
    End If
    
    
    mstrFilter = mstrFilter & str����
    If mstrFilter <> "" Then
        mstrFilter = Mid(mstrFilter, 5)
    End If
    Unload Me
End Sub


Private Sub TxtCode_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub
Private Sub TxtCode_KeyPress(Index As Integer, KeyAscii As Integer)
    zlControl.TxtCheckKeyPress Txt������, KeyAscii, m����ʽ
End Sub

Private Sub TxtCode_LostFocus(Index As Integer)
    ImeLanguage False
End Sub

Private Sub TxtName_GotFocus()
    SetTxtGotFocus TxtName, True
End Sub

Private Sub TxtName_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub TxtName_KeyPress(KeyAscii As Integer)
    zlControl.TxtCheckKeyPress Txt������, KeyAscii, m�ı�ʽ
End Sub

Private Sub TxtName_LostFocus()
    ImeLanguage False
End Sub

Private Sub Txt���ö�_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub
Private Sub Txt���ö�_KeyPress(Index As Integer, KeyAscii As Integer)
    zlControl.TxtCheckKeyPress Txt������, KeyAscii, m���ʽ
End Sub

Private Sub Txt���ö�_LostFocus(Index As Integer)
    ImeLanguage False
End Sub

Private Sub Txt������_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub Txt������_KeyPress(Index As Integer, KeyAscii As Integer)
    zlControl.TxtCheckKeyPress Txt������, KeyAscii, m���ʽ
End Sub

Public Sub GetFiler(ByVal FrmMain As Object, ByRef blnCancel As Boolean, ByRef strFilter As String, ByRef cllFilter As Collection, Optional ByVal strPriv As String)
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:��ȡ����
    '--�����:frmMain-������
    '
    '--������:blnCancel-ȡ��
    '         strFilter-����
    '--��  ��:
    '-----------------------------------------------------------------------------------------------------------
    Dim strReg As String
    Dim i As Integer
    mstrPrivs = strPriv
    strReg = zlDatabase.GetPara("��Ӧ������", glngSys, mlngModule)
    If strReg = "" Then
        strReg = "00000"
    End If
    Err = 0
    On Error Resume Next
    For i = 1 To Len(strReg)
        If Mid(strReg, i, 1) = 1 Then
            chkType(i - 1).Value = 1
        Else
            chkType(i - 1).Value = 0
        End If
    Next
    Call Ȩ�޿���
    Me.Show 1, FrmMain
    blnCancel = mblnCancel
    strFilter = mstrFilter
    Set cllFilter = mcllFilter
End Sub

Private Sub Txt������_LostFocus(Index As Integer)
    ImeLanguage False
End Sub


Private Sub Ȩ�޿���()
    'Ȩ�޿���
    Dim blnҩƷ As Boolean
    Dim bln���� As Boolean
    Dim bln�豸 As Boolean
    Dim bln���� As Boolean
    Dim bln���� As Boolean
    
    blnҩƷ = InStr(1, mstrPrivs, "ҩƷ��Ӧ��") <> 0
    bln���� = InStr(1, mstrPrivs, "���ʹ�Ӧ��") <> 0
    bln�豸 = InStr(1, mstrPrivs, "�豸��Ӧ��") <> 0
    bln���� = InStr(1, mstrPrivs, "������Ӧ��") <> 0
    bln���� = InStr(1, mstrPrivs, "���Ĺ�Ӧ��") <> 0
    
    chkType(0).Enabled = blnҩƷ
    chkType(1).Enabled = bln����
    chkType(2).Enabled = bln�豸
    chkType(3).Enabled = bln����
    chkType(4).Enabled = bln����
End Sub

