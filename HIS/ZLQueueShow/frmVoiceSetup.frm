VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form frmVoiceSetup 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��������"
   ClientHeight    =   6090
   ClientLeft      =   -15
   ClientTop       =   375
   ClientWidth     =   6855
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6090
   ScaleWidth      =   6855
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.Frame frm�����㲥���� 
      Height          =   5325
      Left            =   135
      TabIndex        =   14
      Top             =   90
      Width           =   6615
      Begin VB.CheckBox chkHintSound 
         Caption         =   "����ǰ������ʾ��"
         Height          =   240
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   1860
      End
      Begin RichTextLib.RichTextBox rtbVBS 
         Height          =   3495
         Left            =   120
         TabIndex        =   2
         Top             =   1680
         Width           =   6345
         _ExtentX        =   11192
         _ExtentY        =   6165
         _Version        =   393217
         BorderStyle     =   0
         Enabled         =   -1  'True
         ScrollBars      =   3
         TextRTF         =   $"frmVoiceSetup.frx":0000
      End
      Begin VB.TextBox txt�㲥ʱ�䳤�� 
         Height          =   270
         Left            =   1800
         TabIndex        =   3
         Top             =   570
         Width           =   615
      End
      Begin VB.TextBox txtPlayCount 
         Height          =   270
         Left            =   4920
         TabIndex        =   4
         Top             =   570
         Width           =   615
      End
      Begin VB.ComboBox cboSoundType 
         Height          =   300
         Left            =   4200
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   200
         Width           =   2355
      End
      Begin VB.TextBox txtSpeed 
         Height          =   270
         Left            =   1425
         TabIndex        =   6
         Text            =   "6"
         Top             =   940
         Width           =   495
      End
      Begin VB.TextBox txtLoopQueryTime 
         Height          =   270
         Left            =   5685
         MaxLength       =   3
         TabIndex        =   7
         Text            =   "30"
         Top             =   940
         Width           =   585
      End
      Begin VB.Label Label1 
         Caption         =   "�Զ�����нű��༭��"
         Height          =   225
         Left            =   120
         TabIndex        =   15
         Top             =   1320
         Width           =   1860
      End
      Begin VB.Label Label13 
         Caption         =   "���������ٶ�Ϊ"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   955
         Width           =   1875
      End
      Begin VB.Label Label12 
         Caption         =   "ÿ����������ʱ��Ϊ        ��       ����ѭ�����Ŵ���Ϊ        ��"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   600
         Width           =   6255
      End
      Begin VB.Label Label11 
         Caption         =   "�������ͣ�"
         Height          =   210
         Left            =   3280
         TabIndex        =   10
         Top             =   240
         Width           =   915
      End
      Begin VB.Label Label10 
         Caption         =   "(-10��10֮��) "
         Height          =   255
         Left            =   1965
         TabIndex        =   11
         Top             =   960
         Width           =   1260
      End
      Begin VB.Label Label9 
         Caption         =   "��"
         Height          =   255
         Left            =   6300
         TabIndex        =   12
         Top             =   960
         Width           =   255
      End
      Begin VB.Label Label2 
         Caption         =   "��ѯ������������ʱ����Ϊ"
         Height          =   255
         Left            =   3285
         TabIndex        =   13
         Top             =   965
         Width           =   2400
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ��(&C)"
      Height          =   400
      Left            =   5655
      TabIndex        =   1
      Top             =   5565
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&S)"
      Height          =   400
      Left            =   4470
      TabIndex        =   0
      Top             =   5565
      Width           =   1100
   End
End
Attribute VB_Name = "frmVoiceSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Const M_STR_DEFAULT_VBS As String = "sub CusVoicePlay(lngCallId,strCallContext)" & vbCrLf & _
                                            "    Dim i                                          " & vbCrLf & _
                                            "                                                   " & vbCrLf & _
                                            "    SpVoice.Rate = 0                               " & vbCrLf & _
                                            "    SpVoice.Volume = 100                           " & vbCrLf & _
                                            "                                                   " & vbCrLf & _
                                            "    'Lili�������ĺ�Ӣ��                            " & vbCrLf & _
                                            "    Set SpVoice.Voice = SpVoice.GetVoices(""" & "Name=Microsoft Lili" & """).Item(0)" & vbCrLf & _
                                            "    SpVoice.Speak strCallContext, 1                " & vbCrLf & _
                                            "                                                   " & vbCrLf & _
                                            "    'Annaֻ�ܺ���Ӣ��                              " & vbCrLf & _
                                            "    Set SpVoice.Voice = SpVoice.GetVoices(""" & "Name=Microsoft Anna" & """).Item(0)" & vbCrLf & _
                                            "    SpVoice.Speak strCallContext, 1                " & vbCrLf & _
                                            "End Sub                                            "

Private mlngModule As Long
Private mblnOk As Boolean



Public Function ShowMe(objOwner As Object) As Boolean
    mblnOk = False
    Call Me.Show(1, objOwner)
    
    ShowMe = mblnOk
End Function

Private Sub cbo����_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub cboSoundType_Click()
On Error GoTo errHandle
    If cboSoundType.Text = "�Զ���ű�����" Then
        rtbVBS.Enabled = True
    Else
        rtbVBS.Enabled = False
    End If
Exit Sub
errHandle:
    If gobjComLib.ErrCenter() = 1 Then Resume
End Sub

Private Sub cmdCancel_Click()
    '�رմ���
    mblnOk = False
    Me.Hide
End Sub

Private Sub cmdOK_Click()
On Error GoTo errHandle
    '�����������
    If Val(txtLoopQueryTime.Text) > 65 Then
        MsgBox "��ѯ������������ʱ�������ܴ���65�룬���������á�", vbOKOnly, Me.Caption
        
        txtLoopQueryTime.SetFocus
        txtLoopQueryTime.SelStart = 0
        txtLoopQueryTime.SelLength = Len(txtLoopQueryTime.Text)
        
        Exit Sub
    End If

    SaveSetting "ZLSOFT", G_STR_REGPATH, "��������ʱ��", Val(txt�㲥ʱ�䳤��.Text)            'Call zlDatabase.SetPara("�����㲥ʱ�䳤��", Val(txt�㲥ʱ�䳤��.Text), glngSys, glngModul)
    SaveSetting "ZLSOFT", G_STR_REGPATH, "������������", Val(txtSpeed.Text)                   'Call zlDatabase.SetPara("�����㲥����", Val(txtSpeed.Text), glngSys, glngModul)

    SaveSetting "ZLSOFT", G_STR_REGPATH, "��������ǰ������ʾ��", chkHintSound.value
    
    SaveSetting "ZLSOFT", G_STR_REGPATH, "�������Ŵ���", IIf(Val(txtPlayCount.Text) <= 0, 0, Val(txtPlayCount.Text))   'Call zlDatabase.SetPara("�������Ŵ���", Val(txtPlayCount.Text), glngSys, glngModul)

    '��������
    SaveSetting "ZLSOFT", G_STR_REGPATH, "��������", cboSoundType.Text                        'Call zlDatabase.SetPara("��������", cboSoundType.Text, glngSys, glngModul)
    '��ѯʱ��
    SaveSetting "ZLSOFT", G_STR_REGPATH, "��ѯ���ʱ��", IIf(Val(txtLoopQueryTime.Text) <= 0, 30, Val(txtLoopQueryTime.Text))     'Call zlDatabase.SetPara("��ѯʱ��", Val(txtLoopQueryTime.Text), glngSys, glngModul)

    SaveSetting "ZLSOFT", G_STR_REGPATH, "����VBS�Զ������", IIf(Trim(cboSoundType.Text) = "�Զ���ű�����", 1, 0)
    SaveSetting "ZLSOFT", G_STR_REGPATH, "VBS�ű�", rtbVBS.Text
    
    mblnOk = True
    '�رմ���
    Unload Me
Exit Sub
errHandle:
    If gobjComLib.ErrCenter() = 1 Then Resume
End Sub

Private Sub LoadMSSoundType()
    Dim objVoice As Object
    Dim objToken As Object
    
    Set objVoice = CreateObject("SAPI.SPVoice")

    cboSoundType.Clear
    cboSoundType.AddItem ""
    
    If objVoice Is Nothing Then Exit Sub
    
    For Each objToken In objVoice.GetVoices()
        cboSoundType.AddItem objToken.GetAttribute("Name")
    Next
    
    cboSoundType.AddItem "�Զ���ű�����"
    
    cboSoundType.ListIndex = 0
End Sub

Private Sub ReadLocalPara()
    Dim lng�㲥���� As Long
    Dim strSoundType As String
    Dim i As Integer

    txtLoopQueryTime.Text = Val(GetSetting("ZLSOFT", G_STR_REGPATH, "��ѯ���ʱ��", 30))      ' Val(zlDatabase.GetPara("��ѯʱ��", glngSys, glngModul, "30"))
    txt�㲥ʱ�䳤��.Text = Val(GetSetting("ZLSOFT", G_STR_REGPATH, "��������ʱ��", 15))       'Val(zlDatabase.GetPara("�����㲥ʱ�䳤��", glngSys, glngModul, "15"))
    txtPlayCount.Text = Val(GetSetting("ZLSOFT", G_STR_REGPATH, "�������Ŵ���", 2))           'Val(zlDatabase.GetPara("�������Ŵ���", glngSys, glngModul, "3"))

    strSoundType = GetSetting("ZLSOFT", G_STR_REGPATH, "��������", "")           'zlDatabase.GetPara("��������", glngSys, glngModul, "ϵͳĬ��")
    For i = 0 To cboSoundType.ListCount - 1
        If cboSoundType.List(i) = strSoundType Then
            cboSoundType.ListIndex = i
            Exit For
        End If
    Next
    
    chkHintSound.value = Val(GetSetting("ZLSOFT", G_STR_REGPATH, "��������ǰ������ʾ��", ""))
    
    lng�㲥���� = Val(GetSetting("ZLSOFT", G_STR_REGPATH, "������������", 0))                 'Val(zlDatabase.GetPara("�����㲥����", glngSys, glngModul, "0"))
    txtSpeed.Text = IIf(lng�㲥���� <= 10 And lng�㲥���� >= -10, lng�㲥����, 0)
    
    rtbVBS.Text = GetSetting("ZLSOFT", G_STR_REGPATH, "VBS�ű�", M_STR_DEFAULT_VBS)


    rtbVBS.Enabled = IIf(cboSoundType.Text = "�Զ���ű�����", True, False)
End Sub

Private Sub Form_Load()
    Call LoadMSSoundType

    Call ReadLocalPara
End Sub

Private Sub txt�㲥ʱ�䳤��_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub
