VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form frmSetup 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��������"
   ClientHeight    =   6930
   ClientLeft      =   -15
   ClientTop       =   375
   ClientWidth     =   6855
   Icon            =   "frmSetup.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6930
   ScaleWidth      =   6855
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.OptionButton optCallWay 
      Caption         =   "Զ����������"
      Height          =   360
      Index           =   0
      Left            =   255
      TabIndex        =   2
      Top             =   5400
      Width           =   1410
   End
   Begin VB.Frame framԶ���������� 
      Height          =   855
      Left            =   135
      TabIndex        =   3
      Top             =   5445
      Width           =   6615
      Begin VB.ComboBox cboRemotePlaykStation 
         Height          =   300
         Left            =   1320
         TabIndex        =   4
         Top             =   360
         Width           =   5205
      End
      Begin VB.Label Label14 
         Caption         =   "Զ��վ������"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   400
         Width           =   1215
      End
   End
   Begin VB.OptionButton optCallWay 
      Caption         =   "������������"
      Height          =   270
      Index           =   1
      Left            =   255
      TabIndex        =   6
      Top             =   105
      Value           =   -1  'True
      Width           =   1410
   End
   Begin VB.Frame frm�����㲥���� 
      Height          =   5205
      Left            =   135
      TabIndex        =   19
      Top             =   120
      Width           =   6615
      Begin VB.CheckBox chkHintSound 
         Caption         =   "����ǰ������ʾ��"
         Height          =   240
         Left            =   1665
         TabIndex        =   22
         Top             =   375
         Width           =   1860
      End
      Begin VB.CheckBox chkUseSound 
         Caption         =   "������������"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   360
         Width           =   1455
      End
      Begin RichTextLib.RichTextBox rtbVBS 
         Height          =   3255
         Left            =   390
         TabIndex        =   7
         Top             =   1860
         Width           =   6105
         _ExtentX        =   10769
         _ExtentY        =   5741
         _Version        =   393217
         BorderStyle     =   0
         ScrollBars      =   3
         TextRTF         =   $"frmSetup.frx":06EA
      End
      Begin VB.TextBox txt�㲥ʱ�䳤�� 
         Height          =   270
         Left            =   1800
         TabIndex        =   8
         Top             =   825
         Width           =   615
      End
      Begin VB.TextBox txtPlayCount 
         Height          =   270
         Left            =   4935
         TabIndex        =   9
         Top             =   825
         Width           =   615
      End
      Begin VB.ComboBox cboSoundType 
         Height          =   300
         ItemData        =   "frmSetup.frx":0787
         Left            =   4530
         List            =   "frmSetup.frx":0789
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   330
         Width           =   1995
      End
      Begin VB.TextBox txtSpeed 
         Height          =   270
         Left            =   1425
         TabIndex        =   11
         Text            =   "6"
         Top             =   1170
         Width           =   495
      End
      Begin VB.TextBox txtLoopQueryTime 
         Height          =   270
         Left            =   5685
         MaxLength       =   3
         TabIndex        =   12
         Text            =   "30"
         Top             =   1170
         Width           =   585
      End
      Begin VB.Label Label1 
         Caption         =   "�Զ�����нű��༭��"
         Height          =   225
         Left            =   120
         TabIndex        =   21
         Top             =   1620
         Width           =   1860
      End
      Begin VB.Label Label13 
         Caption         =   "���������ٶ�Ϊ"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   1200
         Width           =   1875
      End
      Begin VB.Label Label12 
         Caption         =   "ÿ����������ʱ��Ϊ        ��       ����ѭ�����Ŵ���Ϊ        ��"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   855
         Width           =   5775
      End
      Begin VB.Label Label11 
         Caption         =   "�������ͣ�"
         Height          =   210
         Left            =   3630
         TabIndex        =   15
         Top             =   390
         Width           =   915
      End
      Begin VB.Label Label10 
         Caption         =   "(-10��10֮��) "
         Height          =   255
         Left            =   1965
         TabIndex        =   16
         Top             =   1215
         Width           =   1260
      End
      Begin VB.Label Label9 
         Caption         =   "��"
         Height          =   255
         Left            =   6300
         TabIndex        =   17
         Top             =   1215
         Width           =   255
      End
      Begin VB.Label Label2 
         Caption         =   "��ѯ������������ʱ����Ϊ"
         Height          =   255
         Left            =   3285
         TabIndex        =   18
         Top             =   1215
         Width           =   2400
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ��(&C)"
      Height          =   400
      Left            =   5655
      TabIndex        =   1
      Top             =   6405
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&S)"
      Height          =   400
      Left            =   4470
      TabIndex        =   0
      Top             =   6405
      Width           =   1100
   End
End
Attribute VB_Name = "frmSetup"
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
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub chkUseSound_Click()
On Error GoTo errHandle
    Dim blnUseLocalPlay As Boolean
    Dim lngBackColor As Long
    
    blnUseLocalPlay = IIf(chkUseSound.value <> 0, True, False)
    lngBackColor = IIf(chkUseSound.value <> 0, &H80000005, Me.BackColor)
    
    Label10.Enabled = blnUseLocalPlay
    Label11.Enabled = blnUseLocalPlay
    Label12.Enabled = blnUseLocalPlay
    Label13.Enabled = blnUseLocalPlay
    Label2.Enabled = blnUseLocalPlay
    Label9.Enabled = blnUseLocalPlay
    
    txt�㲥ʱ�䳤��.Enabled = blnUseLocalPlay
    txt�㲥ʱ�䳤��.BackColor = lngBackColor
    
    txtPlayCount.Enabled = blnUseLocalPlay
    txtPlayCount.BackColor = lngBackColor
    
    txtSpeed.Enabled = blnUseLocalPlay
    txtSpeed.BackColor = lngBackColor
    
    txtLoopQueryTime.Enabled = blnUseLocalPlay
    txtLoopQueryTime.BackColor = lngBackColor
    
    cboSoundType.Enabled = blnUseLocalPlay
    cboSoundType.BackColor = lngBackColor

    rtbVBS.Enabled = blnUseLocalPlay
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub cmdCancel_Click()
    '�رմ���
    mblnOk = False
    Me.Hide
End Sub

Private Sub cmdOK_Click()
On Error GoTo errHandle
    '�����������
    
    If optCallWay(0).value = True Then
        If Trim(cboRemotePlaykStation.Text) = "" Then
            MsgBox "Զ��վ����������Ϊ�գ�������ѡ��", vbOKOnly, Me.Caption
            cboRemotePlaykStation.SetFocus
            Exit Sub
        End If
    End If
    
    
    If Val(txtLoopQueryTime.Text) > 65 Then
        MsgBox "��ѯ������������ʱ�������ܴ���65�룬���������á�", vbOKOnly, Me.Caption
        
        txtLoopQueryTime.SetFocus
        Call zlControl.TxtSelAll(txtLoopQueryTime)
    End If

    SaveSetting "ZLSOFT", gstrRegPath, "��������ʱ��", Val(txt�㲥ʱ�䳤��.Text)            'Call zlDatabase.SetPara("�����㲥ʱ�䳤��", Val(txt�㲥ʱ�䳤��.Text), glngSys, glngModul)
    SaveSetting "ZLSOFT", gstrRegPath, "������������", Val(txtSpeed.Text)                   'Call zlDatabase.SetPara("�����㲥����", Val(txtSpeed.Text), glngSys, glngModul)

    SaveSetting "ZLSOFT", gstrRegPath, "������������", chkUseSound.value                    'Call zlDatabase.SetPara("������������", chkUseSound.value, glngSys, glngModul)
    SaveSetting "ZLSOFT", gstrRegPath, "��������ǰ������ʾ��", chkHintSound.value
    
    SaveSetting "ZLSOFT", gstrRegPath, "Զ�˺���վ��", cboRemotePlaykStation.Text           'Call zlDatabase.SetPara("Զ�˺���վ��", cboWorkStation.Text, glngSys, glngModul)
    SaveSetting "ZLSOFT", gstrRegPath, "�������Ŵ���", IIf(Val(txtPlayCount.Text) <= 0, 0, Val(txtPlayCount.Text))   'Call zlDatabase.SetPara("�������Ŵ���", Val(txtPlayCount.Text), glngSys, glngModul)

    '��������
    SaveSetting "ZLSOFT", gstrRegPath, "��������", cboSoundType.Text                        'Call zlDatabase.SetPara("��������", cboSoundType.Text, glngSys, glngModul)
    '��ѯʱ��
    SaveSetting "ZLSOFT", gstrRegPath, "��ѯ���ʱ��", IIf(Val(txtLoopQueryTime.Text) <= 0, 30, Val(txtLoopQueryTime.Text))     'Call zlDatabase.SetPara("��ѯʱ��", Val(txtLoopQueryTime.Text), glngSys, glngModul)

    SaveSetting "ZLSOFT", gstrRegPath, "����VBS�Զ������", IIf(Trim(cboSoundType.Text) = "�Զ���ű�����", 1, 0)
    SaveSetting "ZLSOFT", gstrRegPath, "VBS�ű�", rtbVBS.Text
    
    '����кŷ�ʽ
    If optCallWay(0).value Then
        SaveSetting "ZLSOFT", gstrRegPath, "���ŷ�ʽ", 1                                    'Call zlDatabase.SetPara("�кŷ�ʽ", 1, glngSys, glngModul)
    Else
        SaveSetting "ZLSOFT", gstrRegPath, "���ŷ�ʽ", 0                                    'Call zlDatabase.SetPara("�кŷ�ʽ", 0, glngSys, glngModul)
    End If

    mblnOk = True
    '�رմ���
    Unload Me
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub ReadWorkStationInf()
'*****************************************************
'��ȡվ����Ϣ
'*****************************************************

    Dim strSql As String
    Dim rsTemp As ADODB.Recordset

    strSql = "select ����վ from zlClients where ��ֹʹ��<>1 order by ����վ"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "��ȡվ����Ϣ")

    cboRemotePlaykStation.Clear
    If rsTemp.EOF Then Exit Sub

    While Not rsTemp.EOF
        Call cboRemotePlaykStation.AddItem(rsTemp("����վ"))
        rsTemp.MoveNext
    Wend
    
End Sub

Private Sub LoadMSSoundType()
    Dim objVoice As Object
    Dim objToken As Object
    
    Set objVoice = CreateObject("SAPI.SPVoice")

    cboSoundType.Clear
    If objVoice Is Nothing Then Exit Sub
    
    For Each objToken In objVoice.GetVoices()
        cboSoundType.AddItem objToken.GetAttribute("Name")
    Next
    
    cboSoundType.AddItem "�Զ���ű�����"
    
    cboSoundType.ListIndex = 0
End Sub

Private Sub ReadLocalPara()
    Dim i As Integer
    Dim lng�㲥���� As Long
    Dim strSoundType As String

   '��ȡ�кŷ�ʽ
    cboRemotePlaykStation.Text = GetSetting("ZLSOFT", gstrRegPath, "Զ�˺���վ��", "")     'zlDatabase.GetPara("Զ�˺���վ��", glngSys, glngModule, "")
    cboRemotePlaykStation.Enabled = optCallWay(0).value

    chkUseSound.value = Val(GetSetting("ZLSOFT", gstrRegPath, "������������", 1))   'zlDatabase.GetPara("������������", glngSys, glngModul, "1")

    txtLoopQueryTime.Text = Val(GetSetting("ZLSOFT", gstrRegPath, "��ѯ���ʱ��", 30))      ' Val(zlDatabase.GetPara("��ѯʱ��", glngSys, glngModul, "30"))
    txt�㲥ʱ�䳤��.Text = Val(GetSetting("ZLSOFT", gstrRegPath, "��������ʱ��", 15))       'Val(zlDatabase.GetPara("�����㲥ʱ�䳤��", glngSys, glngModul, "15"))
    txtPlayCount.Text = Val(GetSetting("ZLSOFT", gstrRegPath, "�������Ŵ���", 2))           'Val(zlDatabase.GetPara("�������Ŵ���", glngSys, glngModul, "3"))

    If cboSoundType.Enabled = True Then                                                     'zlDatabase.GetPara("��������", glngSys, glngModul, "ϵͳĬ��")
        strSoundType = Trim(GetSetting("ZLSOFT", gstrRegPath, "��������", ""))
        
        For i = 0 To cboSoundType.ListCount - 1
            If cboSoundType.List(i) = strSoundType Then
                cboSoundType.ListIndex = i
                Exit For
            End If
        Next
        
        If cboSoundType.ListCount > 0 And cboSoundType.ListIndex < 0 Then cboSoundType.ListIndex = 0
    End If
    
    chkHintSound.value = Val(GetSetting("ZLSOFT", gstrRegPath, "��������ǰ������ʾ��", ""))
    
    lng�㲥���� = Val(GetSetting("ZLSOFT", gstrRegPath, "������������", 0))                 'Val(zlDatabase.GetPara("�����㲥����", glngSys, glngModul, "0"))
    txtSpeed.Text = IIf(lng�㲥���� <= 10 And lng�㲥���� >= -10, lng�㲥����, 0)
    
    rtbVBS.Text = GetSetting("ZLSOFT", gstrRegPath, "VBS�ű�", M_STR_DEFAULT_VBS)


    rtbVBS.Enabled = IIf(cboSoundType.Text = "�Զ���ű�����", True, False)

    optCallWay(0).value = Val(GetSetting("ZLSOFT", gstrRegPath, "���ŷ�ʽ", 1))     'Val(zlDatabase.GetPara("�кŷ�ʽ", glngSys, glngModule, "0"))
    optCallWay(1).value = Not optCallWay(0).value
    
    Call optCallWay_Click(0)
End Sub

Private Sub Form_Load()
    Call LoadMSSoundType
    
    Call ReadWorkStationInf
    
    Call ReadLocalPara
End Sub


Private Sub optCallWay_Click(Index As Integer)
    
    chkUseSound.Enabled = Not optCallWay(0).value
    cboRemotePlaykStation.Enabled = optCallWay(0).value
    
    If optCallWay(0).value Then
        frm�����㲥����.Enabled = False
        
        txt�㲥ʱ�䳤��.BackColor = Me.BackColor
        txtPlayCount.BackColor = Me.BackColor
        txtSpeed.BackColor = Me.BackColor
        rtbVBS.BackColor = Me.BackColor
        txtLoopQueryTime.BackColor = Me.BackColor
        cboSoundType.BackColor = Me.BackColor
        
        cboRemotePlaykStation.BackColor = &H80000005
    Else
        frm�����㲥����.Enabled = True
        
        txt�㲥ʱ�䳤��.BackColor = &H80000005
        txtPlayCount.BackColor = &H80000005
        txtSpeed.BackColor = &H80000005
        rtbVBS.BackColor = &H80000005
        txtLoopQueryTime.BackColor = &H80000005
        cboSoundType.BackColor = &H80000005
        
        cboRemotePlaykStation.BackColor = Me.BackColor
    End If
End Sub


Private Sub txt�㲥ʱ�䳤��_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub
