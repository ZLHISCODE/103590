VERSION 5.00
Begin VB.Form frmSetup 
   Caption         =   "��������"
   ClientHeight    =   9135
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5580
   Icon            =   "frmSetup.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9135
   ScaleWidth      =   5580
   StartUpPosition =   1  '����������
   Begin VB.CheckBox chkOrderStyle 
      Caption         =   "ʹ������ԭʼ˳������"
      Height          =   255
      Left            =   2880
      TabIndex        =   48
      Top             =   8280
      Width           =   2175
   End
   Begin VB.Frame Frame1 
      Caption         =   "�Ŷӷ�����ʾ����"
      Height          =   975
      Left            =   240
      TabIndex        =   45
      Top             =   4800
      Width           =   5175
      Begin VB.OptionButton optGroupType 
         Caption         =   "�����ҷ���"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   49
         Top             =   600
         Width           =   1215
      End
      Begin VB.OptionButton optGroupType 
         Caption         =   "��ҽ����������"
         Height          =   255
         Index           =   1
         Left            =   2880
         TabIndex        =   47
         Top             =   240
         Width           =   1575
      End
      Begin VB.OptionButton optGroupType 
         Caption         =   "���������Ʒ���"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   46
         Top             =   240
         Value           =   -1  'True
         Width           =   1575
      End
   End
   Begin VB.Frame framCalledColumn 
      Caption         =   "�Ѻ���������"
      Height          =   1095
      Left            =   240
      TabIndex        =   37
      Top             =   7080
      Width           =   5175
      Begin VB.CheckBox chkCalledColumn 
         Caption         =   "ҽ������"
         Height          =   255
         Index           =   5
         Left            =   3840
         TabIndex        =   44
         Tag             =   "ҽ������"
         Top             =   360
         Width           =   1095
      End
      Begin VB.CheckBox chkCalledColumn 
         Caption         =   "����"
         Height          =   255
         Index           =   4
         Left            =   2640
         TabIndex        =   43
         Tag             =   "����"
         Top             =   360
         Width           =   855
      End
      Begin VB.CheckBox chkCalledColumn 
         Caption         =   "����"
         Enabled         =   0   'False
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   42
         Tag             =   "����"
         Top             =   360
         Value           =   1  'Checked
         Width           =   975
      End
      Begin VB.CheckBox chkCalledColumn 
         Caption         =   "��������"
         Enabled         =   0   'False
         Height          =   255
         Index           =   1
         Left            =   1320
         TabIndex        =   41
         Tag             =   "��������"
         Top             =   360
         Value           =   1  'Checked
         Width           =   1215
      End
      Begin VB.CheckBox chkCalledColumn 
         Caption         =   "������"
         Height          =   255
         Index           =   2
         Left            =   1320
         TabIndex        =   40
         Tag             =   "����ҽ��"
         Top             =   720
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin VB.CheckBox chkCalledColumn 
         Caption         =   "����ʱ��"
         Height          =   255
         Index           =   3
         Left            =   2640
         TabIndex        =   39
         Tag             =   "����ʱ��"
         Top             =   720
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin VB.CheckBox chkCalledColumn 
         Caption         =   "�������"
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   38
         Tag             =   "�������"
         Top             =   720
         Width           =   1095
      End
   End
   Begin VB.ComboBox cbxComeback 
      Height          =   300
      ItemData        =   "frmSetup.frx":06EA
      Left            =   960
      List            =   "frmSetup.frx":06F4
      Style           =   2  'Dropdown List
      TabIndex        =   29
      Top             =   8235
      Width           =   975
   End
   Begin VB.Frame framColumn 
      Caption         =   "�Ŷ�������"
      Height          =   1095
      Left            =   240
      TabIndex        =   18
      Top             =   5880
      Width           =   5175
      Begin VB.CheckBox chkColumn 
         Caption         =   "�������"
         Height          =   255
         Index           =   7
         Left            =   3840
         TabIndex        =   27
         Tag             =   "�������"
         Top             =   720
         Width           =   1095
      End
      Begin VB.CheckBox chkColumn 
         Caption         =   "�Ŷ�ʱ��"
         Height          =   255
         Index           =   6
         Left            =   2640
         TabIndex        =   25
         Tag             =   "�Ŷ�ʱ��"
         Top             =   720
         Width           =   1095
      End
      Begin VB.CheckBox chkColumn 
         Caption         =   "�Ŷ�״̬"
         Height          =   255
         Index           =   5
         Left            =   1320
         TabIndex        =   24
         Tag             =   "�Ŷ�״̬"
         Top             =   720
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin VB.CheckBox chkColumn 
         Caption         =   "ҽ������"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   23
         Tag             =   "ҽ������"
         Top             =   720
         Width           =   1095
      End
      Begin VB.CheckBox chkColumn 
         Caption         =   "����"
         Height          =   255
         Index           =   3
         Left            =   3840
         TabIndex        =   22
         Tag             =   "����"
         Top             =   360
         Width           =   855
      End
      Begin VB.CheckBox chkColumn 
         Caption         =   "����"
         Height          =   255
         Index           =   2
         Left            =   2640
         TabIndex        =   21
         Tag             =   "����"
         Top             =   360
         Width           =   855
      End
      Begin VB.CheckBox chkColumn 
         Caption         =   "��������"
         Enabled         =   0   'False
         Height          =   255
         Index           =   1
         Left            =   1320
         TabIndex        =   20
         Tag             =   "��������"
         Top             =   360
         Value           =   1  'Checked
         Width           =   1335
      End
      Begin VB.CheckBox chkColumn 
         Caption         =   "����"
         Enabled         =   0   'False
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   19
         Tag             =   "����"
         Top             =   360
         Value           =   1  'Checked
         Width           =   975
      End
   End
   Begin VB.CheckBox chkUseDisplay 
      Caption         =   "��ʾ�ŶӶ���"
      Height          =   255
      Left            =   360
      TabIndex        =   11
      Top             =   120
      Width           =   1455
   End
   Begin VB.Frame Frame3 
      Caption         =   "�кŷ�ʽ����"
      Height          =   3375
      Left            =   240
      TabIndex        =   3
      Top             =   1320
      Width           =   5175
      Begin VB.OptionButton optCallWay 
         Caption         =   "����Զ������"
         Height          =   450
         Index           =   0
         Left            =   360
         TabIndex        =   5
         Top             =   2280
         Width           =   1455
      End
      Begin VB.CheckBox chkUseSound 
         Caption         =   "������������"
         Height          =   255
         Left            =   360
         TabIndex        =   17
         Top             =   720
         Width           =   1455
      End
      Begin VB.OptionButton optCallWay 
         Caption         =   "���ñ�������"
         Height          =   495
         Index           =   1
         Left            =   360
         TabIndex        =   4
         Top             =   240
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.Frame frm�����㲥���� 
         Height          =   1935
         Left            =   240
         TabIndex        =   12
         Top             =   360
         Width           =   4815
         Begin VB.TextBox txtLoopQueryTime 
            Height          =   270
            Left            =   1800
            MaxLength       =   2
            TabIndex        =   34
            Text            =   "30"
            Top             =   1560
            Width           =   615
         End
         Begin VB.TextBox txtSpeed 
            Height          =   270
            Left            =   1320
            TabIndex        =   32
            Text            =   "6"
            Top             =   1200
            Width           =   495
         End
         Begin VB.ComboBox cboSoundType 
            Height          =   300
            ItemData        =   "frmSetup.frx":0706
            Left            =   2760
            List            =   "frmSetup.frx":0710
            TabIndex        =   31
            Text            =   "cboSoundType"
            Top             =   340
            Width           =   1815
         End
         Begin VB.TextBox txtPlayCount 
            Height          =   270
            Left            =   3720
            TabIndex        =   16
            Top             =   840
            Width           =   615
         End
         Begin VB.TextBox txt�㲥ʱ�䳤�� 
            Height          =   270
            Left            =   1800
            TabIndex        =   13
            Top             =   840
            Width           =   615
         End
         Begin VB.Label Label7 
            Caption         =   "������ѯ���ʱ��Ϊ"
            Height          =   255
            Left            =   120
            TabIndex        =   36
            Top             =   1605
            Width           =   1695
         End
         Begin VB.Label Label8 
            Caption         =   "��"
            Height          =   255
            Left            =   2480
            TabIndex        =   35
            Top             =   1605
            Width           =   255
         End
         Begin VB.Label Label6 
            Caption         =   "(���ٷ�Χ��-10��10֮��) "
            Height          =   255
            Left            =   1800
            TabIndex        =   33
            Top             =   1230
            Width           =   2295
         End
         Begin VB.Label Label5 
            Caption         =   "�������ͣ�"
            Height          =   255
            Left            =   1920
            TabIndex        =   30
            Top             =   380
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "ÿ�������㲥����Ϊ        �� ���Ŵ���Ϊ        ��"
            Height          =   255
            Left            =   120
            TabIndex        =   15
            Top             =   855
            Width           =   4455
         End
         Begin VB.Label Label3 
            Caption         =   "�����㲥���٣�"
            Height          =   255
            Left            =   120
            TabIndex        =   14
            Top             =   1215
            Width           =   1755
         End
      End
      Begin VB.Frame Frame4 
         Height          =   855
         Left            =   240
         TabIndex        =   6
         Top             =   2400
         Width           =   4815
         Begin VB.ComboBox cboWorkStation 
            Height          =   300
            Left            =   1320
            TabIndex        =   26
            Top             =   360
            Width           =   3375
         End
         Begin VB.Label labRemoteComputerName 
            Caption         =   "Զ��վ������"
            Height          =   255
            Left            =   240
            TabIndex        =   7
            Top             =   400
            Width           =   1215
         End
      End
   End
   Begin VB.Frame frm��ʾ�豸���� 
      Height          =   1095
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   5175
      Begin VB.ComboBox cbo��ʾӲ����� 
         Height          =   300
         ItemData        =   "frmSetup.frx":0728
         Left            =   240
         List            =   "frmSetup.frx":072A
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   600
         Width           =   3495
      End
      Begin VB.CommandButton cmd��ʾ�豸���� 
         Caption         =   "�豸����"
         Height          =   300
         Left            =   3840
         TabIndex        =   8
         Top             =   600
         Width           =   1100
      End
      Begin VB.Label Label2 
         Caption         =   "��ʾ�豸���"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ��"
      Height          =   375
      Left            =   4320
      TabIndex        =   1
      Top             =   8640
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��"
      Height          =   375
      Left            =   3120
      TabIndex        =   0
      Top             =   8640
      Width           =   1100
   End
   Begin VB.Label labCallBack 
      Caption         =   "���ﲡ��           �Ŷ�"
      Height          =   255
      Left            =   240
      TabIndex        =   28
      Top             =   8280
      Width           =   2175
   End
End
Attribute VB_Name = "frmSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mstrReg As String


Private Sub cbo����_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub


Private Sub chkUseDisplay_Click()
    If chkUseDisplay.Value = 0 Then
        frm��ʾ�豸����.Enabled = False
        
        cbo��ʾӲ�����.BackColor = frm��ʾ�豸����.BackColor
    Else
        frm��ʾ�豸����.Enabled = True
        
        cbo��ʾӲ�����.BackColor = &H80000005
        
        
    End If
End Sub

Private Sub cmdCancel_Click()
    '�رմ���
    Unload Me
End Sub

Private Sub cmdOK_Click()
    '�����������
    Dim strRet As String
    Dim strColumnInf As String
    Dim strCalledColumnInf As String
    
    Dim i As Integer
    
    mstrReg = "����ȫ��\�Ŷӽк�"
    
    If Val(txtLoopQueryTime.Text) > 65 Then
        MsgBox "������ѯ���ʱ�䲻�ܴ���65�룬���������á�", vbOKOnly, Me.Caption
        
        txtLoopQueryTime.SetFocus
        Call zlControl.TxtSelAll(txtLoopQueryTime)
        
        Exit Sub
    End If
    
    'SaveSetting "ZLSOFT", strReg, "�����㲥ʱ�䳤��", Val(txt�㲥ʱ�䳤��.Text)
    Call zlDatabase.SetPara("�����㲥ʱ�䳤��", Val(txt�㲥ʱ�䳤��.Text), glngSys, glngModul)
    'SaveSetting "ZLSOFT", strReg, "�����㲥����", Val(cbo����.Text)
    Call zlDatabase.SetPara("�����㲥����", Val(txtSpeed.Text), glngSys, glngModul)
    'SaveSetting "ZLSOFT", strReg, "��ʾ�ŶӶ���", chkUseDisplay.Value
    Call zlDatabase.SetPara("��ʾ�ŶӶ���", chkUseDisplay.Value, glngSys, glngModul)
    'SaveSetting "ZLSOFT", strReg, "������������", chkUseSound.Value
    Call zlDatabase.SetPara("������������", chkUseSound.Value, glngSys, glngModul)
    
    'SaveSetting "ZLSOFT", strReg, "Զ��վ������", txtRemoteComputerName.Text
    Call zlDatabase.SetPara("Զ�˺���վ��", cboWorkStation.Text, glngSys, glngModul)
    'SaveSetting "ZLSOFT", strReg, "�������Ŵ���", txtPlayCount.Text
    Call zlDatabase.SetPara("�������Ŵ���", Val(txtPlayCount.Text), glngSys, glngModul)
    
    '��������
    Call zlDatabase.SetPara("��������", cboSoundType.Text, glngSys, glngModul)
    '��ѯʱ��
    Call zlDatabase.SetPara("��ѯʱ��", Val(txtLoopQueryTime.Text), glngSys, glngModul)
    
    strColumnInf = ""
    For i = 0 To 7
        If chkColumn(i).Value = vbChecked Then
            If Trim(strColumnInf) <> "" Then strColumnInf = strColumnInf & ","
            strColumnInf = strColumnInf & chkColumn(i).Tag
        End If
    Next i
    
    Call zlDatabase.SetPara("������ʾ��", strColumnInf, glngSys, glngModul)
    
    
    
    strCalledColumnInf = ""
    For i = 0 To 6
        If chkCalledColumn(i).Value = vbChecked Then
            If Trim(strCalledColumnInf) <> "" Then strCalledColumnInf = strCalledColumnInf & ","
            strCalledColumnInf = strCalledColumnInf & chkCalledColumn(i).Tag
        End If
    Next i
    
    Call zlDatabase.SetPara("����������ʾ��", strCalledColumnInf, glngSys, glngModul)
    
    
    '����кŷ�ʽ
    If optCallWay(0).Value Then
        'SaveSetting "ZLSOFT", strReg, "�кŷ�ʽ", 1
         Call zlDatabase.SetPara("�кŷ�ʽ", 1, glngSys, glngModul)
    Else
        'SaveSetting "ZLSOFT", strReg, "�кŷ�ʽ", 0
        Call zlDatabase.SetPara("�кŷ�ʽ", 0, glngSys, glngModul)
    End If
    
    
    '������ʾ�豸
    If cbo��ʾӲ�����.ListIndex <> -1 Then
        'SaveSetting "ZLSOFT", strReg, "��ʾ�豸���", cbo��ʾӲ�����.ItemData(cbo��ʾӲ�����.ListIndex)
        Call zlDatabase.SetPara("��ʾ�豸���", cbo��ʾӲ�����.ItemData(cbo��ʾӲ�����.ListIndex), glngSys, glngModul)
    End If
    
    Call zlDatabase.SetPara("���ﲡ���Ƿ�����", cbxComeback.ListIndex, glngSys, glngModul)
    
    For i = 0 To optGroupType.Count - 1
        If optGroupType(i).Value Then
            Call zlDatabase.SetPara("�Ŷӷ�������", i, glngSys, glngModul)
            Exit For
        End If
    Next
    
    Call zlDatabase.SetPara("ʹ������ԭʼ˳������", chkOrderStyle.Value, glngSys, glngModul)
    '�رմ���
    Unload Me
End Sub

Private Sub cmd��ʾ�豸����_Click()
    If pobjLEDShow Is Nothing Then
        Call frmQueueStation.InitLED(plngLEDModal)
    End If
        
    If Not pobjLEDShow Is Nothing Then
        Call pobjLEDShow.zlSetup(Me)
    End If
End Sub

Private Sub ReadWorkStationInf()
'*****************************************************
'��ȡվ����Ϣ
'*****************************************************

    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    
    strSql = "select ����վ from zlClients where ��ֹʹ��<>1 order by ����վ"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "��ȡվ����Ϣ")
    
    If rsTemp.EOF Then Exit Sub
    
    While Not rsTemp.EOF
        Call cboWorkStation.AddItem(rsTemp("����վ"))
        rsTemp.MoveNext
    Wend
    
End Sub

Private Sub ReadLocalPara()
    Dim lng�㲥���� As Long
    Dim lngLEDModal As Long
    Dim strColumnInf As String
    Dim strCalledColumnInf As String
    
    Dim i As Integer
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    
   '��ȡ�кŷ�ʽ
        
    'optCallWay(0).Value = Val(GetSetting("ZLSOFT", strReg, "�кŷ�ʽ", 0))
    optCallWay(0).Value = Val(zlDatabase.GetPara("�кŷ�ʽ", glngSys, glngModul, "0"))
    optCallWay(1).Value = Not optCallWay(0).Value
    
    'txtRemoteComputerName.Text = GetSetting("ZLSOFT", strReg, "Զ��վ������", "")
    cboWorkStation.Text = zlDatabase.GetPara("Զ�˺���վ��", glngSys, glngModul, "")
    cboWorkStation.Enabled = optCallWay(0).Value
    
    
    txtLoopQueryTime.Text = Val(zlDatabase.GetPara("��ѯʱ��", glngSys, glngModul, "30"))
    'txt�㲥ʱ�䳤��.Text = Val(GetSetting("ZLSOFT", strReg, "�����㲥ʱ�䳤��", 15))
    txt�㲥ʱ�䳤��.Text = Val(zlDatabase.GetPara("�����㲥ʱ�䳤��", glngSys, glngModul, "15"))
    'txtPlayCount.Text = Val(GetSetting("ZLSOFT", strReg, "�������Ŵ���", 3))
    txtPlayCount.Text = Val(zlDatabase.GetPara("�������Ŵ���", glngSys, glngModul, "3"))
    'lng�㲥���� = Val(GetSetting("ZLSOFT", strReg, "�����㲥����", 60))
    lng�㲥���� = Val(zlDatabase.GetPara("�����㲥����", glngSys, glngModul, "60"))
    
    cboSoundType.Text = zlDatabase.GetPara("��������", glngSys, glngModul, "ϵͳĬ��")
    
    strColumnInf = zlDatabase.GetPara("������ʾ��", glngSys, glngModul, ",����,��������,�Ŷ�״̬,")
    strColumnInf = Replace(strColumnInf, "��", ",")
    strColumnInf = "," & strColumnInf & ","
    
    For i = 0 To 7
        chkColumn(i).Value = Int(IIf(InStr(1, strColumnInf, "," & chkColumn(i).Tag & ",") > 0, vbChecked, vbUnchecked))
    Next i
    
    
    
    strCalledColumnInf = zlDatabase.GetPara("����������ʾ��", glngSys, glngModul, ",����,��������,")
    strCalledColumnInf = Replace(strCalledColumnInf, "��", ",")
    strCalledColumnInf = "," & strCalledColumnInf & ","
    
    For i = 0 To 6
        chkCalledColumn(i).Value = Int(IIf(InStr(1, strCalledColumnInf, "," & chkCalledColumn(i).Tag & ",") > 0, vbChecked, vbUnchecked))
    Next i
    
    If optCallWay(0).Value Then
        txt�㲥ʱ�䳤��.BackColor = Me.BackColor
        txtPlayCount.BackColor = Me.BackColor
        txtSpeed.BackColor = Me.BackColor
        
        cboWorkStation.BackColor = &H80000005
        
        frm�����㲥����.Enabled = False
    Else
        txt�㲥ʱ�䳤��.BackColor = &H80000005
        txtPlayCount.BackColor = &H80000005
        txtSpeed.BackColor = &H80000005
        
        cboWorkStation.BackColor = Me.BackColor
        
        frm�����㲥����.Enabled = True
    End If
    
    
    If lng�㲥���� <= 10 And lng�㲥���� >= -10 Then
        txtSpeed.Text = lng�㲥����
    Else
        txtSpeed.Text = 0
    End If
    
    'chkUseSound.Value = GetSetting("ZLSOFT", strReg, "������������", 1)
    chkUseSound.Value = zlDatabase.GetPara("������������", glngSys, glngModul, "1")
    
    'chkUseDisplay.Value = GetSetting("ZLSOFT", strReg, "��ʾ�ŶӶ���", 1)
    chkUseDisplay.Value = zlDatabase.GetPara("��ʾ�ŶӶ���", glngSys, glngModul, "1")
    If chkUseDisplay.Value = 0 Then
        cbo��ʾӲ�����.BackColor = frm��ʾ�豸����.BackColor
    End If
    
    '��д��ʾ�豸���
    'lngLEDModal = GetSetting("ZLSOFT", strReg, "��ʾ�豸���", 101)
    lngLEDModal = zlDatabase.GetPara("��ʾ�豸���", glngSys, glngModul, "101")
    
    cbo��ʾӲ�����.Clear
    
    strSql = "Select ��������,������,Nvl(����,0) AS ����,˵�� From �Ŷ�LED��ʾ����  "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "��ȡ��LED��ʾ�ӿڵ�ע����Ϣ")
    
    While rsTemp.EOF = False
        cbo��ʾӲ�����.AddItem Nvl(rsTemp!˵��)
        cbo��ʾӲ�����.ItemData(cbo��ʾӲ�����.ListCount - 1) = Nvl(rsTemp!��������, 0)
        If lngLEDModal = Nvl(rsTemp!��������, 0) Then
            cbo��ʾӲ�����.ListIndex = cbo��ʾӲ�����.ListCount - 1
        End If
        rsTemp.MoveNext
    Wend
    
    If cbo��ʾӲ�����.ListCount > 0 And cbo��ʾӲ�����.ListIndex = -1 Then
        cbo��ʾӲ�����.ListIndex = 0
    End If
    
    cbxComeback.ListIndex = zlDatabase.GetPara("���ﲡ���Ƿ�����", glngSys, glngModul, "1", Array(labCallBack, cbxComeback), True)
    
    optGroupType(Val(zlDatabase.GetPara("�Ŷӷ�������", glngSys, glngModul, "0"))).Value = True
    
    chkOrderStyle.Value = zlDatabase.GetPara("ʹ������ԭʼ˳������", glngSys, glngModul, "0")
End Sub

Private Sub Form_Load()
    Call ReadWorkStationInf
    
    Call ReadLocalPara
End Sub


Private Sub optCallWay_Click(Index As Integer)
    cboWorkStation.Enabled = optCallWay(0).Value
    
    If optCallWay(0).Value Then
        frm�����㲥����.Enabled = False
        
        txt�㲥ʱ�䳤��.BackColor = Me.BackColor
        txtPlayCount.BackColor = Me.BackColor
        txtSpeed.BackColor = Me.BackColor
        
        cboWorkStation.BackColor = &H80000005
    Else
        frm�����㲥����.Enabled = True
        
        txt�㲥ʱ�䳤��.BackColor = &H80000005
        txtPlayCount.BackColor = &H80000005
        txtSpeed.BackColor = &H80000005
        
        cboWorkStation.BackColor = Me.BackColor
    End If
End Sub

Private Sub txt�㲥ʱ�䳤��_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub
