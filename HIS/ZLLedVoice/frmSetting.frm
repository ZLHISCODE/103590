VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmSetting 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "�豸����"
   ClientHeight    =   6195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6315
   ControlBox      =   0   'False
   FillStyle       =   0  'Solid
   Icon            =   "frmSetting.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6195
   ScaleWidth      =   6315
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin MSComDlg.CommonDialog cdg 
      Left            =   5160
      Top             =   2520
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdTest 
      Caption         =   "����(&T)"
      Height          =   350
      Left            =   5085
      TabIndex        =   11
      Top             =   1560
      Visible         =   0   'False
      Width           =   1100
   End
   Begin VB.CommandButton CmdHelp 
      Caption         =   "����(&H)"
      CausesValidation=   0   'False
      Height          =   350
      Left            =   5085
      TabIndex        =   12
      Top             =   2010
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ��(&C)"
      CausesValidation=   0   'False
      Height          =   350
      Left            =   5085
      TabIndex        =   10
      Top             =   600
      Width           =   1100
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "ȷ��(&O)"
      Default         =   -1  'True
      Height          =   350
      Left            =   5085
      TabIndex        =   9
      Top             =   135
      Width           =   1100
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   6165
      Left            =   120
      TabIndex        =   13
      Top             =   30
      Width           =   4815
      Begin VB.Frame fraTDKJ_BJ_2008�� 
         Caption         =   "��ʾ��Ϣ����"
         Height          =   1830
         Left            =   120
         TabIndex        =   30
         Top             =   1575
         Width           =   4605
         Begin VB.ComboBox cbo�շ� 
            Height          =   300
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   34
            Top             =   1335
            Width           =   4410
         End
         Begin VB.ComboBox cbo�Һ� 
            Height          =   300
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   32
            Top             =   660
            Width           =   4410
         End
         Begin VB.Label lbl�շ� 
            Caption         =   "�շ�������ʱ��ʾ"
            Height          =   345
            Left            =   150
            TabIndex        =   33
            Top             =   1095
            Width           =   1785
         End
         Begin VB.Label lbl�Һ� 
            Caption         =   "�Һ�������ʱ��ʾ"
            Height          =   345
            Left            =   90
            TabIndex        =   31
            Top             =   420
            Width           =   1785
         End
      End
      Begin VB.Frame fraBps 
         Caption         =   "������"
         Height          =   735
         Left            =   120
         TabIndex        =   15
         Top             =   1545
         Width           =   4605
         Begin VB.OptionButton optSpeed 
            Caption         =   "2400"
            Height          =   180
            Index           =   3
            Left            =   3500
            TabIndex        =   8
            Top             =   350
            Width           =   750
         End
         Begin VB.OptionButton optSpeed 
            Caption         =   "4800"
            Height          =   180
            Index           =   2
            Left            =   2400
            TabIndex        =   7
            Top             =   350
            Width           =   750
         End
         Begin VB.OptionButton optSpeed 
            Caption         =   "9600"
            Height          =   180
            Index           =   1
            Left            =   1300
            TabIndex        =   6
            Top             =   350
            Value           =   -1  'True
            Width           =   750
         End
         Begin VB.OptionButton optSpeed 
            Caption         =   "19200"
            Height          =   180
            Index           =   0
            Left            =   200
            TabIndex        =   5
            Top             =   350
            Width           =   750
         End
      End
      Begin VB.CheckBox chkLED 
         Caption         =   "�����豸"
         Height          =   240
         Left            =   3600
         TabIndex        =   17
         Top             =   240
         Width           =   1095
      End
      Begin VB.Frame fraCom 
         Caption         =   "���п�"
         Height          =   735
         Left            =   120
         TabIndex        =   16
         Top             =   720
         Width           =   4605
         Begin VB.OptionButton optCom 
            Caption         =   "���п�1"
            Height          =   180
            Index           =   0
            Left            =   200
            TabIndex        =   1
            Top             =   350
            Value           =   -1  'True
            Width           =   930
         End
         Begin VB.OptionButton optCom 
            Caption         =   "���п�2"
            Height          =   180
            Index           =   1
            Left            =   1300
            TabIndex        =   2
            Top             =   350
            Width           =   930
         End
         Begin VB.OptionButton optCom 
            Caption         =   "���п�3"
            Height          =   180
            Index           =   2
            Left            =   2400
            TabIndex        =   3
            Top             =   350
            Width           =   930
         End
         Begin VB.OptionButton optCom 
            Caption         =   "���п�4"
            Height          =   180
            Index           =   3
            Left            =   3500
            TabIndex        =   4
            Top             =   350
            Width           =   930
         End
      End
      Begin VB.ComboBox cboType 
         Height          =   300
         ItemData        =   "frmSetting.frx":000C
         Left            =   960
         List            =   "frmSetting.frx":000E
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   210
         Width           =   2415
      End
      Begin VB.Frame fraCharac 
         Caption         =   "����"
         Height          =   2535
         Left            =   105
         TabIndex        =   18
         Top             =   2370
         Width           =   4605
         Begin VB.CommandButton cmdSwffile 
            Caption         =   "��"
            Enabled         =   0   'False
            Height          =   250
            Left            =   4140
            TabIndex        =   27
            Top             =   1230
            Width           =   250
         End
         Begin VB.CommandButton cmdPicFile 
            Caption         =   "��"
            Enabled         =   0   'False
            Height          =   250
            Left            =   4140
            TabIndex        =   24
            Top             =   750
            Width           =   250
         End
         Begin VB.TextBox txtBottom 
            Enabled         =   0   'False
            Height          =   300
            Left            =   1320
            MaxLength       =   15
            TabIndex        =   23
            ToolTipText     =   "�������15���ַ�(�����ְ�ȫ��)"
            Top             =   1980
            Width           =   3045
         End
         Begin VB.CheckBox chkBottom 
            Caption         =   "������Ϣ"
            Height          =   195
            Left            =   240
            TabIndex        =   22
            Top             =   2040
            Width           =   1020
         End
         Begin VB.CheckBox chkDDisplay 
            Caption         =   "ʹ��˫����ʾ��"
            Height          =   195
            Left            =   240
            TabIndex        =   21
            Top             =   360
            Width           =   1740
         End
         Begin VB.CheckBox chkNew 
            Caption         =   "ʹ�������豸"
            Height          =   195
            Left            =   240
            TabIndex        =   20
            Top             =   1680
            Width           =   1380
         End
         Begin VB.CheckBox chk������� 
            Caption         =   "���������ʾ"
            Enabled         =   0   'False
            Height          =   195
            Left            =   2400
            TabIndex        =   19
            Top             =   1680
            Width           =   1380
         End
         Begin VB.TextBox txtPicFile 
            Enabled         =   0   'False
            Height          =   300
            Left            =   1365
            MaxLength       =   250
            TabIndex        =   25
            Top             =   720
            Width           =   3045
         End
         Begin VB.TextBox txtSwffile 
            Enabled         =   0   'False
            Height          =   300
            Left            =   1345
            MaxLength       =   250
            TabIndex        =   28
            ToolTipText     =   "����ͣ�����ʱ�Զ�����"
            Top             =   1200
            Width           =   3045
         End
         Begin VB.Label Label1 
            Caption         =   "Flash�ļ�"
            Height          =   180
            Left            =   420
            TabIndex        =   29
            Top             =   1245
            Width           =   840
         End
         Begin VB.Label Label2 
            Caption         =   "����ͼƬ"
            Height          =   180
            Left            =   540
            TabIndex        =   26
            Top             =   765
            Width           =   720
         End
      End
      Begin VB.Label lbltitle 
         Caption         =   "�豸����"
         Height          =   225
         Index           =   0
         Left            =   120
         TabIndex        =   14
         Top             =   270
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmSetting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'Ҫ���������
Private blnInit As Boolean '�Ƿ��Ѿ���ʼ��

Private Sub cboType_Click()
  
    '������֧ͬ�ֵ�
    cmdTest.Visible = True
    fraCom.Visible = True
    optCom(3).Visible = True
    chkDDisplay.Visible = True
    fraCharac.Caption = "����"
    
    '������ͬ��֧�ֵ�
    fraBps.Visible = False
    chkNew.Visible = False
    chk�������.Visible = False
    
    chkBottom.Visible = False
    txtBottom.Visible = False
    txtBottom.MaxLength = 15
    fraTDKJ_BJ_2008��.Visible = False
    '���е�
    Select Case cboType.ItemData(cboType.ListIndex)
        Case Dev_SYC_XII
            cmdTest.Visible = False
        Case Dev_LK822
            cmdTest.Visible = False
            fraBps.Visible = True
            chkBottom.Visible = True: txtBottom.Visible = True
        Case Dev_SHY_II
            cmdTest.Visible = False
            optCom(3).Visible = False
            chkNew.Visible = True: chk�������.Visible = True
        Case Dev_NJF_VH
        Case Dev_TDKJ_BJ
        Case Dev_surpass
        Case Dev_MDT_SD01
        Case Dev_TDKJ_BJ_2008
             fraTDKJ_BJ_2008��.Visible = True
        Case Dev_TDKJ_BJ_IV
        Case Dev_DDisplay
            fraCom.Visible = False
            chkDDisplay.Visible = False
            txtPicFile.Enabled = True
            cmdPicFile.Enabled = True
            txtSwffile.Enabled = True
            cmdSwffile.Enabled = True
            txtBottom.MaxLength = 250
            chkBottom.Visible = True: txtBottom.Visible = True
        Case Dev_SYC_Q9
            
    End Select
    
    '˫����ʾ�������
    If cboType.ItemData(cboType.ListIndex) <> Dev_DDisplay Then Call chkDDisplay_Click
    Call chkBottom_Click
    
    If Visible Then Call AdjustPlace
End Sub

Private Sub AdjustPlace()
    
    fraBps.Top = fraCom.Top + IIf(fraCom.Visible, fraCom.Height + 100, 0)
    
    If fraBps.Visible Then
        fraCharac.Top = fraBps.Top + fraBps.Height + 100
    Else
        fraCharac.Top = fraCom.Top + IIf(fraCom.Visible, fraCom.Height + 100, 0)
    End If
    
    fraTDKJ_BJ_2008��.Top = fraCharac.Top
    If fraTDKJ_BJ_2008��.Visible Then
        fraCharac.Top = fraTDKJ_BJ_2008��.Top + fraTDKJ_BJ_2008��.Height + 100
    End If
    Me.Height = fraCharac.Top + fraCharac.Height + 800  '800�������߶�
    
    If Not fraBps.Visible And Not fraCom.Visible Then
        fraCharac.Caption = "����"
    End If
End Sub

Private Sub chkBottom_Click()
    txtBottom.Enabled = chkBottom.Value = 1
End Sub

Private Sub chkDDisplay_Click()
    txtPicFile.Enabled = (chkDDisplay.Value = 1)
    cmdPicFile.Enabled = (chkDDisplay.Value = 1)
    txtSwffile.Enabled = (chkDDisplay.Value = 1)
    cmdSwffile.Enabled = (chkDDisplay.Value = 1)
    
    If cboType.ListIndex <> -1 Then
        If cboType.ItemData(cboType.ListIndex) <> Dev_LK822 Then
            chkBottom.Visible = (chkDDisplay.Value = 1)
            txtBottom.Visible = (chkDDisplay.Value = 1)
            If chkBottom.Value = 1 And txtBottom.Text = "" Then txtBottom.Text = "�����뵱�����,лл!"
        End If
    End If
End Sub

Private Sub chkNew_Click()
    chk�������.Enabled = chkNew.Value = 1
End Sub

Private Sub cmdCancel_Click()
    If blnInit Then
        CloseDevice
        CloseService
    End If
    Unload Me
End Sub

Private Sub CmdHelp_Click()
    Call ShowHelp(App.ProductName, Me.hWnd, Me.Name)
End Sub

Private Sub cmdOk_Click()
    Dim intType As Integer
    Dim i As Integer
    
    intType = cboType.ItemData(cboType.ListIndex)
    
    SaveSetting "ZLSOFT", "����ȫ��", "�豸����", intType
    SaveSetting "ZLSOFT", "����ȫ��", "ʹ��", chkLED.Value
    
    '����:22341
    If intType = 9 Then
        Call SaveData_TDKJ_BJ_2008��
    End If
    For i = 0 To optCom.UBound
        If optCom(i).Value Then
            SaveSetting "ZLSOFT", "����ȫ��", "�˿�", i + 1
            Exit For
        End If
    Next
    For i = 0 To optSpeed.UBound
        If optSpeed(i).Value Then
            SaveSetting "ZLSOFT", "����ȫ��", "������", optSpeed(i).Caption
            Exit For
        End If
    Next
    
    SaveSetting "ZLSOFT", "����ȫ��", "˫����ʾ��", chkDDisplay.Value
    SaveSetting "ZLSOFT", "����ȫ��", "����ͼƬ", Trim(txtPicFile.Text)
    SaveSetting "ZLSOFT", "����ȫ��", "SWF�ļ�", Trim(txtSwffile.Text)
    If frmDisplay.mblnLoad Then
        If (chkDDisplay.Value = 1 Or intType = 99) And chkLED.Value = 1 Then
            Call frmDisplay.Check_Update_BkPic
        Else
            Unload frmDisplay
        End If
    End If
    
    SaveSetting "ZLSOFT", "����ȫ��", "�е�����Ϣ", chkBottom.Value
    SaveSetting "ZLSOFT", "����ȫ��", "������Ϣ", txtBottom.Text
    
    SaveSetting "ZLSOFT", "����ȫ��", "����SHY-II", chkNew.Value
    SaveSetting "ZLSOFT", "����ȫ��", "���������ʾ", chk�������.Value
    
    If blnInit Then
        CloseDevice
        CloseService
    End If
    Unload Me
End Sub
Private Sub Load_TDKJ_BJ_2008��_BaseCode()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ػ�����������
    '���:
    '����:
    '����:
    '����:22341
    '����:���˺�
    '����:2009-08-03 15:17:35
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strInFor As String
    
    strInFor = Trim(GetSetting("ZLSOFT", "����ȫ��", "�Һ���ʾ", "������������"))
    With cbo�Һ�
        .Clear
        .AddItem "������������"
        If strInFor = "������������" Then .ListIndex = .NewIndex
        .AddItem "���������ӵ�����"
        If strInFor = "���������ӵ�����" Then .ListIndex = .NewIndex
        If .ListIndex < 0 And strInFor <> "" Then
            '�϶�������
            .AddItem strInFor: .ListIndex = .NewIndex
        End If
        If .ListIndex < 0 Then .ListIndex = 0
    End With
    
    strInFor = GetSetting("ZLSOFT", "����ȫ��", "�շ���ʾ", "������������")
    With cbo�շ�
        .Clear
        .AddItem "���ʾ���ĹҺ�Ʊ"
        If strInFor = "���ʾ���ĹҺ�Ʊ" Then .ListIndex = .NewIndex
        .AddItem "������������"
        If strInFor = "������������" Then .ListIndex = .NewIndex
        .AddItem "���������ӵ�����"
        If strInFor = "���������ӵ�����" Then .ListIndex = .NewIndex
        .AddItem "������"
        If strInFor = "������" Then .ListIndex = .NewIndex
        If .ListIndex < 0 And strInFor <> "" Then
            '�϶�������
            .AddItem strInFor: .ListIndex = .NewIndex
        End If
        If .ListIndex < 0 Then .ListIndex = 0
    End With
End Sub
Private Sub SaveData_TDKJ_BJ_2008��()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�����������
    '���:
    '����:
    '����:
    '����:22341
    '����:���˺�
    '����:2009-08-03 15:17:35
    '--------------------------------------------------------------------------------------------------------------------------------------------
    
    SaveSetting "ZLSOFT", "����ȫ��", "�Һ���ʾ", cbo�Һ�.Text
    SaveSetting "ZLSOFT", "����ȫ��", "�շ���ʾ", cbo�շ�.Text
End Sub


Private Sub InitData()
    '��ʼ������
    Dim intType As Integer
    Dim intPort As Integer
    Dim strSpeed As String
    Dim intX As Integer
    Dim i As Integer
    
    blnInit = False
    
    '����:22341
    Call Load_TDKJ_BJ_2008��_BaseCode
    
    intType = GetSetting("ZLSOFT", "����ȫ��", "�豸����", "1")
    chkLED.Value = GetSetting("ZLSOFT", "����ȫ��", "ʹ��", "0")
    intPort = GetSetting("ZLSOFT", "����ȫ��", "�˿�", "1")
    strSpeed = GetSetting("ZLSOFT", "����ȫ��", "������", "9600")
    
    For i = 0 To optCom.UBound
        If i + 1 = intPort Then
            optCom(i).Value = True
        End If
    Next
    For i = 0 To optSpeed.UBound
        If optSpeed(i).Caption = strSpeed Then
            optSpeed(i).Value = True
        End If
    Next
    
    chkBottom.Value = Val(GetSetting("ZLSOFT", "����ȫ��", "�е�����Ϣ", "0"))
    txtBottom.Text = GetSetting("ZLSOFT", "����ȫ��", "������Ϣ", "")
    chkBottom_Click
    
    chkDDisplay.Value = Val(GetSetting("ZLSOFT", "����ȫ��", "˫����ʾ��", 0))
    txtPicFile.Text = GetSetting("ZLSOFT", "����ȫ��", "����ͼƬ", "")
    txtSwffile.Text = GetSetting("ZLSOFT", "����ȫ��", "SWF�ļ�", "")
    
    chkNew.Value = Val(GetSetting("ZLSOFT", "����ȫ��", "����SHY-II", 0))
    chk�������.Value = Val(GetSetting("ZLSOFT", "����ȫ��", "���������ʾ", 0))
            
    '��Ҫ����������,�Ա�Cbo_Click�¼�����
    With cboType
        .Clear
        .AddItem "SYC XII ������ʾ��"
        .ItemData(.NewIndex) = 1
        
        .AddItem "LK822 ����Һ����ʾ�ն�"
        .ItemData(.NewIndex) = 2
        
        .AddItem "SHY-II �������Ա�����"
        .ItemData(.NewIndex) = 3
        
        .AddItem "NJF-VH ����������ʾ��"
        .ItemData(.NewIndex) = 4
        
        .AddItem "TDKJ_BJ_I/II ��������ϵͳ"
        .ItemData(.NewIndex) = 5
        
        .AddItem "�����SD-01������ʾ��"
        .ItemData(.NewIndex) = 6
        
        .AddItem "Dev_SURPASS"
        .ItemData(.NewIndex) = 7
        
        .AddItem "FS-YL01��LED����+������ʾ��"
        .ItemData(.NewIndex) = 8
        
        .AddItem "TDKJ-BJ_2008�� ��������ϵͳ"
        .ItemData(.NewIndex) = 9
        
        '2010-02-24 ZHQ һ����ҽԺ����
        .AddItem "TDKJ_BJ_IV ����������"
        .ItemData(.NewIndex) = 10
        
        .AddItem "SYC-Q9����������"
        .ItemData(.NewIndex) = 11
        
        .AddItem "˫����ʾ��"
        .ItemData(.NewIndex) = 99
        
        zlControl.CboLocate cboType, intType, True
    End With
    
End Sub

Private Sub cmdPicFile_Click()
    With cdg
        .DialogTitle = "��ѡ��ͼƬ�ļ�"
        .CancelError = False
        .FileName = txtPicFile.Text
        .Filter = "JPG�ļ�(*.jpg)|*.jpg|BMP�ļ�(*.bmp)|*.bmp|GIF�ļ�(*.gif)|*.gif"
        .ShowOpen
        If .FileName <> "" Then
            txtPicFile.Text = .FileName
        End If
    End With
End Sub

Private Sub cmdSwfFile_Click()
    With cdg
        .DialogTitle = "��ѡ��Flash�ļ�"
        .CancelError = False
        .FileName = txtSwffile.Text
        .Filter = "SWF�ļ�(*.swf)|*.swf"
        .ShowOpen
        If .FileName <> "" Then
            txtSwffile.Text = .FileName
        End If
    End With
End Sub

Private Sub Form_Activate()
    Call AdjustPlace
End Sub

Private Sub Form_Load()
    InitData
End Sub

Private Sub cmdTest_Click()
    Dim i As Integer
    
    On Error GoTo errH
    
    '˫����ʾ����
    If cboType.ItemData(cboType.ListIndex) = Dev_DDisplay Then
        SaveSetting "ZLSOFT", "����ȫ��", "˫����ʾ��", chkDDisplay.Value
        SaveSetting "ZLSOFT", "����ȫ��", "����ͼƬ", Trim(txtPicFile.Text)
        SaveSetting "ZLSOFT", "����ȫ��", "SWF�ļ�", Trim(txtSwffile.Text)
        
        SaveSetting "ZLSOFT", "����ȫ��", "�е�����Ϣ", chkBottom.Value
        SaveSetting "ZLSOFT", "����ȫ��", "������Ϣ", txtBottom.Text
    
        frmDisplay.mblnTest = True
        frmDisplay.Show 1, Me
        Exit Sub
    End If
    
    For i = 0 To optCom.UBound
        If optCom(i).Value Then Exit For
    Next
    
    Select Case cboType.ItemData(cboType.ListIndex)
        Case Dev_NJF_VH
            On Error Resume Next
            Set gobjLED = CreateObject("CTSVR.Bjq")
            On Error GoTo errH
            
            If Not gobjLED Is Nothing Then
                gobjLED.Comport = i + 1
                gobjLED.DispMode = 0
                gobjLED.Display "~��������:168.88Ԫ"
                gobjLED.stdSpeak "168.88_P"
                gobjLED.Display "~����:200Ԫ,����:31.12Ԫ"
                gobjLED.stdSpeak "200_k"
                gobjLED.stdSpeak "31.12_b"
                gobjLED.Display "~�����뵱�����,лл!"
                gobjLED.stdSpeak "_C"
                gobjLED.Reset
                Set gobjLED = Nothing
            Else
                MsgBox "���ܴ����ӿ�,��ȷ���������������Ƿ���ȷ��װ��", vbInformation, gstrSysName
            End If
        Case Dev_TDKJ_BJ
            Call TDKJ_BJ_FUN(i + 1, "&Sc$")
            Call TDKJ_BJ_FUN(i + 1, "&C21 ��ӭ����Ժ����$")
            Call TDKJ_BJ_FUN(i + 1, "&C31  ף�����տ���$")
            Call TDKJ_BJ_FUN(i + 1, "W")
            Call TDKJ_BJ_FUN(i + 1, "v")
            Call TDKJ_BJ_FUN(i + 1, "&Sc$")
        Case Dev_surpass
            Call LocStringDisplay(0, 32, "��ӭ�㵽��Ժ����" + Chr(0))
        Case Dev_MDT_SD01
            If Not blnInit Then
                InitService
                InitDevice i + 1
                blnInit = True
            End If
            Display_Line "���,��ӭ����Ժ����", 4, 0
            Voices "010208"  '���,��ӭ����,ף�����տ���
        Case Dev_FS_YL01
            Call Dev_FS_YL01_Voice("���˼�", 0, 2)
            Call Dev_FS_YL01_Voice(168.88, 1, 3)
            Call Dev_FS_YL01_Voice(200, 2, 3)
            Call Dev_FS_YL01_Voice(31.12, 3, 0)
        Case Dev_TDKJ_BJ_2008
            Call TDKJ_BJ_2008(i + 1, "&Sc$")
            Call TDKJ_BJ_2008(i + 1, "&C21 ��ӭ����Ժ����$")
            Call TDKJ_BJ_2008(i + 1, "&C31  ף�����տ���$")
            Call TDKJ_BJ_2008(i + 1, "W")
            Call TDKJ_BJ_2008(i + 1, "v")
            Call TDKJ_BJ_2008(i + 1, "&Sc$")
        Case Dev_TDKJ_BJ_IV
            Call TDKJ_BJ_IV(i + 1, "&Sc$")
            Call TDKJ_BJ_IV(i + 1, "&C21 ��ӭ����Ժ����$")
            Call TDKJ_BJ_IV(i + 1, "&C31  ף�����տ���$")
            Call TDKJ_BJ_IV(i + 1, "W")
            Call TDKJ_BJ_IV(i + 1, "v")
            Call TDKJ_BJ_IV(i + 1, "&Sc$")
        Case Dev_SYC_Q9
            Call SYC_Q9(i + 1, "f")
            Call SYC_Q9(i + 1, "��ӭ����Ժ����")
            Call SYC_Q9(i + 1, "ף�����տ���")
            Call SYC_Q9(i + 1, "w")
            Call SYC_Q9(i + 1, "v")
    End Select
    Exit Sub
errH:
    MsgBox "�ӿڵ���ʧ��:" & vbCrLf & vbCrLf & Err.Description, vbInformation, gstrSysName
End Sub


Private Sub txtPicFile_GotFocus()
    If Len(txtPicFile.Text) > 0 Then
        txtPicFile.SelStart = 1
        txtPicFile.SelLength = Len(txtPicFile.Text)
    End If
End Sub

Private Sub txtSwffile_GotFocus()
    If Len(txtSwffile.Text) > 0 Then
        txtSwffile.SelStart = 1
        txtSwffile.SelLength = Len(txtSwffile.Text)
    End If
End Sub
