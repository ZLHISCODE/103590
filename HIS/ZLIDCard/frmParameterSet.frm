VERSION 5.00
Begin VB.Form frmParameterSet 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "�豸����"
   ClientHeight    =   2055
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5910
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2055
   ScaleWidth      =   5910
   StartUpPosition =   1  '����������
   Begin VB.Frame Frame1 
      Caption         =   "��˼ SS728M01 ��������"
      Height          =   2475
      Left            =   135
      TabIndex        =   9
      Top             =   2385
      Width           =   5460
      Begin VB.Frame fraLink 
         Height          =   1035
         Left            =   1500
         TabIndex        =   20
         Top             =   165
         Width           =   3780
         Begin VB.CommandButton cmdLink 
            Caption         =   "����(&I)"
            CausesValidation=   0   'False
            Height          =   315
            Left            =   2880
            TabIndex        =   26
            Top             =   233
            Width           =   780
         End
         Begin VB.CommandButton cmdClose 
            Cancel          =   -1  'True
            Caption         =   "�Ͽ�(&O)"
            CausesValidation=   0   'False
            Height          =   315
            Left            =   2865
            TabIndex        =   25
            Top             =   608
            Width           =   780
         End
         Begin VB.OptionButton optNet 
            Caption         =   "�����ն�"
            Height          =   240
            Left            =   135
            TabIndex        =   24
            Top             =   660
            Width           =   1035
         End
         Begin VB.OptionButton optLocal 
            Caption         =   "�����ն�"
            Height          =   240
            Left            =   135
            TabIndex        =   23
            Top             =   285
            Value           =   -1  'True
            Width           =   1035
         End
         Begin VB.ComboBox cboPort 
            Height          =   300
            Left            =   1200
            Style           =   2  'Dropdown List
            TabIndex        =   22
            Top             =   240
            Width           =   1530
         End
         Begin VB.TextBox txtServerIP 
            Height          =   300
            Left            =   1200
            MaxLength       =   16
            TabIndex        =   21
            Text            =   "192.168.31.169"
            Top             =   615
            Width           =   1530
         End
         Begin VB.Line Line4 
            X1              =   2805
            X2              =   2805
            Y1              =   105
            Y2              =   1050
         End
         Begin VB.Line Line3 
            BorderColor     =   &H80000005&
            X1              =   2820
            X2              =   2820
            Y1              =   105
            Y2              =   1050
         End
      End
      Begin VB.TextBox txtGatewayIP 
         Height          =   300
         Left            =   1500
         MaxLength       =   16
         TabIndex        =   19
         Top             =   2040
         Width           =   2040
      End
      Begin VB.CommandButton cmdLet 
         Caption         =   "����(&T)"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   4365
         TabIndex        =   18
         Top             =   2025
         Width           =   780
      End
      Begin VB.CommandButton cmdLoad 
         Caption         =   "��ȡ(&L)"
         Height          =   315
         Left            =   3570
         TabIndex        =   17
         Top             =   2025
         Width           =   780
      End
      Begin VB.CommandButton cmdSet 
         Caption         =   "����(&S)"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   4380
         TabIndex        =   14
         Top             =   1620
         Width           =   780
      End
      Begin VB.CommandButton cmdRead 
         Caption         =   "��ȡ(&R)"
         Height          =   315
         Left            =   3585
         TabIndex        =   13
         Top             =   1620
         Width           =   780
      End
      Begin VB.TextBox txtNetIP 
         Height          =   300
         Left            =   1500
         MaxLength       =   16
         TabIndex        =   12
         Top             =   1635
         Width           =   2040
      End
      Begin VB.Label labGateway 
         Caption         =   "��������IP"
         Height          =   180
         Left            =   225
         TabIndex        =   16
         Top             =   2070
         Width           =   1080
      End
      Begin VB.Label labNetIP 
         Caption         =   "��������IP"
         Height          =   180
         Left            =   225
         TabIndex        =   15
         Top             =   1710
         Width           =   1080
      End
      Begin VB.Shape shpStatus 
         BorderColor     =   &H00FFC0FF&
         DrawMode        =   14  'Copy Pen
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   180
         Left            =   1500
         Shape           =   4  'Rounded Rectangle
         Top             =   1335
         Width           =   420
      End
      Begin VB.Label labStatus 
         Caption         =   "�豸����״̬"
         Height          =   180
         Left            =   195
         TabIndex        =   11
         Top             =   1335
         Width           =   1080
      End
      Begin VB.Label labLinkNetType 
         AutoSize        =   -1  'True
         Caption         =   "�豸���ӷ�ʽ"
         Height          =   180
         Left            =   195
         TabIndex        =   10
         Top             =   630
         Width           =   1080
      End
   End
   Begin VB.Frame fraSet 
      Caption         =   "�豸����"
      Height          =   1695
      Left            =   120
      TabIndex        =   8
      Top             =   150
      Width           =   4140
      Begin VB.TextBox txtInterval 
         Alignment       =   1  'Right Justify
         Height          =   300
         Left            =   1440
         MaxLength       =   4
         TabIndex        =   2
         Text            =   "300"
         ToolTipText     =   "��С300����"
         Top             =   1155
         Width           =   525
      End
      Begin VB.CheckBox chkIDCard 
         Caption         =   "�����豸"
         Height          =   240
         Left            =   1440
         TabIndex        =   0
         Top             =   360
         Width           =   1095
      End
      Begin VB.ComboBox cboType 
         Height          =   300
         ItemData        =   "frmParameterSet.frx":0000
         Left            =   1440
         List            =   "frmParameterSet.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   720
         Width           =   2175
      End
      Begin VB.Label lbltitle 
         Caption         =   "����"
         Height          =   225
         Index           =   2
         Left            =   2040
         TabIndex        =   7
         Top             =   1200
         Width           =   495
      End
      Begin VB.Label lbltitle 
         Caption         =   "�Զ�ʶ����"
         Height          =   225
         Index           =   1
         Left            =   240
         TabIndex        =   6
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label lbltitle 
         Caption         =   "�豸����"
         Height          =   225
         Index           =   0
         Left            =   600
         TabIndex        =   5
         Top             =   780
         Width           =   735
      End
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   4440
      TabIndex        =   3
      Top             =   225
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ��(&C)"
      CausesValidation=   0   'False
      Height          =   350
      Left            =   4425
      TabIndex        =   4
      Top             =   720
      Width           =   1100
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   6540
      Y1              =   2100
      Y2              =   2100
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   0
      X2              =   6540
      Y1              =   2115
      Y2              =   2115
   End
End
Attribute VB_Name = "frmParameterSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'Ҫ���������
Private mblnSS728M01_LinkSate           As Boolean
Private mlngicdev                       As Long         'SS728M01�����豸ID

Private Sub cboType_Click()
On Error GoTo ErrH
    If cboType.Text = "��˼-SS728M01" Then
        Me.Height = 5500
    Else
        Me.Height = 2500
    End If
    Exit Sub
ErrH:
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdCancel_Click()
    If mlngicdev > 0 Then
        MsgBox "�豸����˼-SS728M01����������״̬�����ȶϿ����ӣ�", vbExclamation, "ϵͳ��Ϣ��"
        Exit Sub
    End If
    Unload Me
End Sub

Private Sub CmdHelp_Click()
    'Call ShowHelp(App.ProductName, Me.hwnd, Me.Name)
    '47522:ȡ��zl9Comlib������
End Sub

Private Sub cmdClose_Click()
    Dim lngReturn       As Long
On Error GoTo ErrH
    If mlngicdev > 0 Then
        If optLocal.Value Then
            '2.2.3�رձ����ն�
            'long __stdcall ss_dev_close(long icdev);
            lngReturn = ss_dev_close(mlngicdev)
            Call ss_error(lngReturn)
        ElseIf optNet.Value Then
            '2.2.5ע�������ն�
            lngReturn = ss_dev_logout(mlngicdev)
            Call ss_error(lngReturn)
        End If
    End If
    mblnSS728M01_LinkSate = False
    Call SS728M01_LinkSate
    mlngicdev = 0
    Exit Sub
ErrH:
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdLink_Click()
    Dim lngReturn       As Long
    Dim pszVerInfo      As String
    Dim szDevCom        As String
    Dim Amount          As Long
    Dim Msec            As Long
    Dim szDevIP         As String
On Error GoTo ErrH
    '2.2.1��ȡ�ӿڿ���Ϣ
    'long __stdcall ss_lib_version(char* pszVerInfo);
    pszVerInfo = Space$(4)
    lngReturn = ss_lib_version(pszVerInfo)
    If ss_error(lngReturn) Then
        Exit Sub
    End If
    If optLocal.Value Then
        '2.2.2�򿪱����ն�
        'long __stdcall ss_dev_open(char* szDevCom);
        szDevCom = cboPort.Text
        lngReturn = ss_dev_open(szDevCom)
        If lngReturn < 0 Then
            If ss_error(lngReturn) Then
                Exit Sub
            End If
        Else
            mlngicdev = lngReturn
            ''2.2.7������
            Amount = 1
            Msec = 2
            lngReturn = ss_dev_beep(mlngicdev, Amount, Msec)
            If ss_error(lngReturn) Then
                Exit Sub
            End If
        End If
        '����˿ں�
        SaveSetting "ZLSOFT", "����ȫ��\IDCard\SS728M01", "LinkType", "Local"
        SaveSetting "ZLSOFT", "����ȫ��\IDCard\SS728M01", "Port", szDevCom
    ElseIf optNet.Value Then
        '2.2.4ע�������ն�
        szDevIP = txtServerIP.Text
        lngReturn = ss_dev_login(szDevIP)
        If lngReturn < 0 Then
            If ss_error(lngReturn) Then
                Exit Sub
            End If
        Else
            mlngicdev = lngReturn
            '2.2.7������
            Amount = 1
            Msec = 2
            lngReturn = ss_dev_beep(mlngicdev, Amount, Msec)
            If ss_error(lngReturn) Then
                Exit Sub
            End If
        End If
        SaveSetting "ZLSOFT", "����ȫ��\IDCard\SS728M01", "LinkType", "Net"
        SaveSetting "ZLSOFT", "����ȫ��\IDCard\SS728M01", "NetIp", szDevIP
    End If
    mblnSS728M01_LinkSate = True
    Call SS728M01_LinkSate
    Exit Sub
ErrH:
    MsgBox Err.Description, vbCritical, "���֤��������ʾ"
    Err.Clear
End Sub

Private Sub cmdOk_Click()
    Dim intType As Integer
    
    If cboType.ListIndex >= 0 Then intType = cboType.ItemData(cboType.ListIndex)
    If mlngicdev > 0 Then
        MsgBox "�豸����˼-SS728M01����������״̬�����ȶϿ����ӣ�", vbExclamation, "ϵͳ��Ϣ��"
        Exit Sub
    End If
    SaveSetting "ZLSOFT", "����ȫ��\IDCard", "�豸����", intType
    SaveSetting "ZLSOFT", "����ȫ��\IDCard", "����", chkIDCard.Value
    SaveSetting "ZLSOFT", "����ȫ��\IDCard", "�Զ�ʶ����", Val(txtInterval.Text)
    
    Unload Me
End Sub

Private Sub cmdRead_Click()
    Dim pszParam        As String
    Dim lngReturn       As Long
    Dim Amount          As Long
    Dim Msec            As Long
On Error GoTo ErrH
    '2.2.9��ȡ�������
    'long __stdcall ss_dev_getnet(long icdev, unsigned char _Type, char* pszParam);
    pszParam = Space(16)
    lngReturn = ss_dev_getnet(mlngicdev, ����IP, pszParam)
    If ss_error(lngReturn) Then
        Exit Sub
    End If
    txtNetIP.Text = TruncZero(pszParam)
    '2.2.7������
    Amount = 1
    Msec = 2
    lngReturn = ss_dev_beep(mlngicdev, Amount, Msec)
    If ss_error(lngReturn) Then
        Exit Sub
    End If
    Exit Sub
ErrH:
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdLoad_Click()
    Dim pszParam        As String
    Dim lngReturn       As Long
    Dim Amount          As Long
    Dim Msec            As Long
On Error GoTo ErrH
    '2.2.9��ȡ�������
    'long __stdcall ss_dev_getnet(long icdev, unsigned char _Type, char* pszParam);
    pszParam = Space(16)
    lngReturn = ss_dev_getnet(mlngicdev, ����IP, pszParam)
    If ss_error(lngReturn) Then
        Exit Sub
    End If
    txtGatewayIP.Text = TruncZero(pszParam)
    '2.2.7������
    Amount = 1
    Msec = 2
    lngReturn = ss_dev_beep(mlngicdev, Amount, Msec)
    If ss_error(lngReturn) Then
        Exit Sub
    End If
    Exit Sub
ErrH:
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdSet_Click()
    Dim pszParam        As String
    Dim lngReturn       As Long
    Dim Amount          As Long
    Dim Msec            As Long
On Error GoTo ErrH
    '2.2.10 �����������
    'long __stdcall ss_dev_setnet(long icdev, unsigned char _Type, char* pszParam);
    pszParam = txtNetIP.Text
    lngReturn = ss_dev_setnet(mlngicdev, ����IP, pszParam)
    If ss_error(lngReturn) Then
        Exit Sub
    End If
    
    txtNetIP.Text = "���óɹ���"
    '2.2.7������
    Amount = 1
    Msec = 2
    lngReturn = ss_dev_beep(mlngicdev, Amount, Msec)
    If ss_error(lngReturn) Then
        Exit Sub
    End If
    Exit Sub
ErrH:
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdLet_Click()
    Dim pszParam        As String
    Dim lngReturn       As Long
    Dim Amount          As Long
    Dim Msec            As Long
On Error GoTo ErrH
    '2.2.10 �����������
    'long __stdcall ss_dev_setnet(long icdev, unsigned char _Type, char* pszParam);
    pszParam = txtGatewayIP.Text
    lngReturn = ss_dev_setnet(mlngicdev, ����IP, pszParam)
    If ss_error(lngReturn) Then
        Exit Sub
    End If
    txtGatewayIP.Text = "���óɹ���"
    '2.2.7������
    Amount = 1
    Msec = 2
    lngReturn = ss_dev_beep(mlngicdev, Amount, Msec)
    If ss_error(lngReturn) Then
        Exit Sub
    End If
    Exit Sub
ErrH:
    Err.Clear
    Exit Sub
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Then
        KeyAscii = 0
    ElseIf KeyAscii = 13 Then
        KeyAscii = 0
        Call PressKey(vbKeyTab)
    End If
End Sub

Private Sub Form_Load()
    Dim intTmp  As Integer
    Dim i       As Integer
    
    intTmp = Val(GetSetting("ZLSOFT", "����ȫ��\IDCard", "����", 0))
    chkIDCard.Value = IIf(intTmp = 1, 1, 0)
    
    cboType.AddItem "���ڻ���CVR-100"
    cboType.ItemData(cboType.NewIndex) = IDCardType.CVR100
    cboType.AddItem "���ڻ���CVR-100U"
    cboType.ItemData(cboType.NewIndex) = IDCardType.CVR100U
    cboType.AddItem "���ڻ���CVR-100D"
    cboType.ItemData(cboType.NewIndex) = IDCardType.CVR100D
    cboType.AddItem "��˼-V1"
    cboType.ItemData(cboType.NewIndex) = IDCardType.SS_V1
    cboType.AddItem "������KDQ-116D"
    cboType.ItemData(cboType.NewIndex) = IDCardType.XZX_KDQ
    cboType.AddItem "����GTICR100"
    cboType.ItemData(cboType.NewIndex) = IDCardType.GTICR100
    cboType.AddItem "����GTICR100_01"
    cboType.ItemData(cboType.NewIndex) = IDCardType.GTICR100_01
    cboType.AddItem "����HX-FDX9"
    cboType.ItemData(cboType.NewIndex) = IDCardType.HX_FDX9
    cboType.AddItem "����GTICR100_����"
    cboType.ItemData(cboType.NewIndex) = IDCardType.GTICR100_1
    cboType.AddItem "���ڻ���CVR-100U_����"
    cboType.ItemData(cboType.NewIndex) = IDCardType.CVR100U_1
    cboType.AddItem "���ڻ���CVR-100D_����"
    cboType.ItemData(cboType.NewIndex) = IDCardType.CVR100D_1
    cboType.AddItem "������DKQ-116D_�ɶ�"           '�ɶ�����ƽ�����ӿ�
    cboType.ItemData(cboType.NewIndex) = IDCardType.DKQ_116D
    cboType.AddItem "ͨ�ö������֤�Ķ����ӿ�"      'ͨ�ö������֤�Ķ����ӿ�
    cboType.ItemData(cboType.NewIndex) = IDCardType.COMMON
    cboType.AddItem "��˼-SS728M01"
    cboType.ItemData(cboType.NewIndex) = IDCardType.SS728M01_B01C
    
    intTmp = Val(GetSetting("ZLSOFT", "����ȫ��\IDCard", "�豸����", 0))
    Call CboLocate(cboType, intTmp, True)
    
    intTmp = Val(GetSetting("ZLSOFT", "����ȫ��\IDCard", "�Զ�ʶ����", 300))
    txtInterval.Text = IIf(intTmp < 300, 300, intTmp)
    
    With cboPort
        .AddItem "AUTO"
        .AddItem "COM1"
        .AddItem "COM2"
        .AddItem "COM3"
        .AddItem "COM4"
        .AddItem "COM5"
        .AddItem "COM6"
        .AddItem "COM7"
        .AddItem "COM8"
        .AddItem "COM9"
        .AddItem "COM10"
        .AddItem "COM11"
        .AddItem "COM12"
        .AddItem "COM13"
        .AddItem "COM14"
        .AddItem "COM15"
        .AddItem "COM16"
        .AddItem "USB1"
        .AddItem "USB2"
        .AddItem "USB3"
        .AddItem "USB4"
        .AddItem "USB5"
        .AddItem "USB6"
        .AddItem "USB7"
        .AddItem "USB8"
        .AddItem "USB9"
        .AddItem "USB10"
        .AddItem "USB11"
        .AddItem "USB12"
        .AddItem "USB13"
        .AddItem "USB14"
        .AddItem "USB15"
        .AddItem "USB16"
    End With
    For i = 0 To cboPort.ListCount - 1
        cboPort.ListIndex = i
        If cboPort.Text = GetSetting("ZLSOFT", "����ȫ��\IDCard\SS728M01", "Port", "AUTO") Then
            Exit For
        End If
    Next
    If i = cboPort.ListCount Then cboPort.ListIndex = -1
    Call SS728M01_LinkSate
End Sub

Private Sub optLocal_Click()
    Call SS728M01_LinkSate
End Sub

Private Sub optNet_Click()
    Call SS728M01_LinkSate
End Sub

Private Sub txtInterval_KeyPress(KeyAscii As Integer)
    If InStr("1234567890" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0: Exit Sub
End Sub

Private Sub txtInterval_Validate(Cancel As Boolean)
    If txtInterval.Text < 300 Then Cancel = True
End Sub

Private Function CboLocate(ByVal cboObj As Object, ByVal strValue As String, Optional ByVal blnItem As Boolean = False) As Boolean
    'blnItem:True-��ʾ����ItemData��ֵ��λ������;False-��ʾ�����ı������ݶ�λ������
    Dim lngLocate As Long
    CboLocate = False
    For lngLocate = 0 To cboObj.ListCount - 1
        If blnItem Then
            If cboObj.ItemData(lngLocate) = Val(strValue) Then
                cboObj.ListIndex = lngLocate
                CboLocate = True
                Exit For
            End If
        Else
            If Mid(cboObj.List(lngLocate), InStr(1, cboObj.List(lngLocate), "-") + 1) = strValue Then
                cboObj.ListIndex = lngLocate
                CboLocate = True
                Exit For
            End If
        End If
    Next
End Function

Private Sub SS728M01_LinkSate()
On Error GoTo ErrH
    cmdLink.Enabled = Not mblnSS728M01_LinkSate
    cmdClose.Enabled = mblnSS728M01_LinkSate
    optLocal.Enabled = Not mblnSS728M01_LinkSate
    optNet.Enabled = Not mblnSS728M01_LinkSate
    cboPort.Enabled = Not mblnSS728M01_LinkSate And optLocal.Value
    txtServerIP.Enabled = Not mblnSS728M01_LinkSate And optNet.Value
    shpStatus.FillColor = IIf(mblnSS728M01_LinkSate, vbGreen, vbRed)
    txtNetIP.Enabled = mblnSS728M01_LinkSate
    cmdRead.Enabled = mblnSS728M01_LinkSate
    cmdSet.Enabled = mblnSS728M01_LinkSate
    txtGatewayIP.Enabled = mblnSS728M01_LinkSate
    cmdLoad.Enabled = mblnSS728M01_LinkSate
    cmdLet.Enabled = mblnSS728M01_LinkSate
    Exit Sub
ErrH:
    Err.Clear
    Exit Sub
End Sub
