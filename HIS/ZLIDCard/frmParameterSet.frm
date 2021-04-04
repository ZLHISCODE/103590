VERSION 5.00
Begin VB.Form frmParameterSet 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "设备配置"
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
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame Frame1 
      Caption         =   "神思 SS728M01 参数设置"
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
            Caption         =   "连接(&I)"
            CausesValidation=   0   'False
            Height          =   315
            Left            =   2880
            TabIndex        =   26
            Top             =   233
            Width           =   780
         End
         Begin VB.CommandButton cmdClose 
            Cancel          =   -1  'True
            Caption         =   "断开(&O)"
            CausesValidation=   0   'False
            Height          =   315
            Left            =   2865
            TabIndex        =   25
            Top             =   608
            Width           =   780
         End
         Begin VB.OptionButton optNet 
            Caption         =   "网络终端"
            Height          =   240
            Left            =   135
            TabIndex        =   24
            Top             =   660
            Width           =   1035
         End
         Begin VB.OptionButton optLocal 
            Caption         =   "本地终端"
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
         Caption         =   "设置(&T)"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   4365
         TabIndex        =   18
         Top             =   2025
         Width           =   780
      End
      Begin VB.CommandButton cmdLoad 
         Caption         =   "读取(&L)"
         Height          =   315
         Left            =   3570
         TabIndex        =   17
         Top             =   2025
         Width           =   780
      End
      Begin VB.CommandButton cmdSet 
         Caption         =   "设置(&S)"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   4380
         TabIndex        =   14
         Top             =   1620
         Width           =   780
      End
      Begin VB.CommandButton cmdRead 
         Caption         =   "读取(&R)"
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
         Caption         =   "本机网关IP"
         Height          =   180
         Left            =   225
         TabIndex        =   16
         Top             =   2070
         Width           =   1080
      End
      Begin VB.Label labNetIP 
         Caption         =   "本机网络IP"
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
         Caption         =   "设备连接状态"
         Height          =   180
         Left            =   195
         TabIndex        =   11
         Top             =   1335
         Width           =   1080
      End
      Begin VB.Label labLinkNetType 
         AutoSize        =   -1  'True
         Caption         =   "设备连接方式"
         Height          =   180
         Left            =   195
         TabIndex        =   10
         Top             =   630
         Width           =   1080
      End
   End
   Begin VB.Frame fraSet 
      Caption         =   "设备配置"
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
         ToolTipText     =   "最小300毫秒"
         Top             =   1155
         Width           =   525
      End
      Begin VB.CheckBox chkIDCard 
         Caption         =   "启用设备"
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
         Caption         =   "毫秒"
         Height          =   225
         Index           =   2
         Left            =   2040
         TabIndex        =   7
         Top             =   1200
         Width           =   495
      End
      Begin VB.Label lbltitle 
         Caption         =   "自动识别间隔"
         Height          =   225
         Index           =   1
         Left            =   240
         TabIndex        =   6
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label lbltitle 
         Caption         =   "设备名称"
         Height          =   225
         Index           =   0
         Left            =   600
         TabIndex        =   5
         Top             =   780
         Width           =   735
      End
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   4440
      TabIndex        =   3
      Top             =   225
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(&C)"
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
Option Explicit '要求变量声明
Private mblnSS728M01_LinkSate           As Boolean
Private mlngicdev                       As Long         'SS728M01连接设备ID

Private Sub cboType_Click()
On Error GoTo ErrH
    If cboType.Text = "神思-SS728M01" Then
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
        MsgBox "设备【神思-SS728M01】处于连接状态，请先断开连接！", vbExclamation, "系统消息！"
        Exit Sub
    End If
    Unload Me
End Sub

Private Sub CmdHelp_Click()
    'Call ShowHelp(App.ProductName, Me.hwnd, Me.Name)
    '47522:取消zl9Comlib的引用
End Sub

Private Sub cmdClose_Click()
    Dim lngReturn       As Long
On Error GoTo ErrH
    If mlngicdev > 0 Then
        If optLocal.Value Then
            '2.2.3关闭本地终端
            'long __stdcall ss_dev_close(long icdev);
            lngReturn = ss_dev_close(mlngicdev)
            Call ss_error(lngReturn)
        ElseIf optNet.Value Then
            '2.2.5注销网络终端
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
    '2.2.1获取接口库信息
    'long __stdcall ss_lib_version(char* pszVerInfo);
    pszVerInfo = Space$(4)
    lngReturn = ss_lib_version(pszVerInfo)
    If ss_error(lngReturn) Then
        Exit Sub
    End If
    If optLocal.Value Then
        '2.2.2打开本地终端
        'long __stdcall ss_dev_open(char* szDevCom);
        szDevCom = cboPort.Text
        lngReturn = ss_dev_open(szDevCom)
        If lngReturn < 0 Then
            If ss_error(lngReturn) Then
                Exit Sub
            End If
        Else
            mlngicdev = lngReturn
            ''2.2.7蜂鸣器
            Amount = 1
            Msec = 2
            lngReturn = ss_dev_beep(mlngicdev, Amount, Msec)
            If ss_error(lngReturn) Then
                Exit Sub
            End If
        End If
        '保存端口号
        SaveSetting "ZLSOFT", "公共全局\IDCard\SS728M01", "LinkType", "Local"
        SaveSetting "ZLSOFT", "公共全局\IDCard\SS728M01", "Port", szDevCom
    ElseIf optNet.Value Then
        '2.2.4注册网络终端
        szDevIP = txtServerIP.Text
        lngReturn = ss_dev_login(szDevIP)
        If lngReturn < 0 Then
            If ss_error(lngReturn) Then
                Exit Sub
            End If
        Else
            mlngicdev = lngReturn
            '2.2.7蜂鸣器
            Amount = 1
            Msec = 2
            lngReturn = ss_dev_beep(mlngicdev, Amount, Msec)
            If ss_error(lngReturn) Then
                Exit Sub
            End If
        End If
        SaveSetting "ZLSOFT", "公共全局\IDCard\SS728M01", "LinkType", "Net"
        SaveSetting "ZLSOFT", "公共全局\IDCard\SS728M01", "NetIp", szDevIP
    End If
    mblnSS728M01_LinkSate = True
    Call SS728M01_LinkSate
    Exit Sub
ErrH:
    MsgBox Err.Description, vbCritical, "身份证读卡器提示"
    Err.Clear
End Sub

Private Sub cmdOk_Click()
    Dim intType As Integer
    
    If cboType.ListIndex >= 0 Then intType = cboType.ItemData(cboType.ListIndex)
    If mlngicdev > 0 Then
        MsgBox "设备【神思-SS728M01】处于连接状态，请先断开连接！", vbExclamation, "系统消息！"
        Exit Sub
    End If
    SaveSetting "ZLSOFT", "公共全局\IDCard", "设备种类", intType
    SaveSetting "ZLSOFT", "公共全局\IDCard", "启用", chkIDCard.Value
    SaveSetting "ZLSOFT", "公共全局\IDCard", "自动识别间隔", Val(txtInterval.Text)
    
    Unload Me
End Sub

Private Sub cmdRead_Click()
    Dim pszParam        As String
    Dim lngReturn       As Long
    Dim Amount          As Long
    Dim Msec            As Long
On Error GoTo ErrH
    '2.2.9读取网络参数
    'long __stdcall ss_dev_getnet(long icdev, unsigned char _Type, char* pszParam);
    pszParam = Space(16)
    lngReturn = ss_dev_getnet(mlngicdev, 本机IP, pszParam)
    If ss_error(lngReturn) Then
        Exit Sub
    End If
    txtNetIP.Text = TruncZero(pszParam)
    '2.2.7蜂鸣器
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
    '2.2.9读取网络参数
    'long __stdcall ss_dev_getnet(long icdev, unsigned char _Type, char* pszParam);
    pszParam = Space(16)
    lngReturn = ss_dev_getnet(mlngicdev, 网关IP, pszParam)
    If ss_error(lngReturn) Then
        Exit Sub
    End If
    txtGatewayIP.Text = TruncZero(pszParam)
    '2.2.7蜂鸣器
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
    '2.2.10 设置网络参数
    'long __stdcall ss_dev_setnet(long icdev, unsigned char _Type, char* pszParam);
    pszParam = txtNetIP.Text
    lngReturn = ss_dev_setnet(mlngicdev, 本机IP, pszParam)
    If ss_error(lngReturn) Then
        Exit Sub
    End If
    
    txtNetIP.Text = "设置成功！"
    '2.2.7蜂鸣器
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
    '2.2.10 设置网络参数
    'long __stdcall ss_dev_setnet(long icdev, unsigned char _Type, char* pszParam);
    pszParam = txtGatewayIP.Text
    lngReturn = ss_dev_setnet(mlngicdev, 网关IP, pszParam)
    If ss_error(lngReturn) Then
        Exit Sub
    End If
    txtGatewayIP.Text = "设置成功！"
    '2.2.7蜂鸣器
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
    
    intTmp = Val(GetSetting("ZLSOFT", "公共全局\IDCard", "启用", 0))
    chkIDCard.Value = IIf(intTmp = 1, 1, 0)
    
    cboType.AddItem "深圳华视CVR-100"
    cboType.ItemData(cboType.NewIndex) = IDCardType.CVR100
    cboType.AddItem "深圳华视CVR-100U"
    cboType.ItemData(cboType.NewIndex) = IDCardType.CVR100U
    cboType.AddItem "深圳华视CVR-100D"
    cboType.ItemData(cboType.NewIndex) = IDCardType.CVR100D
    cboType.AddItem "神思-V1"
    cboType.ItemData(cboType.NewIndex) = IDCardType.SS_V1
    cboType.AddItem "新中新KDQ-116D"
    cboType.ItemData(cboType.NewIndex) = IDCardType.XZX_KDQ
    cboType.AddItem "国腾GTICR100"
    cboType.ItemData(cboType.NewIndex) = IDCardType.GTICR100
    cboType.AddItem "国腾GTICR100_01"
    cboType.ItemData(cboType.NewIndex) = IDCardType.GTICR100_01
    cboType.AddItem "华旭HX-FDX9"
    cboType.ItemData(cboType.NewIndex) = IDCardType.HX_FDX9
    cboType.AddItem "国腾GTICR100_升级"
    cboType.ItemData(cboType.NewIndex) = IDCardType.GTICR100_1
    cboType.AddItem "深圳华视CVR-100U_升级"
    cboType.ItemData(cboType.NewIndex) = IDCardType.CVR100U_1
    cboType.AddItem "深圳华视CVR-100D_升级"
    cboType.ItemData(cboType.NewIndex) = IDCardType.CVR100D_1
    cboType.AddItem "新中新DKQ-116D_成都"           '成都宋献平新增接口
    cboType.ItemData(cboType.NewIndex) = IDCardType.DKQ_116D
    cboType.AddItem "通用二代身份证阅读器接口"      '通用二代身份证阅读器接口
    cboType.ItemData(cboType.NewIndex) = IDCardType.COMMON
    cboType.AddItem "神思-SS728M01"
    cboType.ItemData(cboType.NewIndex) = IDCardType.SS728M01_B01C
    
    intTmp = Val(GetSetting("ZLSOFT", "公共全局\IDCard", "设备种类", 0))
    Call CboLocate(cboType, intTmp, True)
    
    intTmp = Val(GetSetting("ZLSOFT", "公共全局\IDCard", "自动识别间隔", 300))
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
        If cboPort.Text = GetSetting("ZLSOFT", "公共全局\IDCard\SS728M01", "Port", "AUTO") Then
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
    'blnItem:True-表示根据ItemData的值定位下拉框;False-表示根据文本的内容定位下拉框
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
