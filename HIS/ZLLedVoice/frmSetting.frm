VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmSetting 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "设备配置"
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
   StartUpPosition =   2  '屏幕中心
   Begin MSComDlg.CommonDialog cdg 
      Left            =   5160
      Top             =   2520
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdTest 
      Caption         =   "测试(&T)"
      Height          =   350
      Left            =   5085
      TabIndex        =   11
      Top             =   1560
      Visible         =   0   'False
      Width           =   1100
   End
   Begin VB.CommandButton CmdHelp 
      Caption         =   "帮助(&H)"
      CausesValidation=   0   'False
      Height          =   350
      Left            =   5085
      TabIndex        =   12
      Top             =   2010
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(&C)"
      CausesValidation=   0   'False
      Height          =   350
      Left            =   5085
      TabIndex        =   10
      Top             =   600
      Width           =   1100
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "确定(&O)"
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
      Begin VB.Frame fraTDKJ_BJ_2008版 
         Caption         =   "提示信息设置"
         Height          =   1830
         Left            =   120
         TabIndex        =   30
         Top             =   1575
         Width           =   4605
         Begin VB.ComboBox cbo收费 
            Height          =   300
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   34
            Top             =   1335
            Width           =   4410
         End
         Begin VB.ComboBox cbo挂号 
            Height          =   300
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   32
            Top             =   660
            Width           =   4410
         End
         Begin VB.Label lbl收费 
            Caption         =   "收费在姓名时提示"
            Height          =   345
            Left            =   150
            TabIndex        =   33
            Top             =   1095
            Width           =   1785
         End
         Begin VB.Label lbl挂号 
            Caption         =   "挂号在姓名时提示"
            Height          =   345
            Left            =   90
            TabIndex        =   31
            Top             =   420
            Width           =   1785
         End
      End
      Begin VB.Frame fraBps 
         Caption         =   "波特率"
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
         Caption         =   "启用设备"
         Height          =   240
         Left            =   3600
         TabIndex        =   17
         Top             =   240
         Width           =   1095
      End
      Begin VB.Frame fraCom 
         Caption         =   "串行口"
         Height          =   735
         Left            =   120
         TabIndex        =   16
         Top             =   720
         Width           =   4605
         Begin VB.OptionButton optCom 
            Caption         =   "串行口1"
            Height          =   180
            Index           =   0
            Left            =   200
            TabIndex        =   1
            Top             =   350
            Value           =   -1  'True
            Width           =   930
         End
         Begin VB.OptionButton optCom 
            Caption         =   "串行口2"
            Height          =   180
            Index           =   1
            Left            =   1300
            TabIndex        =   2
            Top             =   350
            Width           =   930
         End
         Begin VB.OptionButton optCom 
            Caption         =   "串行口3"
            Height          =   180
            Index           =   2
            Left            =   2400
            TabIndex        =   3
            Top             =   350
            Width           =   930
         End
         Begin VB.OptionButton optCom 
            Caption         =   "串行口4"
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
         Caption         =   "其它"
         Height          =   2535
         Left            =   105
         TabIndex        =   18
         Top             =   2370
         Width           =   4605
         Begin VB.CommandButton cmdSwffile 
            Caption         =   "…"
            Enabled         =   0   'False
            Height          =   250
            Left            =   4140
            TabIndex        =   27
            Top             =   1230
            Width           =   250
         End
         Begin VB.CommandButton cmdPicFile 
            Caption         =   "…"
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
            ToolTipText     =   "最多允许15个字符(不区分半全角)"
            Top             =   1980
            Width           =   3045
         End
         Begin VB.CheckBox chkBottom 
            Caption         =   "底行信息"
            Height          =   195
            Left            =   240
            TabIndex        =   22
            Top             =   2040
            Width           =   1020
         End
         Begin VB.CheckBox chkDDisplay 
            Caption         =   "使用双屏显示器"
            Height          =   195
            Left            =   240
            TabIndex        =   21
            Top             =   360
            Width           =   1740
         End
         Begin VB.CheckBox chkNew 
            Caption         =   "使用新型设备"
            Height          =   195
            Left            =   240
            TabIndex        =   20
            Top             =   1680
            Width           =   1380
         End
         Begin VB.CheckBox chk个帐余额 
            Caption         =   "个帐余额提示"
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
            ToolTipText     =   "当暂停或空闲时自动播放"
            Top             =   1200
            Width           =   3045
         End
         Begin VB.Label Label1 
            Caption         =   "Flash文件"
            Height          =   180
            Left            =   420
            TabIndex        =   29
            Top             =   1245
            Width           =   840
         End
         Begin VB.Label Label2 
            Caption         =   "背景图片"
            Height          =   180
            Left            =   540
            TabIndex        =   26
            Top             =   765
            Width           =   720
         End
      End
      Begin VB.Label lbltitle 
         Caption         =   "设备名称"
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
Option Explicit '要求变量声音
Private blnInit As Boolean '是否已经初始化

Private Sub cboType_Click()
  
    '多数共同支持的
    cmdTest.Visible = True
    fraCom.Visible = True
    optCom(3).Visible = True
    chkDDisplay.Visible = True
    fraCharac.Caption = "其它"
    
    '多数共同不支持的
    fraBps.Visible = False
    chkNew.Visible = False
    chk个帐余额.Visible = False
    
    chkBottom.Visible = False
    txtBottom.Visible = False
    txtBottom.MaxLength = 15
    fraTDKJ_BJ_2008版.Visible = False
    '特有的
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
            chkNew.Visible = True: chk个帐余额.Visible = True
        Case Dev_NJF_VH
        Case Dev_TDKJ_BJ
        Case Dev_surpass
        Case Dev_MDT_SD01
        Case Dev_TDKJ_BJ_2008
             fraTDKJ_BJ_2008版.Visible = True
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
    
    '双屏显示相关内容
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
    
    fraTDKJ_BJ_2008版.Top = fraCharac.Top
    If fraTDKJ_BJ_2008版.Visible Then
        fraCharac.Top = fraTDKJ_BJ_2008版.Top + fraTDKJ_BJ_2008版.Height + 100
    End If
    Me.Height = fraCharac.Top + fraCharac.Height + 800  '800标题栏高度
    
    If Not fraBps.Visible And Not fraCom.Visible Then
        fraCharac.Caption = "设置"
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
            If chkBottom.Value = 1 And txtBottom.Text = "" Then txtBottom.Text = "找零请当面点清,谢谢!"
        End If
    End If
End Sub

Private Sub chkNew_Click()
    chk个帐余额.Enabled = chkNew.Value = 1
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
    
    SaveSetting "ZLSOFT", "公共全局", "设备名称", intType
    SaveSetting "ZLSOFT", "公共全局", "使用", chkLED.Value
    
    '问题:22341
    If intType = 9 Then
        Call SaveData_TDKJ_BJ_2008版
    End If
    For i = 0 To optCom.UBound
        If optCom(i).Value Then
            SaveSetting "ZLSOFT", "公共全局", "端口", i + 1
            Exit For
        End If
    Next
    For i = 0 To optSpeed.UBound
        If optSpeed(i).Value Then
            SaveSetting "ZLSOFT", "公共全局", "波特率", optSpeed(i).Caption
            Exit For
        End If
    Next
    
    SaveSetting "ZLSOFT", "公共全局", "双屏显示器", chkDDisplay.Value
    SaveSetting "ZLSOFT", "公共全局", "背景图片", Trim(txtPicFile.Text)
    SaveSetting "ZLSOFT", "公共全局", "SWF文件", Trim(txtSwffile.Text)
    If frmDisplay.mblnLoad Then
        If (chkDDisplay.Value = 1 Or intType = 99) And chkLED.Value = 1 Then
            Call frmDisplay.Check_Update_BkPic
        Else
            Unload frmDisplay
        End If
    End If
    
    SaveSetting "ZLSOFT", "公共全局", "有底行信息", chkBottom.Value
    SaveSetting "ZLSOFT", "公共全局", "底行信息", txtBottom.Text
    
    SaveSetting "ZLSOFT", "公共全局", "新型SHY-II", chkNew.Value
    SaveSetting "ZLSOFT", "公共全局", "个帐余额提示", chk个帐余额.Value
    
    If blnInit Then
        CloseDevice
        CloseService
    End If
    Unload Me
End Sub
Private Sub Load_TDKJ_BJ_2008版_BaseCode()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载基础参数设置
    '入参:
    '出参:
    '返回:
    '问题:22341
    '编制:刘兴洪
    '日期:2009-08-03 15:17:35
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strInFor As String
    
    strInFor = Trim(GetSetting("ZLSOFT", "公共全局", "挂号提示", "请问您的姓名"))
    With cbo挂号
        .Clear
        .AddItem "请问您的姓名"
        If strInFor = "请问您的姓名" Then .ListIndex = .NewIndex
        .AddItem "请问您孩子的姓名"
        If strInFor = "请问您孩子的姓名" Then .ListIndex = .NewIndex
        If .ListIndex < 0 And strInFor <> "" Then
            '肯定是其他
            .AddItem strInFor: .ListIndex = .NewIndex
        End If
        If .ListIndex < 0 Then .ListIndex = 0
    End With
    
    strInFor = GetSetting("ZLSOFT", "公共全局", "收费提示", "请问您的姓名")
    With cbo收费
        .Clear
        .AddItem "请出示您的挂号票"
        If strInFor = "请出示您的挂号票" Then .ListIndex = .NewIndex
        .AddItem "请问您的姓名"
        If strInFor = "请问您的姓名" Then .ListIndex = .NewIndex
        .AddItem "请问您孩子的姓名"
        If strInFor = "请问您孩子的姓名" Then .ListIndex = .NewIndex
        .AddItem "不播报"
        If strInFor = "不播报" Then .ListIndex = .NewIndex
        If .ListIndex < 0 And strInFor <> "" Then
            '肯定是其他
            .AddItem strInFor: .ListIndex = .NewIndex
        End If
        If .ListIndex < 0 Then .ListIndex = 0
    End With
End Sub
Private Sub SaveData_TDKJ_BJ_2008版()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:保存参数设置
    '入参:
    '出参:
    '返回:
    '问题:22341
    '编制:刘兴洪
    '日期:2009-08-03 15:17:35
    '--------------------------------------------------------------------------------------------------------------------------------------------
    
    SaveSetting "ZLSOFT", "公共全局", "挂号提示", cbo挂号.Text
    SaveSetting "ZLSOFT", "公共全局", "收费提示", cbo收费.Text
End Sub


Private Sub InitData()
    '初始化界面
    Dim intType As Integer
    Dim intPort As Integer
    Dim strSpeed As String
    Dim intX As Integer
    Dim i As Integer
    
    blnInit = False
    
    '问题:22341
    Call Load_TDKJ_BJ_2008版_BaseCode
    
    intType = GetSetting("ZLSOFT", "公共全局", "设备名称", "1")
    chkLED.Value = GetSetting("ZLSOFT", "公共全局", "使用", "0")
    intPort = GetSetting("ZLSOFT", "公共全局", "端口", "1")
    strSpeed = GetSetting("ZLSOFT", "公共全局", "波特率", "9600")
    
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
    
    chkBottom.Value = Val(GetSetting("ZLSOFT", "公共全局", "有底行信息", "0"))
    txtBottom.Text = GetSetting("ZLSOFT", "公共全局", "底行信息", "")
    chkBottom_Click
    
    chkDDisplay.Value = Val(GetSetting("ZLSOFT", "公共全局", "双屏显示器", 0))
    txtPicFile.Text = GetSetting("ZLSOFT", "公共全局", "背景图片", "")
    txtSwffile.Text = GetSetting("ZLSOFT", "公共全局", "SWF文件", "")
    
    chkNew.Value = Val(GetSetting("ZLSOFT", "公共全局", "新型SHY-II", 0))
    chk个帐余额.Value = Val(GetSetting("ZLSOFT", "公共全局", "个帐余额提示", 0))
            
    '需要放在最后加载,以便Cbo_Click事件调用
    With cboType
        .Clear
        .AddItem "SYC XII 语音显示器"
        .ItemData(.NewIndex) = 1
        
        .AddItem "LK822 汉字液晶显示终端"
        .ItemData(.NewIndex) = 2
        
        .AddItem "SHY-II 数码语言报价器"
        .ItemData(.NewIndex) = 3
        
        .AddItem "NJF-VH 智能语音显示器"
        .ItemData(.NewIndex) = 4
        
        .AddItem "TDKJ_BJ_I/II 语音报价系统"
        .ItemData(.NewIndex) = 5
        
        .AddItem "麦迪特SD-01语音显示器"
        .ItemData(.NewIndex) = 6
        
        .AddItem "Dev_SURPASS"
        .ItemData(.NewIndex) = 7
        
        .AddItem "FS-YL01型LED点阵+数码显示器"
        .ItemData(.NewIndex) = 8
        
        .AddItem "TDKJ-BJ_2008版 语音报价系统"
        .ItemData(.NewIndex) = 9
        
        '2010-02-24 ZHQ 一汽总医院加入
        .AddItem "TDKJ_BJ_IV 语音报价器"
        .ItemData(.NewIndex) = 10
        
        .AddItem "SYC-Q9语音报价器"
        .ItemData(.NewIndex) = 11
        
        .AddItem "双屏显示器"
        .ItemData(.NewIndex) = 99
        
        zlControl.CboLocate cboType, intType, True
    End With
    
End Sub

Private Sub cmdPicFile_Click()
    With cdg
        .DialogTitle = "请选择背图片文件"
        .CancelError = False
        .FileName = txtPicFile.Text
        .Filter = "JPG文件(*.jpg)|*.jpg|BMP文件(*.bmp)|*.bmp|GIF文件(*.gif)|*.gif"
        .ShowOpen
        If .FileName <> "" Then
            txtPicFile.Text = .FileName
        End If
    End With
End Sub

Private Sub cmdSwfFile_Click()
    With cdg
        .DialogTitle = "请选择Flash文件"
        .CancelError = False
        .FileName = txtSwffile.Text
        .Filter = "SWF文件(*.swf)|*.swf"
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
    
    '双屏显示窗口
    If cboType.ItemData(cboType.ListIndex) = Dev_DDisplay Then
        SaveSetting "ZLSOFT", "公共全局", "双屏显示器", chkDDisplay.Value
        SaveSetting "ZLSOFT", "公共全局", "背景图片", Trim(txtPicFile.Text)
        SaveSetting "ZLSOFT", "公共全局", "SWF文件", Trim(txtSwffile.Text)
        
        SaveSetting "ZLSOFT", "公共全局", "有底行信息", chkBottom.Value
        SaveSetting "ZLSOFT", "公共全局", "底行信息", txtBottom.Text
    
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
                gobjLED.Display "~请您付款:168.88元"
                gobjLED.stdSpeak "168.88_P"
                gobjLED.Display "~收您:200元,找你:31.12元"
                gobjLED.stdSpeak "200_k"
                gobjLED.stdSpeak "31.12_b"
                gobjLED.Display "~找零请当面点清,谢谢!"
                gobjLED.stdSpeak "_C"
                gobjLED.Reset
                Set gobjLED = Nothing
            Else
                MsgBox "不能创建接口,请确认设置驱动程序是否正确安装！", vbInformation, gstrSysName
            End If
        Case Dev_TDKJ_BJ
            Call TDKJ_BJ_FUN(i + 1, "&Sc$")
            Call TDKJ_BJ_FUN(i + 1, "&C21 欢迎到本院就诊$")
            Call TDKJ_BJ_FUN(i + 1, "&C31  祝您早日康复$")
            Call TDKJ_BJ_FUN(i + 1, "W")
            Call TDKJ_BJ_FUN(i + 1, "v")
            Call TDKJ_BJ_FUN(i + 1, "&Sc$")
        Case Dev_surpass
            Call LocStringDisplay(0, 32, "欢迎你到本院就诊" + Chr(0))
        Case Dev_MDT_SD01
            If Not blnInit Then
                InitService
                InitDevice i + 1
                blnInit = True
            End If
            Display_Line "你好,欢迎到本院就诊", 4, 0
            Voices "010208"  '你好,欢迎光临,祝你早日康复
        Case Dev_FS_YL01
            Call Dev_FS_YL01_Voice("病人甲", 0, 2)
            Call Dev_FS_YL01_Voice(168.88, 1, 3)
            Call Dev_FS_YL01_Voice(200, 2, 3)
            Call Dev_FS_YL01_Voice(31.12, 3, 0)
        Case Dev_TDKJ_BJ_2008
            Call TDKJ_BJ_2008(i + 1, "&Sc$")
            Call TDKJ_BJ_2008(i + 1, "&C21 欢迎到本院就诊$")
            Call TDKJ_BJ_2008(i + 1, "&C31  祝您早日康复$")
            Call TDKJ_BJ_2008(i + 1, "W")
            Call TDKJ_BJ_2008(i + 1, "v")
            Call TDKJ_BJ_2008(i + 1, "&Sc$")
        Case Dev_TDKJ_BJ_IV
            Call TDKJ_BJ_IV(i + 1, "&Sc$")
            Call TDKJ_BJ_IV(i + 1, "&C21 欢迎到本院就诊$")
            Call TDKJ_BJ_IV(i + 1, "&C31  祝您早日康复$")
            Call TDKJ_BJ_IV(i + 1, "W")
            Call TDKJ_BJ_IV(i + 1, "v")
            Call TDKJ_BJ_IV(i + 1, "&Sc$")
        Case Dev_SYC_Q9
            Call SYC_Q9(i + 1, "f")
            Call SYC_Q9(i + 1, "欢迎到本院就诊")
            Call SYC_Q9(i + 1, "祝您早日康复")
            Call SYC_Q9(i + 1, "w")
            Call SYC_Q9(i + 1, "v")
    End Select
    Exit Sub
errH:
    MsgBox "接口调用失败:" & vbCrLf & vbCrLf & Err.Description, vbInformation, gstrSysName
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
