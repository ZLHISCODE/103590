VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "排队显示控制"
   ClientHeight    =   6405
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8520
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6405
   ScaleWidth      =   8520
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame Frame2 
      Caption         =   "窗口显示数量"
      Height          =   4020
      Left            =   3720
      TabIndex        =   20
      Top             =   180
      Width           =   4695
      Begin VB.TextBox txtShowNum 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   120
         TabIndex        =   27
         Text            =   "1"
         Top             =   360
         Width           =   495
      End
      Begin VB.CommandButton cmdShowNum 
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   0
         Left            =   720
         TabIndex        =   26
         Top             =   360
         Width           =   360
      End
      Begin VB.CommandButton cmdShowNum 
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         Left            =   1080
         TabIndex        =   25
         Top             =   360
         Width           =   360
      End
      Begin VB.CommandButton cmdShowNum 
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   2
         Left            =   1440
         TabIndex        =   24
         Top             =   360
         Width           =   360
      End
      Begin VB.CommandButton cmdShowNum 
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   3
         Left            =   1800
         TabIndex        =   23
         Top             =   360
         Width           =   360
      End
      Begin VB.CommandButton cmdShowNum 
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   4
         Left            =   2160
         TabIndex        =   22
         Top             =   360
         Width           =   360
      End
      Begin VB.CommandButton cmdShowNum 
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   5
         Left            =   2520
         TabIndex        =   21
         Top             =   360
         Width           =   360
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfShowStyle 
         Height          =   2970
         Left            =   120
         TabIndex        =   28
         Top             =   885
         Width           =   4440
         _cx             =   7832
         _cy             =   5239
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483636
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   2
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   3
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   360
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   ""
         ScrollTrack     =   0   'False
         ScrollBars      =   2
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
         AutoSearch      =   0
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   0
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   0
         ShowComboButton =   1
         WordWrap        =   0   'False
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         DataMode        =   0
         VirtualData     =   -1  'True
         DataMember      =   ""
         ComboSearch     =   3
         AutoSizeMouse   =   -1  'True
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
         Begin VB.CommandButton cmdConfigWindow 
            Caption         =   "配置"
            Height          =   375
            Left            =   960
            TabIndex        =   29
            Top             =   1080
            Visible         =   0   'False
            Width           =   975
         End
      End
   End
   Begin VB.Frame fraBusinessType 
      Caption         =   "所属业务"
      Height          =   840
      Left            =   90
      TabIndex        =   18
      Top             =   180
      Width           =   3495
      Begin VB.ComboBox cboBusinessType 
         Height          =   300
         ItemData        =   "frmMain.frx":6852
         Left            =   150
         List            =   "frmMain.frx":6854
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   360
         Width           =   3255
      End
   End
   Begin VB.Frame Frame3 
      Height          =   25
      Left            =   -120
      TabIndex        =   17
      Top             =   5740
      Width           =   9375
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "关闭(&C)"
      Height          =   350
      Left            =   7440
      TabIndex        =   16
      Top             =   5920
      Width           =   975
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "应用(&S)"
      Default         =   -1  'True
      Height          =   350
      Left            =   5280
      TabIndex        =   15
      Top             =   5920
      Width           =   975
   End
   Begin VB.Frame fraHospitalInfo 
      Height          =   4500
      Left            =   90
      TabIndex        =   11
      Top             =   1080
      Width           =   3495
      Begin VB.TextBox txtHospitalName 
         Height          =   300
         Left            =   920
         TabIndex        =   13
         Text            =   "重庆市第一人民医院"
         Top             =   240
         Width           =   2445
      End
      Begin VB.CommandButton cmdSetLogo 
         Caption         =   "医院图标设置(&L)"
         Height          =   350
         Left            =   150
         TabIndex        =   12
         Top             =   3960
         Width           =   3195
      End
      Begin VB.Image imgLOGO 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   3315
         Left            =   150
         Picture         =   "frmMain.frx":6856
         Stretch         =   -1  'True
         Top             =   645
         Width           =   3195
      End
      Begin VB.Label lblHospitalName 
         AutoSize        =   -1  'True
         Caption         =   "医院名称"
         Height          =   180
         Left            =   120
         TabIndex        =   14
         Top             =   280
         Width           =   720
      End
   End
   Begin VB.Frame Frame5 
      Height          =   1320
      Left            =   3720
      TabIndex        =   4
      Top             =   4260
      Width           =   4695
      Begin VB.CheckBox chkUserMsgCenter 
         Caption         =   "启用消息服务中心"
         Height          =   180
         Left            =   2880
         TabIndex        =   30
         Top             =   615
         Width           =   1750
      End
      Begin VB.TextBox txtRefreshInterval 
         Height          =   300
         Left            =   1320
         TabIndex        =   8
         Text            =   "30"
         Top             =   915
         Width           =   555
      End
      Begin VB.CheckBox chkAutoLogin 
         Caption         =   "自动登录"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         ToolTipText     =   "自动登录后，不会打开此配置窗口，直接进入排队样式界面。"
         Top             =   600
         Width           =   1020
      End
      Begin VB.CheckBox chkPowerboot 
         Caption         =   "开机自动启动"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   220
         Width           =   1430
      End
      Begin VB.CheckBox chkUseSound 
         Caption         =   "启用语音呼叫"
         Height          =   255
         Left            =   2880
         TabIndex        =   5
         Top             =   220
         Width           =   1420
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "轮询时间间隔"
         Height          =   180
         Left            =   120
         TabIndex        =   10
         Top             =   975
         Width           =   1080
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "秒"
         Height          =   180
         Left            =   1920
         TabIndex        =   9
         Top             =   975
         Width           =   180
      End
   End
   Begin VB.CommandButton cmdImportCfg 
      Caption         =   "导入配置(&I)"
      Height          =   350
      Left            =   120
      TabIndex        =   3
      Top             =   5920
      Width           =   1215
   End
   Begin VB.CommandButton cmdExportCfg 
      Caption         =   "导出配置(&E)"
      Height          =   350
      Left            =   1440
      TabIndex        =   2
      Top             =   5920
      Width           =   1215
   End
   Begin VB.CommandButton cmdView 
      Caption         =   "预览(&V)"
      Height          =   350
      Left            =   6360
      TabIndex        =   1
      Top             =   5920
      Width           =   975
   End
   Begin VB.CommandButton cmdVoiceCfg 
      Caption         =   "语音设置(&V)"
      Height          =   350
      Left            =   2760
      TabIndex        =   0
      Top             =   5920
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog dlgReg 
      Left            =   2880
      Top             =   5460
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'该窗口为整个排队叫号显示的入口，主要功能为进行必要的初始化配置
'在该窗口中主要包含的设置内容如下：
'
'
'1.窗口数量
'2.所属业务
'3.自动登录
'......
'
'当选择窗口数量后，在窗口样式配置列表中，将自动添加对应数量的窗口数据信息，依次顺序编号，且窗口样式默认为单队列
'
'选择一条窗口记录后，内容配置按钮将更具设置的窗口样式，打开配置窗口，调用ShowConfigWindow方法打开对应的样式配置窗口，并返回配置结果
'
'
'
'
'
'
'
'
'
Private mlngOldShowNum As Long      '上一次配置中的显示数量
Private mstrOldShowStyle As String  '样式改变前的值
Private mstrPic As String           '医院图标的十六进制串

Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long

Private Sub ShowConfigWindow(ByVal lngWindowNo As Long, ByVal lngStyleType As TShowStyle)
'显示LCD或LED的配置窗口
'需要根据不同样式打开对应样式的配置窗口
'当未对该过程处理时，默认打开通用配置窗口
'渠道用户在修改排队显示过程中，如果需要修改配置可独立增加窗口然后在该过程中增加

    Dim objConfig As ISty
    Dim objOldCfg As Object
    
On Error GoTo errHandle
    
    Select Case lngStyleType
        Case TShowStyle.ssSingleMan         '单病人样式配置
            Set objConfig = New frmStyle_SingleMan
            
        Case TShowStyle.ssSingleQueue       '单队列样式配置
            Set objConfig = New frmStyle_SingleQueue
            
        Case TShowStyle.ssMultiQueue        '多队列样式配置
            Set objConfig = New frmStyle_MultiQueue
            
        Case TShowStyle.ssOld
            Set objOldCfg = New frmStyle_CommonCfg
            
            Call objOldCfg.OpenShowConfig(lngWindowNo, TShowStyle.ssOld, Me)
            
            Exit Sub
        'Case ....      '如果其他样式需要单独弹出配置窗口，则可以直接添加新的case处理
        '
        '
        '
    End Select

    Call objConfig.ShowCfg(lngWindowNo, Me)
Exit Sub
errHandle:
    Unload objConfig
    Set objConfig = Nothing
    
    '重新抛出异常
    Call Err.Raise(Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Sub

Public Sub zlShowMe()
    Call Me.Show
End Sub

Private Sub InitLocalPars()
'初始化本地参数配置
    Dim i As Integer
    Dim strBusinessType As String

On Error GoTo ErrorHand
    '加载医院图标
    mstrPic = GetSetting("ZLSOFT", G_STR_REGPATH, "医院LOGO")
    Call LoadPictureInfo(imgLOGO, mstrPic)
    
    chkPowerboot.value = Val(GetSetting("ZLSOFT", G_STR_REGPATH, "开机启动", 0))
    chkAutoLogin.value = Val(GetSetting("ZLSOFT", G_STR_REGPATH, "自动登录", 0))
    mlngOldShowNum = Val(GetSetting("ZLSOFT", G_STR_REGPATH, "窗口数量", 1))
    txtShowNum.Text = mlngOldShowNum
    txtHospitalName.Text = GetSetting("ZLSOFT", G_STR_REGPATH, "医院名称", "重庆市第一人民医院")
    txtRefreshInterval.Text = Val(GetSetting("ZLSOFT", G_STR_REGPATH, "轮询间隔", 30))
    chkUseSound.value = Val(GetSetting("ZLSOFT", G_STR_REGPATH, "启用语音呼叫", 1))
    
    If gstrCompareVersion < "010.034.000" Then
        chkUserMsgCenter.value = 0
        chkUserMsgCenter.Enabled = False
    Else
        chkUserMsgCenter.value = Val(GetSetting("ZLSOFT", G_STR_REGPATH, "启用消息服务中心", 1))
    End If
    
    strBusinessType = GetSetting("ZLSOFT", G_STR_REGPATH, "所属业务", "0-临床排队业务")
    
    cboBusinessType.Clear
    cboBusinessType.AddItem "0-临床排队业务"
    cboBusinessType.AddItem "1-PACS排队业务"
    cboBusinessType.AddItem "2-体检排队业务"
    'cboBusinessType.AddItem .....
    
    For i = 0 To cboBusinessType.ListCount - 1
        If cboBusinessType.List(i) = strBusinessType Then
            cboBusinessType.ListIndex = i
            Exit For
        End If
    Next
    
    If cboBusinessType.ListIndex < 0 And cboBusinessType.ListCount > 0 Then cboBusinessType.ListIndex = 0
Exit Sub
ErrorHand:
    MsgBox Err.Description, vbExclamation, gstrSysName
    Err.Clear
End Sub

Private Sub InitStyleSetup()
'初始化样式配置列表
    Dim i As Integer

On Error GoTo ErrorHand

    With vsfShowStyle
        .Cols = 3
        .Rows = 1 '清空数据
        .Rows = Val(txtShowNum.Text) + 1
        
        .ColWidth(0) = 450
        .ColWidth(1) = 2500
        
        .TextMatrix(0, 0) = "编号"
        .TextMatrix(0, 1) = "样式"
        .TextMatrix(0, 2) = "内容配置"
        .Cell(flexcpAlignment, 0, 0, 0, 2) = flexAlignCenterCenter
        
        .Editable = flexEDKbdMouse
        
        If gobjFile.FileExists("C:\APPSOFT\Apply\zl9LCDShow.dll") Then
            .ColComboList(1) = "0-单病人样式|1-单队列样式|2-多队列样式|3-老版本样式"
        Else
            .ColComboList(1) = "0-单病人样式|1-单队列样式|2-多队列样式"
        End If

        For i = 1 To .Rows - 1
            .TextMatrix(i, 0) = i
            .TextMatrix(i, 1) = GetSetting("ZLSOFT", G_STR_REGPATH & "\" & i, "显示样式", "1-单队列样式")    '默认为单队列样式
            
            .Cell(flexcpAlignment, i, 0, i, 2) = flexAlignCenterCenter
        Next
        
        '最后一列自动填充满列表
        .ExtendLastCol = True
        
        If .Rows > 1 Then .RowSel = 1
    End With
Exit Sub
ErrorHand:
    MsgBox Err.Description, vbExclamation, gstrSysName
    Err.Clear
End Sub

Private Sub cboBusinessType_Click()
On Error GoTo ErrorHand
    glngBusinessType = Split(cboBusinessType.Text, "-")(0)
    
    Call frmTrayIcon.setMsgBusinessType(glngBusinessType)
Exit Sub
ErrorHand:
    MsgBox Err.Description, vbExclamation, gstrSysName
    Err.Clear
End Sub

Private Sub chkPowerboot_Click()
    Dim objPowerboot As Object
On Error GoTo ErrorHand
    Set objPowerboot = CreateObject("wscript.shell")
    
    If chkPowerboot.value = 1 Then  '开机启动
        objPowerboot.regwrite "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Run\zlQueueShow", App.Path & "\zlQueueShow.exe"
    Else                            '取消开机启动
        objPowerboot.regdelete "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Run\zlQueueShow"
    End If
Exit Sub
ErrorHand:
    MsgBox Err.Description, vbExclamation, gstrSysName
    Err.Clear
End Sub

Private Sub cmdVoiceCfg_Click()
On Error GoTo ErrorHand
    Call frmVoiceSetup.ShowMe(Me)
Exit Sub
ErrorHand:
    Debug.Print Err.Description
    Err.Clear
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
On Error GoTo ErrorHand
    If KeyAscii = vbKeyEscape Then Unload Me
Exit Sub
ErrorHand:
    Debug.Print Err.Description
    Err.Clear
End Sub

Private Sub cmdView_Click()
On Error GoTo ErrorHand
    '先保存配置
    If Not SaveStyleSetup Then Exit Sub
    
    '关闭样式窗口
    Call CloseStyleWindow
    
    '打开样式窗口
    Call OpenStyleWindow
Exit Sub
ErrorHand:
    MsgBox Err.Description, vbExclamation, gstrSysName
    Err.Clear
End Sub

Private Sub cmdCancel_Click()
    Me.Hide
End Sub

Private Sub cmdConfigWindow_Click()
    Dim lngStyleType As Long

On Error GoTo ErrorHand
    
    Select Case Split(vsfShowStyle.TextMatrix(vsfShowStyle.RowSel, 1), "-")(1)
        Case "单病人样式"
            lngStyleType = TShowStyle.ssSingleMan
        Case "单队列样式"
            lngStyleType = TShowStyle.ssSingleQueue
        Case "多队列样式"
            lngStyleType = TShowStyle.ssMultiQueue
        Case "老版本样式"
            lngStyleType = TShowStyle.ssOld
    End Select
    
    Call ShowConfigWindow(vsfShowStyle.TextMatrix(vsfShowStyle.RowSel, 0), lngStyleType)
Exit Sub
ErrorHand:
    MsgBox Err.Description, vbExclamation, gstrSysName
    Err.Clear
End Sub

Private Sub cmdExportCfg_Click()
'导出配置信息
On Error GoTo ErrorHand
    dlgReg.Filter = "注册文件(*.reg)|*.reg|文本文件(*.txt)|*.txt"
    dlgReg.ShowSave
    
    If dlgReg.FileName = "" Then Exit Sub
    
    Shell "regedit -e """ & dlgReg.FileName & """ ""HKEY_CURRENT_USER\Software\VB and VBA Program Settings\ZLSOFT\公共模块\zl9QueueShow""", vbHide
Exit Sub
ErrorHand:
    MsgBox Err.Description, vbExclamation, gstrSysName
    Err.Clear
End Sub

Private Sub cmdImportCfg_Click()
'导入配置信息
On Error GoTo ErrorHand
    dlgReg.Filter = "注册文件(*.reg)|*.reg|文本文件(*.txt)|*.txt"
    dlgReg.ShowOpen
     
    If dlgReg.FileName = "" Then Exit Sub
     
    Shell "regedit /s """ & dlgReg.FileName & ""
    '根据导入的注册信息重新配置排队显示控制窗口
    Call InitLocalPars
    Call InitStyleSetup
Exit Sub
ErrorHand:
    MsgBox Err.Description, vbExclamation, gstrSysName
    Err.Clear
End Sub

Private Sub cmdSave_Click()
'保存配置
On Error GoTo ErrorHand
    Call SaveStyleSetup
Exit Sub
ErrorHand:
    MsgBox Err.Description, vbExclamation, gstrSysName
    Err.Clear
End Sub

Private Function SaveStyleSetup() As Boolean
'保存样式设置
    Dim i As Integer
    
    SaveStyleSetup = False
On Error GoTo ErrorHand1
    SaveSetting "ZLSOFT", G_STR_REGPATH, "窗口数量", Nvl(txtShowNum.Text, 1)
    SaveSetting "ZLSOFT", G_STR_REGPATH, "开机启动", chkPowerboot.value
    SaveSetting "ZLSOFT", G_STR_REGPATH, "自动登录", chkAutoLogin.value
    SaveSetting "ZLSOFT", G_STR_REGPATH, "所属业务", cboBusinessType.Text
    SaveSetting "ZLSOFT", G_STR_REGPATH, "医院名称", txtHospitalName.Text
    SaveSetting "ZLSOFT", G_STR_REGPATH, "轮询间隔", Nvl(txtRefreshInterval.Text, 30)
    SaveSetting "ZLSOFT", G_STR_REGPATH, "启用语音呼叫", chkUseSound.value
    SaveSetting "ZLSOFT", G_STR_REGPATH, "启用消息服务中心", chkUserMsgCenter.value
    
    For i = 1 To vsfShowStyle.Rows - 1
        SaveSetting "ZLSOFT", G_STR_REGPATH & "\" & i, "显示样式", vsfShowStyle.TextMatrix(i, 1)
    Next
    
    '当新设置的窗口显示数量小于上一次设置的数量时，删除注册表中多余的信息
    If Val(txtShowNum.Text) < mlngOldShowNum Then
        For i = Val(txtShowNum.Text) + 1 To mlngOldShowNum
            RegDeleteKey HKEY_CURRENT_USER, "Software\VB and VBA Program Settings\ZLSOFT\公共模块\zl9QueueShow\" & i
        Next
    End If
        
On Error GoTo ErrorHand2
    SaveSetting "ZLSOFT", G_STR_REGPATH, "医院LOGO", mstrPic
    
    SaveStyleSetup = True
Exit Function
ErrorHand1:
    MsgBox Err.Description, vbExclamation, gstrSysName
    Err.Clear
ErrorHand2:
    MsgBox "您选择的图片过大，请重新选择！", vbExclamation, gstrSysName
    Err.Clear
End Function

Private Sub cmdShowNum_Click(Index As Integer)
On Error GoTo ErrorHand
    txtShowNum.Text = cmdShowNum(Index).Caption
Exit Sub
ErrorHand:
    MsgBox Err.Description, vbExclamation, gstrSysName
    Err.Clear
End Sub

Private Sub cmdSetLogo_Click()
    Dim strFileName As String
    Dim arrByte() As Byte
    Dim arrPic() As String
    Dim lngCount As Long, lngFileSize As Long
On Error GoTo ErrorHand
    dlgReg.Filter = "(*.jpg)|*.jpg|(*.gif)|*.gif|(*.bmp)|*.bmp|(*.*)|*.*"
    dlgReg.ShowOpen

    strFileName = dlgReg.FileName

    If strFileName = "" Then Exit Sub

    '读取文件长度
    lngFileSize = FileLen(strFileName)

    ReDim arrByte(0 To lngFileSize - 1) '定义数值长度
    ReDim arrPic(0 To lngFileSize - 1) '定义数值长度

    Open strFileName For Binary As #1
    Get #1, , arrByte
    Close #1

    '将字节转换为16进制
    For lngCount = LBound(arrByte) To UBound(arrByte)
        arrPic(lngCount) = Hex(arrByte(lngCount))
        If Len(arrPic(lngCount)) = 1 Then arrPic(lngCount) = "0" & arrPic(lngCount)
    Next
    
    mstrPic = Join(arrPic, "")
    
    imgLOGO.Picture = LoadPicture(strFileName)
Exit Sub
ErrorHand:
    MsgBox Err.Description, vbExclamation, gstrSysName
    Err.Clear
End Sub

Private Sub Form_Load()
On Error GoTo ErrorHand
    Call InitLocalPars
    
    Call InitStyleSetup
Exit Sub
ErrorHand:
    MsgBox Err.Description, vbExclamation, gstrSysName
    Err.Clear
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error GoTo ErrorHand
    Unload frmTrayIcon
Exit Sub
ErrorHand:
    MsgBox Err.Description, vbExclamation, gstrSysName
    Err.Clear
End Sub

Private Sub txtRefreshInterval_KeyPress(KeyAscii As Integer)
On Error GoTo ErrorHand
    If InStr("01234567890" & Chr(8), Chr(KeyAscii)) <= 0 Then
        KeyAscii = 0
    End If
Exit Sub
ErrorHand:
    MsgBox Err.Description, vbExclamation, gstrSysName
    Err.Clear
End Sub

Private Sub txtShowNum_Change()
On Error GoTo ErrorHand
    If Val(txtShowNum.Text) <= 0 Then txtShowNum.Text = 1
    txtShowNum.Text = Val(txtShowNum.Text)
    
    Call InitStyleSetup
Exit Sub
ErrorHand:
    MsgBox Err.Description, vbExclamation, gstrSysName
    Err.Clear
End Sub

Private Sub txtShowNum_KeyPress(KeyAscii As Integer)
On Error GoTo ErrorHand
    If InStr("01234567890" & Chr(8), Chr(KeyAscii)) <= 0 Then
        KeyAscii = 0
    End If
Exit Sub
ErrorHand:
    MsgBox Err.Description, vbExclamation, gstrSysName
    Err.Clear
End Sub

Private Sub vsfShowStyle_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim i As Integer
On Error GoTo ErrorHand
    '只能配置一个老版本LCD显示
    If vsfShowStyle.TextMatrix(Row, 1) = "3-老版本样式" Then
        For i = 1 To vsfShowStyle.Rows - 1
            If vsfShowStyle.TextMatrix(i, 1) = "3-老版本样式" And i <> Row Then
                MsgBox "最多只能配置一个老版本显示窗口！", vbExclamation, gstrSysName
                vsfShowStyle.TextMatrix(Row, 1) = mstrOldShowStyle
                Exit Sub
            End If
        Next
    End If
    
    If mstrOldShowStyle <> vsfShowStyle.TextMatrix(Row, 1) Then
        If MsgBox("改变样式后将删除原来对应的内容配置" & vbCrLf & "是否要继续？", vbYesNo + vbDefaultButton2) = vbNo Then
            vsfShowStyle.TextMatrix(Row, 1) = mstrOldShowStyle
            Exit Sub
        End If
        
        RegDeleteKey HKEY_CURRENT_USER, "Software\VB and VBA Program Settings\ZLSOFT\公共模块\zl9QueueShow\" & Row
    End If
Exit Sub
ErrorHand:
    MsgBox Err.Description, vbExclamation, gstrSysName
    Err.Clear
End Sub

Private Sub vsfShowStyle_AfterScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)
On Error GoTo ErrorHand
    Call ShowButton
Exit Sub
ErrorHand:
    MsgBox Err.Description, vbExclamation, gstrSysName
    Err.Clear
End Sub

Private Sub ShowButton()
'在指定单元格显示配置按钮
    cmdConfigWindow.Visible = False
    
    With vsfShowStyle
        If .RowSel < 0 Then Exit Sub
        
        cmdConfigWindow.Left = .Cell(flexcpLeft, .RowSel, 2)
        cmdConfigWindow.Top = .Cell(flexcpTop, .RowSel, 2)
        cmdConfigWindow.Height = .Cell(flexcpHeight, .RowSel, 2) - 10
        cmdConfigWindow.Width = .Cell(flexcpWidth, .RowSel, 2) - 10
    End With
    
    If cmdConfigWindow.Top < vsfShowStyle.RowHeight(0) Then Exit Sub
    
    cmdConfigWindow.Visible = True
End Sub

Private Sub vsfShowStyle_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
On Error GoTo ErrorHand
    mstrOldShowStyle = vsfShowStyle.TextMatrix(Row, 1)
Exit Sub
ErrorHand:
    MsgBox Err.Description, vbExclamation, gstrSysName
    Err.Clear
End Sub

Private Sub vsfShowStyle_ComboCloseUp(ByVal Row As Long, ByVal Col As Long, FinishEdit As Boolean)
    FinishEdit = True
End Sub

Private Sub vsfShowStyle_DblClick()
On Error GoTo ErrorHand
    If vsfShowStyle.ColSel = 1 Then
        vsfShowStyle.Editable = flexEDKbdMouse
    Else
        vsfShowStyle.Editable = flexEDNone
    End If
Exit Sub
ErrorHand:
    MsgBox Err.Description, vbExclamation, gstrSysName
    Err.Clear
End Sub

Private Sub vsfShowStyle_SelChange()
On Error GoTo ErrorHand
    Call ShowButton
Exit Sub
ErrorHand:
    MsgBox Err.Description, vbExclamation, gstrSysName
    Err.Clear
End Sub
