VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmHL7Main 
   Caption         =   "中联HL7服务"
   ClientHeight    =   3330
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   5400
   Icon            =   "frmHL7Main.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3330
   ScaleWidth      =   5400
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer tmFileInput 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   1800
      Top             =   2040
   End
   Begin VB.Timer tmIdle 
      Interval        =   60000
      Left            =   4080
      Top             =   1320
   End
   Begin MSWinsockLib.Winsock wsSend 
      Left            =   120
      Top             =   2520
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Timer tmHISAction 
      Interval        =   10000
      Left            =   3360
      Top             =   1320
   End
   Begin VB.Timer tmLinkTimeOut 
      Index           =   0
      Interval        =   1000
      Left            =   2520
      Top             =   1320
   End
   Begin VB.Timer tmMsgProcess 
      Interval        =   10000
      Left            =   1800
      Top             =   1320
   End
   Begin MSWinsockLib.Winsock wsHL7Server 
      Index           =   0
      Left            =   120
      Top             =   1920
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSComctlLib.Toolbar tbrMain 
      Align           =   1  'Align Top
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5400
      _ExtentX        =   9525
      _ExtentY        =   1296
      ButtonWidth     =   1455
      ButtonHeight    =   1138
      Appearance      =   1
      ImageList       =   "imgGray"
      HotImageList    =   "imgColor"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   6
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "启动服务"
            Key             =   "启动服务"
            ImageKey        =   "启动服务"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "停止服务"
            Key             =   "停止服务"
            ImageKey        =   "停止服务"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "帮助"
            Key             =   "帮助"
            Object.ToolTipText     =   "帮助"
            Object.Tag             =   "帮助"
            ImageKey        =   "帮助"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "退出"
            Key             =   "退出"
            Object.ToolTipText     =   "退出"
            Object.Tag             =   "退出"
            ImageKey        =   "退出"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgColor 
      Left            =   120
      Top             =   1200
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   16
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHL7Main.frx":08CA
            Key             =   "预览"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHL7Main.frx":0AE4
            Key             =   "打印"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHL7Main.frx":0CFE
            Key             =   "帮助"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHL7Main.frx":0F18
            Key             =   "退出"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHL7Main.frx":1132
            Key             =   "记录"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHL7Main.frx":182C
            Key             =   "调整"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHL7Main.frx":1F26
            Key             =   "完成"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHL7Main.frx":2620
            Key             =   "主费"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHL7Main.frx":2D1A
            Key             =   "补费"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHL7Main.frx":3414
            Key             =   "改费"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHL7Main.frx":3B0E
            Key             =   "删费"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHL7Main.frx":4208
            Key             =   "新嘱"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHL7Main.frx":4902
            Key             =   "修改"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHL7Main.frx":4FFC
            Key             =   "停止服务"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHL7Main.frx":56F6
            Key             =   "作废"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHL7Main.frx":5DF0
            Key             =   "启动服务"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgGray 
      Left            =   720
      Top             =   1200
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   16
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHL7Main.frx":64EA
            Key             =   "预览"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHL7Main.frx":6704
            Key             =   "打印"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHL7Main.frx":691E
            Key             =   "帮助"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHL7Main.frx":6B38
            Key             =   "退出"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHL7Main.frx":6D52
            Key             =   "记录"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHL7Main.frx":744C
            Key             =   "调整"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHL7Main.frx":7B46
            Key             =   "完成"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHL7Main.frx":8240
            Key             =   "主费"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHL7Main.frx":893A
            Key             =   "补费"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHL7Main.frx":9034
            Key             =   "改费"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHL7Main.frx":972E
            Key             =   "删费"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHL7Main.frx":9E28
            Key             =   "新嘱"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHL7Main.frx":A522
            Key             =   "修改"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHL7Main.frx":AC1C
            Key             =   "停止服务"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHL7Main.frx":B316
            Key             =   "作废"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHL7Main.frx":BA10
            Key             =   "启动服务"
         EndProperty
      EndProperty
   End
   Begin MSWinsockLib.Winsock wsHL7Link 
      Index           =   0
      Left            =   720
      Top             =   1920
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Menu mnuFile 
      Caption         =   "文件(&F)"
      Begin VB.Menu mmuParaSetup 
         Caption         =   "参数设置"
      End
      Begin VB.Menu mnuFile_3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileQuit 
         Caption         =   "退出(&X)"
      End
   End
   Begin VB.Menu mnuService 
      Caption         =   "服务"
      Begin VB.Menu munStartService 
         Caption         =   "启动服务"
      End
      Begin VB.Menu munStopService 
         Caption         =   "停止服务"
      End
      Begin VB.Menu mnuService_1 
         Caption         =   "-"
      End
      Begin VB.Menu mmuShowService 
         Caption         =   "显示当前服务"
      End
   End
   Begin VB.Menu mnuLog 
      Caption         =   "日志"
      Begin VB.Menu mmuLog 
         Caption         =   "记录日志"
         Begin VB.Menu mmuProcessLog 
            Caption         =   "记录通讯日志"
            Index           =   1
         End
         Begin VB.Menu mmuProcessLog 
            Caption         =   "记录处理日志"
            Index           =   2
         End
         Begin VB.Menu mmuProcessLog 
            Caption         =   "记录详细日志"
            Index           =   3
         End
      End
      Begin VB.Menu mnuFile_1 
         Caption         =   "-"
      End
      Begin VB.Menu mmuShowLog 
         Caption         =   "显示日志"
         Index           =   1
      End
      Begin VB.Menu mmuShowLog 
         Caption         =   "显示消息记录"
         Index           =   2
      End
      Begin VB.Menu mmuShowLog 
         Caption         =   "显示错误日志"
         Index           =   3
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "帮助(&H)"
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "关于(&A)..."
      End
   End
End
Attribute VB_Name = "frmHL7Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private LastState As Integer

Private WithEvents mobjIcon As clsTaskIcon  '托盘类
Attribute mobjIcon.VB_VarHelpID = -1
Private mfrmShowLog As frmShowLog

Private intWSListenerCount As Integer       '启动侦听的winsock的数量
Private intWSLinkCount As Integer           '接受连接请求的winsock的数量

Private mblnActionProcess As Boolean        '进入动作处理
Private mblnSendConnect As Boolean          '发送Socket连接成功
Private mblnSendACK As Boolean              '发送Socket接收到ACK响应

Private mlngIdleTime As Long                '累计Socket空闲的总时间，以秒为单位


Private Sub Form_Load()
        
    On Error Resume Next
    
    If WindowState = vbMinimized Then
        LastState = vbNormal
        Me.Hide
    Else
        LastState = WindowState
    End If
    
    
    '----------加载托盘图标
    Set mobjIcon = New clsTaskIcon
    mobjIcon.frmHwnd = tbrMain.hwnd ' hwnd
    mobjIcon.Icon = Icon.Handle
    mobjIcon.Message = "中联HL7服务网关"
    mobjIcon.AddIcon
    '----------加载托盘图标
    
    '建立到本地Access（日志记录）的连接
    gstrAccessPath = App.Path & "\ZlHL7Log"
    gstrAccessName = gstrAccessPath & ".mdb"
    
    With gcnAccess
        .ConnectionString = "DBQ=" & gstrAccessName & ";DefaultDir=" & App.Path & ";Driver={Microsoft Access Driver (*.mdb)}"
        .Open
        If .State = adStateClosed Then MsgBox "不能打开本地日志文件，系统将无法记录接收过程！", vbInformation, gstrSysName
    End With
    
    ''初始化数据，设置默认值
    intWSListenerCount = 0
    intWSLinkCount = 0
    ReDim gstrMsgQueue(0) As String
    ReDim gintTimeOut(0) As Integer
    gintQueueIndex = 1
    gblnQueueBusy = False
    mblnActionProcess = False
    mblnSendConnect = False
    mblnSendACK = False
    gblnServiceStart = True
    tbrMain.Buttons(3).Enabled = True
    tbrMain.Buttons(2).Enabled = False
    
    gstrRegPath = "私有模块\HL7接口"
    
    '读取数据库和注册表的参数
    Call ReadPara
    
    '启动消息接收服务
    Call subInputServiceSwitch(1)
    
    '设置窗口自动最小化
    Me.WindowState = vbMinimized
    
    '将启动情况记录到消息日志
    Call WriteMessageLog("HL7网关启动", Now & " HL7网关启动，版本为：" & App.Major & "." & App.Minor & "." & App.Revision & "，登录用户为：" & gstrDbUser)
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = vbFormControlMenu Then
        '不退出，只是最小化
        Cancel = True
        Me.WindowState = vbMinimized
    End If
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    If WindowState = 1 Then
        Me.Hide
        Exit Sub
    End If
    
    If WindowState <> vbMinimized Then
        LastState = WindowState
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim wsSocket As Winsock
    '清除托盘图标
    mobjIcon.DelIcon
    Set mobjIcon = Nothing
    
    '停止消息接收服务
    Call subInputServiceSwitch(0)
    
    '清除所有winsock侦听
    For Each wsSocket In wsHL7Server
        If wsSocket.State <> sckClosed Then wsSocket.Close
        If wsSocket.Index <> 0 Then Unload wsSocket
    Next
    For Each wsSocket In wsHL7Link
        If wsSocket.State <> sckClosed Then wsSocket.Close
        If wsSocket.Index <> 0 Then Unload wsSocket
    Next
    
    '将退出情况记录到消息日志
    Call WriteMessageLog("HL7网关退出", Now & " HL7网关退出")
    
End Sub

Private Sub mmuParaSetup_Click()
    Call frmParaSetup.zlSohwMe(Me)
End Sub


Private Sub mmuProcessLog_Click(Index As Integer)
    Dim i As Integer
    
    mmuProcessLog(Index).Checked = Not mmuProcessLog(Index).Checked
    gblnProcessLog = mmuProcessLog(Index).Checked
    glngProcessLogLevel = Index
    
    For i = 1 To 3
        If i <> Index Then
            mmuProcessLog(i).Checked = False
        End If
    Next i
End Sub

Private Sub mmuShowLog_Click(Index As Integer)
    Call subShowLog(Index)
End Sub

Private Sub mmuShowService_Click()
    Call subShowLog(4)
End Sub

Private Sub mnuFileQuit_Click()
    Dim dblSleepTime As Double
    
    gblnServiceStart = False
    
    dblSleepTime = timeGetTime
    
    '延时2秒钟再关闭，以便于停止发送服务
    While timeGetTime < dblSleepTime + 2000
        DoEvents
    Wend
    
    Unload Me
End Sub

Private Sub mnuHelpAbout_Click()
    gzlComLib.ShowAbout Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub mobjIcon_MouseLeftDBClick()
    '如果显示日志的模式窗口已经被打开，则退出，避免出现错误
    If mfrmShowLog Is Nothing Then
        If WindowState <> 1 Then
            WindowState = vbMinimized
            If WindowState = vbMinimized Then
                Me.Hide
            Else
                Me.Show
            End If
        Else
            WindowState = vbNormal
            Me.Show
        End If
    End If
End Sub

Private Sub munStartService_Click()
    gblnServiceStart = True
    
    tmHISAction.Enabled = True
    
    Call subInputServiceSwitch(1)
End Sub

Private Sub munStopService_Click()
    gblnServiceStart = False
    
    tmHISAction.Enabled = False
    
    Call subInputServiceSwitch(0)
End Sub

Private Sub subInputServiceSwitch(itype As Integer)
    'itype=0，停止服务;itype=1启动服务
        
    If itype = 0 Then
        '停止所有服务
        Call funListenPorts(0)
        tmFileInput.Enabled = False
    Else
        '启动服务时，确保只其中其中一种服务
        If gintInputDataType = 0 Then
            '启动HL7监听服务
            Call funListenPorts(1)
            tmFileInput.Enabled = False
        Else
            tmFileInput.Enabled = True
            Call funListenPorts(0)
        End If
    End If
End Sub

Private Sub tbrMain_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "启动服务"
            tbrMain.Buttons(3).Enabled = True
            Button.Enabled = False
            Call munStartService_Click
        Case "停止服务"
            tbrMain.Buttons(2).Enabled = True
            Button.Enabled = False
            Call munStopService_Click
        Case "帮助"
            Call mnuHelpAbout_Click
        Case "退出"
            Me.WindowState = vbMinimized ' mnuFileQuit_Click
    End Select
End Sub

Private Sub tbrMain_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    mobjIcon.MouseState x
End Sub

Private Function ReadPara() As Boolean
'------------------------------------------------
'功能： 读取基本参数
'参数：无
'返回：True --成功；False -- 失败
'------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    Dim strSQL As String
    Dim i As Integer
    
    On Error GoTo err
    
    '获取本机IP地址
    gstrLocalIP = funcGetLocalIP & ",127.0.0.1"
    
    strSQL = "Select IP地址,端口号,服务类型,发送程序名称,发送设备名称,接收程序名称,接收设备名称 From zlhis.hl7服务配置 " & _
                " Where instr([1],IP地址)>0"
    Set rsTemp = gzlDatabase.OpenSQLRecord(strSQL, "提取HL7服务配置", gstrLocalIP)
    
    ReDim HL7Services(rsTemp.RecordCount) As THL7Service
    i = 1
    
    While Not rsTemp.EOF
        HL7Services(i).strIP = Nvl(rsTemp!IP地址)
        HL7Services(i).lngPort = Val(Nvl(rsTemp!端口号, 0))
        HL7Services(i).intServiceType = Nvl(rsTemp!服务类型)
        HL7Services(i).strSendApp = Nvl(rsTemp!发送程序名称)
        HL7Services(i).strSendFacility = Nvl(rsTemp!发送设备名称)
        HL7Services(i).strReceiveApp = Nvl(rsTemp!接收程序名称)
        HL7Services(i).strReceiveFacility = Nvl(rsTemp!接收设备名称)
        i = i + 1
        rsTemp.MoveNext
    Wend
    
    gintTimeOutMax = Val(GetSetting("ZLSOFT", gstrRegPath, "超时", 60))
    gintInputDataType = Val(GetSetting("ZLSOFT", gstrRegPath, "接收消息方式", 0))
    If gintInputDataType <> 1 Then gintInputDataType = 0
    gstrFileDir = GetSetting("ZLSOFT", gstrRegPath, "文件消息目录", "")
    gstrFileSuffix = GetSetting("ZLSOFT", gstrRegPath, "文件消息后缀", "")
    gstrFileBackupDir = GetSetting("ZLSOFT", gstrRegPath, "文件消息备份目录", "")
    
    Exit Function
    
err:
    If gzlComLib.ErrCenter() = 1 Then Resume
    Call gzlComLib.SaveErrLog
End Function

Private Sub funListenPorts(itype As Integer)
'-----------------------------------------------------------------------------
'功能:启动或停止对服务端口的侦听
'参数: iType = 0 停止侦听；iType = 1 启动侦听。
'返回：True -- 成功；False -- 失败
'-----------------------------------------------------------------------------
    Dim strPort As String
    Dim i As Integer
    Dim lngResult As Long
    
    On Error Resume Next
    
    '启动本机接收服务端口
    For i = 1 To UBound(HL7Services)
        If HL7Services(i).intServiceType = 1 Then
            If InStr("," & strPort, HL7Services(i).lngPort) = 0 Then
                strPort = strPort & "," & HL7Services(i).lngPort
                If itype = 0 Then   '停止侦听
                    '停止侦听
                    If funWinsockUnlisten(HL7Services(i).lngPort) = 0 Then
                        HL7Services(i).Started = False
                    End If
                ElseIf itype = 1 Then   '启动侦听
                    '启动侦听
                    lngResult = funWinSockListen(HL7Services(i).lngPort)
                    If lngResult = 0 Then
                        HL7Services(i).Started = True
                    Else
                        HL7Services(i).Started = False
                        MsgBox "端口：" & HL7Services(i).lngPort & "已被使用，" & _
                                " 系统无法监听！请重新设置监听端口。", vbExclamation, gstrSysName
                    End If
                End If
            Else
                HL7Services(i).Started = IIf(itype = 0, False, True)
            End If
        End If
    Next i
        
End Sub

Private Function funWinSockListen(lngPort As Long) As Long
'-----------------------------------------------------------------------------
'功能:启动winsock端口的侦听
'参数:
'       lngPort ---侦听的端口号
'返回：0 -- 成功；1 -- 失败,端口被占用；2-失败，winsock索引不正确，创建新的winsock失败
'       3 -- 失败，其他错误
'-----------------------------------------------------------------------------
    'wsHL7Server()数组中，wsHL7Server（1）之后的实例负责监听，
    '监听到通讯请求后，创建一个新的wsHL7Link实例来接受连接请求
    
    On Error GoTo err
    
    intWSListenerCount = intWSListenerCount + 1
    
    '先检查当前winsock是否需要创建
    If wsHL7Server.Count = intWSListenerCount Then
        Load wsHL7Server(intWSListenerCount)
    Else
        'intWSListenerCount数量不对，直接返回错误
        funWinSockListen = 2
        Exit Function
    End If
    
    
    '先关闭一次winsock
    wsHL7Server(intWSListenerCount).Close
    wsHL7Server(intWSListenerCount).Bind lngPort
    wsHL7Server(intWSListenerCount).Listen
    If wsHL7Server(intWSListenerCount).State = sckListening Then
        funWinSockListen = 0        '侦听成功
    Else
        funWinSockListen = 3        '侦听失败,未知错误
    End If
    
    Exit Function
err:
    '出错不提示，直接返回启动端口失败
    If err.Number = 10048 Then
        funWinSockListen = 1    '端口被占用
    Else
        funWinSockListen = 3    '侦听失败,未知错误
    End If
End Function

Private Function funWinsockUnlisten(lngPort As Long) As Long
'-----------------------------------------------------------------------------
'功能:停止winsock端口的侦听
'参数:
'       lngPort ---侦听的端口号
'返回：0 -- 成功；1 -- 失败，其他错误
'-----------------------------------------------------------------------------
    Dim wsListener As Winsock
    
    On Error GoTo err
    
    '默认返回值是0，如果没有找到端口，认为这个端口的侦听被停止了
    funWinsockUnlisten = 0
    
    '首先查找正在监听该端口的winsock服务
    For Each wsListener In wsHL7Server
        If wsListener.LocalPort = lngPort And wsListener.State = sckListening And wsListener.Index <> 0 Then
            wsListener.Close
            Unload wsListener
            '查找并删除这个服务旗下创建的数据连接实例
            'wsHL7Link
            '待添加
            Exit For
        End If
    Next
    
    '如果只剩下一个wsHL7Server，则重新设置对象计数器
    If wsHL7Server.Count = 1 Then
        intWSListenerCount = 0
    End If
    Exit Function
err:
    funWinsockUnlisten = 1
End Function


Private Sub tmFileInput_Timer()
    '通过文件方式接收消息的定时器
    '定时轮询指定目录，找到文件后，解析文件，备份文件
    
    Dim strFileName As String
    Dim strHL7Msg As String
    
    On Error GoTo err
    
        
    '循环读取这个目录下的所有文件
    '使用dir()函数查询指定目录，是按照文件名来排序的，如果有多个消息并发接收时，有可能出现一个文件因为命名的原因，始终被放到最后。
    strFileName = Dir(gstrFileDir & "\*." & gstrFileSuffix)
    While strFileName <> ""
    
       
        '记录详细日志
        Call WriteProcessLog("tmFileInput_Timer", "1、接收到文件消息，准备处理", "文件名为：" & gstrFileDir & "\" & strFileName, 2)
    
        '提取文件
        strHL7Msg = funReadFile(gstrFileDir & "\" & strFileName)
        
        '记录详细日志
        Call WriteProcessLog("tmFileInput_Timer", "2、提取文件内容", "文件内容前半段：" & Left(strHL7Msg, 150), 3)
        
        
        '将文件内容加入消息队列
        Call MsgInQueue(strHL7Msg)
        
        '记录详细日志
        Call WriteProcessLog("tmFileInput_Timer", "3、消息入队", "文件名：" & strFileName, 3)
                
        '将文件移到备份目录，备份目录每天一个
        Call funBackupHL7File(gstrFileDir, strFileName)
        
        '记录详细日志
        Call WriteProcessLog("tmFileInput_Timer", "4、文件消息备份", "文件名：" & strFileName, 3)
        
        '读取下一个文件
        strFileName = Dir(gstrFileDir & "\*." & gstrFileSuffix)
        
        '记录详细日志
        Call WriteProcessLog("tmFileInput_Timer", "5、下一个文件名", "文件名：" & strFileName, 3)
    Wend
    
    Exit Sub
err:
    '错误处理
    If err.Number = 52 Then
        Call WriteLog(5003, err.Number, "tmFileInput_Timer 出现错误，错误描述是：" & err.Description & "，dir文件名=" & gstrFileDir & "\*." & gstrFileSuffix)
    Else
        Call WriteLog(5003, err.Number, "tmFileInput_Timer 出现错误，错误描述是：" & err.Description)
    End If
End Sub

Private Function funBackupHL7File(strFileDir As String, strFileName As String)
    Dim strPath As String
    
    On Error GoTo err
    '按照当天日期，创建目录，目录为\年\月+日
    strPath = "\" + Format(Date, "yyyy") + "\" + Format(Date, "mmdd") + "\"
    
    
    '创建备份目录
    Call MkLocalDir(gstrFileBackupDir + strPath + "\")
    
    '复制文件
    Call FileCopy(strFileDir + "\" + strFileName, gstrFileBackupDir + strPath + strFileName)
    Call Kill(strFileDir + "\" + strFileName)
    
    Exit Function
err:
    Call WriteLog(4004, err.Number, "funBackupHL7File 出现错误，strFileDir=" & strFileDir & "，strFileName=" & strFileName & "，错误描述是：" & err.Description)
End Function

Public Sub MkLocalDir(ByVal strDir As String)
'------------------------------------------------
'功能：创建本地目录
'参数： strDir－－本地目录
'返回：无
'------------------------------------------------
    Dim objFile As New Scripting.FileSystemObject
    Dim aNestDirs() As String, i As Integer
    Dim strPath As String
    On Error Resume Next
    
    '读取全部需要创建的目录信息
    ReDim Preserve aNestDirs(0)
    aNestDirs(0) = strDir
    
    strPath = objFile.GetParentFolderName(strDir)
    Do While Len(strPath) > 0
        ReDim Preserve aNestDirs(UBound(aNestDirs) + 1)
        aNestDirs(UBound(aNestDirs)) = strPath
        strPath = objFile.GetParentFolderName(strPath)
    Loop
    '创建全部目录
    For i = UBound(aNestDirs) To 0 Step -1
        MkDir aNestDirs(i)
    Next
End Sub

Private Function funReadFile(strFileName As String) As String
'读取并返回文件内容

    Dim fn As Integer
    Dim strContent As String
    
    fn = FreeFile
    
    On Error GoTo err
    
    Open strFileName For Binary Access Read As #fn
    strContent = Space(FileLen(strFileName))
    Get #fn, , strContent
    Close #fn
    
    funReadFile = strContent
        
    Exit Function
err:
    '关闭文件
    Close #fn
    Call WriteLog(4002, err.Number, "funReadFile 出现错误，strFileName=" & "，错误描述是：" & err.Description)
End Function


Private Sub tmHISAction_Timer()
    '定时轮询HIS数据库的HL7消息临时表
    On Error Resume Next
    '确保出错后，还能够把mblnActionProcess设置成true
    
    If mblnActionProcess = False Then
        mblnActionProcess = True
        Call funActionProcess
        mblnActionProcess = False
    End If
End Sub

Private Sub tmIdle_Timer()
    
    On Error GoTo err
    '如果出错，记录错误日志
    
    '空闲计时器1分钟运行一次
    
    '在程序里面再启动10分钟一次的计时器，每天凌晨3点10分，检查日志文件大小
    mlngIdleTime = mlngIdleTime + 1
    If mlngIdleTime > 10 Then
        '如果在凌晨3点15分之后的15分钟之内，则重发消息，检查日志文件
        If DateDiff("n", "03:15:00", Time) > 0 And DateDiff("n", "03:15:00", Time) < 15 Then
        
            '先停止tmIdle，避免在重发的过程中，被多次触发
            tmIdle.Enabled = False
            mlngIdleTime = 1
            

            '判断日志文件是否超过600M，超过则创建新的日志文件
            If FileLen(gstrAccessName) > 600000000 Then
                Call subNewLogFile
            End If
        
            '重发消息
            Call funReSendMessage
            
        End If
    End If
    
    '启动tmIdle
    tmIdle.Enabled = True
    
    
'-----------------------------以下方式无法使用-------------------------------
'由于GE的MUSE系统会自动保持socket连接，就算是在通讯的空闲时间，也会一直连接socket，
'因此不能用socket数量=1来判断socket是否空闲。

'    '累计Socket通讯的空闲时间，空闲时间够长，则启动重发消息的机制
'    'wsHL7Link数量为1，表示没有启动任何Socket来处理消息，是空闲
'    If wsHL7Link.Count = 1 Then
'        mlngIdleTime = mlngIdleTime + 1
'        '空闲十分钟之后，重发一次消息
'        If mlngIdleTime > 600 Then
'            mlngIdleTime = 1
'
'            '重发消息
'            Call funReSendMessage
'
'            '判断日志文件是否超过600M，超过则创建新的日志文件
'            If FileLen(gstrAccessName) > 600000000 Then
'                Call subNewLogFile
'            End If
'
'        End If
'    End If
'---------------------------------------------------------------------------
    Exit Sub
err:
    Call WriteLog(5001, err.Number, "tmIdle_Timer 定时器备份日志或重发消息出错。错误描述是：" & err.Description)
    tmIdle.Enabled = True
End Sub

Private Sub tmLinkTimeOut_Timer(Index As Integer)
    
    On Error GoTo err
    
    '处理计时，以及消息超时,处理ACK响应
    If Index <> 0 Then
        gintTimeOut(Index) = gintTimeOut(Index) + 1
        '超时
        If gintTimeOut(Index) > gintTimeOutMax Then
            Dim strACK As String
            
            Call GetMsgACK(tmLinkTimeOut(Index).Tag, strACK)
                            
            '返回ACK响应
            wsHL7Link(Index).SendData strACK
            
            '关闭和卸载消息计时器
            tmLinkTimeOut(Index).Enabled = False
            Unload tmLinkTimeOut(Index)
            
            gintTimeOut(Index) = 0
            '关闭消息连接
            wsHL7Link(Index).Close
            Unload wsHL7Link(Index)
        End If
    End If
    
    Exit Sub
err:
    Call WriteLog(5002, err.Number, "tmLinkTimeOut_Timer 定时器出现错误，Index=" & Index & "，错误描述是：" & err.Description)
End Sub

Private Sub tmMsgProcess_Timer()
    On Error GoTo err
    
    '先停止Timer计时，集中时间处理消息
    tmMsgProcess.Enabled = False
    Call funMsgProcess
    
    '清理系统缓存
    '如果wsHL7Link只有一个，则重新设置intWSLinkCount索引
    If wsHL7Link.Count = 1 Then
        intWSLinkCount = 0
        ReDim gintTimeOut(0) As Integer
    End If
    
    '消息处理完之后，再启动计时
    tmMsgProcess.Enabled = True
    Exit Sub
err:
    Call WriteLog(5004, err.Number, "tmMsgProcess_Timer 定时器出现错误，错误描述是：" & err.Description)
    tmMsgProcess.Enabled = True
End Sub

Private Sub Winsock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)

End Sub

Private Sub wsHL7Link_Close(Index As Integer)

    On Error Resume Next
    
    '同时卸载超时计时器
    Unload tmLinkTimeOut(Index)
    gintTimeOut(Index) = 0
    
'    wsHL7Link(Index).Close
    '在对方关闭连接的时候，把自己卸载了
    Unload wsHL7Link(Index)
End Sub

Private Sub wsHL7Link_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    
    Dim strACK As String
    Dim strData As String
    Dim lngMsgFullType As Long
    Dim strMessage As String
    Dim strRemain As String
    
    On Error GoTo err
    
    '记录详细日志
    Call WriteProcessLog("wsHL7Link_DataArrival", "接收到消息，准备处理", "wsHL7Link(" & Index & ")接收到消息内容，准备处理", 2)
    
    '接收数据
    wsHL7Link(Index).GetData strData, vbString
    
    '记录通讯日志
    Call WriteProcessLog("wsHL7Link_DataArrival", "wsHL7Link(" & Index & ")接收到消息内容：", strData, 1)
    
    '接收到第一个数据，才开始计时，同时初始化tag
    If tmLinkTimeOut(Index).Enabled = False Then
        tmLinkTimeOut(Index).Enabled = True
        gintTimeOut(Index) = 1
        tmLinkTimeOut(Index).Tag = ""
    End If
    
    '累计消息到tag中
    tmLinkTimeOut(Index).Tag = tmLinkTimeOut(Index).Tag & strData
    
    '提取和更新TAG中的消息，如果有完整消息，则返回ACK并入队
     
    While funGetMessage(tmLinkTimeOut(Index).Tag, strMessage, strRemain) = True
        '返回ACK，'消息入队
        If GetMsgACK(strMessage, strACK) = True Then
            '只有消息正确，才保存数据到HL7消息解析队列中
            '记录消息日志
            Call WriteMessageLog("wsHL7Link(" & Index & ")接收到心电结果消息", strMessage)
            Call MsgInQueue(strMessage)
        End If
        '不论消息是否正确，对完整消息都返回响应
        wsHL7Link(Index).SendData strACK
        '记录消息日志
        Call WriteMessageLog("wsHL7Link(" & Index & ")返回消息ACK", strACK)
        
        '更新tag的内容
        tmLinkTimeOut(Index).Tag = strRemain
    Wend
    
    Exit Sub
err:
    '暂不处理
    Call WriteLog(3001, err.Number, "wsHL7Link(" & Index & ") wsHL7Link_DataArrival 出现错误，错误描述是：" & err.Description)
End Sub


Private Sub wsHL7Server_Close(Index As Integer)
    '查找并删除这个服务旗下创建的数据连接实例
    'wsHL7Link
    '待添加
End Sub

Private Sub wsHL7Server_ConnectionRequest(Index As Integer, ByVal requestID As Long)
    '接收到通讯请求，启动一个新的实例来接收这个请求，并且创建一个Timer来记录连接时间
    
    '不用检查连接winsock的数量，因为每一个连接实例都是创建后，随着连接的关闭而关闭的,实例的索引不连续
    intWSLinkCount = intWSLinkCount + 1
    Load wsHL7Link(intWSLinkCount)
    
    '加载超时和分包处理机制
    Load tmLinkTimeOut(intWSLinkCount)
    
    tmLinkTimeOut(intWSLinkCount).Interval = 1000   '1秒运行一次
    ReDim Preserve gintTimeOut(intWSLinkCount) As Integer
    
    wsHL7Link(intWSLinkCount).LocalPort = 0
    wsHL7Link(intWSLinkCount).Accept requestID
    
    '通讯日志
    Call WriteProcessLog("wsHL7Server_ConnectionRequest", "接收到连接请求", _
        "从端口：" & wsHL7Server(Index).LocalPort & "，接收到连接请求，请求的IP是：" & wsHL7Server(Index).RemoteHostIP, 1)
        
        '记录详细日志
    Call WriteProcessLog("wsHL7Server_ConnectionRequest", "启动子socket-" & intWSLinkCount & "接收消息", "wsHL7Link(" & intWSLinkCount & ")状态为：" & wsHL7Link(intWSLinkCount).State, 2)
End Sub


Public Function funActionProcess() As Long
'-----------------------------------------------------------------------------
'功能:处理HIS的动作，轮询HL7中间表，一个个处理并发送消息，如果有停止服务的信号，则逐步停止消息的发送
'参数：
'返回值：无
'-----------------------------------------------------------------------------
     '定时轮询HIS数据库的HL7消息临时表
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim bln直接发送 As Boolean
    Dim str动作类型 As String
    Dim str业务ID串 As String
    Dim lng消息ID As Long
    Dim lngSendResult As Long
    
    On Error GoTo err
    
    '这里有点不对，一个待发消息，对应多个可以发送的消息配置，每个消息配置都发送成功了，才算是成功；
    
    Call CheckDBConnect
    
    '只发送3天之内的医嘱，如果一个收费单的医嘱3天之内没有缴费，则不再发送
    strSQL = "Select Id ,动作类型,业务ID串,发送次数,直接发送 From zlhis.hl7待发消息 Where 产生时间 >= Sysdate -3 order by ID"
    Set rsTemp = gzlDatabase.OpenSQLRecord(strSQL, "提取待发送的HL7消息")
    
    '记录详细日志
    Call WriteProcessLog("funActionProcess", "1、提取待发送的HL7消息", "strSQL = " + Replace(strSQL, "'", "‘"), 3)
        
    While Not rsTemp.EOF And gblnServiceStart = True
        
        
        
        '一个个组织消息并且发送
        bln直接发送 = IIf(Nvl(rsTemp!直接发送, 0) = 1, True, False)
        str动作类型 = Nvl(rsTemp!动作类型)
        str业务ID串 = Nvl(rsTemp!业务ID串, 0)
        lng消息ID = rsTemp!ID
        
        '记录详细日志
        Call WriteProcessLog("funActionProcess", "2、准备发送消息", "lng消息ID=" & lng消息ID & ",str业务ID串=" & str业务ID串, 3)
    
        lngSendResult = funSendAMessage(lng消息ID, str业务ID串, str动作类型, bln直接发送)
        
        If lngSendResult <> 2 Then
            Call WriteProcessLog("funActionProcess", "3、消息发送结果：" & IIf(lngSendResult = 0, "成功", "失败"), "消息ID：" & lng消息ID & "，业务串ID：" & str业务ID串 & "，动作类型：" & str动作类型, 3)
            
            '如果有发送失败的，先记录下来，发送成功的，删除数据库临时表记录
            strSQL = "zlhis.b_Hl7interface.HL7待发消息_UPDATE(" & lng消息ID & "," & IIf(lngSendResult = 0, "0", "1") & " ) "
            gzlDatabase.ExecuteProcedure strSQL, "消息发送后处理"
        End If
        
        rsTemp.MoveNext
    Wend
    Exit Function
err:
    Call WriteLog(3002, err.Number, "funActionProcess 出现错误，错误描述是：" & err.Description)
End Function

Private Sub wsSend_Connect()
    '这个事件发生，说明连接成功，可以发送消息了
    mblnSendConnect = True
End Sub

Private Sub wsSend_DataArrival(ByVal bytesTotal As Long)
    '发送消息之后，对方返回响应消息
    Dim strData As String
    
    On Error GoTo err
    
    '接收消息
    wsSend.GetData strData, vbString
    '把消息暂存在TAG中
    wsSend.Tag = strData
    mblnSendACK = True
    Exit Sub
err:
    Call WriteLog(3003, err.Number, "wsSend_DataArrival 出现错误，错误描述是：" & err.Description)
End Sub


Private Function funSendHL7Message(strIP As String, lngPort As Long, strMessage As String) As Long
'-----------------------------------------------------------------------------
'功能:发送一个HL7消息
'参数： strIP -- IP地址
'       lngPort -- 端口号
'       strMessage -- 消息内容
'返回值：0 -- 成功； 1 - 失败；2-停止服务
'-----------------------------------------------------------------------------
    Dim dblSleepTime As Double
    
    On Error GoTo err
                    
    funSendHL7Message = 1
    
    '先断开发送的连接，再重新连接
    If wsSend.State <> sckClosed Then wsSend.Close
    mblnSendConnect = False
    
    wsSend.RemoteHost = strIP
    wsSend.RemotePort = lngPort
    wsSend.LocalPort = 0    ''客户端的本地端口指定为0，在连接的时候，会自动使用空闲窗口，避免因为断开之后，需要等待5分钟才能再次连接
    wsSend.Connect
    If wsSend.State = sckConnecting Then
        Call WriteProcessLog("funSendHL7Message", "1、准备发送消息，正在连接", "IP地址是：" & strIP & "，端口是：" & lngPort & "，准备发送的消息是：" & strMessage, 3)
    End If

    ''一个延时,必须doevents才能触发Socket的connect事件
    dblSleepTime = timeGetTime
    '只有时间到了，或者发送成功了，或者退出程序了，才退出循环
    While timeGetTime < dblSleepTime + 2000 And mblnSendConnect = False And gblnServiceStart = True
        DoEvents
    Wend
    
    If mblnSendConnect = True Then
        '连接成功，则发送消息
        wsSend.SendData strMessage
        Call WriteProcessLog("funSendHL7Message", "2、发送HL7消息", "IP地址是：" & strIP & "，端口是：" & lngPort & "，准备发送的消息是：" & strMessage, 1)
        
        '发送消息之后，等待ACK响应
        mblnSendACK = False
        '再来一个延时，等待响应消息
        dblSleepTime = timeGetTime
        '只有时间到了，或者收到ACK成功了，或者退出程序了，才退出循环
        While timeGetTime < dblSleepTime + 2000 And mblnSendACK = False And gblnServiceStart = True
            DoEvents
        Wend
        
        '解析响应消息，如果是AA接收，则记录消息发送成功
        If mblnSendACK = True Then
            '接收到ACK响应，解析这个响应
            Call WriteProcessLog("funSendHL7Message", "3、收到ACK响应", "ACK消息内容是：" & wsSend.Tag, 1)
            If funParseACK(strMessage, wsSend.Tag) = 0 Then
                '正常接收，返回0
                funSendHL7Message = 0
            Else
                funSendHL7Message = 1
            End If
        Else
            '没有收到ACK响应，funSendHL7Message返回默认值1，发送失败
            Call WriteProcessLog("funSendHL7Message", "3.1、接收ACK响应超时", "接收ACK响应超时，消息发送失败", 1)
        End If
    Else
        '连接不成功，那就算是超时了，记录日志
        Call WriteProcessLog("funSendHL7Message", "2.1、连接超时", "IP地址是：" & strIP & "，端口是：" & lngPort & "，准备发送的消息是：" & strMessage, 1)
    End If
            
    If gblnServiceStart = False Then
        '如果中途出现用户干预，停止了发送服务，则返回2
        funSendHL7Message = 2
        Call WriteProcessLog("funSendHL7Message", "4、手工停止发送服务", "消息可能没有发送成功", 1)
    End If
        
    '消息处理完成后，是不是要断开wsSend的连接呢？
    wsSend.Close
    
    Exit Function
err:
    Call WriteLog(3004, err.Number, "funSendHL7Message 出现错误，错误描述是：" & err.Description)
End Function

Private Sub subShowLog(intType As Integer)
    On Error Resume Next
    
    Set mfrmShowLog = New frmShowLog
    mfrmShowLog.intLogType = intType
    mfrmShowLog.Show 1, Me
    Set mfrmShowLog = Nothing
    
End Sub

Private Function funReSendMessage() As Long
'-----------------------------------------------------------------------------
'功能:  重新发送HL7消息，处理“hl7重发消息”表中的消息，一次处理10个消息
'       如果有停止服务的信号，则逐步停止消息的发送
'参数：
'返回值：无
'-----------------------------------------------------------------------------
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim bln直接发送 As Boolean
    Dim str动作类型 As String
    Dim str业务ID串 As String
    Dim lng消息ID As Long
    Dim lngSendResult As Long
    
    On Error GoTo err
    
    Call WriteProcessLog("funReSendMessage", "1、重发消息", "准备重发消息", 3)
    
    '查询“hl7重发消息”表
    strSQL = "Select ID,动作类型,业务ID串,产生时间,发送次数,直接发送,重发时间,消息时效 From zlhis.hl7重发消息  Where  消息时效 = 1  Order By 重发时间 "
    Set rsTemp = gzlDatabase.OpenSQLRecord(strSQL, "提取需要重发的全部HL7消息")
    
    '循环发送消息
    While Not rsTemp.EOF And gblnServiceStart = True
        '一个个组织消息并且发送
        bln直接发送 = IIf(Nvl(rsTemp!直接发送, 0) = 1, True, False)
        str动作类型 = Nvl(rsTemp!动作类型)
        str业务ID串 = Nvl(rsTemp!业务ID串, 0)
        lng消息ID = rsTemp!ID
        
        lngSendResult = funSendAMessage(lng消息ID, str业务ID串, str动作类型, bln直接发送)
        
        If lngSendResult <> 2 Then
            
            Call WriteProcessLog("funReSendMessage", "2、重发消息，发送结果：" & IIf(lngSendResult = 0, "成功", "失败"), "消息ID：" & lng消息ID & "，业务串ID：" & str业务ID串 & "，动作类型：" & str动作类型, 3)
            
            '如果有发送失败的，先记录下来，发送成功的，删除数据库临时表记录
            strSQL = "zlhis.b_Hl7interface.HL7重发消息_UPDATE(" & lng消息ID & "," & IIf(lngSendResult = 0, "0", "1") & " ) "
            gzlDatabase.ExecuteProcedure strSQL, "消息重发后处理"
            
        Else
            '消息没有发送，可能是医嘱未计费
            Call WriteProcessLog("funReSendMessage", "2.2、消息没有发送", "可能是医嘱没有计费", 3)
        End If
        
        rsTemp.MoveNext
    Wend
    Exit Function
err:
    Call WriteLog(3005, err.Number, "funReSendMessage 出现错误，错误描述是：" & err.Description)
End Function

Private Function funSendAMessage(lng消息ID As Long, str业务ID串 As String, str动作类型 As String, bln直接发送 As Boolean) As Long
'-----------------------------------------------------------------------------
'功能:  发送一条HL7消息
'参数： lng消息ID   ---消息ID
'       str业务ID串 ---消息的业务ID串
'       str动作类型 ---消息的动作类型
'       bln直接发送 ---是否直接发送消息
'返回值：0--发送成功；1--发送失败；2--没有发送，比如未交费
'-----------------------------------------------------------------------------
    Dim strSQL As String
    Dim rsFee As ADODB.Recordset
    Dim arrMsgs As THl7Messages
    Dim iMsg As Integer
    Dim lngResult As Long
    Dim blnSendOK As Boolean
    
    On Error GoTo err
    
    '通过直接发送字段判断消息是否直接发送给MUSE，否则就检查收费状态
    If bln直接发送 = False And str动作类型 = HL7_MSG_SEND_NEW_ORDER Then
        '新开医嘱，而且不是直接发送，则查询数据库，获取医嘱的收费情况
        strSQL = "Select 医嘱ID From zlhis.病人医嘱发送  " & _
                 " Where 记录性质=1 And 计费状态 In (1,2) And 医嘱id = [1] And 发送号 = [2]"
        Set rsFee = gzlDatabase.OpenSQLRecord(strSQL, "提取医嘱的计费状态", CLng(Split(str业务ID串, ";")(0)), CLng(Split(str业务ID串, ";")(1)))
        
        If rsFee.RecordCount = 0 Then
            '已收费,可以发送
            bln直接发送 = True
        End If
    End If
    
    If bln直接发送 = True Then
    
        '记录详细日志
        Call WriteProcessLog("funSendAMessage", "1、准备提取消息定义", "str动作类型=" & str动作类型, 3)
        
        '从数据库中填充该动作的HL7消息结构
        '从数据库提取消息定义
        arrMsgs = getMsgDefFromDB(str动作类型)
        
        '查询并填写每一个消息段的字段内容
        If UBound(arrMsgs.arrMsgs) > 0 Then
            '记录待发消息的ID
            arrMsgs.lngActionID = lng消息ID
            '如果是多次发送的消息，则复制多个消息记录
'                If Nvl(rsTemp!发送次数, 1) > 1 Then
'                    Call funDuplicateMsg(arrMsgs, CInt(Nvl(rsTemp!发送次数, 1)))
'                End If
            
            '记录详细日志
            Call WriteProcessLog("funSendAMessage", "2、准备填充消息内容", "str业务ID串=" & str业务ID串, 3)
        
            '根据字段设置填写数据库或者固定内容
            Call funfillMsgValue(arrMsgs, str业务ID串)
            
            '发送每一个消息
            For iMsg = 1 To UBound(arrMsgs.arrMsgs)
                lngResult = funSendHL7Message(arrMsgs.arrMsgs(iMsg).strIP, arrMsgs.arrMsgs(iMsg).lngPort, arrMsgs.arrMsgs(iMsg).strText)
                If lngResult = 0 Then
                    arrMsgs.arrMsgs(iMsg).blnSendOK = True
                    '记录消息记录
                    Call WriteMessageLog(arrMsgs.arrMsgs(iMsg).strActionType, arrMsgs.arrMsgs(iMsg).strText)
                ElseIf lngResult = 2 Then   '停止发送服务
                    Exit For
                End If
            Next iMsg
            
            '在这里检查消息是否发送成功，通过过程处理待发消息记录
            blnSendOK = True
            '检查每一个被配置的消息是否发送成功
            For iMsg = 1 To UBound(arrMsgs.arrMsgs)
                If arrMsgs.arrMsgs(iMsg).blnSendOK = False Then
                    blnSendOK = False
                    Exit For
                End If
            Next iMsg
            
        Else
            blnSendOK = False
        End If
        funSendAMessage = IIf(blnSendOK = True, 0, 1)
    Else
        funSendAMessage = 2 '消息没有发送
    End If
    
    Exit Function
err:
    Call WriteLog(3006, err.Number, "funSendAMessage 出现错误，错误描述是：" & err.Description)
End Function
