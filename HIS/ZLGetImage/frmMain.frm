VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "中联图像下载"
   ClientHeight    =   1500
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3735
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1500
   ScaleWidth      =   3735
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.Timer tmMsg 
      Interval        =   100
      Left            =   120
      Top             =   1080
   End
   Begin VB.Label Label1 
      Caption         =   "正在下载图像，请勿关闭。                                    图像下载完成后会自动关闭。。。。。。"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   3495
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents mobjIcon As clsTaskIcon  '托盘类
Attribute mobjIcon.VB_VarHelpID = -1


Private Sub Form_Load()
    '初始化数据
    ReDim pstrMsgQueue(0) As String
    ReDim pConnectedSharedDir(0) As String
    pintQueueIndex = 1
    
    Set pftpConnect = New clsFtp
    pftpConnect.strLogPath = pstrLogPath
    pftpConnect.lngLogLevel = plngLogLevel
    pftpConnect.blnLogEnable = pblnLogEnable
    
    '自己产生一个消息,测试用
'    Call MsgInQueue("\20100512\3977915000\||D:\HAH||127.0.0.1||||PACS||PACS||FTP||hj||minona")
    
'    Call MsgInQueue("20100512\3977915||D:\HAH||127.0.0.1||||PACS||PACS||||||")
    
     '----------加载托盘图标
    Set mobjIcon = New clsTaskIcon
    mobjIcon.frmHwnd = frmMain.hwnd ' hwnd
    mobjIcon.Icon = Icon.Handle
    mobjIcon.Message = "ZL图像下载"
    mobjIcon.AddIcon
    '----------加载托盘图标
    
'    挂上截获消息的hook
    plngPreWndProc = Hook(Me.hwnd)
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mobjIcon.MouseState X
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    '断开FTP连接
    Call pftpConnect.FuncFtpDisConnect
    
    '清除托盘图标
    mobjIcon.DelIcon
    Set mobjIcon = Nothing
'
'    卸载hook
    Unhook Me.hwnd, plngPreWndProc
End Sub


Private Sub mobjIcon_MouseLeftDBClick()
    If WindowState <> vbMinimized Then
        WindowState = vbMinimized
        Me.Hide
    Else
        WindowState = vbNormal
        Me.Show
    End If
End Sub

Private Sub tmMsg_Timer()
    If funMsgProcess = True Then
        Unload Me
    End If
End Sub
