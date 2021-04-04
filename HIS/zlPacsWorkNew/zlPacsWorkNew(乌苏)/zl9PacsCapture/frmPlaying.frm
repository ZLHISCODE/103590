VERSION 5.00
Object = "{5C493D4E-FD57-4FF4-9BA4-C6C670BFF9A7}#70.0#0"; "zl9PacsControl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{B1790453-7708-48C1-B5CC-75255FA4B066}#1.0#0"; "ZLDSVIDEOPROCESS.OCX"
Begin VB.Form frmPlaying 
   BackColor       =   &H00000000&
   Caption         =   "视频播放"
   ClientHeight    =   7590
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   12855
   Icon            =   "frmPlaying.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7590
   ScaleWidth      =   12855
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox picControl 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1215
      Left            =   0
      ScaleHeight     =   1215
      ScaleWidth      =   12855
      TabIndex        =   1
      Top             =   6375
      Width           =   12855
      Begin zl9PacsCapture.ImageButton imbSound 
         Height          =   330
         Left            =   4680
         TabIndex        =   17
         Top             =   470
         Width           =   330
         _ExtentX        =   582
         _ExtentY        =   582
         hPicture        =   "frmPlaying.frx":038A
         nPicture        =   "frmPlaying.frx":09B4
         dPicture        =   "frmPlaying.frx":0FDE
         wPicture        =   "frmPlaying.frx":1608
         ScaleHeight     =   330
         ScaleWidth      =   330
         ScaleLeft       =   0
         ScaleTop        =   0
         ScaleMode       =   1
         Hwnd            =   3868696
      End
      Begin zl9PacsCapture.ZLScrollBar scbSound 
         Height          =   195
         Left            =   5040
         TabIndex        =   16
         Top             =   520
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   344
         Appearance      =   0
         AutoRedraw      =   -1  'True
         BorderStyle     =   1
         ScaleHeight     =   165
         ScaleWidth      =   1665
         ScaleLeft       =   0
         ScaleTop        =   0
         ScaleMode       =   1
         BackColor       =   4210752
         Hwnd            =   2689518
         Position        =   100
         BeginColor      =   49152
         EndColor        =   49152
         ShpMoveVisible  =   0   'False
         AutoShowBlock   =   0   'False
      End
      Begin zl9PacsCapture.ImageButton imbCapture 
         Height          =   720
         Left            =   3720
         TabIndex        =   15
         Top             =   240
         Width           =   720
         _ExtentX        =   1270
         _ExtentY        =   1270
         hPicture        =   "frmPlaying.frx":1C32
         nPicture        =   "frmPlaying.frx":3784
         dPicture        =   "frmPlaying.frx":52D6
         wPicture        =   "frmPlaying.frx":6E28
         ScaleHeight     =   720
         ScaleWidth      =   720
         ScaleLeft       =   0
         ScaleTop        =   0
         ScaleMode       =   1
         Hwnd            =   2559500
         Hint            =   "图像采集"
      End
      Begin VB.Timer timShow 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   4920
         Top             =   1080
      End
      Begin VB.Timer timPlayer 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   3840
         Top             =   1080
      End
      Begin MSComDlg.CommonDialog comDialog 
         Left            =   3000
         Top             =   1080
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin zl9PacsCapture.ImageButton imbFullScreen 
         Height          =   480
         Left            =   3240
         TabIndex        =   10
         Top             =   360
         Width           =   480
         _ExtentX        =   847
         _ExtentY        =   847
         hPicture        =   "frmPlaying.frx":897A
         nPicture        =   "frmPlaying.frx":95CC
         dPicture        =   "frmPlaying.frx":A21E
         wPicture        =   "frmPlaying.frx":AE70
         ScaleHeight     =   480
         ScaleWidth      =   480
         ScaleLeft       =   0
         ScaleTop        =   0
         ScaleMode       =   1
         Hwnd            =   2361820
         Hint            =   "全屏"
      End
      Begin zl9PacsCapture.ImageButton imbFirst 
         Height          =   480
         Left            =   1320
         TabIndex        =   8
         ToolTipText     =   "第一帧"
         Top             =   360
         Width           =   480
         _ExtentX        =   847
         _ExtentY        =   847
         hPicture        =   "frmPlaying.frx":BAC2
         nPicture        =   "frmPlaying.frx":C714
         dPicture        =   "frmPlaying.frx":D366
         wPicture        =   "frmPlaying.frx":DFB8
         ScaleHeight     =   480
         ScaleWidth      =   480
         ScaleLeft       =   0
         ScaleTop        =   0
         ScaleMode       =   1
         Hwnd            =   6555844
         Hint            =   "第一帧"
      End
      Begin zl9PacsCapture.ImageButton imbLast 
         Height          =   480
         Left            =   1800
         TabIndex        =   7
         ToolTipText     =   "上一帧"
         Top             =   360
         Width           =   480
         _ExtentX        =   847
         _ExtentY        =   847
         hPicture        =   "frmPlaying.frx":EC0A
         nPicture        =   "frmPlaying.frx":F85C
         dPicture        =   "frmPlaying.frx":104AE
         wPicture        =   "frmPlaying.frx":11100
         ScaleHeight     =   480
         ScaleWidth      =   480
         ScaleLeft       =   0
         ScaleTop        =   0
         ScaleMode       =   1
         Hwnd            =   1573040
         Hint            =   "上一帧"
      End
      Begin zl9PacsCapture.ImageButton imbEnd 
         Height          =   480
         Left            =   2760
         TabIndex        =   6
         ToolTipText     =   "最后帧"
         Top             =   360
         Width           =   480
         _ExtentX        =   847
         _ExtentY        =   847
         hPicture        =   "frmPlaying.frx":11D52
         nPicture        =   "frmPlaying.frx":129A4
         dPicture        =   "frmPlaying.frx":135F6
         wPicture        =   "frmPlaying.frx":14248
         ScaleHeight     =   480
         ScaleWidth      =   480
         ScaleLeft       =   0
         ScaleTop        =   0
         ScaleMode       =   1
         Hwnd            =   2821090
         Hint            =   "最后帧"
      End
      Begin zl9PacsCapture.ImageButton imbNext 
         Height          =   480
         Left            =   2280
         TabIndex        =   5
         ToolTipText     =   "下一帧"
         Top             =   360
         Width           =   480
         _ExtentX        =   847
         _ExtentY        =   847
         hPicture        =   "frmPlaying.frx":14E9A
         nPicture        =   "frmPlaying.frx":15AEC
         dPicture        =   "frmPlaying.frx":1673E
         wPicture        =   "frmPlaying.frx":17390
         ScaleHeight     =   480
         ScaleWidth      =   480
         ScaleLeft       =   0
         ScaleTop        =   0
         ScaleMode       =   1
         Hwnd            =   2424886
         Hint            =   "下一帧"
      End
      Begin zl9PacsCapture.ImageButton imbStop 
         Height          =   480
         Left            =   840
         TabIndex        =   4
         ToolTipText     =   "停止"
         Top             =   360
         Width           =   480
         _ExtentX        =   847
         _ExtentY        =   847
         hPicture        =   "frmPlaying.frx":17FE2
         nPicture        =   "frmPlaying.frx":18C34
         dPicture        =   "frmPlaying.frx":19886
         wPicture        =   "frmPlaying.frx":1A4D8
         ScaleHeight     =   480
         ScaleWidth      =   480
         ScaleLeft       =   0
         ScaleTop        =   0
         ScaleMode       =   1
         Hwnd            =   4262290
         Hint            =   "停止"
      End
      Begin zl9PacsCapture.ImageButton imbPlay 
         Height          =   705
         Left            =   120
         TabIndex        =   3
         ToolTipText     =   "播放/暂停"
         Top             =   240
         Width           =   705
         _ExtentX        =   1244
         _ExtentY        =   1244
         hPicture        =   "frmPlaying.frx":1B12A
         nPicture        =   "frmPlaying.frx":1CC7C
         dPicture        =   "frmPlaying.frx":1E73E
         wPicture        =   "frmPlaying.frx":20200
         ScaleHeight     =   705
         ScaleWidth      =   705
         ScaleLeft       =   0
         ScaleTop        =   0
         ScaleMode       =   1
         Hwnd            =   3804182
         Hint            =   "播放"
      End
      Begin zl9PacsCapture.ImageButton imbPause 
         Height          =   720
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   720
         _ExtentX        =   1270
         _ExtentY        =   1270
         hPicture        =   "frmPlaying.frx":21CC2
         nPicture        =   "frmPlaying.frx":23814
         dPicture        =   "frmPlaying.frx":25366
         wPicture        =   "frmPlaying.frx":26EB8
         ScaleHeight     =   720
         ScaleWidth      =   720
         ScaleLeft       =   0
         ScaleTop        =   0
         ScaleMode       =   1
         Hwnd            =   2426926
         Hint            =   "暂停"
      End
      Begin zl9PacsCapture.ImageButton imbNoSound 
         Height          =   330
         Left            =   4680
         TabIndex        =   18
         Top             =   480
         Visible         =   0   'False
         Width           =   330
         _ExtentX        =   582
         _ExtentY        =   582
         hPicture        =   "frmPlaying.frx":28A0A
         nPicture        =   "frmPlaying.frx":29034
         dPicture        =   "frmPlaying.frx":2965E
         wPicture        =   "frmPlaying.frx":29C88
         ScaleHeight     =   330
         ScaleWidth      =   330
         ScaleLeft       =   0
         ScaleTop        =   0
         ScaleMode       =   1
         Hwnd            =   2886950
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00404040&
         X1              =   6840
         X2              =   6840
         Y1              =   0
         Y2              =   1200
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00404040&
         X1              =   4560
         X2              =   4560
         Y1              =   0
         Y2              =   1200
      End
      Begin VB.Label labTime 
         BackStyle       =   0  'Transparent
         Caption         =   "0:00:00/0:00:00"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   7560
         TabIndex        =   11
         Top             =   100
         Width           =   1455
      End
      Begin VB.Image imgBackGround 
         Height          =   1215
         Left            =   0
         Picture         =   "frmPlaying.frx":2A2B2
         Stretch         =   -1  'True
         Top             =   0
         Width           =   12375
      End
   End
   Begin VB.PictureBox picPlayer 
      Align           =   3  'Align Left
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5760
      Left            =   0
      ScaleHeight     =   5760
      ScaleWidth      =   9135
      TabIndex        =   0
      Top             =   0
      Width           =   9135
      Begin ZLDSVideoProcess.DSPlay DSPlayer 
         Height          =   6015
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   9015
         Object.Visible         =   -1  'True
         AutoScroll      =   0   'False
         AutoSize        =   0   'False
         AxBorderStyle   =   0
         Caption         =   "DSPlay"
         Color           =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         KeyPreview      =   0   'False
         PixelsPerInch   =   96
         PrintScale      =   1
         Scaled          =   -1  'True
         DropTarget      =   0   'False
         HelpFile        =   ""
         ScreenSnap      =   0   'False
         SnapBuffer      =   10
         DoubleBuffered  =   0   'False
         Enabled         =   -1  'True
         CurTime         =   -1
         CurFrame        =   -1
         PlayRate        =   -1
         ShowModel       =   1
         IsFullScreen    =   0   'False
         IsFit           =   -1  'True
         IsStretch       =   0   'False
         IsAdjustWindowSize=   0   'False
         IsShowState     =   -1  'True
         IsEscKeyQuitFullScreen=   -1  'True
         IsDblClickQuitFullScreen=   0   'False
         IsClickQuitFullScreen=   0   'False
         CurWidth        =   601
         CurHeight       =   401
         SnatchWay       =   0
         AppHandle       =   0
         Volume          =   0
         Balance         =   0
         IsSoundHint     =   0   'False
         IsDebugFilter   =   0   'False
         videoFile       =   ""
      End
   End
   Begin VB.PictureBox picInf 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   0
      ScaleHeight     =   585
      ScaleWidth      =   12825
      TabIndex        =   9
      Top             =   5760
      Width           =   12855
      Begin zl9PacsCapture.ZLScrollBar scbState 
         Height          =   200
         Left            =   0
         TabIndex        =   12
         Top             =   375
         Width           =   9255
         _ExtentX        =   16325
         _ExtentY        =   344
         Appearance      =   0
         AutoRedraw      =   -1  'True
         BorderStyle     =   1
         ScaleHeight     =   165
         ScaleWidth      =   9225
         ScaleLeft       =   0
         ScaleTop        =   0
         ScaleMode       =   0
         BackColor       =   -2147483643
         Hwnd            =   2034686
         Hint            =   "定位"
         ShpMoveVisible  =   0   'False
      End
      Begin VB.Label labText 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   350
         Left            =   0
         TabIndex        =   14
         Top             =   80
         Width           =   12855
      End
   End
   Begin VB.Menu 文件 
      Caption         =   "文件(&F)"
      Begin VB.Menu 打开 
         Caption         =   "打开(&O)"
      End
      Begin VB.Menu 另存为 
         Caption         =   "另存为(&A)..."
      End
      Begin VB.Menu s6 
         Caption         =   "-"
      End
      Begin VB.Menu 退出 
         Caption         =   "退出(&Q)"
      End
   End
   Begin VB.Menu 播放 
      Caption         =   "播放(&P)"
      Begin VB.Menu 播放暂停 
         Caption         =   "播放/暂停(&P)"
      End
      Begin VB.Menu 停止 
         Caption         =   "停止(&S)"
      End
      Begin VB.Menu 播放速度 
         Caption         =   "播放速度(&R)"
         Begin VB.Menu 快速 
            Caption         =   "快速(&K)"
         End
         Begin VB.Menu 正常 
            Caption         =   "正常(&Z)"
            Checked         =   -1  'True
         End
         Begin VB.Menu 慢速 
            Caption         =   "慢速(&M)"
         End
      End
      Begin VB.Menu s1 
         Caption         =   "-"
      End
      Begin VB.Menu 起始帧 
         Caption         =   "起始帧(&F)"
      End
      Begin VB.Menu 结束帧 
         Caption         =   "结束帧(&E)"
      End
      Begin VB.Menu 上一帧 
         Caption         =   "上一帧(&L)"
      End
      Begin VB.Menu 下一帧 
         Caption         =   "下一帧(&N)"
      End
      Begin VB.Menu s4 
         Caption         =   "-"
      End
      Begin VB.Menu 音量 
         Caption         =   "音量(&U)"
         Begin VB.Menu 增大 
            Caption         =   "+ 增大(&A)"
         End
         Begin VB.Menu 减小 
            Caption         =   "-  减小(&D)"
         End
         Begin VB.Menu 静音 
            Caption         =   "静音(&M)"
         End
      End
      Begin VB.Menu 声道 
         Caption         =   "声道(&H)"
         Begin VB.Menu 左声道 
            Caption         =   "左声道(&L)"
         End
         Begin VB.Menu 正常声道 
            Caption         =   "正常(&N)"
            Checked         =   -1  'True
         End
         Begin VB.Menu 右声道 
            Caption         =   "右声道(&R)"
         End
      End
      Begin VB.Menu s10 
         Caption         =   "-"
      End
      Begin VB.Menu 采集图像 
         Caption         =   "采集图像(&C)"
      End
   End
   Begin VB.Menu 查看 
      Caption         =   "查看(&V)"
      Begin VB.Menu 播放时自动隐藏 
         Caption         =   "播放时自动隐藏(&H)"
         Checked         =   -1  'True
      End
      Begin VB.Menu 控制栏 
         Caption         =   "控制栏(&C)"
         Checked         =   -1  'True
         Enabled         =   0   'False
      End
      Begin VB.Menu 信息栏 
         Caption         =   "信息栏(&I)"
         Checked         =   -1  'True
         Enabled         =   0   'False
      End
      Begin VB.Menu s2 
         Caption         =   "-"
      End
      Begin VB.Menu 全屏 
         Caption         =   "全屏(&F)"
      End
      Begin VB.Menu s3 
         Caption         =   "-"
      End
      Begin VB.Menu 显示方式 
         Caption         =   "显示方式(&Y)"
         Begin VB.Menu 按比例缩放 
            Caption         =   "按比例缩放(&O)"
            Checked         =   -1  'True
         End
         Begin VB.Menu 视频拉伸 
            Caption         =   "视频拉伸(&S)"
         End
         Begin VB.Menu 实际大小 
            Caption         =   "实际大小(&E)"
         End
      End
      Begin VB.Menu 渲染方式 
         Caption         =   "渲染方式(&W)"
         Begin VB.Menu VMR 
            Caption         =   "VMR(&V)"
            Checked         =   -1  'True
         End
         Begin VB.Menu DEVICE 
            Caption         =   "DEVICE(&D)"
         End
      End
      Begin VB.Menu s5 
         Caption         =   "-"
      End
      Begin VB.Menu 媒体信息 
         Caption         =   "媒体信息(&O)"
      End
   End
   Begin VB.Menu 帮助 
      Caption         =   "帮助(&H)"
   End
End
Attribute VB_Name = "frmPlaying"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Declare Function SetMenu Lib "user32" (ByVal hWnd As Long, ByVal hMenu As Long) As Long
Private Declare Function GetMenu Lib "user32" (ByVal hWnd As Long) As Long


  
    
Private mstrInfo As String

Private mlngEventTimeLen As Long
Private hMenu As Long
Private mobjCallBack As Object
Private mblnIsActive As Boolean




Public Event OnCapture(pic As StdPicture)




Property Get IsActive() As Boolean
    IsActive = mblnIsActive
End Property


Private Sub HideMenu()
    If hMenu <> 0 Then Exit Sub
    
    ' 获得菜单句柄，并隐藏菜单。
    hMenu = GetMenu(hWnd)
    SetMenu hWnd, 0
    
    Call RefreshPlayerFace(True)
End Sub

Private Sub ShowMenu()
    If hMenu = 0 Then Exit Sub

    SetMenu hWnd, hMenu
    hMenu = 0
    
    Call RefreshPlayerFace(True)
End Sub


Private Sub SwitchPlayButton(blnPlay As Boolean)
    imbPause.Visible = blnPlay
    imbPlay.Visible = Not blnPlay
End Sub

Private Sub SwitchSoundButton(blnIsSound As Boolean)
    imbSound.Visible = blnIsSound
    imbNoSound.Visible = Not blnIsSound
End Sub


Private Sub DEVICE_Click()
    On Error Resume Next
    
    DSPlayer.SnatchWay = swDEVICE
    
    VMR.Checked = False
    DEVICE.Checked = True
End Sub

Private Sub DSPlayer_OnDblClick()
    If Not DSPlayer.IsFullScreen Then
        Call DSPlayer.ShowFullScreen(App.hInstance, GetMonitorIndex(hWnd))
    Else
        Call DSPlayer.QuitFullScreen
    End If
End Sub

Private Sub DSPlayer_OnMouseMove(ByVal Shift As Long, ByVal X As Long, ByVal Y As Long)
'    If Y > 25 Then
'        HideMenu
'    Else
'        ShowMenu
'    End If
    If mlngEventTimeLen <> -1 Then mlngEventTimeLen = 0
End Sub

Private Sub Form_Activate()
    mblnIsActive = True
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If mlngEventTimeLen <> -1 Then mlngEventTimeLen = 0
End Sub


Private Sub LoadParameterConfig()
    
    '保存注册表参数
      
    '播放时自动隐藏
    播放时自动隐藏.Checked = GetSetting("ZLSOFT", G_STR_REG_PATH_PUBLIC, "播放时自动隐藏", True)
    '控制栏
    控制栏.Checked = GetSetting("ZLSOFT", G_STR_REG_PATH_PUBLIC, "播放控制栏", True)
    '信息栏
    信息栏.Checked = GetSetting("ZLSOFT", G_STR_REG_PATH_PUBLIC, "播放信息栏", True)
    
    
    控制栏.Enabled = Not 播放时自动隐藏.Checked
    信息栏.Enabled = Not 播放时自动隐藏.Checked
    
    
    picControl.Visible = 控制栏.Checked
    picInf.Visible = 信息栏.Checked
    
    
    '按比例缩放
    按比例缩放.Checked = GetSetting("ZLSOFT", G_STR_REG_PATH_PUBLIC, "播放时按比例缩放", True)
    If 按比例缩放.Checked Then DSPlayer.ShowModel = smFit
    '视频拉伸
    视频拉伸.Checked = GetSetting("ZLSOFT", G_STR_REG_PATH_PUBLIC, "播放时视频拉伸", False)
    If 视频拉伸.Checked Then DSPlayer.ShowModel = smStretch
    '实际大小
    实际大小.Checked = GetSetting("ZLSOFT", G_STR_REG_PATH_PUBLIC, "播放时实际大小", False)
    If 实际大小.Checked Then DSPlayer.ShowModel = smNormal
    
    
    'VMR
    VMR.Checked = GetSetting("ZLSOFT", G_STR_REG_PATH_PUBLIC, "播放时VMR模式", True)
    If VMR.Checked Then DSPlayer.SnatchWay = swVMR
    
    'DEVICE
    DEVICE.Checked = GetSetting("ZLSOFT", G_STR_REG_PATH_PUBLIC, "播放时DEVICE模式", False)
    If DEVICE.Checked Then DSPlayer.SnatchWay = swDEVICE
End Sub

Private Sub Form_Load()
    SetWindowPos Me.hWnd, -1, Me.CurrentX, Me.CurrentY, Me.ScaleWidth, Me.ScaleHeight, 3 '将窗口置顶
    
    '恢复窗体设置
    Call zlCL_RestoreWinState(Me, App.ProductName)
    
    '载入参数配置
    Call LoadParameterConfig
    
    mblnIsActive = False
    mlngEventTimeLen = 0
    
    Set mobjCallBack = Nothing
    
    DSPlayer.AppHandle = Me.hWnd
    DSPlayer.IsShowState = False
    DSPlayer.IsEscKeyQuitFullScreen = True
    'DSPlayer.IsDblClickQuitFullScreen = True
'    DSPlayer.ShowModel = smFit
    'DSPlayer.ShowAnimate atQiu
    'DSPlayer.SnatchWay = swDEVICE
    

    If Trim(mstrInfo) = "" Then
        labText.Caption = IIf(DSPlayer.VideoState = vsPlay, "正在播放...", IIf(DSPlayer.VideoState = vsPause, "暂停中...", "准备就绪..."))
    Else
        labText.Caption = mstrInfo
    End If
End Sub

Public Sub OpenVideoFile(Optional strFileName As String = "", Optional ByRef objCallBack As Object = Nothing)
    On Error Resume Next
    Dim strPlayFile As String
    
    Set mobjCallBack = objCallBack
    
    strPlayFile = strFileName
    If Trim(strPlayFile) = "" Then
        comDialog.DefaultExt = ".AVI"
        comDialog.Filter = "(*.AVI)|*.AVI|(*.MPEG)|*.MPEG|(*.*)|*.*"
    
        comDialog.ShowOpen
        
        strPlayFile = comDialog.FileName
    End If
    
    If Trim(strPlayFile) <> "" Then
        If Trim(Dir(strPlayFile)) <> "" Then
            Dim sErrMsg As String
            
            sErrMsg = DSPlayer.Play(strPlayFile)
            
            If Trim(sErrMsg) <> "" Then
                Call MsgboxCus(sErrMsg, vbOKOnly, G_STR_HINT_TITLE)
                Exit Sub
            End If
            
            scbState.position = 0
            scbState.Min = 0
            scbState.Max = DSPlayer.timeLen
            
            scbSound.position = DSPlayer.Volume
            
            timPlayer.Enabled = True
            timShow.Enabled = True
            
            Call SwitchPlayButton(DSPlayer.VideoState = vsPlay)
            
'            If DSPlayer.StreamTypeName = "Audio" Then
'                Call DSPlayer.ShowAnimate(atQiu)
'            End If
            
        End If
    End If
End Sub


Private Sub PlayVideo()
    If DSPlayer.VideoState = vsStop Then
        DSPlayer.RePlay
        
        scbState.position = 0
        scbState.Min = 0
        scbState.Max = DSPlayer.timeLen
        timPlayer.Enabled = True
        
        Call SwitchPlayButton(DSPlayer.VideoState = vsPlay)
        
        Exit Sub
    End If
    
    DSPlayer.Run
    Call SwitchPlayButton(DSPlayer.VideoState = vsPlay)
End Sub



Private Sub PauseVideo()
    DSPlayer.Pause
    Call SwitchPlayButton(DSPlayer.VideoState = vsPlay)
End Sub


Private Sub StopVideo()
    DSPlayer.Stop
    Call SwitchPlayButton(DSPlayer.VideoState = vsPlay)
End Sub


Private Sub RefreshPlayerFace(Optional blnOnlyRefreshPlayer As Boolean = False)
    picPlayer.Width = Me.Width
    
    DSPlayer.Left = 0
    DSPlayer.Top = 0
    DSPlayer.Width = picPlayer.Width
    DSPlayer.Height = picPlayer.Height
    
    Call DSPlayer.RefreshWindow
    
    If blnOnlyRefreshPlayer Then Exit Sub
    
    scbState.Left = 0
    scbState.Top = picInf.Height - scbState.Height - 30
    scbState.Width = picInf.Width - 10
    
    labText.Width = picInf.Width
    
    
    imgBackGround.Left = 0
    imgBackGround.Top = 0
    imgBackGround.Width = picControl.Width
    
    labTime.Left = picControl.Width - labTime.Width - 30
End Sub


Private Sub Form_Resize()
    On Error Resume Next
    
    Call RefreshPlayerFace
End Sub


Private Sub SaveParameterConfig()
  '保存注册表参数

  '播放时自动隐藏
  SaveSetting "ZLSOFT", G_STR_REG_PATH_PUBLIC, "播放时自动隐藏", 播放时自动隐藏.Checked
  '控制栏
  SaveSetting "ZLSOFT", G_STR_REG_PATH_PUBLIC, "播放控制栏", 控制栏.Checked
  '信息栏
  SaveSetting "ZLSOFT", G_STR_REG_PATH_PUBLIC, "播放信息栏", 信息栏.Checked
  
  
  '按比例缩放
  SaveSetting "ZLSOFT", G_STR_REG_PATH_PUBLIC, "播放时按比例缩放", 按比例缩放.Checked
  '视频拉伸
  SaveSetting "ZLSOFT", G_STR_REG_PATH_PUBLIC, "播放时视频拉伸", 视频拉伸.Checked
  '实际大小
  SaveSetting "ZLSOFT", G_STR_REG_PATH_PUBLIC, "播放时实际大小", 实际大小.Checked

  
  'VMR
  SaveSetting "ZLSOFT", G_STR_REG_PATH_PUBLIC, "播放时VMR模式", VMR.Checked
  'DEVICE
  SaveSetting "ZLSOFT", G_STR_REG_PATH_PUBLIC, "播放时DEVICE模式", DEVICE.Checked
  
End Sub

Private Sub Form_Unload(Cancel As Integer)
    '保存窗体设置
    Call zlCL_SaveWinState(Me, App.ProductName)
    
    Call SaveParameterConfig
    
End Sub

Private Sub imbCapture_OnClick()
    '采集播放的图像
    If Not (mobjCallBack Is Nothing) Then
        Call mobjCallBack.subCaptureImg(True, "", DSPlayer.CaptureBmpImage)
    End If
End Sub


Private Sub imbEnd_OnClick()
    On Error Resume Next
    
    DSPlayer.LastFrame
End Sub

Private Sub imbFirst_OnClick()
    On Error Resume Next
    
    DSPlayer.FirstFrame
End Sub

Private Sub imbFullScreen_OnClick()
    '全屏播放
    On Error Resume Next

    Call DSPlayer.ShowFullScreen(App.hInstance, GetMonitorIndex(Me.hWnd))
End Sub

Private Sub imbLast_OnClick()
    On Error Resume Next
    
    DSPlayer.PriorFrame
End Sub

Private Sub imbNext_OnClick()
    On Error Resume Next
    
    DSPlayer.NextFrame
End Sub

Private Sub imbNoSound_OnClick()
    On Error Resume Next
    
    Call 静音_Click
    
'    DSPlayer.Volume = scbSound.Position * 100
'
'    Call SwitchSoundButton(True)
End Sub


Private Sub imbPause_OnClick()
    On Error Resume Next
    
    Call PauseVideo
End Sub

Private Sub imbPlay_OnClick()
    On Error Resume Next
    
    Call PlayVideo
End Sub

Private Sub imbSound_OnClick()
    On Error Resume Next
    
    Call 静音_Click
'    DSPlayer.Volume = 0
'
'    Call SwitchSoundButton(False)
End Sub


Private Sub imbStop_OnClick()
    On Error Resume Next
    
    Call StopVideo
End Sub

Private Function ToTimeFormat(time As Long) As String
    Dim lngHour As Long, lngMinute As Long, lngSecond As Long
    
    lngSecond = time Mod 60
    lngMinute = IIf(Int(time / 60) >= 60, Int(time / 60) Mod 60, Int(time / 60))
    lngHour = Int(time / 3600)
    
    ToTimeFormat = Format(lngHour & ":" & lngMinute & ":" & lngSecond)
End Function


Private Sub imgBackGround_DblClick()
    Call OpenVideoFile("", mobjCallBack)
End Sub


Private Sub imgBackGround_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If mlngEventTimeLen <> -1 Then mlngEventTimeLen = 0
End Sub

Private Sub labText_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If mlngEventTimeLen <> -1 Then mlngEventTimeLen = 0
End Sub

Private Sub scbSound_OnMouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If mlngEventTimeLen <> -1 Then mlngEventTimeLen = 0
End Sub

Private Sub scbSound_OnPositionChange(lngOldPosition As Long, lngNewPostion As Long)
    DSPlayer.Volume = lngNewPostion * 100
End Sub

Private Sub scbState_OnMouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If mlngEventTimeLen <> -1 Then mlngEventTimeLen = 0
End Sub

Private Sub scbState_OnPositionChange(lngOldPosition As Long, lngNewPostion As Long)
    If lngNewPostion = lngOldPosition Then Exit Sub
    
    DSPlayer.CurTime = lngNewPostion
End Sub

Private Sub timPlayer_Timer()
    Select Case DSPlayer.VideoState
        Case vsPlay, vsPause:
            labTime.Caption = ToTimeFormat(DSPlayer.CurTime) & "/" & ToTimeFormat(DSPlayer.timeLen)
            scbState.position = DSPlayer.CurTime
        Case vsStop:
            labTime.Caption = "0:00:00/0:00:00"
            scbState.position = 0
            
            timPlayer.Enabled = False
            
            Call SwitchPlayButton(False)
    End Select
    
    If Trim(mstrInfo) = "" Then
        labText.Caption = IIf(DSPlayer.VideoState = vsPlay, "正在播放...", IIf(DSPlayer.VideoState = vsPause, "暂停中...", "准备就绪..."))
    End If
End Sub

Private Sub timShow_Timer()
    On Error Resume Next
    
    If Not 播放时自动隐藏.Checked Then Exit Sub
    If DSPlayer.IsFullScreen Then Exit Sub
    
    '当视频停止播放时,显示出所有播放界面
    If DSPlayer.VideoState = vsStop Then
        picControl.Visible = True
        picInf.Visible = True
        
        Call ShowMenu
        
        timShow.Enabled = False
        Exit Sub
    End If
         
    
    Dim i As Long
    
    mlngEventTimeLen = mlngEventTimeLen + 1
    
    If mlngEventTimeLen = 16 Then
        mlngEventTimeLen = 15
        Exit Sub
    End If
    
    If mlngEventTimeLen >= 10 Then
        mlngEventTimeLen = -1
        
        picControl.Visible = False
        picInf.Visible = False
        
        Call HideMenu
        
        For i = 0 To 10
            Call Sleep(20)
            DoEvents
        Next i
'        Call Sleep(500)
        
        mlngEventTimeLen = 15
    Else
        picControl.Visible = True
        picInf.Visible = True
        
        Call ShowMenu
    End If
    
    Call RefreshPlayerFace(True)
End Sub

Private Sub VMR_Click()
    On Error Resume Next
    
    DSPlayer.SnatchWay = swVMR
    
    VMR.Checked = True
    DEVICE.Checked = False
    
End Sub


Private Sub 按比例缩放_Click()
    On Error Resume Next
    
    DSPlayer.ShowModel = smFit
    DSPlayer.RefreshWindow
    
    按比例缩放.Checked = True
    视频拉伸.Checked = False
    实际大小.Checked = False
End Sub

Private Sub 播放时自动隐藏_Click()
    播放时自动隐藏.Checked = Not 播放时自动隐藏.Checked
    
    控制栏.Enabled = Not 播放时自动隐藏.Checked
    信息栏.Enabled = Not 播放时自动隐藏.Checked
End Sub

Private Sub 播放暂停_Click()
    On Error Resume Next
    If imbPlay.Visible Then
        Call PlayVideo
        Exit Sub
    End If
    
    If imbPause.Visible Then
        Call PauseVideo
        Exit Sub
    End If
End Sub

Private Sub 采集图像_Click()
    '采集播放的图像
    If Not (mobjCallBack Is Nothing) Then
        Call mobjCallBack.subCaptureImg(True, "", DSPlayer.CaptureBmpImage)
    End If
End Sub

Private Sub 打开_Click()
    On Error Resume Next
    Call OpenVideoFile("", mobjCallBack)
End Sub

Private Sub 减小_Click()
    On Error Resume Next
    
    DSPlayer.Volume = DSPlayer.Volume - 100
    
    scbSound.position = Round(DSPlayer.Volume / 100)
End Sub

Private Sub 结束帧_Click()
    On Error Resume Next
    DSPlayer.LastFrame
End Sub

Private Sub 静音_Click()
    On Error Resume Next
    
    If Not 静音.Checked Then
        DSPlayer.Volume = 0
        静音.Checked = True
    Else
        DSPlayer.Volume = scbSound.position * 100
        静音.Checked = False
    End If
     
    Call SwitchSoundButton(Not 静音.Checked)
End Sub

Private Sub 控制栏_Click()
    On Error Resume Next
    
    picControl.Visible = Not picControl.Visible
    控制栏.Checked = picControl.Visible
    
    Call RefreshPlayerFace(True)
End Sub

Private Sub 快速_Click()
    On Error Resume Next
    
    DSPlayer.PlayRate = 2
    快速.Checked = True
    慢速.Checked = False
    正常.Checked = False
End Sub


Private Sub VideoSaveAs()
    Dim strFileName As String
    Dim strFileType As String
    
    If Trim(DSPlayer.videoFile) = "" Then
        MsgboxCus "没有可另存的视频文件。", vbOKOnly Or vbInformation, G_STR_HINT_TITLE
        Exit Sub
    End If

    comDialog.Filter = "(*.AVI)|*.AVI|(*.MPEG)|*.MPEG|(*.*)|*.*"

    comDialog.ShowSave
    strFileName = comDialog.FileName
    If strFileName <> "" Then
        '复制视频文件到指定路径
        Call FileCopy(DSPlayer.videoFile, strFileName)
    End If

End Sub

Private Sub 另存为_Click()
    '另存为录像
    Call VideoSaveAs
End Sub

Private Sub 慢速_Click()
    On Error Resume Next
    
    DSPlayer.PlayRate = 0.5
    慢速.Checked = True
    快速.Checked = False
    正常.Checked = False
End Sub

Private Sub 媒体信息_Click()
    On Error Resume Next
    
    Call DSPlayer.ShowVideoInfo(hWnd)
End Sub

Private Sub 起始帧_Click()
    On Error Resume Next
    
    DSPlayer.FirstFrame
End Sub

Private Sub 全屏_Click()
    On Error Resume Next
    
    Call DSPlayer.ShowFullScreen(App.hInstance, GetMonitorIndex(hWnd))
End Sub

Private Sub 上一帧_Click()
    On Error Resume Next
    DSPlayer.PriorFrame
End Sub


Private Sub 实际大小_Click()
    On Error Resume Next
    
    DSPlayer.ShowModel = smNormal
    DSPlayer.RefreshWindow
    
    按比例缩放.Checked = False
    视频拉伸.Checked = False
    实际大小.Checked = True
End Sub

Private Sub 视频拉伸_Click()
    On Error Resume Next
    
    DSPlayer.ShowModel = smStretch
    DSPlayer.RefreshWindow
    
    按比例缩放.Checked = False
    视频拉伸.Checked = True
    实际大小.Checked = False
End Sub

Private Sub 停止_Click()
    On Error Resume Next
    Call StopVideo
End Sub

Private Sub 退出_Click()
    On Error Resume Next
    Call Unload(Me)
End Sub

Private Sub 下一帧_Click()
    On Error Resume Next
    DSPlayer.NextFrame
End Sub

Private Sub 信息栏_Click()
    On Error Resume Next
    
    picInf.Visible = Not picInf.Visible
    信息栏.Checked = picInf.Visible
    
    Call RefreshPlayerFace(True)
End Sub


Private Sub 右声道_Click()
    On Error Resume Next
    
    DSPlayer.Balance = -10000
    
    左声道.Checked = False
    正常声道.Checked = False
    右声道.Checked = True
End Sub

Private Sub 增大_Click()
    On Error Resume Next
    
    DSPlayer.Volume = DSPlayer.Volume + 100
    
    scbSound.position = Round(DSPlayer.Volume / 100)
End Sub

Private Sub 正常_Click()
    On Error Resume Next
    
    DSPlayer.PlayRate = 1
    正常.Checked = True
    慢速.Checked = False
    快速.Checked = False
End Sub

Private Sub 正常声道_Click()
    On Error Resume Next
    
    DSPlayer.Balance = 0
    
    左声道.Checked = False
    正常声道.Checked = True
    右声道.Checked = False
End Sub

Private Sub 左声道_Click()
    On Error Resume Next
    
    DSPlayer.Balance = 10000
    
    左声道.Checked = True
    正常声道.Checked = False
    右声道.Checked = False
End Sub
