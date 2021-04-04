VERSION 5.00
Object = "{B1790453-7708-48C1-B5CC-75255FA4B066}#1.0#0"; "ZLDSVideoProcess.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   Caption         =   "Form1"
   ClientHeight    =   5850
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9615
   LinkTopic       =   "Form1"
   ScaleHeight     =   5850
   ScaleWidth      =   9615
   StartUpPosition =   3  '窗口缺省
   Begin ZLDSVideoProcess.DSPlay DSPlay 
      Height          =   5175
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   5415
      Object.Visible         =   -1  'True
      AutoScroll      =   0   'False
      AutoSize        =   0   'False
      AxBorderStyle   =   1
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
      ShowModel       =   2
      IsFullScreen    =   0   'False
      IsFit           =   0   'False
      IsStretch       =   -1  'True
      IsAdjustWindowSize=   0   'False
      IsShowState     =   -1  'True
      IsEscKeyQuitFullScreen=   -1  'True
      IsDblClickQuitFullScreen=   0   'False
      IsClickQuitFullScreen=   0   'False
      CurWidth        =   361
      CurHeight       =   345
      SnatchWay       =   0
      AppHandle       =   0
   End
   Begin VB.Frame fraPanel 
      Height          =   5655
      Left            =   5760
      TabIndex        =   1
      Top             =   120
      Width           =   3855
      Begin VB.Frame fraImgSnatchWay 
         Caption         =   "图像抓取方式"
         Height          =   975
         Left            =   1920
         TabIndex        =   26
         Top             =   3000
         Width           =   1815
         Begin VB.OptionButton optDevice 
            Caption         =   "DEVICE"
            Height          =   255
            Left            =   120
            TabIndex        =   28
            Top             =   600
            Width           =   1335
         End
         Begin VB.OptionButton optVMR 
            Caption         =   "VMR"
            Height          =   255
            Left            =   120
            TabIndex        =   27
            Top             =   240
            Value           =   -1  'True
            Width           =   1215
         End
      End
      Begin VB.Timer tmrState 
         Interval        =   100
         Left            =   3000
         Top             =   -240
      End
      Begin VB.CommandButton cmdRePlay 
         Caption         =   "重复播放"
         Height          =   495
         Left            =   240
         TabIndex        =   24
         Top             =   2400
         Width           =   1455
      End
      Begin VB.CommandButton cmdGetSize 
         Caption         =   "取得视频大小"
         Height          =   495
         Left            =   240
         TabIndex        =   23
         Top             =   4440
         Width           =   1455
      End
      Begin VB.CheckBox chkShowState 
         Caption         =   "显示状态栏"
         Height          =   255
         Left            =   1920
         TabIndex        =   22
         Top             =   4200
         Value           =   1  'Checked
         Width           =   1575
      End
      Begin VB.CommandButton cmdStop 
         Caption         =   "停止"
         Height          =   495
         Left            =   240
         TabIndex        =   12
         Top             =   1210
         Width           =   1455
      End
      Begin VB.Frame fraPlay 
         Caption         =   "播放模式"
         Height          =   1095
         Left            =   1920
         TabIndex        =   19
         Top             =   1800
         Width           =   1815
         Begin VB.OptionButton OptStretch 
            Caption         =   "拉伸播放"
            Height          =   255
            Left            =   120
            TabIndex        =   21
            Top             =   360
            Value           =   -1  'True
            Width           =   1335
         End
         Begin VB.OptionButton OptScale 
            Caption         =   "按比例播放"
            Height          =   255
            Left            =   120
            TabIndex        =   20
            Top             =   720
            Width           =   1335
         End
      End
      Begin MSComDlg.CommonDialog dlgSave 
         Left            =   2400
         Top             =   -240
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         DefaultExt      =   "*.Jpg"
         Filter          =   "*.Jpg|*.Jpg|*.bmp|*.bmp"
      End
      Begin VB.Frame fraImgFormat 
         Caption         =   "图像格式"
         Height          =   735
         Left            =   1920
         TabIndex        =   16
         Top             =   240
         Width           =   1815
         Begin VB.OptionButton optJpg 
            Caption         =   "JPG"
            Height          =   255
            Left            =   120
            TabIndex        =   18
            Top             =   360
            Value           =   -1  'True
            Width           =   615
         End
         Begin VB.OptionButton optBmp 
            Caption         =   "BMP"
            Height          =   255
            Left            =   960
            TabIndex        =   17
            Top             =   360
            Width           =   615
         End
      End
      Begin VB.CommandButton cmdCaptureImg 
         Caption         =   "抓取图像"
         Height          =   495
         Left            =   1920
         TabIndex        =   15
         Top             =   1080
         Width           =   1815
      End
      Begin VB.CommandButton cmdPause 
         Caption         =   "暂停"
         Height          =   495
         Left            =   240
         TabIndex        =   14
         Top             =   720
         Width           =   1455
      End
      Begin VB.CommandButton cmdPlay 
         Caption         =   "播放"
         Height          =   495
         Left            =   240
         TabIndex        =   13
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton cmdVideoInf 
         Caption         =   "视频信息"
         Height          =   495
         Left            =   240
         TabIndex        =   11
         Top             =   3120
         Width           =   1455
      End
      Begin VB.CommandButton cmdFullScreen 
         Caption         =   "全屏播放"
         Height          =   495
         Left            =   240
         TabIndex        =   10
         Top             =   3615
         Width           =   1455
      End
      Begin VB.CommandButton cmdFirstFrame 
         Caption         =   "|<"
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   9
         Top             =   5160
         Width           =   375
      End
      Begin VB.CommandButton cmdFirstFrame 
         Caption         =   "<<"
         Height          =   375
         Index           =   1
         Left            =   615
         TabIndex        =   8
         Top             =   5160
         Width           =   375
      End
      Begin VB.CommandButton cmdFirstFrame 
         Caption         =   ">>"
         Height          =   375
         Index           =   2
         Left            =   990
         TabIndex        =   7
         Top             =   5160
         Width           =   375
      End
      Begin VB.CommandButton cmdFirstFrame 
         Caption         =   ">|"
         Height          =   375
         Index           =   3
         Left            =   1365
         TabIndex        =   6
         Top             =   5160
         Width           =   375
      End
      Begin VB.CommandButton cmdContinue 
         Caption         =   "继续"
         Height          =   495
         Left            =   240
         TabIndex        =   5
         Top             =   1710
         Width           =   1455
      End
      Begin VB.CommandButton cmdAddRate2 
         Caption         =   "2X"
         Height          =   375
         Index           =   0
         Left            =   1740
         TabIndex        =   4
         Top             =   5160
         Width           =   375
      End
      Begin VB.CommandButton cmdAddRate2 
         Caption         =   "Normal"
         Height          =   375
         Index           =   1
         Left            =   2115
         TabIndex        =   3
         Top             =   5160
         Width           =   735
      End
      Begin VB.CommandButton cmdAddRate2 
         Caption         =   "1/2X"
         Height          =   375
         Index           =   2
         Left            =   2850
         TabIndex        =   2
         Top             =   5160
         Width           =   615
      End
      Begin MSComDlg.CommonDialog dlgVideoSelect 
         Left            =   1920
         Top             =   -240
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         DefaultExt      =   "*.avi"
         Filter          =   "*.avi|*.avi|*.*|*.*"
      End
      Begin VB.Label labFrame 
         Caption         =   "当前帧:0"
         Height          =   255
         Left            =   1920
         TabIndex        =   25
         Top             =   4680
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkShowState_Click()
  DSPlay.IsShowState = chkShowState.Value
End Sub

Private Sub cmdAddRate2_Click(Index As Integer)
  '播放速度设置
  Select Case Index
  Case 0
    DSPlay.PlayRate = 2
  Case 1
    DSPlay.PlayRate = 1
  Case 2
    DSPlay.PlayRate = 0.5
  End Select
End Sub

Private Sub cmdCaptureImg_Click()
  '采集图像
  Dim sErrMsg As String
  dlgSave.ShowSave
  
  If optJpg.Value Then
    sErrMsg = DSPlay.CaptureJpgImgToFile(dlgSave.FileName, 100)
  Else
    sErrMsg = DSPlay.CaptureBmpImgToFile(dlgSave.FileName)
  End If
  
  If Trim(sErrMsg) <> "" Then
    MsgBox sErrMsg
    Exit Sub
  End If
  
  Call Form2.ShowCaptureImg(dlgSave.FileName)
End Sub

Private Sub cmdContinue_Click()
  '继续播放
  Dim sErrMsg As String
  
  sErrMsg = DSPlay.Run
  If Trim(sErrMsg) <> "" Then
    MsgBox sErrMsg
  End If
End Sub

Private Sub cmdFirstFrame_Click(Index As Integer)
  '移动指定帧
  Select Case Index
  Case 0
    DSPlay.FirstFrame
  Case 1
    DSPlay.PriorFrame
  Case 2
    DSPlay.NextFrame
  Case 3
    DSPlay.LastFrame
  End Select
  
End Sub

Private Sub cmdFullScreen_Click()
  '全屏播放
  Dim sErrMsg As String
  
  sErrMsg = DSPlay.ShowFullScreen(App.hInstance, GetMonitorIndex(Me.hwnd))
  If Trim(sErrMsg) <> "" Then
    MsgBox sErrMsg
  End If
End Sub

Private Sub cmdGetSize_Click()
  Dim sErrMsg As String
  Dim sWidth As String
  Dim sHeight As String
  
  sErrMsg = DSPlay.GetVideoProperty(vpVideoWidth, sWidth)
  If Trim(sErrMsg) <> "" Then
    Call MsgBox(sErrMsg)
  End If
  
  sErrMsg = DSPlay.GetVideoProperty(vpVideoHeight, sHeight)
  
  MsgBox "宽:" & sWidth & " 高:" & sHeight
End Sub

Private Sub cmdPause_Click()
  '暂停
  Dim sErrMsg As String
  
  sErrMsg = DSPlay.Pause
  
  If Trim(sErrMsg) <> "" Then
    Call MsgBox(sErrMsg)
  End If
End Sub

Private Sub cmdPlay_Click()
  '打开文件播放
  dlgVideoSelect.ShowOpen
  
  If Trim(dlgVideoSelect.FileName) = "" Then Exit Sub
  If Dir(dlgVideoSelect.FileName) = "" Then Exit Sub
  
  Dim sErrMsg As String
  
  sErrMsg = DSPlay.Play(dlgVideoSelect.FileName)
  
  If Trim(sErrMsg) <> "" Then
    Call MsgBox(sErrMsg)
  End If
End Sub

Private Sub cmdRePlay_Click()
  '重复播放
  Dim sErrMsg As String
  
  sErrMsg = DSPlay.RePlay
  
  If Trim(sErrMsg) <> "" Then
    Call MsgBox(sErrMsg)
  End If
End Sub

Private Sub cmdStop_Click()
  '停止播放
  Dim sErrMsg As String
  
  sErrMsg = DSPlay.Stop
  
  If Trim(sErrMsg) <> "" Then
    Call MsgBox(sErrMsg)
  End If
End Sub

Private Sub cmdVideoInf_Click()
  '显示视频信息
  Dim sErrMsg As String
  
  sErrMsg = DSPlay.ShowVideoInfo(Me.hwnd)
  If Trim(sErrMsg) <> "" Then
    MsgBox sErrMsg
  End If
End Sub


Private Sub Form_Resize()
  fraPanel.Top = 0
  fraPanel.Left = Me.Width - fraPanel.Width - 100
    
  DSPlay.AppHandle = Me.hwnd
  DSPlay.Top = 0
  DSPlay.Left = 0
  
  If Me.Width - fraPanel.Width - 140 >= 0 Then
    DSPlay.Width = Me.Width - fraPanel.Width - 140
  End If
  
  If Me.Height - 400 >= 0 Then
    DSPlay.Height = Me.Height - 400
  End If
  
  Call DSPlay.RefreshWindow
End Sub

Private Sub Form_Unload(Cancel As Integer)
  DSPlay.FreeRes
End Sub

Private Sub optDevice_Click()
  If optVMR.Value Then
    DSPlay.SnatchWay = swVMR
  Else
    DSPlay.SnatchWay = swDEVICE
  End If
End Sub

Private Sub OptScale_Click()
  '按比例播放
  If OptScale.Value Then
    DSPlay.ShowModel = smFit
  End If
End Sub

Private Sub OptStretch_Click()
  '拉伸播放
  If OptStretch.Value Then
    DSPlay.ShowModel = smStretch
  End If
End Sub

Private Sub optVMR_Click()
  If optVMR.Value Then
    DSPlay.SnatchWay = swVMR
  Else
    DSPlay.SnatchWay = swDEVICE
  End If
End Sub

Private Sub tmrState_Timer()
  labFrame.Caption = "当前帧:" & CStr(DSPlay.CurFrame)
End Sub
