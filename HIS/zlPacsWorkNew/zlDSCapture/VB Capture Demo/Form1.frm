VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{B1790453-7708-48C1-B5CC-75255FA4B066}#1.0#0"; "ZLDSVI~1.OCX"
Begin VB.Form Form1 
   Caption         =   "�ɼ����Գ���"
   ClientHeight    =   10590
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10950
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10590
   ScaleWidth      =   10950
   StartUpPosition =   2  '��Ļ����
   Begin ZLDSVideoProcess.DSCapture objDSCapture 
      Height          =   4935
      Left            =   2160
      TabIndex        =   0
      Top             =   1320
      Width           =   6495
      Object.Visible         =   -1  'True
      AutoScroll      =   0   'False
      AutoSize        =   0   'False
      AxBorderStyle   =   0
      Caption         =   ""
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
      KeyPreview      =   -1  'True
      PixelsPerInch   =   96
      PrintScale      =   1
      Scaled          =   -1  'True
      DropTarget      =   0   'False
      HelpFile        =   ""
      ScreenSnap      =   0   'False
      SnapBuffer      =   10
      DoubleBuffered  =   0   'False
      Enabled         =   -1  'True
      IsStretch       =   0   'False
      IsShowState     =   -1  'True
      IsFullScreen    =   0   'False
      IsAdjustWindowSize=   0   'False
      IsFit           =   0   'False
      IsEscKeyQuitFullScreen=   -1  'True
      IsDblClickQuitFullScreen=   0   'False
      IsClickQuitFullScreen=   0   'False
      CurWidth        =   433
      CurHeight       =   329
      CurVideoWidth   =   433
      CurVideoHeight  =   311
      ShowModel       =   0
      CapParameterWindPos=   8
      SnatchWay       =   0
      ParameterCfgFileName=   ""
      HideCfgItem     =   0
      AppHandle       =   0
   End
   Begin VB.CommandButton cmdScreen 
      Caption         =   "screenIndex"
      Height          =   375
      Left            =   7200
      TabIndex        =   26
      Top             =   10200
      Width           =   735
   End
   Begin VB.Frame fraControl 
      Height          =   3255
      Left            =   120
      TabIndex        =   1
      Top             =   7320
      Width           =   10695
      Begin VB.CommandButton Command2 
         Caption         =   "RealVideoSize"
         Height          =   375
         Left            =   7800
         TabIndex        =   29
         Top             =   2280
         Width           =   1575
      End
      Begin VB.CommandButton Command1 
         Caption         =   "�����ڴ�"
         Height          =   495
         Left            =   1920
         TabIndex        =   28
         Top             =   840
         Width           =   1575
      End
      Begin VB.CommandButton cmdFromClipBoard 
         Caption         =   "���Լ�����"
         Height          =   495
         Left            =   2760
         TabIndex        =   27
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmdCustomParameter 
         Caption         =   "�ɼ���������1"
         Height          =   495
         Left            =   1800
         TabIndex        =   24
         Top             =   2640
         Width           =   1815
      End
      Begin VB.Frame fraSnatchWay 
         Caption         =   "ͼ��ץȡ��ʽ"
         Height          =   615
         Left            =   3960
         TabIndex        =   21
         Top             =   2520
         Width           =   2415
         Begin VB.OptionButton optDevice 
            Caption         =   "DEVICE"
            Height          =   255
            Left            =   840
            TabIndex        =   23
            Top             =   240
            Width           =   1095
         End
         Begin VB.OptionButton optVMR 
            Caption         =   "VMR"
            Height          =   255
            Left            =   120
            TabIndex        =   22
            Top             =   240
            Value           =   -1  'True
            Width           =   855
         End
      End
      Begin VB.CommandButton Command 
         Caption         =   "�ɼ���������"
         Height          =   495
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton cmdPreview 
         Caption         =   "����Ԥ��ģʽ"
         Height          =   495
         Left            =   120
         TabIndex        =   19
         Top             =   840
         Width           =   1575
      End
      Begin VB.CommandButton cmdFree 
         Caption         =   "�˳���ƵԤ��"
         Height          =   495
         Left            =   120
         TabIndex        =   18
         Top             =   1440
         Width           =   1575
      End
      Begin VB.CommandButton cmdCapture 
         Caption         =   "�ɼ�ͼ��"
         Enabled         =   0   'False
         Height          =   495
         Left            =   1800
         TabIndex        =   17
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdCaptureVideo 
         Caption         =   "��ʼ¼��"
         Enabled         =   0   'False
         Height          =   495
         Left            =   1800
         TabIndex        =   16
         Top             =   1440
         Width           =   1815
      End
      Begin VB.CommandButton cmdStopCaptureVideo 
         Caption         =   "ֹͣ¼��"
         Enabled         =   0   'False
         Height          =   495
         Left            =   1800
         TabIndex        =   15
         Top             =   2040
         Width           =   1815
      End
      Begin VB.CheckBox ckShowState 
         Caption         =   "��ʾ��Ƶ״̬"
         Height          =   375
         Left            =   6600
         TabIndex        =   14
         Top             =   1080
         Value           =   1  'Checked
         Width           =   1455
      End
      Begin VB.CheckBox ckFullScreen 
         Caption         =   "ȫ����ʾ��Ƶͼ��(��ESC���˳�)"
         Height          =   375
         Left            =   6600
         TabIndex        =   13
         Top             =   720
         Width           =   3255
      End
      Begin VB.CommandButton cmdVideoCaptureFilterCfg 
         Caption         =   "VideoCaptureFilterCfg"
         Height          =   495
         Left            =   7680
         TabIndex        =   12
         Top             =   1680
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.CommandButton cmdVideoCapturePinCfg 
         Caption         =   "VideoCapturePinCfg"
         Height          =   495
         Left            =   8160
         TabIndex        =   11
         Top             =   1080
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.CheckBox chkAutoSize 
         Caption         =   "������ʾ��Ƶ����"
         Height          =   255
         Left            =   6600
         TabIndex        =   10
         Top             =   360
         Width           =   2775
      End
      Begin VB.Frame fraViewStyle 
         Caption         =   "��ʾ��ʽ"
         Height          =   1455
         Left            =   3960
         TabIndex        =   6
         Top             =   240
         Width           =   2535
         Begin VB.OptionButton optAutoFitCut 
            Caption         =   "��Ӧ�ü���Χ"
            Height          =   255
            Left            =   120
            TabIndex        =   25
            Top             =   960
            Width           =   2175
         End
         Begin VB.OptionButton optSourceState 
            Caption         =   "ԭʼ��С"
            Height          =   255
            Left            =   120
            TabIndex        =   9
            Top             =   720
            Value           =   -1  'True
            Width           =   2295
         End
         Begin VB.OptionButton optFit 
            Caption         =   "������������Ƶ��ʾͼ��"
            Height          =   255
            Left            =   120
            TabIndex        =   8
            Top             =   480
            Width           =   2295
         End
         Begin VB.OptionButton optStretch 
            Caption         =   "������Ƶ��ʾͼ��"
            Height          =   255
            Left            =   120
            TabIndex        =   7
            Top             =   240
            Width           =   2295
         End
      End
      Begin VB.Frame fraSaveFormat 
         Caption         =   "�ɼ�ͼ�󱣴��ʽ"
         Height          =   735
         Left            =   3960
         TabIndex        =   3
         Top             =   1800
         Width           =   3135
         Begin VB.OptionButton optBmp 
            Caption         =   "BMPλͼ��ʽ"
            Height          =   255
            Left            =   240
            TabIndex        =   5
            Top             =   360
            Value           =   -1  'True
            Width           =   1335
         End
         Begin VB.OptionButton optJpg 
            Caption         =   "JPGͼ���ʽ"
            Height          =   255
            Left            =   1680
            TabIndex        =   4
            Top             =   360
            Width           =   1335
         End
      End
      Begin VB.CommandButton cmdVideoSource 
         Caption         =   "��ƵԴѡ��"
         Height          =   495
         Left            =   120
         TabIndex        =   2
         Top             =   2040
         Width           =   1575
      End
   End
   Begin MSComDlg.CommonDialog diaSaveVideo 
      Left            =   10680
      Top             =   3120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "(*.AVI)|*.AVI|(*.*)|*.*"
   End
   Begin MSComDlg.CommonDialog dlgSaveImg 
      Left            =   10680
      Top             =   1800
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "(*.BMP)|*.BMP|(*.JPG)|*.JPG|(*.*)|*.*"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub chkAutoSize_Click()
  '�Զ���Ӧ��Ƶ���ڴ�С
  objDSCapture.IsAdjustWindowSize = chkAutoSize.Value

  Call Form_Resize
End Sub

Private Sub ckFullScreen_Click()
  'ȫ����ʾ
  If ckFullScreen.Value Then
    Call objDSCapture.ShowFullScreen(Me.hwnd, GetMonitorIndex(Me.hwnd))
  Else
    Call objDSCapture.QuitFullScreen
  End If
End Sub

Private Sub ckShowState_Click()
  objDSCapture.IsShowState = ckShowState.Value
End Sub


Private Sub cmdCapture_Click()
  '�ɼ�ͼ��
  Dim sErrMsg As String
  Dim sImgFile As String
  
  dlgSaveImg.FileName = ""
  dlgSaveImg.DefaultExt = "*.BMP"
  
  Call dlgSaveImg.ShowSave
  sImgFile = dlgSaveImg.FileName
  
  If Trim(sImgFile) = "" Then Exit Sub
    
  If optBmp.Value Then
    sErrMsg = objDSCapture.CaptureBmpImageToFile(sImgFile)
  Else
    sErrMsg = objDSCapture.CaptureJpgImageToFile(sImgFile, 100)
  End If
  
  
  
  If Trim(sErrMsg) <> "" Then
    Call MsgBox(sErrMsg)
    Exit Sub
  End If
  
  Call Form2.ShowCaptureImg(sImgFile)
  
End Sub

Private Sub cmdCaptureVideo_Click()
  '��ʼ¼��
  Dim sErrMsg As String
  Dim sVideoFile As String
  
  diaSaveVideo.FileName = ""
  diaSaveVideo.DefaultExt = "*.AVI"
  
  Call diaSaveVideo.ShowSave
  sVideoFile = diaSaveVideo.FileName
  
  If Trim(sVideoFile) = "" Then Exit Sub
  
  sErrMsg = objDSCapture.StartCaptureVideo(sVideoFile)
  If Trim(sErrMsg) <> "" Then
    Call MsgBox(sErrMsg)
  End If
End Sub

Private Sub cmdCustomParameter_Click()

  Call Form3.ShowCaptureParameterConfig(objDSCapture)
  
End Sub

Private Sub cmdFree_Click()
  'ֹͣ��ƵԤ��
  Dim sErrMsg As String
  
  sErrMsg = objDSCapture.StopPreview()
  If Trim(sErrMsg) <> "" Then
    Call MsgBox(sErrMsg)
  End If
  
  cmdCapture.Enabled = objDSCapture.PreviewState
  cmdStopCaptureVideo.Enabled = objDSCapture.PreviewState
  cmdCaptureVideo.Enabled = objDSCapture.PreviewState
End Sub

Private Sub cmdFromClipBoard_Click()
  '�ɼ�ͼ��
  Dim sErrMsg As String
  
  sErrMsg = objDSCapture.CaptureImgToClipBoard
  
  If Trim(sErrMsg) <> "" Then
    Call MsgBox(sErrMsg)
    Exit Sub
  End If
  
  Call Form2.ShowCaptureImgFromClipBoard
End Sub

Private Sub cmdPreview_Click()
  '��ʼ��ƵԤ��
  Dim sErrMsg As String
  
  sErrMsg = objDSCapture.StartPreview()
  If Trim(sErrMsg) <> "" Then
    Call MsgBox(sErrMsg)
  End If
  
  cmdCapture.Enabled = objDSCapture.PreviewState
  cmdStopCaptureVideo.Enabled = objDSCapture.PreviewState
  cmdCaptureVideo.Enabled = objDSCapture.PreviewState
End Sub

Private Sub cmdScreen_Click()
  
  Call MsgBox(GetMonitorIndex(Me.hwnd))
  
End Sub

Private Sub cmdStopCaptureVideo_Click()
  'ֹͣ¼��
  Dim sVideoFile As String
  Dim sErrMsg As String
  
  sErrMsg = objDSCapture.StopCaptureVideo(sVideoFile)
  If Trim(sErrMsg) <> "" Then
    Call MsgBox(sErrMsg)
  End If
End Sub

Private Sub cmdVideoCaptureFilterCfg_Click()
  '��Ƶ�ɼ�Դ����
  Dim sErrMsg As String
  
  sErrMsg = objDSCapture.ShowVideoCaptureFilterCfg(Me.hwnd)
  If Trim(sErrMsg) <> "" Then
    Call MsgBox(sErrMsg)
  End If
  
End Sub

Private Sub cmdVideoCapturePinCfg_Click()
  '��Ƶ�ɼ��˿�����
  Dim sErrMsg As String

  sErrMsg = objDSCapture.ShowVideoCapturePinCfg(Me.hwnd)
  If Trim(sErrMsg) <> "" Then
    Call MsgBox(sErrMsg)
  End If
End Sub

Private Sub cmdVideoSource_Click()
  '��Ƶ��Դ����
  objDSCapture.Left = 1
  Exit Sub
  Dim sErrMsg As String
  
  sErrMsg = objDSCapture.ShowVfwVideoSourceCfg(Me.hwnd)
  If Trim(sErrMsg) <> "" Then
    Call MsgBox(sErrMsg)
  End If
End Sub

Private Sub Command_Click()
  '�ɼ���������
  Dim sErrMsg As String
  
  sErrMsg = objDSCapture.ShowCaptureParameterCfgDialog(0)
  
  If Trim(sErrMsg) <> "" Then
    Call MsgBox(sErrMsg)
  End If
End Sub


Private Sub Command1_Click()
  '�ɼ�ͼ��
  Call Form2.ShowCaptureImgFromMemory(objDSCapture.CaptureBmpImage)
End Sub

Private Sub Command2_Click()
    Dim vSize As TVideoSize
    
    vSize = objDSCapture.GetRealVideoSize
    
    MsgBox vSize.Width & "X" & vSize.Height
End Sub

Private Sub objDSCapture_OnVideoSizeChange(ByVal videoWidth As Long, ByVal videoHieght As Long, ByVal windowWidth As Long, ByVal windowHeight As Long)
  Call Form_Resize
End Sub

Private Sub Form_Load()
  objDSCapture.AppHandle = Me.hwnd
  objDSCapture.IsDblClickQuitFullScreen = True


  Call objDSCapture.ReadParameterFromFile
  
  Call cmdPreview_Click
  
  Call Form_Resize
  
  clipformat = 0
  RegisterClipboardFormat "ZLDSVIDEOPROCESS10161"
  
  Me.AutoRedraw = True
  
'  ReDim Preserve monitor(1)
'  monitor(1) = -1
'
'  EnumDisplayMonitors ByVal 0&, ByVal 0&, AddressOf MonitorEnumProc, ByVal 0&
End Sub


Private Sub Form_Resize()

  fraControl.Left = 0
  fraControl.Top = Me.Height - fraControl.Height - 400
  fraControl.Width = Me.Width - 100

  objDSCapture.Top = 0
  objDSCapture.Left = 0

  If chkAutoSize.Value Then
    objDSCapture.Left = (Me.Width - ScaleX(objDSCapture.CurWidth, vbPixels, vbTwips)) / 2 - 50
  Else
    objDSCapture.CurHeight = ScaleY(Me.Height - fraControl.Height - 400, vbTwips, vbPixels)
    objDSCapture.CurWidth = ScaleX(Me.Width - 100, vbTwips, vbPixels) - 8
  End If
  
  'objDSCapture.Height = -1 ʹ��Height�������Ϊ������������󣬶�CurHeight�����������
  
  
  '��Ƶ��ʾ���ڴ�С�ı��ˢ�¿ؼ��ڲ�����ʾ����
  objDSCapture.RefreshWindow
  
  Me.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Call objDSCapture.SaveParameterToFile
  
  objDSCapture.FreeRes
End Sub


Private Sub optAutoFitCut_Click()
  
  objDSCapture.ShowModel = smAutoFitCut
  objDSCapture.RefreshWindow
End Sub




Private Sub optDevice_Click()
  If optVMR.Value Then
    objDSCapture.SnatchWay = swVMR
  Else
    objDSCapture.SnatchWay = swDEVICE
  End If
End Sub

Private Sub optFit_Click()
  objDSCapture.IsFit = optFit.Value
  objDSCapture.RefreshWindow
End Sub

Private Sub optSourceState_Click()
  objDSCapture.ShowModel = smNormal
  objDSCapture.RefreshWindow
End Sub

Private Sub optStretch_Click()
  objDSCapture.ShowModel = smStretch
  objDSCapture.RefreshWindow
End Sub

Private Sub optVMR_Click()
  If optVMR.Value Then
    objDSCapture.SnatchWay = swVMR
  Else
    objDSCapture.SnatchWay = swDEVICE
  End If
End Sub
