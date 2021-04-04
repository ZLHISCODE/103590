VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{B1790453-7708-48C1-B5CC-75255FA4B066}#1.0#0"; "ZLDSVIDEOPROCESS.OCX"
Begin VB.Form frmRecordAudio 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "声音录制"
   ClientHeight    =   6285
   ClientLeft      =   150
   ClientTop       =   780
   ClientWidth     =   7215
   Icon            =   "frmRecordAudio.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6285
   ScaleWidth      =   7215
   StartUpPosition =   3  '窗口缺省
   Begin ZLDSVideoProcess.TMCIAudio mciSound 
      Height          =   2535
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6975
      LineColor       =   65280
      MaxColor        =   255
      SampleCount     =   400
      DrawFrequency   =   500
      BackColor       =   0
      Enabled         =   -1  'True
      Object.Visible         =   -1  'True
      Channels        =   1
      BitsPerSample   =   1
      SampleRate      =   44100
      NoSamples       =   1024
      SplitChannels   =   0   'False
      TrigLevel       =   128
      Triggered       =   0   'False
      RecordFile      =   ""
      RecordCurTime   =   0
      RecordPostion   =   0
      AudioDeviceId   =   0
      Object.Width           =   465
      Object.Height          =   169
      Object.Left            =   8
      Object.Top             =   8
      Hint            =   ""
      ShowHint        =   0   'False
      DoubleBuffered  =   0   'False
      AppHandle       =   0
      Title           =   ""
      BufferCount     =   1
      FormatTag       =   1
      IsCompressWav   =   -1  'True
      CompRate        =   64
   End
   Begin MSComDlg.CommonDialog dlgSaveAs 
      Left            =   360
      Top             =   5640
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame famRecordConfig 
      Caption         =   "录制参数配置"
      Height          =   2295
      Left            =   120
      TabIndex        =   8
      Top             =   3480
      Visible         =   0   'False
      Width           =   6975
      Begin VB.ComboBox cbxBufferCount 
         Height          =   300
         ItemData        =   "frmRecordAudio.frx":179A
         Left            =   1560
         List            =   "frmRecordAudio.frx":17B6
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   1320
         Width           =   1815
      End
      Begin VB.CommandButton cmdApply 
         Caption         =   "应 用(&A)"
         Height          =   375
         Left            =   4440
         TabIndex        =   19
         Top             =   1800
         Width           =   1095
      End
      Begin VB.ComboBox cbxRecordDevice 
         Height          =   300
         ItemData        =   "frmRecordAudio.frx":1807
         Left            =   1560
         List            =   "frmRecordAudio.frx":1809
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   360
         Width           =   5175
      End
      Begin VB.TextBox txtBufferSize 
         Height          =   270
         Left            =   5160
         TabIndex        =   15
         Text            =   "1024"
         Top             =   1320
         Width           =   1215
      End
      Begin VB.ComboBox cbxCompressRate 
         Height          =   300
         ItemData        =   "frmRecordAudio.frx":180B
         Left            =   1560
         List            =   "frmRecordAudio.frx":1821
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   840
         Width           =   5175
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "取 消(&S)"
         Height          =   375
         Left            =   5640
         TabIndex        =   10
         Top             =   1800
         Width           =   1095
      End
      Begin VB.CommandButton cmdSure 
         Caption         =   "确 定(&S)"
         Height          =   375
         Left            =   3240
         TabIndex        =   9
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "录音设备选择："
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   405
         Width           =   1455
      End
      Begin VB.Label Label5 
         Caption         =   "B"
         Height          =   255
         Left            =   6480
         TabIndex        =   16
         Top             =   1365
         Width           =   255
      End
      Begin VB.Label Label4 
         Caption         =   "采样缓存大小："
         Height          =   255
         Left            =   3960
         TabIndex        =   14
         Top             =   1365
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "采样缓存数量："
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   1365
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "      压缩率："
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   885
         Width           =   1335
      End
   End
   Begin VB.Timer tmrState 
      Interval        =   1000
      Left            =   1080
      Top             =   5640
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   7
      Top             =   6030
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Text            =   "准备就绪..."
            TextSave        =   "准备就绪..."
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3528
            MinWidth        =   3528
            Text            =   "录制长度:0(秒)"
            TextSave        =   "录制长度:0(秒)"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   6562
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picStop 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   5640
      Picture         =   "frmRecordAudio.frx":187D
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   5
      Top             =   2835
      Width           =   255
   End
   Begin VB.PictureBox picPause 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   3960
      Picture         =   "frmRecordAudio.frx":1BBF
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   3
      Top             =   2835
      Width           =   255
   End
   Begin VB.CommandButton cmdPause 
      Caption         =   "   暂停录制(&P)"
      Enabled         =   0   'False
      Height          =   400
      Left            =   3840
      TabIndex        =   4
      Top             =   2760
      Width           =   1575
   End
   Begin VB.PictureBox picStart 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2280
      Picture         =   "frmRecordAudio.frx":1F01
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   2
      Top             =   2835
      Width           =   255
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "   开始录制(&S)"
      Height          =   400
      Left            =   2160
      TabIndex        =   1
      Top             =   2760
      Width           =   1575
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "   停止录制(&T)"
      Enabled         =   0   'False
      Height          =   400
      Left            =   5520
      TabIndex        =   6
      Top             =   2760
      Width           =   1575
   End
   Begin VB.Menu menu_File 
      Caption         =   "文件(&F)"
      Begin VB.Menu menu_SaveAs 
         Caption         =   "另存为(&A)"
      End
      Begin VB.Menu menu_Split1 
         Caption         =   "-"
      End
      Begin VB.Menu menu_Import 
         Caption         =   "导入(&I)"
      End
      Begin VB.Menu menu_Exit 
         Caption         =   "退出(&E)"
      End
   End
   Begin VB.Menu menu_Edit 
      Caption         =   "编辑(&E)"
      Begin VB.Menu menu_StartRecord 
         Caption         =   "开始录制(&S)"
      End
      Begin VB.Menu menu_PauseRecord 
         Caption         =   "暂停录制(&P)"
      End
      Begin VB.Menu menu_StopRecord 
         Caption         =   "停止录制(&T)"
      End
      Begin VB.Menu menu_Split2 
         Caption         =   "-"
      End
      Begin VB.Menu menu_ParConfig 
         Caption         =   "录制参数配置(&C)"
      End
      Begin VB.Menu menu_SoundFormatConfig 
         Caption         =   "声音格式配置(&U)"
      End
   End
   Begin VB.Menu menu_Help 
      Caption         =   "帮助(&H)"
   End
End
Attribute VB_Name = "frmRecordAudio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mstrRecordBufferFile As String  '声音缓存文件
Private mobjCallBack As Object          '回调事件对象




Private Sub LoadRecordConfig()
On Error GoTo errHandle
    '加载录音配置
    Dim lngCompress As Long
    Dim i As Integer
    
    If cbxRecordDevice.ListCount > 0 Then cbxRecordDevice.ListIndex = GetSetting("ZLSOFT", G_STR_REG_PATH_PUBLIC, "录音设备序号", 0)
    
    If cbxBufferCount.ListCount > 0 Then cbxBufferCount.ListIndex = GetSetting("ZLSOFT", G_STR_REG_PATH_PUBLIC, "录音缓存数量类型", 3)
    
    txtBufferSize.Text = GetSetting("ZLSOFT", G_STR_REG_PATH_PUBLIC, "录音缓存大小", 1024)
    
    lngCompress = GetSetting("ZLSOFT", G_STR_REG_PATH_PUBLIC, "录音压缩比率", 64)
    
    For i = 0 To cbxCompressRate.ListCount - 1
        If Val(cbxCompressRate.List(i)) = lngCompress Then
            cbxCompressRate.ListIndex = i
            Exit For
        End If
    Next i
    
    '应用配置
    Call cmdApply_Click
    
    Exit Sub
errHandle:
    MsgboxCus err.Description, vbOKOnly, G_STR_HINT_TITLE
End Sub

Private Sub SaveRecordConfig()
On Error Resume Next
    '保存录音配置
    
    Call SaveSetting("ZLSOFT", G_STR_REG_PATH_PUBLIC, "录音设备序号", Val(cbxRecordDevice.Text))
    Call SaveSetting("ZLSOFT", G_STR_REG_PATH_PUBLIC, "录音缓存数量类型", Val(cbxBufferCount.Text))
    Call SaveSetting("ZLSOFT", G_STR_REG_PATH_PUBLIC, "录音缓存大小", txtBufferSize.Text)
    Call SaveSetting("ZLSOFT", G_STR_REG_PATH_PUBLIC, "录音压缩比率", Val(cbxCompressRate.Text))
End Sub


Private Sub LoadWavFormat()
On Error GoTo errHandle
    '加载声音格式
    mciSound.FormatTag = GetSetting("ZLSOFT", G_STR_REG_PATH_PUBLIC, "录音格式类型", 1)
    mciSound.SampleRate = GetSetting("ZLSOFT", G_STR_REG_PATH_PUBLIC, "录音采样频率", 44100)
    mciSound.BitsPerSample = GetSetting("ZLSOFT", G_STR_REG_PATH_PUBLIC, "录音采样位数类型", 1)
    mciSound.Channels = GetSetting("ZLSOFT", G_STR_REG_PATH_PUBLIC, "录音通道类型", 1)
    
    Exit Sub
errHandle:
    MsgboxCus err.Description, vbOKOnly, G_STR_HINT_TITLE
End Sub

Private Sub SaveWavFormat()
On Error Resume Next
    '保存声音格式
    Call SaveSetting("ZLSOFT", G_STR_REG_PATH_PUBLIC, "录音采样频率", mciSound.SampleRate)
    Call SaveSetting("ZLSOFT", G_STR_REG_PATH_PUBLIC, "录音采样位数类型", mciSound.BitsPerSample)
    Call SaveSetting("ZLSOFT", G_STR_REG_PATH_PUBLIC, "录音通道类型", mciSound.Channels)
    Call SaveSetting("ZLSOFT", G_STR_REG_PATH_PUBLIC, "录音格式类型", mciSound.FormatTag)
End Sub



Private Sub LoadRecordDevice()
    On Error Resume Next
    
    '加载录音设备
    Dim i As Integer
    Dim lngRecordDeviceCount As Long
    
    
    Call cbxRecordDevice.Clear
    
    lngRecordDeviceCount = mciSound.RecordInputCount
    For i = 0 To lngRecordDeviceCount - 1
        cbxRecordDevice.AddItem (i & "-" & mciSound.RecordInputName(i))
    Next i
    
    If cbxRecordDevice.ListCount > 0 Then cbxRecordDevice.ListIndex = 0
End Sub





Public Sub ShowRecordAudio(Optional ByRef objCallBack As Object = Nothing)
    '显示声音录制
    Call Me.Show(0, objCallBack)
        
    Set mobjCallBack = objCallBack
End Sub



Private Sub cmdApply_Click()
On Error Resume Next
    '应用录音配置
    
    mciSound.AudioDeviceId = Val(cbxRecordDevice.Text)
    mciSound.CompRate = Val(cbxCompressRate.Text)
    mciSound.BufferCount = Val(cbxBufferCount.Text)
    mciSound.NoSamples = Val(txtBufferSize.Text)
End Sub

Private Sub cmdCancel_Click()
On Error Resume Next

    Me.Height = 4200
    famRecordConfig.Visible = False
End Sub

Private Sub cmdStart_Click()
On Error GoTo errHandle

    Call mciSound.StartRecord
    
    Exit Sub
errHandle:
    Call MsgboxCus(err.Description, vbOKOnly, G_STR_HINT_TITLE)
    err.Clear
End Sub

Private Sub cmdPause_Click()
On Error GoTo errHandle
    Call mciSound.PauseRecord
    
    Exit Sub
errHandle:
    Call MsgboxCus(err.Description, vbOKOnly, G_STR_HINT_TITLE)
    err.Clear
End Sub

Private Sub cmdStop_Click()
On Error GoTo errHandle
    Dim lngTimeLen As Long
    
    lngTimeLen = mciSound.RecordTimeLen
    Call mciSound.StopRecord
    
    If Dir(mstrRecordBufferFile) <> "" Then
        '保存录制的音频
        Call mobjCallBack.subSaveAudio(mstrRecordBufferFile, lngTimeLen)
    End If
    
        Exit Sub
errHandle:
   Call MsgboxCus(err.Description, vbOKOnly, G_STR_HINT_TITLE)
   err.Clear
End Sub

Private Sub cmdSure_Click()
    '应用配置
    Call cmdApply_Click
    
    '退出配置
    Call cmdCancel_Click
End Sub



Private Sub Form_Load()
    '将窗口置顶
    SetWindowPos Me.hWnd, -1, Me.CurrentX, Me.CurrentY, Me.ScaleWidth, Me.ScaleHeight, 3
    
    '恢复窗体设置
    Call zlCL_RestoreWinState(Me, App.ProductName)
    
    Me.Height = 4200
    
    
    '判断如果临时目录不存在，则进行创建
    If Dir(IIf(Len(App.Path) > 3, App.Path & "\", App.Path & "") & "TmpImage", vbDirectory) = "" Then
        Call MkDir(IIf(Len(App.Path) > 3, App.Path & "\", App.Path & "") & "TmpImage")
    End If
    
    mstrRecordBufferFile = IIf(Len(App.Path) > 3, App.Path & "\", App.Path & "") & "TmpImage\WavBuffer.wav"
    mciSound.RecordFile = mstrRecordBufferFile
    mciSound.AppHandle = Me.hWnd
    
    Set mobjCallBack = Nothing
    
    '载入录音设备
    Call LoadRecordDevice
    
    '设置默认压缩比例
    cbxCompressRate.ListIndex = 1
    
    '设置默认缓存数量
    cbxBufferCount.ListIndex = 3
    
    
    '加载声音格式
    Call LoadWavFormat
    
    '加载录音配置
    Call LoadRecordConfig
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If mciSound.RecordState <> adsStop Then
        Dim lngState As Long
        lngState = MsgboxCus("录音尚未结束，是否停止录音并保存？", vbYesNoCancel, G_STR_HINT_TITLE)
        
        If lngState = vbYes Then
            '保存录音
            Call cmdStop_Click
        ElseIf lngState = vbCancel Then
            Cancel = 1
        End If
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    If mciSound.RecordState <> adsStop Then mciSound.StopRecord
    
    '保存窗体设置
    Call zlCL_SaveWinState(Me, App.ProductName)
    
    
    '保存录音配置
    Call SaveRecordConfig
    
    '保存声音格式
    Call SaveWavFormat
End Sub

Private Sub menu_Exit_Click()
    Call Unload(Me)
End Sub

Private Sub menu_Import_Click()
On Error GoTo errHandle
    
    dlgSaveAs.FileName = ""
    dlgSaveAs.Filter = "(*.wav)|*.wav|(*.mp3)|*.mp3|(*.*)|*.*"
    
    Call dlgSaveAs.ShowOpen
        
    If Dir(dlgSaveAs.FileName) <> "" And Dir(dlgSaveAs.FileName) <> "0" Then
        Call FileCopy(dlgSaveAs.FileName, mstrRecordBufferFile)


        If Dir(mstrRecordBufferFile) <> "" Then
            '保存录制的音频
            Call mobjCallBack.subSaveAudio(mstrRecordBufferFile, 0)
        End If
    End If
        
    Exit Sub
errHandle:
    MsgboxCus err.Description, vbOKOnly, G_STR_HINT_TITLE
End Sub

Private Sub menu_ParConfig_Click()
On Error Resume Next
    If famRecordConfig.Visible Then
        Call cmdCancel_Click
    Else
        famRecordConfig.Visible = True
        Me.Height = 7005
    End If
End Sub

Private Sub menu_PauseRecord_Click()
    Call cmdPause_Click
End Sub

Private Sub menu_SaveAs_Click()
On Error GoTo errHandle
    
    If Dir(mstrRecordBufferFile) = "" Then
        '弹出提示框
        Call MsgboxCus("未找到需要另存的音频文件。", vbOKOnly Or vbInformation, G_STR_HINT_TITLE)
        
        Exit Sub
    End If
    
    dlgSaveAs.FileName = ""
    dlgSaveAs.Filter = "(*.wav)|*.wav|(*.mp3)|*.mp3|(*.*)|*.*"
    Call dlgSaveAs.ShowSave
    
    If Trim(dlgSaveAs.FileName) <> "" Then
        Call FileCopy(mstrRecordBufferFile, dlgSaveAs.FileName)
    End If
    
    Exit Sub
errHandle:
    MsgboxCus err.Description, vbOKOnly, G_STR_HINT_TITLE
End Sub

Private Sub menu_SoundFormatConfig_Click()
On Error Resume Next
    '声音格式配置
    mciSound.ShowFormatDialog
End Sub

Private Sub menu_StartRecord_Click()
    Call cmdStart_Click
End Sub

Private Sub menu_StopRecord_Click()
    Call cmdStop_Click
End Sub

Private Sub tmrState_Timer()
On Error Resume Next
    '设置按钮状态
    cmdStart.Enabled = mciSound.RecordState <> adsRun
    cmdPause.Enabled = mciSound.RecordState = adsRun
    cmdStop.Enabled = mciSound.RecordState = adsRun
    
    menu_StartRecord.Enabled = cmdStart.Enabled
    menu_StopRecord.Enabled = cmdStop.Enabled
    menu_PauseRecord.Enabled = cmdPause.Enabled
    
    '显示当前录制状态
    If mciSound.RecordState = adsRun Then
      StatusBar1.Panels(1).Text = "正在录音..."
    ElseIf mciSound.RecordState = adsPause Then
      StatusBar1.Panels(1).Text = "暂停中..."
    Else
      StatusBar1.Panels(1).Text = "准备就绪..."
    End If
    
    '显示已录制长度
    If mciSound.RecordState <> adsStop Then
      StatusBar1.Panels(2).Text = "录制长度:" & mciSound.RecordTimeLen & "(秒)"
    End If
End Sub
