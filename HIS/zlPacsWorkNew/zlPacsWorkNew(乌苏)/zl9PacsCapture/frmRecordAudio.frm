VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{B1790453-7708-48C1-B5CC-75255FA4B066}#1.0#0"; "ZLDSVIDEOPROCESS.OCX"
Begin VB.Form frmRecordAudio 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "����¼��"
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
   StartUpPosition =   3  '����ȱʡ
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
      Caption         =   "¼�Ʋ�������"
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
         Caption         =   "Ӧ ��(&A)"
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
         Caption         =   "ȡ ��(&S)"
         Height          =   375
         Left            =   5640
         TabIndex        =   10
         Top             =   1800
         Width           =   1095
      End
      Begin VB.CommandButton cmdSure 
         Caption         =   "ȷ ��(&S)"
         Height          =   375
         Left            =   3240
         TabIndex        =   9
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "¼���豸ѡ��"
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
         Caption         =   "���������С��"
         Height          =   255
         Left            =   3960
         TabIndex        =   14
         Top             =   1365
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "��������������"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   1365
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "      ѹ���ʣ�"
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
            Text            =   "׼������..."
            TextSave        =   "׼������..."
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3528
            MinWidth        =   3528
            Text            =   "¼�Ƴ���:0(��)"
            TextSave        =   "¼�Ƴ���:0(��)"
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
      Caption         =   "   ��ͣ¼��(&P)"
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
      Caption         =   "   ��ʼ¼��(&S)"
      Height          =   400
      Left            =   2160
      TabIndex        =   1
      Top             =   2760
      Width           =   1575
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "   ֹͣ¼��(&T)"
      Enabled         =   0   'False
      Height          =   400
      Left            =   5520
      TabIndex        =   6
      Top             =   2760
      Width           =   1575
   End
   Begin VB.Menu menu_File 
      Caption         =   "�ļ�(&F)"
      Begin VB.Menu menu_SaveAs 
         Caption         =   "���Ϊ(&A)"
      End
      Begin VB.Menu menu_Split1 
         Caption         =   "-"
      End
      Begin VB.Menu menu_Import 
         Caption         =   "����(&I)"
      End
      Begin VB.Menu menu_Exit 
         Caption         =   "�˳�(&E)"
      End
   End
   Begin VB.Menu menu_Edit 
      Caption         =   "�༭(&E)"
      Begin VB.Menu menu_StartRecord 
         Caption         =   "��ʼ¼��(&S)"
      End
      Begin VB.Menu menu_PauseRecord 
         Caption         =   "��ͣ¼��(&P)"
      End
      Begin VB.Menu menu_StopRecord 
         Caption         =   "ֹͣ¼��(&T)"
      End
      Begin VB.Menu menu_Split2 
         Caption         =   "-"
      End
      Begin VB.Menu menu_ParConfig 
         Caption         =   "¼�Ʋ�������(&C)"
      End
      Begin VB.Menu menu_SoundFormatConfig 
         Caption         =   "������ʽ����(&U)"
      End
   End
   Begin VB.Menu menu_Help 
      Caption         =   "����(&H)"
   End
End
Attribute VB_Name = "frmRecordAudio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mstrRecordBufferFile As String  '���������ļ�
Private mobjCallBack As Object          '�ص��¼�����




Private Sub LoadRecordConfig()
On Error GoTo errHandle
    '����¼������
    Dim lngCompress As Long
    Dim i As Integer
    
    If cbxRecordDevice.ListCount > 0 Then cbxRecordDevice.ListIndex = GetSetting("ZLSOFT", G_STR_REG_PATH_PUBLIC, "¼���豸���", 0)
    
    If cbxBufferCount.ListCount > 0 Then cbxBufferCount.ListIndex = GetSetting("ZLSOFT", G_STR_REG_PATH_PUBLIC, "¼��������������", 3)
    
    txtBufferSize.Text = GetSetting("ZLSOFT", G_STR_REG_PATH_PUBLIC, "¼�������С", 1024)
    
    lngCompress = GetSetting("ZLSOFT", G_STR_REG_PATH_PUBLIC, "¼��ѹ������", 64)
    
    For i = 0 To cbxCompressRate.ListCount - 1
        If Val(cbxCompressRate.List(i)) = lngCompress Then
            cbxCompressRate.ListIndex = i
            Exit For
        End If
    Next i
    
    'Ӧ������
    Call cmdApply_Click
    
    Exit Sub
errHandle:
    MsgboxCus err.Description, vbOKOnly, G_STR_HINT_TITLE
End Sub

Private Sub SaveRecordConfig()
On Error Resume Next
    '����¼������
    
    Call SaveSetting("ZLSOFT", G_STR_REG_PATH_PUBLIC, "¼���豸���", Val(cbxRecordDevice.Text))
    Call SaveSetting("ZLSOFT", G_STR_REG_PATH_PUBLIC, "¼��������������", Val(cbxBufferCount.Text))
    Call SaveSetting("ZLSOFT", G_STR_REG_PATH_PUBLIC, "¼�������С", txtBufferSize.Text)
    Call SaveSetting("ZLSOFT", G_STR_REG_PATH_PUBLIC, "¼��ѹ������", Val(cbxCompressRate.Text))
End Sub


Private Sub LoadWavFormat()
On Error GoTo errHandle
    '����������ʽ
    mciSound.FormatTag = GetSetting("ZLSOFT", G_STR_REG_PATH_PUBLIC, "¼����ʽ����", 1)
    mciSound.SampleRate = GetSetting("ZLSOFT", G_STR_REG_PATH_PUBLIC, "¼������Ƶ��", 44100)
    mciSound.BitsPerSample = GetSetting("ZLSOFT", G_STR_REG_PATH_PUBLIC, "¼������λ������", 1)
    mciSound.Channels = GetSetting("ZLSOFT", G_STR_REG_PATH_PUBLIC, "¼��ͨ������", 1)
    
    Exit Sub
errHandle:
    MsgboxCus err.Description, vbOKOnly, G_STR_HINT_TITLE
End Sub

Private Sub SaveWavFormat()
On Error Resume Next
    '����������ʽ
    Call SaveSetting("ZLSOFT", G_STR_REG_PATH_PUBLIC, "¼������Ƶ��", mciSound.SampleRate)
    Call SaveSetting("ZLSOFT", G_STR_REG_PATH_PUBLIC, "¼������λ������", mciSound.BitsPerSample)
    Call SaveSetting("ZLSOFT", G_STR_REG_PATH_PUBLIC, "¼��ͨ������", mciSound.Channels)
    Call SaveSetting("ZLSOFT", G_STR_REG_PATH_PUBLIC, "¼����ʽ����", mciSound.FormatTag)
End Sub



Private Sub LoadRecordDevice()
    On Error Resume Next
    
    '����¼���豸
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
    '��ʾ����¼��
    Call Me.Show(0, objCallBack)
        
    Set mobjCallBack = objCallBack
End Sub



Private Sub cmdApply_Click()
On Error Resume Next
    'Ӧ��¼������
    
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
        '����¼�Ƶ���Ƶ
        Call mobjCallBack.subSaveAudio(mstrRecordBufferFile, lngTimeLen)
    End If
    
        Exit Sub
errHandle:
   Call MsgboxCus(err.Description, vbOKOnly, G_STR_HINT_TITLE)
   err.Clear
End Sub

Private Sub cmdSure_Click()
    'Ӧ������
    Call cmdApply_Click
    
    '�˳�����
    Call cmdCancel_Click
End Sub



Private Sub Form_Load()
    '�������ö�
    SetWindowPos Me.hWnd, -1, Me.CurrentX, Me.CurrentY, Me.ScaleWidth, Me.ScaleHeight, 3
    
    '�ָ���������
    Call zlCL_RestoreWinState(Me, App.ProductName)
    
    Me.Height = 4200
    
    
    '�ж������ʱĿ¼�����ڣ�����д���
    If Dir(IIf(Len(App.Path) > 3, App.Path & "\", App.Path & "") & "TmpImage", vbDirectory) = "" Then
        Call MkDir(IIf(Len(App.Path) > 3, App.Path & "\", App.Path & "") & "TmpImage")
    End If
    
    mstrRecordBufferFile = IIf(Len(App.Path) > 3, App.Path & "\", App.Path & "") & "TmpImage\WavBuffer.wav"
    mciSound.RecordFile = mstrRecordBufferFile
    mciSound.AppHandle = Me.hWnd
    
    Set mobjCallBack = Nothing
    
    '����¼���豸
    Call LoadRecordDevice
    
    '����Ĭ��ѹ������
    cbxCompressRate.ListIndex = 1
    
    '����Ĭ�ϻ�������
    cbxBufferCount.ListIndex = 3
    
    
    '����������ʽ
    Call LoadWavFormat
    
    '����¼������
    Call LoadRecordConfig
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If mciSound.RecordState <> adsStop Then
        Dim lngState As Long
        lngState = MsgboxCus("¼����δ�������Ƿ�ֹͣ¼�������棿", vbYesNoCancel, G_STR_HINT_TITLE)
        
        If lngState = vbYes Then
            '����¼��
            Call cmdStop_Click
        ElseIf lngState = vbCancel Then
            Cancel = 1
        End If
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    If mciSound.RecordState <> adsStop Then mciSound.StopRecord
    
    '���洰������
    Call zlCL_SaveWinState(Me, App.ProductName)
    
    
    '����¼������
    Call SaveRecordConfig
    
    '����������ʽ
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
            '����¼�Ƶ���Ƶ
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
        '������ʾ��
        Call MsgboxCus("δ�ҵ���Ҫ������Ƶ�ļ���", vbOKOnly Or vbInformation, G_STR_HINT_TITLE)
        
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
    '������ʽ����
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
    '���ð�ť״̬
    cmdStart.Enabled = mciSound.RecordState <> adsRun
    cmdPause.Enabled = mciSound.RecordState = adsRun
    cmdStop.Enabled = mciSound.RecordState = adsRun
    
    menu_StartRecord.Enabled = cmdStart.Enabled
    menu_StopRecord.Enabled = cmdStop.Enabled
    menu_PauseRecord.Enabled = cmdPause.Enabled
    
    '��ʾ��ǰ¼��״̬
    If mciSound.RecordState = adsRun Then
      StatusBar1.Panels(1).Text = "����¼��..."
    ElseIf mciSound.RecordState = adsPause Then
      StatusBar1.Panels(1).Text = "��ͣ��..."
    Else
      StatusBar1.Panels(1).Text = "׼������..."
    End If
    
    '��ʾ��¼�Ƴ���
    If mciSound.RecordState <> adsStop Then
      StatusBar1.Panels(2).Text = "¼�Ƴ���:" & mciSound.RecordTimeLen & "(��)"
    End If
End Sub
