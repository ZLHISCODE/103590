VERSION 5.00
Object = "{853AAF94-E49C-11D0-A303-0040C711066C}#4.3#0"; "DicomObjects.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "comct332.ocx"
Object = "{84926CA3-2941-101C-816F-0E6013114B7F}#1.0#0"; "IMGSCAN.OCX"
Begin VB.Form frmImgScan 
   BackColor       =   &H80000008&
   ClientHeight    =   7560
   ClientLeft      =   120
   ClientTop       =   45
   ClientWidth     =   9765
   ClipControls    =   0   'False
   Icon            =   "frmImgScan.frx":0000
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   PaletteMode     =   2  'Custom
   ScaleHeight     =   7560
   ScaleWidth      =   9765
   StartUpPosition =   1  'CenterOwner
   Begin ScanLibCtl.ImgScan ImgScan1 
      Left            =   4650
      Top             =   3540
      _Version        =   65536
      _ExtentX        =   873
      _ExtentY        =   873
      _StockProps     =   0
      PageType        =   6
      CompressionType =   6
      CompressionInfo =   4096
   End
   Begin DicomObjects.DicomViewer DViewer1 
      Height          =   2655
      Left            =   5040
      TabIndex        =   4
      Top             =   2880
      Visible         =   0   'False
      Width           =   2775
      _Version        =   262146
      _ExtentX        =   4895
      _ExtentY        =   4683
      _StockProps     =   35
   End
   Begin ComCtl3.CoolBar cbr 
      Align           =   1  'Align Top
      Height          =   390
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   9765
      _ExtentX        =   17224
      _ExtentY        =   688
      BandCount       =   2
      _CBWidth        =   9765
      _CBHeight       =   390
      _Version        =   "6.7.8988"
      Child1          =   "tbrMain"
      MinWidth1       =   4500
      MinHeight1      =   330
      NewRow1         =   0   'False
      Caption2        =   "存储设备"
      Child2          =   "cboDevice"
      MinHeight2      =   300
      Width2          =   2505
      NewRow2         =   0   'False
      Begin VB.ComboBox cboDevice 
         Height          =   315
         ItemData        =   "frmImgScan.frx":406A
         Left            =   8205
         List            =   "frmImgScan.frx":4077
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   45
         Width           =   1470
      End
      Begin MSComctlLib.Toolbar tbrMain 
         Height          =   330
         Left            =   165
         TabIndex        =   3
         Top             =   30
         Width           =   6975
         _ExtentX        =   12303
         _ExtentY        =   582
         ButtonWidth     =   1349
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Style           =   1
         TextAlignment   =   1
         ImageList       =   "imgGray"
         HotImageList    =   "imgColor"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   8
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "扫描"
               Key             =   "扫描"
               Object.ToolTipText     =   "开始胶片扫描"
               Object.Tag             =   "扫描"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "保存"
               Key             =   "保存"
               Object.ToolTipText     =   "保存扫描的画面"
               Object.Tag             =   "保存"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split1"
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "删除"
               Key             =   "删除"
               Object.ToolTipText     =   "删除当前选择的画面"
               Object.Tag             =   "删除"
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Caption         =   "清除"
               Key             =   "清除"
               Object.ToolTipText     =   "清除所有已扫描的画面"
               Object.Tag             =   "清除"
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split2"
               Style           =   3
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "帮助"
               Key             =   "帮助"
               Object.ToolTipText     =   "当前帮助主题"
               Object.Tag             =   "帮助"
               ImageKey        =   "帮助"
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "退出"
               Key             =   "退出"
               Object.ToolTipText     =   "退出"
               Object.Tag             =   "退出"
               ImageKey        =   "退出"
            EndProperty
         EndProperty
      End
   End
   Begin VB.PictureBox picView 
      Height          =   495
      Left            =   120
      ScaleHeight     =   435
      ScaleWidth      =   4155
      TabIndex        =   0
      Top             =   3960
      Width           =   4215
      Begin DicomObjects.DicomViewer DViewer 
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   0
         Width           =   2175
         _Version        =   262146
         _ExtentX        =   3836
         _ExtentY        =   661
         _StockProps     =   35
         BackColor       =   -2147483636
      End
   End
   Begin VB.Menu MenuFile 
      Caption         =   "&F文件"
      Visible         =   0   'False
      Begin VB.Menu LoadScreen 
         Caption         =   "&L 装入图象文件..."
         Shortcut        =   ^L
      End
      Begin VB.Menu SaveScreen 
         Caption         =   "&S 存屏幕图象文件..."
         Shortcut        =   ^S
      End
      Begin VB.Menu SaveBuffer 
         Caption         =   "&B 存缓存图象文件..."
      End
      Begin VB.Menu p30 
         Caption         =   "-"
      End
      Begin VB.Menu CopyToClipb 
         Caption         =   "&P 拷贝到粘贴板"
      End
      Begin VB.Menu p10 
         Caption         =   "-"
      End
      Begin VB.Menu SetToZero 
         Caption         =   "&C 图象清零 "
      End
      Begin VB.Menu SetToBand 
         Caption         =   "&D 置条带图象 "
      End
      Begin VB.Menu p19 
         Caption         =   "-"
      End
      Begin VB.Menu PrintPic 
         Caption         =   "&P 打印图像... "
      End
      Begin VB.Menu p12 
         Caption         =   "-"
      End
      Begin VB.Menu EXITOKDEMO 
         Caption         =   "&E 退出"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu MenuOption 
      Caption         =   "&O选项"
      Visible         =   0   'False
      Begin VB.Menu SetupVideo 
         Caption         =   "&S 设置参数... "
      End
      Begin VB.Menu CapSequence 
         Caption         =   "&C 序列采集..."
      End
      Begin VB.Menu p1 
         Caption         =   "-"
      End
      Begin VB.Menu CenterScreen 
         Caption         =   "&C 使采到屏幕中心 "
         Checked         =   -1  'True
      End
      Begin VB.Menu EnableMask 
         Caption         =   "&E 允许位屏蔽 "
      End
      Begin VB.Menu p33 
         Caption         =   "-"
      End
      Begin VB.Menu MakeXMirror 
         Caption         =   "&H 水平方向镜象 "
      End
      Begin VB.Menu MakeYMirror 
         Caption         =   "&V 垂直方向镜象 "
      End
      Begin VB.Menu p15 
         Caption         =   "-"
      End
      Begin VB.Menu MatchClient 
         Caption         =   "&M 匹配用户窗 "
      End
      Begin VB.Menu SetMatchSize 
         Caption         =   "&M 设置用户窗大小 "
      End
      Begin VB.Menu p16 
         Caption         =   "-"
      End
      Begin VB.Menu REPLAYBUFFER 
         Caption         =   "&B 回放缓存..."
      End
      Begin VB.Menu REPLAYMEMORY 
         Caption         =   "&M 回放内存..."
      End
      Begin VB.Menu REPLAYFILE 
         Caption         =   "&F 回放序列文件..."
      End
      Begin VB.Menu p3 
         Caption         =   "-"
      End
      Begin VB.Menu BUFFERTOSCREEN 
         Caption         =   "&U 缓存0传送到屏幕 "
      End
      Begin VB.Menu SCREENTOBUFFER 
         Caption         =   "&R 屏幕传送到缓存0 "
      End
      Begin VB.Menu BUFFER0TOBUFFER1 
         Caption         =   "&0 缓存的第0幅传送到第1幅 "
      End
      Begin VB.Menu BUFFER1TOBUFFER0 
         Caption         =   "&1 缓存的第1幅传送到第0幅 "
      End
      Begin VB.Menu p14 
         Caption         =   "-"
      End
      Begin VB.Menu BufferToFrame 
         Caption         =   "&U 缓存的第0幅传到帧存"
      End
      Begin VB.Menu FrameToBuffer 
         Caption         =   "&R 帧存传到缓存的第0幅"
      End
      Begin VB.Menu FrameToScreen 
         Caption         =   "&S 帧存传到屏幕"
      End
      Begin VB.Menu p13 
         Caption         =   "-"
      End
      Begin VB.Menu SELECTCARD 
         Caption         =   "&B 选用图象板..."
      End
   End
   Begin VB.Menu MenuCapture 
      Caption         =   "&C采集"
      Visible         =   0   'False
      Begin VB.Menu BACKTOSCREEN 
         Caption         =   "&E 使回送屏幕.."
      End
      Begin VB.Menu p4 
         Caption         =   "-"
      End
      Begin VB.Menu CAPTOBUFFER 
         Caption         =   "&B 序列采到缓存"
      End
      Begin VB.Menu LOOPTOBUFFER 
         Caption         =   "&L (循环)序列采到缓存"
      End
      Begin VB.Menu p24 
         Caption         =   "-"
      End
      Begin VB.Menu SeqCapToBuf 
         Caption         =   "&C 中断控制序列采到"
      End
      Begin VB.Menu p5 
         Caption         =   "-"
      End
      Begin VB.Menu CapTOMEMORY 
         Caption         =   "&M 序列采到内存"
      End
      Begin VB.Menu CAPTOSEQFILE 
         Caption         =   "&F 序列采到文件"
      End
      Begin VB.Menu p6 
         Caption         =   "-"
      End
      Begin VB.Menu CapToInDirect 
         Caption         =   "&I (经缓存)实时显"
      End
      Begin VB.Menu CapToDirect 
         Caption         =   "&D (待停）实时显示"
      End
      Begin VB.Menu CapToForever 
         Caption         =   "&E (恒久）实时显示"
      End
      Begin VB.Menu p2 
         Caption         =   "-"
      End
      Begin VB.Menu CONTTOBUFFER0 
         Caption         =   "&0 实时采到缓存第0幅"
      End
      Begin VB.Menu CONTTOBUFFER1 
         Caption         =   "&1 实时采到缓存第1幅"
      End
      Begin VB.Menu p11 
         Caption         =   "-"
      End
      Begin VB.Menu CapToFrame 
         Caption         =   "&V 实时采到帧存"
      End
      Begin VB.Menu p17 
         Caption         =   "-"
      End
      Begin VB.Menu MulChanCap 
         Caption         =   "&M 多通道分时实时显"
      End
      Begin VB.Menu MulChanCapSub 
         Caption         =   "&B 多通道分时分区实时显"
      End
      Begin VB.Menu p18 
         Caption         =   "-"
      End
      Begin VB.Menu AsyncMulCap 
         Caption         =   "&A 多卡采集分时送显"
      End
      Begin VB.Menu SyncMulCap 
         Caption         =   "&S 多卡分区实时显示"
      End
      Begin VB.Menu p22 
         Caption         =   "-"
      End
      Begin VB.Menu CaptureAudio 
         Caption         =   "&U 采集音频数据"
      End
   End
   Begin VB.Menu MenuDisp 
      Caption         =   "&D回显"
      Visible         =   0   'False
      Begin VB.Menu DISPFROMBUFFER 
         Caption         =   "&B 序列回显缓存"
      End
      Begin VB.Menu LOOPFROMBUFFER 
         Caption         =   "&L (循环)序列回显缓存"
      End
      Begin VB.Menu p7 
         Caption         =   "-"
      End
      Begin VB.Menu DISPFROMMEMORY 
         Caption         =   "&M (循环)序列回显内存"
      End
      Begin VB.Menu DISPFROMFILE 
         Caption         =   "&F (循环)序列回显文件"
      End
      Begin VB.Menu p8 
         Caption         =   "-"
      End
      Begin VB.Menu CAPTOMONITOR 
         Caption         =   "&V 显示视频输入"
      End
      Begin VB.Menu DISPFROMFRAME 
         Caption         =   "&R 连续回显帧存"
      End
      Begin VB.Menu p9 
         Caption         =   "-"
      End
      Begin VB.Menu NormalLut 
         Caption         =   "&N 正向输出显示"
      End
      Begin VB.Menu InverseLut 
         Caption         =   "&I 反向输出显示"
      End
      Begin VB.Menu AbsoluteLut 
         Caption         =   "&A 绝对值输出显示"
      End
   End
   Begin VB.Menu Freeze 
      Caption         =   "&P停止"
      Enabled         =   0   'False
      Visible         =   0   'False
   End
   Begin VB.Menu Active 
      Caption         =   "&A显示"
      Visible         =   0   'False
   End
   Begin VB.Menu SINGLECAPTO 
      Caption         =   "&S单帧采"
      Visible         =   0   'False
   End
   Begin VB.Menu mnuSend 
      Caption         =   "&P发送"
      Visible         =   0   'False
   End
   Begin VB.Menu MenuHelp 
      Caption         =   "&H帮助"
      Visible         =   0   'False
      Begin VB.Menu SysHelp 
         Caption         =   "系统帮助"
      End
      Begin VB.Menu CORR 
         Caption         =   "&H 使用帮助"
         Shortcut        =   {F1}
      End
      Begin VB.Menu SetAllocBuf 
         Caption         =   "&A 分配缓存"
      End
      Begin VB.Menu RegDevDriver 
         Caption         =   "&R 安装设备驱动"
      End
      Begin VB.Menu p21 
         Caption         =   "-"
      End
      Begin VB.Menu ABOUT 
         Caption         =   "&A 系统信息..."
         Shortcut        =   {F5}
      End
      Begin VB.Menu SIGNALEINFO 
         Caption         =   "&S 信号信息..."
         Shortcut        =   {F2}
      End
      Begin VB.Menu EXTTRIGGER 
         Caption         =   "&T 测试外触发"
      End
   End
End
Attribute VB_Name = "frmImgScan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private lScrnOffset As Long
Private iCurImageIndex As Long
Private strPatientID As String, strStudyUID As String, strImgType As String, strSeriesID As String
Private lngDeviceNO As String
Private aDevices() As Variant
Private mlngAdviceID As Long, mlngSendNO As Long

Private MultiImages As New DicomImages
Private strCachePath As String

Public Sub ShowMe(ByVal lngAdviceID As Long, ByVal lngSendNO As Long, Optional ByVal strType As String = "", _
    Optional ByVal strCheckUID As String = "")
    strPatientID = lngAdviceID: strStudyUID = "": strSeriesID = ""
    mlngAdviceID = lngAdviceID: mlngSendNO = lngSendNO
    strImgType = strType: strStudyUID = strCheckUID
'    lblInfo.Caption = GetPatientInfo(lngAdviceID, lngSendNO, lngPatientID, strStudyUID)
    Me.Show vbModal
End Sub


Private Sub exFreshWindow()
'刷新屏幕
End Sub
'
'Private Sub DViewer_DblClick()
'    If DViewer.Images.count = 0 Then Exit Sub
'    If Me.tbrMain.Buttons("录制").Value = tbrPressed Then Exit Sub
'
'    StopDisp
'    With DViewer1
'        .Images.Clear
'        .Images.Add DViewer.Images(iCurImageIndex)
'    End With
'End Sub

Private Sub cboDevice_Click()
    lngDeviceNO = aDevices(0, cboDevice.ListIndex)
End Sub

Private Sub Form_Activate()
    If Me.Tag = "" Then Exit Sub
    Me.Tag = ""
    
    InitPara
    GetAllImages DViewer, strStudyUID, strSeriesID, strCachePath, iCurImageIndex
End Sub

Private Sub InitPara()
    Dim rsTmp As New ADODB.Recordset
    
    lngDeviceNO = GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\影像扫描", "设备号", "0")
    On Error GoTo DBError
    
    gstrSQL = "Select 设备号,设备名 From 影像设备目录 Where 类型=1"
    OpenRecordset rsTmp, Me.Caption
    If rsTmp.EOF Then
        MsgBox "未定义影像存储设备，请到影像设备目录中设置！", vbInformation, gstrSysName
        Unload Me: Exit Sub
    End If
    
    aDevices = rsTmp.GetRows: rsTmp.MoveFirst
    lngDeviceNO = GetDefaultDev(aDevices, lngDeviceNO)
    
    Me.cboDevice.Clear
    Do While Not rsTmp.EOF
        cboDevice.AddItem Nvl(rsTmp(1))
        rsTmp.MoveNext
    Loop
    cboDevice.ListIndex = GetComboxIndex(aDevices, lngDeviceNO)
    Exit Sub
DBError:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function GetComboxIndex(aSource() As Variant, ByVal SeekString As String) As Long
    Dim i As Long
    
    For i = 0 To UBound(aSource, 2)
        If aSource(0, i) = SeekString Then Exit For
    Next
    If i > UBound(aSource, 2) Then i = 0
    GetComboxIndex = i
End Function

Private Sub Form_Load()
'初始化程序
    Dim i As Integer
    Dim l As Long
    Dim h As Long
    Dim bb(100) As Byte
    Dim total As Integer

    Dim objFileSystem As New Scripting.FileSystemObject
    Call RestoreWinState(Me, App.ProductName)
    
    iCurImageIndex = 0
    
    bActive = 0
    bMaskMode = 0
    total = 2
    iCurrUsedNo = -1
    iVirtCode = 0
    SQFILE = "ok.seq"
    iNumImage = NUMINFILE
    iNum = 2
    NoCapture = 2
    ratio = 25
    
    MaxBoard = 0
    
    strCachePath = App.Path & "\TmpImage\"
    If Not objFileSystem.FolderExists(strCachePath) Then objFileSystem.CreateFolder strCachePath
    
    Me.Tag = "Loading"
End Sub

Private Sub Form_Paint()
    exFreshWindow
End Sub

Private Sub Form_Resize()
    With picView
        .Left = 0
        .Top = 0
        .Width = Me.ScaleWidth
        .Height = Me.ScaleHeight - .Top
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveSetting("ZLSOFT", "公共模块\" & App.ProductName & "\影像扫描", "设备号", lngDeviceNO)
    Call SaveWinState(Me, App.ProductName)
End Sub

Private Sub picView_Resize()
    Dim iCols As Integer, iRows As Integer
    With DViewer
        .Left = 0: .Top = 0
        .Width = picView.ScaleWidth: .Height = picView.ScaleHeight
        
        If .Images.count > 0 Then
            ResizeRegion .Images.count, .Width, .Height, iRows, iCols
            .MultiColumns = iCols: .MultiRows = iRows
        End If
    End With
End Sub

Private Sub AddToDicomImages(ByVal strTmpFile As String)
    Dim iRows As Integer, iCols As Integer, objDicomImage As New DicomImage
    
    With DViewer
        objDicomImage.FileImport strTmpFile, "BMP"
        objDicomImage.PatientID = strPatientID
        '统一检查UID和序列UID
        If .Images.count > 0 Then
            objDicomImage.StudyUID = .Images(1).StudyUID
            objDicomImage.SeriesUID = .Images(1).SeriesUID
        ElseIf Len(strStudyUID) > 0 Then
            objDicomImage.StudyUID = strStudyUID
            If Len(strSeriesID) > 0 Then objDicomImage.SeriesUID = strSeriesID
        Else
            strStudyUID = objDicomImage.StudyUID
        End If
        
        .Images.Add objDicomImage: .CurrentIndex = 1
        
        If iCurImageIndex > 0 Then .Images(iCurImageIndex).BorderColour = vbWhite
        With .Images(.Images.count)
            .BorderStyle = 6: .BorderWidth = 1: .BorderColour = vbRed
        End With
        iCurImageIndex = .Images.count
    
        ResizeRegion .Images.count, .Width, .Height, iRows, iCols
        .MultiColumns = iCols: .MultiRows = iRows
    End With
End Sub

Private Sub DViewer_MouseDown(Button As Integer, Shift As Integer, x As Long, y As Long)
    Dim i As Integer
    If Button <> 1 Then Exit Sub
    
    With DViewer
        i = .ImageIndex(x, y)
        If i > 0 And i <= .Images.count And i <> iCurImageIndex Then
            .Images(iCurImageIndex).BorderColour = vbWhite
            .Images(i).BorderColour = vbRed
            iCurImageIndex = i
        End If
    End With
End Sub

Private Sub tbrMain_ButtonClick(ByVal Button As MSComctlLib.Button)
    On Error GoTo RunError
    Select Case Button.Key
        Case "扫描"
            CaptureImage
        Case "保存"
            SaveImages DViewer.Images, CStr(lngDeviceNO), strCachePath, , strImgType
        Case "删除"
            DeleteImage
        Case "清除"
            DeleteAllImages
        Case "帮助"
            ShowHelp App.ProductName, Me.hwnd, Me.Name
        Case "退出"
            Unload Me
    End Select
    Exit Sub
RunError:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub DeleteImage()
    Dim iCols As Integer, iRows As Integer
    If iCurImageIndex < 1 Then Exit Sub
    
    With DViewer
        .Images.Remove iCurImageIndex
        ResizeRegion .Images.count, .Width, .Height, iRows, iCols
        .MultiColumns = iCols: .MultiRows = iRows
        
        If iCurImageIndex > .Images.count Then iCurImageIndex = .Images.count
        If iCurImageIndex > 0 Then .Images(iCurImageIndex).BorderColour = vbRed
    End With
End Sub

Private Sub DeleteAllImages()
    Dim i As Long
    
    If DViewer.Images.count < 1 Then Exit Sub
    
    With DViewer
        For i = 1 To .Images.count
            .Images.Remove 1
        Next
        .MultiColumns = 1: .MultiRows = 1
        
        iCurImageIndex = 0
    End With
End Sub

Private Function GetDefaultDev(aSource() As Variant, ByVal lngDev As String) As String
    Dim i As Long
    
    For i = 0 To UBound(aSource, 2)
        If aSource(0, i) = lngDev Then Exit For
    Next
    If i > UBound(aSource, 2) Then i = 0
    GetDefaultDev = aSource(0, i)
End Function

Private Sub CaptureImage()
    Dim strTmpFile As String, objFile As New Scripting.FileSystemObject
    
    On Error GoTo CaptureError
    strTmpFile = strCachePath & objFile.GetTempName
    With ImgScan1
        .ScanTo = FileOnly
        .FileType = BMP_Bitmap
        .Image = strTmpFile
    
        .OpenScanner
        .StartScan
        .CloseScanner
    End With
    If objFile.FileExists(strTmpFile) Then
        AddToDicomImages strTmpFile
        objFile.DeleteFile strTmpFile
    End If
    
    Exit Sub
CaptureError:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub
