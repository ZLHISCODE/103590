VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{853AAF94-E49C-11D0-A303-0040C711066C}#4.3#0"; "DicomObjects.ocx"
Object = "{84926CA3-2941-101C-816F-0E6013114B7F}#1.0#0"; "IMGSCAN.OCX"
Begin VB.Form frmPetitionCapture 
   Caption         =   "申请单图像"
   ClientHeight    =   8415
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11940
   Icon            =   "frmPetitionCapture.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8415
   ScaleWidth      =   11940
   StartUpPosition =   3  '窗口缺省
   Begin ScanLibCtl.ImgScan ImageScanner 
      Left            =   0
      Top             =   6240
      _Version        =   65536
      _ExtentX        =   661
      _ExtentY        =   661
      _StockProps     =   0
      StopScanBox     =   -1  'True
      FileType        =   3
      CompressionType =   0
      CompressionInfo =   0
      ScanTo          =   4
   End
   Begin VB.Frame fmeDcmViewer 
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Height          =   6735
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   10695
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   480
         Top             =   6240
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin DicomObjects.DicomViewer dcmMiniature 
         Height          =   4575
         Left            =   1800
         TabIndex        =   8
         ToolTipText     =   "双击选中图像，对图像进行操作。"
         Top             =   120
         Width           =   7530
         _Version        =   262147
         _ExtentX        =   13282
         _ExtentY        =   8070
         _StockProps     =   35
         BackColor       =   -2147483642
      End
      Begin DicomObjects.DicomViewer dcmViewImg 
         Height          =   1575
         Left            =   120
         TabIndex        =   15
         Top             =   3120
         Visible         =   0   'False
         Width           =   2175
         _Version        =   262147
         _ExtentX        =   3836
         _ExtentY        =   2778
         _StockProps     =   35
         BackColor       =   -2147483640
         UseScrollBars   =   0   'False
         UseMouseWheel   =   -1  'True
      End
      Begin DicomObjects.DicomViewer dcmView 
         Height          =   1575
         Left            =   120
         TabIndex        =   14
         Top             =   1440
         Width           =   2175
         _Version        =   262147
         _ExtentX        =   3836
         _ExtentY        =   2778
         _StockProps     =   35
         BackColor       =   0
         UseScrollBars   =   0   'False
      End
      Begin VB.PictureBox picTemp2 
         BackColor       =   &H80000007&
         BorderStyle     =   0  'None
         Height          =   1215
         Left            =   120
         ScaleHeight     =   1215
         ScaleWidth      =   1695
         TabIndex        =   13
         Top             =   120
         Width           =   1695
      End
   End
   Begin VB.Frame fmeInfoCtrl 
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Height          =   1530
      Left            =   0
      TabIndex        =   0
      Top             =   6870
      Width           =   11895
      Begin VB.Frame FmeProcess 
         Height          =   1455
         Left            =   7200
         TabIndex        =   16
         Top             =   0
         Width           =   1215
         Begin VB.CommandButton cmdReduce 
            Caption         =   "缩小"
            Height          =   330
            Left            =   180
            TabIndex        =   22
            Top             =   670
            Width           =   855
         End
         Begin VB.CommandButton cmdMagnify 
            Caption         =   "放大"
            Height          =   330
            Left            =   180
            TabIndex        =   21
            Top             =   240
            Width           =   855
         End
         Begin VB.CommandButton cmdReset 
            Caption         =   "复位"
            Height          =   330
            Left            =   180
            TabIndex        =   17
            Top             =   1100
            Width           =   855
         End
      End
      Begin VB.Frame FmeScanbtn 
         Height          =   1455
         Left            =   8520
         TabIndex        =   6
         Top             =   0
         Width           =   3255
         Begin VB.CommandButton cmdChoiceEqu 
            Caption         =   "选择设备(&C)"
            Height          =   375
            Left            =   1700
            TabIndex        =   12
            Top             =   840
            Width           =   1335
         End
         Begin VB.CommandButton cmdScanSet 
            Caption         =   "扫描设置(&E)"
            Height          =   375
            Left            =   240
            TabIndex        =   11
            Top             =   840
            Width           =   1335
         End
         Begin VB.CommandButton cmdDeleteImg 
            Caption         =   "删除图像(&D)"
            Height          =   375
            Left            =   1700
            TabIndex        =   10
            Top             =   360
            Width           =   1335
         End
         Begin VB.CommandButton cmdScanImg 
            Caption         =   "扫描图像(&S)"
            Height          =   375
            Left            =   240
            TabIndex        =   9
            Top             =   360
            Width           =   1335
         End
      End
      Begin VB.Frame fmePatientInfo 
         Height          =   1455
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   6975
         Begin VB.Label lblCheckNum 
            Caption         =   "检 查 号:"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   480
            TabIndex        =   20
            Top             =   705
            Width           =   2535
         End
         Begin VB.Label lblPatientAge 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "年    龄:"
            Height          =   180
            Left            =   5040
            TabIndex        =   19
            Top             =   240
            Width           =   1530
         End
         Begin VB.Label lblPatientDept 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "病人科室:"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   180
            Left            =   480
            TabIndex        =   18
            Top             =   1170
            Width           =   2565
         End
         Begin VB.Label lblPatientName 
            Caption         =   "姓    名:"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   480
            TabIndex        =   5
            Top             =   240
            Width           =   2535
         End
         Begin VB.Label lblExamineMethod 
            Caption         =   "检查方法:"
            Height          =   240
            Left            =   3120
            TabIndex        =   4
            Top             =   705
            Width           =   3765
         End
         Begin VB.Label lblSpePosition 
            Caption         =   "检查部位:"
            Height          =   240
            Left            =   3120
            TabIndex        =   3
            Top             =   1140
            Width           =   3735
         End
         Begin VB.Label lblPatientSex 
            Caption         =   "性    别:"
            Height          =   255
            Left            =   3120
            TabIndex        =   2
            Top             =   240
            Width           =   1815
         End
      End
   End
End
Attribute VB_Name = "frmPetitionCapture"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'点坐标类型
Private Type TPoint
  X As Integer
  Y As Integer
End Type

Private mstrTempDirOfScan As String          '扫描的临时目录
Private mstrScanDeviceTempDir As String      '扫描设备临时目录
Private mstrBufferDir As String

Private mintScanImageIndex As Integer        '扫描图像索引
Private mintCurImgIndex As Integer           '当前被选中的图象索引
Private mintShowPhotoNumber As Integer       '缩略图的显示数量


Private Const CONST_STR_DEFAULT_SCAN_FILE_TEMPLATE As String = "Scan"  '默认扫描文件临时存储路径
Private Const CONST_STR_DEFAULT_TEMP_SCAN_DIR_NAME As String = "\TempScan"  '默认扫描文件临时存储路径

Private mlngAdviceID As Long           '医嘱ID
Private mlngCurDeptId As Long          '当前科室ID
Private mstrPrivs As String            '当前权限

Private mstrSaveDeviceID As String      '存储设备的设备号
Private miNet As New clsFtp             'FTP类
Private mFtpUser As String              'FTP账号
Private mFtpPass As String              'FTP密码
Private mFtpDir As String               'FTP目录
Private mFtpIp As String               'FTP地址

Private mlngBaseX As Long               'dcmView中鼠标Down时的X坐标
Private mlngBaseY As Long               'dcmView中鼠标Down时的Y坐标
Private mMouseDownPoint As TPoint       '鼠标在DcmImage上按下时的位置
Private mblndcmViewImgDown As Boolean    '用于判断dcmView中鼠标是否被按下
Private mInitScrollPoint As TPoint      '开始拖动时的初始位置
Private mCorpSize As TPoint             '拖动后的相对偏移位置
Private mblnIsExamine As Boolean        '是否查看申请单
Public mblnIsLogin As Boolean           '是否是登录窗口的申请单按钮

'病人基本信息
Private mstrCheckNo As String           '检查号
Private mstrDeptName As String          '科室名称
Private mstrPatientName As String       '病人姓名
Private mstrPatientAge As String        '病人年龄
Private mstrPatientSex As String        '病人性别
Private mstrExamineMethod As String     '检查方法
Private mstrSpePosition As String       '标本部位

Public Sub ShowPetitionCaptureWind(ByVal strPrivs As String, lngCurDeptId As Long, strDeptName As String, _
                                   strPatientName As String, strPatientAge As String, strPatientSex As String, _
                                   strExamineMethod As String, strSpePosition As String, blnIsExamine As Boolean, _
                                   blnIsLogin As Boolean, Optional lngAdviceID As Long = 0, Optional intState As Integer = 0)
Dim rsTemp As ADODB.Recordset
Dim strSql As String
Dim FTPconn As New clsFtp
On Error GoTo errH

    '设置模块变量
    mstrPrivs = strPrivs
    mlngAdviceID = lngAdviceID
    mblnIsExamine = IIf(intState = 0, blnIsExamine, True)
    mblnIsLogin = blnIsLogin
    mlngCurDeptId = lngCurDeptId
    
    '初始化科室级参数
    InitDeptPara mlngCurDeptId
    
    If FTPconn.FuncFtpConnect(mFtpIp, mFtpUser, mFtpPass) = 0 Then
        MsgBox "FTP不能正常连接，请检查网络设置。", vbInformation, gstrSysName
        Exit Sub
    End If
    
    '断开FTP测试连接
    FTPconn.FuncFtpDisConnect
    
    strSql = "select 检查号 from 影像检查记录  where 医嘱id=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "取得检查号", lngAdviceID)
    
    If rsTemp.RecordCount > 0 Then
        mstrCheckNo = Nvl(rsTemp!检查号)
    End If
    
    mstrDeptName = strDeptName
    mstrPatientName = strPatientName
    mstrPatientAge = strPatientAge
    mstrPatientSex = strPatientSex
    mstrExamineMethod = strExamineMethod
    mstrSpePosition = strSpePosition
    
    '初始化病人信息
    Call InitLables
     
    Call Me.Show(1)
    
    Exit Sub
errH:
    '断开FTP连接
    FTPconn.FuncFtpDisConnect
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmdChoiceEqu_Click()
On Error GoTo errHandle

    '打开选择扫描设备窗口
    ImageScanner.ShowSelectScanner

    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmdDeleteImg_Click()
On Error GoTo errHandle

    '删除方法
    If mblnIsLogin Then
        Call subDeleteDcmImage
    Else
        Call subDeleteFTPImage
    End If
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmdScanImg_Click()
On Error GoTo errHandle
    
    Call ScanImages
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmdScanSet_Click()
On Error GoTo errHandle
    '打开扫描设置窗口
    ImageScanner.ShowScanPreferences

    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub ScanImages()

  '删除程序中临时存储的图像目录
  On Error GoTo continue
    If Dir(mstrTempDirOfScan, vbDirectory) <> "" Then
      Call mdlDir.DeleteFolder(mstrTempDirOfScan)
    End If
continue:

  If Dir(mstrTempDirOfScan, vbDirectory) = "" Then
    Call MkDir(mstrTempDirOfScan)
  End If

  '删除twain设备临时存储的目录
  On Error GoTo continue1
    If Dir(mstrScanDeviceTempDir, vbDirectory) <> "" Then
      Call mdlDir.DeleteFolder(mstrScanDeviceTempDir)
    End If
continue1:

  If Dir(mstrScanDeviceTempDir, vbDirectory) = "" Then
    Call MkDir(mstrScanDeviceTempDir)
  End If

  mintScanImageIndex = 0

  '设置扫描后的文件数据类型
  ImageScanner.FileType = BMP_Bitmap
  ImageScanner.StopScanBox = True
  ImageScanner.ShowSetupBeforeScan = True
  ImageScanner.ScanTo = UseFileTemplateOnly
  '设置采集的模板文件
  ImageScanner.Image = mstrTempDirOfScan & "\" & CONST_STR_DEFAULT_SCAN_FILE_TEMPLATE


  If Not ImageScanner.ScannerAvailable Then
    ImageScanner.OpenScanner
  End If

  On Error GoTo errProcess
    Call ImageScanner.StartScan
    Call ImageScanner.StopScan
    Call ImageScanner.CloseScanner

    Exit Sub
errProcess:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub



Private Sub dcmMiniature_DblClick()
On Error GoTo errHandle
    If mintCurImgIndex = 0 Then
        MsgBoxD Me, "该病人没有已扫描的申请单。", vbInformation, gstrSysName
        Exit Sub
    End If
    
   '将选中的图像单独加载到dcmViewImg里去并隐藏自身
    Call LoadViewImg
    
    FmeProcess.Visible = True
    dcmMiniature.Visible = False
    dcmViewImg.Visible = True
    
    If dcmViewImg.Visible Then
        cmdChoiceEqu.Enabled = False
        cmdDeleteImg.Enabled = False
        cmdScanImg.Enabled = False
        cmdScanSet.Enabled = False
    End If
    
    Call Form_Resize

    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub LoadViewImg()
Dim ImgTmpImage As New DicomImage
    
    dcmViewImg.Images.Clear
    '将图像转存至PicBox
    Set picTemp2.Picture = dcmMiniature.Images.Item(mintCurImgIndex).Picture
    '将图像复制到剪贴板
    Call Clipboard.SetData(picTemp2.Picture, 2)
'    从剪切板取回图像
    Call ImgTmpImage.Paste
    
    Call Clipboard.Clear
    '将图像放入缩略图中
    dcmViewImg.Images.Add ImgTmpImage
End Sub

Private Sub dcmViewImg_DblClick()
On Error GoTo errHandle

    dcmMiniature.Visible = True
    dcmViewImg.Visible = False
    FmeProcess.Visible = False
    
     '缩放状态下 不能进行操作
    If dcmViewImg.Visible = False And Not mblnIsExamine Then
        cmdChoiceEqu.Enabled = True
        cmdDeleteImg.Enabled = True
        cmdScanImg.Enabled = True
        cmdScanSet.Enabled = True
    End If
    
     Call Form_Resize
     
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub



Private Sub dcmViewImg_MouseMove(Button As Integer, Shift As Integer, X As Long, Y As Long)
On Error Resume Next

    If mblndcmViewImgDown = True And Button = 1 And dcmViewImg.Images.Count > 0 Then
        dcmViewImg.Images(1).ScrollX = mInitScrollPoint.X - X
        dcmViewImg.Images(1).ScrollY = mInitScrollPoint.Y - Y

        dcmViewImg.Refresh
    End If
End Sub

Private Sub dcmViewImg_MouseDown(Button As Integer, Shift As Integer, X As Long, Y As Long)
On Error GoTo errHandle
    Dim intLabelType As Integer

    If Button = 1 And dcmViewImg.Images.Count > 0 Then
        mMouseDownPoint.X = dcmViewImg.Images(1).ActualScrollX
        mMouseDownPoint.Y = dcmViewImg.Images(1).ActualScrollY
          
        mInitScrollPoint.X = dcmViewImg.Images(1).ScrollX + X
        mInitScrollPoint.Y = dcmViewImg.Images(1).ScrollY + Y
        
        mblndcmViewImgDown = True
        
        '记录当前鼠标坐标
        mlngBaseX = X
        mlngBaseY = Y
    End If
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub dcmViewImg_MouseUp(Button As Integer, Shift As Integer, X As Long, Y As Long)
On Error GoTo errHandle

    If mblndcmViewImgDown = True And Button = 1 And dcmViewImg.Images.Count > 0 Then
        '计算图像漫游的偏移位置
        mCorpSize.X = mCorpSize.X + (dcmViewImg.Images(1).ActualScrollX - mMouseDownPoint.X)
        mCorpSize.Y = mCorpSize.Y + (dcmViewImg.Images(1).ActualScrollY - mMouseDownPoint.Y)
    End If
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub dcmViewImg_MouseWheel(ByVal Shift As Long, ByVal Delta As Integer, ByVal X As Long, ByVal Y As Long)
On Error GoTo errHandle
    '鼠标滚动事件 实现拖动
     Dim dblZoom As Double
    dblZoom = dcmViewImg.Images(1).ActualZoom
    dblZoom = dblZoom * (1 + Delta * 0.001)
    If dblZoom < 64 And dblZoom > 0.01 Then
        subCenterZoom dcmViewImg.Images(1), dcmViewImg, dblZoom, mCorpSize
    End If
    mlngBaseY = Y
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmdMagnify_Click()
On Error GoTo errHandle
'按钮放大
Dim dblZoom As Double

    dblZoom = dcmViewImg.Images(1).ActualZoom
    dblZoom = dblZoom * (1 + 300 * 0.001)
    If dblZoom < 64 And dblZoom > 0.01 Then
        subCenterZoom dcmViewImg.Images(1), dcmViewImg, dblZoom, mCorpSize
    End If
    'mlngBaseY = y
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub


Private Sub cmdReduce_Click()
On Error GoTo errHandle
'按钮缩小
    Dim dblZoom As Double
    
    dblZoom = dcmViewImg.Images(1).ActualZoom
    dblZoom = dblZoom * (1 + (-240) * 0.001)
    If dblZoom < 64 And dblZoom > 0.01 Then
        subCenterZoom dcmViewImg.Images(1), dcmViewImg, dblZoom, mCorpSize
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub


Private Sub cmdReset_Click()
On Error GoTo errHandle
'重置按钮以及图像
    
    '初始化拖动后的相对偏移位置
    mCorpSize.X = 0
    mCorpSize.Y = 0
    
    '重置图像
    Call LoadViewImg
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub


Private Sub subCenterZoom(img As DicomImage, Viewer As DicomViewer, dblZoom As Double, corpSize As TPoint)
'------------------------------------------------
'功能：对图像进行缩放。以当前viewer中心点为缩放中心点。
'参数： img -- 进行缩放的图像
'       viewer －－ 图像所在的viewer
'       dblZoom －－图像新的缩放倍数
'------------------------------------------------
    img.Zoom = dblZoom
    img.StretchToFit = False

    img.ScrollX = (img.SizeX * img.ActualZoom - ScaleX(Viewer.Width, vbTwips, vbPixels) / Viewer.MultiColumns) / 2 + corpSize.X
    img.ScrollY = (img.SizeY * img.ActualZoom - ScaleY(Viewer.Height, vbTwips, vbPixels) / Viewer.MultiRows) / 2 + corpSize.Y
End Sub





Private Sub Form_Load()
'窗体加载事件

Dim strFolder As String
On Error GoTo errHandle
    Call RestoreWinState(Me, App.ProductName)
    
    mstrTempDirOfScan = App.Path + CONST_STR_DEFAULT_TEMP_SCAN_DIR_NAME
    If Len(mstrTempDirOfScan) > 45 Then
        
        Dim pathlen As Long

        strFolder = String(256, 0)
        pathlen = GetTempPath(256, strFolder)
        If pathlen > 0 Then
            mstrTempDirOfScan = Left(strFolder, pathlen - 1) + CONST_STR_DEFAULT_TEMP_SCAN_DIR_NAME
        End If
    End If
    
    '根据参数判断 当前是查看申请单还是扫描申请单,如是查看则禁用四个操作按钮
    If mblnIsExamine Then
        cmdChoiceEqu.Enabled = False
        cmdDeleteImg.Enabled = False
        cmdScanImg.Enabled = False
        cmdScanSet.Enabled = False
    End If
    
    '初始化隐藏 图像高级处理按钮
    FmeProcess.Visible = False
    
    '设置设备临时目录
    mstrScanDeviceTempDir = "C:\Documents and Settings\All Users\Application Data\Microsoft\WIA"

    '如果程序在磁盘的根目录则app.path为“x:\”
    mstrBufferDir = IIf(Len(App.Path) > 3, App.Path & "\TmpImage\", App.Path & "TmpImage\")
    
    '将已有图像加载到DcmViewer控件中显示
    Call GetPetitionImages(dcmMiniature, mlngAdviceID, 100)

    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Public Sub Form_Unload(Cancel As Integer)
On Error GoTo errHandle
    Call SaveWinState(Me, App.ProductName)

    '关闭窗口时 断开当前FTP连接
    miNet.FuncFtpDisConnect
    
    '暂不关闭，只是隐藏窗口
    If mblnIsLogin Then
        Cancel = 1
        Me.Hide
    End If
    
    Exit Sub
errHandle:
    '断开FTP连接
    miNet.FuncFtpDisConnect
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub


Private Sub InitLables()
'根据传入的值给病人基本信息lbl赋值

    lblCheckNum.Caption = "检 查 号：" & mstrCheckNo
    lblPatientDept.Caption = "病人科室：" & mstrDeptName
    lblPatientName.Caption = "姓    名：" & mstrPatientName
    lblPatientAge.Caption = "年    龄：" & mstrPatientAge
    lblPatientSex.Caption = "性    别：" & mstrPatientSex
    lblExamineMethod.Caption = "检查方法：" & mstrExamineMethod
    lblSpePosition.Caption = "检查部位：" & mstrSpePosition

End Sub

Public Sub InitDeptPara(ByVal lngDeptID As Long)
'初始化科室级参数
    Dim rsTmp As New ADODB.Recordset
    On Error GoTo DBError
    mlngCurDeptId = lngDeptID
    
    '读取并检测存储设备号
    mstrSaveDeviceID = GetDeptPara(mlngCurDeptId, "存储设备号")
    gstrSQL = "Select 设备号,设备名 From 影像设备目录 Where 类型=1 and 设备号=[1] and NVL(状态,0)=1"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "得到设备号", mstrSaveDeviceID)
    If rsTmp.EOF Then
        MsgBox "影像存储设备未定义或处于停用，请检查！", vbInformation, gstrSysName
        mstrSaveDeviceID = ""
        Exit Sub
    End If
    
    Call funGetStorageDevice(Me, mstrSaveDeviceID, mFtpDir, mFtpIp, mFtpUser, mFtpPass)
    
    Exit Sub
DBError:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub imageScanner_PageDone(ByVal PageNumber As Long)
On Error GoTo errHandle
      Dim strScanFile As String

      If mintScanImageIndex = -1 Then
        Exit Sub
      End If
    
      '计算扫描文件索引
      mintScanImageIndex = mintScanImageIndex + 1
    
      
      strScanFile = mintScanImageIndex
    
      '取得有效的扫描文件名称
      While Len(strScanFile) < 4
        strScanFile = "0" + strScanFile
      Wend
    
      strScanFile = mstrTempDirOfScan & "\" & CONST_STR_DEFAULT_SCAN_FILE_TEMPLATE & strScanFile & ".bmp"
    
      '保存扫描的图像
      Call subCaptureImg(True, strScanFile)

    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Public Sub subCaptureImg(ByVal RealTimeCap As Boolean, Optional ByVal strFileName As String = "", _
    Optional ByRef picCapture As StdPicture = Nothing, Optional ByVal blnIsAfterCapture As Boolean = False)
'------------------------------------------------
'功能: 扫描并存储图像
'参数：无
'返回：无，直接保存新采集的图像
'------------------------------------------------
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset

    If mblnIsLogin Then
        If funCaptureSingleImage(RealTimeCap, strFileName, picCapture) = False Then
            MsgBoxD Me, "图像加载失败。", vbInformation, gstrSysName
            Exit Sub
        End If
    Else
        If funCaptureSingleImage(RealTimeCap, strFileName, picCapture) = True Then
            '调用保存图像到服务器 方法
            Call subSaveImage(, mlngAdviceID)
        End If
    End If
    
End Sub




Private Function funCaptureSingleImage(ByVal RealTimeCap As Boolean, _
    Optional ByVal strFileName As String = "", Optional ByRef picCapture As StdPicture = Nothing) As Boolean
'------------------------------------------------
'功能：将图像放入缩略图dcmMiniature中。
'参数：无
'返回：无，直接将新采集的图像放入dcmMiniature中
'------------------------------------------------

    Dim ImgTmpImage As New DicomImage

    On Error GoTo SaveFileError

    '扫描图像
    Call Clipboard.Clear

    If Not (picCapture Is Nothing) Then
        Set picTemp2.Picture = Nothing
        Set picTemp2.Picture = picCapture

    ElseIf Trim(strFileName) <> "" And Dir(strFileName) <> "" Then
        '使用dcmView显示的是图片，不需要再裁剪
        Set picTemp2.Picture = Nothing
        Set picTemp2.Picture = LoadPicture(strFileName)

    Else
        Set picTemp2.Picture = Nothing

        If dcmView.Images.Count > 0 Then
            Set picTemp2.Picture = dcmView.CurrentImage.Capture(False).Picture
        End If
    End If

    '将图像再次提交到剪切板
    If picTemp2.Picture Is Nothing Then
        funCaptureSingleImage = False
        Exit Function
    End If


    Call Clipboard.SetData(picTemp2.Picture, 2)
'    从剪切板取回图像
    Call ImgTmpImage.Paste

    Call Clipboard.Clear

    '将图像放入缩略图中
    Call subInsert2Mini(ImgTmpImage)

    funCaptureSingleImage = True

    Exit Function
SaveFileError:
    If ErrCenter() = 1 Then
        Resume
    End If

    Call SaveErrLog
End Function

Private Sub subInsert2Mini(img As DicomImage)
'------------------------------------------------
'功能：将图像添加到缩略图dcmMiniature中
'参数：img－－输入的图像
'返回：无，直接将图像添加到缩略图dcmMiniature中
'------------------------------------------------
    Dim iRows As Integer
    Dim iCols As Integer

    '计算缩略图的图像布局

    ResizeRegion dcmMiniature.Images.Count + 1, dcmMiniature.Width, dcmMiniature.Height, iRows, iCols

    dcmMiniature.MultiColumns = iCols
    dcmMiniature.MultiRows = iRows

    dcmMiniature.Images.Add img

    '处理缩略图中被选中的状态
    If mintCurImgIndex > 0 And mintCurImgIndex <= dcmMiniature.Images.Count Then
        dcmMiniature.Images(mintCurImgIndex).BorderColour = vbWhite
    End If


    With dcmMiniature.Images(dcmMiniature.Images.Count)
        .BorderWidth = 1
        .BorderStyle = 6
        .BorderColour = vbRed
    End With

    mintCurImgIndex = dcmMiniature.Images.Count
End Sub


Public Sub subSaveImage(Optional iEncode As Integer = 0, Optional lngAdviceID As Long)
'------------------------------------------------
'功能：将最后一个缩略图保存到数据库中
'参数：iEncode－－压缩方式，1－Run-Length Encoding行程压缩；2－不处理，保持原图的压缩方式，其他－Lossless JPEG encoding JPEG无损压缩
'返回：无
'------------------------------------------------
    Dim rsTmp As New ADODB.Recordset
    Dim ImgTmp As New DicomImage

    Dim strReceived As String
    Dim strFileTitle As String       '图像文件开头
    Dim strResult As String         'FTP操作结果
    Dim nowTime As Date
    Dim blnInTrans As Boolean       '在事物处理过程中
    Dim strRandom As String
    Dim i As Integer

    
    If dcmMiniature.Images.Count <= 0 Then
        Exit Sub
    End If
    
    '读取最后一个缩略图
    Set ImgTmp = dcmMiniature.Images(dcmMiniature.Images.Count)

    '得到随机数
    strRandom = CInt(Rnd * 100 + 1)

    nowTime = zlDatabase.Currentdate
    strFileTitle = Format(nowTime, "mmddhhmmss")
    strReceived = Format(nowTime, "yyyymmdd")
    
    '创建缓冲目录
    MkLocalDir mstrBufferDir & strReceived & "/" & lngAdviceID & "/"

    '保存图像到缓存目录  Lossless JPEG encoding JPEG无损压缩
    ImgTmp.WriteFile mstrBufferDir & strReceived & "/" & lngAdviceID & "/" & strFileTitle & lngAdviceID & strRandom, True

    ImgTmp.FileExport mstrBufferDir & strReceived & "/" & lngAdviceID & "/" & strFileTitle & lngAdviceID & strRandom & ".jpg", "JPG", 80

    ImgTmp.Tag = strFileTitle & lngAdviceID & strRandom & ".jpg"
    
    strResult = miNet.FuncFtpConnect(mFtpIp, mFtpUser, mFtpPass)

    If strResult = 0 Then
        'FTP操作失败，提示错误，并删除缩略图中的图像
        MsgBoxD Me, "FTP连接失败，图像无法保存，请检查网络设置。", vbInformation, gstrSysName
        dcmMiniature.Images.Remove (i)
        Exit Sub
    End If

    '保存扫描单图像
    WriteToURL mstrBufferDir & strReceived & "/" & lngAdviceID & "/" & strFileTitle & lngAdviceID & strRandom, mFtpDir & _
        strReceived & "/" & lngAdviceID & "/" & strFileTitle & lngAdviceID & strRandom

    miNet.FuncFtpDisConnect

    '图像存储成功后，存储数据库信息
    On Error GoTo DBError
    
    '插入新的图像记录
    gstrSQL = "ZL_影像申请单图像_INSERT ('" & lngAdviceID & "','" & strFileTitle & lngAdviceID & strRandom & ".jpg" & "','" & strReceived & "/" & lngAdviceID & "','" & mstrSaveDeviceID & "','" & UserInfo.姓名 & "',sysdate)"

    '保存图像
    Call zlDatabase.ExecuteProcedure(CStr(gstrSQL), "保存图像")
    
    '如果mblnIsLogin=true 那么说明这是在登记界面的保存图像，需要将该参数设置为false
    If mblnIsLogin Then
        mblnIsLogin = False
    End If

    Exit Sub
DBError:
    '断开FTP连接
    miNet.FuncFtpDisConnect
    '出错，则回退数据库操作，并且删除所采集的图像
    err.Raise err.Number, "检查图像保存"
    dcmMiniature.Images.Remove (dcmMiniature.Images.Count)
End Sub

Public Sub GetPetitionImages(dcmViewer As DicomViewer, lngAdviceID As Long, _
Optional intGetImgNum As Integer = 0, Optional intShowImgNum As Integer = 0)
'------------------------------------------------
'功能：删除dcmViewer中的图像后，将读取的图像文件放入dcmViewer中
'参数： dcmViewer－－打开图像的DicomViewer
'       lngAdviceID －－ 医嘱ID
'       intGetImgNum －－本次读取的图像数量
'       intShowImgNum －－本次显示的图像数量
'返回：无，直接修改dcmViewer中显示的图像
'------------------------------------------------

    Dim strSql As String
    Dim rsTmp As New ADODB.Recordset
    Dim curImage As DicomImage
    Dim iCols As Integer, iRows As Integer
    Dim objFile As New Scripting.FileSystemObject, strTmpFile As String
    Dim Inet1 As New clsFtp
    Dim strFTPUser As String              'FTP账号
    Dim strFtpPass As String              'FTP密码
    Dim strFtpDir As String               'FTP目录
    Dim strFtpIp As String                'FTP地址
    Dim strFirstDevNo As String
    Dim strNextDevNo As String
    Dim strTmpFolder As String
    
    On Error GoTo DBError

       strSql = "select 申请单图像,扫描人,扫描时间,FTP路径,设备号 from 影像申请单图像 where 医嘱ID=[1] order by 设备号"
       Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "提取申请单图像信息", lngAdviceID)

        'DcmViewer.Images.Clear
        If rsTmp.RecordCount > 0 Then
            ResizeRegion rsTmp.RecordCount, dcmViewer.Width, dcmViewer.Height, iRows, iCols

            dcmViewer.MultiColumns = iCols
            dcmViewer.MultiRows = iRows
            
            mstrBufferDir = IIf(Len(App.Path) > 3, App.Path & "\TmpImage\", App.Path & "TmpImage\")

            strFirstDevNo = Nvl(rsTmp("设备号"))
   
            Do While Not rsTmp.EOF
                strTmpFolder = mstrBufferDir & objFile.GetParentFolderName(Nvl(rsTmp("FTP路径")) & "/" & Mid(Nvl(rsTmp("申请单图像")), 1, InStr(Nvl(rsTmp("申请单图像")), ".") - 1))
                '创建本地目录
                If Not objFile.FolderExists(strTmpFolder) Then MkLocalDir (strTmpFolder)
            
                If strFirstDevNo <> strNextDevNo Then
                    Call funGetStorageDevice(Me, Nvl(rsTmp("设备号")), strFtpDir, strFtpIp, strFTPUser, strFtpPass)
                    
                    '判断FTP是否连接成功
                    If Inet1.FuncFtpConnect(strFtpIp, strFTPUser, strFtpPass) = 0 Then
                        MsgBoxD Me, "FTP不能正常连接，请检查网络设置。"
                        Exit Sub
                    End If
                End If
                
                strTmpFile = mstrBufferDir & Nvl(rsTmp("FTP路径")) & "/" & Mid(Nvl(rsTmp("申请单图像")), 1, InStr(Nvl(rsTmp("申请单图像")), ".") - 1)

                If Dir(strTmpFile) = vbNullString Then
                    '本地缓存图像不存在，则读取FTP图像

                    If Inet1.FuncDownloadFile(objFile.GetParentFolderName(strFtpDir & Nvl(rsTmp("FTP路径")) & "/" & Mid(Nvl(rsTmp("申请单图像")), 1, InStr(Nvl(rsTmp("申请单图像")), ".") - 1)), strTmpFile, objFile.GetFileName(Nvl(rsTmp("FTP路径")) & "/" & Mid(Nvl(rsTmp("申请单图像")), 1, InStr(Nvl(rsTmp("申请单图像")), ".") - 1))) <> 0 Then
                        '下载图像失败
                        MsgBoxD Me, "下载过程遇到未知错误，请联系系统管理员。"
                        Exit Sub
                    End If
                End If

                If Dir(strTmpFile) <> vbNullString Then
                        
                        Set curImage = dcmViewer.Images.ReadFile(strTmpFile)
                        curImage.Tag = Nvl(rsTmp("申请单图像"))
                        
                        With curImage
                            .BorderStyle = 6
                            .BorderWidth = 1
                            .BorderColour = vbWhite
                        End With

                    '取消自动剪影,因为DicomObjects控件本身对处理剪影有BUG，存在（0028，6100）时，会自动对图像进行剪影，
                    '导致晋煤的DSA图像不能正常显示
                    '虽然设置图像的mask=0 ,可以取消剪影，但是每次图像被添加到新的Dicomimages之后，自动又将mask设置成1了，
                    '这样在程序中无法很好的控制，因此直接去掉（0028，6100）这个属性。
                    If Not IsNull(curImage.Attributes(&H28, &H6100).value) Then
                        curImage.Attributes.Remove &H28, &H6100
                    End If
                End If

                rsTmp.MoveNext
                If Not rsTmp.EOF Then strNextDevNo = Nvl(rsTmp("设备号"))
            Loop
            If dcmViewer.Images.Count > 0 Then
                dcmViewer.CurrentIndex = 1
                dcmViewer.Images(dcmViewer.Images.Count).BorderColour = vbRed
            End If
        Else
            dcmViewer.MultiColumns = 1
            dcmViewer.MultiRows = 1
        End If
    Inet1.FuncFtpDisConnect
    Exit Sub
DBError:
    '断开FTP连接
    Inet1.FuncFtpDisConnect
    If ErrCenter() = 1 Then
        Resume
    End If

    Call SaveErrLog
End Sub

Private Sub subDeleteFTPImage()
'------------------------------------------------
'功能：删除缩略图中被选中的图像，先从数据库中删除，然后从FTP中删除.
'参数：无
'返回：无，直接删除缩略图中最后一个图像
'------------------------------------------------
Dim blnResult As Boolean
    If mintCurImgIndex > 0 And mintCurImgIndex <= dcmMiniature.Images.Count Then
        
        '从数据库和FTP中删除缩略图中被选中的图像
        blnResult = DelPetitionImg()
        
        If blnResult = True Then    '删除成功，则修改缩略图状态，并触发StateChanged事件
            '在缩略图中删除图像
            dcmMiniature.Images.Remove mintCurImgIndex
            dcmView.Images.Clear
            
            If mintCurImgIndex > dcmMiniature.Images.Count Then
                mintCurImgIndex = dcmMiniature.Images.Count
            End If

            If mintCurImgIndex > 0 Then
                dcmMiniature.Images(mintCurImgIndex).BorderColour = vbRed
            End If
            
            Call fmeDcmViewer_Resize
        End If
    End If
End Sub

Private Sub subDeleteDcmImage()

'在缩略图中删除图像
        dcmMiniature.Images.Remove mintCurImgIndex
        dcmView.Images.Clear
        
        If mintCurImgIndex > dcmMiniature.Images.Count Then
            mintCurImgIndex = dcmMiniature.Images.Count
        End If

        If mintCurImgIndex > 0 Then
            dcmMiniature.Images(mintCurImgIndex).BorderColour = vbRed
        End If
        
        Call fmeDcmViewer_Resize

End Sub


Private Function DelPetitionImg() As Boolean
'------------------------------------------------
'功能：从数据库和FTP中删除图像，删除缩略图中被选中的图像
'参数：无
'返回：True－－删除成功，False－－删除失败

    Dim ImgTmp As New DicomImage
    Dim rsTmp As New ADODB.Recordset
    Dim strReportImage As String
    Dim varTmp As Variant
    Dim strResult As Long
    Dim strSql As String
    Dim strFTPUser As String              'FTP账号
    Dim strFtpPass As String              'FTP密码
    Dim strFtpDir As String               'FTP目录
    Dim strFtpIp As String                'FTP地址
    
    If dcmMiniature.Images.Count < mintCurImgIndex Then Exit Function
    Set ImgTmp = dcmMiniature.Images(mintCurImgIndex)
                
    On Error GoTo errHand
    
    strSql = "select 扫描时间,设备号 from 影像申请单图像 where 医嘱ID=[1] and 申请单图像 = [2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "提取申请单图像信息", mlngAdviceID, ImgTmp.Tag)
    
    If rsTmp.EOF = True Then
        MsgBoxD Me, "没有找到可以删除的图像!", vbInformation, gstrSysName
        DelPetitionImg = False
        Exit Function
    End If
    
    Call funGetStorageDevice(Me, Nvl(rsTmp("设备号")), strFtpDir, strFtpIp, strFTPUser, strFtpPass)
    
    strResult = miNet.FuncFtpConnect(strFtpIp, strFTPUser, strFtpPass)
    If strResult = 0 Then
        MsgBoxD Me, "连接FTP失败，请检查FTP连接。", vbInformation, gstrSysName
        DelPetitionImg = False
        Exit Function
    End If

    gstrSQL = "ZL_影像申请单图像_DELETE(" & mlngAdviceID & ",'" & ImgTmp.Tag & "')"

    zlDatabase.ExecuteProcedure gstrSQL, "影像图像删除"

    '删除图像文件
    RemoveFromURL strFtpDir & _
        Format(Nvl(rsTmp("扫描时间"), CStr(zlDatabase.Currentdate)), "yyyyMMdd") & "/" & _
        mlngAdviceID & "/" & Mid(ImgTmp.Tag, 1, InStr(ImgTmp.Tag, ".") - 1)

    miNet.FuncFtpDisConnect
    DelPetitionImg = True
    Exit Function
errHand:
    '断开FTP连接
    miNet.FuncFtpDisConnect
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Private Sub WriteToURL(ByVal strFileName As String, ByVal strDestFileName As String)
'------------------------------------------------
'功能：将本地文件保存到远程网络上
'参数：strFileName--本地文件名，strDestFileName－－远程文件名
'返回：无
'-----------------------------------------------
'功能：
    Dim objFileSystem As New Scripting.FileSystemObject
    Dim strPath As String
    
    '在FTP中创建目录
    strPath = objFileSystem.GetParentFolderName(strDestFileName)
    miNet.FuncFtpMkDir "/", strPath
    
    '向FTP上传文件
    miNet.FuncUploadFile strPath, strFileName, objFileSystem.GetFileName(strDestFileName)
End Sub


Private Sub RemoveFromURL(ByVal strDestFileName As String)
'------------------------------------------------
'功能：从FTP删除文件
'参数：strDestFileName－－远程文件名
'返回：无
'-----------------------------------------------
    Dim objFileSystem As New Scripting.FileSystemObject
    
    miNet.FuncDelFile objFileSystem.GetParentFolderName(strDestFileName), objFileSystem.GetFileName(strDestFileName)
End Sub

Private Sub dcmMiniature_MouseDown(Button As Integer, Shift As Integer, X As Long, Y As Long)
On Error GoTo errHandle
    Dim i As Integer

    If Button = 1 Then
        mCorpSize.X = 0
        mCorpSize.Y = 0
        
        '被选中图像显示红框
        i = dcmMiniature.ImageIndex(X, Y)
        If i <> 0 Then
        
            If mintCurImgIndex > 0 And mintCurImgIndex <= dcmMiniature.Images.Count Then
                dcmMiniature.Images(mintCurImgIndex).BorderColour = vbWhite
            End If
            dcmMiniature.Images(i).BorderColour = vbRed
            mintCurImgIndex = i
            
        End If
    End If
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub




Private Sub Form_Resize()
On Error Resume Next
    fmeDcmViewer.Top = 0
    fmeDcmViewer.Left = 0
    fmeDcmViewer.Width = Me.ScaleWidth
    fmeDcmViewer.Height = Me.ScaleHeight - fmeInfoCtrl.Height

    dcmMiniature.Top = 60
    dcmMiniature.Left = 60
    dcmMiniature.Width = fmeDcmViewer.Width - 120
    dcmMiniature.Height = fmeDcmViewer.Height
    
    dcmViewImg.Top = 60
    dcmViewImg.Left = 60
    dcmViewImg.Width = fmeDcmViewer.Width - 120
    dcmViewImg.Height = fmeDcmViewer.Height


    fmeInfoCtrl.Top = fmeDcmViewer.Height
    fmeInfoCtrl.Left = 0
    fmeInfoCtrl.Width = fmeDcmViewer.Width

    fmePatientInfo.Top = 0
    fmePatientInfo.Left = 60
    fmePatientInfo.Width = fmeInfoCtrl.Width - FmeScanbtn.Width - IIf(FmeProcess.Visible, FmeProcess.Width, -60) - 220
    fmePatientInfo.Height = fmeInfoCtrl.Height

    FmeProcess.Top = 0
    FmeProcess.Left = fmePatientInfo.Width + 120
    FmeProcess.Height = fmeInfoCtrl.Height

    FmeScanbtn.Top = 0
    FmeScanbtn.Left = fmePatientInfo.Width + IIf(FmeProcess.Visible, FmeProcess.Width, -60) + 180
    FmeScanbtn.Height = fmeInfoCtrl.Height

End Sub


Private Sub fmeDcmViewer_Resize()
    Dim iRows As Integer
    Dim iCols As Integer
    
    On Error Resume Next
    
    dcmMiniature.Left = 0
    dcmMiniature.Top = 0
    dcmMiniature.Width = fmeDcmViewer.Width
    dcmMiniature.Height = fmeDcmViewer.Height
    
    dcmViewImg.Top = 60
    dcmViewImg.Left = 60
    dcmViewImg.Width = fmeDcmViewer.Width - 120
    dcmViewImg.Height = fmeDcmViewer.Height
    
    '自动对图像做布局
    '计算缩略图的图像布局
    ResizeRegion dcmMiniature.Images.Count, dcmMiniature.Width, dcmMiniature.Height, iRows, iCols
    
    dcmMiniature.MultiColumns = iCols
    dcmMiniature.MultiRows = iRows

End Sub


