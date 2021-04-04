VERSION 5.00
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "Mscomm32.ocx"
Object = "{853AAF94-E49C-11D0-A303-0040C711066C}#4.3#0"; "DicomObjects.ocx"
Object = "{B1790453-7708-48C1-B5CC-75255FA4B066}#1.0#0"; "ZLDSVideoProcess.ocx"
Object = "{84926CA3-2941-101C-816F-0E6013114B7F}#1.0#0"; "IMGSCAN.OCX"
Begin VB.Form frmVideoCaptureV2 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   6825
   ClientLeft      =   0
   ClientTop       =   60
   ClientWidth     =   10410
   ClipControls    =   0   'False
   DrawStyle       =   5  'Transparent
   Icon            =   "frmVideoCaptureV2.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   6825
   ScaleWidth      =   10410
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Tag             =   "视频采集"
   Begin ScanLibCtl.ImgScan ImageScanner 
      Left            =   1440
      Top             =   0
      _Version        =   65536
      _ExtentX        =   873
      _ExtentY        =   873
      _StockProps     =   0
      StopScanBox     =   -1  'True
      FileType        =   3
      CompressionType =   0
      CompressionInfo =   0
      ScanTo          =   4
   End
   Begin VB.PictureBox picAfter 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   3840
      ScaleHeight     =   375
      ScaleWidth      =   2655
      TabIndex        =   12
      Top             =   5280
      Visible         =   0   'False
      Width           =   2655
      Begin VB.Label labCloseAfter 
         BackStyle       =   0  'Transparent
         Caption         =   "×"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   2400
         TabIndex        =   14
         Top             =   90
         Width           =   255
      End
      Begin VB.Label labAfterInfo 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "标识:---"
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
         Left            =   0
         TabIndex        =   13
         Top             =   90
         Visible         =   0   'False
         Width           =   2295
      End
   End
   Begin VB.PictureBox picLock 
      BackColor       =   &H0080FFFF&
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   3720
      ScaleHeight     =   375
      ScaleWidth      =   2655
      TabIndex        =   10
      Top             =   840
      Visible         =   0   'False
      Width           =   2655
      Begin VB.Label labCloseLock 
         BackStyle       =   0  'Transparent
         Caption         =   "×"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   2400
         TabIndex        =   15
         Top             =   90
         Width           =   255
      End
      Begin VB.Label labLockInfo 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "锁定:---"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   0
         TabIndex        =   11
         Top             =   90
         Width           =   2415
      End
   End
   Begin VB.PictureBox picView 
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Height          =   3615
      Left            =   1680
      ScaleHeight     =   3615
      ScaleWidth      =   6855
      TabIndex        =   5
      Top             =   1440
      Width           =   6855
      Begin ZLDSVideoProcess.DSCapture wdmCapture 
         Height          =   3375
         Left            =   120
         TabIndex        =   6
         Top             =   120
         Width           =   4215
         Object.Visible         =   -1  'True
         AutoScroll      =   0   'False
         AutoSize        =   0   'False
         AxBorderStyle   =   1
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
         CurWidth        =   281
         CurHeight       =   225
         CurVideoWidth   =   279
         CurVideoHeight  =   205
         ShowModel       =   0
         CapParameterWindPos=   8
         SnatchWay       =   0
         ParameterCfgFileName=   ""
         HideCfgItem     =   0
         AppHandle       =   0
      End
      Begin VB.TextBox txtInputText 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   4440
         TabIndex        =   8
         Text            =   "Text1"
         Top             =   2520
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.PictureBox picCusVideo 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   855
         Left            =   4440
         ScaleHeight     =   57
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   72
         TabIndex        =   7
         Top             =   240
         Width           =   1080
      End
      Begin DicomObjects.DicomViewer dcmView 
         Height          =   855
         Left            =   4440
         TabIndex        =   9
         Top             =   1320
         Width           =   1095
         _Version        =   262147
         _ExtentX        =   1931
         _ExtentY        =   1508
         _StockProps     =   35
         UseScrollBars   =   0   'False
      End
   End
   Begin VB.PictureBox pbxSize 
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      Height          =   75
      Index           =   0
      Left            =   1440
      MousePointer    =   7  'Size N S
      ScaleHeight     =   75
      ScaleWidth      =   7335
      TabIndex        =   4
      Top             =   1200
      Width           =   7335
   End
   Begin VB.PictureBox pbxSize 
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      Height          =   3975
      Index           =   3
      Left            =   8760
      MousePointer    =   9  'Size W E
      ScaleHeight     =   3975
      ScaleWidth      =   75
      TabIndex        =   3
      Top             =   1215
      Width           =   75
   End
   Begin VB.PictureBox pbxSize 
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      Height          =   3975
      Index           =   2
      Left            =   1440
      MousePointer    =   9  'Size W E
      ScaleHeight     =   3975
      ScaleWidth      =   75
      TabIndex        =   2
      Top             =   1200
      Width           =   75
   End
   Begin VB.Timer tmrReg 
      Interval        =   10000
      Left            =   960
      Top             =   0
   End
   Begin VB.Timer tmrComm 
      Interval        =   2
      Left            =   480
      Top             =   0
   End
   Begin MSCommLib.MSComm commListener 
      Left            =   2640
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.PictureBox picTemp2 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   2040
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox pbxSize 
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   75
      Index           =   1
      Left            =   1440
      MousePointer    =   7  'Size N S
      ScaleHeight     =   75
      ScaleWidth      =   7335
      TabIndex        =   1
      Top             =   5160
      Width           =   7335
   End
   Begin XtremeCommandBars.CommandBars cbrMain 
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmVideoCaptureV2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'功能：采集和录制视频图像
'
'
'
'修改历史：
'
'2010-01-19: 将wdm视频组件加入到采集模块中，并支持对指定SDK视频采集的实现
'
'
'
'裁剪原理说明：
'
'
'
'
'                ------------------------------------
'               |原始图象大小                        |
'               |                                    |
'               |                                    |
'               |         ------------------         |
'               |        |                  |        |
'               |<-- L-->|     裁剪大小     |<-- R-->|
'               |        |                  |        |
'               |         ------------------         |
'               |                                    |
'               |                                    |
'               |                                    |
'                ------------------------------------
'
'L表示左边裁剪的大小百分比
'R表示右边裁剪的大小百分比
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


Option Explicit

 
Private Const M_STR_HINT_NoSelectData As String = "无效的检查数据，请选择需要执行的检查记录。"


Private Const conMenu_ImgPro_Window = 501           '亮度对比度
Private Const conMenu_ImgPro_Zoom = 502             '缩放
Private Const conMenu_ImgPro_Corp = 512             '拖动

Private Const conMenu_ImgPro_Rotate_Pop = 503          '顺时针旋转
Private Const conMenu_ImgPro_RRotate = 5030          '顺时针旋转
Private Const conMenu_ImgPro_LRotate = 5031          '逆时针旋转

Private Const conMenu_ImgPro_Smooth_Pop = 504        '锐化
Private Const conMenu_ImgPro_Sharpness = 5040        '锐化
Private Const conMenu_ImgPro_Smooth = 5041           '平滑

Private Const conMenu_ImgPro_Lab_Pop = 505
Private Const conMenu_ImgPro_Text = 5050             '文字标注
Private Const conMenu_ImgPro_Arrow = 5051            '箭头标注
Private Const conMenu_ImgPro_Ellipse = 5052          '圆形标注

Private Const conMenu_ImgPro_Save = 506         '保存
Private Const conMenu_ImgPro_RectSave = 50601        '裁剪保存
Private Const conMenu_ImgPro_DirectSave = 50602        '直接保存
Private Const conMenu_ImgPro_RectCapture = 507         '裁剪后采集
Private Const conMenu_ImgPro_Restore = 508       '恢复



'裁剪范围 modify by tjh 2010-01-19
Private Type TCutRange
  LeftRate As Double
  TopRate As Double
  WidthRate As Double
  HeightRate As Double
End Type


'视频区域 modify by tjh 2010-01-19
Private Type TVideoArea
  Left As Long
  Top As Long
  Width As Long
  Height As Long
End Type


'移动方向 modify by tjh 2010-01-19
Private Enum TMoveOrientation
  moUp = 0    '上
  moDown = 1  '下
  moLeft = 2  '左
  moRight = 3 '右
End Enum


'COM脚踏端口状态
Private Type TComPortState
    intComState As Integer          'COM口的状态
    lngComTime As Long              '记录com口保持状态的时间
    dtLastCapture As Date           '最近脚踏踩下的时间
    blnCTSHolding As Boolean        '记录常态时的CTS线的电平
    
    lngSignalCount As Long
    lngStartTick As Long
    lngEndTick As Long
End Type


Private Type DlgFileInfo
    iCount As Long
    sPath As String
    sFIle() As String
End Type

Private Enum Dkp_ID
    Dkp_ID_Video = 1     '检查列表
    Dkp_ID_Miniature      '当前病人基本信息
End Enum



Private mobjCapHelper As ICapHelper
Private mstrAfterCapTag As String
Private mstrBufferDir As String
Private mblnIsLock As Boolean
 
Private mintCaptureFlag As Integer

Private mobjCustomDevice As Object  '专用视频采集对象

Private dcmglbUID As New DicomGlobal    '定义UIDRoot=1

Private WithEvents mobjDxDevice As clsDxHidDevice   '实现蓝韵手柄之类的采集方式
Attribute mobjDxDevice.VB_VarHelpID = -1
 
Private WithEvents mfrmParameter As frmVideoSetupV2
Attribute mfrmParameter.VB_VarHelpID = -1
Private mfrmOpenStudy As frmOpenStudyList

Private mblnRealTime As Boolean         '记录当前显示的是实时显示还是图像处理窗口。True = 实时视频窗口，False = 图像处理窗口
Private mblnPlayVideo As Boolean        '记录当前显示的图像处理窗口中显示的是图片还是录象？True = 录象；False = 图片
Private mintMouseState As Integer       '记录当前图像处理时的鼠标状态:1=亮度对比度；2=缩放；3=裁剪缩放；11=箭头标注；12=圆形标注；13=文字标注


Private mlngBaseX As Long               'dcmView中鼠标Down时的X坐标
Private mlngBaseY As Long               'dcmView中鼠标Down时的Y坐标
Private mMouseDownPoint As TPoint       '鼠标在DcmImage上按下时的位置
Private mInitScrollPoint As TPoint      '开始拖动时的初始位置
Private mCorpSize As TPoint             '拖动后的相对偏移位置

Private mstrTempDirOfScan As String          '扫描的临时目录
Private mintScanImageIndex As Integer        '扫描图像索引
 

Private mblnMoveDown  As Boolean        '用于判断是否按下鼠标左键
Private mblnDcmViewDown As Boolean      '用于判断dcmView中鼠标是否被按下
Private mintCurImgIndex As Integer      '当前被选中的图象索引
Private mdcmSelectLabel As DicomLabel   '当前被选中的标注

Private mstrAviFileName As String       '录像文件名

Private mstrAfterDir As String

Private mcpsComState As TComPortState       'Com端口使用状态


Private mlngImageSwapWay As Long          '0-内存交换,1-剪贴板交换，2-本地文件交换
Private mblnUseBeforeConvert As Boolean     '提前转换


Private mblnReadOnly As Boolean         '是否只能查看True查看模式，False采集模式
 

Private mVideoCapture As clsVideoCapture '视频采集对象
Private WithEvents mobjPlayWindow As frmPlaying
Attribute mobjPlayWindow.VB_VarHelpID = -1

Private mdblZoomRate As Double  '缩放比率（在cbrMain的cbrMain_ResizeClient事件中需要重新计算该值）
Private mVideoSize As TVideoSize '视频大小（由相关的视频组件保存）
Private mCurCutRange As TCutRange '视频裁剪范围设置（该参数通过GetString和SaveString保存在注册表中）
Private mVideoArea As TVideoArea  '视频客户区域设置（在cbrMain的cbrMain_ResizeClient事件中需要重新计算该值）

Private Const M_LNG_REFRESHINTERVAL As Long = 600 '刷新间隔
Private mstrVideoRegTime As String '保存视频启动注册时间
Private mstrMsg As String
Private mblnRefreshState As Boolean
Private mblnInitState As Boolean

Private mintFontSize As Integer '字号
 
Private mblnImageShield As Boolean   '是否屏蔽大图
Private mblnTimerState As Boolean '计时开启状态

Private Const CAPTURE_PARAMETER_CONFIG_FILE_NAME As String = "ZLVideoProcess.ini"
Private Const CONST_STR_DEFAULT_TEMP_SCAN_DIR_NAME As String = "\TempScan"  '默认扫描文件临时存储路径
Private Const CONST_STR_DEFAULT_SCAN_FILE_TEMPLATE As String = "Scan"  '默认扫描文件临时存储路径



'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'----------------------------------------------------------------------------------------------------------

Private Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Private Declare Function MoveFile Lib "kernel32" Alias "MoveFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String) As Long

Public Event OnControlResize(objControl As Object)


'获取视频采集对象
Property Get videoCapture() As clsVideoCapture
    Set videoCapture = mVideoCapture
End Property


'获取视频采集窗口的初始化状态
Property Get InitState() As Boolean
    InitState = mblnInitState
End Property

'当前锁定状态
Property Get IsLock() As Boolean
    IsLock = mblnIsLock
End Property

'是否后台采集中
Property Get IsAfter() As Boolean
    IsAfter = IIf(Len(mstrAfterCapTag) > 0, True, False)
End Property
'
'Private Sub UnLockStudy()
''解锁检查
'    mblnIsLock = False
'End Sub


Public Sub NotificationRefresh()
'通知刷新
    mblnRefreshState = False
End Sub

Private Function GetTag(ByVal FolderName As String, ByRef strType As String) As Integer
'解析文件夹名称中的标识号，FolderName：目标目录名，strType： 返回“标识” 或 “检查”
On Error GoTo errH
    Dim i As Integer
    Dim strTmp As String
    
    strType = Mid(FolderName, 1, 2)
    strTmp = Mid(FolderName, 3, Len(FolderName) - 2)
    i = InStr(strTmp, "-")
    GetTag = Val(Mid(strTmp, 1, i - 1))
    
    Exit Function
errH:
    GetTag = 0
End Function

Private Function GetStudyUIDFromFolderName(ByVal FolderName As String) As String
'解析文件夹名称中的检查UID并返回，若出错返回文件夹名
On Error GoTo errH
    Dim i As Integer
    Dim j As Integer
    
    i = InStr(FolderName, "-")
    j = Len(FolderName)
    
    GetStudyUIDFromFolderName = Mid(FolderName, i + 1, j - i)
    Exit Function
errH:
    GetStudyUIDFromFolderName = FolderName
End Function


Private Sub Form_Initialize()
'初始化模块变量
    mblnInitState = False
    mblnIsLock = False
    mblnTimerState = False
End Sub




Public Sub ShowVideoConfig()
On Error GoTo errHandle
'视频配置

BUGEX "ShowVideoConfig 1"
    '先保存修改的参数设置
    Call SaveParameterCfg
BUGEX "ShowVideoConfig 2"

    '判断是否处于实时模式显示状态
    If mblnRealTime = False Then
        Call ConfigVideoShowState(True)
    End If
    
    '打开参数配置窗口
    If mfrmParameter.ShowParameterConfig(mVideoCapture, Me) = False Then Exit Sub
    
    '重新读取配置参数------------------------------------------
BUGEX "ShowVideoConfig 3"
    Call InitParameter

    
BUGEX "ShowVideoConfig 6"
        
    If gobjCapturePar.VideoDirverType = vdtCustom Then
        If mobjCustomDevice Is Nothing Then Call InitCustomDevice

        Call mobjCustomDevice.StartPreview
        Call mobjCustomDevice.UpdateWindow(picCusVideo.ScaleWidth, picCusVideo.ScaleHeight)
    End If
    
    Call OpenComm
    
    gstrHotKeyTest = GetSetting("ZLSOFT", "公共模块", "采集热键", "F8")
    
    
BUGEX "ShowVideoConfig End"
Exit Sub
errHandle:
    MsgboxCus err.Description, vbOKOnly, G_STR_HINT_TITLE
    err.Clear
End Sub


Private Sub InitParameter()
'初始化参数设置
    Dim rsTmp As New ADODB.Recordset
    Dim intVideoCapture As Integer
    Dim strSQL As String

    mintCaptureFlag = 0
    mblnRealTime = True
    mintCurImgIndex = 0
    mblnPlayVideo = False
    
 

    '如果程序在磁盘的根目录则app.path为“x:\”
    mstrBufferDir = GetAppPath & "\TmpImage\"
    mstrAfterDir = GetAppPath & "\TmpAfterImage\"
    
'    mstrAviFileName = mstrBufferDir & "TmpVideo.avi"
    
    gint视频设备数量 = getLicenseCount(LOGIN_TYPE_视频设备)
    
    mlngImageSwapWay = Val(GetSetting("ZLSOFT", G_STR_REG_PATH_PUBLIC, "图像交换方式", 0))
    Call SaveSetting("ZLSOFT", G_STR_REG_PATH_PUBLIC, "图像交换方式", mlngImageSwapWay)
    
    mblnUseBeforeConvert = GetSetting("ZLSOFT", G_STR_REG_PATH_PUBLIC, "图像提前转换", 0)
    Call SaveSetting("ZLSOFT", G_STR_REG_PATH_PUBLIC, "图像提前转换", IIf(mblnUseBeforeConvert, 1, 0))

    '读取裁剪比率
    mCurCutRange.LeftRate = Val(GetSetting("ZLSOFT", G_STR_REG_PATH_PUBLIC, "mdblX1Scale", 0))  '使用mdblX1Scale名称是为了保证和以前的参数设置兼容
    mCurCutRange.WidthRate = Val(GetSetting("ZLSOFT", G_STR_REG_PATH_PUBLIC, "mdblX2Scale", 0))
    mCurCutRange.TopRate = Val(GetSetting("ZLSOFT", G_STR_REG_PATH_PUBLIC, "mdblY1Scale", 0))
    mCurCutRange.HeightRate = Val(GetSetting("ZLSOFT", G_STR_REG_PATH_PUBLIC, "mdblY2Scale", 0))

    If (mCurCutRange.LeftRate >= 1) Or (mCurCutRange.LeftRate < 0) Then mCurCutRange.LeftRate = 0
    If (mCurCutRange.WidthRate >= 1) Or (mCurCutRange.WidthRate < 0) Then mCurCutRange.WidthRate = 0
    If (mCurCutRange.TopRate >= 1) Or (mCurCutRange.TopRate < 0) Then mCurCutRange.TopRate = 0
    If (mCurCutRange.HeightRate >= 1) Or (mCurCutRange.HeightRate < 0) Then mCurCutRange.HeightRate = 0

    '定义UIDRoot=1
    dcmglbUID.RegString("UIDRoot") = "1"
    
    '读取采集配置参数
    If gobjCapturePar Is Nothing Then
        Set gobjCapturePar = New clsCaptureParameter
    End If
    
    Call gobjCapturePar.ReadParameter
    
    If gobjCapturePar.IsAllowChangeSize = False Then
        Me.pbxSize.Item(0).MousePointer = 0
        Me.pbxSize.Item(1).MousePointer = 0
        Me.pbxSize.Item(2).MousePointer = 0
        Me.pbxSize.Item(3).MousePointer = 0
    Else
        Me.pbxSize.Item(0).MousePointer = 7
        Me.pbxSize.Item(1).MousePointer = 7
        Me.pbxSize.Item(2).MousePointer = 9
        Me.pbxSize.Item(3).MousePointer = 9
    End If
 
End Sub


Public Sub zlRePreview(Optional ByVal blnForceStop As Boolean = False)
'重新进入视频预览
    If mVideoCapture.IsStartup Then
        If blnForceStop Or mVideoCapture.VideoDriverType <> vdtWDM Then
            Call mVideoCapture.StopPreview
            Call mVideoCapture.StartPreview
        End If
         
        Call wdmCapture.RePreview
    End If
End Sub

Public Sub zlInitModule(objCapHelper As Object)
BUGEX "zlPacsCapture zlInitModule 0"
'初始化模块参数
    Set mobjCapHelper = objCapHelper
    
    If mblnInitState Then Exit Sub
    
    '初始化参数
    Call InitParameter
    
BUGEX "zlInitModule 1"
    '初始化专用视频采集接口
    Call InitCustomDevice
    
BUGEX "zlInitModule 2"
    '打开视频采集设备
    Call OpenVideoCaptureDevice
  
BUGEX "zlInitModule End"
    mblnInitState = True
End Sub

Private Sub InitCustomDevice()
    Dim strCustomDeviceDir As String        '专用视频采集部件路径
    Dim strCustomDeviceDllName As String    '专用视频采集部件名称
    Dim objFile As New FileSystemObject
    
    '初始化专用视频采集接口
    strCustomDeviceDir = gobjCapturePar.CustomDevicePath
    If Trim(strCustomDeviceDir) <> "" And gobjCapturePar.VideoDirverType = vdtCustom Then
        strCustomDeviceDllName = Trim(Replace(objFile.GetFileName(strCustomDeviceDir), ".dll", ""))
        
        Set mobjCustomDevice = CreateObject(strCustomDeviceDllName & ".cls" & strCustomDeviceDllName)
        
        If Not mobjCustomDevice Is Nothing Then
            Call mobjCustomDevice.zlInit(gcnVideoOracle, UserInfo.id, glngDepartId, picCusVideo.hwnd)
        End If
    End If
End Sub


Public Sub zlRestoreWindow(ByVal blnReadOnly As Boolean, Optional ByVal blnIsMain As Boolean = False, _
    Optional ByVal blnIsOnlyState As Boolean = False)
'刷新界面
On Error GoTo errHandle
    mblnReadOnly = blnReadOnly
    
    If blnIsOnlyState Then Exit Sub
    
    If blnIsMain And cbrMain(2).position <> xtpBarRight Then
        cbrMain(2).position = xtpBarRight
        cbrMain.RecalcLayout
    ElseIf blnIsMain = False And cbrMain(2).position <> xtpBarLeft Then
        cbrMain(2).position = xtpBarLeft
        cbrMain.RecalcLayout
    End If
    
    If IsTwainCaptureWay Then Exit Sub
 
    Call ConfigVideoShowState(True)
    
    
    Exit Sub
errHandle:
    MsgboxCus err.Description, vbOKOnly, G_STR_HINT_TITLE
End Sub


Public Sub zlPreviewThumbnail(objImg As Object)
'预览缩略图
    Dim dblTempZoom As Double
    Dim img As DicomImage
    Dim i As Integer
    
    '将被选中图像装载到dcmView中
    dcmView.Images.Clear
    
    If objImg Is Nothing Then Exit Sub
    
    If txtInputText.Visible Then txtInputText.Visible = False
 
    dcmView.Images.Add objImg
    
    '显示dcmView，隐藏picVideo
    dcmView.CurrentImage.BorderWidth = 0
    
    dblTempZoom = dcmView.CurrentImage.ActualZoom
    dcmView.CurrentImage.StretchToFit = False
        
    '判断当进入浮动窗口时，缩放比率不能小于0.1
    If dblTempZoom < 0.1 Then dblTempZoom = 0.1
                  
    Call subCenterZoom(Me, dcmView.CurrentImage, dcmView, dblTempZoom, mCorpSize)
    
    Call ConfigVideoShowState(False)
End Sub


Private Sub StopCapture()
'-----------------------------------------------------------------------------------------
'功能：停止显示视频采集,释放视频采集窗口，
'      释放串口侦听的端口
'参数：无
'返回：无
'-----------------------------------------------------------------------------------------
    
    '关闭COMM口
    If commListener.PortOpen Then commListener.PortOpen = False
    
    '采用Midi接口需在消毁事件句柄
    If Not mobjDxDevice Is Nothing Then
        If mobjDxDevice.Handle <> 0 Then Call mobjDxDevice.CloseDxDevice
    End If
    
    '释放采集设备及窗体
    If Not mVideoCapture Is Nothing Then
        Call mVideoCapture.StopPreview
    End If
End Sub



Public Sub zlUpdateCommandBars(Control As XtremeCommandBars.CommandBarControl)
'只有影像采集工作站才具备后台采集功能

'根据当前状态确定菜单的可视和可操作

    '如果没有初始化视频对象，则视频相关的按钮都不允许使用
    If mVideoCapture Is Nothing Then
        Control.Enabled = False
        Exit Sub
    End If
    
    Select Case Control.id
        Case conMenu_Cap_Dynamic       '动态显示
            Control.Checked = mblnRealTime
            Control.Enabled = (Not mblnReadOnly Or Len(mstrAfterCapTag) > 0 Or mblnIsLock) And (Not IsTwainCaptureWay) And (mVideoCapture.IsStartup Or IsCustomCaptureWay)    ' And (mhCapWnd <> 0) modify by tjh at 2010-01-20
            Control.Visible = Not IsTwainCaptureWay 'And Not IsCustomCaptureWay
            
            If mblnRealTime Then
                Control.IconId = conMenu_Cap_Dynamic
            Else
                Control.IconId = 10023
            End If
            
        Case conMenu_Cap_MarkMap       '影像采集
            Control.Enabled = (Not mblnReadOnly Or Len(mstrAfterCapTag) > 0 Or mblnIsLock) And (mVideoCapture.IsStartup Or IsTwainCaptureWay Or IsCustomCaptureWay)
            
'        Case conMenu_Cap_After_Capture  '后台采集
'            Control.Enabled = mVideoCapture.IsStartup
'            Control.Visible = gobjCapturePar.IsUseAfterCapture And (Not IsCustomCaptureWay)
            
        Case conMenu_Cap_Record        '录像
            Control.Enabled = (Not mblnReadOnly Or mblnIsLock) And ((gobjCapturePar.VideoDirverType = vdtWDM And mVideoCapture.IsStartup) Or gobjCapturePar.VideoDirverType = vdtCustom)
            Control.Visible = Not IsTwainCaptureWay And Len(mstrAfterCapTag) <= 0
            
        Case conMenu_Cap_Timer
            Control.Visible = gobjCapturePar.VideoDirverType = vdtWDM
            Control.Enabled = Not mblnReadOnly
            If mblnTimerState Then
                Control.IconId = 10025
                Control.ToolTipText = "关闭计时"
            Else
                Control.IconId = 10024
                Control.ToolTipText = "开启计时"
            End If
'        Case conMenu_Cap_After_Record   '后台录像
'            Control.Enabled = gobjCapturePar.VideoDirverType = vdtWDM And mVideoCapture.IsStartup
'            Control.Visible = Not IsTwainCaptureWay And Not IsCustomCaptureWay And gobjCapturePar.IsUseAfterCapture And False
            
'        Case conMenu_Cap_Record_Stop '停止录像 modify by tjh at 2010-01-22
'            Control.Enabled = mblnRealTime And Not mblnReadOnly And (gobjCapturePar.VideoDirverType = vdtWDM) And mVideoCapture.IsStartup
'            Control.Visible = Not IsTwainCaptureWay And Not IsCustomCaptureWay
            
        Case conMenu_Cap_RecordAudio '录音
            Control.Enabled = Not mblnReadOnly Or mblnIsLock
            Control.Visible = Not IsCustomCaptureWay And Len(mstrAfterCapTag) <= 0
            
        '录像播放,录像停止,录像快进,录像快退,保存录像
        Case conMenu_Cap_Play, conMenu_Cap_Stop, conMenu_Cap_Forward, _
             conMenu_Cap_Back
            If (mblnRealTime = False) And (dcmView.Images.count > 0) Then
                Control.Visible = dcmView.Images(1).Tag.Tag = VIDEOTAG Or dcmView.Images(1).Tag.Tag = AUDIOTAG
                Control.Enabled = dcmView.Images(1).Tag.Tag = VIDEOTAG Or dcmView.Images(1).Tag.Tag = AUDIOTAG
            Else
                Control.Visible = False
                Control.Enabled = False
            End If
            
         '亮度对比度,缩放,裁剪缩放,顺时针旋转,逆时针旋转,锐化,平滑,高级处理
        Case conMenu_ImgPro_Window, conMenu_ImgPro_Zoom, conMenu_ImgPro_Save, conMenu_ImgPro_RectSave, conMenu_ImgPro_DirectSave, _
             conMenu_ImgPro_Rotate_Pop, conMenu_ImgPro_RRotate, conMenu_ImgPro_LRotate, _
             conMenu_ImgPro_Smooth_Pop, conMenu_ImgPro_Sharpness, conMenu_ImgPro_Smooth, conMenu_ImgPro_Corp

            Control.Visible = dcmView.Visible
            Control.Enabled = (mblnRealTime = False)
        '箭头标注,圆形标注,文字标注,
        Case conMenu_ImgPro_Lab_Pop, conMenu_ImgPro_Arrow, conMenu_ImgPro_Ellipse, conMenu_ImgPro_Text
            Control.Visible = dcmView.Visible
            Control.Enabled = (mblnRealTime = False)
            
'        Case conMenu_Cap_OpenStudyList
'            Control.Enabled = True
'            Control.Visible = gobjCapturePar.IsUseCaptureLock
            
        Case conMenu_Cap_StudySyncState
            Control.Enabled = Not mblnReadOnly Or mblnIsLock
'            Control.Visible = gobjCapturePar.IsUseCaptureLock
            
        Case conMenu_Cap_After_Tag
            Control.Enabled = mVideoCapture.IsStartup
'            Control.Visible = gobjCapturePar.IsUseAfterCapture
            
'        ''''''''''''
'        Case conMenu_Cap_ImgImport
'            Control.Enabled = Not mblnReadOnly
            
    End Select
End Sub


''''''''''''''''''''''''''''''''''
'扫描图像
''''''''''''''''''''''''''''''''''

Private Sub DelScanTmpDir(ByVal strDir As String)
'删除扫描临时文件
On Error GoTo errHandle
    If Dir(strDir, vbDirectory) <> "" Then
      Call mdlDir.DeleteFolder(strDir)
    End If
errHandle:
End Sub

Private Sub DoScanCapture()
'扫描图像
On Error GoTo errProcess
                  
    '删除程序中临时存储的图像目录
    Call DelScanTmpDir(mstrTempDirOfScan)
        
    If Dir(mstrTempDirOfScan, vbDirectory) = "" Then
      Call MkDir(mstrTempDirOfScan)
    End If
    
    '删除twain设备临时存储的目录
    Call DelScanTmpDir(gobjCapturePar.ScanDeviceTmpDir)
    
    If Dir(gobjCapturePar.ScanDeviceTmpDir, vbDirectory) = "" Then
      Call MkDir(gobjCapturePar.ScanDeviceTmpDir)
    End If
    
    mintScanImageIndex = 0
    
    '设置扫描后的文件数据类型
    ImageScanner.FileType = BMP_Bitmap
    ImageScanner.StopScanBox = True
    ImageScanner.ShowSetupBeforeScan = True
    ImageScanner.ScanTo = UseFileTemplateOnly
    
    '设置采集的模板文件
    ImageScanner.Image = mstrTempDirOfScan & "\" & CONST_STR_DEFAULT_SCAN_FILE_TEMPLATE
    
    
    If Not ImageScanner.ScannerAvailable Then ImageScanner.OpenScanner
  
    Call ImageScanner.StartScan
    Call ImageScanner.StopScan
    Call ImageScanner.CloseScanner
    
    Exit Sub
errProcess:
    Call ImageScanner.CloseScanner

    MsgboxCus err.Description, vbOKOnly, G_STR_HINT_TITLE
    err.Clear
End Sub


Public Sub ForeCapture(ByVal blnIsReal As Boolean)
'前台图像采集
    Dim blnIsRealCap As Boolean
    
    If Not ((Not mblnReadOnly Or Len(mstrAfterCapTag) > 0 Or mblnIsLock) And (mVideoCapture.IsStartup Or IsTwainCaptureWay Or IsCustomCaptureWay)) Then Exit Sub '如果为只读，或者视频没有启动，则不允许采集
            
    If IsTwainCaptureWay Then
        Call DoScanCapture  '通过TWAIN接口采集图像
    Else
        If Not blnIsReal Then
            blnIsRealCap = IIf(MsgboxCus("确定要采集当前静态图像吗？选“是”则采集当前处理图像。", _
                                vbQuestion + vbYesNo + vbDefaultButton1, G_STR_HINT_TITLE) = vbYes, False, True)
        End If
        
        If blnIsReal = False Then
            '非实时采集
            Call DoNormalCapture(False)
            Exit Sub
        End If
        
        If IsCustomCaptureWay Then
            '自定义采集
            Call DoCustomCapture
        Else
            '采集图像
            Call DoNormalCapture(True)
        End If
    End If
End Sub



Public Sub zlExecuteCommandBars(Control As XtremeCommandBars.CommandBarControl)
    On Error GoTo errHandle
        Call VideoCaptureMenuProcess(Control)
        
        Call DicomImageMenuProcess(Control)
    Exit Sub
errHandle:
    MsgboxCus err.Description, vbOKOnly, G_STR_HINT_TITLE
End Sub


Private Sub VideoCaptureMenuProcess(Control As XtremeCommandBars.CommandBarControl)
'视频采集菜单处理
    Select Case Control.id
        Case conMenu_Cap_Dynamic       '动态显示
            If IsTwainCaptureWay Then
                Call MsgboxCus("TWAIN采集模式下，不能进行动态视频的显示。", vbOKOnly, G_STR_HINT_TITLE)
            Else
                Call ConfigVideoShowState(True)
            End If
            
        Case conMenu_Cap_MarkMap       '影像采集
            If Len(mstrAfterCapTag) <= 0 Then
                Call ForeCapture(True)
            Else
                Call AfterCapture
            End If
            
        Case conMenu_Cap_Timer
            Call StartTimer
            
'        Case conMenu_Cap_After_Capture  '后台采集
'            Call AfterCapture
            
        Case conMenu_Cap_Record        '录像
            If mstrVideoRegTime = "" Then
                MsgboxCus mstrMsg, vbOKOnly, "提示"
                Exit Sub
            End If
            
            If Control.IconId = conMenu_Cap_Record Then
                Control.IconId = conMenu_Cap_Record_Stop
                
                Call ConfigVideoShowState(True)
                Call StartVideo '开始录像
            Else
                Control.IconId = conMenu_Cap_Record
                Call StopVideo  '停止录像
            End If
            
'        Case conMenu_Cap_Record_Stop  '停止录像 modify by tjh at 2010-01-22
'            If mstrVideoRegTime = "" Then
'                'MsgboxCus  "未检测到有效的注册信息，不能进行录像操作！", vbOKOnly, "提示"
'                Exit Sub
'            End If
'
'            If Len(mstrAviFileName) > 0 Then
'                Call StopVideo
'            End If
            
        Case conMenu_Cap_RecordAudio    '录音
            If mstrVideoRegTime = "" Then
                MsgboxCus mstrMsg, vbOKOnly, "提示"
                Exit Sub
            End If
            
            Call frmRecordAudio.ShowRecordAudio(Me)
            
        Case conMenu_Cap_Play          '录像播放
            Call PlayCurVideo
                
'        Case conMenu_Cap_OpenStudyList      '打开检查采集图像
'            Call mobjCapHelper.OpenLocker
            
        Case conMenu_Cap_StudySyncState     '锁定检查
            If Control.IconId = 10012 Then
                Call CloseAfterCap
                Call LockCapture(Control)
            Else
                Call UnLockCapture(Control)
            End If
        Case conMenu_Cap_After_Tag      '更新后台采集标识
            
            If mstrVideoRegTime = "" Then
                MsgboxCus mstrMsg, vbOKOnly, "提示"
                Exit Sub
            End If
            
            If mblnIsLock Then
                MsgboxCus "锁定状态不能进行后台标记.", vbOKOnly, "提示"
                Exit Sub
            End If
            
            
            Call ResetAfterCaptureTag
            
    End Select
End Sub

Public Sub ResetLockState(ByVal blnIsLock As Boolean)
    Dim objControl As XtremeCommandBars.CommandBarControl
    
    Set objControl = cbrMain.FindControl(, conMenu_Cap_StudySyncState, False, True)
    If objControl Is Nothing Then Exit Sub
    
    If blnIsLock Then
        CloseAfterCap
        Call LockCapture(objControl)
    Else
        Call UnLockCapture(objControl)
    End If
End Sub

Private Sub LockCapture(Control As XtremeCommandBars.CommandBarControl)
    Dim strLocker As String
    
    Control.IconId = 8123
    
    mblnIsLock = True
    
    Call mobjCapHelper.CapLock(strLocker)
    
    
    labLockInfo = "锁定:" & strLocker & ""
    picLock.Visible = True
    
    Call DrawBorderColor(True)
End Sub

Private Sub UnLockCapture(Control As XtremeCommandBars.CommandBarControl)
    Control.IconId = 10012
    
    mblnIsLock = False
    
    Call mobjCapHelper.CapUnlock
    
    labLockInfo.Caption = ""
    picLock.Visible = False
    
    Call DrawBorderColor(False)
End Sub

Private Sub DrawBorderColor(ByVal blnIsLock As Boolean)
    Dim lngColor As Long
    lngColor = IIf(blnIsLock, vbRed, vbBlue)
    
    pbxSize(0).BackColor = lngColor
    pbxSize(1).BackColor = lngColor
    pbxSize(2).BackColor = lngColor
    pbxSize(3).BackColor = lngColor
End Sub

Private Sub DicomImageMenuProcess(Control As XtremeCommandBars.CommandBarControl)
'dicom图像菜单处理
    If mblnRealTime = True Or dcmView.Images.count <= 0 Then Exit Sub
    
    Select Case Control.id
        Case conMenu_ImgPro_Window         '亮度对比度
            subSetMouseState 1
            If mintMouseState = 1 Then
                Control.Checked = True
            End If
            
        Case conMenu_ImgPro_Zoom           '缩放
            subSetMouseState 2
            If mintMouseState = 2 Then
                Control.Checked = True
            End If
            
        Case conMenu_ImgPro_RectSave       '裁剪保存
            subSetMouseState 3
            If mintMouseState = 3 Then
                Control.Checked = True
            End If
            
        Case conMenu_ImgPro_Save, conMenu_ImgPro_DirectSave ' 直接保存
            Call CaptureFrameSelectImage
            
        Case conMenu_ImgPro_RectCapture         '裁剪后采集
            Call CaptureFrameSelectImage
            
        Case conMenu_ImgPro_Rotate_Pop, conMenu_ImgPro_RRotate         '顺时针旋转
            Call subSetRotate(dcmView.Images(1), True)
            
        Case conMenu_ImgPro_LRotate        '逆时针旋转
            Call subSetRotate(dcmView.Images(1), False)
            
        Case conMenu_ImgPro_Sharpness      '锐化
            Call subSetSharp(dcmView.Images(1), True)
            
        Case conMenu_ImgPro_Smooth_Pop, conMenu_ImgPro_Smooth         '平滑
            Call subSetSharp(dcmView.Images(1), False)
            
        Case conMenu_ImgPro_Corp          '拖动
           subSetMouseState 14
            If mintMouseState = 14 Then
                Control.Checked = True
            End If
            
        Case conMenu_ImgPro_Arrow          '箭头标注
            subSetMouseState 11
            If mintMouseState = 11 Then
                Control.Checked = True
            End If
            
        Case conMenu_ImgPro_Ellipse        '圆形标注
            subSetMouseState 12
            If mintMouseState = 12 Then
                Control.Checked = True
            End If
            
        Case conMenu_ImgPro_Lab_Pop, conMenu_ImgPro_Text            '文字标注
            subSetMouseState 13
'            If mintMouseState = 13 Then
'                Control.Checked = True
'            End If
    End Select
    
    If mintMouseState <> 0 Then picView.Refresh
End Sub


Public Sub zlUnloadMe()
    Unload Me
End Sub


Private Sub cbrMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    zlExecuteCommandBars Control
End Sub


'modify by tjh at 2010-01-19
'通过该方法计算缩放比率和视频可显示范围
Private Sub cbrMain_ResizeClient(ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long)
    
    
    mVideoArea.Height = Bottom - Top - 4 * pbxSize(0).Height
    mVideoArea.Width = Right - Left - 4 * pbxSize(2).Width
    mVideoArea.Left = Left
    mVideoArea.Top = Top
    
    '计算缩放比率
    Call CalcVideoZoomRate

    '配置视频显示范围
    Call ConfigVideoDisplay(wdmCapture)
    Call ConfigVideoDisplay(picCusVideo)
    
    '刷新视频显示
    If IsCustomCaptureWay Then
        If Not (mobjCustomDevice Is Nothing) Then
            Call mobjCustomDevice.UpdateWindow(picCusVideo.ScaleWidth, picCusVideo.ScaleHeight)
        End If
    Else
        If Not (mVideoCapture Is Nothing) Then
            Call mVideoCapture.RefreshVideoWindow
        End If
    End If
    
    '刷新DcmView中的图像显示位置
    If dcmView.Images.count > 0 Then
        Call subCenterZoom(Me, dcmView.Images(1), dcmView, dcmView.Images(1).ActualZoom, mCorpSize)
    End If

    '刷新裁剪边线位置
    Call RefreshPbxSizePos
        
    
    If IsTwainCaptureWay Then
      
        '调整图像的显示范围
        dcmView.Left = Left
        dcmView.Top = Top
        dcmView.Width = Right - Left
        dcmView.Height = Bottom - Top
  
        '刷新DcmView中图像的显示位置
        If dcmView.Images.count > 0 Then
            Call subCenterZoom(Me, dcmView.Images(1), dcmView, dcmView.Images(1).ActualZoom, mCorpSize)
        End If
    
    End If
 
    If pbxSize(0).Top - picLock.Height < Top Then
        picLock.Top = Top
    Else
        picLock.Top = pbxSize(0).Top - picLock.Height
    End If
    
    picLock.Left = Left + ((Right - Left) - picLock.Width) / 2
    
    
    If pbxSize(1).Top + picAfter.Height > Bottom - Top Then
        picAfter.Top = Bottom - picAfter.Height
    Else
        picAfter.Top = pbxSize(1).Top
    End If
    
    picAfter.Left = Left + ((Right - Left) - picAfter.Width) / 2
    
End Sub


Private Sub cbrMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
On Error GoTo errHandle
    zlUpdateCommandBars Control
Exit Sub
errHandle:
    cbrMain.Options.UpdatePeriod = 2147483647
    MsgboxCus "菜单状态更新异常:" & err.Description, vbOKOnly, "提示"
End Sub

Private Sub DoListenCommSingal()
    Dim blnIsTouch As Boolean
    Dim lngTickCount As Long
    
    blnIsTouch = False
    
    lngTickCount = GetTickCount - mcpsComState.lngStartTick
'    strTouch = "间隔:" & Lpad(lngTickCount, 7) & "      序号:" & mcpsComState.lngSignalCount & "      "
    
    If lngTickCount <= gobjCapturePar.ComInterval Then
        mcpsComState.lngStartTick = GetTickCount
        mcpsComState.lngSignalCount = 1
    Else
        mcpsComState.lngSignalCount = mcpsComState.lngSignalCount + 1
        
        If GetTickCount - mcpsComState.lngEndTick > gobjCapturePar.ComInterval Then
            '清除因干扰产生的信号计数
            mcpsComState.lngSignalCount = 1
            mcpsComState.lngEndTick = GetTickCount
'            BUGEX ">>             ", True
        End If
        
        If mcpsComState.lngSignalCount >= gobjCapturePar.ComSignalCount Then
            '判断是否指定时间内接受到对应的信号数
            blnIsTouch = True
            
            mcpsComState.lngSignalCount = 0
            mcpsComState.lngStartTick = GetTickCount
            mcpsComState.lngEndTick = mcpsComState.lngStartTick
        End If
        
    End If
    
    If blnIsTouch = True And Not mblnReadOnly Then
'        BUGEX "**********************脚踏踩下*********************", True
        commListener.PortOpen = False
        
        If mstrAfterCapTag = "" Then
            Call ForeCapture(True)
        Else
            Call AfterCapture
        End If
        
        commListener.PortOpen = True
    End If
End Sub


Private Sub DoListenCommStateData()

    Dim strInput As String
    

    strInput = ""
    If commListener.InBufferCount > 0 Then strInput = commListener.Input
    
    If Not (commListener.CommEvent = comEvCTS Or commListener.CommEvent = comEvDSR _
        Or commListener.CommEvent = comEvCD Or commListener.CommEvent = comEvRing Or strInput <> "" _
        Or commListener.CommEvent = comEvSend Or commListener.CommEvent = comEvReceive) Then Exit Sub
        
    
    If gobjCapturePar.CaptureWay = 1 Then '转换触发
        If mcpsComState.intComState <> commListener.CommEvent Then
           '如果累计时间超过了采图时间间隔，则采集图像
           If mcpsComState.lngComTime > gobjCapturePar.ComInterval Then
               'If Me.cbrMain.FindControl(, conMenu_Cap_MarkMap).Enabled Then
               If Not mblnReadOnly Then
                    If mstrAfterCapTag = "" Then
                        Call ForeCapture(True)
                    Else
                        Call AfterCapture
                    End If
               End If
           End If
           
           '记录新的COM状态，计时器清零，启动timer
           mcpsComState.intComState = commListener.CommEvent
           mcpsComState.lngComTime = 0
           
           tmrComm.Enabled = True
        End If
    ElseIf gobjCapturePar.CaptureWay = 0 Then   '直接触发
        '两次踩下脚踏的时间间隔不能少于3秒
        If DateDiff("S", mcpsComState.dtLastCapture, time) < gobjCapturePar.ComInterval Then
            mcpsComState.dtLastCapture = time
            
            Exit Sub
        End If
        
        mcpsComState.dtLastCapture = time
        
        If Not mblnReadOnly Then
            If mstrAfterCapTag = "" Then
                Call ForeCapture(True)
            Else
                Call AfterCapture
            End If
        End If
    Else    '电平触发
        '对于电平触发的情况，当踩下脚踏的时候，对应线的电平会出现（低-高-低）或（高-低-高）的变化
        '通过电平变化，可以确定是否踩了脚踏。
        '当出现电流干扰时，虽然会出现OnComm事件，但是电平不会发生变化。
        '通过判断当前电平跟常态电平是否相同来确定电平是否发生了变化。
        
        '判断电平是否改变，判断CTS线
        If mcpsComState.blnCTSHolding <> commListener.CTSHolding Then
            '过滤振荡，毛刺现象，判断两次触发的时间是否小于设定的间隔
            If DateDiff("S", mcpsComState.dtLastCapture, time) < gobjCapturePar.ComInterval Then
                mcpsComState.dtLastCapture = time
                
                Exit Sub
            End If
            
            mcpsComState.dtLastCapture = time
            
            If Not mblnReadOnly Then
                If mstrAfterCapTag = "" Then
                    Call ForeCapture(True)
                Else
                    Call AfterCapture
                End If
            End If
        End If
    End If
End Sub


Private Sub commListener_OnComm()
On Error GoTo errHandle

    '如果是TWAIN扫描或专用视频采集，则不支持脚踏开关采集
    If IsTwainCaptureWay Or IsCustomCaptureWay Then Exit Sub
    
    If gobjCapturePar.ComPortType <> "COM" Then Exit Sub
    
    If gobjCapturePar.ComSignalCount > 0 Then
        '检测脚踏信号情况的方式进行采集
        Call DoListenCommSingal
    Else
        '检测脚踏状态数据情况的方式进行采集
        Call DoListenCommStateData
    End If
    
    Exit Sub
errHandle:
    If commListener.PortOpen = False Then commListener.PortOpen = True
    MsgboxCus err.Description, vbOKOnly, G_STR_HINT_TITLE
End Sub


Private Sub dcmView_DblClick()
On Error GoTo errHandle
    Call PlayCurVideo
Exit Sub
errHandle:
    MsgboxCus err.Description, vbOKOnly, G_STR_HINT_TITLE
End Sub


'modify by tjh at 2010-01-20
'计算视频缩放比率
Private Sub CalcVideoZoomRate()
  If mVideoSize.Width = 0 Or mVideoSize.Height = 0 Then
    mdblZoomRate = 1
    Exit Sub
  End If
  
    
  If (mVideoArea.Height <= 0) Or (mVideoArea.Width <= 0) Then
    mdblZoomRate = 1
    Exit Sub
  End If
  
  '计算缩放比率
  If (mVideoArea.Height / mVideoArea.Width) > (mVideoSize.Height / mVideoSize.Width) Then
    mdblZoomRate = mVideoArea.Width / mVideoSize.Width
  Else
    mdblZoomRate = mVideoArea.Height / mVideoSize.Height
  End If
  
End Sub


'modify by tjh at 2010-01-20
'配置视频显示
Private Sub ConfigVideoDisplay(videoObj As Object)
  '边框大小
  Const DICOM_VIEWER_BODER_SIZE As Long = 5
  
  If mVideoSize.Width = 0 Or mVideoSize.Height = 0 Then Exit Sub
  If (mVideoArea.Height <= 0) Or (mVideoArea.Width <= 0) Then Exit Sub

  
  '设置视频显示范围
  videoObj.Top = 0 - mdblZoomRate * mVideoSize.Height * mCurCutRange.TopRate
  videoObj.Height = mdblZoomRate * mVideoSize.Height
  picView.Height = mdblZoomRate * mVideoSize.Height * (1 - mCurCutRange.TopRate - mCurCutRange.HeightRate)
  
  videoObj.Left = 0 - mdblZoomRate * mVideoSize.Width * mCurCutRange.LeftRate
  videoObj.Width = mdblZoomRate * mVideoSize.Width
  picView.Width = mdblZoomRate * mVideoSize.Width * (1 - mCurCutRange.LeftRate - mCurCutRange.WidthRate)
  
  'picview的width和right是右端和下端的裁剪范围设置，不能随意改变
  
  If mVideoArea.Width <= picView.Width + pbxSize(2).Width Then
    picView.Left = mVideoArea.Left + pbxSize(2).Width * 2
  Else
    picView.Left = mVideoArea.Left + (mVideoArea.Width - picView.Width - 2 * pbxSize(2).Width) / 2 + 3 * pbxSize(2).Width
  End If
  
  If mVideoArea.Height <= picView.Height + pbxSize(0).Height Then
    picView.Top = mVideoArea.Top + pbxSize(0).Height * 2
  Else
    picView.Top = mVideoArea.Top + (mVideoArea.Height - picView.Height - 2 * pbxSize(0).Height) / 2 + 3 * pbxSize(2).Width
  End If
  
  '设置DICOM显示图像的大小
  dcmView.Left = DICOM_VIEWER_BODER_SIZE
  dcmView.Top = DICOM_VIEWER_BODER_SIZE
  dcmView.Width = picView.Width - DICOM_VIEWER_BODER_SIZE * 2
  dcmView.Height = picView.Height - DICOM_VIEWER_BODER_SIZE * 2
End Sub


Private Sub ConfigTwainDisplay()
  '边框大小
  Const DICOM_VIEWER_BODER_SIZE As Long = 5
  
  dcmView.Left = DICOM_VIEWER_BODER_SIZE
  dcmView.Top = DICOM_VIEWER_BODER_SIZE
  dcmView.Width = picView.Width - DICOM_VIEWER_BODER_SIZE * 2
  dcmView.Height = picView.Height - DICOM_VIEWER_BODER_SIZE * 2
End Sub


Public Sub HideBorder()
    '隐藏窗口的标题框
    Dim lngWindowStyle As Long
    
    lngWindowStyle = GetWindowLong(Me.hwnd, GWL_STYLE)
    lngWindowStyle = lngWindowStyle And Not (WS_SYSMENU Or WS_CAPTION Or WS_MAXIMIZEBOX Or WS_MINIMIZEBOX Or WS_THICKFRAME)
    
    Call SetWindowLong(Me.hwnd, GWL_STYLE, lngWindowStyle Or WS_CHILD)
End Sub

Private Sub OpenVideoCaptureDevice()
'打开视频采集设备
    Dim blnIsStartupVideo As Boolean
    Dim lngCusWidth As Long
    Dim lngCusHeight As Long

BUGEX "OpenVideoCaptureDevice 1"

    If mVideoCapture Is Nothing Then
        '创建视频采集对象
        Set mVideoCapture = New clsVideoCapture
        
        '连接视频相关组件
        Call mVideoCapture.ConnectedVfwDeviceObj(picCusVideo)
        Call mVideoCapture.ConnectedWdmDeviceObj(wdmCapture)
        Call mVideoCapture.ConnectedCustomDeviceObj(mobjCustomDevice)
        
        '读取配置文件
        Call mVideoCapture.ReadCaptureParameterFromFile(GetAppPath & "\" & CAPTURE_PARAMETER_CONFIG_FILE_NAME)

        '设置视频的显示模式
        Call mVideoCapture.SetVideoShowWay(swStretch)

        '在读取文件配置后修改该属性（只有设置该属性，才能根据四条边框进行调节和显示）
        wdmCapture.AppHandle = Me.hwnd
        wdmCapture.IsShowState = False

        mdblZoomRate = 1
    End If
    
    mstrVideoRegTime = funVideoRegTime(Me)
    If mstrVideoRegTime = "" Then mstrMsg = "视频源不允许启动，请联系管理员到服务管理工具中进行配置！"
    
    If UCase(Command()) = "DEBUG" Then
        mstrVideoRegTime = Now
    End If
    
    '设置视频驱动类型
    mVideoCapture.VideoDriverType = gobjCapturePar.VideoDirverType
        
    If (Not mVideoCapture.IsStartup) Then
        
        '读取视频大小
        mVideoSize.Width = ScaleX(mVideoCapture.VideoSize.Width, vbPixels, vbTwips)
        mVideoSize.Height = ScaleX(mVideoCapture.VideoSize.Height, vbPixels, vbTwips)
        
        '配置界面
        Call CaptureSwitchFace(IsTwainCaptureWay) 'Or IsCustomCaptureWay
       

        '*******************************************************
BUGEX "OpenVideoCaptureDevice 5"
        '开始视频预览********************************************
        If Not IsTwainCaptureWay And Not IsCustomCaptureWay Then
            mblnRealTime = True
            
            Call mVideoCapture.StartPreview
                    
'            blnIsStartupVideo = mVideoCapture.IsStartup
        ElseIf IsCustomCaptureWay Then
            '专用采集
            mblnRealTime = True
            
            Call mobjCustomDevice.GetSizeInfo(lngCusWidth, lngCusHeight)
            
            mVideoSize.Width = ScaleX(lngCusWidth, vbPixels, vbTwips)
            mVideoSize.Height = ScaleX(lngCusHeight, vbPixels, vbTwips)
'            blnIsStartupVideo = True
        Else
            'twain采集
            mblnRealTime = False
            
'            blnIsStartupVideo = ImageScanner.ScannerAvailable
        End If
 

        '*********************************************************
    Else
        Call ConfigVideoShowState(True)
    End If
    
    Call OpenComm   '打开采集端口
End Sub


Public Sub ResetAfterCaptureTag()
'更新后台采集信息
    Call mobjCapHelper.AfterTag(mstrAfterCapTag)
    
    labAfterInfo.Caption = "标识:" & mstrAfterCapTag
    labAfterInfo.Visible = True
    picAfter.Visible = True
    
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If (Shift And vbCtrlMask) = 0 Then Exit Sub
    If (Shift And vbShiftMask) = 0 Then Exit Sub
    If (Shift And vbAltMask) = 0 Then Exit Sub
    
    If KeyCode <> 86 Then Exit Sub
    
    Call ShowVideoConfig
End Sub

Private Sub Form_Load()
  On Error GoTo errHandle
    '设置窗口样式
    '在这里必须对该窗口对象进行置顶操作，否则在执行打开或者保存操作时，弹出的文件选择框将位于该窗口之后
    SetWindowPos Me.hwnd, -1, Me.CurrentX, Me.CurrentY, Me.ScaleWidth, Me.ScaleHeight, 3  '将窗口置顶
     
    Call InitCommandBars
    Call InitScanDir
    
    Set mfrmParameter = New frmVideoSetupV2
    
    Exit Sub
errHandle:
    MsgboxCus err.Description, vbOKOnly, G_STR_HINT_TITLE
End Sub

Private Sub InitScanDir()
    mstrTempDirOfScan = GetAppPath + CONST_STR_DEFAULT_TEMP_SCAN_DIR_NAME
    If Len(mstrTempDirOfScan) > 45 Then
        Dim strFolder As String
        Dim pathlen As Long
        
        strFolder = String(256, 0)
        pathlen = GetTempPath(256, strFolder)
        If pathlen > 0 Then
            mstrTempDirOfScan = Left(strFolder, pathlen - 1) + CONST_STR_DEFAULT_TEMP_SCAN_DIR_NAME
        End If
    End If
End Sub


'返回是否为TWAIN的采集方式
Private Function IsTwainCaptureWay() As Boolean
    IsTwainCaptureWay = IIf(gobjCapturePar.VideoDirverType = vdtTWAIN, True, False)
End Function

Private Function IsCustomCaptureWay() As Boolean
    IsCustomCaptureWay = IIf(gobjCapturePar.VideoDirverType = vdtCustom, True, False)
End Function

'配置TWAIN时的采集界面
Private Sub CaptureSwitchFace(ByVal blnUseTwain As Boolean)
    '去掉和TWAIN扫描不相关的一些界面配置
    
    dcmView.Visible = blnUseTwain
    picView.Visible = Not blnUseTwain
    
    pbxSize(0).Visible = Not blnUseTwain
    pbxSize(1).Visible = Not blnUseTwain
    pbxSize(2).Visible = Not blnUseTwain
    pbxSize(3).Visible = Not blnUseTwain
        
    wdmCapture.Visible = False
    picCusVideo.Visible = False
      
    If blnUseTwain Then
      Set dcmView.Container = Me
      Set txtInputText.Container = Me
    Else
      Set dcmView.Container = picView
      Set txtInputText.Container = picView
    End If
    
    Call ConfigVideoShowState(True)
    
    Dim lngLeft As Long, lngTop As Long, lngRight As Long, lngBottom As Long
    
    Call cbrMain.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)
    Call cbrMain_ResizeClient(lngLeft, lngTop, lngRight, lngBottom)
End Sub


'更新采集的驱动程序
Private Sub UpdateCaptureDirver(ByVal videoDirver As TVideoDriverType)

    '先停止视频的预览
    If IsCustomCaptureWay Then
        If Not (mobjCustomDevice Is Nothing) Then
            mobjCustomDevice.StopPreview
        End If
    Else
        Call mVideoCapture.StopPreview
    End If
    
    gobjCapturePar.VideoDirverType = videoDirver
    mVideoCapture.VideoDriverType = videoDirver
       
    '读取视频大小
    mVideoSize.Width = ScaleX(mVideoCapture.VideoSize.Width, vbPixels, vbTwips)
    mVideoSize.Height = ScaleX(mVideoCapture.VideoSize.Height, vbPixels, vbTwips)
       
    Call CaptureSwitchFace(videoDirver = vdtTWAIN)
        
    
    '如果不是Twain采集方式，则重新启动预览
    If videoDirver <> vdtTWAIN And videoDirver <> vdtCustom Then
        mblnRealTime = True
      
        '开始预览
        Call mVideoCapture.StartPreview
        
        '刷新视频预览窗口
        Call mVideoCapture.RefreshVideoWindow
    Else
        If videoDirver = vdtCustom Then
            '初始化专用视频采集接口
            'Call InitCustomDevice
            If Not (mobjCustomDevice Is Nothing) Then
                mobjCustomDevice.StartPreview
                Call mobjCustomDevice.UpdateWindow(picCusVideo.ScaleWidth, picCusVideo.ScaleHeight)
            End If
        End If
        
        mblnRealTime = True
    End If
End Sub


'保存当前参数设置
Private Sub SaveParameterCfg()
BUGEX "SaveParameterCfg 1"
    
  '裁剪参数设置
  SaveSetting "ZLSOFT", G_STR_REG_PATH_PUBLIC, "mdblX1Scale", mCurCutRange.LeftRate
  SaveSetting "ZLSOFT", G_STR_REG_PATH_PUBLIC, "mdblX2Scale", mCurCutRange.WidthRate
  SaveSetting "ZLSOFT", G_STR_REG_PATH_PUBLIC, "mdblY1Scale", mCurCutRange.TopRate
  SaveSetting "ZLSOFT", G_STR_REG_PATH_PUBLIC, "mdblY2Scale", mCurCutRange.HeightRate
  
BUGEX "SaveParameterCfg 2"
        
  '保存采集参数
  If Not mVideoCapture Is Nothing Then Call mVideoCapture.SaveCaptureParameterToFile(GetAppPath & "\" & CAPTURE_PARAMETER_CONFIG_FILE_NAME)
BUGEX "SaveParameterCfg 3"
End Sub


Private Sub OpenComm()
    On Error GoTo err
    
BUGEX "OpenComm 1"
BUGEX "OpenComm ComPortType:" & gobjCapturePar.ComPortType
    If gobjCapturePar.ComPortType = "无" Then Exit Sub
BUGEX "OpenComm 2"
    If gobjCapturePar.ComPortType = "COM" Then
BUGEX "OpenComm 3"
        If commListener.PortOpen Then Exit Sub
BUGEX "OpenComm 4"
        commListener.CommPort = gobjCapturePar.ComPortName
        commListener.Settings = "9600,N,8,1"
        commListener.InputMode = comInputModeText
        commListener.RThreshold = 1
        commListener.InBufferCount = 0
        commListener.InputLen = 0
        commListener.RTSEnable = True
        commListener.DTREnable = True
        commListener.EOFEnable = False
                        
        commListener.PortOpen = True
BUGEX "OpenComm 5"
        '记录常态电平电位
        mcpsComState.blnCTSHolding = commListener.CTSHolding
BUGEX "OpenComm 6"
    Else
BUGEX "OpenComm 7"
        If mobjDxDevice Is Nothing Then
BUGEX "OpenComm 7.1"
            Set mobjDxDevice = New clsDxHidDevice
        Else
BUGEX "OpenComm 7.2"
        End If
BUGEX "OpenComm 8"
        '打开DX设备
        If mobjDxDevice.Handle = 0 Then Call mobjDxDevice.OpenDxDevice(gobjCapturePar.ComPortName)
BUGEX "OpenComm 9"
        tmrComm.Enabled = True
        tmrComm.Interval = 20
    End If
BUGEX "OpenComm 10"
    Exit Sub
err:
BUGEX "OpenComm 11"
    Call MsgboxCus("端口打开错误", vbOKOnly, G_STR_HINT_TITLE)
BUGEX "OpenComm 12"
End Sub


Private Sub dcmView_MouseDown(Button As Integer, Shift As Integer, X As Long, Y As Long)
    If Button = 1 And dcmView.Images.count > 0 Then
        Dim intLabelType As Integer
        
        If mintMouseState = 13 And txtInputText.Text <> "" And txtInputText.Visible Then
            If Not mdcmSelectLabel Is Nothing Then mdcmSelectLabel.Text = txtInputText.Text
        End If
        
        mMouseDownPoint.X = dcmView.Images(1).ActualScrollX
        mMouseDownPoint.Y = dcmView.Images(1).ActualScrollY
          
        mInitScrollPoint.X = dcmView.Images(1).ScrollX + X
        mInitScrollPoint.Y = dcmView.Images(1).ScrollY + Y
        
        mblnDcmViewDown = True
        If mintMouseState <> 0 Then
            '记录当前鼠标位置
            mlngBaseX = X
            mlngBaseY = Y
            
            Select Case mintMouseState
                'Case 14  '图像拖动
                
                Case 11, 12, 13, 3    '箭头，椭圆，文字,框选
                    If mintMouseState = 11 Then
                        intLabelType = doLabelArrow
                    ElseIf mintMouseState = 12 Then
                        intLabelType = doLabelEllipse
                    ElseIf mintMouseState = 13 Then
                        intLabelType = doLabelText
                    ElseIf mintMouseState = 3 Then
                        intLabelType = doLabelRectangle
                    End If
                    
                    dcmView.Images(1).Labels.Add GetNewLabel(intLabelType, dcmView.ImageXPosition(X, Y), dcmView.ImageYPosition(X, Y), 0, 0)
                    
                    Set mdcmSelectLabel = dcmView.Images(1).Labels(dcmView.Images(1).Labels.count)
                    
                    If mintMouseState = 3 Then
                        mdcmSelectLabel.Tag = "CUT"
                    End If
                    
                    mdcmSelectLabel.LineWidth = 2
            End Select
        End If
    End If
End Sub


Private Sub dcmView_MouseMove(Button As Integer, Shift As Integer, X As Long, Y As Long)
    Dim dblZoom As Double
    
    If mblnDcmViewDown = True And Button = 1 And dcmView.Images.count > 0 Then
        Select Case mintMouseState
            Case 1  '亮度对比度
                dcmView.Images(1).Width = dcmView.Images(1).Width + (X - mlngBaseX)
                dcmView.Images(1).Level = dcmView.Images(1).Level + (Y - mlngBaseY)
                
                mlngBaseX = X
                mlngBaseY = Y
            Case 2  '缩放
                dblZoom = dcmView.Images(1).ActualZoom
                dblZoom = dblZoom * (1 + (Y - mlngBaseY) * 0.001)
                
                If dblZoom < 64 And dblZoom > 0.01 Then
                    subCenterZoom Me, dcmView.Images(1), dcmView, dblZoom, mCorpSize
                End If
                mlngBaseY = Y
'            Case 3  '裁剪缩放
'                Dim dcmLabel As DicomLabel
'                dcmView.Labels.Clear
'                Set dcmLabel = dcmView.Labels.AddNew
'                dcmLabel.LabelType = doLabelRectangle
'                dcmLabel.Left = mlngBaseX
'                dcmLabel.Top = mlngBaseY
'                dcmLabel.Width = x - mlngBaseX
'                dcmLabel.Height = y - mlngBaseY
            Case 11, 12, 3 '箭头标注'圆形标注,框选
                mdcmSelectLabel.Width = dcmView.ImageXPosition(X, Y) - mdcmSelectLabel.Left
                mdcmSelectLabel.Height = dcmView.ImageYPosition(X, Y) - mdcmSelectLabel.Top
            Case 14
                '拖动图像......
                dcmView.Images(1).ScrollX = mInitScrollPoint.X - X
                dcmView.Images(1).ScrollY = mInitScrollPoint.Y - Y
        End Select
        dcmView.Refresh
    End If
End Sub


Private Sub dcmView_MouseUp(Button As Integer, Shift As Integer, X As Long, Y As Long)
    If mblnDcmViewDown = True And Button = 1 And dcmView.Images.count > 0 Then
        mblnDcmViewDown = False
        If mintMouseState = 13 Then     '文字标注
            txtInputText.Left = X * Screen.TwipsPerPixelX
            txtInputText.Top = Y * Screen.TwipsPerPixelY
            txtInputText.Text = ""
            txtInputText.Visible = True
            txtInputText.SetFocus
            
        ElseIf mintMouseState = 3 Then  '裁剪缩放
            
            '显示图像保存菜单
            Call ShowFrameSelectImagePopup
            '删除框选用的临时标注
            If dcmView.Images(1).Labels.count > 0 Then
                If dcmView.Images(1).Labels(dcmView.Images(1).Labels.count).Tag = "CUT" Then
                    dcmView.Images(1).Labels.Remove dcmView.Images(1).Labels.count
                End If
            End If
            
            Set mdcmSelectLabel = Nothing
            
            
'            dcmView.Labels.Clear
            
'            dcmView.Labels.Clear
'            RectangleZoom dcmView, dcmView.Images(1), mlngBaseX, mlngBaseY, x - mlngBaseX, y - mlngBaseY
        ElseIf mintMouseState = 14 Then
            '计算图像漫游的偏移位置
            mCorpSize.X = mCorpSize.X + (dcmView.Images(1).ActualScrollX - mMouseDownPoint.X)
            mCorpSize.Y = mCorpSize.Y + (dcmView.Images(1).ActualScrollY - mMouseDownPoint.Y)
        End If
        
        dcmView.Refresh
    End If
End Sub

   
Private Function DoNormalCapture(ByVal blnIsReal As Boolean, Optional objPic As StdPicture = Nothing) As Boolean
'------------------------------------------------
'功能：采集并存储图像
'参数：无
'返回：无，直接保存新采集的图像
'------------------------------------------------
On Error GoTo errHandle
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim objImg As DicomImage
    Dim strError As String
    
    DoNormalCapture = False
    
    If mstrVideoRegTime = "" Then
        MsgboxCus mstrMsg, vbOKOnly, "提示"
        Exit Function
    End If
    
     
    Set objImg = ConvertDcmImage(strError, blnIsReal, "", objPic)
    
    If Not objImg Is Nothing Then
         
        mintCaptureFlag = 2
        
        DoNormalCapture = mobjCapHelper.SaveImg(objImg, "", True)
    Else
        MsgboxCus strError, vbOKOnly, "提示"
    End If
Exit Function
errHandle:
    err.Raise err.Number, err.Description
End Function

Private Function DoScanCaptureDown(ByVal strFile As String) As Boolean
On Error GoTo errHandle
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim objImg As DicomImage
    Dim strError As String
    
    DoScanCaptureDown = False
    
    If mstrVideoRegTime = "" Then
        MsgboxCus mstrMsg, vbOKOnly, "提示"
        Exit Function
    End If
    
     
    Set objImg = ConvertDcmImage(strError, , strFile)
    
    If Not objImg Is Nothing Then
         
        mintCaptureFlag = 2
        
        DoScanCaptureDown = mobjCapHelper.SaveImg(objImg, "", True)
        
        If DoScanCaptureDown Then Call ShowScanImage(objImg)
    Else
        MsgboxCus strError, vbOKOnly, "提示"
    End If
Exit Function
errHandle:
    err.Raise err.Number, err.Description
End Function
 


Private Function DoCustomCapture() As Boolean
    Dim objCapPic As StdPicture
    Dim strCapImgFile As String
    Dim blnIsCusSave As Boolean
    Dim objImg As DicomImage
    Dim strError As String
    
    DoCustomCapture = False
    
    If mobjCustomDevice Is Nothing Then Exit Function
    
    '采集图像
    If Not mobjCustomDevice.zlCaptureImage(mobjCapHelper.GetCustomMainID, _
        objCapPic, strCapImgFile, blnIsCusSave) Then
        Exit Function
    End If
    
    Set objImg = ConvertDcmImage(strError, , strCapImgFile, objCapPic)
    If Not objImg Is Nothing Then
        DoCustomCapture = mobjCapHelper.SaveImg(Nothing, "", Not blnIsCusSave)
    Else
        MsgboxCus strError, vbOKOnly, "提示"
    End If
End Function

Public Sub AfterCapture()
'------------------------------------------------
'功能：后台采集
'参数：无
'返回：无，直接保存新采集的图像
'------------------------------------------------
On Error GoTo errHandle
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim objImg As DicomImage
    Dim strError As String
    
    If mstrVideoRegTime = "" Then
        MsgboxCus mstrMsg, vbOKOnly, "提示"
        Exit Sub
    End If
     
    If Not mVideoCapture.IsStartup Then
        MsgboxCus "视频源尚未启动不能采集。", vbOKOnly, "提示"
        Exit Sub
    End If
    
    '如果没有后台采集标识，则产生新的后台采集标识
    If Len(mstrAfterCapTag) <= 0 Then Call ResetAfterCaptureTag
     
    Set objImg = ConvertDcmImage(strError, True)
    
    If Not objImg Is Nothing Then
         
        mintCaptureFlag = 2
        
        Call mobjCapHelper.SaveImg(objImg, "", True, mstrAfterCapTag)
    Else
        MsgboxCus strError, vbOKOnly, "提示"
    End If
Exit Sub
errHandle:
    err.Raise err.Number, err.Description
End Sub

Private Function PictureToDicomImg(ByVal lngHDC As Long, ByVal lngPictureHandle As Long, _
    objDcmImg As Object, ByRef strError As String) As Boolean
'congpicture中复制图像到dicomimage
    Const bitCount As Long = 3
        
    Dim hBitmap As OLE_HANDLE
    Dim stucbmp As TBitMap
    Dim lngSize As Long
    Dim lngResult As Long
    Dim aryPixels() As Byte
    Dim stuDipInf As BITMAPINFO
    
    Dim i As Long, j As Long, bytTemp As Byte
    
On Error GoTo errHandle
    PictureToDicomImg = False
    hBitmap = lngPictureHandle
    
    '使用GetObject方法将获取32位的格式头信息
    lngResult = GetObject(hBitmap, Len(stucbmp), stucbmp)
    If lngResult = 0 Then Exit Function
    
    
    While stucbmp.bmWidth * 3 Mod 4 <> 0
        '当使用GetDIBits函数时，每一行所需的字节数必须是4的倍数，即按4字节对齐
        stucbmp.bmWidth = stucbmp.bmWidth - 1
    Wend
    
    '按照24位图像格式计算图像的存储大小，以字节为单位
    lngSize = stucbmp.bmWidth * 3 * stucbmp.bmHeight 'stucbmp.bmWidthBytes * stucbmp.bmHeight
    
    stuDipInf.bmiHeader.biSize = 40
    stuDipInf.bmiHeader.biHeight = -stucbmp.bmHeight
    stuDipInf.bmiHeader.biPlanes = stucbmp.bmPlanes
    stuDipInf.bmiHeader.biBitCount = 24 'bmp.bmBitsPixel  '强制使用24位格式，便于后面计算和转换
    stuDipInf.bmiHeader.biWidth = stucbmp.bmWidth
    stuDipInf.bmiHeader.biCompression = BI_RGB
    stuDipInf.bmiHeader.biSizeImage = lngSize
    stuDipInf.bmiHeader.biXPelsPerMeter = 0
    stuDipInf.bmiHeader.biYPelsPerMeter = 0
    stuDipInf.bmiHeader.biClrUsed = 0
    stuDipInf.bmiHeader.biClrImportant = 0
    stuDipInf.bmiColors(0).rgbBlue = 8
    stuDipInf.bmiColors(0).rgbGreen = 8
    stuDipInf.bmiColors(0).rgbRed = 8
    stuDipInf.bmiColors(0).rgbReserved = 0
    
'    ReDim aryPixels(1 To stucbmp.bmWidthBytes, 1 To stucbmp.bmHeight, 1 To 1)
    ReDim aryPixels(1 To stucbmp.bmWidth * 3, 1 To stucbmp.bmHeight, 1 To 1)

'    lngResult = GetBitmapBits(hBitmap, lngSize, aryPixels(1, 1, 1))

    '只能使用该函数获取24位的像素格式，如果使用GetBitmapBits，获取的将是32位的图像格式
    lngResult = GetDIBits(lngHDC, hBitmap, 0, stucbmp.bmHeight, aryPixels(1, 1, 1), stuDipInf, DIB_RGB_COLORS)
    If lngResult = 0 Then Exit Function
    

    '将bmp的brg存储格式转换为dicom的rgb存储格式
    For i = 1 To stucbmp.bmWidth * 3 Step 3
        For j = 1 To stucbmp.bmHeight
            bytTemp = aryPixels(i + 2, j, 1)
            aryPixels(i + 2, j, 1) = aryPixels(i, j, 1)
            aryPixels(i, j, 1) = bytTemp
        Next j
    Next i

    '构造dicom的图像格式
    objDcmImg.Attributes.Add &H28, &H2, 3       'stucbmp.bmBitsPixel        'samples per pixel
    objDcmImg.Attributes.Add &H28, &H4, "RGB"                  'Photometric Interpretation
    objDcmImg.Attributes.Add &H28, &H6, 0                      'planar configuration
    objDcmImg.Attributes.Add &H28, &H100, 8                    'Bits Allocated
    objDcmImg.Attributes.Add &H28, &H101, 8                    'Bits Stored
    objDcmImg.Attributes.Add &H28, &H102, 7                    'High Bit
    objDcmImg.Attributes.Add &H28, &H103, 0                    'Pixel Representation
    objDcmImg.Attributes.Add &H28, &H10, stucbmp.bmHeight          'rows
    objDcmImg.Attributes.Add &H28, &H11, stucbmp.bmWidth           'columns
    
    objDcmImg.Pixels = aryPixels
    objDcmImg.InstanceUID = dcmglbUID.NewUID

    PictureToDicomImg = True
Exit Function
errHandle:
    strError = err.Description
End Function


Private Function ConvertDcmImage(ByRef strError As String, _
                        Optional ByVal blnRealState As Boolean = True, _
                        Optional ByVal strFileName As String = "", _
                        Optional objCapture As StdPicture = Nothing) As DicomImage
'------------------------------------------------
'功能：采集单帧视频图像，将图像转换成DICOM格式，并填写DICOM文件头，最后将图像放入缩略图dcmMiniature中。
'参数：无
'返回：无，直接将新采集的图像放入dcmMiniature中
'------------------------------------------------
'采集单帧图像
On Error GoTo SaveFileError
    Dim ImgTmpImage As DicomImage
    Dim dcmTag As clsImageTagInf
    Dim strFile As String
    
    '采集图像，分为直接视频采集和播放录象采集
    Set ConvertDcmImage = Nothing

    If Not (objCapture Is Nothing) Then
        '从stdPicture读取图像
        Set picTemp2.Picture = Nothing
        
        picTemp2.Width = mVideoSize.Width * (1 - mCurCutRange.WidthRate - mCurCutRange.LeftRate)
        picTemp2.Height = mVideoSize.Height * (1 - mCurCutRange.HeightRate - mCurCutRange.TopRate)
        
        picTemp2.Picture = objCapture
        
    ElseIf Trim(strFileName) <> "" And Dir(strFileName) <> "" Then
        '从文件读取图像
        Set picTemp2.Picture = Nothing
        Set picTemp2.Picture = LoadPicture(strFileName)
        
    Else
        If blnRealState = False And mblnPlayVideo = False Then
            '使用dcmView显示的是图片，不需要再裁剪
            Set picTemp2.Picture = Nothing
            
            If dcmView.Images.count > 0 Then
                Set picTemp2.Picture = dcmView.CurrentImage.Capture(False).Picture
            End If
        Else
            '当处于实时视频显示时，需要对图像进行裁剪操作
            Set picTemp2.Picture = Nothing
                        
            Dim curPic As StdPicture
            Set curPic = mVideoCapture.CaptureImageFromMemory

            If curPic Is Nothing Then
                strError = "视频图像采集失败，请检查视频参数设置是否正确(如视频设备，显示模式等)。"
                Exit Function
            End If
            
            If mCurCutRange.LeftRate > 0.005 Or mCurCutRange.TopRate > 0.005 Then
                picTemp2.Width = mVideoSize.Width * (1 - mCurCutRange.WidthRate - mCurCutRange.LeftRate)
                picTemp2.Height = mVideoSize.Height * (1 - mCurCutRange.HeightRate - mCurCutRange.TopRate)
    
                '应用图像范围裁剪
                Call picTemp2.PaintPicture(curPic, 0, 0, picTemp2.Width, picTemp2.Height, _
                                           mVideoSize.Width * mCurCutRange.LeftRate, mVideoSize.Height * mCurCutRange.TopRate, _
                                           picTemp2.Width, picTemp2.Height, vbSrcCopy)
                                                   
                picTemp2.Picture = picTemp2.Image
            Else
                Set picTemp2.Picture = curPic
            End If

            Set curPic = Nothing
        End If
    End If
    
    '如果没有采集到图像，则直接退出
    If picTemp2.Picture Is Nothing Then
        strError = "未采集到图像."
        Exit Function
    End If
    
    '创建dicom格式图像
    Set ImgTmpImage = New DicomImage

    Select Case mlngImageSwapWay
        Case 0  '内存
            '不使用剪贴板方式，从Picture中复制图像到ImgTmpImage中,不使用剪贴板交换数据
            If PictureToDicomImg(picTemp2.hdc, picTemp2.Picture.Handle, ImgTmpImage, strError) = False Then
                Exit Function
            End If
        Case 1  '剪贴板
            If ClipboardToDicomImg(picTemp2.Picture, ImgTmpImage, dcmglbUID.NewUID, strError) = False Then
                Exit Function
            End If
        Case 2  '文件
            If FileToDicomImg(picTemp2.Picture, ImgTmpImage, strError) = False Then
                Exit Function
            End If
    End Select
    
    Set ConvertDcmImage = ImgTmpImage

    Exit Function
SaveFileError:
    MsgboxCus err.Description, vbOKOnly, G_STR_HINT_TITLE
End Function

Private Function FileToDicomImg(objPic As StdPicture, objDcmImg As Object, _
    ByRef strError As String) As Boolean
'文件到dicomimage
    Dim strFile As String
    
On Error GoTo errHandle
    FileToDicomImg = False

    strFile = mstrBufferDir & "ImageFile.SWAP"
    Call SavePicture(objPic, strFile)
    
    objDcmImg.FileImport strFile, "BMP"
    objDcmImg.InstanceUID = dcmglbUID.NewUID
    
    FileToDicomImg = True
Exit Function
errHandle:
    strError = err.Description
End Function


Private Sub Form_Resize()
On Error GoTo errHandle
    
    '设置图标大小
    If Me.ScaleHeight < 7000 Or Me.ScaleWidth < 4000 Then
        cbrMain.Options.SetIconSize True, 16, 16
    Else
        cbrMain.Options.SetIconSize True, 32, 32
    End If

'    cbrMain.RecalcLayout
    
errHandle:
End Sub

Private Sub Form_Terminate()
    Set mdcmSelectLabel = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim i As Long
    

BUGEX "VideoForm_UnLoad 1"
    tmrComm.Enabled = False
    
BUGEX "VideoForm_UnLoad 3"
    '先关闭采集窗口和COMM口
    Call StopCapture
BUGEX "VideoForm_UnLoad 4"
    '保持裁剪设置
    Call SaveParameterCfg
    
    
BUGEX "VideoForm_UnLoad 6"
    If Not mfrmParameter Is Nothing Then
        Unload mfrmParameter
    End If
    
BUGEX "VideoForm_UnLoad 8"
    wdmCapture.FreeRes
    
BUGEX "VideoForm_UnLoad 9"
    Set mobjCapHelper = Nothing
    
BUGEX "VideoForm_UnLoad 10"
    Set dcmglbUID = Nothing
    Set mobjDxDevice = Nothing
    Set mVideoCapture = Nothing
    Set mfrmParameter = Nothing
    
    If Not mobjCustomDevice Is Nothing Then
        mobjCustomDevice.zlFree
        Set mobjCustomDevice = Nothing
    End If
    
    If Not mobjPlayWindow Is Nothing Then
        Unload mobjPlayWindow
        Set mobjPlayWindow = Nothing
    End If
    
BUGEX "VideoForm_UnLoad End"
End Sub

 

Private Sub subSetMouseState(intMouseState As Integer)
    '改变当前鼠标状态
    mintMouseState = IIf(mintMouseState = intMouseState, 0, intMouseState)
    
    If txtInputText.Visible Then txtInputText.Visible = False
    
    cbrMain.FindControl(xtpControlButton, conMenu_ImgPro_Window, False, True).Checked = False
    cbrMain.FindControl(xtpControlButton, conMenu_ImgPro_Zoom, False, True).Checked = False
    cbrMain.FindControl(xtpControlButton, conMenu_ImgPro_RectSave, False, True).Checked = False
    cbrMain.FindControl(xtpControlButton, conMenu_ImgPro_DirectSave, False, True).Checked = False
    cbrMain.FindControl(xtpControlButton, conMenu_ImgPro_Arrow, False, True).Checked = False
    cbrMain.FindControl(xtpControlButton, conMenu_ImgPro_Ellipse, False, True).Checked = False
    cbrMain.FindControl(xtpControlButton, conMenu_ImgPro_Text, False, True).Checked = False
    cbrMain.FindControl(xtpControlButton, conMenu_ImgPro_Corp, False, True).Checked = False
'    cbrMain.FindControl(xtpControlSplitButtonPopup, conMenu_ImgPro_Lab_Pop, False, True).Checked = False
End Sub


'modify by tjh at 2010-01-20
'配置视频显示状态
Private Sub ConfigVideoShowState(ByVal blnShowState As Boolean)
  mblnRealTime = blnShowState
  
  Select Case gobjCapturePar.VideoDirverType
    Case vdtVFW, vdtCustom
      picCusVideo.Visible = blnShowState
      wdmCapture.Visible = False
      dcmView.Visible = Not blnShowState
    Case vdtWDM
      wdmCapture.Visible = blnShowState
      picCusVideo.Visible = False
      dcmView.Visible = Not blnShowState
    Case vdtTWAIN, vdtCustom
      wdmCapture.Visible = False
      picCusVideo.Visible = False
      mblnRealTime = False
      
      dcmView.Visible = True
  End Select
End Sub


'modify by tjh at 2010-01-20
'更新视频裁剪框的位置
Private Sub RefreshPbxSizePos()
  '上
  pbxSize(0).Top = picView.Top - pbxSize(0).Height + 5
  pbxSize(0).Left = picView.Left
  pbxSize(0).Width = picView.Width
  
  '下
  pbxSize(1).Top = picView.Top + picView.Height - 5
  pbxSize(1).Left = picView.Left
  pbxSize(1).Width = picView.Width
  
  '左
  pbxSize(2).Top = picView.Top - pbxSize(0).Height
  pbxSize(2).Left = picView.Left - pbxSize(2).Width + 5
  pbxSize(2).Height = picView.Height + pbxSize(0).Height * 2
  
  '右
  pbxSize(3).Top = picView.Top - pbxSize(0).Height
  pbxSize(3).Left = picView.Left + picView.Width - 5
  pbxSize(3).Height = picView.Height + pbxSize(0).Height * 2
  
  'pbxsize刷新显示
  Call pbxSize(0).Refresh
  Call pbxSize(1).Refresh
  Call pbxSize(2).Refresh
  Call pbxSize(3).Refresh
End Sub


'modify by tjh at 2010-01-20
'改变视频裁剪范围
Private Sub ChangeCutRanage(videoObj As Object, ByVal lngChangeIndex As Long, ByVal X As Long, ByVal Y As Long)
  Dim lngDistanceX As Long
  Dim lngDistanceY As Long
  
  lngDistanceX = X ' - mStartPoint.X
  lngDistanceY = Y ' - mStartPoint.Y
  
  
  Select Case lngChangeIndex
    Case moUp '上--------------------------------------------------
      If (picView.Height - lngDistanceY) <= 50 * mdblZoomRate Then Exit Sub
      If videoObj.Top - lngDistanceY > 0 Then Exit Sub  'lngDistanceY = 0
     
      videoObj.Top = videoObj.Top - lngDistanceY
      
      picView.Top = picView.Top + lngDistanceY
      picView.Height = (picView.Height - lngDistanceY)
    Case moDown '下--------------------------------------------------
      If (picView.Height + lngDistanceY) <= 50 * mdblZoomRate Then Exit Sub
      'If Abs(0 - DSCapture.Top) + Picturexx.Height >= mVideoSize.Height * mdblVZoomRate Then Exit Sub
            
      picView.Height = (picView.Height + lngDistanceY)
      
      If Abs(0 - videoObj.Top) + picView.Height >= mVideoSize.Height * mdblZoomRate Then
        picView.Height = (picView.Height - lngDistanceY)
      End If
    Case moLeft '左--------------------------------------------------
      If (picView.Width - lngDistanceX) <= 50 * mdblZoomRate Then Exit Sub
      If videoObj.Left - lngDistanceX > 0 Then Exit Sub ' lngDistanceX = 0
      
      videoObj.Left = videoObj.Left - lngDistanceX
      
      picView.Left = picView.Left + lngDistanceX
      picView.Width = picView.Width - lngDistanceX
    
    Case moRight '右--------------------------------------------------
      If (picView.Width + lngDistanceX) <= 50 * mdblZoomRate Then Exit Sub
      'If Abs(0 - DSCapture.Left) + Picturexx.Width >= mVideoSize.Width * mdblHZoomRate Then Exit Sub
            
      picView.Width = picView.Width + lngDistanceX
      
      If Abs(0 - videoObj.Left) + picView.Width >= mVideoSize.Width * mdblZoomRate Then
        picView.Width = picView.Width - lngDistanceX
      End If
  End Select
End Sub


'modify by tjh at 2010-01-20
'应用裁剪范围设置
Private Sub ApplayCutRange(videoObj As Object)

   mCurCutRange.LeftRate = Abs(videoObj.Left) / (mVideoSize.Width * mdblZoomRate)
   mCurCutRange.WidthRate = (mVideoSize.Width * mdblZoomRate - picView.Width + videoObj.Left) / (mVideoSize.Width * mdblZoomRate)

   mCurCutRange.TopRate = Abs(videoObj.Top) / (mVideoSize.Height * mdblZoomRate)
   mCurCutRange.HeightRate = (mVideoSize.Height * mdblZoomRate - picView.Height + videoObj.Top) / (mVideoSize.Height * mdblZoomRate)
End Sub


Private Sub imageScanner_PageDone(ByVal PageNumber As Long)
  If mintScanImageIndex = -1 Then
    Exit Sub
  End If

  '计算扫描文件索引
  mintScanImageIndex = mintScanImageIndex + 1
  
  Dim curScanFile As String
  curScanFile = CStr(mintScanImageIndex)
  
  '取得有效的扫描文件名称
  While Len(curScanFile) < 4
    curScanFile = "0" + curScanFile
  Wend
  
  curScanFile = mstrTempDirOfScan & "\" & CONST_STR_DEFAULT_SCAN_FILE_TEMPLATE & curScanFile & ".bmp"
  
  '保存扫描的图像
  Call DoScanCaptureDown(curScanFile)
End Sub


Private Sub ShowScanImage(objImg As DicomImage)

    '将被选中图像装载到dcmView中
    dcmView.Images.Clear
    dcmView.Images.Add objImg
    
    '显示dcmView，隐藏picVideo
    dcmView.CurrentImage.BorderWidth = 0
    mblnRealTime = False
'    picVideo.Visible = False
'    dcmView.Visible = True
End Sub

 

Private Sub labCloseAfter_Click()
On Error GoTo errHandle
    Call CloseAfterCap
Exit Sub
errHandle:
    MsgboxCus err.Description, vbOKOnly, G_STR_HINT_TITLE
End Sub

Private Sub CloseAfterCap()
    mstrAfterCapTag = ""
    labAfterInfo.Caption = "标识:---"
    picAfter.Visible = False
    
    Call mobjCapHelper.AfterTag("CLOSE")
End Sub

Private Sub labCloseLock_Click()
On Error GoTo errHandle
    Dim Control As XtremeCommandBars.CommandBarControl
    
    Set Control = cbrMain.FindControl(, conMenu_Cap_StudySyncState, False, True)
    
    If Control Is Nothing Then Exit Sub
    
    Call UnLockCapture(Control)
Exit Sub
errHandle:
    MsgboxCus err.Description, vbOKOnly, G_STR_HINT_TITLE
End Sub

Private Sub mobjDxDevice_OnDxKeyPress(ByVal lngButtonNum As Long)
On Error GoTo errHandle
    BUGEX "mobjDxDevice_OnDxKeyPress 1"
    BUGEX "mobjDxDevice_OnDxKeyPress ButtonNum:" & lngButtonNum

    Select Case lngButtonNum
        Case 0  '前台采集
                BUGEX "mobjDxDevice_OnDxKeyPress 2"
                If mstrAfterCapTag <> "" Then Call CloseAfterCap
                
                Call ForeCapture(True)
                
        Case 1  '后台采集
                BUGEX "mobjDxDevice_OnDxKeyPress 3"
'                If gobjCapturePar.IsUseAfterCapture Then
                    Call AfterCapture
'                End If
                
        Case 2  '更新标识
                BUGEX "mobjDxDevice_OnDxKeyPress 4"
'                If gobjCapturePar.IsUseAfterCapture Then
                    Call ResetAfterCaptureTag
'                End If
                
        Case Else
            Call mobjDxDevice_OnDxKeyPress(0)
    End Select
    
    BUGEX "mobjDxDevice_OnDxKeyPress 5"
    Exit Sub
errHandle:
    MsgboxCus err.Description, vbOKOnly, G_STR_HINT_TITLE
End Sub

Private Sub mfrmParameter_OnVideoDirverChange(ByVal vdtDirverType As TVideoDriverType)
'驱动改变后，更新采集界面
On Error GoTo errHandle
    Call mVideoCapture.StopPreview
    
    mVideoCapture.VideoDriverType = vdtDirverType
    
    Call UpdateCaptureDirver(vdtDirverType)
    
'    '如果为TWAIN的方式，则不进行视频的刷新
'    If mVideoCapture.VideoDriverType <> vdtTWAIN Then
'        Call mVideoCapture.StartPreview
'
'        Call mVideoCapture.RefreshVideoWindow
'    End If
Exit Sub
errHandle:
    MsgboxCus err.Description, vbOKOnly, G_STR_HINT_TITLE
End Sub


Private Sub mobjPlayWindow_OnCapture(pic As stdole.StdPicture)
On Error GoTo errHandle
    Call DoNormalCapture(True, pic)
Exit Sub
errHandle:
    
End Sub

'modify by tjh at 2010-01-20
Private Sub pbxSize_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    '开始执行裁剪范围设置
    If Button = 1 And gobjCapturePar.IsAllowChangeSize Then
        mblnMoveDown = True
    End If
  
End Sub


'modify by tjh at 2010-01-20
Private Sub pbxSize_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

  '设置视频裁剪范围
  If mblnMoveDown = True And Button = 1 Then
    If wdmCapture.Visible Then
      Call ChangeCutRanage(wdmCapture, Index, X, Y)
    ElseIf picCusVideo.Visible Then
      Call ChangeCutRanage(picCusVideo, Index, X, Y)
    Else
      Call ChangeCutRanage(dcmView, Index, X, Y)
    End If
      
            
    Call RefreshPbxSizePos

  End If
    
End Sub


'modify by tjh at 2010-01-20
Private Sub pbxSize_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
  If mblnMoveDown = True And Button = 1 Then
          
    '应用裁剪设置
    If wdmCapture.Visible Then
      Call ApplayCutRange(wdmCapture)
    ElseIf picCusVideo.Visible Then
      Call ApplayCutRange(picCusVideo)
    End If
    
    If IsTwainCaptureWay Or IsCustomCaptureWay Then
      ConfigTwainDisplay
    Else
      '设置显示范围
      Call ConfigVideoDisplay(wdmCapture)
      Call ConfigVideoDisplay(picCusVideo)

      '刷新视频显示
      If Not (mVideoCapture Is Nothing) Then
        Call mVideoCapture.RefreshVideoWindow
      End If
    End If

    '设置裁剪边框位置
    Call RefreshPbxSizePos

  End If
    
  mblnMoveDown = False
    
End Sub

Private Sub PlayCurVideo()
'------------------------------------------------
'功能：dcmView中录像图像的播放
'参数：无
'返回：无，直接播放dcmView中的图像
'------------------------------------------------
    Dim strFile As String
    
    If mobjPlayWindow Is Nothing Then
        Set mobjPlayWindow = New frmPlaying
    End If
    
    If dcmView.Images.count > 0 Then
        '下载录像，如果本地存在，则不进行下载
        If dcmView.Images(1).Tag.Tag <> VIDEOTAG And dcmView.Images(1).Tag.Tag <> AUDIOTAG Then Exit Sub
        
        strFile = dcmView.Images(1).Tag.VideoFile
        
        '打开播放··
        Call mobjPlayWindow.Show
        
        '刷新播放窗口
        While Not mobjPlayWindow.IsActive
            Call Sleep(10)
            DoEvents
        Wend
            
        Call mobjPlayWindow.OpenVideoFile(Replace(strFile, "/", "\"), Nothing, True)
    End If
End Sub


Private Sub tmrComm_Timer()
    On Error GoTo errHandle
    If gobjCapturePar.ComPortType = "COM" Then
        mcpsComState.lngComTime = mcpsComState.lngComTime + 2
        
        '大于0.08秒，则自动放弃
        If mcpsComState.lngComTime > 40 Then
            mcpsComState.lngComTime = 0
            
            tmrComm.Enabled = False
        End If
        
    Else
         If Not mobjDxDevice Is Nothing Then Call mobjDxDevice.PollDxDevice
    End If
    
    Exit Sub
errHandle:
    tmrComm.Enabled = False
    MsgboxCus err.Description, vbOKOnly, G_STR_HINT_TITLE
End Sub

Private Sub tmrReg_Timer()
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    
On Error GoTo errHandle:
    If Not mVideoCapture.IsStartup Then
        Exit Sub
    End If
 
    If gint视频设备数量 <= -1 Then Exit Sub
    
    strSQL = "select count(1) 已启用数量 from zltools.zlclients where 启用视频源=1"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "已启用数量")
    
    If rsTemp.RecordCount > gint视频设备数量 Then
        mstrVideoRegTime = ""
        mstrMsg = "授权视频站点数量与当前实际使用数量不匹配，请检查！"
        Exit Sub
    End If
    
    If DateDiff("S", mstrVideoRegTime, Now) >= M_LNG_REFRESHINTERVAL Then
        '判断数据库中是否存在已经注册的ip并且已经启用视频源，如果不存在则认为没有成功注册
        If FunCheckRegInfo(Me) Then
            mstrVideoRegTime = Now
        Else
            mstrVideoRegTime = ""
            mstrMsg = "视频源不允许启动，请联系管理员到服务管理工具中进行配置！"
            Exit Sub
        End If
    End If
Exit Sub
errHandle:
End Sub

Private Sub txtInputText_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Or KeyAscii = 27 Then  '''ESC和回车键退出输入
        txtInputText.Visible = False
        If Trim(txtInputText.Text) = "" Or KeyAscii = 27 Then
            '删除文字标注
            dcmView.Images(1).Labels.Remove dcmView.Images(1).Labels.count
            txtInputText = "1 "
        Else
            mdcmSelectLabel.Text = txtInputText.Text
            dcmView.Refresh
        End If
    End If
End Sub

Private Sub StartVideo()
'------------------------------------------------
'功能：录像
'参数：无
'返回：将录像文件放入缩略图
'------------------------------------------------
    Dim strError As String
    Dim strVideoFile As String
    Dim strEncoderName As String
    Dim lngRecordTimeLen As String
    Dim blnIsSave As Boolean
    
    On Error GoTo continue1
    
      '删除历史的视频文件
    mstrAviFileName = mstrBufferDir & "TmpVideo_" & Format(Now, "HHMMSS") & ".avi"
    If Dir(mstrAviFileName) <> "" Then RemoveFile mstrAviFileName
continue1:
    
    On Error GoTo CapErr
            
    '按现目前的方式,使用vfw的时候不允许进行录像操作
    If mVideoCapture.VideoDriverType = vdtVFW Or mVideoCapture.VideoDriverType = vdtTWAIN Then
        '录像完成(vfw进入录象后，直到结束才执行StartVideo以后的语句)
        '不处理vfw的录像功能
        MsgboxCus "不能再VFW和TWAIN驱动方式下录制视频。", vbOKOnly, "提示"
        Exit Sub
    End If
    
    If IsCustomCaptureWay Then
        If mobjCustomDevice Is Nothing Then
            MsgboxCus "专用视频采集接口调用失败，不能录制视频。", vbOKOnly, "提示"
            Exit Sub
        End If
         
        strVideoFile = mobjCustomDevice.zlStartVideo( _
                        mobjCapHelper.GetCustomMainID, _
                        mstrAviFileName, blnIsSave, _
                        strEncoderName, lngRecordTimeLen)
        
        If FileExists(strVideoFile) = False Then
            Exit Sub
        End If
        
        '如果这里返回了文件名，则直接保存
        If Len(strVideoFile) > 0 Then
            Call mobjCapHelper.SaveVideo(strVideoFile, "", _
                                strEncoderName, lngRecordTimeLen, blnIsSave)
            
            mstrAviFileName = ""
        End If
    Else
        'modify by tjh at 2010-01-20
        strError = mVideoCapture.StartVideo(mstrAviFileName)
        If Trim(strError) <> "" Then MsgboxCus strError, vbInformation, "提示"
    End If
    
    Exit Sub
CapErr:
    Call MsgboxCus(err.Description, vbOKOnly, G_STR_HINT_TITLE)
    err.Clear
End Sub


'modify by tjh at 2010-01-20
'停止视频录像
Private Sub StopVideo()
    Dim strVideoFile As String
    Dim strEncoderName As String
    Dim lngRecordTimeLen As Long
    
    Dim blnIsCusSave As Boolean
            
    If mVideoCapture.VideoDriverType = vdtVFW Or mVideoCapture.VideoDriverType = vdtTWAIN Then Exit Sub
    
On Error GoTo continue1
    If Dir(mstrAviFileName) <> "" Then RemoveFile mstrAviFileName
continue1:
    
    On Error GoTo CapErr
    
    If IsCustomCaptureWay Then
        If mobjCustomDevice Is Nothing Then
            MsgboxCus "专用视频采集接口调用失败。", vbOKOnly, "提示"
            Exit Sub
        End If
         
        strVideoFile = mobjCustomDevice.zlstopVideo( _
                        mobjCapHelper.GetCustomMainID, _
                        mstrAviFileName, blnIsCusSave, _
                        strEncoderName, lngRecordTimeLen)
                        
        If FileExists(strVideoFile) = False Then
            MsgboxCus "专用视频录像文件获取失败。", vbOKOnly, "提示"
            Exit Sub
        End If
    Else
        Call mVideoCapture.StopVideo
        
        strVideoFile = mstrAviFileName
        strEncoderName = mVideoCapture.GetEncoderName
        lngRecordTimeLen = mVideoCapture.GetTimeLen
        
        blnIsCusSave = False
    End If
       
    Call mobjCapHelper.SaveVideo(strVideoFile, "", _
                    strEncoderName, lngRecordTimeLen, Not blnIsCusSave)
    Exit Sub
CapErr:
    MsgboxCus err.Description, vbOKOnly, G_STR_HINT_TITLE
End Sub


'停止音频文件
Public Sub subSaveAudio(ByVal strAudioFile As String, ByVal lngTimeLen As Long)
On Error GoTo CapErr
   
    Call mobjCapHelper.SaveAudio(strAudioFile, "", "", lngTimeLen, True)
    
    Exit Sub
CapErr:
    MsgboxCus err.Description, vbOKOnly, G_STR_HINT_TITLE
End Sub

Private Sub InitCommandBars()
'功能创建工具条
    Dim cbrControl As CommandBarControl
    Dim cbrToolBar As CommandBar
    
    '-----------------------------------------------------
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    Me.cbrMain.VisualTheme = xtpThemeOffice2003
    
    BUGEX "InitCommandBars:Set CommandBar Icon"
    
    Set Me.cbrMain.Icons = zlCommFun.GetPubIcons 'imgPublic.Icons '
    
    BUGEX "InitCommandBars:1"
    
    With Me.cbrMain.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '放在VisualTheme后有效
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 32, 32
    End With
    
    BUGEX "InitCommandBars:2"
    
    Me.cbrMain.EnableCustomization False
    Me.cbrMain.ActiveMenuBar.Visible = False
    
    BUGEX "InitCommandBars:3"
    
    
    Set cbrToolBar = Me.cbrMain.Add("采集工具栏", xtpBarLeft)
 
'    cbrToolBar.EnableDocking xtpFlagStretched
    cbrToolBar.ShowTextBelowIcons = False
    cbrToolBar.Closeable = False

    With cbrToolBar.Controls
    
        '在非TWAIN采集模式的情况下，才显示该按钮
        'If Not GetIsTwainCaptureWay Then
            Set cbrControl = .Add(xtpControlButton, conMenu_Cap_Dynamic, "动态"): cbrControl.ToolTipText = "显示实时视频"
        'End If
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Cap_MarkMap, "采集"): cbrControl.ToolTipText = "采集图像"
        
'        '启用后台采集
'        Set cbrControl = .Add(xtpControlButton, conMenu_Cap_After_Capture, "后台采集"): cbrControl.ToolTipText = "后台采集"
'            cbrControl.IconId = 10020
        
        '在非TWAIN采集模式的情况下，才显示该按钮
        'If Not GetIsTwainCaptureWay Then
            Set cbrControl = .Add(xtpControlButton, conMenu_Cap_Record, "录像"): cbrControl.ToolTipText = "开始录像"
                cbrControl.Enabled = True
                
            
'            Set cbrControl = .Add(xtpControlButton, conMenu_Cap_After_Record, "后台录像"): cbrControl.ToolTipText = "后台录像"
'                cbrControl.IconId = 10021
            
            
'            Set cbrControl = .Add(xtpControlButton, conMenu_Cap_Record_Stop, "停止录像"): cbrControl.ToolTipText = "停止录像"
'                cbrControl.Enabled = False
                
            Set cbrControl = .Add(xtpControlButton, conMenu_Cap_RecordAudio, "录音"): cbrControl.ToolTipText = "录音"
        'End If
        
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Cap_Play, "播放"): cbrControl.ToolTipText = "播放录像"
            cbrControl.BeginGroup = True
            
        Set cbrControl = .Add(xtpControlButton, conMenu_Cap_Timer, "计时"): cbrControl.ToolTipText = "开启计时"
            cbrControl.IconId = 10024
            
'        Set cbrControl = .Add(xtpControlButton, conMenu_Cap_OpenStudyList, "打开检查"): cbrControl.ToolTipText = "打开检查"
'            cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Cap_StudySyncState, "锁定检查"): cbrControl.ToolTipText = "锁定检查"
            cbrControl.IconId = 10012
            cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Cap_After_Tag, "更新标识"): cbrControl.ToolTipText = "更新标识"
            cbrControl.IconId = 10022
            
        
            
            
        Set cbrControl = .Add(xtpControlButton, conMenu_ImgPro_Window, "亮度"): cbrControl.ToolTipText = "调节亮度/对比度": cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_ImgPro_Zoom, "缩放"): cbrControl.ToolTipText = "缩放图像"
        Set cbrControl = .Add(xtpControlButton, conMenu_ImgPro_Corp, "拖动"): cbrControl.ToolTipText = "拖动图像"
        
        Set cbrControl = .Add(xtpControlSplitButtonPopup, conMenu_ImgPro_Save, "保存"): cbrControl.ToolTipText = "保存图像": cbrControl.IconId = 3201
        With cbrControl.CommandBar.Controls
            Set cbrControl = .Add(xtpControlButton, conMenu_ImgPro_RectSave, "裁剪保存"): cbrControl.ToolTipText = "裁剪采集图像并保持": cbrControl.IconId = 0
            Set cbrControl = .Add(xtpControlButton, conMenu_ImgPro_DirectSave, "直接保存"): cbrControl.ToolTipText = "保存当前处理图像": cbrControl.IconId = 0
        End With
        
        Set cbrControl = .Add(xtpControlSplitButtonPopup, conMenu_ImgPro_Rotate_Pop, "旋转"): cbrControl.IconId = 503
        With cbrControl.CommandBar.Controls
          Set cbrControl = .Add(xtpControlButton, conMenu_ImgPro_RRotate, "顺时"): cbrControl.ToolTipText = "顺时针旋转": cbrControl.IconId = 503
          Set cbrControl = .Add(xtpControlButton, conMenu_ImgPro_LRotate, "逆时"): cbrControl.ToolTipText = "逆时针旋转": cbrControl.IconId = 504
        End With
        
        Set cbrControl = .Add(xtpControlSplitButtonPopup, conMenu_ImgPro_Smooth_Pop, "平滑"): cbrControl.IconId = 506
        With cbrControl.CommandBar.Controls
          Set cbrControl = .Add(xtpControlButton, conMenu_ImgPro_Sharpness, "锐化"): cbrControl.ToolTipText = "锐化": cbrControl.IconId = 505
          Set cbrControl = .Add(xtpControlButton, conMenu_ImgPro_Smooth, "平滑"): cbrControl.ToolTipText = "平滑": cbrControl.IconId = 506
        End With
        
        Set cbrControl = .Add(xtpControlSplitButtonPopup, conMenu_ImgPro_Lab_Pop, "标注"): cbrControl.ToolTipText = "标注": cbrControl.IconId = 509
        With cbrControl.CommandBar.Controls
          Set cbrControl = .Add(xtpControlButton, conMenu_ImgPro_Arrow, "箭头"): cbrControl.ToolTipText = "箭头标注"": cbrControl.IconId = 507"
          Set cbrControl = .Add(xtpControlButton, conMenu_ImgPro_Ellipse, "圆形"): cbrControl.ToolTipText = "圆形标注"": cbrControl.IconId = 508"
          Set cbrControl = .Add(xtpControlButton, conMenu_ImgPro_Text, "文本"): cbrControl.ToolTipText = "文字标注": cbrControl.IconId = 509
        End With
        
    End With
    
    For Each cbrControl In cbrToolBar.Controls
        cbrControl.Style = xtpButtonIcon
        cbrControl.Category = "采集处理"
        cbrControl.Enabled = False
    Next
End Sub


Private Sub ShowFrameSelectImagePopup()
'------------------------------------------------
'功能：创建框选图象的时候 ，鼠标右键的弹出菜单
'参数：
'返回：无
'------------------------------------------------

Dim cbrControl As CommandBarControl
Dim cbrToolBar As CommandBar
Dim cbrToolPopup As CommandBarPopup
    
    
    '鼠标右键弹出菜单
    Set cbrToolBar = Me.cbrMain.Add("鼠标右键", xtpBarPopup)
    cbrToolBar.EnableDocking xtpFlagStretched
    cbrToolBar.ShowTextBelowIcons = False
    cbrToolBar.Closeable = False
    With cbrToolBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_ImgPro_RectCapture, "保存")
        cbrControl.IconId = 0
    End With
    cbrToolBar.Visible = True
    cbrToolBar.ShowPopup
End Sub


'DicomViewer裁剪后采集图象
Private Sub CaptureFrameSelectImage()
    Dim imgResult As DicomImage
    
    '采集裁剪图像
    Set imgResult = CutImage(dcmView.Images(1))
    If imgResult Is Nothing Then Exit Sub
    
    '给imgResult一个唯一的 InstanceUID
    imgResult.InstanceUID = dcmglbUID.NewUID
     
    mintCaptureFlag = 1
    
    Call mobjCapHelper.SaveImg(imgResult, "")
End Sub

Private Sub StartTimer()
On Error GoTo errH
    mblnTimerState = Not mblnTimerState
    Call mVideoCapture.StartTimer(mblnTimerState)
    
    Exit Sub
errH:
    MsgboxCus err.Description, vbOKOnly, G_STR_HINT_TITLE
End Sub
 
