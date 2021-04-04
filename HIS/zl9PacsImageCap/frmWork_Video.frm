VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "Mscomm32.ocx"
Object = "{853AAF94-E49C-11D0-A303-0040C711066C}#4.3#0"; "DicomObjects.ocx"
Object = "{B1790453-7708-48C1-B5CC-75255FA4B066}#1.0#0"; "ZLDSVideoProcess.ocx"
Object = "{84926CA3-2941-101C-816F-0E6013114B7F}#1.0#0"; "IMGSCAN.OCX"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "*\A..\ZL9PACSCONTROL\zl9PacsControl.vbp"
Begin VB.Form frmWork_Video 
   BorderStyle     =   0  'None
   ClientHeight    =   8880
   ClientLeft      =   0
   ClientTop       =   60
   ClientWidth     =   10410
   ClipControls    =   0   'False
   DrawStyle       =   5  'Transparent
   Icon            =   "frmWork_Video.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   8880
   ScaleWidth      =   10410
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Tag             =   "视频采集"
   Begin ScanLibCtl.ImgScan ImageScanner 
      Left            =   4080
      Top             =   6360
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
   Begin zl9PacsControl.ucSplitter ucSplitter1 
      Height          =   135
      Left            =   0
      TabIndex        =   1
      Top             =   4620
      Width           =   10410
      _ExtentX        =   18362
      _ExtentY        =   238
      MousePointer    =   7
      SplitType       =   0
      DBClickType     =   2
      SplitLevel      =   3
      Con1MinSize     =   3000
      Con2MinSize     =   1000
      Control1Name    =   "picCapture"
      Control2Name    =   "picReportImage"
   End
   Begin VB.PictureBox picReportImage 
      Height          =   4125
      Left            =   0
      ScaleHeight     =   4065
      ScaleWidth      =   10350
      TabIndex        =   13
      Top             =   4755
      Width           =   10410
      Begin VB.PictureBox picCacheImage 
         BackColor       =   &H8000000D&
         Height          =   3375
         Left            =   4800
         ScaleHeight     =   3315
         ScaleWidth      =   3795
         TabIndex        =   16
         Top             =   720
         Width           =   3855
         Begin MSComCtl2.DTPicker DTPImg 
            Height          =   255
            Left            =   360
            TabIndex        =   19
            Top             =   120
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   450
            _Version        =   393216
            Format          =   110755841
            CurrentDate     =   42587
         End
         Begin VB.ComboBox cboCacheImage 
            Height          =   300
            Left            =   2160
            Style           =   2  'Dropdown List
            TabIndex        =   18
            Top             =   120
            Width           =   1215
         End
         Begin zl9PacsControl.ucImagePreview ucCacheImage 
            Height          =   2565
            Left            =   960
            TabIndex        =   17
            Top             =   600
            Width           =   1890
            _ExtentX        =   3334
            _ExtentY        =   4524
            BackColor       =   4210752
         End
      End
      Begin VB.PictureBox picPreview 
         BackColor       =   &H8000000D&
         Height          =   2535
         Left            =   480
         ScaleHeight     =   2475
         ScaleWidth      =   2955
         TabIndex        =   14
         Top             =   600
         Width           =   3015
         Begin zl9PacsControl.ucImagePreview ucPreview 
            Height          =   2205
            Left            =   480
            TabIndex        =   15
            Top             =   120
            Width           =   1890
            _ExtentX        =   3334
            _ExtentY        =   4524
            BackColor       =   4210752
         End
      End
      Begin XtremeDockingPane.DockingPane dkpReportImage 
         Left            =   3600
         Top             =   360
         _Version        =   589884
         _ExtentX        =   450
         _ExtentY        =   423
         _StockProps     =   0
      End
   End
   Begin VB.Timer tmrReg 
      Interval        =   10000
      Left            =   1560
      Top             =   6120
   End
   Begin MSComDlg.CommonDialog dlgOpen 
      Left            =   720
      Top             =   4950
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer tmrComm 
      Interval        =   2
      Left            =   0
      Top             =   5040
   End
   Begin VB.Timer timerHook 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   15
      Top             =   6090
   End
   Begin VB.Timer timerRePaint 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   135
      Top             =   6780
   End
   Begin VB.PictureBox picCapture 
      ForeColor       =   &H00000000&
      Height          =   4620
      Left            =   0
      ScaleHeight     =   4560
      ScaleWidth      =   10350
      TabIndex        =   2
      Top             =   0
      Width           =   10410
      Begin VB.PictureBox pbxSize 
         BackColor       =   &H00800000&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   75
         Index           =   1
         Left            =   360
         MousePointer    =   7  'Size N S
         ScaleHeight     =   75
         ScaleWidth      =   7335
         TabIndex        =   11
         Top             =   3840
         Width           =   7335
      End
      Begin VB.PictureBox pbxSize 
         BackColor       =   &H00800000&
         BorderStyle     =   0  'None
         Height          =   3975
         Index           =   2
         Left            =   480
         MousePointer    =   9  'Size W E
         ScaleHeight     =   3975
         ScaleWidth      =   75
         TabIndex        =   10
         Top             =   0
         Width           =   75
      End
      Begin VB.PictureBox pbxSize 
         BackColor       =   &H00800000&
         BorderStyle     =   0  'None
         Height          =   3975
         Index           =   3
         Left            =   7560
         MousePointer    =   9  'Size W E
         ScaleHeight     =   3975
         ScaleWidth      =   75
         TabIndex        =   9
         Top             =   15
         Width           =   75
      End
      Begin VB.PictureBox pbxSize 
         BackColor       =   &H00800000&
         BorderStyle     =   0  'None
         Height          =   75
         Index           =   0
         Left            =   360
         MousePointer    =   7  'Size N S
         ScaleHeight     =   75
         ScaleWidth      =   7335
         TabIndex        =   8
         Top             =   120
         Width           =   7335
      End
      Begin VB.PictureBox picView 
         BackColor       =   &H8000000D&
         BorderStyle     =   0  'None
         Height          =   3495
         Left            =   600
         ScaleHeight     =   3495
         ScaleWidth      =   6855
         TabIndex        =   3
         Top             =   240
         Width           =   6855
         Begin ZLDSVideoProcess.DSCapture wdmCapture 
            Height          =   3135
            Left            =   720
            TabIndex        =   4
            Top             =   240
            Width           =   3495
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
            CurWidth        =   233
            CurHeight       =   209
            CurVideoWidth   =   231
            CurVideoHeight  =   189
            ShowModel       =   0
            CapParameterWindPos=   8
            SnatchWay       =   0
            ParameterCfgFileName=   ""
            HideCfgItem     =   0
            AppHandle       =   0
         End
         Begin VB.PictureBox picVideo 
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            Height          =   3015
            Left            =   1200
            ScaleHeight     =   201
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   224
            TabIndex        =   6
            Top             =   120
            Width           =   3360
         End
         Begin VB.TextBox txtInputText 
            Appearance      =   0  'Flat
            Height          =   300
            Left            =   5520
            TabIndex        =   5
            Text            =   "Text1"
            Top             =   840
            Visible         =   0   'False
            Width           =   975
         End
         Begin DicomObjects.DicomViewer dcmView 
            Height          =   1575
            Left            =   4440
            TabIndex        =   7
            Top             =   1440
            Width           =   2175
            _Version        =   262147
            _ExtentX        =   3836
            _ExtentY        =   2778
            _StockProps     =   35
            BackColor       =   0
            UseScrollBars   =   0   'False
         End
      End
      Begin DicomObjects.DicomViewer dcmAfter 
         Height          =   735
         Left            =   8820
         TabIndex        =   12
         Top             =   3195
         Visible         =   0   'False
         Width           =   1035
         _Version        =   262147
         _ExtentX        =   1826
         _ExtentY        =   1296
         _StockProps     =   35
         BackColor       =   0
         UseScrollBars   =   0   'False
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
   Begin MSCommLib.MSComm commListener 
      Bindings        =   "frmWork_Video.frx":06EA
      Left            =   0
      Top             =   3840
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.PictureBox picTemp2 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   1455
      Left            =   8280
      ScaleHeight     =   1455
      ScaleWidth      =   375
      TabIndex        =   0
      Top             =   1200
      Visible         =   0   'False
      Width           =   375
   End
End
Attribute VB_Name = "frmWork_Video"
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

'被锁定检查的基本信息
Private Type TUnLockStudyInf
    lngAdviceId As Long
    lngSendNo As Long
    blnMoved As Boolean
    lngStudyState As Long
End Type

'当前待采集图像检查的基本信息
Private Type TCurStudyBaseInf
    strStudyUid As String          '检查UID
    strModality As String          '影像类别
    strSex As String               '性别
    strBirthDate As String         '出生日期
    strAge As String               '年龄
    strName As String              '姓名
    strCheckNo As String           '检查号
    strPatientID As String         '病人ID
End Type


'后台采集信息
Private Type TAfterCaptureInf
    intAfterTag As Integer    '标识
    strAfterStudyUid As String     '后台采集检查UID
    strAfterSeriesUid As String    '后台采集序列UID

    lngAfterCurImageCount As Long  '当前后台采集图像数量
    strAfterParentTitle As String  '后台采集信息
End Type

'COM脚踏端口状态
Private Type TComPortState
    intComState As Integer          'COM口的状态
    lngComTime As Long              '记录com口保持状态的时间
    dtLastCapture As Date           '最近脚踏踩下的时间
    blnCTSHolding As Boolean        '记录常态时的CTS线的电平
End Type


Private mdcmTmpImg As DicomImage
Private mintCaptureFlag As Integer

Private mobjCustomDevice As Object  '专用视频采集对象

Private dcmglbUID As New DicomGlobal    '定义UIDRoot=1

Private WithEvents mobjDxDevice As clsDxHidDevice   '实现蓝韵手柄之类的采集方式
Attribute mobjDxDevice.VB_VarHelpID = -1
'Private WithEvents mobjHotHook As clsHookKey

Public mhCapWnd As Long                 '采集窗口的句柄
Private WithEvents mfrmParameter As frmVideoSetup
Attribute mfrmParameter.VB_VarHelpID = -1
Private mfrmOpenStudy As frmOpenStudyList
Private mstrAfterStationName As String

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

Private mstrNameInf As String

Private mblnMoveDown  As Boolean        '用于判断是否按下鼠标左键
Private mblnDcmViewDown As Boolean      '用于判断dcmView中鼠标是否被按下
Private mintCurImgIndex As Integer      '当前被选中的图象索引
Private mdcmSelectLabel As DicomLabel   '当前被选中的标注

Private mstrAviFileName As String       '录像文件名
Private mstrEncoderName As String       '
Private mstrBufferDir As String
Private mstrAfterDir As String

Private mcpsComState As TComPortState       'Com端口使用状态

Private mblnUseClipbord As Boolean          '是否使用剪贴板


Private mobjFtpConnection As New clsFtp
Private mobjBakFtpConnection As New clsFtp

Private mobjFtp As TFtpDeviceInf        'ftp设备信息
Private mobjBakFtp As TFtpDeviceInf     'ftp备份存储设备信息


Private mblnReadOnly As Boolean         '是否只能查看True查看模式，False采集模式

Private mblnShowProcessBar As Boolean   '是否显示处理工具栏


'病人基本信息资料
Private mlngTmpAdviceId As Long
Private mlngTmpSendNo As Long

Private mlngAdviceId As Long            '医嘱ID
Private mlngSendNo As Long
Private mblnMoved As Boolean            '是否转储
Private mlngStudyState As Long



Private mAfterCaptureInf As TAfterCaptureInf    '后台采集信息
Private mSelectStudyInf As TUnLockStudyInf      '锁定的检查信息
Private mcurStudyInf As TCurStudyBaseInf        '当前检查信息

Private mVideoCapture As clsVideoCapture '视频采集对象

Private mdblZoomRate As Double  '缩放比率（在cbrMain的cbrMain_ResizeClient事件中需要重新计算该值）
Private mVideoSize As TVideoSize '视频大小（由相关的视频组件保存）
Private mCurCutRange As TCutRange '视频裁剪范围设置（该参数通过GetString和SaveString保存在注册表中）
Private mVideoArea As TVideoArea  '视频客户区域设置（在cbrMain的cbrMain_ResizeClient事件中需要重新计算该值）

Private mblnCaptureLockState As Boolean '视频锁定状态

Private mstrInstitution As String       '单位名称

Private Const M_LNG_REFRESHINTERVAL As Long = 600 '刷新间隔
Private mstrVideoRegTime As String '保存视频启动注册时间
Private mstrMsg As String
Private mblnRefreshState As Boolean
Private mblnInitState As Boolean

Private mrsImageCache As New ADODB.Recordset
Private mobjFile As New FileSystemObject
Private mlngReleationType As Integer    '1--导出；2--导入
Private mdcmUID As New DicomGlobal

Private mdate As Date 'DTP控件选择的时间,用于过滤后台图
Private mblIsNeedRefresh '判断是否需要重新加载后台图(用于保持与报告模块下拉框图像一致)
Private mintTagMax As Integer '最大标识
Private mintTagNow As Integer '当前标识
Private mintFontSize As Integer '字号
Private mblnIsReported As Boolean   '是否已经存在报告了
Private mstrInstanceUID As String

Private mblnCompareSize As Boolean
Private mblnImageShield As Boolean   '是否屏蔽大图

Private Const CAPTURE_PARAMETER_CONFIG_FILE_NAME As String = "ZLVideoProcess.ini"
Private Const CONST_STR_DEFAULT_TEMP_SCAN_DIR_NAME As String = "\TempScan"  '默认扫描文件临时存储路径
Private Const CONST_STR_DEFAULT_SCAN_FILE_TEMPLATE As String = "Scan"  '默认扫描文件临时存储路径

Private Type DlgFileInfo
    iCount As Long
    sPath As String
    sFIle() As String
End Type

Private Enum Dkp_ID
    Dkp_ID_Video = 1     '检查列表
    Dkp_ID_Miniature      '当前病人基本信息
End Enum


'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'----------------------------------------------------------------------------------------------------------

Private Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Private Declare Function MoveFile Lib "kernel32" Alias "MoveFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String) As Long

Public Event OnIsUseAfterImgChange(ByVal blUse As Boolean)
Public Event OnStateChange(ByVal lngEventType As TVideoEventType, ByVal lngAdviceId As Long, ByVal lngSendNo As Long, ByVal strOther As String)
Public Event OnControlResize(objControl As Object)
Public Event OnImgLoadState(ByVal blnLoadFinish As Boolean, ByVal blnUpLoad As Boolean)

Property Get CaptionEx() As String
    CaptionEx = Me.Tag
End Property

Property Let CaptionEx(value As String)
    Dim hWndParent As Long
    
    Me.Tag = value
    
    hWndParent = GetParent(Me.hwnd)
    
    Call SetWindowText(hWndParent, Me.Tag)
End Property


'获取视频采集对象
Property Get videoCapture() As clsVideoCapture
    Set videoCapture = mVideoCapture
End Property


'获取视频采集窗口的初始化状态
Property Get InitState() As Boolean
    InitState = mblnInitState
End Property

'锁定的病人姓名
Property Get LockPatientName() As String
    LockPatientName = mstrNameInf
End Property

'当前锁定状态
Property Get LockState() As Boolean
    LockState = mblnCaptureLockState
End Property

Private Sub UnLockStudy()
'解锁检查
    mblnCaptureLockState = False
End Sub

'是否已经存在报告了
Property Get IsReported() As Boolean
    IsReported = mblnIsReported
End Property
Property Let IsReported(ByVal value As Boolean)
    mblnIsReported = value
End Property

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

Private Function LoadMiniCache() As Boolean
    Dim i As Integer
    Dim strQueryPath As String
    
    Dim objFolder2 As Folder, objFolder3 As Folder, objFolder4 As Folder
    Dim strStudyUid As String, strSeriesUID As String
    Dim lngStudyNo As Long, lngSeriesNo As Long

    Dim strAfterTime As String
    Dim dtChose As Date
    Dim intTmp As Integer
    Dim strTag As String  '三位数的标识
    Dim strType As String
    Dim curDate As Date
        
    If gobjCapturePar.IsUseAfterCapture = False Then Exit Function
    
    curDate = zlDatabase.Currentdate
    dtChose = mdate

    Set mrsImageCache = New ADODB.Recordset
    mrsImageCache.Fields.Append "姓名", adVarChar, 100
    mrsImageCache.Fields.Append "检查号", adVarChar, 64
    mrsImageCache.Fields.Append "检查UID", adVarChar, 64
    mrsImageCache.Fields.Append "序列号", adVarChar, 18
    mrsImageCache.Fields.Append "序列UID", adVarChar, 64
    mrsImageCache.Fields.Append "检查日期", adVarChar, 20

    mrsImageCache.Fields.Append "路径", adVarChar, 4000
    mrsImageCache.Open
    
    If mobjFile.FolderExists(mstrAfterDir) = False Then Exit Function
    
    If mobjFile.GetFolder(mstrAfterDir).SubFolders.Count > 0 Then
        For Each objFolder2 In mobjFile.GetFolder(mstrAfterDir).SubFolders

                        If InStr(objFolder2.Name, Format(mdate, "yyyymmdd")) > 0 Then ''如果不是选择的时间则跳过

                                If objFolder2.SubFolders.Count > 0 Then
                                
                                        For Each objFolder3 In objFolder2.SubFolders                            '检查UID层
                                        
                                                If objFolder3.SubFolders.Count >= 0 Then

                                                        strAfterTime = Format(objFolder3.DateCreated, "YYYY-MM-DD HH:MM:SS")

                                                        strStudyUid = GetStudyUIDFromFolderName(objFolder3.Name)
                                                                                                                  
                                                        lngStudyNo = lngStudyNo + 1
                                                        lngSeriesNo = 0
                                                        
                                                        For Each objFolder4 In objFolder3.SubFolders                    '序列UID层
                                                                
                                                                strSeriesUID = objFolder4.Name
                                                                lngSeriesNo = lngSeriesNo + 1
                                                                
                                                                mrsImageCache.AddNew
                                                                strTag = Lpad(GetTag(objFolder3.Name, strType), 3, "0")
                                                                mrsImageCache!姓名 = strType & strTag
                                                                mrsImageCache!检查号 = lngStudyNo
                                                                mrsImageCache!检查uid = strStudyUid
                                                                mrsImageCache!序列号 = lngSeriesNo
                                                                mrsImageCache!序列Uid = strSeriesUID
                                                                mrsImageCache!检查日期 = strAfterTime
                                                                mrsImageCache!路径 = objFolder4.Path
                                                                mrsImageCache.Update
                                                                
                                                        Next
                                                        
                                                End If
                    Next
                                    Exit For '此时已经遍历完所选时间，跳出时间选择
                                End If
                        End If '时间选择
        Next
    End If
    
    If mrsImageCache.RecordCount > 0 Then
        mrsImageCache.Sort = "检查日期 desc"
        mrsImageCache.MoveFirst
    End If

    cboCacheImage.Clear
    ucCacheImage.ImgViewer.Images.Clear
    
    For i = 0 To mrsImageCache.RecordCount - 1
        If i = 0 Then strQueryPath = Nvl(mrsImageCache!路径)

        cboCacheImage.AddItem Nvl(mrsImageCache!姓名) & "     时间：" & Format(Nvl(mrsImageCache!检查日期), "HH:MM:SS")

        mrsImageCache.MoveNext
    Next
    
    If mrsImageCache.RecordCount > 0 Then
        If cboCacheImage.ListIndex < 0 Then
            cboCacheImage.ListIndex = 0
        End If
    End If
  
End Function

Private Sub cboCacheImage_Click()
    Dim strQueryPath As String
    Dim rsTmp As ADODB.Recordset
    On Error GoTo errH
    
    ucCacheImage.ClearCurrentPageImage
    
    Set rsTmp = mrsImageCache
    rsTmp.Filter = ""
        
    If mrsImageCache.RecordCount <= 0 Then Exit Sub
    
    mrsImageCache.MoveFirst
    Set rsTmp = mrsImageCache

    rsTmp.Filter = "姓名='" & Trim(Mid(cboCacheImage.Text, 1, 5)) & "'"

    If rsTmp.RecordCount < 1 Then Exit Sub
    strQueryPath = Nvl(rsTmp!路径)
    
    If strQueryPath = "" Then Exit Sub
    Call ucCacheImage.RefreshImage(slLocal, strQueryPath, mblnMoved)
    Call ucCacheImage.RedrawSelf
    Call dkpReportImage.RedrawPanes
    
    
    Exit Sub
errH:
    MsgboxCus err.Description, vbOKOnly, G_STR_HINT_TITLE
    err.Clear
End Sub

Private Sub DTPImg_Change()
    On Error GoTo errH
        
    mdate = DTPImg.value
    Call LoadMiniCache
    ucCacheImage.RedrawSelf
    Call dkpReportImage.RedrawPanes
        
    Exit Sub
errH:
    MsgboxCus err.Description, vbOKOnly, G_STR_HINT_TITLE
End Sub

Private Sub Form_Initialize()
'初始化模块变量
    mblnInitState = False
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
    
BUGEX "ShowVideoConfig 4"
    Call ConfigFtpStorageDevice(gobjCapturePar.CurStorageDeviceNo, gobjCapturePar.BakStorageDeviceNo)

BUGEX "ShowVideoConfig 5"
    If gobjCapturePar.IsUseAfterCapture Then
        Call UpdateAfterCaptureInfo
    Else
        Call ShowAfterCaptureInf(False)
    End If
    
BUGEX "ShowVideoConfig 6"
    Call OpenComm
    
    If gobjCapturePar.VideoDirverType = vdtCustom Then Call InitCustomDevice
    
    gstrHotKeyTest = GetSetting("ZLSOFT", "公共模块", "采集热键", "F8")
    
''    重新注册全局热键
'    If gobjCapturePar.strCaptureHot <> "" Then
'        Call mobjHotHook.EnableHook(WM_KEYDOWN, True)
'    Else
'        Call mobjHotHook.FreeHook
'    End If
    '----------------------------------------------------------
    
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
    
    mstrInstitution = GetSetting("ZLSOFT", "注册信息", "单位名称", "")

    mAfterCaptureInf.strAfterParentTitle = ""

    '如果程序在磁盘的根目录则app.path为“x:\”
    mstrBufferDir = GetAppPath & "\TmpImage\"
    mstrAfterDir = GetAppPath & "\TmpAfterImage\"
    
    mstrAviFileName = mstrBufferDir & "TmpVideo.avi"
    
    gint视频设备数量 = getLicenseCount(LOGIN_TYPE_视频设备)
    
    mblnUseClipbord = GetSetting("ZLSOFT", G_STR_REG_PATH_PUBLIC, "UseClipbord", 0)
    Call SaveSetting("ZLSOFT", G_STR_REG_PATH_PUBLIC, "UseClipbord", IIf(mblnUseClipbord, 1, 0))
    
    TimerRePaint.Interval = Val(GetSetting("ZLSOFT", G_STR_REG_PATH_PUBLIC, "界面重绘间隔", 500))
    Call SaveSetting("ZLSOFT", G_STR_REG_PATH_PUBLIC, "界面重绘间隔", TimerRePaint.Interval)
    
    mblnCompareSize = IIf(Val(GetSetting("ZLSOFT", "公共模块\Ftp", "启用FTP文件大小对比", 1)) <> 0, True, False)
    Call SaveSetting("ZLSOFT", "公共模块\Ftp", "启用FTP文件大小对比", IIf(mblnCompareSize, 1, 0))

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

    '鼠标移动时，显示大图
    ucPreview.BigImageWay = gobjCapturePar.ImagePreview
    ucPreview.PreviewTime = gobjCapturePar.PreviewTime
    ucPreview.ImgLoadType = gtFileLoadType
    ucPreview.DoShield = True
    
    Call GetLocalPar
    ucPreview.ImageShield = mblnImageShield
    
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

    '加载所有报告图像
    ucPreview.OnlyLoadReportImage = False
End Sub


Private Sub ConfigFtpStorageDevice(ByVal strCurStorageNo As String, ByVal strBakStorageNo As String)
'配置ftp存储设备
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    
    '配置在线存储设备信息
    strSQL = "Select 设备号,设备名 From 影像设备目录 Where 类型=1 and 设备号=[1] and NVL(状态,0)=1"
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "查询存储设备", strCurStorageNo)
    
    mobjFtp.strDeviceId = ""
    If rsTmp.EOF Then
        MsgboxCus "影像存储设备未定义或处于停用，请检查！", vbInformation, G_STR_HINT_TITLE
        
        mobjFtp.strDeviceId = ""
        mblnReadOnly = True
        Exit Sub
    End If
    
    mobjFtp.strDeviceId = strCurStorageNo
    Call funGetFtpDeviceInf(Me, mobjFtp)
    

    '配置备份设备信息
    mobjBakFtp.strDeviceId = ""
    If Val(strBakStorageNo) > 0 Then
        strSQL = "Select 设备号,设备名 From 影像设备目录 Where 类型=1 and 设备号=[1] and NVL(状态,0)=1"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "查询备份设备", strBakStorageNo)
        
        If rsTmp.EOF Then
            MsgboxCus "未取得有效的备份设备信息，不能对采集图像进行备份操作，请检查备份设备配置是否正确。", vbInformation, G_STR_HINT_TITLE
            
            Exit Sub
        End If
        
        mobjBakFtp.strDeviceId = strBakStorageNo
        Call funGetFtpDeviceInf(Me, mobjBakFtp)
    End If
    
End Sub

Public Sub zlRePreview()
'重新进入视频预览
    If mVideoCapture.IsStartup Then
        Call mVideoCapture.StopPreview
        Call mVideoCapture.StartPreview
    End If
End Sub

Public Sub zlInitModule()
BUGEX "zlPacsCapture zlInitModule 0"
'初始化模块参数
    
    '初始化参数
    Call InitParameter
    
BUGEX "gobjCapturePar.CurStorageDeviceNo = " & gobjCapturePar.CurStorageDeviceNo
    '配置ftp存储设备
    Call ConfigFtpStorageDevice(gobjCapturePar.CurStorageDeviceNo, gobjCapturePar.BakStorageDeviceNo)

BUGEX "zlInitModule 1"
    '打开视频采集设备
    Call OpenVideoCaptureDevice

BUGEX "zlInitModule 2"
    '更新后台采集信息
    If gobjCapturePar.IsUseAfterCapture Then Call UpdateAfterCaptureInfo
    
BUGEX "zlInitModule 3"
    '初始化专用视频采集接口
    Call InitCustomDevice
BUGEX "zlInitModule End"
    mblnInitState = True
End Sub

Private Sub InitCustomDevice()
    Dim strCustomDeviceDir As String        '专用视频采集部件路径
    Dim strCustomDeviceDllName As String    '专用视频采集部件名称
    Dim objFile As New FileSystemObject
    
    '初始化专用视频采集接口
    strCustomDeviceDir = gobjCapturePar.CustomDevicePath
    If strCustomDeviceDir <> "" Then
        strCustomDeviceDllName = Trim(Replace(objFile.GetFileName(strCustomDeviceDir), ".dll", ""))
        
        Set mobjCustomDevice = CreateObject(strCustomDeviceDllName & ".cls" & strCustomDeviceDllName)
        
        If Not mobjCustomDevice Is Nothing Then
            Call mobjCustomDevice.zlInit(gcnVideoOracle, UserInfo.id, glngDepartId, glngRootHandle)
        End If
    End If
End Sub


'----------------------------------------------------------------------------------------------------------
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++


Public Sub zlUpdateAdviceInf(ByVal lngAdviceId As Long, _
                            ByVal lngSendNo As Long, _
                            ByVal lngStudyState As Long, _
                            ByVal blnMoved As Boolean)
'更新医嘱信息
    Dim rsTemp As ADODB.Recordset
    Dim strSQL As String
    
    '保存主界面的当前检查信息
    mSelectStudyInf.lngAdviceId = lngAdviceId
    mSelectStudyInf.blnMoved = blnMoved
    mSelectStudyInf.lngSendNo = lngSendNo
    mSelectStudyInf.lngStudyState = lngStudyState
    
    If mblnCaptureLockState Then Exit Sub
    
    mlngAdviceId = lngAdviceId
    mlngSendNo = lngSendNo
    mblnMoved = blnMoved
    mlngStudyState = lngStudyState
    
    mblnReadOnly = False
    mblnRefreshState = True
    
    '数据被转移时，没有权限时，状态为指定状态时，该模块为只读
    If mlngAdviceId <= 0 Or blnMoved Or lngStudyState = 6 Or lngStudyState = 0 Or lngStudyState = 1 Or InStr(gstrPrivs, "视频采集") <= 0 Then
        mblnReadOnly = True
    End If
    
    '提取病人基本信息,写DICOM参数时用
    strSQL = "Select A.影像类别,A.姓名,A.性别,A.年龄,A.出生日期,A.姓名,A.检查号,A.检查UID,B.病人ID " & _
                " From 影像检查记录 A,病人医嘱记录　B " & _
                " Where A.医嘱ID=B.Id And A.医嘱ID=[1]"
                
    If mblnMoved Then
        strSQL = Replace(strSQL, "影像检查记录", "H影像检查记录")
        strSQL = Replace(strSQL, "病人医嘱记录", "H病人医嘱记录")
    End If
    
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "提取病人基本信息", lngAdviceId)
    
    If Not rsTemp.EOF Then
        mcurStudyInf.strStudyUid = Nvl(rsTemp("检查UID"))
        mcurStudyInf.strModality = Nvl(rsTemp("影像类别"))
        mcurStudyInf.strSex = Nvl(rsTemp("性别"))
        mcurStudyInf.strAge = Nvl(rsTemp("年龄"))
        mcurStudyInf.strBirthDate = Nvl(rsTemp("出生日期"))
        mcurStudyInf.strName = Nvl(rsTemp("姓名"))
        mcurStudyInf.strCheckNo = Nvl(rsTemp("检查号"))
        mcurStudyInf.strPatientID = Nvl(rsTemp("病人ID"))
        
        mstrNameInf = Nvl(rsTemp("姓名"))
        
        mcurStudyInf.strSex = IIf(mcurStudyInf.strSex = "男", "M", IIf(mcurStudyInf.strSex = "女", "F", "O"))
    Else
        mcurStudyInf.strStudyUid = ""
        mcurStudyInf.strModality = ""
        mcurStudyInf.strSex = ""
        mcurStudyInf.strAge = ""
        mcurStudyInf.strCheckNo = ""
        mcurStudyInf.strPatientID = ""
        mcurStudyInf.strBirthDate = ""
        mcurStudyInf.strName = ""
        
        mstrNameInf = ""
    End If
    
    Me.Tag = "图像采集" & IIf(mstrNameInf <> "", "(" & mstrNameInf & ")", "")
    Me.CaptionEx = Me.Tag
End Sub


Public Sub zlRefreshFace(Optional blnForceRefresh As Boolean = False)
'刷新界面
On Error GoTo errHandle
    Dim rsTemp As ADODB.Recordset
    Dim iRows As Integer
    Dim iCols As Integer
    Dim strStudyUid As String
    
BUGEX "zlRefreshFace 0"
    If (mlngTmpAdviceId = mlngAdviceId And mlngTmpSendNo = mlngSendNo And mblnRefreshState) And Not blnForceRefresh And Not mblIsNeedRefresh Then Exit Sub

BUGEX "zlRefreshFace 0.1"
    mlngTmpAdviceId = mlngAdviceId
    mlngTmpSendNo = mlngSendNo
    mblnRefreshState = True

BUGEX "zlRefreshFace 1"
    Call ConfigVideoShowState(True)

BUGEX "zlRefreshFace 2"
'增加刷新过程用于保证与报告模块的显示数据同步
    Call ucPreview.RefreshImage(slStudy, mcurStudyInf.strStudyUid, mblnMoved, blnForceRefresh, False)
    Call ClearEmptyFolder(False) '删除空目录
    Call LoadMiniCache
    mblIsNeedRefresh = False
    
    '切换到 检查图tab
    dkpReportImage.Panes(1).Selected = True
    
BUGEX "zlRefreshFace 3"
    If ucPreview.ImgViewer.Images.Count > 0 Then
BUGEX "zlRefreshFace 4"
        '将被选中图像装载到dcmView中
        Call PreviewThumbnail(ucPreview.SelectIndex)
BUGEX "zlRefreshFace 5"
        '如果是Twain或专用视频采集模式，则设置mblnRealTime为false
        If IsTwainCaptureWay = True Or IsCustomCaptureWay Then mblnRealTime = False
    Else
        Call dcmView.Images.Clear
    End If
BUGEX "zlRefreshFace 6"
    Exit Sub
errHandle:
    MsgboxCus err.Description, vbOKOnly, G_STR_HINT_TITLE
End Sub


Private Sub StopCapture()
'-----------------------------------------------------------------------------------------
'功能：停止显示视频采集,释放视频采集窗口，
'      释放串口侦听的端口
'参数：无
'返回：无
'-----------------------------------------------------------------------------------------
'    Call mobjHotHook.FreeHook
    
    '关闭COMM口
    If commListener.PortOpen Then commListener.PortOpen = False
    
    '释放采集设备及窗体
    If Not mVideoCapture Is Nothing Then
        Call mVideoCapture.StopPreview
    End If
    
    '采用Midi接口需在消毁事件句柄
    If Not mobjDxDevice Is Nothing Then
        If mobjDxDevice.Handle <> 0 Then Call mobjDxDevice.CloseDxDevice
    End If
    
'    Call ucCapHook.FreeHook
End Sub



Public Sub zlUpdateCommandBars(control As XtremeCommandBars.CommandBarControl)
'只有影像采集工作站才具备后台采集功能

'根据当前状态确定菜单的可视和可操作

    '如果没有初始化视频对象，则视频相关的按钮都不允许使用
    If mVideoCapture Is Nothing Then
        control.Enabled = False
        Exit Sub
    End If
    
    Select Case control.id
        Case conMenu_Cap_Dynamic       '动态显示
            control.Checked = mblnRealTime
            control.Enabled = (Not mblnReadOnly) And (Not IsTwainCaptureWay And Not IsCustomCaptureWay) And mVideoCapture.IsStartup ' And (mhCapWnd <> 0) modify by tjh at 2010-01-20
            control.Visible = Not IsTwainCaptureWay And Not IsCustomCaptureWay
            
            If mblnRealTime Then
                control.IconId = conMenu_Cap_Dynamic
            Else
                control.IconId = 10023
            End If
            
        Case conMenu_Cap_MarkMap       '影像采集
            control.Enabled = Not mblnReadOnly And (mVideoCapture.IsStartup Or IsTwainCaptureWay Or IsCustomCaptureWay)
            
        Case conMenu_Cap_After_Capture  '后台采集
            control.Enabled = mVideoCapture.IsStartup
            control.Visible = gobjCapturePar.IsUseAfterCapture And (Not IsCustomCaptureWay)
            
        Case conMenu_Cap_Import        '影像导入
            control.Enabled = Not mblnReadOnly
            
        Case conMenu_Cap_DelImg  '影像删除
            'dkpReportImage.FindPane(1).Selected = true 是指下方展现的标签是"检查图" ，功能：只要点击后台图标签，按钮就不可用
            control.Enabled = (mblnRealTime = False) And (ucPreview.ImgViewer.Images.Count > 0) And (Not mblnReadOnly) And Me.Visible And dkpReportImage.FindPane(1).Selected = True
            
        Case conMenu_Cap_Record        '录像
            control.Enabled = Not mblnReadOnly And ((gobjCapturePar.VideoDirverType = vdtWDM And mVideoCapture.IsStartup) Or gobjCapturePar.VideoDirverType = vdtCustom)
            control.Visible = Not IsTwainCaptureWay
            
        Case conMenu_Cap_After_Record   '后台录像
            control.Enabled = gobjCapturePar.VideoDirverType = vdtWDM And mVideoCapture.IsStartup
            control.Visible = Not IsTwainCaptureWay And Not IsCustomCaptureWay And gobjCapturePar.IsUseAfterCapture And False
            
        Case conMenu_Cap_Record_Stop '停止录像 modify by tjh at 2010-01-22
            control.Enabled = mblnRealTime And Not mblnReadOnly And (gobjCapturePar.VideoDirverType = vdtWDM) And mVideoCapture.IsStartup
            control.Visible = Not IsTwainCaptureWay And Not IsCustomCaptureWay
            
        Case conMenu_Cap_RecordAudio '录音
            control.Enabled = Not mblnReadOnly
            control.Visible = Not IsCustomCaptureWay
            
'        Case conMenu_Cap_Full_Screen '全屏 modify by tjh at 2010-01-22 (如果使用新的视频回放组件，则可以启用该功能)
'            control.Enabled = mblnRealTime And (Not mblnReadOnly) And Not GetIsTwainCaptureWay And mVideoCapture.IsStartup
'            control.Visible = Not GetIsTwainCaptureWay And mstrVideoRegTime <> ""
'
'        Case conMenu_Cap_DevSet        '设置（如果处于浮动状态时，则屏蔽该按钮） modify by tjh at 2010-01-25
'            control.Enabled = gobjCapturePar.IsUseStartupVideo And mstrVideoRegTime <> ""  'mblnEmbedded ' And (Not mblnReadOnly)
'
'            '如果为浮动窗体，则隐藏该设置按钮
'            'control.Visible = mstrVideoRegTime <> ""
'            If Not (mParentContainer Is Nothing) Then
'                If TypeOf mParentContainer Is frmVideoDockWindow Then
'                    control.Enabled = False
'                Else
'                    control.Enabled = True
'                End If
'            End If
            
        '录像播放,录像停止,录像快进,录像快退,保存录像
        Case conMenu_Cap_Play, conMenu_Cap_Stop, conMenu_Cap_Forward, _
             conMenu_Cap_Back
            If (mblnRealTime = False) And (dcmView.Images.Count > 0) Then
                control.Visible = dcmView.Images(1).Tag.Tag = VIDEOTAG Or dcmView.Images(1).Tag.Tag = AUDIOTAG
                control.Enabled = dcmView.Images(1).Tag.Tag = VIDEOTAG Or dcmView.Images(1).Tag.Tag = AUDIOTAG
            Else
                control.Visible = False
                control.Enabled = False
            End If
            
        Case conMenu_Cap_SaveAs
            control.Enabled = Me.Visible
            
         '亮度对比度,缩放,裁剪缩放,顺时针旋转,逆时针旋转,锐化,平滑,高级处理
        Case conMenu_Process_Window, conMenu_Process_Zoom, conMenu_Process_RectZoom, conMenu_Process_RRotate, _
             conMenu_Process_LRotate, conMenu_Process_Sharpness, conMenu_Process_Filter, conMenu_Process_Corp

            control.Enabled = (mblnRealTime = False)
        '箭头标注,圆形标注,文字标注,
        Case conMenu_Process_Arrow, conMenu_Process_Ellipse, conMenu_Process_Text
            control.Enabled = (mblnRealTime = False)
            
'        Case conMenu_Tool_Analyse
'            If mblnObserve Then
'                control.Enabled = Not mblnReadOnly
'            Else
'                control.Visible = False
'                control.Enabled = False
'            End If
'
            
        Case conMenu_Cap_OpenStudyList
            control.Enabled = True
            control.Visible = gobjCapturePar.IsUseCaptureLock
            
        Case conMenu_Cap_StudySyncState
            control.Enabled = Not mblnReadOnly Or mblnCaptureLockState
            control.Visible = gobjCapturePar.IsUseCaptureLock
            
        Case conMenu_Cap_After_Tag
            control.Enabled = mVideoCapture.IsStartup
            control.Visible = gobjCapturePar.IsUseAfterCapture
            
        ''''''''''''
        Case conMenu_Cap_ImgImport
            
        Case conMenu_Cap_ImgExport          '发送到后台
            
        Case conMenu_Cap_ReUpLoad, conMenu_Cap_ReDownLoad
            control.Visible = gtFileLoadType = Service
            control.Enabled = ucPreview.CurImageCount > 0
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

Private Sub ScanImages()
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


Private Function IsVerityCapture() As Boolean
'判断是否为正常的采集方式
    Dim strSQL As String
    Dim rsData As ADODB.Recordset
    
    IsVerityCapture = False
    
    '采集图像时，如果不是后台采集，则需判断当前加载的图像与数据库中的图像记录数是否一致，如果不一致，说明该检查当前可能正被其他设备站点采集
    '该处修改主要是防止设备操作技师误踩脚踏开关造成的图像采集
    strSQL = "select count(*) as 图像数 from 影像检查图象 where 序列uid in(select 序列UID from 影像检查序列 where 检查UID=(select 检查UID from 影像检查记录 where 医嘱id=[1])) "
    Set rsData = zlDatabase.OpenSQLRecord(strSQL, "查询图像数量", mlngAdviceId)
    
    If rsData.RecordCount > 0 Then
        If Val(Nvl(rsData!图像数)) <> ucPreview.ImageTotal Then
            Call MsgboxCus("当前加载的图像数量与实际记录数不一致，请检查是否另有用户对其进行操作，如无操作请刷新后重试。", vbInformation + vbOKOnly, G_STR_HINT_TITLE)
            Exit Function
        End If
    End If
    
    IsVerityCapture = True

End Function


Private Sub CaptureImage()
'************************************************************
'
'从视频或者录像中采集图像
'
'************************************************************
    Dim blnIsRealCap As Boolean
    
    If Not (Not mblnReadOnly And (mVideoCapture.IsStartup Or IsTwainCaptureWay Or IsCustomCaptureWay)) Then Exit Sub  '如果为只读，或者视频没有启动，则不允许采集
    
    If Not IsVerityCapture Then Exit Sub
            
    If IsTwainCaptureWay Then
        Call ScanImages  '通过TWAIN接口采集图像
    ElseIf IsCustomCaptureWay Then
        Call CustomCapture
    Else
        blnIsRealCap = mblnRealTime '为实时显示时自动采实时图
        
        If Not mblnRealTime Then
            blnIsRealCap = IIf(MsgboxCus("确定要采集当前静态图像吗？选“是”则采集当前处理图像。", vbQuestion + vbYesNo + vbDefaultButton1, G_STR_HINT_TITLE) = vbYes, False, True)
        End If
        
        '采集图像
        Call subCaptureImg(blnIsRealCap)
    End If
End Sub

Private Sub CustomCapture()
    Dim objCapPic As StdPicture
    Dim strCapImgFiles As String
    Dim blnUseCustom As Boolean
    
    If mobjCustomDevice Is Nothing Then Exit Sub
    
    '采集图像
    If Not mobjCustomDevice.zlCaptureImage(mlngAdviceId, objCapPic, strCapImgFiles, blnUseCustom) Then
        Exit Sub
    End If
    
    '保存扫描的图像
    Call subCaptureImg(True, strCapImgFiles, objCapPic, False, blnUseCustom)
  
    Call ShowScanImage(ucPreview.CurImageCount)
End Sub

Private Sub CaptureAfterImage()
'后台图像采集
    If Not mVideoCapture.IsStartup Then Exit Sub  '如果为只读，或者视频没有启动，则不允许采集,twain方式不允许后台采集
    
    Call subCaptureImg(True, "", Nothing, True)
End Sub


Public Sub zlExecuteCommandBars(control As XtremeCommandBars.CommandBarControl)
    On Error GoTo errHandle
        Call VideoCaptureMenuProcess(control)
        
        Call DicomImageMenuProcess(control)
        
        Call ControlImageMenuProcess(control)
    Exit Sub
errHandle:
    MsgboxCus err.Description, vbOKOnly, G_STR_HINT_TITLE
End Sub

Private Sub DoStateChange(ByVal lngEventType As TVideoEventType, ByVal lngAdviceId As Long, ByVal lngSendNo As Long, ByVal strOther As String)
'触发StateChange事件
On Error GoTo errHandle

BUGEX "DoStateChange(frmWork_Video) 1"
    RaiseEvent OnStateChange(lngEventType, lngAdviceId, lngSendNo, strOther)
    
BUGEX "DoStateChange(frmWork_Video) 2"
    '广播图像更新消息
    If lngEventType = vetCaptureFirstImg _
        Or lngEventType = vetDelAllImg _
        Or lngEventType = vetUpdateImg Then
        
BUGEX "DoStateChange(frmWork_Video) 3 PostMessage lngAdviceId:" & lngAdviceId
        '发送广播消息
        BoradcastMsg lngAdviceId
    End If
    
BUGEX "DoStateChange(frmWork_Video) End"
Exit Sub
errHandle:
    MsgboxCus err.Description, vbOKOnly, G_STR_HINT_TITLE
End Sub

Private Sub VideoCaptureMenuProcess(control As XtremeCommandBars.CommandBarControl)
'视频采集菜单处理
    Select Case control.id
        Case conMenu_Cap_Dynamic       '动态显示
            If IsTwainCaptureWay Then
                Call MsgboxCus("TWAIN采集模式下，不能进行动态视频的显示。", vbOKOnly, G_STR_HINT_TITLE)
            ElseIf IsCustomCaptureWay Then
                Call MsgboxCus("专用视频采集模式下，不能进行动态视频的显示。", vbOKOnly, G_STR_HINT_TITLE)
            Else
                Call ConfigVideoShowState(True)
            End If
            
        Case conMenu_Cap_MarkMap       '影像采集
            Call CaptureImage
            
        Case conMenu_Cap_After_Capture
            Call CaptureAfterImage
            
        Case conMenu_Cap_Import        '影像导入
            Call InputImageFile
            
        Case conMenu_Cap_DelImg  '影像删除
            Call subDeleteImage
            
        Case conMenu_Cap_Record        '录像
            If mstrVideoRegTime = "" Then
                MsgboxCus mstrMsg, vbOKOnly, "提示"
                Exit Sub
            End If
            
            If IsCustomCaptureWay Then
                Call CustomVideoSave
            Else
                Call subVideoSave
            End If
            
        Case conMenu_Cap_Record_Stop  '停止录像 modify by tjh at 2010-01-22
            If mstrVideoRegTime = "" Then
                'MsgboxCus  "未检测到有效的注册信息，不能进行录像操作！", vbOKOnly, "提示"
                Exit Sub
            End If
            
            Call subStopVideo
            
        Case conMenu_Cap_RecordAudio    '录音
            If mstrVideoRegTime = "" Then
                MsgboxCus mstrMsg, vbOKOnly, "提示"
                Exit Sub
            End If
            
            Call frmRecordAudio.ShowRecordAudio(Me)
            
'        Case conMenu_Cap_Full_Screen '全屏 modify by tjh at 2010-01-22
'            Call subFullCall
            
'        Case conMenu_Cap_DevSet        '设置
'            Call SaveParameterCfg
'            Call subVideoSetup
            
        Case conMenu_Cap_Play          '录像播放
            Call subVideoPlay
'
        Case conMenu_Cap_SaveAs        '文件另存
            Call subVideoSaveAs
            
'        Case conMenu_Process_Cursor
'            subSetMouseState 0
'            control.Checked = True
                
        Case conMenu_Cap_OpenStudyList      '打开检查采集图像
            Call OpenStudy
            
        Case conMenu_Cap_StudySyncState     '锁定检查
            If control.IconId = 10012 Then
                control.IconId = 8123
                
                Call LockStudy(0, 0, 0, 0, 0)
                
                Call DoStateChange(vetLockStudy, mlngAdviceId, mlngSendNo, mstrNameInf)
            Else
                control.IconId = 10012
                
                Call UnLockStudy
                
                If mlngAdviceId <> mSelectStudyInf.lngAdviceId Then
                    Call zlUpdateAdviceInf(mSelectStudyInf.lngAdviceId, mSelectStudyInf.lngSendNo, mSelectStudyInf.lngStudyState, mSelectStudyInf.blnMoved)
                    Call zlRefreshFace
                End If
                
                Call DoStateChange(vetUnLockStudy, mlngAdviceId, mlngSendNo, mstrNameInf)
            End If
        Case conMenu_Cap_After_Tag      '更新后台采集标识
            
            If mstrVideoRegTime = "" Then
                MsgboxCus mstrMsg, vbOKOnly, "提示"
                Exit Sub
            End If
            
            Call UpdateAfterCaptureInfo
            Call ShowAfterCaptureInf(True)
            
    End Select
End Sub

Private Sub DicomImageMenuProcess(control As XtremeCommandBars.CommandBarControl)
'dicom图像菜单处理
    If mblnRealTime = True Or dcmView.Images.Count <= 0 Then Exit Sub
    
    Select Case control.id
        Case conMenu_Process_Window         '亮度对比度
            subSetMouseState 1
            control.Checked = True
            
        Case conMenu_Process_Zoom           '缩放
            subSetMouseState 2
            control.Checked = True
            
        Case conMenu_Process_RectZoom       '裁剪缩放
            subSetMouseState 3
            control.Checked = True
        
        Case conMenu_Process_RectCapture         '裁剪后采集
            Call CaptureFrameSelectImage
            
        Case conMenu_Process_RRotate        '顺时针旋转
            Call subSetRotate(dcmView.Images(1), True)
            
        Case conMenu_Process_LRotate        '逆时针旋转
            Call subSetRotate(dcmView.Images(1), False)
            
        Case conMenu_Process_Sharpness      '锐化
            Call subSetSharp(dcmView.Images(1), True)
            
        Case conMenu_Process_Filter         '平滑
            Call subSetSharp(dcmView.Images(1), False)
            
        Case conMenu_Process_Corp          '拖动
           subSetMouseState 14
           control.Checked = True
            
        Case conMenu_Process_Arrow          '箭头标注
            subSetMouseState 11
            control.Checked = True
            
        Case conMenu_Process_Ellipse        '圆形标注
            subSetMouseState 12
            control.Checked = True
            
        Case conMenu_Process_Text           '文字标注
            subSetMouseState 13
            control.Checked = True
    End Select

End Sub

Private Sub ControlImageMenuProcess(control As XtremeCommandBars.CommandBarControl)
    Select Case control.id
        Case conMenu_Cap_SplitPage
            Call ucPreview.ShowPageConfig
            Call ucCacheImage.ShowPageConfig
            
        Case conMenu_Cap_ImgImport          '发送到检查
            mlngReleationType = 2
            If ReleationImage Then
                mAfterCaptureInf.lngAfterCurImageCount = GetCountOfCapPic(mintTagNow)
                Call ShowAfterCaptureInf(True)
                
                Call ClearEmptyFolder(False)
            End If
            
        Case conMenu_Cap_ImgExport          '发送到后台
            mlngReleationType = 1
            Call ReleationImage
            
        Case conMenu_Cap_AddToReport        '加入报告图
            Call AddImageToReport
            
        Case conMenu_Cap_ReUpLoad           '重新上传
            Call ucPreview_OnReUpload
            
        Case conMenu_Cap_ReDownLoad         '重新下载
            Call ucPreview.ReLoadFailedImage
            
        Case conMenu_Cap_DelCache
            mlngReleationType = 2
            Call DelTempImage
            Call ClearEmptyFolder(False)

            mAfterCaptureInf.lngAfterCurImageCount = GetCountOfCapPic(mintTagNow)
            Call ShowAfterCaptureInf(True)
            
            
        Case conMenu_Cap_RefreshCache
            Call LoadMiniCache
            
        Case conMenu_Cap_ImageProcess
            Call ucPreview.ShowImageProcess(ucPreview.SelectIndex, 3, 0)
        
        Case conMenu_Cap_ImageShield
            control.Checked = Not control.Checked
            
            mblnImageShield = control.Checked
            ucPreview.ImageShield = mblnImageShield
            Call SaveLocalPar
        
        Case conMenu_Cap_ImageDel           '删除图像
            If MsgBox("是否删除选中图像？", vbYesNo, G_STR_HINT_TITLE) = vbYes Then
                Call subDeleteImage
            End If
    End Select
End Sub

Private Function AddImageToReport() As Boolean
'------------------------------------------------
'功能：将图像暂时添加到报告中，设置报告图标记
'参数： lngActionType ： 1 - 加入报告图
'返回：True - 成功； False - 失败
'------------------------------------------------
    On Error GoTo err
    
    Dim i As Long
    Dim strSQL As String
    
    '提取图像UID
    For i = 1 To ucPreview.CurImageCount
        '修改报告图的tag，仅修改程序中的缓存数据
        '加入报告图时，如果>1，不用修改；=0不用修改，=""改成=0
        If ucPreview.ImgChecked(i) = True Then
            If ucPreview.ImgViewer.Images(i).Tag.ReportImage = "" Then
                ucPreview.ImgViewer.Images(i).Tag.ReportImage = 0
                Call AddToReport(ucPreview.ImgViewer.Images(i))
            End If
            AddImageToReport = True
        End If
    Next i
    
    If AddImageToReport = False Then
        MsgBox "请先选择一个图像。", vbOKOnly, "提示信息"
        Exit Function
    End If
    
    Exit Function
err:
    err.Raise err.Number, err.Description
End Function

Private Sub AddToReport(dcmImage As DicomImage)
    Dim strSQL As String
    '更新数据库
    strSQL = "Zl_影像检查_设置报告图('" & dcmImage.InstanceUID & "',1)"
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    '用事件更新报告图显示
    Call DoStateChange(vetAddReportImg, mlngAdviceId, mlngSendNo, dcmImage.InstanceUID)
    '画采集缩略图的报告图标记
    Call ucPreview.DrawReportImgTag(dcmImage)
End Sub

Private Function ReleationImage() As Boolean
    Dim strHint As String
    Dim rsImageDatas As ADODB.Recordset
    Dim strTmpFile As String
    Dim i As Integer
    
    Set rsImageDatas = GetReleationImageIds()
    
    If rsImageDatas Is Nothing Then
        Call MsgboxEx(Me.hwnd, "请选择需要进行关联的检查图像。", vbInformation, G_STR_HINT_TITLE)
        Exit Function
    End If
        
    '当前检查UID在数据库中不存在，则退出本程序
    If rsImageDatas.RecordCount <= 0 Then
        Call MsgboxEx(Me.hwnd, "请选择需要进行关联的检查图像。", vbInformation, G_STR_HINT_TITLE)
        Exit Function
    End If
    
    If mlngReleationType = 2 Then
        '关联图像提示
        strHint = GetReleationHintInfo(mlngAdviceId, rsImageDatas)
        
        If strHint = "" Then
            Call MsgboxEx(Me.hwnd, "不能查询到需要关联的数据信息，结束关联。", vbOKOnly, G_STR_HINT_TITLE)
            Exit Function
        End If
        
        If MsgboxEx(Me.hwnd, strHint, vbDefaultButton2 + vbQuestion + vbYesNo, G_STR_HINT_TITLE) = vbNo Then Exit Function
        
    Else
        '取消关联提示
        If MsgboxEx(Me.hwnd, "是否确认将所选图像发送到后台？", vbDefaultButton2 + vbQuestion + vbYesNo, G_STR_HINT_TITLE) = vbNo Then Exit Function
    End If

    If mlngReleationType = 2 Then '等于2表示关联图像
        ReleationImage = StartReleation(mlngAdviceId, rsImageDatas)
    Else
        ReleationImage = CancelReleation(mlngAdviceId, rsImageDatas)
    End If
        
    '清除红框，防止出现2个红框的BUG
    For i = 1 To ucPreview.CurImageCount
        ucPreview.ImgViewer.Images(i).BorderColour = vbWhite
    Next
    
    Call DoStateChange(IIf(mlngReleationType = 2, vetImportImage, vetExportImage), mlngAdviceId, mlngSendNo, mcurStudyInf.strStudyUid)
End Function

'取得关联提示信息
Private Function GetReleationHintInfo(lngAdviceId As Long, rsReleationImage As ADODB.Recordset) As String
    Dim rsTemp As ADODB.Recordset
    Dim strSQL As String
    
    GetReleationHintInfo = ""
    
    strSQL = "select 检查号,姓名,性别,年龄 from 影像检查记录 where 医嘱ID=[1]"
    
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngAdviceId)

    If rsTemp.RecordCount <= 0 Then Exit Function
    
    GetReleationHintInfo = "是否确认将选择的图像发送到[" & Nvl(rsTemp!姓名) & "(" & Nvl(rsTemp!检查号) & ") " & Nvl(rsTemp!性别) & " " & Nvl(rsTemp!年龄) & "]的检查中？"
End Function

Private Function StartReleation(ByVal lngAdviceId As Long, rsImageDatas As ADODB.Recordset) As Boolean
'开始关联
On Error GoTo errHandle
    Dim strSQL As String
    Dim strNewStudyUID As String, strNewSeriesUid As String
    Dim curDate As Date
    Dim blnBeginTrans As Boolean
    Dim rsTmp As ADODB.Recordset
    Dim objFtp As New clsFtp
    Dim objMoveList As New Collection
    Dim strNewDeviceNo As String, strNewFtpIp As String, strNewFtpUrl As String, strNewFtpVirtualPath As String, strNewFtpUser As String, strNewFtpPwd As String
    Dim arrSql() As String
    Dim i As Long
    
    blnBeginTrans = False
    StartReleation = False
    
    curDate = zlDatabase.Currentdate
    
    strSQL = "select 检查UID,接收日期 from 影像检查记录 where 医嘱ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngAdviceId)
    
    If rsTmp.RecordCount <= 0 Then
        Call MsgboxEx(Me.hwnd, "找不到待关联的检查信息。", vbInformation, G_STR_HINT_TITLE)
        Exit Function
    End If
    
    If Trim(Nvl(rsTmp!检查uid)) = "" Or Trim(Nvl(rsTmp!接收日期)) = "" Then

        '尚未采集图像，需要生成新的检查UID
        strNewStudyUID = CreateStudyUid(rsImageDatas!检查uid)
        
        Call GetStorageDevice(mlngAdviceId, strNewStudyUID, strNewDeviceNo, strNewFtpIp, strNewFtpUrl, strNewFtpVirtualPath, strNewFtpUser, strNewFtpPwd)
        If Trim(strNewFtpIp) = "" Then
            Call MsgboxEx(Me.hwnd, "不能取得有效的存储设备，请检查存储设备配置。", vbInformation, G_STR_HINT_TITLE)
            Exit Function
        End If
        
        '更新存储设备信息
        strSQL = "Zl_影像检查_更新设备(" & mlngAdviceId & ",'" & strNewStudyUID & "','" & strNewDeviceNo & "'," & _
                                        "to_Date('" & Format(curDate, "yyyy-mm-dd hh:mm") & "','YYYY-MM-DD HH24:MI:SS'))"
        Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    Else
        strNewStudyUID = Nvl(rsTmp!检查uid)
        
        Call GetStorageDevice(mlngAdviceId, strNewStudyUID, strNewDeviceNo, strNewFtpIp, strNewFtpUrl, strNewFtpVirtualPath, strNewFtpUser, strNewFtpPwd)
        If Trim(strNewFtpIp) = "" Then
            Call MsgboxEx(Me.hwnd, "不能取得有效的存储设备，请检查存储设备配置。", vbInformation, G_STR_HINT_TITLE)
            Exit Function
        End If
    End If
    
    '连接FTP
    If objFtp.FuncFtpConnect(strNewFtpIp, strNewFtpUser, strNewFtpPwd) = 0 Then
        Call MsgboxEx(Me.hwnd, "FTP连接失败，请检查网络设置。", vbInformation, G_STR_HINT_TITLE)
        Exit Function
    End If

    ReDim arrSql(0)
    rsImageDatas.MoveFirst
    While Not rsImageDatas.EOF
        '创建新的序列UID
        strNewSeriesUid = CreateSeriesUid(rsImageDatas!序列Uid, strNewStudyUID)
        
        ReDim Preserve arrSql(UBound(arrSql) + 1)
        arrSql(UBound(arrSql)) = "Zl_影像检查_图像导入(" & mlngAdviceId & ",'" & strNewStudyUID & "','" & strNewSeriesUid & "','" & mobjFile.GetFileName(Nvl(rsImageDatas!路径)) & "')"
        '修改本地文件
        Call UpdateAfterImage(Nvl(rsImageDatas!路径), strNewStudyUID, strNewSeriesUid)
        rsImageDatas.MoveNext
    Wend
    
    '移动图像文件
    If Not MoveImageToStudy(objFtp, rsImageDatas, strNewFtpVirtualPath, objMoveList) Then Exit Function
    
    gcnVideoOracle.BeginTrans
    
    blnBeginTrans = True
    
    For i = 1 To UBound(arrSql)
        '更新图像关联数据
        Call zlDatabase.ExecuteProcedure(arrSql(i), Me.Caption)
    Next
    
    '提交事务
    Call gcnVideoOracle.CommitTrans
    
    '说明全部上传成功,删除本地临时图像
    Call DelTempImages(rsImageDatas)
    
    StartReleation = True
    
    Exit Function
errHandle:
    If blnBeginTrans Then Call gcnVideoOracle.RollbackTrans
    
    Call CancelImageMove(strNewFtpIp, strNewFtpUser, strNewFtpPwd, objMoveList)
    Call OutputDebug("StartReleation", err)
    
    Call RaiseErr(err)  '继续抛出错误
End Function

Private Sub UpdateAfterImage(strTmpFile As String, strStudyUid As String, strSeriesUID As String)
'修改本地文件
    Dim curImage As DicomImage
    Dim dcmViewer As DicomViewer
    
    If Len(Dir(strTmpFile)) = 0 Then Exit Sub
    
    Set curImage = ReadViewImage(strTmpFile, dcmViewer)
    curImage.StudyUID = strStudyUid
    curImage.SeriesUID = strSeriesUID
    curImage.WriteFile strTmpFile, True
End Sub

Private Function CancelReleation(ByVal lngAdviceId As Long, rsImageDatas As ADODB.Recordset) As Boolean
'撤销关联
On Error GoTo errHandle
    Dim strSQL As String
    Dim strNewStudyUID As String, strNewSeriesUid As String
    Dim curDate As Date
    Dim objFtp As New clsFtp
    Dim blnBeginTrans As Boolean
    Dim rsTmp As ADODB.Recordset
    Dim rsReportImage As ADODB.Recordset
    Dim objMoveList As New Collection
    Dim strNewDeviceNo As String, strNewFtpIp As String, strNewFtpUrl As String, strNewFtpVirtualPath As String, strNewFtpUser As String, strNewFtpPwd As String
    
    blnBeginTrans = False
    CancelReleation = False
    
    curDate = zlDatabase.Currentdate
    
    '撤销图像关联
    strNewStudyUID = CreateStudyUid(mdcmUID.NewUID)
        
    Call GetStorageDevice(mlngAdviceId, strNewStudyUID, strNewDeviceNo, strNewFtpIp, strNewFtpUrl, strNewFtpVirtualPath, strNewFtpUser, strNewFtpPwd)
    If Trim(strNewFtpIp) = "" Then
        Call MsgboxEx(Me.hwnd, "不能取得有效的存储设备，请检查存储设备配置。", vbInformation, G_STR_HINT_TITLE)
        Exit Function
    End If
    
    '连接FTP
    If objFtp.FuncFtpConnect(strNewFtpIp, strNewFtpUser, strNewFtpPwd) = 0 Then
        Call MsgboxEx(Me.hwnd, "FTP连接失败，请检查网络设置。", vbInformation, G_STR_HINT_TITLE)
        Exit Function
    End If
    
    If Not MoveImageToAfter(objFtp, rsImageDatas, objMoveList) Then Exit Function

    gcnVideoOracle.BeginTrans
    
    blnBeginTrans = True
    
    '更新数据
    rsImageDatas.MoveFirst
    While Not rsImageDatas.EOF
        strSQL = "Zl_影像检查_图像导出(" & mlngAdviceId & ",'" & Nvl(rsImageDatas!图像UID) & "')"
         
        Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
        
        rsImageDatas.MoveNext
    Wend
    
    Call gcnVideoOracle.CommitTrans
    
    Call ClearFtpImage(rsImageDatas, strNewStudyUID)
    
    CancelReleation = True
Exit Function
errHandle:
    If blnBeginTrans Then Call gcnVideoOracle.RollbackTrans
    Call OutputDebug("CancelReleation", err)
    Call RaiseErr(err)
End Function

'撤销图像的移动
Private Sub CancelImageMove(ByVal strFTPIP As String, ByVal strFTPUser As String, ByVal strFTPPwd As String, objMoveList As Collection)
    Dim i As Long
    Dim objFtp As New clsFtp
    Dim strDestFile As String
    Dim strMoveFile As String
    
    If objMoveList Is Nothing Then Exit Sub
    If objMoveList.Count <= 0 Then Exit Sub
    
On Error GoTo errHandle

    Call objFtp.FuncFtpConnect(strFTPIP, strFTPUser, strFTPPwd)
    
    For i = 1 To objMoveList.Count
        strDestFile = objMoveList.Item(i)
        
        strMoveFile = Mid(strDestFile, InStr(strDestFile, ">") + 1, 255)
        strDestFile = Mid(strDestFile, 1, InStr(strDestFile, ">") - 1)
        
        Call objFtp.FuncReNameFile(strMoveFile, strDestFile)
    Next i
        
errHandle:
    objFtp.FuncFtpDisConnect
End Sub

Private Function CreateSeriesUid(ByVal strSeriesUID As String, ByVal strStudyUid As String) As String
'创建序列UID
    Dim rsData As New ADODB.Recordset
    Dim strSQL As String
    Dim strNewSeriesUid As String
    
    strNewSeriesUid = strSeriesUID 'M_STR_SERIES_UID & "." & Format(Now, "yymmddhhmmss") & "." & Fix(Rnd(1000) * 1000)
    
    strSQL = "select 序列UID from 影像检查序列 where 序列UID = [1] And 检查UID <> [2]"
              
    Set rsData = zlDatabase.OpenSQLRecord(strSQL, "PACS图像保存", strNewSeriesUid, strStudyUid)
    
    If rsData.RecordCount > 0 Then
        '创建一个新的检查UID
        strSQL = "Select 影像检查UID序号_ID.Nextval From Dual"
        Set rsData = zlDatabase.OpenSQLRecord(strSQL, "PACS图像保存")
        
        If Len(strNewSeriesUid) <= 55 Then
            strNewSeriesUid = strNewSeriesUid & ".A" & rsData(0)
        Else
            strNewSeriesUid = Left(strNewSeriesUid, 55) & ".A" & rsData(0)
        End If
    End If
    
    CreateSeriesUid = strNewSeriesUid
End Function

Private Sub ClearFtpImage(rsImageDatas As ADODB.Recordset, ByVal strNewStudyUID As String)
On Error GoTo errHandle
'转移图像成功后，在删除临时图像和原有FTP的图像和目录，清场操作出现错误可以不处理
    Dim objSrcFtp As New clsFtp
    Dim strTmpFile As String
    Dim strVirtualPath As String
    Dim strImageUID As String
    Dim strCurFtpIp As String, strCurFtpUser As String, strCurFtpPwd As String
    
    strCurFtpIp = ""
    strCurFtpUser = ""
    strCurFtpPwd = ""
    
    Call rsImageDatas.MoveFirst
    
    While Not rsImageDatas.EOF
        strTmpFile = GetAppPath & "\TmpImage\" & Nvl(rsImageDatas!图像UID)
        
        strImageUID = Nvl(rsImageDatas!图像UID)
        
        strVirtualPath = Nvl(rsImageDatas!Root) & Nvl(rsImageDatas!Url)
                
        If strCurFtpIp <> Nvl(rsImageDatas!host) Or strCurFtpUser <> Nvl(rsImageDatas!FtpUser) Or strCurFtpPwd <> Nvl(rsImageDatas!FtpPwd) Then
            strCurFtpIp = Nvl(rsImageDatas!host)
            strCurFtpUser = Nvl(rsImageDatas!FtpUser)
            strCurFtpPwd = Nvl(rsImageDatas!FtpPwd)
            
            Call objSrcFtp.FuncFtpConnect(strCurFtpIp, strCurFtpUser, strCurFtpPwd)
        End If
        
'       为避免重新下载图像，如果本地存在图像文件，则不用进行删除
        
        If FileExists(strTmpFile) Then Call Kill(strTmpFile)
        If FileExists(strTmpFile & ".jpg") Then Call Kill(strTmpFile & ".jpg")
        
        Call objSrcFtp.FuncDelFile(strVirtualPath, strImageUID)
        
        '删除空的ftp目录
        Call objSrcFtp.FuncFtpDelDir(Replace(strVirtualPath, strImageUID, ""), strImageUID)
                
        rsImageDatas.MoveNext
        
        If Not rsImageDatas.EOF Then
            If strCurFtpIp <> Nvl(rsImageDatas!host) Or strCurFtpUser <> Nvl(rsImageDatas!FtpUser) Or strCurFtpPwd <> Nvl(rsImageDatas!FtpPwd) Then
                Call objSrcFtp.FuncFtpDisConnect
            End If
        End If
    Wend
    
    objSrcFtp.FuncFtpDisConnect
Exit Sub
errHandle:
    Call OutputDebug("ClearFtpImage", err)
End Sub

Private Function MoveImageToStudy(objFtp As clsFtp, rsImageDatas As ADODB.Recordset, strNewFtpVirtualPath As String, ByRef objMoveList As Collection) As Boolean
    Dim i As Long

    rsImageDatas.MoveFirst
    While Not rsImageDatas.EOF
        If objFtp.FuncUploadFile(strNewFtpVirtualPath, rsImageDatas!路径, mobjFile.GetFileName(rsImageDatas!路径)) <> 0 Then
            '失败后删除之前上传的文件
            For i = 0 To objMoveList.Count - 1
                Call objFtp.FuncDelFile(strNewFtpVirtualPath, objMoveList(i))
            Next
            
            Call MsgboxEx(Me.hwnd, "图像移动失败，请检查FTP传输是否正常。", vbInformation, G_STR_HINT_TITLE)
            
            Exit Function
        Else
            Call objMoveList.Add(rsImageDatas!路径)
        End If
        
        rsImageDatas.MoveNext
    Wend
    
    MoveImageToStudy = True
End Function

Private Function MoveImageToAfter(objFtp As clsFtp, rsImageDatas As ADODB.Recordset, ByRef objMoveList As Collection) As Boolean
    Dim i As Long
    Dim strDestPath As String
    
    rsImageDatas.MoveFirst

    While Not rsImageDatas.EOF
        strDestPath = GetAfterImagePath(rsImageDatas!图像UID, rsImageDatas!序列Uid, rsImageDatas!检查uid)
        If mobjFile.FolderExists(strDestPath) = False Then Call MkLocalDir(strDestPath)
        
        If objFtp.FuncDownloadFile(rsImageDatas!Root & rsImageDatas!Url, strDestPath & rsImageDatas!图像UID, rsImageDatas!图像UID) <> 0 Then
            '失败后删除之前下载的文件
            For i = 0 To objMoveList.Count - 1
                Call mobjFile.DeleteFile(objMoveList(i))
            Next

            Call MsgboxEx(Me.hwnd, "图像移动失败，请检查FTP传输是否正常。", vbInformation, G_STR_HINT_TITLE)
            
            Exit Function
        Else
            Call objMoveList.Add(strDestPath & rsImageDatas!图像UID)
        End If
        
        rsImageDatas.MoveNext
    Wend
    Call MsgboxEx(Me.hwnd, "已将选中图像发送到[检查" & mintTagNow & "]中", vbInformation, G_STR_HINT_TITLE)
    
    MoveImageToAfter = True
End Function

Public Function GetAfterImagePath(ByVal strImageName As String, ByVal strSeriesUID As String, ByVal strStudyUid As String) As String
'发送到后台获得目录 若已经存在该标识的目录 使用之，否则创建
    Dim strTmpPath As String
    Dim objFolder1 As Folder, objFolder2 As Folder, objFolder3 As Folder
    Dim curDate As Date
    Dim strDate As String
    Dim intTmp As Integer
    
    curDate = zlDatabase.Currentdate
    strDate = Format(curDate, "yyyymmdd")
    
    strTmpPath = ""
    
    If mobjFile.FolderExists(mstrAfterDir & "\") Then
        For Each objFolder1 In mobjFile.GetFolder(mstrAfterDir & "\").SubFolders   '时间层
            If objFolder1.Name = strDate Then '时间只判断当天

                For Each objFolder2 In mobjFile.GetFolder(objFolder1.Path).SubFolders   '检查层
                
                    If InStr(objFolder2.Name, "检查" & mintTagNow) > 0 Then '判断是否有这个检查+当前标识的目录，若有，直接使用，
                        
                        For Each objFolder3 In mobjFile.GetFolder(objFolder2.Path).SubFolders   '序列层
                                strTmpPath = objFolder3.Path & "\"
                                GoTo step2
                        Next
                   
                    End If
                Next
                
                Exit For '终止时间层文件夹的搜索
            End If
        Next
    End If
    

    If strTmpPath = "" Then
        mAfterCaptureInf.strAfterStudyUid = ""
        mintTagNow = mintTagNow + 1
        mintTagMax = mintTagMax + 1
        strTmpPath = mstrAfterDir & "\" & Format(curDate, "yyyymmdd") & "\" & "检查" & mintTagNow & "-" & strStudyUid & "\" & strSeriesUID & "\"
    End If
    
    '找到目录后停止前面的遍历，直接进入step2
step2:
    
    intTmp = mintTagNow + 1
    
    Call ShowAfterCaptureInf(True)
    
    GetAfterImagePath = strTmpPath
End Function


Private Sub GetStorageDevice(ByVal lngAdviceId As Long, ByVal strNewStudyUID As String, _
    ByRef strDeviceNO As String, ByRef strFTPIP As String, _
    ByRef strFtpUrl As String, ByRef strVirtualPath As String, _
    ByRef strFTPUser As String, ByRef strFTPPwd As String)
'获取新的存储设备信息，如果设备存储信息部存在，则需要进行增加
'如果是取消关联，则使用strNewStudyUID将不能从数据库中查找到对应的数据
'strDeviceNum:设备号
'strFtpIp: ftp地址
'strFtpUrl: ftp目录
'strVirtualPath: ftp虚拟存储路径
'strFtpUser: ftp用户名
'strFtpPwd: ftp密码

    Dim strSQL As String
    Dim rsData As ADODB.Recordset
    Dim rsTemp As ADODB.Recordset
    Dim blnIsGetNewDevice As Boolean
    Dim objDestFtp As New clsFtp
    Dim curDate As Date
    
    strFTPIP = ""
    strFtpUrl = ""
    strFTPUser = ""
    strFTPPwd = ""
    strDeviceNO = ""
    
    strSQL = "Select D.FTP用户名 As FtpUser,D.FTP密码 As FtpPwd,C.位置一,C.位置二,C.位置三,C.接收日期," & _
        "D.IP地址 As Host," & _
        "'/'||D.Ftp目录||'/' As Root,Decode(C.接收日期,Null,'',to_Char(C.接收日期,'YYYYMMDD')||'/')" & _
        "||C.检查UID As URL " & _
        "From 影像检查记录 C,影像设备目录 D " & _
        "Where Decode(C.位置一,Null,C.位置二,C.位置一)=D.设备号(+)" & _
        "And C.检查UID= [1]"
        
    If mblnMoved Then
        strSQL = Replace(strSQL, "影像检查记录", "H影像检查记录")
    End If
        
    Set rsData = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNewStudyUID)
    
    blnIsGetNewDevice = False
    
    If rsData.RecordCount <= 0 Then
        blnIsGetNewDevice = True
    Else
        '如果执行到这里，说明是执行图像关联,需要判断当前检查的存储设备是否有效，如果无效需生成新的存储设备
        If Trim(rsData!接收日期) = "" Then
            blnIsGetNewDevice = True
        Else
            strDeviceNO = Nvl(rsData!位置一)
            strFTPIP = Nvl(rsData!host)
            strFtpUrl = Nvl(rsData!Root)
            strFTPUser = Nvl(rsData!FtpUser)
            strFTPPwd = Nvl(rsData!FtpPwd)
            strVirtualPath = strFtpUrl & Nvl(rsData!Url)
        End If
    End If
    
    If blnIsGetNewDevice Then
        '生成新的检查UID和存储设备,如果执行到这里，说明是取消关联
        
        '查询非医技工作站中的图像存储设备
        strDeviceNO = GetDeptPara(glngDepartId, "存储设备号")
        
        If Val(strDeviceNO) <= 0 Then
            MsgboxEx Me.hwnd, "未找到图像存储设备,请确认在影像流程管理中是否对该科室配置了图像采集存储设备。", vbInformation, G_STR_HINT_TITLE
            Exit Sub
        End If
        
        strSQL = "Select 设备号,设备名,'/'||Decode(Ftp目录,Null,'',Ftp目录||'/') As URL,FTP用户名,FTP密码,IP地址 " & _
                    " From 影像设备目录 Where 类型=1 and 设备号=[1] and NVL(状态,0)=1"
                    
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Tag, strDeviceNO)
        
        '如果存储设备停用，则直接退出
        If rsTemp.RecordCount <= 0 Then
            MsgboxEx Me.hwnd, "未找到存储设备,请确认设备号为 [" & strDeviceNO & "] 的设备是否启用。", vbInformation, G_STR_HINT_TITLE
            Exit Sub
        End If
        
        strFtpUrl = Nvl(rsTemp("URL"))
        strFTPIP = Nvl(rsTemp("IP地址"))
        strFTPUser = Nvl(rsTemp("FTP用户名"))
        strFTPPwd = Nvl(rsTemp("FTP密码"))
        
        strFtpUrl = IIf(strFtpUrl = "/", "//", strFtpUrl)
        
        objDestFtp.FuncFtpConnect strFTPIP, strFTPUser, strFTPPwd
        On Error GoTo errHandle
            curDate = zlDatabase.Currentdate
            
            strVirtualPath = strFtpUrl & Format(curDate, "YYYYMMDD") & "/" & strNewStudyUID
            '创建FTP目录
            objDestFtp.FuncFtpMkDir strFtpUrl, Format(curDate, "YYYYMMDD") & "/" & strNewStudyUID
            
        Call objDestFtp.FuncFtpDisConnect
errHandle:
        Call objDestFtp.FuncFtpDisConnect
    End If
End Sub


Private Sub DelTempImage()
    Dim rsImageDatas As ADODB.Recordset
    Dim i As Long
    
    '在数据库中查询图像数据
    Set rsImageDatas = GetReleationImageIds()
    
    If rsImageDatas Is Nothing Then
        Call MsgboxEx(Me.hwnd, "请选择需要删除的检查图像。", vbInformation, G_STR_HINT_TITLE)
        Exit Sub
    End If
    
    '当前检查UID在数据库中不存在，则退出本程序
    If rsImageDatas.RecordCount <= 0 Then
        Call MsgboxEx(Me.hwnd, "请选择需要删除的检查图像。", vbInformation, G_STR_HINT_TITLE)
        Exit Sub
    End If
    
    If MsgboxEx(Me.hwnd, "是否确认删除所选图像？", vbDefaultButton2 + vbQuestion + vbYesNo, G_STR_HINT_TITLE) = vbNo Then Exit Sub
    
    Call DelTempImages(rsImageDatas)
End Sub

Private Function DelTempImages(rsImageDatas As ADODB.Recordset) As Boolean
'删除本地缓存中的文件，在界面上删除ucpre控件的选中图像
    Dim blfinished As Boolean
    Dim i As Long
    Dim curTime As Date
    Dim intTmp As Integer
    
On Error GoTo errHandle
    If rsImageDatas.RecordCount <= 0 Then Exit Function
    
    Call rsImageDatas.MoveFirst
    
    While Not rsImageDatas.EOF
        If mobjFile.FileExists(rsImageDatas!路径) Then Call mobjFile.DeleteFile(rsImageDatas!路径)
        
        rsImageDatas.MoveNext
    Wend
    
   
    '删除界面图像
    blfinished = False
    
    For i = ucCacheImage.CurImageCount To 1 Step -1
        If ucCacheImage.ImgChecked(i) Then
            Call ucCacheImage.DeleteImage(i)
            blfinished = True
        End If
    Next
    

    
    If blfinished = False Then
        Call ucCacheImage.DeleteImage(ucCacheImage.SelectIndex)
    End If
    
    '同时需要删除cbo项目
    
    If ucCacheImage.CurImageCount = 0 Then
        curTime = zlDatabase.Currentdate
        '是当天并且选中的是当前标识，就不进行清空操作
        If Not ((Format(DTPImg.value, "yyyymmdd") = Format(curTime, "yyyymmdd")) And (InStr(cboCacheImage.Text, "标识" & Lpad((mintTagNow), 3, "0")) > 0)) Then
            intTmp = cboCacheImage.ListIndex
            cboCacheImage.RemoveItem (cboCacheImage.ListIndex)
            If cboCacheImage.ListCount > intTmp - 1 Then
                cboCacheImage.ListIndex = intTmp - 1
            Else
                cboCacheImage.ListIndex = 0
            End If
        End If
    End If

    DelTempImages = True
    
    Exit Function
errHandle:
    Call err.Raise(err.Number, err.Source, err.Description, err.HelpFile, err.HelpContext)
End Function

Private Function GetReleationImageIds() As ADODB.Recordset
'查询关联或者要取消关联的图像ID
    Dim i As Long, j As Long
    Dim strSQL As String
    Dim strValues(0 To 80) As String
    Dim strValue As String
    Dim strUninTable As String
    Dim strFilter As String
    Dim strTmpFile As String
    Dim rsImageDatas As ADODB.Recordset

    j = 0
    strUninTable = ""
    strFilter = ""
    strValue = ""
    
    '构造查询语句
    If mlngReleationType = 1 Then
        For i = 1 To ucPreview.CurImageCount
            If ucPreview.ImgChecked(i) Then
                If j > 79 Then
                    strFilter = strFilter & " Or 图像UID ='" & ucPreview.ImgViewer.Images(i).InstanceUID & "'"
                Else
                    If zlCommFun.ActualLen(strValue) > 3600 Then
                         strValues(j) = Mid(strValue, 2)
                         strUninTable = strUninTable & " Union ALL  Select  Column_Value as 图像UID From Table(Cast(f_Str2list([" & j + 1 & "]) As zlTools.t_Strlist))  " & vbCrLf
                         
                         strValue = ""
                         j = j + 1
                    End If
                    
                    strValue = strValue & "," & ucPreview.ImgViewer.Images(i).InstanceUID
                End If
            End If
        Next
                
        '若所有图像都没有被选中的红点，则有红框的图像视为选中
        If Not ucPreview.SelectImage Is Nothing And strValue = "" Then strValue = strValue & "," & ucPreview.SelectImage.InstanceUID
                
    Else
        Set rsImageDatas = New ADODB.Recordset
        rsImageDatas.Fields.Append "序列UID", adVarChar, 4000
        rsImageDatas.Fields.Append "检查UID", adVarChar, 4000
        rsImageDatas.Fields.Append "路径", adVarChar, 4000
        rsImageDatas.Open
            
        For i = 1 To ucCacheImage.CurImageCount
            If ucCacheImage.ImgChecked(i) Then
                strTmpFile = ucCacheImage.ImgViewer.Images(i).Tag.FilePath
                rsImageDatas.AddNew
                rsImageDatas!序列Uid = mobjFile.GetFolder(mobjFile.GetParentFolderName(strTmpFile)).Name
                rsImageDatas!检查uid = GetStudyUIDFromFolderName(mobjFile.GetFolder(mobjFile.GetParentFolderName(mobjFile.GetParentFolderName(strTmpFile))).Name)
                'rsImageDatas!检查uid = mobjFile.GetFolder(mobjFile.GetParentFolderName(mobjFile.GetParentFolderName(strTmpFile))).Name
                rsImageDatas!路径 = strTmpFile
                rsImageDatas.Update
            End If
        Next
        
        '没有图像处于红点状态，就选择有红框的
        If rsImageDatas.RecordCount = 0 Then
            If ucCacheImage.CurImageCount > 0 Then
                If Not ucCacheImage.SelectImage Is Nothing Then
                    strTmpFile = ucCacheImage.SelectImage.Tag.FilePath
                    rsImageDatas.AddNew
                    rsImageDatas!序列Uid = mobjFile.GetFolder(mobjFile.GetParentFolderName(strTmpFile)).Name
                    rsImageDatas!检查uid = GetStudyUIDFromFolderName(mobjFile.GetFolder(mobjFile.GetParentFolderName(mobjFile.GetParentFolderName(strTmpFile))).Name)
                    rsImageDatas!路径 = strTmpFile
                    rsImageDatas.Update
                End If
            End If
        End If
                
        If rsImageDatas.RecordCount > 0 Then rsImageDatas.MoveFirst
        
        Set GetReleationImageIds = rsImageDatas
        Exit Function
    End If
    
    If strValue <> "" Then
        strValues(j) = Mid(strValue, 2)
        strUninTable = strUninTable & " Union ALL  Select  Column_Value as 图像UID From Table(Cast(f_Str2list([" & j + 1 & "]) As zlTools.t_Strlist))  " & vbCrLf
    End If
    
    '如果没有需要查找的图像UID，则返回空数据集
    If strUninTable <> "" Then
        strUninTable = Mid(strUninTable, 11)
    Else
        Set GetReleationImageIds = Nothing
        Exit Function
    End If
    
    If strFilter <> "" Then strUninTable = strUninTable & " Union All Select 图像UID from [影像图象] where  ( " & Mid(strFilter, 4) & ")"
    
    strSQL = "Select /*+ RULE*/ D.FTP用户名 As FtpUser,D.FTP密码 As FtpPwd, Decode(C.位置一,Null,C.位置二,C.位置一) as 设备号," & _
        "D.IP地址 As Host,B.序列UID,B.检查UID,C.影像类别, " & _
        "'/'||D.Ftp目录||'/' As Root,Decode(C.接收日期,Null,'',to_Char(C.接收日期,'YYYYMMDD')||'/')" & _
        "||C.检查UID As URL,A.图像UID, c.姓名,c.性别,c.年龄,c.检查号 " & _
        "From 影像检查图象 A, 影像检查序列 B, 影像检查记录 C,影像设备目录 D,(" & Replace(strUninTable, "[影像图象]", "影像检查图象") & ") E " & _
        "Where Decode(C.位置一,Null,C.位置二,C.位置一)=D.设备号(+)" & _
        "And A.序列UID=B.序列UID and B.检查UID=C.检查UID and A.图像UID = E.图像UID "
        
    If mblnMoved Then
        strSQL = Replace(strSQL, "影像检查图象", "H影像检查图象")
        strSQL = Replace(strSQL, "影像检查序列", "H影像检查序列")
        strSQL = Replace(strSQL, "影像检查记录", "H影像检查记录")
    End If
    
    Set GetReleationImageIds = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strValues(0), strValues(1), strValues(2), strValues(3), _
        strValues(4), strValues(5), strValues(6), strValues(7), strValues(8), strValues(9), strValues(10), _
        strValues(11), strValues(12), strValues(13), strValues(14), strValues(15), strValues(16), strValues(17), _
        strValues(18), strValues(19), strValues(20), strValues(21), strValues(22), strValues(23), strValues(24), strValues(25), strValues(26), _
        strValues(27), strValues(28), strValues(29), strValues(30), strValues(31), strValues(32), strValues(33), strValues(34), strValues(35), strValues(36), _
        strValues(37), strValues(38), strValues(39), strValues(40), strValues(41), strValues(42), strValues(43), strValues(44), strValues(45), strValues(46), _
        strValues(47), strValues(48), strValues(49), strValues(50), strValues(51), strValues(52), strValues(53), strValues(54), strValues(55), strValues(56), _
        strValues(57), strValues(58), strValues(59), strValues(60), strValues(61), strValues(62), strValues(63), strValues(64), strValues(65), strValues(66), _
        strValues(67), strValues(68), strValues(69), strValues(70), strValues(71), strValues(72), strValues(73), strValues(74), strValues(75), strValues(76), _
        strValues(77), strValues(78), strValues(79), strValues(80))
End Function

Private Sub OpenStudy()
    Dim cbrControl As CommandBarControl
    
    Dim lngCurAdviceId As Long
    Dim lngSendNo As Long
    Dim lngStudyState As Long
    Dim blnResult As Boolean
    
    
    If mfrmOpenStudy Is Nothing Then Set mfrmOpenStudy = New frmOpenStudyList
    
    blnResult = mfrmOpenStudy.ShowStudyWindow(lngCurAdviceId, lngSendNo, lngStudyState, Me)
    
    If blnResult = False Then Exit Sub
        
    If lngCurAdviceId > 0 Then
        '开始打开新的检查进行采集
        Call UnLockStudy
        
        Call zlUpdateAdviceInf(lngCurAdviceId, lngSendNo, lngStudyState, 0)
        Call zlRefreshFace
        
        Call LockStudy(0, 0, 0, 0, 0)
                
        '修改锁定状态
        Set cbrControl = cbrMain.FindControl(, conMenu_Cap_StudySyncState)
        cbrControl.IconId = 8123
        
        '触发病人改变事件
        Call DoStateChange(vetLockStudy, mlngAdviceId, mlngSendNo, mstrNameInf)
    End If
    
End Sub


Public Sub zlUnloadMe()
    Unload Me
End Sub


Private Sub cbrMain_Execute(ByVal control As XtremeCommandBars.ICommandBarControl)
    zlExecuteCommandBars control
End Sub



Private Sub cbrMain_Resize()
    If cbrMain.Count >= 3 Then
        If cbrMain.Item(3).Visible <> mblnShowProcessBar Then
            mblnShowProcessBar = cbrMain.Item(3).Visible
        End If
    End If
    
    If cbrMain.Count >= 3 Then
        If picCapture.Width < 4000 Then
            cbrMain.Item(2).position = xtpBarTop
            cbrMain.Item(3).position = xtpBarBottom
        Else
            cbrMain.Item(2).position = xtpBarLeft
            cbrMain.Item(3).position = xtpBarRight
        End If
    End If
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
    Call ConfigVideoDisplay(picVideo)
    
    '刷新视频显示
    If Not (mVideoCapture Is Nothing) Then
        Call mVideoCapture.RefreshVideoWindow
    End If
    
    '刷新DcmView中的图像显示位置
    If dcmView.Images.Count > 0 Then
        Call subCenterZoom(Me, dcmView.Images(1), dcmView, dcmView.Images(1).ActualZoom, mCorpSize)
    End If

    '刷新裁剪边线位置
    Call RefreshPbxSizePos
        
    
    If IsTwainCaptureWay Or IsCustomCaptureWay Then
      
        '调整图像的显示范围
        dcmView.Left = Left
        dcmView.Top = Top
        dcmView.Width = Right - Left
        dcmView.Height = Bottom - Top
  
        '刷新DcmView中图像的显示位置
        If dcmView.Images.Count > 0 Then
            Call subCenterZoom(Me, dcmView.Images(1), dcmView, dcmView.Images(1).ActualZoom, mCorpSize)
        End If
    
    End If
End Sub


Private Sub cbrMain_Update(ByVal control As XtremeCommandBars.ICommandBarControl)
    zlUpdateCommandBars control
End Sub


Private Sub commListener_OnComm()
On Error GoTo errHandle
    Dim strInput As String
    
    '如果是TWAIN扫描或专用视频采集，则不支持脚踏开关采集
    If IsTwainCaptureWay Or IsCustomCaptureWay Then Exit Sub
    
    If gobjCapturePar.ComPortType <> "COM" Then Exit Sub
    
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
                    Call subCaptureImg(True)
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
            Call subCaptureImg(True)
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
                Call subCaptureImg(True)
            End If
        End If
    End If
errHandle:
End Sub


Private Sub dcmView_DblClick()
On Error GoTo errHandle
    Call subVideoPlay
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
  
  picView.Left = mVideoArea.Left + (mVideoArea.Width - picView.Width - 2 * pbxSize(2).Width) / 2 + 3 * pbxSize(2).Width
  picView.Top = mVideoArea.Top + (mVideoArea.Height - picView.Height - 2 * pbxSize(0).Height) / 2 + 3 * pbxSize(0).Height
  
  '设置DICOM显示图像的大小
  dcmView.Left = DICOM_VIEWER_BODER_SIZE
  dcmView.Top = DICOM_VIEWER_BODER_SIZE
  dcmView.Width = picView.Width - DICOM_VIEWER_BODER_SIZE
  dcmView.Height = picView.Height - DICOM_VIEWER_BODER_SIZE
End Sub


Private Sub ConfigTwainDisplay()
  '边框大小
  Const DICOM_VIEWER_BODER_SIZE As Long = 5
  
  dcmView.Left = DICOM_VIEWER_BODER_SIZE
  dcmView.Top = DICOM_VIEWER_BODER_SIZE
  dcmView.Width = picView.Width - DICOM_VIEWER_BODER_SIZE
  dcmView.Height = picView.Height - DICOM_VIEWER_BODER_SIZE
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

BUGEX "OpenVideoCaptureDevice 1"

    If mVideoCapture Is Nothing Then
        '创建视频采集对象
        Set mVideoCapture = New clsVideoCapture
        
        '连接视频相关组件
        Call mVideoCapture.ConnectedVfwDeviceObj(picVideo)
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
    
    If Not mVideoCapture.IsStartup Then
        
        '设置视频驱动类型
        mVideoCapture.VideoDriverType = gobjCapturePar.VideoDirverType
    
        '读取视频大小
        mVideoSize.Width = ScaleX(mVideoCapture.VideoSize.Width, vbPixels, vbTwips)
        mVideoSize.Height = ScaleX(mVideoCapture.VideoSize.Height, vbPixels, vbTwips)
        
        '配置界面
        Call CaptureSwitchFace(IsTwainCaptureWay Or IsCustomCaptureWay)
        

        '*******************************************************
BUGEX "OpenVideoCaptureDevice 5"
        '开始视频预览********************************************
        If Not IsTwainCaptureWay And Not IsCustomCaptureWay Then
            BUGEX "LSQ test1"
            mblnRealTime = True
            
            Call mVideoCapture.StartPreview
            blnIsStartupVideo = mVideoCapture.IsStartup
        Else
            BUGEX "LSQ test2"
            mblnRealTime = False
            
            blnIsStartupVideo = ImageScanner.ScannerAvailable
        End If
 

        '*********************************************************
BUGEX "OpenVideoCaptureDevice 8"
    '    If mVideoCapture.IsStartup Then Call ucCapHook.EnableHook
    Else
        Call ConfigVideoShowState(True)
    End If
    
    Call OpenComm   '打开采集端口
    
'    If gobjCapturePar.strCaptureHot <> "" Then Call mobjHotHook.EnableHook(WM_KEYDOWN, True)
End Sub


Public Sub UpdateAfterCaptureInfo()
'更新后台采集信息
    
    '只有影像采集模块并且启用后后台采集才能使用后台采集
    If Not IsTwainCaptureWay And Not IsCustomCaptureWay Then
        Call CreateNewCaptureTag
        Call ShowAfterCaptureInf(False)
    Else
        Call ShowAfterCaptureInf(False)
    End If
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
'    Call SetWindowStyle
'    Set mobjHotHook = New clsHookKey
    
    DTPImg.value = Format(zlDatabase.Currentdate, "yyyy-mm-dd")
    mdate = DTPImg.value
    Call LoadMiniCache
    
    '在这里必须对该窗口对象进行置顶操作，否则在执行打开或者保存操作时，弹出的文件选择框将位于该窗口之后
    SetWindowPos Me.hwnd, -1, Me.CurrentX, Me.CurrentY, Me.ScaleWidth, Me.ScaleHeight, 3  '将窗口置顶
    
    mstrAfterStationName = AnalyseComputer
    
    InitCommandBars
    
    Call InitReportImageFace
            
    ucPreview.PageImgCount = Val(GetSetting("ZLSOFT", G_STR_REG_PATH_PUBLIC, "视频缩略图数量", 6))
    ucPreview.ShowCheckBox = True
    
    ucCacheImage.PageImgCount = Val(GetSetting("ZLSOFT", G_STR_REG_PATH_PUBLIC, "视频缩略图数量", 6))
    ucCacheImage.ShowCheckBox = True
    
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
    
    Set mfrmParameter = New frmVideoSetup
    
    Exit Sub
errHandle:
    MsgboxCus err.Description, vbOKOnly, G_STR_HINT_TITLE
End Sub

Private Sub InitReportImageFace()
    Dim objImage As Pane, objCache As Pane
    
    With dkpReportImage
        .CloseAll
        .VisualTheme = ThemeOffice2003
        .Options.HideClient = True
        .Options.UseSplitterTracker = False '实时拖动
        .Options.ThemedFloatingFrames = True
        .Options.AlphaDockingContext = True
        .PanelPaintManager.BoldSelected = True
        .TabPaintManager.position = xtpTabPositionLeft  'TAB放到左边显示
        .TabPaintManager.Appearance = xtpTabAppearancePropertyPage2003
        .TabPaintManager.BoldSelected = True
    End With
    
    Set objImage = dkpReportImage.CreatePane(1, 0, 400, DockLeftOf)
    objImage.Title = "检 查 图"
    objImage.Handle = picPreview.hwnd
    objImage.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    
    If gobjCapturePar.IsUseAfterCapture = True Then
        Call UseAfterImageChanged(True)
    Else
        picCacheImage.Visible = False
    End If
    
    objImage.Selected = True
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
    picVideo.Visible = False
      
    If blnUseTwain Then
      Set dcmView.Container = picCapture
      Set txtInputText.Container = picCapture
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
    Call mVideoCapture.StopPreview
    
    gobjCapturePar.VideoDirverType = videoDirver
    mVideoCapture.VideoDriverType = videoDirver
       
    '读取视频大小
    mVideoSize.Width = ScaleX(mVideoCapture.VideoSize.Width, vbPixels, vbTwips)
    mVideoSize.Height = ScaleX(mVideoCapture.VideoSize.Height, vbPixels, vbTwips)
       
    Call CaptureSwitchFace(videoDirver = vdtTWAIN Or videoDirver = vdtCustom)
        
    
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
            Call InitCustomDevice
        End If
        
        mblnRealTime = False
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
  
  
  '显示处理工具栏
  SaveSetting "ZLSOFT", G_STR_REG_PATH_PUBLIC, "显示处理工具栏", mblnShowProcessBar
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
        tmrComm.Interval = 2
    End If
BUGEX "OpenComm 10"
    Exit Sub
err:
BUGEX "OpenComm 11"
    Call MsgboxCus("端口打开错误", vbOKOnly, G_STR_HINT_TITLE)
BUGEX "OpenComm 12"
End Sub


Private Sub dcmView_MouseDown(Button As Integer, Shift As Integer, X As Long, Y As Long)
    If Button = 1 And dcmView.Images.Count > 0 Then
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
                    
                    Set mdcmSelectLabel = dcmView.Images(1).Labels(dcmView.Images(1).Labels.Count)
                    
                    mdcmSelectLabel.LineWidth = 2
            End Select
        End If
    End If
End Sub


Private Sub dcmView_MouseMove(Button As Integer, Shift As Integer, X As Long, Y As Long)
    Dim dblZoom As Double
    
    If mblnDcmViewDown = True And Button = 1 And dcmView.Images.Count > 0 Then
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
    If mblnDcmViewDown = True And Button = 1 And dcmView.Images.Count > 0 Then
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
            If dcmView.Images(1).Labels.Count > 0 Then
                dcmView.Images(1).Labels.Remove dcmView.Images(1).Labels.Count
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

   
Public Sub subCaptureImg(ByVal RealTimeCap As Boolean, _
                        Optional ByVal strFileName As String = "", _
                        Optional ByRef picCapture As StdPicture = Nothing, _
                        Optional ByVal blnIsAfterCapture As Boolean = False, _
                        Optional ByVal blnUseCustom As Boolean = False)
'------------------------------------------------
'功能：采集并存储图像
'参数：无
'返回：无，直接保存新采集的图像
'------------------------------------------------
On Error GoTo errHandle
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    
    If mstrVideoRegTime = "" Then
        MsgboxCus mstrMsg, vbOKOnly, "提示"
        Exit Sub
    End If
    
    If blnIsAfterCapture Then
        If Not mVideoCapture.IsStartup Then Exit Sub
    Else
        If Not (Not mblnReadOnly And (mVideoCapture.IsStartup Or IsTwainCaptureWay)) Then Exit Sub
    End If
    
BUGEX "subCaptureImg 1"
    If funCaptureSingleImage(RealTimeCap, strFileName, picCapture, blnIsAfterCapture) = True Then
        If blnIsAfterCapture Then
            '如果是后台采集，则后台采集成功后，删除后台采集的图像
            If subSaveAfterCaptureImage Then Call dcmAfter.Images.Clear
            
            mAfterCaptureInf.lngAfterCurImageCount = GetCountOfCapPic(mintTagNow)
            
            Call ShowAfterCaptureInf(True)
                        
            Exit Sub
        End If
        
        If IsCustomCaptureWay And blnUseCustom Then Exit Sub
        
BUGEX "subCaptureImg 2"
        mintCaptureFlag = 2
        
        Call subSaveImage(mlngAdviceId, mcurStudyInf.strStudyUid)
    End If
BUGEX "subCaptureImg 5"
Exit Sub
errHandle:
    err.Raise err.Number, err.Description
End Sub

Private Function CopyPictureToDicomImg(ByVal lngHDC As Long, ByVal lngPictureHandle As Long, objDcmImg As Object) As Boolean
'congpicture中复制图像到dicomimage
    Const bitCount As Long = 3
        
    Dim hBitmap As OLE_HANDLE
    Dim stucbmp As TBitMap
    Dim lngSize As Long
    Dim lngResult As Long
    Dim aryPixels() As Byte
    Dim stuDipInf As BITMAPINFO
    
    Dim i As Long, j As Long, bytTemp As Byte
    
    
    CopyPictureToDicomImg = False
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

    CopyPictureToDicomImg = True
End Function


Private Function funCaptureSingleImage(ByVal RealTimeCap As Boolean, _
                                    Optional ByVal strFileName As String = "", _
                                    Optional ByRef picCapture As StdPicture = Nothing, _
                                    Optional ByVal blnIsAfterCapture As Boolean = False) As Boolean
'------------------------------------------------
'功能：采集单帧视频图像，将图像转换成DICOM格式，并填写DICOM文件头，最后将图像放入缩略图dcmMiniature中。
'参数：无
'返回：无，直接将新采集的图像放入dcmMiniature中
'------------------------------------------------
'采集单帧图像
On Error GoTo SaveFileError
    Dim ImgTmpImage As DicomImage
    Dim dcmTag As clsImageTagInf
    
    '采集图像，分为直接视频采集和播放录象采集

    If Not (picCapture Is Nothing) Then
        Set picTemp2.Picture = Nothing
        picTemp2.Width = mVideoSize.Width * (1 - mCurCutRange.WidthRate - mCurCutRange.LeftRate)
        picTemp2.Height = mVideoSize.Height * (1 - mCurCutRange.HeightRate - mCurCutRange.TopRate)
        picTemp2.Picture = picCapture
    ElseIf Trim(strFileName) <> "" And Dir(strFileName) <> "" Then
        '使用dcmView显示的是图片，不需要再裁剪
        Set picTemp2.Picture = Nothing
        Set picTemp2.Picture = LoadPicture(strFileName)
        
    Else
        If RealTimeCap = False And mblnPlayVideo = False Then
            '使用dcmView显示的是图片，不需要再裁剪
            Set picTemp2.Picture = Nothing
            
            If dcmView.Images.Count > 0 Then
                Set picTemp2.Picture = dcmView.CurrentImage.Capture(False).Picture
            End If
        Else
            '当处于实时视频显示时，需要对图像进行裁剪操作
            Set picTemp2.Picture = Nothing
                        
            Dim curPic As StdPicture
            Set curPic = mVideoCapture.CaptureImageFromMemory

            If curPic Is Nothing Then
                Call MsgboxCus("视频图像采集失败，请检查视频参数设置是否正确(如视频设备，显示模式等)。", vbOKOnly, G_STR_HINT_TITLE)
                
                funCaptureSingleImage = False
                Exit Function
            End If
            
            picTemp2.Width = mVideoSize.Width * (1 - mCurCutRange.WidthRate - mCurCutRange.LeftRate)
            picTemp2.Height = mVideoSize.Height * (1 - mCurCutRange.HeightRate - mCurCutRange.TopRate)

            '应用图像范围裁剪
            Call picTemp2.PaintPicture(curPic, 0, 0, picTemp2.Width, picTemp2.Height, _
                                       mVideoSize.Width * mCurCutRange.LeftRate, mVideoSize.Height * mCurCutRange.TopRate, _
                                       picTemp2.Width, picTemp2.Height, vbSrcCopy)
                                               
            picTemp2.Picture = picTemp2.Image

            Set curPic = Nothing
        End If
    End If
    
    If picTemp2.Picture Is Nothing Then
        funCaptureSingleImage = False
        Exit Function
    End If

    '创建dicom格式图像
    Set ImgTmpImage = New DicomImage
    
    If mblnUseClipbord Then
        '使用剪贴板方式
        Call Clipboard.SetData(picTemp2.Picture, 2)
        '从剪切板取回图像
        Call ImgTmpImage.Paste
        
        Call Clipboard.Clear
    Else
        '不使用剪贴板方式，从Picture中复制图像到ImgTmpImage中,不使用剪贴板交换数据
        If Not CopyPictureToDicomImg(picTemp2.hdc, picTemp2.Image.Handle, ImgTmpImage) Then
            funCaptureSingleImage = False
            Exit Function
        End If
    End If
    

    '填写图像参数到DICOM参数
    Call subWriteDicomPara(ImgTmpImage, mlngAdviceId, blnIsAfterCapture)
    
    Set dcmTag = New clsImageTagInf
    dcmTag.Tag = imgTag
    
    Set ImgTmpImage.Tag = dcmTag
    
    If blnIsAfterCapture Then
        Call dcmAfter.Images.Add(ImgTmpImage)
    Else
        '将图像放入缩略图中
        Call subInsert2Mini(ImgTmpImage)
    End If
    
BUGEX "dcmTag:" & ImgTmpImage.Tag.Tag
    
    funCaptureSingleImage = True

    Exit Function
SaveFileError:
    MsgboxCus err.Description, vbOKOnly, G_STR_HINT_TITLE
End Function


Private Sub subWriteDicomPara(img As DicomImage, lngAdviceId As Long, _
    Optional blnIsAfterCapture As Boolean = False)
'------------------------------------------------
'功能：给输入的图像填写DICOM文件头信息
'参数：img－－输入的DICOM文件,lngAdviceID－－医嘱ID
'返回：无，直接文件头信息写入img的文件头
'------------------------------------------------
    Dim curDate As Date

    curDate = zlDatabase.Currentdate
    
    If blnIsAfterCapture Then
        img.Attributes.Add &H10, &H10, ""                           'Name 姓名
        img.Attributes.Add &H10, &H20, ""                           'Patient ID 病人ID
        img.Attributes.Add &H10, &H30, ""                           'BirthDate 生日
        img.Attributes.Add &H10, &H40, ""                           'Sex 性别
        img.Attributes.Add &H10, &H1010, ""                         'Age 年龄
        img.Attributes.Add &H10, &H4000, ""                         'Patient Comment 病人注释
        img.Attributes.Add &H20, &H10, ""                           'Study ID 检查ID
        img.Attributes.Add &H8, &H60, mcurStudyInf.strModality                   'Modality 影像类别
    Else
        img.Attributes.Add &H10, &H10, mcurStudyInf.strName                     'Name 姓名
        img.Attributes.Add &H10, &H20, mcurStudyInf.strPatientID                'Patient ID 病人ID
        img.Attributes.Add &H10, &H30, mcurStudyInf.strBirthDate                'BirthDate 生日
        img.Attributes.Add &H10, &H40, mcurStudyInf.strSex                      'Sex 性别
        img.Attributes.Add &H10, &H1010, mcurStudyInf.strAge                    'Age 年龄
        img.Attributes.Add &H10, &H4000, ""                         'Patient Comment 病人注释
        img.Attributes.Add &H20, &H10, mcurStudyInf.strCheckNo                  'Study ID 检查ID
        img.Attributes.Add &H8, &H60, mcurStudyInf.strModality                   'Modality 影像类别
    End If
    
    img.Attributes.Add &H8, &H8, ""                             'ImageType  空
    img.Attributes.Add &H8, &H16, "1.2.840.10008.5.1.4.1.1.7"   'SOP Class  UID，二次捕捉
    img.Attributes.Add &H8, &H20, Format(curDate, "yyyy-mm-dd")     'Study Date 检查日期
    img.Attributes.Add &H8, &H21, Format(curDate, "yyyy-mm-dd")     'Series Date 序列日期
    img.Attributes.Add &H8, &H22, Format(curDate, "yyyy-mm-dd")     'Acquisition Date 采集日期
    img.Attributes.Add &H8, &H23, Format(curDate, "yyyy-mm-dd")     'Image Date   图像日期
    img.Attributes.Add &H8, &H30, Format(curDate, "HH24:MI:SS")     'Study Time   检查时间
    img.Attributes.Add &H8, &H31, Format(curDate, "HH24:MI:SS")     'Series Time  序列时间
    img.Attributes.Add &H8, &H32, Format(curDate, "HH24:MI:SS")     'Acquisition Time  采集时间
    img.Attributes.Add &H8, &H33, Format(curDate, "HH24:MI:SS")     'Image Time  图像时间
    img.Attributes.Add &H8, &H50, ""                            'Accession Number 空
    img.Attributes.Add &H8, &H70, "ZLSOFT"                      'Manufacturer 厂商
    img.Attributes.Add &H8, &H80, mstrInstitution                'Institution Name 单位名称
    img.Attributes.Add &H8, &H90, ""                            'Referring Physician's Name 空
    img.Attributes.Add &H8, &H1030, ""                          'Study Description 检查描述 空
    img.Attributes.Add &H20, &H11, "1"                          'Series Number 序列号
    img.Attributes.Add &H20, &H13, "1"                          'ImageNumber 图像号
    img.Attributes.Add &H20, &H20, ""                           'Orientation 空
End Sub

Private Sub UniteUID(img As DicomImage)
    Set mdcmTmpImg = img
    
    '如果是视频,或者音频，则不修正序列UID
    '根据缩略图的检查UID和序列UID，修改img的值
    Call subUniteUID(mdcmTmpImg, mdcmTmpImg.Tag.Tag <> VIDEOTAG And mdcmTmpImg.Tag.Tag <> AUDIOTAG)
End Sub

Private Sub subInsert2Mini(img As DicomImage)
'------------------------------------------------
'功能：将图像添加到缩略图dcmMiniature中
'参数：img－－输入的DICOM图像
'      blnIsLocalImg如果为true,则表示为视频
'返回：无，直接将图像添加到缩略图dcmMiniature中
'------------------------------------------------
    
    '如果是视频,或者音频，则不修正序列UID
    '根据缩略图的检查UID和序列UID，修改img的值
    Call subUniteUID(img, img.Tag.Tag <> VIDEOTAG And img.Tag.Tag <> AUDIOTAG)
    
    ucPreview.AddImage img, img.Tag
End Sub

Private Sub Form_Resize()
On Error GoTo errHandle
BUGEX "Form_Resize(frmWork_Video) 1"

    Call ucSplitter1.RePaint(False)
BUGEX "Form_Resize(frmWork_Video) 2"

BUGEX "Form_Resize(frmWork_Video) picCaptureHeight:" & picCapture.Height

Exit Sub
errHandle:
BUGEX "Form_Resize(frmWork_Video) Err:" & err.Description
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    mdate = Format(zlDatabase.Currentdate, "yyyy-mm-dd")
    'LoadMiniCache (True) '用于清除最新的空目录
    ClearEmptyFolder (True)

BUGEX "VideoForm_UnLoad 1"
    tmrComm.Enabled = False
    timerHook.Enabled = False
    
BUGEX "VideoForm_UnLoad 3"
    '先关闭采集窗口和COMM口
    Call StopCapture
BUGEX "VideoForm_UnLoad 4"
    '保持裁剪设置
    Call SaveParameterCfg
BUGEX "VideoForm_UnLoad 5"
    Call SaveSetting("ZLSOFT", G_STR_REG_PATH_PUBLIC, "视频缩略图数量", ucPreview.PageImgCount)
    
BUGEX "VideoForm_UnLoad 6"
    If Not mfrmParameter Is Nothing Then
        Unload mfrmParameter
    End If
BUGEX "VideoForm_UnLoad 7"
    '断开ftp连接
    If Not mobjFtpConnection Is Nothing Then
        Call mobjFtpConnection.FuncFtpDisConnect
        Set mobjFtpConnection = Nothing
    End If
BUGEX "VideoForm_UnLoad 8"
    '断开备份ftp连接
    If Not mobjBakFtpConnection Is Nothing Then
        Call mobjBakFtpConnection.FuncFtpDisConnect
        Set mobjBakFtpConnection = Nothing
    End If
    
BUGEX "VideoForm_UnLoad 9"
    If Not mfrmOpenStudy Is Nothing Then
        Unload mfrmOpenStudy
        Set mfrmOpenStudy = Nothing
    End If
    
BUGEX "VideoForm_UnLoad 10"
    wdmCapture.FreeRes
BUGEX "VideoForm_UnLoad 11"

'    Call mobjHotHook.FreeHook
'    Set mobjHotHook = Nothing
    
    Set dcmglbUID = Nothing
    Set mobjDxDevice = Nothing
    Set mVideoCapture = Nothing
    Set mfrmParameter = Nothing
    
    If Not mobjCustomDevice Is Nothing Then
        mobjCustomDevice.zlFree
        Set mobjCustomDevice = Nothing
    End If
BUGEX "VideoForm_UnLoad End"
End Sub


Private Sub subDeleteImage()
'------------------------------------------------
'功能：删除缩略图中被选中的图像，先从数据库中删除，然后从FTP中删除。删除后触发StateChanged事件
'参数：无
'返回：无，直接删除缩略图中最后一个图像
'------------------------------------------------
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim blnResult As Boolean
    Dim i As Long
    Dim lngCount As Integer '左上角有红标记的图像数
    Dim strImgUids As String
    Dim lngIndexTemp As Long
    Dim lngDelCount As Long
    
    If ucPreview.ImgViewer.Images.Count > 0 Then
        lngCount = 0
        '得到需要删除的图像uid中间用';'隔开
        For i = 1 To ucPreview.CurImageCount
            If ucPreview.ImgChecked(i) = True Then
                lngIndexTemp = i
                lngCount = lngCount + 1
                strImgUids = IIf(lngCount = 1, ucPreview.ImgViewer.Images(i).InstanceUID, strImgUids & ";" & ucPreview.ImgViewer.Images(i).InstanceUID)
            End If
        Next
                
        If lngCount > 100 Then
            MsgboxCus "对不起，程序支持一次性删除的图像不能超过100张，请重新选择", vbInformation, G_STR_HINT_TITLE
            Exit Sub
        End If
      
        '从数据库和FTP中删除缩略图中被选中的图像
        If lngCount > 1 Then
            blnResult = DeleteManyImages(strImgUids)
        ElseIf lngCount = 1 Then
        '如果等于1 可能红点和红框不是同一幅图像，此时删除红点图像。
            blnResult = DeleteImages(Me, 1, strImgUids, "")
        Else
            blnResult = DeleteImages(Me, 1, ucPreview.SelectImage.InstanceUID, "")
        End If
        
        If blnResult = True Then    '删除成功，则修改缩略图状态，并触发StateChanged事件
            '在缩略图中删除图像
            lngDelCount = lngCount
            If lngCount > 1 Then
                For i = ucPreview.CurImageCount To 1 Step -1
                    If ucPreview.ImgChecked(i) = True Then
                        If lngDelCount <> 1 Then
                            Call ucPreview.DeleteImage(i, False)
                            lngDelCount = lngDelCount - 1
                        Else
                            Call ucPreview.DeleteImage(i, True, True)
                        End If
                    End If
                Next
            ElseIf lngCount = 1 Then
                Call ucPreview.DeleteImage(lngIndexTemp)
            ElseIf lngCount = 0 Then
                Call ucPreview.DeleteImage(ucPreview.SelectIndex)
            End If
            dcmView.Images.Clear

            If Not ucPreview.SelectImage Is Nothing Then
                dcmView.Images.Add ucPreview.SelectImage
            End If
            
            
            '设置影像检查状态，如果删除最后一个图，且原检查过程为3，则修改为2
            If ucPreview.CurImageCount = 0 Then
                
                If mlngStudyState = 3 Then
                    strSQL = "Zl_影像检查_State(" & mlngAdviceId & "," & mlngSendNo & ",2,NULL,'" & UserInfo.编号 & "','" & UserInfo.姓名 & "'," & glngDepartId & ")"
                    zlDatabase.ExecuteProcedure strSQL, "删除最后一个图像"
                    mlngStudyState = 2
                End If
                
                Call DoStateChange(vetDelAllImg, mlngAdviceId, mlngSendNo, mcurStudyInf.strStudyUid)
                
                mcurStudyInf.strStudyUid = ""
                
                '当最后的图像删除时，则显示实时视频画面
                Call ConfigVideoShowState(True)
            Else
                Call DoStateChange(vetImgDeled, mlngAdviceId, mlngSendNo, mcurStudyInf.strStudyUid)
            End If
        
        End If
    End If
End Sub


Private Sub subSetMouseState(intMouseState As Integer)
    '改变当前鼠标状态
    mintMouseState = IIf(mintMouseState = intMouseState, 0, intMouseState)
    
    cbrMain.FindControl(xtpControlButton, conMenu_Process_Window).Checked = False
    cbrMain.FindControl(xtpControlButton, conMenu_Process_Zoom).Checked = False
    cbrMain.FindControl(xtpControlButton, conMenu_Process_RectZoom).Checked = False
    cbrMain.FindControl(xtpControlButton, conMenu_Process_Arrow).Checked = False
    cbrMain.FindControl(xtpControlButton, conMenu_Process_Ellipse).Checked = False
    cbrMain.FindControl(xtpControlButton, conMenu_Process_Text).Checked = False
    cbrMain.FindControl(xtpControlButton, conMenu_Process_Corp).Checked = False
End Sub


'modify by tjh at 2010-01-20
'配置视频显示状态
Private Sub ConfigVideoShowState(ByVal blnShowState As Boolean)
  mblnRealTime = blnShowState
  
  Select Case gobjCapturePar.VideoDirverType
    Case vdtVFW
      picVideo.Visible = blnShowState
      wdmCapture.Visible = False
      dcmView.Visible = Not blnShowState
    Case vdtWDM
      wdmCapture.Visible = blnShowState
      picVideo.Visible = False
      dcmView.Visible = Not blnShowState
    Case vdtTWAIN, vdtCustom
      wdmCapture.Visible = False
      picVideo.Visible = False
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
  Call subCaptureImg(True, curScanFile)
  
  Call ShowScanImage(ucPreview.CurImageCount)
End Sub


Private Sub ShowScanImage(imgIndex As Integer)

    '将被选中图像装载到dcmView中
    dcmView.Images.Clear
    dcmView.Images.Add ucPreview.SelectImage
    
    '显示dcmView，隐藏picVideo
    dcmView.CurrentImage.BorderWidth = 0
    mblnRealTime = False
'    picVideo.Visible = False
'    dcmView.Visible = True
End Sub

Public Sub UseAfterImageChanged(ByVal blUse As Boolean)
    Dim objImage As Pane, objCache As Pane, objTmp As Pane
    Dim blHavePane As Boolean
    Dim i As Integer
    
    If blUse = True Then
    '是否需要创建的判断
        blHavePane = False
        For i = 1 To dkpReportImage.PanesCount
            Set objTmp = dkpReportImage.FindPane(i)
            
            If objTmp.Title = "检 查 图" Then
                Set objImage = dkpReportImage.FindPane(i)
            End If
            
            If objTmp.Title = "后 台 图" Then
                blHavePane = False
                Exit Sub
            End If
        Next
        If blHavePane = False Then
            picCacheImage.Visible = True
            Set objCache = dkpReportImage.CreatePane(2, 0, 400, DockLeftOf)
            objCache.Title = "后 台 图"
            objCache.Handle = picCacheImage.hwnd
            objCache.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
            objCache.AttachTo objImage
            objImage.Selected = True
            
            LoadMiniCache
        End If
    Else
    '是否需要销毁的判断
        blHavePane = False
        
        For i = 1 To dkpReportImage.PanesCount
            Set objTmp = dkpReportImage.FindPane(i)
            
            If objTmp.Title = "后 台 图" Then
                blHavePane = True
                Exit For
            End If
        Next

        If blHavePane = True Then Call dkpReportImage.DestroyPane(objTmp)

        picCacheImage.Visible = False

    End If
  
    Exit Sub
errH:
    Call err.Raise(0, , "后台图标签处理错误" & err.Description)
End Sub

Private Sub mfrmParameter_OnIsUseAfterImgChange(ByVal blUse As Boolean)
On Error GoTo errHandle
'后台采集参数设置界面确认后触发，判断是否显示后台图pane或是否删除后台图pane
    gobjCapturePar.IsUseAfterCapture = blUse
    
    Call UseAfterImageChanged(blUse)
    RaiseEvent OnIsUseAfterImgChange(blUse)
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
            If Me.cbrMain.FindControl(xtpControlButton, conMenu_Cap_MarkMap).Enabled And _
                Me.cbrMain.FindControl(xtpControlButton, conMenu_Cap_MarkMap).Visible Then
                Call subCaptureImg(True)
            End If
        Case 1  '后台采集
BUGEX "mobjDxDevice_OnDxKeyPress 3"
            If Me.cbrMain.FindControl(xtpControlButton, conMenu_Cap_After_Capture).Enabled And _
                Me.cbrMain.FindControl(xtpControlButton, conMenu_Cap_After_Capture).Visible Then
                Call subCaptureImg(True, "", Nothing, True)
            Else
                Call mobjDxDevice_OnDxKeyPress(0)
            End If
        Case 2  '更新标识
BUGEX "mobjDxDevice_OnDxKeyPress 4"
            If Me.cbrMain.FindControl(xtpControlButton, conMenu_Cap_After_Tag).Enabled And _
                Me.cbrMain.FindControl(xtpControlButton, conMenu_Cap_After_Tag).Visible Then
                
                If gobjCapturePar.IsUseAfterCapture Then Call UpdateAfterCaptureInfo
            Else
                Call mobjDxDevice_OnDxKeyPress(0)
            End If
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


Private Sub mobjHotHook_OnKeyBoardLHook(ByVal lngMsg As Long, ByVal lngVkCode As Long, ByVal lngScanCode As Long, ByVal lngFlags As Long)
'    Dim lngWindowPID As Long
'    Dim lngParentPID As Long
'    Dim lngCurrentPID As Long
'
'    If lngMsg <> WM_KEYDOWN Then Exit Sub
'
'    '判断触发消息的是否为当前进程
'    Call GetWindowThreadProcessId(GetActiveWindow(), lngWindowPID)
'    Call GetWindowThreadProcessId(glngRootHandle, lngParentPID)
'
'    lngCurrentPID = GetCurrentProcessId
'
'
'    If lngCurrentPID = lngWindowPID Or lngWindowPID = lngParentPID Then
'
'
'
'        '使用热键进行采集
'        If GetKeyAliasEx(lngVkCode) = gobjCapturePar.strCaptureHot Then
'            If Me.cbrMain.FindControl(xtpControlButton, conMenu_Cap_MarkMap).Enabled And _
'                Me.cbrMain.FindControl(xtpControlButton, conMenu_Cap_MarkMap).Visible Then
'                Call subCaptureImg(True)
'            End If
'        End If
'    End If
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
    ElseIf picVideo.Visible Then
      Call ChangeCutRanage(picVideo, Index, X, Y)
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
    ElseIf picVideo.Visible Then
      Call ApplayCutRange(picVideo)
    End If
    
    If IsTwainCaptureWay Or IsCustomCaptureWay Then
      ConfigTwainDisplay
    Else
      '设置显示范围
      Call ConfigVideoDisplay(wdmCapture)
      Call ConfigVideoDisplay(picVideo)

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


Private Sub picCapture_Resize()
On Error GoTo errHandle
    
    '设置图标大小
    If picCapture.Height < 7000 Or picCapture.Width < 4000 Then
        cbrMain.Options.SetIconSize True, 16, 16
    Else
        cbrMain.Options.SetIconSize True, 32, 32
    End If
    
    picCapture.Refresh
    
errHandle:
End Sub


Private Function LoadPlayVideo() As String
'下载播放视频
On Error GoTo errHandle
    If dcmView.Images(1).Tag.Tag <> VIDEOTAG And dcmView.Images(1).Tag.Tag <> AUDIOTAG Then Exit Function
    
    If dcmView.Images(1).Tag.Tag = VIDEOTAG Then
        Call dcmView.Images(1).FileImport(GetAppPath & "\..\附加文件\aviDownload.bmp", "DIB/BMP")
    
        '下载需要播放的视频
        LoadPlayVideo = GetSingleImage(dcmView.Images(1).InstanceUID, dcmView.Images(1).SeriesUID, Me, mblnMoved)
    
        Call dcmView.Images(1).FileImport(GetAppPath & "\..\附加文件\avi.bmp", "DIB/BMP")
    Else
        Call dcmView.Images(1).FileImport(GetAppPath & "\..\附加文件\wavDownload.bmp", "DIB/BMP")
    
        '下载需要播放的视频
        LoadPlayVideo = GetSingleImage(dcmView.Images(1).InstanceUID, dcmView.Images(1).SeriesUID, Me, mblnMoved)
    
        Call dcmView.Images(1).FileImport(GetAppPath & "\..\附加文件\wav.bmp", "DIB/BMP")
    End If
errHandle:
End Function

Private Sub subVideoPlay()
'------------------------------------------------
'功能：dcmView中录像图像的播放
'参数：无
'返回：无，直接播放dcmView中的图像
'------------------------------------------------
    Dim strFile As String
    
    If dcmView.Images.Count > 0 Then
        '下载录像，如果本地存在，则不进行下载
        If dcmView.Images(1).Tag.Tag <> VIDEOTAG And dcmView.Images(1).Tag.Tag <> AUDIOTAG Then Exit Sub
        
        strFile = LoadPlayVideo
        
        '打开播放··
        Call frmPlaying.Show
        
        '刷新播放窗口
'       Call frmPlaying.Refresh
        While Not frmPlaying.IsActive
            Call Sleep(10)
            DoEvents
        Wend
            
        Call frmPlaying.OpenVideoFile(Replace(strFile, "/", "\"), Me)
    End If
End Sub


Private Sub subVideoSaveAs()
'------------------------------------------------
'功能：另存dcmView中的图像,支持的格式为AVI,DCM,BMP,JPE
'参数：无
'返回：无
'------------------------------------------------
    Dim strFileName As String
    Dim strFileType As String
    
    If mblnRealTime = False And dcmView.Images.Count > 0 Then
    
        If dcmView.Images(1).Tag.Tag = VIDEOTAG Then
            dlgOpen.Filter = "(*.AVI)|*.AVI|(*.MPEG)|*.MPEG|(*.*)|*.*"
            
            dlgOpen.ShowSave
            strFileName = dlgOpen.FileName
            
            If strFileName <> "" Then
                '复制视频文件到指定路径
                Call FileCopy(dcmView.Images(1).Tag.VideoFile, strFileName)
            End If
            
            Exit Sub
        End If
            
        If dcmView.Images(1).FrameCount > 1 Then
            dlgOpen.Filter = "录像文件(*.AVI)|*.AVI|DICOM文件(*.dcm)|*.dcm|图像文件 (*.BMP)|*.BMP|图像文件(*.JPG)|*.JPG"
        Else
            dlgOpen.Filter = "DICOM文件(*.dcm)|*.dcm|图像文件 (*.BMP)|*.BMP|图像文件(*.JPG)|*.JPG"
        End If
        
        
        dlgOpen.ShowSave
        strFileName = dlgOpen.FileName
        
        If strFileName <> "" Then
            strFileType = UCase(Right(Trim(strFileName), 3))
            
            Select Case strFileType
                Case "AVI"
                    If dcmView.Images(1).FrameCount > 1 Then
                        dcmView.Images(1).WriteAVI strFileName, 1, dcmView.Images(1).FrameCount, 1, "", 100, False
                    Else
                        MsgboxCus "静态图像无法保存成AVI格式，请重新选择图像格式。", vbInformation, G_STR_HINT_TITLE
                    End If
                Case "DCM"
                    dcmView.Images(1).WriteFile strFileName, True
                Case "BMP"
                    dcmView.Images(1).FileExport strFileName, "BMP"
                Case "JPG"
                    dcmView.Images(1).FileExport strFileName, "JPG"
            End Select
        End If
    End If
End Sub


Private Sub InputImageFile()
'------------------------------------------------
'功能：打开外部文件，放入缩略图中
'参数：无
'返回：无
'------------------------------------------------
On Error Resume Next
    Dim DlgInfo As DlgFileInfo
    Dim i As Integer
    Dim ImgTmpImage As New DicomImage
    Dim j As Integer
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    
    '选择文件
    With Me.dlgOpen
        .CancelError = False
        .MaxFileSize = 32767 '被打开的文件名尺寸设置为最大，即32K
        .flags = cdlOFNHideReadOnly Or cdlOFNAllowMultiselect Or cdlOFNExplorer Or cdlOFNNoDereferenceLinks
        .DialogTitle = "选择文件"
        .Filter = "DICOM文件（*.dcm）(*.img)|*.dcm;*.img|图像文件 (*.BMP)(*.JPG)|*.BMP;*.JPG|所有文件（*.*）|*.*"
        .ShowOpen
        If .FileName <> "" Then
            DlgInfo = GetDlgSelectFileInfo(.FileName)
        End If
        '在打开了*.pif文件后须将Filename属性置空，否则当选取多个*.pif文件后，当前路径会改变
        .FileName = ""
    End With

    For i = 1 To DlgInfo.iCount
        Set ImgTmpImage = ReadViewImage(DlgInfo.sPath & DlgInfo.sFIle(i))
    
        
        '设置图像的DICOM参数
        subWriteDicomPara ImgTmpImage, mlngAdviceId
        
        Dim dcmTag As New clsImageTagInf
        dcmTag.Tag = imgTag
    
        Set ImgTmpImage.Tag = dcmTag
        
        mintCaptureFlag = 3
        
        '将图像插入到缩略图中
        subInsert2Mini ImgTmpImage
            
        '保存图像，并触发图像存储事件
        Call subSaveImage(mlngAdviceId, mcurStudyInf.strStudyUid)
    Next
End Sub

Private Sub subUniteUID(dcmImg As DicomImage, ByVal blnIsUpdateSeriesUid As Boolean)
'------------------------------------------------
'功能：重整输入图像的检查UID和序列UID，保证输入图像的检查UID和序列UID跟缩略图dcmMiniature中的一致。
'       新添加进来的图像采用第一个图像的检查UID和序列UID。
'       如果是第一个图像，则是用输入的检查UID或者是图像本身自带的检查UID，同时给检查UID变量赋值
'参数：dcmImg－－输入的DICOM图像
'返回：无，直接修改图像的检查UID和序列UID
'------------------------------------------------
    Dim i As Integer
    
    '新图像采用跟第一个图像相同的检查UID和序列UID
    If ucPreview.CurImageCount > 0 Then
                
        dcmImg.StudyUID = ucPreview.ImgViewer.Images(1).StudyUID
        
        '如果参数为true，则允许更新img的序列UID，否则使用新的序列
        If blnIsUpdateSeriesUid Then
            '查找为图像的序列UID
            For i = 1 To ucPreview.ImgViewer.Images.Count
                If ucPreview.ImgViewer.Images(i).Tag.Tag = imgTag Then
                    dcmImg.SeriesUID = ucPreview.ImgViewer.Images(i).SeriesUID
                    
                    Exit For
                End If
            Next i
            
        End If
    ElseIf Len(mcurStudyInf.strStudyUid) > 0 Then
        dcmImg.StudyUID = mcurStudyInf.strStudyUid
    Else
        mcurStudyInf.strStudyUid = dcmImg.StudyUID
        
        '当检查uid改变后，需要更新缩略图显示组件中的查询值
        ucPreview.QueryValue = mcurStudyInf.strStudyUid
    End If
End Sub


Private Function GetDlgSelectFileInfo(strFileName As String) As DlgFileInfo
'------------------------------------------------
'功能：将文件名转化为全路径数组
'参数：strFileName--文件名，通过打开文件控件来获得。
'返回：全路径数组
'------------------------------------------------
    Dim sPath, tmpStr As String
    Dim sFIle() As String
    Dim iCount, i As Integer
    On Error GoTo errHandle
    sPath = CurDir()  '获得当前的路径，因为在CommonDialog中改变路径时会改变当前的Path
    tmpStr = Right$(strFileName, Len(strFileName) - Len(sPath)) '将文件名分离出来
    
    If Left$(tmpStr, 1) = Chr$(0) Then
        '选择了多个文件(表现为第一个字符为空格)
        For i = 1 To Len(tmpStr)
            If Mid$(tmpStr, i, 1) = Chr$(0) Then
                iCount = iCount + 1
                ReDim Preserve sFIle(iCount)
            Else
                sFIle(iCount) = sFIle(iCount) & Mid$(tmpStr, i, 1)
            End If
        Next i
    Else
        '只选择了一个文件(注意：根目录下的文件名除去路径后没有"\"）
        iCount = 1
        ReDim Preserve sFIle(iCount)
        If Left$(tmpStr, 1) = "\" Then tmpStr = Right$(tmpStr, Len(tmpStr) - 1)
        sFIle(iCount) = tmpStr
    End If
    
    GetDlgSelectFileInfo.iCount = iCount
    
    ReDim GetDlgSelectFileInfo.sFIle(iCount)
    
    If Right$(sPath, 1) <> "\" Then sPath = sPath & "\"
    GetDlgSelectFileInfo.sPath = sPath
    
    For i = 1 To iCount
        GetDlgSelectFileInfo.sFIle(i) = sFIle(i)
    Next i
    Exit Function
errHandle:
    MsgboxCus "GetDlgSelectFileInfo函数执行错误！", vbOKOnly + vbCritical, G_STR_HINT_TITLE
End Function


Private Sub picDock_Paint()
BUGEX "picDock_Paint(frmWork_Video)"
End Sub

Private Sub picPreview_Resize()
On Error Resume Next
    ucPreview.Left = 0
    ucPreview.Top = 0
    ucPreview.Width = picPreview.ScaleWidth
    ucPreview.Height = picPreview.ScaleHeight
End Sub

Private Sub picCacheImage_Resize()
On Error Resume Next
    DTPImg.Left = 0
    DTPImg.Top = 0
    
    If mintFontSize = 12 Then
        DTPImg.Width = 1750
        DTPImg.Height = 360
    ElseIf mintFontSize = 15 Then
        DTPImg.Width = 2000
        DTPImg.Height = 420
    Else
        '其他情况按字号=9 处理
        DTPImg.Width = 1400
        DTPImg.Height = 300
    End If
    

    cboCacheImage.Left = DTPImg.Width
    cboCacheImage.Top = 0
    cboCacheImage.Width = picCacheImage.ScaleWidth - DTPImg.Width
    cboCacheImage.Height = picCacheImage.ScaleHeight
    
    ucCacheImage.Left = 0
    ucCacheImage.Top = DTPImg.Height
    ucCacheImage.Width = picCacheImage.ScaleWidth
    ucCacheImage.Height = picCacheImage.ScaleHeight - DTPImg.Height
End Sub

Private Sub TimerHook_Timer()
On Error GoTo errHandle
    '当使用hook热键调用采集时，使用timer进行采集操作，避免在执行多次CaptureImage操作后，hook失效
    '造成hook失效的可能原因有hook的处理机制中如果截获hook后的处理时间过长可能造成失效，或者dicomobjects的fileexport方法调用多次造成失效，目前不去细究
    Call CaptureImage
    
    timerHook.Enabled = False
Exit Sub
errHandle:
    MsgboxCus err.Description, vbOKOnly, G_STR_HINT_TITLE
End Sub


Private Sub TimerRePaint_Timer()
 
    TimerRePaint.Enabled = False

    Call cbrMain.RecalcLayout
    Call ucSplitter1.RedrawSelf
    Call ucPreview.RedrawSelf
    Call dcmView.Refresh
    Call picCapture.Refresh

    BUGEX "timerRePaint_Timer 1"
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
            dcmView.Images(1).Labels.Remove dcmView.Images(1).Labels.Count
            txtInputText = "1 "
        Else
            mdcmSelectLabel.Text = txtInputText.Text
            dcmView.Refresh
        End If
    End If
End Sub

Private Sub CustomVideoSave()
    Dim dcmTmpImg As New DicomImage
    Dim strVideoFiles As String
    Dim blnUseCustom As Boolean
    Dim strEncoderName As String '编码器名称
    Dim lngRecordTimeLen As Long '录制视频长度
    
    If mobjCustomDevice Is Nothing Then Exit Sub
    
    Call mobjCustomDevice.zlCaptureVideo(mlngAdviceId, strVideoFiles, blnUseCustom, strEncoderName, lngRecordTimeLen)
    
    '录像完成
    If Trim(strVideoFiles) <> "" And Dir(strVideoFiles) <> "" Then
        dcmTmpImg.FileImport GetAppPath & "\..\附加文件\avi.bmp", "DIB/BMP"
        
        Dim dcmTag As New clsImageTagInf
        dcmTag.EncoderName = strEncoderName
        dcmTag.VideoFile = strVideoFiles
        dcmTag.CaptureTime = zlDatabase.Currentdate
        dcmTag.RecordTimeLen = lngRecordTimeLen
        dcmTag.Tag = VIDEOTAG
        
        Set dcmTmpImg.Tag = dcmTag
        
        subWriteDicomPara dcmTmpImg, mlngAdviceId
        
        mintCaptureFlag = 4
        
        subInsert2Mini dcmTmpImg
        
        '保存视频录像
        Call subSaveImage(mlngAdviceId, mcurStudyInf.strStudyUid)
    End If
End Sub

Private Sub subVideoSave()
'------------------------------------------------
'功能：录像
'参数：无
'返回：将录像文件放入缩略图
'------------------------------------------------
    
    Dim i As Integer
    Dim dcmTmpImg As New DicomImage
    Dim strError As String
            
    On Error GoTo continue1
      '删除历史的视频文件
      If Dir(mstrAviFileName) <> "" Then
        Kill mstrAviFileName
      End If
continue1:
    
    On Error GoTo CapErr
            
    '按现目前的方式,使用vfw的时候不允许进行录像操作
    If mVideoCapture.VideoDriverType = vdtVFW Then
        '录像完成(vfw进入录象后，直到结束才执行StartVideo以后的语句)
        '不处理vfw的录像功能
        Exit Sub
    End If
    
    'modify by tjh at 2010-01-20
    strError = mVideoCapture.StartVideo(mstrAviFileName)
    If Trim(strError) <> "" Then MsgboxCus strError, vbInformation, G_STR_HINT_TITLE
    
    '获取当前录像的编码器名称
    mstrEncoderName = mVideoCapture.GetEncoderName
    
    Exit Sub
CapErr:
    Call MsgboxCus(err.Description, vbOKOnly, G_STR_HINT_TITLE)
    err.Clear
End Sub


'modify by tjh at 2010-01-20
'停止视频录像
Private Sub subStopVideo()
    Dim dcmTmpImg As New DicomImage
            
    If mVideoCapture.VideoDriverType = vdtVFW Then Exit Sub
    
    On Error GoTo continue1
    If Dir(mstrAviFileName) <> "" Then
        Kill mstrAviFileName
    End If
continue1:
    
    On Error GoTo CapErr
    
    Call mVideoCapture.StopVideo
   
    
    '录像完成
    If Dir(mstrAviFileName) <> "" Then
        On Error GoTo continue2
            dcmTmpImg.FileImport GetAppPath & "\..\附加文件\avi.bmp", "DIB/BMP"
continue2:
        
        Dim dcmTag As New clsImageTagInf
        dcmTag.EncoderName = mstrEncoderName
        dcmTag.VideoFile = mstrAviFileName
        dcmTag.CaptureTime = zlDatabase.Currentdate
        dcmTag.RecordTimeLen = mVideoCapture.GetTimeLen
        dcmTag.Tag = VIDEOTAG
        
        Set dcmTmpImg.Tag = dcmTag
        
        subWriteDicomPara dcmTmpImg, mlngAdviceId
        
        mintCaptureFlag = 4
        
        subInsert2Mini dcmTmpImg
        
        '保存视频录像
        Call subSaveImage(mlngAdviceId, mcurStudyInf.strStudyUid)
    End If
    
    Exit Sub
CapErr:
    MsgboxCus err.Description, vbOKOnly, G_STR_HINT_TITLE
End Sub


'停止音频文件
Public Sub subSaveAudio(ByVal strAudioFile As String, ByVal lngTimeLen As Long)

    Dim i As Integer
    Dim dcmTmpImg As New DicomImage
    
    On Error GoTo CapErr
   
    
    '录像完成
    If Dir(strAudioFile) <> "" Then
        dcmTmpImg.FileImport GetAppPath & "\..\附加文件\wav.bmp", "DIB/BMP"
        
        Dim dcmTag As New clsImageTagInf
        dcmTag.EncoderName = ""
        dcmTag.VideoFile = strAudioFile
        dcmTag.CaptureTime = zlDatabase.Currentdate
        dcmTag.RecordTimeLen = lngTimeLen
        dcmTag.Tag = AUDIOTAG
        
        Set dcmTmpImg.Tag = dcmTag
        
        subWriteDicomPara dcmTmpImg, mlngAdviceId
        
        mintCaptureFlag = 5
        
        subInsert2Mini dcmTmpImg
        
        '保存录制的音频
        Call subSaveImage(mlngAdviceId, mcurStudyInf.strStudyUid)
    End If
    
    Exit Sub
CapErr:
    MsgboxCus err.Description, vbOKOnly, G_STR_HINT_TITLE
End Sub

'modify by tjh at 2010-01-22
'全屏显示
Private Sub subFullCall()
  Call mVideoCapture.FullScreen(Me, Me.hwnd)
End Sub

Private Function CreatFolder(ByVal intTag As Integer) As Boolean
    Dim strAfterImgPath As String
    Dim strReceivedTime As String
    Dim curTime As Date
    Dim i As Integer
    
    curTime = zlDatabase.Currentdate
    CreatFolder = False

    '如果uid为空，则创建新的UID
    mAfterCaptureInf.strAfterStudyUid = dcmglbUID.NewUID
    mAfterCaptureInf.strAfterSeriesUid = dcmglbUID.NewUID
    
    mAfterCaptureInf.intAfterTag = intTag
    strReceivedTime = Format(curTime, "yyyyMMdd")
    
    strAfterImgPath = mstrAfterDir & "\" & strReceivedTime & "\" & "标识" & mAfterCaptureInf.intAfterTag & "-" & mAfterCaptureInf.strAfterStudyUid & "\" & mAfterCaptureInf.strAfterSeriesUid & "\"

    MkLocalDir strAfterImgPath


    
    cboCacheImage.AddItem "标识" & Lpad(mAfterCaptureInf.intAfterTag, 3, "0") & "     时间：" & Format(curTime, "HH:MM:SS"), 0

End Function

Private Function GetCaptureTag() As Integer
'获得当前标识
    Dim curDate As Date
    
    curDate = zlDatabase.Currentdate
    Call GetTodayTagMax(curDate)  '更新当天最大标识
           
    If GetCountOfCapPic(mintTagMax) = 0 Then
        '无图像，继续使用这个标志
        mintTagNow = mintTagMax
    Else
        '有图像，更新标识并且创建新目录
        mintTagNow = mintTagMax + 1
        CreatFolder (mintTagNow)
    End If
    
    '如果mintTagNow=1，判断目录标识1是否存在并考虑是否创建
    If mintTagNow = 1 Then
        If JudgeExistFolderOne = False Then
            Call CreatFolder(1)
        End If
    End If

    GetCaptureTag = mintTagNow
End Function

Private Sub CreateNewCaptureTag()
'取得新的采集标识

    mAfterCaptureInf.intAfterTag = GetCaptureTag

    mAfterCaptureInf.lngAfterCurImageCount = 0
End Sub

Public Sub showAfterCapInfo(ByVal intType As Integer, ByVal isNeedRefreshTitle As Boolean)

    If isNeedRefreshTitle = True Then
        mAfterCaptureInf.lngAfterCurImageCount = GetCountOfCapPic(mintTagNow)
        Call ShowAfterCaptureInf(True)
    End If
    
    If intType = 1 Then mblIsNeedRefresh = True
End Sub


Private Sub ShowAfterCaptureInf(ByVal blnShowTag As Boolean)
'更新后台采集图像信息
    Dim objPanel As Pane
    Dim i As Long
    
    If gobjOwner Is Nothing Then
        For i = 1 To dkpReportImage.PanesCount
            Set objPanel = dkpReportImage.Panes(i)
            If InStr(objPanel.Title, "后 台 图") > 0 Or InStr(objPanel.Title, "后台图") > 0 Then
                objPanel.Title = "后台图 - 当前标识:" & mintTagNow
                Exit Sub
            End If
        Next i
        
        Exit Sub
    Else
        For i = 1 To dkpReportImage.PanesCount
            Set objPanel = dkpReportImage.Panes(i)
            If InStr(objPanel.Title, "后 台 图") > 0 Or InStr(objPanel.Title, "后台图") > 0 Then
                objPanel.Title = "后 台 图"
                Exit For
            End If
        Next i
    End If
    
    If Not gobjCapturePar.IsUseAfterCapture Or blnShowTag = False Then
        If InStr(gobjOwner.Caption, "      当前后台采集标识：") > 0 Then
            gobjOwner.Caption = Mid(gobjOwner.Caption, 1, InStr(gobjOwner.Caption, "      当前后台采集标识：") - 1)
        End If
            
        Exit Sub
    End If
    
    If mAfterCaptureInf.strAfterParentTitle = "" Then
        If InStr(gobjOwner.Caption, "      当前后台采集标识：") > 0 Then
            mAfterCaptureInf.strAfterParentTitle = Mid(gobjOwner.Caption, 1, InStr(gobjOwner.Caption, "      当前后台采集标识：") - 1)
        Else
            mAfterCaptureInf.strAfterParentTitle = gobjOwner.Caption
        End If
    End If
    
    gobjOwner.Caption = mAfterCaptureInf.strAfterParentTitle & "      当前后台采集标识：" & mintTagNow & "  当前标识图像数：" & mAfterCaptureInf.lngAfterCurImageCount & "        "
End Sub


Private Function subSaveAfterCaptureImage(Optional iEncode As Integer = 0) As Boolean
'保存后台采集图像
    Dim i As Long
    Dim dtNowTime As Date
    Dim strReceivedTime As String
    Dim ImgTmp As DicomImage
    Dim strAfterImgPath As String
    Dim strTime As String
    Dim objFolder As Folder
    
    subSaveAfterCaptureImage = False
    
    If dcmAfter.Images.Count <= 0 Then Exit Function
    
    dtNowTime = zlDatabase.Currentdate
    strReceivedTime = Format(dtNowTime, "yyyyMMdd")
    
    If mAfterCaptureInf.strAfterStudyUid = "" Then
        '如果uid为空，则创建新的UID
        mAfterCaptureInf.strAfterStudyUid = dcmglbUID.NewUID
        mAfterCaptureInf.strAfterSeriesUid = dcmglbUID.NewUID
        
        mAfterCaptureInf.intAfterTag = GetCaptureTag
    End If


    strAfterImgPath = mstrAfterDir & "\" & strReceivedTime & "\" & "标识" & mAfterCaptureInf.intAfterTag & "-" & mAfterCaptureInf.strAfterStudyUid & "\" & mAfterCaptureInf.strAfterSeriesUid & "\"
    
    MkLocalDir strAfterImgPath
    
    Set objFolder = mobjFile.GetFolder(strAfterImgPath)
    
    strTime = Format(objFolder.DateCreated, "YYYY-MM-DD HH:MM:SS")
    
    mrsImageCache.AddNew
    mrsImageCache!姓名 = "标识" & Lpad(mintTagNow, 3, "0")
    mrsImageCache!检查uid = mAfterCaptureInf.strAfterStudyUid
    mrsImageCache!序列Uid = mAfterCaptureInf.strAfterSeriesUid
    mrsImageCache!检查日期 = strTime
    mrsImageCache!路径 = strAfterImgPath
    mrsImageCache.Update
    

    For i = 1 To dcmAfter.Images.Count
        Set ImgTmp = dcmAfter.Images(i)
        
        ImgTmp.StudyUID = mAfterCaptureInf.strAfterStudyUid
        ImgTmp.SeriesUID = mAfterCaptureInf.strAfterSeriesUid
        
        If ImgTmp.Tag.Tag <> VIDEOTAG And ImgTmp.Tag.Tag <> AUDIOTAG Then
            '保存图像到缓存目录
            Select Case iEncode
                Case 1          'Run-Length Encoding行程压缩
                    ImgTmp.WriteFile strAfterImgPath & ImgTmp.InstanceUID, True, "1.2.840.10008.1.2.5"
                Case 2          '不处理，保持原图的压缩方式
                    ImgTmp.WriteFile strAfterImgPath & ImgTmp.InstanceUID, True
                Case Else       'Lossless JPEG encoding JPEG无损压缩
                    ImgTmp.WriteFile strAfterImgPath & ImgTmp.InstanceUID, True, "1.2.840.10008.1.2.4.70"
            End Select
            
            '存储为报告图像
            'ImgTmp.FileExport strAfterImgPath & ImgTmp.InstanceUID & ".jpg", "JPG", 80
        End If
        
        If ImgTmp.Tag.Tag = VIDEOTAG Then
            '将录像复制到对应的目录中，避免重新下载
            Call MoveFile(ImgTmp.Tag.VideoFile, strAfterImgPath & ImgTmp.InstanceUID & ".avi")
        ElseIf ImgTmp.Tag.Tag = AUDIOTAG Then
            '将音频文件复制到对应的目录中，避免重新下载
            Call MoveFile(ImgTmp.Tag.VideoFile, strAfterImgPath & ImgTmp.InstanceUID & ".wav")
        End If
        
        mAfterCaptureInf.lngAfterCurImageCount = mAfterCaptureInf.lngAfterCurImageCount + 1
    Next i
    
    If ImgTmp.Tag.Tag <> VIDEOTAG And ImgTmp.Tag.Tag <> AUDIOTAG Then
        Call frmCaptureHint.ShowCaptureHint( _
            IIf(gobjCapturePar.IsWindowHint, strAfterImgPath & ImgTmp.InstanceUID, ""), _
            gobjCapturePar.IsSountHint, hpRB, Me)
            
    End If
    
    Call DoStateChange(vetAfterUpdateImg, 0, 0, mAfterCaptureInf.strAfterStudyUid)

    'Call AddDCMImg(strAfterImgPath)
    Dim dcmTag As New zl9PacsControl.clsImageTagInf
    Dim strPathTmp As String
    Set ImgTmp.Tag = dcmTag
    
    strPathTmp = strAfterImgPath & "\" & ImgTmp.InstanceUID
    strPathTmp = Replace(strPathTmp, "/", "\")
    strPathTmp = Replace(strPathTmp, "\\", "\")
    ImgTmp.Tag.FilePath = strPathTmp
    
    Call AddDCMImg(ImgTmp)
    
    subSaveAfterCaptureImage = True
End Function

Private Function ShowMessage(ByVal lngUpLoadResult As Long) As Boolean
'文件上传成功与否的提示,文件上传成功返回true，反之返回false
    ShowMessage = False
    
    If lngUpLoadResult = 0 Then '上传成功，不做处理
        ShowMessage = True
    ElseIf lngUpLoadResult = 1 Then 'FTP链接失败
        MsgboxCus "FTP连接失败，文件无法保存，请检查网络设置。", vbInformation, G_STR_HINT_TITLE
    Else                      '文件上传失败
        MsgboxCus "文件上传失败，可能由于网络不稳定造成。", vbInformation, G_STR_HINT_TITLE
    End If
End Function

Private Sub subSaveImage(ByVal lngAdviceId As Long, ByVal strStudyUid As String, Optional iEncode As Integer = 0)
'------------------------------------------------
'功能：将最后一个缩略图保存到数据库中
'参数：iEncode－－压缩方式，1－Run-Length Encoding行程压缩；2－不处理，保持原图的压缩方式，其他－Lossless JPEG encoding JPEG无损压缩
'返回：无
'------------------------------------------------
    Dim ImgTmp As New DicomImage
    
    '读取最后一个缩略图
    With ucPreview.ImgViewer
        If .Images.Count <= 0 Then Exit Sub
        Set ImgTmp = .Images(.Images.Count)
    End With
    
    Call SaveImage(ImgTmp, lngAdviceId, strStudyUid, iEncode)
End Sub

Private Sub SaveImage(ImgTmp As DicomImage, ByVal lngAdviceId As Long, ByVal strStudyUid As String, Optional iEncode As Integer = 0, Optional blnSave As Boolean)
    Dim rsTmp As New ADODB.Recordset
    
    
    Dim dtReceived As String
    Dim blnFirstImage As String     '是否本次检查的第一张图像
    Dim nowTime As Date
    Dim strReportImages As String
    Dim arrSql() As Variant
    Dim blnInTrans As Boolean       '在事物处理过程中
    Dim i As Integer
    Dim lngSendNo As Long
    Dim strSQL As String
    Dim imgTag As clsImageTagInf
    Dim blnContinue As Boolean
    
    '先保存FTP图像
    '读取接收日期
    strSQL = "select 姓名, 性别, 年龄, 检查UID ,接收日期,报告图象,发送号 from 影像检查记录 where 医嘱ID = [1] "
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, App.ProductName, lngAdviceId)
    
    nowTime = zlDatabase.Currentdate
    
    If IsNull(rsTmp("检查UID")) Then
        dtReceived = Format(nowTime, "yyyyMMdd")
        blnFirstImage = True
    Else
        dtReceived = Format(rsTmp("接收日期"), "yyyyMMdd")
        blnFirstImage = False
    End If
    
    '创建缓冲目录
    MkLocalDir mstrBufferDir & dtReceived & "\" & strStudyUid & "\"
    lngSendNo = rsTmp!发送号
    
    Set imgTag = ImgTmp.Tag
    
    If Not blnSave Then
        blnContinue = imgTag.Tag <> VIDEOTAG And imgTag.Tag <> AUDIOTAG
    Else
        blnContinue = True
    End If
    
    If blnContinue Then
    
        strReportImages = Nvl(rsTmp("报告图象"))
        
        '检查报告图象的长度，如果超过4000个字节，则提示无法保存图像
        If Len(strReportImages & " ;" & ImgTmp.InstanceUID & ".jpg") >= 4000 Then
            MsgboxCus "图像数量超过上限，请先删除部分图像后，再继续采集图像。", vbInformation, G_STR_HINT_TITLE
            Call ucPreview.DeleteImage(ucPreview.CurImageCount)
            Exit Sub
        End If
        
        '保存图像到缓存目录
        Select Case iEncode
            Case 1          'Run-Length Encoding行程压缩
                ImgTmp.WriteFile mstrBufferDir & dtReceived & "\" & strStudyUid & "\" & ImgTmp.InstanceUID, True, "1.2.840.10008.1.2.5"
            Case 2          '不处理，保持原图的压缩方式
                ImgTmp.WriteFile mstrBufferDir & dtReceived & "\" & strStudyUid & "\" & ImgTmp.InstanceUID, True
            Case Else       'Lossless JPEG encoding JPEG无损压缩
                ImgTmp.WriteFile mstrBufferDir & dtReceived & "\" & strStudyUid & "\" & ImgTmp.InstanceUID, True, "1.2.840.10008.1.2.4.70"
        End Select
        
        If gtFileLoadType <> Service Then ImgTmp.FileExport mstrBufferDir & dtReceived & "\" & strStudyUid & "\" & ImgTmp.InstanceUID & ".jpg", "JPG", 80
    End If
    
BUGEX "subSaveImage gtFileLoadType = " & gtFileLoadType

    If gtFileLoadType = Service Then
        If Not SaveImageWithService(lngAdviceId, strStudyUid, dtReceived, rsTmp, ImgTmp) Then Exit Sub
    Else
        If Not SaveImageWithNormal(lngAdviceId, strStudyUid, dtReceived, ImgTmp) Then Exit Sub
    End If
    
    '图像存储成功后，存储数据库信息
    On Error GoTo DBError
    arrSql = Array()
    
    If blnFirstImage Then
        strSQL = "ZL_影像检查记录_SET(" & lngAdviceId & "," & lngSendNo & ",'" & _
            strStudyUid & "',null," & _
            "to_Date('" & Format(nowTime, "yyyy-mm-dd hh:mm") & "','YYYY-MM-DD HH24:MI:SS'),'" & mobjFtp.strDeviceId & "')"
        ReDim Preserve arrSql(UBound(arrSql) + 1)
        arrSql(UBound(arrSql)) = strSQL
    End If
    
    strSQL = "Select 序列UID From 影像检查序列  Where 序列UID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "PACS图像保存", CStr(ImgTmp.SeriesUID))
    
    '插入新的检查序列,如果为录像，则插入新的序列
    If rsTmp.EOF Or ImgTmp.Tag.Tag = VIDEOTAG Or ImgTmp.Tag.Tag = AUDIOTAG Then
        If ImgTmp.Tag.Tag = VIDEOTAG Then
            strSQL = "ZL_影像序列_INSERT('" & strStudyUid & "','" & ImgTmp.SeriesUID & "','视频录像',0)"
        ElseIf ImgTmp.Tag.Tag = AUDIOTAG Then
            strSQL = "ZL_影像序列_INSERT('" & strStudyUid & "','" & ImgTmp.SeriesUID & "','音频数据',0)"
        Else
            strSQL = "ZL_影像序列_INSERT('" & strStudyUid & "','" & ImgTmp.SeriesUID & "','" & ImgTmp.SeriesDescription & "',0)"
        End If
        
        ReDim Preserve arrSql(UBound(arrSql) + 1)
        arrSql(UBound(arrSql)) = strSQL
    End If
    
    If ImgTmp.Tag.Tag <> VIDEOTAG And ImgTmp.Tag.Tag <> AUDIOTAG Then
        '插入新的图像记录
        strSQL = "ZL_影像图象_INSERT('" & ImgTmp.InstanceUID & "','" & ImgTmp.SeriesUID & "',NULL,0, null, sysdate)"
    Else
        If ImgTmp.Tag.Tag = VIDEOTAG Then
            '插入新的视频记录
            strSQL = "ZL_影像图象_INSERT('" & ImgTmp.InstanceUID & "','" & ImgTmp.SeriesUID & "',Null,0" & _
            ",null,sysdate,null,null,null,null,null,null,null,null,null," & VIDEOTAG & ",'" & mstrEncoderName & "'," & ImgTmp.Tag.RecordTimeLen & ")"
        Else
            '插入新的音频记录
            strSQL = "ZL_影像图象_INSERT('" & ImgTmp.InstanceUID & "','" & ImgTmp.SeriesUID & "',Null,0" & _
            ",null,sysdate,null,null,null,null,null,null,null,null,null," & AUDIOTAG & ",''," & ImgTmp.Tag.RecordTimeLen & ")"
        End If
    End If
    
    ReDim Preserve arrSql(UBound(arrSql) + 1)
    arrSql(UBound(arrSql)) = strSQL
    
    gcnVideoOracle.BeginTrans        '----------保存图像
    blnInTrans = True
    For i = 0 To UBound(arrSql)
        Call zlDatabase.ExecuteProcedure(CStr(arrSql(i)), "保存图像")
    Next i
    
    gcnVideoOracle.CommitTrans
    blnInTrans = False
    
    If ImgTmp.Tag.Tag <> VIDEOTAG And ImgTmp.Tag.Tag <> AUDIOTAG Then
        Call frmCaptureHint.ShowCaptureHint( _
            IIf(gobjCapturePar.IsWindowHint, mstrBufferDir & dtReceived & "\" & strStudyUid & "\" & ImgTmp.InstanceUID, ""), _
            gobjCapturePar.IsSountHint, hpRB, Me)
    End If

    If mintCaptureFlag = 1 Or mintCaptureFlag = 4 Or mintCaptureFlag = 5 Then
        If ucPreview.CurImageCount = 1 Then
            Call DoStateChange(vetCaptureFirstImg, lngAdviceId, lngSendNo, strStudyUid)
        Else
            '更新图像
            Call DoStateChange(vetUpdateImg, lngAdviceId, lngSendNo, strStudyUid)
        End If
    ElseIf mintCaptureFlag = 2 Then
        '设置影像检查状态，如果采集第一张图，且原来的状态是已报到，则修改成已检查
        If ucPreview.ImgViewer.Images.Count = 1 Then
            If mlngStudyState < 3 Then
                strSQL = "Zl_影像检查_State(" & lngAdviceId & "," & lngSendNo & ",3,NULL,'" & UserInfo.编号 & "','" & UserInfo.姓名 & "'," & glngDepartId & ")"
                zlDatabase.ExecuteProcedure strSQL, "采集第一个图像"
                mlngStudyState = 3
            End If
        End If
        
        If ucPreview.ImgViewer.Images.Count = 1 Then
            '采集第一副图像
            Call DoStateChange(vetCaptureFirstImg, lngAdviceId, lngSendNo, strStudyUid)
        Else
            '更新图像
            Call DoStateChange(vetUpdateImg, lngAdviceId, lngSendNo, strStudyUid)
        End If
    ElseIf mintCaptureFlag = 3 Then
        '设置影像检查状态，如果采集第一张图，且原来的状态是已报到，则修改成已检查
        If ucPreview.CurImageCount = 1 Then
            If mlngStudyState < 3 Then
                strSQL = "Zl_影像检查_State(" & lngAdviceId & "," & lngSendNo & ",3,NULL,'" & UserInfo.编号 & "','" & UserInfo.姓名 & "'," & glngDepartId & ")"
                zlDatabase.ExecuteProcedure strSQL, "采集第一个图像"
                mlngStudyState = 3
            End If
        End If
        
        If ucPreview.CurImageCount = 1 Then
            Call DoStateChange(vetCaptureFirstImg, lngAdviceId, lngSendNo, strStudyUid)
        Else
            '更新图像
            Call DoStateChange(vetUpdateImg, lngAdviceId, lngSendNo, strStudyUid)
        End If
    End If
    Exit Sub
DBError:
    '出错，则回退数据库操作，并且删除所采集的图像
    If blnInTrans = True Then gcnVideoOracle.RollbackTrans
    err.Raise err.Number, "检查图像保存"
    Call ucPreview.DeleteImage(ucPreview.CurImageCount)
End Sub

Private Function SaveImageWithNormal(ByVal lngAdviceId As Long, ByVal strStudyUid As String, ByVal dtReceived As String, ImgTmp As DicomImage) As Boolean
'使用最原始的方式上传图像
    Dim lngResult As Long
    Dim lngUpLoadResult As Long '成功:0，FTP连接失败:1，上传文件失败:2
    
    lngResult = mobjFtpConnection.FuncFtpConnect(mobjFtp.strFTPIP, mobjFtp.strFTPUser, mobjFtp.strFTPPwd)
    If lngResult = 0 Then
        'FTP操作失败，提示错误，并删除缩略图中的图像
        MsgboxCus "FTP连接失败，图像无法保存，请检查网络设置。", vbInformation, G_STR_HINT_TITLE
        Call ucPreview.DeleteImage(ucPreview.CurImageCount)
    
        Exit Function
    End If
    
    If Val(mobjBakFtp.strDeviceId) > 0 Then
        lngResult = mobjBakFtpConnection.FuncFtpConnect(mobjBakFtp.strFTPIP, mobjBakFtp.strFTPUser, mobjBakFtp.strFTPPwd)
        If lngResult = 0 Then
            mobjBakFtp.strDeviceId = ""
            MsgboxCus "备份ftp设备连接失败，采集的图像将不能进行备份操作，如需备份请检查流程管理中的备份设备配置。", vbInformation, G_STR_HINT_TITLE
        End If
    End If
    
    If ImgTmp.Tag.Tag <> VIDEOTAG And ImgTmp.Tag.Tag <> AUDIOTAG Then
        '保存dicom图像
        lngUpLoadResult = WriteToURL(mstrBufferDir & dtReceived & "\" & strStudyUid & "\" & ImgTmp.InstanceUID, _
                                    mobjFtp.strFtpDir & dtReceived & "/" & strStudyUid & "/" & ImgTmp.InstanceUID, True)
        
        If Not ShowMessage(lngUpLoadResult) Then Exit Function
            
        lngUpLoadResult = WriteToURL(mstrBufferDir & dtReceived & "\" & strStudyUid & "\" & ImgTmp.InstanceUID & ".jpg", _
                                    mobjFtp.strFtpDir & dtReceived & "/" & strStudyUid & "/" & ImgTmp.InstanceUID & ".jpg")
        
        If Not ShowMessage(lngUpLoadResult) Then Exit Function
        
        '备份当前采集的图像
        If mobjBakFtpConnection.hConnection <> 0 Then
            lngUpLoadResult = BakImgToURL(mstrBufferDir & dtReceived & "\" & strStudyUid & "\" & ImgTmp.InstanceUID, _
                                        mobjBakFtp.strFtpDir & dtReceived & "/" & strStudyUid & "/" & ImgTmp.InstanceUID)
            
            If lngUpLoadResult <> 0 Then
                MsgboxCus "文件备份失败，可能由于网络不稳定造成。", vbInformation, G_STR_HINT_TITLE
            End If
        End If
    Else
        '保存录像
        lngUpLoadResult = WriteToURL(ImgTmp.Tag.VideoFile, _
                                    mobjFtp.strFtpDir & dtReceived & "/" & strStudyUid & "/" & ImgTmp.InstanceUID)
        
        If Not ShowMessage(lngUpLoadResult) Then Exit Function
        
        '备份录像
        If mobjBakFtpConnection.hConnection <> 0 Then
            lngUpLoadResult = BakImgToURL(ImgTmp.Tag.VideoFile, _
                                        mobjBakFtp.strFtpDir & dtReceived & "/" & strStudyUid & "/" & ImgTmp.InstanceUID)
            
            If lngUpLoadResult <> 0 Then
                MsgboxCus "文件备份失败，可能由于网络不稳定造成。", vbInformation, G_STR_HINT_TITLE
            End If
        End If
        
        If ImgTmp.Tag.Tag = VIDEOTAG Then
            '将录像复制到对应的目录中，避免重新下载
            Call FileCopy(ImgTmp.Tag.VideoFile, mstrBufferDir & dtReceived & "\" & strStudyUid & "\" & ImgTmp.InstanceUID & ".avi")
            Call Kill(ImgTmp.Tag.VideoFile)
        
            ImgTmp.Tag.VideoFile = mstrBufferDir & dtReceived & "\" & strStudyUid & "\" & ImgTmp.InstanceUID & ".avi"
        ElseIf ImgTmp.Tag.Tag = AUDIOTAG Then
            '将音频文件复制到对应的目录中，避免重新下载
            Call FileCopy(ImgTmp.Tag.VideoFile, mstrBufferDir & dtReceived & "\" & strStudyUid & "\" & ImgTmp.InstanceUID & ".wav")
            Call Kill(ImgTmp.Tag.VideoFile)
        
            ImgTmp.Tag.VideoFile = mstrBufferDir & dtReceived & "\" & strStudyUid & "\" & ImgTmp.InstanceUID & ".wav"
        End If
    End If
    
    mobjFtpConnection.FuncFtpDisConnect
    If mobjBakFtpConnection.hConnection <> 0 Then mobjBakFtpConnection.FuncFtpDisConnect
    
    SaveImageWithNormal = True
End Function

Private Function SaveImageWithService(ByVal lngAdviceId As Long, ByVal strStudyUid As String, ByVal dtReceived As String, rsTmp As ADODB.Recordset, ImgTmp As DicomImage) As Boolean
'使用Service服务后台上传
    Dim strSQL As String
    Dim fileMsg As TransferFileMsg
    
    If ImgTmp.Tag.Tag = VIDEOTAG Then
        '将录像移动到对应的目录中，避免重新下载
        Name ImgTmp.Tag.VideoFile As mstrBufferDir & dtReceived & "/" & strStudyUid & "/" & ImgTmp.InstanceUID & ".avi"
    
        ImgTmp.Tag.VideoFile = mstrBufferDir & dtReceived & "/" & strStudyUid & "/" & ImgTmp.InstanceUID & ".avi"
    ElseIf ImgTmp.Tag.Tag = AUDIOTAG Then
        '将音频文件移动到对应的目录中，避免重新下载
        Name ImgTmp.Tag.VideoFile As mstrBufferDir & dtReceived & "/" & strStudyUid & "/" & ImgTmp.InstanceUID & ".wav"
    
        ImgTmp.Tag.VideoFile = mstrBufferDir & dtReceived & "/" & strStudyUid & "/" & ImgTmp.InstanceUID & ".wav"
    End If
    
'BUGEX "strFTPIP = " & mobjFtp.strFTPIP & " strFTPUser = " & mobjFtp.strFTPUser & " strFTPPwd = " & mobjFtp.strFTPPwd
'BUGEX "strFtpDir = " & mobjFtp.strFtpDir & dtReceived & "/" & strStudyUid & "/"
'BUGEX "lngAdviceId = " & lngAdviceId
'BUGEX "LOCALDIR = " & mstrBufferDir & dtReceived & "/" & strStudyUid & "/" & " FILENAME = " & ImgTmp.InstanceUID
'BUGEX "strBakFTPIP = " & mobjBakFtp.strFTPIP & " strBakFTPUser = " & mobjBakFtp.strFTPUser & " strBakFTPPwd = " & mobjBakFtp.strFTPPwd
'BUGEX "strFtpDir = " & mobjBakFtp.strFtpDir & dtReceived & "/" & strStudyUid & "/"
    
    With fileMsg
        fileMsg.strAdviceId = lngAdviceId
        fileMsg.strName = Nvl(rsTmp("姓名"))
        fileMsg.strSex = Nvl(rsTmp("性别"))
        fileMsg.strAge = Nvl(rsTmp("年龄"))
        
        fileMsg.ftpInfo.strDeviceId = mobjFtp.strDeviceId
        fileMsg.ftpInfo.strFtpDir = mobjFtp.strFtpDir
        fileMsg.ftpInfo.strFTPIP = mobjFtp.strFTPIP
        fileMsg.ftpInfo.strFTPPwd = mobjFtp.strFTPPwd
        fileMsg.ftpInfo.strFTPUser = mobjFtp.strFTPUser
        fileMsg.ftpInfo.strSDDir = mobjFtp.strSDDir
        fileMsg.ftpInfo.strSDPswd = mobjFtp.strSDPswd
        fileMsg.ftpInfo.strSDUser = mobjFtp.strSDUser
        
        fileMsg.bakFtpInfo.strDeviceId = mobjBakFtp.strDeviceId
        fileMsg.bakFtpInfo.strFtpDir = mobjBakFtp.strFtpDir
        fileMsg.bakFtpInfo.strFTPIP = mobjBakFtp.strFTPIP
        fileMsg.bakFtpInfo.strFTPPwd = mobjBakFtp.strFTPPwd
        fileMsg.bakFtpInfo.strFTPUser = mobjBakFtp.strFTPUser
        fileMsg.bakFtpInfo.strSDDir = mobjBakFtp.strSDDir
        fileMsg.bakFtpInfo.strSDPswd = mobjBakFtp.strSDPswd
        fileMsg.bakFtpInfo.strSDUser = mobjBakFtp.strSDUser
        
        fileMsg.strLocalDir = mstrBufferDir & dtReceived & "\" & strStudyUid & "\" & ImgTmp.InstanceUID
        fileMsg.strFileName = ImgTmp.InstanceUID
        fileMsg.strSubDir = dtReceived & "/" & strStudyUid & "/" & ImgTmp.InstanceUID
        fileMsg.strMediaType = ImgTmp.Tag.Tag
    End With

    If Not SendDataToservice("缩略图", COMMAND_CAPIMG_UPLOAD, "图像采集", fileMsg) Then
BUGEX "图像信息未成功发送至服务器"
        MsgboxEx Me.hwnd, "将图像数据发送至服务管理器时出错，可能是ZLPacsServerCenter服务未启用！" & vbCrLf & _
                          "数据将临时保存到本地，待下次打开服务时尝试自动上传！", vbOKOnly, G_STR_HINT_TITLE
            
        SaveImageWithService = True
        Exit Function
    Else
BUGEX "图像信息成功发送至服务器"
        SaveImageWithService = True
    End If
End Function

Private Function WriteToURL(ByVal SrcFileName As String, ByVal DestFileName As String, Optional ByVal blnCheck As Boolean = False) As Long
'------------------------------------------------
'功能：将本地文件保存到远程网络上
'参数：SrcFileName--本地文件名，DestFileName－－远程文件名
'      blnCheck -- 上传文件后应进行检查（比较ftp上和本地的文件大小是否相同）
'返回：成功返回0，连接失败返回1，上传文件失败返回2
'-----------------------------------------------
'功能：
    Dim objFileSystem As New Scripting.FileSystemObject
    Dim strPath As String, StrMessage As String
    Dim lngFtpFileSzie As Long, lngDestFileSize As Long
    
    '在FTP中创建目录
    strPath = objFileSystem.GetParentFolderName(DestFileName)
    mobjFtpConnection.FuncFtpMkDir "/", strPath
    
ReUplcad:
    
    '向FTP上传文件
    WriteToURL = mobjFtpConnection.FuncUploadFile(strPath, Replace(SrcFileName, "/", "\"), objFileSystem.GetFileName(DestFileName))
    
    '上传成功后对比一下大小，判断是否正常上传
    If blnCheck And WriteToURL = 0 And mblnCompareSize Then
        lngDestFileSize = objFileSystem.GetFile(SrcFileName).Size
        lngFtpFileSzie = mobjFtpConnection.FuncFtpGetFileSize(strPath, objFileSystem.GetFileName(DestFileName))
        
        If lngFtpFileSzie < lngDestFileSize Then
            StrMessage = "上传后的文件大小[" & lngFtpFileSzie & "]与原文件大小[" & lngDestFileSize & "]不一致" & vbCrLf & _
                         "原文件：" & SrcFileName & vbCrLf & _
                         "FTP文件：" & strPath & objFileSystem.GetFileName(DestFileName) & vbCrLf & _
                         "是否需要重新上传？"
            
            If MsgBox(StrMessage, vbQuestion + vbYesNo, "提示") = vbYes Then
                GoTo ReUplcad
            End If
        End If
    End If
End Function


Private Function BakImgToURL(ByVal SrcFileName As String, ByVal DestFileName As String) As Long
'------------------------------------------------
'功能：备份图像到远程网络上
'参数：SrcFileName--本地文件名，DestFileName－－远程文件名
'返回：成功返回0，连接失败返回1，上传文件失败返回2
'-----------------------------------------------
    Dim objFileSystem As New Scripting.FileSystemObject
    Dim strPath As String
    
    If mobjBakFtpConnection.hConnection = 0 Then Exit Function
    
    '在FTP中创建目录
    strPath = objFileSystem.GetParentFolderName(DestFileName)
    mobjBakFtpConnection.FuncFtpMkDir "/", strPath
    
    '向FTP上传文件
    BakImgToURL = mobjBakFtpConnection.FuncUploadFile(strPath, Replace(SrcFileName, "/", "\"), objFileSystem.GetFileName(DestFileName))
End Function


Private Sub RemoveFromURL(ByVal DestFileName As String)
'------------------------------------------------
'功能：从FTP删除文件
'参数：DestFileName－－远程文件名
'返回：无
'-----------------------------------------------
    Dim objFileSystem As New Scripting.FileSystemObject
    
    mobjFtpConnection.FuncDelFile objFileSystem.GetParentFolderName(DestFileName), objFileSystem.GetFileName(DestFileName)
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
    
    '是否显示处理工具栏
    mblnShowProcessBar = GetSetting("ZLSOFT", G_STR_REG_PATH_PUBLIC, "显示处理工具栏", "True")
    
    BUGEX "InitCommandBars:4"
    
    '采集工具栏定义
    Set cbrToolBar = Me.cbrMain.Add("采集栏", xtpBarLeft)
    cbrToolBar.EnableDocking xtpFlagStretched
    cbrToolBar.ShowTextBelowIcons = False
    cbrToolBar.Closeable = False
    
    With cbrToolBar.Controls
        '在非TWAIN采集模式的情况下，才显示该按钮
        'If Not GetIsTwainCaptureWay Then
            Set cbrControl = .Add(xtpControlButton, conMenu_Cap_Dynamic, "动态"): cbrControl.ToolTipText = "显示实时视频"
        'End If
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Cap_MarkMap, "采集"): cbrControl.ToolTipText = "采集图像"
        
        '启用后台采集
        Set cbrControl = .Add(xtpControlButton, conMenu_Cap_After_Capture, "后台采集"): cbrControl.ToolTipText = "后台采集"
            cbrControl.IconId = 10020
        
        '在非TWAIN采集模式的情况下，才显示该按钮
        'If Not GetIsTwainCaptureWay Then
            Set cbrControl = .Add(xtpControlButton, conMenu_Cap_Record, "录像"): cbrControl.ToolTipText = "开始录像"
                cbrControl.Enabled = True
                
            
            Set cbrControl = .Add(xtpControlButton, conMenu_Cap_After_Record, "后台录像"): cbrControl.ToolTipText = "后台录像"
                cbrControl.IconId = 10021
            
            
            Set cbrControl = .Add(xtpControlButton, conMenu_Cap_Record_Stop, "停止录像"): cbrControl.ToolTipText = "停止录像"
                cbrControl.Enabled = False
                
            Set cbrControl = .Add(xtpControlButton, conMenu_Cap_RecordAudio, "录音"): cbrControl.ToolTipText = "录音"
        'End If
        
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Cap_Play, "播放"): cbrControl.ToolTipText = "播放录像"
            cbrControl.BeginGroup = True
        
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Cap_Import, "导入"): cbrControl.ToolTipText = "文件导入"
            cbrControl.IconId = 10002
            cbrControl.BeginGroup = True
            
            
        Set cbrControl = .Add(xtpControlButton, conMenu_Cap_SaveAs, "另存"): cbrControl.ToolTipText = "文件另存": cbrControl.IconId = 3091
            cbrControl.IconId = 10004
            
        Set cbrControl = .Add(xtpControlButton, conMenu_Cap_DelImg, "删图"): cbrControl.ToolTipText = "删除图像": cbrControl.IconId = 10001
            
            
        Set cbrControl = .Add(xtpControlButton, conMenu_Cap_OpenStudyList, "打开检查"): cbrControl.ToolTipText = "打开检查"
            cbrControl.BeginGroup = True
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Cap_StudySyncState, "锁定检查"): cbrControl.ToolTipText = "锁定检查"
            cbrControl.IconId = 10012
        
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Cap_After_Tag, "更新标识"): cbrControl.ToolTipText = "更新标识"
            cbrControl.IconId = 10022
        
    End With
    
    For Each cbrControl In cbrToolBar.Controls
        cbrControl.Style = xtpButtonIcon '  xtpButtonIconAndCaption
        cbrControl.Category = "采集"
        cbrControl.Enabled = False
    Next
    
    Set cbrToolBar = Me.cbrMain.Add("处理栏", xtpBarRight)
    cbrToolBar.ShowTextBelowIcons = False
    cbrToolBar.EnableDocking xtpFlagStretched
    With cbrToolBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Process_Window, "亮度"): cbrControl.ToolTipText = "调节亮度/对比度"
        Set cbrControl = .Add(xtpControlButton, conMenu_Process_Zoom, "缩放"): cbrControl.ToolTipText = "缩放图像"
        Set cbrControl = .Add(xtpControlButton, conMenu_Process_Corp, "拖动"): cbrControl.ToolTipText = "拖动图像"
        Set cbrControl = .Add(xtpControlButton, conMenu_Process_RectZoom, "裁剪采集"): cbrControl.ToolTipText = "裁剪采集图像": cbrControl.IconId = 3201
        Set cbrControl = .Add(xtpControlButton, conMenu_Process_RRotate, "顺时"): cbrControl.ToolTipText = "顺时针旋转"
        Set cbrControl = .Add(xtpControlButton, conMenu_Process_LRotate, "逆时"): cbrControl.ToolTipText = "逆时针旋转"
        Set cbrControl = .Add(xtpControlButton, conMenu_Process_Sharpness, "锐化"): cbrControl.ToolTipText = "锐化"
        Set cbrControl = .Add(xtpControlButton, conMenu_Process_Filter, "平滑"): cbrControl.ToolTipText = "平滑"
        Set cbrControl = .Add(xtpControlButton, conMenu_Process_Arrow, "箭头"): cbrControl.ToolTipText = "箭头标注"
        cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Process_Ellipse, "圆形"): cbrControl.ToolTipText = "圆形标注"
        Set cbrControl = .Add(xtpControlButton, conMenu_Process_Text, "文字"): cbrControl.ToolTipText = "文字标注"
'        Set cbrControl = .Add(xtpControlButton, conMenu_Tool_Analyse, "高级"): cbrControl.ToolTipText = "高级处理"
        cbrControl.BeginGroup = True
    End With
    
    For Each cbrControl In cbrToolBar.Controls
        cbrControl.Style = xtpButtonIcon
        cbrControl.Category = "处理"
        cbrControl.Enabled = False
    Next
    cbrToolBar.Visible = mblnShowProcessBar
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
        Set cbrControl = .Add(xtpControlButton, conMenu_Process_RectCapture, "裁剪采集")
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
    
    '把结果图加入到viewer中并且保存
    '设置图像的DICOM参数
    subWriteDicomPara imgResult, mlngAdviceId
    
    Dim dcmTag As New clsImageTagInf
    dcmTag.Tag = imgTag
    
    Set imgResult.Tag = dcmTag
    
    mintCaptureFlag = 1
    
    '将图像插入到缩略图中
    subInsert2Mini imgResult
    
    '保存图像，并触发图像存储事件
    Call subSaveImage(mlngAdviceId, mcurStudyInf.strStudyUid)
End Sub


Private Sub ucCapHook_OnKeyBoardLHook(ByVal lngVkCode As Long, ByVal lngScanCode As Long, ByVal lngFlags As Long)
On Error GoTo errHandle
    Select Case lngVkCode
        Case 66
            '判断键盘按键是否松开，为0表示按下键盘
            If lngScanCode = 128 Then
                '执行快捷采集
'                Call CaptureImage

                If timerHook.Enabled Then Exit Sub
                timerHook.Enabled = True
            End If
    End Select
Exit Sub
errHandle:
    MsgboxCus err.Description, vbOKOnly, G_STR_HINT_TITLE
End Sub

Private Sub ucPreview_OnMouseDown(Button As Integer, Shift As Integer, X As Long, Y As Long)
'    If ucPreview.ImgViewer.Images.Count <= 0 Then Exit Sub
'    If Button = 1 Then ucPreview.ImgChecked(ucPreview.SelectIndex) = Not ucPreview.ImgChecked(ucPreview.SelectIndex)
End Sub

Private Sub ucPreview_OnMouseUp(Button As Integer, Shift As Integer, X As Long, Y As Long)
    If ucPreview.ImgViewer.Images.Count <= 0 Then Exit Sub
    If Button = 2 Then Call ShowPopupImage(1)
End Sub

Private Sub ucCacheImage_OnMouseDown(Button As Integer, Shift As Integer, X As Long, Y As Long)
'    If ucCacheImage.ImgViewer.Images.Count <= 0 Then Exit Sub
'    If Button = 1 Then ucCacheImage.ImgChecked(ucCacheImage.SelectIndex) = Not ucCacheImage.ImgChecked(ucCacheImage.SelectIndex)
End Sub

Private Sub ucCacheImage_OnMouseUp(Button As Integer, Shift As Integer, X As Long, Y As Long)
    If ucCacheImage.ImgViewer.Images.Count <= 0 Then Exit Sub
    If Button = 2 Then Call ShowPopupImage(2)
End Sub

Private Sub ShowPopupImage(ByVal intType As Integer)
'------------------------------------------------
'功能：创建鼠标右键弹出菜单
'intType:0--报告图，1--缩略图，2--缓存图
'------------------------------------------------
    Dim cbrControl As CommandBarControl
    Dim cbrToolBar As CommandBar
    Dim cbrToolPopup As CommandBarPopup
    
    If intType <> 2 Then
        If ucPreview.CurImageCount < 1 Then Exit Sub
    End If
    
    '鼠标右键弹出菜单
    Set cbrToolBar = cbrMain.Add("鼠标右键", xtpBarPopup)
    cbrToolBar.EnableDocking xtpFlagStretched
    cbrToolBar.ShowTextBelowIcons = False
    cbrToolBar.Closeable = False
    With cbrToolBar.Controls
        If intType = 1 Then
            Set cbrControl = .Add(xtpControlButton, conMenu_Cap_SplitPage, "分页设置")
            
            Set cbrControl = .Add(xtpControlButton, conMenu_Cap_ImageShield, "屏蔽大图")
            Call GetLocalPar
            cbrControl.Checked = mblnImageShield
            
            If glngModule = 1291 And gobjCapturePar.IsUseAfterCapture Then
                Set cbrControl = .Add(xtpControlButton, conMenu_Cap_ImgExport, "发送到后台")
                cbrControl.BeginGroup = True
            End If
            
            If mblnIsReported = False Then
                Set cbrControl = .Add(xtpControlButton, conMenu_Cap_AddToReport, "加入报告图")
                cbrControl.BeginGroup = True
            End If
            
            Set cbrControl = .Add(xtpControlButton, conMenu_Cap_ReUpLoad, "重新上传")
                cbrControl.BeginGroup = True
                
            Set cbrControl = .Add(xtpControlButton, conMenu_Cap_ReDownLoad, "重新下载")
            
            Set cbrControl = .Add(xtpControlButton, conMenu_Cap_ImageProcess, "图像处理")
                cbrControl.BeginGroup = True
            Set cbrControl = .Add(xtpControlButton, conMenu_Cap_ImageDel, "删除图像")
        Else
            Set cbrControl = .Add(xtpControlButton, conMenu_Cap_SplitPage, "分页设置")
            
            Set cbrControl = .Add(xtpControlButton, conMenu_Cap_ImgImport, "发送到检查")
                cbrControl.BeginGroup = True
                
            Set cbrControl = .Add(xtpControlButton, conMenu_Cap_DelCache, "删除")
                cbrControl.BeginGroup = True
                
            Set cbrControl = .Add(xtpControlButton, conMenu_Cap_RefreshCache, "刷新")
        End If
    End With
    cbrToolBar.Visible = True
    cbrToolBar.ShowPopup
End Sub

Private Sub ucPreview_OnClick(ByVal lngSelectedIndex As Long)

    mCorpSize.X = 0
    mCorpSize.Y = 0
    
    '被选中图像显示红框
    If lngSelectedIndex <> 0 Then
        
        Call PreviewThumbnail(lngSelectedIndex)

        '设置视频的当前显示状态
        Call ConfigVideoShowState(False)
    End If
    
    '恢复当前控件焦点，以便能够滚动图像
    ucPreview.SetFocus
End Sub


Private Sub PreviewThumbnail(ByVal lngImgIndex As Long)
'预览缩略图
    Dim dblTempZoom As Double
    Dim img As DicomImage
    Dim i As Integer
    
    '将被选中图像装载到dcmView中
    dcmView.Images.Clear
    
    If lngImgIndex <= 0 Then Exit Sub
    dcmView.Images.Add ucPreview.ImgViewer.Images(lngImgIndex)
    
    Set img = dcmView.Images(1)
    
    '去除边框
    For i = 1 To img.Labels.Count
        If img.Labels(i).Tag = "SELECT" Or img.Labels(i).Tag = "BORDER" Or img.Labels(i).Tag = "FAILD" Then
            img.Labels(i).Visible = False
        End If
    Next
    
    '显示dcmView，隐藏picVideo
    dcmView.CurrentImage.BorderWidth = 0
    
    dblTempZoom = dcmView.CurrentImage.ActualZoom
    dcmView.CurrentImage.StretchToFit = False
        
    '判断当进入浮动窗口时，缩放比率不能小于0.1
    If dblTempZoom < 0.1 Then dblTempZoom = 0.1
                  
    Call subCenterZoom(Me, dcmView.CurrentImage, dcmView, dblTempZoom, mCorpSize)
End Sub


Private Sub ucPreview_OnDbClick(ByVal lngSelectedIndex As Long, blnContinue As Boolean)
'双击时播放音视频文件
On Error GoTo errHandle
    If lngSelectedIndex <= 0 Or lngSelectedIndex > ucPreview.CurImageCount Then Exit Sub
    
    If Not ucPreview.SelectImage.Tag Is Nothing Then
        If UCase(TypeName(ucPreview.SelectImage.Tag)) = UCase("clsImageTagInf") Then
            If ucPreview.SelectImage.Tag.Tag = VIDEOTAG Or ucPreview.SelectImage.Tag.Tag = AUDIOTAG Then
                Call subVideoPlay
                blnContinue = False
            End If
        End If
    End If
    
Exit Sub
errHandle:
    MsgboxCus err.Description, vbOKOnly, G_STR_HINT_TITLE
End Sub

Private Sub ucPreview_OnReUpload()
On Error GoTo errHandle
    
    Call ReloadSelectedImg
    
    Exit Sub
errHandle:
    MsgboxCus err.Description, vbOKOnly, G_STR_HINT_TITLE
End Sub

Private Sub ReloadSelectedImg()
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim dtReceived As String
    Dim objSelectedImg As DicomImage
    Dim fileMsg As TransferFileMsg
    
    If ucPreview.IsFailedImg(ucPreview.SelectIndex) = True Then
        MsgboxEx Me.hwnd, "不能对该图像进行上传.", vbOKOnly, G_STR_HINT_TITLE
        Exit Sub
    End If
    
    '重新上传选择的文件
    Set objSelectedImg = ucPreview.SelectImage
    
    strSQL = "select 姓名, 性别, 年龄, 检查UID ,接收日期,报告图象,发送号 from 影像检查记录 where 医嘱ID = [1] "
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, App.ProductName, mlngAdviceId)
    
    If rsTmp.RecordCount <= 0 Or objSelectedImg Is Nothing Then Exit Sub
    
    If IsNull(rsTmp("检查UID")) Then
        dtReceived = Format(zlDatabase.Currentdate, "yyyyMMdd")
    Else
        dtReceived = Format(rsTmp("接收日期"), "yyyyMMdd")
    End If
    
'BUGEX "strFTPIP = " & mobjFtp.strFTPIP & " strFTPUser = " & mobjFtp.strFTPUser & " strFTPPwd = " & mobjFtp.strFTPPwd
'BUGEX "strFtpDir = " & mobjFtp.strFtpDir & dtReceived & "/" & objSelectedImg.StudyUID & "/"
'BUGEX "mlngAdviceId = " & mlngAdviceId
'BUGEX "strBakFTPIP = " & mobjBakFtp.strFTPIP & " strBakFTPUser = " & mobjBakFtp.strFTPUser & " strBakFTPPwd = " & mobjBakFtp.strFTPPwd
'BUGEX "strFtpDir = " & mobjBakFtp.strFtpDir & dtReceived & "/" & objSelectedImg.StudyUID & "/"
'BUGEX "LOCALDIR = " & mstrBufferDir & dtReceived & "/" & objSelectedImg.StudyUID & "/" & " FILENAME = " & objSelectedImg.InstanceUID

    With fileMsg
        fileMsg.strAdviceId = mlngAdviceId
        fileMsg.strName = Nvl(rsTmp("姓名"))
        fileMsg.strSex = Nvl(rsTmp("性别"))
        fileMsg.strAge = Nvl(rsTmp("年龄"))
        
        fileMsg.ftpInfo.strDeviceId = mobjFtp.strDeviceId
        fileMsg.ftpInfo.strFtpDir = mobjFtp.strFtpDir
        fileMsg.ftpInfo.strFTPIP = mobjFtp.strFTPIP
        fileMsg.ftpInfo.strFTPPwd = mobjFtp.strFTPPwd
        fileMsg.ftpInfo.strFTPUser = mobjFtp.strFTPUser
        fileMsg.ftpInfo.strSDDir = mobjFtp.strSDDir
        fileMsg.ftpInfo.strSDPswd = mobjFtp.strSDPswd
        fileMsg.ftpInfo.strSDUser = mobjFtp.strSDUser
        
        fileMsg.bakFtpInfo.strDeviceId = mobjBakFtp.strDeviceId
        fileMsg.bakFtpInfo.strFtpDir = mobjBakFtp.strFtpDir
        fileMsg.bakFtpInfo.strFTPIP = mobjBakFtp.strFTPIP
        fileMsg.bakFtpInfo.strFTPPwd = mobjBakFtp.strFTPPwd
        fileMsg.bakFtpInfo.strFTPUser = mobjBakFtp.strFTPUser
        fileMsg.bakFtpInfo.strSDDir = mobjBakFtp.strSDDir
        fileMsg.bakFtpInfo.strSDPswd = mobjBakFtp.strSDPswd
        fileMsg.bakFtpInfo.strSDUser = mobjBakFtp.strSDUser
        
        fileMsg.strLocalDir = mstrBufferDir & dtReceived & "\" & objSelectedImg.StudyUID & "\" & objSelectedImg.InstanceUID
        fileMsg.strFileName = objSelectedImg.InstanceUID
        fileMsg.strSubDir = dtReceived & "/" & objSelectedImg.StudyUID & "/" & objSelectedImg.InstanceUID
        fileMsg.strMediaType = objSelectedImg.Tag.Tag
    End With
    
    If Not SendDataToservice("缩略图", COMMAND_CAPIMG_UPLOAD, "图像采集", fileMsg) Then
        MsgboxEx Me.hwnd, "将图像数据发送至服务管理器时出错，可能是ZLPacsServerCenter服务未启用！" & vbCrLf & _
                          "数据将临时保存到本地，待下次打开服务时尝试自动上传！", vbOKOnly, G_STR_HINT_TITLE
    Else
BUGEX "图像信息成功发送至服务器"
    End If
End Sub

Private Sub ucPreview_OnSaveImage(ByVal dcmImage As DicomObjects.DicomImage, ByVal lngImageType As Long)
    Select Case lngImageType
        Case 1
            If dcmImage.Tag.ReportImage = "" Then
                dcmImage.Tag.ReportImage = 0
                '先要存为检查图
                Call SaveImageToStady(dcmImage, mlngAdviceId)
                Call AddToReport(dcmImage)
                ucPreview.AfterSaveStudy dcmImage
            End If
        Case 2
            Call SaveImageToStady(dcmImage, mlngAdviceId)
    End Select
End Sub

Private Sub ucSplitter1_OnMoveEnd()
On Error Resume Next
    RaiseEvent OnControlResize(picCapture)
err.Clear
End Sub

Private Sub AddDCMImg(ByVal img As DicomImage)
'如果是当天当前标识，则向右下角图片显示控件中增加图片，可以体现采集成功

On Error GoTo errH
    Dim curTime As Date


    If Not (InStr(cboCacheImage.Text, "标识" & Lpad((mintTagNow), 3, "0")) > 0) Then Exit Sub '不是当前标识退出

    curTime = zlDatabase.Currentdate
    If Format(DTPImg.value, "yyyymmdd") <> Format(curTime, "yyyymmdd") Then Exit Sub '不是当天退出
    
    Call ucCacheImage.AddImage(img, img.Tag)
    
    Exit Sub
errH:
    MsgboxCus err.Description, vbOKOnly, G_STR_HINT_TITLE
End Sub

Private Sub ClearEmptyFolder(ByVal blNoReason As Boolean)
'blNoReason 是否跳过条件执行本过程？用于关闭程序的时候执行
'清空空目录及对应的下拉框，若当前选中的是当天最新标识，则不执行此操作
'首先判断下拉框当前选中是否是当天最新标识
    Dim curTime As Date
    Dim strTime As String
    Dim objFolder1 As Folder, objFolder2 As Folder, objFolder3 As Folder
    Dim strType As String
    Dim strTag As String
    Dim i As Long
    On Error GoTo errH
   
    
    If blNoReason = False Then
    
        '若当前显示有图像，不进行操作
        If Not mblIsNeedRefresh And ucCacheImage.CurImageCount > 0 Then Exit Sub
        
        curTime = zlDatabase.Currentdate
    
        '是当天并且选中的是当前标识，就不进行清空操作
        If Not mblIsNeedRefresh And (Format(DTPImg.value, "yyyymmdd") = Format(curTime, "yyyymmdd")) And (InStr(cboCacheImage.Text, "标识" & Lpad((mintTagNow), 3, "0")) > 0) Then Exit Sub

    End If
    
    If mobjFile.FolderExists(mstrAfterDir) = False Then Exit Sub
    
    If mobjFile.GetFolder(mstrAfterDir).SubFolders.Count > 0 Then
        For Each objFolder1 In mobjFile.GetFolder(mstrAfterDir).SubFolders '''时间
            If objFolder1.SubFolders.Count > 0 Then
                For Each objFolder2 In objFolder1.SubFolders '''检查uid
                    If objFolder2.SubFolders.Count > 0 Then
                    
                        For Each objFolder3 In objFolder2.SubFolders '''序列UID
                            If objFolder3.Files.Count = 0 Then
                                If blNoReason = True Then
                                    Call mobjFile.DeleteFolder(objFolder3.Path)
                                Else
                                    If Not (objFolder1.Name = Format(curTime, "yyyymmdd") And mintTagNow = GetTag(objFolder2.Name, strType)) Then
                                        Call mobjFile.DeleteFolder(objFolder3.Path)
                                    End If
                                End If
                                
                            End If
                        Next
                        
                        If objFolder2.SubFolders.Count = 0 Then Call mobjFile.DeleteFolder(objFolder2.Path)
                    Else
                        Call mobjFile.DeleteFolder(objFolder2.Path)
                    End If
                Next
                
                If objFolder1.SubFolders.Count = 0 Then Call mobjFile.DeleteFolder(objFolder1.Path)
            Else
                Call mobjFile.DeleteFolder(objFolder1.Path)
            End If
        Next
    End If
    Exit Sub
errH:
    MsgboxCus err.Description, vbOKOnly, G_STR_HINT_TITLE
End Sub

Private Function GetTodayTagMax(ByVal curDate As Date) As Integer
    '计算当天最大标识
    Dim strDate As String
    Dim intTmp As Integer
    Dim strType As String
    Dim strStudyUid As String
    Dim objFolder2 As Folder, objFolder3 As Folder
    
    On Error GoTo errH
    
    mintTagMax = 1
    
    strDate = Format(curDate, "yyyymmdd")
    
    If mobjFile.FolderExists(mstrAfterDir) = False Then Exit Function
    
    If mobjFile.GetFolder(mstrAfterDir).SubFolders.Count > 0 Then
        For Each objFolder2 In mobjFile.GetFolder(mstrAfterDir).SubFolders
            If InStr(objFolder2.Name, strDate) > 0 Then                                 '时间选择
            
                If objFolder2.SubFolders.Count > 0 Then                                  '时间层是否有子目录
                
                    For Each objFolder3 In objFolder2.SubFolders                            '检查UID层
                    
                        If objFolder3.SubFolders.Count > 0 Then

                            strStudyUid = GetStudyUIDFromFolderName(objFolder3.Name)
                            
                            intTmp = GetTag(objFolder3.Name, strType)
                            If intTmp > mintTagMax Then mintTagMax = intTmp
                                             
                        End If

                    Next
                    
                End If '时间层是否有子目录
                
            End If '时间选择
        Next
    End If
    
    GetTodayTagMax = mintTagMax
    
    Exit Function
errH:
    BUGEX "GetTodayTagMax output= -1"
    GetTodayTagMax = -1
End Function

Private Function JudgeExistFolderOne() As Boolean
'判断是否存在标识1的目录  t- 存在 f -不存在 默认不存在
    On Error GoTo errH
    Dim objFolder2 As Folder, objFolder3 As Folder, objFolder4 As Folder
    Dim strDate As String

    strDate = Format(zlDatabase.Currentdate, "yyyymmdd")
    JudgeExistFolderOne = False
    ''''''''''''''
    If mobjFile.FolderExists(mstrAfterDir) Then
        For Each objFolder2 In mobjFile.GetFolder(mstrAfterDir).SubFolders
            If objFolder2.Name = strDate Then '时间只判断当天

                For Each objFolder3 In objFolder2.SubFolders   '检查层
                
                    If InStr(objFolder3.Name, "标识1") > 0 Then
                        JudgeExistFolderOne = True
                        Exit Function
                    End If
                Next
                
            End If
        Next
    End If
    
    JudgeExistFolderOne = False
    
    Exit Function
errH:
    BUGEX "JudgeExistFolderOne output= False"
    JudgeExistFolderOne = False
End Function

Private Function GetCountOfCapPic(ByVal intTag As Integer) As Integer '11111
'计算当前标识的图像数量,inttag标识
    On Error GoTo errH
    Dim objFolder2 As Folder, objFolder3 As Folder, objFolder4 As Folder
    Dim strDate As String
    

    strDate = Format(zlDatabase.Currentdate, "yyyymmdd")
    GetCountOfCapPic = 0

    If mobjFile.FolderExists(mstrAfterDir) Then
        For Each objFolder2 In mobjFile.GetFolder(mstrAfterDir).SubFolders
        
            If objFolder2.Name = strDate Then '时间只判断当天

                For Each objFolder3 In objFolder2.SubFolders   '检查层
                
                    If InStr(objFolder3.Name, "检查" & intTag) > 0 Or InStr(objFolder3.Name, "标识" & intTag) > 0 Then '
                    'If InStr(objFolder3.Name, "标识" & inttag) > 0 Then  '
                        For Each objFolder4 In objFolder3.SubFolders   '序列层
                            GetCountOfCapPic = objFolder4.Files.Count
                            Exit Function
                        Next
                        
                    End If
                Next
                      
            End If
        Next
    End If

    Exit Function
errH:
    BUGEX "GetCountOfCapPic output= -1"
    GetCountOfCapPic = -1
End Function

Public Sub setFontSize(ByVal intsize As Integer)
'设置下拉框字号
     
    mintFontSize = intsize
    Call picCacheImage_Resize
    
    DTPImg.Font.Size = intsize
    cboCacheImage.Font.Size = intsize
    dkpReportImage.RedrawPanes

End Sub

Private Function DeleteManyImages(ByVal strArrImgs As String) As Boolean
'删除多个图像(删除数据库内容),strArrImgs:待删除的图像uid
    On Error GoTo errH
    Dim i As Integer
    Dim strSQL As String
    Dim iNet As New clsFtp
    Dim rsTemp As ADODB.Recordset
    Dim strFTPIP As String
    Dim strFtpPass As String
    Dim strFTPUser As String
    Dim strRoot As String
    Dim intDeviceUsed As Integer
    Dim lngResult As Long
    Dim strImageUID As String
    
    DeleteManyImages = False
    strSQL = "Select a.医嘱ID,a.发送号,c.图像UID, " & _
            " Decode(a.接收日期,Null,'',to_Char(a.接收日期,'YYYYMMDD')||'/')||a.检查UID As 图像目录, " & _
            "D.FTP用户名 As User1,D.FTP密码 As Pwd1,D.IP地址 As Host1,'/'||D.Ftp目录||'/' As Root1,d.设备号 as 设备号1," & _
            "E.FTP用户名 As User2,E.FTP密码 As Pwd2,E.IP地址 As Host2,'/'||E.Ftp目录||'/' As Root2,e.设备号 as 设备号2 " & _
            "From 影像检查记录 a,影像检查序列 b,影像检查图象 c,影像设备目录 D,影像设备目录 E, " & _
            "Table(Cast(f_Str2list([1],';') As zlTools.t_Strlist)) F " & _
            "Where a.检查UID=b.检查UID And b.序列UID=c.序列UID And c.图像UID = F.Column_Value " & _
            "And a.位置一=D.设备号(+) And a.位置二=E.设备号(+)"
            
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "PACS删除多个图像", strArrImgs)
    If rsTemp.EOF = True Then
        MsgboxCus "没有找到可以删除的图像!", vbInformation, G_STR_HINT_TITLE
        DeleteManyImages = False
        Exit Function
    End If
    
    '在数据库中删除图像
    gcnVideoOracle.BeginTrans
    rsTemp.MoveFirst
    While Not rsTemp.EOF
        strSQL = "ZL_影像图象_DELETE(" & rsTemp("医嘱ID") & "," & rsTemp("发送号") & ",'" & rsTemp("图像UID") & "','" & "" & "')"
        zlDatabase.ExecuteProcedure strSQL, "影像图像删除"
        rsTemp.MoveNext
    Wend
    gcnVideoOracle.CommitTrans
    '准备在FTP中删除图像，先查找设备一，在查找设备二
    
    rsTemp.MoveFirst
    If Not IsNull(rsTemp!设备号1) Then
        strFTPIP = Nvl(rsTemp!Host1)
        strFTPUser = Nvl(rsTemp!User1)
        strFtpPass = Nvl(rsTemp!Pwd1)

        intDeviceUsed = 1
        lngResult = iNet.FuncFtpConnect(strFTPIP, strFTPUser, strFtpPass)

        If lngResult = 0 Then
            If Not IsNull(rsTemp!设备号2) Then
                strFTPIP = Nvl(rsTemp!Host2)
                strFTPUser = Nvl(rsTemp!User2)
                strFtpPass = Nvl(rsTemp!Pwd2)

                intDeviceUsed = 2
                lngResult = iNet.FuncFtpConnect(strFTPIP, strFTPUser, strFtpPass)

                If lngResult = 0 Then
                    If MsgboxCus("连接FTP失败，是否继续删除图像？" & vbCrLf & "此时继续删除，则只能删除数据库内容，无法删除图像文件。" & vbCrLf & "‘是’则继续删除，‘否’则不删除。", vbQuestion + vbYesNo, G_STR_HINT_TITLE) = vbNo Then
                        DeleteManyImages = False
                        Exit Function
                    End If
                End If
            End If
        End If
    End If
    
    strRoot = IIf(intDeviceUsed = 1, Nvl(rsTemp!Root1), Nvl(rsTemp!Root2))
    
    'FTP中删除图像文件
    If UBound(Split(strArrImgs, ";")) > 0 Then
        For i = 0 To UBound(Split(strArrImgs, ";"))
            strImageUID = Split(strArrImgs, ";")(i)
            Call iNet.FuncDelFile(strRoot, strImageUID)
            Call iNet.FuncDelFile(strRoot, strImageUID & ".jpg")
        Next
    End If
    
    '关闭FTP连接
    iNet.FuncFtpDisConnect
    DeleteManyImages = True
    
    Exit Function
errH:
    gcnVideoOracle.RollbackTrans
    BUGEX "DeleteManyImages err"
    DeleteManyImages = False
    Call RaiseErr(err)
End Function

Public Sub LockStudy(ByVal intType As Integer, ByVal lngAdviceId As Long, ByVal lngSendNo As Long, _
ByVal lngStudyState As Long, ByVal blnMoved As Boolean)
'锁定 或者 解除锁定
'intType=0:以前的功能，只改变变量值 ; 1: 复杂锁定(未启用同步定位到检查列表);   2：解除锁定     3:简单锁定(已经启用同步定位到检查列表)
    Dim cbrControl As CommandBarControl
    
    Set cbrControl = cbrMain.FindControl(, conMenu_Cap_StudySyncState)
     
    '根据intType判断需要锁定还是解锁，影响VideoCaptureMenuProcess中判断条件
    If intType = 0 Then
        mblnCaptureLockState = True
        Exit Sub
    ElseIf intType = 1 Then
        
        mblnCaptureLockState = False
        Call zlUpdateAdviceInf(lngAdviceId, lngSendNo, lngStudyState, blnMoved)
        mblnCaptureLockState = True
    
        cbrControl.IconId = 10012
    ElseIf intType = 2 Then
        cbrControl.IconId = 8123
    ElseIf intType = 3 Then
        cbrControl.IconId = 10012
    End If
    
    If Not cbrControl Is Nothing Then Call VideoCaptureMenuProcess(cbrControl)
End Sub

Public Function RefreshImageCaptureFace(ByVal blIsDock As Boolean)
'刷新采集界面缩略图高度控件位置，更新控件内部变量
    Call ucPreview.RefreshFace(blIsDock)
    Call ucCacheImage.RefreshFace(blIsDock)
End Function

Public Function SaveImageToStady(dcmImage As DicomImage, lngAdviceId As Long)
    Dim dcmTag As New clsImageTagInf
    
    If mstrInstanceUID = dcmImage.InstanceUID Then Exit Function
    '设置图像的DICOM参数
    subWriteDicomPara dcmImage, lngAdviceId
    
    dcmTag.Tag = imgTag
    
    Set dcmImage.Tag = dcmTag
    mintCaptureFlag = 1
    
    '将图像插入到缩略图中
    subInsert2Mini dcmImage
    
    '保存图像，并触发图像存储事件
    Call SaveImage(dcmImage, lngAdviceId, mcurStudyInf.strStudyUid, 0, True)
    
    mstrInstanceUID = dcmImage.InstanceUID
End Function

Private Sub SaveLocalPar()
'保存本地参数
    SaveSetting "ZLSOFT", G_STR_REG_PATH_PUBLIC, "屏蔽大图", IIf(mblnImageShield, 1, 0)
End Sub

Private Sub GetLocalPar()
'读取本地参数

    mblnImageShield = Val(GetSetting("ZLSOFT", G_STR_REG_PATH_PUBLIC, "屏蔽大图", 0)) = 1
End Sub
