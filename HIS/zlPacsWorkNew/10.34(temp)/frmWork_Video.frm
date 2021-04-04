VERSION 5.00
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "Mscomm32.ocx"
Object = "{853AAF94-E49C-11D0-A303-0040C711066C}#4.3#0"; "DicomObjects.ocx"
Object = "{84926CA3-2941-101C-816F-0E6013114B7F}#1.0#0"; "IMGSCAN.OCX"
Object = "{B1790453-7708-48C1-B5CC-75255FA4B066}#1.0#0"; "ZLDSVideoProcess.ocx"
Begin VB.Form frmWork_Video 
   ClientHeight    =   8880
   ClientLeft      =   120
   ClientTop       =   510
   ClientWidth     =   10425
   ClipControls    =   0   'False
   DrawStyle       =   5  'Transparent
   Icon            =   "frmWork_Video.frx":0000
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   8880
   ScaleWidth      =   10425
   StartUpPosition =   3  '窗口缺省
   Begin ScanLibCtl.ImgScan ImageScanner 
      Left            =   120
      Top             =   3120
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
   Begin VB.Timer tmrReg 
      Interval        =   10000
      Left            =   30
      Top             =   6660
   End
   Begin VB.Timer timerHook 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   0
      Top             =   6090
   End
   Begin VB.PictureBox picDock 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   8295
      Left            =   480
      ScaleHeight     =   8295
      ScaleWidth      =   9015
      TabIndex        =   3
      Top             =   120
      Width           =   9015
      Begin zl9PACSWork.ucSplitter ucSplitter1 
         Height          =   135
         Left            =   0
         TabIndex        =   4
         Top             =   6015
         Width           =   9015
         _ExtentX        =   15901
         _ExtentY        =   238
         MousePointer    =   7
         SplitType       =   0
         DBClickType     =   2
         SplitLevel      =   3
         Con1MinSize     =   3000
         Con2MinSize     =   1000
         Control1Name    =   "picCapture"
         Control2Name    =   "ucPreview"
      End
      Begin VB.PictureBox picCapture 
         ForeColor       =   &H00000000&
         Height          =   6015
         Left            =   0
         ScaleHeight     =   5955
         ScaleWidth      =   8955
         TabIndex        =   6
         Top             =   0
         Width           =   9015
         Begin VB.PictureBox picView 
            BackColor       =   &H8000000D&
            BorderStyle     =   0  'None
            Height          =   3495
            Left            =   600
            ScaleHeight     =   3495
            ScaleWidth      =   6855
            TabIndex        =   11
            Top             =   240
            Width           =   6855
            Begin ZLDSVideoProcess.DSCapture wdmCapture 
               Height          =   3135
               Left            =   720
               TabIndex        =   12
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
            Begin VB.TextBox txtInputText 
               Appearance      =   0  'Flat
               Height          =   300
               Left            =   5520
               TabIndex        =   14
               Text            =   "Text1"
               Top             =   840
               Visible         =   0   'False
               Width           =   975
            End
            Begin VB.PictureBox picVideo 
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               Height          =   3015
               Left            =   1200
               ScaleHeight     =   201
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   224
               TabIndex        =   13
               Top             =   120
               Width           =   3360
            End
            Begin DicomObjects.DicomViewer dcmView 
               Height          =   1575
               Left            =   4440
               TabIndex        =   15
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
         Begin VB.PictureBox pbxSize 
            BackColor       =   &H00800000&
            BorderStyle     =   0  'None
            Height          =   75
            Index           =   0
            Left            =   360
            MousePointer    =   7  'Size N S
            ScaleHeight     =   75
            ScaleWidth      =   7335
            TabIndex        =   10
            Top             =   120
            Width           =   7335
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
            Height          =   3975
            Index           =   2
            Left            =   480
            MousePointer    =   9  'Size W E
            ScaleHeight     =   3975
            ScaleWidth      =   75
            TabIndex        =   8
            Top             =   0
            Width           =   75
         End
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
            TabIndex        =   7
            Top             =   3840
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
      Begin zl9PACSWork.ucImagePreview ucPreview 
         Bindings        =   "frmWork_Video.frx":1CCA
         Height          =   2145
         Left            =   0
         TabIndex        =   5
         Top             =   6150
         Width           =   9015
         _ExtentX        =   15901
         _ExtentY        =   3784
         BackColor       =   4210752
      End
   End
   Begin DicomObjects.DicomViewer dcmAfter 
      Height          =   735
      Left            =   8880
      TabIndex        =   2
      Top             =   2160
      Visible         =   0   'False
      Width           =   1035
      _Version        =   262147
      _ExtentX        =   1826
      _ExtentY        =   1296
      _StockProps     =   35
      BackColor       =   12632319
   End
   Begin VB.PictureBox picBackImg 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1920
      Left            =   7680
      Picture         =   "frmWork_Video.frx":1CDE
      ScaleHeight     =   1920
      ScaleWidth      =   1920
      TabIndex        =   1
      Top             =   5760
      Visible         =   0   'False
      Width           =   1920
   End
   Begin VB.Timer tmrComm 
      Interval        =   2
      Left            =   0
      Top             =   5040
   End
   Begin MSCommLib.MSComm commListener 
      Left            =   0
      Top             =   3840
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin MSComDlg.CommonDialog dlgOpen 
      Left            =   0
      Top             =   4440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox picTemp2 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   1455
      Left            =   45
      ScaleHeight     =   1455
      ScaleWidth      =   375
      TabIndex        =   0
      Top             =   480
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

Implements IWorkMenu


Private Const M_STR_HINT_NoSelectData As String = "无效的检查数据，请选择需要执行的检查记录。"
Private Const M_STR_MODULE_MENU_TAG As String = "采集"


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

'点坐标类型
Private Type TPoint
  X As Integer
  Y As Integer
End Type


Private Type TUnLockStudyInf
    lngAdviceID As Long
    lngSendNO As Long
    blnMoved As Boolean
    lngStudyState As Long
End Type

'''视频窗口事件类型
'Public Enum TVideoEventType
'    vetLockStudy = 1
'    vetAddFirstImg = 2
'    vetDelLastImg = 3
'    vetRecVideo = 4
'    vetUpdateImg = 5
'End Enum

Private mstrActiveType                  '激活方式


Private WithEvents mclsDxDevice As clsDxHidDevice   '实现蓝韵手柄之类的采集方式
Attribute mclsDxDevice.VB_VarHelpID = -1

Public mhCapWnd As Long                 '采集窗口的句柄

Private mlngModul As Long
Private mstrPrivs As String              '模块权限
Private mlngCurDeptId As Long          '当前科室
Private mobjOwner As Object

Public pobjPacsCore As zl9PacsCore.clsViewer
Public mblnObserve As Boolean         '是否有观片基本权限   true是  false否


Private mRestoreContainer As Object
Private mParentContainer As Object
Public mIsShowing As Boolean

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

Private mstrInfor As String

Private mblnMoveDown  As Boolean         '用于判断是否按下鼠标左键
Private mblnDcmViewDown As Boolean      '用于判断dcmView中鼠标是否被按下
Private mintCurImgIndex As Integer      '当前被选中的图象索引
Private mdcmSelectLabel As DicomLabel   '当前被选中的标注
Private mstrAviFileName As String       '录像文件名
Private mstrEncoderName As String
Private mstrBufferDir As String

Private mintCapType As Integer            '脚踏触发方式，0-直接触发，1-变换触发，2-电平触发
Private mintComInterval As Integer       '脚踏采图的时间间隔，单位秒
Private mintComState As Integer          'COM口的状态
Private mlngComTime As Long              '记录com口保持状态的时间
Private mdtLastCapture As Date           '最近脚踏踩下的时间
Private mblnCTSHolding As Boolean        '记录常态时的CTS线的电平
Private mstrComPort As Long              '串口启动的端口号
Private mblnUseClipbord As Boolean          '是否使用剪贴板

Private mobjFtpConnection As New clsFtp
Private mobjBakFtpConnection As New clsFtp

Private mblnUseInetFtp As Boolean

Private mobjFtp As TFtpDeviceInf        'ftp设备信息
Private mobjBakFtp As TFtpDeviceInf     'ftp备份存储设备信息

Private dcmglbUID As New DicomGlobal    '定义UIDRoot=1
Private mblnReadOnly As Boolean         '是否只能查看True查看模式，False采集模式

Private mblnShowProcessBar As Boolean   '是否显示处理工具栏
Private mstrScanDeviceTempDir As String '扫描设备临时目录
Private mblnShowImage As Boolean        '鼠标移动时，是否自动显示大图
Private mdblBigImgZoom As Double        '大图放大倍数
Private mblnUnload As Boolean           '是否允许关闭窗口
Private mblnLocalizerBackward As Boolean    '定位片后置
Private mblnChangeUser As Boolean       '是否启用了用户交互

'病人基本信息资料
Private mlngTmpAdviceId As Long
Private mlngTmpSendNo As Long

Private mlngAdviceID As Long            '医嘱ID
Private mlngSendNo As Long
Private mblnMoved As Boolean            '是否转储
Private mlngStudyState As Long

Private mstrStudyUID As String          '检查UID
Private mstrModality As String          '影像类别
Private mstrSex As String               '性别
Private mstrBirthDate As String         '出生日期
Private mstrAge As String               '年龄
Private mstrName As String              '姓名
Private mstrCheckNo As String           '检查号
Private mstrPatientID As String         '病人ID
Private mstrInstitution As String       '单位名称


Private mstrAfterTag As String          '后台采集标记
Private mstrAfterStudyUid As String     '后台采集检查UID
Private mstrAfterSeriesUid As String    '后台采集序列UID
Private mstrAfterModality As String     '后台采集的影像类别
Private mpanAfterInf As Pane            '后台采集信息显示面板
Private mlngAfterCurImageCount As Long  '当前后台采集图像数量
Private mblnAfterIsUse As Boolean       '是否启用后台采集功能
Private mstrAfterParentTitle As String

Private mSelectStudyInf As TUnLockStudyInf


'modify by tjh at 2010-01-20////////////////////////////////////////////

'Private pCurrentfrmCapture As frmVideoCapture    '记录拥有视频源的采集窗口
Private mVideoCapture As clsVideoCapture '视频采集对象

Private mdblZoomRate As Double  '缩放比率（在cbrMain的cbrMain_ResizeClient事件中需要重新计算该值）
Private mVideoSize As TVideoSize '视频大小（由相关的视频组件保存）
Private mCurCutRange As TCutRange '视频裁剪范围设置（该参数通过GetString和SaveString保存在注册表中）
Private mVideoArea As TVideoArea  '视频客户区域设置（在cbrMain的cbrMain_ResizeClient事件中需要重新计算该值）
Private mVideoDriverType As TVideoDriverType '视频驱动类型（该参数通过GetPara和SetPara保存在数据库中）
Private mblnSoundHint As Boolean    '声音提示
Private mblnPoputWindowHint As Boolean  '弹窗提示

Private Const M_LNG_REFRESHINTERVAL As Long = 600 '刷新间隔

Private mstrVideoRegTime As String '保存视频启动注册时间
Private mblnIsExecuteReg As Boolean '判断是否执行注册过程
Private mblnIsAllowStartupVideo As Boolean '是否允许启动视频源
Private mblnIsLockStudy As Boolean
Public mblnCurCaptureState As Boolean         '保存当前采集状态


Private mObjActiveMenuBar As CommandBars

Private mblnRefreshState As Boolean
Private mblnInitState As Boolean


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

Private Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Private Declare Function MoveFile Lib "kernel32" Alias "MoveFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String) As Long

Property Get CaptionEx() As String
    CaptionEx = Me.Tag
End Property

Property Let CaptionEx(value As String)
    Dim hWndParent As Long
    
    Me.Tag = value
    
    hWndParent = GetParent(Me.hWnd)
    
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

'获取菜单接口对象
Property Get zlMenu() As IWorkMenu
    Set zlMenu = Me
End Property

Public Sub NotificationRefresh()
'通知刷新
    mblnRefreshState = False
End Sub


Private Sub Form_Initialize()
'初始化模块变量
    mblnInitState = False
End Sub


'接口实现部分*********************************************************************************

Public Function IWorkMenu_zlGetModuleMenuId() As Long
'获取影像菜单的菜单ID
    IWorkMenu_zlGetModuleMenuId = conMenu_Cap_Group
End Function


Public Function IWorkMenu_zlIsModuleMenu(ByVal objControlMenu As XtremeCommandBars.ICommandBarControl) As Boolean
'判断菜单是否属于该模块菜单
    IWorkMenu_zlIsModuleMenu = IIf(objControlMenu.Category = M_STR_MODULE_MENU_TAG, True, False)
End Function


Public Sub IWorkMenu_zlCreateMenu(objMenuBar As Object)
'创建影像记录对应的菜单
    Dim cbrMenuBar As CommandBarPopup
    Dim cbrPopControl As CommandBarControl
    Dim cbrControl As CommandBarControl
    
    Dim str3DFuncs() As String
    Dim i As Long
    Dim lng3DFunc As Long
    
    
    Set mObjActiveMenuBar = objMenuBar

    If Not HasMenu(objMenuBar, conMenu_Cap_Group) Then
        Set cbrMenuBar = mObjActiveMenuBar.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_Cap_Group, "采集", 3, False)
        cbrMenuBar.ID = conMenu_Cap_Group
        cbrMenuBar.Category = M_STR_MODULE_MENU_TAG
        
        
        With cbrMenuBar.CommandBar
            Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Cap_Dynamic, "动态", "显示实时视频", 0, True)
            Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Cap_MarkMap, "采集", "采集图像", 0, False)
            
            If mlngModul = G_LNG_VIDEOSTATION_MODULE Then
                Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Cap_After_Capture, "后台采集", "后台采集检查图像", 10020, False)
            End If
            
            Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Cap_Record, "录像", "录制检查视频图像", 0, True)
            
            If mlngModul = G_LNG_VIDEOSTATION_MODULE Then
                Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Cap_After_Record, "后台录像", "后台录制检查视频图像", 10021, False)
            End If
            
            Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Cap_Record_Stop, "停止录像", "停止视频录制", 0, False): cbrControl.Enabled = False
            Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Cap_RecordAudio, "录音", "录音", 0, False)
            
            Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Cap_Play, "播放", "播放录像或者录音", 0, True)
            
            Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Cap_Import, "导入", "文件导入", 10002, True)
            Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Cap_SaveAs, "另存", "文件另存", 3091, False)
            Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Cap_DelImg, "删图", "删除图像", 10001, False)
            
            If mlngModul = G_LNG_VIDEOSTATION_MODULE Then
                Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Cap_OpenStudyList, "打开检查", "打开检查", 0, True)
                Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Cap_StudySyncState, "锁定检查", "锁定检查", 10012, False)
                Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Cap_After_Tag, "标记检查", "标记检查", 10022, False)
            End If
            
            Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Cap_DevSet, "视频设置", "视频设置", 815, True)
            
'            If mlngModul = G_LNG_VIDEOSTATION_MODULE Then
'                Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Manage_ChangeUser, "用户交换", "用户交换", 3012, False)
'            End If
            
            '最右边显示浮动采集按钮
            Set cbrControl = CreateModuleMenu(mObjActiveMenuBar.ActiveMenuBar.Controls, xtpControlButton, comMenu_Cap_Process, "浮动采集", "弹出独立采集窗口", 0, False)
            cbrControl.flags = xtpFlagRightAlign
        End With
    End If
End Sub


Public Sub IWorkMenu_zlCreateToolBar(objToolBar As Object)
'创建工具栏
'    Dim cbrControl As CommandBarControl
'
'    '只有视频采集站点才有用户交换功能
'    If mlngModul = G_LNG_VIDEOSTATION_MODULE Then
'        If HasMenu(objToolBar, conMenu_Manage_ChangeUser) Then Exit Sub
'
'        Set cbrControl = CreateModuleMenu(objToolBar.Controls, xtpControlButton, conMenu_Manage_ChangeUser, "交换", "交换检查医生和报告医生", 3012, True, 4)
'    End If
End Sub


Public Sub IWorkMenu_zlClearMenu()
'清除所创建的菜单
'    Dim cbrControl As CommandBarControl
'
'    Set cbrControl = mObjActiveMenuBar.FindControl(, conMenu_Cap_Group)
'    If Not cbrControl Is Nothing Then Call cbrControl.Delete
End Sub


Public Sub IWorkMenu_zlClearToolBar()
'清除创建的工具栏
    Dim cbrControl As CommandBarControl
    
    Set cbrControl = mObjActiveMenuBar.FindControl(, conMenu_Manage_ChangeUser)
    If Not cbrControl Is Nothing Then Call cbrControl.Delete
End Sub

Public Sub IWorkMenu_zlExecuteMenu(ByVal lngMenuId As Long)
'根据菜单ID执行对应功能
    Dim objCbrControl As XtremeCommandBars.CommandBarControl
    
    Select Case lngMenuId
        Case conMenu_Cap_DevSet     '视频参数设置
            Call Menu_Cap_VideoConfig
            
'        Case conMenu_Manage_ChangeUser
'            Call SendMsgToMainWindow(Me, wetChangeUser, mlngAdviceID)
            
        Case comMenu_Cap_Process
            Call Menu_Manage_浮动采集(True)
            
        Case Else
            Set objCbrControl = Me.cbrMain.FindControl(, lngMenuId)
            
            If Not objCbrControl Is Nothing Then Call zlExecuteCommandBars(objCbrControl)
    End Select
End Sub


Public Sub IWorkMenu_zlUpdateMenu(ByVal control As XtremeCommandBars.ICommandBarControl)
'更新菜单
    Select Case control.ID
        Case conMenu_Cap_DevSet
            control.Enabled = Me.Visible
            
'        Case conMenu_Manage_ChangeUser
'            control.Visible = mblnChangeUser
            
        Case Else
            Call zlUpdateCommandBars(control)
    End Select
End Sub


Public Sub IWorkMenu_zlPopupMenu(objPopup As XtremeCommandBars.ICommandBar)
'配置右键菜单
    Exit Sub
End Sub

Public Sub IWorkMenu_zlRefreshSubMenu(objMenuBar As Object)
'刷新弹出的子菜单
    Exit Sub
End Sub

'*********************************************************************************************


Private Function CreateModuleMenu(objMenuControl As CommandBarControls, _
    ByVal lngType As XTPControlType, ByVal lngID As Long, ByVal strCaption As String, _
    Optional strToolTip As String = "", Optional lngIconId As Long = 0, Optional blnStartGroup As Boolean = False, Optional ByVal lngIndex As Long = -1) As CommandBarControl
'创建该模块内的菜单
    
    If lngIndex >= 0 Then
        Set CreateModuleMenu = objMenuControl.Add(lngType, lngID, strCaption, lngIndex)
    Else
        Set CreateModuleMenu = objMenuControl.Add(lngType, lngID, strCaption)
    End If
    
    CreateModuleMenu.ID = lngID '如果这里不指定id，则不能将有些菜单添加到右键菜单中
    
    If lngIconId <> 0 Then CreateModuleMenu.IconId = lngIconId
    If blnStartGroup Then CreateModuleMenu.BeginGroup = True
    If strToolTip <> "" Then CreateModuleMenu.ToolTipText = strToolTip
    
    CreateModuleMenu.Category = M_STR_MODULE_MENU_TAG
End Function


Private Sub Menu_Cap_VideoConfig()
On Error GoTo errHandle
    frmVideoSetup.mlngModul = mlngModul
    frmVideoSetup.strRegName = "frmVideoCapture"    '设置注册表的节点名称
    frmVideoSetup.mstrPrivs = mstrPrivs
    frmVideoSetup.mlngCurDepartId = mlngCurDeptId
         
    Set frmVideoSetup.frmParent = Me
          
    'modify by tjh at 2010-01-20
    'frmVideoSetup.Show 1, Me
    
    Call frmWork_Video.SaveParameterCfg
    Call frmVideoSetup.ShowParameterConfig(frmWork_Video.videoCapture, Me)
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub Menu_Manage_浮动采集(Optional blnUnload As Boolean = True)
On Error GoTo errHandle

    If Not GetIsValidOfStorageDevice(mlngCurDeptId) Then
      MsgBoxD Me, "影像存储设备未定义或处于停用，请检查！", vbInformation, gstrSysName
      Exit Sub
    End If
    
    'Call frmVideoCapture.SetRestoreContainer(picVideoContainer)
    Call frmVideoDockWindow.Show
    
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub InitParameter()
'初始化参数设置
    Dim rsTmp As New ADODB.Recordset
    Dim intVideoCapture As Integer
    Dim strRegPath As String        '注册表参数的保存路径

    mblnRealTime = True
    mintCurImgIndex = 0
    mblnPlayVideo = False
    mstrVideoRegTime = ""
    
    mblnAfterIsUse = False
    mstrAfterModality = "OT"
    mstrAfterParentTitle = ""
        
    '如果程序在磁盘的根目录则app.path为“x:\”
    mstrBufferDir = IIf(Len(App.Path) > 3, App.Path & "\TmpImage\", App.Path & "TmpImage\")
    mstrAviFileName = mstrBufferDir & "TmpVideo.avi"
    
    mblnUnload = False
    mblnIsExecuteReg = False
        
    
    mstrInstitution = GetSetting("ZLSOFT", "注册信息", "单位名称", "")
    
    gint视频设备数量 = getLicenseCount(LOGIN_TYPE_视频设备)
    '读取注册表信息--界面布局
    strRegPath = "公共模块\" & App.ProductName & "\frmVideoCapture"
    
    mblnUseClipbord = GetSetting("ZLSOFT", strRegPath, "UseClipbord", 0)
    Call SaveSetting("ZLSOFT", strRegPath, "UseClipbord", IIf(mblnUseClipbord, 1, 0))
    
    '读取驱动类型
    mVideoDriverType = zlDatabase.GetPara("视频驱动类型", glngSys, mlngModul, "0")
    
    '读取提示类型
    mblnSoundHint = zlDatabase.GetPara("采集后声音提示", glngSys, mlngModul, True)
    mblnPoputWindowHint = zlDatabase.GetPara("采集后弹窗提示", glngSys, mlngModul, True)
    
    '读取扫描设备临时存储的图像目录
    mstrScanDeviceTempDir = GetSetting("ZLSOFT", strRegPath, "扫描设备临时目录", "C:\Documents and Settings\All Users\Application Data\Microsoft\WIA")
     
     
    '读取裁剪比率
    mCurCutRange.LeftRate = Val(GetSetting("ZLSOFT", strRegPath, "mdblX1Scale", 0))  '使用mdblX1Scale名称是为了保证和以前的参数设置兼容
    mCurCutRange.WidthRate = Val(GetSetting("ZLSOFT", strRegPath, "mdblX2Scale", 0))
    mCurCutRange.TopRate = Val(GetSetting("ZLSOFT", strRegPath, "mdblY1Scale", 0))
    mCurCutRange.HeightRate = Val(GetSetting("ZLSOFT", strRegPath, "mdblY2Scale", 0))

    If (mCurCutRange.LeftRate >= 1) Or (mCurCutRange.LeftRate < 0) Then mCurCutRange.LeftRate = 0
    If (mCurCutRange.WidthRate >= 1) Or (mCurCutRange.WidthRate < 0) Then mCurCutRange.WidthRate = 0
    If (mCurCutRange.TopRate >= 1) Or (mCurCutRange.TopRate < 0) Then mCurCutRange.TopRate = 0
    If (mCurCutRange.HeightRate >= 1) Or (mCurCutRange.HeightRate < 0) Then mCurCutRange.HeightRate = 0
  
    
    '读取串口的参数
    On Error GoTo continue1:
    mstrActiveType = zlDatabase.GetPara("脚踏端口", glngSys, mlngModul, "1")
    If IsNumeric(mstrActiveType) Then
        mstrComPort = CLng(mstrActiveType)
        mstrActiveType = "COM"
        
        mintCapType = zlDatabase.GetPara("脚踏采集方式", glngSys, mlngModul, "1")
        If mintCapType < 0 Or mintCapType > 2 Then
            mintCapType = 1
        End If
        '读取脚踏间隔时间
        mintComInterval = zlDatabase.GetPara("脚踏时间间隔", glngSys, mlngModul, "1")
    End If
continue1:

    
    '鼠标移动时，是否自动显示大图
     mblnShowImage = zlDatabase.GetPara("鼠标移动时显示大图", glngSys, mlngModul, "0")
     mdblBigImgZoom = zlDatabase.GetPara("采集大图放大倍数", glngSys, mlngModul, "1")
     
     If mblnShowImage Then ucPreview.MouseMoveZoom = mdblBigImgZoom
     
     
    '定义UIDRoot=1
    dcmglbUID.RegString("UIDRoot") = "1"
    
    '设置视频采集区域大小是否允许修改
    intVideoCapture = Val(zlDatabase.GetPara("允许改变采集区域大小", glngSys, mlngModul, "1", , InStr(mstrPrivs, ";参数设置;") > 0))
    
    If intVideoCapture = 0 Then
    
        pbxSize.Item(0).MousePointer = 0
        pbxSize.Item(1).MousePointer = 0
        pbxSize.Item(2).MousePointer = 0
        pbxSize.Item(3).MousePointer = 0
    Else
    
        pbxSize.Item(0).MousePointer = 7
        pbxSize.Item(1).MousePointer = 7
        pbxSize.Item(2).MousePointer = 9
        pbxSize.Item(3).MousePointer = 9
    
    End If
    
    
    '初始化科室级参数==============================================================================
    mblnAfterIsUse = GetDeptPara(mlngCurDeptId, "启用后台采集", 0)
    mstrAfterModality = GetDeptPara(mlngCurDeptId, "后台影像类别", "OT")
    
    '读取并检测存储设备号
    mobjFtp.strDeviceId = GetDeptPara(mlngCurDeptId, "存储设备号")
    mobjBakFtp.strDeviceId = GetDeptPara(mlngCurDeptId, "备份设备号")
    
    mblnLocalizerBackward = Val(GetDeptPara(mlngCurDeptId, "定位片后置", 0))
    
'    mblnChangeUser = GetDeptPara(mlngCurDeptId, "允许交换用户", 0) = "1"              '允许交换用户
    
    '获取在线存储设备信息
    gstrSQL = "Select 设备号,设备名 From 影像设备目录 Where 类型=1 and 设备号=[1] and NVL(状态,0)=1"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Tag, mobjFtp.strDeviceId)
    
    If rsTmp.EOF Then
        MsgBox "影像存储设备未定义或处于停用，请检查！", vbInformation, gstrSysName
        mobjFtp.strDeviceId = ""
        mblnReadOnly = True
        Exit Sub
    End If
    
    Call funGetFtpDeviceInf(Me, mobjFtp)
    
    '获取备份设备信息
    If Val(mobjBakFtp.strDeviceId) > 0 Then
        gstrSQL = "Select 设备号,设备名 From 影像设备目录 Where 类型=1 and 设备号=[1] and NVL(状态,0)=1"
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Tag, mobjBakFtp.strDeviceId)
        
        If rsTmp.EOF Then
            mobjBakFtp.strDeviceId = ""
            MsgBox "未取得有效的备份设备信息，不能对采集图像进行备份操作，请检查备份设备配置是否正确。", vbInformation, gstrSysName
            
            Exit Sub
        End If
        
        Call funGetFtpDeviceInf(Me, mobjBakFtp)
    End If
    
    
End Sub

Public Sub zlInitModule(ByVal lngModule As Long, ByVal strPrivs As String, ByVal lngDepartId As Long, Optional owner As Object = Nothing)
'初始化模块参数
    mlngModul = lngModule
    mstrPrivs = strPrivs
    mlngCurDeptId = lngDepartId
    
    If Not owner Is Nothing Then Set mobjOwner = owner
    
    '设置窗口样式
    Call SetWindowStyle
    
    Call InitParameter
    
    Call OpenVideoCaptureDevice
    
    Call UpdateAfterCaptureInfo
    
    mblnInitState = True
End Sub




'显示视频窗口
Public Sub ShowVideoWindow(ByRef objContainer As Object)
    Dim strRegPath As String
    
    If Not mIsShowing Then
        Call Me.Show
        
        mIsShowing = True
    End If
    
    If objContainer Is Nothing Then Exit Sub
    
    If Not mParentContainer Is Nothing Then
        Call SaveVideoAreaCfg(mParentContainer.Name)
    End If
    
    Set mParentContainer = objContainer
    Call SetParent(Me.hWnd, mParentContainer.hWnd)

    If Me.Height <> mParentContainer.Height Then
        Call LoadVideoAreaCfg(mParentContainer.Name)
    End If

    Call UpdateSize
    
    If TypeOf mParentContainer Is Form Then
        mParentContainer.Caption = Me.Tag
        mParentContainer.Icon = Me.Icon
        
        Me.Width = Me.Width - 140
        Me.Height = Me.Height - 140
    End If
End Sub


Public Sub HideVideoWindow()
'隐藏视频显示窗口
    Me.Hide
    
    mIsShowing = False
End Sub


'更新当前视频窗口大小
Public Sub UpdateSize()
On Error GoTo errHandle
    If mParentContainer Is Nothing Then Exit Sub
    
    Me.Left = 0
    Me.Top = 0
    
    Me.Height = mParentContainer.Height
    Me.Width = mParentContainer.Width
    
    If TypeOf mParentContainer Is Form Then
        Me.Width = Me.Width - 140
        Me.Height = Me.Height - 500
    End If
errHandle:
End Sub


'设置恢复时的容器对象
Public Sub SetRestoreContainer(ByRef objContainer As Object)
    Set mRestoreContainer = objContainer
    
'    mRestoreContainer.Visible = True
End Sub


'恢复原有的视频显示容器
Public Sub RestoreContainer()
    If mRestoreContainer Is Nothing Then Exit Sub
    
    If Not mParentContainer Is Nothing Then
        Call SaveVideoAreaCfg(mParentContainer.Name)
    End If
    
    Set mParentContainer = mRestoreContainer
    Call SetParent(Me.hWnd, mRestoreContainer.hWnd)
    
    Me.Left = 0
    Me.Top = 0

    If Me.Height <> mRestoreContainer.Height Then
        '当从浮动窗口或其他窗口恢复视频显示位置时，重新从注册表读取视频显示位置大小
        Call LoadVideoAreaCfg(mRestoreContainer.Name)
    End If
    
    Me.Height = mRestoreContainer.Height
    Me.Width = mRestoreContainer.Width
    
    
    If TypeOf mRestoreContainer Is Form Then
        mRestoreContainer.Caption = Me.Tag
        mRestoreContainer.Icon = Me.Icon
    End If
End Sub


Property Get ParentContainerObj() As Object
    Set ParentContainerObj = mParentContainer
End Property

Property Set ParentContainerObj(value As Object)
    Set mParentContainer = value
End Property



Property Get RestoreContainerObj() As Object
    Set RestoreContainerObj = mRestoreContainer
End Property

Property Set RestoreContainerObj(value As Object)
    Set mRestoreContainer = value
End Property



Property Get IsLockStudy() As Boolean
    IsLockStudy = mblnIsLockStudy
End Property



Property Get LockPatientName() As String
    LockPatientName = mstrInfor
End Property



'----------------------------------------------------------------------------------------------------------
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'检测是否运行启动视频源
'当视频源没有正常启动时，则不进行注册，也不进行判断
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function CheckVideoReg(ByVal blnIsStartupVideo As Boolean) As Boolean
  '不论视频启动成功，都需要进行注册
  
    mblnIsExecuteReg = True
  
    mstrVideoRegTime = FunLogIn(Me, LOGIN_TYPE_视频设备)
  
    CheckVideoReg = mstrVideoRegTime <> ""
End Function


Public Sub zlUpdateAdviceInf(ByVal lngAdviceID As Long, ByVal lngSendNO As Long, ByVal lngStudyState As Long, ByVal blnMoved As Boolean)
'更新医嘱信息
    Dim rsTemp As ADODB.Recordset
    
    '保存主界面的当前检查信息
    mSelectStudyInf.lngAdviceID = lngAdviceID
    mSelectStudyInf.blnMoved = blnMoved
    mSelectStudyInf.lngSendNO = lngSendNO
    mSelectStudyInf.lngStudyState = lngStudyState
    
    If mblnIsLockStudy Then Exit Sub
    
    mlngAdviceID = lngAdviceID
    mlngSendNo = lngSendNO
    mblnMoved = blnMoved
    mlngStudyState = lngStudyState
    mblnReadOnly = False
    mblnRefreshState = True
    
    '数据被转移时，没有权限时，状态为指定状态时，该模块为只读
    If mlngAdviceID <= 0 Or blnMoved Or lngStudyState = 6 Or lngStudyState = 0 Or lngStudyState = 1 Or InStr(mstrPrivs, "视频采集") <= 0 Then
        mblnReadOnly = True
    End If
    
    '提取病人基本信息,写DICOM参数时用
    gstrSQL = "Select /*+Rule */ A.影像类别,A.姓名,A.性别,A.年龄,A.出生日期,A.姓名,A.检查号,A.检查UID,B.病人ID " & _
                " From 影像检查记录 A,病人医嘱记录　B " & _
                " Where A.医嘱ID=[1] And A.医嘱ID=B.Id"
                
    If mblnMoved Then
        gstrSQL = Replace(gstrSQL, "影像检查记录", "H影像检查记录")
        gstrSQL = Replace(gstrSQL, "病人医嘱记录", "H病人医嘱记录")
    End If
    
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取病人基本信息", lngAdviceID)
    
    If Not rsTemp.EOF Then
        mstrStudyUID = Nvl(rsTemp("检查UID"))
        mstrModality = Nvl(rsTemp("影像类别"))
        mstrInfor = Nvl(rsTemp("姓名"))
        mstrSex = Nvl(rsTemp("性别"))
        mstrAge = Nvl(rsTemp("年龄"))
        mstrBirthDate = Nvl(rsTemp("出生日期"))
        mstrName = Nvl(rsTemp("姓名"))
        mstrCheckNo = Nvl(rsTemp("检查号"))
        mstrPatientID = Nvl(rsTemp("病人ID"))
        
        If mstrSex = "男" Then
            mstrSex = "M"
        ElseIf mstrSex = "女" Then
            mstrSex = "F"
        Else
            mstrSex = "O"
        End If
    Else
        mstrStudyUID = ""
        mstrModality = ""
        mstrInfor = ""
        mstrSex = ""
        mstrAge = ""
        mstrCheckNo = ""
        mstrPatientID = ""
        mstrBirthDate = ""
        mstrName = ""
    End If
    
    Me.Tag = "图像采集" & IIf(mstrInfor <> "", "(" & mstrInfor & ")", "")
    Me.CaptionEx = Me.Tag
End Sub


Private Sub LockStudy()
'锁定检查
    mblnIsLockStudy = True
End Sub


Private Sub UnLockStudy()
'解锁检查
    mblnIsLockStudy = False
End Sub


Public Sub zlRefreshFace(Optional blnForceRefresh As Boolean = False)
'刷新界面
    Dim rsTemp As ADODB.Recordset
    Dim iRows As Integer
    Dim iCols As Integer
    Dim strStudyUID As String
    
    On Error GoTo errHandle
    
    
    If (mlngTmpAdviceId = mlngAdviceID And mlngTmpSendNo = mlngSendNo And mblnRefreshState) And Not blnForceRefresh Then Exit Sub
    
    mlngTmpAdviceId = mlngAdviceID
    mlngTmpSendNo = mlngSendNo
    mblnRefreshState = True

    Call ConfigVideoShowState(True)
    
'    '提取病人基本信息,写DICOM参数时用
'    gstrSQL = "Select A.检查UID From 影像检查记录 A,病人医嘱记录　B  Where A.医嘱ID=B.Id and A.医嘱ID=[1] "
'
'    If mblnMoved Then
'        gstrSQL = Replace(gstrSQL, "影像检查记录", "H影像检查记录")
'        gstrSQL = Replace(gstrSQL, "病人医嘱记录", "H病人医嘱记录")
'
'        ucPreview.Enable = False
'    End If
'
'    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取病人基本信息", mlngAdviceID)
'
'    If rsTemp.RecordCount <= 0 Then
'        strStudyUID = ""
'    Else
'        strStudyUID = Nvl(rsTemp!检查uid)
'    End If


    Call ucPreview.RefreshImage(slStudy, mstrStudyUID, mblnMoved, blnForceRefresh, False)
    
    If ucPreview.ImgViewer.Images.Count > 0 Then
                    
        '将被选中图像装载到dcmView中
        dcmView.Images.Clear
        dcmView.Images.Add ucPreview.ImgViewer.Images(ucPreview.SelectIndex)
        
        Dim dblTempZoom As Double
              
        dblTempZoom = dcmView.CurrentImage.ActualZoom
        dcmView.CurrentImage.StretchToFit = False
        
        
        '判断当进入浮动窗口时，缩放比率不能小于0.1
        If dblTempZoom < 0.1 Then
            dblTempZoom = 0.1
        End If
        
              
        Call subCenterZoom(dcmView.CurrentImage, dcmView, dblTempZoom, mCorpSize)
        
        '如果是Twain采集模式，则设置mblnRealTime为false
        If IsTwainCaptureWay = True Then mblnRealTime = False

        '显示dcmView，隐藏picVideo
        dcmView.CurrentImage.BorderWidth = 0
    Else
        Call dcmView.Images.Clear
    End If
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub




Private Sub zlStopCapture()
'-----------------------------------------------------------------------------------------
'功能：停止显示视频采集,释放视频采集窗口，
'      释放串口侦听的端口
'参数：无
'返回：无
'-----------------------------------------------------------------------------------------
    '释放采集设备及窗体
    If Not mVideoCapture Is Nothing Then Call mVideoCapture.StopPreview
    
    '关闭COMM口
    If commListener.PortOpen Then
        commListener.PortOpen = False
    End If
    
    '采用Midi接口需在消毁事件句柄
    If Not mclsDxDevice Is Nothing Then
        If mclsDxDevice.Handle <> 0 Then Call mclsDxDevice.CloseDxDevice
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
    
    Select Case control.ID
        Case conMenu_Cap_Dynamic       '动态显示
            control.Checked = mblnRealTime
            control.Enabled = (Not mblnReadOnly) And Not IsTwainCaptureWay And mVideoCapture.IsStartup ' And (mhCapWnd <> 0) modify by tjh at 2010-01-20
            control.Visible = Not IsTwainCaptureWay
            
            If mblnRealTime Then
                control.IconId = conMenu_Cap_Dynamic
            Else
                control.IconId = 10023
            End If
            
        Case conMenu_Cap_MarkMap       '影像采集
            control.Enabled = Not mblnReadOnly And (mVideoCapture.IsStartup Or IsTwainCaptureWay) And Not mblnCurCaptureState
            
        Case conMenu_Cap_After_Capture  '后台采集
            control.Enabled = mVideoCapture.IsStartup And Not mblnCurCaptureState
            control.Visible = mblnAfterIsUse And mlngModul = G_LNG_VIDEOSTATION_MODULE
            
        Case conMenu_Cap_Import        '影像导入
            control.Enabled = Not mblnReadOnly
            
        Case conMenu_Cap_DelImg  '影像删除
            control.Enabled = (mblnRealTime = False) And (ucPreview.ImgViewer.Images.Count > 0) And (Not mblnReadOnly) And Me.Visible
            
        Case conMenu_Cap_Record        '录像
            control.Enabled = Not mblnReadOnly And mVideoDriverType = vdtWDM And mVideoCapture.IsStartup
            control.Visible = Not IsTwainCaptureWay
            
        Case conMenu_Cap_After_Record   '后台录像
            control.Enabled = mVideoDriverType = vdtWDM And mVideoCapture.IsStartup
            control.Visible = Not IsTwainCaptureWay And mblnAfterIsUse And False
            
        Case conMenu_Cap_Record_Stop '停止录像 modify by tjh at 2010-01-22
            control.Enabled = mblnRealTime And Not mblnReadOnly And (mVideoDriverType = vdtWDM) And mVideoCapture.IsStartup
            control.Visible = Not IsTwainCaptureWay
            
        Case conMenu_Cap_RecordAudio '录音
            control.Enabled = Not mblnReadOnly
            
'        Case conMenu_Cap_Full_Screen '全屏 modify by tjh at 2010-01-22 (如果使用新的视频回放组件，则可以启用该功能)
'            control.Enabled = mblnRealTime And (Not mblnReadOnly) And Not GetIsTwainCaptureWay And mVideoCapture.IsStartup
'            control.Visible = Not GetIsTwainCaptureWay And mstrVideoRegTime <> ""
'
        Case conMenu_Cap_DevSet        '设置（如果处于浮动状态时，则屏蔽该按钮） modify by tjh at 2010-01-25
            control.Enabled = mblnIsAllowStartupVideo   'mblnEmbedded ' And (Not mblnReadOnly)
            
            '如果为浮动窗体，则隐藏该设置按钮
            'control.Visible = mstrVideoRegTime <> ""
            If Not (mParentContainer Is Nothing) Then
                If TypeOf mParentContainer Is frmVideoDockWindow Then
                    control.Enabled = False
                Else
                    control.Enabled = True
                End If
            End If
            
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
            
        Case conMenu_Tool_Analyse
            If mblnObserve Then
                control.Enabled = Not mblnReadOnly
            Else
                control.Visible = False
                control.Enabled = False
            End If
            
            
        Case conMenu_Cap_OpenStudyList
            control.Enabled = True
            control.Visible = IIf(mlngModul = G_LNG_VIDEOSTATION_MODULE, True, False)
            
        Case conMenu_Cap_StudySyncState
            control.Enabled = Not mblnReadOnly Or mblnIsLockStudy
            control.Visible = IIf(mlngModul = G_LNG_VIDEOSTATION_MODULE, True, False)
            
        Case conMenu_Cap_After_Tag
            control.Enabled = mVideoCapture.IsStartup
            control.Visible = mblnAfterIsUse And mlngModul = G_LNG_VIDEOSTATION_MODULE
    End Select
End Sub


''''''''''''''''''''''''''''''''''
'扫描图像
''''''''''''''''''''''''''''''''''
Private Sub ScanImages()
  '注册失败则不执行该功能
  If mstrVideoRegTime = "" Then
    Exit Sub
  End If
                
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
    Call ImageScanner.CloseScanner

    MsgBox err.Description
End Sub


Public Sub CaptureImage()
'************************************************************
'
'从视频或者录像中采集图像
'
'************************************************************
    Dim strSql As String
    Dim rsData As ADODB.Recordset
    
    If mstrVideoRegTime = "" Then   '如果没有注册，则不允许采集
        MsgboxEx Me, "未检测到有效的注册信息，不能进行图像采集操作！", vbOKOnly, "提示"
        Exit Sub
    End If
    
    If Not (Not mblnReadOnly And (mVideoCapture.IsStartup Or IsTwainCaptureWay)) Then Exit Sub  '如果为只读，或者视频没有启动，则不允许采集
    
    '采集图像时，如果不是后台采集，则需判断当前加载的图像与数据库中的图像记录数是否一致，如果不一致，说明该检查当前可能正被其他设备站点采集
    strSql = "select count(*) as 图像数 from 影像检查图象 where 序列uid in(select 序列UID from 影像检查序列 where 检查UID=(select 检查UID from 影像检查记录 where 医嘱id=[1])) "
    Set rsData = zlDatabase.OpenSQLRecord(strSql, "查询图像数量", mlngAdviceID)
    
    If rsData.RecordCount > 0 Then
        If Val(Nvl(rsData!图像数)) <> ucPreview.ImageTotal Then
            Call MsgBoxD(Me, "检测到当前加载的图像数量与数据库记录数不一致，如果无其他用户对该检查进行采集，则请在刷新后重试。", vbInformation + vbOKOnly, "提示")
            Exit Sub
        End If
    End If
            
    If IsTwainCaptureWay Then
      Call ScanImages  '通过TWAIN接口采集图像
    Else
        If mblnRealTime Then '为实时显示时自动采实时图
            Call subCaptureImg(True)
        Else
            Call subCaptureImg(MsgBoxD(Me, "确定要采集当前静态图吗？选“否”则采集设备实时图像。", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo)
        End If
    End If
End Sub



Public Sub CaptureAfterImage()
'后台图像采集
    If mstrVideoRegTime = "" Then   '如果没有注册，则不允许采集
        MsgboxEx Me, "未检测到有效的注册信息，不能进行图像采集操作！", vbOKOnly, "提示"
        Exit Sub
    End If
    
    If Not mVideoCapture.IsStartup Then Exit Sub  '如果为只读，或者视频没有启动，则不允许采集,twain方式不允许后台采集
    
    Call subCaptureImg(True, "", Nothing, True)
    
End Sub


Public Sub zlExecuteCommandBars(control As XtremeCommandBars.CommandBarControl)
  On Error GoTo errHandle
    Select Case control.ID
        Case conMenu_Cap_Dynamic       '动态显示
            If IsTwainCaptureWay Then
              Call MsgBoxD(Me, "TWAIN采集模式下，不能进行动态视频的显示。", vbOKOnly, "提示")
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
                MsgboxEx Me, "未检测到有效的注册信息，不能进行录像操作！", vbOKOnly, "提示"
                Exit Sub
            End If
            
            Call subVideoSave
            
        Case conMenu_Cap_Record_Stop  '停止录像 modify by tjh at 2010-01-22
            If mstrVideoRegTime = "" Then
                'MsgboxEx Me, "未检测到有效的注册信息，不能进行录像操作！", vbOKOnly, "提示"
                Exit Sub
            End If
            
            Call subStopVideo
            
        Case conMenu_Cap_RecordAudio    '录音
            If mstrVideoRegTime = "" Then
                MsgboxEx Me, "未检测到有效的注册信息，不能进行录音操作！", vbOKOnly, "提示"
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
            subSetRotate True
            
        Case conMenu_Process_LRotate        '逆时针旋转
            subSetRotate False
            
        Case conMenu_Process_Sharpness      '锐化
            subSetSharp True
            
        Case conMenu_Process_Filter         '平滑
            subSetSharp False
            
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
        Case conMenu_Tool_Analyse           '高级处理
            Call OpenViewer(1, pobjPacsCore, mlngAdviceID, False, Me, "", mblnMoved, mblnLocalizerBackward)
            
        Case conMenu_Cap_OpenStudyList      '打开检查采集图像
            Call OpenStudy
            
        Case conMenu_Cap_StudySyncState     '锁定检查
            If control.IconId = 10012 Then
                control.IconId = 8123
                
                Call LockStudy
                
                Call SendMsgToMainWindow(Me, wetLockStudy, mlngAdviceID, mstrInfor)
            Else
                control.IconId = 10012
                
                Call UnLockStudy
                
                If mlngAdviceID <> mSelectStudyInf.lngAdviceID Then
                    Call zlUpdateAdviceInf(mSelectStudyInf.lngAdviceID, mSelectStudyInf.lngSendNO, mSelectStudyInf.lngStudyState, mSelectStudyInf.blnMoved)
                    Call zlRefreshFace
                End If
                
                Call SendMsgToMainWindow(Me, wetUnLockStudy, mlngAdviceID, mstrInfor)
            End If
        Case conMenu_Cap_After_Tag      '更新后台采集标记
            If mstrVideoRegTime = "" Then
                MsgboxEx Me, "未检测到有效的注册信息，不能进行标记！", vbOKOnly, "提示"
                Exit Sub
            End If
            
            Call UpdateAfterCaptureInfo
    End Select
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume Next
End Sub


Private Sub OpenStudy()
    Dim cbrControl As CommandBarControl
    
    Dim lngCurAdviceId As Long
    Dim lngSendNO As Long
    Dim lngStudyState As Long
    Dim blnMoved As Boolean
    Dim blnResult As Boolean
    
    blnResult = mobjOwner.OpenPatiListWind(lngCurAdviceId, lngSendNO, lngStudyState, blnMoved)
        
    If lngCurAdviceId > 0 Then
        '开始打开新的检查进行采集
        Call UnLockStudy
        
        Call zlUpdateAdviceInf(lngCurAdviceId, lngSendNO, lngStudyState, blnMoved)
        Call zlRefreshFace
        
        Call LockStudy
                
        '修改锁定状态
        Set cbrControl = cbrMain.FindControl(, conMenu_Cap_StudySyncState)
        cbrControl.IconId = 8123
       
        '触发病人改变事件
        Call SendMsgToMainWindow(Me, wetLockStudy, mlngAdviceID, mstrInfor)
    End If
End Sub


Public Sub zlUnloadMe()
    mblnUnload = True
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
            cbrMain.Item(2).Position = xtpBarTop
            cbrMain.Item(3).Position = xtpBarBottom
        Else
            cbrMain.Item(2).Position = xtpBarLeft
            cbrMain.Item(3).Position = xtpBarRight
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
        Call subCenterZoom(dcmView.Images(1), dcmView, dcmView.Images(1).ActualZoom, mCorpSize)
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
        If dcmView.Images.Count > 0 Then
            Call subCenterZoom(dcmView.Images(1), dcmView, dcmView.Images(1).ActualZoom, mCorpSize)
        End If
    
    End If
End Sub


Private Sub cbrMain_Update(ByVal control As XtremeCommandBars.ICommandBarControl)
    zlUpdateCommandBars control
End Sub


Private Sub commListener_OnComm()
On Error GoTo errHandle
    Dim strInput As String
    
    '如果是TWAIN扫描，则不支持脚踏开关采集
    If IsTwainCaptureWay Then Exit Sub
    
    If mstrActiveType <> "COM" Then Exit Sub
    
    strInput = ""
    If commListener.InBufferCount > 0 Then strInput = commListener.Input
    
    If Not (commListener.CommEvent = comEvCTS Or commListener.CommEvent = comEvDSR _
        Or commListener.CommEvent = comEvCD Or commListener.CommEvent = comEvRing Or strInput <> "" _
        Or commListener.CommEvent = comEvSend Or commListener.CommEvent = comEvReceive) Then Exit Sub
    
    If mintCapType = 1 Then '转换触发
        If mintComState <> commListener.CommEvent Then
           '如果累计时间超过了采图时间间隔，则采集图像
           If mlngComTime > mintComInterval Then
               'If Me.cbrMain.FindControl(, conMenu_Cap_MarkMap).Enabled Then
               If Not mblnReadOnly Then
                    Call subCaptureImg(True)
               End If
           End If
           
           '记录新的COM状态，计时器清零，启动timer
           mintComState = commListener.CommEvent
           mlngComTime = 0
           tmrComm.Enabled = True
        End If
    ElseIf mintCapType = 0 Then   '直接触发
        '两次踩下脚踏的时间间隔不能少于3秒
        If DateDiff("S", mdtLastCapture, time) < mintComInterval Then
            mdtLastCapture = time
            Exit Sub
        End If
        
        mdtLastCapture = time
        
        If Not mblnReadOnly Then
            Call subCaptureImg(True)
        End If
    Else    '电平触发
        '对于电平触发的情况，当踩下脚踏的时候，对应线的电平会出现（低-高-低）或（高-低-高）的变化
        '通过电平变化，可以确定是否踩了脚踏。
        '当出现电流干扰时，虽然会出现OnComm事件，但是电平不会发生变化。
        '通过判断当前电平跟常态电平是否相同来确定电平是否发生了变化。
        
        '判断电平是否改变，判断CTS线
        If mblnCTSHolding <> commListener.CTSHolding Then
            '过滤振荡，毛刺现象，判断两次触发的时间是否小于设定的间隔
            If DateDiff("S", mdtLastCapture, time) < mintComInterval Then
                mdtLastCapture = time
                Exit Sub
            End If
            
            mdtLastCapture = time
            
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
    If ErrCenter() = 1 Then Resume
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
  
  If mVideoSize.Width = 0 Or mVideoSize.Height = 0 Then
    Exit Sub
  End If
  
  If (mVideoArea.Height <= 0) Or (mVideoArea.Width <= 0) Then
    Exit Sub
  End If
  
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


Private Sub SetWindowStyle()
    Dim lngWindowStyle As Long
    
    lngWindowStyle = GetWindowLong(Me.hWnd, GWL_STYLE)
    lngWindowStyle = lngWindowStyle And Not (WS_SYSMENU Or WS_CAPTION Or WS_MAXIMIZEBOX Or WS_MINIMIZEBOX Or WS_THICKFRAME)
    
    Call SetWindowLong(Me.hWnd, GWL_STYLE, lngWindowStyle Or WS_CHILD)
End Sub

Private Sub OpenVideoCaptureDevice()
'打开视频采集设备
    Dim blnIsStartupVideo As Boolean

    If mVideoCapture Is Nothing Then
        '创建视频采集对象
        Set mVideoCapture = New clsVideoCapture
        
        '连接视频相关组件
        Call mVideoCapture.ConnectedVfwDeviceObj(picVideo)
        Call mVideoCapture.ConnectedWdmDeviceObj(wdmCapture)
        
        '读取配置文件
        Call mVideoCapture.ReadCaptureParameterFromFile(App.Path & "\" & CAPTURE_PARAMETER_CONFIG_FILE_NAME)
    
        '设置视频的显示模式
        Call mVideoCapture.SetVideoShowWay(swStretch)
    
        '在读取文件配置后修改该属性（只有设置该属性，才能根据四条边框进行调节和显示）
        wdmCapture.AppHandle = Me.hWnd
        wdmCapture.IsShowState = False
        
        mdblZoomRate = 1
    Else
        Call zlStopCapture
    End If

    
    '设置视频驱动类型
    mVideoCapture.VideoDriverType = mVideoDriverType

    '读取视频大小
    mVideoSize.Width = ScaleX(mVideoCapture.VideoSize.Width, vbPixels, vbTwips)
    mVideoSize.Height = ScaleX(mVideoCapture.VideoSize.Height, vbPixels, vbTwips)
    
    '配置界面
    Call CaptureSwitchFace(IsTwainCaptureWay)
    
    mblnIsAllowStartupVideo = FunCheckRegInfo(Me)
    
    '判断是否允许启动视频源********************************
    If Not mblnIsAllowStartupVideo Then
      mVideoCapture.IsAllowStartupVideo = False
      
      '当不运行启动时，进入twain的操作界面
      mVideoDriverType = vdtTWAIN
      mVideoCapture.VideoDriverType = vdtTWAIN
      '配置界面
      Call CaptureSwitchFace(IsTwainCaptureWay)
      
      Exit Sub
    End If
    '*******************************************************
    
    
    '开始视频预览********************************************
    If Not IsTwainCaptureWay Then
        mblnRealTime = True
        
        Call mVideoCapture.StartPreview
                
        blnIsStartupVideo = mVideoCapture.IsStartup
    Else
        mblnRealTime = False
        
        blnIsStartupVideo = ImageScanner.ScannerAvailable
    End If
    
    '注册并判断是否允许正常启用视频，不允许则停止视频显示
    If Not CheckVideoReg(blnIsStartupVideo) Then
        Call mVideoCapture.StopPreview
        
        If mblnIsExecuteReg Then
            mVideoCapture.IsAllowStartupVideo = False
        End If
    Else
        Call OpenComm(False) '打开采集端口
    End If
    
    '注册失败后重置显示界面，进入twain的操作界面
    '*****************************************************
    '该方法由采集参数配置窗口调用
    '这里进行注册是因为可能出现参数配置不对，或者硬件产生的视频不能启动，造成没有对系统进行注册，因而部分功能无法使用
    '当对视频参数进行设置后，有可能相关配置已经被正确修改，所以需要重新进行注册，启用相关功能
    '*****************************************************
    If mstrVideoRegTime = "" Then
      mVideoDriverType = vdtTWAIN
      mVideoCapture.VideoDriverType = vdtTWAIN
      '配置界面
      Call CaptureSwitchFace(IsTwainCaptureWay)
    End If
    '*********************************************************
    
'    If mVideoCapture.IsStartup Then Call ucCapHook.EnableHook
End Sub


Private Sub UpdateAfterCaptureInfo()
'更新后台采集信息
    
    '只有影像采集模块并且启用后后台采集才能使用后台采集
    If mlngModul = G_LNG_VIDEOSTATION_MODULE And Not IsTwainCaptureWay And mblnAfterIsUse And mVideoCapture.IsAllowStartupVideo Then
        Call CreateNewCaptureTag
        Call ShowAfterCaptureInf
    End If
End Sub


Private Sub Form_Load()
  On Error GoTo errHandle
    Dim strRegPath As String
    '设置窗口样式
    Call SetWindowStyle
    
    '该方法在show之后才会触发
    mIsShowing = False
        
    
    InitCommandBars
    
    Call ucPreview.InitImgPreview(gcnOracle)
        
    strRegPath = "公共模块\" & App.ProductName & "\frmVideoCapture"
    ucPreview.PageImgCount = Val(GetSetting("ZLSOFT", strRegPath, "采集缩略图数量", 5))
    
    mstrTempDirOfScan = App.Path + CONST_STR_DEFAULT_TEMP_SCAN_DIR_NAME
    If Len(mstrTempDirOfScan) > 45 Then
        Dim strFolder As String
        Dim pathlen As Long
        
        strFolder = String(256, 0)
        pathlen = GetTempPath(256, strFolder)
        If pathlen > 0 Then
            mstrTempDirOfScan = Left(strFolder, pathlen - 1) + CONST_STR_DEFAULT_TEMP_SCAN_DIR_NAME
        End If
    End If
    
    
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub


'返回是否为TWAIN的采集方式
Private Function IsTwainCaptureWay() As Boolean
  IsTwainCaptureWay = IIf(mVideoDriverType = vdtTWAIN, True, False)
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
Public Sub UpdateCaptureDirver(ByVal videoDirver As TVideoDriverType)
    '如果注册失败，则不允许对驱动类型更新
   If mstrVideoRegTime = "" And mblnIsExecuteReg Then
       Exit Sub
   End If
 
    '先停止视频的预览
    Call mVideoCapture.StopPreview
    
    mVideoDriverType = videoDirver
    mVideoCapture.VideoDriverType = videoDirver
       
    '读取视频大小
    mVideoSize.Width = ScaleX(mVideoCapture.VideoSize.Width, vbPixels, vbTwips)
    mVideoSize.Height = ScaleX(mVideoCapture.VideoSize.Height, vbPixels, vbTwips)
       
    Call CaptureSwitchFace(videoDirver = vdtTWAIN)
        
    
    '如果不是Twain采集方式，则重新启动预览
    If videoDirver <> vdtTWAIN Then
      mblnRealTime = True
      
      '开始预览
      Call mVideoCapture.StartPreview
      
    Else
      mblnRealTime = False
    End If
End Sub


Public Sub SaveVideoAreaCfg(ByVal strAreaName As String)
'保存视频采集区域配置
  Dim strRegPath As String
  
  '保存注册表参数
  strRegPath = "公共模块\" & App.ProductName & "\" & strAreaName
  SaveSetting "ZLSOFT", strRegPath, "CY1", picCapture.Height
End Sub


Public Sub LoadVideoAreaCfg(ByVal strAreaName As String)
'载入视频采集区域配置
    Dim strRegPath As String
     
    strRegPath = "公共模块\" & App.ProductName & "\" & strAreaName
    picCapture.Height = Val(GetSetting("ZLSOFT", strRegPath, "CY1", picCapture.Height))
End Sub


'保存当前参数设置
Public Sub SaveParameterCfg()
  Dim strRegPath As String
  
  '保存注册表参数
  strRegPath = "公共模块\" & App.ProductName & "\frmVideoCapture"
    
  '裁剪参数设置
  SaveSetting "ZLSOFT", strRegPath, "mdblX1Scale", mCurCutRange.LeftRate
  SaveSetting "ZLSOFT", strRegPath, "mdblX2Scale", mCurCutRange.WidthRate
  SaveSetting "ZLSOFT", strRegPath, "mdblY1Scale", mCurCutRange.TopRate
  SaveSetting "ZLSOFT", strRegPath, "mdblY2Scale", mCurCutRange.HeightRate
  
  
  '显示处理工具栏
  SaveSetting "ZLSOFT", strRegPath, "显示处理工具栏", mblnShowProcessBar
    
        
  '保存采集参数
  If Not mVideoCapture Is Nothing Then Call mVideoCapture.SaveCaptureParameterToFile(App.Path & "\" & CAPTURE_PARAMETER_CONFIG_FILE_NAME)
End Sub


Private Sub OpenComm(blnForce As Boolean)
    
    On Error GoTo err
    
    If mstrActiveType = "无" Then Exit Sub
    
    If mstrActiveType = "COM" Then
        
        If commListener.PortOpen Then Exit Sub

        commListener.CommPort = mstrComPort
        commListener.Settings = "9600,N,8,1"
        commListener.InputMode = comInputModeText
        commListener.RThreshold = 1
        commListener.InBufferCount = 0
        commListener.InputLen = 0
        commListener.RTSEnable = True
                        
        commListener.PortOpen = True
            
        '记录常态电平电位
        mblnCTSHolding = commListener.CTSHolding
        
    Else
        
        If mclsDxDevice Is Nothing Then Set mclsDxDevice = New clsDxHidDevice
        
        '打开DX设备
        Call mclsDxDevice.OpenDxDevice(mstrActiveType)
        
        tmrComm.Enabled = True
        tmrComm.Interval = 2
    End If
    
    Exit Sub
err:
    MsgBox "端口打开错误", vbOKOnly, "提示"
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
    If mblnDcmViewDown = True And Button = 1 And dcmView.Images.Count > 0 Then
        Select Case mintMouseState
            Case 1  '亮度对比度
                dcmView.Images(1).Width = dcmView.Images(1).Width + (X - mlngBaseX)
                dcmView.Images(1).Level = dcmView.Images(1).Level + (Y - mlngBaseY)
                mlngBaseX = X
                mlngBaseY = Y
            Case 2  '缩放
                Dim dblZoom As Double
                dblZoom = dcmView.Images(1).ActualZoom
                dblZoom = dblZoom * (1 + (Y - mlngBaseY) * 0.001)
                If dblZoom < 64 And dblZoom > 0.01 Then
                    subCenterZoom dcmView.Images(1), dcmView, dblZoom, mCorpSize
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


Private Sub RectangleZoom(Viewer As DicomViewer, img As DicomImage, lngLeft As Long, lngTop As Long, lngWidth As Long, lngHeight As Long)
    Dim newZoom As Double
    Dim dblRatio As Double
    Dim sX As Long
    Dim sY As Long
    Dim oldZoom As Double
    
    If lngWidth > 0 And lngHeight > 0 Then
        oldZoom = img.ActualZoom
        sX = img.ActualScrollX
        sY = img.ActualScrollY
        
        img.StretchToFit = False
        
        dblRatio = Viewer.Width / Screen.TwipsPerPixelX / lngWidth
        If dblRatio > Viewer.Height / Screen.TwipsPerPixelY / lngHeight Then
            dblRatio = Viewer.Height / Screen.TwipsPerPixelY / lngHeight
        End If
        
        newZoom = oldZoom * dblRatio
        img.Zoom = newZoom
        
        img.ScrollX = (sX + lngLeft) * dblRatio
        img.ScrollY = (sY + lngTop) * dblRatio
    End If
End Sub


Private Sub subCenterZoom(img As DicomImage, Viewer As DicomViewer, dblZoom As Double, corpSize As TPoint)
'------------------------------------------------
'功能：对图像进行缩放。以当前viewer中心点为缩放中心点。
'参数：
'       img -- 进行缩放的图像
'       viewer －－ 图像所在的viewer
'       dblZoom －－图像新的缩放倍数
'返回：无，直接调整图像的缩放倍数
'上级函数或过程：frmViewer.Viewer_MouseMove
'下级函数或过程：无
'引用的外部参数：无
'编制人： 黄捷 2006-2-10
'------------------------------------------------
    img.Zoom = dblZoom
    img.StretchToFit = False

            
    img.ScrollX = (img.SizeX * img.ActualZoom - ScaleX(Viewer.Width, vbTwips, vbPixels) / Viewer.MultiColumns) / 2 + corpSize.X
    img.ScrollY = (img.SizeY * img.ActualZoom - ScaleY(Viewer.Height, vbTwips, vbPixels) / Viewer.MultiRows) / 2 + corpSize.Y
End Sub


Public Function GetNewLabel(lType As Integer, lLeft As Integer, lTop As Integer, lWidth As Integer, lHeight As Integer) As DicomLabel
'------------------------------------------------
'功能：生成一个LABEL对象，并对其做初始化。
'参数：lType--标注的类型；lLeft--标注的Left值；lTop--标注的Top值；lWidth--标注的Width值；lHeight--标注的Height值。
'返回：新生成的标注。
'编制人：黄捷
'------------------------------------------------
    Dim l As New DicomLabel
    l.LabelType = lType
    l.XOR = True
    l.ImageTied = True
    l.Left = lLeft
    l.Top = lTop
    l.Width = lWidth
    l.Height = lHeight
    l.Margin = 0
    l.AutoSize = True
    l.FontSize = 12
    l.LineWidth = 1
    If l.LabelType = 0 Then     '文字
        l.Transparent = False
        l.Width = 200
        l.Height = 10
    End If
    Set GetNewLabel = l
End Function
   
   
Public Sub subCaptureImg(ByVal RealTimeCap As Boolean, Optional ByVal strFileName As String = "", _
    Optional ByRef picCapture As StdPicture = Nothing, Optional ByVal blnIsAfterCapture As Boolean = False)
'------------------------------------------------
'功能：采集并存储图像
'参数：无
'返回：无，直接保存新采集的图像
'------------------------------------------------
On Error GoTo errHandle
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    
    
    '
    If mblnUseInetFtp Then
        If mblnCurCaptureState Then Exit Sub
        
        mblnCurCaptureState = True
    End If
    
    If funCaptureSingleImage(RealTimeCap, strFileName, picCapture, blnIsAfterCapture) = True Then
        If blnIsAfterCapture Then
            '如果是后台采集，则后台采集成功后，删除后台采集的图像
            If subSaveAfterCaptureImage Then Call dcmAfter.Images.Clear
            
            Call ShowAfterCaptureInf
            
            mblnCurCaptureState = False
            Exit Sub
        End If
        
        Call subSaveImage
        
        '设置影像检查状态，如果采集第一张图，且原来的状态是已报到，则修改成已检查
        If ucPreview.ImgViewer.Images.Count = 1 Then
            
            If mlngStudyState < 3 Then
                strSql = "Zl_影像检查_State(" & mlngAdviceID & "," & mlngSendNo & ",3,NULL,'" & UserInfo.编号 & "','" & UserInfo.姓名 & "'," & mlngCurDeptId & ")"
                zlDatabase.ExecuteProcedure strSql, "采集第一个图像"
            End If
        End If
        
        
        If ucPreview.ImgViewer.Images.Count = 1 Then
            Call SendMsgToMainWindow(Me, wetCaptureFirstImg, mlngAdviceID, mstrStudyUID)
        Else
            Call SendMsgToMainWindow(Me, wetUpdateImg, mlngAdviceID, mstrStudyUID)
        End If
    End If
    
    
    mblnCurCaptureState = False
Exit Sub
errHandle:
    mblnCurCaptureState = False
    err.Raise err.Number, err.Description
End Sub

Private Function CopyPictureToDicomImg(ByVal lngHDC As Long, ByVal lngPictureHandle As Long, objDcmImg As Object) As Boolean
    Const bitCount As Long = 3
'congpicture中复制图像到dicomimage
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
    Optional ByVal strFileName As String = "", Optional ByRef picCapture As StdPicture = Nothing, _
    Optional ByVal blnIsAfterCapture As Boolean = False) As Boolean
'------------------------------------------------
'功能：采集单帧视频图像，将图像转换成DICOM格式，并填写DICOM文件头，最后将图像放入缩略图dcmMiniature中。
'参数：无
'返回：无，直接将新采集的图像放入dcmMiniature中
'------------------------------------------------
'采集单帧图像
On Error GoTo SaveFileError
    Dim ImgTmpImage As New DicomImage

    
    '采集图像，分为直接视频采集和播放录象采集

    If Not (picCapture Is Nothing) Then
        Set picTemp2.Picture = Nothing
        Set picTemp2.Picture = picCapture
        
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
                        
            'modify by tjh at 2009-01-20
            Dim curPic As StdPicture
            Set curPic = mVideoCapture.CaptureImageFromMemory

            If curPic Is Nothing Then
                Call MsgBoxD(Me, "视频图像采集失败，请检查视频参数设置是否正确(如视频设备，显示模式等)。", vbOKOnly, "提示")
                
                funCaptureSingleImage = False
                Exit Function
            End If
            
            picTemp2.Width = mVideoSize.Width * (1 - mCurCutRange.WidthRate - mCurCutRange.LeftRate)
            picTemp2.Height = mVideoSize.Height * (1 - mCurCutRange.HeightRate - mCurCutRange.TopRate)

            Call picTemp2.PaintPicture(curPic, 0, 0, picTemp2.Width, picTemp2.Height, _
                                       mVideoSize.Width * mCurCutRange.LeftRate, mVideoSize.Height * mCurCutRange.TopRate, _
                                       picTemp2.Width, picTemp2.Height, vbSrcCopy)
                                               
            picTemp2.Picture = picTemp2.Image

            Set curPic = Nothing
        End If
    End If
    
    
    '将图像再次提交到剪切板
    If picTemp2.Picture Is Nothing Then
        funCaptureSingleImage = False
        Exit Function
    End If
  

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
    Call subWriteDicomPara(ImgTmpImage, mlngAdviceID, blnIsAfterCapture)
    
    Dim dcmTag As New clsImageTagInf
    dcmTag.Tag = IMGTAG
    
    Set ImgTmpImage.Tag = dcmTag
    
    If blnIsAfterCapture Then
        Call dcmAfter.Images.Add(ImgTmpImage)
    Else
        '将图像放入缩略图中
        Call subInsert2Mini(ImgTmpImage)
    End If
    
    
    funCaptureSingleImage = True
    
    Exit Function
SaveFileError:
    If ErrCenter() = 1 Then Resume
End Function


Private Sub subWriteDicomPara(img As DicomImage, lngAdviceID As Long, _
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
        img.Attributes.Add &H8, &H60, mstrModality                   'Modality 影像类别
    Else
        img.Attributes.Add &H10, &H10, mstrName                     'Name 姓名
        img.Attributes.Add &H10, &H20, mstrPatientID                'Patient ID 病人ID
        img.Attributes.Add &H10, &H30, mstrBirthDate                'BirthDate 生日
        img.Attributes.Add &H10, &H40, mstrSex                      'Sex 性别
        img.Attributes.Add &H10, &H1010, mstrAge                    'Age 年龄
        img.Attributes.Add &H10, &H4000, ""                         'Patient Comment 病人注释
        img.Attributes.Add &H20, &H10, mstrCheckNo                  'Study ID 检查ID
        img.Attributes.Add &H8, &H60, mstrModality                   'Modality 影像类别
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
    
    ucPreview.AddImage img
End Sub


Private Sub Form_Resize()
On Error GoTo errHandle
    picDock.Left = 0
    picDock.Top = 0
    picDock.Width = Me.ScaleWidth
    picDock.Height = Me.ScaleHeight

    Call ucSplitter1.RePaint(False)
    
errHandle:
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim strRegPath As String

    
    '卸载视频注册
    Call FunLogOut(Me, LOGIN_TYPE_视频设备, mstrVideoRegTime)

    '先关闭采集窗口和COMM口
    Call zlStopCapture
  
    '保持裁剪设置
    Call SaveParameterCfg
    
    '保存视频采集区域设置
    If Not mRestoreContainer Is Nothing Then
        Call SaveVideoAreaCfg(mRestoreContainer.Name)
    End If

    strRegPath = "公共模块\" & App.ProductName & "\frmVideoCapture"
    Call SaveSetting("ZLSOFT", strRegPath, "采集缩略图数量", ucPreview.PageImgCount)
    
'    Call mobjInetFtp.QuitFtp
    
    Set mclsDxDevice = Nothing
    Set mVideoCapture = Nothing
    Set mParentContainer = Nothing
    Set mRestoreContainer = Nothing
    Set mobjOwner = Nothing
End Sub


Private Sub subDeleteImage()
'------------------------------------------------
'功能：删除缩略图中被选中的图像，先从数据库中删除，然后从FTP中删除。删除后触发StateChanged事件
'参数：无
'返回：无，直接删除缩略图中最后一个图像
'------------------------------------------------
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    
    If ucPreview.ImgViewer.Images.Count > 0 Then
        
        Dim blnResult As Boolean
                 

        '从数据库和FTP中删除缩略图中被选中的图像
        blnResult = DeleteImages(Me, 1, ucPreview.SelectImage.InstanceUID, "")
        
        If blnResult = True Then    '删除成功，则修改缩略图状态，并触发StateChanged事件
            '在缩略图中删除图像
            Call ucPreview.DeleteImage(ucPreview.SelectIndex)
            dcmView.Images.Clear

            If Not ucPreview.SelectImage Is Nothing Then
                dcmView.Images.Add ucPreview.SelectImage
            End If
            
            
            '设置影像检查状态，如果删除最后一个图，且原检查过程为3，则修改为2
            If ucPreview.CurImageCount = 0 Then
                
                If mlngStudyState = 3 Then
                    strSql = "Zl_影像检查_State(" & mlngAdviceID & "," & mlngSendNo & ",2,NULL,'" & UserInfo.编号 & "','" & UserInfo.姓名 & "'," & mlngCurDeptId & ")"
                    zlDatabase.ExecuteProcedure strSql, "删除最后一个图像"
                End If
                
                Call SendMsgToMainWindow(Me, wetDelAllImg, mlngAdviceID, mstrStudyUID)
                
                mstrStudyUID = ""
                
                '当最后的图像删除时，则显示实时视频画面
                Call ConfigVideoShowState(True)
            Else
                Call SendMsgToMainWindow(Me, wetUpdateImg, mlngAdviceID, mstrStudyUID)
            End If
        End If
    End If
End Sub


Private Sub subSetMouseState(intMouseState As Integer)
    '改变当前鼠标状态
    If mintMouseState = intMouseState Then
        mintMouseState = 0
    Else
        mintMouseState = intMouseState
    End If
    cbrMain.FindControl(xtpControlButton, conMenu_Process_Window).Checked = False
    cbrMain.FindControl(xtpControlButton, conMenu_Process_Zoom).Checked = False
    cbrMain.FindControl(xtpControlButton, conMenu_Process_RectZoom).Checked = False
    cbrMain.FindControl(xtpControlButton, conMenu_Process_Arrow).Checked = False
    cbrMain.FindControl(xtpControlButton, conMenu_Process_Ellipse).Checked = False
    cbrMain.FindControl(xtpControlButton, conMenu_Process_Text).Checked = False
    cbrMain.FindControl(xtpControlButton, conMenu_Process_Corp).Checked = False
End Sub


Private Sub subSetSharp(blnSharp As Boolean)
'------------------------------------------------
'功能：dcmView中图像的平滑和锐化
'参数：blnSharp表示图像处理的方向，True=锐化；False=平滑
'返回：无，直接处理dcmView中的图像
'------------------------------------------------
    If mblnRealTime = False And dcmView.Images.Count > 0 Then
        If blnSharp = True Then
            '锐化处理
            If dcmView.Images(1).FilterLength <= 0 Then
                dcmView.Images(1).FilterLength = 0
                '先前没有平滑处理，直接进行锐化处理
                dcmView.Images(1).UnsharpEnhancement = dcmView.Images(1).UnsharpEnhancement + 0.1
            Else
                '如果先前已经有平滑处理，则先淡化平滑效果
                dcmView.Images(1).FilterLength = dcmView.Images(1).FilterLength - 1
            End If
        Else
            '平滑处理
            '判断Zoom是否＝1，如果是，则修改为0.9999
            If dcmView.Images(1).ActualZoom = 1 Then
                dcmView.Images(1).Zoom = 0.9999
            End If
            
            If dcmView.Images(1).UnsharpEnhancement <= 0 Then
                dcmView.Images(1).UnsharpEnhancement = 0
                '先前没有锐化处理，直接开始平滑
                '判断FilterLength是否＝0如果是，则在2/ActualZoom和2×FilterLength之间进行调整
                If dcmView.Images(1).FilterLength = 0 Then
                    dcmView.Images(1).FilterLength = 2 / dcmView.Images(1).ActualZoom + 1
                Else    '正常情况下FilterLength＋1
                    dcmView.Images(1).FilterLength = dcmView.Images(1).FilterLength + 1
                End If
            Else
                '先前已经有了锐化处理，先淡化锐化的效果
                dcmView.Images(1).UnsharpEnhancement = dcmView.Images(1).UnsharpEnhancement - 0.1
            End If
        End If
    End If
End Sub


Private Sub subSetRotate(blnClockwise As Boolean)
'------------------------------------------------
'功能：dcmView中图像的旋转
'参数：blnClockwise旋转的方向,True=顺时针旋转；False=逆时针旋转
'返回：无，直接处理dcmView中的图像
'------------------------------------------------
    If mblnRealTime = False And dcmView.Images.Count > 0 Then
        Dim iRotateState As Integer
        
        iRotateState = dcmView.Images(1).RotateState
        If blnClockwise = True Then
            iRotateState = iRotateState - 1
        Else
            iRotateState = iRotateState + 1
        End If
        If iRotateState = -1 Then iRotateState = 3
        iRotateState = iRotateState Mod 4
        dcmView.Images(1).RotateState = iRotateState
    End If
End Sub


'modify by tjh at 2010-01-20
'配置视频显示状态
Public Sub ConfigVideoShowState(ByVal blnShowState As Boolean)
  mblnRealTime = blnShowState
  
  Select Case mVideoDriverType
    Case vdtVFW
      picVideo.Visible = blnShowState
      wdmCapture.Visible = False
      dcmView.Visible = Not blnShowState
    Case vdtWDM
      wdmCapture.Visible = blnShowState
      picVideo.Visible = False
      dcmView.Visible = Not blnShowState
    Case vdtTWAIN
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


Private Sub mclsDxDevice_OnDxKeyPress(ByVal lngButtonNum As Long)
On Error GoTo errHandle
    
    Select Case lngButtonNum
        Case 0  '前台采集
            If Me.cbrMain.FindControl(xtpControlButton, conMenu_Cap_MarkMap).Enabled And _
                Me.cbrMain.FindControl(xtpControlButton, conMenu_Cap_MarkMap).Visible Then
                Call subCaptureImg(True)
            End If
        Case 1  '后台采集
            If Me.cbrMain.FindControl(xtpControlButton, conMenu_Cap_After_Capture).Enabled And _
                Me.cbrMain.FindControl(xtpControlButton, conMenu_Cap_After_Capture).Visible Then
                Call subCaptureImg(True, "", Nothing, True)
            Else
                Call mclsDxDevice_OnDxKeyPress(0)
            End If
        Case 2  '更新标记
            
            If Me.cbrMain.FindControl(xtpControlButton, conMenu_Cap_After_Tag).Enabled And _
                Me.cbrMain.FindControl(xtpControlButton, conMenu_Cap_After_Tag).Visible Then
                Call UpdateAfterCaptureInfo
            Else
                Call mclsDxDevice_OnDxKeyPress(0)
            End If
        Case Else
            Call mclsDxDevice_OnDxKeyPress(0)
    End Select

    Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
End Sub

'modify by tjh at 2010-01-20
Private Sub pbxSize_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim intVideoCapture As Integer

    intVideoCapture = Val(zlDatabase.GetPara("允许改变采集区域大小", glngSys, mlngModul, "1", , InStr(mstrPrivs, ";参数设置;") > 0))
  '开始执行裁剪范围设置
    If Button = 1 And intVideoCapture = 1 Then
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
    
    If IsTwainCaptureWay Then
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
errHandle:
End Sub


Private Sub subVideoPlay()
'------------------------------------------------
'功能：dcmView中录像图像的播放
'参数：无
'返回：无，直接播放dcmView中的图像
'------------------------------------------------
    If dcmView.Images.Count > 0 Then
        '下载录像，如果本地存在，则不进行下载
        If dcmView.Images(1).Tag.Tag <> VIDEOTAG And dcmView.Images(1).Tag.Tag <> AUDIOTAG Then
            '不是电影格式不能播放,不用提示
            Exit Sub
        End If
        
        On Error GoTo continue1
        
            If dcmView.Images(1).Tag.Tag = VIDEOTAG Then
                Call dcmView.Images(1).FileImport(IIf(Len(App.Path) > 3, App.Path & "\..\附加文件\aviDownload.bmp", App.Path & "..\附加文件\aviDownLoad.bmp"), "DIB/BMP")
        
                '下载需要播放的视频
                Call GetSingleImage(dcmView.Images(1).InstanceUID, dcmView.Images(1).SeriesUID, mblnMoved)
            
                Call dcmView.Images(1).FileImport(IIf(Len(App.Path) > 3, App.Path & "\..\附加文件\avi.bmp", App.Path & "..\附加文件\avi.bmp"), "DIB/BMP")
            Else
                Call dcmView.Images(1).FileImport(IIf(Len(App.Path) > 3, App.Path & "\..\附加文件\wavDownload.bmp", App.Path & "..\附加文件\wavDownLoad.bmp"), "DIB/BMP")
        
                '下载需要播放的视频
                Call GetSingleImage(dcmView.Images(1).InstanceUID, dcmView.Images(1).SeriesUID, mblnMoved)
            
                Call dcmView.Images(1).FileImport(IIf(Len(App.Path) > 3, App.Path & "\..\附加文件\wav.bmp", App.Path & "..\附加文件\wav.bmp"), "DIB/BMP")
            End If
            
continue1:
            '打开播放··
            Call frmPlaying.Show
            
            '刷新播放窗口
'            Call frmPlaying.Refresh
            While Not frmPlaying.IsActive
                Call Sleep(10)
                DoEvents
            Wend
                
            
            Call frmPlaying.OpenVideoFile(Replace(dcmView.Images(1).Tag.VideoFile, "/", "\"), Me)
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
            strFileName = dlgOpen.Filename
            
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
        strFileName = dlgOpen.Filename
        
        If strFileName <> "" Then
            strFileType = UCase(Right(Trim(strFileName), 3))
            
            Select Case strFileType
                Case "AVI"
                    If dcmView.Images(1).FrameCount > 1 Then
                        dcmView.Images(1).WriteAVI strFileName, 1, dcmView.Images(1).FrameCount, 1, "", 100, False
                    Else
                        MsgBoxD Me, "静态图像无法保存成AVI格式，请重新选择图像格式。", vbInformation, gstrSysName
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
    Dim DlgInfo As DlgFileInfo
    Dim i As Integer
    Dim ImgTmpImage As New DicomImage
    Dim ImgTmpImages As New DicomImages
    Dim blDicomFile As Boolean              '是否DICO文件 =True为DICOM文件
    Dim j As Integer
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    
    '选择文件
    With Me.dlgOpen
        .CancelError = False
        .MaxFileSize = 32767 '被打开的文件名尺寸设置为最大，即32K
        .flags = cdlOFNHideReadOnly Or cdlOFNAllowMultiselect Or cdlOFNExplorer Or cdlOFNNoDereferenceLinks
        .DialogTitle = "选择文件"
        .Filter = "DICOM文件（*.dcm）(*.img)|*.dcm;*.img|图像文件 (*.BMP)(*.JPG)|*.BMP;*.JPG|所有文件（*.*）|*.*"
        .ShowOpen
        If .Filename <> "" Then
            DlgInfo = GetDlgSelectFileInfo(.Filename)
        End If
        '在打开了*.pif文件后须将Filename属性置空，否则当选取多个*.pif文件后，当前路径会改变
        .Filename = ""
    End With
    
    On Error Resume Next
    
    For i = 1 To DlgInfo.iCount
        err.Clear
        Set ImgTmpImage = Nothing
        ImgTmpImages.Clear
        ImgTmpImage.FileImport DlgInfo.sPath & DlgInfo.sFIle(i), ""
        If err <> 0 Then
            err.Clear
            ImgTmpImages.ReadFile DlgInfo.sPath & DlgInfo.sFIle(i)
            If err = 0 Then
                blDicomFile = True
            End If
        End If
        
        If blDicomFile = True And ImgTmpImages.Count > 0 Then
            Set ImgTmpImage = ImgTmpImages(1)
        End If
        
        '设置图像的DICOM参数
        subWriteDicomPara ImgTmpImage, mlngAdviceID
        
        Dim dcmTag As New clsImageTagInf
        dcmTag.Tag = IMGTAG
    
        Set ImgTmpImage.Tag = dcmTag
    
        '将图像插入到缩略图中
        subInsert2Mini ImgTmpImage
        '保存图像，并触发图像存储事件
        Call subSaveImage
        
        '设置影像检查状态，如果采集第一张图，且原来的状态是已报到，则修改成已检查
        If ucPreview.CurImageCount = 1 Then
            If mlngStudyState < 3 Then
                strSql = "Zl_影像检查_State(" & mlngAdviceID & "," & mlngSendNo & ",3,NULL,'" & UserInfo.编号 & "','" & UserInfo.姓名 & "'," & mlngCurDeptId & ")"
                zlDatabase.ExecuteProcedure strSql, "采集第一个图像"
            End If
        End If
        
        If ucPreview.CurImageCount = 1 Then
            Call SendMsgToMainWindow(Me, wetCaptureFirstImg, mlngAdviceID)
        End If
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
                If ucPreview.ImgViewer.Images(i).Tag.Tag = IMGTAG Then
                    dcmImg.SeriesUID = ucPreview.ImgViewer.Images(i).SeriesUID
                    
                    Exit For
                End If
            Next i
            
        End If
    ElseIf Len(mstrStudyUID) > 0 Then
        dcmImg.StudyUID = mstrStudyUID
    Else
        mstrStudyUID = dcmImg.StudyUID
        
        '当检查uid改变后，需要更新缩略图显示组件中的查询值
        ucPreview.QueryValue = mstrStudyUID
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
    MsgBox "GetDlgSelectFileInfo函数执行错误！", vbOKOnly + vbCritical, gstrSysName
End Function



Private Sub TimerHook_Timer()
On Error GoTo errHandle
    '当使用hook热键调用采集时，使用timer进行采集操作，避免在执行多次CaptureImage操作后，hook失效
    '造成hook失效的可能原因有hook的处理机制中如果截获hook后的处理时间过长可能造成失效，或者dicomobjects的fileexport方法调用多次造成失效，目前不去细究
    Call CaptureImage
    timerHook.Enabled = False
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub tmrComm_Timer()
    On Error GoTo errHandle
    If mstrActiveType = "COM" Then
        mlngComTime = mlngComTime + 2
        
        '大于0.08秒，则自动放弃
        If mlngComTime > 40 Then
            mlngComTime = 0
            tmrComm.Enabled = False
        End If
        
    Else
         If Not mclsDxDevice Is Nothing Then Call mclsDxDevice.PollDxDevice
    End If
    
    Exit Sub
errHandle:
    tmrComm.Enabled = False
    If ErrCenter = 1 Then
        Resume
    End If
End Sub


Private Sub tmrReg_Timer()
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    
On Error GoTo errHandle:
    If Not mVideoCapture.IsStartup Then
        Exit Sub
    End If
    
    If gint视频设备数量 <= -1 Then Exit Sub
    
    strSql = "select count(1) 已启用数量 from zltools.zlclients where 启用视频源=1"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "已启用数量")
    
    If rsTemp.RecordCount > gint视频设备数量 Then
        mstrVideoRegTime = ""
        
        Exit Sub
    End If
    
    If DateDiff("S", mstrVideoRegTime, Now) >= M_LNG_REFRESHINTERVAL Then
        '判断数据库中是否存在已经注册的ip并且已经启用视频源，如果不存在则认为没有成功注册
        If FunCheckRegInfo(Me) Then
            mstrVideoRegTime = Now
        Else
            mstrVideoRegTime = ""
            
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
    If Trim(strError) <> "" Then MsgBoxD Me, strError, vbInformation, gstrSysName
    
    '获取当前录像的编码器名称
    mstrEncoderName = mVideoCapture.GetEncoderName
    
    Exit Sub
CapErr:
  Call MsgBox(err.Description, vbOKOnly, "提示")
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
            dcmTmpImg.FileImport IIf(Len(App.Path) > 3, App.Path & "\..\附加文件\avi.bmp", App.Path & "..\附加文件\avi.bmp"), "DIB/BMP"
continue2:
        
        Dim dcmTag As New clsImageTagInf
        dcmTag.EncoderName = mstrEncoderName
        dcmTag.VideoFile = mstrAviFileName
        dcmTag.CaptureTime = zlDatabase.Currentdate
        dcmTag.RecordTimeLen = mVideoCapture.GetTimeLen
        dcmTag.Tag = VIDEOTAG
        
        Set dcmTmpImg.Tag = dcmTag
        
        subWriteDicomPara dcmTmpImg, mlngAdviceID
        
        subInsert2Mini dcmTmpImg
        
        '保存视频录像
        Call subSaveImage
    End If
    
    '如果是录像，也需要对状态进行更新
    If ucPreview.CurImageCount = 1 Then
        Call SendMsgToMainWindow(Me, wetCaptureFirstImg, mlngAdviceID)
    End If
    
    Exit Sub
CapErr:
    If ErrCenter() = 1 Then Resume
End Sub


'停止音频文件
Public Sub subSaveAudio(ByVal strAudioFile As String, ByVal lngTimeLen As Long)

    Dim i As Integer
    Dim dcmTmpImg As New DicomImage
    
    On Error GoTo CapErr
   
    
    '录像完成
    If Dir(strAudioFile) <> "" Then
        dcmTmpImg.FileImport IIf(Len(App.Path) > 3, App.Path & "\..\附加文件\wav.bmp", App.Path & "..\附加文件\wav.bmp"), "DIB/BMP"
        
        Dim dcmTag As New clsImageTagInf
        dcmTag.EncoderName = ""
        dcmTag.VideoFile = strAudioFile
        dcmTag.CaptureTime = zlDatabase.Currentdate
        dcmTag.RecordTimeLen = lngTimeLen
        dcmTag.Tag = AUDIOTAG
        
        Set dcmTmpImg.Tag = dcmTag
        
        subWriteDicomPara dcmTmpImg, mlngAdviceID
        
        subInsert2Mini dcmTmpImg
        
        '保存录制的音频
        Call subSaveImage
    End If
    
    '如果是录像，也需要对状态进行更新
    If ucPreview.CurImageCount = 1 Then
        Call SendMsgToMainWindow(Me, wetCaptureFirstImg, mlngAdviceID)
    End If
    
    Exit Sub
CapErr:
    If ErrCenter() = 1 Then Resume
End Sub

'modify by tjh at 2010-01-22
'全屏显示
Private Sub subFullCall()
  Call mVideoCapture.FullScreen(Me, Me.hWnd)
End Sub


Private Function GetCaptureTag() As String
'取得后台采集标记
    Dim i As Long
    Dim strSql As String
    Dim rsData As ADODB.Recordset
        
    GetCaptureTag = "001"
        
    strSql = "select 检查号 from 影像临时记录 where 姓名='后台'"
    Set rsData = zlDatabase.OpenSQLRecord(strSql, Me.Caption)
    
    If rsData.RecordCount <= 0 Then Exit Function
    
    For i = 1 To 999
        rsData.Filter = " 检查号='" & Lpad(i, 3, "0") & "'"
        If rsData.RecordCount <= 0 Then
            GetCaptureTag = Lpad(i, 3, "0")
            Exit Function
        End If
    Next i
    
    GetCaptureTag = ""
End Function



Private Sub CreateNewCaptureTag()
'取得新的采集标记
    mstrAfterStudyUid = CreateStudyUid(dcmglbUID.NewUID)
    mstrAfterSeriesUid = CreateSeriesUid(dcmglbUID.NewUID)
    
    mstrAfterTag = GetCaptureTag
    
    mlngAfterCurImageCount = 0
End Sub


Private Sub ShowAfterCaptureInf()
'更新后台采集图像信息
    If Not mblnAfterIsUse Then Exit Sub
    
    If mobjOwner Is Nothing Then Exit Sub
    
    If mstrAfterParentTitle = "" Then
        If InStr(mobjOwner.Caption, "      后台采集标记：") > 0 Then
            mstrAfterParentTitle = Mid(mobjOwner.Caption, 1, InStr(mobjOwner.Caption, "      后台采集标记：") - 1)
        Else
            mstrAfterParentTitle = mobjOwner.Caption
        End If
    End If
    
    mobjOwner.Caption = mstrAfterParentTitle & "      后台采集标记：" & mstrAfterTag & "  当前后台采集数：" & mlngAfterCurImageCount & "        "
End Sub


Private Function subSaveAfterCaptureImage(Optional iEncode As Integer = 0) As Boolean
'保存后台采集图像
    Dim i As Long
    Dim lngResult As Long
    Dim strSql As String
    Dim dtNowTime As Date
    Dim strReceivedTime As String
    Dim ImgTmp As DicomImage

    subSaveAfterCaptureImage = False
    
    If dcmAfter.Images.Count <= 0 Then Exit Function
    
    dtNowTime = zlDatabase.Currentdate
    strReceivedTime = Format(dtNowTime, "yyyyMMdd")
    
    If mstrAfterStudyUid = "" Then
        '如果uid为空，则创建新的UID
        mstrAfterStudyUid = dcmglbUID.NewUID
        mstrAfterSeriesUid = dcmglbUID.NewUID
        
        mstrAfterTag = GetCaptureTag()
    End If
    
    If Trim(mstrAfterTag) = "" Then
        Call MsgBoxD(Me, "不能获取有效的后台采集标记，请检查后台采集的检查数量是否已满，后台采集检查数不能超过1000。", vbOKOnly, Me.Caption)
        Exit Function
    End If

    '创建缓冲目录
    MkLocalDir mstrBufferDir & strReceivedTime & "/" & mstrAfterStudyUid & "/"
 
    '不使用inet方式时，需要先初始化ftp连接
    lngResult = mobjFtpConnection.FuncFtpConnect(mobjFtp.strFTPIP, mobjFtp.strFTPUser, mobjFtp.strFTPPwd)

    If lngResult = 0 Then
        'FTP操作失败，提示错误，并删除缩略图中的图像
        MsgBoxD Me, "FTP连接失败，后台采集图像无法保存，请检查网络设置。", vbInformation, gstrSysName
        Exit Function
    End If
        
    For i = 1 To dcmAfter.Images.Count
    
        Set ImgTmp = dcmAfter.Images(i)
        
        ImgTmp.StudyUID = mstrAfterStudyUid
        ImgTmp.SeriesUID = mstrAfterSeriesUid
        
        If ImgTmp.Tag.Tag <> VIDEOTAG And ImgTmp.Tag.Tag <> AUDIOTAG Then
            '保存图像到缓存目录
            Select Case iEncode
                Case 1          'Run-Length Encoding行程压缩
                    ImgTmp.WriteFile mstrBufferDir & strReceivedTime & "/" & mstrAfterStudyUid & "/" & ImgTmp.InstanceUID, True, "1.2.840.10008.1.2.5"
                Case 2          '不处理，保持原图的压缩方式
                    ImgTmp.WriteFile mstrBufferDir & strReceivedTime & "/" & mstrAfterStudyUid & "/" & ImgTmp.InstanceUID, True
                Case Else       'Lossless JPEG encoding JPEG无损压缩
                    ImgTmp.WriteFile mstrBufferDir & strReceivedTime & "/" & mstrAfterStudyUid & "/" & ImgTmp.InstanceUID, True, "1.2.840.10008.1.2.4.70"
            End Select
            
            '存储为报告图像
            ImgTmp.FileExport mstrBufferDir & strReceivedTime & "/" & mstrAfterStudyUid & "/" & ImgTmp.InstanceUID & ".jpg", "JPG", 80
        End If
        
        If ImgTmp.Tag.Tag <> VIDEOTAG And ImgTmp.Tag.Tag <> AUDIOTAG Then
            '保存dicom图像
            WriteToURL mstrBufferDir & strReceivedTime & "/" & mstrAfterStudyUid & "/" & ImgTmp.InstanceUID, mobjFtp.strFtpDir & _
                strReceivedTime & "/" & mstrAfterStudyUid & "/" & ImgTmp.InstanceUID
                
            '上传报告图
            WriteToURL mstrBufferDir & strReceivedTime & "/" & mstrAfterStudyUid & "/" & ImgTmp.InstanceUID & ".jpg", mobjFtp.strFtpDir & _
                strReceivedTime & "/" & mstrAfterStudyUid & "/" & ImgTmp.InstanceUID & ".jpg"
        Else
            '保存录像
            WriteToURL ImgTmp.Tag.VideoFile, mobjFtp.strFtpDir & strReceivedTime & "/" & mstrAfterStudyUid & "/" & ImgTmp.InstanceUID
            
            If ImgTmp.Tag.Tag = VIDEOTAG Then
                '将录像复制到对应的目录中，避免重新下载
                Call MoveFile(ImgTmp.Tag.VideoFile, mstrBufferDir & strReceivedTime & "/" & mstrAfterStudyUid & "/" & ImgTmp.InstanceUID & ".avi")
                
            ElseIf ImgTmp.Tag.Tag = AUDIOTAG Then
                '将音频文件复制到对应的目录中，避免重新下载
                Call MoveFile(ImgTmp.Tag.VideoFile, mstrBufferDir & strReceivedTime & "/" & mstrAfterStudyUid & "/" & ImgTmp.InstanceUID & ".wav")
                
            End If
        End If
        
        '图像存储成功后，存储数据库信息
        strSql = "ZL_影像检查_后台采集('" & mstrAfterModality & "','" & mstrAfterStudyUid & "','" & mstrAfterSeriesUid & "','" & _
                                        ImgTmp.InstanceUID & "','" & mstrAfterTag & "','" & mobjFtp.strDeviceId & "'," & _
                                        "to_Date('" & Format(dtNowTime, "yyyy-mm-dd hh:mm") & "','YYYY-MM-DD HH24:MI:SS'))"
        
        Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
        
        mlngAfterCurImageCount = mlngAfterCurImageCount + 1
    Next i
    
    If mblnUseInetFtp Then
        '使用inet ftp方式时，这里不需要断开连接
    Else
        mobjFtpConnection.FuncFtpDisConnect
    End If
    
    If ImgTmp.Tag.Tag <> VIDEOTAG And ImgTmp.Tag.Tag <> AUDIOTAG Then
        Call frmCaptureHint.ShowCaptureHint( _
            IIf(mblnPoputWindowHint, mstrBufferDir & strReceivedTime & "/" & mstrAfterStudyUid & "/" & ImgTmp.InstanceUID, ""), _
            mblnSoundHint, hpRB, Me)
            
    End If
    
    subSaveAfterCaptureImage = True
End Function


Private Sub subSaveImage(Optional iEncode As Integer = 0)
'------------------------------------------------
'功能：将最后一个缩略图保存到数据库中
'参数：iEncode－－压缩方式，1－Run-Length Encoding行程压缩；2－不处理，保持原图的压缩方式，其他－Lossless JPEG encoding JPEG无损压缩
'返回：无
'------------------------------------------------
    Dim rsTmp As New ADODB.Recordset
    Dim ImgTmp As New DicomImage
    
    Dim dtReceived As String
    Dim blnFirstImage As String     '是否本次检查的第一张图像
    Dim lngResult As String         'FTP操作结果
    Dim nowTime As Date
    Dim strReportImages As String
    Dim arrSQL() As Variant
    Dim blnInTrans As Boolean       '在事物处理过程中
    Dim i As Integer
    Dim lngSendNO As Long
    
    '读取最后一个缩略图
    With ucPreview.ImgViewer
        If .Images.Count <= 0 Then Exit Sub
        Set ImgTmp = .Images(.Images.Count)
    End With
    
    '先保存FTP图像
    '读取接收日期
    gstrSQL = "select 检查UID ,接收日期,报告图象,发送号 from 影像检查记录 where 医嘱ID = [1] "
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, App.ProductName, mlngAdviceID)
    nowTime = zlDatabase.Currentdate
    
    If IsNull(rsTmp("检查UID")) Then
        dtReceived = Format(nowTime, "yyyyMMdd")
        blnFirstImage = True
    Else
        dtReceived = Format(rsTmp("接收日期"), "yyyyMMdd")
        blnFirstImage = False
    End If
    
    '创建缓冲目录
    MkLocalDir mstrBufferDir & dtReceived & "/" & mstrStudyUID & "/"
    lngSendNO = rsTmp!发送号
    
    
    If ImgTmp.Tag.Tag <> VIDEOTAG And ImgTmp.Tag.Tag <> AUDIOTAG Then
        strReportImages = Nvl(rsTmp("报告图象"))
    
    
        '检查报告图象的长度，如果超过4000个字节，则提示无法保存图像
        If Len(strReportImages & " ;" & ImgTmp.InstanceUID & ".jpg") >= 4000 Then
            MsgBoxD Me, "图像数量超过上限，请先删除部分图像后，再继续采集图像。", vbInformation, gstrSysName
            Call ucPreview.DeleteImage(ucPreview.CurImageCount)
            Exit Sub
        End If
    
        '保存图像到缓存目录
        Select Case iEncode
            Case 1          'Run-Length Encoding行程压缩
                ImgTmp.WriteFile mstrBufferDir & dtReceived & "/" & mstrStudyUID & "/" & ImgTmp.InstanceUID, True, "1.2.840.10008.1.2.5"
            Case 2          '不处理，保持原图的压缩方式
                ImgTmp.WriteFile mstrBufferDir & dtReceived & "/" & mstrStudyUID & "/" & ImgTmp.InstanceUID, True
            Case Else       'Lossless JPEG encoding JPEG无损压缩
                ImgTmp.WriteFile mstrBufferDir & dtReceived & "/" & mstrStudyUID & "/" & ImgTmp.InstanceUID, True, "1.2.840.10008.1.2.4.70"
        End Select

        ImgTmp.FileExport mstrBufferDir & dtReceived & "/" & mstrStudyUID & "/" & ImgTmp.InstanceUID & ".jpg", "JPG", 80
    End If
    
    '不使用inet方式时，需要先初始化ftp连接
    lngResult = mobjFtpConnection.FuncFtpConnect(mobjFtp.strFTPIP, mobjFtp.strFTPUser, mobjFtp.strFTPPwd)
    If lngResult = 0 Then
        'FTP操作失败，提示错误，并删除缩略图中的图像
        MsgBoxD Me, "FTP连接失败，图像无法保存，请检查网络设置。", vbInformation, gstrSysName
        Call ucPreview.DeleteImage(ucPreview.CurImageCount)
    
        Exit Sub
    End If

    If Val(mobjBakFtp.strDeviceId) > 0 Then
        lngResult = mobjBakFtpConnection.FuncFtpConnect(mobjBakFtp.strFTPIP, mobjBakFtp.strFTPUser, mobjBakFtp.strFTPPwd)
        If lngResult = 0 Then
            mobjBakFtp.strDeviceId = ""
            MsgBoxD Me, "备份ftp设备连接失败，采集的图像将不能进行备份操作，如需备份请检查流程管理中的备份设备配置。", vbInformation, gstrSysName
        End If
    End If


    If ImgTmp.Tag.Tag <> VIDEOTAG And ImgTmp.Tag.Tag <> AUDIOTAG Then
        '保存dicom图像
        WriteToURL mstrBufferDir & dtReceived & "/" & mstrStudyUID & "/" & ImgTmp.InstanceUID, mobjFtp.strFtpDir & _
            dtReceived & "/" & mstrStudyUID & "/" & ImgTmp.InstanceUID
            
        WriteToURL mstrBufferDir & dtReceived & "/" & mstrStudyUID & "/" & ImgTmp.InstanceUID & ".jpg", mobjFtp.strFtpDir & _
            dtReceived & "/" & mstrStudyUID & "/" & ImgTmp.InstanceUID & ".jpg"
            
        '备份当前采集的图像
        If mobjBakFtpConnection.hConnection <> 0 Then
            BakImgToURL mstrBufferDir & dtReceived & "/" & mstrStudyUID & "/" & ImgTmp.InstanceUID, mobjBakFtp.strFtpDir & _
                dtReceived & "/" & mstrStudyUID & "/" & ImgTmp.InstanceUID
        End If
    Else
        '保存录像
        WriteToURL ImgTmp.Tag.VideoFile, mobjFtp.strFtpDir & dtReceived & "/" & mstrStudyUID & "/" & ImgTmp.InstanceUID

        '备份录像
        If mobjBakFtpConnection.hConnection <> 0 Then
            BakImgToURL ImgTmp.Tag.VideoFile, mobjBakFtp.strFtpDir & dtReceived & "/" & mstrStudyUID & "/" & ImgTmp.InstanceUID
        End If
        
        
        If ImgTmp.Tag.Tag = VIDEOTAG Then
            '将录像复制到对应的目录中，避免重新下载
            Call FileCopy(ImgTmp.Tag.VideoFile, mstrBufferDir & dtReceived & "/" & mstrStudyUID & "/" & ImgTmp.InstanceUID & ".avi")
            Call Kill(ImgTmp.Tag.VideoFile)
        
            ImgTmp.Tag.VideoFile = mstrBufferDir & dtReceived & "/" & mstrStudyUID & "/" & ImgTmp.InstanceUID & ".avi"
        ElseIf ImgTmp.Tag.Tag = AUDIOTAG Then
            '将音频文件复制到对应的目录中，避免重新下载
            Call FileCopy(ImgTmp.Tag.VideoFile, mstrBufferDir & dtReceived & "/" & mstrStudyUID & "/" & ImgTmp.InstanceUID & ".wav")
            Call Kill(ImgTmp.Tag.VideoFile)
        
            ImgTmp.Tag.VideoFile = mstrBufferDir & dtReceived & "/" & mstrStudyUID & "/" & ImgTmp.InstanceUID & ".wav"
        End If
    End If
    
    If mblnUseInetFtp Then
        '使用inetftp时，不需要每次都断开连接
    Else
        mobjFtpConnection.FuncFtpDisConnect
        
        If mobjBakFtpConnection.hConnection <> 0 Then mobjBakFtpConnection.FuncFtpDisConnect
    End If
    

    '图像存储成功后，存储数据库信息
    On Error GoTo DBError
    arrSQL = Array()
    
    If blnFirstImage Then
        gstrSQL = "ZL_影像检查记录_SET(" & mlngAdviceID & "," & lngSendNO & ",'" & _
            mstrStudyUID & "',null," & _
            "to_Date('" & Format(nowTime, "yyyy-mm-dd hh:mm") & "','YYYY-MM-DD HH24:MI:SS'),'" & mobjFtp.strDeviceId & "')"
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = gstrSQL
    End If
    
    
    gstrSQL = "Select 序列UID From 影像检查序列  Where 序列UID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "PACS图像保存", CStr(ImgTmp.SeriesUID))
    
    '插入新的检查序列,如果为录像，则插入新的序列
    If rsTmp.EOF Or ImgTmp.Tag.Tag = VIDEOTAG Or ImgTmp.Tag.Tag = AUDIOTAG Then
        If ImgTmp.Tag.Tag = VIDEOTAG Then
            gstrSQL = "ZL_影像序列_INSERT('" & mstrStudyUID & "','" & ImgTmp.SeriesUID & "','视频录像',0)"
        ElseIf ImgTmp.Tag.Tag = AUDIOTAG Then
            gstrSQL = "ZL_影像序列_INSERT('" & mstrStudyUID & "','" & ImgTmp.SeriesUID & "','音频数据',0)"
        Else
            gstrSQL = "ZL_影像序列_INSERT('" & mstrStudyUID & "','" & ImgTmp.SeriesUID & "','" & ImgTmp.SeriesDescription & "',0)"
        End If
        
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = gstrSQL
    End If
    
    
    
    If ImgTmp.Tag.Tag <> VIDEOTAG And ImgTmp.Tag.Tag <> AUDIOTAG Then
        '插入新的图像记录
        gstrSQL = "ZL_影像图象_INSERT('" & ImgTmp.InstanceUID & "','" & ImgTmp.SeriesUID & "',NULL,0, null, sysdate)"
    Else
        If ImgTmp.Tag.Tag = VIDEOTAG Then
            '插入新的视频记录
            gstrSQL = "ZL_影像图象_INSERT('" & ImgTmp.InstanceUID & "','" & ImgTmp.SeriesUID & "',Null,0" & _
            ",null,sysdate,null,null,null,null,null,null,null,null,null," & VIDEOTAG & ",'" & mstrEncoderName & "'," & ImgTmp.Tag.RecordTimeLen & ")"
        Else
            '插入新的音频记录
            gstrSQL = "ZL_影像图象_INSERT('" & ImgTmp.InstanceUID & "','" & ImgTmp.SeriesUID & "',Null,0" & _
            ",null,sysdate,null,null,null,null,null,null,null,null,null," & AUDIOTAG & ",''," & ImgTmp.Tag.RecordTimeLen & ")"
        End If
    End If
        
    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
    arrSQL(UBound(arrSQL)) = gstrSQL
    
    If ImgTmp.Tag.Tag <> VIDEOTAG And ImgTmp.Tag.Tag <> AUDIOTAG Then
        '如果不是检查图像，则不保存报告图
        gstrSQL = "ZL_影像检查报告_ADD('" & mstrStudyUID & "','" & ImgTmp.InstanceUID & ".jpg')"
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = gstrSQL
    End If
    
    gcnOracle.BeginTrans        '----------保存图像
    blnInTrans = True
    For i = 0 To UBound(arrSQL)
        Call zlDatabase.ExecuteProcedure(CStr(arrSQL(i)), "保存图像")
    Next i
    gcnOracle.CommitTrans
    blnInTrans = False
    
    If ImgTmp.Tag.Tag <> VIDEOTAG And ImgTmp.Tag.Tag <> AUDIOTAG Then
        Call frmCaptureHint.ShowCaptureHint( _
            IIf(mblnPoputWindowHint, mstrBufferDir & dtReceived & "/" & mstrStudyUID & "/" & ImgTmp.InstanceUID, ""), _
            mblnSoundHint, hpRB, Me)
    End If
    
    Exit Sub
DBError:
    '出错，则回退数据库操作，并且删除所采集的图像
    If blnInTrans = True Then gcnOracle.RollbackTrans
    err.Raise err.Number, "检查图像保存"
    Call ucPreview.DeleteImage(ucPreview.CurImageCount)
End Sub



Private Sub WriteToURL(ByVal SrcFileName As String, ByVal DestFileName As String)
'------------------------------------------------
'功能：将本地文件保存到远程网络上
'参数：SrcFileName--本地文件名，DestFileName－－远程文件名
'返回：无
'-----------------------------------------------
'功能：
    Dim objFileSystem As New Scripting.FileSystemObject
    Dim strPath As String

    '在FTP中创建目录
    strPath = objFileSystem.GetParentFolderName(DestFileName)
    mobjFtpConnection.FuncFtpMkDir "/", strPath
    
    '向FTP上传文件
    mobjFtpConnection.FuncUploadFile strPath, SrcFileName, objFileSystem.GetFileName(DestFileName)
End Sub


Private Sub BakImgToURL(ByVal SrcFileName As String, ByVal DestFileName As String)
'------------------------------------------------
'功能：备份图像到远程网络上
'参数：SrcFileName--本地文件名，DestFileName－－远程文件名
'返回：无
'-----------------------------------------------
    Dim objFileSystem As New Scripting.FileSystemObject
    Dim strPath As String
    
    If mobjBakFtpConnection.hConnection = 0 Then Exit Sub
    
    '在FTP中创建目录
    strPath = objFileSystem.GetParentFolderName(DestFileName)
    mobjBakFtpConnection.FuncFtpMkDir "/", strPath
    
    '向FTP上传文件
    mobjBakFtpConnection.FuncUploadFile strPath, SrcFileName, objFileSystem.GetFileName(DestFileName)
End Sub


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
    Dim strRegPath As String
    
    '-----------------------------------------------------
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    Me.cbrMain.VisualTheme = xtpThemeOffice2003
    Set Me.cbrMain.Icons = zlCommFun.GetPubIcons
    
    
    With Me.cbrMain.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '放在VisualTheme后有效
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 32, 32
    End With
    Me.cbrMain.EnableCustomization False
    Me.cbrMain.ActiveMenuBar.Visible = False
    
    strRegPath = "公共模块\" & App.ProductName & "\frmVideoCapture"
    
    '是否显示处理工具栏
    mblnShowProcessBar = GetSetting("ZLSOFT", strRegPath, "显示处理工具栏", "True")
    
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
        
        If mlngModul = G_LNG_VIDEOSTATION_MODULE Then
            Set cbrControl = .Add(xtpControlButton, conMenu_Cap_After_Capture, "后台采集")
                cbrControl.ToolTipText = "后台采集"
                cbrControl.IconId = 10020
        End If
        
        '在非TWAIN采集模式的情况下，才显示该按钮
        'If Not GetIsTwainCaptureWay Then
            Set cbrControl = .Add(xtpControlButton, conMenu_Cap_Record, "录像"): cbrControl.ToolTipText = "开始录像"
                cbrControl.Enabled = True
                
            If mlngModul = G_LNG_VIDEOSTATION_MODULE Then
                Set cbrControl = .Add(xtpControlButton, conMenu_Cap_After_Record, "后台录像")
                    cbrControl.ToolTipText = "后台录像"
                    cbrControl.IconId = 10021
            End If
            
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
            
            
        If mlngModul = G_LNG_VIDEOSTATION_MODULE Then
            Set cbrControl = .Add(xtpControlButton, conMenu_Cap_OpenStudyList, "打开检查"): cbrControl.ToolTipText = "打开检查"
            cbrControl.BeginGroup = True
            
            Set cbrControl = .Add(xtpControlButton, conMenu_Cap_StudySyncState, "锁定检查"): cbrControl.ToolTipText = "锁定检查"
            cbrControl.IconId = 10012
            
            
            Set cbrControl = .Add(xtpControlButton, conMenu_Cap_After_Tag, "标记检查")
                cbrControl.ToolTipText = "标记检查"
                cbrControl.IconId = 10022
        End If
        
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
        Set cbrControl = .Add(xtpControlButton, conMenu_Tool_Analyse, "高级"): cbrControl.ToolTipText = "高级处理"
        cbrControl.BeginGroup = True
    End With
    
    For Each cbrControl In cbrToolBar.Controls
        cbrControl.Style = xtpButtonIcon
        cbrControl.Category = "处理"
        cbrControl.Enabled = False
    Next
    cbrToolBar.Visible = mblnShowProcessBar
End Sub


Public Sub ShowFrameSelectImagePopup()
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
    Dim imgs As New DicomImages
    Dim iPlane As Integer
    Dim dblZoom As Double
    Dim iLeft As Integer
    Dim iRight As Integer
    Dim iTop As Integer
    Dim iBottom As Integer
    Dim iMax As Integer
    Dim img As DicomImage
    Dim lblFrame As DicomLabel
    
    If Me.dcmView.Images.Count <> 1 Then Exit Sub
    If Me.dcmView.Images(1).Labels.Count < 1 Then Exit Sub
    
    Set img = Me.dcmView.Images(1)
    Set lblFrame = Me.dcmView.Images(1).Labels(Me.dcmView.Images(1).Labels.Count)
    
    If Abs(lblFrame.Width) = 0 Or Abs(lblFrame.Height) = 0 Then
        MsgBoxD Me, "请选择图像区域后再保存", vbExclamation, gstrSysName
        Exit Sub
    End If
    
    '图象最大宽高=300
    iMax = 300
    
    '根据label来提取被框选中的图像
    '图象位数,黑白图像为1，彩色图像为3
    iPlane = 1
    If Not IsNull(img.Attributes(&H28, &H4).value) And img.Attributes(&H28, &H4).Exists Then
        If img.Attributes(&H28, &H4).value = "RGB" Then
            iPlane = 3
        End If
    End If
    
    '图象框的位置
    If lblFrame.Width >= 0 Then
        iLeft = lblFrame.Left
        iRight = iLeft + lblFrame.Width
    Else
        iLeft = lblFrame.Left + lblFrame.Width
        iRight = lblFrame.Left
    End If
    
    If lblFrame.Height >= 0 Then
        iTop = lblFrame.Top
        iBottom = iTop + lblFrame.Height
    Else
        iTop = lblFrame.Top + lblFrame.Height
        iBottom = lblFrame.Top
    End If
    
    '控制结果图象的大小在300*300之内
    If (iRight - iLeft) > iMax Or (iBottom - iTop) > iMax Then
        dblZoom = iMax / (iRight - iLeft)
        If dblZoom > iMax / (iBottom - iTop) Then dblZoom = iMax / (iBottom - iTop)
    Else
        dblZoom = 1
    End If
    
    img.Labels(img.Labels.Count).Visible = False
    If (img.RotateState = doRotateLeft And img.FlipState = doFlipNormal) _
        Or (img.RotateState = doRotateRight And img.FlipState = doFlipBoth) _
        Or (img.RotateState = doRotate180 And img.FlipState = doFlipVertical) _
        Or (img.RotateState = doRotateNormal And img.FlipState = doFlipHorizontal) Then
        'X方向对调
        Set imgResult = img.PrinterImage(8, iPlane, True, dblZoom, img.SizeX - iRight, img.SizeX - iLeft, iTop, iBottom)
    ElseIf (img.RotateState = doRotateLeft And img.FlipState = doFlipBoth) _
        Or (img.RotateState = doRotateRight And img.FlipState = doFlipNormal) _
        Or (img.RotateState = doRotate180 And img.FlipState = doFlipHorizontal) _
        Or (img.RotateState = doRotateNormal And img.FlipState = doFlipVertical) Then
        'Y方向对调
        Set imgResult = img.PrinterImage(8, iPlane, True, dblZoom, iLeft, iRight, img.SizeY - iBottom, img.SizeY - iTop)
    ElseIf (img.RotateState = doRotateRight And img.FlipState = doFlipHorizontal) _
        Or (img.RotateState = doRotateLeft And img.FlipState = doFlipVertical) _
        Or (img.RotateState = doRotate180 And img.FlipState = doFlipNormal) _
        Or (img.RotateState = doRotateNormal And img.FlipState = doFlipBoth) Then
        'X，Y方向对调
        Set imgResult = img.PrinterImage(8, iPlane, True, dblZoom, img.SizeX - iRight, img.SizeX - iLeft, img.SizeY - iBottom, img.SizeY - iTop)
    Else
        Set imgResult = img.PrinterImage(8, iPlane, True, dblZoom, iLeft, iRight, iTop, iBottom)
    End If
    
    '给imgResult一个唯一的 InstanceUID
    imgResult.InstanceUID = dcmglbUID.NewUID
    
    '把结果图加入到viewer中并且保存
    '设置图像的DICOM参数
    subWriteDicomPara imgResult, mlngAdviceID
    
    Dim dcmTag As New clsImageTagInf
    dcmTag.Tag = IMGTAG
    
    Set imgResult.Tag = dcmTag
    
    '将图像插入到缩略图中
    subInsert2Mini imgResult
    
    '保存图像，并触发图像存储事件
    Call subSaveImage
    
    If ucPreview.CurImageCount = 1 Then
        Call SendMsgToMainWindow(Me, wetCaptureFirstImg, mlngAdviceID)
    End If
End Sub


Private Sub ucCapHook_OnKeyBoardLHook(ByVal lngVkCode As Long, ByVal lngScanCode As Long, ByVal lngFlags As Long)
On Error GoTo errHandle
    Select Case lngVkCode
        Case 66
            '判断键盘按键是否松开，为0表示按下键盘
            If lngScanCode = 128 Then
                '执行快捷采集
'                Call CaptureImage

                If timerHook.Enabled Or mblnCurCaptureState Then Exit Sub
                timerHook.Enabled = True
            End If
    End Select
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub ucPreview_OnClick(ByVal lngSelectedIndex As Long)

    mCorpSize.X = 0
    mCorpSize.Y = 0
    
    '被选中图像显示红框
    If lngSelectedIndex <> 0 Then
        
        '将被选中图像装载到dcmView中
        dcmView.Images.Clear
        dcmView.Images.Add ucPreview.ImgViewer.Images(lngSelectedIndex)

        '显示dcmView，隐藏picVideo
        dcmView.CurrentImage.BorderWidth = 0
        
        '使图像居中显示，并可以自由拖动图像
        Dim dblTempZoom As Double
              
        dblTempZoom = dcmView.CurrentImage.ActualZoom
        dcmView.CurrentImage.StretchToFit = False
              
        Call subCenterZoom(dcmView.CurrentImage, dcmView, dblTempZoom, mCorpSize)
        
        '设置视频的当前显示状态
        Call ConfigVideoShowState(False)
    End If
    
    '恢复当前控件焦点，以便能够滚动图像
    ucPreview.SetFocus
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
    If ErrCenter() = 1 Then Resume
End Sub

