VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsVfwCaptureParameterCfg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


'视频显示类型
Public Enum TShowWay
  swNormal = 0 '原大小显示
  swStretch = 1 '拉伸显示
  swFit = 2  '自动适应
  swAutoFitCut = 3 '按裁剪范围显示
  swWindowAutoFit = 4 '自适应视频大小
End Enum


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'采集相关参数设置
'
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private mlngCaptureDeviceIndex            '采集设备索引
Private mlngCapBit As Long                '采集位数设置
Private mlngCapVideoWidth As Long         '采集视频宽度
Private mlngCapVideoHeight As Long        '采集视频高度
Private mlngCapPreviewRate As Long        '采集预览速度
Private mblnIsScale As Boolean            '是否自动适应
Private mblnCaptureAudio As Boolean       '是否采集音频
Private mblnAbortLeftMouse As Boolean     '是否左键中止采集
Private mblnAbortRightMouse As Boolean    '是否右键中止采集
Private mblnLimitEnabled As Boolean      '是否设置时间限制
Private mlngAVStreamMaster As Long        '音频同步方式
Private mlngRequestMicroSecPerFrame As Long '采集帧频率
Private mlngVideoShowWay As TShowWay        '视频显示方式 (0： 原大小显示,1：拉伸显示,2：按比例缩放，3：按裁剪范围缩放，4：自适应视频大小)

Private mdblLeftCutRate As Double        '左裁剪比率
Private mdblTopCutRate As Double       '
Private mdblWidthCutRate As Double       '宽度裁剪比率
Private mdblHeightCutRate As Double      '

Private Const M_LNG_DEFAULT_MAX_CAPDEV_COUNT As Long = 20  '默认最大采集设备数量
Private Const M_LNG_DEFAULT_CAPBIT As Long = 24            '默认采集位数

Private Const M_STR_DEFAULT_CONFIG_NAME As String = "VfwCapture.ini" '默认采集参数的配置文件名称




''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'载入默认参数配置
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub LoadDefaultCfg()
  mlngCaptureDeviceIndex = 0
  mlngCapBit = M_LNG_DEFAULT_CAPBIT
  mlngCapVideoWidth = 320
  mlngCapVideoHeight = 240
  mlngCapPreviewRate = 66
  mblnIsScale = True
  mblnCaptureAudio = False
  mblnAbortLeftMouse = True
  mblnAbortRightMouse = True
  mblnLimitEnabled = True
  mlngAVStreamMaster = AVSTREAMMASTER_NONE
  mlngRequestMicroSecPerFrame = 15
  mlngVideoShowWay = 0
  
  mdblHeightCutRate = 0
  mdblWidthCutRate = 0
  mdblTopCutRate = 0
  mdblLeftCutRate = 0
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'克隆对象
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub CloneTo(ByRef curCaptureParameter As clsVfwCaptureParameterCfg)
  curCaptureParameter.AbortLeftMouse = Me.AbortLeftMouse
  curCaptureParameter.AbortRightMouse = Me.AbortRightMouse
  curCaptureParameter.AVStreamMaster = Me.AVStreamMaster
  curCaptureParameter.CapBit = Me.CapBit
  curCaptureParameter.CapPreviewRate = Me.CapPreviewRate
  curCaptureParameter.CaptureAudio = Me.CaptureAudio
  curCaptureParameter.CaptureDeviceIndex = Me.CaptureDeviceIndex
  curCaptureParameter.CapVideoHeight = Me.CapVideoHeight
  curCaptureParameter.CapVideoWidth = Me.CapVideoWidth
  curCaptureParameter.IsScale = Me.IsScale
  curCaptureParameter.LimitEnabled = Me.LimitEnabled
  curCaptureParameter.RequestMicroSecPerFrame = Me.RequestMicroSecPerFrame
  curCaptureParameter.VideoShowWay = Me.VideoShowWay
  
  curCaptureParameter.LeftCutRate = Me.LeftCutRate
  curCaptureParameter.TopCutRate = Me.TopCutRate
  curCaptureParameter.WidthCutRate = Me.WidthCutRate
  curCaptureParameter.HeightCutRate = Me.HeightCutRate
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'保存参数配置
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SaveParameterCfg(ByVal sFileName As String)
  Dim strCurrentFile As String
  Dim iniFile As clsIniFile
  
  strCurrentFile = sFileName
  
  If Trim(strCurrentFile) = "" Then
    strCurrentFile = App.Path & "\" & M_STR_DEFAULT_CONFIG_NAME
  End If
  
  Set iniFile = New clsIniFile
  
  Call iniFile.SetIniFile(strCurrentFile)
  
  Call iniFile.WriteValue("VfwCaptureConfig", "AbortLeftMouse", CStr(Me.AbortLeftMouse))
  Call iniFile.WriteValue("VfwCaptureConfig", "AbortRightMouse", CStr(Me.AbortRightMouse))
  Call iniFile.WriteValue("VfwCaptureConfig", "AVStreamMaster", CStr(Me.AVStreamMaster))
  Call iniFile.WriteValue("VfwCaptureConfig", "CapBit", CStr(Me.CapBit))
  Call iniFile.WriteValue("VfwCaptureConfig", "CapPreviewRate", CStr(Me.CapPreviewRate))
  Call iniFile.WriteValue("VfwCaptureConfig", "CaptureAudio", CStr(Me.CaptureAudio))
  Call iniFile.WriteValue("VfwCaptureConfig", "CaptureDeviceIndex", CStr(Me.CaptureDeviceIndex))
  Call iniFile.WriteValue("VfwCaptureConfig", "CapVideoHeight", CStr(Me.CapVideoHeight))
  Call iniFile.WriteValue("VfwCaptureConfig", "CapVideoWidth", CStr(Me.CapVideoWidth))
  Call iniFile.WriteValue("VfwCaptureConfig", "IsScale", CStr(Me.IsScale))
  Call iniFile.WriteValue("VfwCaptureConfig", "LimitEnabled", CStr(Me.LimitEnabled))
  Call iniFile.WriteValue("VfwCaptureConfig", "RequestMicroSecPerFrame", CStr(Me.RequestMicroSecPerFrame))
  Call iniFile.WriteValue("VfwCaptureConfig", "VideoShowWay", CStr(Me.VideoShowWay))
  
  Call iniFile.WriteValue("VfwCaptureConfig", "LeftCutRate", CStr(Me.LeftCutRate))
  Call iniFile.WriteValue("VfwCaptureConfig", "TopCutRate", CStr(Me.TopCutRate))
  Call iniFile.WriteValue("VfwCaptureConfig", "WidthCutRate", CStr(Me.WidthCutRate))
  Call iniFile.WriteValue("VfwCaptureConfig", "HeightCutRate", CStr(Me.HeightCutRate))
  
  Set iniFile = Nothing
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'读取参数配置
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ReadParameterCfg(ByVal sFileName As String)
  Dim strCurrentFile As String
  Dim iniFile As clsIniFile
  
  strCurrentFile = sFileName
  
  If Trim(strCurrentFile) = "" Then
    strCurrentFile = App.Path & "\" & M_STR_DEFAULT_CONFIG_NAME
  End If
  
  Set iniFile = New clsIniFile
  
  Call iniFile.SetIniFile(strCurrentFile)
  
  Me.AbortLeftMouse = CBool(iniFile.ReadValue("VfwCaptureConfig", "AbortLeftMouse", CStr(True)))
  Me.AbortRightMouse = CBool(iniFile.ReadValue("VfwCaptureConfig", "AbortRightMouse", CStr(True)))
  Me.AVStreamMaster = CLng(iniFile.ReadValue("VfwCaptureConfig", "AVStreamMaster", CStr(0)))
  Me.CapBit = CLng(iniFile.ReadValue("VfwCaptureConfig", "CapBit", CStr(24)))
  Me.CapPreviewRate = CLng(iniFile.ReadValue("VfwCaptureConfig", "CapPreviewRate", CStr(66)))
  Me.CaptureAudio = CBool(iniFile.ReadValue("VfwCaptureConfig", "CaptureAudio", CStr(False)))
  Me.CaptureDeviceIndex = CLng(iniFile.ReadValue("VfwCaptureConfig", "CaptureDeviceIndex", CStr(0)))
  Me.CapVideoHeight = CLng(iniFile.ReadValue("VfwCaptureConfig", "CapVideoHeight", CStr(240)))
  Me.CapVideoWidth = CLng(iniFile.ReadValue("VfwCaptureConfig", "CapVideoWidth", CStr(320)))
  Me.IsScale = CBool(iniFile.ReadValue("VfwCaptureConfig", "IsScale", CStr(True)))
  Me.LimitEnabled = CBool(iniFile.ReadValue("VfwCaptureConfig", "LimitEnabled", CStr(False)))
  Me.RequestMicroSecPerFrame = CLng(iniFile.ReadValue("VfwCaptureConfig", "RequestMicroSecPerFrame", CStr(15)))
  Me.VideoShowWay = CLng(iniFile.ReadValue("VfwCaptureConfig", "VideoShowWay", CStr(0)))
  
  Me.LeftCutRate = CDbl(iniFile.ReadValue("VfwCaptureConfig", "LeftCutRate", CStr(0)))
  Me.TopCutRate = CDbl(iniFile.ReadValue("VfwCaptureConfig", "TopCutRate", CStr(0)))
  Me.WidthCutRate = CDbl(iniFile.ReadValue("VfwCaptureConfig", "WidthCutRate", CStr(0)))
  Me.HeightCutRate = CDbl(iniFile.ReadValue("VfwCaptureConfig", "HeightCutRate", CStr(0)))
  
  Set iniFile = Nothing
End Sub

'//采集设备索引
Property Get CaptureDeviceIndex() As Long
  CaptureDeviceIndex = mlngCaptureDeviceIndex
End Property

Property Let CaptureDeviceIndex(value As Long)
  mlngCaptureDeviceIndex = value
End Property



'//颜色深度
Property Get CapBit() As Long
  CapBit = mlngCapBit
End Property

Property Let CapBit(value As Long)
  mlngCapBit = value
End Property


'//取得视频宽度
Property Get CapVideoWidth() As Long
  CapVideoWidth = mlngCapVideoWidth
End Property

Property Let CapVideoWidth(value As Long)
  mlngCapVideoWidth = value
End Property


'//取得视频高度
Property Get CapVideoHeight() As Long
  CapVideoHeight = mlngCapVideoHeight
End Property

Property Let CapVideoHeight(value As Long)
  mlngCapVideoHeight = value
End Property


'//取得预览速度
Property Get CapPreviewRate() As Long
  CapPreviewRate = mlngCapPreviewRate
End Property

Property Let CapPreviewRate(value As Long)
  mlngCapPreviewRate = value
End Property


'//是否自动适应
Property Get IsScale() As Boolean
  IsScale = mblnIsScale
End Property

Property Let IsScale(value As Boolean)
  mblnIsScale = value
End Property


'//是否采集音频
Property Get CaptureAudio() As Boolean
  CaptureAudio = mblnCaptureAudio
End Property

Property Let CaptureAudio(value As Boolean)
  mblnCaptureAudio = value
End Property


'//左键是否终止采集
Property Get AbortLeftMouse() As Boolean
  AbortLeftMouse = mblnAbortLeftMouse
End Property

Property Let AbortLeftMouse(value As Boolean)
  mblnAbortLeftMouse = value
End Property


'//右键是否终止采集
Property Get AbortRightMouse() As Boolean
  AbortRightMouse = mblnAbortRightMouse
End Property

Property Let AbortRightMouse(value As Boolean)
  mblnAbortRightMouse = value
End Property
  
  
'//是否有时间限制
Property Get LimitEnabled() As Boolean
  LimitEnabled = mblnLimitEnabled
End Property

Property Let LimitEnabled(value As Boolean)
  mblnLimitEnabled = value
End Property

  
'//音频同步方式
Property Get AVStreamMaster() As Long
  AVStreamMaster = mlngAVStreamMaster
End Property

Property Let AVStreamMaster(value As Long)
  mlngAVStreamMaster = value
End Property

  
'//采集帧频率
Property Get RequestMicroSecPerFrame() As Long
  RequestMicroSecPerFrame = mlngRequestMicroSecPerFrame
End Property

Property Let RequestMicroSecPerFrame(value As Long)
  mlngRequestMicroSecPerFrame = value
End Property


'视频显示方式
Property Get VideoShowWay() As Long
  VideoShowWay = mlngVideoShowWay
End Property

Property Let VideoShowWay(value As Long)
  mlngVideoShowWay = value
End Property


'图象裁剪设置
Property Get HeightCutRate() As Double
  HeightCutRate = mdblHeightCutRate
End Property
   
Property Let HeightCutRate(value As Double)
  mdblHeightCutRate = value
End Property

   
Property Get WidthCutRate() As Double
  WidthCutRate = mdblWidthCutRate
End Property

Property Let WidthCutRate(value As Double)
  mdblWidthCutRate = value
End Property


Property Get TopCutRate() As Double
  TopCutRate = mdblTopCutRate
End Property

Property Let TopCutRate(value As Double)
  mdblTopCutRate = value
End Property


Property Get LeftCutRate() As Double
  LeftCutRate = mdblLeftCutRate
End Property

Property Let LeftCutRate(value As Double)
  mdblLeftCutRate = value
End Property
