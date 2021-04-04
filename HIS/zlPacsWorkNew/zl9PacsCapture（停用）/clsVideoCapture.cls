VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsVideoCapture"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*************************************************************************************************************
'*
'*  功能：采集包装对象（对vfw，wdm以及其他视频驱动类型进行支持）
'*  创建时间：2010-01-19
'*  创建人：TJH
'*
'*************************************************************************************************************





Option Explicit






'视频驱动类型
Public Enum TVideoDriverType
  vdtWDM = 0
  vdtVFW = 1
  vdtTWAIN = 2
  vdtCustom = 3
  '其他需要支持的驱动类型......
End Enum









Private mVideoDriverType As TVideoDriverType '保存当前要使用的驱动类型

Private mVfwWindow As PictureBox  'vfw的视频显示窗口
Private mVfwControl As clsVfwCapture 'vfw的视频控制对象

Private mWdmControl As ZLDSVideoProcess.DSCapture  'wdm的视频采集组件

Private mCustomControl As Object    '专用视频采集组件

Private mblnIsAllowStartupVideo As Boolean '是否允许启动视频，当为FALSE时，该对象的所有方法将不能使用







''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'对象属性定义
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Property Get VideoDriverType() As TVideoDriverType
  VideoDriverType = mVideoDriverType
End Property

Property Get Visible() As Boolean
  Visible = mWdmControl.Visible Or mVfwWindow.Visible
End Property

Property Let WdmVisible(value As Boolean)
  mWdmControl.Visible = True
End Property

Property Let VfwVisible(value As Boolean)
  mVfwWindow.Visible = value
End Property

Property Let VideoDriverType(value As TVideoDriverType)
  mVideoDriverType = value
  
  '设置采集组件的显示状态
  Select Case value
    Case vdtWDM
      mWdmControl.Visible = True
      mVfwWindow.Visible = False
    Case vdtVFW
      mVfwWindow.Visible = True
      mWdmControl.Visible = False
    'Case ...
  End Select
End Property

'判断采集组件是否启动
Property Get IsStartup() As Boolean
  Select Case mVideoDriverType
    Case vdtWDM
      IsStartup = mWdmControl.PreviewState
    Case vdtVFW
      IsStartup = mVfwControl.hWnd <> 0
    Case Else
      IsStartup = False
  End Select
End Property

Property Get IsAllowStartupVideo() As Boolean
  IsAllowStartupVideo = mblnIsAllowStartupVideo
End Property

Property Let IsAllowStartupVideo(value As Boolean)
  mblnIsAllowStartupVideo = value
End Property


Property Get VideoSize() As TVideoSize

  '取得视频宽度
  Select Case mVideoDriverType
    Case vdtVFW
      Dim vfwParameter As clsVfwParameterCfg
      
      Set vfwParameter = mVfwControl.GetCaptureParameter()
      
      VideoSize.Width = vfwParameter.CapVideoWidth
      VideoSize.Height = vfwParameter.CapVideoHeight
      
    Case vdtWDM
      Dim wdmParameter As TCaptureParameter
      Dim wdmParameterEnum As DSCapParameterEnum
      Dim wdmVideoSize As TVideoSize
            
      Set wdmParameterEnum = New DSCapParameterEnum
            
      Call mWdmControl.GetCaptureParameter(wdmParameter)
      
      VideoSize = wdmParameterEnum.VideoSizeConvert(wdmParameter.VideoSize)
    Case Else
      VideoSize.Width = 800
      VideoSize.Height = 600
  End Select
  
End Property












''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'对象方法定义
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''



'开始视频预览
Public Sub StartPreview()

  '不允许启动视频，则直接退出
  If Not mblnIsAllowStartupVideo Then
    Exit Sub
  End If
  
  Select Case mVideoDriverType
    Case vdtWDM 'wdm的视频采集
      Call mWdmControl.StartPreview
      Call mWdmControl.RefreshWindow
    Case vdtVFW 'vfw的视频采集
      Call mVfwControl.StartPreview(mVfwWindow.hWnd)
      Call mVfwControl.UpdateCaptureWindowPos(mVfwWindow.ScaleWidth, mVfwWindow.ScaleHeight)
    'Case ...
  End Select
  
End Sub


'停止视频预览
Public Sub StopPreview()

  Select Case mVideoDriverType
    Case vdtWDM 'wdm的视频采集
      Call mWdmControl.StopPreview
    Case vdtVFW 'vfw的视频采集
      Call mVfwControl.StopPreview
    'Case   ...
  End Select
  
End Sub


'全屏显示
Public Sub FullScreen(owner As Object, ByVal ownerHandle As Long)
  
  '不允许启动视频，则直接退出
  If Not mblnIsAllowStartupVideo Then
    Exit Sub
  End If
  
  Select Case mVideoDriverType
    Case vdtWDM 'wdm的视频采集
      mWdmControl.ShowModel = smFit
      
      Call mWdmControl.ShowFullScreen(0, GetMonitorIndex(ownerHandle))
      
      mWdmControl.ShowModel = smStretch
    Case vdtVFW 'vfw的视频采集
      'Dim vfwParameter As clsVfwCaptureParameterCfg
      'Set vfwParameter = mVfwControl.GetCaptureParameter()
      
      'vfwParameter.VideoShowWay = swFit
      
      'Call mVfwControl.SetCaptureParameter(vfwParameter)
      
      mVfwControl.CaptureParameterInf.VideoShowWay = swFit
      
      Call mVfwControl.ShowFullScreen(owner, mVfwWindow, GetMonitorIndex(ownerHandle))
      
      mVfwControl.CaptureParameterInf.VideoShowWay = swStretch
      'vfwParameter.VideoShowWay = swStretch
      
      'Call mVfwControl.SetCaptureParameter(vfwParameter)
    'Case ...
  End Select
  
End Sub


'刷新视频窗口
Public Sub RefreshVideoWindow()

  '不允许启动视频，则直接退出
  If Not mblnIsAllowStartupVideo Then
    Exit Sub
  End If
  
  Select Case mVideoDriverType
    Case vdtWDM 'wdm的视频采集
      Call mWdmControl.RefreshWindow
    Case vdtVFW 'vfw的视频采集
      Call mVfwControl.UpdateCaptureWindowPos(mVfwWindow.ScaleWidth, mVfwWindow.ScaleHeight)
    'Case ...
  End Select
  
End Sub


'采集图象到文件
Public Sub CaptureImgToFile(ByVal strFileName As String)

  '不允许启动视频，则直接退出
  If Not mblnIsAllowStartupVideo Then
    Exit Sub
  End If
  
  '采集图像
  Select Case mVideoDriverType
    Case vdtVFW
      Call mVfwControl.CaptureBmpImgToFile(strFileName)
    Case vdtWDM
      Call mWdmControl.CaptureBmpImageToFile(strFileName)
    'Case Other......
  End Select
  
End Sub


'将采集图像保存到剪贴板
Public Function CaptureImgToClipBoard() As Boolean
  '不允许启动视频，则直接退出
  If Not mblnIsAllowStartupVideo Then
    CaptureImgToClipBoard = False
    Exit Function
  End If
  
  On Error GoTo errHandle
    
    '采集图像
    Select Case mVideoDriverType
        Case vdtVFW
            CaptureImgToClipBoard = mVfwControl.CaptureImageToClipBoard
        Case vdtWDM
            CaptureImgToClipBoard = (mWdmControl.CaptureImgToClipBoard = "")
            'Case Other......
    End Select
  
    
  Exit Function
errHandle:
  CaptureImgToClipBoard = False
End Function


Public Function CaptureImageFromMemory() As IPictureDisp
  '不允许启动视频，则直接退出
  If Not mblnIsAllowStartupVideo Then
    Set CaptureImageFromMemory = Nothing
    Exit Function
  End If
  
  On Error GoTo errHandle
    
    Dim iPic As IPictureDisp
    Set iPic = Nothing
    
    '采集图像
    Select Case mVideoDriverType
        Case vdtVFW
        
            Clipboard.Clear
            
            Call mVfwControl.CaptureImageToClipBoard
            
            Set iPic = Clipboard.GetData(2)
            
            Clipboard.Clear
        Case vdtWDM
            Set iPic = mWdmControl.CaptureBmpImage
        'Case Other......
    End Select
  
    Set CaptureImageFromMemory = iPic
    
  Exit Function
errHandle:
  Set CaptureImageFromMemory = Nothing
End Function


Public Function CaptureImg() As IPictureDisp
  '不允许启动视频，则直接退出
  If Not mblnIsAllowStartupVideo Then
    Set CaptureImg = Nothing
    Exit Function
  End If
  
  On Error GoTo errHandle
  
    Dim strTempFile As String
  
    '创建采集目录
    If Dir(App.Path & "\TmpImage\", vbDirectory) = "" Then
      Call MkDir(App.Path & "\TmpImage\")
    End If
    strTempFile = App.Path & "\TmpImage\" & "DAEAE94F2CBC46018D6F1496C00447E3.dat"
    
  
    '采集图像
    Select Case mVideoDriverType
        Case vdtVFW
            Call mVfwControl.CaptureBmpImgToFile(strTempFile)
        Case vdtWDM
            Call mWdmControl.CaptureBmpImageToFile(strTempFile)
            'Case Other......
    End Select
  
    If Dir(strTempFile) <> "" Then
        '从文件载入图像
        Set CaptureImg = LoadPicture(strTempFile)
  
        '删除采集的图像文件
        Call Kill(strTempFile)
    Else
        Set CaptureImg = Nothing
    End If
  
  Exit Function
errHandle:
  Set CaptureImg = Nothing
End Function


'开始录像
Public Function StartVideo(ByVal strFileName As String) As String
  StartVideo = ""
  
  '不允许启动视频，则直接退出
  If Not mblnIsAllowStartupVideo Then
    StartVideo = "不允许启动当前视频源。"
    Exit Function
  End If
  
  '开始视频录像
  Select Case mVideoDriverType
    Case vdtVFW
      Call mVfwControl.StartVideo(strFileName)
    Case vdtWDM
      mWdmControl.IsShowState = True
      
      StartVideo = mWdmControl.StartCaptureVideo(strFileName)
      
      If Trim(StartVideo) <> "" Then mWdmControl.IsShowState = False
    'Case Other......
  End Select
  
End Function


'停止录像
Public Sub StopVideo()

  '不允许启动视频，则直接退出
  If Not mblnIsAllowStartupVideo Then
    Exit Sub
  End If

  Dim strTempFileName As String
  
  '停止视频录像
  Select Case mVideoDriverType
    Case vdtVFW
      Call mVfwControl.StopVideo
    Case vdtWDM
      Call mWdmControl.StopCaptureVideo(strTempFileName)
      
      mWdmControl.IsShowState = False
    'Case Other......
  End Select
  
End Sub


'显示采集参数配置窗口
Public Sub ShowCaptureParameterCfgDialog(objOwner As Object, Optional hWndParent As Long = 0)

  '不允许启动视频，则直接退出
  If Not mblnIsAllowStartupVideo Then
    Exit Sub
  End If

  Select Case mVideoDriverType
    Case vdtWDM 'wdm的视频采集
      mWdmControl.HideCfgItem = hciVideoShowWay + hciVideoState + hciImageCapture
      
      Call mWdmControl.ShowCaptureParameterCfgDialog(hWndParent)
    Case vdtVFW 'vfw的视频采集
      Call mVfwControl.ShowCaptureParameterCfgDialog(objOwner, vpiShowWay)
    Case vdtCustom
        If Not mCustomControl Is Nothing Then Call mCustomControl.zlConfigDevice
    'Case Other......
  End Select
End Sub


'获取视频录像的编码器名称
Function GetEncoderName() As String
    Dim curCap As TCaptureParameter
    
    GetEncoderName = ""
    If mVideoDriverType <> vdtWDM Then Exit Function
    
    Call mWdmControl.GetCaptureParameter(curCap)
    GetEncoderName = curCap.EncoderName
End Function


'获取视频录像的录制长度
Function GetTimeLen() As Long
    GetTimeLen = 0
    If mVideoDriverType <> vdtWDM Then Exit Function
    
    GetTimeLen = mWdmControl.RecordTimeLen
End Function


'连接VFW设备对象
Public Sub ConnectedVfwDeviceObj(ByRef vfwWindow As Object)
   Set mVfwWindow = vfwWindow
End Sub


'连接WDM设备对象
Public Sub ConnectedWdmDeviceObj(ByRef wdmControl As Object)
  Set mWdmControl = wdmControl
End Sub

'连接专用视频采集设备对象
Public Sub ConnectedCustomDeviceObj(ByRef customControl As Object)
    Set mCustomControl = customControl
End Sub

'连接指定采集卡对象
Public Sub ConnectedTrdDeviceObj()
  '该方法目前不需要实现......
  Call err.Raise(-1, "ConnectedTrdDeviceObj", "不支持该方法。")
End Sub


'设置视频显示模式
Public Sub SetVideoShowWay(ByVal curShowWay As Long)
    Select Case curShowWay
      Case swNormal
        mWdmControl.ShowModel = smNormal
        mVfwControl.CaptureParameterInf.VideoShowWay = swNormal
      Case swFit
        mWdmControl.ShowModel = smFit
        mVfwControl.CaptureParameterInf.VideoShowWay = swFit
      Case swStretch
        mWdmControl.ShowModel = smStretch
        mVfwControl.CaptureParameterInf.VideoShowWay = swStretch
      Case swAutoFitCut
        mWdmControl.ShowModel = smAutoFitCut
        mVfwControl.CaptureParameterInf.VideoShowWay = swAutoFitCut
      Case swWindowAutoFit
        mWdmControl.ShowModel = smWindAutoFit
        mVfwControl.CaptureParameterInf.VideoShowWay = swWindowAutoFit
      Case Else
        mWdmControl.ShowModel = smStretch
        mVfwControl.CaptureParameterInf.VideoShowWay = swStretch
    End Select
End Sub


'从配置文件读取采集参数
Public Sub ReadCaptureParameterFromFile(ByVal strCfgName As String)

  '读取vfw参数
  mVfwControl.ParameterCfgFileName = strCfgName
  Call mVfwControl.ReadVfwCaptureParameterFromFile
 

  '读取wdm参数
  mWdmControl.ParameterCfgFileName = strCfgName
  Call mWdmControl.ReadParameterFromFile
  
End Sub


'保存采集参数到文件
Public Sub SaveCaptureParameterToFile(ByVal strCfgName As String)
  
  '保存vfw参数
  mVfwControl.ParameterCfgFileName = strCfgName
  Call mVfwControl.SaveVfwCaptureParameterToFile
  
  '保存wdm参数
  mWdmControl.ParameterCfgFileName = strCfgName
  Call mWdmControl.SaveParameterToFile

End Sub


'对象初始化
Private Sub Class_Initialize()
  mblnIsAllowStartupVideo = True
  
  Set mVfwControl = New clsVfwCapture
    
  '将mVfwControl的显示模式设置为拉伸模式，以适应图像裁剪的调节
  Dim vfwParameter As clsVfwParameterCfg
  Set vfwParameter = mVfwControl.GetCaptureParameter()
  
  vfwParameter.VideoShowWay = swStretch
  
  Call mVfwControl.SetCaptureParameter(vfwParameter)
End Sub


'释放资源
Private Sub Class_Terminate()
  Set mVfwControl = Nothing
End Sub
