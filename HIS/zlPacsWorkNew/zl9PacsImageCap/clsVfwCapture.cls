VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsVfwCapture"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------------------------
'*********************************************************************************************************
'
'说明：封装vfw的采集相关功能
'创建日期：2009-10-29
'创建人：TJH
'
'其他描述：
'
'*********************************************************************************************************
'---------------------------------------------------------------------------------------------------------


Option Explicit

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'类所需变量或者常量定义
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Const M_LNG_DEFAULT_CAP_WINDOW_WIDTH As Long = 320 '160      '默认视频采集窗口的宽度（注意和采集视频宽度的区别）
Private Const M_LNG_DEFAULT_CAP_WINDOW_HEIGHT As Long = 240 '120     '默认视频采集窗口的高度



Private Const M_LNG_DEFAULT_CUSTOM_ERRNUM As Long = -1 '默认自定义错误号


Private mCaptureParameterInf As clsVfwParameterCfg  '保存当前采集参数


Private mlngCurCaptureWindowHandle As Long '创建采集窗口后的窗口句柄
Private mlngVideoDisplayWindowHandle As Long '视频显示窗口的句柄

Private mlngDisplayWindowWidth As Long '当前负责视频显示的窗口宽度
Private mlngDisplayWindowHeight As Long '当前负责视频显示的窗口高度

Private mstrParameterCfgFileName As String '参数配置文件名称


'当视频显示窗口大小调整后产生的事件
Public Event OnVideoWindowChange(ByVal lngWidth As Long, ByVal lngHeight As Long, ByRef blnIsChangeSize As Boolean)




'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'卸载采集设备
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub StopPreview()
  If mlngCurCaptureWindowHandle = 0 Then
    Exit Sub
  End If
  
  Call SendMessageAsLong(mlngCurCaptureWindowHandle, WM_CAP_DRIVER_DISCONNECT, mCaptureParameterInf.CaptureDeviceIndex, 0&) '断开与采集设备的连接
  Call DestroyWindow(mlngCurCaptureWindowHandle) '释放采集窗口资源
  
  mlngCurCaptureWindowHandle = 0
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'根据fps取得MicrosSec值
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function GetMicrosSecFromFPS(ByVal lngFps As Long) As Long
  If lngFps = 0 Then Exit Function
  GetMicrosSecFromFPS = 1000000 / lngFps
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'取得VFW的相关采集参数
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetCaptureParameter() As clsVfwParameterCfg
  '//返回vfw视频采集参数
  Set GetCaptureParameter = mCaptureParameterInf
End Function



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'设置相关采集参数
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetCaptureParameter(ByRef captureParmeter As clsVfwParameterCfg)

  If captureParmeter Is Nothing Then
    Exit Sub
  End If
        
    
  '克隆采集参数对象
  Call captureParmeter.CloneTo(mCaptureParameterInf)
  
  '应用参数配置
  'Call ConfigCaptureParameter(mCaptureParameterInf)
End Sub



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'配置采集所需要的参数
'
'参数说明：
'capBit: 采集窗口位数
'capWidth: 采集窗口宽度
'capHeight: 采集窗口高度
'isScale: 是否自动适应窗口大小
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ConfigCaptureParameter(ByRef captureParmeter As clsVfwParameterCfg)
    
  If captureParmeter Is Nothing Then
    Exit Sub
  End If
    
  If mlngCurCaptureWindowHandle = 0 Then
    Exit Sub
  End If
            
    
  '取得采集相关设置参数
  Dim biCaptureInf As BITMAPINFO
    
  Call SendMessage(mlngCurCaptureWindowHandle, WM_CAP_GET_VIDEOFORMAT, Len(biCaptureInf), biCaptureInf)
    
  '设置采集参数
  If biCaptureInf.bmiHeader.biBitCount <> 0 Then
      With biCaptureInf.bmiHeader
          .biBitCount = captureParmeter.CapBit
          .biWidth = captureParmeter.CapVideoWidth
          .biHeight = captureParmeter.CapVideoHeight
          .biSizeImage = captureParmeter.CapVideoWidth * captureParmeter.CapVideoHeight * CInt(captureParmeter.CapBit / 8)
      End With
      Call SendMessage(mlngCurCaptureWindowHandle, WM_CAP_SET_VIDEOFORMAT, 0, biCaptureInf)
  End If

  'preview 模式
  Call SendMessageAsLong(mlngCurCaptureWindowHandle, WM_CAP_SET_PREVIEWRATE, captureParmeter.CapPreviewRate, 0&) '设置预览速度
  Call SendMessageAsLong(mlngCurCaptureWindowHandle, WM_CAP_SET_PREVIEW, 1, 0&) '进入预览
  Call SendMessageAsLong(mlngCurCaptureWindowHandle, WM_CAP_SET_SCALE, -(captureParmeter.IsScale), 0&) '自动适应窗口大小
    
  'overlay模式(需要根据CAPSTATUS结构的fOverlayWindow判断是否支持覆盖模式)
  'Call SendMessageAsLong(curCaptureWindowHandle, WM_CAP_SET_OVERLAY, True, 0&)
    
    
  '设置视频采集参数
  Dim videoParameter As CAPTUREPARMS
    
  '取得视频采集相关参数
  Call SendMessage(mlngCurCaptureWindowHandle, WM_CAP_GET_CAPSTREAMPTR, Len(videoParameter), videoParameter)
    
  With videoParameter
      .wPercentDropForError = 10  '帧丢失的最大百分比
      .fMakeUserHitOKToCapture = False '是否弹出用户提示
      .fUsingDOSMemory = True  '是否使用dos内存模式
      .wNumVideoRequested = 32
      .fAbortLeftMouse = -(captureParmeter.AbortLeftMouse) '用鼠标左键或者右键停止采集
      .fAbortRightMouse = -(captureParmeter.AbortRightMouse)
      .wChunkGranularity = 0
      .dwAudioBufferSize = 0
      .fDisableWriteCache = False 'Not used in Win32 applications.
      .fMCIControl = False  'not using MCI
      .fStepCaptureAt2x = False
      .fYield = False '   not multi-threading
      .wNumAudioRequested = 4 '10 is max limit
        
        
      .AVStreamMaster = captureParmeter.AVStreamMaster
        
      .dwIndexSize = INDEX_15_MINUTES
      .dwRequestMicroSecPerFrame = GetMicrosSecFromFPS(captureParmeter.RequestMicroSecPerFrame)
      .fCaptureAudio = captureParmeter.CaptureAudio
      .fLimitEnabled = captureParmeter.LimitEnabled
      .wTimeLimit = INDEX_15_MINUTES
  End With
    
  Call SendMessage(mlngCurCaptureWindowHandle, WM_CAP_SET_SEQUENCE_SETUP, Len(videoParameter), videoParameter)
    
End Sub




'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'读取采集参数
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ReadVfwCaptureParameterFromFile()
  mCaptureParameterInf.ReadParameterCfg (mstrParameterCfgFileName)
End Sub



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'保存采集参数
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SaveVfwCaptureParameterToFile()
  mCaptureParameterInf.SaveParameterCfg (mstrParameterCfgFileName)
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'刷新参数
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub RefreshParameter()
  Call StartPreview(mlngVideoDisplayWindowHandle)
End Sub


Public Sub ShowFullScreen(owner As Object, ByRef sourceWindow As PictureBox, ByVal monitorIndex As Integer)
  '如果句柄为零，则退出全屏显示
  If mlngCurCaptureWindowHandle = 0 Then
    Exit Sub
  End If
  
  Call StopPreview
  Call frmFullScreen.ShowFullScreen(Me, mCaptureParameterInf, owner, sourceWindow, monitorIndex)
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'加载采集设备
'
'参数说明：
'deviceIndex   设备索引
'capDisplyWindowHwnd   视频显示窗口句柄
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub StartPreview(ByVal lngCapDisplyWindowHwnd As Long)
    
  If lngCapDisplyWindowHwnd = 0 Then
    Exit Sub
  End If
    
    
  '判断采集窗口句柄是否为0，如果不为零则需要先卸载采集设备
  If mlngCurCaptureWindowHandle <> 0 Then
    'Call StopPreview
    Exit Sub
  End If
  
  mlngVideoDisplayWindowHandle = lngCapDisplyWindowHwnd
  
  '创建采集窗口
  mlngCurCaptureWindowHandle = capCreateCaptureWindow("VideoCapture", WS_CHILD Or WS_VISIBLE, 0, 0, M_LNG_DEFAULT_CAP_WINDOW_WIDTH, M_LNG_DEFAULT_CAP_WINDOW_HEIGHT, mlngVideoDisplayWindowHandle, 0)
  If mlngCurCaptureWindowHandle = 0 Then
    'Call err.Raise(M_LNG_DEFAULT_CUSTOM_ERRNUM, "capCreateCaptureWindow", "采集窗口创建失败。")
    Exit Sub
  End If
    
  '连接采集设备
  Dim bIsSendSuccessed As Boolean
  bIsSendSuccessed = SendMessageAsLong(mlngCurCaptureWindowHandle, WM_CAP_DRIVER_CONNECT, mCaptureParameterInf.CaptureDeviceIndex, 0&)
    
  If Not bIsSendSuccessed Then
    bIsSendSuccessed = SendMessageAsLong(mlngCurCaptureWindowHandle, WM_CAP_DRIVER_CONNECT, mCaptureParameterInf.CaptureDeviceIndex, 0&)
    
    If Not bIsSendSuccessed Then
        mlngCurCaptureWindowHandle = 0
    
        'Call err.Raise(M_LNG_DEFAULT_CUSTOM_ERRNUM, "SendMessageAsLong", "采集设备连接失败。")
        Exit Sub
    End If
  End If
  
 
  '设置默认的采集参数
  Call ConfigCaptureParameter(mCaptureParameterInf)
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'更新采集窗口大小和位置
'
'参数说明：
'displayWindowWidth   当前采集窗口宽度(以像素为单位)
'displayWindowHeight   当前采集窗口高度(以像素为单位)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub UpdateCaptureWindowPos(ByVal lngDisplayWindowWidth As Long, ByVal lngDisplayWindowHeight As Long)
  
  If mlngCurCaptureWindowHandle = 0 Then
    Exit Sub
  End If
  
  '//保存当前视频显示窗口大小
  mlngDisplayWindowWidth = lngDisplayWindowWidth
  mlngDisplayWindowHeight = lngDisplayWindowHeight
      
      
  Dim dblWCutRate As Double
  Dim dblHCutRate As Double
  Dim lngCutWidth As Long
  Dim lngCutHeight As Long
  Dim blnIsChangeSize As Boolean
      
      
  Select Case mCaptureParameterInf.VideoShowWay
    Case 0
      '原始大小显示
      '设置采集窗口位置，并使其居中
      Call SetWindowPos(mlngCurCaptureWindowHandle, _
                0&, _
                (mlngDisplayWindowWidth - mCaptureParameterInf.CapVideoWidth) / 2, _
                (mlngDisplayWindowHeight - mCaptureParameterInf.CapVideoHeight) / 2, _
                mCaptureParameterInf.CapVideoWidth, _
                mCaptureParameterInf.CapVideoHeight, _
                SWP_NOZORDER Or SWP_NOSENDCHANGING)
    
    Case 1
      '拉伸显示
      Call SetWindowPos(mlngCurCaptureWindowHandle, _
                0&, 0, 0, mlngDisplayWindowWidth, mlngDisplayWindowHeight, _
                SWP_NOZORDER Or SWP_NOSENDCHANGING)
              
    Case 2
      '按比例缩放
      Dim rate As Double
      
      If (mCaptureParameterInf.CapVideoHeight / mCaptureParameterInf.CapVideoWidth) > (mlngDisplayWindowHeight / mlngDisplayWindowWidth) Then
        rate = (mlngDisplayWindowHeight) / mCaptureParameterInf.CapVideoHeight
      Else
        rate = mlngDisplayWindowWidth / mCaptureParameterInf.CapVideoWidth
      End If
      
      If (mCaptureParameterInf.CapVideoHeight = mlngDisplayWindowHeight) And (mCaptureParameterInf.CapVideoWidth = mlngDisplayWindowWidth) Then
        rate = 1
      End If
      
      
      Call SetWindowPos(mlngCurCaptureWindowHandle, _
                0&, _
                (mlngDisplayWindowWidth - mCaptureParameterInf.CapVideoWidth * rate) / 2, _
                (mlngDisplayWindowHeight - mCaptureParameterInf.CapVideoHeight * rate) / 2, _
                mCaptureParameterInf.CapVideoWidth * rate, _
                mCaptureParameterInf.CapVideoHeight * rate, _
                SWP_NOZORDER Or SWP_NOSENDCHANGING)
    Case 3
      '按裁剪范围大小进行显示
      If mCaptureParameterInf.WidthCutRate <= 0 Or mCaptureParameterInf.HeightCutRate <= 0 Then
        dblWCutRate = 1
        dblWCutRate = 1
      Else
        dblWCutRate = mCaptureParameterInf.WidthCutRate
        dblHCutRate = mCaptureParameterInf.HeightCutRate
      End If
      
      
      lngCutWidth = mCaptureParameterInf.CapVideoWidth * dblWCutRate
      lngCutHeight = mCaptureParameterInf.CapVideoHeight * dblHCutRate
      
      Call SetWindowPos(mlngCurCaptureWindowHandle, _
                0&, _
                CLng(0 - mCaptureParameterInf.CapVideoWidth * mCaptureParameterInf.LeftCutRate), _
                CLng(0 - mCaptureParameterInf.CapVideoHeight * mCaptureParameterInf.TopCutRate), _
                mCaptureParameterInf.CapVideoWidth, _
                mCaptureParameterInf.CapVideoHeight, _
                SWP_NOZORDER Or SWP_NOSENDCHANGING)
      
      blnIsChangeSize = False
      RaiseEvent OnVideoWindowChange(lngCutWidth, lngCutHeight, blnIsChangeSize)
      
      If blnIsChangeSize Then
        mlngDisplayWindowWidth = lngCutWidth
        mlngDisplayWindowHeight = lngCutHeight
      End If
      
    Case 4
      '自适应视频大小
      Call SetWindowPos(mlngCurCaptureWindowHandle, _
                0&, _
                0, _
                0, _
                mCaptureParameterInf.CapVideoWidth, _
                mCaptureParameterInf.CapVideoHeight, _
                SWP_NOZORDER Or SWP_NOSENDCHANGING)
      
      blnIsChangeSize = False
      RaiseEvent OnVideoWindowChange(mCaptureParameterInf.CapVideoWidth, mCaptureParameterInf.CapVideoHeight, blnIsChangeSize)
      
      If blnIsChangeSize Then
        mlngDisplayWindowWidth = mCaptureParameterInf.CapVideoWidth
        mlngDisplayWindowHeight = mCaptureParameterInf.CapVideoHeight
      End If
      
    Case Else
      '原始大小显示
      '设置采集窗口位置，并使其居中
      Call SetWindowPos(mlngCurCaptureWindowHandle, _
                0&, _
                (mlngDisplayWindowWidth - mCaptureParameterInf.CapVideoWidth) / 2, _
                (mlngDisplayWindowHeight - mCaptureParameterInf.CapVideoHeight) / 2, _
                mCaptureParameterInf.CapVideoWidth, _
                mCaptureParameterInf.CapVideoHeight, _
                SWP_NOZORDER Or SWP_NOSENDCHANGING)
    
  End Select
      

End Sub



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'图像采集
'
'参数说明：
'captureWindow   图像采集窗口的窗口句柄
'
'返回值说明：返回采集到的图像接口对象
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function CaptureImage() As IPictureDisp
      
  On Error GoTo errHandle
    '判断采集窗口句柄是否有效
    If mlngCurCaptureWindowHandle = 0 Then
        Set CaptureImage = Nothing
        Exit Function
    End If
    
    Call SendMessageAsLong(mlngCurCaptureWindowHandle, WM_CAP_GRAB_FRAME_NOSTOP, 0, 0&) '抓取一副图像到内存
    Call SendMessageAsLong(mlngCurCaptureWindowHandle, WM_CAP_SET_PREVIEW, 1, 0&) '继续浏览采集视频图像
    
    Call SendMessageAsLong(mlngCurCaptureWindowHandle, WM_CAP_EDIT_COPY, 0&, 0&) '从内存复制图像到剪贴板
    
    Set CaptureImage = Clipboard.GetData(2) '从剪贴板中返回数据
    
    Clipboard.Clear
    
    Exit Function
errHandle:
  Set CaptureImage = Nothing
End Function



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'图像采集
'
'将图像采集到剪贴板
'
'
'返回值说明：
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function CaptureImageToClipBoard() As Boolean

  On Error GoTo errHandle
    '判断采集窗口句柄是否有效
    If mlngCurCaptureWindowHandle = 0 Then
        CaptureImageToClipBoard = False
        Exit Function
    End If
    
    Call SendMessageAsLong(mlngCurCaptureWindowHandle, WM_CAP_GRAB_FRAME_NOSTOP, 0, 0&) '抓取一副图像到内存
    Call SendMessageAsLong(mlngCurCaptureWindowHandle, WM_CAP_SET_PREVIEW, 1, 0&) '继续浏览采集视频图像
    
    Call SendMessageAsLong(mlngCurCaptureWindowHandle, WM_CAP_EDIT_COPY, 0&, 0&) '从内存复制图像到剪贴板
    
    CaptureImageToClipBoard = True
    
    Exit Function
errHandle:
    CaptureImageToClipBoard = False
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'图像采集到指定文件中
'
'参数说明：
'sFileName:   图像采集后的保存文件名
'
'返回值说明：返回为真说明采集成功
Public Function CaptureBmpImgToFile(ByVal strFileName As String) As Boolean
  On Error GoTo errHandle
    CaptureBmpImgToFile = False
      
    '判断采集窗口句柄是否有效
    If mlngCurCaptureWindowHandle = 0 Then
        Exit Function
    End If
    
    Call SendMessageAsLong(mlngCurCaptureWindowHandle, WM_CAP_GRAB_FRAME_NOSTOP, 0, 0&) '抓取一副图像到内存
    Call SendMessageAsLong(mlngCurCaptureWindowHandle, WM_CAP_SET_PREVIEW, 1, 0&) '继续浏览采集视频图像
    
    Call SendMessageAsLong(mlngCurCaptureWindowHandle, WM_CAP_EDIT_COPY, 0&, 0&) '从内存复制图像到剪贴板
    
    '//从剪贴板中取得采集的图像
    Dim iPicture As IPictureDisp
    Set iPicture = Clipboard.GetData(2) '从剪贴板中返回数据
  
    '//将采集的图象保存到指定的文件中
    If Not (iPicture Is Nothing) Then
        Call SavePicture(iPicture, strFileName)
    End If
    
    Clipboard.Clear
  
    CaptureBmpImgToFile = True
    
    Exit Function
errHandle:
    CaptureBmpImgToFile = False
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'弹出视频源对话框
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ShowCaptureVideoSourceDialog()

  If mlngCurCaptureWindowHandle = 0 Then
    Exit Sub
  End If
  
  '//显示对话框
  Call SendMessageAsLong(mlngCurCaptureWindowHandle, WM_CAP_DLG_VIDEOSOURCE, 0&, 0&)
End Sub


Public Sub ShowCaptureParameterCfgDialog(objOwner As Object, Optional ByVal lngHideItem As Long = 0)
  Call frmVideoParameter.ShowVfwParameter(Me, mlngDisplayWindowWidth, mlngDisplayWindowHeight, objOwner, lngHideItem)
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'弹出图像压缩设置对话框
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ShowCaptureCompressionDialog()

  If mlngCurCaptureWindowHandle = 0 Then
    Exit Sub
  End If
  
  '//显示对话框
  Call SendMessageAsLong(mlngCurCaptureWindowHandle, WM_CAP_DLG_VIDEOCOMPRESSION, 0&, 0&)
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'弹出视频格式设置对话框
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ShowCaptureVideoFormatDialog()
  If mlngCurCaptureWindowHandle = 0 Then
    Exit Sub
  End If
  
  '//显示对话框
  Call SendMessageAsLong(mlngCurCaptureWindowHandle, WM_CAP_DLG_VIDEOFORMAT, 0&, 0&)
    
  Dim capState As CAPSTATUS
  Dim retVal As Boolean
  
  '取得设置后的捕获视频大小
  retVal = SendMessage(mlngCurCaptureWindowHandle, WM_CAP_GET_STATUS, Len(capState), capState)
  If retVal Then
    mCaptureParameterInf.CapVideoWidth = capState.uiImageWidth
    mCaptureParameterInf.CapVideoHeight = capState.uiImageHeight
  End If
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'弹出视频显示设置对话框（如果采集卡支持）
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ShowCaptureVideoDisplayDialog()

  If mlngCurCaptureWindowHandle = 0 Then
    Exit Sub
  End If
  
  '//显示对话框
  Call SendMessageAsLong(mlngCurCaptureWindowHandle, WM_CAP_DLG_VIDEODISPLAY, 0&, 0&)
End Sub


'开始视频采集
Public Sub StartVideo(ByVal strVideoFile As String)
  
  '如果没有载入采集设备，则直接退出
  If mlngCurCaptureWindowHandle = 0 Then
    Exit Sub
  End If
  
    
  '设置视频文件保存位置
  Call SendMessageAsString(mlngCurCaptureWindowHandle, WM_CAP_FILE_SET_CAPTURE_FILE, 0, strVideoFile)
    
  '发送视频采集开始消息
  Call SendMessageAsLong(mlngCurCaptureWindowHandle, WM_CAP_SEQUENCE, 0&, 0&)
    
End Sub


'停止视频采集（该方法没有测试）
Public Sub StopVideo()

  If mlngCurCaptureWindowHandle = 0 Then
    Exit Sub
  End If
    
  '发送消息停止VIDEO的采集
  Call SendMessageAsLong(mlngCurCaptureWindowHandle, WM_CAP_ABORT, 0&, 0&)
End Sub


'*********************************************************************************************
'
'类属性定义
'
'*********************************************************************************************


'采集窗口句柄（只读）
Property Get hWnd() As Long
  hWnd = mlngCurCaptureWindowHandle
End Property


'当前捕获视频宽度（只读）
Property Get CaptureVideoWidth() As Long
  CaptureVideoWidth = mCaptureParameterInf.CapVideoWidth
End Property


'当前捕获视频高度（只读）
Property Get CaptureVideoHeight() As Long
  CaptureVideoHeight = mCaptureParameterInf.CapVideoHeight
End Property

'参数配置文件名称
Property Get ParameterCfgFileName() As String
  ParameterCfgFileName = mstrParameterCfgFileName
End Property

Property Let ParameterCfgFileName(value As String)
  mstrParameterCfgFileName = value
End Property

'获取采集参数设置对象
Property Get CaptureParameterInf() As clsVfwParameterCfg
  Set CaptureParameterInf = mCaptureParameterInf
End Property



'*********************************************************************************************
'
'类事件
'
'*********************************************************************************************


Private Sub Class_Initialize()
  Set mCaptureParameterInf = New clsVfwParameterCfg
  
  '读取默认值
  Call mCaptureParameterInf.LoadDefaultCfg
End Sub

Private Sub Class_Terminate()
  Set mCaptureParameterInf = Nothing
End Sub




