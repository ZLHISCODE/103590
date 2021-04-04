Attribute VB_Name = "mdlVideoCapture"
Option Explicit
'Public pCurrentfrmCapture As frmVideoCapture    '记录拥有视频源的采集窗口
'
'Public Sub zlCapUnoladCapWnd(ByRef hCapWnd As Long)
'    '释放采集窗口
'    If hCapWnd = 0 Then Exit Sub
'    DestroyWindow hCapWnd
'    hCapWnd = 0
'End Sub
'
'Public Sub zlCapDisConnectCapDevice(miDeviceIndex As Integer, ByRef hCapWnd As Long)
'    '断开对设备的连接
'    If hCapWnd = 0 Then Exit Sub
'    Call mDisConnectDevice(hCapWnd, miDeviceIndex)
'End Sub
'
'Public Sub zlCapCopyImageToClipboard(ByRef hCapWnd As Long)
'    '把图像复制到剪贴板
'    If hCapWnd = 0 Then Exit Sub
'    Call mCopyImageToClipBoard(hCapWnd)
'End Sub
'
'Public Sub zlCapGrabNoStop(ByRef hCapWnd As Long)
'    '采集图像之前，先把一帧采集进入缓存
'    If hCapWnd = 0 Then Exit Sub
'    Call mGrapNoStopAndPreview(hCapWnd)
'End Sub
'
'Public Sub zlCapSetCaptureFile(mstrAviFileName As String, ByRef hCapWnd As Long)
'    '设置录像文件的名称
'    If hCapWnd = 0 Then Exit Sub
'    capFileSetCaptureFile hCapWnd, mstrAviFileName
'End Sub
'
'Public Sub zlCapCaptureSetSetup(ByRef hCapWnd As Long)
'    '设置录像的参数，现在的参数是，录像开始后，只有单击左键或者右键才会停止
'    Dim CapParams As CAPTUREPARMS
'
'    If hCapWnd = 0 Then Exit Sub
'
'    With CapParams
'        .wPercentDropForError = 10
'        .fMakeUserHitOKToCapture = True
'        .fUsingDOSMemory = True
'        .wNumVideoRequested = 32
'        .fAbortLeftMouse = -(True)
'        .fAbortRightMouse = -(True)
'        .wChunkGranularity = 0
'        .dwAudioBufferSize = 0
'        .fDisableWriteCache = False
'        .fMCIControl = False
'        .fStepCaptureAt2x = False
'        .fYield = False
'        .wNumAudioRequested = 4 '10 is max limit
'        .AVStreamMaster = AVSTREAMMASTER_NONE
'        '保存到哪里去了？读取哪里的参数？
'        .dwIndexSize = Val(GetSetting(App.Title, "preferences", "maxframes", INDEX_3_HOURS))
'        .dwRequestMicroSecPerFrame = microsSecFromFPS(Val(PlayFPS))
'        .fCaptureAudio = False
'        .fLimitEnabled = True
'        .wTimeLimit = Val(INDEX_3_HOURS)
'    End With
'    On Error GoTo CapErr
'    mcapCaptureSetSetup hCapWnd, CapParams
'    Exit Sub
'CapErr:
'End Sub
'
'Public Sub zlCapCaptureSequence(ByRef hCapWnd As Long)
'    '开始录像
'    If hCapWnd = 0 Then Exit Sub
'    capCaptureSequence hCapWnd
'End Sub
'
'Private Function microsSecFromFPS(ByVal fps As Long) As Long
'    If fps = 0 Then Exit Function
'    microsSecFromFPS = 1000000 / fps
'End Function
'
'Public Sub zlCapDlgVideoCompression(ByRef hCapWnd As Long)
'    '打开压缩设置窗口
'    If hCapWnd = 0 Then Exit Sub
'    capDlgVideoCompression hCapWnd
'End Sub
'
'Public Sub zlCapDlgVideoFormat(ByRef hCapWnd As Long)
'    '打开格式设置对话框
'    If hCapWnd = 0 Then Exit Sub
'    mViewerFormat hCapWnd
'End Sub
'
'Public Sub zlCapSaveVideoFormat(strRegName As String, ByRef hCapWnd As Long)
'    '保存格式设置到注册表中
'    Dim BITCapTureInfo As BITMAPINFO
'    Dim strRegPath As String
'    Dim blnRet As Boolean
'
'    If hCapWnd = 0 Then Exit Sub
'
'    blnRet = mGetVideoFormat(hCapWnd, BITCapTureInfo)
'
'    If blnRet = True Then
'        strRegPath = "公共模块\" & App.ProductName & "\" & strRegName
'        If BITCapTureInfo.bmiHeader.biBitCount <> 0 Then
'            SaveSetting "ZLSOFT", strRegPath, "CapBitCount", BITCapTureInfo.bmiHeader.biBitCount
'            SaveSetting "ZLSOFT", strRegPath, "CapBiWidth", BITCapTureInfo.bmiHeader.biWidth
'            SaveSetting "ZLSOFT", strRegPath, "CapBiHeight", BITCapTureInfo.bmiHeader.biHeight
'        End If
'    End If
'End Sub
'
'Public Sub zlCapDlgVideoSource(ByRef hCapWnd As Long)
'    '打开视频源设置对话框
'    If hCapWnd = 0 Then Exit Sub
'
'    mViewerSource hCapWnd
'End Sub
'
'Function zlCapConnCapDevice(ByRef hCapWnd As Long, ParentWindowWnd As Long, CapDeviceIndex As Integer, _
'    ByRef mlngVideoSizeX As Long, ByRef mlngVideoSizeY As Long) As Boolean
''-----------------------------------------------------------------------------------------
''功能：连接到设备，如果hCapWnd=0，则使用ParentWindowWnd创建采集窗体
''参数：hCapWnd 采集窗体句柄；ParentWindowWnd 父窗体句柄 ; CapDeviceIndex 设备索引号
''返回：True = 成功 False = 失败
''上级函数或过程：
''下级函数或过程：capCreateCaptureWindow;SendMessageAsLong;SendMessageAsAny;SetWindowPos
''引用的外部参数：
''编制人：曾超
''修改人：黄捷 2009-2-6
''-----------------------------------------------------------------------------------------
'
'    Dim strTmp() As String
'    Dim i  As Integer
'    Dim CaptureWinSize As CAPSTATUS
'
'    Dim intCapBitCount As Integer
'    Dim intCapBiWidth As Integer
'    Dim intCapBiHeight As Integer
'    Dim strRegPath As String
'
'    strRegPath = "公共模块\" & App.ProductName & "\frmVideoCapture"
'    intCapBitCount = Val(GetSetting("ZLSOFT", strRegPath, "CapBitCount", 0))
'    intCapBiWidth = Val(GetSetting("ZLSOFT", strRegPath, "CapBiWidth", 0))
'    intCapBiHeight = Val(GetSetting("ZLSOFT", strRegPath, "CapBiHeight", 0))
'
'    Call mConnectCapDevice(hCapWnd, ParentWindowWnd, CapDeviceIndex, intCapBitCount, intCapBiWidth, intCapBiHeight)
'
'    '获取视频窗口的分辨率
'    CaptureWinSize = mGetCaptureWindowStatus(hCapWnd)
'    mlngVideoSizeX = CaptureWinSize.uiImageWidth * Screen.TwipsPerPixelX
'    mlngVideoSizeY = CaptureWinSize.uiImageHeight * Screen.TwipsPerPixelY
'
'    zlCapConnCapDevice = True
'End Function
'
