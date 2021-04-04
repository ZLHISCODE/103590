Attribute VB_Name = "mdlAvicap32"

Option Explicit



'-------------------跟avicap32相关的API库函数和变量的声明--------------------
Public Const WM_USER As Long = &H400
Public Const WM_CAP_START As Long = WM_USER

Public Const WM_CAP_GET_CAPSTREAMPTR As Long = WM_CAP_START + 1

Public Const WM_CAP_SET_CALLBACK_ERROR As Long = WM_CAP_START + 2
Public Const WM_CAP_SET_CALLBACK_STATUS As Long = WM_CAP_START + 3
Public Const WM_CAP_SET_CALLBACK_YIELD As Long = WM_CAP_START + 4
Public Const WM_CAP_SET_CALLBACK_FRAME As Long = WM_CAP_START + 5
Public Const WM_CAP_SET_CALLBACK_VIDEOSTREAM As Long = WM_CAP_START + 6
Public Const WM_CAP_SET_CALLBACK_WAVESTREAM As Long = WM_CAP_START + 7
Public Const WM_CAP_GET_USER_DATA As Long = WM_CAP_START + 8
Public Const WM_CAP_SET_USER_DATA As Long = WM_CAP_START + 9
    
Public Const WM_CAP_DRIVER_CONNECT As Long = WM_CAP_START + 10
Public Const WM_CAP_DRIVER_DISCONNECT As Long = WM_CAP_START + 11
Public Const WM_CAP_DRIVER_GET_NAME As Long = WM_CAP_START + 12
Public Const WM_CAP_DRIVER_GET_VERSION As Long = WM_CAP_START + 13
Public Const WM_CAP_DRIVER_GET_CAPS As Long = WM_CAP_START + 14

Public Const WM_CAP_FILE_SET_CAPTURE_FILE As Long = WM_CAP_START + 20
Public Const WM_CAP_FILE_GET_CAPTURE_FILE As Long = WM_CAP_START + 21
Public Const WM_CAP_FILE_ALLOCATE As Long = WM_CAP_START + 22
Public Const WM_CAP_FILE_SAVEAS As Long = WM_CAP_START + 23
Public Const WM_CAP_FILE_SET_INFOCHUNK As Long = WM_CAP_START + 24
Public Const WM_CAP_FILE_SAVEDIB As Long = WM_CAP_START + 25

Public Const WM_CAP_EDIT_COPY As Long = WM_CAP_START + 30

Public Const WM_CAP_SET_AUDIOFORMAT As Long = WM_CAP_START + 35
Public Const WM_CAP_GET_AUDIOFORMAT As Long = WM_CAP_START + 36

Public Const WM_CAP_DLG_VIDEOFORMAT As Long = WM_CAP_START + 41
Public Const WM_CAP_DLG_VIDEOSOURCE As Long = WM_CAP_START + 42
Public Const WM_CAP_DLG_VIDEODISPLAY As Long = WM_CAP_START + 43
Public Const WM_CAP_GET_VIDEOFORMAT As Long = WM_CAP_START + 44
Public Const WM_CAP_SET_VIDEOFORMAT As Long = WM_CAP_START + 45
Public Const WM_CAP_DLG_VIDEOCOMPRESSION As Long = WM_CAP_START + 46

Public Const WM_CAP_SET_PREVIEW As Long = WM_CAP_START + 50
Public Const WM_CAP_SET_OVERLAY As Long = WM_CAP_START + 51
Public Const WM_CAP_SET_PREVIEWRATE As Long = WM_CAP_START + 52
Public Const WM_CAP_SET_SCALE As Long = WM_CAP_START + 53
Public Const WM_CAP_GET_STATUS As Long = WM_CAP_START + 54
Public Const WM_CAP_SET_SCROLL As Long = WM_CAP_START + 55

Public Const WM_CAP_GRAB_FRAME As Long = WM_CAP_START + 60
Public Const WM_CAP_GRAB_FRAME_NOSTOP As Long = WM_CAP_START + 61

Public Const WM_CAP_SEQUENCE As Long = WM_CAP_START + 62
Public Const WM_CAP_SEQUENCE_NOFILE As Long = WM_CAP_START + 63
Public Const WM_CAP_SET_SEQUENCE_SETUP As Long = WM_CAP_START + 64
Public Const WM_CAP_GET_SEQUENCE_SETUP As Long = WM_CAP_START + 65
Public Const WM_CAP_SET_MCI_DEVICE As Long = WM_CAP_START + 66
Public Const WM_CAP_GET_MCI_DEVICE As Long = WM_CAP_START + 67
Public Const WM_CAP_STOP As Long = WM_CAP_START + 68
Public Const WM_CAP_ABORT As Long = WM_CAP_START + 69

Public Const WM_CAP_SINGLE_FRAME_OPEN As Long = WM_CAP_START + 70
Public Const WM_CAP_SINGLE_FRAME_CLOSE As Long = WM_CAP_START + 71
Public Const WM_CAP_SINGLE_FRAME As Long = WM_CAP_START + 72

Public Const WM_CAP_PAL_OPEN As Long = WM_CAP_START + 80
Public Const WM_CAP_PAL_SAVE As Long = WM_CAP_START + 81
Public Const WM_CAP_PAL_PASTE As Long = WM_CAP_START + 82
Public Const WM_CAP_PAL_AUTOCREATE As Long = WM_CAP_START + 83
Public Const WM_CAP_PAL_MANUALCREATE As Long = WM_CAP_START + 84

Public Const WM_CAP_SET_CALLBACK_CAPCONTROL As Long = WM_CAP_START + 85

Public Const WS_VISIBLE As Long = &H10000000
Public Const SWP_NOSIZE As Long = &H1&
Public Const SWP_NOMOVE As Long = &H2&
Public Const SWP_NOSENDCHANGING As Long = &H400&

Public Const AVSTREAMMASTER_NONE  As Long = 1
Public Const INDEX_15_MINUTES As Long = 27000
Public Const INDEX_3_HOURS As Long = 324000

Public Const PlayFPS = 15                            '每秒多少FPS

Type VFWPOINT
        X As Long
        Y As Long
End Type

Type CAPSTATUS
    uiImageWidth As Long
    uiImageHeight As Long
    fLiveWindow As Long
    fOverlayWindow As Long
    fScale As Long
    ptScroll As VFWPOINT
    fUsingDefaultPalette As Long
    fAudioHardware As Long
    fCapFileExists As Long
    dwCurrentVideoFrame As Long
    dwCurrentVideoFramesDropped As Long
    dwCurrentWaveSamples As Long
    dwCurrentTimeElapsedMS As Long
    hPalCurrent As Long
    fCapturingNow As Long
    dwReturn As Long
    wNumVideoAllocated As Long
    wNumAudioAllocated As Long
End Type

Type CAPTUREPARMS
    dwRequestMicroSecPerFrame As Long       '// Requested capture rate
    fMakeUserHitOKToCapture As Long         '// Show "Hit OK to cap" dlg?
    wPercentDropForError As Long            '// Give error msg if > (10% default)
    fYield As Long                          '// Capture via background task?
    dwIndexSize As Long                     '// Max index size in frames (32K default)
    wChunkGranularity As Long               '// Junk chunk granularity (2K default)
    fUsingDOSMemory As Long                 '// Use DOS buffers? (obsolete)
    wNumVideoRequested As Long              '// # video buffers, If 0, autocalc
    fCaptureAudio As Long                   '// Capture audio?
    wNumAudioRequested As Long              '// # audio buffers, If 0, autocalc
    vKeyAbort As Long                       '// Virtual key causing abort
    fAbortLeftMouse As Long                 '// Abort on left mouse?
    fAbortRightMouse As Long                '// Abort on right mouse?
    fLimitEnabled As Long                   '// Use wTimeLimit?
    wTimeLimit As Long                      '// Seconds to capture
    fMCIControl As Long                     '// Use MCI video source?
    fStepMCIDevice As Long                  '// Step MCI device?
    dwMCIStartTime As Long                  '// Time to start in MS
    dwMCIStopTime As Long                   '// Time to stop in MS
    fStepCaptureAt2x As Long                '// Perform spatial averaging 2x
    wStepCaptureAverageFrames As Long       '// Temporal average n Frames
    dwAudioBufferSize As Long               '// Size of audio bufs (0 = default)
    fDisableWriteCache As Long              '// Attempt to disable write cache
    AVStreamMaster As Long                  '// Which stream controls length?
End Type

'得到采集驱动列表
Declare Function capGetDriverDescription Lib "avicap32.dll" Alias "capGetDriverDescriptionA" _
                                        (ByVal dwDriverIndex As Long, _
                                        ByVal lpszName As String, _
                                        ByVal cbName As Long, _
                                        ByVal lpszVer As String, _
                                        ByVal cbVer As Long) As Long
'创建采集窗口
Declare Function capCreateCaptureWindow Lib "avicap32.dll" Alias "capCreateCaptureWindowA" _
                                        (ByVal lpszWindowName As String, _
                                        ByVal dwStyle As Long, _
                                        ByVal X As Long, _
                                        ByVal Y As Long, _
                                        ByVal nWidth As Long, _
                                        ByVal nHeight As Long, _
                                        ByVal hWndParent As Long, _
                                        ByVal nID As Long) As Long




'-------------------过程-------------------------------------------



Function mResizeCaptureWindow(hCapWnd As Long) As Boolean
'---------------------------------------------------------------------
'功能：根据采集的窗体大小来设置窗体大小
'参数：
'返回：True = 成功 False = 失败
'上级函数或过程：
'下级函数或过程：SendMessageAsAny
'引用的外部参数：hCapWnd
'编制人：曾超
'修改人：
'---------------------------------------------------------------------
    Dim capStat As CAPSTATUS
    mResizeCaptureWindow = SendMessageAsAny(hCapWnd, WM_CAP_GET_STATUS, Len(capStat), capStat)

    Call SetWindowPos(hCapWnd, _
                0&, _
                0&, _
                0&, _
                capStat.uiImageWidth, _
                capStat.uiImageHeight, _
                SWP_NOMOVE Or SWP_NOZORDER Or SWP_NOSENDCHANGING)

End Function

Function mGetCaptureWindowStatus(hCapWnd As Long) As CAPSTATUS
'---------------------------------------------------------------------
'功能：返回窗体的状态
'参数：
'返回：CAPSTATUS自定义类型
'上级函数或过程：
'下级函数或过程：SendMessageAsAny
'引用的外部参数：hCapWnd
'编制人：曾超
'修改人：
'---------------------------------------------------------------------
    Dim capStat As CAPSTATUS
    
    Call SendMessageAsAny(hCapWnd, WM_CAP_GET_STATUS, Len(capStat), capStat)
    
    mGetCaptureWindowStatus.dwCurrentTimeElapsedMS = capStat.dwCurrentTimeElapsedMS
    mGetCaptureWindowStatus.dwCurrentVideoFrame = capStat.dwCurrentVideoFrame
    mGetCaptureWindowStatus.dwCurrentVideoFramesDropped = capStat.dwCurrentVideoFramesDropped
    mGetCaptureWindowStatus.dwCurrentWaveSamples = capStat.dwCurrentWaveSamples
    mGetCaptureWindowStatus.dwReturn = capStat.dwReturn
    mGetCaptureWindowStatus.fAudioHardware = capStat.fAudioHardware
    mGetCaptureWindowStatus.fCapFileExists = capStat.fCapFileExists
    mGetCaptureWindowStatus.fLiveWindow = capStat.fLiveWindow
    mGetCaptureWindowStatus.fOverlayWindow = capStat.fOverlayWindow
    mGetCaptureWindowStatus.fScale = capStat.fScale
    mGetCaptureWindowStatus.fUsingDefaultPalette = capStat.fUsingDefaultPalette
    mGetCaptureWindowStatus.hPalCurrent = capStat.hPalCurrent
    mGetCaptureWindowStatus.ptScroll = capStat.ptScroll
    mGetCaptureWindowStatus.uiImageHeight = capStat.uiImageHeight
    mGetCaptureWindowStatus.uiImageWidth = capStat.uiImageWidth
    mGetCaptureWindowStatus.wNumAudioAllocated = capStat.wNumAudioAllocated
    mGetCaptureWindowStatus.wNumVideoAllocated = capStat.wNumVideoAllocated
    
End Function

Function mSelectCapDevice(hCapWnd As Long, CapDeviceIndex As Integer) As Boolean
'---------------------------------------------------------------------
'功能：连接到指定设备
'参数：CapDeviceIndex 设备索引(0--8)
'返回：True = 成功 False = 失败
'上级函数或过程：
'下级函数或过程：SendMessageAsLong
'引用的外部参数：hCapWnd
'编制人：曾超
'修改人：黄捷 2007-4-2
'---------------------------------------------------------------------
    mSelectCapDevice = SendMessageAsLong(hCapWnd, WM_CAP_DRIVER_CONNECT, CapDeviceIndex, 0&)
    
End Function


Sub mCopyImageToClipBoard(hCapWnd As Long)
'---------------------------------------------------------------------
'功能：复制图像到粘贴板
'参数：
'返回：
'上级函数或过程：
'下级函数或过程：SendMessageAsLong
'引用的外部参数：
'编制人：曾超
'修改人：黄捷 2007-4-2
'---------------------------------------------------------------------
    Call SendMessageAsLong(hCapWnd, WM_CAP_EDIT_COPY, 0&, 0&)
End Sub

Function mViewerFormat(hCapWnd As Long) As Boolean
'---------------------------------------------------------------------
'功能：显示图像格式
'参数：
'返回：True = 成功 False = 失败
'上级函数或过程：
'下级函数或过程：SendMessageAsLong
'引用的外部参数：
'编制人：曾超
'修改人：黄捷 2007-4-2
'---------------------------------------------------------------------
    mViewerFormat = SendMessageAsLong(hCapWnd, WM_CAP_DLG_VIDEOFORMAT, 0&, 0&)
    'mViewerFormat = mResizeCaptureWindow
End Function


Function mViewerSource(hCapWnd As Long) As Boolean
'---------------------------------------------------------------------
'功能：显示图像来源
'参数：
'返回：True = 成功 False = 失败
'上级函数或过程：
'下级函数或过程：SendMessageAsLong
'引用的外部参数：
'编制人：曾超
'修改人：黄捷 2007-4-2
'---------------------------------------------------------------------
    mViewerSource = SendMessageAsLong(hCapWnd, WM_CAP_DLG_VIDEOSOURCE, 0&, 0&)
End Function

Function capDlgVideoCompression(ByVal hCapWnd As Long) As Boolean

   capDlgVideoCompression = SendMessageAsLong(hCapWnd, WM_CAP_DLG_VIDEOCOMPRESSION, 0&, 0&)
End Function

Function capFileSetCaptureFile(ByVal hCapWnd As Long, ByVal FilePath As String) As Boolean
'---------------------------------------------------------------------
'功能：设置采集录像的路径
'参数：
'返回：
'上级函数或过程：
'下级函数或过程：SendMessageAsString
'引用的外部参数：
'编制人：曾超
'修改人：
'---------------------------------------------------------------------
   capFileSetCaptureFile = SendMessageAsString(hCapWnd, WM_CAP_FILE_SET_CAPTURE_FILE, 0&, FilePath)
End Function


Function mcapCaptureSetSetup(ByVal hCapWnd As Long, ByRef capParms As CAPTUREPARMS) As Boolean
'---------------------------------------------------------------------
'功能：设置采集录像参数设置
'参数：
'返回：
'上级函数或过程：
'下级函数或过程：SendMessageAsString
'引用的外部参数：hCapWnd
'编制人：曾超
'修改人：
'---------------------------------------------------------------------
   Call SendMessageAsAny(hCapWnd, WM_CAP_SET_SEQUENCE_SETUP, Len(capParms), capParms)
End Function

Function capCaptureSequence(ByVal hCapWnd As Long) As Boolean
'---------------------------------------------------------------------
'功能：设置采集录像参数设置
'参数：
'返回：
'上级函数或过程：
'下级函数或过程：SendMessageAsString
'引用的外部参数：hCapWnd
'编制人：曾超
'修改人：
'---------------------------------------------------------------------
   capCaptureSequence = SendMessageAsLong(hCapWnd, WM_CAP_SEQUENCE, 0&, 0&)
End Function


Public Function mGetCapSureDevice() As String
'---------------------------------------------------------------------
'功能：获取视频设备清单
'参数：
'返回：设备清单用";"分开
'上级函数或过程：
'下级函数或过程：capGetDriverDescription
'引用的外部参数：
'编制人：曾超
'修改人：
'---------------------------------------------------------------------
    '获取驱动列表
    Const MAXVIDDRIVERS As Long = 9
    Const CAP_STRING_MAX As Long = 128
    Dim Index As Long
    Dim DEVICE As String
    Dim VERSION As String
    Dim strTmp As String
    
    DEVICE = String$(CAP_STRING_MAX, 0)
    VERSION = String$(CAP_STRING_MAX, 0)
    For Index = 0 To 8
        If 0 <> capGetDriverDescription(Index, DEVICE, CAP_STRING_MAX, VERSION, CAP_STRING_MAX) Then
             strTmp = Left(DEVICE, InStr(DEVICE, vbNullChar) - 1) & Left$(VERSION, InStr(VERSION, vbNullChar) - 1)
             If Len(Trim(mGetCapSureDevice)) > 0 Then
                mGetCapSureDevice = mGetCapSureDevice & ";"
             End If
             mGetCapSureDevice = mGetCapSureDevice & strTmp
        End If
    Next
End Function


Function mDisConnectDevice(hCapWnd As Long, CapDeviceIndex As Integer)
'---------------------------------------------------------------------
'功能：断开跟设备的连接
'参数：CapDeviceIndex 设备索引(0--8)
'返回：True = 成功 False = 失败
'上级函数或过程：
'下级函数或过程：SendMessageAsLong
'编制人：黄捷 2009-2-6
'---------------------------------------------------------------------
    mDisConnectDevice = SendMessageAsLong(hCapWnd, WM_CAP_DRIVER_DISCONNECT, CapDeviceIndex, 0&)
End Function


Public Function mGrapNoStopAndPreview(hCapWnd As Long) As Boolean
'---------------------------------------------------------------------
'功能：抓取一个图像，并继续预览状态
'参数：无
'返回：True = 成功 False = 失败
'上级函数或过程：
'下级函数或过程：SendMessageAsLong
'编制人：黄捷 2009-2-6
'---------------------------------------------------------------------
    mGrapNoStopAndPreview = SendMessageAsLong(hCapWnd, WM_CAP_GRAB_FRAME_NOSTOP, 0, 0&)
    mGrapNoStopAndPreview = SendMessageAsLong(hCapWnd, WM_CAP_SET_PREVIEW, -(True), 0&)
End Function

Public Function mGetVideoFormat(hCapWnd As Long, ByRef BITCapTureInfo As BITMAPINFO) As Boolean
'---------------------------------------------------------------------
'功能：读取视频格式参数
'参数：BITCapTureInfo 视频格式参数结构体
'返回：True = 成功 False = 失败
'上级函数或过程：
'下级函数或过程：SendMessageAsLong
'编制人：黄捷 2009-2-6
'---------------------------------------------------------------------
    Dim blnRet As Boolean
    
    mGetVideoFormat = SendMessage(hCapWnd, WM_CAP_GET_VIDEOFORMAT, Len(BITCapTureInfo), BITCapTureInfo)
End Function

Public Function mConnectCapDevice(hCapWnd As Long, hWndParent As Long, mintDeviceIndex As Integer, _
        intCapBitCount As Integer, intCapBiWidth As Integer, intCapBiHeight As Integer) As Boolean
        
    Dim BITCapTureInfo As BITMAPINFO
    Dim retVal As Boolean
    
    If hCapWnd = 0 Then
        hCapWnd = capCreateCaptureWindow("ZLSOFT_CAPTURE", WS_CHILD Or WS_VISIBLE, 0, 0, 100, 100, hWndParent, 0)
    End If
    
    If hCapWnd = 0 Then
        MsgboxCus "创建采集窗体失败！", vbInformation, gstrSysName
        Exit Function
    End If

    retVal = SendMessageAsLong(hCapWnd, WM_CAP_DRIVER_CONNECT, mintDeviceIndex, 0&)
    
    If retVal = False Then
        MsgboxCus "连接设备失败!", vbInformation, gstrSysName
        DestroyWindow hCapWnd
        hCapWnd = 0
        Exit Function
    End If
    
    SendMessage hCapWnd, WM_CAP_GET_VIDEOFORMAT, Len(BITCapTureInfo), BITCapTureInfo
    
    If intCapBitCount <> 0 And BITCapTureInfo.bmiHeader.biBitCount <> 0 Then
        With BITCapTureInfo.bmiHeader
            .biBitCount = intCapBitCount
            .biWidth = intCapBiWidth
            .biHeight = intCapBiHeight
            .biSizeImage = .biWidth * .biHeight * CInt(.biBitCount / 8)
        End With
        SendMessage hCapWnd, WM_CAP_SET_VIDEOFORMAT, 0, BITCapTureInfo
    End If

    Call SendMessageAsLong(hCapWnd, WM_CAP_SET_PREVIEWRATE, 66, 0&)
    Call SendMessageAsLong(hCapWnd, WM_CAP_SET_PREVIEW, -(True), 0&)
    Call SendMessageAsLong(hCapWnd, WM_CAP_SET_SCALE, -(True), 0&)
End Function
        
        

