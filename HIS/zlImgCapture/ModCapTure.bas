Attribute VB_Name = "ModCapTure"
Option Explicit
'--------------------------------------------------------
'功  能：连接视频输入设备显示保存图像等。
'编制人：曾超
'编制日期：2005.11.8
'过程函数清单：
'       mCapturePosition         采集窗体位置设置
'       mConnCapDevice           连接到设备
'       mGetCapSureDevice        获取视频设备清单
'       mGetCaptureWindowStatus  返回窗体的状态
'       mPaintPicture            图片重绘
'       mParentWindowResize      设置显示窗口的位置在父窗体中心
'       mResizeCaptureWindow     根据采集的窗体大小来设置窗体大小
'       mSaveImageFile           保存当前显示的图像
'       mSelectCapDevice         连接到指定设备
'       mViewerFormat            显示图像格式
'       mViewerSource            显示图像来源
'       mCopyImageToClipBoard    复制图像到粘贴板
'修改记录：
'
'-------------------------------------------------------
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

Public Const WS_CHILD As Long = &H40000000
Public Const WS_VISIBLE As Long = &H10000000
Public Const SWP_NOSIZE As Long = &H1&
Public Const SWP_NOMOVE As Long = &H2&
Public Const SWP_NOZORDER As Long = &H4&
Public Const SWP_NOSENDCHANGING As Long = &H400&
Public Const HWND_BOTTOM As Long = 1&

Public Const AVSTREAMMASTER_NONE  As Long = 1
Public Const INDEX_15_MINUTES As Long = 27000
Public Const INDEX_3_HOURS As Long = 324000
Public Const IDS_CAP_BEGIN As Long = 300
Public Const IDS_CAP_END As Long = 301
Public Const IDS_CAP_STAT_VIDEOAUDIO As Long = 511
Public Const IDS_CAP_STAT_VIDEOONLY As Long = 512

Public hCapWnd As Long, gintDeviceIndex As Integer
Public Const DSTINVERT = &HCC0020

Type VFWPOINT
        x As Long
        y As Long
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
                                        ByVal x As Long, _
                                        ByVal y As Long, _
                                        ByVal nWidth As Long, _
                                        ByVal nHeight As Long, _
                                        ByVal hWndParent As Long, _
                                        ByVal nID As Long) As Long
'消息发送
Declare Function SendMessageAsLong Lib "user32" Alias "SendMessageA" _
                                            (ByVal hwnd As Long, _
                                            ByVal wMsg As Long, _
                                            ByVal wParam As Long, _
                                            ByVal lParam As Long) As Long
Declare Function SendMessageAsAny Lib "user32" Alias "SendMessageA" _
                                            (ByVal hwnd As Long, _
                                            ByVal wMsg As Long, _
                                            ByVal wParam As Long, _
                                            ByRef lParam As Any) As Long
Declare Function SendMessageAsString Lib "user32" Alias "SendMessageA" _
                                            (ByVal hwnd As Long, _
                                            ByVal wMsg As Long, _
                                            ByVal wParam As Long, _
                                            ByVal lParam As String) As Long
                                            
Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long
Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal MSG As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function EnumChildWindows Lib "user32" (ByVal hWndParent As Long, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Public Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Declare Function SetWindowTextAsLong Lib "user32" Alias "SetWindowTextA" (ByVal hwnd As Long, ByVal LPCSTR As Long) As Long ' C BOOL
''''''''''''''''''''''''''''''''''''''''''''''''''保存录像时的压缩设置''''''''''''''''''''''''''''''''
Public preWinProc As Long
Public Const GWL_WNDPROC = (-4)
Public Const CB_GETCURSEL = &H147
Public Const CB_SETCURSEL = &H14E
Public Const CB_GETLBTEXT = &H148
Public Const CB_FINDSTRINGEXACT = &H158
Public blCompressionStup As Boolean                  '是否设置了录像压缩设置
Public blClosefrm As Boolean                         '是否关闭窗体
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Const MCI_RESUME = &H855                      '暂停后重新开始
Public Const PlayFPS = 15                            '每秒多少FPS
Function mGetCapSureDevice() As String
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
    Dim Device As String
    Dim VERSION As String
    Dim strtmp As String
    
    Device = String$(CAP_STRING_MAX, 0)
    VERSION = String$(CAP_STRING_MAX, 0)
    For Index = 0 To 8
        If 0 <> capGetDriverDescription(Index, Device, CAP_STRING_MAX, VERSION, CAP_STRING_MAX) Then
             strtmp = Left(Device, InStr(Device, vbNullChar) - 1) & Left$(VERSION, InStr(VERSION, vbNullChar) - 1)
             If Len(Trim(mGetCapSureDevice)) > 0 Then
                mGetCapSureDevice = mGetCapSureDevice & ";"
             End If
             mGetCapSureDevice = mGetCapSureDevice & strtmp
        End If
    Next
End Function


Function mConnCapDevice(ParentWindowWnd As Long, CapDeviceIndex As Integer) As Boolean
'-----------------------------------------------------------------------------------------
'功能：连接到设备
'参数：ParentWindowWnd 父窗体句柄 ; CapDeviceIndex 设备索引号
'返回：True = 成功 False = 失败
'上级函数或过程：
'下级函数或过程：capCreateCaptureWindow;SendMessageAsLong;SendMessageAsAny;SetWindowPos
'引用的外部参数：hCapWnd
'编制人：曾超
'修改人：
'-----------------------------------------------------------------------------------------
    Dim retVal As Boolean
    Dim strtmp() As String
    Dim i  As Integer
    Dim BITCapTureInfo As BITMAPINFO
    Dim intCaptureTYPE As Integer
    Dim intCaptureWidth As Integer
    Dim intCaptureHeight As Integer
    
    If hCapWnd = 0 Then
        hCapWnd = capCreateCaptureWindow("ZLSOFT_CAPTURE", WS_CHILD Or WS_VISIBLE, 0, 0, 100, 100, ParentWindowWnd, 0)
    End If
    
    If hCapWnd = 0 Then
        MsgBox "创建采集窗体失败！", vbInformation, "ZlPacsWork"
        Exit Function
    End If

    retVal = SendMessageAsLong(hCapWnd, WM_CAP_DRIVER_CONNECT, CapDeviceIndex, 0&)
    
    If retVal = False Then
        MsgBox "连接设备失败!", vbInformation, "ZlPacsWork"
        DestroyWindow hCapWnd
        Exit Function
    End If
    gintDeviceIndex = CapDeviceIndex
    
    SendMessage hCapWnd, WM_CAP_GET_VIDEOFORMAT, Len(BITCapTureInfo), BITCapTureInfo
    
    intCaptureTYPE = GetSetting("ZLSOFT", "私有模块\ZLHIS\" & App.ProductName & "\frmCapture", "CaptureType", 0)
    intCaptureWidth = GetSetting("ZLSOFT", "私有模块\ZLHIS\" & App.ProductName & "\frmCapture", "CaptureWidth", 0)
    intCaptureHeight = GetSetting("ZLSOFT", "私有模块\ZLHIS\" & App.ProductName & "\frmCapture", "CaptureHeight", 0)
    
    If intCaptureTYPE <> 0 And BITCapTureInfo.bmiHeader.biBitCount <> 0 Then
        With BITCapTureInfo.bmiHeader
            .biBitCount = intCaptureTYPE
            .biWidth = intCaptureWidth
            .biHeight = intCaptureHeight
            .biSizeImage = .biWidth * .biHeight * CInt(.biBitCount / 8)
        End With
        SendMessage hCapWnd, WM_CAP_SET_VIDEOFORMAT, 0, BITCapTureInfo
    End If

    Call SendMessageAsLong(hCapWnd, WM_CAP_SET_PREVIEWRATE, 66, 0&)

    Call SendMessageAsLong(hCapWnd, WM_CAP_SET_PREVIEW, -(True), 0&)
    
    
    Call mResizeCaptureWindow
               
    mConnCapDevice = True

End Function

Function mSelectCapDevice(CapDeviceIndex As Integer) As Boolean
'---------------------------------------------------------------------
'功能：连接到指定设备
'参数：CapDeviceIndex 设备索引(0--8)
'返回：True = 成功 False = 失败
'上级函数或过程：
'下级函数或过程：SendMessageAsLong;SetWindowPos
'引用的外部参数：hCapWnd
'编制人：曾超
'修改人：
'---------------------------------------------------------------------
    mSelectCapDevice = SendMessageAsLong(hCapWnd, WM_CAP_DRIVER_CONNECT, CapDeviceIndex, 0&)
    
    Call mResizeCaptureWindow
End Function
Function mParentWindowResize(ParentWindowWidth As Long, ParentWindowHeight As Long) As Boolean
'---------------------------------------------------------------------
'功能：设置显示窗口的位置在父窗体中心
'参数：ParentWindowWidth 父窗体宽度 ParentWindowHeight 父窗体高度
'返回：True = 成功 False = 失败
'上级函数或过程：
'下级函数或过程：SendMessageAsAny
'引用的外部参数：hCapWnd
'编制人：曾超
'修改人：
'---------------------------------------------------------------------
    Dim retVal As Boolean
    Dim capStat As CAPSTATUS
    Dim lngWidth As Long
    Dim lngHeight As Long
    
    retVal = SendMessageAsAny(hCapWnd, WM_CAP_GET_STATUS, Len(capStat), capStat)
    
    If retVal Then
        If ParentWindowWidth - capStat.uiImageWidth <= 0 Then
            lngWidth = ParentWindowWidth
        Else
            lngWidth = (ParentWindowWidth - capStat.uiImageWidth) / 2
        End If
        If ParentWindowHeight - capStat.uiImageHeight <= 0 Then
            lngHeight = ParentWindowHeight
        Else
            lngHeight = (ParentWindowHeight - capStat.uiImageHeight) / 2
        End If
        Call SetWindowPos(hCapWnd, 0&, lngWidth, lngHeight, 0&, 0&, SWP_NOSIZE Or SWP_NOZORDER Or SWP_NOSENDCHANGING)
    End If
    mParentWindowResize = True
End Function
Function mSaveImageFile(SavePath As String) As Boolean
'---------------------------------------------------------------------
'功能：保存当前显示的图像
'参数：SavePath=保存路径
'返回：True = 成功 False = 失败
'上级函数或过程：
'下级函数或过程：SendMessageAsString
'引用的外部参数：hCapWnd
'编制人：曾超
'修改人：
'---------------------------------------------------------------------
    mSaveImageFile = SendMessageAsString(hCapWnd, WM_CAP_FILE_SAVEDIB, 0&, SavePath)
End Function
Function mViewerFormat() As Boolean
'---------------------------------------------------------------------
'功能：显示图像格式
'参数：
'返回：True = 成功 False = 失败
'上级函数或过程：
'下级函数或过程：SendMessageAsLong
'引用的外部参数：hCapWnd
'编制人：曾超
'修改人：
'---------------------------------------------------------------------
    mViewerFormat = SendMessageAsLong(hCapWnd, WM_CAP_DLG_VIDEOFORMAT, 0&, 0&)
    mViewerFormat = mResizeCaptureWindow
End Function
Function mViewerSource() As Boolean
'---------------------------------------------------------------------
'功能：显示图像来源
'参数：
'返回：True = 成功 False = 失败
'上级函数或过程：
'下级函数或过程：SendMessageAsLong
'引用的外部参数：hCapWnd
'编制人：曾超
'修改人：
'---------------------------------------------------------------------
    mViewerSource = SendMessageAsLong(hCapWnd, WM_CAP_DLG_VIDEOSOURCE, 0&, 0&)
End Function
Function mResizeCaptureWindow() As Boolean
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

Function mGetCaptureWindowStatus() As CAPSTATUS
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
Sub mPaintPicture(DestionationhDC As Long, DestionationX As Long, DestionationY As Long, DestionationWidth As Long, _
    DestionationHeight As Long, SourcehDC As Long, Optional SourceX As Long = 0, Optional SourceY As Long = 0)
'---------------------------------------------------------------------
'功能：图片重绘
'参数：
'返回：
'上级函数或过程：
'下级函数或过程：SendMessageAsAny
'引用的外部参数：hCapWnd
'编制人：曾超
'修改人：
'---------------------------------------------------------------------
    BitBlt DestionationhDC, DestionationX, DestionationY, DestionationWidth, DestionationHeight, SourcehDC, SourceX, SourceY, DSTINVERT
    
End Sub
Sub mCapturePosition(CapX As Long, CapY As Long, CapWidth As Long, CapHeight As Long)
'---------------------------------------------------------------------
'功能：采集窗体位置设置
'参数：
'返回：
'上级函数或过程：
'下级函数或过程：SetWindowPos
'引用的外部参数：hCapWnd
'编制人：曾超
'修改人：
'---------------------------------------------------------------------
    Call SetWindowPos(hCapWnd, _
                0&, _
                CapX, _
                CapY, _
                CapWidth, _
                CapHeight, _
                SWP_NOMOVE Or SWP_NOZORDER Or SWP_NOSENDCHANGING)

End Sub
Sub mCopyImageToClipBoard()
'---------------------------------------------------------------------
'功能：复制图像到粘贴板
'参数：
'返回：
'上级函数或过程：
'下级函数或过程：SendMessageAsLong
'引用的外部参数：hCapWnd
'编制人：曾超
'修改人：
'---------------------------------------------------------------------
    Call SendMessageAsLong(hCapWnd, WM_CAP_EDIT_COPY, 0&, 0&)
End Sub

Function capFileSetCaptureFile(ByVal hCapWnd As Long, ByVal FilePath As String) As Boolean
'---------------------------------------------------------------------
'功能：设置采集录像的路径
'参数：
'返回：
'上级函数或过程：
'下级函数或过程：SendMessageAsString
'引用的外部参数：hCapWnd
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
Function capDlgVideoCompression(ByVal hCapWnd As Long) As Boolean

   capDlgVideoCompression = SendMessageAsLong(hCapWnd, WM_CAP_DLG_VIDEOCOMPRESSION, 0&, 0&)
End Function
Public Function Wndproc(ByVal hwnd As Long, ByVal MSG As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim hw As Long
    hw = FindWindow(vbNullString, "视频压缩")
    If hw <> 0 And blCompressionStup = False Then
        EnumChildWindows hw, AddressOf GetOkButton, 0
    End If
    Wndproc = CallWindowProc(preWinProc, hwnd, MSG, wParam, lParam)
End Function

Public Function GetOkButton(ByVal hwnd As Long, ByVal lParam As Long) As Long
    Dim strClassName As String
    Dim strWindowsName As String
    Dim strItemTxt As String
    Dim lngComboBox As Long
    Dim strLoadCompressionSetup As String
    On Error Resume Next
    strClassName = Space(255)
    strWindowsName = Space(255)
    strItemTxt = Space(255)

    If hwnd <> 0 Then
        GetClassName hwnd, strClassName, 255
        GetWindowText hwnd, strWindowsName, 255
        strClassName = Mid$(strClassName, 1, InStr(1, strClassName, Chr(0)) - 1)
        strWindowsName = Mid$(strWindowsName, 1, InStr(1, strWindowsName, Chr(0)) - 1)
        If strClassName = "ComboBox" Then
            If blClosefrm = False Then
                lngComboBox = CInt(Val(GetSetting("ZLSOFT", "私有模块\ZLHIS\" & App.ProductName & "\frmCapture", "ComPressionSetup")))
                SendMessage hwnd, CB_SETCURSEL, lngComboBox, 0
            Else
                lngComboBox = SendMessage(hwnd, CB_GETCURSEL, 0, 0)
                SendMessage hwnd, CB_GETLBTEXT, lngComboBox, ByVal strItemTxt
                strItemTxt = Mid$(strItemTxt, 1, InStr(1, strItemTxt, Chr(0)) - 1)
                SaveSetting "ZLSOFT", "私有模块\ZLHIS\" & App.ProductName & "\frmCapture", "ComPressionSetup", lngComboBox
            End If
        End If
        If strClassName = "Button" And strWindowsName = "确定" Then
'            AppActivate "视频压缩"
'            SendKeys "{ENTER}"
            SendMessage hwnd, &HF5, 0, 0
            blCompressionStup = True
        End If
    End If
    GetOkButton = 1
End Function
Public Function StatusProc(ByVal hCapWnd As Long, ByVal StatusCode As Long, ByVal lpStatusString As Long) As Long
    Dim strtmp As String
    Dim strTime As String
    Dim lngTmp As Long
    On Error Resume Next
    Select Case StatusCode
        Case 0 'this is recommended in docs
            'when zero is sent, clear old status messages
            'frmMain.Caption = App.Title
        Case IDS_CAP_END ' Video Capture has finished
            frmImgCapture.stbThis.Panels(2).Text = "状态:录像完成"
        Case IDS_CAP_STAT_VIDEOAUDIO, IDS_CAP_STAT_VIDEOONLY
            MsgBox "录像完成", vbInformation, "zl9ImgCapture"
        Case Else
            'use this function if you need a real VB string
            'frmMain.Caption = LPSTRtoVBString(lpStatusString)
            
            'or, just pass the LPCSTR to a WINAPI function
            Call SetWindowTextAsLong(frmImgCapture.txtState.hwnd, lpStatusString)
            frmImgCapture.txtState.Refresh
            strtmp = frmImgCapture.txtState
            strtmp = Mid(strtmp, 1, InStr(1, strtmp, "帧"))
            strTime = frmImgCapture.txtState
            If InStr(1, strTime, "帧") > 0 Then
                strTime = Mid(strTime, InStr(1, strTime, "在") + 1, InStr(1, strTime, ".") - 1 - InStr(1, strTime, "在"))
                frmImgCapture.stbThis.Panels(2).Text = "状态:采集中(按下鼠标左键或右键结束采集)" & strtmp & " 捕获时间：" & strLalcTime(CLng(Val(strTime)) * PlayFPS)
            End If
    End Select
    StatusProc = -(True) '- converts Boolean to C BOOL
End Function
Function capSetCallbackOnStatus(ByVal hCapWnd As Long, ByVal lpProc As Long) As Boolean
   capSetCallbackOnStatus = SendMessageAsLong(hCapWnd, WM_CAP_SET_CALLBACK_STATUS, 0&, lpProc)
End Function
Public Function strLalcTime(strTime As Long) As String
    '计算出当前播放的时间格式(00:00)
    Dim intHour As Integer
    Dim intMinute As Integer
    Dim intSecond As Integer
    Dim intTmp As Integer
    intSecond = (strTime / PlayFPS) Mod 60
    intMinute = Int((strTime / PlayFPS) / 60)
    intHour = Int((strTime / PlayFPS) / 60 / 60)
    strLalcTime = Format(intHour & ":" & intMinute & ":" & intSecond, "hh:mm:ss")
End Function


Function capSetScale(ByVal hCapWnd As Long, blnScale As Boolean) As Boolean
'---------------------------------------------------------------------
'功能：设置是否使用缩放方式采集
'参数：hCapWnd--采集窗口句柄；blnScale--True缩放，Fasle 不缩放
'返回：
'上级函数或过程：
'下级函数或过程：SendMessageAsString
'引用的外部参数：hCapWnd
'编制人：黄捷
'修改人：
'---------------------------------------------------------------------
    capSetScale = SendMessageAsLong(hCapWnd, WM_CAP_SET_SCALE, -(blnScale), 0&)
End Function
