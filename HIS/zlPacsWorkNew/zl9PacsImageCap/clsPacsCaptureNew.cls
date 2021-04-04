VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPacsCaptureNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


Private Const M_STR_PARAMETER_VIDEO_DOCK As String = "VideoDock"

Private WithEvents mobjVideoWindow As frmVideoCapture  ' frmWork_Video '视频采集窗口对象
Attribute mobjVideoWindow.VB_VarHelpID = -1
Private WithEvents mobjPane As Form                 '视频容器对象
Attribute mobjPane.VB_VarHelpID = -1
Private WithEvents mobjDock As Form                 '浮动窗口对象
Attribute mobjDock.VB_VarHelpID = -1

Private WithEvents mobjNotifyEvent As clsNotifyEvent
Attribute mobjNotifyEvent.VB_VarHelpID = -1

Private mRegEvent As ICaptureEvent

Private mstrParentWindowKey As String
Private mlngCurPicHeight As Long
Private mblnHasVideo As Boolean             '判断当前类中是否拥有视频采集模块
Private mblnInitFinishState As Boolean
Private mblnAllowEventNotify As Boolean
Private mblnRefreshState As Boolean
Private mobjEvent As Object
Private mblnIsReported As Boolean           '是否已经存在报告了

Public Event OnDockClose()

Property Get ModuleNo() As Long
    ModuleNo = glngModule
End Property

Property Let ModuleNo(ByVal lngValue As Long)
    glngModule = lngValue
    
    gobjCapturePar.IsUseAfterCapture = GetDeptPara(glngDepartId, "启用后台采集", 0)
    
    If Not mobjVideoWindow Is Nothing Then
'        Call mobjVideoWindow.UseAfterImageChanged(gobjCapturePar.IsUseAfterCapture And lngValue = 1291)
    End If
End Property


'容器对象句柄
Property Get ContainerHwnd() As Long
    ContainerHwnd = 0
    
    If Not mobjPane Is Nothing Then
        ContainerHwnd = mobjPane.hwnd
    End If
End Property

'容器对象
Property Get ContainerObj() As Object
    Set ContainerObj = mobjPane
End Property


'采集窗口句柄
Property Get VideoHwnd() As Long
    VideoHwnd = mobjVideoWindow.hwnd
End Property


'采集窗口对象
Property Get VideoWindow() As Object
    Set VideoWindow = mobjVideoWindow
End Property

'允许事件通知
Property Get AllowEventNotify() As Boolean
    AllowEventNotify = mblnAllowEventNotify
End Property


Property Let AllowEventNotify(value As Boolean)
    mblnAllowEventNotify = value
End Property

'文件加载方式
Property Get ImgLoadType() As FileLoadType
    ImgLoadType = gtFileLoadType
End Property

Property Let ImgLoadType(value As FileLoadType)
    gtFileLoadType = value
End Property

'获取当前进程ID
Property Get ProcessId() As Long
    ProcessId = GetCurrentProcessId
End Property

Property Get HasVideo() As Boolean
'判断该对象是否拥有视频采集
    HasVideo = mblnHasVideo
End Property

'浮动状态
Property Get VideoDockState() As Boolean
    VideoDockState = gblnDockingState
End Property

Property Get VideoTitle() As String
    If mobjVideoWindow Is Nothing Then
        VideoTitle = "视频采集"
    Else
        If mobjVideoWindow.LockState Then
'            VideoTitle = "【" & mobjVideoWindow.LockPatientName & "】视频采集"
        Else
            VideoTitle = "视频采集"
        End If
    End If
End Property

'保存使用采集模块的窗口关键字
Property Get ParentWindowKey() As String
    ParentWindowKey = mstrParentWindowKey
End Property

'是否已经存在报告了
Property Get IsReported() As Boolean
    IsReported = mblnIsReported
End Property
Property Let IsReported(ByVal value As Boolean)
    mblnIsReported = value
    If Not mobjVideoWindow Is Nothing Then
'        mobjVideoWindow.IsReported = value
    End If
End Property

Property Let ParentWindowKey(ByVal value As String)
    mstrParentWindowKey = value
End Property


Private Sub DoDockClose()
On Error GoTo errHandle
BUGEX_ "DoDockClose(clsPacsCapture) 1"
    If Not mobjEvent Is Nothing Then
BUGEX_ "DoDockClose(clsPacsCapture) 2"
        Call mobjEvent.OnDockClose
    ElseIf Not mRegEvent Is Nothing Then
BUGEX_ "DoDockClose(clsPacsCapture) 3"
    Else
BUGEX_ "DoDockClose(clsPacsCapture) 4"
        RaiseEvent OnDockClose
    End If
BUGEX_ "DoDockClose(clsPacsCapture) End"
Exit Sub
errHandle:
    BUGEX_ "DoDockClose(clsPacsCapture) Err:" & err.Description
End Sub


'******************************************************************************************************************************
'对外调用接口
'******************************************************************************************************************************


'显示浮动视频窗口
Public Sub zlShowPopupVideo()
    Dim lngPicHeight As Long
    Dim lngWindowStyle As Long
    
    '如果浮动采集被打开，则不允许再次打开采集界面
    If gblnDockingState = dsOpen Then Exit Sub
    
    '打开浮动采集窗口前，通知其他实例保存对应参数或状态
    Call mobjNotifyEvent.NotifyContainerLoad(mobjDock)

    Set mobjDock = New frmVideoDockWindow
    
'    If mobjVideoWindow.LockState = True Then
'        mobjDock.Caption = "视频采集 [" & mobjVideoWindow.LockPatientName & "]"
'    End If

    '如果处于调试状态，则需要在任务栏显示状态
    If gblnOpenDebug Then
        lngWindowStyle = GetWindowLong(mobjDock.hwnd, -20)
        lngWindowStyle = lngWindowStyle Or WS_EX_APPWINDOW
    
        SetWindowLong mobjDock.hwnd, -20, lngWindowStyle
    End If
    
    mobjDock.Caption = mobjVideoWindow.CaptionEx
    Call mobjDock.Show(0)
        
    lngPicHeight = Val(LoadVideoAreaCfg(M_STR_PARAMETER_VIDEO_DOCK))
    
    Call UpdateVideoContainer(mobjDock, lngPicHeight, True)
    
'    Call SetForegroundWindow(mobjDock.hwnd)
'    Call mobjVideoWindow.ucSplitter1.RePaint(False)
    
    '打开后需要刷新重绘视频显示界面
    Call mobjDock.Refresh
    
    gblnDockingState = dsOpen
    
    mblnHasVideo = True
End Sub


'当窗口进入ActiveWindow状态时，需要调用此方法
Public Sub zlRefreshVideoWindow()
On Error GoTo errHandle
    '如果浮动采集被打开，则不切换视频窗口的显示及刷新
BUGEX_ "zlRefreshVideoWindow 1"
    If gblnDockingState = dsOpen Then
    
        '注册打开的先后顺序，当关闭浮动窗口时，可以在正确的窗体中加载
        Call RegVideoWindowOrder
    
        Exit Sub
    End If
    
    If mblnRefreshState = True Then Exit Sub
    
    mblnRefreshState = True
    
    '需要先设置该变量值为false，避免重复刷新当前窗口时，重复触发mobjNotifyEvent_OnContainerLoad事件处理，
    mblnHasVideo = False
    
BUGEX_ "zlRefreshVideoWindow 2"
    '通知其他类实例，保存picCapture高度等操作
    Call mobjNotifyEvent.NotifyContainerLoad(mobjPane)
    
BUGEX_ "zlRefreshVideoWindow 3"
    '读取mobjVideoWindow.picCapture高度配置
    If mstrParentWindowKey <> "" And mlngCurPicHeight <= 0 Then
        '第一次加载窗口时，需要读取该参数值
        mlngCurPicHeight = Val(LoadVideoAreaCfg(mstrParentWindowKey))
    End If
        
BUGEX_ "zlRefreshVideoWindow 4"
    Call UpdateVideoContainer(mobjPane, mlngCurPicHeight)
    
    
    '在调用Refresh方法时，会触发mobjPane的Paint事件，因此需要先将mblnHasVideo设置为true
    mblnHasVideo = True
    
BUGEX_ "zlRefreshVideoWindow 5"
    '打开后需要刷新重绘视频显示界面
'    Call mobjPane.Refresh
    
BUGEX_ "zlRefreshVideoWindow 6"
    '设置内部的ZORDER处理
    Call RegVideoWindowOrder
    
    If gobjCapturePar.IsAutoRepreview Then Call zlRePreview
        
    '设置是否已经存在报告了
'    mobjVideoWindow.IsReported = mblnIsReported
    
    mblnRefreshState = False
    
BUGEX_ "zlRefreshVideoWindow End"

Exit Sub
errHandle:
    mblnRefreshState = False
    err.Raise err.Number, err.Source, err.Description, err.HelpFile, err.HelpContext
End Sub

'验证对象是否有效
Public Function IsValid() As Boolean
'当ActiveExe进程被意外退出后，所有方法将不可用
    IsValid = True
End Function


Public Sub zlInitModule(objDbAccess As Object, _
                        ByVal lngSysNo As Long, _
                        ByVal lngCallModule As Long, _
                        ByVal strPrivs As String, _
                        ByVal lngCurDepartId As Long, _
                        ByVal lngOwnerHwnd As Long, _
                        Optional objOwner As Object = Nothing, _
                        Optional blnForceRefresh As Boolean = False, _
                        Optional blnOpenDebugMode As Boolean = False)
'初始化模块参数
    Dim blnIsEqualDB As Boolean


    gblnOpenDebug = IIf(blnOpenDebugMode = False, gblnOpenDebug, blnOpenDebugMode)
    
BUGEX_ "zlInitModule 1", True
    If objDbAccess Is Nothing Then Exit Sub

        
BUGEX_ "zlInitModule 2", True
    '保存所需参数
    If Not (objOwner Is Nothing) Then
        Set gobjOwner = objOwner
    End If
    
    glngSys = lngSysNo
    glngModule = lngCallModule
    gstrPrivs = strPrivs
    glngDepartId = lngCurDepartId
        
BUGEX_ "zlInitModule3:Connect DB.", True
    '初始化用户控件及数据库连接
    Call InitCommonLib(objDbAccess)
    
BUGEX_ "zlInitModule 4", True
    glngRootHandle = lngOwnerHwnd
    If lngOwnerHwnd <= 0 Then
        glngRootHandle = GetAncestor(mobjPane.hwnd, GA_ROOT)
    End If
BUGEX_ "zlInitModule 5", True
    Call gobjCapturePar.ReadParameter
    
BUGEX_ "zlInitModule 6", True
    If blnForceRefresh And Not VideoIsNull Then
BUGEX_ "zlInitModule 7", True
        Call mobjVideoWindow.zlInitModule
    End If
    
    mblnInitFinishState = True
    
BUGEX_ "zlInitModule End", True
End Sub


Public Sub zlUpdateStudyInf(ByVal lngAdviceId As Long, _
                            ByVal lngSendNo As Long, _
                            ByVal lngStudyState As Long, _
                            ByVal blnMoved As Boolean, _
                            ByVal blnIsReported As Boolean)
'更新采集时所需的检查信息
    If VideoIsNull = True Then Exit Sub
    
'    Call mobjVideoWindow.zlUpdateAdviceInf(lngAdviceId, lngSendNo, lngStudyState, blnMoved)

    IsReported = blnIsReported
    
    If gblnDockingState = dsOpen Then
        If mobjDock Is Nothing Then Exit Sub
        mobjDock.Caption = mobjVideoWindow.CaptionEx
    End If
End Sub


Public Sub zlCaptureImg()
'图像采集
    If VideoIsNull = True Then Exit Sub
    
    Call mobjVideoWindow.ForeCapture(True)
End Sub

Public Sub zlRePreview()
'重新进入预览
    Call mobjVideoWindow.zlRePreview
End Sub

Public Sub zlCaptureAfterImg()
'后台采集
    If VideoIsNull = True Then Exit Sub
    If Not gobjCapturePar.IsUseAfterCapture Then Exit Sub
    
    Call mobjVideoWindow.AfterCapture
End Sub

Public Sub zlUpdateAfterCaptureInfo()
'更新标记
    If VideoIsNull = True Then Exit Sub
    If Not gobjCapturePar.IsUseAfterCapture Then Exit Sub

    Call mobjVideoWindow.ResetAfterCaptureTag
End Sub

Public Sub zlRefreshData(Optional blnForceRefresh As Boolean = False)
'刷新图像采集界面
BUGEX_ "zlRefreshData 1"
    If VideoIsNull = True Then Exit Sub
    
BUGEX_ "zlRefreshData 2"
    Call mobjVideoWindow.zlRefreshFace(blnForceRefresh)
    
BUGEX_ "zlRefreshData 3"
End Sub


Public Sub zlNotifyRefresh()
'通知界面数据刷新
    If VideoIsNull = True Then Exit Sub
    
    Call mobjVideoWindow.NotificationRefresh
End Sub

Public Sub zlNotifyQuit()
'通知退出Active视频模块
    gblnIsQuitModule = True
End Sub


Public Sub zlShowVideoConfig()
'显示视频配置窗口

BUGEX_ "zlShowVideoConfig 1"
    If VideoIsNull = True Then Exit Sub
    
'    If gblnDockingState = dsOpen Then Exit Sub
    
BUGEX_ "zlShowVideoConfig 2"
    Call mobjVideoWindow.ShowVideoConfig
BUGEX_ "zlShowVideoConfig End"
End Sub



'******************************************************************************************************************************
'内部私有处理方法
'******************************************************************************************************************************

Private Function VideoIsNull() As Boolean
    VideoIsNull = IIf(gobjVideoNew Is Nothing, True, False)
End Function


Private Sub RegVideoWindowOrder()
'注册视频窗口打开顺序
    If Not gobjZOrder Is Nothing Then
        Call DelCurZOrderInf
        Call gobjZOrder.Add(mobjPane.hwnd, mobjPane.hwnd)
        
'        Debug.Print "RegHwnd:" & mobjPane.hwnd
    End If
End Sub

Private Sub CreateGlobalResource()
'初始化视频采集对象
    
BUGEX_ "CreateGlobalResource 1", True
    If glngInstanceCount <= 0 Then

        gblnOpenDebug = GetSetting("ZLSOFT", G_STR_REG_PATH_PUBLIC, "IsOpenDebug", 0)
        Call SaveSetting("ZLSOFT", G_STR_REG_PATH_PUBLIC, "IsOpenDebug", IIf(gblnOpenDebug, 1, 0))
    
        If gblnOpenDebug Then
BUGEX_ "CreateResource Open Debug Successed", True
        Else
BUGEX_ "CreateResource Open Debug Failed", True
        End If
    End If
    
    '创建公共资源
    If gobjCapturePar Is Nothing Then
BUGEX_ "CreateGlobalResource gobjCapturePar", True
        Set gobjCapturePar = New clsCaptureParameter
    End If
    
BUGEX_ "CreateGlobalResource 2", True
    If gobjZOrder Is Nothing Then
BUGEX_ "CreateGlobalResource gobjZOrder", True
        Set gobjZOrder = New Scripting.Dictionary
    End If
    

BUGEX_ "CreateGlobalResource 3", True
    If gobjVideoNew Is Nothing Then
BUGEX_ "CreateGlobalResource gobjVideoNew", True
        Set gobjVideoNew = New frmVideoCapture
    End If
    
BUGEX_ "CreateGlobalResource 4", True
    If gobjNotifyEvent Is Nothing Then
BUGEX_ "CreateGlobalResource gobjNotifyEvent", True
        Set gobjNotifyEvent = New clsNotifyEvent
    End If
    
BUGEX_ "CreateGlobalResource End", True

End Sub



Private Sub Class_Initialize()
On Error GoTo errHandle
BUGEX_ "Class_Initialize 1", True
    '初始化模块变量
    mstrParentWindowKey = ""
    mlngCurPicHeight = 0
    mblnHasVideo = False
    mblnInitFinishState = False
    gblnIsQuitModule = False
    mblnAllowEventNotify = False
    
BUGEX_ "Class_Initialize 2"
    
    '创建公共资源
    Call CreateGlobalResource
    
BUGEX_ "Class_Initialize 3"
    '创建类资源
    If mobjPane Is Nothing Then
        BUGEX_ "CreateResource mobjPane", True
        Set mobjPane = New frmVideoPaneWindow
    End If
    
BUGEX_ "Class_Initialize 4"
    '设置所需对象变量
    Set mobjVideoWindow = gobjVideoNew
    Set mobjNotifyEvent = gobjNotifyEvent
    
    glngInstanceCount = glngInstanceCount + 1
    
BUGEX_ "Class_Initialize End", True
Exit Sub
errHandle:
    BUGEX_ "Class_Initialize Err:" & err.Description, True, True
End Sub


Public Sub FreeGlobalResource()
'释放视频采集资源
        
BUGEX_ "FreeGlobalSource 1", True
        '释放事件通知对象
        If Not gobjNotifyEvent Is Nothing Then
            Set gobjNotifyEvent = Nothing
        End If
                
BUGEX_ "FreeGlobalSource 2", True
        '只有实例数量为一的时候，才释放全局公共资源
        If Not gobjVideoNew Is Nothing Then
            Unload gobjVideoNew
            Set gobjVideoNew = Nothing
        End If
        
BUGEX_ "FreeGlobalSource 3", True
        If Not gobjZOrder Is Nothing Then
            Call gobjZOrder.RemoveAll
            Set gobjZOrder = Nothing
        End If
        
BUGEX_ "FreeGlobalSource 4", True

        '设置参数配置对象
        If Not gobjCapturePar Is Nothing Then
            Set gobjCapturePar = Nothing
        End If
        
BUGEX_ "FreeGlobalSource 5", True
        '释放oracle对象
        If Not gcnVideoOracle Is Nothing Then
            Set gcnVideoOracle = Nothing
        End If
        
BUGEX_ "FreeGlobalSource End", True
End Sub

Private Sub Class_Terminate()
'释放资源
    Dim i As Integer
On Error GoTo errHandle

BUGEX_ "FreeSource 1", True
    Call DelCurZOrderInf
    
    '释放类资源--------------------------------------
    
    '释放事件处理
    If Not mRegEvent Is Nothing Then
        Set mRegEvent = Nothing
    End If
    
    
BUGEX_ "FreeSource 2", True
    '断开视频窗口与容器窗口的连接
    If glngCurVideoContainerHwnd = mobjPane.hwnd Then
        mobjVideoWindow.Visible = False
        SetParent mobjVideoWindow.hwnd, 0
    End If
    
    '释放弹出窗口
    If Not mobjDock Is Nothing Then
        Unload mobjDock
        Set mobjDock = Nothing
    End If
    
    '释放容器窗口
    If Not mobjPane Is Nothing Then
        Unload mobjPane
        Set mobjPane = Nothing
    End If
    
    
    '释放全局资源--------------------------------------
    If glngInstanceCount = 1 Then Call FreeGlobalResource
    
    
    Set mobjNotifyEvent = Nothing
    Set mobjVideoWindow = Nothing
    Set mobjEvent = Nothing
    
BUGEX "zlPacsCapture Class_Terminate"

    glngInstanceCount = glngInstanceCount - 1
    glngModule = 0
errHandle:
End Sub

'更新当前视频窗口大小
Private Sub UpdateVideoContainer(objPane As Object, Optional ByVal lngPicHeight As Long = 0, Optional ByVal blnIsDock As Boolean = False)
On Error GoTo errHandle
    Dim lngParentHwnd As Long
    Dim lngActiveHwnd As Long
    
    If VideoIsNull = True Then Exit Sub
    
    lngParentHwnd = GetAncestor(objPane.hwnd, GA_PARENT)
    
BUGEX_ "UpdateVideoContainer 1"

    If objPane Is Nothing Then Exit Sub
    If (lngParentHwnd <= 0 Or lngParentHwnd = GetDesktopWindow) And Not blnIsDock Then Exit Sub
    
BUGEX_ "UpdateVideoContainer 2"

    If GetAncestor(mobjVideoWindow.hwnd, GA_PARENT) <> objPane.hwnd Then
        mobjVideoWindow.Visible = False
        SetParent mobjVideoWindow.hwnd, 0
'    BUGEX_ "UpdateVideoContainer 3 picCaptureHeight:" & mobjVideoWindow.picCapture.Height
    
        mobjVideoWindow.Left = 0
        mobjVideoWindow.Top = 0
    
        mobjVideoWindow.Height = objPane.Height
        mobjVideoWindow.Width = objPane.Width
    
        If blnIsDock Then
            mobjVideoWindow.Width = mobjVideoWindow.Width - 160
            mobjVideoWindow.Height = mobjVideoWindow.Height - 500
        End If
        
'        Call mobjVideoWindow.RefreshImageCaptureFace(blnIsDock)
    
    BUGEX_ "UpdateVideoContainer 4"
        SetParent mobjVideoWindow.hwnd, objPane.hwnd
        Call ShowObject(mobjVideoWindow)
    Else
        Call MoveWindow(mobjVideoWindow.hwnd, 0, 0, _
                        mobjVideoWindow.ScaleX(objPane.Width, vbTwips, vbPixels), _
                        mobjVideoWindow.ScaleY(objPane.Height, vbTwips, vbPixels), _
                        0)
    End If

    
BUGEX_ "UpdateVideoContainer 5"
    '如果mobjPane设定了parent，则可进行显示
    If objPane.Tag <> "1" Then Call ShowObject(objPane)
    
    '需要在这里设置mobjVideoWindow.picCapture的高度，如果在后面进行设置可能造成mobjVideoWindow.picCapture高度不能生效
'    mobjVideoWindow.Height = IIf(lngPicHeight > 0, lngPicHeight, mobjVideoWindow.Height)

    lngActiveHwnd = GetTopHwnd(objPane.hwnd)
'    Call SetFocusEx(GetTopHwnd(objPane.hwnd))
'    Call SetActiveWindow(lngActiveHwnd)
BUGEX_ "UpdateVideoContainer 6 ActiveHwnd:" & lngActiveHwnd
    
    glngCurVideoContainerHwnd = objPane.hwnd
        
BUGEX_ "UpdateVideoContainer End"
    Exit Sub
errHandle:
    BUGEX_ "UpdateVideoContainer Err:" & err.Description
End Sub


Private Sub ShowObject(objShow As Object)
'显示object对象
On Error GoTo errHandle
    objShow.Visible = True
Exit Sub
errHandle:
    ShowWindow objShow.hwnd, 9
End Sub


Private Sub mobjDock_Load()
    '恢复窗口状态
BUGEX_ "mobjDock_Load 1"
    SetWindowPos mobjDock.hwnd, -1, mobjDock.CurrentX, mobjDock.CurrentY, mobjDock.ScaleWidth, mobjDock.ScaleHeight, 3 '将窗口置顶
    
BUGEX_ "mobjDock_Load 2"
    Call RestoreWinState(mobjDock, App.ProductName)
    
BUGEX_ "mobjDock_Load End"
End Sub

Private Sub mobjDock_Resize()
On Error GoTo errHandle
    '浮动窗口大小改变时，需要调整视频采集窗口大小
    '如果窗口进行最小化时，不进行视频布局调整
    If Not mblnInitFinishState Then Exit Sub
    If mobjDock.WindowState = 1 Then Exit Sub
    
    Call UpdateVideoContainer(mobjDock, 0, True)
    
    Exit Sub
errHandle:
    BUGEX_ "mobjDock_Resize Err:" & err.Description
End Sub

Private Sub BUGEX_(ByVal strInf As String, Optional ByVal blnIsForce As Boolean = False, Optional ByVal blnRaiseErr As Boolean = False)
'输出调试消息
    If gblnOpenDebug Then blnIsForce = True
    
    If mobjPane Is Nothing Then
        BUGEX "HWND:" & 0 & " " & strInf, blnIsForce
    Else
        BUGEX "HWND:" & mobjPane.hwnd & " " & strInf, blnIsForce
    End If
    
    '继续触发错误
    If blnRaiseErr Then
        err.Raise err.Number, err.Source, err.Description, err.HelpFile, err.HelpContext
    End If
End Sub

Private Sub mobjNotifyEvent_OnContainerLoad(objOldPane As Object)
On Error GoTo errHandle
BUGEX_ "mobjNotifyEvent_OnContainerLoad 1"
    If VideoIsNull = True Then Exit Sub
    
BUGEX_ "mobjNotifyEvent_OnContainerLoad 2"
    '在其他容器界面中打开视频采集页面显示之前，需要先保存当前采集模块的缩略图高度
    If mblnHasVideo And gblnDockingState = dsClosed Then
BUGEX_ "mobjNotifyEvent_OnContainerLoad 3"
        mlngCurPicHeight = mobjVideoWindow.Height
        
        If mstrParentWindowKey <> "" Then
BUGEX_ "mobjNotifyEvent_OnContainerLoad 4 mlngCurPicHeight:" & mlngCurPicHeight
            Call SaveVideoAreaCfg(mstrParentWindowKey, mlngCurPicHeight)
        End If
        
        mblnHasVideo = False
    End If
    
BUGEX_ "mobjNotifyEvent_OnContainerLoad End"
Exit Sub
errHandle:
    MsgboxCus err.Description, vbOKOnly, G_STR_HINT_TITLE
End Sub

Private Sub mobjNotifyEvent_OnContainerUnLoad(objOldPane As Object)
    Dim lngParentHwnd As Long
    
    If VideoIsNull = True Then Exit Sub                       '当视频对象无效时，则不切换视频显示
    If gblnDockingState = dsOpen Then Exit Sub          '当处于浮动窗口显示时，则不切换视频显示
    If objOldPane.hwnd = mobjPane.hwnd Then Exit Sub    '当退出的窗口句柄与当前对象句柄相同时，则不切换显示视频
    If mobjNotifyEvent.IsChangePane Then Exit Sub       '当视频已经被切换时，则不需要在重复的对视频进行切换
    
    If gobjZOrder.Count >= 1 Then
BUGEX_ "mobjNotifyEvent_OnContainerUnLoad 1"
        '由于已经加载的视频窗口或者对象在释放之前，都会先移除在gobjZOrder中的句柄信息，
        '因此只需判断移除后所需加载的视频窗口是否为该列表的最后一个窗口句柄即可
        If mobjPane.hwnd <> gobjZOrder.Keys(gobjZOrder.Count - 1) Then Exit Sub
BUGEX_ "mobjNotifyEvent_OnContainerUnLoad 2"
'    Else
'BUGEX_ "mobjNotifyEvent_OnContainerUnLoad 3"
'        '判断当前的父窗口句柄是否为之前视频窗口的下一个窗口句柄。
'        If GetTopHwnd(mobjPane.hWnd) <> glngNextVideoContainerHwnd Then Exit Sub
'
'        lngParentHwnd = GetAncestor(mobjPane.hWnd, GA_PARENT)
'
'        '当mobjpane没有父窗口时，即只创建了clspacscapture类，并没有使用该类的containerhwnd属性将该控件附加到界面中，
'        '因此当关闭弹出窗口后，不需要将视频在附加到mobjpanel上
'        If lngParentHwnd <= 0 Or lngParentHwnd = GetDesktopWindow Then Exit Sub
'BUGEX_ "mobjNotifyEvent_OnContainerUnLoad 4"
    End If
    
    lngParentHwnd = GetAncestor(mobjPane.hwnd, GA_PARENT)
    
    '当mobjpane没有父窗口时，即只创建了clspacscapture类，并没有使用该类的containerhwnd属性将该控件附加到界面中，
    '因此当关闭弹出窗口后，不需要将视频在附加到mobjpanel上
    If lngParentHwnd <= 0 Or lngParentHwnd = GetDesktopWindow Or mobjPane.Visible = False Then
        If gobjZOrder.Count >= 1 Then Call gobjZOrder.Remove(gobjZOrder.Keys(gobjZOrder.Count - 1))
        Exit Sub
    End If
        
    Call zlRefreshVideoWindow
    mobjNotifyEvent.IsChangePane = True
    
BUGEX_ "mobjNotifyEvent_OnContainerUnLoad End"
End Sub


Private Sub mobjDock_Unload(Cancel As Integer)
On Error GoTo errHandle
    Dim lngParentHwnd As Long
    
    gblnDockingState = dsClosing
    
BUGEX_ "mobjDock_Unload 1"

    If VideoIsNull = True Then Exit Sub
    
BUGEX_ "mobjDock_Unload 2"
    Call SaveWinState(mobjDock, App.ProductName)
    Call SaveVideoAreaCfg(M_STR_PARAMETER_VIDEO_DOCK, mobjVideoWindow.Height)
    
    mobjVideoWindow.Visible = False
    SetParent mobjVideoWindow.hwnd, 0

'    glngNextVideoContainerHwnd = GetTopHwnd(mobjPane.hWnd)
    
    lngParentHwnd = GetAncestor(mobjPane.hwnd, GA_PARENT)
    
BUGEX_ "mobjDock_Unload 3"
    '当关闭弹出窗口时，且只创建了clsPacsCapture的唯一实例，同时并没有将mobjPane容器组件通过ContainerHwnd属性附加到其他界面中，
    '因此在这种情况下，可以直接释放视频采集资源
    If glngInstanceCount <= 1 And (lngParentHwnd <= 0 Or lngParentHwnd = GetDesktopWindow) Then
BUGEX_ "mobjDock_Unload 4"
        Unload mobjVideoWindow
        Set mobjVideoWindow = Nothing
        
    Else
BUGEX_ "mobjDock_Unload 5"
        If gblnIsQuitModule = False Then
            '通知其他实例加载视频模块显示
            Call mobjNotifyEvent.NotifyContainerUnLoad(mobjDock)
        End If
    End If
    
BUGEX_ "mobjDock_Unload 6"
    Set mobjDock = Nothing
    
    gblnDockingState = dsClosed
    
    Call DoDockClose
    
BUGEX_ "mobjDock_Unload End"
Exit Sub
errHandle:
    MsgboxCus err.Description, vbOKOnly, G_STR_HINT_TITLE
End Sub


Private Sub mobjPane_Activate()
'On Error GoTo errHandle
'    Dim lngParentHwnd As Long
'
'BUGEX_ "mobjPane_Activate 1"
'    If VideoIsNull = True Then Exit Sub
'
'BUGEX_ "mobjPane_Activate 2"
'    lngParentHwnd = GetAncestor(mobjPane.hWnd, GA_PARENT)
'
'    '如果浮动采集被打开，则不切换视频窗口的显示及刷新
'    If Not mblnInitFinishState Or lngParentHwnd <= 0 Or lngParentHwnd = GetDesktopWindow Then Exit Sub
'    If gblnDockingState = dsOpen Then Exit Sub      '如果浮动窗口打开，则不刷新视频窗口
'
'BUGEX_ "mobjPane_Activate 3"
'    If mblnHasVideo = True Then
'BUGEX_ "mobjPane_Activate 4"
'        '如果该容器组件内已经加载了视频，则只需重新调整对应界面大小即可
'        '容器窗口大小改变时，需要调整视频采集窗口大小
'        Call UpdateVideoContainer(mobjPane, mlngCurPicHeight)
''        mlngCurPicHeight = mobjVideoWindow.picCapture.Height
'    Else
'BUGEX_ "mobjPane_Activate 5"
'        Call zlRefreshVideoWindow
'    End If
'
'BUGEX_ "mobjPane_Activate End"
'    Exit Sub
'errHandle:
End Sub


Private Sub mobjPane_DblClick()
On Error GoTo errHandle
    Dim lngParentHwnd As Long
    
    If VideoIsNull = True Then Exit Sub
    
    lngParentHwnd = GetAncestor(mobjPane.hwnd, GA_PARENT)
    
    '如果浮动窗口打开，则不恢复当前容器的视频显示
    If Not mblnInitFinishState Or lngParentHwnd <= 0 Or lngParentHwnd = GetDesktopWindow Then Exit Sub
    If gblnDockingState = dsOpen Then
        Call RegVideoWindowOrder
        
        MsgboxCus "浮动窗口已打开，不能在此显示视频。", vbOKOnly + vbInformation, G_STR_HINT_TITLE
        Exit Sub
    End If
    
    Call zlRefreshVideoWindow
Exit Sub
errHandle:
End Sub

Private Sub DrawBackground()
'绘制背景图像
    Dim i As Integer
    Dim Count As Integer
    Dim wordRect As RECT
    
    Count = 2
    wordRect.Bottom = 45
    wordRect.Right = 200

    If mobjPane.picBackImg.Height * 3 >= mobjPane.Height Then Count = 1

    Call mobjPane.Cls
    
    For i = 0 To Count
        Call mobjPane.PaintPicture(mobjPane.picBackImg.Picture, _
            Round(mobjPane.Width / (i + 1)) - mobjPane.picBackImg.Width + 200, _
            Round((mobjPane.Height / 3) * (i + 1) - mobjPane.picBackImg.Height), _
            mobjPane.picBackImg.Width, mobjPane.picBackImg.Height)

        If mobjVideoWindow Is Nothing Then
            wordRect.Left = mobjPane.ScaleX(Round(mobjPane.Width / (i + 1)) - mobjPane.picBackImg.Width, vbTwips, vbPixels) + 35
        Else
            wordRect.Left = mobjPane.ScaleX(Round(mobjPane.Width / (i + 1)) - mobjPane.picBackImg.Width, vbTwips, vbPixels) + 17
        End If
        
        wordRect.Top = mobjPane.ScaleY(Round((mobjPane.Height / 3) * (i + 1) - mobjPane.picBackImg.Height), vbTwips, vbPixels) - 30

        wordRect.Right = wordRect.Left + 200
        wordRect.Bottom = wordRect.Top + 90

        If mobjVideoWindow Is Nothing Then
            Call DrawText(mobjPane.hdc, "视频未被注册" & vbCrLf & "已禁用视频源", 27, wordRect, 0)
        Else
            Call DrawText(mobjPane.hdc, "视频已被其他窗口占用" & vbCrLf & " 双击鼠标可恢复显示", 44, wordRect, 0)
        End If
    Next i
End Sub

Private Sub mobjPane_Paint()
On Error GoTo errHandle

BUGEX_ "mobjPane_Paint 1"
    
    If mobjPane.Visible = False Then Exit Sub
    
    '如果当前的mobjPane没有嵌入视频采集，则重绘背景图像
    If mblnHasVideo = False Or glngCurVideoContainerHwnd <> mobjPane.hwnd Then
BUGEX_ "mobjPane_Paint 2"
        Call DrawBackground
    Else
BUGEX_ "mobjPane_Paint 3"
'        Call SetFocusEx(GetTopHwnd(mobjPane.hwnd))
'        Call SetActiveWindow(GetTopHwnd(mobjPane.hwnd))

        '当mobjPane_Raise中不触发视频显示后，则不需要更新视频界面大小
'        If mobjVideoWindow.Width <> mobjPane.Width Or _
'            mobjVideoWindow.Height <> mobjPane.Height Then
'            Call UpdateVideoContainer(mobjPane, mlngCurPicHeight)
'        End If
        
'        mobjVideoWindow.TimerRePaint.Enabled = True
    End If
    
BUGEX_ "mobjPane_Paint End"

Exit Sub
errHandle:
    
BUGEX_ "mobjPane_Paint Err:" & err.Description
End Sub

Private Sub mobjPane_Resize()
On Error GoTo errHandle
    Dim lngParentHwnd As Long
    
BUGEX_ "mobjPane_Resize 1"
    If VideoIsNull = True Then Exit Sub
    
BUGEX_ "mobjPane_Resize 2"
    lngParentHwnd = GetAncestor(mobjPane.hwnd, GA_PARENT)
    
    '如果浮动采集被打开，则不切换视频窗口的显示及刷新
    If Not mblnInitFinishState Or lngParentHwnd <= 0 Or lngParentHwnd = GetDesktopWindow Then Exit Sub
'    If mobjPane.Visible = False Then Exit Sub       '如果当前窗口未显示，则不刷新视频窗口
    If gblnDockingState = dsOpen Then Exit Sub      '如果浮动窗口打开，则不刷新视频窗口

BUGEX_ "mobjPane_Resize 3"
    If mblnHasVideo = True Then
BUGEX_ "mobjPane_Resize 4"

        '如果是调整界面大小，但本身视频页面被隐藏，则调整大小时，不自动显示视频页面
        mobjPane.Tag = IIf(mobjPane.Visible = False, "1", "")
        
        '如果该容器组件内已经加载了视频，则只需重新调整对应界面大小即可
        '容器窗口大小改变时，需要调整视频采集窗口大小
        Call UpdateVideoContainer(mobjPane, mlngCurPicHeight)
'        mlngCurPicHeight = mobjVideoWindow.picCapture.Height
    Else
'BUGEX_ "mobjPane_Resize 5"
'        Call zlRefreshVideoWindow
    End If
    
BUGEX_ "mobjPane_Resize End"
    Exit Sub
errHandle:
    BUGEX_ "mobjPane_Resize Err:" & err.Description
End Sub

Private Sub DelCurZOrderInf()
'从gobjZOrder中删除当前容器窗口的句柄信息
    If Not gobjZOrder Is Nothing Then
        If gobjZOrder.Exists(mobjPane.hwnd) Then
            Call gobjZOrder.Remove(mobjPane.hwnd)
'            Debug.Print "RemoveHwnd:" & mobjPane.hwnd & " ItemCount:" & gobjZOrder.Count
            
        End If
    End If
End Sub



Public Sub RegEventObj(objEvent As Object)
'注册事件对象
    Set mobjEvent = objEvent
    
    Set mRegEvent = Nothing
End Sub


Private Sub mobjPane_Unload(Cancel As Integer)
On Error GoTo errHandle
    If VideoIsNull = True Then Exit Sub
    
'    Debug.Print mobjPane.hwnd
    '如果当前容器控件中包含视频采集界面，则在容器卸载时，需要通知其他界面加载视频,反之则不需要通知。
    If glngCurVideoContainerHwnd <> mobjPane.hwnd Then Exit Sub
    If mblnHasVideo = False Then Exit Sub
    
    '在卸载当前窗口时，如果包含视频窗口，则保存当前窗口的视频高度
    If mstrParentWindowKey <> "" Then
        Call SaveVideoAreaCfg(mstrParentWindowKey, mobjVideoWindow.Height)
    End If
    
    mobjVideoWindow.Visible = False
    SetParent mobjVideoWindow.hwnd, 0
    
'    Debug.Print mobjPane.hwnd
    '需要移除在gobjZOrder中存在的句柄信息
    Call DelCurZOrderInf
    
'    glngNextVideoContainerHwnd = GetNextWindowCus(GetTopHwnd(mobjPane.hWnd))

    '只有clsPacsCapture的实例唯一时，可直接在mobjpanel的unload事件中释放视频采集资源
    If glngInstanceCount <= 1 Then
        Unload mobjVideoWindow
        Set mobjVideoWindow = Nothing
    Else
        If gblnIsQuitModule = False Then
            '通知其他实例加载视频模块显示
            Call mobjNotifyEvent.NotifyContainerUnLoad(mobjPane)
        End If
    End If
    
Exit Sub
errHandle:
    MsgboxCus err.Description, vbOKOnly, G_STR_HINT_TITLE
End Sub


Private Sub mobjVideoWindow_OnControlResize(objControl As Object)
On Error GoTo errHandle
BUGEX_ "mobjVideoWindow_OnControlResize 1"

    If mblnHasVideo = False Then Exit Sub
    If gblnDockingState = dsOpen Then Exit Sub
    
    If objControl.Name = "picCapture" Then mlngCurPicHeight = objControl.Height
BUGEX_ "mobjVideoWindow_OnControlResize End"
Exit Sub
errHandle:
    BUGEX_ "mobjVideoWindow_OnControlResize Err:" & err.Description
End Sub

Public Sub setFontSize(ByVal intsize As Integer)
'设置字号
    If Not mobjVideoWindow Is Nothing Then
'        Call mobjVideoWindow.setFontSize(intsize)
    End If
End Sub


Public Sub LockStudy(ByVal intType As Integer, ByVal lngAdviceId As Long, ByVal lngSendNo As Long, _
ByVal lngStudyState As Long, ByVal blnMoved As Boolean)
'锁定 或者 解除锁定
'intType=1: 锁定   2：解除锁定
'    If Not mobjVideoWindow Is Nothing Then Call mobjVideoWindow.LockStudy(intType, lngAdviceId, lngSendNo, lngStudyState, blnMoved)
End Sub













