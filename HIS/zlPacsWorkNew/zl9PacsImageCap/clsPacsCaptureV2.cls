VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPacsCaptureV2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


Private Const M_STR_PARAMETER_VIDEO_DOCK As String = "VideoDock"

Private WithEvents mobjVideoWindow As frmVideoCaptureV2  ' frmWork_Video '��Ƶ�ɼ����ڶ���
Attribute mobjVideoWindow.VB_VarHelpID = -1
Private WithEvents mobjPane As Form                 '��Ƶ��������
Attribute mobjPane.VB_VarHelpID = -1
Private WithEvents mobjDock As Form                 '�������ڶ���
Attribute mobjDock.VB_VarHelpID = -1

Private WithEvents mobjNotifyEvent As clsNotifyEvent
Attribute mobjNotifyEvent.VB_VarHelpID = -1

Private mRegEvent As ICaptureEvent

Private mstrParentWindowKey As String
Private mlngCurPicHeight As Long
Private mblnHasVideo As Boolean             '�жϵ�ǰ�����Ƿ�ӵ����Ƶ�ɼ�ģ��
Private mblnInitFinishState As Boolean
Private mblnAllowEventNotify As Boolean
Private mblnRefreshState As Boolean
'Private mobjEvent As Object
Private mobjCapHelper As ICapHelper

Private WithEvents mobjPlay As frmPlaying
Attribute mobjPlay.VB_VarHelpID = -1

'Public Event OnDockClose()

Property Get DeptId() As Long
    DeptId = glngDepartId
End Property


Property Get ModuleNo() As Long
    ModuleNo = glngModule
End Property

Property Let ModuleNo(ByVal lngValue As Long)
    glngModule = lngValue
    
    gobjCapturePar.IsUseAfterCapture = GetDeptPara(glngDepartId, "���ú�̨�ɼ�", 0)
    
    If Not mobjVideoWindow Is Nothing Then
'        Call mobjVideoWindow.UseAfterImageChanged(gobjCapturePar.IsUseAfterCapture And lngValue = 1291)
    End If
End Property


'����������
Property Get ContainerHwnd() As Long
    ContainerHwnd = 0
    
    If Not mobjPane Is Nothing Then
        ContainerHwnd = mobjPane.hwnd
    End If
End Property

'��������
Property Get ContainerObj() As Object
    Set ContainerObj = mobjPane
End Property


'�ɼ����ھ��
Property Get VideoHwnd() As Long
    VideoHwnd = mobjVideoWindow.hwnd
End Property


'�ɼ����ڶ���
Property Get VideoWindow() As Object
    Set VideoWindow = mobjVideoWindow
End Property

'�����¼�֪ͨ
Property Get AllowEventNotify() As Boolean
    AllowEventNotify = mblnAllowEventNotify
End Property


Property Let AllowEventNotify(value As Boolean)
    mblnAllowEventNotify = value
End Property

'�ļ����ط�ʽ
Property Get ImgLoadType() As FileLoadType
    ImgLoadType = gtFileLoadType
End Property

Property Let ImgLoadType(value As FileLoadType)
    gtFileLoadType = value
End Property

'��ȡ��ǰ����ID
Property Get ProcessId() As Long
    ProcessId = GetCurrentProcessId
End Property

Property Get HasVideo() As Boolean
'�жϸö����Ƿ�ӵ����Ƶ�ɼ�
    HasVideo = mblnHasVideo
End Property

'����״̬
Property Get VideoDockState() As Boolean
    VideoDockState = gblnDockingState
End Property


'����ʹ�òɼ�ģ��Ĵ��ڹؼ���
Property Get ParentWindowKey() As String
    ParentWindowKey = mstrParentWindowKey
End Property


Property Let ParentWindowKey(ByVal value As String)
    mstrParentWindowKey = value
End Property


'��ǰ����״̬
Property Get IsLock() As Boolean
    IsLock = mobjVideoWindow.IsLock
End Property

'�Ƿ��̨�ɼ���
Property Get IsAfter() As Boolean
    IsAfter = mobjVideoWindow.IsAfter
End Property


Private Function CheckVideoObj() As Boolean
    CheckVideoObj = False
    
    If mobjVideoWindow Is Nothing Then
        Set mobjVideoWindow = gobjVideoV2
        
        If mobjVideoWindow Is Nothing Then
            err.Raise 0, "CheckVideoObj", "��Ƶ�ɼ�������ʧЧ��"
            Exit Function
        End If
        
        Call mobjVideoWindow.zlRePreview
    End If
    
    CheckVideoObj = True
End Function

'Private Sub DoDockClose()
'On Error GoTo errHandle
'BUGEX_ "DoDockClose(clsPacsCapture) 1"
'    If Not mobjEvent Is Nothing Then
'BUGEX_ "DoDockClose(clsPacsCapture) 2"
'        Call mobjEvent.OnDockClose
'    ElseIf Not mRegEvent Is Nothing Then
'BUGEX_ "DoDockClose(clsPacsCapture) 3"
'    Else
'BUGEX_ "DoDockClose(clsPacsCapture) 4"
'        RaiseEvent OnDockClose
'    End If
'BUGEX_ "DoDockClose(clsPacsCapture) End"
'Exit Sub
'errHandle:
'    BUGEX_ "DoDockClose(clsPacsCapture) Err:" & err.Description
'End Sub


'******************************************************************************************************************************
'������ýӿ�
'******************************************************************************************************************************


'��ʾ������Ƶ����
Public Sub zlShowPopupVideo(ByVal blnReadOnly As Boolean)
    Dim lngPicHeight As Long
    Dim lngWindowStyle As Long
    
    '��������ɼ����򿪣��������ٴδ򿪲ɼ�����
    If gblnDockingState = dsOpen Then Exit Sub
    
    '�򿪸����ɼ�����ǰ��֪ͨ����ʵ�������Ӧ������״̬
    Call mobjNotifyEvent.NotifyContainerLoad(mobjDock)

    Set mobjDock = New frmVideoDockWindow
    
'    If mobjVideoWindow.LockState = True Then
'        mobjDock.Caption = "��Ƶ�ɼ� [" & mobjVideoWindow.LockPatientName & "]"
'    End If

    '������ڵ���״̬������Ҫ����������ʾ״̬
    If gblnOpenDebug Then
        lngWindowStyle = GetWindowLong(mobjDock.hwnd, -20)
        lngWindowStyle = lngWindowStyle Or WS_EX_APPWINDOW
    
        SetWindowLong mobjDock.hwnd, -20, lngWindowStyle
    End If
    
'    mobjDock.Caption = mobjVideoWindow.CaptionEx
    Call mobjDock.Show(0)
        
    lngPicHeight = Val(LoadVideoAreaCfg(M_STR_PARAMETER_VIDEO_DOCK))
    
    Call UpdateVideoContainer(mobjDock, lngPicHeight, True)
    
'    Call SetForegroundWindow(mobjDock.hwnd)
'    Call mobjVideoWindow.ucSplitter1.RePaint(False)
    
    '�򿪺���Ҫˢ���ػ���Ƶ��ʾ����
    Call mobjDock.Refresh
    
    Call zlRestoreWindow(blnReadOnly, True)
    
    glngInstanceCount = glngInstanceCount + 1
    
    gblnDockingState = dsOpen
    
    mblnHasVideo = True
End Sub

Public Sub zlRestoreWindow(ByVal blnReadOnly As Boolean, Optional ByVal blnIsMain As Boolean = False, _
    Optional ByVal blnIsOnlyState As Boolean = False)
    If CheckVideoObj = False Then Exit Sub
    
    Call mobjVideoWindow.zlRestoreWindow(blnReadOnly, blnIsMain, blnIsOnlyState)
End Sub

Public Sub zlPreviewThumbnail(objImg As Object)
    If CheckVideoObj = False Then Exit Sub
    
    Call mobjVideoWindow.zlPreviewThumbnail(objImg)
End Sub

Public Function IsCurVideo() As Boolean
    Dim lngHwnd As Long
    
    IsCurVideo = False
    
    If mobjVideoWindow Is Nothing Then Exit Function
    
    lngHwnd = GetAncestor(mobjVideoWindow.hwnd, GA_PARENT)
    
    IsCurVideo = IIf(lngHwnd = mobjPane.hwnd, True, False)

End Function

'�����ڽ���ActiveWindow״̬ʱ����Ҫ���ô˷���
Public Sub zlRefreshVideoWindow()
On Error GoTo errHandle
    '��������ɼ����򿪣����л���Ƶ���ڵ���ʾ��ˢ��
BUGEX_ "zlRefreshVideoWindow 1"
    If gblnDockingState = dsOpen Then
    
        'ע��򿪵��Ⱥ�˳�򣬵��رո�������ʱ����������ȷ�Ĵ����м���
        Call RegVideoWindowOrder
    
        Exit Sub
    End If
    
    If mblnRefreshState = True Then Exit Sub
    
    mblnRefreshState = True
    
    '��Ҫ�����øñ���ֵΪfalse�������ظ�ˢ�µ�ǰ����ʱ���ظ�����mobjNotifyEvent_OnContainerLoad�¼�������
    mblnHasVideo = False
    
BUGEX_ "zlRefreshVideoWindow 2"
    '֪ͨ������ʵ��������picCapture�߶ȵȲ���
    Call mobjNotifyEvent.NotifyContainerLoad(mobjPane)
    
BUGEX_ "zlRefreshVideoWindow 3"
    '��ȡmobjVideoWindow.picCapture�߶�����
    If mstrParentWindowKey <> "" And mlngCurPicHeight <= 0 Then
        '��һ�μ��ش���ʱ����Ҫ��ȡ�ò���ֵ
        mlngCurPicHeight = Val(LoadVideoAreaCfg(mstrParentWindowKey))
    End If
        
BUGEX_ "zlRefreshVideoWindow 4"
    Call UpdateVideoContainer(mobjPane, mlngCurPicHeight)
    
    
    '�ڵ���Refresh����ʱ���ᴥ��mobjPane��Paint�¼��������Ҫ�Ƚ�mblnHasVideo����Ϊtrue
    mblnHasVideo = True
    
BUGEX_ "zlRefreshVideoWindow 5"
    '�򿪺���Ҫˢ���ػ���Ƶ��ʾ����
'    Call mobjPane.Refresh
    
BUGEX_ "zlRefreshVideoWindow 6"
    '�����ڲ���ZORDER����
    Call RegVideoWindowOrder
    
    If gobjCapturePar.IsAutoRepreview Then Call zlRePreview
    
    mblnRefreshState = False
    
BUGEX_ "zlRefreshVideoWindow End"

Exit Sub
errHandle:
    mblnRefreshState = False
    err.Raise err.Number, err.Source, err.Description, err.HelpFile, err.HelpContext
End Sub

'��֤�����Ƿ���Ч
Public Function IsValid() As Boolean
'��ActiveExe���̱������˳������з�����������
    IsValid = True
End Function


Public Sub zlInitModule(objDbAccess As Object, _
                        objCapHelper As Object, _
                        ByVal lngSysNo As Long, _
                        ByVal lngCallModule As Long, _
                        ByVal strPrivs As String, _
                        ByVal lngCurDepartId As Long, _
                        ByVal lngOwnerHwnd As Long, _
                        Optional blnOpenDebugMode As Boolean = False)
'��ʼ��ģ�����
    Dim blnIsEqualDB As Boolean


    gblnOpenDebug = IIf(blnOpenDebugMode = False, gblnOpenDebug, blnOpenDebugMode)
    
BUGEX_ "zlInitModule 1", True
    If objDbAccess Is Nothing Then Exit Sub
    
    glngSys = lngSysNo
    glngModule = lngCallModule
    gstrPrivs = strPrivs
    glngDepartId = lngCurDepartId
    Set mobjCapHelper = objCapHelper
        
BUGEX_ "zlInitModule3:Connect DB.", True
    '��ʼ���û��ؼ������ݿ�����
    Call InitCommonLib(objDbAccess)
    
BUGEX_ "zlInitModule 4", True
    glngRootHandle = lngOwnerHwnd
    If lngOwnerHwnd <= 0 Then
        glngRootHandle = GetAncestor(mobjPane.hwnd, GA_ROOT)
    End If
    
BUGEX_ "zlInitModule 5", True
    Call gobjCapturePar.ReadParameter
    
 
BUGEX_ "zlInitModule 6", True
    Call mobjVideoWindow.zlInitModule(objCapHelper)
    
    mblnInitFinishState = True
    
BUGEX_ "zlInitModule End", True
End Sub


Public Sub zlCaptureImg()
'ͼ��ɼ�
    If VideoIsNull = True Then Exit Sub
    
    Call mobjVideoWindow.ForeCapture(True)
End Sub

Public Sub zlRePreview()
'���½���Ԥ��
    Call mobjVideoWindow.zlRePreview
End Sub

Public Sub zlCaptureAfterImg()
'��̨�ɼ�
    If VideoIsNull = True Then Exit Sub
    If Not gobjCapturePar.IsUseAfterCapture Then Exit Sub
    
    Call mobjVideoWindow.AfterCapture
End Sub

Public Sub zlUpdateAfterCaptureInfo()
'���±��
    If VideoIsNull = True Then Exit Sub
    If Not gobjCapturePar.IsUseAfterCapture Then Exit Sub

    Call mobjVideoWindow.ResetAfterCaptureTag
End Sub


Public Sub zlNotifyRefresh()
'֪ͨ��������ˢ��
    If VideoIsNull = True Then Exit Sub
    
    Call mobjVideoWindow.NotificationRefresh
End Sub

Public Sub zlNotifyQuit()
'֪ͨ�˳�Active��Ƶģ��
    gblnIsQuitModule = True
End Sub


Public Sub zlShowVideoConfig()
'��ʾ��Ƶ���ô���

BUGEX_ "zlShowVideoConfig 1"
    If VideoIsNull = True Then Exit Sub
    
'    If gblnDockingState = dsOpen Then Exit Sub
    
BUGEX_ "zlShowVideoConfig 2"
    Call mobjVideoWindow.ShowVideoConfig
BUGEX_ "zlShowVideoConfig End"
End Sub


Public Sub PlayVideo(ByVal strMediaFile As String)
'------------------------------------------------
'���ܣ�dcmView��¼��ͼ��Ĳ���
'��������
'���أ��ޣ�ֱ�Ӳ���dcmView�е�ͼ��
'------------------------------------------------
On Error GoTo errHandle
     If mobjPlay Is Nothing Then
        Set mobjPlay = New frmPlaying
    End If
    
    mobjPlay.Show
 
    Call mobjPlay.OpenVideoFile(Replace(strMediaFile, "/", "\"), Nothing, True)
    
Exit Sub
errHandle:
    MsgboxCus err.Description, vbOKOnly, "��ʾ"
End Sub


'******************************************************************************************************************************
'�ڲ�˽�д�������
'******************************************************************************************************************************

Private Function VideoIsNull() As Boolean
    VideoIsNull = IIf(gobjVideoV2 Is Nothing, True, False)
End Function


Private Sub RegVideoWindowOrder()
'ע����Ƶ���ڴ�˳��
    If Not gobjZOrder Is Nothing Then
        Call DelCurZOrderInf
        Call gobjZOrder.Add(mobjPane.hwnd, mobjPane.hwnd)
        
'        Debug.Print "RegHwnd:" & mobjPane.hwnd
    End If
End Sub

Private Sub CreateGlobalResource()
'��ʼ����Ƶ�ɼ�����
    
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
    
    '����������Դ
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
    If gobjVideoV2 Is Nothing Then
BUGEX_ "CreateGlobalResource gobjVideoNew", True
        Set gobjVideoV2 = New frmVideoCaptureV2
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
    '��ʼ��ģ�����
    mstrParentWindowKey = ""
    mlngCurPicHeight = 0
    mblnHasVideo = False
    mblnInitFinishState = False
    gblnIsQuitModule = False
    mblnAllowEventNotify = False
    
BUGEX_ "Class_Initialize 2"
    
    '����������Դ
    Call CreateGlobalResource
    
BUGEX_ "Class_Initialize 3"
    '��������Դ
    If mobjPane Is Nothing Then
        BUGEX_ "CreateResource mobjPane", True
        Set mobjPane = New frmVideoPaneWindow
    End If
    
BUGEX_ "Class_Initialize 4"
    '��������������
    Set mobjVideoWindow = gobjVideoV2
    Set mobjNotifyEvent = gobjNotifyEvent
    
    glngInstanceCount = glngInstanceCount + 1
    
BUGEX_ "Class_Initialize End", True
Exit Sub
errHandle:
    BUGEX_ "Class_Initialize Err:" & err.Description, True, True
End Sub


Public Sub FreeGlobalResource()
'�ͷ���Ƶ�ɼ���Դ
        
BUGEX_ "FreeGlobalSource 1", True
        '�ͷ��¼�֪ͨ����
        If Not gobjNotifyEvent Is Nothing Then
            Set gobjNotifyEvent = Nothing
        End If
                
BUGEX_ "FreeGlobalSource 2", True
        'ֻ��ʵ������Ϊһ��ʱ�򣬲��ͷ�ȫ�ֹ�����Դ
        If Not gobjVideoV2 Is Nothing Then
            Unload gobjVideoV2
            Set gobjVideoV2 = Nothing
        End If
        
BUGEX_ "FreeGlobalSource 3", True
        If Not gobjZOrder Is Nothing Then
            Call gobjZOrder.RemoveAll
            Set gobjZOrder = Nothing
        End If
        
BUGEX_ "FreeGlobalSource 4", True

        '���ò������ö���
        If Not gobjCapturePar Is Nothing Then
            Set gobjCapturePar = Nothing
        End If
        
BUGEX_ "FreeGlobalSource 5", True
        '�ͷ�oracle����
        If Not gcnVideoOracle Is Nothing Then
            Set gcnVideoOracle = Nothing
        End If
        
BUGEX_ "FreeGlobalSource End", True
End Sub

Private Sub Class_Terminate()
'�ͷ���Դ
    Dim i As Integer
On Error GoTo errHandle

BUGEX_ "FreeSource 1", True
    Call DelCurZOrderInf
    
    '�ͷ�����Դ--------------------------------------
    
    '�ͷ��¼�����
    If Not mRegEvent Is Nothing Then
        Set mRegEvent = Nothing
    End If
    
    
BUGEX_ "FreeSource 2", True
    '�Ͽ���Ƶ�������������ڵ�����
    If glngCurVideoContainerHwnd = mobjPane.hwnd Then
        mobjVideoWindow.Visible = False
        SetParent mobjVideoWindow.hwnd, 0
    End If
    
    '�ͷŵ�������
    If Not mobjDock Is Nothing Then
        Unload mobjDock
        Set mobjDock = Nothing
    End If
    
    '�ͷ���������
    If Not mobjPane Is Nothing Then
        Unload mobjPane
        Set mobjPane = Nothing
    End If
    
    '�ͷ���Ƶ���Ŵ���
    If Not mobjPlay Is Nothing Then
        Unload mobjPlay
        Set mobjPlay = Nothing
    End If
    
    
    '�ͷ�ȫ����Դ--------------------------------------
    If glngInstanceCount = 1 Then Call FreeGlobalResource
    
    
    Set mobjNotifyEvent = Nothing
    Set mobjVideoWindow = Nothing
'    Set mobjEvent = Nothing
    Set mobjCapHelper = Nothing
    
    
BUGEX "zlPacsCapture Class_Terminate"

    glngInstanceCount = glngInstanceCount - 1
    glngModule = 0
errHandle:
End Sub

'���µ�ǰ��Ƶ���ڴ�С
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
    
        mobjVideoWindow.Height = objPane.ScaleHeight
        mobjVideoWindow.Width = objPane.ScaleWidth
    
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
                        mobjVideoWindow.ScaleX(objPane.ScaleWidth, vbTwips, vbPixels), _
                        mobjVideoWindow.ScaleY(objPane.ScaleHeight, vbTwips, vbPixels), _
                        0)
    End If

    
BUGEX_ "UpdateVideoContainer 5"
    '���mobjPane�趨��parent����ɽ�����ʾ
    If objPane.Tag <> "1" Then Call ShowObject(objPane)
    
    '��Ҫ����������mobjVideoWindow.picCapture�ĸ߶ȣ�����ں���������ÿ������mobjVideoWindow.picCapture�߶Ȳ�����Ч
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
'��ʾobject����
On Error GoTo errHandle
    objShow.Visible = True
Exit Sub
errHandle:
    ShowWindow objShow.hwnd, 9
End Sub


Private Sub mobjDock_Load()
    '�ָ�����״̬
BUGEX_ "mobjDock_Load 1"
    SetWindowPos mobjDock.hwnd, -1, mobjDock.CurrentX, mobjDock.CurrentY, mobjDock.ScaleWidth, mobjDock.ScaleHeight, 3 '�������ö�
    
BUGEX_ "mobjDock_Load 2"
    Call RestoreWinState(mobjDock) ', App.ProductName
    
BUGEX_ "mobjDock_Load End"
End Sub

Private Sub mobjDock_Resize()
On Error GoTo errHandle
    '�������ڴ�С�ı�ʱ����Ҫ������Ƶ�ɼ����ڴ�С
    '������ڽ�����С��ʱ����������Ƶ���ֵ���
    If Not mblnInitFinishState Then Exit Sub
    If mobjDock.WindowState = 1 Then Exit Sub
    
    Call UpdateVideoContainer(mobjDock, 0, True)
    
    Exit Sub
errHandle:
    BUGEX_ "mobjDock_Resize Err:" & err.Description
End Sub

Private Sub BUGEX_(ByVal strInf As String, Optional ByVal blnIsForce As Boolean = False, Optional ByVal blnRaiseErr As Boolean = False)
'���������Ϣ
    If gblnOpenDebug Then blnIsForce = True
    
    If mobjPane Is Nothing Then
        BUGEX "HWND:" & 0 & " " & strInf, blnIsForce
    Else
        BUGEX "HWND:" & mobjPane.hwnd & " " & strInf, blnIsForce
    End If
    
    '������������
    If blnRaiseErr Then
        err.Raise err.Number, err.Source, err.Description, err.HelpFile, err.HelpContext
    End If
End Sub

Private Sub mobjNotifyEvent_OnContainerLoad(objOldPane As Object)
On Error GoTo errHandle
BUGEX_ "mobjNotifyEvent_OnContainerLoad 1"
    If VideoIsNull = True Then Exit Sub
    
BUGEX_ "mobjNotifyEvent_OnContainerLoad 2"
    '���������������д���Ƶ�ɼ�ҳ����ʾ֮ǰ����Ҫ�ȱ��浱ǰ�ɼ�ģ�������ͼ�߶�
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
    
    If VideoIsNull = True Then Exit Sub                       '����Ƶ������Чʱ�����л���Ƶ��ʾ
    If gblnDockingState = dsOpen Then Exit Sub          '�����ڸ���������ʾʱ�����л���Ƶ��ʾ
    If objOldPane.hwnd = mobjPane.hwnd Then Exit Sub    '���˳��Ĵ��ھ���뵱ǰ��������ͬʱ�����л���ʾ��Ƶ
    If mobjNotifyEvent.IsChangePane Then Exit Sub       '����Ƶ�Ѿ����л�ʱ������Ҫ���ظ��Ķ���Ƶ�����л�
    
    If gobjZOrder.count >= 1 Then
BUGEX_ "mobjNotifyEvent_OnContainerUnLoad 1"
        '�����Ѿ����ص���Ƶ���ڻ��߶������ͷ�֮ǰ���������Ƴ���gobjZOrder�еľ����Ϣ��
        '���ֻ���ж��Ƴ���������ص���Ƶ�����Ƿ�Ϊ���б������һ�����ھ������
        If mobjPane.hwnd <> gobjZOrder.Keys(gobjZOrder.count - 1) Then Exit Sub
BUGEX_ "mobjNotifyEvent_OnContainerUnLoad 2"
'    Else
'BUGEX_ "mobjNotifyEvent_OnContainerUnLoad 3"
'        '�жϵ�ǰ�ĸ����ھ���Ƿ�Ϊ֮ǰ��Ƶ���ڵ���һ�����ھ����
'        If GetTopHwnd(mobjPane.hWnd) <> glngNextVideoContainerHwnd Then Exit Sub
'
'        lngParentHwnd = GetAncestor(mobjPane.hWnd, GA_PARENT)
'
'        '��mobjpaneû�и�����ʱ����ֻ������clspacscapture�࣬��û��ʹ�ø����containerhwnd���Խ��ÿؼ����ӵ������У�
'        '��˵��رյ������ں󣬲���Ҫ����Ƶ�ڸ��ӵ�mobjpanel��
'        If lngParentHwnd <= 0 Or lngParentHwnd = GetDesktopWindow Then Exit Sub
'BUGEX_ "mobjNotifyEvent_OnContainerUnLoad 4"
    End If
    
    lngParentHwnd = GetAncestor(mobjPane.hwnd, GA_PARENT)
    
    '��mobjpaneû�и�����ʱ����ֻ������clspacscapture�࣬��û��ʹ�ø����containerhwnd���Խ��ÿؼ����ӵ������У�
    '��˵��رյ������ں󣬲���Ҫ����Ƶ�ڸ��ӵ�mobjpanel��
    If lngParentHwnd <= 0 Or lngParentHwnd = GetDesktopWindow Or mobjPane.Visible = False Then
        If gobjZOrder.count >= 1 Then Call gobjZOrder.Remove(gobjZOrder.Keys(gobjZOrder.count - 1))
        Exit Sub
    End If
        
    Call zlRefreshVideoWindow
    mobjNotifyEvent.IsChangePane = True
    
BUGEX_ "mobjNotifyEvent_OnContainerUnLoad End"
End Sub

Public Sub SetPopupTitle(ByVal strAttachTitle As String)
    mobjDock.Caption = "��Ƶ�ɼ�"
    
    If mobjDock Is Nothing Then Exit Sub
    If mobjVideoWindow.IsLock Then Exit Sub
    If mobjVideoWindow.IsAfter Then Exit Sub
    If Len(strAttachTitle) <= 0 Then Exit Sub
    
    mobjDock.Caption = "��Ƶ�ɼ�(" & strAttachTitle & ")"
End Sub


Private Sub mobjDock_Unload(Cancel As Integer)
On Error GoTo errHandle
    Dim lngParentHwnd As Long
    
    gblnDockingState = dsClosing
    
BUGEX_ "mobjDock_Unload 1"

    If VideoIsNull = True Then Exit Sub
    
BUGEX_ "mobjDock_Unload 2"
    Call SaveWinState(mobjDock) ', App.ProductName
    Call SaveVideoAreaCfg(M_STR_PARAMETER_VIDEO_DOCK, mobjVideoWindow.Height)
    
    mobjVideoWindow.Visible = False
    SetParent mobjVideoWindow.hwnd, 0

'    glngNextVideoContainerHwnd = GetTopHwnd(mobjPane.hWnd)
    
    lngParentHwnd = GetAncestor(mobjPane.hwnd, GA_PARENT)
    
BUGEX_ "mobjDock_Unload 3"
    '���رյ�������ʱ����ֻ������clsPacsCapture��Ψһʵ����ͬʱ��û�н�mobjPane�������ͨ��ContainerHwnd���Ը��ӵ����������У�
    '�������������£�����ֱ���ͷ���Ƶ�ɼ���Դ
    If glngInstanceCount <= 1 And (lngParentHwnd <= 0 Or lngParentHwnd = GetDesktopWindow) Then
BUGEX_ "mobjDock_Unload 4"
        Unload mobjVideoWindow
        Set mobjVideoWindow = Nothing
        
    Else
BUGEX_ "mobjDock_Unload 5"
        If gblnIsQuitModule = False Then
            '֪ͨ����ʵ��������Ƶģ����ʾ
            Call mobjNotifyEvent.NotifyContainerUnLoad(mobjDock)
        End If
    End If
    
BUGEX_ "mobjDock_Unload 6"
    Set mobjDock = Nothing
    
    gblnDockingState = dsClosed
    
'    Call DoDockClose
    
    glngInstanceCount = glngInstanceCount - 1
    
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
'    '��������ɼ����򿪣����л���Ƶ���ڵ���ʾ��ˢ��
'    If Not mblnInitFinishState Or lngParentHwnd <= 0 Or lngParentHwnd = GetDesktopWindow Then Exit Sub
'    If gblnDockingState = dsOpen Then Exit Sub      '����������ڴ򿪣���ˢ����Ƶ����
'
'BUGEX_ "mobjPane_Activate 3"
'    If mblnHasVideo = True Then
'BUGEX_ "mobjPane_Activate 4"
'        '���������������Ѿ���������Ƶ����ֻ�����µ�����Ӧ�����С����
'        '�������ڴ�С�ı�ʱ����Ҫ������Ƶ�ɼ����ڴ�С
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


Private Sub mobjPane_Click()
On Error GoTo errHandle
    Dim lngParentHwnd As Long
    
    If VideoIsNull = True Then Exit Sub
    If Not mblnInitFinishState Then Exit Sub
    
    '����������ڴ򿪣��򲻻ָ���ǰ��������Ƶ��ʾ
    If gblnDockingState = dsOpen Then
        Call RegVideoWindowOrder

        MsgboxCus "���������Ѵ򿪣������ڴ���ʾ��Ƶ��", vbOKOnly + vbInformation, "��ʾ"
        Exit Sub
    End If

    lngParentHwnd = GetAncestor(mobjPane.hwnd, GA_ROOT)
    If lngParentHwnd = 0 Or lngParentHwnd = GetDesktopWindow Then Exit Sub
    
    SetForegroundWindow lngParentHwnd
    
Exit Sub
errHandle:
    
End Sub

Private Sub mobjPane_DblClick()
On Error GoTo errHandle
'    Dim lngParentHwnd As Long
'
'    If VideoIsNull = True Then Exit Sub
'
'    lngParentHwnd = GetAncestor(mobjPane.hwnd, GA_PARENT)
'
'    '����������ڴ򿪣��򲻻ָ���ǰ��������Ƶ��ʾ
'    If Not mblnInitFinishState Or lngParentHwnd <= 0 Or lngParentHwnd = GetDesktopWindow Then Exit Sub
'    If gblnDockingState = dsOpen Then
'        Call RegVideoWindowOrder
'
'        MsgboxCus "���������Ѵ򿪣������ڴ���ʾ��Ƶ��", vbOKOnly + vbInformation, G_STR_HINT_TITLE
'        Exit Sub
'    End If
'
'    Call zlRefreshVideoWindow
Exit Sub
errHandle:

End Sub

Private Sub DrawBackground()
'���Ʊ���ͼ��
    Dim i As Integer
    Dim count As Integer
    Dim wordRect As RECT
    Dim lngImgTop As Long
    Dim lngImgLeft As Long
    
    count = 2
    wordRect.Bottom = 45
    wordRect.Right = 200

    If mobjPane.picBackImg.Height * 3 >= mobjPane.Height Then count = 1

    Call mobjPane.Cls
    
    For i = 0 To count
        lngImgLeft = Round(mobjPane.Width / (i + 1)) - mobjPane.picBackImg.Width + 200
        lngImgTop = Round((mobjPane.Height / 3) * (i + 1) - mobjPane.picBackImg.Height)
        
        Call mobjPane.PaintPicture(mobjPane.picBackImg.Picture, _
            lngImgLeft, _
            lngImgTop, _
            mobjPane.picBackImg.Width, mobjPane.picBackImg.Height)

        If mobjVideoWindow Is Nothing Then
            wordRect.Left = mobjPane.ScaleX(Round(mobjPane.Width / (i + 1)) - mobjPane.picBackImg.Width, vbTwips, vbPixels) + 25
        Else
            wordRect.Left = mobjPane.ScaleX(Round(mobjPane.Width / (i + 1)) - mobjPane.picBackImg.Width, vbTwips, vbPixels) + 8 ' + 17
        End If
        
        wordRect.Top = mobjPane.ScaleY(lngImgTop + mobjPane.picBackImg.Width, vbTwips, vbPixels)

        wordRect.Right = wordRect.Left + 150
        wordRect.Bottom = wordRect.Top + 60

        If mobjVideoWindow Is Nothing Then
            Call DrawText(mobjPane.hdc, "��Ƶδ��ע��" & vbCrLf & "�ѽ�����ƵԴ", 27, wordRect, 0)
        Else
            Call DrawText(mobjPane.hdc, "��Ƶ�ѱ���������ռ��", 20, wordRect, 0)
        End If
    Next i
End Sub

Private Sub mobjPane_Paint()
On Error GoTo errHandle

BUGEX_ "mobjPane_Paint 1"
    
    If mobjPane.Visible = False Then Exit Sub
    
    '�����ǰ��mobjPaneû��Ƕ����Ƶ�ɼ������ػ汳��ͼ��
    If mblnHasVideo = False Or glngCurVideoContainerHwnd <> mobjPane.hwnd Then
BUGEX_ "mobjPane_Paint 2"
        Call DrawBackground
    Else
BUGEX_ "mobjPane_Paint 3"
'        Call SetFocusEx(GetTopHwnd(mobjPane.hwnd))
'        Call SetActiveWindow(GetTopHwnd(mobjPane.hwnd))

        '��mobjPane_Raise�в�������Ƶ��ʾ������Ҫ������Ƶ�����С
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
    
    '��������ɼ����򿪣����л���Ƶ���ڵ���ʾ��ˢ��
    If Not mblnInitFinishState Or lngParentHwnd <= 0 Or lngParentHwnd = GetDesktopWindow Then Exit Sub
'    If mobjPane.Visible = False Then Exit Sub       '�����ǰ����δ��ʾ����ˢ����Ƶ����
    If gblnDockingState = dsOpen Then Exit Sub      '����������ڴ򿪣���ˢ����Ƶ����

BUGEX_ "mobjPane_Resize 3"
    If mblnHasVideo = True Then
BUGEX_ "mobjPane_Resize 4"

        '����ǵ��������С����������Ƶҳ�汻���أ��������Сʱ�����Զ���ʾ��Ƶҳ��
        mobjPane.Tag = IIf(mobjPane.Visible = False, "1", "")
        
        '���������������Ѿ���������Ƶ����ֻ�����µ�����Ӧ�����С����
        '�������ڴ�С�ı�ʱ����Ҫ������Ƶ�ɼ����ڴ�С
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
'��gobjZOrder��ɾ����ǰ�������ڵľ����Ϣ
    If Not gobjZOrder Is Nothing Then
        If gobjZOrder.Exists(mobjPane.hwnd) Then
            Call gobjZOrder.Remove(mobjPane.hwnd)
'            Debug.Print "RemoveHwnd:" & mobjPane.hwnd & " ItemCount:" & gobjZOrder.Count
            
        End If
    End If
End Sub



'Public Sub RegEventObj(objEvent As Object)
''ע���¼�����
'    Set mobjEvent = objEvent
'
'    Set mRegEvent = Nothing
'End Sub


Private Sub mobjPane_Unload(Cancel As Integer)
On Error GoTo errHandle
    If VideoIsNull = True Then Exit Sub
    
'    Debug.Print mobjPane.hwnd
    '�����ǰ�����ؼ��а�����Ƶ�ɼ����棬��������ж��ʱ����Ҫ֪ͨ�������������Ƶ,��֮����Ҫ֪ͨ��
    If glngCurVideoContainerHwnd <> mobjPane.hwnd Then Exit Sub
    If mblnHasVideo = False Then Exit Sub
    
    '��ж�ص�ǰ����ʱ�����������Ƶ���ڣ��򱣴浱ǰ���ڵ���Ƶ�߶�
    If mstrParentWindowKey <> "" Then
        Call SaveVideoAreaCfg(mstrParentWindowKey, mobjVideoWindow.Height)
    End If
    
    mobjVideoWindow.Visible = False
    SetParent mobjVideoWindow.hwnd, 0
    
'    Debug.Print mobjPane.hwnd
    '��Ҫ�Ƴ���gobjZOrder�д��ڵľ����Ϣ
    Call DelCurZOrderInf
    
'    glngNextVideoContainerHwnd = GetNextWindowCus(GetTopHwnd(mobjPane.hWnd))

    'ֻ��clsPacsCapture��ʵ��Ψһʱ����ֱ����mobjpanel��unload�¼����ͷ���Ƶ�ɼ���Դ
    If glngInstanceCount <= 1 Then
        Unload mobjVideoWindow
        Set mobjVideoWindow = Nothing
    Else
        If gblnIsQuitModule = False Then
            '֪ͨ����ʵ��������Ƶģ����ʾ
            Call mobjNotifyEvent.NotifyContainerUnLoad(mobjPane)
        End If
    End If
    
Exit Sub
errHandle:
    MsgboxCus err.Description, vbOKOnly, G_STR_HINT_TITLE
End Sub


'����UID
Private Function CreateUID() As String
    Dim objDcmG As New DicomGlobal
     
    objDcmG.RegString("UIDRoot") = "1"
    
    CreateUID = objDcmG.NewUID
End Function


Private Sub mobjPlay_OnCapture(pic As stdole.StdPicture)
    Dim ImgTmpImage As DicomImage
    Dim strError As String
    
On Error GoTo errHandle
    '����dicom��ʽͼ��
    Set ImgTmpImage = New DicomImage
    
    '��Ƶ���Ųɼ�����ֱ��ʹ�ü����壬��Ϊ�ô˹��ܵ�Ŀ����ԱȽ���ȷ��һ����ٳ��ֱ߱༭�����ڿ���Ƶ��Ȼ�󻹽��вɼ�
    If ClipboardToDicomImg(pic, ImgTmpImage, CreateUID, strError) Then
        Call mobjCapHelper.SaveImg(ImgTmpImage, "", True)
    End If
Exit Sub
errHandle:
    MsgboxCus err.Description, vbOKOnly, "��ʾ"
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

Public Sub ReSetFormFontSize(ByVal intsize As Integer)
'�����ֺ�
'    If Not mobjVideoWindow Is Nothing Then
'        Call mobjVideoWindow.setFontSize(intsize)
'    End If
End Sub

Public Sub ResetLockState(ByVal blnIsLock As Boolean)
    If mobjVideoWindow Is Nothing Then Exit Sub
    
    Call mobjVideoWindow.ResetLockState(blnIsLock)
End Sub


Public Sub LockStudy(ByVal intType As Integer, ByVal lngAdviceId As Long, ByVal lngSendNo As Long, _
ByVal lngStudyState As Long, ByVal blnMoved As Boolean)
'���� ���� �������
'intType=1: ����   2���������
'    If Not mobjVideoWindow Is Nothing Then Call mobjVideoWindow.LockStudy(intType, lngAdviceId, lngSendNo, lngStudyState, blnMoved)
End Sub












