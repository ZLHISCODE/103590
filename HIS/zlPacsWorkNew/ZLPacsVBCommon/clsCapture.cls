VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCapture"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mobjWork_ImageCap As Object

Property Get HasVideo() As Boolean
'判断该对象是否拥有视频采集
    BUGEX "zlPacsVbCommon>>clsCapture.HasVideo: Step 1", True
     If Not mobjWork_ImageCap Is Nothing Then
        BUGEX "zlPacsVbCommon>>clsCapture.HasVideo: Step 1.1", True
        HasVideo = mobjWork_ImageCap.HasVideo
    End If
    
    BUGEX "zlPacsVbCommon>>clsCapture.HasVideo: Step End.", True
End Property

Public Function ContainerHwnd() As Long
    BUGEX "zlPacsVbCommon>>clsCapture.ContainerHwnd: Step 1", True
    If Not mobjWork_ImageCap Is Nothing Then
        BUGEX "zlPacsVbCommon>>clsCapture.ContainerHwnd: Step 1.1", True
        ContainerHwnd = mobjWork_ImageCap.ContainerHwnd
    End If
    
    BUGEX "zlPacsVbCommon>>clsCapture.ContainerHwnd: Step End.", True
End Function

Public Function ContainerObj() As Object
    BUGEX "zlPacsVbCommon>>clsCapture.ContainerObj: Step 1", True
    If Not mobjWork_ImageCap Is Nothing Then
        BUGEX "zlPacsVbCommon>>clsCapture.ContainerObj: Step 1.1", True
        Set ContainerObj = mobjWork_ImageCap.ContainerObj
    End If
    
    BUGEX "zlPacsVbCommon>>clsCapture.ContainerObj: Step End.", True
End Function

Public Function VideoTitle() As String
    BUGEX "zlPacsVbCommon>>clsCapture.VideoTitle: Step 1", True
    If Not mobjWork_ImageCap Is Nothing Then
        BUGEX "zlPacsVbCommon>>clsCapture.VideoTitle: Step 1.1", True
        VideoTitle = mobjWork_ImageCap.VideoTitle
    End If
    
    BUGEX "zlPacsVbCommon>>clsCapture.VideoTitle: Step End.", True
End Function

Public Sub zlUpdateAdviceInf(ByVal lngAdviceID As Long)
    Dim lngSendNO As Long
    Dim lngStudyState As Long
    Dim blnMoved As Boolean
    Dim strSql As String
    Dim rsData As ADODB.Recordset
    
    BUGEX "zlPacsVbCommon>>clsCapture.zlUpdateAdviceInf: Step 1", True
    If Not mobjWork_ImageCap Is Nothing Then
    
        BUGEX "zlPacsVbCommon>>clsCapture.zlUpdateAdviceInf: Step 1.1", True
        strSql = "Select 发送号,执行过程 From 病人医嘱发送 Where 医嘱ID = [1]"
        Set rsData = zlDatabase.OpenSQLRecord(strSql, "", lngAdviceID)
        
        BUGEX "zlPacsVbCommon>>clsCapture.zlUpdateAdviceInf: Step 1.2", True
        If rsData.RecordCount > 0 Then
            lngSendNO = Nvl(rsData!发送号, 0)
            lngStudyState = Nvl(rsData!执行过程, 0)
        End If
        
        BUGEX "zlPacsVbCommon>>clsCapture.zlUpdateAdviceInf: Step 1.3", True
        Call mobjWork_ImageCap.zlUpdateStudyInf(lngAdviceID, lngSendNO, lngStudyState, blnMoved, False)
    End If
    
    BUGEX "zlPacsVbCommon>>clsCapture.zlUpdateAdviceInf: Step End.", True
End Sub

Public Sub zlRefreshData()
    BUGEX "zlPacsVbCommon>>clsCapture.zlRefreshData: Step 1", True
    If Not mobjWork_ImageCap Is Nothing Then
        BUGEX "zlPacsVbCommon>>clsCapture.zlRefreshData: Step 1.1", True
        Call mobjWork_ImageCap.zlRefreshData
    End If
    
    BUGEX "zlPacsVbCommon>>clsCapture.zlRefreshData: Step End.", True
End Sub

Public Sub zlRefreshVideoWindow(Optional blnSingleWindow As Boolean = False)
    BUGEX "zlPacsVbCommon>>clsCapture.zlRefreshVideoWindow: Step 1", True
    If Not mobjWork_ImageCap Is Nothing Then
    
        BUGEX "zlPacsVbCommon>>clsCapture.zlRefreshVideoWindow: Step 1.1", True
        mobjWork_ImageCap.ParentWindowKey = "采集" & IIf(blnSingleWindow = True, "Dock", "")
        
        BUGEX "zlPacsVbCommon>>clsCapture.zlRefreshVideoWindow: Step 1.2", True
        mobjWork_ImageCap.zlRefreshVideoWindow
    End If
    
    BUGEX "zlPacsVbCommon>>clsCapture.zlRefreshVideoWindow: Step End.", True
End Sub

Public Sub InitActiveVideoModuleObj(objDbAccess As Object, _
                                    ByVal strDbUser As String, _
                                    ByVal lngSysNo As Long, _
                                    ByVal lngCallModule As Long, _
                                    ByVal lngCurDepartId As Long, _
                                    Optional blnForceRefresh As Boolean = False, _
                                    Optional blnOpenDebugMode As Boolean = False)
'初始化ActivexExe视频采集模块对象
    Dim strPrivs As String
    
    BUGEX "zlPacsVbCommon>>clsCapture.InitActiveVideoModuleObj: Step 1", True
    If mobjWork_ImageCap Is Nothing Then
    
        BUGEX "zlPacsVbCommon>>clsCapture.InitActiveVideoModuleObj: Step 1.1", True
        Set mobjWork_ImageCap = CreateObject("zl9PacsImageCap.clsPacsCapture")
        
        BUGEX "zlPacsVbCommon>>clsCapture.InitActiveVideoModuleObj: Step 1.2", True
        InitComLib objDbAccess, strDbUser
        
'        BUGEX "zlPacsVbCommon>>clsCapture.InitActiveVideoModuleObj: Step 1.3", True
        
        BUGEX "zlPacsVbCommon>>clsCapture.InitActiveVideoModuleObj: Step 1.4", True
        strPrivs = ";" & gobjComlib.GetPrivFunc(lngSysNo, lngCallModule) & ";"
        
        BUGEX "zlPacsVbCommon>>clsCapture.InitActiveVideoModuleObj: Step 1.5", True
        Call mobjWork_ImageCap.zlInitModule(objDbAccess, lngSysNo, lngCallModule, strPrivs, lngCurDepartId, 0, Nothing, blnForceRefresh, blnOpenDebugMode)
    End If
    
    BUGEX "zlPacsVbCommon>>clsCapture.InitActiveVideoModuleObj: Step End.", True
End Sub

Public Sub ClassTerminate()
    BUGEX "zlPacsVbCommon>>clsCapture.ClassTerminate: Step 1", True
    If Not mobjWork_ImageCap Is Nothing Then
        BUGEX "zlPacsVbCommon>>clsCapture.ClassTerminate: Step 1.1", True
        Set mobjWork_ImageCap = Nothing
    End If
    BUGEX "zlPacsVbCommon>>clsCapture.ClassTerminate: Step End.", True
End Sub
