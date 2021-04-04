VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCapLinker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'实现ICapHelper接口，传递给视频采集对象进行对应处理
Implements ICapHelper


Private mlngMainAdviceId As Long
Private mlngReportAdviceId As Long
Private mlngLockAdviceId As Long

Private mblnMainReadOnly As Boolean
Private mblnReportReadOnly As Boolean
Private mblnLockReadOnly As Boolean

Private mlngDepartId As Long
Private mstrLineDeviceNO As String
Private mstrBackDeviceNO As String
Private mstrPrivs As String
  
Private mobjMainHelper As ucPacsHelper
Private mobjRepPacsHelpers() As ucPacsHelper
Private mobjOwner As Object
Private mObjNotify As IEventNotify

Public Event OnLockChange(ByVal blnIsLock As Boolean)
Public Event OnAfterChange(ByVal blnIsAfter As Boolean)
  

'报告窗口关闭时，会使用该属性将视频采集嵌入到主界面窗口中
Property Get MainHelper() As Object
    Set MainHelper = mobjMainHelper
End Property

Property Set MainHelper(value As Object)
    Set mobjMainHelper = value
End Property


'主界面列表选择的医嘱ID
Property Get MainAdviceId() As Long
    MainAdviceId = mlngMainAdviceId
End Property
Property Let MainAdviceId(ByVal value As Long)
    mlngMainAdviceId = value
    mblnMainReadOnly = IsReadOnly(value)
End Property

Property Get MainAdvReadOnly() As Boolean
    MainAdvReadOnly = mblnMainReadOnly
End Property

'打开的报告窗口对应的医嘱ID
Property Get ReportAdviceId() As Long
    ReportAdviceId = mlngReportAdviceId
End Property
Property Let ReportAdviceId(ByVal value As Long)
    mlngReportAdviceId = value
    mblnReportReadOnly = IsReadOnly(value)
End Property

Property Get ReportAdvReadOnly() As Boolean
    ReportAdvReadOnly = mblnReportReadOnly
End Property

'锁定采集后的医嘱ID
Property Get LockAdviceId() As Long
    LockAdviceId = mlngLockAdviceId
End Property
Property Let LockAdviceId(ByVal value As Long)
    mlngLockAdviceId = value
    mblnLockReadOnly = IsReadOnly(value)
End Property

Property Get LockAdvReadOnly() As Boolean
    LockAdvReadOnly = mblnLockReadOnly
End Property


Public Function IsMoved(ByVal lngAdviceId As Long) As Boolean
    Dim strSQL As String
    Dim rsData As ADODB.Recordset
    
On Error GoTo errhandle
    IsMoved = False
    
    strSQL = "select 1 from H影像检查记录 where 医嘱ID=[1]"
    Set rsData = zlDatabase.OpenSQLRecord(strSQL, "查询医嘱是否转储", lngAdviceId)
    If rsData.RecordCount > 0 Then IsMoved = True
    
Exit Function
errhandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function IsReadOnly(ByVal lngAdviceId As Long)
    Dim strSQL As String
    Dim rsData As ADODB.Recordset
    Dim lngStep As Long
On Error GoTo errhandle
    '默认为只读
    IsReadOnly = True
    
    If lngAdviceId <= 0 Then Exit Function
    
    '已经转储的数据为只读
    If IsMoved(lngAdviceId) Then Exit Function
    
    '查询当前医嘱的执行过程是否为已报到且没有审核
    strSQL = "Select 执行过程 From 病人医嘱发送 Where 医嘱ID=[1]"
    Set rsData = zlDatabase.OpenSQLRecord(strSQL, "查询医嘱过程", lngAdviceId)
    
    If rsData.RecordCount <= 0 Then Exit Function
    
    lngStep = Val(nvl(rsData!执行过程))
    
    '已完成或未报到的检查不允许图像采集
    If lngStep = 6 Or lngStep = 1 Or lngStep = 0 Then Exit Function
    
    IsReadOnly = False
Exit Function
errhandle:
    IsReadOnly = True
    
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Public Sub AddRepPacsHelper(ByVal lngAdviceId As Long, objPacsHelper As Object)
'注册弹出式报告窗口中的pacshelper对象，便于图像采集时的同步显示
    Dim i As Long
    Dim lngBound As Long
    
    For i = 1 To UBound(mobjRepPacsHelpers)
        If mobjRepPacsHelpers(i).AdviceId = lngAdviceId Then
            Set mobjRepPacsHelpers(i) = Nothing
            Set mobjRepPacsHelpers(i) = objPacsHelper
            
            Exit Sub
        End If
    Next
    
    '新增PacsHelper数组
    lngBound = UBound(mobjRepPacsHelpers) + 1
    ReDim Preserve mobjRepPacsHelpers(lngBound)
    
    Set mobjRepPacsHelpers(lngBound) = objPacsHelper
End Sub

Public Sub RemoveRepPacsHelper(ByVal lngAdviceId As Long)
    Dim i As Long
    Dim blnStartMove As Boolean
    Dim lngBound As Long
    
    blnStartMove = False
    lngBound = UBound(mobjRepPacsHelpers)
    
    For i = 1 To lngBound
        If mobjRepPacsHelpers(i).AdviceId = lngAdviceId Then
            blnStartMove = True
        End If
        
        If blnStartMove And i + 1 <= lngBound Then
        '开始移动数组
            Set mobjRepPacsHelpers(i) = Nothing
            Set mobjRepPacsHelpers(i) = mobjRepPacsHelpers(i + 1)
        End If
    Next
    
    If blnStartMove Then
        Set mobjRepPacsHelpers(lngBound) = Nothing
        
        ReDim Preserve mobjRepPacsHelpers(lngBound - 1)
    End If
    
End Sub



Public Sub Init(objNotify As IEventNotify, ByVal lngDeptId As Long, ByVal strPrivs As String)
    mlngDepartId = lngDeptId
    mstrPrivs = strPrivs
    
    Set mobjOwner = objNotify.Owner
    Set mObjNotify = objNotify
    
    '查询非医技工作站中的图像存储设备
    mstrLineDeviceNO = GetDeptPara(mlngDepartId, "存储设备号")
    mstrBackDeviceNO = GetDeptPara(mlngDepartId, "备份设备号")
         
End Sub

Private Function GetCapAdviceId() As Long
'获取采集医嘱ID

    '判断视频采集窗口是否嵌入到了报告书写窗口中，如果是则需要遍历报告窗口对应的医嘱ID,并获取对应的采集医嘱ID
    '如果没有嵌入报告窗口，可能嵌入了主窗口，或者是弹出式采集窗口，这样就直接使用主医嘱ID
    '如果是锁定了采集，则直接取锁定的医嘱ID
    If mlngLockAdviceId <> 0 Then
        GetCapAdviceId = mlngLockAdviceId
        Exit Function
    End If
    
    If mlngReportAdviceId <> 0 Then
        GetCapAdviceId = mlngReportAdviceId
        Exit Function
    End If
    
    If mlngMainAdviceId <> 0 Then
        GetCapAdviceId = mlngMainAdviceId
        Exit Function
    End If
    
End Function


Private Sub Class_Initialize()
    ReDim mobjRepPacsHelpers(0)
End Sub

Private Sub Class_Terminate()
    Dim i As Long
    
    For i = 1 To UBound(mobjRepPacsHelpers)
        Set mobjRepPacsHelpers(i) = Nothing
    Next
    
    Erase mobjRepPacsHelpers
    
    Set mobjMainHelper = Nothing
    
    Set mObjNotify = Nothing
    Set mobjOwner = Nothing
End Sub

'接口实现********************************************************************************************

Public Function ICapHelper_IsAllowCapture() As Boolean
    '需要判断检查状态
    '需要判断采集数量
    ICapHelper_IsAllowCapture = True
    
    If GetCapAdviceId <= 0 Or InStr(gstrPrivs, "视频采集") <= 0 Then
        ICapHelper_IsAllowCapture = False
        Exit Function
    End If
    
'    If mlngLockAdviceId <> 0 And mblnLockReadOnly Then
'        ICapHelper_IsAllowCapture = False
'        Exit Function
'    End If
'
'    If mlngReportAdviceId <> 0 And mblnReportReadOnly Then
'        ICapHelper_IsAllowCapture = False
'        Exit Function
'    End If
'
'    If mlngMainAdviceId <> 0 And mblnMainReadOnly Then
'        ICapHelper_IsAllowCapture = False
'        Exit Function
'    End If

    If mlngLockAdviceId <> 0 Then
        ICapHelper_IsAllowCapture = Not mblnLockReadOnly
        Exit Function
    End If
    
    If mlngReportAdviceId <> 0 Then
        ICapHelper_IsAllowCapture = Not mblnReportReadOnly
        Exit Function
    End If
    
    If mlngMainAdviceId <> 0 Then
        ICapHelper_IsAllowCapture = Not mblnMainReadOnly
        Exit Function
    End If
    
End Function

Public Function ICapHelper_GetCustomMainID() As Long
'获取专用采集所需的医嘱ID
    ICapHelper_GetCustomMainID = GetCapAdviceId
End Function


Private Function GetRepPacsHelper(ByVal lngAdviceId As Long) As Object
    Dim i As Long
    
    For i = 1 To UBound(mobjRepPacsHelpers)
        Set GetRepPacsHelper = mobjRepPacsHelpers(i)
        If GetRepPacsHelper.AdviceId = lngAdviceId Then Exit Function
    Next
    
    Set GetRepPacsHelper = Nothing
End Function


Private Function LoadCaptureImg(img As Object, _
    dcmInfo As TDicomBaseInfo, lineFtpInfo As TFtpDeviceInf, backFtpInfo As TFtpDeviceInf, _
    Optional ByVal blnIsUpload As Boolean = True) As String
'载入采集图像
    Dim objPacsHelper As ucPacsHelper
    Dim objReportBgImgInfo As clsBgImgInfo
    Dim objMainBgImgInfo As clsBgImgInfo
    Dim strDcmFile As String
    
    Dim objDcm As DicomImage
    Dim blnIsTransProcess As Boolean
    
    LoadCaptureImg = ""
    
    Set objReportBgImgInfo = GetBgImgInfo(dcmInfo, lineFtpInfo, backFtpInfo, blnIsUpload)
    Set objMainBgImgInfo = objReportBgImgInfo.CopyNew
    
    If DirExists(objReportBgImgInfo.FilePath) = False Then
        '创建本地路径
        Call MkLocalDir(objReportBgImgInfo.FilePath)
    End If
    
    strDcmFile = objReportBgImgInfo.FilePath & objReportBgImgInfo.Filename
    If FileExists(strDcmFile) = False And Not img Is Nothing Then
        '将img图像存储到本地
        Set objDcm = img
        'Lossless JPEG encoding JPEG无损压缩    'jpeg2000与其他应用兼容性较弱, 默认不进行压缩传输写入文件性能可提高10倍
        objDcm.WriteFile strDcmFile, True, "1.2.840.10008.1.2.1" '"1.2.840.10008.1.2.4.70"
    End If
    
    blnIsTransProcess = False   '是否进行图像传输处理
    
    Set objPacsHelper = GetRepPacsHelper(dcmInfo.lngAdviceId)
    If Not objPacsHelper Is Nothing Then
        Call objPacsHelper.SyncCaptureImage(objReportBgImgInfo)
        blnIsTransProcess = True
    End If
    
'    If objReportBgImgInfo.LoadState <> lsLocal Then MsgBox "Step1:" & objReportBgImgInfo.ErrorInfo
    
    '如果在弹出报告窗口中对命令进行了处理，则后续不需要在进行处理
    If blnIsTransProcess Then
        If objPacsHelper.IsValid Then
            objMainBgImgInfo.LoadState = lsLocal
            objMainBgImgInfo.ImgCommand = icReadly
        End If
    End If
    
    If Not mobjMainHelper Is Nothing Then
'        如果进行了采集锁定，医嘱id会判断为不同，因此以下处理不会被执行，造成锁定采集的图像不会被上传
        If dcmInfo.lngAdviceId = mobjMainHelper.AdviceId Then
            '同步mainhelper的图像显示
            Call mobjMainHelper.SyncCaptureImage(objMainBgImgInfo)
        Else
            If blnIsTransProcess = False Then   '如果没有进行图像传输处理且医嘱id与当前采集的医嘱id不同时，则需要使用控件的代理传输
                Call mobjMainHelper.SyncCaptureImage(objMainBgImgInfo, True)
            End If
        End If
        
    End If
    
    LoadCaptureImg = strDcmFile
End Function

Private Sub SaveAfterImg(img As Object, ByVal strBackTag As String)
'保存后台采集的图像到本地...
    Dim strPath As String
    Dim objDcmImg As DicomImage
    Dim dcmInfo As TDicomBaseInfo
    Dim i As Long
    
    strPath = FormatFilePath(SysRootPath & "\Apply\TmpAfterImage\" & Format(Now, "yyyymmdd") & "\" & strBackTag)
    
    If DirExists(strPath) = False Then MkLocalDir (strPath)
    
    Set objDcmImg = img
    
    WriteDicomPara objDcmImg, dcmInfo, True
    
    objDcmImg.WriteFile strPath & "\" & objDcmImg.InstanceUID, True, "1.2.840.10008.1.2.1"
    
    If FileExists(strPath & "\" & objDcmImg.InstanceUID) = True Then
        '0-文件名,1-句柄,2-是否声音提示,3-是否后台采集,4-说明
        ShellExecute mobjOwner.hwnd, "open", SysRootPath & "zl9PacsImageHint.exe", strPath & "\" & objDcmImg.InstanceUID & "," & mobjOwner.hwnd & ",1,1," & strBackTag, "", 1
    End If
        
    '后台图像同步显示
    If Not mobjMainHelper Is Nothing Then
        Call mobjMainHelper.SyncAfterCapture(objDcmImg, strBackTag)
    End If
    
    '同步弹出报告窗口中的后台图像
    For i = 1 To UBound(mobjRepPacsHelpers)
        Call mobjRepPacsHelpers(i).SyncAfterCapture(objDcmImg, strBackTag)
    Next
End Sub



Public Function ICapHelper_SaveImg(dcmImg As Object, ByVal strDes As String, _
    Optional ByVal blnIsNeedSave As Boolean = True, _
    Optional ByVal strBackTag As String = "") As Boolean
'保存图像
'blnIsSave如果为true，则不需要进行保存操作
'strBackTag如果不为空，表示后台采集的图像
 
    Dim lngAdviceId As Long
    Dim dicomInfo As TDicomBaseInfo
    Dim lineFtpInfo As TFtpDeviceInf
    Dim backFtpInfo As TFtpDeviceInf
    Dim strDcmFile As String
    Dim strErr As String
    
    Dim strCurDeviceNo As String
    
 
    If Len(strBackTag) > 0 Then
        '后台采集
        Call SaveAfterImg(dcmImg, strBackTag)
        Exit Function
    End If
    
    lngAdviceId = GetCapAdviceId
    
    '判断医嘱ID是否有效
    If lngAdviceId <> 0 Then
        If blnIsNeedSave Then
            dicomInfo = GetDicomBaseInfo(lngAdviceId, False)
            
            dicomInfo.strImageDes = strDes
            
            Call WriteDicomPara(dcmImg, dicomInfo)
            
            lineFtpInfo = GetLineFtpInfo(mstrLineDeviceNO, False, dicomInfo, strErr)
            If Len(strErr) > 0 Then
                MsgboxH mobjOwner.hwnd, strErr, vbOKOnly, "提示"
                Exit Function
            End If
            
            backFtpInfo = GetBackFtpInfo(mstrBackDeviceNO, dicomInfo, strErr)
            If Len(strErr) > 0 Then
                MsgboxH mobjOwner.hwnd, strErr, vbOKOnly, "提示"
                Exit Function
            End If
            
            
            Call SaveImageInfo(dicomInfo, lineFtpInfo)
            
            
        Else
            '如果已经保存（比如使用三方自定义采集时），则从Dicom图像中提取Dicom信息...
            dicomInfo = GetDicomBaseInfoEx(lngAdviceId, dcmImg, strCurDeviceNo)
            
            If Len(strCurDeviceNo) <= 0 Then strCurDeviceNo = mstrLineDeviceNO
            
            lineFtpInfo = GetLineFtpInfo(strCurDeviceNo, False, dicomInfo, strErr)
            If Len(strErr) > 0 Then
                MsgboxH mobjOwner.hwnd, strErr, vbOKOnly, "提示"
                Exit Function
            End If
             
        End If
        
        strDcmFile = LoadCaptureImg(dcmImg, dicomInfo, lineFtpInfo, backFtpInfo, blnIsNeedSave)
  
        '声音提示
        If FileExists(strDcmFile) = True Then
            '0-文件名,1-句柄,2-是否声音提示,3-是否后台采集,4-说明
            ShellExecute mobjOwner.hwnd, "open", SysRootPath & "zl9PacsImageHint.exe", strDcmFile & "," & mobjOwner.hwnd & ",1,0,", "", 4
        End If

        
        '发送消息到主窗口刷新数据行s
        If dicomInfo.lngImgNo = 1 Then
            mObjNotify.Broadcast BM_IMAGE_EVENT_FIRST, , lngAdviceId, -1, , -1    '-1表示首次采集
        Else
            mObjNotify.Broadcast BM_IMAGE_EVENT_FIRST, , lngAdviceId, , , -1
        End If
    Else
        '医嘱ID无效的处理
        If MsgboxH(mobjOwner.hwnd, "当前医嘱无效，是否将采集图像另存到后台缓存？", vbYesNo, "提示") = vbNo Then Exit Function
        
        Call SaveAfterImg(dcmImg, "无效医嘱(临时)")
    End If
    
   
    ICapHelper_SaveImg = True

End Function


Public Function ICapHelper_SaveVideo(ByVal strVideoFile As String, ByVal strDes As String, _
    Optional ByVal strEncodeName As String, Optional ByVal lngRecordTimeLen As Long, _
    Optional ByVal blnIsNeedSave As Boolean = True) As Boolean
'保存视频
    Dim lngAdviceId As Long
    Dim dicomInfo As TDicomBaseInfo
    Dim lineFtpInfo As TFtpDeviceInf
    Dim backFtpInfo As TFtpDeviceInf
    Dim strTargetFile As String
    Dim strErr As String
    
    If FileExists(strVideoFile) = False Then
        MsgboxEx mobjOwner, "未找到录制的视频文件，保存失败。", vbOKOnly, "提示"
        ICapHelper_SaveVideo = False
        Exit Function
    End If
    
'暂不支持后台录像
'    If Len(strBackTag) > 0 Then
'
'    End If
    
    lngAdviceId = GetCapAdviceId
    
    '判断医嘱ID是否有效
    If lngAdviceId <> 0 Then
        dicomInfo = GetDicomBaseInfo(lngAdviceId, False)
         
        dicomInfo.strSeriesDes = "视频录像"
        dicomInfo.strMediaEncode = strEncodeName
        dicomInfo.lngMediaTag = VIDEOTAG
        dicomInfo.lngMediaLen = lngRecordTimeLen
            
        If blnIsNeedSave Then

            lineFtpInfo = GetLineFtpInfo(mstrLineDeviceNO, False, dicomInfo, strErr)
            If Len(strErr) > 0 Then
                MsgboxH mobjOwner.hwnd, strErr, vbOKOnly, "提示"
                Exit Function
            End If
            
            backFtpInfo = GetBackFtpInfo(mstrBackDeviceNO, dicomInfo, strErr)
            If Len(strErr) > 0 Then
                MsgboxH mobjOwner.hwnd, strErr, vbOKOnly, "提示"
                Exit Function
            End If
            
            Call SaveImageInfo(dicomInfo, lineFtpInfo)
        End If
        
        '复制视频文件到检查目录
        strTargetFile = GetStudyImgPath(dicomInfo) & dicomInfo.strInstanceUID
         
        MoveFile strVideoFile, strTargetFile
        
        strTargetFile = LoadCaptureImg(Nothing, dicomInfo, lineFtpInfo, backFtpInfo, blnIsNeedSave)
        
        '声音提示
        If FileExists(strTargetFile) = True Then
            '0-文件名,1-句柄,2-是否声音提示,3-是否后台采集,4-说明
            ShellExecute mobjOwner.hwnd, "open", SysRootPath & "zl9PacsImageHint.exe", "AVI," & mobjOwner.hwnd & ",1,0,", "", 4
        End If

        '发送消息到主窗口刷新数据行
        mObjNotify.Broadcast BM_IMAGE_EVENT_FIRST, , lngAdviceId, -2 '视频
    Else
        strTargetFile = App.Path & "\TmpImage\视频_" & Format(Now, "yymmdd_hhmmss") & ".avi"
        
        If DirExists(App.Path & "\TmpImage\") = False Then Call MkLocalDir(App.Path & "\TmpImage\")
        
        '医嘱ID无效的处理
        If MsgboxH(mobjOwner.hwnd, "当前医嘱无效，是否将当前录像另存到" & strTargetFile & " ？", vbYesNo, "提示") = vbNo Then Exit Function
        
        MoveFile strVideoFile, strTargetFile
    End If
    
    ICapHelper_SaveVideo = True
End Function


Public Function ICapHelper_SaveAudio(ByVal strAudioFile As String, ByVal strDes As String, _
    Optional ByVal strEncodeName As String, Optional ByVal lngRecordTimeLen As Long, _
    Optional ByVal blnIsNeedSave As Boolean = True) As Boolean
'保存音频
    Dim lngAdviceId As Long
    Dim dicomInfo As TDicomBaseInfo
    Dim lineFtpInfo As TFtpDeviceInf
    Dim backFtpInfo As TFtpDeviceInf
    Dim strTargetFile As String
    Dim strErr As String
    
    If FileExists(strAudioFile) = False Then
        MsgboxEx mobjOwner, "未找到录制的音频文件，保存失败。", vbOKOnly, "提示"
        ICapHelper_SaveAudio = False
        Exit Function
    End If
    
'暂不支持后台录像
'    If Len(strBackTag) > 0 Then
'
'    End If
    
    lngAdviceId = GetCapAdviceId
    
    '判断医嘱ID是否有效
    If lngAdviceId <> 0 Then
        dicomInfo = GetDicomBaseInfo(lngAdviceId, False)
         
        dicomInfo.strSeriesDes = "音频数据"
        dicomInfo.strMediaEncode = ""
        dicomInfo.lngMediaTag = AUDIOTAG
        dicomInfo.lngMediaLen = lngRecordTimeLen
            
        If blnIsNeedSave Then

            lineFtpInfo = GetLineFtpInfo(mstrLineDeviceNO, False, dicomInfo, strErr)
            If Len(strErr) > 0 Then
                MsgboxH mobjOwner.hwnd, strErr, vbOKOnly, "提示"
                Exit Function
            End If
            
            backFtpInfo = GetBackFtpInfo(mstrBackDeviceNO, dicomInfo, strErr)
            If Len(strErr) > 0 Then
                MsgboxH mobjOwner.hwnd, strErr, vbOKOnly, "提示"
                Exit Function
            End If
            
            Call SaveImageInfo(dicomInfo, lineFtpInfo)
        End If
        
        '复制视频文件到检查目录
        strTargetFile = GetStudyImgPath(dicomInfo) & dicomInfo.strInstanceUID
         
        MoveFile strAudioFile, strTargetFile
        
        strTargetFile = LoadCaptureImg(Nothing, dicomInfo, lineFtpInfo, backFtpInfo, blnIsNeedSave)
        
        '声音提示
        If FileExists(strTargetFile) = True Then
            '0-文件名,1-句柄,2-是否声音提示,3-是否后台采集,4-说明
            ShellExecute mobjOwner.hwnd, "open", SysRootPath & "zl9PacsImageHint.exe", "WAV," & mobjOwner.hwnd & ",1,0,", "", 4
        End If

        '发送消息到主窗口刷新数据行
        mObjNotify.Broadcast BM_IMAGE_EVENT_FIRST, , lngAdviceId, -3 '音频
    Else
        strTargetFile = App.Path & "\TmpImage\音频_" & Format(Now, "yymmdd_hhmmss") & ".avi"
        
        If DirExists(App.Path & "\TmpImage\") = False Then Call MkLocalDir(App.Path & "\TmpImage\")
        
        '医嘱ID无效的处理
        If MsgboxH(mobjOwner.hwnd, "当前医嘱无效，是否将当前音频另存到" & strTargetFile & " ？", vbYesNo, "提示") = vbNo Then Exit Function
        
        MoveFile strAudioFile, strTargetFile
    End If
    
    ICapHelper_SaveAudio = True
End Function

Private Function GetLockPatientInfo(ByVal lngAdviceId As Long) As String
'获取锁定患者信息
    Dim strSQL As String
    Dim rsData As ADODB.Recordset
    
    GetLockPatientInfo = ""
    
    strSQL = "Select 姓名 from 影像检查记录 Where 医嘱ID=[1]"
    Set rsData = zlDatabase.OpenSQLRecord(strSQL, "查询锁定患者名称", lngAdviceId)
    
    If rsData.RecordCount <= 0 Then Exit Function
    
    GetLockPatientInfo = nvl(rsData!姓名)
End Function

Public Sub ICapHelper_CapLock(ByRef strLocker As String)
'采集锁定
    Dim lngLockAdviceId As Long
    
    If mlngMainAdviceId <> 0 Then
        lngLockAdviceId = mlngMainAdviceId
    End If
    
    If mlngReportAdviceId <> 0 Then
        lngLockAdviceId = mlngReportAdviceId
    End If
     
    LockAdviceId = lngLockAdviceId
    
    strLocker = GetLockPatientInfo(lngLockAdviceId)
    
    RaiseEvent OnLockChange(True)
End Sub


Public Sub ICapHelper_CapUnlock()
'采集解锁
    mlngLockAdviceId = 0
    mblnLockReadOnly = True
    
    RaiseEvent OnLockChange(False)
End Sub


'Public Sub ICapHelper_OpenLocker()
''打开锁定患者
'
'End Sub

Public Sub ICapHelper_AfterTag(ByRef strBackTag As String)
'刷新标记
    Dim i As Long
    
    If strBackTag = "CLOSE" Then
        RaiseEvent OnAfterChange(False)
        Exit Sub
    End If
    
    strBackTag = Format(Now, "HHMMSS")
    
    If Not mobjMainHelper Is Nothing Then Call mobjMainHelper.SyncAfterTag(strBackTag)
    
    For i = 1 To UBound(mobjRepPacsHelpers)
        Call mobjRepPacsHelpers(i).SyncAfterTag(strBackTag)
    Next
    
    RaiseEvent OnAfterChange(True)
End Sub



'接口实现********************************************************************************************
