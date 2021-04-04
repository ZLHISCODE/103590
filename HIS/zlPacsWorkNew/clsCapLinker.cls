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

'ʵ��ICapHelper�ӿڣ����ݸ���Ƶ�ɼ�������ж�Ӧ����
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
  

'���洰�ڹر�ʱ����ʹ�ø����Խ���Ƶ�ɼ�Ƕ�뵽�����洰����
Property Get MainHelper() As Object
    Set MainHelper = mobjMainHelper
End Property

Property Set MainHelper(value As Object)
    Set mobjMainHelper = value
End Property


'�������б�ѡ���ҽ��ID
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

'�򿪵ı��洰�ڶ�Ӧ��ҽ��ID
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

'�����ɼ����ҽ��ID
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
    
    strSQL = "select 1 from HӰ�����¼ where ҽ��ID=[1]"
    Set rsData = zlDatabase.OpenSQLRecord(strSQL, "��ѯҽ���Ƿ�ת��", lngAdviceId)
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
    'Ĭ��Ϊֻ��
    IsReadOnly = True
    
    If lngAdviceId <= 0 Then Exit Function
    
    '�Ѿ�ת��������Ϊֻ��
    If IsMoved(lngAdviceId) Then Exit Function
    
    '��ѯ��ǰҽ����ִ�й����Ƿ�Ϊ�ѱ�����û�����
    strSQL = "Select ִ�й��� From ����ҽ������ Where ҽ��ID=[1]"
    Set rsData = zlDatabase.OpenSQLRecord(strSQL, "��ѯҽ������", lngAdviceId)
    
    If rsData.RecordCount <= 0 Then Exit Function
    
    lngStep = Val(nvl(rsData!ִ�й���))
    
    '����ɻ�δ�����ļ�鲻����ͼ��ɼ�
    If lngStep = 6 Or lngStep = 1 Or lngStep = 0 Then Exit Function
    
    IsReadOnly = False
Exit Function
errhandle:
    IsReadOnly = True
    
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Public Sub AddRepPacsHelper(ByVal lngAdviceId As Long, objPacsHelper As Object)
'ע�ᵯ��ʽ���洰���е�pacshelper���󣬱���ͼ��ɼ�ʱ��ͬ����ʾ
    Dim i As Long
    Dim lngBound As Long
    
    For i = 1 To UBound(mobjRepPacsHelpers)
        If mobjRepPacsHelpers(i).AdviceId = lngAdviceId Then
            Set mobjRepPacsHelpers(i) = Nothing
            Set mobjRepPacsHelpers(i) = objPacsHelper
            
            Exit Sub
        End If
    Next
    
    '����PacsHelper����
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
        '��ʼ�ƶ�����
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
    
    '��ѯ��ҽ������վ�е�ͼ��洢�豸
    mstrLineDeviceNO = GetDeptPara(mlngDepartId, "�洢�豸��")
    mstrBackDeviceNO = GetDeptPara(mlngDepartId, "�����豸��")
         
End Sub

Private Function GetCapAdviceId() As Long
'��ȡ�ɼ�ҽ��ID

    '�ж���Ƶ�ɼ������Ƿ�Ƕ�뵽�˱�����д�����У����������Ҫ�������洰�ڶ�Ӧ��ҽ��ID,����ȡ��Ӧ�Ĳɼ�ҽ��ID
    '���û��Ƕ�뱨�洰�ڣ�����Ƕ���������ڣ������ǵ���ʽ�ɼ����ڣ�������ֱ��ʹ����ҽ��ID
    '����������˲ɼ�����ֱ��ȡ������ҽ��ID
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

'�ӿ�ʵ��********************************************************************************************

Public Function ICapHelper_IsAllowCapture() As Boolean
    '��Ҫ�жϼ��״̬
    '��Ҫ�жϲɼ�����
    ICapHelper_IsAllowCapture = True
    
    If GetCapAdviceId <= 0 Or InStr(gstrPrivs, "��Ƶ�ɼ�") <= 0 Then
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
'��ȡר�òɼ������ҽ��ID
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
'����ɼ�ͼ��
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
        '��������·��
        Call MkLocalDir(objReportBgImgInfo.FilePath)
    End If
    
    strDcmFile = objReportBgImgInfo.FilePath & objReportBgImgInfo.Filename
    If FileExists(strDcmFile) = False And Not img Is Nothing Then
        '��imgͼ��洢������
        Set objDcm = img
        'Lossless JPEG encoding JPEG����ѹ��    'jpeg2000������Ӧ�ü����Խ���, Ĭ�ϲ�����ѹ������д���ļ����ܿ����10��
        objDcm.WriteFile strDcmFile, True, "1.2.840.10008.1.2.1" '"1.2.840.10008.1.2.4.70"
    End If
    
    blnIsTransProcess = False   '�Ƿ����ͼ���䴦��
    
    Set objPacsHelper = GetRepPacsHelper(dcmInfo.lngAdviceId)
    If Not objPacsHelper Is Nothing Then
        Call objPacsHelper.SyncCaptureImage(objReportBgImgInfo)
        blnIsTransProcess = True
    End If
    
'    If objReportBgImgInfo.LoadState <> lsLocal Then MsgBox "Step1:" & objReportBgImgInfo.ErrorInfo
    
    '����ڵ������洰���ж���������˴��������������Ҫ�ڽ��д���
    If blnIsTransProcess Then
        If objPacsHelper.IsValid Then
            objMainBgImgInfo.LoadState = lsLocal
            objMainBgImgInfo.ImgCommand = icReadly
        End If
    End If
    
    If Not mobjMainHelper Is Nothing Then
'        ��������˲ɼ�������ҽ��id���ж�Ϊ��ͬ��������´������ᱻִ�У���������ɼ���ͼ�񲻻ᱻ�ϴ�
        If dcmInfo.lngAdviceId = mobjMainHelper.AdviceId Then
            'ͬ��mainhelper��ͼ����ʾ
            Call mobjMainHelper.SyncCaptureImage(objMainBgImgInfo)
        Else
            If blnIsTransProcess = False Then   '���û�н���ͼ���䴦����ҽ��id�뵱ǰ�ɼ���ҽ��id��ͬʱ������Ҫʹ�ÿؼ��Ĵ�������
                Call mobjMainHelper.SyncCaptureImage(objMainBgImgInfo, True)
            End If
        End If
        
    End If
    
    LoadCaptureImg = strDcmFile
End Function

Private Sub SaveAfterImg(img As Object, ByVal strBackTag As String)
'�����̨�ɼ���ͼ�񵽱���...
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
        '0-�ļ���,1-���,2-�Ƿ�������ʾ,3-�Ƿ��̨�ɼ�,4-˵��
        ShellExecute mobjOwner.hwnd, "open", SysRootPath & "zl9PacsImageHint.exe", strPath & "\" & objDcmImg.InstanceUID & "," & mobjOwner.hwnd & ",1,1," & strBackTag, "", 1
    End If
        
    '��̨ͼ��ͬ����ʾ
    If Not mobjMainHelper Is Nothing Then
        Call mobjMainHelper.SyncAfterCapture(objDcmImg, strBackTag)
    End If
    
    'ͬ���������洰���еĺ�̨ͼ��
    For i = 1 To UBound(mobjRepPacsHelpers)
        Call mobjRepPacsHelpers(i).SyncAfterCapture(objDcmImg, strBackTag)
    Next
End Sub



Public Function ICapHelper_SaveImg(dcmImg As Object, ByVal strDes As String, _
    Optional ByVal blnIsNeedSave As Boolean = True, _
    Optional ByVal strBackTag As String = "") As Boolean
'����ͼ��
'blnIsSave���Ϊtrue������Ҫ���б������
'strBackTag�����Ϊ�գ���ʾ��̨�ɼ���ͼ��
 
    Dim lngAdviceId As Long
    Dim dicomInfo As TDicomBaseInfo
    Dim lineFtpInfo As TFtpDeviceInf
    Dim backFtpInfo As TFtpDeviceInf
    Dim strDcmFile As String
    Dim strErr As String
    
    Dim strCurDeviceNo As String
    
 
    If Len(strBackTag) > 0 Then
        '��̨�ɼ�
        Call SaveAfterImg(dcmImg, strBackTag)
        Exit Function
    End If
    
    lngAdviceId = GetCapAdviceId
    
    '�ж�ҽ��ID�Ƿ���Ч
    If lngAdviceId <> 0 Then
        If blnIsNeedSave Then
            dicomInfo = GetDicomBaseInfo(lngAdviceId, False)
            
            dicomInfo.strImageDes = strDes
            
            Call WriteDicomPara(dcmImg, dicomInfo)
            
            lineFtpInfo = GetLineFtpInfo(mstrLineDeviceNO, False, dicomInfo, strErr)
            If Len(strErr) > 0 Then
                MsgboxH mobjOwner.hwnd, strErr, vbOKOnly, "��ʾ"
                Exit Function
            End If
            
            backFtpInfo = GetBackFtpInfo(mstrBackDeviceNO, dicomInfo, strErr)
            If Len(strErr) > 0 Then
                MsgboxH mobjOwner.hwnd, strErr, vbOKOnly, "��ʾ"
                Exit Function
            End If
            
            
            Call SaveImageInfo(dicomInfo, lineFtpInfo)
            
            
        Else
            '����Ѿ����棨����ʹ�������Զ���ɼ�ʱ�������Dicomͼ������ȡDicom��Ϣ...
            dicomInfo = GetDicomBaseInfoEx(lngAdviceId, dcmImg, strCurDeviceNo)
            
            If Len(strCurDeviceNo) <= 0 Then strCurDeviceNo = mstrLineDeviceNO
            
            lineFtpInfo = GetLineFtpInfo(strCurDeviceNo, False, dicomInfo, strErr)
            If Len(strErr) > 0 Then
                MsgboxH mobjOwner.hwnd, strErr, vbOKOnly, "��ʾ"
                Exit Function
            End If
             
        End If
        
        strDcmFile = LoadCaptureImg(dcmImg, dicomInfo, lineFtpInfo, backFtpInfo, blnIsNeedSave)
  
        '������ʾ
        If FileExists(strDcmFile) = True Then
            '0-�ļ���,1-���,2-�Ƿ�������ʾ,3-�Ƿ��̨�ɼ�,4-˵��
            ShellExecute mobjOwner.hwnd, "open", SysRootPath & "zl9PacsImageHint.exe", strDcmFile & "," & mobjOwner.hwnd & ",1,0,", "", 4
        End If

        
        '������Ϣ��������ˢ��������s
        If dicomInfo.lngImgNo = 1 Then
            mObjNotify.Broadcast BM_IMAGE_EVENT_FIRST, , lngAdviceId, -1, , -1    '-1��ʾ�״βɼ�
        Else
            mObjNotify.Broadcast BM_IMAGE_EVENT_FIRST, , lngAdviceId, , , -1
        End If
    Else
        'ҽ��ID��Ч�Ĵ���
        If MsgboxH(mobjOwner.hwnd, "��ǰҽ����Ч���Ƿ񽫲ɼ�ͼ�����浽��̨���棿", vbYesNo, "��ʾ") = vbNo Then Exit Function
        
        Call SaveAfterImg(dcmImg, "��Чҽ��(��ʱ)")
    End If
    
   
    ICapHelper_SaveImg = True

End Function


Public Function ICapHelper_SaveVideo(ByVal strVideoFile As String, ByVal strDes As String, _
    Optional ByVal strEncodeName As String, Optional ByVal lngRecordTimeLen As Long, _
    Optional ByVal blnIsNeedSave As Boolean = True) As Boolean
'������Ƶ
    Dim lngAdviceId As Long
    Dim dicomInfo As TDicomBaseInfo
    Dim lineFtpInfo As TFtpDeviceInf
    Dim backFtpInfo As TFtpDeviceInf
    Dim strTargetFile As String
    Dim strErr As String
    
    If FileExists(strVideoFile) = False Then
        MsgboxEx mobjOwner, "δ�ҵ�¼�Ƶ���Ƶ�ļ�������ʧ�ܡ�", vbOKOnly, "��ʾ"
        ICapHelper_SaveVideo = False
        Exit Function
    End If
    
'�ݲ�֧�ֺ�̨¼��
'    If Len(strBackTag) > 0 Then
'
'    End If
    
    lngAdviceId = GetCapAdviceId
    
    '�ж�ҽ��ID�Ƿ���Ч
    If lngAdviceId <> 0 Then
        dicomInfo = GetDicomBaseInfo(lngAdviceId, False)
         
        dicomInfo.strSeriesDes = "��Ƶ¼��"
        dicomInfo.strMediaEncode = strEncodeName
        dicomInfo.lngMediaTag = VIDEOTAG
        dicomInfo.lngMediaLen = lngRecordTimeLen
            
        If blnIsNeedSave Then

            lineFtpInfo = GetLineFtpInfo(mstrLineDeviceNO, False, dicomInfo, strErr)
            If Len(strErr) > 0 Then
                MsgboxH mobjOwner.hwnd, strErr, vbOKOnly, "��ʾ"
                Exit Function
            End If
            
            backFtpInfo = GetBackFtpInfo(mstrBackDeviceNO, dicomInfo, strErr)
            If Len(strErr) > 0 Then
                MsgboxH mobjOwner.hwnd, strErr, vbOKOnly, "��ʾ"
                Exit Function
            End If
            
            Call SaveImageInfo(dicomInfo, lineFtpInfo)
        End If
        
        '������Ƶ�ļ������Ŀ¼
        strTargetFile = GetStudyImgPath(dicomInfo) & dicomInfo.strInstanceUID
         
        MoveFile strVideoFile, strTargetFile
        
        strTargetFile = LoadCaptureImg(Nothing, dicomInfo, lineFtpInfo, backFtpInfo, blnIsNeedSave)
        
        '������ʾ
        If FileExists(strTargetFile) = True Then
            '0-�ļ���,1-���,2-�Ƿ�������ʾ,3-�Ƿ��̨�ɼ�,4-˵��
            ShellExecute mobjOwner.hwnd, "open", SysRootPath & "zl9PacsImageHint.exe", "AVI," & mobjOwner.hwnd & ",1,0,", "", 4
        End If

        '������Ϣ��������ˢ��������
        mObjNotify.Broadcast BM_IMAGE_EVENT_FIRST, , lngAdviceId, -2 '��Ƶ
    Else
        strTargetFile = App.Path & "\TmpImage\��Ƶ_" & Format(Now, "yymmdd_hhmmss") & ".avi"
        
        If DirExists(App.Path & "\TmpImage\") = False Then Call MkLocalDir(App.Path & "\TmpImage\")
        
        'ҽ��ID��Ч�Ĵ���
        If MsgboxH(mobjOwner.hwnd, "��ǰҽ����Ч���Ƿ񽫵�ǰ¼�����浽" & strTargetFile & " ��", vbYesNo, "��ʾ") = vbNo Then Exit Function
        
        MoveFile strVideoFile, strTargetFile
    End If
    
    ICapHelper_SaveVideo = True
End Function


Public Function ICapHelper_SaveAudio(ByVal strAudioFile As String, ByVal strDes As String, _
    Optional ByVal strEncodeName As String, Optional ByVal lngRecordTimeLen As Long, _
    Optional ByVal blnIsNeedSave As Boolean = True) As Boolean
'������Ƶ
    Dim lngAdviceId As Long
    Dim dicomInfo As TDicomBaseInfo
    Dim lineFtpInfo As TFtpDeviceInf
    Dim backFtpInfo As TFtpDeviceInf
    Dim strTargetFile As String
    Dim strErr As String
    
    If FileExists(strAudioFile) = False Then
        MsgboxEx mobjOwner, "δ�ҵ�¼�Ƶ���Ƶ�ļ�������ʧ�ܡ�", vbOKOnly, "��ʾ"
        ICapHelper_SaveAudio = False
        Exit Function
    End If
    
'�ݲ�֧�ֺ�̨¼��
'    If Len(strBackTag) > 0 Then
'
'    End If
    
    lngAdviceId = GetCapAdviceId
    
    '�ж�ҽ��ID�Ƿ���Ч
    If lngAdviceId <> 0 Then
        dicomInfo = GetDicomBaseInfo(lngAdviceId, False)
         
        dicomInfo.strSeriesDes = "��Ƶ����"
        dicomInfo.strMediaEncode = ""
        dicomInfo.lngMediaTag = AUDIOTAG
        dicomInfo.lngMediaLen = lngRecordTimeLen
            
        If blnIsNeedSave Then

            lineFtpInfo = GetLineFtpInfo(mstrLineDeviceNO, False, dicomInfo, strErr)
            If Len(strErr) > 0 Then
                MsgboxH mobjOwner.hwnd, strErr, vbOKOnly, "��ʾ"
                Exit Function
            End If
            
            backFtpInfo = GetBackFtpInfo(mstrBackDeviceNO, dicomInfo, strErr)
            If Len(strErr) > 0 Then
                MsgboxH mobjOwner.hwnd, strErr, vbOKOnly, "��ʾ"
                Exit Function
            End If
            
            Call SaveImageInfo(dicomInfo, lineFtpInfo)
        End If
        
        '������Ƶ�ļ������Ŀ¼
        strTargetFile = GetStudyImgPath(dicomInfo) & dicomInfo.strInstanceUID
         
        MoveFile strAudioFile, strTargetFile
        
        strTargetFile = LoadCaptureImg(Nothing, dicomInfo, lineFtpInfo, backFtpInfo, blnIsNeedSave)
        
        '������ʾ
        If FileExists(strTargetFile) = True Then
            '0-�ļ���,1-���,2-�Ƿ�������ʾ,3-�Ƿ��̨�ɼ�,4-˵��
            ShellExecute mobjOwner.hwnd, "open", SysRootPath & "zl9PacsImageHint.exe", "WAV," & mobjOwner.hwnd & ",1,0,", "", 4
        End If

        '������Ϣ��������ˢ��������
        mObjNotify.Broadcast BM_IMAGE_EVENT_FIRST, , lngAdviceId, -3 '��Ƶ
    Else
        strTargetFile = App.Path & "\TmpImage\��Ƶ_" & Format(Now, "yymmdd_hhmmss") & ".avi"
        
        If DirExists(App.Path & "\TmpImage\") = False Then Call MkLocalDir(App.Path & "\TmpImage\")
        
        'ҽ��ID��Ч�Ĵ���
        If MsgboxH(mobjOwner.hwnd, "��ǰҽ����Ч���Ƿ񽫵�ǰ��Ƶ���浽" & strTargetFile & " ��", vbYesNo, "��ʾ") = vbNo Then Exit Function
        
        MoveFile strAudioFile, strTargetFile
    End If
    
    ICapHelper_SaveAudio = True
End Function

Private Function GetLockPatientInfo(ByVal lngAdviceId As Long) As String
'��ȡ����������Ϣ
    Dim strSQL As String
    Dim rsData As ADODB.Recordset
    
    GetLockPatientInfo = ""
    
    strSQL = "Select ���� from Ӱ�����¼ Where ҽ��ID=[1]"
    Set rsData = zlDatabase.OpenSQLRecord(strSQL, "��ѯ������������", lngAdviceId)
    
    If rsData.RecordCount <= 0 Then Exit Function
    
    GetLockPatientInfo = nvl(rsData!����)
End Function

Public Sub ICapHelper_CapLock(ByRef strLocker As String)
'�ɼ�����
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
'�ɼ�����
    mlngLockAdviceId = 0
    mblnLockReadOnly = True
    
    RaiseEvent OnLockChange(False)
End Sub


'Public Sub ICapHelper_OpenLocker()
''����������
'
'End Sub

Public Sub ICapHelper_AfterTag(ByRef strBackTag As String)
'ˢ�±��
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



'�ӿ�ʵ��********************************************************************************************