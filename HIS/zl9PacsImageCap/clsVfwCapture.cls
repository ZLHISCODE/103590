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
'˵������װvfw�Ĳɼ���ع���
'�������ڣ�2009-10-29
'�����ˣ�TJH
'
'����������
'
'*********************************************************************************************************
'---------------------------------------------------------------------------------------------------------


Option Explicit

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'������������߳�������
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Const M_LNG_DEFAULT_CAP_WINDOW_WIDTH As Long = 320 '160      'Ĭ����Ƶ�ɼ����ڵĿ��ȣ�ע��Ͳɼ���Ƶ���ȵ�����
Private Const M_LNG_DEFAULT_CAP_WINDOW_HEIGHT As Long = 240 '120     'Ĭ����Ƶ�ɼ����ڵĸ߶�



Private Const M_LNG_DEFAULT_CUSTOM_ERRNUM As Long = -1 'Ĭ���Զ�������


Private mCaptureParameterInf As clsVfwParameterCfg  '���浱ǰ�ɼ�����


Private mlngCurCaptureWindowHandle As Long '�����ɼ����ں�Ĵ��ھ��
Private mlngVideoDisplayWindowHandle As Long '��Ƶ��ʾ���ڵľ��

Private mlngDisplayWindowWidth As Long '��ǰ������Ƶ��ʾ�Ĵ��ڿ���
Private mlngDisplayWindowHeight As Long '��ǰ������Ƶ��ʾ�Ĵ��ڸ߶�

Private mstrParameterCfgFileName As String '���������ļ�����


'����Ƶ��ʾ���ڴ�С������������¼�
Public Event OnVideoWindowChange(ByVal lngWidth As Long, ByVal lngHeight As Long, ByRef blnIsChangeSize As Boolean)




'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'ж�زɼ��豸
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub StopPreview()
  If mlngCurCaptureWindowHandle = 0 Then
    Exit Sub
  End If
  
  Call SendMessageAsLong(mlngCurCaptureWindowHandle, WM_CAP_DRIVER_DISCONNECT, mCaptureParameterInf.CaptureDeviceIndex, 0&) '�Ͽ���ɼ��豸������
  Call DestroyWindow(mlngCurCaptureWindowHandle) '�ͷŲɼ�������Դ
  
  mlngCurCaptureWindowHandle = 0
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'����fpsȡ��MicrosSecֵ
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function GetMicrosSecFromFPS(ByVal lngFps As Long) As Long
  If lngFps = 0 Then Exit Function
  GetMicrosSecFromFPS = 1000000 / lngFps
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'ȡ��VFW����زɼ�����
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetCaptureParameter() As clsVfwParameterCfg
  '//����vfw��Ƶ�ɼ�����
  Set GetCaptureParameter = mCaptureParameterInf
End Function



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'������زɼ�����
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetCaptureParameter(ByRef captureParmeter As clsVfwParameterCfg)

  If captureParmeter Is Nothing Then
    Exit Sub
  End If
        
    
  '��¡�ɼ���������
  Call captureParmeter.CloneTo(mCaptureParameterInf)
  
  'Ӧ�ò�������
  'Call ConfigCaptureParameter(mCaptureParameterInf)
End Sub



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'���òɼ�����Ҫ�Ĳ���
'
'����˵����
'capBit: �ɼ�����λ��
'capWidth: �ɼ����ڿ���
'capHeight: �ɼ����ڸ߶�
'isScale: �Ƿ��Զ���Ӧ���ڴ�С
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ConfigCaptureParameter(ByRef captureParmeter As clsVfwParameterCfg)
    
  If captureParmeter Is Nothing Then
    BUGEX "LSQ test8"
    Exit Sub
  End If
    
  If mlngCurCaptureWindowHandle = 0 Then
    BUGEX "LSQ test9"
    Exit Sub
  End If
            
    
  'ȡ�òɼ�������ò���
  Dim biCaptureInf As BITMAPINFO
    
  Call SendMessage(mlngCurCaptureWindowHandle, WM_CAP_GET_VIDEOFORMAT, Len(biCaptureInf), biCaptureInf)
    
  '���òɼ�����
  If biCaptureInf.bmiHeader.biBitCount <> 0 Then
      With biCaptureInf.bmiHeader
          .biBitCount = captureParmeter.CapBit
          .biWidth = captureParmeter.CapVideoWidth
          .biHeight = captureParmeter.CapVideoHeight
          .biSizeImage = captureParmeter.CapVideoWidth * captureParmeter.CapVideoHeight * CInt(captureParmeter.CapBit / 8)
      End With
      Call SendMessage(mlngCurCaptureWindowHandle, WM_CAP_SET_VIDEOFORMAT, 0, biCaptureInf)
  End If

  'preview ģʽ
  Call SendMessageAsLong(mlngCurCaptureWindowHandle, WM_CAP_SET_PREVIEWRATE, captureParmeter.CapPreviewRate, 0&) '����Ԥ���ٶ�
  Call SendMessageAsLong(mlngCurCaptureWindowHandle, WM_CAP_SET_PREVIEW, 1, 0&) '����Ԥ��
  Call SendMessageAsLong(mlngCurCaptureWindowHandle, WM_CAP_SET_SCALE, -(captureParmeter.IsScale), 0&) '�Զ���Ӧ���ڴ�С
    
  'overlayģʽ(��Ҫ����CAPSTATUS�ṹ��fOverlayWindow�ж��Ƿ�֧�ָ���ģʽ)
  'Call SendMessageAsLong(curCaptureWindowHandle, WM_CAP_SET_OVERLAY, True, 0&)
    
    
  '������Ƶ�ɼ�����
  Dim videoParameter As CAPTUREPARMS
    
  'ȡ����Ƶ�ɼ���ز���
  Call SendMessage(mlngCurCaptureWindowHandle, WM_CAP_GET_CAPSTREAMPTR, Len(videoParameter), videoParameter)
    
  With videoParameter
      .wPercentDropForError = 10  '֡��ʧ�����ٷֱ�
      .fMakeUserHitOKToCapture = False '�Ƿ񵯳��û���ʾ
      .fUsingDOSMemory = True  '�Ƿ�ʹ��dos�ڴ�ģʽ
      .wNumVideoRequested = 32
      .fAbortLeftMouse = -(captureParmeter.AbortLeftMouse) '�������������Ҽ�ֹͣ�ɼ�
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
'��ȡ�ɼ�����
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ReadVfwCaptureParameterFromFile()
  mCaptureParameterInf.ReadParameterCfg (mstrParameterCfgFileName)
End Sub



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'����ɼ�����
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SaveVfwCaptureParameterToFile()
  mCaptureParameterInf.SaveParameterCfg (mstrParameterCfgFileName)
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'ˢ�²���
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub RefreshParameter()
  Call StartPreview(mlngVideoDisplayWindowHandle)
End Sub


Public Sub ShowFullScreen(owner As Object, ByRef sourceWindow As PictureBox, ByVal monitorIndex As Integer)
  '������Ϊ�㣬���˳�ȫ����ʾ
  If mlngCurCaptureWindowHandle = 0 Then
    Exit Sub
  End If
  
  Call StopPreview
  Call frmFullScreen.ShowFullScreen(Me, mCaptureParameterInf, owner, sourceWindow, monitorIndex)
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'���زɼ��豸
'
'����˵����
'deviceIndex   �豸����
'capDisplyWindowHwnd   ��Ƶ��ʾ���ھ��
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub StartPreview(ByVal lngCapDisplyWindowHwnd As Long)
    
  If lngCapDisplyWindowHwnd = 0 Then
    BUGEX "LSQ test4"
    Exit Sub
  End If
    
    
  '�жϲɼ����ھ���Ƿ�Ϊ0�������Ϊ������Ҫ��ж�زɼ��豸
  If mlngCurCaptureWindowHandle <> 0 Then
    'Call StopPreview
    BUGEX "LSQ test5"
    Exit Sub
  End If
  
  mlngVideoDisplayWindowHandle = lngCapDisplyWindowHwnd
  
  '�����ɼ�����
  mlngCurCaptureWindowHandle = capCreateCaptureWindow("VideoCapture", WS_CHILD Or WS_VISIBLE, 0, 0, M_LNG_DEFAULT_CAP_WINDOW_WIDTH, M_LNG_DEFAULT_CAP_WINDOW_HEIGHT, mlngVideoDisplayWindowHandle, 0)
  If mlngCurCaptureWindowHandle = 0 Then
    'Call err.Raise(M_LNG_DEFAULT_CUSTOM_ERRNUM, "capCreateCaptureWindow", "�ɼ����ڴ���ʧ�ܡ�")
    BUGEX "LSQ test6"
    Exit Sub
  End If
    
  '���Ӳɼ��豸
  Dim bIsSendSuccessed As Boolean
  bIsSendSuccessed = SendMessageAsLong(mlngCurCaptureWindowHandle, WM_CAP_DRIVER_CONNECT, mCaptureParameterInf.CaptureDeviceIndex, 0&)
    
  If Not bIsSendSuccessed Then
    bIsSendSuccessed = SendMessageAsLong(mlngCurCaptureWindowHandle, WM_CAP_DRIVER_CONNECT, mCaptureParameterInf.CaptureDeviceIndex, 0&)
    
    If Not bIsSendSuccessed Then
        mlngCurCaptureWindowHandle = 0
    
        'Call err.Raise(M_LNG_DEFAULT_CUSTOM_ERRNUM, "SendMessageAsLong", "�ɼ��豸����ʧ�ܡ�")
        BUGEX "LSQ test7"
        Exit Sub
    End If
  End If
  
 
  '����Ĭ�ϵĲɼ�����
  Call ConfigCaptureParameter(mCaptureParameterInf)
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'���²ɼ����ڴ�С��λ��
'
'����˵����
'displayWindowWidth   ��ǰ�ɼ����ڿ���(������Ϊ��λ)
'displayWindowHeight   ��ǰ�ɼ����ڸ߶�(������Ϊ��λ)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub UpdateCaptureWindowPos(ByVal lngDisplayWindowWidth As Long, ByVal lngDisplayWindowHeight As Long)
  
  If mlngCurCaptureWindowHandle = 0 Then
    Exit Sub
  End If
  
  '//���浱ǰ��Ƶ��ʾ���ڴ�С
  mlngDisplayWindowWidth = lngDisplayWindowWidth
  mlngDisplayWindowHeight = lngDisplayWindowHeight
      
      
  Dim dblWCutRate As Double
  Dim dblHCutRate As Double
  Dim lngCutWidth As Long
  Dim lngCutHeight As Long
  Dim blnIsChangeSize As Boolean
      
      
  Select Case mCaptureParameterInf.VideoShowWay
    Case 0
      'ԭʼ��С��ʾ
      '���òɼ�����λ�ã���ʹ�����
      Call SetWindowPos(mlngCurCaptureWindowHandle, _
                0&, _
                (mlngDisplayWindowWidth - mCaptureParameterInf.CapVideoWidth) / 2, _
                (mlngDisplayWindowHeight - mCaptureParameterInf.CapVideoHeight) / 2, _
                mCaptureParameterInf.CapVideoWidth, _
                mCaptureParameterInf.CapVideoHeight, _
                SWP_NOZORDER Or SWP_NOSENDCHANGING)
    
    Case 1
      '������ʾ
      Call SetWindowPos(mlngCurCaptureWindowHandle, _
                0&, 0, 0, mlngDisplayWindowWidth, mlngDisplayWindowHeight, _
                SWP_NOZORDER Or SWP_NOSENDCHANGING)
              
    Case 2
      '����������
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
      '���ü���Χ��С������ʾ
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
      '����Ӧ��Ƶ��С
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
      'ԭʼ��С��ʾ
      '���òɼ�����λ�ã���ʹ�����
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
'ͼ��ɼ�
'
'����˵����
'captureWindow   ͼ��ɼ����ڵĴ��ھ��
'
'����ֵ˵�������زɼ�����ͼ��ӿڶ���
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function CaptureImage() As IPictureDisp
      
  On Error GoTo errHandle
    '�жϲɼ����ھ���Ƿ���Ч
    If mlngCurCaptureWindowHandle = 0 Then
        Set CaptureImage = Nothing
        Exit Function
    End If
    
    Call SendMessageAsLong(mlngCurCaptureWindowHandle, WM_CAP_GRAB_FRAME_NOSTOP, 0, 0&) 'ץȡһ��ͼ���ڴ�
    Call SendMessageAsLong(mlngCurCaptureWindowHandle, WM_CAP_SET_PREVIEW, 1, 0&) '��������ɼ���Ƶͼ��
    
    Call SendMessageAsLong(mlngCurCaptureWindowHandle, WM_CAP_EDIT_COPY, 0&, 0&) '���ڴ渴��ͼ�񵽼�����
    
    Set CaptureImage = Clipboard.GetData(2) '�Ӽ������з�������
    
    Clipboard.Clear
    
    Exit Function
errHandle:
  Set CaptureImage = Nothing
End Function



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'ͼ��ɼ�
'
'��ͼ��ɼ���������
'
'
'����ֵ˵����
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function CaptureImageToClipBoard() As Boolean

  On Error GoTo errHandle
    '�жϲɼ����ھ���Ƿ���Ч
    If mlngCurCaptureWindowHandle = 0 Then
        CaptureImageToClipBoard = False
        Exit Function
    End If
    
    Call SendMessageAsLong(mlngCurCaptureWindowHandle, WM_CAP_GRAB_FRAME_NOSTOP, 0, 0&) 'ץȡһ��ͼ���ڴ�
    Call SendMessageAsLong(mlngCurCaptureWindowHandle, WM_CAP_SET_PREVIEW, 1, 0&) '��������ɼ���Ƶͼ��
    
    Call SendMessageAsLong(mlngCurCaptureWindowHandle, WM_CAP_EDIT_COPY, 0&, 0&) '���ڴ渴��ͼ�񵽼�����
    
    CaptureImageToClipBoard = True
    
    Exit Function
errHandle:
    CaptureImageToClipBoard = False
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'ͼ��ɼ���ָ���ļ���
'
'����˵����
'sFileName:   ͼ��ɼ���ı����ļ���
'
'����ֵ˵��������Ϊ��˵���ɼ��ɹ�
Public Function CaptureBmpImgToFile(ByVal strFileName As String) As Boolean
  On Error GoTo errHandle
    CaptureBmpImgToFile = False
      
    '�жϲɼ����ھ���Ƿ���Ч
    If mlngCurCaptureWindowHandle = 0 Then
        Exit Function
    End If
    
    Call SendMessageAsLong(mlngCurCaptureWindowHandle, WM_CAP_GRAB_FRAME_NOSTOP, 0, 0&) 'ץȡһ��ͼ���ڴ�
    Call SendMessageAsLong(mlngCurCaptureWindowHandle, WM_CAP_SET_PREVIEW, 1, 0&) '��������ɼ���Ƶͼ��
    
    Call SendMessageAsLong(mlngCurCaptureWindowHandle, WM_CAP_EDIT_COPY, 0&, 0&) '���ڴ渴��ͼ�񵽼�����
    
    '//�Ӽ�������ȡ�òɼ���ͼ��
    Dim iPicture As IPictureDisp
    Set iPicture = Clipboard.GetData(2) '�Ӽ������з�������
  
    '//���ɼ���ͼ�󱣴浽ָ�����ļ���
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
'������ƵԴ�Ի���
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ShowCaptureVideoSourceDialog()

  If mlngCurCaptureWindowHandle = 0 Then
    Exit Sub
  End If
  
  '//��ʾ�Ի���
  Call SendMessageAsLong(mlngCurCaptureWindowHandle, WM_CAP_DLG_VIDEOSOURCE, 0&, 0&)
End Sub


Public Sub ShowCaptureParameterCfgDialog(objOwner As Object, Optional ByVal lngHideItem As Long = 0)
  Call frmVideoParameter.ShowVfwParameter(Me, mlngDisplayWindowWidth, mlngDisplayWindowHeight, objOwner, lngHideItem)
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'����ͼ��ѹ�����öԻ���
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ShowCaptureCompressionDialog()

  If mlngCurCaptureWindowHandle = 0 Then
    Exit Sub
  End If
  
  '//��ʾ�Ի���
  Call SendMessageAsLong(mlngCurCaptureWindowHandle, WM_CAP_DLG_VIDEOCOMPRESSION, 0&, 0&)
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'������Ƶ��ʽ���öԻ���
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ShowCaptureVideoFormatDialog()
  If mlngCurCaptureWindowHandle = 0 Then
    Exit Sub
  End If
  
  '//��ʾ�Ի���
  Call SendMessageAsLong(mlngCurCaptureWindowHandle, WM_CAP_DLG_VIDEOFORMAT, 0&, 0&)
    
  Dim capState As CAPSTATUS
  Dim retVal As Boolean
  
  'ȡ�����ú�Ĳ�����Ƶ��С
  retVal = SendMessage(mlngCurCaptureWindowHandle, WM_CAP_GET_STATUS, Len(capState), capState)
  If retVal Then
    mCaptureParameterInf.CapVideoWidth = capState.uiImageWidth
    mCaptureParameterInf.CapVideoHeight = capState.uiImageHeight
  End If
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'������Ƶ��ʾ���öԻ�������ɼ���֧�֣�
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ShowCaptureVideoDisplayDialog()

  If mlngCurCaptureWindowHandle = 0 Then
    Exit Sub
  End If
  
  '//��ʾ�Ի���
  Call SendMessageAsLong(mlngCurCaptureWindowHandle, WM_CAP_DLG_VIDEODISPLAY, 0&, 0&)
End Sub


'��ʼ��Ƶ�ɼ�
Public Sub StartVideo(ByVal strVideoFile As String)
  
  '���û������ɼ��豸����ֱ���˳�
  If mlngCurCaptureWindowHandle = 0 Then
    Exit Sub
  End If
  
    
  '������Ƶ�ļ�����λ��
  Call SendMessageAsString(mlngCurCaptureWindowHandle, WM_CAP_FILE_SET_CAPTURE_FILE, 0, strVideoFile)
    
  '������Ƶ�ɼ���ʼ��Ϣ
  Call SendMessageAsLong(mlngCurCaptureWindowHandle, WM_CAP_SEQUENCE, 0&, 0&)
    
End Sub


'ֹͣ��Ƶ�ɼ����÷���û�в��ԣ�
Public Sub StopVideo()

  If mlngCurCaptureWindowHandle = 0 Then
    Exit Sub
  End If
    
  '������ϢֹͣVIDEO�Ĳɼ�
  Call SendMessageAsLong(mlngCurCaptureWindowHandle, WM_CAP_ABORT, 0&, 0&)
End Sub


'*********************************************************************************************
'
'�����Զ���
'
'*********************************************************************************************


'�ɼ����ھ����ֻ����
Property Get hwnd() As Long
  hwnd = mlngCurCaptureWindowHandle
End Property


'��ǰ������Ƶ���ȣ�ֻ����
Property Get CaptureVideoWidth() As Long
  CaptureVideoWidth = mCaptureParameterInf.CapVideoWidth
End Property


'��ǰ������Ƶ�߶ȣ�ֻ����
Property Get CaptureVideoHeight() As Long
  CaptureVideoHeight = mCaptureParameterInf.CapVideoHeight
End Property

'���������ļ�����
Property Get ParameterCfgFileName() As String
  ParameterCfgFileName = mstrParameterCfgFileName
End Property

Property Let ParameterCfgFileName(value As String)
  mstrParameterCfgFileName = value
End Property

'��ȡ�ɼ��������ö���
Property Get CaptureParameterInf() As clsVfwParameterCfg
  Set CaptureParameterInf = mCaptureParameterInf
End Property



'*********************************************************************************************
'
'���¼�
'
'*********************************************************************************************


Private Sub Class_Initialize()
  Set mCaptureParameterInf = New clsVfwParameterCfg
  
  '��ȡĬ��ֵ
  Call mCaptureParameterInf.LoadDefaultCfg
End Sub

Private Sub Class_Terminate()
  Set mCaptureParameterInf = Nothing
End Sub



