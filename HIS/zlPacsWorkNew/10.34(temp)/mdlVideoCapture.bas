Attribute VB_Name = "mdlVideoCapture"
Option Explicit
'Public pCurrentfrmCapture As frmVideoCapture    '��¼ӵ����ƵԴ�Ĳɼ�����
'
'Public Sub zlCapUnoladCapWnd(ByRef hCapWnd As Long)
'    '�ͷŲɼ�����
'    If hCapWnd = 0 Then Exit Sub
'    DestroyWindow hCapWnd
'    hCapWnd = 0
'End Sub
'
'Public Sub zlCapDisConnectCapDevice(miDeviceIndex As Integer, ByRef hCapWnd As Long)
'    '�Ͽ����豸������
'    If hCapWnd = 0 Then Exit Sub
'    Call mDisConnectDevice(hCapWnd, miDeviceIndex)
'End Sub
'
'Public Sub zlCapCopyImageToClipboard(ByRef hCapWnd As Long)
'    '��ͼ���Ƶ�������
'    If hCapWnd = 0 Then Exit Sub
'    Call mCopyImageToClipBoard(hCapWnd)
'End Sub
'
'Public Sub zlCapGrabNoStop(ByRef hCapWnd As Long)
'    '�ɼ�ͼ��֮ǰ���Ȱ�һ֡�ɼ����뻺��
'    If hCapWnd = 0 Then Exit Sub
'    Call mGrapNoStopAndPreview(hCapWnd)
'End Sub
'
'Public Sub zlCapSetCaptureFile(mstrAviFileName As String, ByRef hCapWnd As Long)
'    '����¼���ļ�������
'    If hCapWnd = 0 Then Exit Sub
'    capFileSetCaptureFile hCapWnd, mstrAviFileName
'End Sub
'
'Public Sub zlCapCaptureSetSetup(ByRef hCapWnd As Long)
'    '����¼��Ĳ��������ڵĲ����ǣ�¼��ʼ��ֻ�е�����������Ҽ��Ż�ֹͣ
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
'        '���浽����ȥ�ˣ���ȡ����Ĳ�����
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
'    '��ʼ¼��
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
'    '��ѹ�����ô���
'    If hCapWnd = 0 Then Exit Sub
'    capDlgVideoCompression hCapWnd
'End Sub
'
'Public Sub zlCapDlgVideoFormat(ByRef hCapWnd As Long)
'    '�򿪸�ʽ���öԻ���
'    If hCapWnd = 0 Then Exit Sub
'    mViewerFormat hCapWnd
'End Sub
'
'Public Sub zlCapSaveVideoFormat(strRegName As String, ByRef hCapWnd As Long)
'    '�����ʽ���õ�ע�����
'    Dim BITCapTureInfo As BITMAPINFO
'    Dim strRegPath As String
'    Dim blnRet As Boolean
'
'    If hCapWnd = 0 Then Exit Sub
'
'    blnRet = mGetVideoFormat(hCapWnd, BITCapTureInfo)
'
'    If blnRet = True Then
'        strRegPath = "����ģ��\" & App.ProductName & "\" & strRegName
'        If BITCapTureInfo.bmiHeader.biBitCount <> 0 Then
'            SaveSetting "ZLSOFT", strRegPath, "CapBitCount", BITCapTureInfo.bmiHeader.biBitCount
'            SaveSetting "ZLSOFT", strRegPath, "CapBiWidth", BITCapTureInfo.bmiHeader.biWidth
'            SaveSetting "ZLSOFT", strRegPath, "CapBiHeight", BITCapTureInfo.bmiHeader.biHeight
'        End If
'    End If
'End Sub
'
'Public Sub zlCapDlgVideoSource(ByRef hCapWnd As Long)
'    '����ƵԴ���öԻ���
'    If hCapWnd = 0 Then Exit Sub
'
'    mViewerSource hCapWnd
'End Sub
'
'Function zlCapConnCapDevice(ByRef hCapWnd As Long, ParentWindowWnd As Long, CapDeviceIndex As Integer, _
'    ByRef mlngVideoSizeX As Long, ByRef mlngVideoSizeY As Long) As Boolean
''-----------------------------------------------------------------------------------------
''���ܣ����ӵ��豸�����hCapWnd=0����ʹ��ParentWindowWnd�����ɼ�����
''������hCapWnd �ɼ���������ParentWindowWnd �������� ; CapDeviceIndex �豸������
''���أ�True = �ɹ� False = ʧ��
''�ϼ���������̣�
''�¼���������̣�capCreateCaptureWindow;SendMessageAsLong;SendMessageAsAny;SetWindowPos
''���õ��ⲿ������
''�����ˣ�����
''�޸��ˣ��ƽ� 2009-2-6
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
'    strRegPath = "����ģ��\" & App.ProductName & "\frmVideoCapture"
'    intCapBitCount = Val(GetSetting("ZLSOFT", strRegPath, "CapBitCount", 0))
'    intCapBiWidth = Val(GetSetting("ZLSOFT", strRegPath, "CapBiWidth", 0))
'    intCapBiHeight = Val(GetSetting("ZLSOFT", strRegPath, "CapBiHeight", 0))
'
'    Call mConnectCapDevice(hCapWnd, ParentWindowWnd, CapDeviceIndex, intCapBitCount, intCapBiWidth, intCapBiHeight)
'
'    '��ȡ��Ƶ���ڵķֱ���
'    CaptureWinSize = mGetCaptureWindowStatus(hCapWnd)
'    mlngVideoSizeX = CaptureWinSize.uiImageWidth * Screen.TwipsPerPixelX
'    mlngVideoSizeY = CaptureWinSize.uiImageHeight * Screen.TwipsPerPixelY
'
'    zlCapConnCapDevice = True
'End Function
'
