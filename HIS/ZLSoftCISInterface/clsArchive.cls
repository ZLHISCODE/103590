VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsArchive"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'�ӿڹ���˵��:
'
'1��RIS�鿴���Ӳ���
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'�ӿڵ���˳��˵��
'1��֧�����ַ�ʽ�鿴���Ӳ�����Ƕ��ʽ���ںͶ�������
'2���������ڵ��÷�������1����ʼ��zlInitCommon ��2�����Ӳ�������zlSoftCISInterface
'3��Ƕ��ʽ���ڵ��÷�������1����ʼ��zlInitCommon��2���������ڣ���ȡ���ھ��zlGetFormHwnd��3��ˢ�´���zlRefresh

Private mfrmDockArchive As Object ' ��һ���ޱ߿��壬�ɰ󶨵����������ϡ��˹�����ʱֻ֧��35��֮��İ汾
Private mfrmArchive As Object '���˲��˵��õ��Ӳ������Ĵ������

Private mobjArchive As Object '�� zlPublicAdvice.clsPublicAdvice
Private mfrmArchiveView As frmArchiveView   '���Ӳ������Ĵ���

'******************************************************************************************************************
'���ܣ��򿪵��Ӳ������Ĵ���
'������
'       lngPatientID - ����ID
'       lngClinicID - ����ID������Ϊ�Һ�ID ���˹Һż�¼.ID��סԺΪ��ҳID
'���أ�True - �ɹ���False - ʧ��
'˵��������ֱ���ֹ��رմ��ڣ�����Ҫ����zlCloseArchiveForm���ֹ��رյĴ��ڻᱻ����
'******************************************************************************************************************
Public Function zlOpenArchiveForm(ByVal lngPatientID As Long, ByVal lngClinicID As Long, Optional ByVal blnLoad As Boolean = False)
    Dim lngLeft As Long
    Dim lngTop As Long
    
    
    
    '��������Ѿ��򿪣���������ʾ�����ٴ���
    If mfrmArchiveView Is Nothing Then
        '��һ�ε��õ��Ӳ�������ѯ����ֵ������������
        Set mfrmArchiveView = New frmArchiveView
        Call mfrmArchiveView.ShowMe(lngPatientID, lngClinicID, Not blnLoad)
        mfrmArchiveView.WindowState = vbNormal
    End If
    
    '���Ӳ�����ѯ�����Ѿ����ڣ�������ʾ��������������Ϣ
    Call mfrmArchiveView.zlRefresh(lngPatientID, lngClinicID)
    
    If blnLoad Then
        zlOpenArchiveForm = True
        Exit Function
    End If
    Call mfrmArchiveView.Show
    mfrmArchiveView.WindowState = 0
    
    '�Ѵ���չʾ����ǰ�ˣ��в��ֻ��������޷���ʾ����ǰ�ˣ������APIǿ����ʾ
    If (Screen.Width > mfrmArchiveView.Width) And (Screen.Height > mfrmArchiveView.Height) Then
        lngLeft = (Screen.Width - mfrmArchiveView.Width) / 2 / Screen.TwipsPerPixelX
        lngTop = (Screen.Height - mfrmArchiveView.Height) / 2 / Screen.TwipsPerPixelY
    End If
    SetWindowPos mfrmArchiveView.hWnd, HWND_TOPMOST, lngLeft, lngTop, 0, 0, &H10 Or &H1
    SetWindowPos mfrmArchiveView.hWnd, HWND_NOTOPMOST, lngLeft, lngTop, 0, 0, &H10 Or &H1
    BringWindowToTop mfrmArchiveView.hWnd
    SetForegroundWindow mfrmArchiveView.hWnd
    
    
    zlOpenArchiveForm = True
    
End Function

'******************************************************************************************************************
'���ܣ��رյ��Ӳ������Ĵ���
'��������
'���أ�True - �ɹ���False - ʧ��
'˵����
'******************************************************************************************************************
Public Function zlCloseArchiveForm()
    
    '������ڲ����ڣ���ֱ�ӷ��سɹ�
    If mfrmArchiveView Is Nothing Then
        zlCloseArchiveForm = True
        Exit Function
    End If
    
    '�رմ���
    mfrmArchiveView.zlCloseMe
    
    Set mobjArchive = Nothing
    Set mfrmArchiveView = Nothing
    
    zlCloseArchiveForm = True
    
End Function

'******************************************************************************************************************
'���ܣ� ��ʼ������
'������ ��
'���أ� True���ɹ���False��ʧ��
'˵���� �����ʼ��������ʼ��ʱ������:zlInitCommon,һ��ֻ����һ�Ρ�
'******************************************************************************************************************
Public Function zlInitCommon() As Boolean
    Dim objEmr As Object
    
    If CreateArchive = False Then Exit Function
    
    Set objEmr = UpdateEmrInterface
    
    If mobjArchive.InitCommon(gcnOracle, glngSys, , , , , , objEmr) = False Then Exit Function
    zlInitCommon = True
    
 
End Function

'******************************************************************************************************************
'���ܣ� ��ȡǶ��ʽ�������Ĵ��ھ��
'������ ��
'���أ� ���ھ����0ʧ��
'˵����
'******************************************************************************************************************
Public Function zlGetFormHwnd() As Long
    On Error GoTo err

    If CreateArchive = False Then Exit Function
    
    zlGetFormHwnd = zlGetForm.hWnd
    
    Exit Function
err:
    MsgBox err.Description
   
End Function

'******************************************************************************************************************
'���ܣ� ˢ��Ƕ��ʽ�������Ĵ���
'������ lngPatientID - ����ID
'       lngClinicID - ����ID������Ϊ�Һ�ID ���˹Һż�¼.ID��סԺΪ��ҳID
'���أ� 0 �ɹ�����0��ʧ��
'˵����
'******************************************************************************************************************
Public Function zlRefresh(ByVal lngPatientID As Long, ByVal lngClinicID As Long) As Long
    
    zlRefresh = 1
    
    
    If CreateArchive = False Then Exit Function
    Call mfrmDockArchive.zlRefresh(lngPatientID, lngClinicID)
    zlRefresh = 0
    
End Function


'******************************************************************************************************************
'���ܣ� �������ڵĵ��Ӳ�������
'������ lngPatientID - ����ID
'       lngClinicID - ����ID������Ϊ�Һ�ID ���˹Һż�¼.ID��סԺΪ��ҳID
'���أ� 0 �ɹ�����0��ʧ��
'˵����
'******************************************************************************************************************
Public Function zlShowArchive(ByVal lngPatientID As Long, ByVal lngClinicID As Long) As Long
'���ܣ����Ӳ�������  frmParent ���Դ� Nothing
    
    Dim frmParent As Object
    Set frmParent = Nothing
    
    zlShowArchive = 1
    

    If CreateArchive = False Then Exit Function
    Call mobjArchive.ShowArchive(frmParent, lngPatientID, lngClinicID)
    '��ȡ��ǰ��ʾ�Ĵ����������ж��
    If mfrmArchive Is Nothing Then
        Set mfrmArchive = mobjArchive.GetArchiveOneFrom
    End If
    zlShowArchive = 0
   
End Function

Private Function CreateArchive() As Boolean
'���ܣ���������
    err = 0: On Error Resume Next
    If Not mobjArchive Is Nothing Then CreateArchive = True: Exit Function
    Set mobjArchive = CreateObject("zlPublicAdvice.clsPublicAdvice")
    If err <> 0 Then
        MsgBox "δ�ҵ�zlPublicAdvice����,�����վ���ǲ����˴˲���!", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    CreateArchive = False
End Function

Private Function zlGetForm() As Object
'���ܣ����Ӳ��������޿���
    If CreateArchive = False Then Exit Function
    If mfrmDockArchive Is Nothing Then
        Set mfrmDockArchive = mobjArchive.GetArchiveFrom
    End If
    Set zlGetForm = mfrmDockArchive
End Function

Private Sub Class_Initialize()
    Call CreateArchive
End Sub

Private Sub Class_Terminate()
    On Error Resume Next
    Unload mfrmDockArchive
    Unload mfrmArchive
    Set mfrmDockArchive = Nothing
    Set mfrmArchive = Nothing
    Set mobjArchive = Nothing
End Sub



'****************�ٴ��鿴����
Public Sub zlViewEPRReport(frmParent As Object, ByVal lng����ID As Long, ByVal CanPrint As Boolean)
'���ܣ��鿴���Ӳ�������  frmParent ���Դ� Nothing
    If CreateArchive = False Then Exit Sub
    Call mobjArchive.ViewEPRReport(frmParent, lng����ID, CanPrint)
End Sub

'****************�ٴ��鿴PACSͼ��
Public Sub zlViewPACSImage(frmParent As Object, ByVal lngҽ��id As Long, ByVal blnMoved As Boolean)
'���ܣ�PACS��Ƭ���� frmParent ���Դ� Nothing
    If CreateArchive = False Then Exit Sub
    Call mobjArchive.ViewPACSImage(frmParent, lngҽ��id, blnMoved)
End Sub