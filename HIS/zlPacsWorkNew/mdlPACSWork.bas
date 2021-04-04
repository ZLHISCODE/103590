Attribute VB_Name = "mdlPACSWork"
Option Explicit



'����վ���������峣��
Public gbytFontSize As Byte '����վ�����С
Public Const C_INT_FONTSISE_SMALL = 9
Public Const C_INT_FONTSISE_MEDIUM = 12
Public Const C_INT_FONTSISE_BIG = 15

'����Զ�ִ�����
Public Const C_STR_INTERFACE_0 = "���Զ�ִ��"
Public Const C_STR_INTERFACE_1 = "�ǼǺ�"
Public Const C_STR_INTERFACE_2 = "������"
Public Const C_STR_INTERFACE_3 = "��ͼ��"
Public Const C_STR_INTERFACE_4 = "���汣���"
Public Const C_STR_INTERFACE_5 = "����ǩ����"
Public Const C_STR_INTERFACE_6 = "������˺�"
Public Const C_STR_INTERFACE_7 = "�����ɺ�"
Public Const C_STR_INTERFACE_11 = "ȡ���Ǽ�ʱ"
Public Const C_STR_INTERFACE_12 = "ȡ������ʱ"
Public Const C_STR_INTERFACE_13 = "ɾ��ͼ��ʱ"
Public Const C_STR_INTERFACE_14 = "ȡ������ʱ"
Public Const C_STR_INTERFACE_15 = "ȡ��ǩ��ʱ"
Public Const C_STR_INTERFACE_16 = "ȡ�����ʱ"
Public Const C_STR_INTERFACE_17 = "ȡ�����ʱ"
Public Const C_STR_INTERFACE_21 = "����л���"
Public Const C_STR_INTERFACE_22 = "���沵�غ�"

Public Enum EPlugInState
    δ���� = 0
    ͨ�� = 1
    δͨ�� = 2
End Enum

Public Enum EInterfaceExeTime
    ���Զ�ִ�� = 0
    �ǼǺ� = 1
    ������ = 2
    ��ͼ�� = 3
    ���汣��� = 4
    ����ǩ���� = 5
    ������˺� = 6
    �����ɺ� = 7
    ȡ���Ǽ�ʱ = 11
    ȡ������ʱ = 12
    ɾ��ͼ��ʱ = 13
    ȡ������ʱ = 14
    ȡ��ǩ��ʱ = 15
    ȡ�����ʱ = 16
    ȡ�����ʱ = 17
    ����л��� = 21
    ���沵�غ� = 22
End Enum

'ģ��ų�������
Public Const G_LNG_XWPACSVIEW_MODULE As Long = 1288     'XWPACS���
Public Const G_LNG_PACSSTATION_MODULE As Long = 1290    'Ӱ��ҽ��ϵͳ���
Public Const G_LNG_VIDEOSTATION_MODULE As Long = 1291   'Ӱ��ɼ�ϵͳ���
Public Const G_LNG_PATHSTATION_MODULE As Long = 1294    'Ӱ����ϵͳ���

Public Const IMGTAG = 0   'ͼ����
Public Const MULFRAMETAG = 1 '����ͼ
Public Const VIDEOTAG = 2 '��Ƶ���
Public Const AUDIOTAG = 3 '��Ƶ���


Public Type TInterface
    intID As Integer
    strVBS As String
    intType As Integer '���� Ԥ��
    intExeTime As Integer 'ִ��ʱ��
    strName As String '�����Ϣ�� [������:������]
End Type

Public Type TFtpDeviceInf
    strDeviceId As String
    strFtpIp As String
    strFTPUser As String
    strFTPPwd As String
    strFtpDir As String
End Type

'�ɼ�ģ�鴥�����¼�����
Public Enum TVideoEventType
    vetDelAllImg = 0        'ɾ������ͼ��
    vetGetImg = 1           '��ȡͼ��

    vetLockStudy = 2        '�������
    vetUnLockStudy = 3      '�������

    vetCaptureFirstImg = 4  '�ɼ���һ��ͼ��
    vetUpdateImg = 5        '����ͼ��
    
    vetAfterUpdateImg = 6   '���º�̨ͼ��
    
    vetImportImage = 7      '����ͼ��
    vetExportImage = 8      '����ͼ��
        
    vetUseAfterImage = 9
    vetNotUseAfterImage = 10
        
    vetImgCaped = 11
    vetImgDeled = 12
    
    vetAddReportImg = 13    '���뱨��ͼ
End Enum

Public Enum ChargeState
    δ�շ�
    ���շ�
    �޷���
    �Ѽ���
    �Ѳ���
    ���˷�
    ������
    �ѵ���
End Enum

'�༭������
Public Enum ReportType
    ���Ӳ����༭�� = 0
    PACS����༭��
    �����ĵ��༭��
End Enum


'ZLHIS_CIS_017(���߼������)
Public Const G_STR_MSG_ZLHIS_CIS_017 As String = "ZLHIS_CIS_017"

'ZLHIS_PACS_024(����ҽ������)
Public Const G_STR_MSG_ZLHIS_CIS_024 As String = "ZLHIS_CIS_024"

'ZLHIS_CIS_005(ҽ������ִ�����)
Public Const G_STR_MSG_ZLHIS_CIS_005 As String = "ZLHIS_CIS_005"

'ZLHIS_PACS_001(��鱨�����)
Public Const G_STR_MSG_ZLHIS_PACS_001 As String = "ZLHIS_PACS_001"
      
'ZLHIS_PACS_002(���״̬ͬ��)
Public Const G_STR_MSG_ZLHIS_PACS_002 As String = "ZLHIS_PACS_002"

'ZLHIS_PACS_003(���״̬����)
Public Const G_STR_MSG_ZLHIS_PACS_003 As String = "ZLHIS_PACS_003"

'ZLHIS_PACS_004(��鱨�泷��)
Public Const G_STR_MSG_ZLHIS_PACS_004 As String = "ZLHIS_PACS_004"

'ZLHIS_PACS_005(���Σ��ֵ֪ͨ)
Public Const G_STR_MSG_ZLHIS_PACS_005 As String = "ZLHIS_PACS_005"

'���ﻼ�߻��۵���
Public Const G_STR_MSG_ZLHIS_CHARGE_003 As String = "ZLHIS_CHARGE_003"

'Σ��ֵ�Ķ�֪ͨ
Public Const G_STR_MSG_ZLHIS_CIS_025 As String = "ZLHIS_CIS_025"
        
Public gobjMsgCenter As clsPacsMsgProcess
Public gobjRegister As Object
Public gstrUserPswd As String
Public gstrUserName As String
Public gstrServerName As String
Public gcnOracle As ADODB.Connection        '�������ݿ����ӣ��ر�ע�⣺��������Ϊ�µ�ʵ��
Public gstrPrivs As String                   '��ǰ�û����еĵ�ǰģ��Ĺ���
Public gstrSysName As String                'ϵͳ����
Public glngModul As Long
Public glngSys As Long
Public gstrIme As String                    '�Ƿ��Զ��������뷨
Public gstrDBUser As String                 '��ǰ���ݿ��û�
Public gstrInputPwd As String
Public gobjEvent As clsEvent

Public gstrUnitName As String '�û���λ����
Public gfrmMain As Object
Public glngMainHwnd As Long
Public gstrSQL As String
Public glngTXTProc As Long
Public gbln�Ӱ�Ӽ� As Boolean
Public grsDuty As ADODB.Recordset '���ҽ��ְ��
Public grsSysPars As ADODB.Recordset

'ϵͳ����
Public gbytDec As Byte '���ý���С����λ��
Public gstrDec As String '��С��λ������ĸ�ʽ����,��"0.0000"

Public gobjKernel As New zlCISKernel.clsCISKernel 'ҽ������
Public gobjRichEPR As New zlRichEPR.cRichEPR
Public gobjEmr As Object    '���Ӳ���

Public gbytCardLen As Byte '���￨�ų���
Public gblnCardHide As Boolean '���￨��������ʾ
Public gstrCardMask As String  '���￨�������ĸǰ׺:AA|BB|CC...
Public gint�Һ����� As Integer '�Һŵ���Ч����

Public glng������֤ As Long '����һ��ͨ���Ѽ���ʣ����ʱ�Ƿ���Ҫ��֤
Public gblnִ��ǰ�Ƚ��� As Boolean '����һ��ͨ,��Ŀִ��ǰ�������շѻ��ȼ������
Public gblnִ�к���� As Boolean    'ִ�к��Զ���˻��۵�
Public gobjESign As Object                  '����ǩ���ӿڲ���

'�б���ɫ����
Public gdblColor�ѵǼ� As Double
Public gdblColor�ѱ��� As Double
Public gdblColor�Ѽ�� As Double
Public gdblColor�ѱ��� As Double
Public gdblColor����� As Double
Public gdblColor����� As Double
Public gdblColor������ As Double
Public gdblColor������ As Double
Public gdblColor����� As Double
Public gdblColor�Ѿܾ� As Double
Public gdblColor�Ѳ��� As Double


Public gConnectedShardDir() As String   '�Ѿ����ӹ��Ĺ���Ŀ¼���豸������

'---------------------------�豸�������ƣ�ע��-------------------------------
Public Const LOGIN_TYPE_��Ƶ�豸 As String = "Ӱ����Ƶ�豸����"
Public Const LOGIN_TYPE_��Ƭ��ӡ�� As String = "Ӱ��Ƭ��ӡ������"
Public Const LOGIN_TYPE_DICOM�豸 As String = "Ӱ��DICOM�豸����"
Public gint��Ƶ�豸���� As Integer
Public gint��Ƭ��ӡ������ As Integer
Public gintDICOM�豸���� As Integer


Public mrsDeptParas As ADODB.Recordset '���Ʋ�������
'-----------------------------------------------------------
Public Type TYPE_USER_INFO
    ID As Long
    ����ID As Long
    ��� As String
    ���� As String
    ���� As String
    �û��� As String
End Type
Public UserInfo As TYPE_USER_INFO


''''''''''''''''''''''''''''''''''''''ͼ��Ԥ����'''''''''''''''''''''''''''''''''''''
'������Ϣ���ݵĽṹ
Public Type TGetImgMsg
    strSubDir As String          'ͼ�����ڵ���Ŀ¼
    strDestMainDir As String            '����ͼ���Ŀ��Ŀ¼������Ŀ¼
    strIP As String                 'ͼ���������IP��ַ
    strFtpDir As String             'FTPĿ¼
    strFTPUser As String            'FTP�û���
    strFTPPswd As String            'FTP����
    strSDDir As String              '����Ŀ¼����
    strSDUser As String             '����Ŀ¼�û���
    strSDPswd As String             '����Ŀ¼����
    blnEnable As Boolean            '����Ϣ����
End Type

'-------------------------һ��ͨ���������--------------------------------------
Public Const IDKind_���� = "����"
Public Const IDKind_ҽ���� = "ҽ����"
Public Const IDKind_���֤�� = "���֤��"
Public Const IDKind_IC���� = "IC����"
Public Const IDKind_����� = "�����"
Public Const IDKind_סԺ�� = "סԺ��"
Public Const IDKind_�Һŵ� = "�Һŵ���"
Public Const IDKind_�շѵ��ݺ� = "�շѵ��ݺ�"
Public Const IDKindItem_�����ID = "�����ID"
Public Const IDKindItem_���ų��� = "���ų���"

'����һ�������㲿�����࣬����������ص�����
Public Type TSquardCard
    intҽ�ƿ����� As Integer
    lng�����ID As Long
    lngȱʡ�����ID As Long
    blnȱʡ�������� As Boolean
    intȱʡ���ų���  As Integer
    bln�������� As Boolean
End Type

'ͼ���ע
Public Const m_LabelTag_Circle = "NumberCircle"
Public Const m_LabelTag_Back = "NumberBak"
Public Const m_LabelTag_Number = "Number"
Public glngColor(10) As Long             '���ͼ��Բ�α��ʹ�õ�9����ɫ

Public Function IsUseClearType() As Boolean
    Dim lngCurType As Long

    Call SystemParametersInfo(SPI_GETFONTSMOOTHINGTYPE, 0, lngCurType, 0)
    IsUseClearType = IIf(lngCurType = FE_FONTSMOOTHINGCLEARTYPE, True, False)
   
End Function


'*********************************************************************************************************************
'
'�˵���ش������
'
'*********************************************************************************************************************


'��ѯ��ݼ�����
Public Sub BindMenuShortcut(ByVal strProjectName As String, ByVal lngModule As Long, objMenu As Object)
    Dim strSql As String
    Dim rsShoftcutCfg As ADODB.Recordset
    Dim objMain As Object

    strSql = "select a.id, nvl(b.���Ƽ�, a.���Ƽ�) as ���Ƽ�, nvl(b.�ַ���, a.�ַ���) as �ַ���, " & _
             "decode(nvl(b.��ݹ���ID,''),'',a.�����,b.�����) as �����, a.�˵�ID " & _
             "from ��ݹ�����Ϣ a, (select ��ݹ���ID, ���Ƽ�, �ַ���, ����� from ��ݹ��ܹ��� where �û�id=[1] )b " & _
             "where a.id=b.��ݹ���ID(+) and a.��Ŀ=[2] and a.ģ���=[3]"

    Set rsShoftcutCfg = zlDatabase.OpenSQLRecord(strSql, "�󶨲˵���ݼ�", UserInfo.ID, UCase(strProjectName), lngModule)
    
    Set objMain = objMenu
    
    Call RecursionBindMenu(objMain, objMenu.ActiveMenuBar, rsShoftcutCfg)
End Sub


'�󶨲˵���ݷ�ʽ(�ݹ���ð󶨿�ݲ˵�)
Private Sub RecursionBindMenu(cbrMain As Object, objMenu As Object, rsShoftcutCfg As ADODB.Recordset)
    Dim i As Long
    
    If objMenu Is Nothing Then Exit Sub
    If objMenu.Controls.Count <= 0 Then Exit Sub
    
    For i = 1 To objMenu.Controls.Count
        Call BindMenuItemShortcut(cbrMain, objMenu.Controls.Item(i), rsShoftcutCfg)

        If objMenu.Controls.Item(i).type = xtpControlPopup Or objMenu.Controls.Item(i).type = xtpControlButtonPopup Then
            If objMenu.Controls.Item(i).CommandBar.Controls.Count > 0 Then
                Call RecursionBindMenu(cbrMain, objMenu.Controls.Item(i).CommandBar, rsShoftcutCfg)
            End If
        End If
    Next i
End Sub

'�󶨵����˵��Ŀ�ݷ�ʽ
Private Sub BindMenuItemShortcut(cbrMain As Object, cbrControl As Object, rsShoftcutCfg As ADODB.Recordset)
    If rsShoftcutCfg Is Nothing Then Exit Sub
    
    Dim lngFuncKey As Long
    Dim lngCharKey As Long
    Dim lngCommandKey As Long
    
    Dim strKeyAlias As String

    rsShoftcutCfg.Filter = "�˵�ID=" & cbrControl.ID
    
    If rsShoftcutCfg.RecordCount > 0 Then
        lngFuncKey = Val(NVL(rsShoftcutCfg!���Ƽ�))
        lngCharKey = Val(NVL(rsShoftcutCfg!�ַ���))
        strKeyAlias = NVL(rsShoftcutCfg!�����)

        'F8�̶�Ϊ��ݼ��ɼ�ʹ��
        If lngFuncKey = vbKeyF8 Or lngCharKey = vbKeyF8 Then Exit Sub
        
        If (lngFuncKey <> 0 Or lngCharKey <> 0) And InStr(strKeyAlias, "MENU") <= 0 Then
            lngCommandKey = 0
 
            If (lngFuncKey And vbCtrlMask) <> 0 Then
                lngCommandKey = lngCommandKey + FCONTROL
            End If
    
            If (lngFuncKey And vbShiftMask) <> 0 Then
                lngCommandKey = lngCommandKey + FSHIFT
            End If
    
            If (lngFuncKey And vbAltMask) <> 0 Then
                lngCommandKey = lngCommandKey + FALT
            End If
            
            '�󶨲˵���ݼ�
            Call cbrMain.KeyBindings.Add(lngCommandKey, lngCharKey, cbrControl.ID)
            
        ElseIf InStr(strKeyAlias, "MENU") > 0 Then
            If InStr(cbrControl.Caption, "(&") <= 0 Then
                cbrControl.Caption = cbrControl.Caption & "(&" & Replace(strKeyAlias, "MENU+", "") & ")"
            End If
        End If
    End If
    
End Sub



Public Sub CreateViewAndHelpMenu(objViewMenu As Object, objHelpMenu As Object, _
    Optional ByVal strMenuTag As String = "")
    
    Dim cbrControl As CommandBarControl
    Dim cbrMenuBar As CommandBarPopup
    Dim cbrPopControl As CommandBarControl
    
    
    'Begin----------------------�鿴�˵�--------------------------------------
    If Not (objViewMenu Is Nothing) Then
        Set cbrMenuBar = objViewMenu
        
        With cbrMenuBar.CommandBar
            Set cbrControl = .Controls.Add(xtpControlButtonPopup, conMenu_View_ToolBar, "������(T)")
                cbrControl.Category = strMenuTag
                cbrControl.IconId = 1
            
                With cbrControl.CommandBar '�����˵�
                    Set cbrPopControl = .Controls.Add(xtpControlButton, conMenu_View_ToolBar_Button, "��׼��ť(0)")
                        cbrPopControl.Checked = True
                        cbrPopControl.Category = strMenuTag
                        cbrPopControl.IconId = 1
                        
                    Set cbrPopControl = .Controls.Add(xtpControlButton, conMenu_View_ToolBar_Text, "�ı���ǩ(1)")
                        cbrPopControl.Checked = True
                        cbrPopControl.Category = strMenuTag
                        cbrPopControl.IconId = 1
                End With
            Set cbrControl = .Controls.Add(xtpControlButton, conMenu_View_StatusBar, "״̬��(S)")
                cbrControl.Checked = True
                        cbrPopControl.Category = strMenuTag
                        cbrPopControl.IconId = 1
        End With
    End If

    'Begin----------------------�����˵�--------------------------------------Ĭ�Ͽɼ�
    If Not (objHelpMenu Is Nothing) Then
        Set cbrMenuBar = objHelpMenu
        
        With cbrMenuBar.CommandBar
            Set cbrControl = .Controls.Add(xtpControlButton, conMenu_Help_Help, "��������(M)")
                cbrControl.Category = strMenuTag
                cbrControl.IconId = 901
                
            Set cbrControl = .Controls.Add(xtpControlButtonPopup, conMenu_Help_Web, "WEB�ϵ�����(W)")
                cbrControl.Category = strMenuTag
                cbrControl.IconId = 1
                
                With cbrControl.CommandBar
                    Set cbrPopControl = .Controls.Add(xtpControlButton, conMenu_Help_Web_Forum, "������̳(0)")
                        cbrPopControl.Category = strMenuTag
                        cbrPopControl.IconId = 1
                        
                    Set cbrPopControl = .Controls.Add(xtpControlButton, conMenu_Help_Web_Home, "������ҳ(1)")
                        cbrPopControl.Category = strMenuTag
                        cbrPopControl.IconId = 1
                        
                    Set cbrPopControl = .Controls.Add(xtpControlButton, conMenu_Help_Web_Mail, "���ͷ���(2)")
                        cbrPopControl.Category = strMenuTag
                        cbrPopControl.IconId = 9022
                End With
                
            Set cbrControl = .Controls.Add(xtpControlButton, conMenu_Help_About, "���ڡ�(A)")
                cbrControl.Category = strMenuTag
                cbrControl.IconId = 1
        End With
    End If
End Sub

'*********************************************************************************************************************


Public Sub SendMsgToMainWindow(objParameter As Object, _
    ByVal lngWorkType As TWorkEventType, ByVal lngAdviceID As Long, Optional other As Variant = "")
'������Ϣ�������ڵ�Ԫ
    If gobjEvent Is Nothing Then Exit Sub
    
    Call gobjEvent.DoWork(objParameter, lngWorkType, lngAdviceID, other)
End Sub


Public Sub ReadStudyListColor(ByVal lngDeptID As Long)

    gdblColor������ = GetStudyListColor(lngDeptID, "������")
    If gdblColor������ < 0 Then
        gdblColor������ = ColorConstants.vbWhite
    End If
    gdblColor������ = GetColor(gdblColor������)
    
    gdblColor������ = GetStudyListColor(lngDeptID, "������")
    If gdblColor������ < 0 Then
        gdblColor������ = ColorConstants.vbWhite
    End If
    gdblColor������ = GetColor(gdblColor������)
    
    gdblColor����� = GetStudyListColor(lngDeptID, "�����")
    If gdblColor����� < 0 Then
        gdblColor����� = ColorConstants.vbWhite
    End If
    gdblColor����� = GetColor(gdblColor�����)
    
    gdblColor�ѱ��� = GetStudyListColor(lngDeptID, "�ѱ���")
    If gdblColor�ѱ��� < 0 Then
        gdblColor�ѱ��� = ColorConstants.vbWhite
    End If
    gdblColor�ѱ��� = GetColor(gdblColor�ѱ���)
    
    gdblColor�ѵǼ� = GetStudyListColor(lngDeptID, "�ѵǼ�")
    If gdblColor�ѵǼ� < 0 Then
        gdblColor�ѵǼ� = ColorConstants.vbWhite
    End If
    gdblColor�ѵǼ� = GetColor(gdblColor�ѵǼ�)
    
    gdblColor�Ѽ�� = GetStudyListColor(lngDeptID, "�Ѽ��")
    If gdblColor�Ѽ�� < 0 Then
        gdblColor�Ѽ�� = ColorConstants.vbWhite
    End If
    gdblColor�Ѽ�� = GetColor(gdblColor�Ѽ��)
    
    gdblColor����� = GetStudyListColor(lngDeptID, "�����")
    If gdblColor����� < 0 Then
        gdblColor����� = ColorConstants.vbWhite
    End If
    gdblColor����� = GetColor(gdblColor�����)
    
    gdblColor����� = GetStudyListColor(lngDeptID, "�����")
    If gdblColor����� < 0 Then
        gdblColor����� = ColorConstants.vbGreen
    End If
    gdblColor����� = GetColor(gdblColor�����)
    
    gdblColor�ѱ��� = GetStudyListColor(lngDeptID, "�ѱ���")
    If gdblColor�ѱ��� < 0 Then
        gdblColor�ѱ��� = ColorConstants.vbWhite
    End If
    gdblColor�ѱ��� = GetColor(gdblColor�ѱ���)
    
    gdblColor�Ѿܾ� = GetStudyListColor(lngDeptID, "�Ѿܾ�")
    If gdblColor�Ѿܾ� < 0 Then
        gdblColor�Ѿܾ� = ColorConstants.vbRed
    End If
    gdblColor�Ѿܾ� = GetColor(gdblColor�Ѿܾ�)
    
    gdblColor�Ѳ��� = GetStudyListColor(lngDeptID, "�Ѳ���")
    If gdblColor�Ѳ��� < 0 Then
        gdblColor�Ѳ��� = ColorConstants.vbYellow
    End If
    gdblColor�Ѳ��� = GetColor(gdblColor�Ѳ���)
End Sub

Private Function GetColor(ByVal lngColor As Long) As Long
    Dim lngR As Long, lngG As Long, lngB As Long
    Dim lngMaxVal As Long
    
    GetColor = 0
    lngMaxVal = 225
    
    lngR = lngColor Mod 256
    lngG = (Fix(lngColor \ 256)) Mod 256
    lngB = Fix(lngColor \ 256 \ 256)
    
    If lngR = 255 And lngG = 255 And lngB = 255 Then Exit Function
    
    If lngR > lngMaxVal Then lngR = lngMaxVal
    If lngG > lngMaxVal Then lngG = lngMaxVal
    If lngB > lngMaxVal Then lngB = lngMaxVal
    
    GetColor = RGB(lngR, lngG, lngB)
End Function

Public Function GetUserInfo() As Boolean
'���ܣ���ȡ��½�û���Ϣ
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String
    
    Set rsTmp = zlDatabase.GetUserInfo
    
    UserInfo.�û��� = gstrDBUser
    UserInfo.���� = gstrDBUser
    If Not rsTmp.EOF Then
        UserInfo.ID = rsTmp!ID
        UserInfo.��� = rsTmp!���
        UserInfo.����ID = IIf(IsNull(rsTmp!����ID), 0, rsTmp!����ID)
        UserInfo.���� = IIf(IsNull(rsTmp!����), "", rsTmp!����)
        UserInfo.���� = IIf(IsNull(rsTmp!����), "", rsTmp!����)
        UserInfo.�û��� = IIf(IsNull(rsTmp!�û���), "", rsTmp!�û���)
        GetUserInfo = True
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetUser����IDs(Optional ByVal bln���� As Boolean) As String
'���ܣ���ȡ����Ա�����Ŀ���(�������ڿ���+�������������Ŀ���),�����ж��
'�������Ƿ�ȡ���������µĿ���
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String, i As Long
    
    strSql = "Select ����ID From ������Ա Where ��ԱID=[1]"
    If bln���� Then
        strSql = strSql & " Union" & _
            " Select Distinct B.����ID From ������Ա A,��λ״����¼ B" & _
            " Where A.����ID=B.����ID And A.��ԱID=[1]"
    End If
    
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "mdlCISWork", UserInfo.ID)
    For i = 1 To rsTmp.RecordCount
        GetUser����IDs = GetUser����IDs & "," & rsTmp!����ID
        rsTmp.MoveNext
    Next
    GetUser����IDs = Mid(GetUser����IDs, 2)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

'ȡ�ü���б�ָ����������ɫ
Public Function GetStudyListColor(ByVal lngDeptID As Long, ByVal strParameterName As String) As Double
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    Dim lngTemp As Long
             
    On Error GoTo err
        
    strSql = "select ID ,����ID,������,����ֵ from Ӱ�����̲��� where ����ID = [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "ȡ�ü���б���ɫ", lngDeptID)
        
    GetStudyListColor = -1
    
    While Not rsTemp.EOF
        If rsTemp!������ = strParameterName Then
          GetStudyListColor = Val(rsTemp!����ֵ)
          Exit Function
        End If
        rsTemp.MoveNext
    Wend
    
    Exit Function
err:
    If ErrCenter() = 1 Then Resume Next
    Call SaveErrLog
End Function

Public Function getID_TO_����(ByVal lngID As Long, ByVal strDict As String) As String
Dim rsTemp As ADODB.Recordset
    On Error GoTo errHandle
    gstrSQL = "select ���� FROM " & strDict & " WHERE ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ͨ��������ȡID", lngID)
    If Not rsTemp.EOF Then
        getID_TO_���� = NVL(rsTemp!����)
    End If
    Exit Function
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function getID_TO_����(ByVal lngID As Long, ByVal strDict As String) As String
Dim rsTemp As ADODB.Recordset
    On Error GoTo errHandle
    gstrSQL = "select ���� FROM " & strDict & " WHERE ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ͨ��ID��ȡ����", lngID)
    If Not rsTemp.EOF Then
        getID_TO_���� = NVL(rsTemp!����)
    End If
    Exit Function
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function getID_TO_����(ByVal lngID As Long, ByVal strDict As String) As String
Dim rsTemp As ADODB.Recordset
    On Error GoTo errHandle
    gstrSQL = "select ���� FROM " & strDict & " WHERE ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ͨ��ID��ȡ����", lngID)
    If Not rsTemp.EOF Then
        getID_TO_���� = NVL(rsTemp!����)
    End If
    Exit Function
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Sub RemoveCheckImages(ByVal lngҽ��ID As Long, ByVal lng���ͺ� As Long)
    'ɾ��ָ��ҽ���ļ��Ӱ��
    Dim strSql As String, rsTmp As New ADODB.Recordset
    
    Dim Inte As New clsFtp
    Dim strDeviceNO As String
    On Error GoTo ProcError
    '��ɾ��ͼ��
    strSql = "select a.IP��ַ, a.FTPĿ¼, a.FTP�û���, a.FTP����, a.ҽ��ID, a.���ͺ�, a.���UID, a.λ��, a.�������� ,a.�豸�� ,c.ͼ��UID" & _
             " from (select IP��ַ, FTPĿ¼, FTP�û���, FTP����, ҽ��ID, ���ͺ�, ���UID, λ��һ as λ��, ��������, a.�豸�� " & _
             "       from Ӱ���豸Ŀ¼ a, Ӱ�����¼ b " & _
             "       Where a.�豸�� = B.λ��һ " & _
             "       Union All " & _
             "       select IP��ַ, FTPĿ¼, FTP�û���, FTP����, ҽ��ID, ���ͺ�, ���UID, λ�ö� as λ��, ��������, a.�豸��" & _
             "       from Ӱ���豸Ŀ¼ a, Ӱ�����¼ b " & _
             "       Where a.�豸�� = B.λ�ö� " & _
             "       Union All " & _
             "       select IP��ַ, FTPĿ¼, FTP�û���, FTP����, ҽ��ID, ���ͺ�, ���UID, λ���� as λ��, ��������, a.�豸�� " & _
             "       from Ӱ���豸Ŀ¼ a, Ӱ�����¼ b " & _
             "       Where a.�豸�� = B.λ���� " & _
             "       ) a , Ӱ�������� b , Ӱ����ͼ�� c " & _
             " Where a.���uid = B.���uid " & _
             " and b.����uid = c.����uid " & _
             " and a.ҽ��ID = [1] And ���ͺ� = [2] "
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "��ͼ", lngҽ��ID, lng���ͺ�)
    Do Until rsTmp.EOF
        If strDeviceNO <> NVL(rsTmp("�豸��")) Then
            strDeviceNO = NVL(rsTmp("�豸��"))
            Inte.FuncFtpConnect NVL(rsTmp("IP��ַ")), NVL(rsTmp("FTP�û���")), NVL(rsTmp("FTP����"))
        End If
        Inte.FuncDelFile IIf(IsNull(rsTmp("FTPĿ¼")), "", rsTmp("FTPĿ¼") & "/") & Format(rsTmp("��������"), "YYYYMMDD") & "/" & rsTmp("���UID"), rsTmp("ͼ��UID")
        rsTmp.MoveNext
    Loop
    strDeviceNO = ""
    Inte.FuncFtpDisConnect
    'ɾ��Ŀ¼
    strSql = "select IP��ַ,FTPĿ¼,FTP�û���,FTP����,ҽ��ID,���ͺ�,���UID,�豸��,λ��,�������� from " & _
             "      (select IP��ַ,FTPĿ¼,FTP�û���,FTP����,ҽ��ID,���ͺ�,���UID,a.�豸��,λ��һ as λ��,�������� from Ӱ���豸Ŀ¼ a , Ӱ�����¼ b " & _
             "      Where a.�豸�� = B.λ��һ " & _
             "      Union All " & _
             "      select IP��ַ,FTPĿ¼,FTP�û���,FTP����,ҽ��ID,���ͺ�,���UID,a.�豸��,λ�ö� as λ��,�������� from Ӱ���豸Ŀ¼ a , Ӱ�����¼ b " & _
             "      Where a.�豸�� = B.λ�ö� " & _
             "      Union All " & _
             "      select IP��ַ,FTPĿ¼,FTP�û���,FTP����,ҽ��ID,���ͺ�,���UID,a.�豸��,λ���� as λ��,�������� from Ӱ���豸Ŀ¼ a , Ӱ�����¼ b " & _
             "      where a.�豸�� = b.λ���� ) a " & _
             " Where a.ҽ��ID = [1] And ���ͺ� = [2] "
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "��Ŀ¼", lngҽ��ID, lng���ͺ�)
    Do Until rsTmp.EOF
        If strDeviceNO <> NVL(rsTmp("�豸��")) Then
            strDeviceNO = NVL(rsTmp("�豸��"))
            Inte.FuncFtpConnect NVL(rsTmp("IP��ַ")), NVL(rsTmp("FTP�û���")), NVL(rsTmp("FTP����"))
        End If
        Inte.FuncFtpDelDir IIf(IsNull(rsTmp("FTPĿ¼")), "", rsTmp("FTPĿ¼")), Format(rsTmp("��������"), "YYYYMMDD") & "/" & rsTmp("���UID")
        rsTmp.MoveNext
    Loop
    Inte.FuncFtpDisConnect
    Exit Sub
ProcError:
    Inte.FuncFtpDisConnect
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Function MovedByDate(ByVal vDate As Date) As Boolean
'���ܣ��ж�ָ������֮ǰ���Ƿ�����Ѿ�ִ��������ת��,����ָ�������ڡ���š�ϵͳ�ж�ָ�����ڵ������Ƿ���ת���������ݱ���
'������vDate=ʱ����ʱ��εĿ�ʼʱ��

    MovedByDate = zlDatabase.DateMoved(CStr(vDate), 1, glngSys)
    
End Function

Public Function CheckOneDuty(ByVal strҽ�� As String, ByVal strְ�� As String, ByVal strҽ�� As String, ByVal blnҽ�� As Boolean) As String
'���ܣ���鵱ǰָ��ҩƷ����ְ���Ƿ����
'������strҽ��=ҩƷҽ����ʾ����
'      strְ��=ҩƷ����ְ��
'      strҽ��=����ҽ��
'      blnҽ��=�Ƿ񹫷ѻ�ҽ������
'      grsDuty=��¼ҽ��ְ�񻺴�
'���أ�ְ���������ʾ��Ϣ����������򷵻ؿա�
    Const STR_ְ�� = "����,����,�м�,����/ʦ��,Ա/ʿ,,,,��Ƹ"
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String, strMsg As String
    Dim intְ��A As Integer, intְ��B As Integer
    
    If Len(strְ��) <> 2 Or strҽ�� = "" Then Exit Function
    
    'ȡҩƷ����ְ��
    If blnҽ�� Then
        intְ��B = Val(Right(strְ��, 1))
    Else
        intְ��B = Val(Left(strְ��, 1))
    End If
    If intְ��B = 0 Then Exit Function '������
    
    'ȡҽ��ְ��
    If grsDuty Is Nothing Then
        Set grsDuty = New ADODB.Recordset
        grsDuty.Fields.Append "ҽ��", adVarChar, 50
        grsDuty.Fields.Append "ְ��", adInteger
        grsDuty.CursorLocation = adUseClient
        grsDuty.LockType = adLockOptimistic
        grsDuty.CursorType = adOpenStatic
        grsDuty.Open
        Set grsDuty.ActiveConnection = Nothing
    End If
    grsDuty.Filter = "ҽ��='" & strҽ�� & "'"
    If grsDuty.EOF Then
        On Error GoTo errH
        strSql = "Select ����,Nvl(Ƹ�μ���ְ��,0) as ְ�� From ��Ա�� Where ����=[1]"
        'Set rsTmp = New ADODB.Recordset
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "mdlCISWork", strҽ��)
        On Error GoTo 0
        If Not rsTmp.EOF Then
            grsDuty.AddNew
            grsDuty!ҽ�� = rsTmp!����
            grsDuty!ְ�� = rsTmp!ְ��
            grsDuty.Update
        End If
    End If
    If Not grsDuty.EOF Then
        intְ��A = grsDuty!ְ��
    End If
        
    '���ְ��Ҫ��
    If intְ��A = 0 Then
        'ҽ��δ����ְ������
        strMsg = """" & strҽ�� & """Ҫ��Ĵ���ְ�����㣺" & vbCrLf & vbCrLf & IIf(blnҽ��, "��ҽ���򹫷Ѳ���,", "") & _
            "��ҩƷҪ��ְ������Ϊ""" & Split(STR_ְ��, ",")(intְ��B - 1) & """�����´�,��ҽ��""" & strҽ�� & """δ����ְ��"
    ElseIf intְ��B < intְ��A Then
        '��ֵԽСְ��Խ��
        strMsg = """" & strҽ�� & """Ҫ��Ĵ���ְ�����㣺" & vbCrLf & vbCrLf & IIf(blnҽ��, "��ҽ���򹫷Ѳ���,", "") & _
            "��ҩƷҪ��ְ������Ϊ""" & Split(STR_ְ��, ",")(intְ��B - 1) & """�����´�,��ҽ��""" & strҽ�� & """��ְ��Ϊ""" & Split(STR_ְ��, ",")(intְ��A - 1) & """��"
    End If
    CheckOneDuty = strMsg
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetFullNO(ByVal strNO As String, ByVal intNum As Integer) As String
'���ܣ����û�����Ĳ��ݵ��ţ�����ȫ���ĵ��š�
'������intNum=��Ŀ���,Ϊ0ʱ�̶��������
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String, intType As Integer
    Dim dtCurDate As Date
    Dim strMaxNo As String
    
    On Error GoTo errH
    
    If Len(strNO) >= 8 Then
        GetFullNO = Right(strNO, 8)
        Exit Function
    ElseIf Len(strNO) = 7 Then
        GetFullNO = PreFixNO & strNO
        Exit Function
    End If

    GetFullNO = strNO
    
    strSql = "Select ��Ź���,Sysdate as ����,������ From ������Ʊ� Where ��Ŀ���=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, App.ProductName, intNum)
    dtCurDate = date
    If Not rsTmp.EOF Then
        intType = NVL(rsTmp!��Ź���, 0)
        dtCurDate = rsTmp!����
        strMaxNo = NVL(rsTmp!������)
    End If

    If strMaxNo = "" Then
        strMaxNo = PreFixNO & "0000001"
    End If
    
    If intType = 1 Then
        '���ձ��
        strSql = Format(CDate("1992-" & Format(rsTmp!����, "MM-dd")) - CDate("1992-01-01") + 1, "000")
        GetFullNO = PreFixNO & strSql & Format(Right(strNO, 4), "0000")
    Else
        '������,ȡ�������ǰ��λ
        GetFullNO = Left(strMaxNo, 2) & Format(Right(strNO, 6), "000000")
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function InitSysPar() As Boolean
'��ʼ��ȫ�ֲ���
    Dim strValue As String
    On Error Resume Next
    
    gstrSysName = GetSetting("ZLSOFT", "ע����Ϣ", "gstrSysName", "")
    strValue = zlDatabase.GetPara("���뷨")
    gstrIme = IIf(strValue = "", "���Զ�����", strValue)
    
        '���ý��С����λ��
    gbytDec = Val(zlDatabase.GetPara(9, glngSys, , 2))
    gstrDec = "0." & String(gbytDec, "0")
    
    'ִ�к��Զ����
    '��35�汾���Ѿ�ɾ���˸ò�����ʹ�øò����ĳ��򣬰��ղ���ֵΪ1��true�ķ�ʽ���д���
    gblnִ�к���� = True ' Val(zlDatabase.GetPara(81, glngSys)) <> 0
    
    'һ��ͨ������֤
    glng������֤ = Val(Split(zlDatabase.GetPara(28, glngSys), "|")(0))
    
    '����һ��ͨ,��Ŀִ��ǰ�������շѻ��ȼ������
    gblnִ��ǰ�Ƚ��� = Val(zlDatabase.GetPara(163, glngSys)) <> 0
    
    If Not GetUserInfo Then
        MsgBox "��ǰ�û�δ���ö�Ӧ����Ա��Ϣ,����ϵͳ����Ա��ϵ,�ȵ��û���Ȩ���������ã�", vbInformation, gstrSysName
        Exit Function
    End If

    gstrUnitName = GetUnitName
    
    '����Ĭ����ɫ
    glngColor(1) = RGB(186, 186, 186)
    glngColor(2) = RGB(255, 215, 0)
    glngColor(3) = RGB(255, 0, 255)
    glngColor(4) = RGB(255, 0, 130)
    glngColor(5) = RGB(0, 255, 0)
    glngColor(6) = RGB(130, 255, 255)
    glngColor(7) = RGB(255, 255, 0)
    glngColor(8) = RGB(0, 0, 255)
    glngColor(9) = RGB(0, 160, 0)
    
    InitSysPar = True
End Function
Public Function MergeImageFiles(ByVal strCurrUID As String, ByVal strNewUID As String, _
    Optional ByVal strReceiveDate As String = "", Optional ByVal strMoveFiles As String = "", _
    Optional ByVal blnSaveReportImg As Boolean = False) As Boolean
'------------------------------------------------
'���ܣ���һ������Ӱ���ļ�ת�Ƶ���������ȥ��֧��Ӱ�������ȡ������
'������ strCurrUID ����Դ���UID
'       strNewUID ����ת�ƺ��µ�Ŀ�ļ��UID
'       strReceiveDate ���� �������ڣ���������ͼ��洢·������strNewUID�������ݿ���ʱ������Ҫʹ�ñ�����
'       strMoveFiles ���� ��Ҫ�ƶ����ļ����б�ʹ��"|"�ָ��ļ��������û�У���ת��Դ���UIDָ���Ŀ¼�µ�����ͼ��
'���أ�True--�ɹ���False��ʧ��
'------------------------------------------------
    Dim objSrcFtp As New clsFtp, objDestFtp As New clsFtp
    Dim strSrcPath As String, strDestPath As String
    Dim rsTmp As New ADODB.Recordset, strSql As String, strTmpFile As String
    Dim aFiles() As String, i As Integer, objFile As New Scripting.FileSystemObject
    Dim strFTPUser As String, strFTPPassw As String, strFTPHost As String, strFTPRoot As String
    Dim lngResult As Long       '��¼FTP�����Ľ��
        
    '����¼��UID���ɼ��UID������Ϊ�ϲ���ɣ���ֱ���˳�
    If strCurrUID = strNewUID Then
        MergeImageFiles = True
        Exit Function
    End If
    
    On Error GoTo errH

    '�����ƶ��ķ���ͬ��Դͼ�п����ڡ�Ӱ����ʱ��¼�����ߡ�Ӱ�����¼����
    '����ʱ����ʱ��¼���Ƶ�������¼��ȡ������ʱ��������¼���Ƶ���ʱ��¼
    
    strSql = "Select D.FTP�û��� As FtpUser,D.FTP���� As FtpPwd," & _
        "D.IP��ַ As Host," & _
        "'/'||D.FtpĿ¼||'/' As Root,Decode(C.��������,Null,'',to_Char(C.��������,'YYYYMMDD')||'/')" & _
        "||C.���UID As URL " & _
        "From Ӱ�����¼ C,Ӱ���豸Ŀ¼ D " & _
        "Where Decode(C.λ��һ,Null,C.λ�ö�,C.λ��һ)=D.�豸��(+)" & _
        "And C.���UID= [1] Union All " & _
        "Select D.FTP�û��� As FtpUser,D.FTP���� As FtpPwd," & _
        "D.IP��ַ As Host," & _
        "'/'||D.FtpĿ¼||'/' As Root,Decode(C.��������,Null,'',to_Char(C.��������,'YYYYMMDD')||'/')" & _
        "||C.���UID As URL " & _
        "From Ӱ����ʱ��¼ C,Ӱ���豸Ŀ¼ D " & _
        "Where Decode(C.λ��һ,Null,C.λ�ö�,C.λ��һ)=D.�豸��(+)" & _
        "And C.���UID= [1]"
    '�����ݿ��в�ѯ�ɼ��UID
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "ZLPACSWork", strCurrUID)
    '��ǰ���UID�����ݿ��в����ڣ����˳�������
    If rsTmp.EOF Then
        Exit Function
    End If
    
    '�洢������FTP��������
    strFTPHost = NVL(rsTmp("Host"))
    strFTPPassw = NVL(rsTmp("FtpPwd"))
    strFTPRoot = NVL(rsTmp("Root"))
    strFTPUser = NVL(rsTmp("FtpUser"))
    strSrcPath = NVL(rsTmp("Root")) & NVL(rsTmp("URL"))
    lngResult = objSrcFtp.FuncFtpConnect(strFTPHost, strFTPUser, strFTPPassw)
    If lngResult = 0 Then Exit Function     'FTP����ʧ�ܣ��˳�����
    
    '�����ݿ��в�ѯ�¼��UID����ʼ��Ŀ��Ftp,���Ŀ��UID�����ڣ�����һ����·��
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "ZLPACSWork", strNewUID)
    If rsTmp.EOF Then
    '������ͼ��ת����ʱͼ���ʱ��Ŀ�ļ��UID��ʱ������������ݿ��У���ʱֱ��ʹ��ԭ�е�FTP����
    '�������ݿ���ת�Ƽ�¼��ʱ�򣬻���ʹ��ԭ����FTP����
        If strReceiveDate <> "" Then
                objDestFtp.FuncFtpConnect strFTPHost, strFTPUser, strFTPPassw
                strDestPath = strFTPRoot & Format(strReceiveDate, "YYYYMMDD") & "/" & strNewUID
                '����FTPĿ¼
                objDestFtp.FuncFtpMkDir strFTPRoot, Format(strReceiveDate, "YYYYMMDD") & "/" & strNewUID
        Else
            Exit Function
        End If
    Else
        objDestFtp.FuncFtpConnect NVL(rsTmp("Host")), NVL(rsTmp("FtpUser")), NVL(rsTmp("FtpPwd"))
        strDestPath = NVL(rsTmp("Root")) & NVL(rsTmp("URL"))
    End If
    
    '��ȡ��Ҫ�ƶ����ļ���
    If strMoveFiles <> "" Then
        aFiles = Split(strMoveFiles, "|")
    Else
        aFiles = Split(objSrcFtp.FuncDirFiles(strSrcPath), "|")
    End If
    
    
    Dim dcmImages As New DicomImages
    Dim dcmImg As New DicomImage
    
    '��ת��ͼ��
    For i = 0 To UBound(aFiles)
        strTmpFile = App.Path & "\TmpImage\" & aFiles(i)
        lngResult = objSrcFtp.FuncDownloadFile(strSrcPath, strTmpFile, aFiles(i))
        If lngResult <> 0 Then Exit Function
        lngResult = objDestFtp.FuncUploadFile(strDestPath, strTmpFile, aFiles(i))
        If lngResult <> 0 Then Exit Function
                       
        On Error Resume Next
        If blnSaveReportImg And strMoveFiles <> "" Then
            'ɾ��ftp�еı���ͼ
            Call objSrcFtp.FuncDelFile(strSrcPath, aFiles(i) & ".jpg")
            
            Call dcmImages.Clear
            Set dcmImg = dcmImages.ReadFile(strTmpFile)

            Call dcmImg.FileExport(strTmpFile & ".jpg", "JPG")
            Call objDestFtp.FuncUploadFile(strDestPath, strTmpFile & ".jpg", aFiles(i) & ".jpg")
            
            Kill strTmpFile
        End If
        On Error GoTo 0
    Next i
    
    
    'ת��ͼ��ɹ�����ɾ����ʱͼ���ԭ��FTP��ͼ���Ŀ¼���峡�������ִ�����Բ�����
    On Error Resume Next
    For i = 0 To UBound(aFiles)
        strTmpFile = App.Path & "\TmpImage\" & aFiles(i)
        
        Call Kill(strTmpFile)
        Call objSrcFtp.FuncDelFile(strSrcPath, aFiles(i))
        
        
        If strMoveFiles <> "" Then
            If Dir(strTmpFile & ".jpg") <> "" Then Call Kill(strTmpFile & ".jpg")
            
            'ɾ�����ر���ͼ��
            strTmpFile = App.Path & "\TmpImage" & Mid(Replace(strSrcPath, "/", "\"), 2) & "\" & aFiles(i) & ".jpg"
            If Dir(strTmpFile) <> "" Then Call Kill(strTmpFile)
        End If
        
'        'ɾ������Dicomͼ��(����ͼ����Բ���ɾ��)
'        strTmpFile = App.Path & "\TmpImage" & Mid(Replace(strSrcPath, "/", "\"), 2) & "\" & aFiles(i)
'        If Dir(strTmpFile) <> "" Then Call Kill(strTmpFile)
    Next i
    Call objSrcFtp.FuncFtpDelDir(Replace(strSrcPath, strCurrUID, ""), strCurrUID)
    
    objSrcFtp.FuncFtpDisConnect
    objDestFtp.FuncFtpDisConnect
    MergeImageFiles = True
    Exit Function
errH:
    objSrcFtp.FuncFtpDisConnect
    objDestFtp.FuncFtpDisConnect
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Sub MkLocalDir(ByVal strDir As String)
'------------------------------------------------
'���ܣ���������Ŀ¼
'������ strDir��������Ŀ¼
'���أ���
'------------------------------------------------
    Dim objFile As New Scripting.FileSystemObject
    Dim aNestDirs() As String, i As Integer
    Dim strPath As String
    On Error Resume Next
    
    '��ȡȫ����Ҫ������Ŀ¼��Ϣ
    ReDim Preserve aNestDirs(0)
    aNestDirs(0) = strDir
    
    strPath = objFile.GetParentFolderName(strDir)
    Do While Len(strPath) > 0
        ReDim Preserve aNestDirs(UBound(aNestDirs) + 1)
        aNestDirs(UBound(aNestDirs)) = strPath
        strPath = objFile.GetParentFolderName(strPath)
    Loop
    '����ȫ��Ŀ¼
    For i = UBound(aNestDirs) To 0 Step -1
        MkDir aNestDirs(i)
    Next
End Sub


Public Sub ClearCacheFolder(ByVal strCacheFolder As String)
'------------------------------------------------
'���ܣ���ָ��Ŀ¼�Ĵ�С�ﵽһ���ٷֱ�ʱ����ո�Ŀ¼
'������ strCacheFolder--��Ҫ����Ƿ���յ�Ŀ¼
'���أ���
'------------------------------------------------
    Dim objFile As New Scripting.FileSystemObject
    Dim objCurFolder As Scripting.Folder, objCurFile As Scripting.File, objFiles As Scripting.Files
    Dim strDriver As String
    
    On Error Resume Next
    strDriver = objFile.GetDriveName(strCacheFolder)
    Set objCurFolder = objFile.GetFolder(strCacheFolder)
    If objCurFolder.Size / objFile.GetDrive(strDriver).FreeSpace > 0.2 Then
        zl9comlib.zlCommFun.ShowFlash "�����ͼ�񻺳�Ŀ¼����ȴ���", gfrmMain
        objCurFolder.Delete True
        zl9comlib.zlCommFun.StopFlash
    End If
End Sub

Public Function GetTrayHeight() As Long
    '------------------------------------------------------------------------------------------------------------------
    '����:��ȡ�������ĸ߶�
    '------------------------------------------------------------------------------------------------------------------
    Dim lngHwd As Long
    Dim objRect As RECT
    
    On Error Resume Next
    
    lngHwd = FindWindow("shell_traywnd", "")
    objRect = zlControl.GetControlRect(lngHwd)

    GetTrayHeight = Screen.TwipsPerPixelX * (objRect.Bottom - objRect.Top)
    
    If GetTrayHeight < 0 Then GetTrayHeight = 0
    
End Function

Public Sub ResizeRegion(ByVal ImageCount As Integer, ByVal RegionWidth As Long, _
    ByVal RegionHeight As Long, Rows As Integer, Cols As Integer)
'-----------------------------------------------------------------------------
'���ܣ����������ͼ��������ͼ������Ŀ�Ⱥ͸߶ȣ�������ѵ�ͼ����������������
'������ ImageCount����ͼ������
'       RegionWidth--ͼ����ʾ����Ŀ��
'       RegionHeight--ͼ����ʾ����ĸ߶�
'       Rows����[����]�������
'       Cols����[����]�������
'���أ������������Rows���������Cols
'-----------------------------------------------------------------------------
    Dim iCols As Integer, iRows As Integer
    Dim iBase As Integer, blnDoLoop As Integer
    Dim lngFreeCount As Long
    
    If RegionHeight = 0 Then RegionHeight = 1
    If RegionWidth = 0 Then RegionWidth = 1
    
    On Error GoTo err
    iCols = CInt(Sqr(ImageCount * RegionWidth / RegionHeight))
    iRows = CInt(Sqr(ImageCount * RegionHeight / RegionWidth))

    If iRows < 1 Then iRows = 1
    If iCols < 1 Then iCols = 1
    
    '��ͼ���ʽΪ���µ���ʽʱ����Ҫ�����н�������
    
    '��ʽ1��
    'ͼ1  ͼ2  ͼ3  ͼ4
    'ͼ5  ͼ6  ͼ7  ͼ8
    '��1  ��2  ��3  ��4
    
    '��ʽ2��
    'ͼ1  ͼ2  ͼ3  ͼ4
    'ͼ5  ͼ6  ͼ7  ͼ8
    'ͼ9  ��1  ��2  ��3
    
    lngFreeCount = iRows * iCols - ImageCount
    Do While lngFreeCount >= iCols Or lngFreeCount >= iRows
        If lngFreeCount >= iCols Then
            iRows = iRows - 1
        Else
            iCols = iCols - 1
        End If
        
        lngFreeCount = iRows * iCols - ImageCount
    Loop
    
    If iRows < 1 Then iRows = 1
    If iCols < 1 Then iCols = 1
    
    Do While iRows * iCols < ImageCount
        If RegionWidth / iCols > RegionHeight > iRows Then
            iCols = iCols + 1
        Else
            iRows = iRows + 1
        End If
    Loop
    
    '�ٴ�����������
    lngFreeCount = iRows * iCols - ImageCount
    Do While lngFreeCount >= iCols Or lngFreeCount >= iRows
        If lngFreeCount >= iCols Then
            iRows = iRows - 1
        Else
            iCols = iCols - 1
        End If
        
        lngFreeCount = iRows * iCols - ImageCount
    Loop
    
    Rows = iRows: Cols = iCols
    
err:
End Sub


Public Function funGetStudyUID(ByVal strOldStudyUID As String) As String
'-----------------------------------------------------------------------------
'����:��ѯ���ݿ⣬�жϵ�ǰͼ��ļ��UID�Ƿ��Ѿ����������������ʱ���У�
'     ������ڣ����ڼ��UID�������Ӻ�׺����������ֱ�ӷ�������ļ��UID
'�޸���:�ƽ�
'�޸�����:2007-1-27
'-----------------------------------------------------------------------------
    '
    Dim rsMatch As New ADODB.Recordset
    
    funGetStudyUID = strOldStudyUID
    gstrSQL = "select ���UID from Ӱ�����¼ where ���UID = [1]" & _
              " Union All Select ���UID from Ӱ����ʱ��¼ where ���UID = [1]"
    Set rsMatch = zlDatabase.OpenSQLRecord(gstrSQL, "PACSͼ�񱣴�", strOldStudyUID)
    If Not rsMatch.EOF Then
        '����һ���µļ��UID
        gstrSQL = "Select Ӱ����UID���_ID.Nextval From Dual"
        Set rsMatch = zlDatabase.OpenSQLRecord(gstrSQL, "PACSͼ�񱣴�")
        If Len(strOldStudyUID) <= 55 Then
            funGetStudyUID = strOldStudyUID & ".A" & rsMatch(0)
        Else
            funGetStudyUID = Left(strOldStudyUID, 55) & ".A" & rsMatch(0)
        End If
    End If
End Function


Public Function GetImageAttribute(objAttr As DicomAttributes, ByVal AttrName As String) As Variant
'-----------------------------------------------------------------------------
'����:��ȡDICOM���Լ��е�ָ������ֵ
'�޸���:�ƽ�
'�޸�����:2007-2-6
'-----------------------------------------------------------------------------
    Dim AttrTag() As String
    
    GetImageAttribute = ""
    AttrTag = Split(AttrName, ":")
    If objAttr("&h" & AttrTag(0), "&h" & AttrTag(1)).Exists Then
        GetImageAttribute = NVL(objAttr("&h" & AttrTag(0), "&h" & AttrTag(1)).value)
    End If
End Function

Public Function SetDeptPara(ByVal lngDeptID As Long, ByVal varPara As String, ByVal strValue As String) As Boolean
'���ܣ�����ָ���Ĳ���ֵ
'������lngDept=����ID
'      varPara=������
'      strValue=������ֵ
'���أ������Ƿ�ɹ�
    Dim strSql As String
    
    On Error GoTo errH
        
    strSql = "ZL_Ӱ�����̲���_UPDATE(" & lngDeptID & ",'" & varPara & "','" & strValue & "')"
    Call zlDatabase.ExecuteProcedure(strSql, "SetPara")
    
    '���óɹ����������
    Set mrsDeptParas = Nothing
    
    SetDeptPara = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
End Function

Public Function GetDeptPara(ByVal lngDeptID As Long, ByVal varPara As String, Optional ByVal strDefault As String, Optional ByVal blnNotCache As Boolean) As String
'���ܣ���ȡָ���Ĳ���ֵ
'������lngDept=����ID
'      varPara=������
'      strDefault=�����ݿ���û�иò���ʱʹ�õ�ȱʡֵ(ע�ⲻ��Ϊ��ʱ)
'      blnNotCache=�Ƿ񲻴ӻ����ж�ȡ
'���أ�����ֵ���ַ�����ʽ
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String, blnNew As Boolean
    
    On Error GoTo errH
    
    If blnNotCache Then
        Set rsTmp = New ADODB.Recordset
        strSql = "Select ����ֵ from Ӱ�����̲��� where ����ID = [1] and ������=[2]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "��ȡ����", lngDeptID, varPara)
        
        If Not rsTmp.EOF Then
            GetDeptPara = NVL(rsTmp!����ֵ)
        Else
            GetDeptPara = strDefault
        End If
    Else
        '��һ�μ��ز�������
        If mrsDeptParas Is Nothing Then
            blnNew = True
        ElseIf mrsDeptParas.State = 0 Then
            blnNew = True
        End If
        If blnNew Then
            strSql = "Select ����ֵ,������,����ID from Ӱ�����̲���"
            Set mrsDeptParas = New ADODB.Recordset
            Set mrsDeptParas = zlDatabase.OpenSQLRecord(strSql, "��ȡ����")
        End If
        
        '���ݻ����ȡ����ֵ
        mrsDeptParas.Filter = "������='" & CStr(varPara) & "' AND ����ID=" & lngDeptID
        If Not mrsDeptParas.EOF Then
            GetDeptPara = NVL(mrsDeptParas!����ֵ)
        Else
            GetDeptPara = strDefault
        End If
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
End Function

Public Function GetIsValidOfStorageDevice(ByVal lngDeptID As Long) As Boolean
'��ʼ�����Ҽ�����
    Dim rsTmp As New ADODB.Recordset
    Dim strSaveDeviceID As String
    
    On Error GoTo DBError
    
    '��ȡ�����洢�豸��
    strSaveDeviceID = GetDeptPara(lngDeptID, "�洢�豸��")
    
    gstrSQL = "Select �豸��,�豸�� From Ӱ���豸Ŀ¼ Where ����=1 and �豸��=[1] and NVL(״̬,0)=1"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ�洢�豸��Ϣ", strSaveDeviceID)
    
    
    GetIsValidOfStorageDevice = Not rsTmp.EOF
    
    Exit Function
DBError:
    GetIsValidOfStorageDevice = False
    
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function

Public Sub subCancelSeriesRelate(frmParent As Form, ByVal lngAdviceNo As Long, ByVal lngSendNO As Long, _
    ByVal strSeriesNo As String, Optional ByVal blnSaveReportImg = False)
'-----------------------------------------------------------------------------
'����:ȡ������ͼ��Ĺ������ƶ�FTPͼ���µ�λ�ã��޸����ݿ��¼������ʽ���ƶ�����ʱ����
'������ frmParent -- ������
'       lngAdviceNo ����ҽ��ID
'       lngSendNO ���� ���ͺ�
'       strSeriesNo ��������UID
'���أ���
'-----------------------------------------------------------------------------
    
    Dim mcnFTP As New clsFtp
    Dim strSql As String
    Dim rsTmp As ADODB.Recordset
    Dim strCachePath As String
    Dim strCacheFileName As String
    Dim objFile As New Scripting.FileSystemObject
    Dim imgs As New DicomImages
    Dim img As New DicomImage
    Dim strNewStudyUID As String    '�����ɵļ��UID
    Dim strOldStudyUID As String    'ͼ������ԭ���ļ��UID
    Dim strDBStudyUID As String     '���ݿ��б���ļ��UID����ͼ��洢·�����
    Dim strMoveFiles As String  '�洢��Ҫ�ƶ���ͼ���ļ�����ʹ�á�|���ָ�
    Dim blnNoImage As Boolean   '1û��ͼ��ֱ�Ӷ�ȡ���ݿ���Ϣ��0��ͼ��ʹ��ͼ����Ϣ
    Dim lngResult As Long    '��¼FTP���ؽ��
    
    'ͼ���еĲ��˻�����Ϣ
    Dim strModality As String
    Dim strPatientId As String
    Dim strPatientName As String
    Dim strSex As String
    Dim strAge As String
    Dim strDateOfBirth As String
    Dim strManufacturer As String
    Dim strReceiveDateTime As String
    
    
    On Error GoTo DBError
    
    '���������е�һ��ͼ��� ����ID��Ӣ�������Ա����䣬�������ڣ����UID������豸������ʱ��
    strCachePath = App.Path & "\TmpImage\"
    strSql = "Select A.ͼ���,D.FTP�û��� As User1,D.FTP���� As Pwd1,a.ͼ��UID, " & _
        "D.IP��ַ As Host1,c.���uid," & _
        "'/'||D.FtpĿ¼||'/' As Root1,Decode(C.��������,Null,'',to_Char(C.��������,'YYYYMMDD')||'/')" & _
        "||C.���UID||'/'||A.ͼ��UID As URL,d.�豸�� as �豸��1, " & _
        "E.FTP�û��� As User2,E.FTP���� As Pwd2," & "E.IP��ַ As Host2," & _
        "'/'||E.FtpĿ¼||'/' As Root2,e.�豸�� as �豸��2 " & _
        "From Ӱ����ͼ�� A,Ӱ�������� B,Ӱ�����¼ C,Ӱ���豸Ŀ¼ D,Ӱ���豸Ŀ¼ E " & _
        "Where A.����UID=B.����UID And B.���UID=C.���UID And C.λ��һ=D.�豸��(+) And C.λ�ö�=E.�豸��(+) " & _
        "And A.����UID= [1] Order By A.ͼ���"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "������", strSeriesNo)
    
    If Not rsTmp.EOF Then   '�����д���ͼ��
        strDBStudyUID = NVL(rsTmp("���uid"))
        '�½�����Ŀ¼
        strCacheFileName = strCachePath & rsTmp("URL")
        MkLocalDir objFile.GetParentFolderName(strCacheFileName)
        
        '����ͼ��
        If rsTmp("�豸��1") <> "" And mcnFTP.FuncFtpConnect(NVL(rsTmp("Host1")), NVL(rsTmp("User1")), NVL(rsTmp("Pwd1"))) <> 0 Then
            mcnFTP.FuncDownloadFile objFile.GetParentFolderName(NVL(rsTmp("Root1")) & rsTmp("URL")), strCacheFileName, objFile.GetFileName(rsTmp("URL"))
            mcnFTP.FuncFtpDisConnect
        ElseIf rsTmp("�豸��2") <> "" And mcnFTP.FuncFtpConnect(NVL(rsTmp("Host2")), NVL(rsTmp("User2")), NVL(rsTmp("Pwd2"))) <> 0 Then
            mcnFTP.FuncDownloadFile objFile.GetParentFolderName(NVL(rsTmp("Root2")) & rsTmp("URL")), strCacheFileName, objFile.GetFileName(rsTmp("URL"))
            mcnFTP.FuncFtpDisConnect
        Else
            'FTP���Ӵ�����ʾ���˳�����ȡ����������
            MsgBoxD frmParent, "FTP���Ӵ��󣬲���ȡ��������" & vbCrLf & vbCrLf & "�������������ӳ������⡣"
            Exit Sub
        End If
                    
        '��ȡͼ����Ϣ
        If Dir(strCacheFileName) <> vbNullString Then
            Set img = imgs.ReadFile(strCacheFileName)
            'ʹ�ñ�����ͼ�������Ϣ��ȡ����
            strOldStudyUID = img.StudyUID
            strModality = GetImageAttribute(img.Attributes, ATTR_Ӱ�����)
            strPatientId = img.PatientID
            strPatientName = img.Name
            strSex = img.Sex
            If IsDate(img.DateOfBirthAsDate) Then
                If img.Attributes(&H10, &H1010).Exists And Not IsNull(img.Attributes(&H10, &H1010)) Then
                    strAge = img.Attributes(&H10, &H1010).value
                Else
                    strAge = CStr(Year(date) - Year(img.DateOfBirthAsDate))
                End If
                        
                If img.DateOfBirthAsDate <> "0:00:00" Then
                    strDateOfBirth = Format(img.DateOfBirthAsDate, "YYYY-MM-DD")
                Else
                    strDateOfBirth = ""
                End If
            Else
                strAge = "": strDateOfBirth = ""
            End If
            strManufacturer = GetImageAttribute(img.Attributes, ATTR_����豸)
            strReceiveDateTime = GetImageAttribute(img.Attributes, ATTR_�������) & " " & _
                        Format(GetImageAttribute(img.Attributes, ATTR_���ʱ��), "HH:MM")
            'ɾ����ʱͼ��
            Set img = Nothing
            imgs.Remove (1)
            On Error Resume Next
            objFile.DeleteFile strCacheFileName
            On Error GoTo 0
        Else
            '�����һ��ͼ�����ز���ȷ����ȡ���ݿ���Ϣ���������������
            blnNoImage = True
        End If
    Else
        '������û��ͼ�󣬲���Ҫȡ��������Ӧ�ò�������������
        Exit Sub
    End If
    
    '����û��ͼ����Ϣ�ɶ�ȡ������ͼ����Ҫ��Ϣ��ȡ�������ģ�ֱ�Ӷ�ȡ���ݿ��е���Ϣ
    If blnNoImage = True Or Trim(strReceiveDateTime) = "" Then
        strSql = "select a.Ӱ�����,a.����,a.����,a.Ӣ����,a.�Ա�,a.����,a.��������,a.���uid," & _
                " a.����豸,a.�������� from Ӱ�����¼ a where a.ҽ��id =[1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "��ȡ����", lngAdviceNo)
        If Not rsTmp.EOF Then
            If blnNoImage = True Then
                strOldStudyUID = NVL(rsTmp("���uid"))
                strDBStudyUID = NVL(rsTmp("���uid"))
                strPatientId = NVL(rsTmp("����"))
                strPatientName = NVL(rsTmp("Ӣ����"))
                strSex = NVL(rsTmp("�Ա�"))
                strAge = NVL(rsTmp("����"))
                strDateOfBirth = NVL(rsTmp("��������"), "")
                strManufacturer = NVL(rsTmp("����豸"))
            End If
            strModality = NVL(rsTmp("Ӱ�����"))
            strReceiveDateTime = NVL(rsTmp("��������"))
        End If
    End If
    '��֯ͼ���ļ����ƴ�
    strSql = "select ͼ��UID from Ӱ�������� a,Ӱ����ͼ�� b where a.����UID =[1] and a.����UID = b.����UID"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "��ȡ����", strSeriesNo)
    If Not rsTmp.EOF Then
        strMoveFiles = rsTmp(0)
        rsTmp.MoveNext
        While Not rsTmp.EOF
            strMoveFiles = strMoveFiles & "|" & rsTmp(0)
            rsTmp.MoveNext
        Wend
    End If
    
    '������UID�����ݿ����ִ�ļ��UID��ͬ���򴴽��µļ��UID�����޸�ͼ��FTP·��
    strNewStudyUID = funGetStudyUID(strOldStudyUID)
    If strNewStudyUID <> strDBStudyUID Then
        If MergeImageFiles(strDBStudyUID, strNewStudyUID, Format(strReceiveDateTime, "YYYY-MM-DD"), strMoveFiles, blnSaveReportImg) = False Then
            MsgBoxD frmParent, "ͼ��ת�Ʋ��ɹ�������ȡ��������"
            Exit Sub
        End If
    End If
    
    '�޸����ݿ⣬������¼ת����ʱ��¼
    strSql = "ZL_Ӱ����_PhotoCancel(" & lngAdviceNo & "," & lngSendNO & ",'" & strNewStudyUID & "','" & _
              strSeriesNo & "','" & strModality & "','" & strPatientId & "','" & _
              strPatientName & "','" & strSex & "','" & strAge & "'," & _
              IIf(Len(strDateOfBirth) = 0, "null", "to_date('" & strDateOfBirth & "','YYYY-MM-DD')") & _
              ",'" & strManufacturer & "',to_date('" & strReceiveDateTime & "','YYYY-MM-DD HH24:MI:SS'))"
              
    zlDatabase.ExecuteProcedure strSql, "ȡ������"
    
    Exit Sub
DBError:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub


Public Sub GetAllImages(frmParent As Form, dcmViewer As DicomViewer, blnMoved As Boolean, intSearchType As Integer, _
    Optional lngAdviceID As Long, Optional strSeriesUID As String, Optional intGetImgNum As Integer = 0, _
    Optional intShowImgNum As Integer = 0, Optional blnTempDB As Boolean = False, _
    Optional strStudyUID As String = "", Optional strImageUID As String = "")
'------------------------------------------------
'���ܣ�ɾ��dcmViewer�е�ͼ��󣬽���ȡ��ͼ���ļ�����dcmViewer��
'������ frmParent -- ������
'       dcmViewer������ͼ���DicomViewer
'       blnMoved ���� �Ƿ�ת����
'       intSearchType ������������,ֻ����ʽ���ѯ��Ч  1������ҽ��ID�ͷ��ͺŲ飬2����������UID�飬3 - ����ͼ��UID��
'       lngAdviceID ���� ҽ��ID
'       strSeriesUID ���� ����UID
'       intGetImgNum �������ζ�ȡ��ͼ������
'       intShowImgNum ����������ʾ��ͼ������
'       blnTempDB - - �Ƿ����ʱ������ȡͼ��
'       strStudyUID - - ���UID,ֻ�д���ʱ����ҵ�ʱ�򣬲�ʹ���������
'       strImageUID - - ͼ��UID��ֻ�д���ʽ����ҵ�ʱ�򣬲�ʹ���������
'���أ��ޣ�ֱ���޸�dcmViewer����ʾ��ͼ��
'------------------------------------------------
    
    Dim strSql As String
    Dim rsTmp As New ADODB.Recordset
    Dim curImage As DicomImage, i As Integer
    Dim iCols As Integer, iRows As Integer
    Dim objFile As New Scripting.FileSystemObject, strTmpFile As String
    Dim Inet1 As New clsFtp
    Dim Inet2 As New clsFtp
    Dim strCachePath As String
    Dim iCurrentIndex As Integer
    Dim dcmTag As clsImageTagInf
    
    On Error GoTo DBError
    If blnTempDB = False Then       '����ʽͼ����в���ͼ��
        strSql = "Select A.ͼ���,D.FTP�û��� As User1,D.FTP���� As Pwd1," & _
            "D.IP��ַ As Host1,'/'||D.FtpĿ¼||'/' As Root1," & _
            "Decode(C.��������,Null,'',to_Char(C.��������,'YYYYMMDD')||'/')" & _
            "||C.���UID||'/'||A.ͼ��UID As URL,d.�豸�� as �豸��1, " & _
            "E.FTP�û��� As User2,E.FTP���� As Pwd2," & _
            "E.IP��ַ As Host2,'/'||E.FtpĿ¼||'/' As Root2," & _
            "e.�豸�� as �豸��2, A.ͼ��UID,C.���UID,B.����UID,A.��̬ͼ,A.��������,A.�ɼ�ʱ��, A.¼�Ƴ��� " & _
            "From Ӱ����ͼ�� A,Ӱ�������� B,Ӱ�����¼ C,Ӱ���豸Ŀ¼ D,Ӱ���豸Ŀ¼ E " & _
            "Where A.����UID=B.����UID And B.���UID=C.���UID And C.λ��һ=D.�豸��(+) And C.λ�ö�=E.�豸��(+) "
        If blnMoved Then
            strSql = Replace(strSql, "Ӱ����ͼ��", "HӰ����ͼ��")
            strSql = Replace(strSql, "Ӱ��������", "HӰ��������")
            strSql = Replace(strSql, "Ӱ�����¼", "HӰ�����¼")
        End If
        If intShowImgNum <> 0 Then
            strSql = strSql & " And Rownum<=[2] "
        End If
        
        If intSearchType = 1 Then       '1������ҽ��ID�ͷ��ͺŲ�
            strSql = strSql & "And C.ҽ��ID=[1] Order By A.ͼ���"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "��ȡͼ��", lngAdviceID, intGetImgNum)
        ElseIf intSearchType = 2 Then   '2����������UID��
            strSql = strSql & "And A.����UID= [1] Order By A.ͼ���"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "��ȡͼ��", strSeriesUID, intGetImgNum)
        ElseIf intSearchType = 3 Then   '3 - ����ͼ��UID��
            strSql = strSql & "And A.ͼ��UID = [1] "
            Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "��ȡͼ��", strImageUID, intGetImgNum)
        End If
        
    Else                '����ʱ���в���ͼ��
        
        strSql = "Select c.ͼ���,d.FTP�û��� As User1, d.FTP���� As Pwd1, d.Ip��ַ As Host1," _
                & "'/' || d.FtpĿ¼ || '/' As Root1," _
                & " Decode(a.��������, Null, '', To_Char(a.��������, 'YYYYMMDD') || '/') || a.���uid || '/' || c.ͼ��uid As URL," _
                & " d.�豸�� As �豸��1,C.ͼ��UID,a.���UID,b.����UID,d.FTP�û��� As User2, d.FTP���� As Pwd2, " _
                & " d.Ip��ַ As Host2, '/' || d.FtpĿ¼ || '/' As Root2, " _
                & " d.�豸�� As �豸��2,C.��̬ͼ,C.��������, C.�ɼ�ʱ��, C.¼�Ƴ��� " _
                & " From Ӱ����ʱ��¼ a,Ӱ����ʱ���� b,Ӱ����ʱͼ�� c ,Ӱ���豸Ŀ¼ d " _
                & " Where a.���UID=b.���UID And b.����UID = c.����UID And a.λ��һ = d.�豸�� "
                
        If strStudyUID <> "" Then   '���ռ��uid����
            strSql = strSql & "And a.���UID=[1] and Rownum<=[2] Order By c.ͼ���  "
            Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "��ȡͼ��", strStudyUID, CLng(6))
        Else        '��������UID����
            strSql = strSql & "And b.����UID=[1] and Rownum<=[2] Order By c.ͼ���  "
            Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "��ȡͼ��", strSeriesUID, CLng(6))
        End If
    End If
    
        dcmViewer.Images.Clear
        If rsTmp.RecordCount > 0 Then
            If intShowImgNum = 0 Or intShowImgNum >= rsTmp.RecordCount Then
                ResizeRegion rsTmp.RecordCount, dcmViewer.Width, dcmViewer.Height, iRows, iCols
            Else
                ResizeRegion intShowImgNum, dcmViewer.Width, dcmViewer.Height, iRows, iCols
            End If
            dcmViewer.MultiColumns = iCols
            dcmViewer.MultiRows = iRows
            
            '��������Ŀ¼
            strCachePath = IIf(Len(App.Path) > 3, App.Path & "\TmpImage\", App.Path & "TmpImage\")
            MkLocalDir strCachePath & objFile.GetParentFolderName(NVL(rsTmp("URL")))
            
            Do While Not rsTmp.EOF
                
                strTmpFile = strCachePath & NVL(rsTmp("URL"))
                If NVL(rsTmp("��̬ͼ"), IMGTAG) = VIDEOTAG Then
                    strTmpFile = IIf(Len(App.Path) > 3, App.Path & "\..\�����ļ�\Avi.bmp", App.Path & "..\�����ļ�\Avi.bmp")
                ElseIf NVL(rsTmp("��̬ͼ"), IMGTAG) = AUDIOTAG Then
                    strTmpFile = IIf(Len(App.Path) > 3, App.Path & "\..\�����ļ�\wav.bmp", App.Path & "..\�����ļ�\wav.bmp")
                End If
                
                If Dir(strTmpFile) = vbNullString Then
                    '���ػ���ͼ�񲻴��ڣ����ȡFTPͼ��
                    
                    '����FTP����
                    If NVL(rsTmp("�豸��1")) <> vbNullString And Inet1.hConnection = 0 Then
                        If Inet1.FuncFtpConnect(NVL(rsTmp("Host1")), NVL(rsTmp("User1")), NVL(rsTmp("Pwd1"))) = 0 Then
                            If NVL(rsTmp("�豸��2")) <> vbNullString Then
                                If Inet2.FuncFtpConnect(NVL(rsTmp("Host2")), NVL(rsTmp("User2")), NVL(rsTmp("Pwd2"))) = 0 Then
                                    MsgBoxD frmParent, "FTP�����������ӣ������������á�"
                                    Exit Sub
                                End If
                            Else
                                MsgBoxD frmParent, "FTP�����������ӣ������������á�"
                                Exit Sub
                            End If
                        End If
                    End If
                    
                    If Inet1.FuncDownloadFile(objFile.GetParentFolderName(NVL(rsTmp("Root1")) & rsTmp("URL")), strTmpFile, objFile.GetFileName(rsTmp("URL"))) <> 0 Then
                        '���豸��1��ȡͼ��ʧ�ܣ�����豸��2��ȡͼ��
                        If NVL(rsTmp("�豸��2")) <> vbNullString Then
                            If Inet2.hConnection = 0 Then Inet2.FuncFtpConnect NVL(rsTmp("Host2")), NVL(rsTmp("User2")), NVL(rsTmp("Pwd2"))
                            Call Inet2.FuncDownloadFile(objFile.GetParentFolderName(NVL(rsTmp("Root2")) & rsTmp("URL")), strTmpFile, objFile.GetFileName(rsTmp("URL")))
                        End If
                    End If
                End If
      
                If Dir(strTmpFile) <> vbNullString Then
                    
                
                    
                    If NVL(rsTmp("��̬ͼ"), IMGTAG) <> VIDEOTAG And NVL(rsTmp("��̬ͼ"), IMGTAG) <> AUDIOTAG Then
                        Set curImage = dcmViewer.Images.ReadFile(strTmpFile)
                        
    
                        Set dcmTag = New clsImageTagInf
                        dcmTag.tag = NVL(rsTmp("��̬ͼ"), IMGTAG)
                        
                        
                        Set curImage.tag = dcmTag
                        
                        With curImage
                            .BorderStyle = 6
                            .BorderWidth = 1
                            .BorderColour = vbWhite
                        End With
                    Else
                        Set curImage = New DicomImage
                        
                        On Error GoTo continue
                            Call curImage.FileImport(strTmpFile, "DIB/BMP")
continue:
                        
                        Set dcmTag = New clsImageTagInf
                        dcmTag.tag = NVL(rsTmp("��̬ͼ"), VIDEOTAG)
                        dcmTag.EncoderName = NVL(rsTmp("��������"), "")
                        dcmTag.CaptureTime = NVL(rsTmp("�ɼ�ʱ��"))
                        
                        If NVL(rsTmp("��̬ͼ"), VIDEOTAG) = VIDEOTAG Then
                            dcmTag.videoFile = strCachePath & NVL(rsTmp("URL")) & ".avi"
                        Else
                            dcmTag.videoFile = strCachePath & NVL(rsTmp("URL")) & ".wav"
                        End If
                        
                        dcmTag.RecordTimeLen = Val(NVL(rsTmp("¼�Ƴ���"), "0"))
                        
'                        '�������Ƶ¼���ļ������ڲ���ʱ��������
'                        If Trim(dcmTag.VideoFile) <> "" And Dir(dcmTag.VideoFile) <> "" Then
'                            Name dcmTag.VideoFile As dcmTag.VideoFile & ".avi"
'                        End If
                        
                        Set curImage.tag = dcmTag
                        
                        curImage.InstanceUID = NVL(rsTmp("ͼ��UID"))
                        curImage.SeriesUID = NVL(rsTmp("����UID"))
                        curImage.StudyUID = NVL(rsTmp("���UID"))
                        
                        
                        
                        Call AddVideoLabelToDicomImage(curImage, _
                            IIf(dcmTag.tag = VIDEOTAG, "¼��ʱ�䣺", "¼��ʱ�䣺") & NVL(rsTmp("�ɼ�ʱ��")), _
                            IIf(dcmTag.tag = VIDEOTAG, "¼�񳤶ȣ�", "¼�����ȣ�") & NVL(rsTmp("¼�Ƴ���"), "0") & " ��", _
                            "�������ƣ�" & NVL(rsTmp("��������")))
                        
                        With curImage
                            .BorderStyle = 6
                            .BorderWidth = 1
                            .BorderColour = vbWhite
                        End With
                        
                        Call dcmViewer.Images.Add(curImage)
                    End If
                    
                    
                    'ȡ���Զ���Ӱ,��ΪDicomObjects�ؼ�����Դ����Ӱ��BUG�����ڣ�0028��6100��ʱ�����Զ���ͼ����м�Ӱ��
                    '���½�ú��DSAͼ����������ʾ
                    '��Ȼ����ͼ���mask=0 ,����ȡ����Ӱ������ÿ��ͼ����ӵ��µ�Dicomimages֮���Զ��ֽ�mask���ó�1�ˣ�
                    '�����ڳ������޷��ܺõĿ��ƣ����ֱ��ȥ����0028��6100��������ԡ�
                    If Not IsNull(curImage.Attributes(&H28, &H6100).value) Then
                        curImage.Attributes.Remove &H28, &H6100
                    End If
                End If
                
                rsTmp.MoveNext
            Loop
            If dcmViewer.Images.Count > 0 Then
                dcmViewer.CurrentIndex = 1
                dcmViewer.Images(1).BorderColour = vbRed
            End If
        Else
            dcmViewer.MultiColumns = 1
            dcmViewer.MultiRows = 1
        End If
    Inet1.FuncFtpDisConnect
    Inet2.FuncFtpDisConnect
    Exit Sub
DBError:
    Inet1.FuncFtpDisConnect
    Inet2.FuncFtpDisConnect
    If ErrCenter() = 1 Then
        Resume
    End If
    
    Call SaveErrLog
End Sub


Public Sub AddVideoLabelToDicomImage(dcmImage As DicomImage, ByVal strCaptureTimeText As String, _
    ByVal strTimeLenText As String, ByVal strEncoderName As String)
    '����:���label
    '����:dcmImage��dicomͼ��
    '     strCaption�� label�ı�
    Dim labCaption As New DicomLabel
    
    labCaption.LabelType = doLabelText
    '����ʾ������������
    labCaption.Text = strCaptureTimeText & vbCrLf & strTimeLenText '& vbCrLf & strEncoderName
    labCaption.Font.Bold = True
    labCaption.Font.Name = "����"
    labCaption.Font.Size = 10
    labCaption.ForeColour = vbYellow
    labCaption.AutoSize = False

    
    labCaption.Left = 0
    labCaption.Top = 0
    
    Call dcmImage.Labels.Add(labCaption)
End Sub


Public Function GetSingleImage(lngImageUID As String, lngSerialUID As String, Optional blnMoved As Boolean = False) As Boolean
    '����:��FTP�����ļ�
    '����:����UID
    '�������سɹ�����ļ�·��
    Dim strSql As String
    Dim rsTmp As New ADODB.Recordset
    Dim strCachePath As String
    Dim Inet1 As New clsFtp
    Dim Inet2 As New clsFtp
    Dim strDeviceNO1 As String
    Dim strDeviceNO2 As String
    Dim strTmpFile As String
    Dim objFile As New Scripting.FileSystemObject
    Dim DicomImg As New DicomImages
    
    On Error GoTo WriteFileErr
    
    GetSingleImage = True
    
    strSql = "Select A.ͼ���, A.��̬ͼ, D.FTP�û��� As User1,D.FTP���� As Pwd1,a.ͼ��UID, " & _
        "D.IP��ַ As Host1," & _
        "'/'||D.FtpĿ¼||'/' As Root1,Decode(C.��������,Null,'',to_Char(C.��������,'YYYYMMDD')||'/')" & _
        "||C.���UID||'/'||A.ͼ��UID As URL1,d.�豸�� as �豸��1, " & _
        "E.FTP�û��� As User2,E.FTP���� As Pwd2," & _
        "E.IP��ַ As Host2," & _
        "'/'||E.FtpĿ¼||'/' As Root2,Decode(C.��������,Null,'',to_Char(C.��������,'YYYYMMDD')||'/')" & _
        "||C.���UID||'/'||A.ͼ��UID As URL2 , e.�豸�� as �豸��2, A.��̬ͼ,A.�������� " & _
        "From Ӱ����ͼ�� A,Ӱ�������� B,Ӱ�����¼ C,Ӱ���豸Ŀ¼ D,Ӱ���豸Ŀ¼ E " & _
        "Where A.����UID=B.����UID And B.���UID=C.���UID And C.λ��һ=D.�豸��(+) And C.λ�ö�=E.�豸��(+) " & _
        "And A.ͼ��UID= [1]  and a.����UID = [2]  Order By A.ͼ���"
    If blnMoved Then
        strSql = Replace(strSql, "Ӱ����ͼ��", "HӰ����ͼ��")
        strSql = Replace(strSql, "Ӱ��������", "HӰ��������")
        strSql = Replace(strSql, "Ӱ�����¼", "HӰ�����¼")
    End If
            
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "�����ļ�", lngImageUID, lngSerialUID)
    strCachePath = App.Path & "\TmpImage\"
    ClearCacheFolder strCachePath
    
    If rsTmp.EOF <> True Then
        MkLocalDir strCachePath & objFile.GetParentFolderName(NVL(rsTmp("URL1")))
    End If
    
    Do While Not rsTmp.EOF
        If strDeviceNO1 <> rsTmp("�豸��1") Then
            strDeviceNO1 = rsTmp("�豸��1")
            Inet1.FuncFtpConnect NVL(rsTmp("Host1")), NVL(rsTmp("User1")), NVL(rsTmp("Pwd1"))
        End If
        
        If strDeviceNO2 <> rsTmp("�豸��2") Then
            strDeviceNO2 = rsTmp("�豸��2")
            Inet2.FuncFtpConnect NVL(rsTmp("Host2")), NVL(rsTmp("User2")), NVL(rsTmp("Pwd2"))
        End If
        
        If rsTmp("��̬ͼ") = VIDEOTAG Then
            strTmpFile = strCachePath & NVL(rsTmp("URL1")) & ".avi"
        ElseIf rsTmp("��̬ͼ") = AUDIOTAG Then
            strTmpFile = strCachePath & NVL(rsTmp("URL1")) & ".wav"
        Else
            strTmpFile = strCachePath & NVL(rsTmp("URL1"))
        End If
        
        If Dir(strTmpFile) = "" Then
'            Inet.strIPAddress = Nvl(rsTmp("Host1")): Inet.strUser = Nvl(rsTmp("User1")): Inet.strPsw = Nvl(rsTmp("Pwd1"))
            If Inet1.FuncDownloadFile(objFile.GetParentFolderName(NVL(rsTmp("Root1")) & rsTmp("URL1")), strTmpFile, objFile.GetFileName(rsTmp("URL1"))) <> 0 Then
                strTmpFile = strCachePath & NVL(rsTmp("URL2"))
'                Inet.strIPAddress = Nvl(rsTmp("Host2")): Inet.strUser = Nvl(rsTmp("User2")): Inet.strPsw = Nvl(rsTmp("Pwd2"))
                Call Inet2.FuncDownloadFile(objFile.GetParentFolderName(NVL(rsTmp("Root2")) & rsTmp("URL2")), strTmpFile, objFile.GetFileName(rsTmp("URL2")))
            End If
        End If

        DoEvents
        rsTmp.MoveNext
    Loop
    
    Inet1.FuncFtpDisConnect
    Inet2.FuncFtpDisConnect
    
    Exit Function
WriteFileErr:
    Inet1.FuncFtpDisConnect
    Inet2.FuncFtpDisConnect
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function funGetStorageDevice(frmParent As Form, strSaveDeviceID As String, ByRef strDirURL As String, ByRef strIP As String, _
        ByRef strUser As String, ByRef strPwd As String) As Boolean
'------------------------------------------------
'���ܣ������ݿ��ж�ȡ�ƶ��洢�豸ID��FTP���ʲ���
'������ frmParent  -- ������
'       strSaveDeviceID �����洢�豸ID
'       strDirURL����[OUT] FTPĿ¼
'       strIp ����[OUT] IP��ַ
'       strUser ���� [OUT]�û���
'       strPwd ����[OUT]�û���
'���أ�True������ȡ�ɹ���False������ȡʧ��
'-----------------------------------------------
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    
    '���洢�豸�Ƿ����
    strSql = "Select '/'||Decode(FtpĿ¼,Null,'',FtpĿ¼||'/') As URL,FTP�û���,FTP����,IP��ַ " & _
        "From Ӱ���豸Ŀ¼ " & "Where �豸��=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, App.ProductName, strSaveDeviceID)
     'û�д洢�豸ʱ�˳�
    If rsTemp.EOF = True Then
        MsgBoxD frmParent, "û���ҵ��洢�豸,������ѡ��洢�豸!", vbInformation, gstrSysName
        funGetStorageDevice = False
        Exit Function
    End If
    strDirURL = NVL(rsTemp("URL"))
    strIP = NVL(rsTemp("IP��ַ"))
    strUser = NVL(rsTemp("FTP�û���"))
    strPwd = NVL(rsTemp("FTP����"))
    funGetStorageDevice = True
End Function

Public Function funGetFtpDeviceInf(frmParent As Form, objFtp As TFtpDeviceInf) As Boolean
'------------------------------------------------
'���ܣ������ݿ��ж�ȡ�ƶ��洢�豸ID��FTP���ʲ���
'������ frmParent  -- ������
'       strSaveDeviceID �����洢�豸ID
'       strDirURL����[OUT] FTPĿ¼
'       strIp ����[OUT] IP��ַ
'       strUser ���� [OUT]�û���
'       strPwd ����[OUT]�û���
'���أ�True������ȡ�ɹ���False������ȡʧ��
'-----------------------------------------------
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    
    objFtp.strFtpDir = ""
    objFtp.strFtpIp = ""
    objFtp.strFTPUser = ""
    objFtp.strFTPPwd = ""
    
    '���洢�豸�Ƿ����
    strSql = "Select '/'||Decode(FtpĿ¼,Null,'',FtpĿ¼||'/') As URL,FTP�û���,FTP����,IP��ַ From Ӱ���豸Ŀ¼ Where �豸��=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, App.ProductName, objFtp.strDeviceId)
    
     'û�д洢�豸ʱ�˳�
    If rsTemp.EOF = True Then
        MsgBoxD frmParent, "û���ҵ��洢�豸,������ѡ��洢�豸!", vbInformation, gstrSysName
        funGetFtpDeviceInf = False
        
        Exit Function
    End If
    
    objFtp.strFtpDir = NVL(rsTemp("URL"))
    objFtp.strFtpIp = NVL(rsTemp("IP��ַ"))
    objFtp.strFTPUser = NVL(rsTemp("FTP�û���"))
    objFtp.strFTPPwd = NVL(rsTemp("FTP����"))
    
    funGetFtpDeviceInf = True
End Function

Public Function Open3DViewer(lngAdviceID As Long, objParent As Object, Optional ByVal blnMoved As Boolean = False) As Long
'���ܣ�3D��Ƭ
'������
'   lngAdviceId---ҽ��ID
    
    On Error GoTo DBError
    
    If lngAdviceID > 0 Then
        Open3DViewer = XWShow3DImage(lngAdviceID, objParent)
    Else
        If gblnXWLog = True Then
            Call WriteCommLog("Open3DViewer", "����XWShow3DImage�ӿ�", "ҽ��IDΪ��")
        End If
    End If
    
    Exit Function
DBError:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function OpenViewer(ByVal lngViewerType As Long, ByRef objPacsCore As Object, lngAdviceID As Long, _
        blnAddImage As Boolean, objParent As Object, Optional ByVal strSerials As String = "", _
        Optional ByVal blnMoved As Boolean = False, Optional ByVal blnLocalizerBackward As Boolean = False, _
        Optional ByVal intImageInterval As Integer = 0, Optional ByVal strImageString As String = "") As Boolean
'------------------------------------------------
'���ܣ����ݴ����ҽ��ID�ͷ��ͺţ���objPacsCoreָ��Ĺ�Ƭվ
'������
'       lngViewerType -- չ��ͼ���Viewer���ͣ�1-�����ר��Viewer��2-�ٴ������Viewer
'       objPacsCore ������Ƭվ����
'       lngAdviceID ����ҽ��ID
'       blnAddImage--True ��ԭ��ͼ����������ӵ�ǰͼ��Falseɾ��ԭ��ͼ�񣬴򿪵�ǰͼ��
'       objParent -- ������
'       strSerials������ѡ������UID���ƴ����ö��ŷָ�����������룬��ѡ��ȫ������
'       blnMoved������ѡ���Ƿ�ת��
'       blnLocalizerBackward--��ѡ����λ�����,��strImageString����
'       intImageInterval ---��ѡ����ͼ��ļ��������5����ʾÿ5��ͼ��һ��ͼ,��strImageString����
'       strImageString --- ��ѡ��ÿ����������Ҫ�򿪵�ͼ�����ϣ���intImageInterval��blnLocalizerBackward���⣬
'                           ��strImageStringΪ��
'                           �����ǡ�����UID1|1-3;5-27;33-100+����UID2|ȫ����,ȫ����ʾ��ȫ��ͼ��
'���أ�ͼ���ļ���������
'------------------------------------------------
    Dim strSql As String
    Dim rsTmp As New ADODB.Recordset
    Dim strFTPHost As String, strFtpPath As String, strFTPUser As String, strFTPPswd As String
    Dim strSDPath As String, strSDUser As String, strSDPwd As String
    Dim strDeviceNO As String
    Dim i As Integer
    Dim blnConnectDS As Boolean         '�Ƿ����ӵ�ǰ�Ĺ���Ŀ¼
    Dim oneMessage As TGetImgMsg        'Ԥȡͼ�����Ϣ�ṹ
    Dim intImageLocation As Integer
    Dim strXWViewerFilter As String
    Dim strStudyUID As String
    
    On Error GoTo DBError
    
    '��ѯͼ��������PACS����������PACS
    strSql = "Select ���UID,ͼ��λ��,Ӱ����� from Ӱ�����¼ where ҽ��ID =[1]"
    
    If blnMoved Then
        strSql = Replace(strSql, "Ӱ�����¼", "HӰ�����¼")
    End If
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "��ѯͼ�����ڵ�λ��", lngAdviceID)
    
    If rsTmp.RecordCount <> 0 Then
        intImageLocation = NVL(rsTmp!ͼ��λ��, 0)
        strStudyUID = NVL(rsTmp!���UID, "")
    Else
        intImageLocation = 1    '�鲻�����ݣ�˵��ʹ����������RIS
    End If
    
    'ͼ�����������ݿ⣬��������DLL��ʾͼ��
    If intImageLocation = 1 Or intImageLocation = 2 Then
        strXWViewerFilter = lngAdviceID & IIf(strSerials <> "", "[;]" & strSerials, "")
        
        If gblnXWLog = True Then
            Call WriteCommLog("OpenViewer", "����XWShowImage�ӿ�", "��ѯ���˲���Ϊ��" & strXWViewerFilter & "��ͼ��λ��Ϊ��" & intImageLocation)
        End If
        
        '��������ADViewer����WEB Viewer
        Call XWShowImage(lngViewerType, strXWViewerFilter, lngAdviceID, IIf(strSerials <> "", glngSeriesSchemeNo, glngStudySchemeNo))
        
        OpenViewer = True
        
        '���ͼ�񱣴�����ƽ̨������ʾ�û���Ҫ�ȴ������Ҵ���PACSͼ������
        If intImageLocation = 2 Then
            Call XWDownLoadImage(strStudyUID)
        End If
        
        Exit Function
    End If
    
    '�ж��Ƿ��������°�pacs��Ƭ�����ʹ�����°��Ƭ�������°��Ƭ��������ͼ��
    If gblnUseXinWangView = True Then
        Call OpenViewerWithInXWPacs(lngAdviceID, NVL(rsTmp!Ӱ�����), blnMoved)
        
        OpenViewer = True
        Exit Function
    End If
    
    
    'ͼ�����������ݿ⣬��������zl9PacsCore��ʾͼ��
    strFTPHost = ""
           
    '������Ҫ�򿪵�����ͼ����Ϣ
    strSql = "Select D.IP��ַ As Host1,d.�豸�� as �豸��1," & _
        "Decode(C.��������,Null,'',to_Char(C.��������,'YYYYMMDD')||'/')" & _
        "||C.���UID||'/' As Path,E.IP��ַ As Host2,e.�豸�� as �豸��2, " & _
        "D.����Ŀ¼ AS ����Ŀ¼1, E.����Ŀ¼ AS ����Ŀ¼2,D.����Ŀ¼�û��� as ����Ŀ¼�û���1, " & _
        "E.����Ŀ¼�û��� AS ����Ŀ¼�û���2,D.����Ŀ¼���� AS ����Ŀ¼����1,E.����Ŀ¼���� AS ����Ŀ¼����2, " & _
        "D.FTPĿ¼ as FTPĿ¼1, E.FTPĿ¼ as FTPĿ¼2,D.FTP�û��� as FTP�û���1, E.FTP�û��� AS FTP�û���2,  " & _
        "D.FTP���� as FTP����1, E.FTP���� AS FTP����2 " & _
        "From Ӱ�����¼ C,Ӱ���豸Ŀ¼ D,Ӱ���豸Ŀ¼ E " & _
        "Where C.λ��һ=D.�豸��(+) And C.λ�ö�=E.�豸��(+) And C.ҽ��ID=[1] "
    
    '�����ת����־�����ȡת������ʷ��
    If blnMoved Then
        strSql = Replace(strSql, "Ӱ����ͼ��", "HӰ����ͼ��")
        strSql = Replace(strSql, "Ӱ��������", "HӰ��������")
        strSql = Replace(strSql, "Ӱ�����¼", "HӰ�����¼")
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "��ȡ����Ŀ¼��Ϣ", lngAdviceID)
    
    If rsTmp.RecordCount > 0 Then
        '�������صĻ���Ŀ¼����Ҫ�ڵ��ù�Ƭվ֮ǰ�ȴ������Ŀ¼����Ƭվ��ֻ�����أ����������ػ���Ŀ¼
        MkLocalDir App.Path & "\TmpImage\" & rsTmp("Path")
        
        '��ȡFTP�����������û��������룬IP��ַ��
        If rsTmp("�豸��1") <> "" Then
            strDeviceNO = rsTmp("�豸��1")
            strFTPHost = rsTmp("Host1")
            strFtpPath = NVL(rsTmp("FTPĿ¼1"))
            strFTPUser = NVL(rsTmp("FTP�û���1"))
            strFTPPswd = NVL(rsTmp("FTP����1"))
            strSDPath = NVL(rsTmp("����Ŀ¼1"))
            strSDUser = NVL(rsTmp("����Ŀ¼�û���1"))
            strSDPwd = NVL(rsTmp("����Ŀ¼����1"))
        ElseIf NVL(rsTmp("�豸��2")) <> "" Then
            strDeviceNO = rsTmp("�豸��2")
            strFTPHost = rsTmp("Host2")
            strFtpPath = NVL(rsTmp("FTPĿ¼2"))
            strFTPUser = NVL(rsTmp("FTP�û���2"))
            strFTPPswd = NVL(rsTmp("FTP����2"))
            strSDPath = NVL(rsTmp("����Ŀ¼2"))
            strSDUser = NVL(rsTmp("����Ŀ¼�û���2"))
            strSDPwd = NVL(rsTmp("����Ŀ¼����2"))
        End If
        
        '�жϹ���Ŀ¼�Ƿ��Ѿ����ӣ����û�����ӣ����������
        blnConnectDS = True
        For i = 1 To UBound(gConnectedShardDir)
            If gConnectedShardDir(i) = strDeviceNO Then
                blnConnectDS = False
                Exit For
            End If
        Next i
        If blnConnectDS = True And strSDPath <> "" Then
            If funcConnectShardDir(objParent, "\\" & strFTPHost & "\" & strSDPath, strSDUser, strSDPwd) = 0 Then
                ReDim Preserve gConnectedShardDir(UBound(gConnectedShardDir) + 1) As String
                gConnectedShardDir(UBound(gConnectedShardDir)) = strDeviceNO
            End If
        End If
        
        '�򿪹�Ƭվ
        If objPacsCore Is Nothing Then
            Exit Function
        Else
            objPacsCore.CallOpenViewer strImageString, lngAdviceID, objParent, gcnOracle, blnMoved, blnAddImage, intImageInterval, glngSys
        End If
        
        '�ȴ򿪹�Ƭվ����Ԥȡ
        oneMessage.strSubDir = rsTmp("Path")
        oneMessage.strDestMainDir = App.Path & "\TmpImage\"
        oneMessage.strIP = strFTPHost
        oneMessage.strFtpDir = strFtpPath
        oneMessage.strFTPUser = strFTPUser
        oneMessage.strFTPPswd = strFTPPswd
        oneMessage.strSDDir = strSDPath
        oneMessage.strSDUser = strSDUser
        oneMessage.strSDPswd = strSDPwd
        
        Call funPreDownLoadImages(oneMessage)
        
    Else    'û�в��ҵ�ͼ���¼����ر�ԭ���Ѿ��򿪵Ĺ�Ƭ����
        If Not objPacsCore Is Nothing Then objPacsCore.Closefrom
    End If
    
    OpenViewer = True
    Exit Function
DBError:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Public Function GetStudyImageData(ByVal lngAdviceID As Long, ByVal blnMoved As Boolean) As ADODB.Recordset
'��ȡ���ͼ������

    Dim strSql As String
        
    strSql = "Select rownum as ˳���, A.ͼ���,D.FTP�û��� As User1,D.FTP���� As Pwd1," & _
        "D.IP��ַ As Host1,'/'||D.FtpĿ¼||'/' As Root1," & _
        "Decode(C.��������,Null,'',to_Char(C.��������,'YYYYMMDD')||'/')" & _
        "||C.���UID||'/'||A.ͼ��UID As URL,d.�豸�� as �豸��1, " & _
        "E.FTP�û��� As User2,E.FTP���� As Pwd2," & _
        "E.IP��ַ As Host2,'/'||E.FtpĿ¼||'/' As Root2," & _
        "e.�豸�� as �豸��2, A.ͼ��UID,C.���UID,B.����UID,A.��̬ͼ,A.��������,A.�ɼ�ʱ��, A.¼�Ƴ��� " & _
        "From Ӱ����ͼ�� A,Ӱ�������� B,Ӱ�����¼ C,Ӱ���豸Ŀ¼ D,Ӱ���豸Ŀ¼ E " & _
        "Where A.����UID=B.����UID And B.���UID=C.���UID And C.λ��һ=D.�豸��(+) And C.λ�ö�=E.�豸��(+) and c.ҽ��ID=[1] "
        

    If blnMoved Then
        strSql = Replace(strSql, "Ӱ����ͼ��", "HӰ����ͼ��")
        strSql = Replace(strSql, "Ӱ��������", "HӰ��������")
        strSql = Replace(strSql, "Ӱ�����¼", "HӰ�����¼")
    End If
    
    Set GetStudyImageData = zlDatabase.OpenSQLRecord(strSql, "��ѯͼ����Ϣ", lngAdviceID)
End Function


Public Function OpenViewerWithInXWPacs(ByVal lngAdviceID As Long, ByVal strModalityType As String, ByVal blnMoved As Boolean)
'���°�pacs�д�����PACS��ͼƬ��Ƭ��������ϰ汾�����ݣ���ʹ����������Ƭϵͳ����ֱ�Ӵ���Զ��Ŀ¼�ļ���
    Dim rsTemp As ADODB.Recordset

    Dim strFtpUrl As String
    Dim strImages As String
    
    Set rsTemp = GetStudyImageData(lngAdviceID, blnMoved)
    
    strImages = ""

    While Not rsTemp.EOF
        If NVL(rsTemp!�豸��1) <> "" Then
            strFtpUrl = "\\" & NVL(rsTemp!Host1) & "\" & gstrImageShareDir & NVL(rsTemp!Root1) & NVL(rsTemp!Url)
        Else
            strFtpUrl = "\\" & NVL(rsTemp!Host2) & "\" & gstrImageShareDir & NVL(rsTemp!Root2) & NVL(rsTemp!Url)
        End If
        
        If strImages <> "" Then strImages = strImages & "[;]"
        
        strFtpUrl = Replace(strFtpUrl, "//", "/")
        strImages = strImages & Replace(strFtpUrl, "/", "\")
        
        rsTemp.MoveNext
    Wend

    
    '��Զ��Ŀ¼�ļ����жԱȹ�Ƭ
    Call OEMViewOpen(0, strImages, 0, strModalityType)
End Function

Public Function CheckChargeState(ByVal lngҽ��ID As Long, ByVal lng��Դ As Long) As ChargeState
'���ܣ� �жϵ�ǰ��ҽ���Ƿ��շ�
'������ lngҽ��ID --ҽ��ID
'       lng��Դ -- ������Դ

'һ��ҽ�����жಿλ����ҽ��

    Dim rsData As New ADODB.Recordset, rsTemp As ADODB.Recordset, rsClone As ADODB.Recordset
    Dim strTable As String
    Dim lngTempCharged As ChargeState
    
    lngTempCharged = ChargeState.�޷���

    gstrSQL = "Select B.������Դ, A.NO,B.id as ҽ��ID,B.���ID, A.�Ʒ�״̬, A.��¼����,D.����ģʽ" & vbNewLine & _
                "From ����ҽ������ A, ����ҽ����¼ B,  ������Ϣ D" & vbNewLine & _
                "Where A.ҽ��Id=B.ID And B.����ID=D.����ID And (B.id = [1] or B.���Id=[1]) "
                
    Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, "�Ƿ��շ�", lngҽ��ID)
    
    rsData.Filter = "���ID=NULL"
    
    If NVL(rsData!��¼����, 2) = 2 Then '����
        If NVL(rsData!�Ʒ�״̬, -1) = -1 Or NVL(rsData!�Ʒ�״̬, -1) = 0 Then   '��
            lngTempCharged = ChargeState.�޷���
        Else
            If NVL(rsData!�Ʒ�״̬, -1) = 1 Then                                '��
                lngTempCharged = ChargeState.�Ѽ���
            ElseIf NVL(rsData!�Ʒ�״̬, -1) = 2 Then                            '��
                lngTempCharged = ChargeState.�ѵ���
            ElseIf NVL(rsData!�Ʒ�״̬, -1) = 4 Then                            '��
                lngTempCharged = ChargeState.������
            End If
        End If
    Else                                '�շ�
        If NVL(rsData!�Ʒ�״̬, -1) = -1 Or NVL(rsData!�Ʒ�״̬, -1) = 0 Then   '��
            lngTempCharged = ChargeState.�޷���
        Else
            If NVL(rsData!�Ʒ�״̬, -1) = 1 Then                                'Ƿ
                lngTempCharged = ChargeState.δ�շ�
            ElseIf NVL(rsData!�Ʒ�״̬, -1) = 2 Then                            '��
                lngTempCharged = ChargeState.�ѵ���
            ElseIf NVL(rsData!�Ʒ�״̬, -1) = 3 Then                            '��
                lngTempCharged = ChargeState.���շ�
            ElseIf NVL(rsData!�Ʒ�״̬, -1) = 4 Then                            '��
                lngTempCharged = ChargeState.���˷�
            End If
        End If
    End If
    
    CheckChargeState = lngTempCharged
End Function

Public Function CheckConcurrentReport(frmParent As Form, ByVal lngOrderID As Long, Optional blnSilence As Boolean = False) As Boolean
'------------------------------------------------
'���ܣ���鵱ǰ�����Ƿ���ҽ�����ڲ�������
'������ frmParent  -- ������
'       lngOrderID ����ҽ��ID
'       blnSilence--True �����ֲ�����ʾ��False ����ʱ������ʾ��Ϣ
'���أ�True-���˲����˱��棻False--�������ڲ����˱���
'------------------------------------------------

Dim rsTemp As ADODB.Recordset
    
    On Error GoTo errHandle
    CheckConcurrentReport = True
    gstrSQL = "Select ������� From Ӱ�����¼ Where ҽ��ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ��¼", lngOrderID)
    
    If Not rsTemp Is Nothing Then
        If Not rsTemp.EOF Then
            If NVL(rsTemp!�������) <> "" And NVL(rsTemp!�������) <> UserInfo.���� Then
                If blnSilence = False Then
                    MsgBoxD frmParent, "��ǰ���˵ı������ڱ� " & NVL(rsTemp!�������) & " ���������Ժ����ԡ�", vbInformation, gstrSysName
                End If
                CheckConcurrentReport = False
            End If
        End If
    End If
    Exit Function
    
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Sub UpdateReporter(ByVal lngOrderID As Long, ByVal Reporter As String)
    On Error GoTo errHandle
    
    gstrSQL = "ZL_Ӱ�񱨸����_Update(" & lngOrderID & ",'" & Reporter & "')"
    zlDatabase.ExecuteProcedure gstrSQL, "���²�����"
    Exit Sub
    
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Public Function bln����δ�󻮼۵�(ByVal lngҽ��ID As Long, ByVal lng��Դ As Long) As Boolean
    Dim rsTemp As ADODB.Recordset
    Dim strFeeTable As String
    Dim strFilter As String
    
    'סԺ���˲�סԺ���ü�¼���������Ȳ��˲�������ü�¼
    If lng��Դ = 2 Then
        strFeeTable = "סԺ���ü�¼"
        strFilter = " A.��¼����"
    Else
        strFeeTable = "������ü�¼"
        strFilter = " decode(A.��¼����,1,1,11,1,A.��¼����)"
    End If

    On Error GoTo errHandle
    gstrSQL = "Select /*+ RULE */ A.ID" & vbNewLine & _
            "From " & strFeeTable & " A" & vbNewLine & _
            "Where A.ҽ����� + 0 In (Select ID From ����ҽ����¼ Where ID = [1] Or ���id = [1]) And (" & strFilter & ", A.NO) In" & vbNewLine & _
            "      (Select ��¼����, NO" & vbNewLine & _
            "       From ����ҽ������" & vbNewLine & _
            "       Where ҽ��id In (Select ID From ����ҽ����¼ Where ID = [1] Or ���id = [1])" & vbNewLine & _
            "       Union All" & vbNewLine & _
            "       Select ��¼����, NO" & vbNewLine & _
            "       From ����ҽ������" & vbNewLine & _
            "       Where ҽ��id In (Select ID From ����ҽ����¼ Where ID = [1] Or ���id = [1])" & vbNewLine & _
            "       ) And A.���ʷ��� = 1 And A.��¼״̬ = 0"

    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡδ�󻮼۵�", lngҽ��ID)
    If rsTemp.EOF Then
        Exit Function
    Else
        bln����δ�󻮼۵� = True
    End If
    Exit Function
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function bln������Ժ(ByVal lng����ID As Long, ByVal lng��ҳID As Long) As Boolean
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo errHandle
    
    strSql = "SELECT to_char(��Ժ����,'YYYY-MM-DD HH24:MI:SS') as ��Ժ���� from ������ҳ where ����ID=[1] AND ��ҳID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "��ȡ��Ժʱ��", lng����ID, lng��ҳID)
    If rsTemp.EOF Then
        Exit Function
    Else
        If NVL(rsTemp!��Ժ����) = "" Then
            bln������Ժ = True
        End If
    End If
    Exit Function
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function GetRptImages(ByRef RptViewer As DicomViewer, ByVal lngOrderID As Long, ByVal blnMoved As Boolean)
'------------------------------------------------
'���ܣ���ȡ����ͼ�񵽱��أ���ˢ����ʾ
'������ RptViewer ������ʾͼ��Ŀؼ�
'       lngOrderID -- ҽ��ID
'       blnMoved -- �Ƿ�ת��
'���أ��ޣ�ֱ����RptViewer �м���ͼ��
'------------------------------------------------
    Dim objFileSystem As New Scripting.FileSystemObject
    Dim strSql As String, rsTemp As New ADODB.Recordset
    Dim aryFiles() As String    '����ͼ������
    Dim strFiles As String      '���ֺŷָ��ĳɹ����ص��ļ�
    Dim aryRptFileName() As String    '�����ļ�������
    
    Dim cFtpNet As New clsFtp
    Dim strVirtualPath As String
    Dim strLocalPath As String
    Dim intCount As Integer
    Dim curImage As DicomImage
    
    '�����RptViewer �е�ͼ��
    RptViewer.Images.Clear
    
    '��鱾�ػ���ͼ��ĸ�Ŀ¼�Ƿ���ڣ�����������򴴽����ظ�Ŀ¼���������ʧ�ܣ���ֱ���˳�����
    If objFileSystem.FolderExists(App.Path & "\TmpImage\") = False Then objFileSystem.CreateFolder App.Path & "\TmpImage\"
    If objFileSystem.FolderExists(App.Path & "\TmpImage\") = False Then GetRptImages = False: Exit Function
    
    '�����ݿ��ȡͼ����Դ��Ϣ
    err = 0: On Error Resume Next
    strSql = "Select To_Char(L.��������, 'yyyymmdd') As ��Ŀ¼, L.���uid, L.����ͼ��, A1.FtpĿ¼ As Root1, A1.Ip��ַ As Ip1," & vbNewLine & _
            "       A1.FTP�û��� As Usr1, A1.FTP���� As Pwd1, A2.FtpĿ¼ As Root2, A2.Ip��ַ As Ip2, A2.FTP�û��� As Usr2, A2.FTP���� As Pwd2" & vbNewLine & _
            "From Ӱ�����¼ L, Ӱ���豸Ŀ¼ A1, Ӱ���豸Ŀ¼ A2" & vbNewLine & _
            "Where L.λ��һ = A1.�豸��(+) And L.λ�ö� = A2.�豸��(+) And L.ҽ��id = [1]"
    If blnMoved = True Then
        strSql = Replace(strSql, "Ӱ�����¼", "HӰ�����¼")
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "��ȡ����ͼ��", lngOrderID)
    If rsTemp.RecordCount <= 0 Then GetRptImages = False: Exit Function
    aryFiles = Split("" & rsTemp!����ͼ��, ";")
    aryRptFileName = Split("" & rsTemp!����ͼ��, ";")
    If UBound(aryFiles) < 0 Then GetRptImages = False: Exit Function
        
    '��鱾�������б��μ���Ŀ¼�Ƿ���ڣ�����������򴴽����ش洢Ŀ¼���������ʧ�ܣ����˳�����
    err = 0: On Error Resume Next
    strLocalPath = App.Path & "\TmpImage\" & rsTemp!��Ŀ¼
    If objFileSystem.FolderExists(strLocalPath) = False Then objFileSystem.CreateFolder strLocalPath
    If objFileSystem.FolderExists(strLocalPath) = False Then GetRptImages = False: Exit Function
    strLocalPath = strLocalPath & "\" & rsTemp!���UID
    If objFileSystem.FolderExists(strLocalPath) = False Then objFileSystem.CreateFolder strLocalPath
    If objFileSystem.FolderExists(strLocalPath) = False Then GetRptImages = False: Exit Function
        
    strFiles = ""
    '��鱾�ػ���ͼ���Ƿ���ڣ�������ڣ��򲻴�FTP���أ���ֱ�Ӷ�ȡ��������ͼ��
    For intCount = 0 To UBound(aryFiles)
        '����ļ����ڣ�����Ҫ���أ����ñ��
        If Dir(strLocalPath & "\" & Trim(aryFiles(intCount))) <> "" Then
            strFiles = strFiles & ";" & strLocalPath & "\" & Trim(aryFiles(intCount))
            aryFiles(intCount) = ""
        End If
    Next intCount
    
    If strFiles <> "" Then
        If Left(strFiles, 1) = ";" Then strFiles = Mid(strFiles, 2)
    End If
    
    
    '������δ��ڵ��ļ���������Ҫ�򿪵��ļ�������һ�£����FTP���ر��������ڵ�ͼ��
    If UBound(Split(strFiles, ";")) <> UBound(aryFiles) Then
        '���ȴ��豸1����ͼ��
        If "" & rsTemp!Ip1 <> "" Then
            If cFtpNet.FuncFtpConnect("" & rsTemp!Ip1, "" & rsTemp!Usr1, "" & rsTemp!Pwd1) <> 0 Then
                strVirtualPath = rsTemp!Root1 & "/" & rsTemp!��Ŀ¼ & "/" & rsTemp!���UID
                For intCount = 0 To UBound(aryFiles)
                    If aryFiles(intCount) <> "" Then
                        If cFtpNet.FuncDownloadFile(strVirtualPath, strLocalPath & "\" & Trim(aryFiles(intCount)), Trim(aryFiles(intCount))) = 0 Then
                            If Dir(strLocalPath & "\" & Trim(aryFiles(intCount))) <> "" Then
                                strFiles = strFiles & ";" & strLocalPath & "\" & Trim(aryFiles(intCount))
                                aryFiles(intCount) = ""
                            End If
                        End If
                    End If
                Next intCount
            End If
            cFtpNet.FuncFtpDisConnect
        End If
        
        '����豸1����ͼ���������ٴ��豸2����ͼ��
        If strFiles <> "" Then
            If Left(strFiles, 1) = ";" Then strFiles = Mid(strFiles, 2)
        End If
        
        If UBound(Split(strFiles, ";")) <> UBound(aryFiles) And "" & rsTemp!Ip2 <> "" Then
            If cFtpNet.FuncFtpConnect("" & rsTemp!Ip2, "" & rsTemp!Usr2, "" & rsTemp!Pwd2) <> 0 Then
                strVirtualPath = rsTemp!Root2 & "/" & rsTemp!��Ŀ¼ & "/" & rsTemp!���UID
                For intCount = 0 To UBound(aryFiles)
                    If aryFiles(intCount) <> "" Then
                        If cFtpNet.FuncDownloadFile(strVirtualPath, strLocalPath & "\" & Trim(aryFiles(intCount)), Trim(aryFiles(intCount))) = 0 Then
                            If Dir(strLocalPath & "\" & Trim(aryFiles(intCount))) <> "" Then
                                strFiles = strFiles & ";" & strLocalPath & "\" & Trim(aryFiles(intCount))
                            End If
                        End If
                    End If
                Next intCount
            End If
            cFtpNet.FuncFtpDisConnect
        End If
        If strFiles <> "" Then
            If Left(strFiles, 1) = ";" Then strFiles = Mid(strFiles, 2)
        End If
    End If
    
    '����õ��ļ�װ��
    Dim iRows As Integer, iCols As Integer
    aryFiles = Split(strFiles, ";")
    With RptViewer
        For intCount = 0 To UBound(aryFiles)
            Set curImage = New DicomImage
            Call ImportImgToDicom(curImage, aryFiles(intCount))
            
            curImage.BorderWidth = 3: curImage.BorderColour = vbWhite
            curImage.tag = aryRptFileName(intCount)
            .Images.Add curImage
        Next
        If UBound(aryFiles) >= 0 Then
            .CurrentIndex = 1
            .Images(.CurrentIndex).BorderColour = vbBlue
        End If
        If .Images.Count > 0 Then
            ResizeRegion .Images.Count, .Width, .Height, iRows, iCols
            .MultiColumns = iCols: .MultiRows = iRows
        Else
            '��������
        End If
    End With
    
    GetRptImages = True: Exit Function

errHand:
    cFtpNet.FuncFtpDisConnect
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Private Sub ImportImgToDicom(objDcmImage As DicomImage, ByVal strImgFile As String)
On Error GoTo errHandle
    Dim objTmp As StdPicture
    Dim objFs As New FileSystemObject
    
    Set objTmp = LoadPicture(strImgFile)
    
    Call objDcmImage.FileImport(strImgFile, "JPG")
Exit Sub
errHandle:
    Call objFs.DeleteFile(strImgFile, True)
End Sub

Public Sub PromptResult(lngOrderID As Long, lngModul As Long, frmParent As Form, lngCur����ID As Long, strResultInput As String, Optional strDocID As String = "")
    Dim strResult As String
    Dim strSql As String
    Dim objMsgCenter As New clsPacsMsgProcess
    
    strResult = frmResult.zlGetResult(frmParent, lngModul, lngOrderID, lngCur����ID, strResultInput)   '��ʾ���������Ժ�Ӱ������
    
    If strResult = "" Then Exit Sub
    
    If InStr(strResultInput, "Σ��״̬") > 0 Then
        If Split(strResult, "-")(0) = 2 Then     'Σ��ֵ
            strSql = "zl_Ӱ����_Σ������(" & lngOrderID & ",1)"
            
            Call objMsgCenter.OpenMsgCenter(lngModul, lngCur����ID, gstrPrivs)
            Call objMsgCenter.Send_Msg_Critical(lngOrderID)
        ElseIf Split(strResult, "-")(0) = 1 Then
            strSql = "zl_Ӱ����_Σ������(" & lngOrderID & ",0)"
        Else
            strSql = "zl_Ӱ����_Σ������(" & lngOrderID & ",'')"
        End If
        zlDatabase.ExecuteProcedure strSql, "���Σ��ֵ"
    End If
    
    If InStr(strResultInput, "�������") > 0 Then
        If Split(strResult, "-")(1) = 1 Then    '������
            strSql = "ZL_Ӱ����_���(" & lngOrderID & ",1)"
        ElseIf Split(strResult, "-")(1) = 2 Then
            strSql = "ZL_Ӱ����_���(" & lngOrderID & ",0)"
        Else
            strSql = "ZL_Ӱ����_���(" & lngOrderID & ",'')"
        End If
        zlDatabase.ExecuteProcedure strSql, "���������"
    End If
    
    If lngModul = 1290 Then         'Ӱ��ҽ��վ�ż�¼Ӱ������
        If InStr(strResultInput, "Ӱ������") > 0 Then
            Select Case Split(strResult, "-")(2)
                Case 1
                    strSql = "Zl_Ӱ������_Update(" & lngOrderID & ",1)"
                Case 2
                    strSql = "Zl_Ӱ������_Update(" & lngOrderID & ",2)"
                Case 3
                    strSql = "Zl_Ӱ������_Update(" & lngOrderID & ",3)"
                Case 4
                    strSql = "Zl_Ӱ������_Update(" & lngOrderID & ",4)"
                Case Else
                    strSql = "Zl_Ӱ������_Update(" & lngOrderID & ",'')"
            End Select
            zlDatabase.ExecuteProcedure strSql, "Ӱ������"
        End If
    End If
    
    '��¼��������
    If InStr(strResultInput, "��������") > 0 Then
        If Split(strResult, "-")(3) <> "" Then
            Select Case Split(strResult, "-")(3)
                Case 1
                    strSql = "Zl_��������_Update(" & lngOrderID & ",1)"
                Case 2
                    strSql = "Zl_��������_Update(" & lngOrderID & ",2)"
                Case 3
                    strSql = "Zl_��������_Update(" & lngOrderID & ",3)"
                Case 4
                    strSql = "Zl_��������_Update(" & lngOrderID & ",4)"
                Case Else
                    strSql = "Zl_��������_Update(" & lngOrderID & ",'')"
            End Select
            zlDatabase.ExecuteProcedure strSql, "��������"
        End If
    End If
    
    If lngModul <> 1294 Then    '�������⣬���¼��Ϸ������
        If InStr(strResultInput, "�������") > 0 Then
            If Split(strResult, "-")(4) = 1 Then    '�������
                strSql = "Zl_�������_Update(" & lngOrderID & ",'����')"
            ElseIf Split(strResult, "-")(4) = 2 Then
                strSql = "Zl_�������_Update(" & lngOrderID & ",'��������')"
            ElseIf Split(strResult, "-")(4) = 3 Then
                strSql = "Zl_�������_Update(" & lngOrderID & ",'������')"
            Else
                strSql = "Zl_�������_Update(" & lngOrderID & ",'')"
            End If
            zlDatabase.ExecuteProcedure strSql, "�������"
        End If
    End If
End Sub

Public Sub PromptResultRich(lngOrderID As Long, strDocID As String, lngModul As Long, frmParent As Form, lngCur����ID As Long, strResultInput As String)
    Dim strResult As String
    Dim strSql As String
    Dim objRichReport As New clsRichReport
    Dim objMsgCenter As New clsPacsMsgProcess
    
    strResult = frmResult.zlGetResult(frmParent, lngModul, strDocID, lngCur����ID, strResultInput)
    
    If strResult = "" Then Exit Sub
    
    objRichReport.CreatePacsInterface
    
    If InStr(strResultInput, "Σ��״̬") > 0 Then
        If Split(strResult, "-")(0) = 2 Then     'Σ��ֵ
            strSql = "zl_Ӱ����_Σ������(" & lngOrderID & ",1)"
            
            Call objMsgCenter.OpenMsgCenter(lngModul, lngCur����ID, gstrPrivs)
            Call objMsgCenter.Send_Msg_Critical(lngOrderID)
        ElseIf Split(strResult, "-")(0) = 1 Then
            strSql = "zl_Ӱ����_Σ������(" & lngOrderID & ",0)"
        Else
            strSql = "zl_Ӱ����_Σ������(" & lngOrderID & ",'')"
        End If
        zlDatabase.ExecuteProcedure strSql, "���Σ��ֵ"
    End If
    
    If InStr(strResultInput, "�������") > 0 Then
        If Split(strResult, "-")(1) = 1 Then    '������
            Call objRichReport.EvaluatResult(strDocID, "1")
        ElseIf Split(strResult, "-")(1) = 2 Then
            Call objRichReport.EvaluatResult(strDocID, "0")
        Else
            Call objRichReport.EvaluatResult(strDocID, "0")
        End If
    End If
    
    If lngModul = 1290 Then         'Ӱ��ҽ��վ�ż�¼Ӱ������
        If InStr(strResultInput, "Ӱ������") > 0 Then
            Select Case Split(strResult, "-")(2)
                Case 1
                    strSql = "Zl_Ӱ������_Update(" & lngOrderID & ",1)"
                Case 2
                    strSql = "Zl_Ӱ������_Update(" & lngOrderID & ",2)"
                Case 3
                    strSql = "Zl_Ӱ������_Update(" & lngOrderID & ",3)"
                Case 4
                    strSql = "Zl_Ӱ������_Update(" & lngOrderID & ",4)"
                Case Else
                    strSql = "Zl_Ӱ������_Update(" & lngOrderID & ",'')"
            End Select
            zlDatabase.ExecuteProcedure strSql, "Ӱ������"
        End If
    End If
    
    '��¼��������
    If InStr(strResultInput, "��������") > 0 Then
        If Split(strResult, "-")(3) <> "" Then
            Select Case Split(strResult, "-")(3)
                Case 1
                    Call objRichReport.EvaluatReportQuality(strDocID, "1")
                Case 2
                    Call objRichReport.EvaluatReportQuality(strDocID, "2")
                Case 3
                    Call objRichReport.EvaluatReportQuality(strDocID, "3")
                Case 4
                    Call objRichReport.EvaluatReportQuality(strDocID, "4")
                Case Else
                    Call objRichReport.EvaluatReportQuality(strDocID, "0")
            End Select
        End If
    End If
    
    If lngModul <> 1294 Then    '�������⣬���¼��Ϸ������
        If InStr(strResultInput, "�������") > 0 Then
            If Split(strResult, "-")(4) = 1 Then    '�������
                strSql = "Zl_�������_Update(" & lngOrderID & ",'����')"
            ElseIf Split(strResult, "-")(4) = 2 Then
                strSql = "Zl_�������_Update(" & lngOrderID & ",'��������')"
            ElseIf Split(strResult, "-")(4) = 3 Then
                strSql = "Zl_�������_Update(" & lngOrderID & ",'������')"
            Else
                strSql = "Zl_�������_Update(" & lngOrderID & ",'')"
            End If
            zlDatabase.ExecuteProcedure strSql, "�������"
        End If
    End If
End Sub


Public Function BillingWarn(frmParent As Object, ByVal strPrivs As String, _
    rsWarn As ADODB.Recordset, ByVal str���� As String, ByVal curʣ���� As Currency, _
    ByVal cur���ս�� As Currency, ByVal cur���ʽ�� As Currency, ByVal cur������� As Currency, _
    ByVal str�շ���� As String, ByVal str������� As String, str�ѱ���� As String, _
    intWarn As Integer, Optional ByVal bln���� As Boolean) As Integer
'����:�Բ��˼��ʽ��б�����ʾ
'����:rsWarn=���������������õļ�¼��(�ò��˲���,�����ֺ���ҽ��)
'     str�շ����=��ǰҪ�������,���ڷ��౨��
'     str�������=�������,������ʾ
'     bln����=���ɻ��۷���ʱ�ı��������ƾ���ǿ�Ƽ���Ȩ��ʱ�Ĵ���
'     intWarn=�Ƿ���ʾѯ���Ե���ʾ,-1=Ҫ��ʾ,0=ȱʡΪ��,1-ȱʡΪ��
'����:str�ѱ����="CDE":�����ڱ��α�����һ�����,"-"Ϊ������𡣸÷������ڴ����ظ�����
'     intWarn=����ѯ������ʾ�е�ѡ����,0=Ϊ��,1-Ϊ��
'     0;û�б���,����
'     1:������ʾ���û�ѡ�����
'     2:������ʾ���û�ѡ���ж�
'     3:������ʾ�����ж�
'     4:ǿ�Ƽ��ʱ���,����
    Dim bln�ѱ��� As Boolean, byt��־ As Byte
    Dim byt��ʽ As Byte, byt�ѱ���ʽ As Byte
    Dim arrTmp As Variant, vMsg As VbMsgBoxResult
    Dim str���� As String, i As Long
    
    BillingWarn = 0
    
    '�����������:NULL��û������,0�������˵�
    If rsWarn.State = 0 Then Exit Function
    If rsWarn.EOF Then Exit Function
    If IsNull(rsWarn!����ֵ) Then Exit Function
    
    '��Ӧ���λ��Ч��������
    If Not IsNull(rsWarn!������־1) Then
        If rsWarn!������־1 = "-" Or InStr(rsWarn!������־1, str�շ����) > 0 Then byt��־ = 1
        If rsWarn!������־1 = "-" Then str������� = "" '�������ʱ,������ʾ��������
    End If
    If byt��־ = 0 And Not IsNull(rsWarn!������־2) Then
        If rsWarn!������־2 = "-" Or InStr(rsWarn!������־2, str�շ����) > 0 Then byt��־ = 2
        If rsWarn!������־2 = "-" Then str������� = "" '�������ʱ,������ʾ��������
    End If
    If byt��־ = 0 And Not IsNull(rsWarn!������־3) Then
        If rsWarn!������־3 = "-" Or InStr(rsWarn!������־3, str�շ����) > 0 Then byt��־ = 3
        If rsWarn!������־3 = "-" Then str������� = "" '�������ʱ,������ʾ��������
    End If
    If byt��־ = 0 Then Exit Function '����Ч����
    
    '������־2ʵ�����������жϢ٢�,����ֻ��һ���жϢ�
    '���ִ����ǰ����һ�����ֻ������һ�ֱ�����ʽ(������������ʱ)
    'ʾ����"-" �� ",ABC,567,DEF"
    '������־2ʾ����"-��" �� ",ABC��,567��,DEF��"
    bln�ѱ��� = InStr(str�ѱ����, str�շ����) > 0 Or str�ѱ���� Like "-*"
    
    If bln�ѱ��� Then '��intWarn = -1ʱ,Ҳ��ǿ���ٱ���
        If byt��־ = 2 Then
            If str�ѱ���� Like "-*" Then
                byt�ѱ���ʽ = IIf(Right(str�ѱ����, 1) = "��", 2, 1)
            Else
                arrTmp = Split(str�ѱ����, ",")
                For i = 0 To UBound(arrTmp)
                    If InStr(arrTmp(i), str�շ����) > 0 Then
                        byt�ѱ���ʽ = IIf(Right(arrTmp(i), 1) = "��", 2, 1)
                        'Exit For 'ȡ��˵����סԺ����ģ��
                    End If
                Next
            End If
        Else
            Exit Function
        End If
    End If
    
    If str������� <> "" Then str������� = """" & str������� & """����"
    str���� = IIf(cur������� = 0, "", "(��������:" & Format(cur�������, "0.00") & ")")
    curʣ���� = curʣ���� + cur������� - cur���ʽ��
    cur���ս�� = cur���ս�� + cur���ʽ��
        
    '---------------------------------------------------------------------
    If rsWarn!�������� = 1 Then  '�ۼƷ��ñ���(����)
        Select Case byt��־
            Case 1 '���ڱ���ֵ(����Ԥ����ľ�)��ʾѯ�ʼ���
                If curʣ���� < rsWarn!����ֵ Then
                    If Not (InStr(";" & strPrivs & ";", ";ǿ�Ƽ���;") > 0 Or bln����) Then
                        If intWarn = -1 Then
                            vMsg = frmMsgBox.ShowMsgBox(str���� & " ��ǰʣ���" & str���� & ":" & Format(curʣ����, "0.00") & ",����" & str������� & "����ֵ:" & Format(rsWarn!����ֵ, "0.00") & ",����ò��˼�����", frmParent)
                            If vMsg = vbNo Or vMsg = vbCancel Then
                                If vMsg = vbCancel Then intWarn = 0
                                BillingWarn = 2
                            ElseIf vMsg = vbYes Or vMsg = vbIgnore Then
                                If vMsg = vbIgnore Then intWarn = 1
                                BillingWarn = 1
                            End If
                        Else
                            If intWarn = 0 Then
                                BillingWarn = 2
                            ElseIf intWarn = 1 Then
                                BillingWarn = 1
                            End If
                        End If
                    Else
                        If intWarn = -1 Then
                            vMsg = frmMsgBox.ShowMsgBox(IIf(bln����, "", "ǿ�Ƽ���") & "����:" & vbCrLf & vbCrLf & str���� & " ��ǰʣ���" & str���� & ":" & Format(curʣ����, "0.00") & " ����" & str������� & "����ֵ:" & Format(rsWarn!����ֵ, "0.00") & "��", frmParent, True)
                            If vMsg = vbIgnore Then intWarn = 1
                        End If
                        BillingWarn = 4
                    End If
                End If
            Case 2 '���ڱ���ֵ��ʾѯ�ʼ���,Ԥ����ľ�ʱ��ֹ����
                If Not bln�ѱ��� Then
                    If curʣ���� < 0 Then
                        byt��ʽ = 2
                        If Not (InStr(";" & strPrivs & ";", ";ǿ�Ƽ���;") > 0 Or bln����) Then
                            If intWarn = -1 Then
                                vMsg = frmMsgBox.ShowMsgBox(str���� & " ��ǰʣ���" & str���� & "�Ѿ��ľ�," & str������� & "��ֹ���ʡ�", frmParent, True)
                                If vMsg = vbIgnore Then intWarn = 1
                            End If
                            BillingWarn = 3
                        Else
                            If intWarn = -1 Then
                                vMsg = frmMsgBox.ShowMsgBox(str������� & IIf(bln����, "", "ǿ�Ƽ���") & "����:" & vbCrLf & vbCrLf & str���� & " ��ǰʣ���" & str���� & "�Ѿ��ľ���", frmParent, True)
                                If vMsg = vbIgnore Then intWarn = 1
                            End If
                            BillingWarn = 4
                        End If
                    ElseIf curʣ���� < rsWarn!����ֵ Then
                        byt��ʽ = 1
                        If Not (InStr(";" & strPrivs & ";", ";ǿ�Ƽ���;") > 0 Or bln����) Then
                            If intWarn = -1 Then
                                vMsg = frmMsgBox.ShowMsgBox(str���� & " ��ǰʣ���" & str���� & ":" & Format(curʣ����, "0.00") & ",����" & str������� & "����ֵ:" & Format(rsWarn!����ֵ, "0.00") & ",����ò��˼�����", frmParent)
                                If vMsg = vbNo Or vMsg = vbCancel Then
                                    If vMsg = vbCancel Then intWarn = 0
                                    BillingWarn = 2
                                ElseIf vMsg = vbYes Or vMsg = vbIgnore Then
                                    If vMsg = vbIgnore Then intWarn = 1
                                    BillingWarn = 1
                                End If
                            Else
                                If intWarn = 0 Then
                                    BillingWarn = 2
                                ElseIf intWarn = 1 Then
                                    BillingWarn = 1
                                End If
                            End If
                        Else
                            If intWarn = -1 Then
                                vMsg = frmMsgBox.ShowMsgBox(IIf(bln����, "", "ǿ�Ƽ���") & "����:" & vbCrLf & vbCrLf & str���� & " ��ǰʣ���" & str���� & ":" & Format(curʣ����, "0.00") & ",����" & str������� & "����ֵ:" & Format(rsWarn!����ֵ, "0.00") & "��", frmParent, True)
                                If vMsg = vbIgnore Then intWarn = 1
                            End If
                            BillingWarn = 4
                        End If
                    End If
                Else
                    '�ϴ��ѱ�����ѡ�������ǿ�Ƽ���
                    If byt�ѱ���ʽ = 1 Then
                        '�ϴε��ڱ���ֵ��ѡ�������ǿ�Ƽ���,���ٴ�����ڵ����,������Ҫ�ж�Ԥ�����Ƿ�ľ�
                        If curʣ���� < 0 Then
                            byt��ʽ = 2
                            If Not (InStr(";" & strPrivs & ";", ";ǿ�Ƽ���;") > 0 Or bln����) Then
                                If intWarn = -1 Then
                                    vMsg = frmMsgBox.ShowMsgBox(str���� & " ��ǰʣ���" & str���� & "�Ѿ��ľ�," & str������� & "��ֹ���ʡ�", frmParent, True)
                                    If vMsg = vbIgnore Then intWarn = 1
                                End If
                                BillingWarn = 3
                            Else
                                If intWarn = -1 Then
                                    vMsg = frmMsgBox.ShowMsgBox(str������� & IIf(bln����, "", "ǿ�Ƽ���") & "����:" & vbCrLf & vbCrLf & str���� & " ��ǰʣ���" & str���� & "�Ѿ��ľ���", frmParent, True)
                                    If vMsg = vbIgnore Then intWarn = 1
                                End If
                                BillingWarn = 4
                            End If
                        End If
                    ElseIf byt�ѱ���ʽ = 2 Then
                        '�ϴ�Ԥ�����Ѿ��ľ���ǿ�Ƽ���,���ٴ���
                        Exit Function
                    End If
                End If
            Case 3 '���ڱ���ֵ��ֹ����
                If curʣ���� < rsWarn!����ֵ Then
                    If Not (InStr(";" & strPrivs & ";", ";ǿ�Ƽ���;") > 0 Or bln����) Then
                        If intWarn = -1 Then
                            vMsg = frmMsgBox.ShowMsgBox(str���� & " ��ǰʣ���" & str���� & ":" & Format(curʣ����, "0.00") & ",����" & str������� & "����ֵ:" & Format(rsWarn!����ֵ, "0.00") & ",��ֹ���ʡ�", frmParent, True)
                            If vMsg = vbIgnore Then intWarn = 1
                        End If
                        BillingWarn = 3
                    Else
                        If intWarn = -1 Then
                            vMsg = frmMsgBox.ShowMsgBox(IIf(bln����, "", "ǿ�Ƽ���") & "����:" & vbCrLf & vbCrLf & str���� & " ��ǰʣ���" & str���� & ":" & Format(curʣ����, "0.00") & ",����" & str������� & "����ֵ:" & Format(rsWarn!����ֵ, "0.00") & "��", frmParent, True)
                            If vMsg = vbIgnore Then intWarn = 1
                        End If
                        BillingWarn = 4
                    End If
                End If
        End Select
    ElseIf rsWarn!�������� = 2 Then  'ÿ�շ��ñ���(����)
        Select Case byt��־
            Case 1 '���ڱ���ֵ��ʾѯ�ʼ���
                If cur���ս�� > rsWarn!����ֵ Then
                    If Not (InStr(";" & strPrivs & ";", ";ǿ�Ƽ���;") > 0 Or bln����) Then
                        If intWarn = -1 Then
                            vMsg = frmMsgBox.ShowMsgBox(str���� & " ���շ���:" & Format(cur���ս��, gstrDec) & ",����" & str������� & "����ֵ:" & Format(rsWarn!����ֵ, "0.00") & ",����ò��˼�����", frmParent)
                            If vMsg = vbNo Or vMsg = vbCancel Then
                                If vMsg = vbCancel Then intWarn = 0
                                BillingWarn = 2
                            ElseIf vMsg = vbYes Or vMsg = vbIgnore Then
                                If vMsg = vbIgnore Then intWarn = 1
                                BillingWarn = 1
                            End If
                        Else
                            If intWarn = 0 Then
                                BillingWarn = 2
                            ElseIf intWarn = 1 Then
                                BillingWarn = 1
                            End If
                        End If
                    Else
                        If intWarn = -1 Then
                            vMsg = frmMsgBox.ShowMsgBox(IIf(bln����, "", "ǿ�Ƽ���") & "����:" & vbCrLf & vbCrLf & str���� & " ���շ���:" & Format(cur���ս��, gstrDec) & ",����" & str������� & "����ֵ:" & Format(rsWarn!����ֵ, "0.00") & "��", frmParent, True)
                            If vMsg = vbIgnore Then intWarn = 1
                        End If
                        BillingWarn = 4
                    End If
                End If
            Case 3 '���ڱ���ֵ��ֹ����
                If cur���ս�� > rsWarn!����ֵ Then
                    If Not (InStr(";" & strPrivs & ";", ";ǿ�Ƽ���;") > 0 Or bln����) Then
                        If intWarn = -1 Then
                            vMsg = frmMsgBox.ShowMsgBox(str���� & " ���շ���:" & Format(cur���ս��, gstrDec) & ",����" & str������� & "����ֵ:" & Format(rsWarn!����ֵ, "0.00") & ",��ֹ���ʡ�", frmParent, True)
                            If vMsg = vbIgnore Then intWarn = 1
                        End If
                        BillingWarn = 3
                    Else
                        If intWarn = -1 Then
                            vMsg = frmMsgBox.ShowMsgBox(IIf(bln����, "", "ǿ�Ƽ���") & "����:" & vbCrLf & vbCrLf & str���� & " ���շ���:" & Format(cur���ս��, gstrDec) & ",����" & str������� & "����ֵ:" & Format(rsWarn!����ֵ, "0.00") & "��", frmParent, True)
                            If vMsg = vbIgnore Then intWarn = 1
                        End If
                        BillingWarn = 4
                    End If
                End If
        End Select
    End If
    
    '���ڼ�����Ĳ���,�����ѱ������
    If BillingWarn = 1 Or BillingWarn = 4 Then
        If byt��־ = 1 Then
            If rsWarn!������־1 = "-" Then
                str�ѱ���� = "-"
            Else
                str�ѱ���� = str�ѱ���� & "," & rsWarn!������־1
            End If
        ElseIf byt��־ = 2 Then
            If rsWarn!������־2 = "-" Then
                str�ѱ���� = "-"
            Else
                str�ѱ���� = str�ѱ���� & "," & rsWarn!������־2
            End If
            '���ӱ�ע���ж��ѱ����ľ��巽ʽ
            str�ѱ���� = str�ѱ���� & IIf(byt��ʽ = 2, "��", "��")
        ElseIf byt��־ = 3 Then
            If rsWarn!������־3 = "-" Then
                str�ѱ���� = "-"
            Else
                str�ѱ���� = str�ѱ���� & "," & rsWarn!������־3
            End If
        End If
    End If
End Function

Public Function FinishBillingWarn(ByVal frmParent As Object, ByVal strPrivs As String, ByVal lng����ID As Long, _
    ByVal lng��ҳID As Long, ByVal lng����ID As Long, ByVal cur��� As Currency, ByVal str��� As String, ByVal str����� As String) As Boolean
'���ܣ���ִ��������Զ���˵ķ���ʱ���Բ��˷��ý��м��ʱ�����
'������str���="CDE..."����������漰�����շ����
'      str�����="���,����,..."����Ӧ�������������ʾ
    Dim rsPati As ADODB.Recordset
    Dim rsWarn As ADODB.Recordset
    Dim strWarn As String, intWarn As Integer
    Dim strSql As String, intR As Integer, i As Long
    Dim cur���� As Currency
    
    On Error GoTo errH
    
    If lng��ҳID <> 0 Then
        'סԺ���˱���
        strSql = _
            " Select ����ID,Ԥ�����,�������,0 as Ԥ����� From ������� Where ����=1 And ����ID=[1]" & _
            " Union ALL" & _
            " Select A.����ID,0,0,Sum(���) From ����ģ����� A,������ҳ B" & _
            " Where A.����ID=B.����ID And A.��ҳID=B.��ҳID And B.���� Is Not Null And A.����ID=[1] And A.��ҳID=[2] Group by A.����ID"
        strSql = "Select ����ID,Nvl(Sum(Ԥ�����),0)-Nvl(Sum(�������),0)+Nvl(Sum(Ԥ�����),0) as ʣ��� From (" & strSql & ") Group by ����ID"
        
        strSql = "Select A.����,zl_PatiWarnScheme(A.����ID,B.��ҳID) as ���ò���,C.ʣ���," & _
            " Decode(A.������,Null,Null,zl_PatientSurety(A.����ID,B.��ҳID)) as ������" & _
            " From ������Ϣ A,������ҳ B,(" & strSql & ") C" & _
            " Where A.����ID=B.����ID And A.����ID=C.����ID(+)" & _
            " And A.����ID=[1] And B.��ҳID=[2]"
        Set rsPati = zlDatabase.OpenSQLRecord(strSql, "FinishBillingWarn", lng����ID, lng��ҳID)
    Else
        '���������ﱨ��
        strSql = "Select ����ID,Ԥ�����,������� From ������� Where ����=1 And ����ID=[1]"
        strSql = "Select A.����,zl_PatiWarnScheme(A.����ID) as ���ò���,A.������," & _
            " Nvl(B.Ԥ�����,0)-Nvl(B.�������,0)+Nvl(E.�ʻ����,0) as ʣ���" & _
            " From ������Ϣ A,(" & strSql & ") B,ҽ�����˹����� D,ҽ�����˵��� E" & _
            " Where A.����ID=B.����ID(+) And A.����id = D.����id(+) And A.����=D.����(+)" & _
            " And D.����=E.����(+) And D.����=E.����(+) And D.ҽ����=E.ҽ����(+) And D.��־(+)=1" & _
            " And A.����ID=[1]"
        Set rsPati = zlDatabase.OpenSQLRecord(strSql, "FinishBillingWarn", lng����ID)
    End If
    
    intWarn = -1 '���ʱ���ʱȱʡҪ��ʾ
    'ִ�б���:���ﲡ�˲���ID=0
    strSql = "Select Nvl(��������,1) as ��������,����ֵ,������־1,������־2,������־3 From ���ʱ����� Where Nvl(����ID,0)=[1] And ���ò���=[2]"
    Set rsWarn = zlDatabase.OpenSQLRecord(strSql, "FinishBillingWarn", lng����ID, CStr(NVL(rsPati!���ò���)))
    If Not rsWarn.EOF Then
        If rsWarn!�������� = 2 Then cur���� = GetPatiDayMoney(lng����ID)
        str����� = Mid(str�����, 2)
        For i = 1 To Len(str���)
            intR = BillingWarn(frmParent, strPrivs, rsWarn, NVL(rsPati!����), NVL(rsPati!ʣ���, 0), cur����, cur���, NVL(rsPati!������, 0), Mid(str���, i, 1), Split(str�����, ",")(i - 1), strWarn, intWarn)
            If InStr(",2,3,", intR) > 0 Then Exit Function
        Next
    End If
    
    FinishBillingWarn = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function ItemHaveCash(ByVal int������Դ As Integer, ByVal bln����ִ�� As Boolean, ByVal lngҽ��ID As Long, ByVal lng���ID As Long, _
    ByVal lng���ͺ� As Long, ByVal str��� As String, ByVal str���ݺ� As String, ByVal int��¼���� As Integer, ByVal int������� As Integer, ByVal int��ʽ As Integer, _
    Optional ByVal blnMove As Boolean, Optional ByVal dat����ʱ�� As Date, Optional ByRef strҽ��IDs As String, Optional ByRef strNOs As String, Optional ByRef blnIsAbnormal As Boolean) As Boolean
'���ܣ��жϵ�ǰ��ִ��ҽ���Ƿ����շѻ���ʻ��۵��Ƿ������
'������ int������Դ -- 1-����,2-סԺ
'       bln����ִ�� -- True = ����ִ�У�False = ����ҽ��ȫ��ִ��
'       lngҽ��ID -- ҽ��ID
'       lng���ID -- ���ID
'       lng���ͺ� -- ���ͺ�
'       str���=����������ڴ�һ��ҽ�������ַֿ�ִ�е�����
'       str���ݺ� -- ���ݺ�
'       int��¼���� -- ��¼����
'       int������� -- ������ʣ�1=סԺ���͵��������
'       int��ʽ --  0-����Ƿ����δ�շѼ�¼
'                   1-����Ƿ�������շѼ�¼
'       blnMove -- �Ƿ�ת��
'       dat����ʱ�� -- ����ʱ��
'       strҽ��IDs -- �����ز���������ҽ������ص�ҽ��ID
'       strNOs -- �����ز�������ҽ�����͵ĵ��ݺźͲ��ĸ����еĵ��ݺ�
'       blnIsAbnormal -- �����ز��������Ƿ��쳣����
'      ���أ�True--���շѣ�False--δ�շѣ�
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String, strTab As String
    
    If int������Դ = 2 And int��¼���� = 2 And int������� = 0 Then
        strTab = "סԺ���ü�¼"
    Else
        strTab = "������ü�¼"
    End If
    
    ItemHaveCash = True
    strҽ��IDs = ""
    strNOs = ""
    
    '��Ӧ�ķ������Ƿ����δ�շ�[��������]������
    '���嵥ֻ��ʾ���շѲ�ͬ��
    '1.�����ҽ������(���Ӽ�¼���ʵ���������Ϊ���ܲ��շѵ�����ʵ�)
    '2.���ʻ���Ҳ��ʾΪδ��(�嵥��Ҫ���Գ���ִ�к����)
    '3.��NO��Ӧ�����ҽ���ķ��ü��(�嵥�ǰ���ʾ��ҽ��ID)
    strSql = _
        " Select A.��¼״̬,Nvl(B.���ID,B.ID) as ҽ��ID,B.�������,A.ִ��״̬,A.NO" & IIf(strTab = "סԺ���ü�¼", ",0 as ����״̬", ",NVL(A.����״̬,0) as ����״̬") & _
        " From " & strTab & " A,����ҽ����¼ B" & _
        " Where A.NO=[4] And A.��¼״̬ IN(0,1,3) And A.ҽ�����+0=B.ID And MOD(A.��¼����,10)=[5] " & IIf(bln����ִ��, " And B.ID=[2]", "") & _
        " Union ALL " & _
        " Select B.��¼״̬,Nvl(C.���ID,C.ID) as ҽ��ID,C.�������,B.ִ��״̬,A.NO" & IIf(strTab = "סԺ���ü�¼", ",0 as ����״̬", ",NVL(b.����״̬,0) as ����״̬") & _
        " From ����ҽ����¼ C," & strTab & " B,����ҽ������ A" & _
        " Where A.NO=B.NO And A.��¼����=MOD(B.��¼����,10) And A.ҽ��ID=B.ҽ�����+0" & IIf(bln����ִ��, " And A.ҽ��ID=[2]", _
            " And A.ҽ��ID IN (Select ID From ����ҽ����¼ Where (ID=[1] Or ���ID=[1]) And �������=[6])") & _
        " And A.���ͺ�=[3] And B.��¼״̬ IN(0,1,3) And A.ҽ��ID=C.ID And A.��¼����=[5] "
    If blnMove Then
        strSql = Replace(strSql, "����ҽ����¼", "H����ҽ����¼")
        strSql = Replace(strSql, "����ҽ������", "H����ҽ������")
        strSql = Replace(strSql, strTab, "H" & strTab)
    ElseIf zlDatabase.DateMoved(dat����ʱ��) Then
        strSql = strSql & " Union ALL " & Replace(strSql, strTab, "H" & strTab)
    End If
    
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "ItemHaveCash", IIf(lng���ID <> 0, lng���ID, lngҽ��ID), lngҽ��ID, lng���ͺ�, str���ݺ�, int��¼����, str���)
    If Not rsTmp.EOF Then
        If int��ʽ = 0 Then
            rsTmp.Filter = "ҽ��ID=" & IIf(lng���ID <> 0, lng���ID, lngҽ��ID) & " And �������='" & str��� & "' And ����״̬=1"
            If Not rsTmp.EOF Then
                blnIsAbnormal = True
                ItemHaveCash = False
            Else
                rsTmp.Filter = "ҽ��ID=" & IIf(lng���ID <> 0, lng���ID, lngҽ��ID) & " And �������='" & str��� & "' And ��¼״̬=0"
                If Not rsTmp.EOF Then ItemHaveCash = False
            End If
            
            While Not rsTmp.EOF
                If InStr("," & strҽ��IDs & ",", "," & rsTmp!ҽ��ID & ",") = 0 Then
                    strҽ��IDs = strҽ��IDs & "," & rsTmp!ҽ��ID
                End If
                If InStr("," & strNOs & ",", "," & rsTmp!NO & ",") = 0 Then
                    strNOs = strNOs & "," & rsTmp!NO
                End If
                rsTmp.MoveNext
            Wend
                
            strNOs = Mid(strNOs, 2)
            strҽ��IDs = Mid(strҽ��IDs, 2)
        ElseIf int��ʽ = 1 Then
            rsTmp.Filter = "ҽ��ID=" & IIf(lng���ID <> 0, lng���ID, lngҽ��ID) & " And �������='" & str��� & "' And ��¼״̬<>1 And ����״̬<>1"
            If Not rsTmp.EOF Then ItemHaveCash = False
        End If
    ElseIf int��ʽ = 1 Then
        ItemHaveCash = False
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetAdviceMoney(ByVal lngAdviceID As Long, ByVal lng��Դ As Long, str��� As String, str����� As String) As Currency
'���ܣ�����ָ����ҽ��ID����ȡҽ����Ӧδ��˵ļ��ʷ��úϼ�
'������lngAdviceID,strSendNo
'���أ�str���,str�����=���ڱ�����ʾ

    Dim rsTmp As New ADODB.Recordset
    Dim curMoney As Currency
    Dim strFeeTable As String
    Dim strFilter As String
    
    str��� = "": str����� = ""
    
    On Error GoTo errH
    
    '��Ҫ����ϵͳ�����жϣ�81�Ų�����"ִ�к��Զ���˻��۵�"
    If gblnִ�к���� = False Then Exit Function
    
    'סԺ���˲�סԺ���ü�¼���������Ȳ��˲�������ü�¼
    If lng��Դ = 2 Then
        strFeeTable = "סԺ���ü�¼"
        strFilter = " A.��¼����"
    Else
        strFeeTable = "������ü�¼"
        strFilter = " decode(A.��¼����,1,1,11,1,A.��¼����)"
    End If
    
    gstrSQL = "Select /*+ RULE */" & vbNewLine & _
                " B.����, B.����, Sum(A.ʵ�ս��) As ���" & vbNewLine & _
                "From " & strFeeTable & " A, �շ���Ŀ��� B" & vbNewLine & _
                "Where A.ҽ����� + 0 In (Select ID From ����ҽ����¼ Where ID = [1] Or ���id = [1]) And" & vbNewLine & _
                "      (" & strFilter & ", A.NO) In" & vbNewLine & _
                "      ( Select ��¼����, NO" & vbNewLine & _
                "        From ����ҽ������" & vbNewLine & _
                "        Where ҽ��id In (Select ID From ����ҽ����¼ Where ID = [1] Or ���id = [1])" & vbNewLine & _
                "        Union All" & vbNewLine & _
                "        Select ��¼����, NO" & vbNewLine & _
                "        From ����ҽ������" & vbNewLine & _
                "        Where ҽ��id In (Select ID From ����ҽ����¼ Where ID = [1] Or ���id = [1] )" & vbNewLine & _
                "       ) And A.���ʷ��� = 1 And A.��¼״̬ = 0 And A.�շ���� = B.���� " & vbNewLine & _
                "Group By B.����, B.����"
                
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "GetAdviceMoney", lngAdviceID)
    
    curMoney = 0
    Do While Not rsTmp.EOF
        curMoney = curMoney + NVL(rsTmp!���, 0)
        str��� = str��� & rsTmp!����
        str����� = str����� & "," & rsTmp!����
        rsTmp.MoveNext
    Loop
    
    str����� = Mid(str�����, 2)
    GetAdviceMoney = curMoney
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetPinyinName(ByVal strPatiName As String, ByVal intCapital As Integer, ByVal blnUseSplitter As Boolean) As String
'���ܣ����������ֵ��е����ã���ȡ������Ӧ��ƴ����
'strPatiName:��������
    Dim strTempName As String
    Dim strSql As String
    Dim rsReccord As ADODB.Recordset
    
On Error GoTo errHandle
    If strPatiName = "" Then Exit Function
    
    If blnUseSplitter Then
        strSql = "select Zlpinyincode([1],[2],[3],[4]) as ƴ���� from dual"
        Set rsReccord = zlDatabase.OpenSQLRecord(strSql, "��ȡƴ��", strPatiName, 1, 1, " ")
    Else
        strSql = "select Zlpinyincode([1],[2],[3]) as ƴ���� from dual"
        Set rsReccord = zlDatabase.OpenSQLRecord(strSql, "��ȡƴ��", strPatiName, 1, 1)
    End If
    
    If rsReccord.RecordCount > 0 Then
        strTempName = NVL(rsReccord!ƴ����)
    End If
    
    If intCapital = 0 Then
        GetPinyinName = UCase(Trim(strTempName))
    ElseIf intCapital = 1 Then
        GetPinyinName = LCase(Trim(strTempName))
    Else
        GetPinyinName = Trim(strTempName)
    End If
    
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
End Function

Public Function GetPatiDayMoney(lng����ID As Long) As Currency
'���ܣ���ȡָ�����˵��췢���ķ����ܶ�
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String
    
    On Error GoTo errH
    
    strSql = "Select zl_PatiDayCharge([1]) as ��� From Dual"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "mdlCISKernel", lng����ID)
    If Not rsTmp.EOF Then GetPatiDayMoney = NVL(rsTmp!���, 0)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function funcConnectShardDir(frmParent As Form, strShareRemoteDir As String, strUserName As String, _
    strPassWord As String) As Long
'------------------------------------------------
'���ܣ�����������Դ
'������ frmParent  -- ������
'       strShareRemoteDir -- ����Ŀ¼
'       strUserName -- ����Ŀ¼�û���
'       strPassWord -- ����Ŀ¼����
'���أ��ޣ����ӹ���Ŀ¼
'------------------------------------------------
    
    Dim NetR As NETRESOURCE
    Dim lngResult As Long
    
    NetR.dwType = RESOURCETYPE_ANY
    NetR.lpLocalName = vbNullString
    NetR.lpRemoteName = strShareRemoteDir
    NetR.lpProvider = vbNullString
    lngResult = WNetAddConnection2(NetR, strPassWord, strUserName, 0)
    
    If lngResult <> 0 Then
        MsgBoxD frmParent, "��������ʧ�ܣ��������������Ƿ���ȷ��"
    End If
    funcConnectShardDir = lngResult
End Function

Public Function bln����δ��˳�Ժ(ByVal lng����ID As Long, ByVal lng��ҳID As Long, ByVal lngҽ��ID As Long, ByVal lng��Դ As Long) As Boolean
'------------------------------------------------
'���ܣ��жϲ����Ƿ��ѳ�Ժ�����м��˷���δ���
'������ lng����ID  -- ����ID
'       lng��ҳID -- ��ҳID
'       lngҽ��ID -- ҽ��ID
'       lng��Դ -- ������Դ
'���أ�True -- �����ѳ�Ժ����δ��˷��ã�False --����
'------------------------------------------------
'��Ҫ����ϵͳ�����жϣ�81�Ų�����"ִ�к��Զ���˻��۵�"
    
    bln����δ��˳�Ժ = False
    
    If gblnִ�к���� = True Then
        If Not bln������Ժ(lng����ID, lng��ҳID) And bln����δ�󻮼۵�(lngҽ��ID, lng��Դ) Then
            bln����δ��˳�Ժ = True
        End If
    End If
End Function

Public Function blnδ�ɷ���(lngOrderID As Long) As Boolean
'------------------------------------------------
'���ܣ��жϼ��ҽ���Ƿ���δ���ķ���
'������ lngOrderID  -- ҽ��ID
'���أ�True -- δ�շѣ�False --���շ�
'------------------------------------------------
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    Dim intSourceType As Integer
    Dim lngSendNO As Long
    Dim str������� As String
    Dim str���ݺ� As String
    Dim int��¼���� As Integer
    Dim int�������  As Integer
    
    On Error GoTo err
    
    blnδ�ɷ��� = False
    
    strSql = "Select A.��¼����,A.�������,A.���ͺ�,A.NO,B.�������,B.������Դ from ����ҽ������ A,����ҽ����¼ B  where A.ҽ��ID=B.ID and  B.ID =[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "blnδ�ɷ���", lngOrderID)
    If rsTemp.EOF = False Then
        int��¼���� = NVL(rsTemp!��¼����, 0)
        int������� = NVL(rsTemp!�������, 0)
        str������� = NVL(rsTemp!�������)
        lngSendNO = rsTemp!���ͺ�
        str���ݺ� = NVL(rsTemp!NO)
        intSourceType = NVL(rsTemp!������Դ)
        
        blnδ�ɷ��� = Not ItemHaveCash(intSourceType, False, lngOrderID, 0, lngSendNO, str�������, str���ݺ�, int��¼����, int�������, 0)
    End If
    
    Exit Function
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function AssembleImage(AssembleViewer As DicomImages, ByVal intRows As Integer, ByVal intCols As Integer, _
    ByVal lngHeight As Long, ByVal lngWidth As Long) As DicomImage

'���viewer�е���ʾ������ͼ���һ��ͼ��

    Dim Image As New DicomImage '��ͼ��
    Dim imgs As New DicomImages '��ʱ�洢��Ļ�ɼ���ͼ��
    Dim intWidth As Integer     '��ͼ��Ŀ��
    Dim intHeight As Integer    '��ͼ��ĸ߶�
    Dim Simg As New DicomImage
    Dim sZoom As Single
    Dim intImgRectWidth As Integer  '����ͼ���ռ�õ�������
    Dim intImgRectHeight As Integer '����ͼ���ռ�õ�����߶�
    Dim i As Integer
    Dim intMaxWidth As Integer      'ƴ�Ӻ�ͼ��������
    Dim intMaxHeight As Integer     'ƴ�Ӻ�ͼ������߶�
    Dim intBorder As Integer        'ͼ��֮��ı߾�
    Dim intOffsetX As Integer       'ƴ��ʱX�����λ��
    Dim intOffsetY As Integer       'ƴ��ʱY�����λ��
    Dim lngWhiteX As Long           '��ͼ���ɫ�ĳɰ�ɫ��X���
    Dim lngWhiteY As Long           '��ͼ���ɫ�ĳɰ�ɫ��Y�߶�
    Dim iPlane As Integer
    Dim dblZoom As Double
    Dim iMax As Integer
    
    If AssembleViewer.Count <= 0 Then
        '����һ����ͼ**************
        Exit Function
    End If

    On Error GoTo err
    '������ͼ��Ŀ�Ⱥ͸߶�

    '��ͼ��Ŀ�Ⱥ͸߶Ȳ��ܹ�����intMaxWidth��intMaxHeight����ȡ��߶ȣ�
    intMaxWidth = 3073
    intMaxHeight = 3073
    intBorder = 10

    intImgRectWidth = 0
    intImgRectHeight = 0

    '������ͼ��Ŀ�Ⱥ͸߶�

    'ʹ��ԭͼ��Ŀ�Ⱥ͸߶Ⱥͣ�����Viewer�ı�����������

    '����ͼ����¿��
    For i = 1 To AssembleViewer.Count
        If intImgRectWidth < AssembleViewer(i).SizeX Then intImgRectWidth = AssembleViewer(i).SizeX
        If intImgRectHeight < AssembleViewer(i).SizeY Then intImgRectHeight = AssembleViewer(i).SizeY
    Next i
    
    
    If AssembleViewer.Count = 1 Then
        If intImgRectWidth >= intMaxWidth Or intImgRectHeight >= intMaxHeight Then
            iMax = intMaxWidth
        Else
            iMax = IIf(intImgRectWidth > intImgRectHeight, intImgRectWidth, intImgRectHeight)
        End If
    ElseIf AssembleViewer.Count <= 4 Then
        If intImgRectWidth >= 2048 Or intImgRectHeight >= 2048 Then
            iMax = IIf(intImgRectWidth > intImgRectHeight, intImgRectWidth, intImgRectHeight) / 2
        Else
            iMax = IIf(intImgRectWidth > intImgRectHeight, intImgRectWidth, intImgRectHeight) / 1.5
        End If
    Else
        iMax = IIf(intImgRectWidth > intImgRectHeight, intImgRectWidth, intImgRectHeight) / 3
    End If
    
    If iMax < 512 Then iMax = 512
    
    '������������ͼ������
    intWidth = intImgRectWidth * intCols
    intHeight = intImgRectHeight * intRows
    
    '����ͼ��Ŀ�ߣ����ܴ������ֵ
    '�������intMaxWidth��intMaxHeight�򣬰���ͼ���ܳ���ȣ�ʹ��С�ڵ���intMaxWidth��intMaxHeight��Ϊ�¿��,
    If intWidth > intMaxWidth Or intHeight > intMaxHeight Then
        If intHeight / intWidth > intMaxHeight / intMaxWidth Then
            intWidth = intWidth / intHeight * intMaxHeight
            intHeight = intMaxHeight
        Else
            intHeight = intHeight / intWidth * intMaxWidth
            intWidth = intMaxWidth
        End If
    End If
    
    '�ɼ�ͼ��
    '��ͼ��ɼ�����ʱͼ��
    For i = 1 To AssembleViewer.Count
        '�������ű��� hj�޸�,�����ͼ�ϲ�ʱ���Ŵ��ͼ���޷������Ŵ������
        sZoom = intImgRectHeight / AssembleViewer(i).SizeY
        If sZoom > intImgRectWidth / AssembleViewer(i).SizeX Then
            sZoom = intImgRectWidth / AssembleViewer(i).SizeX
        End If
        
        AssembleViewer(i).StretchToFit = False
        AssembleViewer(i).Zoom = sZoom
        '�ɼ�ͼ��
        Set Simg = AssembleViewer(i).PrinterImage(8, 3, True, sZoom, 0, AssembleViewer(i).SizeX, 0, AssembleViewer(i).SizeY)
        imgs.Add Simg
    Next i

    '��ȷ������ͼ��Ŀ�Ⱥ͸߶�
    intImgRectWidth = 0
    intImgRectHeight = 0
    
    '����ͼ��ķֱ�����500*500֮��
    Dim imgsTmp As New DicomImages
    For i = 1 To imgs.Count
        
        iPlane = 1
        If Not IsNull(imgs(i).Attributes(&H28, &H4).value) And imgs(i).Attributes(&H28, &H4).Exists Then
            If imgs(i).Attributes(&H28, &H4).value = "RGB" Then
                iPlane = 3
            End If
        End If
        
        '����imax�������ű���
        If imgs(i).SizeX > iMax Or imgs(i).SizeY > iMax Then
            dblZoom = iMax / imgs(i).SizeX
            If dblZoom > iMax / imgs(i).SizeY Then dblZoom = iMax / imgs(i).SizeY
        Else
            dblZoom = 1
        End If
        
        imgsTmp.Add imgs(i).PrinterImage(8, iPlane, True, dblZoom, 0, imgs(i).SizeX, 0, imgs(i).SizeY)
    Next
    
    Set imgs = imgsTmp

    For i = 1 To imgs.Count
        If intImgRectWidth < imgs(i).SizeX Then intImgRectWidth = imgs(i).SizeX
        If intImgRectHeight < imgs(i).SizeY Then intImgRectHeight = imgs(i).SizeY
        imgs(i).Attributes.Add &H8, &H16, "doSOP_SecondaryCapture"
    Next i
    intImgRectWidth = intImgRectWidth + intBorder
    intImgRectHeight = intImgRectHeight + intBorder
    intWidth = intImgRectWidth * intCols
    intHeight = intImgRectHeight * intRows

    '������ͼ��
    Image.Name = "print"
    Image.PatientID = "print001"
    
    Image.Attributes.Add &H8, &H16, doSOP_SecondaryCapture
    Image.Attributes.Add &H28, &H2, 3 ' samples/pixel
    Image.Attributes.Add &H28, &H4, "RGB" ' photometric interpreation  'CT����MONOCHROME2,CR����MONOCHROME1��
    Image.Attributes.Add &H28, &H10, intHeight  'x,Rows
    Image.Attributes.Add &H28, &H11, intWidth 'Y,Columns
    Image.Attributes.Add &H28, &H100, 8 'bits allocated
    Image.Attributes.Add &H28, &H101, 8 ' bits stored
    Image.Attributes.Add &H28, &H102, 7 ' high bit
    ReDim pix(intWidth * 3, intHeight * 3) As Byte
    For lngWhiteX = 0 To intWidth * 3
        For lngWhiteY = 0 To intHeight * 3
            pix(lngWhiteX, lngWhiteY) = 255
        Next lngWhiteY
    Next lngWhiteX
    Image.Attributes.Add &H7FE0, &H10, pix

    'ƴ����ͼ��
    For i = 1 To imgs.Count
        '����ͼ����λ��
        intOffsetX = (intImgRectWidth - imgs(i).SizeX - intBorder) / 2
        intOffsetY = (intImgRectHeight - imgs(i).SizeY - intBorder) / 2
        Image.Blt imgs(i), 0, 0, ((i - 1) Mod intCols) * intImgRectWidth + intOffsetX, ((i - 1) \ intCols) * intImgRectHeight + intOffsetY, imgs(i).SizeX, imgs(i).SizeY, 1, 1, 1, False
    Next i

    Set AssembleImage = Image
    Exit Function
err:
End Function

Public Function FunLogIn(frmParent As Form, str���� As String) As String
'���ܣ��Գ������ע�ᣬ���ע��ɹ����򷵻�ע��ʱ��
'������ frmParent ---������
'       str���� ---'��ע������ʹ�õ���������
'����ֵ��ע��ɹ�ע�����ڣ�ע��ʧ�ܷ��ؿ�

    Dim intNum As Integer
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    Dim strIP��ַ As String         '��Ҫע���IP��ַ
    
    On Error GoTo err
    
    strIP��ַ = OS.IP
    
    '��ע��������ȡ��Ȩ��������-1--�����ƣ�0--��ֹ��X��X>0��--������������
    intNum = gint��Ƶ�豸����
    
    'intNUM >0 ,����ù���ע�����
    If intNum > 0 Then  '����������
        strSql = "Zl_Ӱ�������¼_Update('" & strIP��ַ & "','" & str���� & "'," & intNum & ")"
        zlDatabase.ExecuteProcedure strSql, "ע��" & str����
        '���ע���Ƿ�ɹ�
        strSql = "Select ����ʱ��,IP��ַ from Ӱ�������¼ where  ����=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "��ȡ����ʱ��", str����)
        
        If rsTemp.RecordCount <= intNum Then
            rsTemp.Filter = "IP��ַ='" & strIP��ַ & "'"
            If rsTemp.RecordCount = 1 Then  'ע��ɹ�
                FunLogIn = rsTemp!����ʱ��
                Exit Function
            End If
        End If
    ElseIf intNum = -1 Then     '������
        FunLogIn = Now
        Exit Function
    Else    '=0����������ֵ����ֹ�������κδ�����������ʾ
    
    End If
    
    'ע��ʧ�ܣ�����������ԭ��
    '1��ע���������������ɵ��������޷�ע��IP��ַ
    '2��ֱ��ͨ��SQL����������IP��ַ�����±��еļ�¼��������������ɵ�����
    Call MsgBoxD(frmParent, "�򿪵�" & str���� & "�������������������" & intNum & "�������������Ӧ����ϵ��", vbOKOnly, gstrSysName)
    FunLogIn = ""
    
    Exit Function
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function FunCheckRegInfo(frmParent As Form) As Boolean
'���ܣ�����Ƿ����ע���ip��ַ����������ƵԴ
On Error GoTo errHandle
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    Dim strIP��ַ As String         '��Ҫע���IP��ַ
    
    FunCheckRegInfo = False
    
    strIP��ַ = OS.IP
    
    strSql = "select ����վ from zltools.zlclients where ip=[1] and ������ƵԴ=1"
    
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "��ȡע����Ϣ", strIP��ַ)
    
    If rsTemp.EOF = False Then FunCheckRegInfo = True
    
Exit Function
errHandle:
End Function

Public Function FunCheckIp(frmParent As Form, str���� As String) As Boolean
'���ܣ�����Ƿ����ע���ip��ַ
On Error GoTo errHandle
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    Dim strIP��ַ As String         '��Ҫע���IP��ַ
    
    FunCheckIp = False
    
    strIP��ַ = OS.IP
    
    strSql = "Select ����ʱ�� from Ӱ�������¼ where ����=[2] and IP��ַ=[1] "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "��ȡ����ʱ��", strIP��ַ, str����)
    
    If rsTemp.EOF = False Then FunCheckIp = True
    
Exit Function
errHandle:
End Function


Public Function FunLogOut(frmParent As Form, str���� As String, str����ʱ�� As String) As Boolean
'���ܣ��˳������ʱ�򣬼������Ƿ�Ϸ�ע�������������ͨ�����������ֶζ�ʱɾ����Ӱ�������¼�����еļ�¼��
'������ frmParent ---������
'       str���� ---'��ע������ʹ�õ���������
'       str����ʱ�� --- ע�Ṥ��վʱ���ص�ʱ��
'����ֵ���Ϸ�ע��True���Ƿ�������False
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    Dim strIP��ַ As String         '��Ҫע���IP��ַ
    Dim intNum As Integer
    
    On Error GoTo err
    strIP��ַ = OS.IP
    
    '����ʱ��Ϊ�գ���ʾע��ʧ�ܣ�û����������������˳���ʱ���ټ�����ݿ�
    If str����ʱ�� = "" Then
        FunLogOut = True
        Exit Function
    End If
    
    '��ע��������ȡ��Ȩ��������-1--�����ƣ�0--��ֹ��X��X>0��--������������
    intNum = gint��Ƶ�豸����
    
    If intNum > 0 Then '������������
        strSql = "Select ����ʱ�� from Ӱ�������¼ where IP��ַ=[1] and ����=[2]"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "��ȡ����ʱ��", strIP��ַ, str����)
        If rsTemp.EOF = False Then
            FunLogOut = True
        Else
            '�Ա�����ʱ������ݿ��ʱ�䣬�������ͬһ�죬˵����ǰһ�쿪�������ע����Ϣ��ɾ���ˣ�
            '���������Ϊ�ǺϷ�ע��
            strSql = "Select sysdate from dual"
            Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "��ȡ���ݿ�ʱ��")
            If Format(rsTemp!sysdate, "yyyy-mm-dd") <> Format(str����ʱ��, "yyyy-mm-dd") Then
                FunLogOut = True
            Else
                FunLogOut = False
            End If
        End If
    ElseIf intNum = -1 Then     '������
        FunLogOut = True
    Else    '=0����������ֵ����ֹ
    
    End If
    If FunLogOut = False Then
        Call MsgBoxD(frmParent, "�򿪵�" & str���� & "�������������������" & intNum & "�������������Ӧ����ϵ��", vbOKOnly, gstrSysName)
    End If
    Exit Function
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function getLicenseCount(strLicenseName As String) As Integer
'��ȡ��Ȩ������
'������ strLicenseName --- ��Ȩ����
    Dim strLiceseCount As String
    
    On Error GoTo err
    
    strLiceseCount = zl9comlib.zlRegInfo(strLicenseName)
    If strLiceseCount = "" Then '������
        getLicenseCount = -1
    ElseIf Val(strLiceseCount) > 0 Then '������������
        getLicenseCount = Val(strLiceseCount)
    Else '��ֹ
        getLicenseCount = 0
    End If
    
    Exit Function
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function getStudyStateRich(ByVal lngOrderID As Long, ByVal strDocID As String, Optional ByVal blnIsCancelComplete As Boolean = False, Optional ByRef blnAllReportFinished As Boolean = False, _
                            Optional ByRef lngSendNO As Long, Optional ByRef bln���������� As Boolean, Optional ByRef blnCriticalValues As Boolean, _
                            Optional ByRef blnImageQuality As Boolean, Optional ByRef blnReportQuality As Boolean, Optional ByRef blnConformDetermine As Boolean) As Integer
'------------------------------------------------
'���ܣ���鱨���ǩ�������ȷ����������еĳ̶ȡ�
'������ lngOrderID -- ҽ��ID
'       strDocId -- ����ID
'       blnIsCancelComplete -- ��ѡ���Ƿ���ִ�е�ȡ����ɲ���
'       blnAllReportFinished -- ��ѡ�����ز���
'       lngSendNO -- ��ѡ�����ز���
'       bln���������� -- ��ѡ�����ز���
'       blnCriticalValues -- ��ѡ�����ز���
'       blnImageQuality -- ��ѡ�����ز���
'       blnReportQuality -- ��ѡ�����ز���
'       blnConformDetermine -- ��ѡ�����ز���
'���أ�1--�ѵǼǣ�2--�ѱ�����3--�Ѽ�飻4--�ѱ��棻5--����ˣ�6--����ɣ�-1--�Ѳ���
'------------------------------------------------

    Dim i As Integer
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    Dim intFinishReport As Integer
    Dim intReportCount As Integer
    Dim blnIsReject As Boolean

    On Error GoTo err
    
    strSql = "Select c.ִ�й���,d.ҽ��id As Ӱ��ҽ��ID,e.ҽ��id As ����ҽ��ID,c.���ͺ�,d.���uid,d.Ӱ������,d.�������,d.Σ��״̬,e.��������," & _
             "e.id,e.������, e.���༭�� As ������,e.���༭ʱ�� As ���ʱ��,e.��������, e.�������, e.����״̬,e.id as ����ID " & _
             "From ����ҽ������ c, Ӱ�����¼ d, Ӱ�񱨸��¼ e " & _
             "Where e.ҽ��id = [1] And d.ҽ��id(+) = c.ҽ��id And e.ҽ��id(+) = c.ҽ��id"
    
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "��ȡ�Ƿ�ǩ��", lngOrderID)
    
    '�����ѯû�н������û�б���
    If rsTemp.EOF = True Then
        strSql = "Select a.���uid, a.���ͺ�, b.ִ�й��� From Ӱ�����¼ a, ����ҽ������ b Where a.ҽ��ID = [1] and a.ҽ��ID = b.ҽ��ID"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "��ȡ�Ƿ�ǩ��", lngOrderID)
        
        If Not rsTemp.EOF Then
            If NVL(rsTemp!���UID) = "" Then
                getStudyStateRich = 2
            Else
                getStudyStateRich = 3
            End If
            
            lngSendNO = NVL(rsTemp!���ͺ�, 0)
            
            '����������Ҫ����תԺ�Ĳ��ˣ�������������ɱ���������������飬��ʱִ�й���Ϊ6�Կ��ܴ���δ��ɵı���
            If NVL(rsTemp!ִ�й���, 0) = 6 And Not blnIsCancelComplete Then
                getStudyStateRich = 6
            End If
        Else
            getStudyStateRich = 1
        End If
         
        Exit Function
    End If
    
    rsTemp.Filter = "����ID='" & strDocID & "'"
    
    If rsTemp.RecordCount > 0 Then
        lngSendNO = NVL(rsTemp!���ͺ�, 0)
        bln���������� = Not IsNull(rsTemp!�������)
        blnCriticalValues = Not IsNull(rsTemp!Σ��״̬)
        blnImageQuality = Not IsNull(rsTemp!Ӱ������)
        blnReportQuality = Not IsNull(rsTemp!��������)
        blnConformDetermine = Not IsNull(rsTemp!�������)
    End If
    
    rsTemp.Filter = ""
    
    '������Ҫ����תԺ�Ĳ��ˣ�������������ɱ���������������飬��ʱִ�й���Ϊ6�Կ��ܴ���δ��ɵı���
    If NVL(rsTemp!ִ�й���, 0) = 6 And Not blnIsCancelComplete Then
        getStudyStateRich = 6
        Exit Function
    End If
    
    strSql = "Select ����״̬ From Ӱ�񱨸��¼ Where ҽ��ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "", lngOrderID)
    
    intReportCount = rsTemp.RecordCount
    
    For i = 0 To rsTemp.RecordCount - 1
        If NVL(rsTemp!����״̬) = 3 Or NVL(rsTemp!����״̬) = 4 Then '����˻�����
            intFinishReport = intFinishReport + 1
        End If
        '��¼�Ƿ��б��汻����
        If NVL(rsTemp!����״̬) = 6 Or NVL(rsTemp!����״̬) = 5 Then blnIsReject = True
        rsTemp.MoveNext
    Next
    
    If intFinishReport = rsTemp.RecordCount Then blnAllReportFinished = True    '���б��涼����˻�����
    
    rsTemp.Filter = "����״̬ = 4"  '������ı���
    If intReportCount = rsTemp.RecordCount Then '���б��涼������
        getStudyStateRich = 5
    Else
        getStudyStateRich = 4
    End If
    
    If blnIsReject = True Then getStudyStateRich = -1
    
    Exit Function
err:
    If ErrCenter() = 1 Then Resume
    
    Call SaveErrLog
End Function

Public Function getStudyState(ByVal lngOrderID As Long, Optional ByVal blnIsCancelComplete As Boolean = False, Optional ByRef lngSendNO As Long, _
                            Optional ByRef str������ As String, Optional ByRef strǩ�� As String, Optional ByRef str������ As String, _
                            Optional ByRef bln���������� As Boolean, Optional ByRef blnCriticalValues As Boolean, Optional ByRef blnImageQuality As Boolean, _
                            Optional ByRef blnReportQuality As Boolean, Optional ByRef blnConformDetermine As Boolean) As Integer
'��鱨���ǩ�������ȷ�����μ����еĳ̶ȡ�
'������ lngOrderID [IN] --- ҽ��id
'       lngSendNo [OUT] --- ���أ����ͺ�
'       str������ [OUT] --- ���أ�����Ĵ�����
'       strǩ��   [OUT] --- ���أ���������ǩ��
'       str������ [OUT] --- ���أ��������󱣴���
'       bln����������[OUT]--- ���أ���������Ƿ��Ѿ�����,True-�����룬False-δ����
'blnIsCancelComplete:�Ƿ���ִ�е�ȡ����ɲ���
'����ֵ��1--�ѵǼǣ�2--�ѱ�����3--�Ѽ�飻4--�ѱ��棻5--����ˣ�6--�����
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    Dim rsSign As ADODB.Recordset
    
    On Error GoTo err
    
    strSql = "Select d.ҽ��id As Ӱ��ҽ��ID,e.ҽ��id As ����ҽ��ID,c.���ͺ�,d.���uid,d.Ӱ������,d.�������,d.Σ��״̬,d.��������, " _
             & " e.����id,e.������, e.������, e.ǩ������, e.���ʱ��, e.���汾,c.�������,c.ִ�й��� " _
             & " From ����ҽ������ c, Ӱ�����¼ d, " _
             & " (Select a.ҽ��id,a.����id,b.������, b.������, b.ǩ������, b.���ʱ��, b.���汾 " _
             & "  From ����ҽ������ a, ���Ӳ�����¼ b Where a.ҽ��id = [1] And a.����id = b.Id) e " _
             & " Where c.ҽ��id = [1] And d.ҽ��id(+) = c.ҽ��id And e.ҽ��id(+) = c.ҽ��id "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "��ȡ�Ƿ�ǩ��", CLng(lngOrderID))
    
    '�����ѯû�н�������˳�
    If rsTemp.EOF = True Then Exit Function
    
    lngSendNO = rsTemp!���ͺ�
    str������ = NVL(rsTemp!������)
    str������ = NVL(rsTemp!������)
    bln���������� = Not IsNull(rsTemp!�������)
    blnCriticalValues = Not IsNull(rsTemp!Σ��״̬)
    blnImageQuality = Not IsNull(rsTemp!Ӱ������)
    blnReportQuality = Not IsNull(rsTemp!��������)
    blnConformDetermine = Not IsNull(rsTemp!�������)
    
    '���Ӱ��ҽ��IDΪ�գ������=1,�ѵǼ�
    '�������ҽ��IDΪ�գ��� ���UIDΪ�գ������=2���ѱ���
    '�������ҽ��IDΪ�գ����UID��Ϊ�գ������=3���Ѽ��
    '�������ǩ���ͱ�����������ȷ������Ϊ2,3,4��5���ѱ���,�Ѽ��,�ѱ��棬�����
    '������Ҫ����תԺ�Ĳ��ˣ�������������ɱ���������������飬��ʱִ�й���Ϊ6�Կ��ܴ���δ��ɵı���
    If NVL(rsTemp!ִ�й���, 0) = 6 And Not blnIsCancelComplete Then
        getStudyState = 6
        Exit Function
    End If
    
    If NVL(rsTemp!Ӱ��ҽ��ID) = "" Then     '����=1,�ѵǼ�
        getStudyState = 1
    ElseIf NVL(rsTemp!����ҽ��ID) = "" And NVL(rsTemp!���UID) = "" Then    '����=2���ѱ���
        getStudyState = 2
    ElseIf NVL(rsTemp!����ҽ��ID) = "" And NVL(rsTemp!���UID) <> "" Then    '����=3���Ѽ��
        getStudyState = 3
    Else    '���ǩ���ͱ���������,ȷ������Ϊ2,3,4��5���ѱ���,�Ѽ��,�ѱ��棬�����
        If NVL(rsTemp!���ʱ��) = "" And rsTemp!���汾 = 1 Then
            'δǩ������ �����һ��ҽʦ��ǩ��ִ�й�����ͼ��Ϊ�Ѽ�飬��ͼ��Ϊ�ѱ���
            getStudyState = IIf(NVL(rsTemp!���UID) = "", 2, 3)
        Else
            '�жϵ�ǰ�����ǩ���������������Ӳ������ݡ����д���1��ǩ��������������ˡ�
            If rsTemp!ǩ������ > 1 Then '�����
                getStudyState = 5
                '����˵��������Ҫ����ǩ���ˡ��������ݵ�����������˲�һ����ǩ���ˣ����Ҫ�������һ��ǩ����
                strSql = "Select Ҫ�ر�ʾ As ǩ������,�����ı� as ǩ��  From ���Ӳ������� Where �ļ�ID=[1] " _
                        & " And ��������= 8 And ��ʼ�� = [2] "
                Set rsSign = zlDatabase.OpenSQLRecord(strSql, "��ȡǩ������", CLng(rsTemp!����Id), CLng(rsTemp!���汾))
                
                If rsSign.EOF = False Then
                    strǩ�� = Split(NVL(rsSign!ǩ��), ";")(0)
                End If
            Else
                '�����������1������ǩ��������û�л������ݣ���2���޶�ģʽ�±��汨�棬��û��ǩ����
                '����������Ĳ�ѯ�����rsTemp!ǩ������ = 0 And rsTemp!���汾 > 1
                '�����˻��ˣ������޶�������û��ǩ�������������Ӧ�ô��ڡ������С���״̬��
                getStudyState = 4
            End If
        End If
    End If
    
    Exit Function
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function hibyte(ByVal wParam As Integer)
    hibyte = wParam \ &H100 And &HFF&
End Function

Private Function lobyte(ByVal wParam As Integer)
    lobyte = wParam And &HFF&
End Function

Private Function funPreDownLoadImages(thisMsg As TGetImgMsg) As Boolean
'------------------------------------------------
'���ܣ���̨����ͼ��
'������ thisMsg  -- Ҫ���ص�ͼ����Ϣ
'���أ�True -- �ɹ��� False -- ʧ��
'------------------------------------------------
    Dim lngWinHandle As Long        '��Ҫ������Ϣ�ġ�����ͼ�����ء�����Ĵ��ھ��
    Dim strMsg As String
    Dim wParam As Long
    Dim lResult As Long
    Dim strTemp As String
    Dim buf(1 To 1024) As Byte
    Dim dss As COPYDATASTRUCT
    
    On Error Resume Next
    
    '��֯��Ϣ
    strMsg = thisMsg.strSubDir & "||" & thisMsg.strDestMainDir & "||" & thisMsg.strIP & "||" _
            & thisMsg.strFtpDir & "||" & thisMsg.strFTPUser & "||" & thisMsg.strFTPPswd & "||" _
            & thisMsg.strSDDir & "||" & thisMsg.strSDUser & "||" & thisMsg.strSDPswd
    
    '����COPYDATA��Ϣ
    
    On Error GoTo err
    
    'ʹ��BUF������ʹ��lstrcpy�������������������ַ���Ϣ
   '��Ϣ���壺wParam = 123��dss��dwData = 3
    wParam = 123
   
    Call CopyMemory(buf(1), ByVal strMsg, LenB(StrConv(strMsg, vbFromUnicode)))
    dss.dwData = 3               '�����Ϣ���ã�3ֻ��˫�������һ����Ƕ���
    dss.cbData = LenB(StrConv(strMsg, vbFromUnicode)) + 1
    
    dss.lpData = VarPtr(buf(1))                    'ʹ��buf���ͣ����Կ�����Ϣ��1024֮��
'    dss.lpData = lstrcpy(strMsg, strMsg)            '����������͵���Ϣ��Ҳ����ȷ�ġ�
    
    
    '����ͼ�����ش���
    Shell App.Path & "\zlGetImage.exe"
        
    '���ش����ʱ�򣬲���ͼ�����س���
    lngWinHandle = FindWindow(vbNullString, "����ͼ������")
    
    lResult = SendMessage(lngWinHandle, WM_COPYDATA, wParam, dss)
    
    funPreDownLoadImages = True
    Exit Function
err:
    '�ݲ�����
End Function

Public Function DynamicCreate(ByVal strclass As String, ByVal strCaption As String) As Object
'��̬��������
    On Error Resume Next
    Set DynamicCreate = CreateObject(strclass)
   
    If err <> 0 Then
        MsgBox strCaption & "�������ʧ�ܣ�����ϵ����Ա����Ƿ���ȷ��װ!", vbInformation, gstrSysName
        Set DynamicCreate = Nothing
    End If
    err.Clear
End Function

Public Function CreateStudyUid(ByVal strUID As String) As String
'�������UID
    Dim rsData As New ADODB.Recordset
    Dim strSql As String
    Dim strNewStudyUID As String
    
    strNewStudyUID = strUID 'M_STR_STUDY_UID & "." & Format(Now, "yymmddhhmmss") & "." & Fix(Rnd(1000) * 1000)

    strSql = "select ���UID from Ӱ�����¼ where ���UID = [1]" & _
              " Union All Select ���UID from Ӱ����ʱ��¼ where ���UID = [1]"
              
    Set rsData = zlDatabase.OpenSQLRecord(strSql, "PACSͼ�񱣴�", strNewStudyUID)
    
    If rsData.RecordCount > 0 Then
        '����һ���µļ��UID
        strSql = "Select Ӱ����UID���_ID.Nextval From Dual"
        Set rsData = zlDatabase.OpenSQLRecord(strSql, "PACSͼ�񱣴�")
        
        If Len(strNewStudyUID) <= 55 Then
            strNewStudyUID = strNewStudyUID & ".A" & rsData(0)
        Else
            strNewStudyUID = Left(strNewStudyUID, 55) & ".A" & rsData(0)
        End If
    End If
    
    CreateStudyUid = strNewStudyUID
End Function


Public Function CreateSeriesUid(ByVal strUID As String) As String
'��������UID
    Dim rsData As New ADODB.Recordset
    Dim strSql As String
    Dim strNewSeriesUid As String
    
    strNewSeriesUid = strUID 'M_STR_SERIES_UID & "." & Format(Now, "yymmddhhmmss") & "." & Fix(Rnd(1000) * 1000)
    
    strSql = "select ����UID from Ӱ�������� where ����UID = [1]" & _
              " Union All Select ����UID from Ӱ����ʱ���� where ����UID = [1]"
              
    Set rsData = zlDatabase.OpenSQLRecord(strSql, "PACSͼ�񱣴�", strNewSeriesUid)
    
    If rsData.RecordCount > 0 Then
        '����һ���µļ��UID
        strSql = "Select Ӱ����UID���_ID.Nextval From Dual"
        Set rsData = zlDatabase.OpenSQLRecord(strSql, "PACSͼ�񱣴�")
        
        If Len(strNewSeriesUid) <= 55 Then
            strNewSeriesUid = strNewSeriesUid & ".A" & rsData(0)
        Else
            strNewSeriesUid = Left(strNewSeriesUid, 55) & ".A" & rsData(0)
        End If
    End If
    
    CreateSeriesUid = strNewSeriesUid
End Function

Public Function DeleteImages(frmParent As Form, intType As Integer, strImageUID As String, _
    strSeriesUID As String) As Boolean
'------------------------------------------------
'���ܣ�ɾ��FTP�е�һ��ͼ�����һ������
'������ frmParent -- ������
'       intType -- ɾ��ͼ������ͣ�1-ɾ��ͼ��2-ɾ������
'       strImageUID -- Ҫɾ��ͼ���UID��intType=1ʱ����Ҫ��ֵ
'       strSeriesUID - Ҫɾ������UID��intType=2ʱ����Ҫ��ֵ
'���أ�True -- �ɹ��� False -- ʧ��
'------------------------------------------------
    '�����ɾ��һ��ͼ��ͬʱɾ��ͬ������ͼ�����ù��� ZL_Ӱ��ͼ��_DELETE
    '�����ɾ��һ�����е�ͼ��ͬʱɾ��ͬ���ı���ͼ
    
    Dim Inet As New clsFtp             'FTP��
    Dim lngResult As Long
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    Dim strFtpIp As String
    Dim strFTPUser As String
    Dim strFtpPass As String
    Dim arrTmp() As String
    Dim strReportImage As String
    Dim intDeviceUsed As Integer
    Dim i As Integer
    Dim strRoot As String
    Dim strImagePath As String
    
    On Error GoTo err
    If intType = 1 And strImageUID = "" Then Exit Function
    If intType = 2 And strSeriesUID = "" Then Exit Function
    
    If intType = 1 Then         'ɾ��ͼ��
        strSql = "Select a.ҽ��ID,a.���ͺ�,c.ͼ��UID,a.����ͼ��, " & _
            " Decode(a.��������,Null,'',to_Char(a.��������,'YYYYMMDD')||'/')||a.���UID As ͼ��Ŀ¼, " & _
            "D.FTP�û��� As User1,D.FTP���� As Pwd1,D.IP��ַ As Host1,'/'||D.FtpĿ¼||'/' As Root1,d.�豸�� as �豸��1," & _
            "E.FTP�û��� As User2,E.FTP���� As Pwd2,E.IP��ַ As Host2,'/'||E.FtpĿ¼||'/' As Root2,e.�豸�� as �豸��2 " & _
            "From Ӱ�����¼ a,Ӱ�������� b,Ӱ����ͼ�� c,Ӱ���豸Ŀ¼ D,Ӱ���豸Ŀ¼ E " & _
            "Where a.���UID=b.���UID And b.����UID=c.����UID And c.ͼ��UID = [1] " & _
            "And a.λ��һ=D.�豸��(+) And a.λ�ö�=E.�豸��(+)"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "PACSɾ��ͼ��", strImageUID)
    ElseIf intType = 2 Then
        strSql = "Select a.ҽ��ID,a.���ͺ�,c.ͼ��UID, " & _
            " Decode(a.��������,Null,'',to_Char(a.��������,'YYYYMMDD')||'/')||a.���UID As ͼ��Ŀ¼, " & _
            "D.FTP�û��� As User1,D.FTP���� As Pwd1,D.IP��ַ As Host1,'/'||D.FtpĿ¼||'/' As Root1,d.�豸�� as �豸��1," & _
            "E.FTP�û��� As User2,E.FTP���� As Pwd2,E.IP��ַ As Host2,'/'||E.FtpĿ¼||'/' As Root2,e.�豸�� as �豸��2 " & _
            "From Ӱ�����¼ a,Ӱ�������� b,Ӱ����ͼ�� c,Ӱ���豸Ŀ¼ D,Ӱ���豸Ŀ¼ E " & _
            "Where a.���UID=b.���UID And b.����UID=c.����UID And b.����UID = [1] " & _
            "And a.λ��һ=D.�豸��(+) And a.λ�ö�=E.�豸��(+)"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "PACSɾ������", strSeriesUID)
    End If
    
    If rsTemp.EOF = True Then
        MsgBoxD frmParent, "û���ҵ�����ɾ����ͼ��!", vbInformation, gstrSysName
        DeleteImages = False
        Exit Function
    End If
    
    '�Ȳ����豸һ���ڲ����豸��
    If Not IsNull(rsTemp!�豸��1) Then
        strFtpIp = NVL(rsTemp!Host1)
        strFTPUser = NVL(rsTemp!User1)
        strFtpPass = NVL(rsTemp!Pwd1)
        lngResult = Inet.FuncFtpConnect(strFtpIp, strFTPUser, strFtpPass)
        intDeviceUsed = 1
        If lngResult = 0 Then
            If Not IsNull(rsTemp!�豸��2) Then
                strFtpIp = NVL(rsTemp!Host2)
                strFTPUser = NVL(rsTemp!User2)
                strFtpPass = NVL(rsTemp!Pwd2)
                lngResult = Inet.FuncFtpConnect(strFtpIp, strFTPUser, strFtpPass)
                intDeviceUsed = 2
                If lngResult = 0 Then
                    If MsgBoxD(frmParent, "����FTPʧ�ܣ��Ƿ����ɾ��ͼ��" & vbCrLf & "��ʱ����ɾ������ֻ��ɾ�����ݿ����ݣ��޷�ɾ��ͼ���ļ���" & vbCrLf & "���ǡ������ɾ����������ɾ����", vbQuestion + vbYesNo, gstrSysName) = vbNo Then
                        DeleteImages = False
                        Exit Function
                    End If
                End If
            End If
        End If
    End If
    
    strRoot = IIf(intDeviceUsed = 1, NVL(rsTemp!Root1), NVL(rsTemp!Root2))
    strImagePath = rsTemp!ͼ��Ŀ¼
    
    If intType = 1 Then
        '�����ɾ������ͼ����ɾ��ͬ������ͼ
        If Not IsNull(rsTemp("����ͼ��")) Then
            arrTmp = Split(rsTemp("����ͼ��"), ";")
            For i = 0 To UBound(arrTmp)
                If Trim(arrTmp(i)) <> strImageUID & ".jpg" Then
                    strReportImage = strReportImage & ";" & arrTmp(i)
                End If
            Next
            strReportImage = Mid(strReportImage, 2)
        End If
        
        strSql = "ZL_Ӱ��ͼ��_DELETE(" & rsTemp("ҽ��ID") & "," & rsTemp("���ͺ�") & ",'" & strImageUID & "','" & strReportImage & "')"
        zlDatabase.ExecuteProcedure strSql, "Ӱ��ͼ��ɾ��"
        
        'ɾ��ͼ���ļ�
        Call Inet.FuncDelFile(strRoot & strImagePath, strImageUID)
        Call Inet.FuncDelFile(strRoot & strImagePath, strImageUID & ".jpg")
    ElseIf intType = 2 Then
        '��ɾ��ͼ���ļ�,ͬʱɾ��ͬ���ı���ͼ
        rsTemp.MoveFirst
        While Not rsTemp.EOF
            Call Inet.FuncDelFile(strRoot & strImagePath, rsTemp!ͼ��UID)
            Call Inet.FuncDelFile(strRoot & strImagePath, rsTemp!ͼ��UID & ".jpg")
            rsTemp.MoveNext
        Wend
        
        '�����ɾ�����У���ֱ��ɾ�������е�ͼ��
        rsTemp.MoveFirst
        strSql = "Zl_Ӱ������_Delete(" & rsTemp("ҽ��ID") & ",'" & strSeriesUID & "')"
        zlDatabase.ExecuteProcedure strSql, "Ӱ������ɾ��"
        
        '���ɾ������֮�󣬱��μ��û��ͼ����ɾ��FTPĿ¼
        strSql = "Select ���UID from Ӱ�����¼ where ҽ��ID =[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "����Ƿ���ͼ��", CStr(rsTemp!ҽ��ID))
        If IsNull(rsTemp!���UID) Then
            'ɾ��Ŀ¼
            Call Inet.FuncFtpDelDir(strRoot, strImagePath)
        End If
    End If
    
    '�ر�FTP����
    Inet.FuncFtpDisConnect
    
    DeleteImages = True
    Exit Function
err:
    Inet.FuncFtpDisConnect
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    
End Function


Public Function GetDataToLocal(ByVal strSql As String, ByVal strTitle As String, ParamArray arrInput() As Variant) As ADODB.Recordset
'���ܣ�ͨ��Command����򿪴�����SQL�ļ�¼��
'������strSQL=�����а���������SQL���,������ʽΪ"[x]"
'             x>=1Ϊ�Զ��������,"[]"֮�䲻���пո�
'             ͬһ�������ɶദʹ��,�����Զ���ΪADO֧�ֵ�"?"����ʽ
'             ʵ��ʹ�õĲ����ſɲ�����,������Ĳ���ֵ��������(��SQL���ʱ��һ��Ҫ�õ��Ĳ���)
'      arrInput=���������Ĳ���ֵ,��������˳�����δ���,��������ȷ����
'               ��Ϊʹ�ð󶨱���,�Դ�"'"���ַ�����,����Ҫʹ��"''"��ʽ��
'      strTitle=����SQLTestʶ��ĵ��ô���/ģ�����
'���أ���¼����CursorLocation=adUseClient,LockType=adLockReadOnly,CursorType=adOpenStatic
'������
'SQL���Ϊ="Select ���� From ������Ϣ Where (����ID=[3] Or �����=[3] Or ���� Like [4]) And �Ա�=[5] And �Ǽ�ʱ�� Between [1] And [2] And ���� IN([6],[7])"
'���÷�ʽΪ��Set rsPati=OpenSQLRecord(strSQL, Me.Caption, CDate(Format(rsMove!ת������,"yyyy-MM-dd")),dtpʱ��.Value, lng����ID, "��%", "��", 20, 21)
    Dim cmdData As New ADODB.Command
    Dim rsData As New ADODB.Recordset
    Dim strPar As String, arrPar As Variant
    Dim lngLeft As Long, lngRight As Long
    Dim strSeq As String, intMax As Integer, i As Integer
    Dim strLog As String, varValue As Variant
    
    '�����Զ���[x]����
    lngLeft = InStr(1, strSql, "[")
    Do While lngLeft > 0
        lngRight = InStr(lngLeft + 1, strSql, "]")
        If lngRight = 0 Then Exit Do
        '������������"[����]����"
        strSeq = Mid(strSql, lngLeft + 1, lngRight - lngLeft - 1)
        If IsNumeric(strSeq) Then
            i = CInt(strSeq)
            strPar = strPar & "," & i
            If i > intMax Then intMax = i
        End If
        
        lngLeft = InStr(lngRight + 1, strSql, "[")
    Loop

    '�滻Ϊ"?"����
    strLog = strSql
    For i = 1 To intMax
        strSql = Replace(strSql, "[" & i & "]", "?")
        
        '��������SQL���ٵ����
        varValue = arrInput(i - 1)
        Select Case TypeName(varValue)
        Case "Byte", "Integer", "Long", "Single", "Double", "Currency" '����
            strLog = Replace(strLog, "[" & i & "]", varValue)
        Case "String" '�ַ�
            strLog = Replace(strLog, "[" & i & "]", "'" & Replace(varValue, "'", "''") & "'")
        Case "Date" '����
            strLog = Replace(strLog, "[" & i & "]", "To_Date('" & Format(varValue, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')")
        End Select
    Next

    '���ԭ�в���:��Ȼ�����ظ�ִ��
'    cmdData.CommandText = "" '��Ϊ����ʱ�����������
'    Do While cmdData.Parameters.Count > 0
'        cmdData.Parameters.Delete 0
'    Loop
    
    '�����µĲ���
    lngLeft = 0: lngRight = 0
    arrPar = Split(Mid(strPar, 2), ",")
    For i = 0 To UBound(arrPar)
        varValue = arrInput((arrPar(i) - 1))
        Select Case TypeName(varValue)
        Case "Byte", "Integer", "Long", "Single", "Double", "Currency" '����
            cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i, adVarNumeric, adParamInput, 30, varValue)
        Case "String" '�ַ�
            intMax = LenB(StrConv(varValue, vbFromUnicode))
            If intMax <= 2000 Then
                intMax = IIf(intMax <= 200, 200, 2000)
                cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i, adVarChar, adParamInput, intMax, varValue)
            Else
                If intMax < 4000 Then intMax = 4000
                cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i, adLongVarChar, adParamInput, intMax, varValue)
            End If
        Case "Date" '����
            cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i, adDBTimeStamp, adParamInput, , varValue)
        Case "Variant()" '����
            '���ַ�ʽ������һЩIN�Ӿ��Union���
            '��ʾͬһ�������Ķ��ֵ,�����Ų�������������Ĳ����Ž���,��Ҫ��֤�����ֵ��������
            If arrPar(i) <> lngRight Then lngLeft = 0
            lngRight = arrPar(i)
            Select Case TypeName(varValue(lngLeft))
            Case "Byte", "Integer", "Long", "Single", "Double", "Currency" '����
                cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i & "_" & lngLeft, adVarNumeric, adParamInput, 30, varValue(lngLeft))
                strLog = Replace(strLog, "[" & lngRight & "]", varValue(lngLeft), 1, 1)
            Case "String" '�ַ�
                intMax = LenB(StrConv(varValue(lngLeft), vbFromUnicode))
                If intMax <= 2000 Then
                    intMax = IIf(intMax <= 200, 200, 2000)
                    cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i & "_" & lngLeft, adVarChar, adParamInput, intMax, varValue(lngLeft))
                Else
                    If intMax < 4000 Then intMax = 4000
                    cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i & "_" & lngLeft, adLongVarChar, adParamInput, intMax, varValue(lngLeft))
                End If
                
                strLog = Replace(strLog, "[" & lngRight & "]", "'" & Replace(varValue(lngLeft), "'", "''") & "'", 1, 1)
            Case "Date" '����
                cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i & "_" & lngLeft, adDBTimeStamp, adParamInput, , varValue(lngLeft))
                strLog = Replace(strLog, "[" & lngRight & "]", "To_Date('" & Format(varValue(lngLeft), "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')", 1, 1)
            End Select
            lngLeft = lngLeft + 1 '�ò������������õ��ڼ���ֵ��
        End Select
    Next

    'ִ�з��ؼ�¼��
    'If cmdData.ActiveConnection Is Nothing Then
        Set cmdData.ActiveConnection = gcnOracle '���Ƚ���(���ִ��1000��Լ0.5x��)
    'End If
    cmdData.CommandText = strSql
    
    Call SQLTest(App.ProductName, strTitle, strLog)
    
    Set rsData.ActiveConnection = gcnOracle
    
    rsData.CursorLocation = adUseClient
    rsData.CursorType = adOpenDynamic
    rsData.LockType = adLockOptimistic
    
    rsData.Open cmdData
    
    Set rsData.ActiveConnection = Nothing
    
    Set GetDataToLocal = rsData
    
    Call SQLTest
End Function


Public Sub GetDeptStorageDevice(frmParent As Form, ByVal lngAdviceID As Long, ByVal strNewStudyUID As String, _
    ByVal lngCurDeptId As Long, ByVal lngModule As Long, ByVal blnMoved As Boolean, _
    ByRef strDeviceNO As String, ByRef strFtpIp As String, _
    ByRef strFtpUrl As String, ByRef strVirtualPath As String, _
    ByRef strFTPUser As String, ByRef strFTPPwd As String)
'------------------------------------------------------------------------------------------
'��ȡ�µĴ洢�豸��Ϣ������豸�洢��Ϣ�����ڣ�����Ҫ��������
'�����ȡ����������ʹ��strNewStudyUID�����ܴ����ݿ��в��ҵ���Ӧ������
'������ frmParent ---��IN����������
'       lngAdviceID---��IN����ҽ��ID
'       strNewStudyUID---��IN�������UID
'       lngCurDeptId ---��IN������ǰ����ID
'       lngModule---��IN����ģ���
'       blnMoved ---��IN�����Ƿ�ת��
'       strDeviceNO---��OUT�����豸��
'       strFtpIp---��OUT����ftp��ַ
'       strFtpUrl---��OUT����ftpĿ¼
'       strVirtualPath---��OUT����ftp����洢·��
'       strFtpUser---��OUT���� ftp�û���
'       strFtpPwd---��OUT����ftp����
'------------------------------------------------------------------------------------------
    Dim strSql As String
    Dim rsData As ADODB.Recordset
    Dim rsTemp As ADODB.Recordset
    Dim blnIsGetNewDevice As Boolean
    Dim objDestFtp As New clsFtp
    Dim curDate As Date
    Dim strDate As String
    
    strFtpIp = ""
    strFtpUrl = ""
    strFTPUser = ""
    strFTPPwd = ""
    strDeviceNO = ""
    
    strSql = "Select D.FTP�û��� As FtpUser,D.FTP���� As FtpPwd,C.λ��һ,C.λ�ö�,C.λ����,C.��������," & _
        "D.IP��ַ As Host," & _
        "'/'||D.FtpĿ¼||'/' As Root,Decode(C.��������,Null,'',to_Char(C.��������,'YYYYMMDD')||'/')" & _
        "||C.���UID As URL " & _
        "From Ӱ�����¼ C,Ӱ���豸Ŀ¼ D " & _
        "Where Decode(C.λ��һ,Null,C.λ�ö�,C.λ��һ)=D.�豸��(+)" & _
        "And C.���UID= [1] Union All " & _
        "Select D.FTP�û��� As FtpUser,D.FTP���� As FtpPwd,C.λ��һ,C.λ�ö�,C.λ����,C.��������," & _
        "D.IP��ַ As Host," & _
        "'/'||D.FtpĿ¼||'/' As Root,Decode(C.��������,Null,'',to_Char(C.��������,'YYYYMMDD')||'/')" & _
        "||C.���UID As URL " & _
        "From Ӱ����ʱ��¼ C,Ӱ���豸Ŀ¼ D " & _
        "Where Decode(C.λ��һ,Null,C.λ�ö�,C.λ��һ)=D.�豸��(+)" & _
        "And C.���UID= [1]"
        
    If blnMoved Then
        strSql = Replace(strSql, "Ӱ�����¼", "HӰ�����¼")
    End If
        
    Set rsData = zlDatabase.OpenSQLRecord(strSql, frmParent.Caption, strNewStudyUID)
    
    blnIsGetNewDevice = False
    
    If rsData.RecordCount <= 0 Then
        blnIsGetNewDevice = True
    Else
        '���ִ�е����˵����ִ��ͼ�����,��Ҫ�жϵ�ǰ���Ĵ洢�豸�Ƿ���Ч�������Ч�������µĴ洢�豸
        If Trim(rsData!��������) = "" Or (lngModule = G_LNG_PACSSTATION_MODULE And NVL(rsData!λ��һ) = "") Then
            blnIsGetNewDevice = True
            strDate = NVL(rsData!��������)
        Else
            strDeviceNO = NVL(rsData!λ��һ)
            strFtpIp = NVL(rsData!host)
            strFtpUrl = NVL(rsData!Root)
            strFTPUser = NVL(rsData!FtpUser)
            strFTPPwd = NVL(rsData!FtpPwd)
            strVirtualPath = strFtpUrl & NVL(rsData!Url)
        End If
    End If
    
    
    If blnIsGetNewDevice Then
        '�����µļ��UID�ʹ洢�豸,���ִ�е����˵����ȡ������
        
        If lngModule = 1290 Then
            '��ѯҽ������վ�У��������Ӧ�Ĵ洢�豸
            strSql = "select d.����ֵ " & _
                        " from ҽ��ִ�з��� a, ����ҽ������ b, Ӱ��DICOM����� c, Ӱ��DICOM������� d " & _
                        " Where a.����ID = b.ִ�в���id And a.ִ�м� = b.ִ�м� And a.����豸 = c.�豸�� " & _
                        " and c.������='ͼ�����' and c.����ID=d.����ID and d.��������='�洢�豸' and b.ҽ��id=[1]"
                        
            Set rsTemp = zlDatabase.OpenSQLRecord(strSql, frmParent.Caption, lngAdviceID)
            
            If rsTemp.RecordCount <= 0 Then
                MsgBoxD frmParent, "δ�ҵ�ͼ��洢�豸,��ȷ�ϵ�ǰ��������豸�Ƿ���Ӱ���豸Ŀ¼�ķ���������������ͼ��洢��", vbInformation, gstrSysName
                Exit Sub
            End If
            
            strDeviceNO = NVL(rsTemp!����ֵ)
        Else
            '��ѯ��ҽ������վ�е�ͼ��洢�豸
            strDeviceNO = GetDeptPara(lngCurDeptId, "�洢�豸��")
            
            If Val(strDeviceNO) <= 0 Then
                MsgBoxD frmParent, "δ�ҵ�ͼ��洢�豸,��ȷ����Ӱ�����̹������Ƿ�Ըÿ���������ͼ��ɼ��洢�豸��", vbInformation, gstrSysName
                Exit Sub
            End If
        End If
        
        
        strSql = "Select �豸��,�豸��,'/'||Decode(FtpĿ¼,Null,'',FtpĿ¼||'/') As URL,FTP�û���,FTP����,IP��ַ " & _
                    " From Ӱ���豸Ŀ¼ Where ����=1 and �豸��=[1] and NVL(״̬,0)=1"
                    
        Set rsTemp = zlDatabase.OpenSQLRecord(strSql, frmParent.Caption, strDeviceNO)
        
        '����洢�豸ͣ�ã���ֱ���˳�
        If rsTemp.RecordCount <= 0 Then
            MsgBoxD frmParent, "δ�ҵ��洢�豸,��ȷ���豸��Ϊ [" & strDeviceNO & "] ���豸�Ƿ����á�", vbInformation, gstrSysName
            Exit Sub
        End If

        strFtpUrl = NVL(rsTemp("URL"))
        strFtpIp = NVL(rsTemp("IP��ַ"))
        strFTPUser = NVL(rsTemp("FTP�û���"))
        strFTPPwd = NVL(rsTemp("FTP����"))
        
        strFtpUrl = IIf(strFtpUrl = "/", "//", strFtpUrl)
        
        On Error GoTo errHandle
        
        objDestFtp.FuncFtpConnect strFtpIp, strFTPUser, strFTPPwd
        
        If lngModule = G_LNG_PACSSTATION_MODULE Then
            strDate = Format(strDate, "YYYYMMDD")
        Else
            curDate = zlDatabase.Currentdate
            strDate = Format(curDate, "YYYYMMDD")
        End If
        
        strVirtualPath = strFtpUrl & strDate & "/" & strNewStudyUID
        '����FTPĿ¼
        objDestFtp.FuncFtpMkDir strFtpUrl, strDate & "/" & strNewStudyUID
            
        Call objDestFtp.FuncFtpDisConnect
    End If
    
    Exit Sub
    
errHandle:
        Call objDestFtp.FuncFtpDisConnect
End Sub


Public Sub WriteCommLog(logSubName As String, logTitle As String, logDesc As String)
'------------------------------------------------
'���ܣ���¼ͨѶ��־
'������ logSubName  --  ������־�ĺ�����
'       logTitle   -- ��־����
'       logDesc   --  ��־����
'���أ���
'------------------------------------------------
    Dim strLog As String
    
    strLog = Now() & " ���⣺ " & logTitle & vbCrLf & "      ������ " & logSubName & vbCrLf & "     ��־���ݣ�" & logDesc & vbCrLf

    LogWrite "PACS��Ҫ���ܵ�����־", glngModul, logTitle, strLog

End Sub

'����:WritTextFile
'����:FileName Ŀ���ļ���.WritStr д��Ŀ����ַ���.
'����ֵ:�ɹ� T.ʧ��  F
Public Function WritTextFile(ByVal strFileName As String, ByVal strWritStr As String) As Boolean
    Dim FileID As Long, ConTents As String
    Dim a As Long, b As Long
    Dim objFSO As New Scripting.FileSystemObject
    
    On Error Resume Next
    
    If objFSO.FileExists(strFileName) Then objFSO.DeleteFile (strFileName)
    
    FileID = FreeFile
    Open strFileName For Append As #FileID
         Print #FileID, strWritStr
    Close #FileID
    
    WritTextFile = (err.Number = 0)
    err.Clear
End Function

Public Function InitRegister() As Boolean
    If gobjRegister Is Nothing Then
        On Error Resume Next
        Set gobjRegister = VBA.Interaction.GetObject("", "zlRegister.clsRegister")
        err.Clear
    
        If gobjRegister Is Nothing Then
            Set gobjRegister = CreateObject("zlRegister.clsRegister")
            err.Clear
            If gobjRegister Is Nothing Then
                MsgBox "����zlRegister��������ʧ��,�����ļ��Ƿ���ڲ�����ȷע�ᡣ", vbExclamation, gstrSysName
                Exit Function
            End If
        End If
    End If
    
    SaveSetting "ZLSOFT", "����ȫ��", "��Ȩ���Ƴ���", "ZL9PACSWORK"
    
    gstrUserName = gobjRegister.GetUserName
    
    '�����Դ������������ֱ����������ΪHIS
    If App.LogMode = 0 Then
        gstrUserPswd = "HIS"
    Else
        gstrUserPswd = gobjRegister.GetPassword(App.hInstance)
    End If
    
    gstrServerName = gobjRegister.GetServerName
    
    InitRegister = True
End Function

Public Function GetNewLabel(lType As Integer, lLeft As Integer, lTop As Integer, lWidth As Integer, lHeight As Integer) As DicomLabel
'------------------------------------------------
'���ܣ�����һ��LABEL���󣬲���������ʼ����
'������lType--��ע�����ͣ�lLeft--��ע��Leftֵ��lTop--��ע��Topֵ��lWidth--��ע��Widthֵ��lHeight--��ע��Heightֵ��
'���أ������ɵı�ע��
'------------------------------------------------
    Dim l As New DicomLabel
    l.LabelType = lType
    l.ImageTied = True
    l.Left = lLeft
    l.Top = lTop
    l.Width = lWidth
    l.Height = lHeight
    l.Margin = 0
    l.AutoSize = True
    l.FontSize = 10
    l.LineWidth = 1
    'l.ForeColour = vbBlack
    l.XOR = True
    
    Set GetNewLabel = l
End Function

Public Sub subLabelCopyRebuild(Simg As DicomImage, oImg As DicomImage)
'------------------------------------------------
'���ܣ��ؽ�ͼ��ı�ע������ϵ
'������sImg--Դͼ��oImg--Ŀ��ͼ��
'���أ���
'------------------------------------------------
    Dim l As DicomLabel
    For Each l In oImg.Labels
        If Not l.TagObject Is Nothing Then
            If Simg.Labels.IndexOf(l.TagObject) <> 0 Then
                Set l.TagObject = oImg.Labels(Simg.Labels.IndexOf(l.TagObject))
            End If
        End If
    Next
End Sub
