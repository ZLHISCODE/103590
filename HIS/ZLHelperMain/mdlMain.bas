Attribute VB_Name = "mdlMain"
Option Explicit
'==================================================================================================
'��д           lshuo
'����           2019/1/21
'ģ��           mdlMain
'˵��
'==================================================================================================
Public gstrComputerName         As String
Public gstrIP                   As String
Public gstrAppsOFT              As String                       'APPSOFT·��
Public gobjFSO                  As New FileSystemObject         '����ȫ���ļ���������
Public gcllExecption            As New Collection               '�������쳣
Public gobjCurJob               As clsJob                       '��ǰ�ķ���������
Public Const G_LNG_MAX_JOBTRY   As Long = 4
Public Const G_LNG_MAX_SUBJOBTRY As Long = 4
Public gblnMsgBox               As Boolean
Public gstrCommand              As String
Public glngWinSessionID         As Long
Public gblnSessionIsAdimin      As Boolean                      '��ǰ�Ự��Ӧ���û��Ƿ����Ա��
Public Enum StartType
    ST_Service = 0              '������������������
    ST_SendServer = 1           '���������ַ�����Ϣ
    ST_SaveServer = 2           '������������˳�
    ST_Exit = 3                 '
End Enum
Public gstStartType             As StartType
'ʹ��SendMesage������Ϊ���������Ľ�����SYSTEMȨ�ޣ����Զ���������������ͨ�û�Ȩ�޽��̡�������Ϣ���Ͳ��ɹ�
Public gobjMetux                As New clsMutex
Public gobjHelperMain           As New clsMemoryShareFP         '�������ּ��໥ͨѶ
Public gobjSessionSID           As New clsMemoryShareFP         '��ǰ�ỰSID
Private Const M_HELPER_MAIN         As String = "67332524-C38A-4318-85C4-FA8151C85EDD"          '�������ּ����ͨѶID
Private Const M_SINGLE_INSTANCE     As String = "DEF43DC9-722D-48E0-9CBD-73E20E373E86"          '��֤��ʵ������
Private Const M_CURRENT_SESSION_SID As String = "DB130B49-92AA-488D-9D58-C1671CD21673"          '�洢SID
Public Const G_HISCUST_RUNNINH      As String = "7AA64FD7-9966-46D9-A10C-420AD5CEC766"          '�Զ�������������
Public gstrCurSID              As String       '��ǰ�Ự��sid
Public gobjLog                  As New clsLog
Private Const M_MAX_LOG_COUNT   As Long = 1000000
Public Const G_MAX_MEMORY_SIZE  As Long = 2048
Private mblnNormalTime          As Boolean          '��־�Ƿ���Ȼʱ���ʽ
Public gblnExit                 As Boolean              '�Ƿ񱣴���Ϣ���˳�
Public gstrOSUser                          As String                       '����ϵͳ�û�
Public gstrOSPWD                           As String                       '����ϵͳ����
Public gstrSystems                              As String                      '����ϵͳ
Sub Main()
    Dim objService      As clsService
    Dim blnNormalTime   As Boolean
    On Error GoTo ErrH
    glngWinSessionID = GetCurrentSessionID()
    '���������������֣���ǵ�ǰ�ỰID,����ϵͳ�û�������ϵͳ�û���SVRSTART SESSIONID=5 USERNAME=Administrator DOMAIN=Win7Base-PC
    '�Զ�����������²���:��ʱ����״̬��Ϣ��HELPERUPGRADE SAVEANDEXIT
    '��ѵһ�����ݿ⣺EXCFUNC DB=192.168.33.201:1521/TESTBASE35
    '��ǰ�����˳�����EXIT
    If IsDesinMode Then
        gstrCommand = "EXCFUNC DB=192.168.33.201:1521/TESTBASE35"
'        gstrCommand = "SVRSTART SESSIONID=5 USERNAME=Administrator DOMAIN=Win7Base-PC"
        gstrCommand = "HELPERUPGRADE SAVEANDEXIT"
'        gstrCommand = "EXIT"
    Else
        gstrCommand = Trim(CStr(Command()))
    End If
    mblnNormalTime = Val(GetSetting("ZLSOFT", "����ģ��\��������", "��Ȼʱ���ʽ", 0)) <> 0
    If App.PrevInstance Then
        Call gobjLog.LogOpen("ZLHelperMain_SessionID_" & glngWinSessionID & "_SedProc", , M_MAX_LOG_COUNT, Val(GetSetting("ZLSOFT", "����ģ��\��������", "���ָ�����־����", RLL_AllLog)), mblnNormalTime)
    Else
        If gstrCommand Like "SVRSTART *" Then
            Call gobjLog.LogOpen("ZLHelperMain_SessionID_" & glngWinSessionID, , M_MAX_LOG_COUNT, Val(GetSetting("ZLSOFT", "����ģ��\��������", "���ָ�����־����", RLL_AllLog)), mblnNormalTime)
        Else
            Call gobjLog.LogOpen("ZLHelperMain_SessionID_" & glngWinSessionID & "_FirstProc", , M_MAX_LOG_COUNT, Val(GetSetting("ZLSOFT", "����ģ��\��������", "���ָ�����־����", RLL_AllLog)), mblnNormalTime)
        End If
    End If
    Call gobjLog.PushMethod(RLL_AllLog, "ZLHelperMain.mdlMain.Main")
    gstStartType = ST_Service
    If gstrCommand Like "SVRSTART *" Then
        Call gobjMetux.CheckMutex(M_SINGLE_INSTANCE)
        gobjLog.LogInfo RLL_LogInfo, "��ǰ���������ַ��񿪱ٵ��������ֺ�̨���̡�", "����", gstrCommand
        If UCase(GetProcessUserName) <> "SYSTEM" Then
            gstStartType = ST_Exit
            Set gobjMetux = Nothing
        Else
            Call gobjHelperMain.CreateMemoryShare(M_HELPER_MAIN, G_MAX_MEMORY_SIZE)
        End If
        gstrCurSID = GetUserSID(GetSessionUser(gstrCommand))
        If gstrCurSID = "" Then
            gstrCurSID = GetUserSID(GetSessionUser(gstrCommand, True))
        End If
        If gstrCurSID <> "" Then
            Call gobjSessionSID.CreateMemoryShare(M_CURRENT_SESSION_SID, G_MAX_MEMORY_SIZE)
            Call gobjSessionSID.WriteMemory(gstrCurSID)
        End If
        gobjLog.LogInfo RLL_LogInfo, "���������ỰSID", gstrCurSID
    Else
        If gobjMetux.CheckMutex(M_SINGLE_INSTANCE) Then
            '��ʱ���������ֽ��̷�����������Ϣ
            gobjLog.LogInfo RLL_LogInfo, "��ǰ�Ѿ������������ַ��񿪱ٵ��������ֺ�̨���̡�", "����", gstrCommand
            gstStartType = ST_SendServer
            Call gobjHelperMain.OpenMemoryShare(M_HELPER_MAIN)
        Else
            gobjLog.LogInfo RLL_LogInfo, "��ǰ�����������ַ��񿪱ٵ��������ֺ�̨���̡�", "����", gstrCommand
            '��ʱû�к�̨�������У���ʱֻ��Ҫ���浽���������ļ����ɡ�
            gstStartType = ST_SaveServer
        End If
        Set gobjMetux = Nothing
    End If
    Call InitEnv
    If gstStartType > ST_Service Then
      If Not EnablePrivilege(, SE_DEBUG_NAME) Then
            gobjLog.LogInfo RLL_LogInfo, SE_DEBUG_NAME & "ʧ��"
        Else
            gobjLog.LogInfo RLL_LogInfo, SE_DEBUG_NAME & "�ɹ�"
        End If
    End If
    'ִ���˳���������ʱ������Ҫ��װ����
    If Not (gstrCommand Like "HELPERUPGRADE *" Or gstStartType = ST_SaveServer Or gstrCommand = "EXIT") Then
        Set objService = New clsService
        If Not objService.IsInstalled("ZLHelperService") Then
            If gobjFSO.FileExists(gstrAppsOFT & "\ZLHelperService.exe") Then
                Call objService.Install("ZLHelperService", "ZLSOFT Upgrade Helper Service", "�����������ַ���", gstrAppsOFT & "\ZLHelperService.exe")
            End If
        End If
        If objService.IsInstalled("ZLHelperService") Then
            If Not objService.IsRunning("ZLHelperService") Then
                Call objService.Start("ZLHelperService")
            End If
        End If
        Set objService = Nothing
    End If
    If (gstrCommand Like "HELPERUPGRADE *" And gstStartType = ST_SaveServer) Or gstStartType = ST_Exit Or gstrCommand = "" Then
        Set gobjHelperMain = Nothing
        Set gobjSessionSID = Nothing
        Set gobjMetux = Nothing
        Set gcllExecption = Nothing
        Call gobjLog.PopMethod(RLL_AllLog, "ZLHelperMain.mdlMain.Main")
        Call gobjLog.LogClose
        Set gobjLog = Nothing
        End
    Else
        Load frmMain
    End If
    Call gobjLog.PopMethod(RLL_AllLog, "ZLHelperMain.mdlMain.Main")
    Exit Sub
ErrH:
    If gobjLog.ErrCenter(RLL_RunError, "ZLHelperMain.mdlMain.Main") = 1 Then
        Resume
    End If
    Call gobjLog.LogClose
End Sub

'--------------------------------------------------------------------------------------------------
'����           GetSessionUser
'����           ���������л�ȡ��ǰ�Ự�û�
'����ֵ         String
'����б�:
'������         ����                    ˵��
'strCommand     String                  ������
'blnOnlyUser    Boolean                 ����ֻ�����û���
'-------------------------------------------------------------------------------------------------
Public Function GetSessionUser(ByVal strCommand As String, Optional ByVal blnOnlyUser As Boolean) As String
    Dim arrTmp  As Variant
    Dim strTmp  As String
    Dim strUser As String, strDomain    As String
    On Error GoTo ErrH
    Call gobjLog.PushMethod(RLL_AllLog, RLL_AllLog, "ZLHelperMain.mdlMain.GetSessionUser", strCommand, blnOnlyUser)
    arrTmp = Split(strCommand, "USERNAME=")
    strTmp = arrTmp(1)
    arrTmp = Split(strTmp, "DOMAIN=")
    strUser = Trim(arrTmp(0))
    strDomain = Trim(arrTmp(1))
    If blnOnlyUser Then
        GetSessionUser = strUser
    Else
        GetSessionUser = strDomain & "\" & strUser
    End If
    Call gobjLog.PopMethod(RLL_AllLog, RLL_AllLog, "ZLHelperMain.mdlMain.GetSessionUser", GetSessionUser)
    Exit Function
ErrH:
    If gobjLog.ErrCenter(RLL_RunError, "ZLHelperMain.mdlMain.GetSessionUser") = 1 Then
        Resume
    End If
End Function

'--------------------------------------------------------------------------------------------------
'����           InitEnv
'����           ��ʼ������
'����ֵ
'����б�:
'������         ����                    ˵��
'
'-------------------------------------------------------------------------------------------------
Private Sub InitEnv()
    Dim rllLogLevel     As RunLogLevel
    Dim objReg          As New clsRegistry
    Dim strRet          As String
    On Error GoTo ErrH
    Call gobjLog.PushMethod(RLL_AllLog, "ZLHelperMain.mdlMain.InitEnv")
    gstrIP = IP
    gstrComputerName = ComputerName
    If IsDesinMode Then
        gstrAppsOFT = "C:\APPSOFT"
    Else
        gstrAppsOFT = App.Path
    End If
    If gstrCurSID <> "" Then
        If objReg.GetRegValue("HKEY_USERS\" & gstrCurSID & "\Software\VB and VBA Program Settings\ZLSOFT\����ģ��\��������", "���ָ�����־����", strRet) Then
            If strRet = "" Then
                rllLogLevel = RLL_LogInfo
            Else
                rllLogLevel = Val(strRet)
            End If
            gobjLog.LogInfo RLL_LogInfo, "��ȡ" & "HKEY_USERS\" & gstrCurSID & "\Software\VB and VBA Program Settings\ZLSOFT\����ģ��\��������", "�ɹ�", "���ָ�����־����", rllLogLevel
        Else
            rllLogLevel = RLL_LogInfo
            gobjLog.LogInfo RLL_LogInfo, "��ȡ" & "HKEY_USERS\" & gstrCurSID & "\Software\VB and VBA Program Settings\ZLSOFT\����ģ��\��������", "ʧ��"
        End If
        gobjLog.CurrentLogLevel = rllLogLevel
        strRet = ""
        If objReg.GetRegValue("HKEY_USERS\" & gstrCurSID & "\Software\VB and VBA Program Settings\ZLSOFT\����ģ��\��������", "��Ȼʱ���ʽ", strRet) Then
            gobjLog.LogInfo RLL_LogInfo, "��ȡ" & "HKEY_USERS\" & gstrCurSID & "\Software\VB and VBA Program Settings\ZLSOFT\����ģ��\��������", "�ɹ�", "��Ȼʱ���ʽ", strRet
        Else
            gobjLog.LogInfo RLL_LogInfo, "��ȡ" & "HKEY_USERS\" & gstrCurSID & "\Software\VB and VBA Program Settings\ZLSOFT\����ģ��\��������", "ʧ��"
        End If
        mblnNormalTime = Val(strRet) <> 0
        gobjLog.IsNormalTime = mblnNormalTime
    End If
    gblnSessionIsAdimin = IsAdministrator
    gobjLog.LogInfo RLL_LogInfo, "���������û�", GetProcessUserName
    gobjLog.LogInfo RLL_LogInfo, "�����Ƿ�AdminȨ��", IsProcessRunAsAdmin
    gobjLog.LogInfo RLL_LogInfo, "�����Ƿ�Admin���û�", gblnSessionIsAdimin
    Call gobjLog.PopMethod(RLL_AllLog, "ZLHelperMain.mdlMain.InitEnv", gstrIP, gstrComputerName, gstrAppsOFT)
    Exit Sub
ErrH:
    If gobjLog.ErrCenter(RLL_RunError, "ZLHelperMain.mdlMain.InitEnv") = 1 Then
        Resume
    End If
End Sub

'--------------------------------------------------------------------------------------------------
'����           CheckSystem
'����           ���ݲ�����鵱ǰϵͳ�ܷ�֧�ֹ�����֤
'����ֵ
'����б�:
'������         ����                    ˵��
'
'-------------------------------------------------------------------------------------------------
Public Sub CheckSystem()
    On Error GoTo ErrH
    Call gobjLog.PushMethod(RLL_AllLog, "ZLHelperMain.mdlMain.CheckSystem")
    If gobjFSO.FileExists(gstrAppsOFT & "\APPLY\ZLBRW.DLL") Then
        If VerFull(gobjFSO.GetFileVersion(gstrAppsOFT & "\APPLY\ZLBRW.DLL")) >= VerFull("10.35.0.130") Then
            gstrSystems = gstrSystems & ",100"
        End If
    End If
    If gobjFSO.FileExists(gstrAppsOFT & "\APPLY\ZLLISBRW.DLL") Then
        If VerFull(gobjFSO.GetFileVersion(gstrAppsOFT & "\APPLY\ZLLISBRW.DLL")) >= VerFull("10.35.0.140") Then
            gstrSystems = gstrSystems & ",2500"
        End If
    End If
    If gobjFSO.FileExists(gstrAppsOFT & "\ZLHEALTHSTART.EXE") Then
        If VerFull(gobjFSO.GetFileVersion(gstrAppsOFT & "\ZLHEALTHSTART.EXE")) >= VerFull("10.35.0.130") Then
            gstrSystems = gstrSystems & ",2700"
        End If
    End If
    gstrSystems = Mid(gstrSystems, 2)
    Call gobjLog.PopMethod(RLL_AllLog, "ZLHelperMain.mdlMain.CheckSystem", gstrSystems)
    Exit Sub
ErrH:
    If gobjLog.ErrCenter(RLL_RunError, "ZLHelperMain.mdlMain.CheckSystem") = 1 Then
        Resume
    End If
End Sub
