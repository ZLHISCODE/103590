Attribute VB_Name = "mdlMain"
Option Explicit
'==================================================================================================
'��д           lshuo
'����           2019/1/21
'ģ��           mdlMain
'˵��
'==================================================================================================
Public gstrComputerName                     As String
Public gstrIP                               As String
Public gstrAppSOFT                          As String                       'APPSOFT·��
Public gobjFSO                              As New FileSystemObject         '����ȫ���ļ���������
Public gcllExecption                        As New Collection               '�������쳣
Public gobjCurJob                           As clsJob                       '��ǰ�ķ���������
Public Const G_LNG_MAX_JOBTRY               As Long = 4
Public Const G_LNG_MAX_SUBJOBTRY            As Long = 4
Public gblnMsgBox                           As Boolean
Public gstrCommand                          As String
Public glngWinSessionID                     As Long
Public Enum StartType
    ST_Service = 0              '������������������
    ST_SendServer = 1           '���������ַ�����Ϣ
    ST_SaveServer = 2           '������������˳�
    ST_Exit = 3                 '
End Enum
Public gstStartType                         As StartType
'ʹ��SendMesage������Ϊ���������Ľ�����SYSTEMȨ�ޣ����Զ���������������ͨ�û�Ȩ�޽��̡�������Ϣ���Ͳ��ɹ�
Public gobjMetux                            As New clsMutex
Public gobjHelperMain                       As New clsMemoryShareFP         '�������ּ��໥ͨѶ
Public gobjSessionSID                       As New clsMemoryShareFP         '��ǰ�ỰSID
Private Const M_HELPER_MAIN                 As String = "67332524-C38A-4318-85C4-FA8151C85EDD"          '�������ּ����ͨѶID
Private Const M_SINGLE_INSTANCE             As String = "DEF43DC9-722D-48E0-9CBD-73E20E373E86"          '��֤��ʵ������
Private Const M_CURRENT_SESSION_SID         As String = "DB130B49-92AA-488D-9D58-C1671CD21673"          '�洢SID
Public Const G_HISCUST_RUNNINH              As String = "7AA64FD7-9966-46D9-A10C-420AD5CEC766"          '�Զ�������������
Public gstrCurSID                           As String       '��ǰ�Ự��sid
Public Logger                               As New clsLogger
Private Const M_MAX_LOG_COUNT               As Long = 600000
Public Const G_MAX_MEMORY_SIZE              As Long = 2048
Private mblnNormalTime                      As Boolean              '��־�Ƿ���Ȼʱ���ʽ
Public gblnExit                             As Boolean              '�Ƿ񱣴���Ϣ���˳�
Public gstrSystems                          As String               '����ϵͳ
Public Const gstrSysName                    As String = "�������"
Public gstrZLSOFTRegKey                     As String

Sub Main()
    Dim objService      As clsService
    Dim blnNormalTime   As Boolean
    On Error GoTo ErrH
    Logger.IsUseCache = True        '���ڿ�ͷ����ʱ��־��û�п�������ʹ�û����¼
    Call Logger.PushMethod("ZLHelperMain.mdlMain.Main")
    glngWinSessionID = GetCurrentSessionID()
    '���������������֣���ǵ�ǰ�ỰID,����ϵͳ�û�������ϵͳ�û���SVRSTART SESSIONID=5 USERNAME=Administrator DOMAIN=Win7Base-PC
    '�Զ�����������²���:��ʱ����״̬��Ϣ��HELPERUPGRADE SAVEANDEXIT
    '��ѵһ�����ݿ⣺EXCFUNC DB=192.168.33.201:1521/TESTBASE35
    '��ǰ�����˳�����EXIT
    If IsDesinMode Then
'        gstrCommand = "EXCFUNC DB=192.168.33.201:1521/TESTBASE35"
        gstrCommand = "SVRSTART SESSIONID=1 USERNAME=Administrator DOMAIN=Win7Work-PC"
'        gstrCommand = "HELPERUPGRADE SAVEANDEXIT"
'        gstrCommand = "EXIT"
    Else
        gstrCommand = Trim(CStr(Command()))
    End If
    gstStartType = ST_Service
    If gstrCommand Like "SVRSTART *" Then
        Call gobjMetux.CheckMutex(M_SINGLE_INSTANCE)
        Logger.Info "��ǰ���������ַ��񿪱ٵ��������ֺ�̨���̡�", "����", gstrCommand
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
        Logger.Info "���������ỰSID", gstrCurSID
    Else
        If gobjMetux.CheckMutex(M_SINGLE_INSTANCE) Then
            '��ʱ���������ֽ��̷�����������Ϣ
            Logger.Info "��ǰ�Ѿ������������ַ��񿪱ٵ��������ֺ�̨���̡�", "����", gstrCommand
            gstStartType = ST_SendServer
            Call gobjHelperMain.OpenMemoryShare(M_HELPER_MAIN)
        Else
            Logger.Info "��ǰ�����������ַ��񿪱ٵ��������ֺ�̨���̡�", "����", gstrCommand
            '��ʱû�к�̨�������У���ʱֻ��Ҫ���浽���������ļ����ɡ�
            gstStartType = ST_SaveServer
        End If
        Set gobjMetux = Nothing
    End If
    Call InitEnv
    'ִ���˳���������ʱ������Ҫ��װ����
    If Not (gstrCommand Like "HELPERUPGRADE *" Or gstStartType = ST_SaveServer Or gstStartType = ST_SendServer Or gstrCommand = "EXIT") Then
        If gstStartType > ST_Service Then
          If Not EnablePrivilege(, SE_DEBUG_NAME) Then
                Logger.Info SE_DEBUG_NAME & "ʧ��"
            Else
                Logger.Info SE_DEBUG_NAME & "�ɹ�"
            End If
        End If
        Set objService = New clsService
        If Not objService.IsInstalled("ZLHelperService") Then
            If gobjFSO.FileExists(gstrAppSOFT & "\ZLHelperService.exe") Then
                Call objService.Install("ZLHelperService", "ZLSOFT Upgrade Helper Service", "�����������ַ���", gstrAppSOFT & "\ZLHelperService.exe")
            End If
        End If
        If objService.IsInstalled("ZLHelperService") Then
            If Not objService.IsRunning("ZLHelperService") Then
                Call objService.Start("ZLHelperService")
            End If
        End If
        Set objService = Nothing
    End If
    If (gstrCommand Like "HELPERUPGRADE *" And gstStartType = ST_SaveServer) Or gstStartType = ST_Exit Or gstrCommand = "" Or gstrCommand = "EXIT" Then
        Set gobjHelperMain = Nothing
        Set gobjSessionSID = Nothing
        Set gobjMetux = Nothing
        Set gcllExecption = Nothing
        Call Logger.PopMethod("ZLHelperMain.mdlMain.Main")
        Call Logger.CloseEx
        Set Logger = Nothing
        End
    Else
        Load frmMain
    End If
    Call Logger.PopMethod("ZLHelperMain.mdlMain.Main")
    Exit Sub
ErrH:
    If Logger.ErrCenter("ZLHelperMain.mdlMain.Main") = 1 Then
        Resume
    End If
    Call Logger.CloseEx
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
    Call Logger.PushMethod("ZLHelperMain.mdlMain.GetSessionUser", strCommand, blnOnlyUser)
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
    Call Logger.PopMethod("ZLHelperMain.mdlMain.GetSessionUser", GetSessionUser)
    Exit Function
ErrH:
    If Logger.ErrCenter("ZLHelperMain.mdlMain.GetSessionUser") = 1 Then
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
    Dim objReg          As New clsRegistry
    Dim strRet          As String
    Dim blnLoopTime     As Boolean, llLogLevel       As LogLevel
    On Error GoTo ErrH
    Call Logger.PushMethod("ZLHelperMain.mdlMain.InitEnv")
    gstrIP = IP
    gstrComputerName = ComputerName
    If IsDesinMode Then
        gstrAppSOFT = "C:\APPSOFT"
    Else
        gstrAppSOFT = App.Path
    End If
    If gstrCurSID <> "" Then
        If Is64bit Then
            gstrZLSOFTRegKey = "HKEY_USERS\" & gstrCurSID & "\Software\WOW6432Node\VB and VBA Program Settings\ZLSOFT"
        Else
            gstrZLSOFTRegKey = "HKEY_USERS\" & gstrCurSID & "\Software\VB and VBA Program Settings\ZLSOFT"
        End If
        If objReg.GetRegValue(gstrZLSOFTRegKey & "\����ģ��\��������", "���ָ�����־����", strRet) Then
            If strRet = "" Then
                llLogLevel = LogLevel_Trace
            Else
                llLogLevel = Val(strRet)
            End If
            Logger.Trace "��ȡ" & gstrZLSOFTRegKey & "\����ģ��\��������", "�ɹ�", "���ָ�����־����", llLogLevel
        Else
            llLogLevel = LogLevel_Trace
            Logger.Warn "��ȡ" & gstrZLSOFTRegKey & "\����ģ��\��������", "ʧ��"
        End If
        strRet = ""
        If objReg.GetRegValue(gstrZLSOFTRegKey & "\����ģ��\��������", "��Ȼʱ���ʽ", strRet) Then
            Logger.Trace "��ȡ" & gstrZLSOFTRegKey & "\����ģ��\��������", "�ɹ�", "��Ȼʱ���ʽ", strRet
        Else
            Logger.Warn "��ȡ" & gstrZLSOFTRegKey & "\����ģ��\��������", "ʧ��"
        End If
        If LenB(strRet) = 0 Then
            strRet = "1"
        End If
        blnLoopTime = Val(strRet) = 0
    Else
        llLogLevel = Val(GetSetting("ZLSOFT", "����ģ��\��������", "���ָ�����־����", LogLevel.LogLevel_Trace))
        blnLoopTime = Val(GetSetting("ZLSOFT", "����ģ��\��������", "��Ȼʱ���ʽ", 1)) = 0
    End If
    '�����ڲ�����ľ�����󼶱�
    Logger.SetComponentLogLevel "ZLHelperMain.clsMutex", LogLevel_Info
    Logger.SetComponentLogLevel "ZLHelperMain.clsMemoryShareFP", LogLevel_Info
    Logger.SetComponentLogLevel "ZLHelperMain.clsException", LogLevel_Info
    Logger.SetComponentLogLevel "ZLHelperMain.frmMain", LogLevel_Info
    Logger.SetComponentLogLevel "ZLHelperMain.mdlRunas", LogLevel_Info
    '����ע�������
    Call LoadRegistryLogLevel
    Call Logger.OpenEx("ZLHelperMain_SessionID_" & glngWinSessionID & "_" & Decode(gstStartType, ST_Service, "Service", ST_SendServer, "SendServer", ST_SaveServer, "SaveServer", ST_Exit, "Exit"), , M_MAX_LOG_COUNT, , blnLoopTime)
    Logger.DebugEx "���������û�", GetProcessUserName
    Logger.DebugEx "�����Ƿ�AdminȨ��", IsProcessRunAsAdmin
    Logger.DebugEx "�����Ƿ�Admin���û�", IsAdministrator

    Call Logger.PopMethod("ZLHelperMain.mdlMain.InitEnv", gstrIP, gstrComputerName, gstrAppSOFT)
    Exit Sub
ErrH:
    If Logger.ErrCenter("ZLHelperMain.mdlMain.InitEnv") = 1 Then
        Resume
    End If
End Sub

'@����    LoadRegistryLogLevel
'   ����ע����е�LogLevel����
'@����ֵ
'
'@����:
'@��ע
'   HKEY_USERS\[SID]\Software\WOW6432Node\VB and VBA Program Settings\ZLSOFT\����ģ��\��������
'       ������1
'           default=��־����
'           ģ��1=��־����2
Private Sub LoadRegistryLogLevel()
    Dim objReg          As New clsRegistry
    Dim arrSubKey       As Variant
    Dim strParentKey    As String
    Dim i               As Long, j          As Long
    Dim arrSubKeyValue  As Variant
    
    On Error GoTo ErrH
    Call Logger.PushMethod("ZLHelperMain.mdlMain.LoadRegistryLogLevel")
    If LenB(gstrZLSOFTRegKey) <> 0 Then
        arrSubKey = objReg.GetAllSubKey(gstrZLSOFTRegKey & "\����ģ��\��������")
        If TypeName(arrSubKey) <> "Empty" Then
            For i = LBound(arrSubKey) To UBound(arrSubKey)
                arrSubKeyValue = objReg.GetAllKeyValue(gstrZLSOFTRegKey & "\����ģ��\��������\" & arrSubKey(i))
                If TypeName(arrSubKeyValue) <> "Empty" Then
                    For j = LBound(arrSubKeyValue) To UBound(arrSubKeyValue) Step 2
                        If arrSubKeyValue(j) = "" Then
                            Logger.SetComponentLogLevel arrSubKey(i) & "", Val(arrSubKeyValue(j + 1))
                            Logger.DebugEx arrSubKey(i) & "", Val(arrSubKeyValue(j + 1))
                        Else
                            Logger.SetComponentLogLevel arrSubKey(i) & "." & arrSubKeyValue(j), Val(arrSubKeyValue(j + 1))
                            Logger.DebugEx arrSubKey(i) & "." & arrSubKeyValue(j), Val(arrSubKeyValue(j + 1))
                        End If
                    Next
                End If
            Next
        End If
    End If
    Call Logger.PopMethod("ZLHelperMain.mdlMain.LoadRegistryLogLevel")
    Exit Sub
ErrH:
    If Logger.ErrCenter("ZLHelperMain.mdlMain.LoadRegistryLogLevel") = 1 Then
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
    Call Logger.PushMethod("ZLHelperMain.mdlMain.CheckSystem")
    If gobjFSO.FileExists(gstrAppSOFT & "\APPLY\ZLBRW.DLL") Then
        If VerFull(gobjFSO.GetFileVersion(gstrAppSOFT & "\APPLY\ZLBRW.DLL")) >= VerFull("10.35.0.130") Then
            gstrSystems = gstrSystems & ",100"
        Else
            Logger.Warn gstrAppSOFT & "\APPLY\ZLBRW.DLL", "�汾�ţ�С��10.35.0.130��", gobjFSO.GetFileVersion(gstrAppSOFT & "\APPLY\ZLBRW.DLL")
        End If
    End If
    If gobjFSO.FileExists(gstrAppSOFT & "\APPLY\ZLLISBRW.DLL") Then
        If VerFull(gobjFSO.GetFileVersion(gstrAppSOFT & "\APPLY\ZLLISBRW.DLL")) >= VerFull("10.35.0.140") Then
            gstrSystems = gstrSystems & ",2500"
        Else
            Logger.Warn gstrAppSOFT & "\APPLY\ZLLISBRW.DLL", "�汾�ţ�С��10.35.0.130��", gobjFSO.GetFileVersion(gstrAppSOFT & "\APPLY\ZLLISBRW.DLL")
        End If
    End If
    If gobjFSO.FileExists(gstrAppSOFT & "\ZLHEALTHSTART.EXE") Then
        If VerFull(gobjFSO.GetFileVersion(gstrAppSOFT & "\ZLHEALTHSTART.EXE")) >= VerFull("10.35.0.130") Then
            gstrSystems = gstrSystems & ",2700"
        Else
            Logger.Warn gstrAppSOFT & "\ZLHEALTHSTART.EXE", "�汾�ţ�С��10.35.0.130��", gobjFSO.GetFileVersion(gstrAppSOFT & "\ZLHEALTHSTART.EXE")
        End If
    End If
    gstrSystems = Mid(gstrSystems, 2)
    Call Logger.PopMethod("ZLHelperMain.mdlMain.CheckSystem", gstrSystems)
    Exit Sub
ErrH:
    If Logger.ErrCenter("ZLHelperMain.mdlMain.CheckSystem") = 1 Then
        Resume
    End If
End Sub
