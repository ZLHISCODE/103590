Attribute VB_Name = "mdlHelperMain"
'@ģ�� mdlHelperMain-2019/7/5
'@��д lshuo
'@����
'
'@����
'
'@��ע
'
Option Explicit
'---------------------------------------------------------------------------
'                0��API�ͳ�������
'---------------------------------------------------------------------------
Public Const G_HELPER_RECEIVE              As String = "67332524-C38A-4318-85C4-FA8151C85EDD"          '���������ַ���������Ϣ���ڴ湲��
Public Const G_HELPER_SEND                 As String = "012CB54B-017C-4B9B-A0D5-956C0E90EFA6"          '������������̷�����Ϣ���ڴ湲��
Public Enum Helper_Share_Type
    HST_Waitting = 0            '�ȴ�״̬
    HST_UpgradeToServer = 1     '�Զ�������������������������ַ�����Ϣ��
    HST_ServerToUpgrade = 2     '�����������Զ�����������Ϣ
    HST_SelfToServer = 3        '�ֹ�������������������������ַ�����Ϣ
End Enum
Public Enum Helper_Share_State
    HST_UnWrited = 0            'δд��
    HST_Writed = 1              '�Ѿ�д��
End Enum
Public Const G_SINGLE_INSTANCE              As String = "DEF43DC9-722D-48E0-9CBD-73E20E373E86"          '��֤��ʵ������
Private Const M_MAX_LOG_COUNT               As Long = 600000
'Private Const M_MAX_LOG_COUNT               As Long = 100
Public Const G_MAX_MEMORY_SIZE              As Long = 2048
Public Const G_LNG_MAX_JOBTRY               As Long = 4
Public Const G_LNG_MAX_SUBJOBTRY            As Long = 3


Public Const gstrSysName                    As String = "�������"
'---------------------------------------------------------------------------
'                1���������
'---------------------------------------------------------------------------
'ȫ�ֶ���
Public Logger                               As New clsLogger
Public Process                              As New clsProcess
Public Environment                          As New clsEnvironment
Public Registry                             As New clsRegistry

Public gobjFSO                              As New FileSystemObject


'ʹ��SendMesage������Ϊ���������Ľ�����SYSTEMȨ�ޣ����Զ���������������ͨ�û�Ȩ�޽��̡�������Ϣ���Ͳ��ɹ�
Public gobjMetux                            As New clsMutex
Public gobjHelperMainRECEIVE                As New clsMemoryShare               '������ȡ��Ϣ
Public gobjHelperMainSend                   As New clsMemoryShare               '����������Ϣ
Public gstrCommand                          As String

Public gobjServerQueue                      As New clsQueue
Public gcllExecption                        As New Collection               '�������쳣

Public gblnMsgBox                           As Boolean
Public gblnExitProcess                      As Boolean
Public glngSendProcess                      As Long
Public Const G_VESRION                      As Long = 1
'---------------------------------------------------------------------------
'                2�����Ա����붨��
'---------------------------------------------------------------------------

'---------------------------------------------------------------------------
'                3����������
'---------------------------------------------------------------------------
Sub Main()
    Dim blnServer           As Boolean
    Dim arrVar              As Variant, i       As Long
    Dim objService          As clsService
    
    Const SE_DEBUG_NAME     As String = "SeDebugPrivilege"
    On Error GoTo ErrH
    Call Logger.PushMethod("ZLHelperMain.mdlHelperMain.Main")
    '�����ڲ�����ľ�����󼶱�
    Logger.AddComponentLogLevel "ZLHelperMain.clsMutex", LogLevel_Info
    Logger.AddComponentLogLevel "ZLHelperMain.clsMemoryShare", LogLevel_Info
    Logger.AddComponentLogLevel "ZLHelperMain.clsServerInfo", LogLevel_Info
    Logger.AddComponentLogLevel "ZLHelperMain.frmMain", LogLevel_Info
    Logger.AddComponentLogLevel "ZLHelperMain.clsProcess", LogLevel_Info
    Logger.AddComponentLogLevel "ZLHelperMain.clsEnvironment", LogLevel_Info
    Logger.AddComponentLogLevel "ZLHelperMain.clsRegistry", LogLevel_Info
    '���������������֣���ǵ�ǰ�ỰID,����ϵͳ�û�������ϵͳ�û���SVRSTART SESSIONID=5 USERNAME=Administrator DOMAIN=Win7Base-PC
    '�Զ�����������²���:��ʱ����״̬��Ϣ��HELPERUPGRADE SAVEANDEXIT
    '��ѵһ�����ݿ⣺EXCFUNC DB=192.168.33.201:1521/TESTBASE35
    '��ǰ�����˳�����EXIT
    If Environment.IsDesinMode Then
        gstrCommand = "EXCFUNC DB=127.0.0.1:1521/TESTBASE"
'        gstrCommand = "SVRSTART SESSIONID=1 USERNAME=Administrator DOMAIN=Win7Work-PC"
'        gstrCommand = "HELPERUPGRADE SAVEANDEXIT"
'        gstrCommand = "EXIT"
    Else
        gstrCommand = Trim(CStr(Command()))
    End If

    If gstrCommand Like "SVRSTART *" Then
        If Environment.ProcessUser = "SYSTEM" Then
            blnServer = True
        End If
    End If
    If blnServer Then
        Logger.OpenEx , "V" & G_VESRION, LogType_SingleInstance, False, M_MAX_LOG_COUNT, Logger.CurrentLogLevel, Logger.IsLoopTimeFormat
        '����������ZLHelperMain.EXE����
        arrVar = Process.ProcessesByProcessName("ZLHelperMain.EXE", GetCurrentProcessId())
        For i = LboundEx(arrVar) To UboundEx(arrVar)
            Process.ProcessTerminate arrVar(i)
        Next
        If Not gobjMetux.CheckMutex(G_SINGLE_INSTANCE) Then
            If Not gobjHelperMainRECEIVE.CreateMemoryShare(G_HELPER_RECEIVE) Or Not gobjHelperMainSend.CreateMemoryShare(G_HELPER_SEND) Then
                Set gobjHelperMainRECEIVE = Nothing
                Set gobjHelperMainSend = Nothing
                Set gobjMetux = Nothing
            End If
        Else
            Set gobjMetux = Nothing
        End If
        Call RestoreEnv
        Load frmMain
    Else
        If gobjMetux.CheckMutex(G_SINGLE_INSTANCE) Then
            Logger.OpenEx , "First_V" & G_VESRION, LogType_SingleInstance, False, M_MAX_LOG_COUNT, Logger.CurrentLogLevel, Logger.IsLoopTimeFormat
            If gstrCommand <> "" Then
                If gobjHelperMainRECEIVE.OpenMemoryShare(G_HELPER_RECEIVE) Then
                    'ʧ��Ҳ���ٴ������߼���������˵����ʧ��
                    Call gobjHelperMainRECEIVE.WriteMemory(gstrCommand, GetCurrentProcessId(), HST_SelfToServer, HST_Writed, False)
                End If
            End If
        Else
            Set gobjMetux = Nothing
            Logger.OpenEx , "Second_V" & G_VESRION, LogType_SingleInstance, False, M_MAX_LOG_COUNT, Logger.CurrentLogLevel, Logger.IsLoopTimeFormat
            'TODO,����������
            Call RestoreEnv
            gobjServerQueue.EnQueue gstrCommand
            Call SaveEnv
            If Not Process.EnablePrivilege(SE_DEBUG_NAME) Then
                Logger.DebugEx SE_DEBUG_NAME, "���", "ʧ��"
            Else
                Logger.DebugEx SE_DEBUG_NAME, "���", "�ɹ�"
            End If
        End If
        If Not gstrCommand Like "EXCFUNC DB=*" And gstrCommand <> "HELPERUPGRADE SAVEANDEXIT" And gstrCommand <> "EXIT" And Not gstrCommand Like "SVRSTART *" Then
            Set objService = New clsService
            If Not objService.IsInstalled("ZLHelperService") Then
                If gobjFSO.FileExists(AppsoftPath & "\ZLHelperService.exe") Then
                    Call objService.Install("ZLHelperService", "ZLSOFT Upgrade Helper Service", "�����������ַ���", AppsoftPath & "\ZLHelperService.exe")
                End If
            End If
            If objService.IsInstalled("ZLHelperService") Then
                If Not objService.IsRunning("ZLHelperService") Then
                    Call objService.Start("ZLHelperService")
                End If
            End If
            Set objService = Nothing
        End If
        Call ProcessExit
    End If
    Call Logger.PopMethod("ZLHelperMain.mdlHelperMain.Main")
    Exit Sub
ErrH:
    If Logger.ErrCenter("ZLHelperMain.mdlHelperMain.Main") = 1 Then
        Resume
    End If
    Call ProcessExit
End Sub

'@����    RestoreEnv
'   �ָ���������ѵ��Ϣ
'@����ֵ
'
'@����:
'@��ע
'
Public Sub RestoreEnv()
    Dim i           As Long, j As Long
    Dim strTmp      As String
    Dim objText     As TextStream
    Dim arrServer   As Variant
    Dim strFile     As String
    On Error GoTo ErrH
    Call Logger.PushMethod("ZLHelperMain.mdlHelperMain.RestoreEnv")
    strFile = AppsoftPath & "\ZLHelperMain_SessionID_" & Environment.SessionID & ".ini"
    If gobjFSO.FileExists(strFile) Then
        Set objText = gobjFSO.OpenTextFile(strFile, ForReading)
        strTmp = objText.ReadLine
        objText.Close
        Set objText = Nothing
        If strTmp Like "SERVER=*" Then
            strTmp = Mid(strTmp, Len("SERVER=*"))
            arrServer = UnSerialize(strTmp)
            If IsArray(arrServer) Then
                For i = LBound(arrServer) To UBound(arrServer)
                    If IsArray(arrServer(i)) Then
                        For j = LBound(arrServer(i)) To UBound(arrServer(i))
                            Call gobjServerQueue.EnQueue(arrServer(i)(j))
                        Next
                    Else
                        Call gobjServerQueue.EnQueue(arrServer(i))
                    End If
                Next
            Else
                Call gobjServerQueue.EnQueue(arrServer)
            End If
        End If
        gobjFSO.DeleteFile strFile, True
    End If
    Logger.Info "��ȡ�����������", "Count", gobjServerQueue.Count
    Call Logger.PopMethod("ZLHelperMain.mdlHelperMain.RestoreEnv")
    Exit Sub
ErrH:
    If Logger.ErrCenter("ZLHelperMain.mdlHelperMain.RestoreEnv") = 1 Then
        Resume
    End If
    
    Call Logger.PopMethod("ZLHelperMain.mdlHelperMain.RestoreEnv")
End Sub

'--------------------------------------------------------------------------------------------------
'����           SaveEnv
'����           �����������ѵ��Ϣ��ֻ��׼���˳�ʱ�ű���
'����ֵ
'����б�:
'������         ����                    ˵��
'strServer      String                  TCP���˳�ʱ�������������ݡ�
'-------------------------------------------------------------------------------------------------
Public Sub SaveEnv()
    Dim objQueue    As New clsQueue
    Dim i           As Long
    Dim strTmp      As String
    Dim objText     As TextStream
    Dim strServer   As String
    On Error GoTo ErrH
    Call Logger.PushMethod("mdlHelperMain.mdlHelperMain.SaveEnv")
    objQueue.QueueSize = 5
    Do While Not gobjServerQueue.IsEmpty
        If gobjServerQueue.Current = "HELPERUPGRADE SAVEANDEXIT" Then
            'Do Nothing
        ElseIf Not gobjServerQueue.Current Like "SVRSTART *" Then
            'EXCFUNC DB=192.168.33.201:1521/TESTBASE35
            If gobjServerQueue.Current Like "EXCFUNC DB=*" Then
                strServer = Mid(gobjServerQueue.Current, Len("EXCFUNC DB=*"))
            Else
                strServer = gobjServerQueue.Current
            End If
            objQueue.EnQueue strServer
        End If
        gobjServerQueue.DeQueue
    Loop
    
    For i = gcllExecption.Count To 1 Step -1
        objQueue.EnQueue gcllExecption(i).Server
        gcllExecption.Remove i
    Next
    If Not objQueue.IsEmpty() Then
        strTmp = SerializeEx(objQueue.Data)
        Logger.Info "���滺�������", "Count", objQueue.Count
        Set objText = gobjFSO.OpenTextFile(AppsoftPath & "\ZLHelperMain_SessionID_" & Environment.SessionID & ".ini", ForWriting, True)
        objText.WriteLine "SERVER=" & strTmp
        objText.Close
        Set objText = Nothing
    End If
    
    Call Logger.PopMethod("ZLHelperMain.mdlHelperMain.SaveEnv")
    Exit Sub
ErrH:
    If Logger.ErrCenter("ZLHelperMain.mdlHelperMain.SaveEnv") = 1 Then
        Resume
    End If
    
    Call Logger.PopMethod("mdlHelperMain.mdlHelperMain.SaveEnv")
End Sub

Public Sub ProcessExit()
    Dim objService      As New clsService
    Set gobjHelperMainSend = Nothing
    Set gobjHelperMainRECEIVE = Nothing
    Set gobjMetux = Nothing
    Set gcllExecption = Nothing
    '֪ͨ�˳�ʱ�رշ���
    If gblnExitProcess Then
        If objService.IsInstalled("ZLHelperService") Then
            If objService.IsRunning("ZLHelperService") Then
                Call objService.Stopping("ZLHelperService")
            End If
        End If
    End If
    Call Logger.PopMethod("ZLHelperMain.mdlHelperMain.Main")
    Set Registry = Nothing
    Set Process = Nothing
    Set Environment = Nothing
    Call Logger.CloseEx
    Set Logger = Nothing
End Sub
'---------------------------------------------------------------------------
'                4��˽�з���
'---------------------------------------------------------------------------

'---------------------------------------------------------------------------
'                5�����󷽷����¼�
'---------------------------------------------------------------------------

