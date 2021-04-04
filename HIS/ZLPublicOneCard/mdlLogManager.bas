Attribute VB_Name = "mdlLogManager"
Option Explicit
'*********************************************************************************************************************************************
'����:��־����
'�ӿ�˵��:
'   1.zlWritLog:д��־
'����:���˺�
'����:2019*01*25 15:14:00
'*********************************************************************************************************************************************

Public gobjLog  As Object  '������־����
Private Const G_STR_LOG_NAME = "һ��ͨ�ӿڵ�����־"
Private Const G_STR_PROJECT = "zlOneCardComLib"
Public Enum gLogCallState
    LogCallState_CallBegin = 0
    LogCallState_CallEnd = 1
End Enum
Public Enum gLogLevel
    EM_UnDefined = -1                 '��δ���ã�Ӧ����ģ�鲿������
    EM_LogOFF = 0                     '����¼��־
    EM_Error = 1                      'ֻ��¼����
    EM_Warn = 2                       '��¼����
    EM_Info = 3                       '��¼��Ҫ��Ϣ
    EM_Trace = 4                      '��¼������Ϣ
    EM_All = 5                        '��¼������־��Ϣ
End Enum
Public Enum LogCallState
    LogCallState_CallBegin = 0
    LogCallState_CallEnd = 1
End Enum
Private mblnCreateLog As Boolean    '�����ظ�����
Private mblnSetBusinessDB As Boolean '��ʼ��������

Private Sub SetLogBusinessDB()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ��ҵ�����Ӹ���־����
    '����:���˺�
    '����:2020-02-06 15:43:41
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If mblnSetBusinessDB Or gobjLog Is Nothing Or gcnOracle Is Nothing Then Exit Sub
    If gcnOracle.State <> 1 Then Exit Sub
    Call gobjLog.SetBusinessDB(gcnOracle)
    mblnSetBusinessDB = True
End Sub


Public Sub WritLogCall(strFuncName As String, strCallName As String, lcsCurentLogCallState As gLogCallState, ParamArray arrPars() As Variant)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��־д��
    '���: strFuncName-��������
    '     strCallName-������
    '     lcsCurentLogCallState-��ʶ���õ�ʱ������ʼ���û��߽������á�
    '     arrPars-д�����־��Ϣ
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2019-01-15 15:08:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strMoudle As String, varPara() As Variant
    
    varPara = arrPars
    'strLogName In String
    '   ��־��ҵ��������ơ���һ��ͨ��־�ȡ���""ʱȡ���һ�β�Ϊ��
    'strComponentName In strComponentName
    '   ��־�����Ĳ������ơ�ʹ��App.EXEName����""ʱȡ���һ�β�Ϊ��
    'strModule In String
    '   ��־������ģ�顣������ZLHIS��ϵ��ģ�������VB��ģ��ȡ���""ʱȡ���һ�β�Ϊ��
    'strFuncName In String
    '   ��־�ķ����Ĺ����������߷�����VB��������""ʱȡ���һ�β�Ϊ��
    'strCallName In String
    '   WebAPI���ƻ��ߴ洢��������
    'lcsCurentLogCallState In LogCallState
    '   ��ʶ���õ�ʱ������ʼ���û��߽������á�
    'arrPars In ParamArray
    '   ������ʽ�� arrPars(0),arrPars(1),...,arrPars(n)
    '@��ע
    If gobjLog Is Nothing Then
        If mblnCreateLog = True Then Exit Sub  'ֻ����һ��.
        Err = 0: On Error Resume Next
        Set gobjLog = CreateObject("zlLog.clsLog")
        mblnCreateLog = True
        If Err <> 0 Then
            Err = 0: On Error GoTo 0: Exit Sub
        End If
    End If
    Call SetLogBusinessDB
    gobjLog.LogCall G_STR_LOG_NAME, G_STR_PROJECT, CStr(glngModul), strFuncName, strCallName, lcsCurentLogCallState, varPara
End Sub
Public Sub zlWritLog(strFuncName As String, strLogInfor As String, ParamArray arrPars() As Variant)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��־����
    '���:
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2020-01-08 19:05:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strMoudle As String, varPara() As Variant
    varPara = arrPars
    '@����    Log
    '   ��¼ָ����־�������־��
    '@����ֵ  Boolean
    '
    '@����:
    'strLogName In String
    '   ��־��ҵ��������ơ���һ��ͨ��־�ȡ���""ʱȡ���һ�β�Ϊ��
    'strComponentName In strComponentName
    '   ��־�����Ĳ������ơ�ʹ��App.EXEName����""ʱȡ���һ�β�Ϊ��
    'strModule In String
    '   ��־������ģ�顣������ZLHIS��ϵ��ģ�������VB��ģ��ȡ���""ʱȡ���һ�β�Ϊ��
    'strFuncName In String
    '   ��־�ķ����Ĺ����������߷�����VB��������""ʱȡ���һ�β�Ϊ��
    'llLogLevel In LogLevel
    '   ��ǰ��¼��־���ʡ�
    '   LogLevel_Error������־������ҵ��������VB���󡢳������ȣ���Ӱ��������еġ�
    '   LogLevel_Warn������־�����ڴ��󣬲�Ӱ��������У����ǿ���������̱䶯���߳����ܲ�ȫ�����������ݿ��ƻ��ߵ�ǰ������أ���ȱʧĳ��������Ȼ���Լ���ʹ�ã����Ƕ�Ӧ����ȱʧ��
    '   LogLevel_Info������־������Ҫ��Ϣ��¼��������Ҫ��Ϣ�ļ�¼������ã����׵����ݡ�
    '   LogLevel_Trace������Ϣ�ǳ�������и�����Ϣ�����ڸ��ٳ������У��Ա㷽������֤��
    'arrPars In ParamArray
    '   ������ʽ��strLogTilte: arrPars(0),arrPars(1),arrPars(2),arrPars(3)...
    '@��ע
    If gobjLog Is Nothing Then
        If mblnCreateLog = True Then Exit Sub  'ֻ����һ��.
        Err = 0: On Error Resume Next
        Set gobjLog = CreateObject("zlLog.clsLog")
        mblnCreateLog = True
        If Err <> 0 Then
            Err = 0: On Error GoTo 0: Exit Sub
        End If
    End If
    Call SetLogBusinessDB
    Call gobjLog.Log(G_STR_LOG_NAME, G_STR_PROJECT, CStr(glngModul), strFuncName, EM_Trace, strLogInfor, varPara)
End Sub
'
'
'Public Sub zlWritLog(ByVal lngModule As Long, ByVal strFunName As String, ByVal strCallFunName As String, _
'    ByVal strLogInfor As String, Optional ByVal intLogType As Integer = 0, Optional strLogName As String = "һ��ͨ�ӿڵ�����־", _
'    Optional strGroupName As String)
'    '---------------------------------------------------------------------------------------------------------------------------------------------
'    '����:��־д��
'    '���:lngModule-��ǰģ���
'    '     strCallFunName-����������
'    '     strFunName-��������
'    '     intLogType-��־����:0-������־;1-����SQL;2-������Ϣ
'    '     strLogInfor-д�����־����
'    '     strLogName-��־����
'    '     strGroupName-����
'    '����:
'    '����:�ɹ�����true,���򷵻�False
'    '����:���˺�
'    '����:2019-01-15 15:08:36
'    '---------------------------------------------------------------------------------------------------------------------------------------------
'    Dim objLogManager As Object
'    On Error GoTo errHandle
'    If zlGetLogManagerObject(objLogManager) = False Then
'        On Error Resume Next
'        Call gobjComLib.LogWrite(strLogName, lngModule, strFunName, "������:" & strCallFunName & IIf(strGroupName = "", "", "-" & strGroupName) & vbTab & strLogInfor)
'        If Err.Number <> 0 And Err.Number <> 438 Then GoTo errHandle
'        Exit Sub
'    End If
'    If objLogManager Is Nothing Then Exit Sub
'    Call gobjLogManager.zlWritLog(lngModule, strFunName, strCallFunName, strLogInfor, intLogType, strLogName)
'    Set objLogManager = Nothing
'    Exit Sub
'errHandle:
'    If gobjComLib.ErrCenter() = 1 Then Resume
'End Sub
'
'Public Function zlGetLogManagerObject(ByRef objLogManager As Object) As Boolean
'    '---------------------------------------------------------------------------------------------------------------------------------------------
'    '����:��ȡ���ù�����������
'    '����:objLogManager-������־����������
'    '����:��ȡ����true,���򷵻�False
'    '����:���˺�
'    '����:2019-01-25 09:57:11
'    '---------------------------------------------------------------------------------------------------------------------------------------------
'
'    On Error GoTo errHandle
'    If Not gobjLogManager Is Nothing Then
'        Set objLogManager = gobjLogManager: zlGetLogManagerObject = True
'        Exit Function
'    End If
'    If gblnCreateLogManager Or gcnOracle Is Nothing Then Exit Function  'ֻ��ʼ��һ��,����ʱ�����ٳ�ʼ��
'
'
'    Err = 0: On Error Resume Next
'    If gobjLogManager Is Nothing Then
'        Set gobjLogManager = CreateObject("zlLogManager.clsLogManager")
'        gblnCreateLogManager = True
'        If Err <> 0 Then Exit Function
'    End If
'
'    Err.Clear:  On Error GoTo errHandle
'    If gobjLogManager Is Nothing Then Exit Function
'
'    'zlInitCommon(ByVal lngSys As Long, _
'     ByVal cnOracle As ADODB.Connection, Optional ByVal strDbUser As String) As Boolean
'    '---------------------------------------------------------------------------------------------------------------------------------------------
'    '����:��ʼ����ص�ϵͳ�ż��������
'    '���:lngSys-ϵͳ��
'    '     cnOracle-���ݿ����Ӷ���
'    '     strDBUser-���ݿ�������
'    '����:��ʼ���ɹ�,����true,���򷵻�False
'    If gobjLogManager.InitCommon(gcnOracle, gstrDBUser) = False Then Exit Function
'    Set objLogManager = gobjLogManager: zlGetLogManagerObject = True
'    Exit Function
'errHandle:
'    Exit Function
'End Function
 

