Attribute VB_Name = "mdlLogManager"
Option Explicit
'*********************************************************************************************************************************************
'����:��־����
'�ӿ�˵��:
'   1.zlWriteLog:д��־
'����:��ΰ��
'����:2019*01*25 15:14:00
'*********************************************************************************************************************************************

Public gobjLog  As Object  '������־����

Private Const G_STR_PROJECT = "zl9CISJob"

Public Enum LogLevel
    LOGLEVEL_UnDefined = -1                 '��δ���ã�Ӧ����ģ�鲿������
    LOGLEVEL_LogOFF = 0                     '����¼��־
    LOGLEVEL_Error = 1                      'ֻ��¼����
    LOGLEVEL_Warn = 2                       '��¼����
    LOGLEVEL_Info = 3                       '��¼��Ҫ��Ϣ
    LOGLEVEL_Trace = 4                      '��¼������Ϣ
    LOGLEVEL_All = 5                        '��¼������־��Ϣ
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
    '����:��ΰ��
    '����:2020-02-28 15:43:41
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If mblnSetBusinessDB Or gobjLog Is Nothing Or gcnOracle Is Nothing Then Exit Sub
    If gcnOracle.State <> 1 Then Exit Sub
    Call gobjLog.SetBusinessDB(gcnOracle)
    mblnSetBusinessDB = True
End Sub


Public Sub WriteLogCall(ByVal strLogName As String, ByVal strModule As String, ByVal strFuncName As String, ByVal strCallName As String, ByVal lcsCurentLogCallState As LogCallState, ParamArray arrPars() As Variant)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��־д��
    '���: strFuncName-��������
    '     strCallName-������
    '     lcsCurentLogCallState-��ʶ���õ�ʱ������ʼ���û��߽������á�
    '     arrPars-д�����־��Ϣ
    '����:�ɹ�����true,���򷵻�False
    '����:��ΰ��
    '����:2020-02-28 15:08:36
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
    gobjLog.LogCall strLogName, G_STR_PROJECT, strModule, strFuncName, strCallName, lcsCurentLogCallState, varPara
End Sub

Public Sub zlWriteLog(ByVal strLogName As String, ByVal strModule As String, ByVal strFuncName As String, ByVal bytLogLevel As LogLevel, ByVal strLogInfor As String, ParamArray arrPars() As Variant)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��־����
    '���:
    '����:�ɹ�����true,���򷵻�False
    '����:��ΰ��
    '����:2020-02-28 19:05:36
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
    Call gobjLog.Log(strLogName, G_STR_PROJECT, strModule, strFuncName, bytLogLevel, strLogInfor, varPara)
End Sub

