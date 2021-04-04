Attribute VB_Name = "mdlLog"
Option Explicit

Public gobjComLib As Object
Public gobjLogComLib As Object  '������־����

Public gobjlog As Object
Private Const G_STR_LOG_NAME = "PACS��Ҫ���ܵ�����־"
Private Const G_STR_PROJECT = "PACSQUERY"
Private ggg As Object
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
Public Sub zlWritLog(ByVal strFuncName As String, ByVal strLogInfor As String, ParamArray arrPars() As Variant)

    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strMoudle As String, varPara() As Variant
    Set gobjlog = Nothing
    If gobjlog Is Nothing Then
        Set gobjlog = CreateObject("zlLog.clsLog")
        Call gobjlog.SetBusinessDB(gcnOracle)
    End If
    varPara = arrPars
    If gobjlog Is Nothing Then Exit Sub
    Call gobjlog.Log(G_STR_LOG_NAME, G_STR_PROJECT, 1290, strFuncName, EM_Trace, strLogInfor, varPara)
End Sub
