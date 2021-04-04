Attribute VB_Name = "mdlLog"
Option Explicit

Public gobjComLib As Object
Public gobjLogComLib As Object  '公共日志管理

Public gobjlog As Object
Private Const G_STR_LOG_NAME = "PACS主要功能调试日志"
Private Const G_STR_PROJECT = "PACSQUERY"
Private ggg As Object
Public Enum gLogLevel
    EM_UnDefined = -1                 '尚未设置，应用于模块部件级别
    EM_LogOFF = 0                     '不记录日志
    EM_Error = 1                      '只记录错误
    EM_Warn = 2                       '记录警告
    EM_Info = 3                       '记录重要信息
    EM_Trace = 4                      '记录跟踪信息
    EM_All = 5                        '记录所有日志信息
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
