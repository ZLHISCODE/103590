Attribute VB_Name = "mdlLogManager"
Option Explicit
'*********************************************************************************************************************************************
'功能:日志管理
'接口说明:
'   1.zlWriteLog:写日志
'编制:余伟节
'日期:2019*01*25 15:14:00
'*********************************************************************************************************************************************

Public gobjLog  As Object  '公共日志管理

Private Const G_STR_PROJECT = "zl9CISJob"

Public Enum LogLevel
    LOGLEVEL_UnDefined = -1                 '尚未设置，应用于模块部件级别
    LOGLEVEL_LogOFF = 0                     '不记录日志
    LOGLEVEL_Error = 1                      '只记录错误
    LOGLEVEL_Warn = 2                       '记录警告
    LOGLEVEL_Info = 3                       '记录重要信息
    LOGLEVEL_Trace = 4                      '记录跟踪信息
    LOGLEVEL_All = 5                        '记录所有日志信息
End Enum

Public Enum LogCallState
    LogCallState_CallBegin = 0
    LogCallState_CallEnd = 1
End Enum

Private mblnCreateLog As Boolean    '避免重复创建
Private mblnSetBusinessDB As Boolean '初始化了连接

Private Sub SetLogBusinessDB()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化业务连接给日志对象
    '编制:余伟节
    '日期:2020-02-28 15:43:41
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If mblnSetBusinessDB Or gobjLog Is Nothing Or gcnOracle Is Nothing Then Exit Sub
    If gcnOracle.State <> 1 Then Exit Sub
    Call gobjLog.SetBusinessDB(gcnOracle)
    mblnSetBusinessDB = True
End Sub


Public Sub WriteLogCall(ByVal strLogName As String, ByVal strModule As String, ByVal strFuncName As String, ByVal strCallName As String, ByVal lcsCurentLogCallState As LogCallState, ParamArray arrPars() As Variant)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:日志写入
    '入参: strFuncName-功能名称
    '     strCallName-调用者
    '     lcsCurentLogCallState-标识调用的时机。开始调用或者结束调用。
    '     arrPars-写入的日志信息
    '返回:成功返回true,否则返回False
    '编制:余伟节
    '日期:2020-02-28 15:08:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strMoudle As String, varPara() As Variant
    
    varPara = arrPars
    'strLogName In String
    '   日志的业务分类名称。如一卡通日志等。传""时取最近一次不为空
    'strComponentName In strComponentName
    '   日志发生的部件名称。使用App.EXEName。传""时取最近一次不为空
    'strModule In String
    '   日志发生的模块。可以是ZLHIS体系的模块可以是VB的模块等。传""时取最近一次不为空
    'strFuncName In String
    '   日志的发生的功能名。或者发生的VB函数。传""时取最近一次不为空
    'strCallName In String
    '   WebAPI名称或者存储过程名称
    'lcsCurentLogCallState In LogCallState
    '   标识调用的时机。开始调用或者结束调用。
    'arrPars In ParamArray
    '   产生格式： arrPars(0),arrPars(1),...,arrPars(n)
    '@备注
    If gobjLog Is Nothing Then
        If mblnCreateLog = True Then Exit Sub  '只创建一次.
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
    '功能:日志跟踪
    '入参:
    '返回:成功返回true,否则返回False
    '编制:余伟节
    '日期:2020-02-28 19:05:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strMoudle As String, varPara() As Variant
    varPara = arrPars
    '@方法    Log
    '   记录指定日志级别的日志。
    '@返回值  Boolean
    '
    '@参数:
    'strLogName In String
    '   日志的业务分类名称。如一卡通日志等。传""时取最近一次不为空
    'strComponentName In strComponentName
    '   日志发生的部件名称。使用App.EXEName。传""时取最近一次不为空
    'strModule In String
    '   日志发生的模块。可以是ZLHIS体系的模块可以是VB的模块等。传""时取最近一次不为空
    'strFuncName In String
    '   日志的发生的功能名。或者发生的VB函数。传""时取最近一次不为空
    'llLogLevel In LogLevel
    '   当前记录日志性质。
    '   LogLevel_Error：该日志是属于业务错误或者VB错误、程序错误等，会影响程序运行的。
    '   LogLevel_Warn：该日志不属于错误，不影响程序运行，但是可能造成流程变动或者程序功能不全。可能与数据控制或者当前环境相关，如缺失某部件，仍然可以继续使用，但是对应功能缺失。
    '   LogLevel_Info：该日志属于重要信息记录，用于重要信息的记录，如费用，交易等数据。
    '   LogLevel_Trace：该信息是程序的运行跟踪信息。用于跟踪程序运行，以便方便错误查证。
    'arrPars In ParamArray
    '   产生格式：strLogTilte: arrPars(0),arrPars(1),arrPars(2),arrPars(3)...
    '@备注
    If gobjLog Is Nothing Then
        If mblnCreateLog = True Then Exit Sub  '只创建一次.
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

