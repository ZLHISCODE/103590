Attribute VB_Name = "mdlPubLog"
Option Explicit
'日志保存模块
Public Enum LOGTYPE
    LOG_ERR = 3
    LOG_WARNING = 4
    LOG_INFO = 6
    LOG_DEBUG = 7
End Enum
Public Enum LOGSource
    LOG_LISDEV = 16
    LOG_RECEIV = 17
    LOG_LISWORK = 18
    LOG_LISCOM = 19
    LOG_PRINTSVR = 20
End Enum
Public Type LOGITEM
    LOG_Key As String
    LOG_DATE As Date
    LOG_SOURCE As LOGSource
    LOG_TYPE As LOGTYPE
    LOG_TAG As String
    LOG_PID As Long
    LOG_MSG As String
    LOG_IP  As String
    LOG_ID  As Long
End Type

Private mObjLog As New clsLog
Public gstrStep As String   '记录当前步骤，用于日志输出
Private mblnDel31Day As Boolean '是否已执行过删除以前文件

Public Sub ShowLog(lngLogSource As LOGSource, _
                   lngLogType As LOGTYPE, ByVal strTag, ByVal lngPid As Long, _
                   ByVal strInfo As String)
    Dim strErr As String, strShowString As String, strTitle As String
    Dim strIniFile As String, strLogFile As String
    If gSysParameter.LogLevel >= lngLogType Then
        '要记录与显示日志
        strIniFile = App.Path & "\" & App.EXEName & ".ini"
        strLogFile = App.Path & "\Run_log.txt"
        
        If mObjLog.LogInit(App.EXEName, strErr, strIniFile, True, strLogFile) Then
            
            
            strTitle = LogSourceToString(lngLogSource) & "-" & strTag
            
            '显示日志
            strShowString = mObjLog.FormatLogInfo(lngLogType, strTitle, lngPid, strInfo)
'            rtbLog = rtbLog.Text & strShowString
'            rtbLog.SelStart = Len(rtbLog.Text)
            
            'If Len(rtbLog.Text) > 8192 Then rtbLog.Text = ""
            strErr = ""
            '写日志
            If lngLogType = LOG_DEBUG Then
                mObjLog.LogDebug strTitle, lngPid, strInfo, strErr
            ElseIf lngLogType = LOG_INFO Then
                mObjLog.LogInfo strTitle, lngPid, strInfo, strErr
            ElseIf lngLogType = LOG_WARNING Then
                mObjLog.LogWarn strTitle, lngPid, strInfo, strErr
            ElseIf lngLogType = LOG_ERR Then
                mObjLog.LogError strTitle, lngPid, strInfo, strErr
            End If
            If strErr <> "" Then
                strShowString = mObjLog.FormatLogInfo(lngLogType, strTitle, lngPid, strErr)
'                rtbLog = rtbLog.Text & strShowString
'                rtbLog.SelStart = Len(rtbLog.Text)
            End If
        Else
            '初始化日志部件失败！
'            rtbLog = rtbLog.Text & strErr
'            rtbLog.SelStart = Len(rtbLog.Text)
            
        End If
    End If
End Sub

Public Function LogSourceToString(ByVal lngLogSource As LOGSource) As String
    '将日志来源转换为文本方式
    If lngLogSource = LOG_LISCOM Then
        LogSourceToString = "数据接收程序"
    ElseIf lngLogSource = LOG_LISDEV Then
        LogSourceToString = "接口"
    ElseIf lngLogSource = LOG_LISWORK Then
        LogSourceToString = "技师工作站"
    ElseIf lngLogSource = LOG_RECEIV Then
        LogSourceToString = "通讯程序"
    ElseIf lngLogSource = LOG_PRINTSVR Then
        LogSourceToString = "打印接口"
    Else
        LogSourceToString "其他"
    End If
End Function

Public Function LogTypeToString(ByVal lngLogType As LOGTYPE) As String
    '将日志类型转换为文本方式
    
    If lngLogType = LOG_DEBUG Then
        LogTypeToString = "调试"
    ElseIf lngLogType = LOG_ERR Then
        LogTypeToString = "错误"
    ElseIf lngLogType = LOG_INFO Then
        LogTypeToString = "提示"
    ElseIf lngLogType = LOG_WARNING Then
        LogTypeToString = "警告"
    Else
        LogTypeToString = "其他"
    End If
        
End Function

'-----

Private Function GetFreeSpace(ByVal strPath As String) As Double
    Dim strDriv As String, Drv As Drive
    Dim strDir As String
    
    If gobjFSO.FolderExists(strPath) Then
        strDriv = gobjFSO.GetDriveName(gobjFSO.GetAbsolutePathName(strPath))
        Set Drv = gobjFSO.GetDrive(strDriv)
        If Drv.IsReady Then
            GetFreeSpace = Drv.FreeSpace
        End If
        Set Drv = Nothing
    End If
End Function






