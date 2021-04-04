Attribute VB_Name = "mdlPubLog"
Option Explicit
'��־����ģ��
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
Public gstrStep As String   '��¼��ǰ���裬������־���
Private mblnDel31Day As Boolean '�Ƿ���ִ�й�ɾ����ǰ�ļ�

Public Sub ShowLog(lngLogSource As LOGSource, _
                   lngLogType As LOGTYPE, ByVal strTag, ByVal lngPid As Long, _
                   ByVal strInfo As String)
    Dim strErr As String, strShowString As String, strTitle As String
    Dim strIniFile As String, strLogFile As String
    If gSysParameter.LogLevel >= lngLogType Then
        'Ҫ��¼����ʾ��־
        strIniFile = App.Path & "\" & App.EXEName & ".ini"
        strLogFile = App.Path & "\Run_log.txt"
        
        If mObjLog.LogInit(App.EXEName, strErr, strIniFile, True, strLogFile) Then
            
            
            strTitle = LogSourceToString(lngLogSource) & "-" & strTag
            
            '��ʾ��־
            strShowString = mObjLog.FormatLogInfo(lngLogType, strTitle, lngPid, strInfo)
'            rtbLog = rtbLog.Text & strShowString
'            rtbLog.SelStart = Len(rtbLog.Text)
            
            'If Len(rtbLog.Text) > 8192 Then rtbLog.Text = ""
            strErr = ""
            'д��־
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
            '��ʼ����־����ʧ�ܣ�
'            rtbLog = rtbLog.Text & strErr
'            rtbLog.SelStart = Len(rtbLog.Text)
            
        End If
    End If
End Sub

Public Function LogSourceToString(ByVal lngLogSource As LOGSource) As String
    '����־��Դת��Ϊ�ı���ʽ
    If lngLogSource = LOG_LISCOM Then
        LogSourceToString = "���ݽ��ճ���"
    ElseIf lngLogSource = LOG_LISDEV Then
        LogSourceToString = "�ӿ�"
    ElseIf lngLogSource = LOG_LISWORK Then
        LogSourceToString = "��ʦ����վ"
    ElseIf lngLogSource = LOG_RECEIV Then
        LogSourceToString = "ͨѶ����"
    ElseIf lngLogSource = LOG_PRINTSVR Then
        LogSourceToString = "��ӡ�ӿ�"
    Else
        LogSourceToString "����"
    End If
End Function

Public Function LogTypeToString(ByVal lngLogType As LOGTYPE) As String
    '����־����ת��Ϊ�ı���ʽ
    
    If lngLogType = LOG_DEBUG Then
        LogTypeToString = "����"
    ElseIf lngLogType = LOG_ERR Then
        LogTypeToString = "����"
    ElseIf lngLogType = LOG_INFO Then
        LogTypeToString = "��ʾ"
    ElseIf lngLogType = LOG_WARNING Then
        LogTypeToString = "����"
    Else
        LogTypeToString = "����"
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






