VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMspServiceLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mobjDataSQLite As zlDataSQLite.clsDataSQLite

Public Function OpenFile(ByVal strLogFile As String) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim blnNew As Boolean
    Dim strParentFolder As String
    Dim objFso As New FileSystemObject
    
    If mobjDataSQLite.Initialize Then
        
        If strLogFile <> "" And Dir(strLogFile) = "" Then
            blnNew = True
            
            strParentFolder = objFso.GetParentFolderName(strLogFile)
            If objFso.FolderExists(strParentFolder) = False Then
                Call objFso.CreateFolder(strParentFolder)
            End If
        
        End If
        
        If mobjDataSQLite.OpenSQLiteFile(strLogFile) Then
            
            If blnNew Then
                
                Call mobjDataSQLite.ExecuteSQL("Create Table zlmsp_runlog(ID Varchar(32),Log_Time Varchar(20),Log_Type Varchar(10),Log_Desc Varchar(200))")
                Call mobjDataSQLite.ExecuteSQL("Create Table zlmsp_para(Code Varchar(10),Content Varchar(1000))")
                
            End If

        End If

        
    End If
    
    OpenFile = True
    
End Function

Public Function CloseFile() As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    
    CloseFile = mobjDataSQLite.CloseSQLiteFile
        
End Function

Public Function WriteRunLog(ByVal strLogType As String, ByVal strLogDesc As String) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim strLogTime As String
    
    If Not (mobjDataSQLite Is Nothing) Then
        strLogTime = Format(Now, "yyyy-MM-dd HH:mm:ss")
        strLogDesc = Replace(strLogDesc, "'", "''")
        Call mobjDataSQLite.ExecuteSQL("Insert Into zlmsp_runlog(ID,Log_Time,Log_Type,Log_Desc) Values('" & zlCommFun.GetGUID & "','" & strLogTime & "','" & strLogType & "','" & strLogDesc & "')")
    End If
    
    WriteRunLog = True
    
End Function

Public Function WriteSendLog(ByVal strMsgSource As String, ByVal strMsgTarget As String, ByVal strMsgBody As String, Optional ByVal strMsgState As String = "发送成功") As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim strLogTime As String
    Dim strSQL As String
    
    If Not (mobjDataSQLite Is Nothing) Then
        strLogTime = Format(Now, "yyyy-MM-dd HH:mm:ss")
        strMsgBody = Replace(strMsgBody, "'", "''")
        
        strSQL = "Insert Into m_SendLog(ID,Log_Time,Msg_Source,Msg_Target,Msg_State,Log_Desc) Values("
        strSQL = strSQL & "'" & zlCommFun.GetGUID & "'"
        strSQL = strSQL & ",'" & strLogTime & "'"
        strSQL = strSQL & ",'" & strMsgSource & "'"
        strSQL = strSQL & ",'" & strMsgTarget & "'"
        strSQL = strSQL & ",'" & strMsgState & "'"
        strSQL = strSQL & ",'" & Replace(strMsgBody, "'", "''") & "')"
        
        Call mobjDataSQLite.ExecuteSQL(strSQL)
    End If
    
    WriteSendLog = True
    
End Function

Public Function EditRunLog(Optional ByVal strConditionMode As String = "", Optional ByVal rsCondition As ADODB.Recordset) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim strSQL As String
            
    If Not (mobjDataSQLite Is Nothing) Then
        
        Select Case UCase(strConditionMode)
        Case UCase("Insert")
            
        Case UCase("Delete")
            strSQL = "Delete From zlmsp_runlog"
            strSQL = strSQL & " Where Log_Time>='" & zlCommFun.GetCondition(rsCondition, "开始时间") & "'"
            strSQL = strSQL & " And Log_Time<='" & zlCommFun.GetCondition(rsCondition, "结束时间") & "'"
        Case UCase("DeleteAll")
            strSQL = "Delete From zlmsp_runlog"
        End Select
        
        EditRunLog = mobjDataSQLite.ExecuteSQL(strSQL)
                
    End If
        
End Function

Public Function EditPara(ByVal strParaCode As String, ByVal strParaValue As String) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim strSQL As String
    
    If Not (mobjDataSQLite Is Nothing) Then
                
        strSQL = "Delete From zlmsp_para Where Code='" & strParaCode & "'"
        EditPara = mobjDataSQLite.ExecuteSQL(strSQL)
        
        strSQL = "Insert Into zlmsp_para(Code,Content) Values("
        strSQL = strSQL & "'" & strParaCode & "'"
        strSQL = strSQL & ",'" & strParaValue & "')"
        
        EditPara = mobjDataSQLite.ExecuteSQL(strSQL)
    End If
        
End Function

Public Function GetPara(Optional ByVal strConditionMode As String = "", Optional ByVal rsCondition As ADODB.Recordset) As zlDataSQLite.SQLiteRecordset
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim strSQL As String
    Dim rs As zlDataSQLite.SQLiteRecordset
        
    If Not (mobjDataSQLite Is Nothing) Then
        
        Select Case UCase(strConditionMode)
        Case UCase("Filter")
            strSQL = "Select Code,Content From zlmsp_para Where Code='" & zlCommFun.GetCondition(rsCondition, "参数编号") & "'"
        Case Else
            strSQL = "Select Code,Content From zlmsp_para"
        End Select
        Set rs.DataSet = mobjDataSQLite.OpenRecordset(strSQL)
                
    End If
    
    GetPara = rs
    
End Function

Public Function GetRunLog(Optional ByVal strConditionMode As String = "", Optional ByVal rsCondition As ADODB.Recordset) As zlDataSQLite.SQLiteRecordset
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim strSQL As String
    Dim rs As zlDataSQLite.SQLiteRecordset
        
    If Not (mobjDataSQLite Is Nothing) Then
        
        Select Case UCase(strConditionMode)
        Case UCase("Filter")
            strSQL = "Select ID,Log_Time,Log_Type,Log_Desc From zlmsp_runlog"
            strSQL = strSQL & " Where Log_Time>='" & zlCommFun.GetCondition(rsCondition, "开始时间") & "'"
            strSQL = strSQL & " And Log_Time<='" & zlCommFun.GetCondition(rsCondition, "结束时间") & "'"
        Case Else
            strSQL = "Select ID,Log_Time,Log_Type,Log_Desc From zlmsp_runlog"
        End Select
        Set rs.DataSet = mobjDataSQLite.OpenRecordset(strSQL)
                
    End If
    
    GetRunLog = rs
    
End Function

Public Function EditSendLog(Optional ByVal strConditionMode As String = "", Optional ByVal rsCondition As ADODB.Recordset) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim strSQL As String
            
    If Not (mobjDataSQLite Is Nothing) Then
        
        Select Case UCase(strConditionMode)
        Case UCase("Insert")
            
        Case UCase("Delete")
            strSQL = "Delete From m_SendLog"
            strSQL = strSQL & " Where Log_Time>='" & zlCommFun.GetCondition(rsCondition, "开始时间") & "'"
            strSQL = strSQL & " And Log_Time<='" & zlCommFun.GetCondition(rsCondition, "结束时间") & "'"
        Case UCase("DeleteAll")
            strSQL = "Delete From m_SendLog"
        End Select
        
        EditSendLog = mobjDataSQLite.ExecuteSQL(strSQL)
                
    End If
        
End Function

Public Function GetSendLog(Optional ByVal strConditionMode As String = "", Optional ByVal rsCondition As ADODB.Recordset) As zlDataSQLite.SQLiteRecordset
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    
    Dim strSQL As String
    Dim rs As zlDataSQLite.SQLiteRecordset
        
    If Not (mobjDataSQLite Is Nothing) Then
        
        Select Case UCase(strConditionMode)
        Case UCase("Filter")
            strSQL = "Select ID,Log_Time,Msg_Source,Msg_Target,Msg_State,Log_Desc From m_SendLog"
            strSQL = strSQL & " Where Log_Time>='" & zlCommFun.GetCondition(rsCondition, "开始时间") & "'"
            strSQL = strSQL & " And Log_Time<='" & zlCommFun.GetCondition(rsCondition, "结束时间") & "'"
        Case Else
            strSQL = "Select ID,Log_Time,Msg_Source,Msg_Target,Msg_State,Log_Desc From m_SendLog"
        End Select
        Set rs.DataSet = mobjDataSQLite.OpenRecordset(strSQL)
        
    End If
    
    GetSendLog = rs
    
End Function


Private Sub Class_Initialize()
    Set mobjDataSQLite = New zlDataSQLite.clsDataSQLite
End Sub

Private Sub Class_Terminate()
    Call mobjDataSQLite.CloseSQLiteFile
    Set mobjDataSQLite = Nothing
End Sub
