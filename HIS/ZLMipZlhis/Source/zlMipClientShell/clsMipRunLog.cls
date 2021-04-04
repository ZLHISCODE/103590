VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMipRunLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mblnDataOpened As Boolean
Private mobjDataSQLite As zlDataSQLite.clsDataSQLite
Private mstrLogFile As String

Public Function Initialize(ByVal strLogFile As String) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    mstrLogFile = strLogFile
    mblnDataOpened = False
    
    Initialize = True
End Function

Public Function OpenRunLogFile() As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim blnNew As Boolean
    Dim strParentFolder As String
    Dim objFso As New FileSystemObject
    
    If mobjDataSQLite.Initialize Then
        
        If mstrLogFile <> "" And Dir(mstrLogFile) = "" Then
            blnNew = True
            strParentFolder = objFso.GetParentFolderName(mstrLogFile)
            If objFso.FolderExists(strParentFolder) = False Then
                Call objFso.CreateFolder(strParentFolder)
            End If
        End If
        
        If mobjDataSQLite.OpenSQLiteFile(mstrLogFile) Then
            If blnNew Then
                Call mobjDataSQLite.ExecuteSQL("Create Table zlmsp_runlog(ID Varchar(32),Log_Time Varchar(20),Log_Type Varchar(10),Log_Desc Varchar(200))")
            End If
            
            mblnDataOpened = True
        End If
    End If

    OpenRunLogFile = mblnDataOpened
    
End Function

Public Function CloseRunLogFile() As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    
    CloseRunLogFile = mobjDataSQLite.CloseSQLiteFile
        
End Function

Public Function WriteRunLog(ByVal strLogType As String, ByVal strLogDesc As String) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim strLogTime As String
    
    If Not (mobjDataSQLite Is Nothing) Then
        strLogTime = Format(Now, "yyyy-MM-dd HH:mm:ss")
        strLogDesc = Replace(strLogDesc, "'", "''")
        Call mobjDataSQLite.ExecuteSQL("Insert Into zlmsp_runlog(ID,Log_Time,Log_Type,Log_Desc) Values('" & zlCommFun.GetGUID & "','" & strLogTime & "','" & strLogType & "','" & strLogDesc & "')")
    End If
    
    WriteRunLog = True
    
End Function

Public Function EditRunLog(Optional ByVal strConditionMode As String = "", Optional ByVal rsCondition As ADODB.Recordset) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim strSQL As String
            
    If Not (mobjDataSQLite Is Nothing) Then
        
        Select Case UCase(strConditionMode)
        Case UCase("Insert")
            
        Case UCase("Delete")
            strSQL = "Delete From zlmsp_runlog"
            strSQL = strSQL & " Where Log_Time>='" & zlCommFun.GetCondition(rsCondition, "��ʼʱ��") & "'"
            strSQL = strSQL & " And Log_Time<='" & zlCommFun.GetCondition(rsCondition, "����ʱ��") & "'"
        Case UCase("DeleteAll")
            strSQL = "Delete From zlmsp_runlog"
            
        Case UCase("DeleteID")
            strSQL = "Delete From zlmsp_runlog Where ID='" & zlCommFun.GetCondition(rsCondition, "ID") & "'"
        End Select
        
        EditRunLog = mobjDataSQLite.ExecuteSQL(strSQL)
                
    End If
        
End Function

Public Function GetRunLog(Optional ByVal strConditionMode As String = "", Optional ByVal rsCondition As ADODB.Recordset) As zlDataSQLite.SQLiteRecordset
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim strSQL As String
    Dim rs As zlDataSQLite.SQLiteRecordset
        
    If Not (mobjDataSQLite Is Nothing) Then
        
        Select Case UCase(strConditionMode)
        Case UCase("Filter")
            strSQL = "Select 0 As ѡ��,ID, case when Log_Type='��Ϣ' then 'info' when Log_Type='����' then 'error' end As ͼ��,Log_Time,Log_Type,Log_Desc From zlmsp_runlog"
            strSQL = strSQL & " Where Log_Time>='" & zlCommFun.GetCondition(rsCondition, "��ʼʱ��") & "'"
            strSQL = strSQL & " And Log_Time<='" & zlCommFun.GetCondition(rsCondition, "����ʱ��") & "'"
        Case Else
            strSQL = "Select ID,Log_Time,Log_Type,Log_Desc From zlmsp_runlog"
        End Select
        Set rs.DataSet = mobjDataSQLite.OpenRecordset(strSQL)
                
    End If
    
    GetRunLog = rs
    
End Function

Private Sub Class_Initialize()
    Set mobjDataSQLite = New zlDataSQLite.clsDataSQLite
End Sub

Private Sub Class_Terminate()
    Call mobjDataSQLite.CloseSQLiteFile
    Set mobjDataSQLite = Nothing
End Sub
