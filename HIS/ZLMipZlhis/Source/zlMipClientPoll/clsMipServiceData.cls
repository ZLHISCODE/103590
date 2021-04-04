VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMipServiceData"
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

Private Sub Class_Initialize()
    Set mobjDataSQLite = New zlDataSQLite.clsDataSQLite
End Sub

Private Sub Class_Terminate()
    Call mobjDataSQLite.CloseSQLiteFile
    Set mobjDataSQLite = Nothing
End Sub

