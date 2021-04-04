VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMspClientData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mobjDataSQLite As zlDataSQLite.clsDataSQLite
Private mstrSQL As String
Private mstrTitle As String
Private mstrError As String
Private mstrDataFile As String
Private mblnOpened As Boolean
Private mobjFso As FileSystemObject

'######################################################################################################################
Public Function Initialize(ByVal strDataFile As String) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    mstrDataFile = strDataFile
    mblnOpened = False
    
    Initialize = True
End Function

Public Function OpenDataFile() As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim blnNewDataFile As Boolean
    Dim strParentFolder As String
        
    If mblnOpened = False Then
    
        If mobjDataSQLite.Initialize Then
            If mstrDataFile <> "" And Dir(mstrDataFile) = "" Then
                blnNewDataFile = True
                Set mobjFso = New FileSystemObject
                strParentFolder = mobjFso.GetParentFolderName(mstrDataFile)
                If mobjFso.FolderExists(strParentFolder) = False Then
                    Call mobjFso.CreateFolder(strParentFolder)
                End If
                Set mobjFso = Nothing
            End If
            If mobjDataSQLite.OpenSQLiteFile(mstrDataFile) Then
                If blnNewDataFile Then Call CreateStruct
            End If
        End If
        
        mblnOpened = True
    
    End If
    
    OpenDataFile = True
End Function

Public Function CloseDataFile() As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    
    CloseDataFile = mobjDataSQLite.CloseSQLiteFile
    mblnOpened = False
    
End Function

Public Function InsertReceiveMessage(ByVal strReceiveText As String, ByVal strReceiveTopic As String, ByVal strReceiveLnkType As String, ByVal strReceiveLnkTitle As String, ByVal strReceiveLnkPara As String) As String
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim strTime As String
    Dim strKey As String
    
    On Error GoTo errHand
    
    If Not (mobjDataSQLite Is Nothing) Then
        strTime = Format(Now, "yyyy-MM-dd HH:mm:ss")
        strKey = zlCommFun.GetGUID
        mstrSQL = "Insert Into m_Receives(ID,receive_date,receive_topic,receive_text,receive_lnk_type,receive_lnk_title,receive_lnk_para,receive_read) Values("
        mstrSQL = mstrSQL & "'" & strKey & "'"
        mstrSQL = mstrSQL & ",'" & strTime & "'"
        mstrSQL = mstrSQL & ",'" & strReceiveText & "'"
        mstrSQL = mstrSQL & ",'" & strReceiveTopic & "'"
        mstrSQL = mstrSQL & ",'" & strReceiveLnkType & "'"
        mstrSQL = mstrSQL & ",'" & strReceiveLnkTitle & "'"
        mstrSQL = mstrSQL & ",'" & strReceiveLnkPara & "',0)"
        If mobjDataSQLite.ExecuteSQL(mstrSQL) Then
            InsertReceiveMessage = strKey
        End If
    End If
    
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
    mstrError = Err.Description
End Function

Public Function DeleteReceiveMessage(ByVal strKey As String) As Boolean
    '******************************************************************************************************************
    '功能：读取MsgEvent数据
    '参数：
    '返回：
    '******************************************************************************************************************
    
    On Error GoTo errHand
    
    If Not (mobjDataSQLite Is Nothing) Then
        mstrSQL = "Delete From m_Receives Where ID='" & strKey & "'"
        DeleteReceiveMessage = mobjDataSQLite.ExecuteSQL(mstrSQL)
    End If
    
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
    mstrError = Err.Description
End Function

Public Function UpdateReceiveMessageReaded(ByVal strKey As String) As Boolean
    '******************************************************************************************************************
    '功能：读取MsgEvent数据
    '参数：
    '返回：
    '******************************************************************************************************************
    
    On Error GoTo errHand
    
    If Not (mobjDataSQLite Is Nothing) Then
        mstrSQL = "Update m_Receives Set receive_read=1 Where ID='" & strKey & "'"
        UpdateReceiveMessageReaded = mobjDataSQLite.ExecuteSQL(mstrSQL)
    End If
    
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
    mstrError = Err.Description
End Function

Public Function IsUnReadMessge() As Boolean
    '******************************************************************************************************************
    '功能：读取MsgEvent数据
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim rsTmp As zlDataSQLite.SQLiteRecordset
    
    On Error GoTo errHand
    
    If Not (mobjDataSQLite Is Nothing) Then
                    
        mstrSQL = "Select 1 From m_Receives Where receive_read<>1"
        Set rsTmp.DataSet = mobjDataSQLite.OpenRecordset(mstrSQL)
        IsUnReadMessge = (rsTmp.DataSet.BOF = False)
        
    End If
        
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
    mstrError = Err.Description
End Function

Public Function ReadReceiveMessage(Optional ByVal strConditionMode As String = "", Optional ByVal rsCondition As ADODB.Recordset) As zlDataSQLite.SQLiteRecordset
    '******************************************************************************************************************
    '功能：读取MsgEvent数据
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim strStyle As String
    Dim strTemp As String
    Dim rsTmp As zlDataSQLite.SQLiteRecordset
    
    On Error GoTo errHand
    
    If Not (mobjDataSQLite Is Nothing) Then
            
        Select Case UCase(strConditionMode)
        '------------------------------------------------------------------------------------------------------------------
        Case UCase("FilterData")
            
            strStyle = zlCommFun.GetCondition(rsCondition, "FilterStyle")
            strTemp = zlCommFun.GetCondition(rsCondition, "FilterText")
                                        
            mstrSQL = "Select 0 As 选择,ID,receive_date,receive_topic,receive_text,receive_lnk_type,receive_lnk_title,receive_lnk_para,receive_read From m_Receives "
            mstrSQL = mstrSQL & " Where receive_date>='" & zlCommFun.GetCondition(rsCondition, "Start_Date") & " 00:00:00'"
            mstrSQL = mstrSQL & " And receive_date<='" & zlCommFun.GetCondition(rsCondition, "End_Date") & " 23:59:59'"
            mstrSQL = mstrSQL & " Order By receive_date Desc"
            
            Select Case strStyle
            Case ""
                Set rsTmp.DataSet = mobjDataSQLite.OpenRecordset(mstrSQL)
            Case "编码"
            
            Case "名称"
    
            End Select
        End Select
    
    End If
    
    ReadReceiveMessage = rsTmp
    
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
    mstrError = Err.Description
End Function

Public Function EditPara(ByVal strParaCode As String, ByVal strParaValue As String) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim strSQL As String

    On Error GoTo errHand

    If Not (mobjDataSQLite Is Nothing) Then

        strSQL = "Delete From m_Options Where Para_No=" & Val(strParaCode)
        EditPara = mobjDataSQLite.ExecuteSQL(strSQL)

        strSQL = "Insert Into m_Options(Para_No,Para_Value) Values("
        strSQL = strSQL & Val(strParaCode)
        strSQL = strSQL & ",'" & strParaValue & "')"

        EditPara = mobjDataSQLite.ExecuteSQL(strSQL)
    End If
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
    mstrError = Err.Description
End Function

Public Function GetPara(Optional ByVal strConditionMode As String = "", Optional ByVal rsCondition As ADODB.Recordset) As zlDataSQLite.SQLiteRecordset
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim strSQL As String
    Dim rs As zlDataSQLite.SQLiteRecordset
    
    On Error GoTo errHand
    
    If Not (mobjDataSQLite Is Nothing) Then

        Select Case UCase(strConditionMode)
        Case UCase("Filter")
            strSQL = "Select Para_No,Para_Value From m_Options Where Para_No=" & GetCondition(rsCondition, "参数编号")
        Case Else
            strSQL = "Select Para_No,Para_Value From m_Options"
        End Select
        Set rs.DataSet = mobjDataSQLite.OpenRecordset(strSQL)

    End If

    GetPara = rs
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
    mstrError = Err.Description
End Function

'######################################################################################################################
Private Function CreateStruct() As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    
    On Error GoTo errHand
    
    mstrSQL = "Create Table m_Options(Para_No Number(5),Para_Value Varchar(1000),Para_Note Varchar(300))"
    Call mobjDataSQLite.ExecuteSQL(mstrSQL)
    
    mstrSQL = "Create Table m_Receives("
    mstrSQL = mstrSQL & "ID varchar(32),"
    mstrSQL = mstrSQL & "receive_date varchar(20),"
    mstrSQL = mstrSQL & "receive_topic Varchar(50),"
    mstrSQL = mstrSQL & "receive_text Varchar(100),"
    mstrSQL = mstrSQL & "receive_lnk_type Varchar(20),"
    mstrSQL = mstrSQL & "receive_lnk_title Varchar(50),"
    mstrSQL = mstrSQL & "receive_lnk_para Varchar(1000),"
    mstrSQL = mstrSQL & "receive_read Number(1)"
    mstrSQL = mstrSQL & ")"
    Call mobjDataSQLite.ExecuteSQL(mstrSQL)
    
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
    mstrError = Err.Description
End Function

Private Sub Class_Initialize()
    Set mobjDataSQLite = New zlDataSQLite.clsDataSQLite
End Sub

Private Sub Class_Terminate()
    Call mobjDataSQLite.CloseSQLiteFile
    Set mobjDataSQLite = Nothing
    mblnOpened = False
End Sub

