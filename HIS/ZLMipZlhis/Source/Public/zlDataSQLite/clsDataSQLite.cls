VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDataSQLite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'######################################################################################################################
Private mblnOpened As Boolean
Private mobjSQLite As dhRichClient3.cConnection

Public Type SQLiteRecordset
    DataSet As cRecordset
End Type

'######################################################################################################################
Public Function Initialize() As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    
    On Error GoTo errHand
    
    mblnOpened = False
        
    Initialize = True
    Exit Function
    
errHand:
    MsgBox Err.Description
    
End Function

Public Function OpenSQLiteFile(ByVal strSQLiteFile As String) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    On Error GoTo errHand
    
    If strSQLiteFile <> "" And Dir(strSQLiteFile) = "" Then
        If mobjSQLite.CreateNewDB(strSQLiteFile) = False Then Exit Function
    End If
    
    OpenSQLiteFile = mobjSQLite.OpenDB(strSQLiteFile)
    
    mblnOpened = True
    
    Exit Function
    
errHand:
    MsgBox Err.Description
    
End Function

Public Function CloseSQLiteFile() As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    If Not (mobjSQLite Is Nothing) Then
        
        Set mobjSQLite = Nothing
        Set mobjSQLite = New dhRichClient3.cConnection
        
    End If
    
    mblnOpened = False
    
    CloseSQLiteFile = True
    
End Function

Public Function ExecuteSQL(ByVal strSQL As String) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    
    If mblnOpened = False Then Exit Function
    
    Call mobjSQLite.Execute(strSQL)
    
    ExecuteSQL = True
    
End Function

Public Function OpenRecordset(ByVal strSQL As String) As cRecordset
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    
    On Error GoTo errHand
    
    If mblnOpened = False Then Exit Function
    
    Set OpenRecordset = mobjSQLite.OpenRecordset(strSQL)
    
    Exit Function
    
errHand:
    MsgBox Err.Description
    
End Function

Public Function BeginTrans() As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    On Error GoTo errHand
    
    If mblnOpened = False Then Exit Function
    
    Call mobjSQLite.BeginTrans
    
'    Call mobjSQLite.Execute("BEGIN TRANSACTION")
    BeginTrans = True
    
    Exit Function
    
errHand:
    MsgBox Err.Description
    
End Function

Public Function CommitTrans() As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    On Error GoTo errHand
    
    If mblnOpened = False Then Exit Function
    
    Call mobjSQLite.CommitTrans
    
'    Call mobjSQLite.Execute("COMMIT TRANSACTION")
    CommitTrans = True
    
    Exit Function
    
errHand:
    MsgBox Err.Description
End Function


Public Function RollbackTrans() As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    On Error GoTo errHand
    
    If mblnOpened = False Then Exit Function
    
    Call mobjSQLite.RollbackTrans
'    Call mobjSQLite.Execute("ROLLBACK TRANSACTION")
    RollbackTrans = True
    
    Exit Function
    
errHand:
    MsgBox Err.Description
End Function

Private Sub Class_Initialize()
    Set mobjSQLite = New dhRichClient3.cConnection
End Sub

Private Sub Class_Terminate()
    mblnOpened = False
    
    If Not (mobjSQLite Is Nothing) Then Set mobjSQLite = Nothing
End Sub
