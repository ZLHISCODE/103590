Attribute VB_Name = "mdlPacsInterface"

Option Explicit

'######################################################################################################################

Public gcnOracle As New ADODB.Connection            '公共数据库连接
Public gstrSysName As String

'######################################################################################################################

Public Function OraDataOpen(ByVal strServerName As String, ByVal strUserName As String, ByVal strUserPwd As String) As Boolean
    '******************************************************************************************************************
    '功能： 打开指定的数据库
    '参数：
    '   strServerName：主机字符串
    '   strUserName：用户名
    '   strUserPwd：密码
    '返回： 数据库打开成功，返回true；失败，返回false
    '******************************************************************************************************************
    On Error GoTo errHand
    
    OraDataOpen = False
    DoEvents
        
    With gcnOracle
        If .State = adStateOpen Then .Close '
        .CursorLocation = adUseClient
        .ConnectionString = "Provider=OraOLEDB.Oracle;" & _
                            "Data Source=" & strServerName & ";" & _
                            "User ID=" & strUserName & ";" & _
                            "Password=" & TranPasswd(strUserPwd) & ";" & _
                            "PLSQLRSet=1;" & _
                            "Persist Security Info=True"
        .Open
        
    End With
        
    OraDataOpen = True
    
    Exit Function
errHand:
    If InStr(Err.Description, "自动化错误") > 0 Then
        Err.Description = "连接串无法创建，请检查数据访问部件是否正常安装。"
    ElseIf InStr(Err.Description, "ORA-12154") > 0 Then
        Err.Description = "无法分析服务器名，请检查在Oracle配置中是否存在该本地网络服务名（主机字符串）。"
    ElseIf InStr(Err.Description, "ORA-12541") > 0 Then
        Err.Description = "无法连接，请检查服务器上的Oracle监听器服务是否启动。"
    ElseIf InStr(Err.Description, "ORA-01033") > 0 Then
        Err.Description = "ORACLE正在初始化或在关闭，请稍候再试。"
    ElseIf InStr(Err.Description, "ORA-01034") > 0 Then
        Err.Description = "ORACLE不可用，请检查服务或数据库实例是否启动。"
    ElseIf InStr(Err.Description, "ORA-02391") > 0 Then
        Err.Description = "用户" & UCase(strUserName) & "已经登录，不允许重复登录(已达到系统所允许的最大登录数)。"
    ElseIf InStr(Err.Description, "ORA-01017") > 0 Then
        Err.Description = "由于用户、口令或服务器指定错误，无法登录。"
    ElseIf InStr(Err.Description, "ORA-28000") > 0 Then
        Err.Description = "由于用户已经被禁用，无法登录。"
    Else
        Err.Description = Err.Description
    End If
    
    Call Err.Raise(Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function


Public Function OraDataClose() As Boolean
    '******************************************************************************************************************
    '功能： 关闭数据库
    '参数：
    '返回： 关闭数据库，返回True；失败，返回False
    '******************************************************************************************************************
    On Error GoTo errHand
            
    gcnOracle.Close
    OraDataClose = True
    
    Exit Function
errHand:
    OraDataClose = False

End Function


Public Function TranPasswd(strOld As String) As String
    '******************************************************************************************************************
    '功能： 密码转换函数
    '参数：
    '   strOld：原密码
    '返回： 加密生成的密码
    '******************************************************************************************************************
    On Error Resume Next
    
    Dim iBit As Integer, strBit As String
    Dim strNew As String
    If Len(Trim(strOld)) = 0 Then TranPasswd = "": Exit Function
    strNew = ""
    For iBit = 1 To Len(Trim(strOld))
        strBit = UCase(Mid(Trim(strOld), iBit, 1))
        Select Case (iBit Mod 3)
        Case 1
            strNew = strNew & _
                Switch(strBit = "0", "W", strBit = "1", "I", strBit = "2", "N", strBit = "3", "T", strBit = "4", "E", strBit = "5", "R", strBit = "6", "P", strBit = "7", "L", strBit = "8", "U", strBit = "9", "M", _
                   strBit = "A", "H", strBit = "B", "T", strBit = "C", "I", strBit = "D", "O", strBit = "E", "K", strBit = "F", "V", strBit = "G", "A", strBit = "H", "N", strBit = "I", "F", strBit = "J", "J", _
                   strBit = "K", "B", strBit = "L", "U", strBit = "M", "Y", strBit = "N", "G", strBit = "O", "P", strBit = "P", "W", strBit = "Q", "R", strBit = "R", "M", strBit = "S", "E", strBit = "T", "S", _
                   strBit = "U", "T", strBit = "V", "Q", strBit = "W", "L", strBit = "X", "Z", strBit = "Y", "C", strBit = "Z", "X", True, strBit)
        Case 2
            strNew = strNew & _
                Switch(strBit = "0", "7", strBit = "1", "M", strBit = "2", "3", strBit = "3", "A", strBit = "4", "N", strBit = "5", "F", strBit = "6", "O", strBit = "7", "4", strBit = "8", "K", strBit = "9", "Y", _
                   strBit = "A", "6", strBit = "B", "J", strBit = "C", "H", strBit = "D", "9", strBit = "E", "G", strBit = "F", "E", strBit = "G", "Q", strBit = "H", "1", strBit = "I", "T", strBit = "J", "C", _
                   strBit = "K", "U", strBit = "L", "P", strBit = "M", "B", strBit = "N", "Z", strBit = "O", "0", strBit = "P", "V", strBit = "Q", "I", strBit = "R", "W", strBit = "S", "X", strBit = "T", "L", _
                   strBit = "U", "5", strBit = "V", "R", strBit = "W", "D", strBit = "X", "2", strBit = "Y", "S", strBit = "Z", "8", True, strBit)
        Case 0
            strNew = strNew & _
                Switch(strBit = "0", "6", strBit = "1", "J", strBit = "2", "H", strBit = "3", "9", strBit = "4", "G", strBit = "5", "E", strBit = "6", "Q", strBit = "7", "1", strBit = "8", "X", strBit = "9", "L", _
                   strBit = "A", "S", strBit = "B", "8", strBit = "C", "5", strBit = "D", "R", strBit = "E", "7", strBit = "F", "M", strBit = "G", "3", strBit = "H", "A", strBit = "I", "N", strBit = "J", "F", _
                   strBit = "K", "O", strBit = "L", "4", strBit = "M", "K", strBit = "N", "Y", strBit = "O", "D", strBit = "P", "2", strBit = "Q", "T", strBit = "R", "C", strBit = "S", "U", strBit = "T", "P", _
                   strBit = "U", "B", strBit = "V", "Z", strBit = "W", "0", strBit = "X", "V", strBit = "Y", "I", strBit = "Z", "W", True, strBit)
        End Select
    Next
    TranPasswd = strNew

End Function


Public Sub ShowSimpleMsg(ByVal strInfo As String)
    '******************************************************************************************************************
    '功能：
    '******************************************************************************************************************
    MsgBox strInfo, vbInformation, gstrSysName
    
End Sub


Public Function ZVal(ByVal varValue As Variant) As String
    '******************************************************************************************************************
    '功能：将0零转换为"NULL"串,在生成SQL语句时用
    '******************************************************************************************************************
    ZVal = IIf(Val(varValue) = 0, "NULL", Val(varValue))
End Function
    
    
Public Function SQLRecord(ByRef rs As ADODB.Recordset) As Boolean
    '******************************************************************************************************************
    '功能:
    '参数:
    '返回:
    '******************************************************************************************************************
    On Error GoTo errHand
    
    Set rs = New ADODB.Recordset
    
    With rs
        
        .Fields.Append "SQL", adVarChar, 300
        .Fields.Append "Trans", adTinyInt                   '1表示开始;2表示结束
        .Fields.Append "Custom", adTinyInt
        .Fields.Append "Parameter", adVarChar, 500
        
        .Open
    End With
    
    SQLRecord = True
    
    Exit Function
errHand:
    SQLRecord = False
    
End Function

Public Function SQLRecordAdd(ByRef rs As ADODB.Recordset, ByVal strSql As String, Optional ByVal intTrans As Integer = 0, Optional ByVal intCustom As Integer = 0, Optional ByVal strParameter As String = "") As Boolean
    '******************************************************************************************************************
    '功能:
    '参数:
    '返回:
    '******************************************************************************************************************
    On Error GoTo errHand
    
    rs.AddNew
    rs("SQL").value = strSql
    rs("Trans").value = intTrans
    rs("Custom").value = intCustom
    rs("Parameter").value = strParameter
    
    SQLRecordAdd = True
    
    Exit Function
errHand:
    SQLRecordAdd = False
End Function

Public Function SQLRecordExecute(ByVal rs As ADODB.Recordset, Optional ByVal strTitle As String, Optional ByVal blnHaveTrans As Boolean = True) As Boolean
    '******************************************************************************************************************
    '功能:
    '参数:
    '返回:
    '******************************************************************************************************************
    On Error GoTo errHand
    
    Dim blnTran As Boolean
    Dim intLoop As Integer
    Dim strSql As String
    
    If rs.RecordCount > 0 Then
        If Len(strTitle) = 0 Then strTitle = gstrSysName
        blnTran = True
        
        If blnHaveTrans Then gcnOracle.BeginTrans
        
        rs.MoveFirst
    
        For intLoop = 1 To rs.RecordCount
        
            strSql = CStr(rs("SQL").value)
            Call ExecuteProcedure(strSql, strTitle)
            
            rs.MoveNext
        Next
    
        If blnHaveTrans Then gcnOracle.CommitTrans
        blnTran = False
    End If
    
    SQLRecordExecute = True
    
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
    
    If blnTran And blnHaveTrans Then gcnOracle.RollbackTrans
End Function


