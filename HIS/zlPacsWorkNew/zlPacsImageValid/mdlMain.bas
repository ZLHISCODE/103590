Attribute VB_Name = "mdlMain"
Option Explicit
Public gcnOracle As ADODB.Connection    '定义公共连接
Public gobjComlib As Object
Public glngState As Long
Public gstrUserName As String
Public gstrServer As String
Public gstrPwd As String
Public gstrDept As String
Public gstrOpen As String
Public glngTime As Long
Public gobjLogin As Object

Sub Main()
    Dim strCmdLine As String
    Dim strOpen As String
    
    
    
    strCmdLine = Command()
    
    If Len(strCmdLine) = 0 Then
        
        If Not InitLogin Then
            frmUserLogin.Show
        Else
            On Error Resume Next
            
            Set gcnOracle = gobjLogin.login
            
            If Not gcnOracle Is Nothing Then
                frmMain.zlShowMe strCmdLine
            Else
                Err.Clear
                frmUserLogin.Show
            End If
        End If

    Else
        Call GetState(strCmdLine)
        If Not MyLogin Then Exit Sub
    End If
    
    If gcnOracle Is Nothing Then
        Exit Sub
    Else
        frmMain.zlShowMe strCmdLine
    End If
End Sub

Private Function InitLogin() As Boolean
    On Error Resume Next
    
    Set gobjLogin = CreateObject("ZLLogin.clsLogin")
   
    If Err <> 0 Then
        InitLogin = False
    Else
        InitLogin = True
    End If
    Err.Clear
End Function

Private Sub GetState(strCommand As String)
    Dim arrPara() As String
    
    arrPara = Split(strCommand, "||")
    
    If UBound(arrPara) >= 6 Then
        gstrServer = arrPara(0)
        gstrUserName = arrPara(1)
        gstrPwd = arrPara(2)
        gstrDept = arrPara(3)
        glngTime = arrPara(4)
        gstrOpen = arrPara(5)
        glngState = Val(arrPara(6))
    End If
End Sub


'定义公共判断连接方法
Public Function myConn(strUser As String, strKey As String, strServe As String) As Boolean
    Dim strOpen As String
    Dim strError As String
    On Error Resume Next
    
    Set gcnOracle = New ADODB.Connection
'    strOpen = "Provider=MSDAORA.1;Password=" & strKey & ";User ID=" & strUser & ";Data Source=" & strServe & ";Persist Security Info=True"
    gcnOracle.Open "Driver={Microsoft ODBC for Oracle};Server=" & strServe, strUser, strKey
    If Err <> 0 Then
        '保存错误信息
        strError = Err.Description
        If InStr(strError, "自动化错误") > 0 Then
            MsgBox "连接串无法创建，请检查数据访问部件是否正常安装。", vbInformation, "ZLSOFT"
        ElseIf InStr(strError, "ORA-12154") > 0 Then
            MsgBox "无法分析服务器名，" & vbCrLf & "请检查在Oracle配置中是否存在该本地网络服务名（主机字符串）。", vbInformation, "ZLSOFT"
        ElseIf InStr(strError, "ORA-12541") > 0 Then
            MsgBox "无法连接，请检查服务器上的Oracle监听器服务是否启动。", vbInformation, "ZLSOFT"
        ElseIf InStr(strError, "ORA-01033") > 0 Then
            MsgBox "ORACLE正在初始化或在关闭，请稍候再试。", vbInformation, "ZLSOFT"
        ElseIf InStr(strError, "ORA-01034") > 0 Then
            MsgBox "ORACLE不可用，请检查服务或数据库实例是否启动。", vbInformation, "ZLSOFT"
        ElseIf InStr(strError, "ORA-02391") > 0 Then
            MsgBox "用户" & UCase(strUser) & "已经登录，不允许重复登录(已达到系统所允许的最大登录数)。", vbExclamation, "ZLSOFT"
        ElseIf InStr(strError, "ORA-01017") > 0 Then
            MsgBox "由于用户、口令或服务器指定错误，无法登录。", vbInformation, "ZLSOFT"
        ElseIf InStr(strError, "ORA-28000") > 0 Then
            MsgBox "由于用户已经被禁用，无法登录。", vbInformation, "ZLSOFT"
        Else
            MsgBox strError, vbInformation, "ZLSOFT"
        End If
        
        Exit Function
    End If
'    gcnOracle.Open strOpen
    gcnOracle.CursorLocation = adUseClient
    myConn = Err.Number = 0
    
    Exit Function
errH:
    myConn = False
End Function

Private Function MyLogin() As Boolean
    Dim strOpen As String

    If Len(gstrOpen) = 0 Then
        MyLogin = myConn(gstrUserName, gstrPwd, gstrServer)
    Else
        Set gcnOracle = New ADODB.Connection
        strOpen = gstrOpen
        gcnOracle.Open strOpen
        gcnOracle.CursorLocation = adUseClient
        MyLogin = True
    End If
End Function

Public Sub getUser(strTmp As String)
    Dim arrTmp() As String
    
    arrTmp = Split(strTmp, "User ID=")
    If UBound(arrTmp) > 0 Then
        gstrUserName = Split(arrTmp(1), ";")(0)
    End If
    
End Sub
