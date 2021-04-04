Attribute VB_Name = "mdlLogManager"
Option Explicit
'*********************************************************************************************************************************************
'功能:日志管理
'接口说明:
'   1.zlWritLog:写日志
'编制:刘兴洪
'日期:2019*01*25 15:14:00
'*********************************************************************************************************************************************
Public gobjLogManager As Object
Public gblnCreateLogManager As Boolean
Public Sub zlWritLog(ByVal lngModule As Long, ByVal strFunName As String, ByVal strCallFunName As String, _
    ByVal strLogInfor As String, Optional ByVal intLogType As Integer = 0, Optional strLogName As String = "一卡通接口调试日志", _
    Optional strGroupName As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:日志写入
    '入参:lngModule-当前模块号
    '     strCallFunName-调用者名称
    '     strFunName-功能名称
    '     intLogType-日志类型:0-正常日志;1-数据SQL;2-错误信息
    '     strLogInfor-写入的日志名称
    '     strLogName-日志名称
    '     strGroupName-组名
    '出参:
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-01-15 15:08:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objLogManager As Object
    On Error GoTo errHandle
    If zlGetLogManagerObject(objLogManager) = False Then
        Call LogWrite(strLogName, lngModule, strFunName, "调用者:" & strCallFunName & IIf(strGroupName = "", "", "-" & strGroupName) & vbTab & strLogInfor)
        Exit Sub
    End If
    If objLogManager Is Nothing Then Exit Sub
    Call gobjLogManager.zlWritLog(lngModule, strFunName, strCallFunName, strLogInfor, intLogType, strLogName)
    Set objLogManager = Nothing
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Public Sub zlWritLogEx(ByVal objCallMain As Object, ByVal lngModule As Long, ByVal strFunName As String, ByVal strLogClassify As String, _
    ByVal strLogInfor As String, Optional ByVal intLogType As Integer = 0, Optional strLogName As String = "电子票据调试日志", _
    Optional strGroupName As String, Optional strBusinessName As String = "电子票据")
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:日志写入
    '入参:objCallMain-调用者对象（可能是类，也可能是窗体）
    '     lngModule-当前模块号
    '     strLogType-函数名或方法名
    '     strLogClassify-日志类别，比如：开始，结束等
    '     intLogType-日志类型:0-正常日志;1-数据SQL;2-错误信息
    '     strLogInfor-写入的日志名称
    '     strLogName-日志名称
    '     strGroupName-组名
    '     strBusinessName-业务名称
    '出参:
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-01-15 15:08:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objLogManager As Object
    Dim strCallFunName As String, strLogInforEx As String
    On Error GoTo errHandle
    If Not objCallMain Is Nothing Then
        strCallFunName = App.ProductName & "." & TypeName(objCallMain) & "." & strFunName
    ElseIf InStr(strFunName, ".") > 0 Then
        strCallFunName = App.ProductName & "." & strFunName
    Else
        strCallFunName = App.ProductName & ".无法确定调用者." & strFunName
    End If
    
    '日志信息构成:
    ' 函数名 +　( strLogClassify  ）+ strLogName
    strLogInforEx = strFunName & "(" & strLogClassify & ")" & strLogInfor
    If zlGetLogManagerObject(objLogManager) = False Then
        Call LogWrite(strLogName, lngModule, strBusinessName, "调用者:" & strCallFunName & IIf(strGroupName = "", "", "-" & strGroupName) & vbTab & strLogInforEx)
        Exit Sub
    End If
    If objLogManager Is Nothing Then Exit Sub
    Call gobjLogManager.zlWritLog(lngModule, strFunName, strCallFunName, strLogInforEx, intLogType, strLogName)
    Set objLogManager = Nothing
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Public Function zlGetLogManagerObject(ByRef objLogManager As Object) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取费用公共部件对象
    '出参:objLogManager-返回日志管理部件对象
    '返回:获取返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-01-25 09:57:11
    '---------------------------------------------------------------------------------------------------------------------------------------------
    
    On Error GoTo errHandle
    If Not gobjLogManager Is Nothing Then
        Set objLogManager = gobjLogManager: zlGetLogManagerObject = True
        Exit Function
    End If
    If gblnCreateLogManager Or gcnOracle Is Nothing Then Exit Function  '只初始化一次,其他时候不用再初始化
    
    
    Err = 0: On Error Resume Next
    If gobjLogManager Is Nothing Then
        Set gobjLogManager = CreateObject("zlLogManager.clsLogManager")
        gblnCreateLogManager = True
        If Err <> 0 Then Exit Function
    End If
    
    Err.Clear:  On Error GoTo errHandle
    If gobjLogManager Is Nothing Then Exit Function
    
    'zlInitCommon(ByVal lngSys As Long, _
     ByVal cnOracle As ADODB.Connection, Optional ByVal strDbUser As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化相关的系统号及相关连接
    '入参:lngSys-系统号
    '     cnOracle-数据库连接对象
    '     strDBUser-数据库所有者
    '返回:初始化成功,返回true,否则返回False
    If gobjLogManager.InitCommon(gcnOracle, UserInfo.用户名) = False Then Exit Function
    Set objLogManager = gobjLogManager: zlGetLogManagerObject = True
    Exit Function
errHandle:
    Exit Function
End Function
 

