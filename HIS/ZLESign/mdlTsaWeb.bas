Attribute VB_Name = "mdlTsaWeb"
Option Explicit
'联合信任时间戳

Private mobjTSA As Object       '用于准格尔医院的时间戳接口

Public Function TSA_initObj() As Boolean
    On Error Resume Next
    Set mobjTSA = Nothing
    Set mobjTSA = CreateObject("tsaMiddleware.UtilUdp")
    If Err.Number <> 0 Then
        'MsgboxEx "时间戳控件没有安装！", vbExclamation, gstrSysName
        Exit Function
    End If
    TSA_initObj = True
End Function

Public Function TSA_UnloadObj()
    '釋放對象
    If Not mobjTSA Is Nothing Then Set mobjTSA = Nothing
End Function

Private Function GetReturnInfo(ByVal strSign As String) As String
    '时间戳返回信息转换函数
    If strSign = "0001" Then
        GetReturnInfo = "网络通信异常"
    ElseIf strSign = "0002" Then
        GetReturnInfo = "系统异常"
    ElseIf strSign = "0003" Then
        GetReturnInfo = "系统繁忙"
    ElseIf strSign = "0004" Then
        GetReturnInfo = "传递参数不合法"
    ElseIf strSign = "0005" Then
        GetReturnInfo = "用户名或密码错误"
    ElseIf strSign = "0006" Then
        GetReturnInfo = "数据库异常"
    ElseIf strSign = "0007" Then
        GetReturnInfo = "DLL配置文件读取错误"
    ElseIf strSign = "1001" Then
        GetReturnInfo = "请求响应失败"
    ElseIf strSign = "1002" Then
        GetReturnInfo = "请求数据已加盖过时间戳"
    ElseIf strSign = "1003" Then
        GetReturnInfo = "请求数据等待加盖时间戳"
    ElseIf strSign = "2001" Then
        GetReturnInfo = "未申请时间戳"
    ElseIf strSign = "2002" Then
        GetReturnInfo = "校验失败"
    ElseIf strSign = "2010" Then
        GetReturnInfo = "验证成功"
    Else
        GetReturnInfo = strSign
    End If
    If GetReturnInfo <> "" Then
        GetReturnInfo = "时间戳接口返回提示：" & GetReturnInfo
    End If
End Function

Public Function Times_Tamp(ByVal strSource As String, ByRef strTimeStamp As String) As Boolean
        '取时间戳
        Dim intCount As Integer, strSign As String
        On Error GoTo hErr
        
        If mobjTSA Is Nothing Then Exit Function
        
100     strSign = mobjTSA.sendTimestamp(strSource, "sha1")
102     If strSign <> "1000" And strSign <> "1002" And strSign <> "1003" Then
104         strSign = GetReturnInfo(strSign)
106         MsgBoxEx "申请时间戳失败！" & strSign, vbExclamation, gstrSysName
            Times_Tamp = False
            Exit Function
        Else
108         intCount = 0
110         Do While intCount <= 100
112             strSign = mobjTSA.gettimestampinfo(strSource, "sha1")
                '签名有花点时间
114             If InStr(strSign, "#") > 0 Then
116                 strTimeStamp = Split(strSign, "#")(0)
118                 If IsDate(strTimeStamp) Then
120                     strTimeStamp = Format(CDate(strTimeStamp), "yyyy-MM-dd HH:mm:ss")
                        Times_Tamp = True
                        Exit Function
                    Else
122                     MsgBoxEx "获取的时间戳不是一个日期！" & strTimeStamp, vbExclamation, gstrSysName
                    End If
124             ElseIf strSign <> "1003" And strSign <> "2001" Then
126                 strSign = GetReturnInfo(strSign)
128                 MsgBoxEx "获取时间戳失败！" & strSign, vbExclamation, gstrSysName
                    Exit Function
                End If
130             intCount = intCount + 1
            Loop
        End If
132     Times_Tamp = True
        Exit Function
hErr:
134    MsgBoxEx "取时间戳-第" & CStr(Erl()) & "行," & Err.Description, vbExclamation, gstrSysName
End Function

Public Function verify_Timestamp(ByVal strSource As String) As Boolean
    '验证时间戳
    Dim strData As String
    If mobjTSA Is Nothing Then Exit Function
    strData = mobjTSA.verifyTimeStamp(strSource, "sha1")
    If strData <> "2010" Then
        MsgBoxEx "验证时间戳失败！" & GetReturnInfo(strData), vbExclamation, gstrSysName
        Exit Function
    End If
    verify_Timestamp = True
End Function

Private Function verify_getTimestamp(ByVal strSource As String) As String
    '获取时间戳  这个是我加的。
    Dim strData As String
    Dim strTimeStamp As String
    If mobjTSA Is Nothing Then Exit Function
    
    strData = mobjTSA.gettimestampinfo(strSource, "sha1")
    If strData = "2001" Then
        MsgBoxEx "获取验证时间戳失败！" & GetReturnInfo(strData), vbExclamation, gstrSysName
        verify_getTimestamp = "空"
        Exit Function
    End If
    
    If InStr(strData, "#") > 0 Then
        strTimeStamp = Split(strData, "#")(0)
        If IsDate(strTimeStamp) Then
            strTimeStamp = Format(CDate(strTimeStamp), "yyyy-MM-dd HH:mm:ss")
        Else
            MsgBoxEx "获取的时间戳不是一个日期！" & strTimeStamp, vbExclamation, gstrSysName
            verify_getTimestamp = "空"
            Exit Function
        End If
    End If
    verify_getTimestamp = strTimeStamp
    
End Function



