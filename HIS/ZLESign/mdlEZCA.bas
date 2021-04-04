Attribute VB_Name = "mdlEZCA"
Option Explicit

Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
'联合信任时间戳webservice方式
Private mobjTSA As Object
Private strUrl As String
Private userid As String
Private userkey As String
Private lngUSETSA As Long

Public Function TSAWEB_initObj() As Boolean
    On Error Resume Next
    Set mobjTSA = Nothing
    Set mobjTSA = CreateObject("MSSOAP.SoapClient30")
    If Err.Number <> 0 Then
        Exit Function
    End If
    strUrl = ReadIni("TSA", "URL", App.Path & "\TSA.ini")
    userid = ReadIni("TSA", "USERID", App.Path & "\TSA.ini")
    userkey = ReadIni("TSA", "USERKEY", App.Path & "\TSA.ini")
    lngUSETSA = ReadIni("TSA", "USETSA", App.Path & "\TSA.ini")
    If strUrl = "" Or userid = "" Or userkey = "" Then
        Err.Raise -1, , "TSA.ini文件不存在或配置错误！"
        Exit Function
    End If
    If lngUSETSA = 0 Then
        Exit Function  'TSA.ini文件中的USETSA值为0代表不启用时间戳
    End If
    mobjTSA.MSSoapInit ReadIni("TSA", "URL", App.Path & "\TSA.ini")
    TSAWEB_initObj = True
End Function

Public Function TSAWEB_UnloadObj()
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
    ElseIf strSign = "1001" Then
        GetReturnInfo = "请求响应失败"
    ElseIf strSign = "1002" Then
        GetReturnInfo = "请求数据已加盖过时间戳"
    ElseIf strSign = "1003" Then
        GetReturnInfo = "请求数据等待加盖时间戳"
    ElseIf strSign = "2000" Then
        GetReturnInfo = "验证成功"
    ElseIf strSign = "2001" Then
        GetReturnInfo = "未申请时间戳"
    ElseIf strSign = "2002" Then
        GetReturnInfo = "签名校验失败"
    ElseIf strSign = "3010" Then
        GetReturnInfo = "含时间参数验证成功"
    ElseIf strSign = "3020" Then
        GetReturnInfo = "含时间戳文件参数验证成功"
    ElseIf strSign = "3030" Then
        GetReturnInfo = "含时间参数和时间戳文件参验证成功"
    Else
        GetReturnInfo = strSign
    End If
    If GetReturnInfo <> "" Then
        GetReturnInfo = "时间戳接口返回提示：" & GetReturnInfo
    End If
End Function

Public Function TimesWEB_Tamp(ByVal strSource As String, ByRef strTimeStamp As String) As Boolean
        '取时间戳
        Dim intCount As Integer, strSign As String
        Dim sz返回信息
        On Error GoTo hErr
        
        If mobjTSA Is Nothing Then Exit Function
        
100     strSign = mobjTSA.applyTimeStamp(userid, userkey, "sha1", StringSHA1(strSource))(0)
102     If strSign <> "1000" And strSign <> "1002" And strSign <> "1003" Then
104         strSign = GetReturnInfo(strSign)
106         MsgBoxEx "申请时间戳失败！" & strSign, vbExclamation, gstrSysName
            TimesWEB_Tamp = False
            Exit Function
        Else
108         intCount = 0
110         Do While intCount <= 100
                sz返回信息 = mobjTSA.GetTimeStamp(userid, userkey, 1, "sha1", StringSHA1(strSource))
112             strSign = sz返回信息(0)
                '签名有花点时间
114             If strSign = 3010 Then
                    strTimeStamp = sz返回信息(1)
118                 If IsDate(strTimeStamp) Then
120                     strTimeStamp = Format(CDate(strTimeStamp), "yyyy-MM-dd HH:mm:ss")
                        TimesWEB_Tamp = True
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
132     TimesWEB_Tamp = True
        Exit Function
hErr:
134    MsgBoxEx "取时间戳-第" & CStr(Erl()) & "行," & Err.Description, vbExclamation, gstrSysName
End Function

Public Function verifyWEB_Timestamp(ByVal strSource As String) As Boolean
    '验证时间戳
    Dim strData As String
    If mobjTSA Is Nothing Then Exit Function
    strData = mobjTSA.verifyTimeStamp(userid, userkey, "sha1", StringSHA1(strSource))(0)
    If strData <> "2010" Then
        MsgBoxEx "验证时间戳失败！" & GetReturnInfo(strData), vbExclamation, gstrSysName
        Exit Function
    End If
    verifyWEB_Timestamp = True
End Function

Public Function verifyWEB_getTimestamp(ByVal strSource As String) As String
    '获取时间戳
    Dim strData As String
    Dim strTimeStamp As String
    If mobjTSA Is Nothing Then Exit Function
    
    strData = mobjTSA.GetTimeStamp(userid, userkey, 1, "sha1", StringSHA1(strSource))(0)
    strTimeStamp = mobjTSA.GetTimeStamp(userid, userkey, 1, "sha1", StringSHA1(strSource))(1)
    If strData = "2001" Then
        MsgBoxEx "获取验证时间戳失败！" & GetReturnInfo(strData), vbExclamation, gstrSysName
        verifyWEB_getTimestamp = "空"
        Exit Function
    End If

    If IsDate(strTimeStamp) Then
        strTimeStamp = Format(CDate(strData), "yyyy-MM-dd HH:mm:ss")
    Else
        MsgBoxEx "获取的时间戳不是一个日期！" & strData, vbExclamation, gstrSysName
        verifyWEB_getTimestamp = "空"
        Exit Function
    End If

    verifyWEB_getTimestamp = strTimeStamp
    
End Function

Public Function ReadIni(strItem As String, strKey As String, strPath As String) As String
    Dim GetStr As String
    On Error GoTo errH

    GetStr = VBA.String(128, 0)
    GetPrivateProfileString strItem, strKey, "", GetStr, 256, strPath
    GetStr = VBA.Replace(GetStr, VBA.Chr(0), "")
    ReadIni = GetStr
    Exit Function
errH:
    Err.Clear
    ReadIni = ""
End Function

