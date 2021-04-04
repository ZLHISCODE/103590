Attribute VB_Name = "mdlBJCA"
Option Explicit

Private mobjTSA As Object       '用于准格尔医院的时间戳接口
Private mobjAXCA As Object      '安信CA部件
Private mLastPWD As String      '缓存密码
Private mobjAXSVR As Object     '安信服务端部件

Public Function ZGRYY_initObj() As Boolean
    Err.Clear: On Error Resume Next
    Set mobjAXCA = Nothing
    Set mobjAXCA = CreateObject("AXSECURITY.AXSecurityCtrl.1")
    
    If Err.Number <> 0 Then
        MsgBoxEx "安信签名控件没有安装！", vbExclamation, gstrSysName
        Exit Function
    End If
    
    Set mobjAXSVR = CreateObject("AnXEgovCom.AnXEgovSOF")
    If Err.Number <> 0 Then
        MsgBoxEx "安信服务端组件没有安装！", vbExclamation, gstrSysName
        Exit Function
    End If
    
    Err.Clear: On Error Resume Next
    Set mobjTSA = Nothing
    Set mobjTSA = CreateObject("tsaMiddleware.UtilUdp")
    If Err.Number <> 0 Then
        MsgBoxEx "时间戳控件没有安装！", vbExclamation, gstrSysName
        Exit Function
    End If
    ZGRYY_initObj = True
End Function
Public Function ZGRYY_UnloadObj()
    '釋放對象
    If Not mobjTSA Is Nothing Then Set mobjTSA = Nothing
    If Not mobjAXCA Is Nothing Then Set mobjAXCA = Nothing
End Function

Public Function ZGRYY_RegCert(arrCertInfo As Variant) As Boolean
    
        
        Dim bSuccess As Boolean
        Dim strDn As String, strSN As String, strUser As String, i As Integer
        Dim strBase64 As String '保存吉林省医院接口返回的图片数据(问题号34527)
        Dim strImage  As String '保存吉林省医院接口返回的图片文件路径
        On Error GoTo errH
100     If mobjAXCA Is Nothing Then Set mobjAXCA = CreateObject("AXSECURITY.AXSecurityCtrl.1")

102     If mobjAXCA Is Nothing Then
104         MsgBoxEx "电子签名控件未正确安装。", vbExclamation, gstrSysName
            Exit Function
        End If
106     bSuccess = mobjAXCA.OpenCert(0, "", 1)
108     If (bSuccess) Then
110         strDn = mobjAXCA.GetCertInfo(1, "")
112         i = InStr(1, strDn, "O=")
114         strUser = Mid(strDn, 4, i - 6)
116         strSN = mobjAXCA.GetCertInfo(2, "")
        '  这是得到图片签章的
118         strBase64 = mobjAXCA.ReadFileFromKey(0, 2)
120         If strBase64 <> "" Then
122             strImage = App.Path & "\" & strSN & ".gif"
124             If Dir(strImage) <> "" Then Kill strImage
126             If Not mobjAXCA.B64DecodeSToFile(strBase64, strImage) Then strImage = ""
            Else
128             strImage = ""
            End If
        Else
130         MsgBoxEx mobjAXCA.GetLastError
            Exit Function
        End If

132     arrCertInfo(0) = strUser
134     arrCertInfo(1) = strDn
136     arrCertInfo(2) = strSN
        
138     arrCertInfo(5) = strImage
140     ZGRYY_RegCert = True
    Exit Function
errH:
142 MsgBoxEx "注册证书-第" & CStr(Erl()) & "行," & Err.Description, vbQuestion, "电子签名"
End Function

Public Function ZGRYY_CheckCert(ByVal strCurrCertSn As String) As Boolean
        '验证当前的USB
        Dim strPIN As String, blnClientChk As Boolean
        
        
        On Error GoTo hErr
100     If mLastPWD <> "" Then strPIN = mLastPWD

102     Call mobjAXCA.SetKeyPWD(2, strPIN)  '请根据用户处的ＵＳＢ型号修改第１个参数,取值为： 0:明华 1：海泰 2:飞天3000 3:握奇
104     mLastPWD = strPIN
106     blnClientChk = mobjAXCA.SetSignerCert(2, strCurrCertSn)
108     If Not blnClientChk Then
110         mLastPWD = ""
112         MsgBoxEx "当前数字证书验证失败。" & mobjAXCA.GetLastError, vbExclamation, gstrSysName
            Exit Function
        End If
        
114     ZGRYY_CheckCert = True
        Exit Function
hErr:
116     MsgBoxEx "检查证书-第" & CStr(Erl()) & "行," & Err.Description, vbExclamation, gstrSysName
End Function
Public Function ZGRYY_Sign(ByVal strCurrCertSn As String, ByVal strSource As String, ByRef strSignData As String, ByRef strTimeStamp As String) As Boolean
        '签名
        
    Dim strClientSignData As String '客户端签名后的数据
    Dim strGetTimeDate As String           '时间戳接口返回数据
    Dim intSvrVerfy    As Integer           '安信服务端P7验证
    On Error GoTo errH
100 If mobjAXCA Is Nothing Then Set mobjAXCA = CreateObject("AXSECURITY.AXSecurityCtrl.1")
    
102 If Not ZGRYY_CheckCert(strCurrCertSn) Then Exit Function
    
104 If Not AxServer_VerifyCert Then Exit Function   '服务端证书验证 2012-12-28
    
106 If Not Times_Tamp(strSource, strTimeStamp) Then Exit Function

108 strClientSignData = mobjAXCA.SignString(strSource, True)
    ' MsgboxEx strTimeStamp '不对，签名时没取到时间，验证时才取的时间戳。
110 If Len(strTimeStamp) = 0 Then
            '再得到时间戳
112     strGetTimeDate = verify_getTimestamp(strSource)
114     If strGetTimeDate <> "空" Then
116         strTimeStamp = Format(CDate(strGetTimeDate), "yyyy-MM-dd HH:mm:ss")
        End If
    End If
118 If strClientSignData = "" Then
120     MsgBoxEx mobjAXCA.GetLastError, vbExclamation, gstrSysName
        Exit Function
    End If
    
    '服务端签名验证，传原文还是客户端签名后的数据文档中没有明确，此处暂时传签名后数据  2012-12-28
    
122 intSvrVerfy = mobjAXSVR.SOF_VerifySignedDataByP7(strClientSignData)
124 If intSvrVerfy <> 0 Then
126     MsgBoxEx "安信服务端签名验证失败，错误码" & intSvrVerfy
        Exit Function
    End If
    
128 strSignData = strClientSignData
130 ZGRYY_Sign = False
    
132 ZGRYY_Sign = True
    Exit Function
errH:
134 MsgBoxEx "签名-第" & CStr(Erl()) & "行," & Err.Description, vbExclamation, gstrSysName
End Function

Public Function ZGRYY_VerifySign(ByVal strCurrCertSn As String, ByVal strSignData As String, ByVal strSource As String, ByVal strTime As String) As Boolean
    '验证签名
    Dim aaa As String
    Dim csdate As Date
    If mobjAXCA Is Nothing Then Set mobjAXCA = CreateObject("AXSECURITY.AXSecurityCtrl.1")
    
    If verify_Timestamp(strSource) = False Then
        MsgBoxEx "时间戳验证失败！", vbExclamation, gstrSysName
        Exit Function '时间戳验证
    End If
    If Not ZGRYY_CheckCert(strCurrCertSn) Then Exit Function
    
    ZGRYY_VerifySign = mobjAXCA.VerifyString(strSignData, True, False, strSource)
    If Not ZGRYY_VerifySign Then
        MsgBoxEx "签名验证失败！" & mobjAXCA.GetLastError, vbExclamation, gstrSysName
        Exit Function
    Else
        MsgBoxEx "签名验证成功！"
    End If
End Function
Private Function GetReturnInfo(ByVal strSign As String) As String
    '准格尔时间戳返回信息转换函数
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

Private Function Times_Tamp(ByVal strSource As String, ByRef strTimeStamp As String) As Boolean
        '取时间戳
        Dim intCount As Integer, strSign As String
        On Error GoTo hErr
    
100     strSign = mobjTSA.sendTimestamp(strSource, "sha1")
102     If strSign <> "1000" And strSign <> "1002" And strSign <> "1003" Then
104         strSign = GetReturnInfo(strSign)
106         MsgBoxEx "申请时间戳失败！" & strSign, vbExclamation, gstrSysName
            Times_Tamp = False
            Exit Function
        Else
108         intCount = 0
110         Do While intCount <= 10
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
124             ElseIf strSign <> "1003" Then
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
134     MsgBoxEx "取时间戳-第" & CStr(Erl()) & "行," & Err.Description, vbExclamation, gstrSysName
End Function

Private Function verify_Timestamp(ByVal strSource As String) As Boolean
    '验证时间戳
    Dim strData As String
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

Private Function AxServer_VerifyCert() As Boolean

    '验证当前USBKey的服务端证书验证
        
    Dim strBase64Cert As String, intCheck As Integer
    
    
    AxServer_VerifyCert = False
    If mobjAXSVR Is Nothing Then Set mobjAXSVR = CreateObject("AnXEgovCom.AnXEgovSOF")
    
    '读取USB中的证书base64编码字符串
    strBase64Cert = mobjAXCA.GetSignerCertInfo(5, "")
    
    '调用安信服务端证书验证功能
    intCheck = mobjAXSVR.SOF_ValidateCert(strBase64Cert)
    If intCheck <> 0 Then
        mLastPWD = ""
        MsgBoxEx "当前数字证书服务端验证失败！错误码为" & intCheck
        Exit Function
    End If
    
    AxServer_VerifyCert = True
    
End Function
