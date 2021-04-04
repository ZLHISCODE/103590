Attribute VB_Name = "mdlSCCA"
Option Explicit
'四川CA功能模块
Private mblnInit As Boolean         '是否已初始化成功
Private SCCA_Client As Object       '客户端证书部件
Private SCCA_Server As Object       '服务器端证书部件

Public Function SCCA_InitObj() As Boolean
    '证书部件初始化
        Dim progID As String
        
        On Error GoTo errH
        SCCA_InitObj = mblnInit
        If mblnInit Then Exit Function
        Set SCCA_Client = CreateObject("wsb.WsbClient")
        Set SCCA_Server = CreateObject("wsbServer.wsbServerClass")
        If SCCA_Client Is Nothing Then
            SCCA_InitObj = False
            MsgBoxEx "初始化CA" & "wsb.WsbClient控件失败!", vbInformation, gstrSysName
            Exit Function
        End If
        If SCCA_Server Is Nothing Then
            SCCA_InitObj = False
            MsgBoxEx "初始化CA" & "wsbServer.wsbServerClass控件失败!", vbInformation, gstrSysName
            Exit Function
        End If
        '初始化成功了
        SCCA_InitObj = True
    
        mblnInit = SCCA_InitObj
        Exit Function
errH:
    Call GetErrMsg(Erl())
End Function

Public Function SCCA_RegCert(arrCertInfo As Variant) As Boolean
        '功能：提供在HIS数据库中注册个人证书的必要信息,用于新发放或更换证书,,需要插入USB-Key
        '返回：arrCertInfo作为数组返回证书相关信息
        '      0-ClientSignCertCN:客户端签名证书公共名称(姓名),用于注册证书时程序验证身份
        '      1-ClientSignCertDN:客户端签名证书主题(每人唯一)
        '      2-ClientSignCertSN:客户端签名证书序列号(每证书唯一)
        '      3-ClientSignCert:客户端签名证书内容
        '      4-ClientEncCert:客户端加密证书内容
        '      5-签名图片文件名,空串表示没有签名图片
        
        Dim strKeyId As String, strCertUserName As String, strEncCert As String, strCertSn As String
        Dim strSigCert As String, i As Integer, strCACert As String, lngOk As Long
        Dim strPicData As String
        On Error GoTo errH
    
100     For i = LBound(arrCertInfo) To UBound(arrCertInfo)
102         arrCertInfo(i) = ""
        Next
    
104     If GetCertList(strCertUserName, strKeyId, strSigCert, strCertSn) Then
106         arrCertInfo(0) = strCertUserName
108         arrCertInfo(1) = GetCertDN(strSigCert)
110         arrCertInfo(2) = strCertSn
112         arrCertInfo(3) = strSigCert
124         SCCA_RegCert = True
        End If

        Exit Function
errH:
126     MsgBoxEx "获取证书信息失败！" & Err.Description, vbInformation, gstrSysName
End Function

Public Function GetCertDN(strCert As String) As String
    Dim strInfo As String
    Dim i As Long
    strInfo = SCCA_Client.SOF_GetCertInfo(strCert, 33)
    If strInfo <> "" Then
        GetCertDN = strInfo
    Else
        Exit Function
    End If
End Function

''' 获取客户端证书列表
''' 返回boolean
Private Function GetCertList(ByRef strName As String, ByRef strUniqueID As String, ByRef strCert As String, ByRef strCertSn As String) As Boolean
    '-入参:无
    '-出参
    'strName :      保存接口返回的证书所有者姓名
    'strUniqueID:   保存接口返回的证书所有者唯一标识
    'strCert:       保存接口返回的签名证书
    'strCertSn      加密证书信息
    Dim strPassas As String
    Dim strList As String '已安装证书用户列表
    Dim arrList() As String
    On Error GoTo errH
    strList = SCCA_Client.SOF_GetUserList()
    If Trim(strList) <> "" Then
        strList = Replace(strList, "||", "|")
        strList = Replace(strList, "&&&", "&")
        arrList = Split(strList, "|")
        '证书信息
        strCert = SCCA_Client.SOF_ExportUserCert(arrList(1)) '证书字符串
        If strCert <> "" Then
            strUniqueID = SCCA_Client.SOF_GetCertInfo(strCert, 53) '唯一标识
            strName = SCCA_Client.SOF_GetCertInfo(strCert, 23) '证书通用者名称
            strCertSn = SCCA_Client.SOF_GetCertInfo(strCert, 2) '证书序列号
        End If
        GetCertList = True
    Else
        MsgBoxEx "没有找到Key盘，请检查！", vbInformation, gstrSysName
        Exit Function
    End If
    Exit Function
errH:
    GetCertList = False
End Function

Public Function SCCA_CheckCert(ByVal strCurrCertSn As String, Optional ByRef strSigCert As String, Optional ByRef strCertSn As String, Optional ByRef blnReDo As Boolean) As Boolean
    '功能：读取USB进行设备初始化并登录
    Dim strKey As String, strPIN As String, strUserName As String, strDate As String
    Dim blnRet As Boolean, intDate As Date
    Dim udtUser As USER_INFO
    Dim intPoint As Integer
    Dim strArry() As String
    On Error GoTo errH
    If Not SCCA_InitObj() Then
        MsgBoxEx "部件未初始化！"
        Exit Function
    End If
    If Not GetCertList(strUserName, strKey, strSigCert, strCertSn) Then Exit Function
    '证书唯一标识
    intPoint = InStr(strKey, "F")
    If mUserInfo.strUserID = "" Then
        MsgBoxEx "您的身份证号为空,请联系管理员到人员管理中录入！", vbOKOnly + vbInformation, gstrSysName
        Exit Function
    ElseIf UCase(Mid(strKey, intPoint + 2)) <> mUserInfo.strUserID Then
        MsgBoxEx "您的身份证号：" & _
                   vbCrLf & vbTab & "【" & mUserInfo.strUserID & "】" & vbCrLf & _
                   "当前证书唯一标识:" & _
                   vbCrLf & vbTab & "【" & UCase(Mid(strKey, intPoint + 2)) & "】" & vbCrLf & _
                   "用户身份证号与当前证书唯一标识不相等,不能使用！", vbInformation, gstrSysName
        Exit Function
    End If
    If Not GetCertLogin(strSigCert, strCertSn, intDate) Then
        blnRet = False
    Else
        blnRet = True
    End If
    If blnRet Then
            '判断是否需要更新注册证书
            udtUser.strName = strUserName
            udtUser.strSignName = strUserName
            udtUser.strUserID = UCase(Mid(strKey, intPoint + 2)) 'SF+身份证号
            udtUser.strCertSn = strCertSn
            udtUser.strCertDN = GetCertDN(strSigCert)
            udtUser.strCert = strSigCert
            udtUser.strEncCert = ""
            udtUser.strCertID = strKey
            Call GetEndDate(strSigCert, strDate)
            If IsUpdateRegCert(udtUser, strDate, blnReDo) Then
                blnRet = True
            Else
                blnRet = False
            End If
        End If
     
        SCCA_CheckCert = blnRet
        Exit Function
errH:
124     MsgBoxEx "检查USBKEY失败！" & Err.Description, vbInformation, gstrSysName
End Function

Private Function GetCertLogin(ByVal strCert As String, ByVal strCertSn As String, ByVal dDate As Date) As Boolean
    '- 入参
    'strUniqueID : 证书唯一标识
    'strPassword : 证书密码
    'strWebserviceUrl:签名服务器地址，即为证书验证
    '- 出参
    'dDate       : 返回证书有效时间
    On Error GoTo errH
    Dim result As Boolean
    Dim strDate As String '证书有效期字符串
    Dim strArry() As String
    Dim lng时间 As Long
    Dim strLogin As String
    If SCCA_Client Is Nothing Then Set SCCA_Client = CreateObject("wsb.WsbClient.1")
    If SCCA_Server Is Nothing Then Set SCCA_Client = CreateObject("wsbServer.wsbServerClass.1")
    '证书安全登录
    'strLogin:不为空:成功
    'strLogin:为空:不成功
    strLogin = SCCA_Client.SOF_SignDataByP7(strCertSn, 1)
    '验证证书结果信息表示
    If strLogin <> "" Then
        '验证证书有效性
        result = SCCA_Client.SOF_ValidateCert(strCert)
        If result Then
            '获取客户端证书有效期截止时间
            Call GetEndDate(strCert, strDate)
            dDate = CDate(strDate)
            lng时间 = CheckValidaty(dDate)
            If lng时间 < 0 Then
                MsgBoxEx "您的证书已过期!"
                GetCertLogin = False
            ElseIf (lng时间 <= 30 And lng时间 > 0) And Not gblnShow Then
                MsgBoxEx "您的证书还有" & lng时间 & "天过期"
                gblnShow = True
                GetCertLogin = True
            Else
                GetCertLogin = True
            End If
        Else
            MsgBoxEx "验证证书失败！" & "SCCA_Client.GetCertInfo", vbInformation, gstrSysName
        End If
    Else
        MsgBoxEx "初始登陆错误！" & "SCCA_Client.SOF_Login", vbInformation, gstrSysName
    End If
    Exit Function
errH:
    MsgBoxEx "调用证书接口错误!" & Err.Description, vbInformation, gstrSysName
    GetCertLogin = False
End Function

Private Function GetEndDate(ByVal strCert As String, ByRef strDate As String)
    Dim strArry() As String
    strDate = SCCA_Client.SOF_GetCertInfo(strCert, 18)
    If strDate <> "" Then
        strArry = Split(strDate, " ")
        If InStr(strArry(0), "Jan") > 0 Then
            strDate = strArry(3) & "-01-" & strArry(1) & " " & strArry(2)
        ElseIf InStr(strArry(0), "Feb") > 0 Then
            strDate = strArry(3) & "-02-" & strArry(1) & " " & strArry(2)
        ElseIf InStr(strArry(0), "Mar") > 0 Then
            strDate = strArry(3) & "-03-" & strArry(1) & " " & strArry(2)
        ElseIf InStr(strArry(0), "Apr") > 0 Then
            strDate = strArry(3) & "-04-" & strArry(1) & " " & strArry(2)
        ElseIf InStr(strArry(0), "May") > 0 Then
            strDate = strArry(3) & "-05-" & strArry(1) & " " & strArry(2)
        ElseIf InStr(strArry(0), "Jun") > 0 Then
            strDate = strArry(3) & "-06-" & strArry(1) & " " & strArry(2)
        ElseIf InStr(strArry(0), "Jul") > 0 Then
            strDate = strArry(3) & "-07-" & strArry(1) & " " & strArry(2)
        ElseIf InStr(strArry(0), "Aug") > 0 Then
            strDate = strArry(3) & "-08-" & strArry(1) & " " & strArry(2)
        ElseIf InStr(strArry(0), "Sep") > 0 Then
            strDate = strArry(3) & "-09-" & strArry(1) & " " & strArry(2)
        ElseIf InStr(strArry(0), "Oct") > 0 Then
            strDate = strArry(3) & "-10-" & strArry(1) & " " & strArry(2)
        ElseIf InStr(strArry(0), "Nov") > 0 Then
            strDate = strArry(3) & "-11-" & strArry(1) & " " & strArry(2)
        ElseIf InStr(strArry(0), "Dec") > 0 Then
            strDate = strArry(3) & "-12-" & strArry(1) & " " & strArry(2)
        End If
    Else
        Exit Function
    End If
End Function

Public Function SCCA_Sign(ByVal strCurrCertSn As String, ByVal strSource As String, ByRef strSignData As String, ByRef strTimeStamp As String, ByRef blnReDo As Boolean) As Boolean
        '签名
        Dim strSigCert As String, strCertSn As String
        Dim blnCheck As Boolean
        Dim datTime As Date
        Dim strDate As String
        Dim udtUser As USER_INFO

        On Error GoTo errH
        blnCheck = SCCA_CheckCert("", strSigCert, strCertSn, blnReDo)
        If blnReDo Then Exit Function
        If blnCheck Then
            datTime = gobjComLib.zlDatabase.Currentdate()
            strDate = Format(datTime, "yyyyMMddhhmmss")
            strTimeStamp = Format(datTime, "yyyy-MM-dd HH:mm:ss")
            strSource = StringSHA1(strSource)
            strSignData = SCCA_Client.SOF_SignDataByP7(strCertSn, strSource)
            If strSignData <> "" Then
                 SCCA_Sign = True
            Else
                MsgBoxEx "签名失败！"
            End If
        Else
            MsgBoxEx "签名失败！", vbInformation, "电子签名部件"
        End If
        Exit Function
errH:
114     MsgBoxEx "签名失败！" & Err.Description, vbInformation, gstrSysName
End Function

Public Function SCCA_VerifySign(ByVal strSignData As String, ByVal strSource As String) As Boolean
        '验证签名
        Dim strTmp As String
        Dim str签名原文 As String
        On Error GoTo errH
        str签名原文 = SCCA_Server.SOF_GetP7SignDataInfo(strSignData, 1)
        strSource = StringSHA1(strSource)
        If str签名原文 = strSource Then
            strTmp = SCCA_Server.SOF_VerifySignedDataByP7(strSignData)
            If Val(strTmp) = 0 Then
                 MsgBoxEx "验证签名成功！", vbInformation, gstrSysName
            Else
                 MsgBoxEx "验证签名失败！", vbInformation, gstrSysName
            End If
        Else
            MsgBoxEx "签名原文与签名值中的原文不一致，请检查原文是否被修改过！", vbInformation, gstrSysName
        End If
        Exit Function
errH:
104     MsgBoxEx "验证签名失败！" & Err.Description, vbInformation, gstrSysName
End Function

Public Sub SCCA_UnloadObj()
    Set SCCA_Client = Nothing
    Set SCCA_Server = Nothing
    mblnInit = False
End Sub

