Attribute VB_Name = "mdlSHCA"
Option Explicit
'上海CA中心功能模块
Private mblnInit As Boolean         '是否已初始化成功
Private mLastPWD As String          '缓存输入的密码

Private SHCA_Client As Object       '证书部件
Private mLogin As Long              '输入密码错误次数

Public Enum SH_Version
    V_SEH = 0
    V_ESE = 1
End Enum

Public Function SHCA_InitObj() As Boolean
    '证书部件初始化
        Dim progID As String
        
        On Error GoTo errH
100
102     SHCA_InitObj = mblnInit
104     If mblnInit Then Exit Function
105     mLastPWD = ""
        If Not SHCA_GetPar(1) Then Exit Function
108     Set SHCA_Client = CreateObject("SafeEngineCOM.SafeEngineCtl")
        If gudtPara.bytSignVersion = V_SEH Then
            Call SHCA_Client.SEH_InitialSession(2, "", "", 0, 2, "", "") '初始化CA接口
        Else
            Call SHCA_Client.ESE_InitialSession(2, "", "", 0, 2, "", "") '初始化CA接口
        End If
        If SHCA_Client.errorCode <> 0 Then
            GoTo errH
        End If
114     SHCA_InitObj = True
    
116     mblnInit = SHCA_InitObj
        mLogin = 0
        Exit Function
errH:
118     MsgBoxEx "创建接口部件失败！" & vbNewLine & Err.Description, vbQuestion, gstrSysName
End Function

Public Function SHCA_RegCert(arrCertInfo As Variant) As Boolean
        '功能：提供在HIS数据库中注册个人证书的必要信息,用于新发放或更换证书,,需要插入USB-Key
        '返回：arrCertInfo作为数组返回证书相关信息
        '      0-ClientSignCertCN:客户端签名证书公共名称(姓名),用于注册证书时程序验证身份
        '      1-ClientSignCertDN:客户端签名证书主题(每人唯一)
        '      2-ClientSignCertSN:客户端签名证书序列号(每证书唯一)
        '      3-ClientSignCert:客户端签名证书内容
        '      4-ClientEncCert:客户端加密证书内容
        '      5-签名图片文件名,空串表示没有签名图片
        
        Dim strKeyId As String, strCertTime As String, strCertUserName As String, strCertDN As String, strCertSn As String
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
124         SHCA_RegCert = True
        End If

        Exit Function
errH:
126     MsgBoxEx "获取证书信息失败！" & vbNewLine & "第" & CStr(Erl()) & "行 " & Err.Description, vbQuestion, gstrSysName

End Function

Public Function GetCertDN(strCert As String) As String
    Dim strCertDN As String
    Dim strCN As String, strO As String, strOU As String, strS As String, strL As String, strC As String, strE As String
    If gudtPara.bytSignVersion = V_SEH Then
        strC = SHCA_Client.SEH_GetCertDetail(strCert, 13)
        strO = SHCA_Client.SEH_GetCertDetail(strCert, 14)
        strOU = SHCA_Client.SEH_GetCertDetail(strCert, 15)
        strS = SHCA_Client.SEH_GetCertDetail(strCert, 16)
        strCN = SHCA_Client.SEH_GetCertDetail(strCert, 17)
        strL = SHCA_Client.SEH_GetCertDetail(strCert, 18)
    Else
        strC = SHCA_Client.ESE_GetCertDetail(strCert, 13)
        strO = SHCA_Client.ESE_GetCertDetail(strCert, 14)
        strOU = SHCA_Client.ESE_GetCertDetail(strCert, 15)
        strS = SHCA_Client.ESE_GetCertDetail(strCert, 16)
        strCN = SHCA_Client.ESE_GetCertDetail(strCert, 17)
        strL = SHCA_Client.ESE_GetCertDetail(strCert, 18)
        strE = SHCA_Client.ESE_GetCertDetail(strCert, 19)
    End If
    strCertDN = IIf(strS = "", "", "S=" & strS & ",") & IIf(strL = "", "", "L=" & strL & ",") & IIf(strO = "", "", "O=" & strO & ",") _
    & IIf(strOU = "", "", "OU=" & strOU & ",") & IIf(strCN = "", "", "CN=" & strCN & ",") & IIf(strE = "", "", "E=" & strE)
    GetCertDN = strCertDN
End Function
Public Function SHCA_Sign(ByVal strCurrCertSn As String, ByVal strSource As String, ByRef strSignData As String, ByRef strTimeStamp As String, ByRef blnReDo As Boolean) As Boolean
        '签名
        Dim strSigCert As String
        Dim blnCheck As Boolean
        Dim datTime As Date
        Dim strDate As String

        On Error GoTo errH
        blnCheck = SHCA_CheckCert("", "", blnReDo)
        If blnReDo Then Exit Function
        If blnCheck Then
            '证书ID进行签名
            datTime = gobjComLib.zlDatabase.Currentdate()
            strDate = Format(datTime, "yyyyMMddhhmmss")
            strTimeStamp = Format(datTime, "yyyy-MM-dd HH:mm:ss")
            If gudtPara.bytSignVersion = V_SEH Then
                strSignData = SHCA_Client.SEH_SignData(strSource, 3)
            Else
                strSignData = SHCA_Client.ESE_SignData(strSource, "")
            End If
            If strSignData <> "" And SHCA_Client.errorCode = 0 Then
                 SHCA_Sign = True
            Else
                MsgBoxEx "签名失败！" & ValidateCertView(SHCA_Client.errorCode)
            End If
        Else
            If mLastPWD = "" Then
                Exit Function
            Else
                MsgBoxEx "签名失败！", vbInformation, "电子签名部件"
            End If
        End If
        Exit Function
errH:
114     MsgBoxEx "签名失败！" & vbNewLine & "第" & CStr(Erl()) & "行 " & Err.Description, vbQuestion, gstrSysName
End Function

Public Function SHCA_VerifySign(ByVal strCurrCertSn As String, ByVal strSignData As String, ByVal strSource As String, ByVal strSignCert As String) As Boolean
        '验证签名
        Dim strTmp As String
        On Error GoTo errH
102     If gudtPara.bytSignVersion = V_SEH Then
            Call SHCA_Client.SEH_InitialSession(2, "", "", 0, 2, "", "") '初始化CA接口
            Call SHCA_Client.SEH_VerifySignData(strSource, 3, strSignData, strSignCert)
            If SHCA_Client.errorCode <> 0 Then
                '兼容老版
                Call SHCA_Client.ESE_InitialSession(2, "", "", 0, 2, "", "") '初始化CA接口
                Call SHCA_Client.ESE_VerifySignData(strSource, "", strSignData, strSignCert)
            End If
        Else
            Call SHCA_Client.ESE_InitialSession(2, "", "", 0, 2, "", "") '初始化CA接口
            Call SHCA_Client.ESE_VerifySignData(strSource, "", strSignData, strSignCert)
            If SHCA_Client.errorCode <> 0 Then
                '兼容新版
                Call SHCA_Client.SEH_InitialSession(2, "", "", 0, 2, "", "") '初始化CA接口
                Call SHCA_Client.SEH_VerifySignData(strSource, 3, strSignData, strSignCert)
            End If
        End If
        If SHCA_Client.errorCode = 0 Then
             MsgBoxEx "验证签名成功！"
        Else
             MsgBoxEx "验证签名失败！" & ValidateCertView(SHCA_Client.errorCode)
        End If
        Exit Function
errH:
104     MsgBoxEx "验证签名失败！" & vbNewLine & "第" & CStr(Erl()) & "行 " & Err.Description, vbQuestion, gstrSysName
End Function

Public Function SHCA_GetPar(Optional ByVal bytFunc As Byte)
    Dim arrList As Variant
    
    On Error GoTo errH
    gstrPara = gobjComLib.zlDatabase.GetPara(90000, glngSys)  '读取URLs 固定读取ZLHIS'gstrPara = "0&&&0"   '0-SEH:1-ESE
    If Val(gstrPara) = 1 Then
        gudtPara.bytSignVersion = V_ESE
    ElseIf Val(gstrPara) = 0 Then
        gudtPara.bytSignVersion = V_SEH
    End If
    SHCA_GetPar = True
    Exit Function
errH:
    MsgBoxEx "读取参数失败！" & vbNewLine & "第" & CStr(Erl()) & "行 " & Err.Description, vbInformation, gstrSysName
End Function

Public Function SHCA_SetParaStr() As String
    SHCA_SetParaStr = IIf(gudtPara.bytSignVersion = 0, "0", "1")
End Function

Public Function SHCA_CheckCert(ByVal strCurrCertSn As String, Optional ByRef strSigCert As String, Optional ByRef blnReDo As Boolean) As Boolean
        '功能：读取USB进行设备初始化并登录
        Dim strKey As String, strPIN As String, strUserName As String, strCertSn As String, strDate As String
        Dim strWebUrl As String, intDate   As Integer
        Dim blnRet As Boolean
        Dim udtUser As USER_INFO
        Dim intPoint As Integer
        On Error GoTo errH
        If Not SHCA_InitObj() Then
102         MsgBoxEx "部件未初始化！"
            Exit Function
        End If
104     If Not GetCertList(strUserName, strKey, strSigCert, strCertSn) Then Exit Function
        intPoint = InStr(strKey, "F")
        If mUserInfo.strUserID = "" Then
            MsgBoxEx "您的身份证号为空,请联系管理员到人员管理中录入！", vbOKOnly + vbInformation, gstrSysName
            Exit Function
        ElseIf Mid(strKey, intPoint + 2) <> mUserInfo.strUserID Then
            MsgBoxEx "您的身份证号：" & _
                       vbCrLf & vbTab & "【" & mUserInfo.strUserID & "】" & vbCrLf & _
                       "当前证书唯一标识:" & _
                       vbCrLf & vbTab & "【" & Mid(strKey, intPoint + 2) & "】" & vbCrLf & _
                       "用户身份证号与当前证书唯一标识不相等,不能使用！", vbInformation, gstrSysName
            Exit Function
        End If
110     If mLastPWD <> "" Then strPIN = mLastPWD

112     If strPIN = "" Then
114         If Not frmPassword.ShowMe(strPIN) Then Exit Function
        End If
        
116     If Not GetCertLogin(strKey, strPIN, strSigCert, intDate, strWebUrl) Then
118         strPIN = ""
            blnRet = False
        Else
            blnRet = True
        End If
        
        If blnRet Then
            '判断是否需要更新注册证书
            udtUser.strName = strUserName
            udtUser.strSignName = strUserName
            udtUser.strUserID = Mid(strKey, intPoint + 2) 'SF+身份证号
            udtUser.strCertSn = strCertSn
            udtUser.strCertDN = GetCertDN(strSigCert)
            udtUser.strCert = strSigCert
            udtUser.strEncCert = ""
            udtUser.strCertID = strKey
            '获取已经注册证书的有效结束日期 日期格式:axBJCASecCOMV21 这个版本解析出来的都是2015/09/15
            If gudtPara.bytSignVersion = V_SEH Then
                strDate = SHCA_Client.SEH_GetCertValidDate(mUserInfo.strCert)
            Else
                strDate = SHCA_Client.ESE_GetCertValidDate(mUserInfo.strCert)
            End If
            If IsUpdateRegCert(udtUser, strDate, blnReDo) Then
                blnRet = True
            Else
                blnRet = False
            End If
        End If
     
        mLastPWD = strPIN
        SHCA_CheckCert = blnRet
        Exit Function
errH:
124     MsgBoxEx "检查USBKEY失败！" & vbNewLine & "第" & CStr(Erl()) & "行 " & Err.Description, vbQuestion, gstrSysName
End Function

Public Sub SHCA_UloadObj()
    Set SHCA_Client = Nothing
    mblnInit = False
End Sub
'----- 以下是内部函数

''' 获取客户端证书列表
''' 返回boolean
Private Function GetCertList(ByRef strName As String, ByRef strUniqueID As String, ByRef strCert As String, Optional ByRef strCertSn As String) As Boolean
    '-入参:无
    '-出参
    'strName :      保存接口返回的证书所有者姓名
    'strUniqueID:   保存接口返回的证书所有者唯一标识
    'strCert:       保存接口返回的签名证书
    Dim strPassas As String
    On Error GoTo errH
    If gudtPara.bytSignVersion = V_SEH Then
        SHCA_Client.SEH_InitialSession 2, "", "", 0, 2, "", "" '初始化CA接口
        strCert = SHCA_Client.SEH_GetSelfCertificate(10, "com1", "")
        If SHCA_Client.errorCode <> 0 Then
            '兼容SM2与RSA混用的情况 从RSA切换到SM2
            SHCA_Client.ESE_InitialSession 2, "", "", 0, 2, "", ""
            strCert = SHCA_Client.ESE_GetSelfCertificate(36, "com1")
            If SHCA_Client.errorCode = 0 Then
                gudtPara.bytSignVersion = V_ESE
                GoTo LineESE
            Else
                MsgBoxEx ValidateCertView(SHCA_Client.errorCode)
            End If
            Exit Function
        End If
LineSEH:
        strName = SHCA_Client.SEH_GetCertDetail(strCert, 17)
        If SHCA_Client.errorCode <> 0 Then
            MsgBoxEx ValidateCertView(SHCA_Client.errorCode)
            Exit Function
        End If
        
        strUniqueID = SHCA_Client.SEH_GetCertInfoByOID(strCert, "1.2.156.112570.148")
        If SHCA_Client.errorCode <> 0 Then
            MsgBoxEx ValidateCertView(SHCA_Client.errorCode)
            Exit Function
        End If
        
        strCertSn = SHCA_Client.SEH_GetCertDetail(strCert, 2)
        If SHCA_Client.errorCode <> 0 Then
            MsgBoxEx ValidateCertView(SHCA_Client.errorCode)
            Exit Function
        End If
    ElseIf gudtPara.bytSignVersion = V_ESE Then
        SHCA_Client.ESE_InitialSession 2, "", "", 0, 2, "", ""
        strCert = SHCA_Client.ESE_GetSelfCertificate(36, "com1")
        If SHCA_Client.errorCode <> 0 Then
            '兼容SM2与RSA混用的情况
            SHCA_Client.SEH_InitialSession 2, "", "", 0, 2, "", "" '初始化CA接口
            strCert = SHCA_Client.SEH_GetSelfCertificate(10, "com1", "")
            If SHCA_Client.errorCode = 0 Then
                gudtPara.bytSignVersion = V_SEH
                GoTo LineSEH
            Else
                MsgBoxEx ValidateCertView(SHCA_Client.errorCode)
            End If
            Exit Function
        End If
LineESE:
        strName = SHCA_Client.ESE_GetCertDetail(strCert, 17)
        If SHCA_Client.errorCode <> 0 Then
            MsgBoxEx ValidateCertView(SHCA_Client.errorCode)
            Exit Function
        End If
        
        strUniqueID = SHCA_Client.ESE_GetCertInfoByOID(strCert, "1.2.156.112570.148")
        If SHCA_Client.errorCode <> 0 Then
            MsgBoxEx ValidateCertView(SHCA_Client.errorCode)
            Exit Function
        End If
        
        strCertSn = SHCA_Client.ESE_GetCertDetail(strCert, 2)
        If SHCA_Client.errorCode <> 0 Then
            MsgBoxEx ValidateCertView(SHCA_Client.errorCode)
            Exit Function
        End If
    End If
    GetCertList = True
    Exit Function
errH:
    GetCertList = False
End Function

Private Function GetCertLogin(ByVal strUniqueID As String, ByVal strPassword As String, ByVal strCert As String, ByRef dDate As Integer, ByRef strWebserviceUrl As String) As Boolean
    '- 入参
    'strUniqueID : 证书唯一标识
    'strPassword : 证书密码
    'strWebserviceUrl:签名服务器地址，即为证书验证
    '- 出参
    'dDate       : 返回证书有效时间
    On Error GoTo errH
    Dim result As Boolean
    If SHCA_Client Is Nothing Then Set SHCA_Client = CreateObject("SafeEngineCOM.SafeEngineCtl")
    If (strPassword = "") Then
        MsgBoxEx "请输入证书密码！"
    Else
        '证书安全登录
        'result:0:成功
        'result:非0:不成功
        If mLogin >= 8 Then
            MsgBoxEx "已经输入了" & mLogin & "次错误密码，超过了最大输入次数！"
            Exit Function
        End If
        If gudtPara.bytSignVersion = V_SEH Then
            Call SHCA_Client.SEH_InitialSession(27, "com1", strPassword, 0, 27, "com1", "") '初始化CA接口(密码)
        Else
            Call SHCA_Client.ESE_InitialSession(36, "com1", strPassword, 0, 36, "com1", "") '初始化CA接口(密码)
        End If
        If SHCA_Client.errorCode = 0 Then
             '验证证书结果信息表示
            If gudtPara.bytSignVersion = V_SEH Then
                Call SHCA_Client.SEH_VerifyCertificate(strCert)
            Else
                Call SHCA_Client.ESE_VerifyCertificate(strCert)
            End If
            If SHCA_Client.errorCode = 0 Then
                
                '获取客户端证书有效期截止时间
                If gudtPara.bytSignVersion = V_SEH Then
                    dDate = SHCA_Client.SEH_GetCertValidDate(strCert)
                Else
                    dDate = SHCA_Client.ESE_GetCertValidDate(strCert)
                End If
                If (dDate <= 30 And dDate > 0) And Not gblnShow Then
                    MsgBoxEx "您的证书还有" & dDate & "天过期"
                    gblnShow = True
                    GetCertLogin = True
                ElseIf (dDate <= 0) Then
                    MsgBoxEx "您的证书已过期 " & Abs(dDate) & " 天"
                    GetCertLogin = False
                Else
                    GetCertLogin = True
                End If
            Else
               MsgBoxEx "验证证书错误！" & ValidateCertView(SHCA_Client.errorCode)
            End If
        Else
            mLogin = mLogin + 1
            MsgBoxEx "初始登陆错误！" & ValidateCertView(SHCA_Client.errorCode)
        End If
       
    End If
    Exit Function
errH:
    mLogin = mLogin + 1
    MsgBoxEx "证书密码可能不正确，您已经输入了" & mLogin & "次密码，还可以输入" & 8 - mLogin & "次!"
    GetCertLogin = False
End Function

''' <summary>
''' 验证证书结果信息表示
''' </summary>
''' <remarks></remarks>
Private Function ValidateCertView(retValidateCert) As String
    Dim strErrorMsg As String
    Select Case retValidateCert
        Case 0
            strErrorMsg = ""
        Case -2113667072:
            strErrorMsg = "装载动态库错误(-2113667072)"
            
        Case -2113667071:
            strErrorMsg = "内存分配错误(-2113667071)"
            
        Case -2113601536:
            strErrorMsg = "读取文件错误(-2113601536)"
            
        Case -2113601535:
            strErrorMsg = "密码错误(-2113601535)"
            
        Case -2113601534:
            strErrorMsg = "非法句柄错误(-2113601534)"
            
        Case -2113601533:
            strErrorMsg = "缺少ECC KEY错误(-2113601533)"
        
        Case -2113601532:
            strErrorMsg = "ECC KEY 与算法不匹配错误(-2113601532)"
            
        Case -2113601531:
            strErrorMsg = "非法算法错误(-2113601531)"
            
        Case -2113601530:
            strErrorMsg = "数字签名错误(-2113601530)"
            
        Case -2113601529:
            strErrorMsg = "摘要错误(-2113601529)"
            
        Case -2113601528:
            strErrorMsg = "缓冲区太小(-2113601528)"
            
        Case -2113601527:
            strErrorMsg = "证书格式错误(-2113601527)"
            
        Case -2113601526:
            strErrorMsg = "缺少公钥错误(-2113601526)"
            
        Case -2113601525:
            strErrorMsg = "验证签名错误(-2113601525)"
            
        Case -2113601524:
            strErrorMsg = "产生公私钥对错误(-2113601524)"
            
        Case -2113601523:
            strErrorMsg = "PKCS12编码错误(-2113601523)"
            
        Case -2113601522:
            strErrorMsg = "PKCS12格式错误(-2113601522)"
            
        Case -2113601520:
            strErrorMsg = "SE_ECC_ERROR_LOAD_BUILTIN_EC(-2113601520)"
            
        Case -2113601519:
            strErrorMsg = "公私钥不匹配错误(-2113601519)"
            
        Case -2113601518:
            strErrorMsg = "PKCS10编码错误(-2113601518)"
            
        Case -2113601517:
            strErrorMsg = "PKCS10解码错误(-2113601517)"
            
        Case -2113601516:
            strErrorMsg = "公钥格式错误(-2113601516)"
            
        Case -2113601515:
            strErrorMsg = "PKCS10格式错误(-2113601515)"
            
        Case -2113601514:
            strErrorMsg = "验证PKCS10错误(-2113601514)"
            
        Case -2113601513:
            strErrorMsg = "ECC KEY格式错误(-2113601513)"
            
        Case -2113601512:
            strErrorMsg = "公钥解码错误(-2113601512)"
            
        Case -2113601511:
            strErrorMsg = "签名格式错误(-2113601511)"
            
        Case -2113601510:
            strErrorMsg = "EC格式错误(-2113601510)"
            
        Case -2113601509:
            strErrorMsg = "ECC KEY解码错误(-2113601509)"
            
        Case -2113601508:
            strErrorMsg = "写入文件错误(-2113601508)"
            
        Case -2113601507:
            strErrorMsg = "证书链非法错误(-2113601507)"
            
        Case -2113601506:
            strErrorMsg = "内存分配错误(-2113601506)"
            
        Case -2113601505:
            strErrorMsg = "初始化环境错误(-2113601505)"
            
        Case -2113601504:
            strErrorMsg = "读取配置文件错误(-2113601504)"
            
        Case -2113601503:
            strErrorMsg = "打开设备错误(-2113601503)"
            
        Case -2113601502:
            strErrorMsg = "打开会话错误(-2113601502)"
            
        Case -2113601501:
            strErrorMsg = "装载动态库错误(-2113601501)"
            
        Case -2113601500:
            strErrorMsg = "设备类型错误(-2113601500)"
         
        Case -2113601499:
            strErrorMsg = "算法不支持错误(-2113601499)"
            
        Case -2113601498:
            strErrorMsg = "产生PKCS10错误(-2113601498)"
            
        Case -2113601497:
            strErrorMsg = "导出公钥错误(-2113601497)"
            
        Case -2113601496:
            strErrorMsg = "EC_POINT非法错误(-2113601496)"
    
        Case -2113601495:
            strErrorMsg = "对称加密错误(-2113601495)"
            
        Case -2113601494:
            strErrorMsg = "对称解密错误(-2113601494)"
            
        Case -2113601493:
            strErrorMsg = "PEM解码错误(-2113601493)"
            
        Case -2113601492:
            strErrorMsg = "获取证书细目错误(-2113601492)"
            
        Case -2113601491:
            strErrorMsg = "PEM编码错误(-2113601491)"
            
        Case -2113601490:
            strErrorMsg = "获取证书扩展项错误(-2113601490)"
            
        Case -2113601489:
            strErrorMsg = "非法接口类型错误(-2113601489)"
            
        Case -2113601488:
            strErrorMsg = "非法参数错误(-2113601488)"
            
        Case -2113601487:
            strErrorMsg = "枚举设备错误(-2113601487)"
            
        Case -2113601486:
            strErrorMsg = "没有设备(-2113601486)"
            
        Case -2113601485:
            strErrorMsg = "设备连接错误(-2113601485)"
            
        Case -2113601484:
            strErrorMsg = "产生随机数错误(-2113601484)"
            
        Case -2113601483:
            strErrorMsg = "SE_ECC_ERROR_SKF_SET_SYMKEY(-2113601483)"
            
        Case -2113601482:
            strErrorMsg = "对称加密初始化错误(-2113601482)"
            
        Case -2113601481:
            strErrorMsg = "对称加密错误(-2113601481)"
            
        Case -2113601480:
            strErrorMsg = "设备管理员口令错误(-2113601480)"
            
        Case -2113601479:
            strErrorMsg = "打开应用错误(-2113601479)"
            
        Case -2113601478:
            strErrorMsg = "设备已锁(-2113601478)"
            
        Case -2113601477:
            strErrorMsg = "设备口令错误(-2113601477)"
            
        Case -2113601476:
            strErrorMsg = "枚举应用错误(-2113601476)"
            
        Case -2113601475:
            strErrorMsg = "删除应用错误(-2113601475)"
            
        Case -2113601474:
            strErrorMsg = "创建应用错误(-2113601474)"
            
        Case -2113601473:
            strErrorMsg = "创建容器错误(-2113601473)"
            
        Case -2113601472:
            strErrorMsg = "设备不支持错误(-2113601472)"
            
        Case -2113601471:
            strErrorMsg = "打开容器错误(-2113601471)"
            
        Case -2113601470:
            strErrorMsg = "导出公钥错误(-2113601470)"
            
        Case -2113601466:
            strErrorMsg = "对称加密错误(-2113601466)"
            
        Case -2113601465:
            strErrorMsg = "导入密钥对错误(-2113601465)"
            
        Case -2113601464:
            strErrorMsg = "修改设备口令错误(-2113601464)"
            
        Case -2113601463:
            strErrorMsg = "导入证书错误(-2113601463)"
            
        Case -2113601462:
            strErrorMsg = "导出证书错误(-2113601462)"
            
        Case -2113601461:
            strErrorMsg = "创建文件错误(-2113601461)"
            
        Case -2113601460:
            strErrorMsg = "写入文件错误(-2113601460)"
            
        Case -2113601459:
            strErrorMsg = "获取文件信息错误(-2113601459)"
            
        Case -2113601458:
            strErrorMsg = "读取文件错误(-2113601458)"
            
        Case -2113601457:
            strErrorMsg = "获取公钥错误(-2113601457)"
            
        Case -2113601454:
            strErrorMsg = "生成密钥对错误(-2113601454)"
            
        Case -2113601453:
            strErrorMsg = "证书已过期(-2113601453)"
            
        Case -2113601452:
            strErrorMsg = "多个设备错误(-2113601452)"
            
        Case -2113601451:
            strErrorMsg = "没有设备(-2113601451)"
            
        Case -2113601450:
            strErrorMsg = "自动检测设备错误(-2113601450)"
            
        Case -2113601449:
            strErrorMsg = "设备无法识别(-2113601449)"
            
        Case -2113601448:
            strErrorMsg = "获取会话密钥错(-2113601448)"
            
        Case -2113601447:
            strErrorMsg = "导入会话密钥错(-2113601447)"
            
        Case -2113601446:
            strErrorMsg = "初始化摘要错误(-2113601446)"
            
        Case -2113601445:
            strErrorMsg = "更新摘要错误(-2113601445)"
            
        Case -2113601444:
            strErrorMsg = "生成会话密钥错(-2113601444)"
            
        Case -2113601442:
            strErrorMsg = "导入会话密钥错(-2113601442)"
            
        Case -2113601441:
            strErrorMsg = "缓冲区太小(-2113601441)"
            
        Case -2113601440:
            strErrorMsg = "P7签名数据初始化错误(-2113601440)"
            
        Case -2113601439:
            strErrorMsg = "产生随机数错误(-2113601439)"
            
        Case -2113601438:
            strErrorMsg = "对称加密错误(-2113601438)"
            
        Case -2113601437:
            strErrorMsg = "对称解密错误(-2113601437)"
            
        Case -2113601436:
            strErrorMsg = "导出公钥错误(-2113601436)"
            
        Case -2113601435:
            strErrorMsg = "添加p7算法错误(-2113601435)"
            
        Case -2113601434:
            strErrorMsg = "P7数据处理错误(-2113601434)"
            
        Case -2113601433:
            strErrorMsg = "SE_ECC_ERROR_ENVELOPE_ADD_RECIP(-2113601433)"
            
        Case -2113601432:
            strErrorMsg = "签名数据错误(-2113601432)"
            
        Case -2113601431:
            strErrorMsg = "摘要数据处理错误(-2113601431)"
            
        Case -2113601430:
            strErrorMsg = "加密更新错误(-2113601430)"
            
        Case -2113601429:
            strErrorMsg = "加密处理错误(-2113601429)"
            
        Case -2113601428:
            strErrorMsg = "解密初始化错误(-2113601428)"
            
        Case -2113601427:
            strErrorMsg = "解密更新错误(-2113601427)"
            
        Case -2113601426:
            strErrorMsg = "解密处理错误(-2113601426)"
            
        Case -2113601425:
            strErrorMsg = "p7格式错误(-2113601425)"
            
        Case -2113601424:
            strErrorMsg = "SE_ECC_ERROR_P7_NO_RECIP(-2113601424)"
            
        Case -2113601423:
            strErrorMsg = "算法非法(-2113601423)"
            
        Case -2113601422:
            strErrorMsg = "私钥长度错误(-2113601422)"
            
        Case -2113601421:
            strErrorMsg = "P7签名错误(-2113601421)"
            
        Case -2113601420:
            strErrorMsg = "验证P7签名错误(-2113601420)"
            
        Case -2113601419:
            strErrorMsg = "P7签名设置版本错误(-2113601419)"
            
        Case -2113601418:
            strErrorMsg = "锁设备错误(-2113601418)"
            
        Case -2113601417:
            strErrorMsg = "缓冲区太小(-2113601417)"
            
        Case -2113601416:
            strErrorMsg = "从LDAP获取证书错误(-2113601416)"
            
        Case -2113601415:
            strErrorMsg = "连接OCSP服务器错误(-2113601415)"
            
        Case -2113601414:
            strErrorMsg = "参数错误(-2113601414)"
            
        Case -2113601413:
            strErrorMsg = "CRL格式错误(-2113601413)"
            
        Case -2113601412:
            strErrorMsg = "证书废除(-2113601412)"
            
        Case -2113601411:
            strErrorMsg = "证书链格式错误(-2113601411)"
            
        Case -2113601410:
            strErrorMsg = "验证证书链错误(-2113601410)"
            
        Case -2113601409:
            strErrorMsg = "管理员密码错误(-2113601409)"
            
        Case -2113601408:
            strErrorMsg = "设备标签格式错误(-2113601408)"
            
        Case -2113601407:
            strErrorMsg = "删除容器错误(-2113601407)"
            
        Case -2113601406:
            strErrorMsg = "枚举文件错误(-2113601406)"
            
        Case -2113601405:
            strErrorMsg = "删除文件错误(-2113601405)"
            
        Case -2113601404:
            strErrorMsg = "枚举容器错误(-2113601404)"
            
        Case -2113601403:
            strErrorMsg = "关闭应用错误(-2113601403)"
        
        Case -2113568768:
            strErrorMsg = "SE_ECC_ERROR_FUNC_LOCAL(-2113568768)"
            
        Case -2113667070:
            strErrorMsg = "读私钥设备错误(-2113667070)"
            
        Case -2113667069:
            strErrorMsg = "私钥密码错误(-2113667069)"
            
        Case -2113667068:
            strErrorMsg = "读证书链设备错误(-2113667068)"
            
        Case -2113667067:
            strErrorMsg = "证书链密码错误(-2113667067)"
            
        Case -2113667066:
            strErrorMsg = "读证书设备错误(-2113667066)"
            
        Case -2113667065:
            strErrorMsg = "证书密码错误(-2113667065)"
            
        Case -2113667064:
            strErrorMsg = "私钥超时(-2113667064)"
            
        Case -2113667063:
            strErrorMsg = "缓冲区太小(-2113667063)"
            
        Case -2113667062:
            strErrorMsg = "初始化环境错误(-2113667062)"
            
        Case -2113667061:
            strErrorMsg = "清除环境错误(-2113667061)"
            
        Case -2113667060:
            strErrorMsg = "数字签名错误(-2113667060)"
            
        Case -2113667059:
            strErrorMsg = "验证签名错误(-2113667059)"
            
        Case -2113667058:
            strErrorMsg = "摘要错误(-2113667058)"
            
        Case -2113667057:
            strErrorMsg = "证书格式错误(-2113667057)"
            
        Case -2113667056:
            strErrorMsg = "数字信封错误(-2113667056)"
            
        Case -2113667055:
            strErrorMsg = "从LDAP获取证书错误(-2113667055)"
            
        Case -2113667054:
            strErrorMsg = "证书已过期(-2113667054)"
            
        Case -2113667053:
            strErrorMsg = "获取证书链错误(-2113667053)"
            
        Case -2113667052:
            strErrorMsg = "证书链格式错误(-2113667052)"
            
        Case -2113667051:
            strErrorMsg = "验证证书链错误(-2113667051)"
            
        Case -2113667050:
            strErrorMsg = "证书已废除(-2113667050)"
            
        Case -2113667049:
            strErrorMsg = "CRL格式错误(-2113667049)"
            
        Case -2113667048:
            strErrorMsg = "连接OCSP服务器错误(-2113667048)"
            
        Case -2113667047:
            strErrorMsg = "OCSP请求编码错误(-2113667047)"
            
        Case -2113667046:
            strErrorMsg = "OCSP回包错误(-2113667046)"
            
        Case -2113667045:
            strErrorMsg = "OCSP回包格式错误(-2113667045)"
            
        Case -2113667044:
            strErrorMsg = "OCSP回包过期(-2113667044)"
            
        Case -2113667043:
            strErrorMsg = "OCSP回包验证签名错误(-2113667043)"
            
        Case -2113667042:
            strErrorMsg = "证书状态未知(-2113667042)"
            
        Case -2113667041:
            strErrorMsg = "对称加解密错误(-2113667041)"
            
        Case -2113667040:
            strErrorMsg = "获取证书信息错误(-2113667040)"
            
        Case -2113667039:
            strErrorMsg = "获取证书细目错误(-2113667039)"
            
        Case -2113667038:
            strErrorMsg = "获取证书唯一标识错误(-2113667038)"
            
        Case -2113667037:
            strErrorMsg = "获取证书扩展项错误(-2113667037)"
            
        Case -2113667036:
            strErrorMsg = "PEM编码错误(-2113667036)"
            
        Case -2113667035:
            strErrorMsg = "PEM解码错误(-2113667035)"
            
        Case -2113667034:
            strErrorMsg = "产生随机数错误(-2113667034)"
            
        Case -2113667033:
            strErrorMsg = "PKCS12参数错误(-2113667033)"
            
        Case -2113667032:
            strErrorMsg = "私钥格式错误(-2113667032)"
            
        Case -2113667031:
            strErrorMsg = "公私钥不匹配(-2113667031)"
            
        Case -2113667030:
            strErrorMsg = "PKCS12编码错误(-2113667030)"
            
        Case -2113667029:
            strErrorMsg = "PKCS12格式错误(-2113667029)"
            
        Case -2113667028:
            strErrorMsg = "PKCS12解码错误(-2113667028)"
            
        Case -2113667027:
            strErrorMsg = "非对称加解密错误(-2113667027)"
            
        Case -2113667026:
            strErrorMsg = "OID格式错误(-2113667026)"
            
        Case -2113667025:
            strErrorMsg = "LDAP地址格式错误(-2113667025)"
            
        Case -2113667024:
            strErrorMsg = "LDAP地址错误(-2113667024)"
            
        Case -2113667023:
            strErrorMsg = "连接LDAP服务器错误(-2113667023)"

        Case -2113667022:
            strErrorMsg = "LDAP绑定错误(-2113667022)"
            
        Case -2113667021:
            strErrorMsg = "没有OID对应的扩展项(-2113667021)"
            
        Case -2113667020:
            strErrorMsg = "获取证书级别错误(-2113667020)"
            
        Case -2113667019:
            strErrorMsg = "读取配置文件错误(-2113667019)"
            
        Case -2113667018:
            strErrorMsg = "私钥未载入(-2113667018)"
            
  ' 以下错误用于登录
        Case -2113666824:
            strErrorMsg = "无效的登录凭证(-2113666824)"
            
        Case -2113666823:
            strErrorMsg = "参数错误(-2113666823)"
            
        Case -2113666822:
            strErrorMsg = "不是服务器证书(-2113666822)"
            
        Case -2113666821:
            strErrorMsg = "登录错误(-2113666821)"
            
        Case -2113666820:
            strErrorMsg = "证书验证方式错误(-2113666820)"
            
        Case -2113666819:
            strErrorMsg = "随机数验证错误(-2113666819)"
            
        Case -2113666818:
            strErrorMsg = "与单点登录客户端代理通信(-2113666818)"
    End Select
    ValidateCertView = strErrorMsg
End Function





