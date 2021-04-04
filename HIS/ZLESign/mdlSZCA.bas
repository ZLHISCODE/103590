Attribute VB_Name = "mdlSZCA"
Option Explicit
'深圳电子电子签名接口
'
Private mSZCAClient As Object           '证书部件
Private mobjSignPic As Object           '签章图片控件
Private mobjSoapClient As Object       'soap连接对象
Private mobjTSA As Object               '确信时间戳对象
Private mblnInit As Boolean

Private Const M_STR_SN As String = "SN"
Private Const M_STR_DN As String = "DN"
Private Const M_STR_TB As String = "TIMEB"
Private Const M_STR_TE As String = "TIMEE"
Private Const M_STR_VER As String = "VER"
Private Const M_STR_OID As String = "1.2.156.1002"

Public Function SZCA_InitObj() As Boolean
'功能:电子签名对象初始化
'     SOAP连接对象初始化
        Dim strUrl As String
        Dim arrPara As Variant
        
        On Error GoTo errH

1000    SZCA_InitObj = mblnInit
1001    If mblnInit Then Exit Function
        
1002    gstrPara = gobjComLib.zlDatabase.GetPara(90000, glngSys)   '读取URLs 固定读取ZLHIS 系统默认100
        '"http://127.0.0.1:8080/SZCAJavaCAS/service/SZCASafeService?wsdl&&&时间戳IP&&&时间戳端口号"
        'gstrPara = "127.0.0.1&&&8080&&&124.133.51.13&&&8888"
        If gstrPara = "" Then
            Err.Raise -1, , "当前系统【" & glngSys & "】没有配置电子签名参数,请到启用电子签名接口处设置。"
            Exit Function
        End If
        arrPara = Split(gstrPara, G_STR_SPLIT)
        If UBound(arrPara) <> 3 Then
            Err.Raise -2, , "当前系统【" & glngSys & "】配置电子签名参数有误,请到启用电子签名接口处设置。"
            Exit Function
        End If
        
1005    Set mSZCAClient = CreateObject("SZCAPKI.SZCAPKICtrl.1")
1006    Set mobjSignPic = CreateObject("SZCAPDFSIGNCTRL.SZCAPdfSignCtrlCtrl.1")
1007    Set mobjTSA = CreateObject("SuresecTsaClass.tsa.1")
1008    Set mobjSoapClient = CreateObject("MSSOAP.SoapClient30")  'SOAP连接对象，
1009    mobjSoapClient.ClientProperty("ServerHTTPRequest") = True
        strUrl = "http://" & arrPara(0) & ":" & arrPara(1) & "/SZCAJavaCAS/services/szcaCAValidate?wsdl" '正式环境地址 问题：112774
1010    mobjSoapClient.MSSoapInit (strUrl)
         
        mobjTSA.ISetTcpServerInfo arrPara(2), arrPara(3), 20
1030    mblnInit = True
1031    SZCA_InitObj = True

1090    Exit Function

errH:
118     MsgBoxEx "初始化电子签名部件失败！" & vbNewLine & "第" & CStr(Erl()) & "行：" & Err.Description, vbInformation, gstrSysName
    
End Function

Public Function SZCA_RegCert(arrCertInfo As Variant) As Boolean
        '功能：提供在HIS数据库中注册个人证书的必要信息,用于新发放或更换证书,,需要插入USB-Key
        '返回：arrCertInfo作为数组返回证书相关信息
        '      0-ClientSignCertCN:客户端签名证书公共名称(姓名),用于注册证书时程序验证身份
        '      1-ClientSignCertDN:客户端签名证书主题(每人唯一)
        '      2-ClientSignCertSN:客户端签名证书序列号(每证书唯一)
        '      3-ClientSignCert:客户端签名证书内容
        '      4-ClientEncCert:客户端加密证书内容
        '      5-签名图片文件名,空串表示没有签名图片
        '      6-时间戳证书
        Dim strKeyId As String, strCertTime As String, strCertUserName As String, strCertDN As String
        Dim strSigCert As String, strUserKey As String, strTSCert As String
        Dim strFile As String
        Dim i As Long
        On Error GoTo errH
    
100     For i = LBound(arrCertInfo) To UBound(arrCertInfo)
101         arrCertInfo(i) = ""
102     Next
        
108     If SZCA_GetCertList(strCertUserName, strKeyId, strCertDN, strSigCert, strUserKey, strFile, strTSCert) Then
200         arrCertInfo(0) = strCertUserName
201         arrCertInfo(1) = strCertDN
202         arrCertInfo(2) = strKeyId
203         arrCertInfo(3) = strSigCert
205         arrCertInfo(4) = strUserKey
206         arrCertInfo(5) = strFile
            arrCertInfo(6) = strTSCert
            SZCA_RegCert = True
        End If
        
300     Exit Function

errH:
    MsgBoxEx "证书注册失败！" & vbNewLine & "第" & CStr(Erl()) & "行 " & Err.Description, vbInformation, gstrSysName
End Function

Private Function SZCA_GetCertList(Optional ByRef strName As String, Optional ByRef strSN As String, _
    Optional ByRef strDn As String, Optional ByRef strCert As String = "-1", _
    Optional ByRef strUserUnigueID As String = "-1", Optional ByRef strFile As String = "-1", _
    Optional ByRef strTSCert As String = "-1") As Boolean
'功能:获取证书信息
    Dim blnRet As Boolean
    Dim arrList As Variant
    Dim strPic As String
    Dim strPath As String, strSource As String
    Dim lngRet As Long, lngTSALen As Long, lngCertLen As Long
    Dim arrTSData(2048) As Byte
    Dim arrCertData(2048) As Byte
    Dim bytSource() As Byte
    
    On Error GoTo errH
1000    Call mSZCAClient.AxInit          '初始化接口
1001    Call mSZCAClient.AxSetCertFilterStr("SC;SZCA;#;#;#;")   '证书过滤
1002    blnRet = mSZCAClient.AxSetKeyStore()   ' 设置签名,解密的证书
1003    If blnRet Then
1004        strSN = mSZCAClient.AxGetCertInfo(M_STR_SN)
1005        strDn = mSZCAClient.AxGetCertInfo(M_STR_DN)  'CN=支太祥,OU=429320496,O=六盘水市人民医院,O=检验科,L=六盘水市,ST=贵州省,C=CN
            arrList = Split(strDn, ",")
            strName = Mid(arrList(0), 4)
1006        If strCert <> "-1" Then strCert = mSZCAClient.AxGetCertData() '证书内容
1007        If strUserUnigueID <> "-1" Then strUserUnigueID = mSZCAClient.AxGetCertExt(M_STR_OID)   '扩展项 用户医疗卫生唯一标识
            
            If strFile <> "-1" Then
1008            strPic = mobjSignPic.SZCA_GetSealDataFromKey() 'PNG格式的数据 返回数据格式为 1-@@@印章图像1的Base64数据2-@@@印章图像2的Base64数据......n-@@@印章图像n的Base64数据
                strPic = Split(strPic, "@@@")(1)
1020            strFile = SaveBase64ToFile("BMP", strSN, strPic)
1030            Call SaveStdPicToFile(LoadPictureGDIPlus(strFile), strFile, BMP, 100)
            End If
            If strTSCert <> "-1" Then
1040            strSource = "测试ABCabc123"
1041            ReDim bytSource(LenB(StrConv(strSource, vbFromUnicode)))
1042            bytSource = StrConv(strSource, vbFromUnicode)

1050            lngRet = mobjTSA.IGenTokenByPlain("SHA1", bytSource(0), UBound(bytSource) + 1, 1, arrTSData(0), lngTSALen)
                If lngRet = 0 Then
1051                lngRet = mobjTSA.IGetTokenCertificate(arrTSData(0), lngTSALen, arrCertData(0), lngCertLen)
1052                strTSCert = FuncEncodeBase64Byte(arrCertData, lngCertLen)
                Else
                    MsgBoxEx "时间戳证书获取失败！", vbOKOnly + vbInformation, gstrSysName
                    Exit Function
                End If
            End If
        Else
            MsgBoxEx "请插入证书Key！", vbOKOnly + vbInformation, gstrSysName
            Exit Function
        End If
1090
        SZCA_GetCertList = True
    Exit Function
errH:
    MsgBoxEx "获取证书信息失败！" & vbNewLine & "第" & CStr(Erl()) & "行 " & Err.Description, vbInformation, gstrSysName
End Function

Private Function GetCertLogin(ByVal strCert As String) As Boolean
'功能:证书登录验证
    Dim strSource As String, strBase64 As String
    Dim strSign As String, strMsg As String
    Dim strRet As String
    
1000    On Error GoTo errH
1001
1002    Randomize
1003    strSource = Int((100000 * Rnd) + 1)
1004    strSign = mSZCAClient.AxSign(strSource)
1005    If strSign <> "" Then
1006        strBase64 = mobjSoapClient.szcaWSSignatureValidatePkcs7String(strSign)
1007        strBase64 = DecodeBase64String(strBase64)
1008
1009        If strBase64 = "1" Then
            '1 证书有效 -1 证书无效，不是所信任的根 -2 证书无效，超过有效期 -3 证书无效，已加入黑名单
1010            strRet = mobjSoapClient.szcaWSCertificateValidateString(strCert)
1011            strRet = DecodeBase64String(strRet)
1012            Select Case strRet
                Case "-1"
                    strMsg = "证书无效，不是所信任的根"
                Case "-2"
                    strMsg = "证书无效，超过有效期"
                Case "-3"
                    strMsg = "证书无效，已加入黑名单"
1020            End Select
            Else
                strMsg = "登录验证失败，验证信息与登录信息不符"
            End If
        Else
            strMsg = "登录失败！"
        End If
        If strMsg <> "" Then
            MsgBoxEx strMsg, vbOKOnly + vbInformation, gstrSysName
        End If
1050
    GetCertLogin = True
    Exit Function
errH:
    MsgBoxEx "证书登录失败！" & vbNewLine & "第" & CStr(Erl()) & "行 " & Err.Description, vbInformation, gstrSysName
End Function

Public Function SZCA_CheckCert(ByVal strCurrCertSn As String) As Boolean
'功能：读取USB进行设备初始化并登录
'返回值:
'  strSigCert -签名证书内容

        Dim strSN As String, strSigCert As String

        On Error GoTo errH
1000    If Not SZCA_InitObj() Then
1002        MsgBoxEx "部件未初始化！"
            Exit Function
        End If
        
1004    If Not SZCA_GetCertList(, strSN, , strSigCert) Then Exit Function
1006    If strCurrCertSn <> strSN Then
1008        MsgBoxEx "该证书未注册在您的名下，不能使用！", vbInformation + vbOKOnly, gstrSysName
            Exit Function
        End If
       
1010
        '登录验证
        If Not InStr(gstrLogins & "|", "|" & strCurrCertSn & "|") > 0 Then '首次验证通过后，下次不在继续验证
            If Not GetCertLogin(strSigCert) Then
                Exit Function
            Else
                If InStr(gstrLogins & "|", "|" & strCurrCertSn & "|") = 0 Then gstrLogins = gstrLogins & "|" & strCurrCertSn
            End If
        End If
1016
        SZCA_CheckCert = True
        Exit Function
errH:
124     MsgBoxEx "检查USBKEY失败！" & vbNewLine & "第" & CStr(Erl()) & "行 " & Err.Description, vbInformation, gstrSysName
End Function

Public Function SZCA_Sign(ByVal strSN As String, ByVal strSource As String, ByRef strSignData As String, ByRef strTimeStamp As String, ByRef strTimeStampCode As String)
'功能:签名
    Dim blnRet As Boolean
    Dim strMsg As String
    Dim strDigest As String
    Dim strTime As String
    Dim bytSource() As Byte
    Dim arrTSData(2048) As Byte
    Dim arrCertData(2048) As Byte
    Dim lngTSALen As Long
    Dim lngRet As Long
    
    On Error GoTo errH
    
1000    If SZCA_CheckCert(strSN) Then
            strDigest = StringSHA1(strSource)
1005        strSignData = mSZCAClient.AxSign(strDigest)
            If strSignData = "" Then
                strMsg = "签名失败：返回签名值为空。"
            Else
1008            strTimeStamp = Format(gobjComLib.zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
            End If
1010
1020    End If

        If strMsg <> "" Then
            MsgBoxEx strMsg, vbInformation + vbOKOnly, gstrSysName
            Exit Function
        End If
1030
        '获取时间戳
        ReDim bytSource(LenB(StrConv(strSource, vbFromUnicode)))
        bytSource = StrConv(strSource, vbFromUnicode)
        '通过明文获取时间戳
        Call mobjTSA.IGenTokenByPlain("SHA1", bytSource(0), UBound(bytSource) + 1, 0, arrTSData(0), lngTSALen)
        '获取时间戳中时间
        lngRet = mobjTSA.IGetTokenGenerateTime(arrTSData(0), lngTSALen, strTime)
1050
        If lngRet = 0 Then
            strTimeStamp = String14ToDate(strTime, strMsg)
            strTimeStampCode = FuncEncodeBase64Byte(arrTSData, lngTSALen)
        Else
            strMsg = "时间戳失败：获取时间戳中时间失败。"
        End If
        If strMsg <> "" Then
            MsgBoxEx strMsg, vbInformation + vbOKOnly, gstrSysName
            Exit Function
        End If
1080
        SZCA_Sign = True
        Exit Function
errH:
    MsgBoxEx "签名失败！" & vbNewLine & "第" & CStr(Erl()) & "行 " & Err.Description, vbInformation, gstrSysName
End Function

Public Function SZCA_VerifySign(ByVal strCert As String, ByVal strSign As String, ByVal strSource As String, _
    ByVal strTSCert As String, ByVal strTStampCode As String) As Boolean
'功能:验证签名
    Dim strBase64 As String
    Dim strRet As String, strDigest As String
    Dim strMsg As String
    Dim blnRet As Boolean
    Dim lngRet As Long
    Dim bytSource() As Byte, bytTStamp() As Byte, bytTSCert() As Byte
    
    On Error GoTo errH
    
1000    strBase64 = mobjSoapClient.szcaWSSignatureValidatePkcs7String(strSign)
1001    strRet = DecodeBase64String(strBase64)
1002    blnRet = False
        If strRet = "1" Then
            '1 证书有效 -1 证书无效，不是所信任的根 -2 证书无效，超过有效期 -3 证书无效，已加入黑名单
            strBase64 = ""
1005        strBase64 = mobjSoapClient.szcaWSCertificateValidateString(strCert)
1006        strRet = DecodeBase64String(strBase64)
1007        Select Case strRet
                Case "1":
                    strMsg = "验签成功"
                    blnRet = True
                Case "-1":
                    strMsg = "证书无效，不是所信任的根"
                Case "-2":
                    strMsg = "证书无效，超过有效期"
                Case "-3":
                    strMsg = "证书无效，已加入黑名单"
            End Select
        Else
            strMsg = "登录验证失败，验证信息与登录信息不符"
        End If
        If blnRet Then
1010
            ReDim bytSource(LenB(StrConv(strSource, vbFromUnicode)))
            bytSource = StrConv(strSource, vbFromUnicode)
            bytTSCert = DecodeBase64Byte(strTSCert)
            bytTStamp = DecodeBase64Byte(strTStampCode)
1020
            lngRet = mobjTSA.IVerifyTimeStampTokenByPlain("SHA1", bytSource(0), UBound(bytSource) + 1, bytTStamp(0), UBound(bytTStamp) + 1, bytTSCert(0), UBound(bytTSCert) + 1)
            If lngRet = 0 Then
                blnRet = True
            Else
                strMsg = "时间戳验证失败"
                blnRet = False
            End If
        End If
1050
        If strMsg <> "" Then
            MsgBoxEx strMsg, vbInformation + vbOKOnly, gstrSysName
        End If
        SZCA_VerifySign = blnRet
1090
    Exit Function
errH:
    MsgBoxEx "验签失败！" & vbNewLine & "第" & CStr(Erl()) & "行 " & Err.Description, vbInformation, gstrSysName
End Function

Public Function SZCA_GetPara() As Boolean
'设置深圳CA服务器地址
    Dim arrList As Variant
    
    On Error GoTo errH
    gstrPara = gobjComLib.zlDatabase.GetPara(90000, glngSys)
    If gstrPara = "" Then gstrPara = "127.0.0.1&&&8080&&&124.133.51.13&&&8888"
    If gstrPara <> "" Then
        arrList = Split(gstrPara, G_STR_SPLIT)
        If UBound(arrList) = 3 Then
             gudtPara.strSIGNIP = Trim(arrList(0))
             gudtPara.strSignPort = Trim(arrList(1))
             gudtPara.strTSIP = Trim(arrList(2))
             gudtPara.strTSPort = Trim(arrList(3))
        End If
    End If
    Exit Function
errH:
    MsgBoxEx "读取参数失败！" & vbNewLine & "第" & CStr(Erl()) & "行 " & Err.Description, vbInformation, gstrSysName
End Function

Public Function SZCA_SetParaStr() As String
    SZCA_SetParaStr = gudtPara.strSIGNIP & G_STR_SPLIT & gudtPara.strSignPort & G_STR_SPLIT & gudtPara.strTSIP & G_STR_SPLIT & gudtPara.strTSPort
End Function

Public Sub SZCA_UnloadObj()
    Set mSZCAClient = Nothing
    Set mobjSignPic = Nothing
    Set mobjSoapClient = Nothing
    Set mobjTSA = Nothing
    mblnInit = False
End Sub

Private Function FuncEncodeBase64Byte(bytArr() As Byte, ByVal lngLength As Long) As String
'功能:将一个字节数组进行Base64编码，并返回字符串
    Dim strRet As String
    Dim i As Long
    Dim bytBuffer() As Byte
    
    ReDim bytBuffer(lngLength)
    
    For i = 0 To lngLength - 1
        bytBuffer(i) = bytArr(i)
    Next
    strRet = EncodeBase64Byte(bytBuffer)
    FuncEncodeBase64Byte = strRet
End Function


