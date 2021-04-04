Attribute VB_Name = "mdlBJCAZX"
Option Explicit
'北京CA中心功能模块
Private mblnInit As Boolean         '是否已初始化成功
Private mLastPWD As String          '缓存输入的密码

Private BJCA_Client As Object       '证书部件
Private BJCA_svs As Object          '
Private BJCA_Pic As Object          '读取证书图片部件
Private BJCA_TS  As Object          '时间戳对象
Private mblnTs As Boolean           '启用时间戳
Private mbytTSVer As Byte           'BJCA_TS_CLIENTCOMLib.BJCATSEngine/BJCA_TS_ClientCom.BJCATSEngine.1
                                    '"BJCA_TS_ClientCom.BJCATSEngine.1"-驻马店精神病医院时间戳对象;河南息县人民医院
                                    '"BJCA_TS_CLIENTCOMLib.BJCATSEngine" -西安儿童医院
Private mLogin As Long              '输入密码错误次数

Private Const BASE64CHR As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/="
Private Const STR_TS_VER_0 As String = "BJCA_TS_CLIENTCOMLIB.BJCATSENGINE"
Private Const STR_TS_VER_1 As String = "BJCA_TS_CLIENTCOM.BJCATSENGINE.1"


Public Function BJCA_InitObj() As Boolean
    '证书部件初始化
        Dim progID As String
        Dim strVer As String
        
        On Error GoTo errH
100
102     BJCA_InitObj = mblnInit
104     If mblnInit Then Exit Function
        mLastPWD = ""
106     Set BJCA_svs = CreateObject("BJCA_SVS_ClientCOM.BJCASVSEngine.1")
108     Set BJCA_Client = CreateObject("BJCASecCOM.BJCASecCOMV2.1")
110     Set BJCA_Pic = CreateObject("GetKeyPic.GetPic")

112     gstrPara = gobjComLib.zlDatabase.GetPara(90000, glngSys)  '读取URLs
        'gstrPara = "0&&&1"   '格式:启用时间戳&&&时间戳版本
        '"BJCA_TS_ClientCom.BJCATSEngine.1"新版\"BJCA_TS_CLIENTCOMLib.BJCATSEngine"老版
        If gstrPara = "" Then
            Err.Raise -1, , "当前系统【" & glngSys & "】没有配置电子签名参数【格式:启用时间戳|启用时间戳版本】，请先配置。"
            Exit Function
        End If
        mblnTs = (Val(Split(gstrPara, G_STR_SPLIT)(0)) = 1)
        mbytTSVer = Val(Split(gstrPara, G_STR_SPLIT)(1))
        
        If mblnTs Then
            If mbytTSVer = 0 Then
                strVer = STR_TS_VER_0
            ElseIf mbytTSVer = 1 Then
                strVer = STR_TS_VER_1
            End If
113         Set BJCA_TS = CreateObject(strVer)
        End If
        
114     BJCA_InitObj = True
    
116     mblnInit = BJCA_InitObj
        mLogin = 0
        Exit Function
errH:
118     MsgBoxEx "创建接口部件失败！" & vbNewLine & Err.Description, vbQuestion, gstrSysName
    
End Function

Public Function BJCA_RegCert(arrCertInfo As Variant) As Boolean
        '功能：提供在HIS数据库中注册个人证书的必要信息,用于新发放或更换证书,,需要插入USB-Key
        '返回：arrCertInfo作为数组返回证书相关信息
        '      0-ClientSignCertCN:客户端签名证书公共名称(姓名),用于注册证书时程序验证身份
        '      1-ClientSignCertDN:客户端签名证书主题(每人唯一)
        '      2-ClientSignCertSN:客户端签名证书序列号(每证书唯一)
        '      3-ClientSignCert:客户端签名证书内容
        '      4-ClientEncCert:客户端加密证书内容
        '      5-签名图片文件名,空串表示没有签名图片
        
        Dim strKeyId As String, strCertTime As String, strCertUserName As String, strCertDN As String
        Dim strSigCert As String, i As Integer, strCACert As String, lngOk As Long
        Dim strPicData As String
        
        On Error GoTo errH
    
100     For i = LBound(arrCertInfo) To UBound(arrCertInfo)
102         arrCertInfo(i) = ""
        Next

104     If GetCertList(strCertUserName, strKeyId, strSigCert) Then
106         arrCertInfo(0) = strCertUserName
108         arrCertInfo(1) = strCertDN
110         arrCertInfo(2) = strKeyId
112         arrCertInfo(3) = strSigCert

114         If Not BJCA_Pic Is Nothing Then
116             If UBound(arrCertInfo) >= 5 Then
118                 strPicData = BJCA_Pic.getpic()
120                 If strPicData <> "" Then
                        '新版电子病历gif格式签名报错,要求改成bmp
122                     arrCertInfo(5) = SaveBase64ToFile("bmp", strKeyId, strPicData) '图片路径
                    End If
                End If
            End If
124         BJCA_RegCert = True
        End If

        Exit Function
errH:
126     MsgBoxEx "获取证书信息失败！" & vbNewLine & "第" & CStr(Erl()) & "行 " & Err.Description, vbQuestion, gstrSysName

End Function

Public Function BJCA_Sign(ByVal strSource As String, ByRef strSignData As String, ByRef strTimeStamp As String, ByRef strTimeStampCode As String, ByRef blnReDo As Boolean) As Boolean
        '签名
        Dim strTiemRequest As String
        Dim strTmp As String
        Dim blnCheck As Boolean
        
        On Error GoTo errH
        blnCheck = BJCA_CheckCert(blnReDo)
        If blnReDo Then Exit Function
        
100     If blnCheck Then               '验证当前USB是否是签名用户的，并获取签名证书
            If mblnTs Then
                If mbytTSVer = 0 Then
                    '成功返回经base64编码的时间戳请求，失败返回空值
                    strTiemRequest = BJCA_TS.CreateTimeStampRequest(strSource)
                    '成功返回经base64编码的时间戳（不带证书），失败返回空值
                    strTimeStampCode = BJCA_TS.CreateTimeStampNoCert(strTiemRequest)
                    If strTimeStampCode = "" Then
                        MsgBoxEx "获取时间戳信息失败！"
                        Exit Function
                    Else
                        strTmp = BJCA_TS.gettimestampinfo(strTimeStampCode, 1) '解析时间
                        '时间戳返回格式：20140911192555，转换成 2014-09-11 19:25:55
                        strTimeStamp = Mid(strTmp, 1, 4) & "-" & Mid(strTmp, 5, 2) & "-" & Mid(strTmp, 7, 2) & " " & Mid(strTmp, 9, 2) & ":" & Mid(strTmp, 11, 2) & ":" & Mid(strTmp, 13, 2)
                    End If
                ElseIf mbytTSVer = 1 Then
                    strTiemRequest = BJCA_TS.CreateTSRequest(strSource, 0)   '不带证书
                    strTimeStampCode = BJCA_TS.CreateTS(strTiemRequest)
                    If strTimeStampCode = "" Then
                        MsgBoxEx "获取时间戳信息失败！"
                        Exit Function
                    Else
                        strTmp = BJCA_TS.GetTSInfo(strTimeStampCode, 1) '解析时间
                        '时间戳返回格式：20140911192555，转换成 2014-09-11 19:25:55
                        strTimeStamp = Mid(strTmp, 1, 4) & "-" & Mid(strTmp, 5, 2) & "-" & Mid(strTmp, 7, 2) & " " & Mid(strTmp, 9, 2) & ":" & Mid(strTmp, 11, 2) & ":" & Mid(strTmp, 13, 2)
                    End If
                End If
            End If
            '证书ID进行签名
110         strSignData = BJCA_Client.SignData(mUserInfo.strCertSn, strSource)
112
        Else
            MsgBoxEx "签名失败！", vbInformation, "电子签名部件"
            Exit Function
        End If
        BJCA_Sign = True
        Exit Function
errH:
114     MsgBoxEx "签名失败！" & vbNewLine & "第" & CStr(Erl()) & "行 " & Err.Description, vbQuestion, gstrSysName
End Function

Public Function BJCA_CheckCert(ByRef blnReDo As Boolean) As Boolean
'--------------------------------------------------------------------------------------------------------------------------
'功能：读取USB进行设备初始化并登录
'参数:
'   出参:blnRedo-证书更新需要重新检查
'返回:
'--------------------------------------------------------------------------------------------------------------------------
        Dim strPIN As String, strUserName As String
        Dim strWebUrl As String, intDate   As Integer
        Dim strCertSn As String
        Dim strPicData As String, strSigCert As String
        Dim strTmp As String, strFileName As String
        Dim blnRet As Boolean
        Dim udtUser As USER_INFO
        Dim strDate As String
        Dim strCertID As String
        
1       On Error GoTo errH
2       If Not BJCA_InitObj() Then
3           MsgBoxEx "部件未初始化！"
4           Exit Function
5       End If
    
6       Call GetCertList(strUserName, strCertID, strSigCert, strCertSn)
7       If mUserInfo.strUserID = "" Then
8           MsgBoxEx "您的身份证号为空,请联系管理员到人员管理中录入！", vbOKOnly + vbInformation, gstrSysName
9           Exit Function
10      ElseIf mUserInfo.strUserID <> Mid(strCertSn, 3) Then
11          MsgBoxEx "该证书未注册在您的名下，不能使用！"
12          Exit Function
13      End If
        
14      If mLastPWD <> "" Then strPIN = mLastPWD
'        strPIN = ""  'CA工程师提出 每次签名都要求输入密码
15      If strPIN = "" Then
16          If Not frmPassword.ShowMe(strPIN) Then Exit Function
17      End If
        
18      If Not GetCertLogin(strCertID, strPIN, strSigCert, intDate, strWebUrl) Then
19          strPIN = ""
20          blnRet = False
21      Else
22          blnRet = True
23      End If
24      LogWrite "BJCA_CheckCert", "GetCertLogin返回值 blnRet=" & blnRet
25      If blnRet Then
            '判断是否需要更新注册证书
26          udtUser.strName = strUserName
27          udtUser.strSignName = strUserName
28          udtUser.strUserID = Mid(strCertSn, 3) 'SF+身份证号
29          udtUser.strCertSn = strCertID
30          udtUser.strCertDN = ""
31          udtUser.strCert = strSigCert
32          udtUser.strEncCert = ""
33          udtUser.strCertID = strCertID
34          udtUser.strPicCode = BJCA_Pic.getpic()
            '获取已经注册证书的有效结束日期 日期格式:axBJCASecCOMV21 这个版本解析出来的都是2015/09/15
35          strDate = BJCA_Client.GetCertInfo(mUserInfo.strCert, 12)
36          If IsUpdateRegCert(udtUser, strDate, blnReDo) Then
37              blnRet = True
38          Else
39              blnRet = False
40          End If
41          LogWrite "BJCA_CheckCert", "IsUpdateRegCert返回值 blnRet=" & blnRet
42      End If
        
43      mLastPWD = strPIN
44      BJCA_CheckCert = blnRet
    
45      Exit Function
errH:
46      MsgBoxEx "检查USBKEY失败！" & vbNewLine & "第" & CStr(Erl()) & "行 " & Err.Description, vbQuestion, gstrSysName
End Function

Public Sub BJCA_UloadObj()
    Set BJCA_Client = Nothing
    Set BJCA_svs = Nothing
    Set BJCA_Pic = Nothing
    Set BJCA_TS = Nothing
    mblnInit = False
End Sub
'----- 以下是内部函数

''' 获取客户端证书列表
''' 返回boolean
Private Function GetCertList(ByRef strName As String, ByRef strUniqueID As String, ByRef strCert As String, Optional ByRef strCertSn As String) As Boolean
    '河科大第一附属医院获取数字证书列表函数
    '-入参:无
    '-出参
    'strName :      保存接口返回的证书所有者姓名
    'strUniqueID:   保存接口返回的证书所有者唯一标识
    'strCert:       保存接口返回的签名证书
      
    Dim strUsbkeyList As String
    Dim arrUserListLength As Integer
    Dim arrUserList() As String
    If BJCA_Client Is Nothing Then Set BJCA_Client = CreateObject("BJCASecCOM.BJCASecCOMV2.1")
    If BJCA_Pic Is Nothing Then Set BJCA_Pic = CreateObject("GetKeyPic.GetPic")
    
    strUsbkeyList = BJCA_Client.getUserList()
    arrUserList = Split(strUsbkeyList, "&&&")
    arrUserListLength = UBound(arrUserList)
    If (arrUserListLength = -1) Then
        MsgBoxEx "请您插入Key！", vbInformation, gstrSysName
        Exit Function
    End If
    If (arrUserListLength <> 0) Then
        Dim i As Integer
        For i = 0 To arrUserListLength - 1
            Dim strOption As String
            strOption = arrUserList(i)
            strName = Split(strOption, "||")(0)
            strUniqueID = Split(strOption, "||")(1)
            strCert = BJCA_Client.ExportUserCert(strUniqueID)
            strCertSn = BJCA_Client.GetCertInfoByOid(strCert, "1.2.156.112562.2.1.1.1")
            If strCertSn = "" Then
                'value="1.2.156.112562.2.1.1.1" 该标识为北京CA SM2证书中唯一标识
                'value="2.16.840.1.113732.2" 该标识为北京CA RSA证书中唯一标识
                strCertSn = BJCA_Client.GetCertInfoByOid(strCert, "2.16.840.1.113732.2") '当第一种方式取不到时缺省按第二种方式取
            End If
        Next
    End If
    GetCertList = True
End Function

Private Function GetCertLogin(ByVal strUniqueID As String, ByVal strPassword As String, ByVal strCert As String, ByRef dDate As Integer, ByRef strWebserviceUrl As String) As Boolean
    '河科大第一附属医院数字证书登录函数
    '- 入参
    'strUniqueID : 证书唯一标识
    'strPassword : 证书密码
    'strWebserviceUrl:签名服务器地址，即为证书验证
    '- 出参
    'dDate       : 返回证书有效时间

    Dim result As Boolean
    If BJCA_Client Is Nothing Then Set BJCA_Client = CreateObject("BJCASecCOM.BJCASecCOMV2.1")
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
        result = BJCA_Client.userLogin(strUniqueID, strPassword)
        LogWrite "GetCertLogin", "接口userLogin" & "    返回值:" & result
        If (result) Then
            mLogin = 0
            Dim strExtLib As String
            strExtLib = BJCA_Client.GetUserInfo(strUniqueID, 15)
            Dim intFlg As Integer
            
            '服务器端验证证书
            '从组件中导出证书
            Dim retValidateCert As Long
            retValidateCert = 100
            retValidateCert = ValidateCert(strCert, strWebserviceUrl)
            LogWrite "GetCertLogin", "ValidateCert" & vbCrLf & _
                        "参数1=" & strCert & vbCrLf & _
                        "参数2=" & strWebserviceUrl & vbCrLf & _
                        "返回值:" & retValidateCert
            '验证证书结果信息表示
            If retValidateCert <> 0 Then Call ValidateCertView(retValidateCert)

            If (retValidateCert = 0) Then
                Dim uniqueIdStr As String
                Dim oid As String
                oid = "2.16.840.1.113732.2"
                Dim s As String
                '获取客户端证书有效期截止时间
                s = BJCA_Client.GetCertInfo(strCert, 12)
                '验证客户端证书有效期剩余天数
                dDate = CheckValidaty(s)
            
                If (dDate <= 30 And dDate > 0 And Not gblnShow) Then
                    MsgBoxEx "您的证书还有" & dDate & "天过期"
                    uniqueIdStr = BJCA_Client.GetCertInfoByOid(strCert, oid)
                    gblnShow = True
                    GetCertLogin = True
                ElseIf (dDate <= 0) Then
                    MsgBoxEx "您的证书已过期 " & Abs(dDate) & " 天"
                    GetCertLogin = False
                Else
                    uniqueIdStr = BJCA_Client.GetCertInfoByOid(strCert, oid)
                    GetCertLogin = True
                End If
            Else
                GetCertLogin = False
            End If
        Else
            mLogin = mLogin + 1
            MsgBoxEx "证书密码可能不正确，您已经输入了" & mLogin & "次密码，还可以输入" & 8 - mLogin & "次!"
            GetCertLogin = False
        End If
    End If

End Function

Private Function ValidateCert(ByRef userCert As String, Optional webserviceUrl As String) As Integer
    '服务器端验证证书
 
    If BJCA_svs Is Nothing Then Set BJCA_svs = CreateObject("BJCA_SVS_ClientCOM.BJCASVSEngine.1")
    ValidateCert = BJCA_svs.ValidateCertificate(userCert)
 
End Function

''' <summary>
''' 验证证书结果信息表示
''' </summary>
''' <remarks></remarks>
Private Sub ValidateCertView(retValidateCert)
    Select Case retValidateCert
        Case 0
            MsgBoxEx "证书有效！"
        Case -1
            MsgBoxEx "不是所信任的根！"
        Case -2
            MsgBoxEx "超过有效期！"
        Case -3
            MsgBoxEx "作废证书！"
        Case -4
            MsgBoxEx "已加入黑名单！"
    End Select
End Sub

''' 客户端验证签名函数
''' 返回boolean值
Public Function BJCA_VerifySign(ByVal strCert As String, ByVal strInData As String, ByRef strData As String, ByVal strTimeStampCode As String) As Boolean
    '河科大第一附属医院数字证书签名验证函数
    '- 入参
    'strInData     : 签名结果
    'strCert       : 签名证书
    'strData       : 签名原文
    'strTimeStampCode :时间戳信息
    '-返回值
    'result:true:  成功
    'result:false: 失败
        Dim intVerifyRet As Integer
        Dim lngResult As Long
        Dim strInfo As String
        Dim blnRet As Boolean
        
        On Error GoTo errH
        '返回值  成功返回0，失败返回如下值：
        '-1为时间戳验证不通过
        '-2为原文验证不通过
        '-3为不是所信任的根
        '-4证书未生效
        '-5查询不到此证书
        '-6为签发时间戳时服务器证书过期
        If mblnTs Then
            If mbytTSVer = 0 Then
                lngResult = BJCA_TS.VerifyTimeStampData(strTimeStampCode, "") '只验证时间戳,不验证源文
            ElseIf mbytTSVer = 1 Then
                lngResult = BJCA_TS.VerifyTS(strTimeStampCode, strData)
            End If
            If lngResult <> 0 Then
                strInfo = "验证时间戳失败！详情:"
                Select Case lngResult
                Case -1
                    MsgBoxEx strInfo & "时间戳验证不通过！"
                Case -2
                    MsgBoxEx strInfo & "原文验证不通过！"
                Case -3
                    MsgBoxEx strInfo & "不是所信任的根！"
                Case -4
                    MsgBoxEx strInfo & "证书未生效！"
                Case -5
                    MsgBoxEx strInfo & "查询不到此证书！"
                Case -6
                    MsgBoxEx strInfo & "签发时间戳服务器证书过期！"
                End Select
                Exit Function
            End If
        End If

'100     If BJCA_Client Is Nothing Then Set BJCA_Client = CreateObject("BJCASecCOM.BJCASecCOMV2.1")
'101     verifySignResult = BJCA_Client.verifySignedData(strCert, strData, strInData)
'上面注释代码错误：西安儿童医院启用北京CA时，北京CA工程师：车利斌 指出，验证签名时采用服务器验证的方式,
'应该使用 此对象（"BJCA_SVS_ClientCOM.BJCASVSEngine.1"） 经行验证签名
        intVerifyRet = -1
100        If BJCA_svs Is Nothing Then Set BJCA_svs = CreateObject("BJCA_SVS_ClientCOM.BJCASVSEngine.1")
101        intVerifyRet = BJCA_svs.VerifySignedData(strCert, strData, strInData)

        If intVerifyRet = 0 Then
            MsgBoxEx "验证签名成功！", vbInformation, gstrSysName
            blnRet = True
        Else
            MsgBoxEx "验证签名失败！", vbInformation, gstrSysName
            blnRet = False
        End If
        BJCA_VerifySign = blnRet
    Exit Function
errH:
     MsgBoxEx "验证签名失败！" & vbNewLine & "第" & CStr(Erl()) & "行 " & Err.Description, vbQuestion, gstrSysName
End Function

''' 检查证书有效性
''' 返回证书有效期天数
Private Function CheckValidaty(ByVal endDate As Date) As Integer
    '河科大第一附属医院检查证书有效性接口
    '-入参: 证书有效截止日期
    '-出参：有效天数
        Dim dblAllSp    As Double
        Dim result      As Integer
        dblAllSp = CDbl(CDate(endDate)) - CDbl(Now)
        result = Int(dblAllSp)
        CheckValidaty = result
End Function

Public Function BJCA_GetPara() As Boolean
'设置湖北CA服务器地址
    Dim arrList As Variant
    
    On Error GoTo errH
    gstrPara = gobjComLib.zlDatabase.GetPara(90000, glngSys)
    If gstrPara = "" Then gstrPara = "0&&&1"   '格式:启用时间戳&&&时间戳版本
    arrList = Split(gstrPara, "&&&")
    If UBound(arrList) = 1 Then
        gudtPara.blnISTS = Val(arrList(0)) = 1
        gudtPara.strTSVersion = arrList(1)
    End If
    Exit Function
errH:
    MsgBoxEx "读取参数失败！" & vbNewLine & "第" & CStr(Erl()) & "行 " & Err.Description, vbInformation, gstrSysName
End Function

Public Function BJCA_SetParaStr() As String
    BJCA_SetParaStr = IIf(gudtPara.blnISTS, 1, 0) & G_STR_SPLIT & gudtPara.strTSVersion
End Function






