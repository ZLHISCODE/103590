Attribute VB_Name = "mdlHuBeiCA"
Option Explicit

Private HUBEI_Client As Object  'CSVS_C_SDK.1
Private HUBEI_SVS As Object  'SVS_S_SDK.1
Private HUBEI_TS As Object  'SVS_S_SDK.1
Private HUBEI_PIC As Object  'HBCA_SOFSeal.Seal.1

Private mblnInit As Boolean     '标记对象是否卸载
Private mintLogin As Integer
Private mstrMethod As String        'RSA-江夏区中医院;SM2-武昌医院

Public Function HUBEI_InitObj() As Boolean
    '证书部件初始化
        On Error GoTo errH
        Dim strSIGNIP As String, intSignPort As Integer, strSignURL As String
        Dim strTSIP As String, intTSPort As Integer, strTSURL As String
        Dim arrList As Variant
        Dim strTmp As String
        Dim lngRet As Long
        
        If mblnInit Then HUBEI_InitObj = True: Exit Function
        
        On Error GoTo 0
    
1000    Set HUBEI_Client = CreateObject("CSVS_C_SDK")
2000    Set HUBEI_SVS = CreateObject("SVS_S_SDK")   '创建证书验证控件
3000    Set HUBEI_TS = CreateObject("SVS_S_SDK")   '创建证书验证控件
4000    Set HUBEI_PIC = CreateObject("HBCA_SOFSeal.Seal")
        
        gstrPara = gobjComLib.zlDatabase.GetPara(90000, glngSys)
         'gstrPara = "221.232.224.75&&&8082&&&221.232.224.75&&&8084&&&RSA|SM2"   '格式:签名服务器&&&时间戳服务器 举例：IP&&&端口号&&&IP&&&端口号&&&KYE算法类型
        
        If gstrPara = "" Then
            Err.Raise -1, , "当前系统【" & 100 & "】没有配置电子签名参数,请到【公共参数设置】设置。"
            Exit Function
        End If
        '签名服务器URL:/hbcaDSS/hbusiness
        '时间戳服务器URL:/hbcaTSS/hbusiness
        arrList = Split(gstrPara, "&&&")
        If UBound(arrList) < 3 Then
            MsgBoxEx "签名服务器地址配置格式有误,请到基础参数进行重新设置。", vbOKOnly + vbInformation, gstrSysName
            Exit Function
        End If
        lngRet = -1
        strSIGNIP = arrList(0)
        intSignPort = CInt(arrList(1))
        strSignURL = "/hbcaDSS/hbusiness"
        lngRet = HUBEI_SVS.SOF_SetServerInfo(strSIGNIP, intSignPort, strSignURL, 80)
        If lngRet <> 0 Then
            MsgBoxEx "签名服务器初始化失败！", vbOKOnly + vbInformation, gstrSysName
            Exit Function
        End If
        lngRet = -1
        strTSIP = arrList(2)
        intTSPort = CInt(arrList(3))
        strTSURL = "/hbcaTSS/hbusiness"
        lngRet = HUBEI_TS.SOF_SetServerInfo(strTSIP, intTSPort, strTSURL, 80)
        If lngRet <> 0 Then
            MsgBoxEx "时间戳服务器初始化失败！", vbOKOnly + vbInformation, gstrSysName
            Exit Function
        End If
        If UBound(arrList) >= 4 Then
            mstrMethod = arrList(4)  'SM2-武昌医院
        Else
            mstrMethod = "RSA"   'RSA-江夏区中医院
        End If
        mintLogin = 0
        gstrLogins = ""
        mblnInit = True
        HUBEI_InitObj = True
    
    Exit Function
errH:
     MsgBoxEx "创建接口部件失败！" & vbNewLine & Err.Description, vbQuestion, gstrSysName
End Function

Public Function HUBEI_RegCert(arrCertInfo As Variant) As Boolean
'功能：提供在HIS数据库中注册个人证书的必要信息,用于新发放或更换证书,,需要插入USB-Key
'返回：arrCertInfo作为数组返回证书相关信息
'      0-ClientSignCertCN:客户端签名证书公共名称(姓名),用于注册证书时程序验证身份
'      1-ClientSignCertDN:客户端签名证书主题(每人唯一)
'      2-ClientSignCertSN:客户端签名证书序列号(每证书唯一)
'      3-ClientSignCert:客户端签名证书内容
'      4-ClientEncCert:客户端加密证书内容
'      5-签名图片文件名,空串表示没有签名图片
        
        Dim strCertID As String, strCertUserName As String, strPicPath As String
        Dim strCert As String, i As Integer
        Dim strCertSn As String
        On Error GoTo errH
        
        For i = LBound(arrCertInfo) To UBound(arrCertInfo)
             arrCertInfo(i) = ""
        Next
        
        If HUBEI_GetCertList(strCertUserName, strCertSn, strCertID, strPicPath) Then
            arrCertInfo(0) = strCertUserName
            arrCertInfo(1) = strCertSn '证书DN
            arrCertInfo(2) = strCertSn '证书序列号(证书主题作为唯一值)
            arrCertInfo(3) = ""
            arrCertInfo(4) = ""
            arrCertInfo(5) = strPicPath
                
            HUBEI_RegCert = True
        End If

        Exit Function
errH:
     MsgBoxEx "获取证书信息失败！" & vbNewLine & "第" & CStr(Erl()) & "行 " & Err.Description, vbQuestion, gstrSysName

End Function


Public Function HUBEI_GetCertList(ByRef strName As String, Optional ByRef strCertSn As String, Optional ByRef strCertID As String, Optional ByRef strPicPath As String = "0") As Boolean
'湖北版获取数字证书列表函数
'-入参:无
'-出参
'strName :      保存接口返回的证书所有者姓名
'strCertSN      保存接口返回的证书SN(实际返回的是证书主题DN,CA主题是唯一键,身份证号不是必填项
'strCertID      返回证书ID 签名时候要用
'strPicPath     缺省不读取

    Dim strUsbkeyList As String
    Dim arrUserList() As String
    Dim strUser As String
    Dim strUserID As String
    Dim strTmp As String
    Dim strCertBase As String
    Dim strBase64 As String
    
    Dim i As Integer
    On Error GoTo errH:
    '获取证书
    Call HUBEI_Client.SOF_SetCertAppPolicy("SIGN")
    If mstrMethod = "SM2" Then
        HUBEI_Client.SOF_SetHashMethod ("SM3")
    End If
    strUsbkeyList = HUBEI_Client.SOF_GetUserList()
    
    If (strUsbkeyList = "") Then
        strName = ""
        MsgBoxEx "请插入证书Key！", vbOKOnly + vbInformation, gstrSysName
        HUBEI_GetCertList = False
        Exit Function
    Else
        '用户1(CertID||Subject||IssuerSubject||CertBase64)&&&用户(CertID||Subject||IssuerSubject||CertBase64)&&&用户&&&用户
        '1419118795628E856CD1B3C0DD607693||CN=测试证书2, OU=急诊, O=江夏中医院, L=武汉, S=湖北, C=CN||CN=HBCA, O=Hubei Digital Certificate Authority Center CO Ltd., L=Wuhan, S=Hubei, C=CN||MIIEkzCCA3ugAwIBAg
        arrUserList = Split(strUsbkeyList, "&&&")

        If UBound(arrUserList) > 1 Then  '多个KEY
            For i = LBound(arrUserList) To UBound(arrUserList) - 1
                strTmp = Split(arrUserList(i), "||")(1)
                strTmp = Split(strTmp, ",")(0)
                strTmp = Mid(strTmp, 4)
                strUser = strUser & "&&&" & strTmp
            Next
            If strUser <> "" Then strUser = Mid(strUser, 4)
            strName = frmSelectUser.ShowMe(strUser)
            
            For i = LBound(arrUserList) To UBound(arrUserList) - 1
                strTmp = Split(arrUserList(i), "||")(1)
                strTmp = Split(strTmp, ",")(0)
                strTmp = Mid(strTmp, 4)
                If strName = strTmp Then
                     strCertSn = Split(arrUserList(i), "||")(1)
                     strCertID = Split(arrUserList(i), "||")(0)
                     strCertBase = Split(arrUserList(i), "||")(3)
                     Exit For
                End If
            Next
        Else
            arrUserList = Split(arrUserList(0), "||")
            strCertSn = arrUserList(1)      '证书DN
            strCertID = arrUserList(0)    '证书ID
            strCertBase = arrUserList(3)   '证书内容
            strName = Mid(Split(arrUserList(1), ",")(0), 4)
        End If
        
    End If
    If strPicPath = "" Then
        strUserID = HUBEI_Client.SOF_GetCertInfoByOidEx(strCertBase, "2.4.16.11.7.3")
        strBase64 = HUBEI_PIC.SOF_GetKeyPictureEx(strCertID, strUserID)
        strPicPath = SaveBase64ToFile("gif", strCertID, strBase64)
    End If
    
    HUBEI_GetCertList = True
    
    Exit Function
errH:
     MsgBoxEx "读取证书信息失败！" & vbNewLine & "第" & CStr(Erl()) & "行 " & Err.Description, vbQuestion, gstrSysName
End Function

Public Function HUBEI_CheckCert(ByVal strCurrCertSn As String, Optional ByRef strCertID As String) As Boolean
'---------------------------------------------------------------------------------------------------------------------
'功能：读取USB进行设备初始化并登录
'---------------------------------------------------------------------------------------------------------------------
    Dim strKey As String, strPIN As String, strUserName As String
    Dim strCertName As String, strCertDN As String
    Dim strCertSn As String
    Dim strCertUserID As String    '包含身份证号信息
    Dim strDate As String
    Dim strCert As String
    Dim blnOk As Boolean
    Dim blnRet As Boolean
    Dim lngRet As Long
    
    On Error GoTo errH
    

     '获取证书信息同时检查Key盘是否插入
    If Not HUBEI_GetCertList(strCertName, strCertSn, strCertID) Then
        HUBEI_CheckCert = False: Exit Function
    End If
    '未注册在当前用户名下的Key
    If strCurrCertSn <> strCertSn Then
        MsgBoxEx "该证书:" & vbCrLf & _
                vbTab & "【" & strCertSn & "】" & vbCrLf _
                & "未注册在您的名下，不能使用！", vbInformation, gstrSysName
        Exit Function
    End If
    '密码验证如果不调用,首次调用签名接口时会触发CA的密码窗口
    
    '登录验证
    If InStr(gstrLogins & "|", "|" & strCertSn & "|") > 0 Then '首次验证通过后，下次不在继续验证
        blnOk = True
    Else
        If Not GetCertLogin(strCertID) Then
            strPIN = ""
            blnOk = False
        Else
            If InStr(gstrLogins & "|", "|" & strCertSn & "|") = 0 Then gstrLogins = gstrLogins & "|" & strCertSn
            blnOk = True
        End If
    End If
    HUBEI_CheckCert = blnOk
    Exit Function
errH:
     MsgBoxEx "检查USBKEY失败！" & vbNewLine & "第" & CStr(Erl()) & "行 " & Err.Description, vbInformation, gstrSysName
End Function

Public Function HUBEI_Sign(ByVal strCurrCertSn As String, ByVal strSource As String, ByRef strSignData As String, _
            ByRef strTimeStamp As String, ByRef strTimeStampCodeID As String) As Boolean
    '签名
    '参数：
    '   strPID --用户身份标识（一般为身份证号）
        Dim strCertID As String
        Dim strPicPath As String
        Dim CertID As String
        Dim strTimeStampCode As String    '时间戳编码
        Dim strDate As String
        Dim strMsg As String
        Dim blnCheck As Boolean
        Dim lngRet As Long
    
        On Error GoTo errH

100     blnCheck = HUBEI_CheckCert(strCurrCertSn, strCertID)
    
102     If blnCheck Then                '验证当前USB是否是签名用户的，并获取签名证书
104         strSource = HUBEI_Client.SOF_HashData(strSource)   '原文转HASH
            'detach
106         Call HUBEI_Client.SOF_SetP7SignMode(1)
108         strSignData = HUBEI_Client.SOF_SignDataByP7(strCertID, strSource)
110         If strSignData <> "" Then
112             lngRet = -1
                'detach
114             Call HUBEI_SVS.SOF_SetP7SignMode(1)
116             lngRet = HUBEI_SVS.SOF_VerifyDetachSignedData(strSignData, strSource)
118             If lngRet = 0 Then
120                 strTimeStampCodeID = HUBEI_TS.SOF_CreateTimeStampResponse(strSignData)  '时间戳ID保存数据库
122                 If strTimeStampCodeID = "" Then
124                     MsgBoxEx "生成时间戳ID失败！" & HUBEI_TS.SOF_GetErrorMsg(), vbOKOnly + vbInformation, gstrSysName
                        Exit Function
                    End If
126                 strTimeStamp = Format(gobjComLib.zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
                Else
128                 MsgBoxEx "验证签名失败！" & HUBEI_SVS.SOF_GetErrorMsg(), vbOKOnly + vbInformation, gstrSysName
                    Exit Function
                End If
            Else
130             MsgBoxEx "签名失败！" & HUBEI_Client.SOF_GetErrorMsg(), vbOKOnly + vbInformation, gstrSysName
                Exit Function
            End If
        Else
            Exit Function
        End If
    
132     HUBEI_Sign = True
        Exit Function
errH:
134      MsgBoxEx "签名失败！" & vbNewLine & "第" & CStr(Erl()) & "行 " & Err.Description, vbInformation, gstrSysName
End Function


Public Function HUBEI_VerifySign(ByVal strSignData As String, ByVal strSource As String, ByVal strTimeStampID As String) As Boolean
    '验证签名
    Dim lngRet As Long
    Dim strTmp As String
    Dim blnRet As Boolean
    Dim strDate As String
    Dim strTimeStamp As String
    
    On Error GoTo errH
    
    lngRet = -1
    strSource = HUBEI_Client.SOF_HashData(strSource)   '原文转HASH
    'detach
    Call HUBEI_SVS.SOF_SetP7SignMode(1)
    lngRet = HUBEI_SVS.SOF_VerifyDetachSignedData(strSignData, strSource)
    
    If lngRet = 0 Then
        strTmp = "验证数据签名成功！"
        blnRet = True
    Else
        strTmp = "验证数据签名失败！" & HUBEI_SVS.SOF_GetErrorMsg()
        blnRet = False
    End If
    '时间戳 只是出现医疗事故之后由医疗机构进行验证 （CA提出不提供该功能）
    If strTmp <> "" Then
        MsgBoxEx strTmp, vbOKOnly + vbInformation, gstrSysName
    End If
    
    HUBEI_VerifySign = blnRet
    
    Exit Function
errH:
104     MsgBoxEx "验证签名失败！" & vbNewLine & "第" & CStr(Erl()) & "行 " & Err.Description, vbInformation, gstrSysName
End Function

Private Function GetCertLogin(ByVal strCertID As String) As Boolean
    '湖北CA数字证书登录函数
    '- 入参
    'strCertID            :证书ID
    'strPin              ：密码
    Dim strRandom As String
    Dim strSignValue As String
    Dim lngRet As Long
    Dim strDate As String
    Dim intDay As Integer
    
    On Error GoTo errH

    strRandom = HUBEI_Client.SOF_GenRandom(10)
    Call HUBEI_Client.SOF_SetP7SignMode(1)
    strSignValue = HUBEI_Client.SOF_SignDataByP7(strCertID, strRandom)

    Call HUBEI_SVS.SOF_SetP7SignMode(1)
    lngRet = -1
    lngRet = HUBEI_SVS.SOF_VerifyDetachSignedData(strSignValue, strRandom)

    If lngRet <> 0 Then
        MsgBoxEx "登陆失败！" & HUBEI_Client.SOF_GetErrorMsg(), vbOKOnly + vbInformation, gstrSysName
        Exit Function
    End If
    
    GetCertLogin = True
    Exit Function
errH:
    MsgBoxEx "登录验证失败！" & vbNewLine & "第" & CStr(Erl()) & "行 " & Err.Description, vbInformation, gstrSysName
End Function

Public Function HUBEI_GetPara() As Boolean
'设置湖北CA服务器地址
    Dim arrList As Variant
    
    On Error GoTo errH
    gstrPara = gobjComLib.zlDatabase.GetPara(90000, glngSys)
    If gstrPara = "" Then gstrPara = "221.232.224.75&&&8082&&&221.232.224.75&&&8084&&&SM2"
    If gstrPara <> "" Then
        arrList = Split(gstrPara, G_STR_SPLIT)
        If UBound(arrList) >= 3 Then
             gudtPara.strSIGNIP = Trim(arrList(0))
             gudtPara.strSignPort = Trim(arrList(1))
             gudtPara.strTSIP = Trim(arrList(2))
             gudtPara.strTSPort = Trim(arrList(3))
             If UBound(arrList) >= 4 Then
                gudtPara.bytSignVersion = IIf(Trim(arrList(4)) = "RSA", 0, 1)
            Else
                gudtPara.bytSignVersion = 0
             End If
        End If
    Else
        gudtPara.strSIGNIP = "221.232.224.75"
        gudtPara.strSignPort = 8082
        gudtPara.strTSIP = "221.232.224.75"
        gudtPara.strTSPort = 8084
        gudtPara.bytSignVersion = 0
    End If
    
    Exit Function
errH:
    MsgBoxEx "读取参数失败！" & vbNewLine & "第" & CStr(Erl()) & "行 " & Err.Description, vbInformation, gstrSysName
End Function

Public Function HUBEI_SetParaStr() As String
    HUBEI_SetParaStr = gudtPara.strSIGNIP & G_STR_SPLIT & gudtPara.strSignPort & G_STR_SPLIT & gudtPara.strTSIP & G_STR_SPLIT & gudtPara.strTSPort & G_STR_SPLIT & IIf(gudtPara.bytSignVersion = 0, "RSA", "SM2")
End Function
'销毁对象
Public Sub HUBEI_Unload()
    Set HUBEI_Client = Nothing
    Set HUBEI_SVS = Nothing
    Set HUBEI_TS = Nothing
    Set HUBEI_PIC = Nothing
    mblnInit = False
End Sub
