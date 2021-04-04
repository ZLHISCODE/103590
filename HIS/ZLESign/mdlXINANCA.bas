Attribute VB_Name = "mdlXINANCA"
Option Explicit
Private mobjSecCtrl As Object      'npmobjSecCtrl.dll
'Private mobjSecCtrl As New SecCtrlLib.CACtrlCom
Private Const mstrTitle As String = "信安CA"
Private mblnInit As Boolean
Private mblnLogin As Boolean    '是否需要登录验证

Public Function XINANCA_InitObj() As Boolean
    '证书部件初始化
    '时间戳外网测试地址 218.29.120.82 port:9198
        Dim lngRet As Long
        
        On Error GoTo ErrH
    
100     If mblnInit Then XINANCA_InitObj = True: Exit Function
        Call XINANCA_GetPara
        If gudtPara.strTSIP <> "" Then gudtPara.blnISTS = True   '开启时间戳
102     Set mobjSecCtrl = CreateObject("SecCtrl.CACtrlCom") '动态创建

106     XINANCA_InitObj = True
108     mblnInit = True
        Exit Function
ErrH:
    GetErrMsg Erl()
End Function

Public Function XINANCA_RegCert(arrCertInfo As Variant, Optional ByVal strUserID As String) As Boolean
'功能：提供在HIS数据库中注册个人证书的必要信息,用于新发放或更换证书,,需要插入USB-Key
'参数:strUserID-身份证号
'返回：arrCertInfo作为数组返回证书相关信息
'      0-ClientSignCertCN:客户端签名证书公共名称(姓名),用于注册证书时程序验证身份
'      1-ClientSignCertDN:客户端签名证书主题(每人唯一)
'      2-ClientSignCertSN:客户端签名证书序列号(每证书唯一)
'      3-ClientSignCert:客户端签名证书内容
'      4-ClientEncCert:客户端加密证书内容
'      5-签名图片文件名,空串表示没有签名图片
        
    Dim strCertUserID As String, strCertUserName As String, strCertDN As String
    Dim strCert As String, i As Integer
    Dim strCertSn As String
    Dim strPicFile As String
    On Error GoTo ErrH
    
    For i = LBound(arrCertInfo) To UBound(arrCertInfo)
        arrCertInfo(i) = ""
    Next
    
    If GetCertList(strCertUserName, strCertSn, strCertDN, strCertUserID, strCert, strPicFile) Then
        arrCertInfo(0) = strCertUserName
        arrCertInfo(1) = strCertDN '证书DN
        arrCertInfo(2) = strCertSn '证书序列号 签名时要用
        arrCertInfo(3) = strCert
        arrCertInfo(4) = ""
        arrCertInfo(5) = strPicFile
        XINANCA_RegCert = True
    End If
    Exit Function
ErrH:
    MsgBoxEx "获取证书信息失败！" & vbNewLine & _
        "第" & CStr(Erl()) & "行 " & Err.Description, vbExclamation, mstrTitle
End Function

Private Function GetCertList(ByRef strName As String, Optional ByRef strCertSn As String = "0", Optional ByRef strCertDN As String = "0", _
           Optional ByRef strCertUserID As String = "0", Optional ByRef strCert As String = "0", Optional ByRef strPicFile As String = "0") As Boolean
'功能:信安CA获取证书详情
'-入参:无
'-出参
'strName :      保存接口返回的证书所有者姓名
'strCertSN      保存接口返回的证书SN
'strCertDN:     保存接口返回的证书DN
'strCertUserID:  保存接口返回的证书所有者唯一标识
'strCert:       保存接口返回的签名证书

        On Error GoTo ErrH
   
        Dim lngRet As Long
        Dim strPic As String
        
100     lngRet = mobjSecCtrl.KS_SetProv("XACA", 0, "") '初始化 未查KEY时会弹出警告提示
102     If mobjSecCtrl.KS_GetLastErrorCode() <> 0 Then
104         MsgBoxEx mobjSecCtrl.KS_GetLastErrorMsg(), vbExclamation + vbOKOnly, mstrTitle
            Exit Function
        End If
106     strCert = GetCert(2) 'type: 1-加密证书，2-签名证书
108     If strCert = "" Then MsgBoxEx "读取证书信息失败！", vbExclamation, mstrTitle: Exit Function
        '15-证书拥有者部门名(OU): 小龙测试
110      If Not GetCertInfo(strCert, 15, strName) Then Exit Function
        '17-证书拥有者通用名称(CN):4127589685665
112      If strCertSn <> "0" Then If Not GetCertInfo(strCert, 17, strCertSn) Then Exit Function
    
        '21-证书拥有者DN:C=CN,S=河南省,L=郑州市,O=河南省地方税务局,OU=小龙测试,CN=4127589685665
114     If strCertDN <> "0" Then
116          If Not GetCertInfo(strCert, 21, strCertDN) Then Exit Function
        End If
        
        If strPicFile <> "0" Then
            If Not XINANCA_GetSeal(strPic) Then Exit Function
            strPicFile = FormatPic("gif", strCertSn, strPic)
        End If
        
        GetCertList = True
        Exit Function
ErrH:
        MsgBoxEx Err.Description & vbCrLf & _
                "在GetCertList 错误行: " & Erl, _
                    vbExclamation + vbOKOnly, mstrTitle
         
End Function
 
Private Function XINANCA_GetSeal(ByRef strSeal As String) As Boolean
      Dim strFileName As String
      Dim strTemp As String


10       On Error GoTo ErrH
      'mobjSecCtrl.KS_SetProv("XACA", 0, "")初始化成功后再调用
20    strFileName = mobjSecCtrl.KS_GetSealList()
30    If mobjSecCtrl.KS_GetLastErrorCode() Then
40        MsgBoxEx "得到印章列表错误:" & mobjSecCtrl.KS_GetLastErrorMsg(), vbExclamation + vbOKOnly, mstrTitle
50        Exit Function
60    End If
70    strTemp = mobjSecCtrl.KS_GetSeal(strFileName)
80    If mobjSecCtrl.KS_GetLastErrorCode() Then
90        MsgBoxEx "得到印章数据失败:" & mobjSecCtrl.KS_GetLastErrorMsg(), vbExclamation + vbOKOnly, mstrTitle
100       Exit Function
110   End If
120   strSeal = mobjSecCtrl.KS_GetInfoFromSeal(strTemp, 1)
130   If mobjSecCtrl.KS_GetLastErrorCode() Then
140       MsgBoxEx "读取图片错误:" & mobjSecCtrl.KS_GetLastErrorMsg(), vbExclamation + vbOKOnly, mstrTitle
150       Exit Function
160   End If
      strSeal = Replace(strSeal, vbLf, "")  '对方返回字符包含换行符
      WriteLog "图片信息:" & strSeal
170   XINANCA_GetSeal = True
180   Exit Function
ErrH:
190           MsgBoxEx Err.Description & vbCrLf & _
                      "在XINANCA_GetSeal 错误行: " & Erl, _
                          vbExclamation + vbOKOnly, mstrTitle
End Function

Public Function XINANCA_CheckCert() As Boolean
    '功能:
    '   1-检查证书是否插上
    '   2-检查当前证书是否注册在当前用户名下
        Dim strName As String
        Dim strCertSn As String
        Dim strCert As String
        Dim strPIN As String
        Dim lngResult As Long
        
        On Error GoTo ErrH
100     If Not GetCertList(strName, strCertSn) Then XINANCA_CheckCert = False: Exit Function
102     If strCertSn <> mUserInfo.strCertSn Then
104         MsgBoxEx "该证书未注册在您的名下，不能使用！" & vbCrLf & _
                    "用户注册证书唯一标识:" & mUserInfo.strCertSn & vbCrLf & _
                    "当前所选证书唯一标识:" & strCertSn, vbInformation + vbOKOnly, gstrSysName
            Exit Function
        End If
106     If Not mblnLogin Then
108         If Not Login() Then
                Exit Function
            Else
110             mblnLogin = True
            End If
        End If
112     XINANCA_CheckCert = True
        Exit Function
ErrH:
114     MsgBoxEx Err.Description & vbCrLf & _
                "XINANCA_CheckCert 错误行: " & Erl, _
                    vbExclamation + vbOKOnly, mstrTitle
End Function

Private Function Login() As Boolean
    '功能:信安CA数字证书登录函数
    '- 入参
    'strCertID            :证书ID
    'strCert              证书内容BASE64编码
    Dim strRandom As String, strSignVal As String
    Dim strDate As String
    Dim intDay As Integer
 
    Dim lngRet As Long
    
        On Error GoTo ErrH
         
100     strRandom = GenRandom(16)  '获取随机数
102     If strRandom = "" Then Exit Function
104     strSignVal = SignDataByP7(strRandom, 0)
106     If strSignVal = "" Then Exit Function
108     lngRet = VerifySignData(strRandom, strSignVal)
110     If mobjSecCtrl.KS_GetLastErrorCode() <> 0 Then
112         MsgBoxEx "随机数验签失败！" & vbNewLine & mobjSecCtrl.KS_GetLastErrorMsg(), vbExclamation + vbOKOnly, mstrTitle
            Exit Function
        End If
114     Login = True
        Exit Function
ErrH:
116     MsgBoxEx "登录验证失败！" & vbNewLine & "第" & CStr(Erl()) & "行 " & Err.Description, vbExclamation, gstrSysName
End Function

Public Function XINANCA_Sign(ByVal strSource As String, ByRef strSignData As String, _
            ByRef strTimeStamp As String, ByRef strTimeStampCode As String, Optional ByVal blnCheck As Boolean) As Boolean
    '功能:
    Dim strURL As String
    Dim strParameter As String
    Dim bytRet() As Byte
    Dim varTemp As Variant
    
        On Error GoTo ErrH
100     If Not blnCheck Then
102         blnCheck = XINANCA_CheckCert()
        End If
    
104     If blnCheck Then
106         strSource = StringSHA1(strSource)
108         strSignData = SignDataByP7(strSource, 0)
110         If strSignData = "" Then
112             MsgBoxEx "签名失败！", vbExclamation, mstrTitle
                Exit Function
            End If
        Else
114         MsgBoxEx "签名失败！", vbExclamation, mstrTitle
            Exit Function
        End If
        '启用时间戳
116     If gudtPara.blnISTS Then
118         strURL = "http://" & gudtPara.strTSIP & ":" & gudtPara.strTSPort & "/tsac.svr"
120         strParameter = "digest=" & strSource
122         bytRet = HttpPost(strURL, strParameter, responseBody)
124         strTimeStampCode = EncodeBase64Byte(bytRet)
126         If strTimeStampCode = "" Then
128             MsgBoxEx "获取时间戳信息失败！", vbExclamation, mstrTitle
                Exit Function
            End If
130         strURL = "http://" & gudtPara.strTSIP & ":" & gudtPara.strTSPort & "/tsav.svr"
132         strParameter = "tsr=" & Replace(strTimeStampCode, "+", "%2B")
134         strTimeStamp = HttpPost(strURL, strParameter, responseText)
            LogWrite "XINANCA_Sign", "时间戳返回值：" & strTimeStamp
136         If strTimeStamp = "" Then
138             MsgBoxEx "获取时间戳失败！", vbExclamation, mstrTitle
                Exit Function
            Else
140             strTimeStamp = Mid(strTimeStamp, InStr(strTimeStamp, "<timestamp>") + Len("<timestamp>"))
142             strTimeStamp = Mid(strTimeStamp, 1, InStr(strTimeStamp, "</timestamp>") - 1)
                strTimeStamp = Replace(strTimeStamp, Space(2), Space(1))  '日期为一位数时前面可能存在空格导致解析失败
144             varTemp = Split(strTimeStamp, Space(1)) 'Jan 21 06:34:28.865495 2019 GMT 时间只取前八位字符
146             strTimeStamp = varTemp(3) & "-" & ConvMonth(varTemp(0)) & "-" & varTemp(1) & " " & Mid(varTemp(2), 1, 8) '格林威治时间
148             strTimeStamp = Format(DateAdd("h", 8, strTimeStamp), "YYYY-MM-DD HH:MM:SS")
            End If
        Else
150         strTimeStamp = Format(gobjComLib.zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
        End If
152     XINANCA_Sign = True
        Exit Function
ErrH:
154       MsgBoxEx "签名失败！" & vbNewLine & "第" & CStr(Erl()) & "行 " & Err.Description, vbExclamation, mstrTitle
End Function

Public Function XINANCA_VerifySign(ByVal strSignData As String, ByVal strSource As String, ByVal strTimeStampCode As String) As Boolean
    '功能:
        Dim lngRet As Long
        Dim strURL As String
        Dim strParameter As String
        
        On Error GoTo ErrH
100     strSource = StringSHA1(strSource)
102     lngRet = VerifySignData(strSource, strSignData)
104     If lngRet <> 0 Then
106         MsgBoxEx "验证失败，该电子签名数据有效!", vbInformation, mstrTitle
            Exit Function
        End If
108     If gudtPara.blnISTS Then
110         strURL = "http://" & gudtPara.strTSIP & ":" & gudtPara.strTSPort & "/tsav.svr"
112         strParameter = "tsr=" & Replace(strTimeStampCode, "+", "%2B")
114         strTimeStampCode = HttpPost(strURL, strParameter, responseText)
116         If strTimeStampCode = "" Then
118             MsgBoxEx "验证时间戳失败！", vbExclamation, mstrTitle
                Exit Function
            End If
        End If
120     MsgBoxEx "验证成功，该电子签名数据有效!", vbInformation, mstrTitle
122     XINANCA_VerifySign = True
    Exit Function
ErrH:
124  MsgBoxEx "验证签名失败！" & vbNewLine & "第" & CStr(Erl()) & "行 " & Err.Description, vbExclamation, mstrTitle
End Function


Public Function XINANCA_UnLoad()
    Set mobjSecCtrl = Nothing
    mblnInit = False
End Function
'/**
' * 获取BASE64编码证书
' * type: 1-加密证书，2-签名证书
' */
Private Function GetCert(ByVal lngType As Long) As String
    Dim strResult As String
    
    strResult = mobjSecCtrl.KS_GetCert(lngType)
    If mobjSecCtrl.KS_GetLastErrorCode() <> 0 Then
        MsgBoxEx mobjSecCtrl.KS_GetLastErrorMsg(), vbExclamation + vbOKOnly, mstrTitle
        Exit Function
    End If
    GetCert = strResult
End Function

'/**
' * 获取证书信息
' * cert: Base64编码证书
' * item: 解析项。
' * 1、证书版本 2、证书序列号 3、证书签名算法标识 4、证书颁发者国家(C)  5、证书颁发者组织名(O)
' * 6、证书颁发者部门名(OU)  7、证书颁发者所在的省、自治区、直辖市(S)  8、证书颁发者通用名称(CN)  9、证书颁发者所在的城市、地区(L)
' * 10、证书颁发者Email  11、证书有效期：起始日期:180410101818  12、证书有效期：终止日期:190410101818  13、证书拥有者国家(C )  14、证书拥有者组织名(O)
' * 15、证书拥有者部门名(OU)  16、证书拥有者所在的省、自治区、直辖市(S)  17、证书拥有者通用名称(CN)  18、证书拥有者所在的城市、地区(L)
' * 19、证书拥有者Email  20、证书颁发者DN  21、证书拥有者DN  22、证书公钥信息  23、CRL发布点.
' */
Private Function GetCertInfo(ByVal strCert As String, ByVal lngItem As Long, ByRef strResult As String) As Boolean
     
    strResult = mobjSecCtrl.KS_GetCertInfo(strCert, lngItem)
    If mobjSecCtrl.KS_GetLastErrorCode() <> 0 Then
        MsgBoxEx mobjSecCtrl.KS_GetLastErrorMsg(), vbExclamation + vbOKOnly, mstrTitle
        Exit Function
    End If
    GetCertInfo = True
End Function
 
'/**
' * 获取证书扩展信息
' * cert: Base64编码证书
' * oid: oid值
' */
Private Function GetCertInfoByOid(ByVal strCert As String, ByVal strOid As String) As String
    Dim strResult As String
    strResult = mobjSecCtrl.KS_GetCertInfoByOid(strCert, strOid)
    If mobjSecCtrl.KS_GetLastErrorCode() <> 0 Then
        MsgBoxEx mobjSecCtrl.KS_GetLastErrorMsg(), vbExclamation + vbOKOnly, mstrTitle
    End If
    GetCertInfoByOid = strResult
End Function

'/**
' * 生成随机数
' * len: 随机数长度
' */
Private Function GenRandom(ByVal lngLen As Long) As String
    Dim strResult As String
    strResult = mobjSecCtrl.KS_GenRandom(lngLen)
    If mobjSecCtrl.KS_GetLastErrorCode() <> 0 Then
        MsgBoxEx mobjSecCtrl.KS_GetLastErrorMsg(), vbExclamation + vbOKOnly, mstrTitle
        Exit Function
    End If
    GenRandom = strResult
End Function

Private Function VerifySignData(ByVal strSource As String, ByVal strSignData As String) As Long
'功能:服务器验证签名
    Dim lngResult As Long

    lngResult = mobjSecCtrl.KS_P7RemoteVerify(1, strSignData, strSource)  '返回的结果是整形，0为成功，非0为失败
    If mobjSecCtrl.KS_GetLastErrorCode() <> 0 Then
        MsgBoxEx mobjSecCtrl.KS_GetLastErrorMsg(), vbExclamation + vbOKOnly, mstrTitle
    End If
    VerifySignData = lngResult
End Function

'/**
' * 数据签名P7
' * indata：明文数据
' * hashAlg:0. AUTO(自动选择，当RSA时为SHA1, SM2时为SM3), 1-SHA1, 2-SHA256, 3-SHA512, 4-MD5, 5-MD4, 6-SM3
' * return：签名数据
' */
Private Function SignDataByP7(ByVal strSource As String, ByVal lngHashAlg As Long) As String
    Dim strResult As String

    strResult = mobjSecCtrl.KS_SetParam("signtype", "pksc7")
    strResult = mobjSecCtrl.KS_SignData(strSource, lngHashAlg)
    If mobjSecCtrl.KS_GetLastErrorCode() <> 0 Then
        MsgBoxEx mobjSecCtrl.KS_GetLastErrorMsg(), vbExclamation + vbOKOnly, mstrTitle
        Exit Function
    End If
    SignDataByP7 = strResult
End Function

 Public Function XINANCA_GetPara() As Boolean
    '设置服务器地址
    
    On Error GoTo ErrH
     
    'If gstrPara = "" Then gstrPara = "192.168.20.203" & G_STR_SPLIT & "9198"
    '外网测试地址 218.29.120.82 port:9198
    'gudtPara.strTSIP="" 代表不开启时间戳功能
    gudtPara.strTSIP = GetThirdPara(CON_PAR_信安, "时间戳IP")
    gudtPara.strTSPort = GetThirdPara(CON_PAR_信安, "时间戳端口")
   
    Exit Function
ErrH:
    MsgBoxEx "读取参数失败！" & vbNewLine & "第" & CStr(Erl()) & "行 " & Err.Description, vbInformation, gstrSysName
End Function

Private Function URLEncode(ByVal strParameter As String) As String
    Dim strTemp As String
    Dim i As Integer
    Dim intValue As Integer

    Dim bytData() As Byte

    strTemp = ""
    bytData = StrConv(strParameter, vbFromUnicode)
    For i = 0 To UBound(bytData)
        intValue = bytData(i)
        If (intValue >= 48 And intValue <= 57) Or _
            (intValue >= 65 And intValue <= 90) Or _
            (intValue >= 97 And intValue <= 122) Then
            strTemp = strTemp & Chr(intValue)
        ElseIf intValue = 32 Then
            strTemp = strTemp & "+"
        Else
            strTemp = strTemp & "%" & LCase(Hex(intValue))
        End If
    Next
    URLEncode = strTemp
End Function
