Attribute VB_Name = "mdlXJCA"
Option Explicit
'新疆CA中心功能模块

Private mblnInit As Boolean         '是否已初始化成功
Private mobjXJCA_Client As Object
Private mstrPara  As String
Private mobjGseal As Object
Private mbytType As Byte '1-精河县人民医院使用的是海泰KEY
                         '2-奇台医院使用的是华大KEY
'
Private Const M_STR_CSP  As String = "HaiTai Cryptographic Service Provider for xjca"
Private Const M_STR_CSP_HD  As String = "CIDC Cryptographic Service Provider v1.0.0"

'电子签章接口声明(C++)
'Private Declare Function XJCA_SignSeal Lib "XJCA_HOS.dll" (ByVal strSrc As String, ByVal lngxml As Long, ByVal lngLen As Long) As Boolean
'说明：bool   XJCA_SignSeal(char* src,char* signxml, DWORD* len)
'参数说明:strSrc-数据源,lngXml--传地址（用byte数组接收字符串,用chr函数转换）,lngLen-传地址 因为传人的是地址,所以用long型
'返回值：True\false
Private Declare Function XJCA_GetSealBMPB Lib "XJCA_HOS.dll" (ByVal strFilePath As String, ByVal lngTimes As Long) As Boolean
'功能:获取签章图片
'参数:
Private Declare Function XJCA_VerifySeal Lib "XJCA_HOS.dll" (ByVal strSrc As String, ByVal strxml As String, ByVal strPic As String, ByVal strCert As String) As Boolean
'功能:验证签章数据.接口原型  bool   XJCA_VerifySeal((char* src,char* xml,char* pct,char* cert)
'参数：
'返回：True\false


Public Function XJCA_InitObj() As Boolean
'功能:证书部件初始化
    Dim strUrl As String

102     XJCA_InitObj = mblnInit
104     If mblnInit Then Exit Function
        On Error Resume Next
        mstrPara = gobjComLib.zlDatabase.GetPara(90000, glngSys)  '读取配置内容
        If mstrPara = "" Then
            Err.Raise -1, , "配置文件读取失败，请先配置。"
            Exit Function
        End If
        On Error GoTo 0: Err.Clear
        
        On Error GoTo errH
        If InStr(mstrPara, "华大") = 0 Then
            mbytType = 1
            '缺省海泰 KEY
108         Set mobjXJCA_Client = CreateObject("xjcaTechATL.xjcaTechATLLib.1")
            Set mobjGseal = CreateObject("Signature.SignatureForm")      '采用SignatureForm控件,而不使用XJCA_HOS.dll中的签章函数 存在内存问题
            If mobjXJCA_Client Is Nothing Or mobjGseal Is Nothing Then
                MsgBoxEx "CA对象创建失败！", vbOKOnly, gstrSysName
                Exit Function
            End If
        Else
            mbytType = 2
            '华大KEY
            Set mobjXJCA_Client = CreateObject("XjcaFgwATL.XjcaFgwATLLib.1")
            Set mobjGseal = CreateObject("XJFormSeal.XJFormSealX")
            If mobjXJCA_Client Is Nothing Or mobjGseal Is Nothing Then
                MsgBoxEx "CA对象创建失败！", vbOKOnly, gstrSysName
                Exit Function
            End If
        End If
114     XJCA_InitObj = True
        
116     mblnInit = XJCA_InitObj
        Exit Function
errH:
118     MsgBoxEx "创建新疆CA接口部件失败！" & vbNewLine & Err.Description, vbQuestion, gstrSysName

End Function

Public Function XJCA_RegCert(arrCertInfo As Variant) As Boolean
        '功能：提供在HIS数据库中注册个人证书的必要信息,用于新发放或更换证书,,需要插入USB-Key
        '返回：arrCertInfo作为数组返回证书相关信息
        '      0-ClientSignCertCN:客户端签名证书公共名称(姓名),用于注册证书时程序验证身份
        '      1-ClientSignCertDN:客户端签名证书主题(每人唯一)
        '      2-ClientSignCertSN:客户端签名证书序列号(每证书唯一)
        '      3-ClientSignCert:客户端签名证书内容
        '      4-ClientEncCert:客户端加密证书内容
        '      5-签名图片文件名,空串表示没有签名图片
        Dim strKeyId As String, strCertTime As String, strCertUserName As String, strCertDN As String
        Dim strSigCert As String
        Dim strFile As String
        Dim blnRet As Boolean
        Dim i As Long
        On Error GoTo errH
        If Not CheckIsXJCA Then Exit Function
100     For i = LBound(arrCertInfo) To UBound(arrCertInfo)
101         arrCertInfo(i) = ""
102     Next
       
108     If XJCA_GetCertList(strCertUserName, strKeyId, strSigCert, strCertDN) Then
200         arrCertInfo(0) = strCertUserName
201         arrCertInfo(1) = strCertDN
202         arrCertInfo(2) = strKeyId
203         arrCertInfo(3) = strSigCert
205         arrCertInfo(4) = ""
            strFile = App.Path & "\" & Format(Now, "yyyyMMdd") & "_" & strKeyId & ".BMP"
            blnRet = XJCA_GetSealBMPB(strFile, 2)
            If blnRet = False Then Exit Function
206         arrCertInfo(5) = strFile
            XJCA_RegCert = True
        End If

300     Exit Function

errH:
    MsgBoxEx "获取证书信息失败！" & vbNewLine & "第" & CStr(Erl()) & "行 " & Err.Description, vbQuestion, gstrSysName
End Function

Public Function XJCA_CheckCert(ByVal strCurrCertSn As String, Optional ByRef strSigCert As String) As Boolean
'功能：读取USB进行设备初始化并登录
'返回值:
'  strSigCert -签名证书内容

        Dim strKey As String
        Dim strUserName As String
        Dim strCertDN As String
        
        Dim lngRet As Long
        
        On Error GoTo errH
        
        If Not XJCA_InitObj() Then
             MsgBoxEx "部件未初始化！", vbInformation + vbOKOnly, gstrSysName
            Exit Function
        End If
        
100     If Not CheckIsXJCA Then Exit Function
   
104     If XJCA_GetCertList(strUserName, strKey) Then
106        If strCurrCertSn <> strKey Then
108            MsgBoxEx "该证书未注册在您的名下，不能使用！"
               Exit Function
           End If
110
116        If Not GetCertLogin(strKey, strUserName) Then
                Exit Function
           End If
122     End If
        
        XJCA_CheckCert = True
        Exit Function
errH:
124     MsgBoxEx "检查USBKEY失败！" & vbNewLine & "第" & CStr(Erl()) & "行 " & Err.Description, vbQuestion, gstrSysName
End Function

Public Function XJCA_GetCertList(Optional ByRef strName As String = "1", Optional ByRef strUniqueID As String = "1", Optional ByRef strCert As String = "1", Optional ByRef strCertDN As String = "1") As Boolean
    '-入参:无
    '-出参
    'strName :      保存接口返回的证书所有者姓名
    'strUniqueID:   保存接口返回的证书所有者唯一标识
    'strCert:       保存接口返回的签名证书
    On Error GoTo errH
    Dim strSrc As String
    Dim strTmp As String
    Dim blnRet As Boolean
    Dim arrTmp As Variant
    Dim i As Long
    
    On Error GoTo errH
    If mbytType = 1 Then
        If strUniqueID <> "1" Then
            frmXJCA.txtValue.Text = CStr(mobjXJCA_Client.XJCA_GetCertSN)   '证书序号
            strUniqueID = Trim(frmXJCA.txtValue.Text)
        End If
        If strName <> "1" Or strCertDN <> "1" Then
            arrTmp = Array()
            ReDim Preserve arrTmp(UBound(arrTmp) + 1)
            arrTmp(UBound(arrTmp)) = CStr(mobjXJCA_Client.XJCA_GetCertDN())     'C=CN, S=650105197001010026, L=0026, O=新疆CA, OU=CA中心, E=xjcaxmss@xjca.com.cn, CN=新疆CA接口测试0026
            strTmp = arrTmp(UBound(arrTmp))
            strCertDN = strTmp
            strName = Mid(strCertDN, InStr(strCertDN, "CN=") + 3)   '获取证书持有者姓名
        End If
        
        If strCert <> "1" Then
            strSrc = "1234567890"
            Call mobjGseal.XJCASetFieldByName("IsNeedCert", "true") '调用签章接口前先调用它,否则报错,XJCASowSignInSvr
            Call mobjGseal.XJCASowSignInSvr(strSrc, strTmp)
            If strTmp <> "" Then
                strCert = Split(strTmp, ",")(1) & "," & Split(strTmp, ",")(2) '证书信息 '证书ID
            Else
                Exit Function
            End If
        End If
    Else
        '华大KEY
        If strUniqueID <> "1" Then
            strUniqueID = CStr(mobjXJCA_Client.XJCA_GetCertSN(M_STR_CSP_HD))   '证书序号
        End If
        If strName <> "1" Or strCertDN <> "1" Then
            arrTmp = Array()
            ReDim Preserve arrTmp(UBound(arrTmp) + 1)
            arrTmp(UBound(arrTmp)) = CStr(mobjXJCA_Client.XJCA_GetCertDN(M_STR_CSP_HD))      'C=CN, S=650105197001010026, L=0026, O=新疆CA, OU=CA中心, E=xjcaxmss@xjca.com.cn, CN=新疆CA接口测试0026
            strTmp = arrTmp(UBound(arrTmp))
            strCertDN = strTmp
            strName = Mid(strCertDN, InStr(strCertDN, "CN=") + 3)   '获取证书持有者姓名
        End If
        
        If strCert <> "1" Then
            strSrc = "1234567890"
            strTmp = ""
            Call mobjGseal.XJCASowSignInSvr(strSrc, strTmp)
            If strTmp <> "" Then
                strCert = Split(strTmp, ",")(1) & "," & Split(strTmp, ",")(2) '证书信息 '证书ID
            Else
                Exit Function
            End If
        End If
    End If
    XJCA_GetCertList = True
    Exit Function
errH:
    MsgBoxEx "读取证书信息失败！" & vbNewLine & "第" & CStr(Erl()) & "行 " & Err.Description, vbQuestion, gstrSysName

End Function

Public Function XJCA_Sign(ByVal strCurrCertSn As String, ByVal strSource As String, ByRef strSignData As String) As Boolean
'功能:新疆CA电子签名
'参数：strCurrCertSn  -证书ID(唯一序列)
'     strSource-需要签名的源数据
'     strTimeStamp-时间戳
'     strTimeStampCode-时间戳信息
'返回值：true 成功,False -失败
'       strSignData-签名后返回的签名数据
'       strTimeStamp-返回的时间戳

        Dim strTmp As String
        Dim bytXml(40000)  As Byte    '签章信息
        Dim lngLen As Long
        Dim i As Long
        Dim J As Long
        
        Dim blnRet As Boolean
        
        On Error GoTo errH
        
100     If XJCA_CheckCert(strCurrCertSn) Then                '验证当前USB是否是签名用户的，并获取签名证书
            '空格、vbTAb,vbCrLF 传人签名接口时，统一数据源返回的签名值有可能不一致
'            strSource = Replace(strSource, " ", "")
'            strSource = Replace(strSource, vbTab, "|")
'            strSource = Replace(strSource, vbCrLf, "||")
            If mbytType = 1 Then
                Call mobjGseal.XJCASetFieldByName("IsNeedCert", "true") '调用签章接口前先调用它,否则报错,XJCASowSignInSvr
                Call mobjGseal.XJCASowSignInSvr(strSource, strTmp)
            Else
                Call mobjGseal.XJCASowSignInSvr(strSource, strTmp)
            End If
            If strTmp <> "" Then
                strSignData = Split(strTmp, ",")(0) '签名数据
            Else
                MsgBoxEx "签名失败！": Exit Function
            End If
        End If
 
        XJCA_Sign = True
        Exit Function
errH:
114     MsgBoxEx "签名失败！" & vbNewLine & "第" & CStr(Erl()) & "行 " & Err.Description, vbQuestion, gstrSysName
End Function


Private Function GetCertLogin(ByVal strUniqueID As String, ByVal strName As String) As Boolean
    'strUniqueID : 证书唯一标识
    'strWebURL：
    
    Dim strTmp As String
    Dim strWebUrl As String
    Dim strAppId As String
         
'    lngResult = mobjXJCA_Client.XJCA_VerifyPin(strPassword, Len(strPassword))

    '服务器端验证证书
'   mstrPara = "http://124.117.245.71:18080/webServices/authService|4028e48a39dd529a0139dd5c383d0010"
'   mstrPara =http://124.117.245.71:48080/webServices/ssoService|4028f6d24a2d7182014a2d83333e001a|华大   华大KEY
    On Error Resume Next
    strWebUrl = Split(mstrPara, "|")(0)
    strAppId = Split(mstrPara, "|")(1)
    Err.Clear: On Error GoTo 0

    strTmp = mobjXJCA_Client.XJCA_CertAuth(strWebUrl, strAppId, strName)
    
    '验证证书结果信息表示
    If strTmp <> "" Then
        strTmp = UCase(strTmp)
        strTmp = Mid(strTmp, InStr(strTmp, UCase("<success>")) + 9)
        strTmp = Mid(strTmp, 1, InStr(strTmp, UCase("</success>")) - 1)
        If strTmp = "FALSE" Then
            MsgBoxEx "登录认证失败!", vbInformation + vbOKOnly, gstrSysName
            GetCertLogin = False: Exit Function
        End If
    Else
        MsgBoxEx "证书验证返回值为空！"
        GetCertLogin = False: Exit Function
    End If
    GetCertLogin = True
End Function

Public Function XJCA_VerifySign(ByVal strCert As String, ByVal strSignData As String, ByVal strSource As String) As Boolean
'功能;验证签名
'参数:strCurrCertSn -证书ID(唯一序列)
'     strCert -证书信息（含公钥信息）
'     strSignData-签名值
'     strSource-待验证源文

        Dim blnRet As Boolean
        
        On Error GoTo errH
'        '空格、vbTAb,vbCrLF 传人签名接口时，统一数据源返回的签名值有可能不一致
'        strSource = Replace(strSource, " ", "")
'        strSource = Replace(strSource, vbTab, "|")
'        strSource = Replace(strSource, vbCrLf, "||")
'
        blnRet = XJCA_VerifySeal(strSource, strSignData & "," & strCert, "", "")
        If blnRet Then
            MsgBoxEx "验证成功，该电子签名数据有效!", vbInformation, gstrSysName
        Else
            MsgBoxEx "验证签名失败！", vbExclamation, gstrSysName
            Exit Function
        End If
       
        XJCA_VerifySign = True
        Exit Function
errH:
104     MsgBoxEx "验证签名失败！" & vbNewLine & "第" & CStr(Erl()) & "行 " & Err.Description, vbQuestion, gstrSysName
End Function

Private Function CheckIsXJCA() As Boolean
'功能：检查CA环境
    Dim lngRet As Long
    Dim strTmp As String
    
    If mbytType = 1 Then
        strTmp = M_STR_CSP
    Else
        strTmp = M_STR_CSP_HD
    End If
    '1-判断证书驱动是否安装
    lngRet = mobjXJCA_Client.XJCA_CspInstalled(strTmp)
    If lngRet <> 10000 Then
        MsgBoxEx "证书驱动未安装！", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    '2-判断证书是否就绪
    lngRet = mobjXJCA_Client.XJCA_KeyInsert(strTmp)
    If lngRet <> 10000 Then
        MsgBoxEx "证书KEY未插入！", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    CheckIsXJCA = True
End Function

Public Sub XJCA_UnLoadObj()
    Set mobjXJCA_Client = Nothing
    Set mobjGseal = Nothing
    mblnInit = False
End Sub

Public Function XJCA_GetPara() As Boolean
'设置湖北CA服务器地址
    
    On Error GoTo errH
    gstrPara = gobjComLib.zlDatabase.GetPara(90000, glngSys)  '读取URLs 固定读取ZLHIS 系统默认100
    If gstrPara = "" Then gstrPara = "http://124.117.245.71:48080/webServices/ssoService|4028f6d24a2d7182014a2d83333e001a|华大"
    If gstrPara <> "" Then
        gudtPara.strSignURL = gstrPara
    End If
    Exit Function
errH:
    MsgBoxEx "读取参数失败！" & vbNewLine & "第" & CStr(Erl()) & "行 " & Err.Description, vbInformation, gstrSysName
End Function

Public Function XJCA_SetParaStr() As String
    XJCA_SetParaStr = gudtPara.strSignURL
End Function
