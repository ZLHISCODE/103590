Attribute VB_Name = "mdlJSCA"
Option Explicit
'江苏CA中心功能模块（江宁医院）
Private mblnInit As Boolean         '是否已初始化成功
Private JSCA_Client As Object       '电子签名部件
Private JSCA_Seal As Object         '电子签章部件
Private mobjSoapClient As Object        'soap连接对象
Private mstrKeys  As String             '插入的key标识
Private mstrSelectKey As String         '多个key的情况下,记录选择的key
'
Private Const GET_CERT_AFTER As Integer = 18   '证书有效期：终止日期
Private Const GET_USER_NAME  As Integer = 23            '证书持有者姓名

Public Function JSCA_InitObj() As Boolean
    '证书部件初始化
    Dim strUrl As String
    
        On Error GoTo errH

102     JSCA_InitObj = mblnInit
104     If mblnInit Then Exit Function

        On Error Resume Next
106     Set mobjSoapClient = CreateObject("MSSOAP.SoapClient30")  'SOAP连接对象，
        mobjSoapClient.ClientProperty("ServerHTTPRequest") = True
        
        If Err.Number <> 0 Then
            MsgBoxEx "系统初始化失败！" & vbCrLf & vbCrLf & "客户端未安装SOAP！" & vbCrLf & vbCrLf & "错误信息如下：" & vbCrLf & vbCrLf & Err.Description, vbCritical, "电子签名部件"
            Exit Function
        End If
        Err.Clear: On Error GoTo 0
        strUrl = gobjComLib.zlDatabase.GetPara(90000, glngSys)  '读取URL
'        strURL = "http://202.102.85.153:8080/HealthWebService.asmx?WSDL"
        If strUrl = "" Then
            Err.Raise -1, , "没有配置签名服务器地址，请先配置。"
            Exit Function
        End If
        On Error Resume Next
        mobjSoapClient.MSSoapInit (strUrl)
        If Err.Number <> 0 Then
            MsgBoxEx "系统初始化失败！" & vbCrLf & vbCrLf & "服务器地址出错！" & vbCrLf & vbCrLf & "错误信息如下：" & vbCrLf & vbCrLf & Err.Description, vbCritical, "电子签名部件"
            Exit Function
        End If
        Err.Clear: On Error GoTo 0
108     Set JSCA_Client = CreateObject("CACltCore.CltCore")
        Set JSCA_Seal = CreateObject("GSEAL.GSealCtrl.1")
        JSCA_Client.IsShowError = 0     '=0时,调用SOF_ShowErrMsg()可以弹出错误信息
        
114     JSCA_InitObj = True
    
116     mblnInit = JSCA_InitObj
        Exit Function
errH:
118     MsgBoxEx "创建江苏CA接口部件失败！" & vbNewLine & Err.Description, vbQuestion, "电子签名部件"
    
End Function

Public Function JSCA_RegCert(arrCertInfo As Variant) As Boolean
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
        Dim i As Long
        On Error GoTo errH
    
100     For i = LBound(arrCertInfo) To UBound(arrCertInfo)
101         arrCertInfo(i) = ""
102     Next
        
108     If JSCA_GetCertList(strCertUserName, strKeyId, strSigCert) Then
200         arrCertInfo(0) = strCertUserName
201         arrCertInfo(1) = JSCA_Client.SOF_GetUserInfo(strKeyId, 4) 'C=CN, S=江苏省, L=南京市, O=江苏省电子商务证书认证中心有限责任公司, OU=JSCA, CN=JSCA_CA
202         arrCertInfo(2) = strKeyId
203         arrCertInfo(3) = strSigCert
205         arrCertInfo(4) = ""
            strFile = App.Path & "\" & Format(Now, "yyyyMMdd") & "_" & strKeyId & ".gif"
            Call JSCA_Seal.JSCAGetSealPath(strFile)     '从key盘读取图片
            
206         arrCertInfo(5) = strFile
            JSCA_RegCert = True
        End If
        
300     Exit Function

errH:
    MsgBoxEx "获取证书信息失败！" & vbNewLine & "第" & CStr(Erl()) & "行 " & Err.Description, vbQuestion, "电子签名部件"
End Function

Public Function JSCA_GetCertList(ByRef strName As String, Optional ByRef strUniqueID As String, Optional ByRef strCert As String, Optional ByVal strUserUnigueID As String) As Boolean
    '江苏CA 江宁医院
    '-入参:无
    '     strUserUnigueID-当前用户绑定的Key，多key情况下用
    '-出参
    'strName :      保存接口返回的证书所有者姓名
    'strUniqueID:   保存接口返回的证书所有者唯一标识
    'strCert:       保存接口返回的签名证书
    Dim strTmp As String
    Dim strkeys As String
    Dim arrKey As Variant
    Dim intCount As Integer
    Dim blnChange As Boolean
    Dim i As Integer
    On Error GoTo errH
    
    If JSCA_Client Is Nothing Then Set JSCA_Client = CreateObject("CACltCore.CltCore")
    If Not JSCA_Client Is Nothing Then
        '只允许操作一个key,避免多个key多次弹出人员选择窗体
        strTmp = JSCA_Client.SOF_GetUserList() '返回数据格式：(用户1||标识1&&&用户2||标识2&&&…)一个key返回两条相同数据
        If strTmp <> "" Then
            arrKey = Split(strTmp, "&&&")
            intCount = (UBound(arrKey) + 1) \ 2
            If intCount > 1 Then
                strkeys = ""
                For i = LBound(arrKey) To UBound(arrKey)
                    If InStr(1, strkeys & ",", "," & arrKey(i) & ",") = 0 Then
                        strkeys = strkeys & "," & arrKey(i)    '记录下当前电脑上所有key标识
                    End If
                    If InStr(1, mstrKeys, "," & arrKey(i) & ",") = 0 And mstrKeys <> "" Then
                        mstrSelectKey = ""
                        blnChange = True  'key变动要重新选择
                    End If
                Next
                
                If strkeys <> "" And (mstrKeys = "" Or blnChange) Then
                    mstrKeys = strkeys & ","
                End If
            Else
                mstrKeys = ""
                mstrSelectKey = ""
            End If
        End If
        
        If intCount > 1 And mstrSelectKey <> "" And mstrSelectKey = strUserUnigueID Then
            strUniqueID = mstrSelectKey '多key的情况下,操作员不变且key盘未变动
        Else
            strUniqueID = JSCA_Client.SOF_SelectCert(3)   '证书ID 多key的情况下会 触发弹出选择窗体
        End If
        
        If intCount > 1 Then
            mstrSelectKey = strUniqueID    '记录下首次选择
        End If
        
        If strUniqueID = "" Then
            MsgBoxEx "请您插入Key！", vbInformation + vbOKOnly, "电子签名部件"
            Exit Function
        End If
        strCert = JSCA_Client.SOF_ExportUserCert(strUniqueID)   '证书内容
        strName = JSCA_Client.SOF_GetCertInfo(strCert, GET_USER_NAME)   '获取证书持有者姓名
    Else
        MsgBoxEx "江苏CA部件初始化失败。", vbInformation, "电子签名部件"
        Exit Function
    End If
    JSCA_GetCertList = True
    Exit Function
errH:
    MsgBoxEx "读取证书信息失败！" & vbNewLine & "第" & CStr(Erl()) & "行 " & Err.Description, vbQuestion, "电子签名部件"
    
End Function


Public Function JSCA_Sign(ByVal strCurrCertSn As String, ByVal strSource As String, ByRef strSignData As String, ByRef strTimeStamp As String, ByRef strTimeStampCode As String) As Boolean
        '签名
'功能:江苏CA电子签名
'参数：strCurrCertSn  -证书ID(唯一序列)
'     strSource-需要签名的源数据
'
'返回值：true 成功,False -失败
'       strSignData-签名后返回的签名数据
'       strTimeStamp-返回的时间戳

        Dim strRequest As String    '时间戳请求
        
        On Error GoTo errH
        
100     If JSCA_CheckCert(strCurrCertSn) Then               '验证当前USB是否是签名用户的，并获取签名证书
            strSource = JSCA_Client.SOF_EncodeBase64(strSource)   'base64编码,防止字符串带有空串或换行符验证失效
110         strSignData = JSCA_Client.SOF_SignData(strCurrCertSn, strSource)
            If strSignData <> "" Then
                JSCA_Sign = True
'                strRequest = mobjSoapClient.CreateTimeStampRequest(strSource)  '产生时间戳请求
'                If strRequest <> "" Then
'112                 strTimeStampCode = mobjSoapClient.CreateTimeStampResponse(strRequest)  '获取时间戳响应（base64编码格式）
'                    strTimeStamp = mobjSoapClient.GetTimeStampInfo(strTimeStampCode, 1)    '解析时间戳type =1：返回时间；type ＝2：返回签名值；type ＝3：返回签名证书
'                    strTimeStamp = GetTimeStamp(strTimeStamp)
'                    JSCA_Sign = True
'                Else
'                    MsgboxEx "时间戳请求失败！", vbExclamation, "电子签名部件"
'                End If
            Else
                MsgBoxEx "签名失败！", vbExclamation, "电子签名部件"
            End If
        Else
            MsgBoxEx "签名失败！", vbExclamation, "电子签名部件"
        End If
        Exit Function
errH:
114     MsgBoxEx "签名失败！" & vbNewLine & "第" & CStr(Erl()) & "行 " & Err.Description, vbQuestion, "电子签名部件"
End Function

Public Function JSCA_CheckCert(ByVal strCurrCertSn As String) As Boolean
'功能：读取USB进行设备初始化并登录
'返回值:
'  strSigCert -签名证书内容

        Dim strKey As String, strUserName As String, strSigCert As String
        Dim strWebUrl As String
        
        On Error GoTo errH
100     If Not mblnInit Then
102         MsgBoxEx "部件未初始化！"
            Exit Function
        End If
        
104     If JSCA_GetCertList(strUserName, strKey, strSigCert, strCurrCertSn) Then
106        If strCurrCertSn <> strKey Then
108            MsgBoxEx "该证书未注册在您的名下，不能使用！", vbInformation + vbOKOnly, "电子签名部件"
               Exit Function
           End If
110
116        If Not GetCertLogin(strKey, strSigCert) Then
                Exit Function
           Else
               JSCA_CheckCert = True
           End If

122     End If
    
        Exit Function
errH:
124     MsgBoxEx "检查USBKEY失败！" & vbNewLine & "第" & CStr(Erl()) & "行 " & Err.Description, vbQuestion, "电子签名部件"
End Function

Private Function GetCertLogin(ByVal strUniqueID As String, ByVal strCert As String) As Boolean
    '江苏CA
    '- 入参
    'strUniqueID : 证书唯一标识
    'strCert-证书内容 64BASE编码

    Dim datEnd  As Date
    Dim intDay  As Integer
    Dim lngTimes As Long
    
    '验证密码 首次调用江苏CA签名接口是自动弹出密码录入窗体，只录入一次。

'        lngTimes = JSCA_Client.SOF_Login(strUniqueID, strPassword)   '校验成功返回 0,失败返回剩余次数,-1 锁死
'        If lngTimes < -1 Then
'            MsgboxEx "密码验证失败！请重新注册CA部件【CACltCore.dll】。", vbInformation + vbOKOnly, "电子签名部件"
'            Exit Function
'        End If

    '获取客户端证书有效期截止时间
    datEnd = JSCA_Client.SOF_GetCertInfo(strCert, GET_CERT_AFTER)
    '验证客户端证书有效期剩余天数
     intDay = Int(CDbl(datEnd) - CDbl(Now))
    
    If (intDay <= 30 And intDay > 0 And Not gblnShow) Then
        MsgBoxEx "您的证书还有" & intDay & "天过期。", vbInformation + vbOKOnly, "电子签名部件"
        gblnShow = True
        GetCertLogin = True
    ElseIf (intDay <= 0) Then
        MsgBoxEx "您的证书已过期 " & Abs(intDay) & " 天。", vbInformation + vbOKOnly, "电子签名部件"
        GetCertLogin = False
    End If
    GetCertLogin = True
End Function

Public Function JSCA_VerifySign(ByVal strSigCert As String, ByVal strSignData As String, ByVal strSource As String, ByVal strTimeStamp As String, ByVal strTimeStampCode As String) As Boolean
'功能;验证签名
'参数: strSigCert -证书内容
'     strSignData-签名值
'     strSource-待验证源文
'     strTimeStampCode -时间戳BASE64编码
        Dim strTmp As String
        
        On Error GoTo errH
100
        strSource = JSCA_Client.SOF_EncodeBase64(strSource) 'base64编码,防止字符串带有空串或换行符验证失效
'        strTmp = mobjSoapClient.VerifyTimeStamp(strSource, strTimeStampCode)
'        If strTmp <> "0" Then
'            MsgboxEx "时间戳验证失败！", vbExclamation, "电子签名部件"
'            Exit Function
'        End If
        strTmp = mobjSoapClient.VerifySignedData(strSigCert, strSource, strSignData)
        If strTmp = "0" Then
            MsgBoxEx "验证成功，该电子签名数据有效!", vbInformation, "电子签名部件"
        Else
            MsgBoxEx "验证签名失败！", vbExclamation, "电子签名部件"
            Exit Function
        End If
       
        JSCA_VerifySign = True
        Exit Function
errH:
104     MsgBoxEx "验证签名失败！" & vbNewLine & "第" & CStr(Erl()) & "行 " & Err.Description, vbQuestion, "电子签名部件"
End Function


Private Function GetTimeStamp(ByVal strTimeStamp As String) As String
'功能：获取时间戳中的时间
    Dim arrTime As Variant
    Dim i As Long
    Dim strTime As String
    
    strTimeStamp = Replace(strTimeStamp, " ", "|")
    strTimeStamp = Replace(strTimeStamp, "||", "|") '当日期为一位数时,防止月份和日期之间存在两个空格的情况
    arrTime = Split(strTimeStamp, "|")  '传人格式：Aug 19 13:07:25 2014 GMT
    strTime = arrTime(0) & " " & arrTime(1) & " " & arrTime(3)  '月/日/年
    strTime = CDate(strTime) & ""  '年 月 日  2014/8/19
    GetTimeStamp = strTime & " " & arrTime(2)  ' 年-月-日 时:分:秒
End Function

Public Function JSCA_GetPara() As Boolean
'设置湖北CA服务器地址
    Dim arrList As Variant
    
    On Error GoTo errH
    gstrPara = gobjComLib.zlDatabase.GetPara(90000, glngSys)
    If gstrPara = "" Then gstrPara = "http://202.102.85.153:8080/HealthWebService.asmx?WSDL"
    If gstrPara <> "" Then
        gudtPara.strSignURL = gstrPara
    End If
    Exit Function
errH:
    MsgBoxEx "读取参数失败！" & vbNewLine & "第" & CStr(Erl()) & "行 " & Err.Description, vbInformation, gstrSysName
End Function

Public Function JSCA_SetParaStr() As String
    JSCA_SetParaStr = gudtPara.strSignURL
End Function


Public Sub JSCA_UnloadObj()
    Set JSCA_Client = Nothing
    Set JSCA_Seal = Nothing      '电子签章部件
    Set mobjSoapClient = Nothing        'soap连接对象
    mblnInit = False
End Sub
