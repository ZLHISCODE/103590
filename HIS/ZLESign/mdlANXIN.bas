Attribute VB_Name = "mdlANXIN"

Option Explicit
'吉林国投安信CA         七台河人民医院
'KEY类型:飞天epass3000gm auto  参数ANXIN3KGM 龙脉GM3000   参数 ANXINLONGMAIGM
Private mobjJLClient As Object   'JITComVCTKEx.JITVCTKEx.1
Private mobjJLServer As Object  'JITClientCOMAPI.JITClientProc.1
Private mobjCertInfo As Object    '签章 JITCertActiveX.CertInfo.1

'Private mobjJLClient As New JITComVCTKExLib.JITVCTKEx
'Private mobjJLServer As New JITClientCOMAPILib.JITClientProc
'Private mobjCertInfo As New JITCertActiveXLib.CertInfo    '签章
Private mblnInit As Boolean

Private mstrPWD As String          '缓存输入的密码
Private mstrKey As String

Private Const M_STR_PARA As String = "<?xml version=""1.0"" encoding=""gb2312""?><authinfo><liblist>" & _
        "<lib type=""CSP"" version=""1.0"" dllname="""" ><algid val=""SHA1"" sm2_hashalg=""sm3""/></lib>" & _
        "<lib type=""SKF"" version=""1.1"" dllname=""SERfR01DQUlTLmRsbA=="" ><algid val=""SHA1"" sm2_hashalg=""sm3""/></lib>" & _
        "<lib type=""SKF"" version=""1.1"" dllname=""U2h1dHRsZUNzcDExXzMwMDBHTS5kbGw="" ><algid val=""SHA1"" sm2_hashalg=""sm3""/></lib>" & _
        "<lib type=""SKF"" version=""1.1"" dllname=""QU5YSU5Dc3AxMV8zMDAwR01BLmRsbA=="" ><algid val=""SHA1"" sm2_hashalg=""sm3""/></lib>" & _
        "</liblist><checkkeytimes><item times=""3"" ></item></checkkeytimes></authinfo>"

Private Enum E_KEY_TYPE
    K_飞天 = 0
    K_龙脉 = 1
End Enum

Public Function ANXIN_InitObj() As Boolean
     '证书部件初始化
    Dim lngRet As Long
    Dim strTSAIP As String
    Dim strPara As String
    Dim varTmp As Variant
    Dim i As Long
    
100     If glngSign > 1 Then ANXIN_InitObj = True: Exit Function
        On Error Resume Next
102     If mobjJLClient Is Nothing Then
104         Set mobjJLClient = CreateObject("JITComVCTKEx.JITVCTKEx.1")
106         If Err.Number <> 0 Then
108             MsgBoxEx "创建安信签名对象【JITComVCTK_S.dll】失败！请检查该控件是否安装并注册。", vbInformation, gstrSysName
                Exit Function
            End If
        End If
110     Err.Clear
112     If mobjJLServer Is Nothing Then
114         Set mobjJLServer = CreateObject("JITClientCOMAPI.JITClientProc.1")
116         If Err.Number <> 0 Then
118             MsgBoxEx "创建安信签名对象【JITClientCOMAPI.dll】失败！请检查该控件是否安装并注册。", vbInformation, gstrSysName
                Exit Function
            End If
        End If
120     Err.Clear
122     If mobjCertInfo Is Nothing Then
124         Set mobjCertInfo = CreateObject("JITCertActiveX.SZZSCertInfo.1")
126         If Err.Number <> 0 Then
128             MsgBoxEx "创建安信签章对象【JITCertActiveX.dll】失败！请检查该控件是否安装并注册。", vbInformation, gstrSysName
                Exit Function
            End If
        End If
130     Err.Clear: On Error GoTo 0
        On Error GoTo errH
        '参数信息:是否启用时间戳[0-不启用;1-启用]&&&签名服务器IP&&&签名服务器端口&&&时间戳IP&&&时间戳端口&&&可选参数(dllname1&dllname2)
        '第一位 签名服务器，第二位时间戳服务器，第三位网关。安信就这3个硬件。如果没有硬件就是“000”，只有签名服务器就是“100”
        '为了兼容以前如果第一个参数=0;则代表只有签名服务器就是“100” ;=1代表启用签名服务器启用时间戳服务器 就是"110"第三位网关暂时未启用，预留参数
        '连接时间戳服务器
        'gstrPara = "1&&&175.17.252.155&&&8000&&&175.17.252.156&&&8000" '七台河人民医院
        'gstrPara = "000&&&175.17.252.155&&&8000&&&175.17.252.156&&&8000" '七台河妇幼医院
        
132     gstrPara = gobjComLib.zlDatabase.GetPara(90000, glngSys, , "") '读取配置内容
        LogWrite "ANXIN_InitObj", "CA参数:" & gstrPara
134     If gstrPara = "" Then
136         Err.Raise -1, , "当前系统【" & glngSys & "】没有配置电子签名参数,请到启用电子签名接口处设置。"
            Exit Function
        End If
138     If UBound(Split(gstrPara, G_STR_SPLIT)) <> 6 Then
140         MsgBoxEx "电子签名参数值设置有误,请检查。" & vbCrLf & _
                "当前参数值:" & gstrPara & vbCrLf & _
                "正确格式:是否启用时间戳[0-不启用;1-启用]&&&签名服务器IP&&&签名服务器端口&&&时间戳IP&&&时间戳端口&&&KEY类型(0-飞天;1-龙脉)&&&可选参数(dllname1&dllname2)", vbInformation, gstrSysName
            Exit Function
        Else
142         Call ANXIN_GetPara
        End If
        
144     If gudtPara.intKeyType = K_飞天 Then
146         mstrKey = "ANXIN3KGM"
148     ElseIf gudtPara.intKeyType = K_龙脉 Then
150         mstrKey = "ANXINLONGMAIGM"
        End If
152     If gudtPara.strOption <> "" Then
154         varTmp = Split(gudtPara.strOption, "&")
156         strPara = "<?xml version=""1.0"" encoding=""gb2312""?><authinfo><liblist>" & _
                        "<lib type=""CSP"" version=""1.0"" dllname="""" ><algid val=""SHA1"" sm2_hashalg=""sm3""/></lib>"
158         For i = LBound(varTmp) To UBound(varTmp)
160             strPara = strPara & "<lib type=""SKF"" version=""1.1"" dllname=""" & varTmp(i) & """ ><algid val=""SHA1"" sm2_hashalg=""sm3""/></lib>"
            Next
162         strPara = strPara & "</liblist><checkkeytimes><item times=""3"" ></item></checkkeytimes></authinfo>"
        Else
164         strPara = M_STR_PARA
        End If
166     lngRet = mobjJLClient.Initialize(strPara)
168     If Not GetErrorInfo("Initialize") Then Exit Function
170     mblnInit = True
174     mstrPWD = ""
176     ANXIN_InitObj = True
        
        Exit Function
errH:
178  MsgBoxEx "创建接口部件失败！" & vbNewLine & "第" & CStr(Erl()) & "行 " & Err.Description, vbExclamation, gstrSysName
End Function

Public Function ANXIN_Sign(ByVal strSource As String, ByRef strSignData As String, ByRef strTimeStamp As String, ByRef strTimeStampCode As String, Optional ByRef blnReDo As Boolean) As Boolean
'签名
    Dim lngRet As Long
    Dim strErr As String
    Dim strHash As String
    
    Dim blnCheck As Boolean
        On Error GoTo errH
        blnCheck = ANXIN_CheckCert(blnReDo)
        If blnReDo Then Exit Function
100     If blnCheck Then                 '验证当前USB是否是签名用户的，并获取签名证书
            '证书ID进行签名
            lngRet = mobjJLClient.SetPinCode(mstrPWD)
110         strSignData = mobjJLClient.DetachSignStr("", strSource)      'DetachSignStr-不带原文签名;AttachSignStr-带原文签名
            If Not GetErrorInfo("DetachSignStr") Then Exit Function
            If strSignData <> "" Then
                If gudtPara.blnISTS Then
                    If Not ConnectToTsaServer() Then Exit Function
                    strHash = StringSHA1(strSource)
                    strTimeStampCode = mobjJLServer.TsaSign("", 1, strHash)            '申请时间戳 传入签名值过长，签名时比较耗时,故采用固定值
                    strTimeStamp = mobjJLServer.VerifyTsaSign(strTimeStampCode)
                    Call mobjJLServer.FinalizeServerConnectEx    '断开时间戳服务器连接
                    If strTimeStampCode = "" Then MsgBoxEx "获取时间戳失败！", vbInformation, gstrSysName: Exit Function
                    '日期格式化
                    strTimeStamp = Mid(strTimeStamp, 1, 14)
                    strTimeStamp = String14ToDate(strTimeStamp, strErr)
                    If strErr <> "" Then MsgBoxEx strErr, vbInformation, gstrSysName: Exit Function
                    '转东八区时间
                    strTimeStamp = Format(DateAdd("h", 8, strTimeStamp), "YYYY-MM-DD HH:MM:SS")
                Else
                    strTimeStamp = Format(gobjComLib.zlDatabase.Currentdate & "", "yyyy-MM-dd HH:mm:ss")
                End If
            Else
                MsgBoxEx "签名失败！", vbInformation, gstrSysName
                Exit Function
            End If
112
        Else
            MsgBoxEx "签名失败！", vbInformation, gstrSysName
            Exit Function
        End If
        ANXIN_Sign = True
        Exit Function
errH:
114     MsgBoxEx "签名失败！" & vbNewLine & "第" & CStr(Erl()) & "行 " & Err.Description, vbExclamation, gstrSysName
End Function

Public Function ANXIN_VerifySign(ByVal strSource As String, ByVal strSignData As String, ByVal strTimeStampCode As String) As Boolean
    '功能;验证签名
    '参数:strSignData -签名值
     Dim blnRet As Boolean
     Dim lngRet As Long
     Dim strTS As String
     On Error GoTo errH
     LogWrite "ANXIN_VerifySign", "验证签名原文:" & strSource & vbCrLf & "验证签名值:" & strSignData & vbCrLf & "签名时间戳信息:" & strTimeStampCode
100 If gudtPara.blnIsSign Then
        '服务器验签
102     If Not ConnectToSignServer() Then Exit Function
104     lngRet = mobjJLServer.VerifyDetachedSign(strSignData, strSource) '服务器验证数据 不带原文签名:VerifyDetachedSign(string, string);带原文签名  VerifyAttachedSign
106     If lngRet <> 0 Then
108         MsgBoxEx "签名验证失败:" & mobjJLServer.GetErrorMessage(lngRet), vbInformation, gstrSysName
110         Call mobjJLServer.FinalizeServerConnectEx   '断开签名服务器链接
            Exit Function
        End If
112     Call mobjJLServer.FinalizeServerConnectEx   '断开签名服务器链接
    Else
114     Call mobjJLClient.VerifyDetachedSignStr(strSignData, strSource)  '客户端验证签名
116     lngRet = mobjJLClient.GetErrorCode()
118     If lngRet <> 0 Then
120         MsgBoxEx "验证签名失败，错误码：" & lngRet & " 错误信息：" & mobjJLClient.GetErrorMessage(lngRet), vbInformation, gstrSysName
            Exit Function
        End If
    End If

    '连接时间戳服务器
122 If gudtPara.blnISTS Then
124     If Not ConnectToTsaServer() Then Exit Function
126     strTS = mobjJLServer.VerifyTsaSign(strTimeStampCode)
128     If strTS = "" Then
130           MsgBoxEx "时间戳验证失败！", vbInformation, gstrSysName
132           Call mobjJLServer.FinalizeServerConnectEx   '断开时间戳服务器
              Exit Function
        End If
134     Call mobjJLServer.FinalizeServerConnectEx   '断开时间戳服务器
    End If
136 MsgBoxEx "验证成功，该电子签名数据有效!", vbInformation, gstrSysName
    
138  ANXIN_VerifySign = True
     Exit Function
errH:
140     MsgBoxEx "验证签名失败！" & vbNewLine & "第" & CStr(Erl()) & "行 " & Err.Description, vbExclamation, gstrSysName

End Function


Public Function ANXIN_CheckCert(ByRef blnReDo As Boolean) As Boolean
    '功能：读取USB进行设备初始化并登录
    Dim strKeySN As String, strUserID As String, strUserName As String, strCertDN As String
    Dim strDate As String
    Dim arrDN As Variant
    Dim udtUser As USER_INFO
    Dim blnRet As Boolean
    Dim i As Long
    
    On Error GoTo errH
    If Not GetCertList(strKeySN, strUserName, strCertDN, , strUserID) Then Exit Function
    If mUserInfo.strUserID = "" Then
        MsgBoxEx "您的身份证号为空,请联系管理员到人员管理中录入！", vbOKOnly + vbInformation, gstrSysName
        Exit Function
     ElseIf mUserInfo.strUserID <> strUserID Then
        MsgBoxEx "该证书未注册在您的名下，不能使用！"
        Exit Function
    End If
    
    '判断是否需要更新注册证书
    udtUser.strName = strUserName
    udtUser.strSignName = strUserName
    udtUser.strUserID = strUserID '身份证号
    udtUser.strCertSn = strKeySN
    udtUser.strCertDN = strCertDN
    udtUser.strCert = ""
    udtUser.strEncCert = ""
    udtUser.strCertID = ""
    udtUser.strPicPath = ""
    arrDN = Split(mUserInfo.strCertDN, ",")     'CN=王二小U3294, O=七台河人民医院, L=七台河市, S=黑龙江省, C=CN, 有效日期=
    For i = 0 To UBound(arrDN)
        If Trim(arrDN(i)) Like "有效日期*" Then
            strDate = Trim(Split(arrDN(i), "=")(1))
            Exit For
        End If
    Next
    If IsUpdateRegCert(udtUser, strDate, blnReDo) Then
        blnRet = True
    Else
        blnRet = False
    End If

    ANXIN_CheckCert = blnRet
    Exit Function
errH:
     MsgBoxEx "检查USBKEY失败！" & vbNewLine & "第" & CStr(Erl()) & "行 " & Err.Description, vbExclamation, gstrSysName
End Function


Public Function ANXIN_RegCert(arrCertInfo As Variant) As Boolean
    '功能：提供在HIS数据库中注册个人证书的必要信息,用于新发放或更换证书,,需要插入USB-Key
    '返回：arrCertInfo作为数组返回证书相关信息
'      0-ClientSignCertCN:客户端签名证书公共名称(姓名),用于注册证书时程序验证身份
'      1-ClientSignCertDN:客户端签名证书主题(每人唯一)
'      2-ClientSignCertSN:客户端签名证书序列号(每证书唯一)
'      3-ClientSignCert:客户端签名证书内容
'      4-ClientEncCert:客户端加密证书内容
'      5-签名图片文件名,空串表示没有签名图片

        Dim strKeyId As String, strCertUserName As String, strCertDN As String, strPicPath As String
        Dim i As Integer
        On Error GoTo errH

        For i = LBound(arrCertInfo) To UBound(arrCertInfo)
             arrCertInfo(i) = ""
        Next

        If GetCertList(strKeyId, strCertUserName, strCertDN, strPicPath) Then
            arrCertInfo(0) = strCertUserName
            arrCertInfo(1) = strCertDN
            arrCertInfo(2) = strKeyId
            arrCertInfo(5) = strPicPath
            ANXIN_RegCert = True
        End If

        Exit Function
errH:
     MsgBoxEx "获取证书信息失败！" & vbNewLine & "第" & CStr(Erl()) & "行 " & Err.Description, vbExclamation, gstrSysName

End Function

Private Function GetCertList(Optional ByRef strUniqueID As String = "-1", Optional ByRef strName As String = "-1", Optional ByRef strCertDN As String = "-1", _
    Optional ByRef strPicPath As String = "-1", Optional ByRef strUserID As String = "-1", Optional ByRef strDate As String = "-1") As Boolean
    '功能:获取安信证书详情
    'strUserID-身份证号
        Dim lngRet As Long
        Dim strRet As String
        Dim datCurrent As Date
        Dim arrDN As Variant
        Dim i As Long
        Dim strKeyCount As String
        Dim strPic As String, strPIN As String
        Dim strTmp As String
        Dim lngDay As Long
        Dim colTmp As Collection
        
        On Error GoTo errH
        
100     If Not mblnInit Then
102         lngRet = mobjJLClient.Initialize(M_STR_PARA)
104         If Not GetErrorInfo("Initialize") Then Exit Function
106         mblnInit = True
        End If
108     lngRet = mobjJLClient.SetCertChooseType(1)
110     lngRet = mobjJLClient.SetCert("SC", "", "", "", "CN = AnXin SM2 CA,O = AnXin CA,C = CN", "")
112     If Not GetErrorInfo("SetCert") Then Exit Function
114     strDate = mobjJLClient.GetCertInfo("SC", 6, "")     '有效日期
116     If IsDate(strDate) Then
            '检查证书是否过期
118         lngDay = CheckValidaty(strDate)
120         If (lngDay <= 30 And lngDay > 0 And Not gblnShow) Then
122             MsgBoxEx "您的证书还有" & lngDay & "天过期", vbInformation, gstrSysName
124             gblnShow = True
126         ElseIf (lngDay <= 0) Then
128             MsgBoxEx "您的证书已过期 " & Abs(lngDay) & " 天"
                Exit Function
            End If
        End If
130     If strUniqueID <> "-1" Then strUniqueID = mobjJLClient.GetCertInfo("SC", 2, "")     '证书序列号
132     If strCertDN <> "-1" Or strName <> "-1" Then
134         strCertDN = mobjJLClient.GetCertInfo("SC", 0, "") 'CN=王二小U3294, O=七台河人民医院, L=七台河市, S=黑龙江省, C=CN, 有效日期=
136         If strCertDN <> "" Then
138             arrDN = Split(strCertDN, ",")
140             For i = 0 To UBound(arrDN)
142                 If Trim(arrDN(i)) Like "CN*" Then
144                     strName = Trim(Split(arrDN(i), "=")(1))
                        Exit For
                    End If
                Next
            End If
146         strCertDN = strCertDN & ", 有效日期=" & strDate
        End If
    
148     If strUserID <> "-1" Then
150         strUserID = ""
152         strTmp = mobjJLClient.GetCertInfo("SC", 7, "1.2.86.11.7.1")  '身份证号需要转ASCII :31 16 a0 14 13 12 34 33 32 35 30 33 31 39 38 36 30 31 31 32 36 32 31 35
154         If Not GetErrorInfo("GetCertInfo") Then Exit Function
156         If strTmp <> "" Then
158             arrDN = Split(strTmp, " ")
160             For i = 6 To UBound(arrDN)    '前6个字符为前缀
162                 strUserID = strUserID & Chr(Val("&H" & arrDN(i)))
                Next
            End If
        End If
    
164     If mstrPWD = "" Then
CheckPWD:
166         If Not frmPassword.ShowMe(mstrPWD, 6, 16) Then Exit Function
            'strRet = mobjCertInfo.VerifyUserPin(mstrKey, mstrPWD)
            'VB调试的时候单步跟踪返回乱码；直接运行返回正确字符串'{"RetryCount":"0","VerifyValue":"1"}
            '首次验证密码,通过签名接口来处理
168         lngRet = mobjJLClient.SetPinCode(mstrPWD)
170         strRet = mobjJLClient.DetachSignStr("", "123")
172         If Not GetErrorInfo("DetachSignStr") Then
174             mstrPWD = ""
                Exit Function
            End If
        End If

        '获得Key数量
188     If strPicPath <> "-1" Then
            '获取签章耗时,签名时不读取，只在注册的时候获取
            'strKeyCount = [{"KeyName":"安信电子钥匙 ","KeyType":"ANXIN3KGM","UsbKeySerialNumber":"AX00010415"},{"KeyName":"安信电子钥匙","KeyType":"ANXIN3KGM","UsbKeySerialNumber":"AX00010414"}]
190         strKeyCount = mobjCertInfo.GetKeyCount(mstrKey)
192         strRet = strKeyCount 'VB调试的时候单步跟踪返回乱码；直接运行返回正确字符串
194         If strRet <> "" Then
196             If UBound(Split(strRet, "},{")) = 0 Then
198                 strPic = mobjCertInfo.ReadImageData(mstrKey, mstrPWD)
200                 If Len(strPic) > 1 Then
202                     strPicPath = SaveBase64ToFile("gif", strUniqueID, strPic)
                    Else
204                     strPicPath = ""
                    End If
206             ElseIf Val(strKeyCount) > 0 Then
208                 MsgBoxEx "请选择唯一的KEY盘插入！", vbInformation, gstrSysName
                    Exit Function
                End If
            End If
        End If

210     GetCertList = True
        Exit Function
errH:
212     MsgBoxEx "获取证书信息失败！" & vbNewLine & "第" & CStr(Erl()) & "行 " & Err.Description, vbExclamation, gstrSysName
End Function

Public Function ANXIN_GetSeal() As String
'获取签章图片
    Dim strPicPath As String
    Call GetCertList(, , , strPicPath)
    ANXIN_GetSeal = strPicPath
End Function

Private Function GetErrorInfo(ByVal strName As String) As Boolean
        Dim lngRet As Long

        On Error GoTo errH
100     lngRet = mobjJLClient.GetErrorCode  'lngRet -536870826 密码不对;-536870823  指定的密码太长或太短
102     If lngRet <> 0 Then
104         MsgBoxEx "调用接口：【" & strName & "】后出错,错误描述:" & vbCrLf & mobjJLClient.GetErrorMessage(lngRet), vbInformation, gstrSysName
            Exit Function
        End If
106     GetErrorInfo = True
        Exit Function
errH:
108     MsgBoxEx "获取错误描述！" & vbNewLine & "第" & CStr(Erl()) & "行 " & Err.Description, vbExclamation, gstrSysName
End Function

Private Function ConnectToTsaServer() As Boolean
        Dim lngRet As Long

        On Error GoTo errH

100     lngRet = mobjJLServer.InitServerConnectEx(gudtPara.strTSIP, CInt(gudtPara.strTSPort))
102     If lngRet <> 0 Then
104         MsgBoxEx mobjJLServer.GetErrorMessage(lngRet), vbInformation, gstrSysName
            Exit Function
        End If
106     lngRet = mobjJLServer.SetServerUriEx("/signserver/service/xml")
108     If lngRet <> 0 Then
110         MsgBoxEx mobjJLServer.GetErrorMessage(lngRet), vbInformation, gstrSysName
112         Call mobjJLServer.FinalizeServerConnectEx '终止服务器连接，以释放连接句柄
            Exit Function
        End If
114     ConnectToTsaServer = True
        Exit Function
errH:
116     MsgBoxEx "连接时间戳服务器！" & vbNewLine & "第" & CStr(Erl()) & "行 " & Err.Description, vbExclamation, gstrSysName

End Function

Private Function ConnectToSignServer() As Boolean
        Dim lngRet As Long

        On Error GoTo errH
100     lngRet = mobjJLServer.InitServerConnectEx(gudtPara.strSIGNIP, CInt(gudtPara.strSignPort))
102     If lngRet <> 0 Then
104         MsgBoxEx mobjJLServer.GetErrorMessage(lngRet), vbInformation, gstrSysName  '连接服务器失败
            Exit Function
        End If
106     lngRet = mobjJLServer.SetServerUriEx("/signserver/service/xml")
108     If lngRet <> 0 Then
110         MsgBoxEx mobjJLServer.GetErrorMessage(lngRet), vbInformation, gstrSysName
112         Call mobjJLServer.FinalizeServerConnectEx '终止服务器连接，以释放连接句柄
            Exit Function
        End If
114     lngRet = mobjJLServer.SetCertAliasEx("")  '设置服务器签名时的签名证书标识,空为默认证书
116     If lngRet <> 0 Then
118         MsgBoxEx mobjJLServer.GetErrorMessage(lngRet), vbInformation, gstrSysName
120         Call mobjJLServer.FinalizeServerConnectEx '终止服务器连接，以释放连接句柄
            Exit Function
        End If
122     ConnectToSignServer = True
        Exit Function
errH:
124     MsgBoxEx "连接签名戳服务器！" & vbNewLine & "第" & CStr(Erl()) & "行 " & Err.Description, vbExclamation, gstrSysName
End Function


Public Function ANXIN_GetPara() As Boolean
        Dim arrList As Variant
    
        On Error GoTo errH
100     If gstrPara = "" Then gstrPara = gobjComLib.zlDatabase.GetPara(90000, glngSys, , "") '读取URLs 固定读取ZLHIS 系统默认100
        '格式是否启用设备[000-都不启用;111-都启用]&&&签名服务器IP&&&签名服务器端口&&&时间戳IP&&&时间戳端口&&&KEY类型(0-飞天;1-龙脉)&&&部件名称
102     If gstrPara = "" Then gstrPara = "110&&&175.17.252.155&&&8000&&&175.17.252.156&&&8000&&&0&&&" & _
            "SERfR01DQUlTLmRsbA==&U2h1dHRsZUNzcDExXzMwMDBHTS5kbGw=&U2h1dHRsZUNzcDExXzMwMDBHTS5kbGw="
104     arrList = Split(gstrPara, "&&&")
106     If UBound(arrList) >= 6 Then
108         If Len(arrList(0)) = 3 Then
110             gudtPara.blnIsSign = Mid(arrList(0), 1, 1) = "1"
112             gudtPara.blnISTS = Mid(arrList(0), 2, 1) = "1"
            Else
114             gudtPara.blnISTS = Val(arrList(0)) = 1
116             gudtPara.blnIsSign = True
            End If
118         gudtPara.strSIGNIP = arrList(1)
120         gudtPara.strSignPort = arrList(2)
122         gudtPara.strTSIP = arrList(3)
124         gudtPara.strTSPort = arrList(4)
        
126         gudtPara.intKeyType = arrList(5)
128         gudtPara.strOption = arrList(6)
        Else
130         gudtPara.blnISTS = True
132         gudtPara.blnIsSign = True
134         gudtPara.strSIGNIP = "175.17.252.155"
136         gudtPara.strSignPort = "8000"
138         gudtPara.strTSIP = "175.17.252.156"
140         gudtPara.strTSPort = "8000"
142         gudtPara.intKeyType = K_飞天    '默认飞天
144         gudtPara.strOption = "SERfR01DQUlTLmRsbA==&U2h1dHRsZUNzcDExXzMwMDBHTS5kbGw=&U2h1dHRsZUNzcDExXzMwMDBHTS5kbGw="
        End If
        Exit Function
errH:
146     MsgBoxEx "读取参数失败！" & vbNewLine & "第" & CStr(Erl()) & "行 " & Err.Description, vbInformation, gstrSysName
End Function

Public Function ANXIN_SetParaStr() As String
    With gudtPara
        ANXIN_SetParaStr = IIf(.blnIsSign, "1", "0") & IIf(.blnISTS, "1", "0") & "0" & G_STR_SPLIT & _
            IIf(Trim(.strSIGNIP) = "", "175.17.252.155", .strSIGNIP) & G_STR_SPLIT & IIf(Trim(.strSignPort) = "", "8000", .strSignPort) & _
            G_STR_SPLIT & IIf(Trim(.strTSIP) = "", "175.17.252.156", .strTSIP) & G_STR_SPLIT & IIf(Trim(.strTSPort) = "", "8000", .strTSPort) & _
            G_STR_SPLIT & .intKeyType & _
            G_STR_SPLIT & IIf(Trim(.strOption) = "", "SERfR01DQUlTLmRsbA==&U2h1dHRsZUNzcDExXzMwMDBHTS5kbGw=&U2h1dHRsZUNzcDExXzMwMDBHTS5kbGw=", .strOption)

    End With
End Function

Public Sub ANXIN_UnLoadObj()
    On Error Resume Next
    Set mobjJLServer = Nothing
    Set mobjCertInfo = Nothing
    Call mobjJLClient.Finalize
    Set mobjJLClient = Nothing
End Sub




