Attribute VB_Name = "mdlLNCASY"
Option Explicit

Private mobjUSBKEY As Object     '辽宁省数字签名 华大 KEYUSBKEYACTIVE.USBKeyActiveCtrl.1
Private mobjMSScriptCtl As Object    'MSScriptControl.ScriptControl.1 微软提供脚本控件 用到javaScript中encodeURI方法获取URL串
Private mblnInit As Boolean
Private mstrLastPwd As String          '缓存输入的密码
Private mintLogin As Integer
Public gobjLNCAPenSign As Object '辽宁CA手写签名对象
'20170817 SM2算法集成  辽宁嘉鸿
Private mbytModel           As Byte             '0-RSA算法;1-SM2算法
Private mobjKeyManager      As Object           '证书管理对象
Private mobjCert            As Object           '证书对象
Private mobjKeyStore        As Object           'UKey操作类KeyStore
Private mobjKeySealArray    As Object
Private mobjKeySeal         As Object           '签章类
Private mobjKeyGateOper     As Object
Private mobjKeyDetector     As Object           'JHKey.KeyDetector.1.1
Private Enum E_Model
    E_RSA = 0
    E_SM2 = 1
End Enum

Public Function LNCA_Initialize() As Boolean
        '功能:创建辽宁CA控件对象
    
        Dim varTmp As Variant
    
        On Error GoTo errH
   
100         If mblnInit Then LNCA_Initialize = True: Exit Function
        
102         gstrPara = gobjComLib.zlDatabase.GetPara(90000, glngSys) '读取URL 服务器
            'gstrPara = "http://218.25.86.214:2010/ssoworker"  '测试地址
104         If gstrPara = "" Then
106             MsgBoxEx "没有配置认证服务器地址，请到【公共参数设置】先配置:" & vbCrLf & vbTab & "系统号100,参数号90000" & _
                        vbCrLf & vbTab & "参数值格式""http://218.25.86.214:2010/ssoworker""", vbInformation, gstrSysName
                Exit Function
            End If
108         varTmp = Split(gstrPara, G_STR_SPLIT)
110         gudtPara.strSignURL = varTmp(0)
112         If UBound(varTmp) >= 1 Then
114             mbytModel = Val(varTmp(1) & "")
            Else
116             mbytModel = E_RSA
            End If
        
118         If UBound(varTmp) >= 2 Then
120             gudtPara.strSIGNIP = varTmp(2)
            Else
122             gudtPara.strSIGNIP = ""
            End If
        
124         If mbytModel = E_RSA Then   'RSA
126             Set mobjUSBKEY = CreateObject("USBKEYACTIVE.USBKeyActiveCtrl.1") '签名对象
128             Set mobjMSScriptCtl = CreateObject("MSScriptControl.ScriptControl.1")
130             mobjMSScriptCtl.Language = "JavaScript"
            Else                    'SM2
132             Set mobjKeyManager = CreateObject("JHKey.KeyManager.1")
134             Set mobjCert = CreateObject("JHKey.Cert.1")
136             Set mobjKeyStore = CreateObject("JHKey.KeyStore.1")
138             Set mobjKeySealArray = CreateObject("JHKey.SealArray.1")
140             Set mobjKeySeal = CreateObject("JHKey.Seal.1")
142             Set mobjKeyGateOper = CreateObject("JHKey.GateOper.1")
144             Set mobjKeyDetector = CreateObject("JHKey.KeyDetector.1")
146             Call mobjKeyGateOper.SetTimeout(10)
148             Call mobjKeyGateOper.SetURL(gudtPara.strSignURL)
            
            End If
        
150         gstrLogins = ""
152         mblnInit = True
154         LNCA_Initialize = True
            Exit Function
errH:
156         LogWrite "LNCA_Initialize", "创建接口部件失败！" & vbNewLine & "第" & CStr(Erl()) & "行 " & Err.Description
End Function

Public Function LNCA_RegCert(arrCertInfo As Variant, Optional ByVal strUserID As String) As Boolean
'功能：提供在HIS数据库中注册个人证书的必要信息,用于新发放或更换证书,,需要插入USB-Key
'参数:strUserID-身份证号
'返回：arrCertInfo作为数组返回证书相关信息
'      0-ClientSignCertCN:客户端签名证书公共名称(姓名),用于注册证书时程序验证身份
'      1-ClientSignCertDN:客户端签名证书主题(每人唯一)
'      2-ClientSignCertSN:客户端签名证书序列号(每证书唯一)
'      3-ClientSignCert:客户端签名证书内容
'      4-ClientEncCert:客户端加密证书内容
'      5-签名图片文件名,空串表示没有签名图片
        
        Dim strCertSn As String, strCertUserName As String, strCertDN As String
        Dim strSigCert As String, i As Integer
        Dim strPic As String
        Dim strCertUserID As String
10      On Error GoTo errH
    
20      For i = LBound(arrCertInfo) To UBound(arrCertInfo)
30          arrCertInfo(i) = ""
40      Next
    
50      If GetCertList(strCertUserName, strCertSn, strSigCert, strCertDN, strPic, strCertUserID) Then
60          If UCase(strCertUserID) <> UCase(strUserID) And strUserID <> "" Then
70              MsgBoxEx "用户身份证号：" & _
                        vbCrLf & vbTab & "【" & UCase(strUserID) & "】" & vbCrLf & _
                        "当前证书唯一标识:" & _
                        vbCrLf & vbTab & "【" & UCase(strCertUserID) & "】" & vbCrLf & _
                        "用户身份证号与当前证书唯一标识不相等,不能注册！", vbInformation, gstrSysName
80              Exit Function
90          End If
100         arrCertInfo(0) = strCertUserName
110         arrCertInfo(1) = strCertDN
120         arrCertInfo(2) = strCertSn
130         arrCertInfo(3) = strSigCert
140         If strPic <> "" Then
150             arrCertInfo(5) = SaveBase64ToFile("gif", strCertSn, strPic)
160         End If
170         LNCA_RegCert = True
180     End If

190     Exit Function
errH:
200     MsgBoxEx "获取证书信息失败！" & vbNewLine & "第" & CStr(Erl()) & "行 " & Err.Description, vbExclamation, gstrSysName

End Function

Public Function LNCA_CheckCert(ByRef blnReDo As Boolean) As Boolean
'功能：读取USB进行设备初始化并登录
    Dim strCertName As String
    Dim strCertSn As String
    Dim strCertUserID As String    '包含身份证号信息
    Dim strDate As String
    Dim udtUser As USER_INFO
    Dim strCertID As String
    Dim strCert As String
    Dim blnOk As Boolean
    
    On Error GoTo errH

     '获取证书信息同时检查Key盘是否插入
    If Not GetCertList(, strCertSn, , , , strCertUserID) Then Exit Function
        
    '未注册在当前用户名下的Key
    If mUserInfo.strUserID = "" Then
        MsgBoxEx "您的身份证号为空,请联系管理员到人员管理中录入！", vbOKOnly + vbInformation, gstrSysName
        Exit Function
    ElseIf UCase(strCertUserID) <> UCase(mUserInfo.strUserID) Then
        MsgBoxEx "您的身份证号：" & _
                   vbCrLf & vbTab & "【" & mUserInfo.strUserID & "】" & vbCrLf & _
                   "当前证书唯一标识:" & _
                   vbCrLf & vbTab & "【" & strCertUserID & "】" & vbCrLf & _
                   "用户身份证号与当前证书唯一标识不相等,不能使用！", vbInformation, gstrSysName
        Exit Function
    End If
    'CA首次签名时会自动弹出密码框
    '登录验证
    If InStr(gstrLogins & "|", "|" & strCertSn & "|") > 0 Then '首次验证通过后，下次不在继续验证
        blnOk = True
    Else
        If Not GetCertList(, , strCert, , , , strDate, 1) Then Exit Function
        If Not GetCertLogin(strCert, strDate) Then
            blnOk = False
        Else
            blnOk = True
            If InStr(gstrLogins & "|", "|" & strCertSn & "|") = 0 Then gstrLogins = gstrLogins & "|" & strCertSn
        End If
    End If

    If blnOk And mUserInfo.strCertSn <> strCertSn Then
        '判断是否需要更新注册证书
        '耗时操作放到注册时读取如签章图片的处理的处理
        If Not GetCertList(strCertName, , udtUser.strCert, udtUser.strCertDN, udtUser.strPicCode, , strDate, 1) Then Exit Function
        udtUser.strName = strCertName
        udtUser.strSignName = strCertName
        udtUser.strUserID = strCertUserID
        udtUser.strCertSn = strCertSn
        If IsUpdateRegCert(udtUser, strDate, blnReDo) Then
            blnOk = True
        Else
            blnOk = False
        End If
    End If
    LNCA_CheckCert = blnOk
    Exit Function
errH:
     MsgBoxEx "检查USBKEY失败！" & vbNewLine & "第" & CStr(Erl()) & "行 " & Err.Description, vbInformation, gstrSysName
End Function

Private Function GetCertList(Optional ByRef strName As String = "-1", Optional ByRef strCertSn As String = "-1", Optional ByRef strCert As String = "-1", _
                Optional ByRef strCertDN As String = "-1", Optional strPic As String = "-1", _
                Optional strUserID As String = "-1", Optional strEndDate As String = "-1", Optional ByVal bytMode As Byte) As Boolean
          '功能:获取证书信息
          '-出参
          '    strName 证书持有者姓名
          '   strCertSN 证书唯一标识
          '   strCert 签名证书
          '   strCertDN 证书描述信息  证书注册用到
          '   strPic      证书图片
          '   bytMode =0 初始化密码;1-不初始化
          Dim strMsg As String
          Dim lngRet As Long
          Dim strPIN As String
          Dim i As Integer
          Dim indexS As Integer
          
10        On Error GoTo errH
20        If Not LNCA_Initialize() Then Exit Function
          
30        If mbytModel = E_RSA Then
              '输入密码
40            If bytMode = 0 Then
50                If mstrLastPwd <> "" Then strPIN = mstrLastPwd
60                If strPIN = "" Then
70                    If Not frmPassword.ShowMe(strPIN) Then Exit Function
80                End If
                  
90                If strPIN = "" Then
100                  MsgBoxEx "请输入证书密码！", vbOKOnly + vbInformation, gstrSysName
110                  Exit Function
120               Else
130                   If mintLogin >= 8 Then
140                       MsgBoxEx "已经输入了" & mintLogin & "次错误密码，超过了最大输入次数！", vbOKOnly + vbInformation, gstrSysName
150                       Exit Function
160                   End If
170                   On Error Resume Next
180                   lngRet = mobjUSBKEY.MNGInit(strPIN)
190                   If Err.Number <> 0 Then
200                       MsgBoxEx "请您插入KEY盘！", vbOKOnly + vbInformation, gstrSysName
210                       Exit Function
220                   End If
230                   Err.Clear: On Error GoTo 0
                      
240                   If lngRet = 0 Then
250                      mstrLastPwd = strPIN
260                   Else
270                       mintLogin = mintLogin + 1
280                       MsgBoxEx "证书密码可能不正确，您已经输入了" & mintLogin & "次密码，还可以输入" & 8 - mintLogin & "次!", vbOKOnly + vbInformation, gstrSysName
290                       mstrLastPwd = ""
300                       Exit Function
310                   End If
320               End If
330           End If
340           On Error GoTo errH
              
350           If bytMode = 0 Then Call mobjUSBKEY.MNGLogin
                
360           If strCertSn <> "-1" Then strCertSn = mobjUSBKEY.MNGGetSignCertSN '唯一标识号
370           If strCert <> "-1" Then strCert = mobjUSBKEY.MNGGetSignCert()    '获取签名证书
380           If strName <> "-1" Then strName = mobjUSBKEY.MNGGetSignCertCN()          ''获取名称
390           If strCertDN <> "-1" Then strCertDN = mobjUSBKEY.MNGGetSignCertDN      '描述
              
400           If strPic <> "-1" Then
410               strPic = mobjUSBKEY.MNGGetSESCount '法人签名|个人章
420               If strPic = "" Then
430                   strMsg = "读取签名图片失败！"
440                   GoTo msgINFO
450               End If
460               strPic = Split(strPic, "|")(0)
470               strPic = mobjUSBKEY.MNGReadSESealByLabelEx(strPic)         '获取签章图片BASE64
480               If strPic = "" Then
490                   strMsg = "读取签名图片失败！"
500                   GoTo msgINFO
510               End If
520           End If
530           If strUserID <> "-1" Then
540               strUserID = mobjUSBKEY.MNGGetSignCertDN_OUa   '身份证号
550               If strUserID <> "" And Len(strUserID) >= 18 Then
560                   strUserID = Right(strUserID, 18)
570               Else
580                   strMsg = "读取身份证号码失败！"
590                   GoTo msgINFO
600               End If
610           End If
              '获取客户端证书有效期截止时间
620           If strEndDate <> "-1" Then
630               strEndDate = mobjUSBKEY.MNGGetSignCertEndValidityTime()
640               strEndDate = CDate(Format(strEndDate, "YYYY-MM-DD HH:MM:SS"))
650           End If
660       ElseIf mbytModel = E_SM2 Then
670           lngRet = mobjKeyDetector.EnumUKey()
680           If lngRet = 0 Then
690               MsgBoxEx "请您插入KEY盘！", vbOKOnly + vbInformation, gstrSysName
700               Exit Function
710           ElseIf lngRet = 1 Then
720               Call mobjKeyManager.EnumKeyStore               '调用EnumKeyStore枚举出所有证书
                  '默认选取SM2的key签名
730               lngRet = mobjKeyManager.GetCertCount()
740               For i = 0 To lngRet - 1 '调用GetCertCount得到全部证书的个数
750                   mobjCert.SetCert (mobjKeyManager.GetCert(i)) '轮询获得所有的证书
760                   If mobjCert.CertUsage = 1 Then 'CertUsage == 1:签名证书,2:加密证书。CertType == 1:RSA,2:SM2
770                       Call mobjKeyManager.InitKeyStoreByIndex(i, mobjKeyStore)
780                       indexS = i
790                       Exit For
800                   End If
810               Next
820           Else
830               Call mobjKeyManager.EnumKeyStore               '调用EnumKeyStore枚举出所有证书
                  '显示证书列表框，用于让用户选取证书。
                  '第一个参数，表示显示的证书类型，1：RSA，2：SM2，3：RSA和SM2都显示;
                  '第二个参数表示证书用途，1：签名，2：加密，3：签名加密均可
                  '0-用户选择了证书;1-未选择证书
840               lngRet = mobjKeyManager.ShowCertsDlg(3, 1)
850               If lngRet <> 0 Then
860                   strMsg = "未选择任何证书！"
870                   GoTo msgINFO
880               End If
890               Call mobjKeyManager.GetSelectedCert(mobjCert)    '根据用户的选择得到指定的证书
                  '根据用户选择的证书，初始化UKey操作类KeyStore 签名时直接调用签名接口
900               Call mobjKeyManager.InitKeyStore(mobjKeyStore)
                  'mobjCert.CertType证书类型，1:RSA，2:SM2
910               indexS = mobjKeyManager.getDlgSelId()
920           End If
930           strName = mobjCert.CertCN
940           If strCertSn <> "-1" Then strCertSn = mobjCert.CertSN
950           If strCertDN <> "-1" Then strCertDN = mobjCert.CertSubject
960           If strCert <> "-1" Then strCert = mobjCert.Body
970           If strUserID <> "-1" Then strUserID = Mid(mobjCert.CertOuA, 2)
980           If strEndDate <> "-1" Then
990               strEndDate = mobjCert.CertNotAfter
1000          End If
              
1010          If mobjCert.CertType = 2 Then  'SM2 允许缓存密码
                  '输入密码
1020              If mstrLastPwd <> "" Then strPIN = mstrLastPwd
1030              If strPIN = "" Then
1040                  If Not frmPassword.ShowMe(strPIN) Then Exit Function
1050              End If
                      
1060              If strPIN = "" Then
1070                 MsgBoxEx "请输入证书密码！", vbOKOnly + vbInformation, gstrSysName
1080                 Exit Function
1090              Else
1100                  Call mobjKeyStore.SetWorkPin(strPIN)  '添加默认PING码，使PIN码框不再弹出，达到静默操作的目的（只对SM2Key有效。RSA的PIN码框由各驱动厂家各自实现，无法控制）
1110                  lngRet = mobjKeyStore.SignData("123")    '0-成功;非0-失败
1120                  If lngRet = 0 Then
1130                      mstrLastPwd = strPIN
1140                  Else
1150                      mstrLastPwd = ""
1160                      Exit Function
1170                  End If
1180              End If
1190          End If
            
              '取印章：根据所选证书，查询到证书所在Key里的所有印章，存入印章数组SealArray
1200          If strPic <> "-1" Then
                  'Call mobjKeyManager.InitSealStore(mobjKeySealArray)
1210              Call mobjKeyManager.InitSealArrByIndex(indexS, mobjKeySealArray) '岫岩医院启用辽宁CA时调整
1220              lngRet = mobjKeySealArray.GetSealCount()         '得到印章数组所存印章个数
1230              If lngRet = 0 Then
1240                  strMsg = "所选证书无对应的印章！"
1250                  GoTo msgINFO
1260              End If
1270              For i = 0 To lngRet - 1
1280                  Call mobjKeySealArray.GetSeal(i, mobjKeySeal)     '从印章数组中取得印章
1290                  strPic = mobjKeySeal.getpic()              '得到印章图片的base64数据
1300                  If strPic <> "" Then
1310                      Exit For
1320                  Else
1330                      strMsg = "读取签名图片失败！"
1340                      GoTo msgINFO
1350                  End If
1360              Next
1370          End If
1380      End If
1390      GetCertList = True
1400      Exit Function

msgINFO:
1410    If strMsg <> "" Then
1420        MsgBoxEx strMsg, vbInformation, gstrSysName
1430        Exit Function
1440    End If
errH:

1450  MsgBoxEx "获取证书信息失败！" & vbNewLine & "第" & CStr(Erl()) & "行 " & Err.Description, vbInformation, gstrSysName
End Function

Private Function GetCertLogin(ByVal strSignCert As String, ByVal strEnd As String) As Boolean
    Dim strSignResult As String
    Dim strToken As String, strParameter As String, strRandom As String
    Dim strRet As String
    Dim datEnd As Date
    Dim intDay As Integer
    Dim lngRet As Long
    
    On Error GoTo errH
    If mbytModel = E_RSA Then
        '获取随机数
        strRandom = HttpPost(gudtPara.strSignURL, "cmd=getrand", responseText)   '获取随机数返回值: {"ret":1,"errinfo":"","rand":"3XO9JCXVJ6LF05M51165"}
        strRandom = GetSubString(strRandom, "rand")
        
        '随机数签名
        strSignResult = mobjUSBKEY.MNGSignData(strRandom, Len(strRandom))      '控件签名
        If strSignResult = "" Then
            MsgBoxEx "随机数签名失败！", vbInformation + vbOKOnly, gstrSysName
            Exit Function
        End If
    
        '随机数签名验证
        strParameter = "cmd=sm2certlogin" & "&rand=" & EnCodeURL(strRandom) & "&cert=" & EnCodeURL(strSignCert) & "&signed=" & EnCodeURL(strSignResult)  '服务器验证KEY 签名结果
        strToken = HttpPost(gudtPara.strSignURL, strParameter, responseText)  '
        strRet = GetSubString(strToken, "ret")
        If strRet <> "1" Then
            MsgBoxEx "证书登录验证失败！", vbInformation + vbOKOnly, gstrSysName
            Exit Function
        End If
    ElseIf mbytModel = E_SM2 Then
'        Call mobjKeyStore.SetWorkPin(mstrLastPWD)
'       重复设置密码导致登陆验证返回值=9 提示失败 故注释
        lngRet = mobjKeyGateOper.ReqCertLogin(mobjKeyStore, "123")  '证书登陆
        If lngRet <> 0 Then
            MsgBoxEx "证书登录验证失败！" & vbCrLf & "错误描述:" & mobjKeyGateOper.GetLastErrText & "返回值:" & lngRet, vbInformation + vbOKOnly, gstrSysName
            Exit Function
        End If
    End If
    
    '获取客户端证书有效期截止时间
    datEnd = CDate(strEnd)
    '验证客户端证书有效期剩余天数
    intDay = Int(CDbl(datEnd) - CDbl(Now))
    
    If (intDay <= 30 And intDay > 0 And Not gblnShow) Then
        MsgBoxEx "您的证书还有" & intDay & "天过期。", vbInformation + vbOKOnly, gstrSysName
        gblnShow = True
        GetCertLogin = True
    ElseIf (intDay <= 0) Then
        MsgBoxEx "您的证书已过期 " & Abs(intDay) & " 天。", vbInformation + vbOKOnly, gstrSysName
        GetCertLogin = False
    End If
        
    GetCertLogin = True
    Exit Function
errH:
    MsgBoxEx "登录服务器验证失败！" & vbNewLine & "第" & CStr(Erl()) & "行 " & Err.Description, vbInformation, gstrSysName
End Function

Public Function LNCA_Sign(ByVal strSource As String, ByRef strSignData As String, ByRef strTimeStamp As String, _
        ByRef strTimeStampCode As String, ByRef blnReDo As Boolean, ByVal blnCheck As Boolean) As Boolean
'签名
        Dim strParameter As String, strMsg As String, strDate As String, strRet As String
        Dim intRet As Integer
        Dim blnRet As Boolean
        Dim datTime As Date
        
10      On Error GoTo errH
        
20      If Not blnCheck Then
30          blnCheck = LNCA_CheckCert(blnReDo)
40          If blnReDo Then Exit Function
50      End If
60      If blnCheck Then
70          If mbytModel = E_RSA Then
                '验证当前USB是否是签名用户的，并获取签名证书
80              strSource = EncodeBase64String(strSource) '源文中包含特殊字符串需要加密转换
90              strSignData = mobjUSBKEY.MNGSignData(strSource, Len(strSource))        '控件签名
100             If strSignData <> "" Then
                    '存储网关服务器
110                 datTime = gobjComLib.zlDatabase.Currentdate()
120                 strDate = Format(datTime, "yyyyMMddhhmmss")
130                 strTimeStamp = Format(datTime, "yyyy-MM-dd HH:mm:ss")
140                 strParameter = "cmd=insert_sign_record" & "&appid=" & EnCodeURL("100") & "&docid=" & EnCodeURL("100") & "&docname=" & _
                    EnCodeURL("ZLHIS") & "&textinfo=" & EnCodeURL(strSource) & "&signdata=" & EnCodeURL(strSignData) & "&signcert=" & _
                    EnCodeURL(mUserInfo.strCert) & "&signdate=" & EnCodeURL(strDate)
150                 strRet = HttpPost(gudtPara.strSignURL, strParameter, responseText)
160                 blnRet = GetSubString(strRet, "ret") = "1"
170                 If Not blnRet Then strMsg = "签名失败！"
180             Else
190                 strMsg = "签名失败！"
200                 blnRet = False
210             End If
220         ElseIf mbytModel = E_SM2 Then
                'mobjKeyStore在LNCA_CheckCert 已经实例化
230             Call mobjKeyStore.SetWorkPin(mstrLastPwd)
240             intRet = mobjKeyStore.SignData(strSource)   '0-成功;非0-失败
250             If intRet = 0 Then
260                 strSignData = mobjKeyStore.GetSignData()               '得到签名数据
270             Else
280                 strMsg = "签名失败！"
290                 blnRet = False
300             End If
                '签名后存储网关服务器
310             datTime = gobjComLib.zlDatabase.Currentdate()
320             strDate = Format(datTime, "yyyyMMddhhmmss")
330             strTimeStamp = Format(datTime, "yyyy-MM-dd HH:mm:ss")
340             intRet = mobjKeyGateOper.ReqUploadMedRecord("01", "appid", "docid", "docname", strSource, strSignData, mUserInfo.strCert, strDate)
350             blnRet = (intRet = 0)
360             If Not blnRet Then strMsg = "签名失败！"
370         End If
380     Else
390         strMsg = "签名失败！"
400         blnRet = False
410     End If
420     If strMsg <> "" Then
430         MsgBoxEx strMsg, vbInformation, gstrSysName
440     End If
                
450     LNCA_Sign = blnRet
460     Exit Function
errH:
470     MsgBoxEx "签名失败！" & vbNewLine & "第" & CStr(Erl()) & "行 " & Err.Description, vbQuestion, gstrSysName
End Function

Public Function LNCA_VerifySign(ByVal strCert As String, ByVal strSignData As String, ByVal strSource As String) As Boolean
'验证签名
'
    Dim strParameter As String
    Dim strRet As String
    Dim blnRet As Boolean
    Dim strMsg As String
    Dim lngRet As Long
    
    On Error GoTo errH
    If mbytModel = E_RSA Then
        strSource = EncodeBase64String(strSource)
        strCert = strCert & "123"
        strParameter = "cmd=verifysm2" & "&text=" & EnCodeURL(strSource) & "&cert=" & EnCodeURL(strCert) & "&signed=" & EnCodeURL(strSignData)
        strRet = HttpPost(gudtPara.strSignURL, strParameter, responseText)
        blnRet = GetSubString(strRet, "ret") = "1"    '返回值=1验证签名成功
    ElseIf mbytModel = E_SM2 Then
        '参数 签名结果;签名原文;预留;服务器证书
        lngRet = mobjKeyGateOper.ReqVerifySig(strSignData, strSource, 0, strCert)
        blnRet = lngRet = 0
    End If
    If blnRet Then    '验证签名失败
        strMsg = "验证成功，该电子签名数据有效！"
    Else
        strMsg = "验签失败！"
    End If
    
    If strMsg <> "" Then
        MsgBoxEx strMsg, vbInformation, gstrSysName
    End If
        
    LNCA_VerifySign = blnRet
    
    Exit Function
errH:
104     MsgBoxEx "验证签名失败！" & vbNewLine & "第" & CStr(Erl()) & "行 " & Err.Description, vbQuestion, gstrSysName
End Function

Public Sub LNCA_UnLoadObj()
    If mbytModel = E_RSA Then
        Set mobjUSBKEY = Nothing
        Set mobjMSScriptCtl = Nothing
    Else
        Set mobjCert = Nothing
        Set mobjKeyGateOper = Nothing
        Set mobjKeyManager = Nothing
        Set mobjKeySeal = Nothing
        Set mobjKeySealArray = Nothing
        Set mobjKeyStore = Nothing
        Set mobjKeyDetector = Nothing
    End If
    
    mblnInit = False
End Sub

Private Function EnCodeURL(ByVal strUrl As String) As String
'功能:将传人字符串按UTF编码方式转换成十六进制的转义序列
'说明：encodeURI-javaScript方法不会对 ASCII 字母和数字进行编码，也不会对这些 ASCII 标点符号进行编码： - _ . ! ~ * ' ( )
    Dim i As Long
    Dim strChar As String
    Dim intAsc As Integer
    Dim strRet As String
    
    For i = 1 To Len(strUrl)
        strChar = Mid(strUrl, i, 1)
        intAsc = Asc(strChar)
        If intAsc >= 0 And intAsc <= 127 Then
           strChar = "%" & Hex(intAsc)
        Else
            strChar = mobjMSScriptCtl.Eval("encodeURI(""" & strChar & """)")
        End If
        strRet = strRet & strChar
    Next
    
    EnCodeURL = strRet
End Function

Public Function LNCA_GetPara() As Boolean
'设置服务器地址
    Dim arrTmp As Variant
    
    On Error GoTo errH
    gstrPara = gobjComLib.zlDatabase.GetPara(90000, glngSys)
    If gstrPara = "" Then gstrPara = "http://218.25.86.214:2010/ssoworker"
    arrTmp = Split(gstrPara, G_STR_SPLIT)
    If UBound(arrTmp) > 0 Then
        gudtPara.strSignURL = arrTmp(0)
        gudtPara.bytSignVersion = Val(arrTmp(1) & "")
        If UBound(arrTmp) > 1 Then
            gudtPara.strSIGNIP = arrTmp(2)
        End If
    Else
        gudtPara.strSignURL = arrTmp(0)  '
        gudtPara.strSIGNIP = ""
    End If

    Exit Function
errH:
    MsgBoxEx "读取参数失败！" & vbNewLine & "第" & CStr(Erl()) & "行 " & Err.Description, vbInformation, gstrSysName
End Function

Public Function LNCA_SetParaStr() As String
    LNCA_SetParaStr = gudtPara.strSignURL & G_STR_SPLIT & gudtPara.bytSignVersion & G_STR_SPLIT & gudtPara.strSIGNIP
End Function

Private Function GetSubString(ByVal strSource As String, ByVal strNode As String) As String
'功能:获取返回字符串中某个节点值
'参数:strSource -传人字符串 {"ret":1,"errinfo":"","rand":"3XO9JCXVJ6LF05M51165"}
'    strNode-标识要获取的节点名称
    Dim arrMain As Variant
    Dim arrSub As Variant
    Dim strRet As String
    Dim i As Long
    
    arrMain = Split(strSource, ",")
    For i = LBound(arrMain) To UBound(arrMain)
        Select Case UCase(strNode)
        Case UCase("rand"), UCase("token")
            If InStr(UCase(arrMain(i)), UCase(strNode)) > 0 Then
                arrSub = Split(arrMain(i), ":")
                strRet = Mid(arrSub(1), 2)
                strRet = left(strRet, Len(strRet) - 2)
                Exit For
            End If
        Case UCase("ret")
            If InStr(UCase(arrMain(i)), UCase(strNode)) > 0 Then
                arrSub = Split(arrMain(i), ":")
                strRet = arrSub(1)
                Exit For
            End If
        Case UCase$("errinfo")
            If InStr(UCase(arrMain(i)), UCase(strNode)) > 0 Then
                arrSub = Split(arrMain(i), ":")
                strRet = Mid(arrSub(1), 2)
                strRet = left(strRet, Len(strRet) - 1)
                Exit For
            End If
        End Select
    Next
    GetSubString = strRet
End Function





