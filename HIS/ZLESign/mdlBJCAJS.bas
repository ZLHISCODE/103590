Attribute VB_Name = "mdlBJCAJS"
Option Explicit

'北京CA中心功能模块(新江苏版_宿迁市人民医院项目)
Private mstrLastPwd As String    '缓存上次输入密码
Private mintLogin As Integer     '输入密码次数
Private BJCAJS_Pic As Object
Private BJCAJS_Client As Object       '客户端证书部件
Private BJCAJS_svs As Object          '签名验证控件
Private BJCAJS_TS As Object           '时间戳控件
Private mstrLogins As String          '标记已经通过登录验证的key的序列号
Private mblnInit As Boolean

Public Function BJCAJS_InitObj() As Boolean
        '证书部件初始化
        On Error GoTo errH
    
100     If mblnInit Then BJCAJS_InitObj = True: Exit Function
    
102     Set BJCAJS_Client = CreateObject("XTXAppCOM.XTXApp.1")
104     Set BJCAJS_Pic = CreateObject("GetKeyPic.GetPic")
106     Set BJCAJS_svs = CreateObject("BJCA_SVS_ClientCOM.BJCASVSEngine.1")   '创建证书验证控件
108     Set BJCAJS_TS = CreateObject("BJCA_TS_ClientCom.BJCATSEngine")    '创建时间戳控件
    
110     mintLogin = 0
112     mstrLogins = ""
114     BJCAJS_InitObj = True
116     mblnInit = True
        Exit Function
errH:
118      MsgBoxEx "创建接口部件失败！" & vbNewLine & Err.Description, vbQuestion, gstrSysName
End Function

Public Function BJCAJS_RegCert(arrCertInfo As Variant, Optional ByVal strUserID As String) As Boolean
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
            On Error GoTo errH
        
100         For i = LBound(arrCertInfo) To UBound(arrCertInfo)
102              arrCertInfo(i) = ""
            Next
        
104         If BJCAJS_GetCertList(strCertUserName, strCertSn, strCertDN, strCertUserID, strCert) Then
106             If UCase(strCertUserID) <> UCase(strUserID) Then
108                 MsgBoxEx "用户身份证号：" & _
                               vbCrLf & vbTab & "【" & UCase(strUserID) & "】" & vbCrLf & _
                               "当前证书唯一标识:" & _
                               vbCrLf & vbTab & "【" & UCase(strCertUserID) & "】" & vbCrLf & _
                               "用户身份证号与当前证书唯一标识不相等,不能注册！", vbInformation, gstrSysName
                    Exit Function
                End If
110             arrCertInfo(0) = strCertUserName
112             arrCertInfo(1) = strCertDN '证书DN
114             arrCertInfo(2) = strCertSn '证书序列号 签名时要用
116             arrCertInfo(3) = strCert
118             arrCertInfo(4) = ""
120             arrCertInfo(5) = SaveBase64ToFile("gif", strCertUserID, BJCAJS_Pic.getpic())
122             BJCAJS_RegCert = True
            End If

            Exit Function
errH:
124      MsgBoxEx "获取证书信息失败！" & vbNewLine & "第" & CStr(Erl()) & "行 " & Err.Description, vbQuestion, gstrSysName

End Function

Public Function BJCAJS_CheckCert(ByRef blnReDo As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------
    '功能：读取USB进行设备初始化并登录
    '参数：
    '   出参 blnRedo:True-重新注册成功
    '---------------------------------------------------------------------------------------------------------------------
        Dim strKey As String, strPIN As String, strUserName As String
        Dim strCertName As String, strCertDN As String
        Dim strCertSn As String
        Dim strCertUserID As String    '包含身份证号信息
        Dim strDate As String
        Dim udtUser As USER_INFO
        Dim strCert As String, strCertID As String
        Dim blnOk As Boolean
        Dim blnRet As Boolean
    
        On Error GoTo errH
    
100     If BJCAJS_Client Is Nothing Then Set BJCAJS_Client = CreateObject("XTXAppCOM.XTXApp.1")
102     If BJCAJS_Pic Is Nothing Then Set BJCAJS_Pic = CreateObject("GetKeyPic.GetPic")
104     If BJCAJS_svs Is Nothing Then Set BJCAJS_svs = CreateObject("BJCA_SVS_ClientCOM.BJCASVSEngine.1")   '创建证书验证控件
106     If BJCAJS_TS Is Nothing Then Set BJCAJS_TS = CreateObject("BJCA_TS_ClientCom.BJCATSEngine")    '创建时间戳控件
    
         '获取证书信息同时检查Key盘是否插入
108     If Not BJCAJS_GetCertList(strCertName, strCertSn, strCertDN, strCertUserID, strCert, strCertID) Then
110         BJCAJS_CheckCert = False: Exit Function
        End If
        '未注册在当前用户名下的Key
112     If mUserInfo.strUserID = "" Then
114         MsgBoxEx "您的身份证号为空,请联系管理员到人员管理中录入！", vbOKOnly + vbInformation, gstrSysName
            Exit Function
116     ElseIf UCase(strCertUserID) <> UCase(mUserInfo.strUserID) Then
118         MsgBoxEx "您的身份证号：" & _
                       vbCrLf & vbTab & "【" & UCase(mUserInfo.strUserID) & "】" & vbCrLf & _
                       "当前证书唯一标识:" & _
                       vbCrLf & vbTab & "【" & UCase(strCertUserID) & "】" & vbCrLf & _
                       "用户身份证号与当前证书唯一标识不相等,不能使用！", vbInformation, gstrSysName
            Exit Function
        End If
        '输入密码
120     If mstrLastPwd <> "" Then strPIN = mstrLastPwd
122     If strPIN = "" Then
124         If Not frmPassword.ShowMe(strPIN) Then Exit Function
        End If
        '密码验证如果不调用,首次调用签名接口时会触发CA的密码窗口
126     If strPIN = "" Then
128        MsgBoxEx "请输入证书密码！", vbOKOnly + vbInformation, gstrSysName
           Exit Function
        Else
130         If mintLogin >= 8 Then
132             MsgBoxEx "已经输入了" & mintLogin & "次错误密码，超过了最大输入次数！", vbOKOnly + vbInformation, gstrSysName
                Exit Function
            End If
134         blnRet = BJCAJS_Client.SOF_Login(strCertID, strPIN)
136         If Not blnRet Then
138             mintLogin = mintLogin + 1
140             MsgBoxEx "证书密码可能不正确，您已经输入了" & mintLogin & "次密码，还可以输入" & 8 - mintLogin & "次!", vbOKOnly + vbInformation, gstrSysName
142             mstrLastPwd = ""
                Exit Function
            End If
         End If
     
        '登录验证
144     If InStr(mstrLogins & "|", "|" & strCertSn & "|") > 0 Then '首次验证通过后，下次不在继续验证
146         blnOk = True
        Else
148         If Not GetCertLogin(strCertID, strCert) Then
150             blnOk = False
            Else
152             blnOk = True
154             If InStr(mstrLogins & "|", "|" & strCertSn & "|") = 0 Then mstrLogins = mstrLogins & "|" & strCertSn
            End If
        End If
    
156     If blnOk Then
            '判断是否需要更新注册证书
158         udtUser.strName = strCertName
160         udtUser.strSignName = strCertName
162         udtUser.strUserID = strCertUserID
164         udtUser.strCertSn = strCertSn
166         udtUser.strCertDN = strCertDN
168         udtUser.strCert = strCert
170         udtUser.strEncCert = ""
172         udtUser.strCertID = strCertID
174         udtUser.strPicCode = BJCAJS_Pic.getpic()
            '获取已经注册证书的有效结束日期
176         strDate = BJCAJS_Client.SOF_GetCertInfo(mUserInfo.strCert, 12)
178         strDate = String14ToDate(strDate)
180         If IsUpdateRegCert(udtUser, strDate, blnReDo) Then
182             BJCAJS_CheckCert = True
            Else
184             BJCAJS_CheckCert = False
            End If
        End If
    
186     mUserInfo.strCertID = strCertID
    
188     mstrLastPwd = strPIN
        Exit Function
errH:
190      MsgBoxEx "检查USBKEY失败！" & vbNewLine & "第" & CStr(Erl()) & "行 " & Err.Description, vbInformation, gstrSysName
End Function

Public Function BJCAJS_Sign(ByVal strCurrCertSn As String, ByVal strSource As String, ByRef strSignData As String, _
            ByRef strTimeStamp As String, ByRef strTimeStampCode As String, ByRef blnReDo As Boolean) As Boolean
    '签名
    '参数：
    '   strPID --用户身份标识（一般为身份证号）
        Dim strSigCert As String
        Dim CertID As String
        Dim strRequest As String    '时间戳请求
        Dim strDate As String
        Dim strMsg As String
        Dim blnCheck As Boolean
        On Error GoTo errH
    
100     blnCheck = BJCAJS_CheckCert(blnReDo)
102     If blnReDo Then Exit Function
    
104     If blnCheck Then                '验证当前USB是否是签名用户的，并获取签名证书
106         strSignData = BJCAJS_Client.SOF_SignData(mUserInfo.strCertID, strSource) '产生签名数据
108         If strSignData <> "" Then
                '源文加签名值产生时间戳请求(不带证书)
110             strRequest = BJCAJS_TS.CreateTSRequest(strSource & strSignData, 0)
112             If strRequest <> "" Then
114                 strTimeStampCode = BJCAJS_TS.CreateTS(strRequest)  '产生时间戳（带证书）
116                 If strTimeStampCode = "" Then
118                     MsgBoxEx "生成时间戳失败！", vbOKOnly + vbInformation, gstrSysName
120                     BJCAJS_Sign = False
                        Exit Function
                    End If
122                 strDate = BJCAJS_TS.gettimestampinfo(strTimeStampCode, 1)  '取得时间戳时间
124                 strTimeStamp = String14ToDate(strDate)
                Else
126                 MsgBoxEx "时间戳请求失败！", vbOKOnly + vbInformation, gstrSysName
128                 BJCAJS_Sign = False
                    Exit Function
                End If
          
130             BJCAJS_Sign = True
            Else
132             MsgBoxEx "签名失败！", vbOKOnly + vbInformation, gstrSysName
            End If
        End If
    
        Exit Function
errH:
134      MsgBoxEx "签名失败！" & vbNewLine & "第" & CStr(Erl()) & "行 " & Err.Description, vbInformation, gstrSysName
End Function

'''验证签名函数
Public Function BJCAJS_VerifySign(ByVal strCertSn As String, ByVal strSignData As String, ByVal strSource As String, ByVal strTimeStampCode As String) As Boolean
    '验证签名
    Dim strTmp As String
    Dim intRet As Integer
    Dim blnOk As Boolean
    Dim strDate As String
    Dim strTimeStamp As String
    On Error GoTo errH
    
100 If BJCAJS_svs Is Nothing Then Set BJCAJS_svs = CreateObject("BJCA_SVS_ClientCOM.BJCASVSEngine.1")   '创建证书验证控件
102 If BJCAJS_TS Is Nothing Then Set BJCAJS_TS = CreateObject("BJCA_TS_ClientCom.BJCATSEngine")    '创建时间戳控件
    '验证时间戳
104 strTmp = ""
106 intRet = BJCAJS_TS.VerifyTS(strTimeStampCode, "")  '不验证源文
108 If intRet <> 0 Then
110     strTmp = "验证时间戳失败！" & GetReturnInfo(intRet)
112     blnOk = False
    Else
114     strDate = BJCAJS_TS.gettimestampinfo(strTimeStampCode, 1)  '取得时间戳时间
116     strTimeStamp = String14ToDate(strDate)
118     strTmp = "验证时间戳成功！" & vbTab & "签名时间:" & strTimeStamp
120     blnOk = True
    End If
    
    '验证签名
122 intRet = BJCAJS_svs.VerifySignatureBySN(strCertSn, strSource, strSignData)
124 If (intRet = 0) Then
126     strTmp = IIf(strTmp <> "", strTmp & vbCrLf, "") & "数据签名验证成功！"
128     blnOk = True And blnOk
    Else
130     strTmp = IIf(strTmp <> "", strTmp & vbCrLf, "") & "数据签名验证失败！"
132     blnOk = False
    End If

134 If strTmp <> "" Then
136     MsgBoxEx strTmp, vbOKOnly + vbInformation, gstrSysName
    End If
    
138 BJCAJS_VerifySign = blnOk
    Exit Function
errH:
140     MsgBoxEx "验证签名失败！" & vbNewLine & "第" & CStr(Erl()) & "行 " & Err.Description, vbInformation, gstrSysName
End Function

'销毁对象
Public Sub BJCAJS_UloadObj()
    Set BJCAJS_Client = Nothing
    Set BJCAJS_svs = Nothing
    Set BJCAJS_TS = Nothing
    mblnInit = False
End Sub

'----- 以下是内部函数
Private Function GetCertLogin(ByVal strCertID As String, ByVal strCert As String) As Boolean
        '北京CA江苏版数字证书登录函数
        '- 入参
        'strCertID            :证书ID
        'strCert              证书内容BASE64编码
        Dim random As String
        Dim serverCert As String
        Dim serverSign As String, strSignVal As String
        Dim blnRet As Boolean
        Dim strDate As String
        Dim intDay As Integer, intRetSign As Integer, intRetVal As Integer
        Dim strTmp As String
        Dim lngRet As Long
    
        On Error GoTo errH
        '1)（BJCA_SVS_ClientCOM组件）HIS系统调用CA接口，获取随机数、服务器证书，并通过服务器证书对随机数进行签名；
100     random = BJCAJS_svs.GenRandom(16) '获取随机数
102     serverCert = BJCAJS_svs.GetServerCertificate() '获取服务器证书
104     serverSign = BJCAJS_svs.SignData(random) '服务端对随机数签名
    
106     blnRet = BJCAJS_Client.SOF_VerifySignedData(serverCert, random, serverSign) '客户端验证服务端签名
108     If Not blnRet Then
110         MsgBoxEx "服务端签名验证失败！", vbOKOnly + vbInformation, gstrSysName
            Exit Function
        End If
        '验证证书是否过期
112     strDate = BJCAJS_Client.SOF_GetCertInfo(strCert, 12)
114     strDate = String14ToDate(strDate)
116     If strDate <> "" Then
        '验证客户端证书有效期剩余天数
118         intDay = CheckValidaty(CDate(strDate))
    
120         If (intDay <= 30 And intDay > 0 And Not gblnShow) Then
122             MsgBoxEx "您的证书还有" & intDay & "天过期。", vbOKOnly + vbInformation, gstrSysName
124             gblnShow = True
126         ElseIf (intDay <= 0) Then
128             MsgBoxEx "您的证书已过期 " & Abs(intDay) & " 天。", vbOKOnly + vbInformation, gstrSysName
                Exit Function
            End If
        End If
    
130     lngRet = BJCAJS_svs.ValidateCertificate(strCert)
132     If lngRet <> 0 Then
134         If lngRet = -1 Then
136             MsgBoxEx "不是所信任的根!", vbOKOnly + vbInformation, gstrSysName
                Exit Function
138         ElseIf lngRet = -2 Then
140             MsgBoxEx "证书超过有效期！", vbOKOnly + vbInformation, gstrSysName
                Exit Function
142         ElseIf lngRet = -3 Then
144             MsgBoxEx "证书已经作废！", vbOKOnly + vbInformation, gstrSysName
                Exit Function
146         ElseIf lngRet = -4 Then
148             MsgBoxEx "证书被放入黑名单！", vbOKOnly + vbInformation, gstrSysName
                Exit Function
150         ElseIf lngRet = -5 Then
152             MsgBoxEx "证书未生效！", vbOKOnly + vbInformation, gstrSysName
                Exit Function
            End If
        End If
    
        '验证证书是否过期
154     strSignVal = BJCAJS_Client.SOF_SignData(strCertID, random)  '客户端随机数签名
156     intRetSign = BJCAJS_svs.VerifySignedData(strCert, random, strSignVal)   '服务端验证客户端签名
158     intRetVal = BJCAJS_svs.ValidateAndSaveCertificate(strCert)  '服务端验证客户端证书有效性并保存证书
    
160     If Not (intRetSign = 0 And (intRetVal = 0 Or intRetVal = 1)) Then
162         MsgBoxEx "客户端证书验失败！", vbOKOnly + vbInformation, gstrSysName
            Exit Function
        End If
    
164     GetCertLogin = True
        Exit Function
errH:
166     MsgBoxEx "登录验证失败！" & vbNewLine & "第" & CStr(Erl()) & "行 " & Err.Description, vbInformation, gstrSysName
End Function

''' 获取客户端证书列表
''' 返回boolean
Public Function BJCAJS_GetCertList(ByRef strName As String, Optional ByRef strCertSn As String = "0", Optional ByRef strCertDN As String = "0", _
            Optional ByRef strCertUserID As String = "0", Optional ByRef strCert As String, Optional ByRef strCertID As String) As Boolean
        '北京CA江苏版获取数字证书列表函数
        '-入参:无
        '-出参
        'strName :      保存接口返回的证书所有者姓名
        'strCertSN      保存接口返回的证书SN
        'strCertDN:     保存接口返回的证书DN
        'strCertUserID:  保存接口返回的证书所有者唯一标识
        'strCert:       保存接口返回的签名证书
        'strCertID      返回证书ID
        Dim strUsbkeyList As String
        Dim arrUserListLength As Integer
        Dim arrUserList() As String
        Dim strUser As String
        Dim i As Integer
        
        On Error GoTo errH
        
100     If BJCAJS_Client Is Nothing Then Set BJCAJS_Client = CreateObject("XTXAppCOM.XTXApp.1")
102     If BJCAJS_Pic Is Nothing Then Set BJCAJS_Pic = CreateObject("GetKeyPic.GetPic")
        '获取证书
104     strUsbkeyList = BJCAJS_Client.SOF_GetUserList()
106     If (strUsbkeyList = "") Then
108         strName = ""
110         MsgBoxEx "请插入证书Key！", vbOKOnly + vbInformation, gstrSysName
112         BJCAJS_GetCertList = False
            Exit Function
        Else
114         arrUserList = Split(strUsbkeyList, "&&&") '宿迁人民四(测试)||999000100089956/6001201312021788&&&宿迁人民二(测试)||999000100089948/6002201309019595&&&
116         If UBound(arrUserList) > 1 Then  '多个KEY
118             For i = LBound(arrUserList) To UBound(arrUserList) - 1
120                 strUser = strUser & "&&&" & Split(arrUserList(i), "||")(0)
                Next
122             If strUser <> "" Then strUser = Mid(strUser, 4)
124             strName = frmSelectUser.ShowMe(strUser)
            
126             For i = LBound(arrUserList) To UBound(arrUserList) - 1
128                If strName = Split(arrUserList(i), "||")(0) Then
130                     strCertID = Split(arrUserList(i), "||")(1)
                        Exit For
                   End If
                Next
            Else
132             arrUserList = Split(arrUserList(0), "||")
134             strName = arrUserList(0)      '证书CN通用名
136             strCertID = arrUserList(1)    '证书ID
            End If
        
138         strCert = BJCAJS_Client.SOF_ExportUserCert(strCertID) '3.导出签名证书。
        
140         If strCertSn <> "0" Then strCertSn = BJCAJS_Client.SOF_GetCertInfo(strCert, 2) '证书序列号 签名时要用
142         If strCertDN <> "0" Then strCertDN = BJCAJS_Client.SOF_GetCertInfo(strCert, 33) '证书DN
144         If strCertUserID <> "0" Then
146             strCertUserID = BJCAJS_Client.SOF_GetCertInfoByOid(strCert, "1.2.156.112562.2.1.1.1") '2.获取证书唯一标识（一般为身份证号）SF+身份证号
148             strCertUserID = Mid(strCertUserID, 3)
            End If
        End If
150     BJCAJS_GetCertList = True
        Exit Function
errH:
    MsgBoxEx "获取证书信息失败！" & vbNewLine & "第" & CStr(Erl()) & "行 " & Err.Description, vbQuestion, gstrSysName
End Function

Private Function GetReturnInfo(ByVal intErrNum As Integer) As String
    '准格尔时间戳返回信息转换函数
    If intErrNum = -1 Then
        GetReturnInfo = "时间戳验证不通过"
    ElseIf intErrNum = -2 Then
        GetReturnInfo = "原文验证不通过"
    ElseIf intErrNum = -3 Then
        GetReturnInfo = "不是所信任的根"
    ElseIf intErrNum = -4 Then
        GetReturnInfo = "证书未生效"
    ElseIf intErrNum = -5 Then
        GetReturnInfo = "查询不到此证书"
    ElseIf intErrNum = -6 Then
        GetReturnInfo = "签发时间戳时服务器证书过期"
    ElseIf intErrNum = 0 Then
        GetReturnInfo = "验证成功"
    Else
        GetReturnInfo = "未知错误"
    End If
    If GetReturnInfo <> "" Then
        GetReturnInfo = "时间戳接口返回提示：" & GetReturnInfo
    End If
End Function



