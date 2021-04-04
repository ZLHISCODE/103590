Attribute VB_Name = "mdlBJCAGX"
Option Explicit

'北京CA中心功能模块(新广西版)
Private mblnInit As Boolean         '是否已初始化成功

Private BJCAGX_Client As Object       '客户端证书部件
Private BJCAGX_svs As Object          '签名验证控件
Private BJCAGX_TS As Object           '时间戳控件
Private mobjPic As Object             '获取签章图片             '
Private mstrLastPwd As String    '缓存上次输入密码
Private mintLogin As Integer     '输入密码次数
Private mstrLogins As String          '标记已经通过登录验证的key的序列号
Public gobjGXCAPenSign As Object

Public Enum Version
    V_RSA = 0
    V_SM2 = 1
End Enum

Public Function BJCAGX_InitObj() As Boolean
        '证书部件初始化
        Dim progID As String
    
        On Error GoTo errH
100 BJCAGX_InitObj = mblnInit
102 If mblnInit Then Exit Function
104    If Not BJCAGX_GetPara(1) Then Exit Function

106    If gudtPara.blnSignPic Then
108        Set mobjPic = CreateObject("GetKeyPic.GetPic.1")
       End If

110    Set BJCAGX_svs = CreateObject("BJCA_SVS_ClientCOM.BJCASVSEngine.1")   '创建证书验证控件

112    Set BJCAGX_TS = CreateObject("BJCA_TS_ClientCom.BJCATSEngine")    '创建时间戳控件

114    If gudtPara.bytSignVersion = V_RSA Then
116        Set BJCAGX_Client = CreateObject("BJCAAPPCTRL.BjcaAppCtrlCtrl.1") '创建签名控件
       Else
118        Set BJCAGX_Client = CreateObject("XTXAppCOM.XTXApp.1")
       End If

120 BJCAGX_InitObj = True
122    mblnInit = BJCAGX_InitObj
       Exit Function
errH:
124 MsgBoxEx "创建接口部件失败！" & vbNewLine & "第" & CStr(Erl()) & "行 " & Err.Description, vbQuestion, gstrSysName
End Function

Public Function BJCAGX_RegCert(arrCertInfo As Variant) As Boolean
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
        Dim strFilePath As String
        On Error GoTo errH
    
        For i = LBound(arrCertInfo) To UBound(arrCertInfo)
             arrCertInfo(i) = ""
        Next
        
        If GetCertList(strCertUserName, strKeyId, strSigCert, strFilePath, strCertDN) Then
            arrCertInfo(0) = strCertUserName
            arrCertInfo(1) = strCertDN
            arrCertInfo(2) = strKeyId
            arrCertInfo(3) = strSigCert
            arrCertInfo(5) = strFilePath
            BJCAGX_RegCert = True
        End If

        Exit Function
errH:
     MsgBoxEx "获取证书信息失败！" & vbNewLine & "第" & CStr(Erl()) & "行 " & Err.Description, vbQuestion, gstrSysName

End Function

Public Function BJCAGX_CheckCert(ByRef blnReDo As Boolean) As Boolean
        '功能：读取USB进行设备初始化并登录
    Dim strCertSN As String, strPIN As String, strCert As String
    Dim strCertUserID As String, strCertName As String
    Dim strCertID As String, strCertDN As String, strPicCode As String
    Dim blnRet As Boolean
    Dim blnOk As Boolean
    Dim udtUser As USER_INFO
    Dim strDate As String
    
    On Error GoTo errH
    
    If Not BJCAGX_InitObj() Then
        MsgBoxEx "部件未初始化！", vbInformation, gstrSysName
        Exit Function
    End If

    '获取证书信息同时检查Key盘是否插入
    If Not GetCertList(strCertName, strCertSN, strCert, 0, strCertDN, strCertUserID, strCertID, strPicCode) Then
        BJCAGX_CheckCert = False: Exit Function
    End If
    
    If gudtPara.bytSignVersion = V_RSA Then
        If mUserInfo.strCertSN <> strCertSN Then
            MsgBoxEx "该证书未注册在您的名下，不能使用！", vbOKOnly + vbInformation, gstrSysName
            Exit Function
        End If
        
        If Not GetCertLogin(strCertSN, strCert) Then
            BJCAGX_CheckCert = False
        Else
            BJCAGX_CheckCert = True
        End If
        blnReDo = False
        
    ElseIf gudtPara.bytSignVersion = V_SM2 Then
        '未注册在当前用户名下的Key
        If mUserInfo.strUserID = "" Then
            MsgBoxEx "您的身份证号为空,请联系管理员到人员管理中录入！", vbOKOnly + vbInformation, gstrSysName
            Exit Function
        ElseIf strCertUserID <> mUserInfo.strUserID Then
            MsgBoxEx "您的身份证号：" & _
                       vbCrLf & vbTab & "【" & mUserInfo.strUserID & "】" & vbCrLf & _
                       "当前证书唯一标识:" & _
                       vbCrLf & vbTab & "【" & strCertUserID & "】" & vbCrLf & _
                       "用户身份证号与当前证书唯一标识不相等,不能使用！", vbInformation, gstrSysName
            Exit Function
        End If
        '输入密码
        If mstrLastPwd <> "" Then strPIN = mstrLastPwd
        If strPIN = "" Then
            If Not frmPassword.ShowMe(strPIN) Then Exit Function
        End If
        '密码验证如果不调用,首次调用签名接口时会触发CA的密码窗口
        If strPIN = "" Then
           MsgBoxEx "请输入证书密码！", vbOKOnly + vbInformation, gstrSysName
           Exit Function
        Else
            If mintLogin >= 8 Then
                MsgBoxEx "已经输入了" & mintLogin & "次错误密码，超过了最大输入次数！", vbOKOnly + vbInformation, gstrSysName
                Exit Function
            End If
            blnRet = BJCAGX_Client.SOF_Login(strCertID, strPIN)
            If Not blnRet Then
                mintLogin = mintLogin + 1
                MsgBoxEx "证书密码可能不正确，您已经输入了" & mintLogin & "次密码，还可以输入" & 8 - mintLogin & "次!", vbOKOnly + vbInformation, gstrSysName
                mstrLastPwd = ""
                Exit Function
            End If
         End If
         
        '登录验证
        If InStr(mstrLogins & "|", "|" & strCertSN & "|") > 0 Then '首次验证通过后，下次不在继续验证
            blnOk = True
        Else
            If Not GetCertLogin(strCertSN, strCert, strCertID) Then
                blnOk = False
            Else
                blnOk = True
                If InStr(mstrLogins & "|", "|" & strCertSN & "|") = 0 Then mstrLogins = mstrLogins & "|" & strCertSN
            End If
        End If
        
        If blnOk Then
            '判断是否需要更新注册证书
            udtUser.strName = strCertName
            udtUser.strSignName = strCertName
            udtUser.strUserID = strCertUserID
            udtUser.strCertSN = strCertSN
            udtUser.strCertDN = strCertDN
            udtUser.strCert = strCert
            udtUser.strEncCert = ""
            udtUser.strCertID = strCertID
            udtUser.strPicCode = strPicCode
            '获取已经注册证书的有效结束日期
            strDate = BJCAGX_Client.SOF_GetCertInfo(mUserInfo.strCert, 12)
            strDate = String14ToDate(strDate)
            If IsUpdateRegCert(udtUser, strDate, blnReDo) Then
                BJCAGX_CheckCert = True
            Else
                BJCAGX_CheckCert = False
            End If
        End If
        
        mUserInfo.strCertID = strCertID
        
        mstrLastPwd = strPIN
    End If
    Exit Function
errH:
     MsgBoxEx "检查USBKEY失败！" & vbNewLine & "第" & CStr(Erl()) & "行 " & Err.Description, vbQuestion, gstrSysName
End Function

Public Function BJCAGX_Sign(ByVal strCurrCertSn As String, ByVal strSource As String, ByRef strSignData As String, _
    ByRef strTimeStamp As String, ByRef strTimeStampCode As String, ByRef blnReDo As Boolean) As Boolean
            '签名
        Dim strSigCert As String
        Dim strRequest As String    '时间戳请求
        Dim strDate As String
        Dim strMsg As String
        Dim blnRet As Boolean, blnCheck As Boolean
        Dim bytTsVer    As Byte  '1-启用老版本时间戳接口
        
        On Error GoTo errH
100     If gudtPara.bytSignVersion = V_RSA Then
102         If BJCAGX_CheckCert(blnReDo) Then                  '验证当前USB是否是签名用户的，并获取签名证书
104             strSignData = BJCAGX_Client.signedData(mUserInfo.strCertSN, strSource)  '产生签名数据
106             If strSignData <> "" Then
108                 blnRet = True
110                 strRequest = BJCAGX_TS.CreateTimeStampRequest(strSource)  '产生时间戳请求
112                 If strRequest <> "" Then
114                     strTimeStampCode = BJCAGX_TS.CreateTimeStamp(strRequest)   '产生时间戳（带证书）
116                     If strTimeStampCode <> "" Then
118                         strDate = BJCAGX_TS.gettimestampinfo(strTimeStampCode, 1)
120                         strTimeStamp = GetTimeStamp(strDate)          '取得时间戳时间
122                         If strTimeStamp = "" Then
124                             strMsg = "解析时间戳失败！"
126                             blnRet = False
                            End If
                        Else
128                         strMsg = "产生时间戳失败！"
130                         blnRet = False
                        End If
                    Else
132                     strMsg = "时间戳请求失败！"
134                     blnRet = False
                    End If
                Else
136                 strMsg = "签名失败！"
138                 blnRet = False
                End If
            Else
140             strMsg = "验证签名失败！"
142             blnRet = False
            End If
        Else
144         blnCheck = BJCAGX_CheckCert(blnReDo)
146         If blnReDo Then Exit Function
148         If blnCheck Then                '验证当前USB是否是签名用户的，并获取签名证书
150             strSignData = BJCAGX_Client.SOF_SignData(mUserInfo.strCertID, strSource) '产生签名数据
152             If strSignData <> "" Then
                    '源文加签名值产生时间戳请求(不带证书)
154                 blnRet = True
                    On Error Resume Next
156                 strRequest = BJCAGX_TS.CreateTSRequest(strSource & strSignData, 0)
158                 If Err.Number = 438 Or strRequest = "" Then '对象不支持该属性或方法 兼容老版本
160                     strRequest = BJCAGX_TS.CreateTimeStampRequest(strSource & strSignData)  '产生时间戳请求
162                     bytTsVer = 1
                    End If
164                 Err.Clear: On Error GoTo errH
166                 If strRequest <> "" Then
168                     If bytTsVer = 0 Then
170                         strTimeStampCode = BJCAGX_TS.CreateTS(strRequest)  '产生时间戳（带证书）
172                         If strTimeStampCode = "" Then
174                             MsgBoxEx "生成时间戳失败！", vbOKOnly + vbInformation, gstrSysName
176                             blnRet = False
                            Else
178                             strDate = BJCAGX_TS.gettimestampinfo(strTimeStampCode, 1)  '取得时间戳时间
180                             strTimeStamp = String14ToDate(strDate)
182                             blnRet = True
                            End If
                        Else
184                         strTimeStampCode = BJCAGX_TS.CreateTimeStamp(strRequest)   '产生时间戳（带证书）
186                         If strTimeStampCode <> "" Then
188                             strDate = BJCAGX_TS.gettimestampinfo(strTimeStampCode, 1)
190                             strTimeStamp = GetTimeStamp(strDate)          '取得时间戳时间
192                             If strTimeStamp = "" Then
194                                 strMsg = "解析时间戳失败！"
196                                 blnRet = False
                                End If
                            Else
198                             strMsg = "产生时间戳失败！"
200                             blnRet = False
                            End If
                        End If
                    Else
202                     strMsg = "时间戳请求失败！"
204                     blnRet = False
                    End If

                Else
206                 strMsg = "验证签名失败！"
208                 blnRet = False
                End If
            End If
        End If
    
210     If strMsg <> "" Then
212         MsgBoxEx strMsg, vbOKOnly + vbInformation, gstrSysName
        End If
    
214     BJCAGX_Sign = blnRet
    
        Exit Function
errH:
216      MsgBoxEx "签名失败！" & vbNewLine & "第" & CStr(Erl()) & "行 " & Err.Description, vbQuestion, gstrSysName
End Function

Public Function BJCAGX_VerifySign(ByVal strCertSN As String, ByVal strSigCert As String, ByVal strSignData As String, ByVal strSource As String, ByVal strTStampCode As String) As Boolean
        '验证签名
        Dim strTmp As String
        Dim blnRet As Boolean
        Dim lngRuslt As Long
        Dim intRet As Integer
        Dim strDate As String
        Dim strTimeStamp As String
        Dim bytVer  As Byte
        
        On Error GoTo errH

100     Call BJCAGX_InitObj
        
102     If gudtPara.bytSignVersion = V_RSA Then
104         blnRet = BJCAGX_Client.VerifySignedData(strSigCert, strSource, strSignData)
106         If blnRet Then
108             strTmp = "验证签名成功！"
            Else
110             strTmp = "验证签名失败！"
            End If
112         If blnRet And strTStampCode <> "" Then
114             lngRuslt = BJCAGX_TS.verifyTimeStamp(strTStampCode)
116             If lngRuslt <> 0 Then
118                 strTmp = "验证时间戳失败！" & GetReturnInfo(lngRuslt)
120                 blnRet = False
                End If
            End If
        Else
            '验证时间戳
122         If strTStampCode <> "" Then
124             strTmp = ""
                On Error Resume Next
126             lngRuslt = BJCAGX_TS.verifyTimeStamp(strTStampCode): bytVer = 0
128             If lngRuslt <> 0 Then strTmp = "【verifyTimeStamp】验证时间戳失败！" & GetReturnInfo(lngRuslt): blnRet = False
130             If Err.Number <> 0 Or lngRuslt <> 0 Then
132                 lngRuslt = BJCAGX_TS.VerifyTS(strTStampCode, "")  '不验证源文
134                 bytVer = 1
                End If
136             Err.Clear: On Error GoTo errH
138             If bytVer = 1 Then
140                 If lngRuslt <> 0 Then
142                     strTmp = strTmp & vbCrLf & _
                            "【VerifyTS】验证时间戳失败！" & GetReturnInfo(lngRuslt)
144                         blnRet = False
                    Else
146                     strDate = BJCAGX_TS.gettimestampinfo(strTStampCode, 1)  '取得时间戳时间
148                     strTimeStamp = String14ToDate(strDate)
150                     strTmp = "验证时间戳成功！" & vbTab & "签名时间:" & strTimeStamp
152                     blnRet = True
                    End If
                End If
            End If

            '验证签名
154         intRet = BJCAGX_svs.VerifySignatureBySN(strCertSN, strSource, strSignData)
156         If (intRet = 0) Then
158             strTmp = IIf(strTmp <> "", strTmp & vbCrLf, "") & "数据签名验证成功！"
160             blnRet = True And blnRet
            Else
162             strTmp = IIf(strTmp <> "", strTmp & vbCrLf, "") & "数据签名验证失败！"
164             blnRet = False
            End If
        End If
        
166     If strTmp <> "" Then
168         MsgBoxEx strTmp, vbOKOnly + vbInformation, gstrSysName
        End If
        
170     BJCAGX_VerifySign = blnRet
        Exit Function
errH:
172     MsgBoxEx "验证签名失败！" & vbNewLine & "第" & CStr(Erl()) & "行 " & Err.Description, vbQuestion, gstrSysName
End Function

'销毁对象
Public Sub BJCAGX_UloadObj()
    Set BJCAGX_Client = Nothing
    Set BJCAGX_svs = Nothing
    Set BJCAGX_TS = Nothing
    Set mobjPic = Nothing
    mblnInit = False
End Sub

'----- 以下是内部函数
Private Function GetCertLogin(ByVal strCertSN As String, ByVal strCert As String, Optional ByVal strCertID As String) As Boolean
        Dim random As String
        Dim serverCert As String
        Dim serverSign As String, strSignVal As String
        Dim blnRet As Boolean
        Dim strDate As String
        Dim intDay As Integer, intRetSign As Integer, intRetVal As Integer
        Dim strTmp As String
        Dim retValidateCert As Long
    
        On Error GoTo errH
    
100     If gudtPara.bytSignVersion = V_RSA Then
102         If BJCAGX_svs Is Nothing Then Set BJCAGX_svs = CreateObject("BJCA_SVS_ClientCOM.BJCASVSEngine.1")
104         random = BJCAGX_Client.GenRandom(24)
106         strSignVal = BJCAGX_Client.signedData(strCertSN, random)
            '证书安全登录
            'strSignVal:非空:成功
            'strSignVal:空:不成功
108         If (strSignVal <> "") Then
                '服务器端验证证书
                '从组件中导出证书
110             retValidateCert = ValidateCert(strCert)
            
                '验证证书结果信息表示
112             If retValidateCert <> 0 Then Call ValidateCertView(retValidateCert)
    
114             If (retValidateCert = 0) Then
                    Dim s As String
                    '获取客户端证书有效期截止时间
116                 s = BJCAGX_Client.GetUserInfo(strCert, 12)
                    '验证客户端证书有效期剩余天数
118                 intDay = CheckValidaty(s)
            
120                 If (intDay <= 30 And intDay > 0) And Not gblnShow Then
122                     MsgBoxEx "您的证书还有" & intDay & "天过期"
124                     gblnShow = True '不再提示
126                     GetCertLogin = True
128                 ElseIf (intDay <= 0) Then
130                     MsgBoxEx "您的证书已过期 " & Abs(intDay) & " 天"
132                     GetCertLogin = False
                    Else
134                     GetCertLogin = True
                    End If
                Else
136                 GetCertLogin = False
                End If
            Else
138             GetCertLogin = False
            End If
140     ElseIf gudtPara.bytSignVersion = V_SM2 Then
            '1)（BJCA_SVS_ClientCOM组件）HIS系统调用CA接口，获取随机数、服务器证书，并通过服务器证书对随机数进行签名；
142         random = BJCAGX_svs.GenRandom(16) '获取随机数
144         serverCert = BJCAGX_svs.GetServerCertificate() '获取服务器证书
146         serverSign = BJCAGX_svs.SignData(random) '服务端对随机数签名
        
148         blnRet = BJCAGX_Client.SOF_VerifySignedData(serverCert, random, serverSign) '客户端验证服务端签名
150         If Not blnRet Then
152             MsgBoxEx "服务端签名验证失败！", vbOKOnly + vbInformation, gstrSysName
                Exit Function
            End If
            '验证证书是否过期
154         strDate = BJCAGX_Client.SOF_GetCertInfo(strCert, 12)
156         strDate = String14ToDate(strDate)
158         If strDate <> "" Then
            '验证客户端证书有效期剩余天数
160             intDay = CheckValidaty(CDate(strDate))
        
162             If (intDay <= 30 And intDay > 0 And Not gblnShow) Then
164                 MsgBoxEx "您的证书还有" & intDay & "天过期。", vbOKOnly + vbInformation, gstrSysName
166                 gblnShow = True
168             ElseIf (intDay <= 0) Then
170                 MsgBoxEx "您的证书已过期 " & Abs(intDay) & " 天。", vbOKOnly + vbInformation, gstrSysName
                    Exit Function
                End If
            End If
            '验证证书是否过期
172         strSignVal = BJCAGX_Client.SOF_SignData(strCertID, random)  '客户端随机数签名
174         intRetSign = BJCAGX_svs.VerifySignedData(strCert, random, strSignVal)   '服务端验证客户端签名
176         intRetVal = BJCAGX_svs.ValidateAndSaveCertificate(strCert)  '服务端验证客户端证书有效性并保存证书
        
178         If Not (intRetSign = 0 And (intRetVal = 0 Or intRetVal = 1)) Then
180             MsgBoxEx "客户端证书验失败！", vbOKOnly + vbInformation, gstrSysName
                Exit Function
            End If
    
182         GetCertLogin = True
        End If
    
        Exit Function
errH:
184     MsgBoxEx "登录验证失败！" & vbNewLine & "第" & CStr(Erl()) & "行 " & Err.Description, vbInformation, gstrSysName
End Function

Private Function ValidateCert(ByRef userCert As String) As Integer
    '服务器端验证证书
    If BJCAGX_svs Is Nothing Then Set BJCAGX_svs = CreateObject("BJCA_SVS_ClientCOM.BJCASVSEngine.1")
    ValidateCert = BJCAGX_svs.ValidateCertificate(userCert)
 
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
        Case -5
            MsgBoxEx "证书未生效！"
    End Select
End Sub

''' 获取客户端证书列表
''' 返回boolean
Private Function GetCertList(ByRef strName As String, ByRef strCertSN As String, ByRef strCert As String, Optional ByRef strFilePath As String = "0", _
        Optional ByRef strCertDN As String = "0", Optional ByRef strCertUserID As String = "0", Optional ByRef strCertID As String = "0", _
        Optional ByRef strPicCode As String = "0") As Boolean
        '北京CA广西版获取数字证书列表函数
        '-入参:无
        '-出参
        'strName :      保存接口返回的证书所有者姓名
        'strCertSn:   保存接口返回的证书所有者唯一标识
        'strCert:       保存接口返回的签名证书
        Dim strUsbkeyList As String
        Dim arrUserListLength As Integer
        Dim arrUserList() As String
        Dim strPic As String, strUser As String
        Dim strPicID As String
    
        On Error GoTo errH
    
100     If gudtPara.bytSignVersion = V_RSA Then

102    strUsbkeyList = BJCAGX_Client.getUserList()
104    arrUserList = Split(strUsbkeyList, "&&&")
106    arrUserListLength = UBound(arrUserList)
108    If (arrUserListLength = -1) Then
110        MsgBoxEx "请您插入Key！", vbOKOnly + vbInformation, gstrSysName
           Exit Function
       End If
112    If (arrUserListLength <> 0) Then
           Dim i As Integer
114        For i = 0 To arrUserListLength - 1
               Dim strOption As String
116            strOption = arrUserList(i)
118            strName = Split(strOption, "||")(0)
120            strCertSN = Split(strOption, "||")(1)
122            strCert = BJCAGX_Client.ExportUserCert(strCertSN)
124            If strCertDN <> "0" Then strCertDN = BJCAGX_Client.GetUserInfo(strCert, 20)
           Next
       End If

126     ElseIf gudtPara.bytSignVersion = V_SM2 Then

128        strUsbkeyList = BJCAGX_Client.SOF_GetUserList()
130    If (strUsbkeyList = "") Then
132        strName = ""
134        MsgBoxEx "请您插入Key！", vbOKOnly + vbInformation, gstrSysName
136        GetCertList = False
           Exit Function
       Else
138            arrUserList = Split(strUsbkeyList, "&&&") 'sm2测试2||216000000279373/1003201510002131&&&sm2测试1||216000000279349/1003201510002370&&&
140        If UBound(arrUserList) > 1 Then  '多个KEY
142            For i = LBound(arrUserList) To UBound(arrUserList) - 1
144                strUser = strUser & "&&&" & Split(arrUserList(i), "||")(0)
               Next
146            If strUser <> "" Then strUser = Mid(strUser, 4)
148            strName = frmSelectUser.ShowMe(strUser)
                
150            For i = LBound(arrUserList) To UBound(arrUserList) - 1
152               If strName = Split(arrUserList(i), "||")(0) Then
154                    strCertID = Split(arrUserList(i), "||")(1)
                       Exit For
                  End If
               Next
           Else
156            arrUserList = Split(arrUserList(0), "||")
158            strName = arrUserList(0)      '证书CN通用名
160            strCertID = arrUserList(1)    '证书ID
           End If

           '获取图片容器名
162        strPicID = Split(strCertID, "/")(0)
    
164            strCert = BJCAGX_Client.SOF_ExportUserCert(strCertID) '3.导出签名证书。
166            If strCertSN <> "0" Then strCertSN = BJCAGX_Client.SOF_GetCertInfo(strCert, 2) '证书序列号 签名时要用
168            If strCertDN <> "0" Then strCertDN = BJCAGX_Client.SOF_GetCertInfo(strCert, 33) '证书DN
170        If strCertUserID <> "0" Then
172                strCertUserID = BJCAGX_Client.SOF_GetCertInfoByOid(strCert, "2.16.840.1.113732.2") '2.获取证书唯一标识（一般为身份证号）SF+身份证号
174                strCertUserID = Mid(strCertUserID, 3)
           End If
       End If
        End If
176     If gudtPara.blnSignPic Then
178         If strFilePath <> "0" Or strPicCode <> "0" Then
180            strPic = mobjPic.GetPic(strPicID)
182            If strPic <> "" Then
184                 strPicCode = mobjPic.ConvertPicFormat(strPic, 5)
186                 If strPicCode <> "" And strFilePath <> "0" Then
188                     strFilePath = SaveBase64ToFile("BMP", strCertSN, strPicCode)
                    Else
190                     strFilePath = ""
                    End If
                End If
            End If
        End If
192     GetCertList = True
    
        Exit Function
errH:
194 MsgBoxEx "读取证书信息失败！" & vbNewLine & "第" & CStr(Erl()) & "行 " & Err.Description, vbQuestion, gstrSysName
End Function

''' 检查证书有效性
''' 返回证书有效期天数
Private Function CheckValidaty(ByVal endDate As Date) As Integer
    '北京CA广西版检查证书有效性接口
    '-入参: 证书有效截止日期
    '-出参：有效天数
        Dim dblAllSp    As Double
        Dim result      As Integer
        dblAllSp = CDbl(CDate(endDate)) - CDbl(Now)
        result = Int(dblAllSp)
        CheckValidaty = result
End Function

Private Function GetReturnInfo(ByVal strSign As Long) As String
    '准格尔时间戳返回信息转换函数
    If strSign = -1 Then
        GetReturnInfo = "时间戳验证不通过"
    ElseIf strSign = -2 Then
        GetReturnInfo = "原文验证不通过"
    ElseIf strSign = -3 Then
        GetReturnInfo = "不是所信任的根"
    ElseIf strSign = -4 Then
        GetReturnInfo = "证书未生效"
    ElseIf strSign = -5 Then
        GetReturnInfo = "查询不到此证书"
    ElseIf strSign = -6 Then
        GetReturnInfo = "签发时间戳时服务器证书过期"
    ElseIf strSign = 0 Then
        GetReturnInfo = "验证成功"
    Else
        GetReturnInfo = "未知错误" & "错误码:" & strSign
    End If
    If GetReturnInfo <> "" Then
        GetReturnInfo = "时间戳接口返回提示：" & GetReturnInfo
    End If
End Function

Private Function GetTimeStamp(ByVal strData As String) As String
    Dim year As String, mouth As String, day As String, hour As String, mm As String, ss As String
    Dim strTimeStamp As String

    '获取时间戳
    If Len(strData) = 14 Then
        year = Mid(strData, 1, 4)
        mouth = Mid(strData, 5, 2)
        day = Mid(strData, 7, 2)
        hour = Mid(strData, 9, 2)
        mm = Mid(strData, 11, 2)
        ss = Mid(strData, 13, 2)
        strTimeStamp = year & "-" & mouth & "-" & day & " " & hour & ":" & mm & ":" & ss
        If Not IsDate(strTimeStamp) Then
            MsgBoxEx "获取的时间戳不是一个日期！" & strTimeStamp, vbExclamation, gstrSysName
            GetTimeStamp = ""
            Exit Function
        End If
    End If
    GetTimeStamp = strTimeStamp
End Function

Public Function BJCAGX_GetPara(Optional ByVal bytFunc As Byte) As Boolean
    Dim arrList As Variant
    
    On Error GoTo errH
    gstrPara = gobjComLib.zlDatabase.GetPara(90000, glngSys)  '读取URLs 固定读取ZLHIS'gstrPara = "0&&&0"   '0-RSA;1-SM2&&&0-不启用签章图片;1-启用签章
    arrList = Split(gstrPara, G_STR_SPLIT)
    If bytFunc = 1 Then
        If gstrPara = "" Or UBound(arrList) < 1 Then
            Err.Raise -1, , "当前系统【" & glngSys & "】没有配置电子签名参数，请先配置。"
            Exit Function
        End If
    End If
    
    If UBound(arrList) = 0 Then
        gudtPara.bytSignVersion = V_RSA
        gudtPara.blnSignPic = False
        gudtPara.strSignURL = "|"
    ElseIf UBound(arrList) = 1 Then
        gudtPara.bytSignVersion = Val(arrList(0))
        gudtPara.blnSignPic = Val(arrList(1)) = 1
        gudtPara.strSignURL = "|"
    ElseIf UBound(arrList) = 2 Then
        gudtPara.bytSignVersion = Val(arrList(0))
        gudtPara.blnSignPic = Val(arrList(1)) = 1
        gudtPara.strSignURL = arrList(2) '以|分隔方式存放手签上传URL和获取URL
    End If

    BJCAGX_GetPara = True
    Exit Function
errH:
    MsgBoxEx "读取参数失败！" & vbNewLine & "第" & CStr(Erl()) & "行 " & Err.Description, vbInformation, gstrSysName
End Function

Public Function BJCAGX_SetParaStr() As String
    BJCAGX_SetParaStr = IIf(gudtPara.bytSignVersion = 0, "0", "1") & G_STR_SPLIT & IIf(gudtPara.blnSignPic, "1", "0") & G_STR_SPLIT & IIf(gudtPara.strSignURL = "", "|", gudtPara.strSignURL)
End Function



