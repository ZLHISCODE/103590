Attribute VB_Name = "mdlHBCA"
Option Explicit

'河北邯郸市第三医院   河北CA
'2018-02-26唐山南湖医院调整时间戳证书存储规则:
'CA时间戳服务器是负载均衡的，两台，两个证书;签名时获取的可能是这一台，也可能是另外一台
'所以要改下时间戳证书存储模式，两种建议，一种是新增一个时间戳证书信息表，根据签名时获取到的时间戳证书信息去数据库查，这种不会冗余;
'第二种是把每次签名时获取到的时间戳信息存到签名信息表(此种方式要求签名信息值+分隔符("[;]")+时间戳证书内容长度小于4000个字符)

Private mblnInit As Boolean         '是否已初始化成功
Private mCertMgr As Object          'HebcaP11XLib.certMgr
Private mSignCert As Object         'HebcaP11XLib.cert
Private mFormSeal As Object         'FormSealCtrl1 电子签章控件
Private mSVSClient As Object        'SVS_SOFT_COMLib.SvsVerify '定义并实例化SVS客户端组件
Private mblnTs As Boolean           '是否启用时间戳

Private Const M_STR_LICENCE As String = "amViY55oZWKcZmhlnWxhaGViY2GXGmJjYWhlYnGH1QQ5GcNqnW6z3vohVnE+nTJr"
Private Const M_STR_SUMMARY As String = "[SUMMARY]"
Private Const M_STR_SPLIT As String = "[;]"

Public Function HBCA_InitObject() As Boolean
    '----------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化电子签名需要用到对象
    '返回:True-初始化成功;False-初始化失败
    '编制:余伟节
    '日期:2015-08-31
    '----------------------------------------------------------------------------------------------------------------------------------
      Dim arrList As Variant
    
100   If mblnInit Then HBCA_InitObject = True: Exit Function
      On Error GoTo errH
      '参数信息:IP|端口号|是否启用时间戳(0-不启用/1-启用)
102   gstrPara = gobjComLib.zlDatabase.GetPara(90000, glngSys, , "121.28.49.158&&&5000&&&0")  '读取配置内容
104   If gstrPara = "" Then
106       Err.Raise -1, , "配置文件读取失败，请到启用电子签名接口处设置。"
          Exit Function
      End If
108   arrList = Split(gstrPara, "&&&")
110   If UBound(arrList) <> 2 Then
112       MsgBoxEx "签名服务器地址配置格式有误,请到启用电子签名接口处设置。", vbOKOnly + vbInformation, gstrSysName
          Exit Function
      End If
114  mblnTs = (Val(CStr(Split(gstrPara, G_STR_SPLIT)(2))) = 1)
116 If mCertMgr Is Nothing Then Set mCertMgr = CreateObject("HebcaP11X.CertMgr.1")
 
118 If mSignCert Is Nothing Then Set mSignCert = CreateObject("HebcaP11X.Cert.1")
120 If mFormSeal Is Nothing Then Set mFormSeal = CreateObject("HebcaFormSeal.FormSealCtrl.1")
122 If mSVSClient Is Nothing Then Set mSVSClient = CreateObject("Svs_soft_com.SvsVerify.1")
124   mCertMgr.Licence = M_STR_LICENCE
126   gstrLogins = ""
128   mblnInit = True
130   HBCA_InitObject = True
      Exit Function
errH:
132 MsgBoxEx "创建河北CA接口部件失败！" & vbNewLine & "第" & CStr(Erl()) & "行 " & vbNewLine & _
              Err.Description, vbQuestion, gstrSysName
End Function

Public Function HBCA_RegCert(arrCertInfo As Variant) As Boolean
'----------------------------------------------------------------------------------------------------------------------------------
'功能:提供在HIS数据库中注册个人证书的必要信息,用于新发放或更换证书,需要插入USB-Key
'返回：arrCertInfo作为数组返回证书相关信息
'      0-ClientSignCertCN:客户端签名证书公共名称(姓名),用于注册证书时程序验证身份
'      1-ClientSignCertDN:客户端签名证书主题(每人唯一)
'      2-ClientSignCertSN:客户端签名证书序列号(每证书唯一)
'      3-ClientSignCert:客户端签名证书内容
'      4-ClientEncCert:客户端加密证书内容
'      5-签名图片文件名,空串表示没有签名图片
'      6-时间戳证书
'      7-签章信息
'编制:余伟节
'日期:2015-08-31
'----------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    Dim strCertName As String
    Dim strCertDN As String, strUserID As String, strSealB64 As String
    Dim strCertSn As String, strTSCert As String
    Dim strSignCert As String, strPic As String
    Dim strEncCert As String
    
    On Error GoTo errH
    
    If mFormSeal Is Nothing Then Set mFormSeal = CreateObject("HebcaFormSeal.FormSealCtrl.1")
    
100     For i = LBound(arrCertInfo) To UBound(arrCertInfo)
102         arrCertInfo(i) = ""
        Next

104     If GetCertList(strCertName, strCertSn, strSignCert, strUserID, strSealB64, strTSCert, strPic) Then
106         arrCertInfo(0) = strCertName
108         arrCertInfo(1) = strCertDN
110         arrCertInfo(2) = strCertSn
112         arrCertInfo(3) = strSignCert
113         arrCertInfo(4) = strEncCert
            arrCertInfo(5) = strPic
            arrCertInfo(6) = strTSCert
            arrCertInfo(7) = strSealB64
114
124         HBCA_RegCert = True
        End If

        Exit Function
errH:
126     MsgBoxEx "获取证书信息失败！" & vbNewLine & "第" & CStr(Erl()) & "行 " & Err.Description, vbQuestion, gstrSysName
   
        
End Function

Private Function GetCertList(ByRef strName As String, ByRef strCertSn As String, ByRef strSignCert As String, _
                ByRef strUserID As String, Optional ByRef strSealBase64 As String, Optional ByRef strTSCert As String, _
                Optional ByRef strPicFile As String = "0", Optional ByRef strPic As String) As Boolean
    '----------------------------------------------------------------------------------------------------------------------------------
    '功能:获取证书信息
    '参数:strName-证书用户名
    '     strCertSn-证书序列号
    '     strSignCert-证书内容
    '     strUserID-用户唯一标识  截取身份证号
    '     strSealBase64-签章BASE64
    '     strTSCert-时间戳证书BASE64
    '返回：
    '编制:余伟节
    '日期:2015-08-31
    '----------------------------------------------------------------------------------------------------------------------------------
        Dim intCount As Integer
        Dim strSignData As String
        Dim strSealSn As String
        Dim objPic As Picture
        
        On Error GoTo errH
        mCertMgr.Licence = M_STR_LICENCE
100     intCount = mCertMgr.GetDeviceCount
102      If intCount < 1 Then
104          MsgBoxEx "未发现KEY,请您插入Key！", vbInformation, gstrSysName
             Exit Function
         End If
106     Set mSignCert = mCertMgr.SelectSignCert
108     intCount = mFormSeal.GetSealCount()
110     If intCount < 1 Then
112          MsgBoxEx "当前设备没有签章！", vbInformation, gstrSysName
             Exit Function
         End If
         'CN=持有者姓名
114     strName = mSignCert.GetSubjectItem("cn")
         'strSignCert = mSignCert.GetCertB64        '得到签名证书内容
     '        strCertDN = mSignCert.GetSubjectItem("DN")
        '获取数字证书的唯一标识,用于和用户建立绑定
116     strUserID = mSignCert.GetCertExtensionByOid("1.2.156.112586.1.4") '"2@6021SF0130637201507090001"
118     strUserID = Mid(strUserID, 10) '获取身份证号
        '获取证书信息  验证签名需要证书信息：
        '签章BASE64编码,签章证书,时间戳证书BASE64编码
120     strSignData = mFormSeal.SignAndSealWithoutTimeStampCert("测试20150901", "", 0, True, mblnTs)
121     If strSignData = "" Then MsgBoxEx "随机数签名失败！", vbInformation, gstrSysName: Exit Function
        '获取签章的SN\签章Bases64
122     strSealSn = mFormSeal.GetSelectedSeal()
124     strSealBase64 = mFormSeal.GetSeal(strSealSn) '获取章的B64
126     strPic = mFormSeal.GetSealPicFromB64(strSealBase64)
128     If strPicFile <> "0" Then
130            strPicFile = SaveBase64ToFile("gif", strSealSn, strPic)
        End If
         '将图片转换成指定bmp格式
    '        Set objPic = LoadPicture(strPicFile)
    '        SavePicture objPic, strPicFile
     '
         '获取证书各项信息，证书的SN和证书的有效期需要存入数据库
132     strSignCert = mFormSeal.GetCert(strSealSn) '获取证书
134     Set mSignCert = mCertMgr.CreateCertFromB64(strSignCert)
136     strCertSn = mSignCert.GetSerialNumber '证书SN
         'dCertDate:=mSignCert.NotAfter    有效期
          '时间戳
138     If mblnTs Then
140         strTSCert = mFormSeal.GetTimeStampCert '获取时间戳证书内容
         End If

        
142     GetCertList = True
        Exit Function
errH:
144     MsgBoxEx "获取证书信息失败！" & vbNewLine & "第" & CStr(Erl()) & "行 " & Err.Description, vbQuestion, gstrSysName
End Function

Public Function HBCA_Sign( _
    ByVal strSource As String, ByRef strSignData As String, ByRef strTimeStamp As String, _
    ByRef strTimeStampCode As String, ByRef blnReDo As Boolean) As Boolean
    '----------------------------------------------------------------------------------------------------------------------------------
    '功能:签名
    '参数:strSource-数据源
    '     strSignData-签名值
    '     strTimeStamp-时间戳值
    '     strTimeStampCode-时间戳信息
    '返回：True/false
    '编制:余伟节
    '日期:2015-08-31
    '----------------------------------------------------------------------------------------------------------------------------------
        '签名
        Dim strTiemRequest As String
        Dim strTmp As String
        Dim strSealSn As String
        Dim strSealBase64 As String
        Dim strTSCert As String
        Dim blnCheck As Boolean
        
        On Error GoTo errH
100     blnCheck = HBCA_CheckCert(blnReDo)
102     If blnReDo Then Exit Function
104     If blnCheck Then                '验证当前USB是否是签名用户的，并获取签名证书
            '调用SignAndSealWithoutTimeStampCert对原文进行盖章，可以对原文数据进行组织，
106         strSource = mCertMgr.util.HashText(strSource, 1)
108         strSignData = mFormSeal.SignAndSealWithoutTimeStampCert(strSource, "", 0, True, mblnTs)
109         If strSignData = "" Then MsgBoxEx "签名失败,签名值为空！", vbInformation, gstrSysName: Exit Function
110         If mblnTs Then
112             strTimeStampCode = mFormSeal.GetTimeStamp() '获取时间戳信息
114             strTimeStamp = mFormSeal.GetTimeStampInfoByB64(strTimeStampCode, "time")
116             strTSCert = mFormSeal.GetTimeStampCert '获取时间戳证书内容
118             If strTimeStampCode = "" Then
120                 MsgBoxEx "签名失败,时间戳B64为空！", vbInformation, gstrSysName
                    Exit Function
122             ElseIf strTSCert = "" Then
124                 MsgBoxEx "签名失败,时间戳证书内容为空！", vbInformation, gstrSysName
                    Exit Function
                End If
            Else
126             strTimeStamp = CStr(gobjComLib.zlDatabase.Currentdate)
            End If
        Else
128         MsgBoxEx "签名失败！", vbInformation, gstrSysName
            Exit Function
        End If
130     If Trim(strSignData) <> "" Then strSignData = M_STR_SUMMARY & strSignData  '此标识[SUMMARY]用于验证签名时区分按原文验证签名还是按摘要验证签名
132     If strTSCert <> "" Then strSignData = strSignData & M_STR_SPLIT & strTSCert     'M_STR_SPLIT 分隔符[;]
134     HBCA_Sign = True
        Exit Function
errH:
136     MsgBoxEx "签名失败！" & vbNewLine & "第" & CStr(Erl()) & "行 " & Err.Description, vbQuestion, gstrSysName
End Function

Public Function HBCA_VerifySign(ByVal strSignData As String, ByVal strSource As String, ByVal strTimeStampCode As String, _
                            ByVal strCert As String, ByVal strTSCert As String, ByVal strSealCert As String) As Boolean
    '----------------------------------------------------------------------------------------------------------------------------------
    '功能:验证签名
    '参数:
    '   strSignData-签名值
    '   strSource-源文
    '   strTimeStampCode-时间戳信息
    '   strCert-证书内容
    '   strTSCert-时间戳证书内容
    '   strSealCert-签章证书内容
    '返回：True/false
    '编制:余伟节
    '日期:2015-08-31
    '----------------------------------------------------------------------------------------------------------------------------------
      Dim strTmp As String
      Dim varTmp As Variant
      Dim lngRet As Long
      Dim blnOk As Boolean
      On Error GoTo errH
    
100   If mFormSeal Is Nothing Then Set mFormSeal = CreateObject("HebcaFormSeal.FormSealCtrl.1")
102   strTmp = ""

      '电子签章验证
104   If UCase(left(strSignData, Len(M_STR_SUMMARY))) = M_STR_SUMMARY Then
          '按摘要签名时验证签名时需要取摘要
106       strSignData = Mid(strSignData, Len(M_STR_SUMMARY) + 1)
107       mCertMgr.Licence = M_STR_LICENCE
108       strSource = mCertMgr.util.HashText(strSource, 1)
          '从签名信息中解析时间戳证书 M_STR_SPLIT 分隔符[;]
110       varTmp = Split(strSignData, M_STR_SPLIT)
112       If UBound(varTmp) = 1 Then
114           strSignData = varTmp(0)
116           strTSCert = varTmp(1)
          End If
      End If

118    Call mFormSeal.VerifyAndShowSeal(strSealCert, strCert, strSource, 1, strSignData, IIf(mblnTs, 0, -1), strTimeStampCode, strTSCert, 0)
120    lngRet = mFormSeal.GetVerifyResult()

122   If lngRet = 0 Then
124       strTmp = "签章验证成功！"
126       blnOk = True
      Else
128       strTmp = "签章验证失败！"
130       blnOk = False
      End If

132   If strTmp <> "" Then
134       MsgBoxEx strTmp, vbOKOnly + vbInformation, gstrSysName
      End If
    
136    HBCA_VerifySign = blnOk
      Exit Function
errH:
138     MsgBoxEx "验证签名失败！" & vbNewLine & "第" & CStr(Erl()) & "行 " & Err.Description, vbInformation, gstrSysName
End Function

Public Function HBCA_CheckCert(ByRef blnReDo As Boolean) As Boolean
    '----------------------------------------------------------------------------------------------------------------------------------
    '功能：读取USB进行设备初始化并登录
    '参数:
    '   出参:blnRedo-True：重新注册证书成功,False-未重新注册证书
    '返回：True/false
    '编制:余伟节
    '日期:2015-08-31
    '----------------------------------------------------------------------------------------------------------------------------------
        Dim strCertUserID As String, strPIN As String, strUserName As String
        Dim strWebUrl As String, intDate   As Integer
        Dim strCertName As String, strCertSn As String, strCert As String, strCertDN As String
        Dim strTSCert As String, strSealCode As String, strPic As String, strDate As String
        Dim blnOk As Boolean
        Dim udtUser As USER_INFO
        
        On Error GoTo errH
100     If Not mblnInit Then
102         Call HBCA_InitObject
104         If Not mblnInit Then
106             MsgBoxEx "部件未初始化！"
                Exit Function
            End If
        End If
         '获取证书信息同时检查Key盘是否插入
108     If Not GetCertList(strCertName, strCertSn, strCert, strCertUserID, strSealCode, strTSCert, , strPic) Then
110         HBCA_CheckCert = False: Exit Function
        End If
        '未注册在当前用户名下的Key
112     If mUserInfo.strUserID = "" Then
114         MsgBoxEx "您的身份证号为空,请联系管理员到人员管理中录入！", vbOKOnly + vbInformation, gstrSysName
            Exit Function
116     ElseIf strCertUserID <> mUserInfo.strUserID Then
118         MsgBoxEx "您的身份证号：" & _
                   vbCrLf & vbTab & "【" & mUserInfo.strUserID & "】" & vbCrLf & _
                   "当前证书唯一标识:" & _
                   vbCrLf & vbTab & "【" & strCertUserID & "】" & vbCrLf & _
                   "用户身份证号与当前证书唯一标识不相等,不能使用！", vbInformation, gstrSysName
            Exit Function
        End If
        
        '登录验证
120     If InStr(gstrLogins & "|", "|" & strCertSn & "|") > 0 Then '首次验证通过后，下次不在继续验证
122         blnOk = True
        Else
124         If Not GetCertLogin() Then
126             blnOk = False
            Else
128             blnOk = True
130             If InStr(gstrLogins & "|", "|" & strCertSn & "|") = 0 Then gstrLogins = gstrLogins & "|" & strCertSn
            End If
        End If
        
132     If blnOk Then
            '判断是否需要更新注册证书
134         udtUser.strName = strCertName
136         udtUser.strSignName = strCertName
138         udtUser.strUserID = strCertUserID
140         udtUser.strCertSn = strCertSn
142         udtUser.strCertDN = strCertDN
144         udtUser.strCert = strCert
146         udtUser.strEncCert = ""
148         udtUser.strCertID = ""
150         udtUser.strPicCode = strPic
152         udtUser.strTSCert = strTSCert
154         udtUser.strSealCode = strSealCode
            '获取已经注册证书的有效结束日期
                '获取证书各项信息，证书的SN和证书的有效期需要存入数据库
156         Set mSignCert = mCertMgr.CreateCertFromB64(mUserInfo.strCert)
158         strDate = mSignCert.NotAfter
160         If IsUpdateRegCert(udtUser, strDate, blnReDo) Then
162             HBCA_CheckCert = True
            Else
164             HBCA_CheckCert = False
            End If
        End If
    
     
    
        Exit Function
errH:
166     MsgBoxEx "检查USBKEY失败！" & vbNewLine & "第" & CStr(Erl()) & "行 " & Err.Description, vbQuestion, gstrSysName
End Function

Private Function GetCertLogin() As Boolean
    '----------------------------------------------------------------------------------------------------------------------------------
    '功能：登录验证
    '参数:
    '返回：True/false
    '编制:余伟节
    '日期:2015-08-31
    '----------------------------------------------------------------------------------------------------------------------------------
        Dim strText As String
        Dim strMsg As String
        Dim lngRetVal As Long
        Dim strSignData As String
        Dim strCertB64 As String
        Dim strSealSn As String
        Dim strDate As String
        Dim intDay As Integer
        Dim strIP As String, lngPort As Long
        Dim arrTmp As Variant
    
        On Error GoTo errH
100     If mCertMgr Is Nothing Then Set mCertMgr = CreateObject("HebcaP11X.CertMgr.1") '实例化P11组件的CertMgr类
102     If mSVSClient Is Nothing Then Set mSVSClient = CreateObject("Svs_soft_com.SvsVerify.1")
104     If mFormSeal Is Nothing Then Set mFormSeal = CreateObject("HebcaFormSeal.FormSealCtrl.1")
106     strText = "hebca2013" '原始字符串

108     mCertMgr.Licence = M_STR_LICENCE

110     Set mSignCert = mCertMgr.SelectSignCert  '得到签名证书对象
112     strSignData = mSignCert.SignText(strText, 1)   '进行数字签名,将签名值存放到signdata
114     strCertB64 = mSignCert.GetCertB64         '得到签名证书内容
        'gstrPara = "121.28.49.158&&&5000"
116     arrTmp = Split(gstrPara, G_STR_SPLIT)     'IP&&&端口号 "121.28.49.158", 5000
118     strIP = arrTmp(0): lngPort = Val(CStr(arrTmp(1)))
120     lngRetVal = mSVSClient.InitialVerify(strIP, lngPort) '初始化SVS客户端

        Dim r As Boolean
    
122     If lngRetVal < 0 Then
124         MsgBoxEx "无法连接SVS服务器!", vbInformation, gstrSysName
            Exit Function
        End If
    
126     lngRetVal = mSVSClient.VerifyCertSign(-1, 0, strText, Len(strText), strCertB64, strSignData, 1, lngRetVal)     '验证
128     Select Case lngRetVal
            Case 0
130             strMsg = "验证成功"
132         Case 1
134             strMsg = "您的证书未生效!"
136         Case 2
138             strMsg = "您的证书已经过期!"
140         Case 4
142             strMsg = "您的证书非河北CA颁发!"
144         Case 1002
146             strMsg = "您的证书非河北CA颁发!"
148         Case 7
150             strMsg = "您的证书已经被吊销!"
152         Case -6406
154             strMsg = "签名验证失败,请重试!"
156         Case Else
158             strMsg = "签名验证失败!错误码:" & lngRetVal
        End Select
160     If strMsg <> "验证成功" Then
162         MsgBoxEx "错误信息:" & strMsg, vbInformation, gstrSysName
            Exit Function
        End If
    '
    '     strSignData = mFormSeal.SignAndSealWithoutTimeStampCert("测试20150901", "", 0, True, True)
    '
    '     '获取签章的SN\签章Bases64
    '     strSealSn = mFormSeal.GetSelectedSeal()
    '     strSignCert = mFormSeal.GetCert(strSealSn) '获取证书
    '     Set mSignCert = mCertMgr.CreateCertFromB64(strSignCert)
164      strDate = mSignCert.NotAfter  '有效期
166     If strDate <> "" Then
        '验证客户端证书有效期剩余天数
168         intDay = CheckValidaty(CDate(strDate))
    
170         If (intDay <= 30 And intDay > 0 And Not gblnShow) Then
172             MsgBoxEx "您的证书还有" & intDay & "天过期。", vbOKOnly + vbInformation, gstrSysName
174             gblnShow = True
176         ElseIf (intDay <= 0) Then
178             MsgBoxEx "您的证书已过期 " & Abs(intDay) & " 天。", vbOKOnly + vbInformation, gstrSysName
                Exit Function
            End If
        End If
    
180     GetCertLogin = True
        Exit Function
errH:
182     MsgBoxEx "登录验证失败！" & vbNewLine & "第" & CStr(Erl()) & "行 " & Err.Description, vbInformation, gstrSysName

End Function

Public Function HBCA_GetPara() As Boolean
    '设置服务器地址
        Dim arrList As Variant
    
        On Error GoTo errH
100     gstrPara = gobjComLib.zlDatabase.GetPara(90000, glngSys)
102     If gstrPara = "" Then gstrPara = "121.28.49.158&&&5000&&&0" '参数信息:IP&&&端口号&&&是否启用时间戳(0-不启用/1-启用)
104     If gstrPara <> "" Then
106         arrList = Split(gstrPara, "&&&")
108         If UBound(arrList) = 2 Then
110              gudtPara.strSIGNIP = Trim(arrList(0))
112              gudtPara.strSignPort = Trim(arrList(1))
114              gudtPara.blnISTS = (Val(arrList(2)) = 1)
            End If
        End If
        Exit Function
errH:
116     MsgBoxEx "读取参数失败！" & vbNewLine & "第" & CStr(Erl()) & "行 " & Err.Description, vbInformation, gstrSysName
End Function

Public Function HBCA_SetParaStr() As String
    HBCA_SetParaStr = gudtPara.strSIGNIP & "&&&" & gudtPara.strSignPort & "&&&" & IIf(gudtPara.blnISTS, "1", "0")
End Function

Public Sub HBCA_UnloadObj()
'----------------------------------------------------------------------------------------------------------------------------------
'功能:卸载对象
'返回:无
'编制:余伟节
'日期:2015-08-31
'----------------------------------------------------------------------------------------------------------------------------------
    Set mCertMgr = Nothing
    Set mSVSClient = Nothing
    Set mSignCert = Nothing
    Set mFormSeal = Nothing
    mblnInit = False
End Sub
