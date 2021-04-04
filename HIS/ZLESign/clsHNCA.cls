VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsHNCA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'河南商丘项目
Private mobjHNCACTRL As Object
Private mobjHNCATS As Object
Private mblnTs As Boolean
Private mbytFunc As Byte   '0-RSA算法,1-SM2算法

Private Declare Function doPostEx Lib "HNCACertTools6.dll" (ByVal strUrl As String, ByVal strAPPCode As String, ByVal strSource As String, _
        ByVal strSignData As String, ByVal strCertPUB As String, ByVal strGBK As String) As Integer

Public Function HNCA_CheckCert(ByRef blnReDo As Boolean) As Boolean
'功能：读取USB进行设备初始化并登录
    Dim strKey As String, strPIN As String, strUserName As String
    Dim strCertName As String, strCertDN As String, strPicPath As String
    Dim strCertSn As String
    Dim strCertUserID As String    '包含身份证号信息
    Dim strDate As String
    Dim udtUser As USER_INFO
    Dim strCert As String, strCertID As String
    Dim blnOk As Boolean
    Dim blnRet As Boolean
    
    On Error GoTo errH
1000
     '获取证书信息同时检查Key盘是否插入
    If Not GetCertList(strCertName, strCertSn, strCert, strCertDN, strPicPath, strCertUserID) Then
        HNCA_CheckCert = False: Exit Function
    End If
    '未注册在当前用户名下的Key
    If mbytFunc = 0 Then
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
    Else
        If mUserInfo.strCertSn <> strCertSn Then
            MsgBoxEx "该证书未注册在您的名下，不能使用！", vbInformation + vbOKOnly, gstrSysName
            Exit Function
        End If
    End If
1100
    '登录验证
    If InStr(gstrLogins & "|", "|" & strCertSn & "|") > 0 Then '首次验证通过后，下次不在继续验证
        blnOk = True
    Else
        If Not GetCertLogin(strCert, strCertDN) Then
            blnOk = False
        Else
            blnOk = True
            If InStr(gstrLogins & "|", "|" & strCertSn & "|") = 0 Then gstrLogins = gstrLogins & "|" & strCertSn
        End If
    End If
1200
    If blnOk And mbytFunc = 0 Then
        '判断是否需要更新注册证书
        udtUser.strName = strCertName
        udtUser.strSignName = strCertName
        udtUser.strUserID = strCertUserID
        udtUser.strCertSn = strCertSn
        udtUser.strCertDN = strCertDN
        udtUser.strCert = strCert
        udtUser.strPicPath = strPicPath
        udtUser.strPicCode = ""
        '获取已经注册证书的有效结束日期 日期格式:2005年12月23日 18:09:31
        strDate = mobjHNCACTRL.GetCertInfo(6, "")
        strDate = Format(strDate, "YYYY-MM-DD HH:MM:SS")
        If IsUpdateRegCert(udtUser, strDate, blnReDo) Then
            HNCA_CheckCert = True
        Else
            HNCA_CheckCert = False
        End If
    Else
        HNCA_CheckCert = blnOk
    End If
1300
    
    Exit Function
errH:
     MsgBoxEx "检查USBKEY失败！" & vbNewLine & "第" & CStr(Erl()) & "行 " & Err.Description, vbInformation, gstrSysName
End Function

Public Function HNCA_RegCert(arrCertInfo As Variant) As Boolean
        '功能：提供在HIS数据库中注册个人证书的必要信息,用于新发放或更换证书,,需要插入USB-Key
        '返回：arrCertInfo作为数组返回证书相关信息
        '      0-ClientSignCertCN:客户端签名证书公共名称(姓名),用于注册证书时程序验证身份
        '      1-ClientSignCertDN:客户端签名证书主题(每人唯一)
        '      2-ClientSignCertSN:客户端签名证书序列号(每证书唯一)
        '      3-ClientSignCert:客户端签名证书内容
        '      4-ClientEncCert:客户端加密证书内容
        '      5-签名图片文件名,空串表示没有签名图片
        
        Dim strCertSn As String, strCertTime As String, strCertUserName As String, strCertDN As String
        Dim strSigCert As String, i As Integer, strCACert As String, lngOk As Long
        Dim strPicPath As String
        On Error GoTo errH
    
100     For i = LBound(arrCertInfo) To UBound(arrCertInfo)
102         arrCertInfo(i) = ""
        Next
    
104     If GetCertList(strCertUserName, strCertSn, strSigCert, strCertDN, strPicPath) Then
106         arrCertInfo(0) = strCertUserName
108         arrCertInfo(1) = strCertDN
110         arrCertInfo(2) = strCertSn
112         arrCertInfo(3) = strSigCert
            arrCertInfo(4) = ""
113         arrCertInfo(5) = strPicPath

            
124         HNCA_RegCert = True
        End If

        Exit Function
errH:
126     MsgBoxEx "获取证书信息失败！" & vbNewLine & "第" & CStr(Erl()) & "行 " & Err.Description, vbQuestion, gstrSysName

End Function

Public Function HNCA_Sign(ByVal strCurrCertSn As String, ByVal strSource As String, ByRef strSignData As String, _
        ByRef strTimeStamp As String, ByRef strTimeStampCode As String, ByRef blnReDo As Boolean) As Boolean
        '签名
        Dim strTSURL As String, strUserName As String, strUserPWD As String
        Dim strRequest As String
        Dim intRet As Integer, intPort As Integer, intSSLPort As Integer
        Dim blnCheck As Boolean, blnRet As Boolean
        Dim strMsg As String
        Dim arrTmp As Variant
        
        On Error GoTo errH
        blnCheck = HNCA_CheckCert(blnReDo)
        If blnReDo Then Exit Function
1100    If blnCheck Then
            '验证当前USB是否是签名用户的，并获取签名证书
            If mbytFunc = 0 Then
1110            strSignData = mobjHNCACTRL.Sign(strCurrCertSn, strSource)    '签名值不含数据源
            Else
1111            strSignData = mobjHNCACTRL.SM2Sign("", strSource)    '签名值不含数据源
            End If
            If strSignData <> "" Then
                If mblnTs Then
                    '由原文申请时间戳
                    intRet = -1
                    strRequest = mobjHNCATS.SOF_CreateTimeStampRequestByPlain(strSource)
                    intRet = mobjHNCATS.SOF_GetLastError()
                    If (intRet <> 0) Then
                        strMsg = "创建时间戳请求失败，失败代码是：" & intRet
                        blnRet = False
                    Else
                        intRet = -1
                        strTimeStampCode = mobjHNCATS.SOF_CreateTimeStampResponse(strRequest)
                        intRet = mobjHNCATS.SOF_GetLastError()
                        If (intRet <> 0) Then
                            strMsg = "创建时间戳响应失败，失败代码是：" & intRet
                            blnRet = False
                        Else
                            strTimeStamp = mobjHNCATS.SOF_GetTimeStampInfo(strTimeStampCode, 1) '格式;2015年10月20日 02:18:45.534  GMT
                            strTimeStamp = Format(Mid(strTimeStamp, 1, 20), "YYYY-MM-DD HH:MM:SS")
                            blnRet = True
                        End If
                    End If
                Else
                    blnRet = True
                    strTimeStamp = Format(gobjComLib.zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
                End If
            Else
                strMsg = "签名失败！"
                blnRet = False
1112        End If
        Else
            strMsg = "签名失败！"
            blnRet = False
        End If
        If strMsg <> "" Then
            MsgBoxEx strMsg, vbInformation, gstrSysName
        End If
                
        HNCA_Sign = blnRet
        Exit Function
errH:
114     MsgBoxEx "签名失败！" & vbNewLine & "第" & CStr(Erl()) & "行 " & Err.Description, vbQuestion, gstrSysName
End Function

Public Function HNCA_VerifySign(ByVal strCert As String, ByVal strSignData As String, ByVal strSource As String, ByVal strTimeStampCode As String) As Boolean
'验证签名
'
        Dim strSigCert As String, strMsg As String
        Dim intRet As Integer
        Dim blnRet As Boolean
        On Error GoTo errH
     
        '获取客户端签名(公钥)证书Base64编码 客户端验证签名,不含数据源的验证签名
        If mbytFunc = 0 Then
1000        intRet = mobjHNCACTRL.VerifySign(strSignData, strSource, strCert)
        Else
1001        intRet = mobjHNCACTRL.SM2Verify(strSignData, strSource, strCert)
        End If
        If intRet = 0 Then   '验证签名成功
            If mblnTs And strTimeStampCode <> "" Then
                '验证时间戳
                intRet = -1
                intRet = mobjHNCATS.SOF_VerifyTimeStampByPlain(strSource, strTimeStampCode)
            End If
            If (intRet = 0) Then
                strMsg = "验签成功！"
                blnRet = True
            Else
                strMsg = "时间戳验证失败，错误代码：" & intRet
                blnRet = False
            End If
        Else
            strMsg = "验签失败！"
            blnRet = False
        End If
            
        If strMsg <> "" Then
            MsgBoxEx strMsg, vbInformation, gstrSysName
        End If
        HNCA_VerifySign = blnRet
        
2000    Exit Function
errH:
104     MsgBoxEx "验证签名失败！" & vbNewLine & "第" & CStr(Erl()) & "行 " & Err.Description, vbQuestion, gstrSysName
End Function


Private Function GetCertList(ByRef strName As String, ByRef strCertSn As String, ByRef strCertPUB As String, _
                ByRef strCertDN As String, Optional strPicPath As String, Optional strUserID As String) As Boolean
'功能:获取证书信息
'-出参
'    strName 证书持有者姓名
'   strCertSN 证书唯一标识
'   strCertPUB 证书公钥  验证签名用到
'   strCertDN 证书描述信息  证书注册用到
'   strPicPath 证书图片保存位置

    Dim strCertCN As String, strPic As String
    Dim intRet As Integer, intCnIndex As Integer
    Dim strPara As String, strSign As String
    
    On Error GoTo errH

10    If mobjHNCACTRL Is Nothing Then Set mobjHNCACTRL = CreateObject("HNCACTRL.HNCACtrlCtrl.1")
      If mobjHNCATS Is Nothing Then Set mobjHNCATS = CreateObject("TSAACTIVEXVER4.TsaActiveXver4Ctrl.1")    '时间戳对象
      
    '0 表示只有一张证书时也弹出证书选择框，1 表示只有一张证书时将不弹出证书选择框，默认值为0
20    mobjHNCACTRL.SetCertChooseType (1)
    '弹出HNCA颁发的签名证书列表
    If mbytFunc = 0 Then
30      intRet = mobjHNCACTRL.SetCert("SC", "", "", "", "CN=HNCA,O=HNCA,L=ZZ,S=HN,C=CN", "")
    Else
31      intRet = mobjHNCACTRL.SetCertEx("SC", "", "", "", "", "", "SM2")
    End If
    '判断是否初始化成功
    If intRet <> 0 Then
        MsgBoxEx "初始化客户端证书错误：" & mobjHNCACTRL.GetErrorMessage(), vbOKOnly, "河南CA"
        Exit Function
    End If

    '以下为业务系统中从主题(DN)项
40     strCertDN = mobjHNCACTRL.GetCertInfo(0, "")
            
    '以下为业务系统中从主题(DN)项中提取CN项

50      intCnIndex = InStr(strCertDN, "CN=") '获取"CN="项位置
51      strCertCN = Mid(strCertDN, intCnIndex) '将CN项前的信息去除
52      intCnIndex = InStr(strCertCN, ",") '第1个逗号所在位置
53      If (intCnIndex <> -1) Then
54          strCertCN = Mid(strCertCN, 1, intCnIndex - 1)
55      End If
56      strCertCN = Mid(strCertCN, 4)
        '证书CN项 姓名
57      strName = Trim(strCertCN)

        '获取证书序列号,SN项目，每个证书唯一标识
60      strCertSn = mobjHNCACTRL.GetCertInfo(2, "")
        '获取客户端签名(公钥)证书Base64编码
80      strCertPUB = mobjHNCACTRL.GetCertInfo(8, "")
        '获取扩展项（根据项目确定扩展域的参数）
81      strPara = IIf(mbytFunc = 0, "1.2.86.11.7.12", "1.2.86.11.7.11")
82      strUserID = mobjHNCACTRL.GetCertInfo(7, strPara)  '身份证号
        If mbytFunc = 0 Then
90          strPic = mobjHNCACTRL.LoadSealFromKey(strCertSn, "111111")   '返回印章数据的BASE64
        Else
91          strPara = "测试"
92          strSign = mobjHNCACTRL.SignSealForSM2(strCertSn, strPara)
93          intRet = mobjHNCACTRL.VerifySealForSM2(strSign, strPara)
94          strPic = mobjHNCACTRL.GetSealForSM2(strSign, intRet)
        End If
        If strPic <> "" Then
100        If mbytFunc = 0 Then strPic = mobjHNCACTRL.GetBmpFromSealData(strPic, 2) '返回印章图片数据的BASE64
120        strPicPath = SaveBase64ToFile("bmp", strCertSn, strPic) '返回印章图片数据的BASE64转换成图片文件并返回图片位置
        Else
            MsgBoxEx "读取图片失败！请检查是否插上key盘！", vbInformation, gstrSysName
            Exit Function
        End If
    
130     GetCertList = True
        Exit Function
errH:
500     MsgBoxEx "获取证书信息失败！" & vbNewLine & "第" & CStr(Erl()) & "行 " & Err.Description, vbInformation, gstrSysName
End Function

Private Function GetCertLogin(ByVal strCertPUB As String, ByVal strCertDN As String) As Boolean
    Dim blnRet As Boolean
    Dim strSignData As String
    Dim strUrl As String, strAPPCode As String, strRand As String, strMsg As String
    Dim strTSIP As String, strUserName As String, strUserPWD As String
    Dim intRet As Integer, intPort As Integer, intSSLPort As Integer
    Dim arrTmp As Variant
        On Error GoTo errH
1000
         arrTmp = Split(gstrPara, "|")
        'gstrPara= http://218.28.16.104:7080/CAServer/servlet/CertChechServlet|218.28.16.104|8080|443|时间戳（0-不启用;1-启用）|签名算法(0-RSA\1-SM2)"
        intRet = -1
        strUrl = CStr(Trim(arrTmp(0)))
        strAPPCode = "1101"
        strTSIP = CStr(Trim(arrTmp(1)))
        intPort = CInt(arrTmp(2))
        intSSLPort = CInt(arrTmp(3))
        strUserName = "abc"
        strUserPWD = "123"
        
1100
        '获取随机数
        Randomize
        strRand = CStr(Int((9001) * Rnd + 1000))


        '随机数签名
        If mbytFunc = 0 Then
1101        strSignData = mobjHNCACTRL.AttachSign(strCertDN, strRand)
        Else
1102        strSignData = mobjHNCACTRL.SM2Sign("", strRand)
        End If
        
        If strSignData = "" Then
            strMsg = mobjHNCACTRL.GetErrorMessage()  ' 获取错误信息
            MsgBoxEx "随机数签名失败:" + strMsg, vbOKOnly, gstrSysName
            Exit Function
        End If
1200
        strMsg = ""
        '发送服务器身份验证
        intRet = doPostEx(strUrl, strAPPCode, strRand, strSignData, strCertPUB, "GBK")

        strMsg = ""

        If (intRet = 0) Then
            blnRet = True
        ElseIf (intRet < 0) Then
            '随机数验签失败
            blnRet = False
            strMsg = "登录服务器验签失败！"
        Else
            blnRet = True
            strMsg = "登录服务器验签成功！您的证书还剩：" & intRet & " 天过期！"
        End If
        If strMsg <> "" Then
           MsgBoxEx strMsg, vbInformation, gstrSysName
        End If
        '验证失败直接退出
        If Not blnRet Then Exit Function
1300
        '时间戳设置系统参数
        If mblnTs Then
            strMsg = ""
            intRet = mobjHNCATS.SetSystemParameter(strTSIP, intPort, intSSLPort)
            If (intRet = 0) Then        'MessageBox.Show("设置成功！")
                
                '2.设置哈希算法，如果不调用默认是sha1
                Call mobjHNCATS.SetMD(1)
                
                '3.用户登录
                If (mobjHNCATS.ISLogin = 1) Then
                    blnRet = True
                Else
                    intRet = -1
                    intRet = mobjHNCATS.LoginByUsername(strUserName, strUserPWD)    '
                    If (intRet = 0) Then
                        blnRet = True
                    Else
                        blnRet = False
                        strMsg = "登录时间戳服务器失败！错误代码为：" & intRet
                    End If
                End If
            Else
                strMsg = "设置时间戳服务器的系统参数失败，错误代码：" & intRet
                blnRet = False
            End If
                
            If strMsg <> "" Then
               MsgBoxEx strMsg, vbInformation, gstrSysName
            End If
        End If
1600
    GetCertLogin = blnRet
    Exit Function
errH:
    MsgBoxEx "登录服务器验证失败！" & vbNewLine & "第" & CStr(Erl()) & "行 " & Err.Description, vbInformation, gstrSysName
End Function

Public Function HNCA_Initialize() As Boolean
    '功能:创建河南CA控件对象
    
    Dim intRet As Integer
    
    On Error GoTo errH
    
1000
        Set mobjHNCACTRL = CreateObject("HNCACTRL.HNCACtrlCtrl.1") '签名对象
1001
        Set mobjHNCATS = CreateObject("TSAACTIVEXVER4.TsaActiveXver4Ctrl.1")  '时间戳对象
1100
        gstrPara = gobjComLib.zlDatabase.GetPara(90000, glngSys) '读取URL 服务器
        'gstrPara = "http://218.28.16.104:7080/CAServer/servlet/CertChechServlet|218.28.16.104|8080|443|时间戳(0-不启用/1-启用)|签名算法(0-RSA\1-SM2)"
        If gstrPara = "" Then
            MsgBoxEx "没有配置签名服务器地址，请到启用电子签名接口处设置。", vbInformation, gstrSysName
            Exit Function
        Else
            If UBound(Split(gstrPara, "|")) <> 5 Then
               MsgBoxEx "签名服务器地址配置有误，请到启用电子签名接口处设置。" & _
                    vbCrLf & vbTab & "参数值格式""服务器URL|TSIP|TSPORT|SSLPORT|是否启用时间戳(0-不启用/1-启用)|签名算法(0-RSA\1-SM2)""", vbInformation, gstrSysName
                Exit Function
            End If
        End If
        mblnTs = Val(Split(gstrPara, "|")(4)) = 1
        mbytFunc = CByte(Split(gstrPara, "|")(5))
1500
        HNCA_Initialize = True
        Exit Function

errH:
     MsgBoxEx "创建接口部件失败！" & vbNewLine & "第" & CStr(Erl()) & "行 " & Err.Description, vbQuestion, gstrSysName
End Function

Private Sub Class_Terminate()
    Set mobjHNCACTRL = Nothing
    Set mobjHNCATS = Nothing
End Sub

Public Function HNCA_GetPara() As Boolean
'设置服务器地址
    
    On Error GoTo errH
    gstrPara = gobjComLib.zlDatabase.GetPara(90000, glngSys)  '读取URLs 固定读取ZLHIS 系统默认100
    'gstrPara = "http://218.28.16.104:7080/CAServer/servlet/CertChechServlet|218.28.16.104|8080|443|时间戳(0-不启用/1-启用)|签名算法(0-RSA\1-SM2)"
    If gstrPara = "" Then gstrPara = "http://218.28.16.104:7080/CAServer/servlet/CertChechServlet|218.28.16.104|8080|443|0|0"
    If gstrPara <> "" Then
        gudtPara.strSignURL = gstrPara
    End If
    Exit Function
errH:
    MsgBoxEx "读取参数失败！" & vbNewLine & "第" & CStr(Erl()) & "行 " & Err.Description, vbInformation, gstrSysName
End Function

Public Function HNCA_SetParaStr() As String
    HNCA_SetParaStr = gudtPara.strSignURL
End Function


