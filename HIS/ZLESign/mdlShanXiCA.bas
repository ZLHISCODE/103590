Attribute VB_Name = "mdlShanXiCA"
Option Explicit

'Private mobjCertMgr As HebcaP11XLibCtl.CertMgr
Private mobjCertMgr As Object          'HebcaP11XLib.certMgr  证书对象

Private Const M_STR_LICENCE As String = "amViY55oZWKcZmhlnWxhaGViY2GXGmJjYWhlYnGH1QQ5GcNqnW6z3vohVnE+nTJr"
Private Const M_STR_SPLIT As String = "<SPLIT>"
Private mstrWSDL As String

Private Function CheckP7(ByVal strSignData As String) As Boolean
          Dim strEnvelope As String
          Dim strResult As String
          
1         On Error GoTo ErrH

2         strEnvelope = "<soap:Envelope xmlns:soap=""http://www.w3.org/2003/05/soap-envelope"" xmlns:snca=""http://snca.CertificateAuthorityServices/"">" & vbNewLine & _
                  "   <soap:Header/>" & vbNewLine & _
                  "   <soap:Body>" & vbNewLine & _
                  "      <snca:checkSNCAPKCS7Certificate>" & vbNewLine & _
                  "         <snca:PKCS7Info>" & strSignData & "</snca:PKCS7Info>" & vbNewLine & _
                  "      </snca:checkSNCAPKCS7Certificate>" & vbNewLine & _
                  "   </soap:Body>" & vbNewLine & _
                  "</soap:Envelope>"
3         LogWrite "CheckP7", "MXL:" & strEnvelope
4         strResult = httpPostSOAP(mstrWSDL, strEnvelope, ".//ns:return", "application/soap+xml;charset=UTF-8")
5         CheckP7 = IIf(strResult = "true", True, False)

6         Exit Function

ErrH:
7         MsgBox "在CheckP7的第" & Erl() & "行出错：" & vbCrLf & _
                  "错误号: " & Err.Number & vbCrLf & _
                  "错误描述：" & Err.Description, vbExclamation, gstrSysName
End Function

'Private Function GetTrustNumber(ByVal strSignData As String) As String
'    Dim strEnvelope As String
'    Dim strResult As String
'
'    strEnvelope = "<soap:Envelope xmlns:soap=""http://www.w3.org/2003/05/soap-envelope"" xmlns:snca=""http://snca.CertificateAuthorityServices/"">" & vbNewLine & _
'                "   <soap:Header/>" & vbNewLine & _
'                "   <soap:Body>" & vbNewLine & _
'                "      <snca:getSNCATrustNumber>" & vbNewLine & _
'                "         <snca:in>" & strSignData & "</snca:in>" & vbNewLine & _
'                "         <snca:type>PKCS7</snca:type>" & vbNewLine & _
'                "         <snca:expendingItemKey></snca:expendingItemKey>" & vbNewLine & _
'                "      </snca:getSNCATrustNumber>" & vbNewLine & _
'                "   </soap:Body>" & vbNewLine & _
'                "</soap:Envelope>"
'    strResult = httpPostSOAP(mstrWSDL, strEnvelope, ".//ns:return", "application/soap+xml;charset=UTF-8")
'    GetTrustNumber = strResult
'End Function

Private Function GetCertList(ByRef strName As String, ByRef strCertSn As String, Optional ByRef strCertDN As String, Optional ByRef objCert As Object) As Boolean
          '----------------------------------------------------------------------------------------------------------------------------------
          '功能:获取证书信息
          '参数:strName-证书用户名
          '     strCertSn-客服信任号
          '     strCertDN-DN
          '----------------------------------------------------------------------------------------------------------------------------------
          Dim intCount As Integer
                          
1         On Error GoTo ErrH
2         intCount = mobjCertMgr.GetDeviceCount
3         If intCount < 1 Then
4             MsgBoxEx "未发现KEY,请您插入Key！", vbInformation, gstrSysName
5             Exit Function
6         ElseIf intCount > 1 Then
7             MsgBoxEx "您的电脑上插入了多把陕西CA数字证书，请将多余的证书移除!", vbInformation, gstrSysName
8             Exit Function
9         End If

      '    intCount = mobjCertMgr.GetSignCertCount
      '    If intCount > 1 Then
      '        MsgBoxEx "您的电脑上插入了多把陕西CA数字证书，请将多余的证书移除!", vbInformation, gstrSysName
      '        Exit Function
      '    End If
      '     CN = 持有者姓名
      '    For i = 0 To intCount
      '        Set objCert = mobjCertMgr.GetCert(i)
      '        If Not objCert Is Nothing Then
      '            strDn = objCert.GetSubject()
      '        End If
      '    Next
10        Set objCert = mobjCertMgr.SelectSignCert
11        If objCert Is Nothing Then
12            MsgBoxEx "获取证书失败！", vbInformation, gstrSysName
13            Exit Function
14        End If
          
15        strName = objCert.GetSubjectItem("cn")
16        If strName = "" Then
17            MsgBoxEx "获取证书CN失败！", vbInformation, gstrSysName
18            Exit Function
19        End If
20        strCertDN = objCert.GetSubject()
21        strCertSn = objCert.GetCertExtensionByOid("1.2.86.11.7.11")
22        If strCertSn = "" Then
23            MsgBoxEx "获取客服信任号失败！", vbInformation, gstrSysName
24            Exit Function
25        End If
26        LogWrite "GetCertList", "CN:" & strName & vbCrLf & "SN:" & strCertSn & vbCrLf & "DN:" & strCertDN
27        GetCertList = True
28        Exit Function

ErrH:
29        MsgBox "在GetCertList的第" & Erl() & "行出错：" & vbCrLf & _
                  "错误号: " & Err.Number & vbCrLf & _
                  "错误描述：" & Err.Description, vbExclamation, gstrSysName
       
End Function

 
Private Function GetCertLogin(ByVal objCert As Object) As Boolean
      '功能:证书登录验证
          Dim strServiceTime As String
          Dim objDevice As Object
          Dim strRan As String
          
          Dim strSignData As String
          Dim strSource As String
          Dim blnResutl As Boolean
          
          '获取服务器时间
1         On Error GoTo ErrH

2         strServiceTime = GetCurrentTime()
          '产生随机数
3         Set objDevice = mobjCertMgr.GetDevice(0)
4         strRan = objDevice.GenRandom(128)
5         strSource = "TIME" & strServiceTime & "TIME" & strRan
          '验证证书
6         strSignData = SignedData(strSource, objCert)
7         If strSignData = "" Then Exit Function
8         blnResutl = CheckP7(strSignData)
9         If Not blnResutl Then
10            MsgBoxEx "服务器认证失败！", vbInformation, gstrSysName
11            Exit Function
12        End If
13        GetCertLogin = True

14        Exit Function

ErrH:
15        MsgBox "在GetCertLogin的第" & Erl() & "行出错：" & vbCrLf & _
                  "错误号: " & Err.Number & vbCrLf & _
                  "错误描述：" & Err.Description, vbExclamation, gstrSysName
End Function

'End Function
'
Private Function GetCurrentTime() As String
          Dim strEnvelope As String
          
1         On Error GoTo ErrH

2         strEnvelope = "<soap:Envelope xmlns:soap=""http://www.w3.org/2003/05/soap-envelope"" xmlns:snca=""http://snca.CertificateAuthorityServices/"">" & vbNewLine & _
                  "   <soap:Header/>" & vbNewLine & _
                  "   <soap:Body>" & vbNewLine & _
                  "      <snca:getCurrentTime>" & vbNewLine & _
                  "         <snca:time xsi:nil=""true"" xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance""/>" & vbNewLine & _
                  "      </snca:getCurrentTime>" & vbNewLine & _
                  "   </soap:Body>" & vbNewLine & _
                  "</soap:Envelope>"
          '2019-07-31-15-02-05
3         LogWrite "GetCurrentTime", "MXL:" & strEnvelope
4         GetCurrentTime = httpPostSOAP(mstrWSDL, strEnvelope, ".//ns:return", "application/soap+xml;charset=UTF-8")

5         Exit Function

ErrH:
6         MsgBox "在GetCurrentTime的第" & Erl() & "行出错：" & vbCrLf & _
                  "错误号: " & Err.Number & vbCrLf & _
                  "错误描述：" & Err.Description, vbExclamation, gstrSysName
End Function

Private Function SaveToDB(ByVal strSignData As String, ByVal strServerID As String, ByVal strAppID As String, ByVal strExtendId As String) As Boolean
      '功能:上传签名值
      '    strSignData=
      '    strServerID = 系统ID
      '    strAppID = 用GUID代替(因为无法获取业务系统ID以及签名记录ID,故将此值设置为GUID,方便后期取证的时候通过该值获取对方服务器签名值)
      '    strExtendId =默认为"00"
          Dim strEnvelope As String
          Dim strResult As String
          
1         On Error GoTo ErrH

2         strEnvelope = "<soap:Envelope xmlns:soap=""http://www.w3.org/2003/05/soap-envelope"" xmlns:snca=""http://snca.CertificateAuthorityServices/"">" & vbNewLine & _
              "   <soap:Header/>" & vbNewLine & _
              "   <soap:Body>" & vbNewLine & _
              "      <snca:checkSNCAPKCS7SignAndSaveToDB>" & vbNewLine & _
              "         <snca:PKCS7Info>" & strSignData & "</snca:PKCS7Info>" & vbNewLine & _
              "         <snca:Service_id>" & strServerID & "</snca:Service_id>" & vbNewLine & _
              "         <snca:app_id>" & strAppID & "</snca:app_id>" & vbNewLine & _
              "         <snca:extend_id>" & strExtendId & "</snca:extend_id>" & vbNewLine & _
              "      </snca:checkSNCAPKCS7SignAndSaveToDB>" & vbNewLine & _
              "   </soap:Body>" & vbNewLine & _
              "</soap:Envelope>"
          LogWrite "SaveToDB", "MXL:" & strEnvelope
3         strResult = httpPostSOAP(mstrWSDL, strEnvelope, ".//ns:return", "application/soap+xml;charset=UTF-8")
4         SaveToDB = IIf(strResult = "true", True, False)

5         Exit Function

ErrH:
6         MsgBox "在SaveToDB的第" & Erl() & "行出错：" & vbCrLf & _
                  "错误号: " & Err.Number & vbCrLf & _
                  "错误描述：" & Err.Description, vbExclamation, gstrSysName
End Function

Public Function ShanXi_CheckCert(Optional ByRef objCert As Object) As Boolean
          '功能：读取USB进行设备初始化并登录
          '返回值:
          '  strSigCert -签名证书内容

          Dim strSN As String
          Dim strName As String
          
1         On Error GoTo ErrH

2         If Not GetCertList(strName, strSN, , objCert) Then Exit Function
3         If mUserInfo.strCertSn <> strSN Then
4             MsgBoxEx "该证书未注册在您的名下，不能使用！", vbInformation + vbOKOnly, gstrSysName
5             gstrLogins = ""
6             Exit Function
7         End If

          '登录验证
8         If gstrLogins <> strSN Then '切换KEY后需要重新登录验证
9             If Not GetCertLogin(objCert) Then
10                gstrLogins = ""
11                Exit Function
12            Else
13                gstrLogins = strSN '标记上一验证通过的KEY
14            End If
15        End If
          
16        ShanXi_CheckCert = True
17        Exit Function

ErrH:
18        MsgBox "在ShanXi_CheckCert的第" & Erl() & "行出错：" & vbCrLf & _
                  "错误号: " & Err.Number & vbCrLf & _
                  "错误描述：" & Err.Description, vbExclamation, gstrSysName

End Function

 
Public Function ShanXi_GetPara() As Boolean
      '设置深圳CA服务器地址
          
1         On Error GoTo ErrH

2         On Error GoTo ErrH

3         With gudtPara
4             .strSignURL = GetThirdPara(CON_PAR_陕西, "签名服务WSDL")
5             .strOption = GetThirdPara(CON_PAR_陕西, "系统标识") '医院全称
6             .strTSIP = GetThirdPara(CON_PAR_陕西, "时间戳服务WSDL")  '时间戳服务WSDL
7             If .strSignURL = "" Or .strOption = "" Or .strTSIP = "" Then
8                 Exit Function
9             End If
10        End With

11        ShanXi_GetPara = True
12        Exit Function

ErrH:
13        MsgBox "在ShanXi_GetPara的第" & Erl() & "行出错：" & vbCrLf & _
                  "错误号: " & Err.Number & vbCrLf & _
                  "错误描述：" & Err.Description, vbExclamation, gstrSysName
End Function

Public Function ShanXi_InitObj() As Boolean
          
          Dim strMsg As String
          
1         On Error Resume Next
2         Set mobjCertMgr = CreateObject("HebcaP11X.CertMgr.1")
3         If Err.Number <> 0 Then
              'C:\Windows\System32\HebcaP11X.dll
4             strMsg = "创建证书管理对象失败！请检查部件【HebcaP11X.dll】是否正确安装并注册。"
5             GoTo ErrH
6         End If
12        On Error GoTo ErrH
13        mobjCertMgr.Licence = M_STR_LICENCE
14        If Not ShanXi_GetPara() Then
15            strMsg = "没有配置电子签名参数，请先到【公共参数设置】配置。"
16            GoTo ErrH
17        End If
      '    gudtPara.strSignURL = "http://111.20.164.185:8771/SNCA_CertificateAuthorityPlatform/services/CertificateAuthorityServices?wsdl"
18        mstrWSDL = gudtPara.strSignURL
19        ShanXi_InitObj = True
20        Exit Function
ErrH:
21       Call GetErrMsg(Erl(), strMsg)
End Function

Public Function ShanXi_RegCert(arrCertInfo As Variant) As Boolean
      '功能：提供在HIS数据库中注册个人证书的必要信息,用于新发放或更换证书,,需要插入USB-Key
      '返回：arrCertInfo作为数组返回证书相关信息
      '      0-ClientSignCertCN:客户端签名证书公共名称(姓名),用于注册证书时程序验证身份
      '      1-ClientSignCertDN:客户端签名证书主题(每人唯一)
      '      2-ClientSignCertSN:客户端签名证书序列号(每证书唯一)
      '      3-ClientSignCert:客户端签名证书内容
      '      4-ClientEncCert:客户端加密证书内容
      '      5-签名图片文件名,空串表示没有签名图片
      '      6-时间戳证书
          Dim strCertSn As String, strCertUserName As String, strCertDN As String
          Dim strSigCert As String, strTSCert As String
          Dim objSeal As Object
          Dim strBase64 As String
          Dim strFile As String
          
          Dim i As Long
1         On Error GoTo ErrH
            
2         For i = LBound(arrCertInfo) To UBound(arrCertInfo)
3             arrCertInfo(i) = ""
4         Next
          '取签章图片
          'ESEALREAD.ESealReadCtrl 0.1
5         On Error Resume Next
6         Set objSeal = CreateObject("ESEALREAD.ESealReadCtrl.1")
7         If Err.Number <> 0 Then
              'C:\Windows\System32\HebcaP11X.dll
8             MsgBoxEx "创建签章对象失败！请检查部件【ESealRead.ocx】是否正确安装并注册。", vbInformation, gstrSysName
9             Exit Function
10        End If
11        On Error GoTo ErrH
12        strBase64 = objSeal.ReadESeal(-3)
13        If strBase64 = "" Then
14            MsgBoxEx "获取签章BASE64失败！", vbInformation, gstrSysName
15            Exit Function
16        End If
          
17        If GetCertList(strCertUserName, strCertSn, strCertDN) Then
18            strFile = FormatPic("bmp", strCertSn, strBase64)
19            If strFile = "" Then
20                MsgBoxEx "生成签名图片失败！", vbInformation, gstrSysName
21                Exit Function
22            End If
23            arrCertInfo(0) = strCertUserName
24            arrCertInfo(1) = strCertDN
25            arrCertInfo(2) = strCertSn
26            arrCertInfo(3) = strSigCert
27            arrCertInfo(4) = ""
28            arrCertInfo(5) = strFile
29            arrCertInfo(6) = strTSCert
30            ShanXi_RegCert = True
31        End If
          
32        Exit Function

ErrH:
33        MsgBoxEx "证书注册失败！" & vbNewLine & "第" & CStr(Erl()) & "行 " & Err.Description, vbInformation, gstrSysName
End Function

Public Sub ShanXi_SetPara(ByVal strURL As String, ByVal strHosName As String, ByVal strTSURL As String)
1         On Error GoTo ErrH

2         With gudtPara
3             gudtPara.strSignURL = strURL
4             gudtPara.strOption = strHosName
5             gudtPara.strTSIP = strTSURL
6             Call UpdateThirdPara(CON_PAR_陕西, 1, "签名服务WSDL", .strSignURL, "签名服务WSDL")
7             Call UpdateThirdPara(CON_PAR_陕西, 2, "系统标识", .strOption, "系统唯一标识")
8             Call UpdateThirdPara(CON_PAR_陕西, 3, "时间戳服务WSDL", .strTSIP, "时间戳服务WSDL")
9         End With

10        Exit Sub

ErrH:
11        MsgBox "在ShanXi_SetPara的第" & Erl() & "行出错：" & vbCrLf & _
                  "错误号: " & Err.Number & vbCrLf & _
                  "错误描述：" & Err.Description, vbExclamation, gstrSysName
End Sub

Public Function ShanXi_Sign(ByVal strSource As String, ByRef strSignData As String, ByRef strTimeStamp As String, ByRef strTimeStampCode As String)
      '功能:签名
          Dim strDigest As String
          Dim strGUID As String
          Dim strBase64 As String
          Dim strGUIDTemp As String
          
          Dim objCert As Object
          
1         On Error GoTo ErrH
          
2         If ShanXi_CheckCert(objCert) Then
              '对原文数据产生hash值
3             If objCert.IsRSACert Then
4                 strDigest = mobjCertMgr.util.HashText(strSource, 1)
5             ElseIf objCert.IsSM2Cert Then
6                 strDigest = mobjCertMgr.util.HashText(strSource, 2)
7             End If
              '对hash值做签名
8             strSignData = SignedData(strDigest, objCert)
9             If strSignData = "" Then Exit Function
              
              '调用接口checkSNCAPKCS7SignAndSaveToDB 上传到认证服务器
10            strGUID = GUID()
11            If strGUID = "" Then
12                MsgBoxEx "获取GUID失败！", vbInformation, gstrSysName
13                Exit Function
14            End If
15            If Not SaveToDB(strSignData, gudtPara.strOption, strGUID, "00") Then
16                MsgBoxEx "上传服务器失败！", vbInformation, gstrSysName
17                Exit Function
18            End If
19            strSignData = strSignData & M_STR_SPLIT & strGUID
              '获取时间戳
20            strBase64 = EncodeBase64String(strSource)
21            strDigest = getHashValue(strBase64)
22            strGUIDTemp = left(strGUID, 1) & "," & Mid(strGUID, 2, Len(strGUID) - 2) & "," & Right(strGUID, 1)
23            strTimeStampCode = SignByTSA(strGUIDTemp, strDigest)
24            If strTimeStampCode = "" Then
25                MsgBoxEx "获取时间戳信息失败！", vbInformation, gstrSysName
26                Exit Function
27            End If
28            strTimeStamp = getSignTime(strTimeStampCode)
29            If strTimeStamp = "" Then
30                MsgBoxEx "从时间戳签名值中获取签名时间失败！", vbInformation, gstrSysName
31                Exit Function
32            End If
33            strTimeStamp = Format(strTimeStamp, "yyyy-MM-dd HH:mm:ss")
34        End If
35        ShanXi_Sign = True
          
36        Exit Function

ErrH:
37        MsgBox "在ShanXi_Sign的第" & Erl() & "行出错：" & vbCrLf & _
                  "错误号: " & Err.Number & vbCrLf & _
                  "错误描述：" & Err.Description, vbExclamation, gstrSysName

End Function

Public Sub ShanXi_UnloadObj()
    Set mobjCertMgr = Nothing
End Sub

 
Public Function ShanXi_VerifySign(ByVal strSign As String, ByVal strSource As String) As Boolean
      '功能:验证签名
      '时间戳验证签名
          Dim arrSign As Variant
          Dim strGUID As String
          Dim strGUIDTemp As String
          Dim strBase64 As String
          Dim strDigest As String
          Dim blnRet As Boolean
          Dim strMsg As String
1         On Error GoTo ErrH

2         arrSign = Split(strSign, M_STR_SPLIT) '签名值<SPLIT>GUID
3         If CheckP7(arrSign(0)) Then
4             strMsg = "验证成功，签名数据有效。"
5         Else
6             MsgBoxEx "验证失败，签名数据无效！", vbInformation, gstrSysName
7             Exit Function
8         End If
          
9         If UBound(arrSign) >= 1 Then
10            strGUID = arrSign(1)
11            strBase64 = EncodeBase64String(strSource)
12            strDigest = getHashValue(strBase64)
13            strGUIDTemp = left(strGUID, 1) & "," & Mid(strGUID, 2, Len(strGUID) - 2) & "," & Right(strGUID, 1)
14            blnRet = verifyContentByTSA(strGUIDTemp, strDigest)
15            If blnRet Then
16                strMsg = strMsg & vbCrLf & "时间戳验证成功。"
17            Else
18                MsgBoxEx "时间戳验证失败！", vbInformation, gstrSysName
19                Exit Function
20            End If
21        End If
22        If strMsg <> "" Then
23           MsgBoxEx strMsg, vbInformation, gstrSysName
24        End If
25        ShanXi_VerifySign = True

26        Exit Function

ErrH:
27        MsgBox "在ShanXi_VerifySign的第" & Erl() & "行出错：" & vbCrLf & _
                  "错误号: " & Err.Number & vbCrLf & _
                  "错误描述：" & Err.Description, vbExclamation, gstrSysName
End Function

Private Function SignedData(ByVal strSource As String, ByVal objCert As Object) As String
          Dim objPkcs7 As Object
          Dim strSignData As String
          Dim strCertB64 As String
          
1         On Error GoTo ErrH

2         strCertB64 = objCert.GetCertB64()
3         Set objPkcs7 = mobjCertMgr.CreatePkcs7()
4         objPkcs7.AddRecipientCert (strCertB64)
5         If objCert.IsRSACert Then
6             On Error Resume Next
7             strSignData = objPkcs7.SignText(0, strSource, 1)
8             If Err.Number = -536145911 Then Exit Function '取消密码窗口
9             If Err.Number > 0 Then GoTo ErrH
10            On Error GoTo ErrH
11        ElseIf objCert.IsSM2Cert Then
12            On Error Resume Next
13            strSignData = objPkcs7.SignText(0, strSource, 2)
14            If Err.Number = -536145911 Then Exit Function  '用户取消操作
15            If Err.Number > 0 Then GoTo ErrH
16            On Error GoTo ErrH
17        Else
18            MsgBoxEx "证书不支持RSA/SM2算法！", vbInformation, gstrSysName
19        End If
20        If strSignData <> "" Then
21            If Not objPkcs7.VerifyB64(strSignData) Then
22                MsgBoxEx "验证签名失败(RSA)！", vbInformation, gstrSysName
23                Exit Function
24            End If
25        Else
26            MsgBoxEx "签名失败：签名值为空！", vbInformation, gstrSysName
27        End If
28        SignedData = strSignData
29        Exit Function

ErrH:
30        MsgBox "在SignedData的第" & Erl() & "行出错：" & vbCrLf & _
                  "错误号: " & Err.Number & vbCrLf & _
                  "错误描述：" & Err.Description, vbExclamation, gstrSysName
End Function

'Private Function SNCAGetCertForPwd(ByVal strPWD As String) As Object
'      ' 输入客户端用户证书PIN码获取证书信息
'      ' 参数1:证书类型默认为0
'      ' 参数2:用户证书PIN码
'      ' 返回:用户客户端签名证书对象
'          Dim intCount As Integer
'          Dim i As Long
'
'          Dim objCert As Object
'
'1         On Error GoTo ErrH
'
'2         intCount = mobjCertMgr.GetCertCount
'3         For i = 0 To intCount
'4             Set objCert = mobjCertMgr.GetSignCert(i)
'5             On Error Resume Next
'6             Call objCert.Login(strPWD)
'7             If Err.Number <> 0 Then
'8               MsgBox Err.Description
'9             End If
'10            On Error GoTo 0
'11            If Not objCert Is Nothing Then
'12                Exit For
'13            End If
'14        Next
'15        Set SNCAGetCertForPwd = objCert
'
'16        Exit Function
'
'ErrH:
'17        MsgBox "在SNCAGetCertForPwd的第" & Erl() & "行出错：" & vbCrLf & _
'                  "错误号: " & Err.Number & vbCrLf & _
'                  "错误描述：" & Err.Description, vbExclamation, gstrSysName

Public Function VerifyB64(ByVal strSignData As String) As Boolean
      '功能:验证签名
          Dim objCert As Object
          Dim objPkcs7 As Object
          
          Dim strCertB64 As String
          Dim blnResult As Boolean
          
1         On Error GoTo ErrH

2         Set objCert = mobjCertMgr.SelectSignCert
3         strCertB64 = objCert.GetCertB64
4         Set objPkcs7 = mobjCertMgr.CreatePkcs7()
5         Call objPkcs7.AddRecipientCert(strCertB64)
6         blnResult = objPkcs7.VerifyB64(strSignData)
          VerifyB64 = blnResult
7         Exit Function

ErrH:
8         MsgBox "在VerifyB64的第" & Erl() & "行出错：" & vbCrLf & _
                  "错误号: " & Err.Number & vbCrLf & _
                  "错误描述：" & Err.Description, vbExclamation, gstrSysName
End Function

Public Function SignByTSA(ByVal strBusinessID As String, ByVal strHashSource As String) As String
 '功能:时间戳签名
 '参数：strBusinessID 首位后加","末位前加"," 例如：6,1234567,8
    Dim strEnvelope As String
    Dim strBase64 As String
   
    On Error GoTo ErrH

    strEnvelope = "<soapenv:Envelope xmlns:soapenv=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:web=""http://webservice.client.tsp.snca.com/"">" & vbNewLine & _
                "   <soapenv:Header/>" & vbNewLine & _
                "   <soapenv:Body>" & vbNewLine & _
                "      <web:signByTSA>" & vbNewLine & _
                "         <arg0>HASH</arg0>" & vbNewLine & _
                "         <arg1>" & strBusinessID & "</arg1>" & vbNewLine & _
                "         <arg2>" & strHashSource & "</arg2>" & vbNewLine & _
                "         <arg3>SHA</arg3>" & vbNewLine & _
                "      </web:signByTSA>" & vbNewLine & _
                "   </soapenv:Body>" & vbNewLine & _
                "</soapenv:Envelope>"
    LogWrite "SignByTSA", "调用【SignByTSA】传入值:" & strEnvelope
    strBase64 = httpPostSOAP(gudtPara.strTSIP, strEnvelope, ".//return")
    LogWrite "SignByTSA", "调用【SignByTSA】返回值:" & strBase64
    SignByTSA = strBase64

    Exit Function

ErrH:
    MsgBox "在SignByTSA的第" & Erl() & "行出错：" & vbCrLf & _
            "错误号: " & Err.Number & vbCrLf & _
            "错误描述：" & Err.Description, vbExclamation, gstrSysName
End Function

Public Function verifyContentByTSA(ByVal strBusinessID As String, ByVal strHash As String) As Boolean
'功能:验证时间戳签名
'参数：strBusinessID 首位后加","末位前加"," 例如：6,1234567,8
'      strHash-原文摘要
    Dim strEnvelope As String
    Dim strResult As String
    
    On Error GoTo ErrH

    strEnvelope = "<soapenv:Envelope xmlns:soapenv=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:web=""http://webservice.client.tsp.snca.com/"">" & vbNewLine & _
                "   <soapenv:Header/>" & vbNewLine & _
                "   <soapenv:Body>" & vbNewLine & _
                "      <web:verifyContentByTSA>" & vbNewLine & _
                "         <arg0>HASH</arg0>" & vbNewLine & _
                "         <arg1>" & strBusinessID & "</arg1>" & vbNewLine & _
                "         <arg2>" & strHash & "</arg2>" & vbNewLine & _
                "         <arg3>SHA</arg3>" & vbNewLine & _
                "      </web:verifyContentByTSA>" & vbNewLine & _
                "   </soapenv:Body>" & vbNewLine & _
                "</soapenv:Envelope>"
    LogWrite "verifyContentByTSA", "调用【verifyContentByTSA】传入值:" & strEnvelope
    strResult = httpPostSOAP(gudtPara.strTSIP, strEnvelope, ".//return")
    LogWrite "verifyContentByTSA", "调用【verifyContentByTSA】返回值:" & strResult
    verifyContentByTSA = IIf(UCase(strResult) = UCase("True"), True, False)
    Exit Function

ErrH:
    MsgBox "在verifyContentByTSA的第" & Erl() & "行出错：" & vbCrLf & _
            "错误号: " & Err.Number & vbCrLf & _
            "错误描述：" & Err.Description, vbExclamation, gstrSysName
End Function


Public Function getHashValue(ByVal strBase64Source As String) As String
'功能:获得指定算法的明文的摘要值，支持SHA1 摘要算法
'参数:strBase64Source-原文转Base64字符串

    Dim strEnvelope As String
    Dim strBase64 As String

    On Error GoTo ErrH

    strEnvelope = "<soapenv:Envelope xmlns:soapenv=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:web=""http://webservice.client.tsp.snca.com/"">" & vbNewLine & _
                    "   <soapenv:Header/>" & vbNewLine & _
                    "   <soapenv:Body>" & vbNewLine & _
                    "      <web:getHashValue>" & vbNewLine & _
                    "         <arg0>" & strBase64Source & "</arg0>" & vbNewLine & _
                    "         <arg1>SHA</arg1>" & vbNewLine & _
                    "      </web:getHashValue>" & vbNewLine & _
                    "   </soapenv:Body>" & vbNewLine & _
                    "</soapenv:Envelope>"

    LogWrite "getHashValue", "调用【getHashValue】传入值:" & strEnvelope
    strBase64 = httpPostSOAP(gudtPara.strTSIP, strEnvelope, ".//return")
    LogWrite "getHashValue", "调用【getHashValue】返回值:" & strBase64
    getHashValue = strBase64

    Exit Function

ErrH:
    MsgBox "在getHashValue的第" & Erl() & "行出错：" & vbCrLf & _
            "错误号: " & Err.Number & vbCrLf & _
            "错误描述：" & Err.Description, vbExclamation, gstrSysName
End Function

Public Function getSignTime(ByVal strTimeStampCode As String) As String
'功能:从时间戳签名值中获取签名时间
'参数:strTimeStampCode-时间戳签名值

    Dim strEnvelope As String
    Dim strResult As String

    On Error GoTo ErrH

    strEnvelope = "<soapenv:Envelope xmlns:soapenv=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:web=""http://webservice.client.tsp.snca.com/"">" & vbNewLine & _
                    "   <soapenv:Header/>" & vbNewLine & _
                    "   <soapenv:Body>" & vbNewLine & _
                    "      <web:getSignTime>" & vbNewLine & _
                    "         <arg0>" & strTimeStampCode & "</arg0>" & vbNewLine & _
                    "      </web:getSignTime>" & vbNewLine & _
                    "   </soapenv:Body>" & vbNewLine & _
                    "</soapenv:Envelope>"

    LogWrite "getSignTime", "调用【getSignTime】传入值:" & strEnvelope
    strResult = httpPostSOAP(gudtPara.strTSIP, strEnvelope, ".//return")
    LogWrite "getSignTime", "调用【getSignTime】返回值:" & strResult
    getSignTime = strResult

    Exit Function

ErrH:
    MsgBox "在getSignTime的第" & Erl() & "行出错：" & vbCrLf & _
            "错误号: " & Err.Number & vbCrLf & _
            "错误描述：" & Err.Description, vbExclamation, gstrSysName
End Function


