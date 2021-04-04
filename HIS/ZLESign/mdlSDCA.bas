Attribute VB_Name = "mdlSDCA"
Option Explicit
Private mobjSDCA As Object  '用于山东签名
Private mobjTSA As Object   '用于山东：时间戳
Private mobjSVS As Object   '用于山东：验签
Private mobjBase64 As Object   'BASE64对象

Public Function SDCA_CheckCert(ByRef blnReDo As Boolean, Optional ByRef strCertID As String) As Boolean
      '--------------------------------------------------------------------------------------------------------------------------
      '功能：读取USB进行设备初始化并登录
      '参数:
      '   出参:blnRedo-证书更新需要重新检查
      '返回:
      '--------------------------------------------------------------------------------------------------------------------------
          Dim strUniqueID As String
          Dim strName As String
          Dim strCertSn As String
          Dim strCertDN As String
          Dim strSigCert As String
          Dim strDate As String
          Dim blnRet As Boolean
          
          Dim udtUser As USER_INFO
          

1         On Error GoTo ErrH

2         If Not GetCertList(, strCertSn, , , strUniqueID, strCertID) Then Exit Function
3         If mUserInfo.strUserID = "" Then
4             MsgBoxEx "您的身份证号为空,请联系管理员到人员管理中录入！", vbOKOnly + vbInformation, gstrSysName
5             Exit Function
6         ElseIf mUserInfo.strUserID <> strUniqueID Then
7             MsgBoxEx "您的身份证号：" & _
                     vbCrLf & vbTab & "【" & mUserInfo.strUserID & "】" & vbCrLf & _
                     "证书身份证号:" & _
                     vbCrLf & vbTab & "【" & strUniqueID & "】" & vbCrLf & _
                     "用户身份证号与证书身份证号不相等,不能使用！", vbInformation, gstrSysName
8             Exit Function
9         End If
          
          '登录验证
10        blnRet = True
11        If mUserInfo.strCertSn <> strCertSn Then
12            If Not GetCertList(strName, , strSigCert, strCertDN) Then Exit Function
              '判断是否需要更新注册证书
13            udtUser.strName = strName
14            udtUser.strSignName = strName
15            udtUser.strUserID = strUniqueID
16            udtUser.strCertSn = strCertSn
17            udtUser.strCertDN = strCertDN
18            udtUser.strCert = strSigCert
19            udtUser.strEncCert = ""
20            udtUser.strCertID = strCertID
21            udtUser.strPicCode = "" '证书注册时不更新图片,图片更新可通过人员管理【更新签名图片】功能完成。
22            udtUser.strPicPath = ""
              '获取已经注册证书的有效结束日期
23            strDate = mobjSDCA.SOF_GetCertInfo(mUserInfo.strCert, 18)
24            If IsUpdateRegCert(udtUser, strDate, blnReDo) Then
25                blnRet = True
26            Else
27                blnRet = False
28            End If
29        End If
          
30        SDCA_CheckCert = blnRet

31        Exit Function

ErrH:
32        MsgBox "在zl9ESign.mdlSDCA.SDCA_CheckCert的第" & Erl() & "行出错：" & vbCrLf & _
                  "错误号: " & Err.Number & vbCrLf & _
                  "错误描述：" & Err.Description, vbExclamation, gstrSysName


End Function


Public Function SDCA_GetPara() As Boolean
      '获取参数设置

1         On Error GoTo ErrH

2         With gudtPara
3             .strTSIP = GetThirdPara(CON_PAR_山东, "时间戳地址")
4             .strTSPort = GetThirdPara(CON_PAR_山东, "时间戳端口")
5             .strSIGNIP = GetThirdPara(CON_PAR_山东, "签名地址")
6             .strSignPort = GetThirdPara(CON_PAR_山东, "签名端口")
7             .bytSignVersion = Val(GetThirdPara(CON_PAR_山东, "版本"))
8             If .bytSignVersion = 1 Then
9                 If .strSIGNIP = "" Or .strTSIP = "" Then Exit Function
10            End If
11        End With
          
12        SDCA_GetPara = True
13        Exit Function

ErrH:
14        MsgBox "在zl9ESign.mdlSDCA.SDCA_GetPara的第" & Erl() & "行出错：" & vbCrLf & _
                  "错误号: " & Err.Number & vbCrLf & _
                  "错误描述：" & Err.Description, vbExclamation, gstrSysName

       
End Function

Public Function SDCA_InitObj() As Boolean
          Dim strMsg As String
          
1         On Error Resume Next
2         Set mobjSDCA = CreateObject("SDCASecurityClient.CASecurityClient.1")
3         If Err.Number <> 0 Then
              'C:\Windows\system32\SDCASecurityClient.dll
4             strMsg = "创建签名对象【SDCASecurityClient.CASecurityClient.1】失败！请检查部件【SDCASecurityClient.dll】是否正确安装并注册。"
5             GoTo ErrH
6         End If
7         On Error GoTo ErrH
          
8         On Error Resume Next
9         Set mobjSVS = CreateObject("NetONEX.SVSClientX.1")
10        If Err.Number <> 0 Then
              'C:\Windows\system32\NetONEX.dll
11            strMsg = "创建验签对象【NetONEX.SVSClientX.1】失败！请检查部件【NetONEX.dll】是否正确安装并注册。"
12            GoTo ErrH
13        End If
14        On Error GoTo ErrH
          '外网地址："60.216.5.244" 端口 9189
15        mobjSVS.ServerAddress = gudtPara.strSIGNIP
16        mobjSVS.ServerPort = Val(gudtPara.strSignPort)
17        On Error Resume Next
18        Set mobjTSA = CreateObject("NetONEX.TSAClientX.1")
19        If Err.Number <> 0 Then
              'C:\Windows\system32\NetONEX.dll
20            strMsg = "创建时间戳对象【NetONEX.TSAClientX.1】失败！请检查部件【NetONEX.dll】是否正确安装并注册。"
21            GoTo ErrH
22        End If
23        On Error GoTo ErrH
          '外网地址："60.216.5.244" 端口 9198
24        mobjTSA.ServerAddress = gudtPara.strTSIP
25        mobjTSA.ServerPort = Val(gudtPara.strTSPort)
          
26        On Error Resume Next
27        Set mobjBase64 = CreateObject("NetONEX.Base64X.1")
28        If Err.Number <> 0 Then
               'C:\Windows\system32\NetONEX.dll
29            strMsg = "创建验签对象【NetONEX.Base64X.1】失败！请检查部件【NetONEX.dll】是否正确安装并注册。"
30            GoTo ErrH
31        End If
32        On Error GoTo ErrH
          
33        SDCA_InitObj = True
34        Exit Function
ErrH:
35       Call GetErrMsg(Erl(), strMsg)
End Function

Public Function SDCA_RegCert(arrCertInfo As Variant) As Boolean
'功能:       提供在HIS数据库中注册个人证书的必要信息 , 用于新发放或更换证书, , 需要插入USB - Key
'返回:       arrCertInfo作为数组返回证书相关信息
'            0-ClientSignCertCN:客户端签名证书公共名称(姓名),用于注册证书时程序验证身份
'            1-ClientSignCertDN:客户端签名证书主题(每人唯一)
'            2-ClientSignCertSN:客户端签名证书序列号(每证书唯一)
'            3-ClientSignCert:客户端签名证书内容
'            4-ClientEncCert:客户端加密证书内容
'            5-签名图片文件名,空串表示没有签名图片
'            6-时间戳证书
          Dim strSN As String, strName As String, strDn As String
          Dim strCert As String
          Dim strPic As String
          
          Dim i As Long
            
1         On Error GoTo ErrH

2         For i = LBound(arrCertInfo) To UBound(arrCertInfo)
3             arrCertInfo(i) = ""
4         Next
5         If GetCertList(strName, strSN, strCert, strDn, , , strPic) Then
6             arrCertInfo(0) = strName
7             arrCertInfo(1) = strDn
8             arrCertInfo(2) = strSN
9             arrCertInfo(3) = strCert
10            arrCertInfo(4) = ""
11            arrCertInfo(5) = strPic
12            arrCertInfo(6) = ""
13            SDCA_RegCert = True
14        End If

15        Exit Function

ErrH:
16        MsgBox "在zl9ESign.mdlSDCA.SDCA_RegCert的第" & Erl() & "行出错：" & vbCrLf & _
                  "错误号: " & Err.Number & vbCrLf & _
                  "错误描述：" & Err.Description, vbExclamation, gstrSysName

End Function

Public Sub SDCA_SetPara()
          
1         On Error GoTo ErrH

2         With gudtPara
3             Call UpdateThirdPara(CON_PAR_山东, 1, "签名地址", .strSIGNIP, "签名服务地址")
4             Call UpdateThirdPara(CON_PAR_山东, 2, "签名端口", .strSignPort, "签名服务端口")
5             Call UpdateThirdPara(CON_PAR_山东, 3, "时间戳地址", .strTSIP, "时间戳服务地址")
6             Call UpdateThirdPara(CON_PAR_山东, 4, "时间戳端口", .strTSPort, "时间戳服务端口")
7             Call UpdateThirdPara(CON_PAR_山东, 5, "版本", .bytSignVersion, "接口版本")
8         End With

9         Exit Sub

ErrH:
10        MsgBox "在zl9ESign.mdlSDCA.SDCA_SetPara的第" & Erl() & "行出错：" & vbCrLf & _
                  "错误号: " & Err.Number & vbCrLf & _
                  "错误描述：" & Err.Description, vbExclamation, gstrSysName
End Sub

Public Function SDCA_Sign(ByVal strSource As String, ByRef strSignData As String, ByRef strTimeStamp As String, ByRef strTimeStampCode As String, ByRef blnReDo As Boolean)
      '功能:签名
          Dim strCertID As String
          Dim blnCheck As Boolean
          
          Dim objTsaRes As Object
          Dim objCert As Object

1         On Error GoTo ErrH
          
2         If Not SDCA_CheckCert(blnReDo, strCertID) Then Exit Function
3         If blnReDo Then Exit Function
          
          '适用于国密库SM2算法证书，默认0 密码校验一次
4         Call mobjSDCA.SOF_InitialVar("CERTID", "KeyLoginPolicy", "0")
          '获取签名值
5         strSignData = mobjSDCA.SOF_SignData(strCertID, strSource)
6         If strSignData = "" Then Exit Function
          
          '获取时间戳请求（时间戳对象初始化的时候已经实例化）
7         Set objTsaRes = mobjTSA.TSACreate(strSource)
                  
          '获取时间戳编码
8         strTimeStampCode = objTsaRes.ToBASE64
          
9         If strTimeStampCode = "" Then
10            MsgBoxEx "获取时间戳信息失败！", vbInformation, gstrSysName
11            Exit Function
12        End If
          'Oct 13 10:14:47 2019 GMT
13        strTimeStamp = objTsaRes.TimeStamp
14        LogWrite "SDCA_Sign", "接口【TimeStamp】返回值:" & strTimeStamp
15        If strTimeStamp = "" Then
16            MsgBoxEx "从时间戳签名值中获取签名时间失败！", vbInformation, gstrSysName
17            Exit Function
18        End If
19        strTimeStamp = GetTimeStamp(strTimeStamp)
20        strTimeStamp = Format(strTimeStamp, "yyyy-MM-dd HH:mm:ss")
       
21        SDCA_Sign = True

22        Exit Function

ErrH:
23        MsgBox "在zl9ESign.mdlSDCA.SDCA_Sign的第" & Erl() & "行出错：" & vbCrLf & _
                  "错误号: " & Err.Number & vbCrLf & _
                  "错误描述：" & Err.Description, vbExclamation, gstrSysName


End Function

Public Sub SDCA_UnloadObj()
    Set mobjSDCA = Nothing
    Set mobjSVS = Nothing
    Set mobjTSA = Nothing
    Set mobjBase64 = Nothing
End Sub

Public Function SDCA_VerifySign(ByVal strCert As String, ByVal strSign As String, ByVal strSource As String, ByVal strTStampCode As String) As Boolean
          '功能:验证签名
          '时间戳验证签名
          Dim lngRet As Long
          Dim strEncode As String
          
          '验证签名及证书有效性,返回值为200 则正确,其余为错误
1         On Error GoTo ErrH
2         strEncode = mobjBase64.EncodeString(strSource)
3         lngRet = mobjSVS.SVSVerifyPKCS1(strCert, strSign, strEncode)
4         If lngRet <> 200 Then
5             MsgBoxEx "签名验签失败！", vbInformation, gstrSysName
6             Exit Function
7         End If
        
8         lngRet = mobjTSA.TSAVerify(strTStampCode)
9         If lngRet <> 200 Then
10            MsgBoxEx "时间戳验证失败！", vbInformation, gstrSysName
11            Exit Function
12        End If
13        MsgBoxEx "验签成功！", vbInformation, gstrSysName
14        SDCA_VerifySign = True
          
15        Exit Function
ErrH:
16        MsgBox "在zl9ESign.mdlSDCA.SDCA_VerifySign的第" & Erl() & "行出错：" & vbCrLf & _
                  "错误号: " & Err.Number & vbCrLf & _
                  "错误描述：" & Err.Description, vbExclamation, gstrSysName

End Function

Private Function GetCertList(Optional ByRef strName As String = "1", Optional ByRef strCertSn As String = "1", _
   Optional ByRef strCert As String, Optional ByRef strCertDN As String = "1", Optional ByRef strUniqueID As String = "1", _
   Optional ByRef strCertID As String, Optional ByRef strPic As String = "1") As Boolean
          'strPic-文件路径
          Dim lngRet As Long
          
1         On Error GoTo ErrH

2         strCertID = mobjSDCA.SOF_GetUserList()
3         If strCertID <> "" Then
4             strCert = mobjSDCA.SOF_ExportUserCert(strCertID) '证书字符串
5             If strCert <> "" Then
6                 If strCertSn <> "1" Then strCertSn = mobjSDCA.SOF_GetCertInfo(strCert, 2) '证书序列号
7                 If strName <> "1" Then strName = mobjSDCA.SOF_GetCertInfo(strCert, 23) '证书通用者名称
8                 If strCertDN <> "1" Then strCertDN = mobjSDCA.SOF_GetCertInfo(strCert, 33) '证书拥有者DN
9                 If strUniqueID <> "1" Then
                    strUniqueID = mobjSDCA.SOF_GetCertInfo(strCert, 53) '唯一标识
                    strUniqueID = Right(strUniqueID, 18)
                  End If
10            End If
11            If strPic <> "1" Then
12                strPic = App.Path & "\pic.bmp"
13                lngRet = mobjSDCA.SOF_ShowSeal(strCertID, 0, strPic, 3)
14                If lngRet = 0 Then strPic = "" '读取失败
15            End If
16            GetCertList = True
17        Else
18            MsgBoxEx "没有找到Key盘，请检查！", vbInformation, gstrSysName
19            Exit Function
20        End If

21        Exit Function

ErrH:
22        MsgBox "在zl9ESign.mdlSDCA.GetCertList的第" & Erl() & "行出错：" & vbCrLf & _
                  "错误号: " & Err.Number & vbCrLf & _
                  "错误描述：" & Err.Description, vbExclamation, gstrSysName
End Function

Private Function GetTimeStamp(ByVal strTimeStamp As String) As String
      '功能：获取时间戳中的时间
          Dim arrTime As Variant
          Dim strTime As String
          
1         On Error GoTo ErrH

2         strTimeStamp = Replace(strTimeStamp, " ", "|")
3         strTimeStamp = Replace(strTimeStamp, "||", "|") '当日期为一位数时,防止月份和日期之间存在两个空格的情况
4         arrTime = Split(strTimeStamp, "|")  '传人格式：Aug 19 13:07:25 2014 GMT
5         strTime = arrTime(0) & " " & arrTime(1) & " " & arrTime(3)  '月/日/年
6         strTime = CDate(strTime) & ""  '年 月 日  2014/8/19
7         GetTimeStamp = strTime & " " & arrTime(2)  ' 年-月-日 时:分:秒

8         Exit Function

ErrH:
9         MsgBox "在zl9ESign.mdlSDCA.GetTimeStamp的第" & Erl() & "行出错：" & vbCrLf & _
                  "错误号: " & Err.Number & vbCrLf & _
                  "错误描述：" & Err.Description, vbExclamation, gstrSysName

End Function
