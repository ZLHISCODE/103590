Attribute VB_Name = "mdlNMGCA"

Option Explicit
'包头市中心医院
Private mobjJLClient As Object   'JITComVCTKEx.JITVCTKEx.1
Private mobjJLServer As Object  'JITClientCOMAPI.JITClientProc.1
Private mobjCertInfo As Object    '签章 BicengEsealInterface.CEsealInterface
Private mblnInit As Boolean
'此字符串和驱动有关
Private Const M_STR_PARA_NM As String = "<?xml version=""1.0"" encoding=""utf-8""?>" & _
        "<authinfo><liblist>" & _
        "<lib type=""SKF"" version=""1.1"" dllname=""bXRva2VuX2dtMzAwMC5kbGw="" ><algid val=""SHA1"" sm2_hashalg=""sm3""/></lib>" & _
        "<lib type=""SKF"" version=""1.1"" dllname=""U21hcnRDVENBUEkuZGxs"" ><algid val=""SHA1"" sm2_hashalg=""sm3""/></lib>" & _
        "<lib type=""SKF"" version=""1.1"" dllname=""U0tGQVBJMjA1NDkuZGxs"" ><algid val=""SHA1"" sm2_hashalg=""sm3""/></lib>" & _
        "</liblist></authinfo>"

Public Function NMG_InitObj() As Boolean
     '证书部件初始化
    Dim lngRet As Long
    Dim strTSAIP As String
    Dim varTmp As Variant
    
100     If glngSign > 1 Then NMG_InitObj = True: Exit Function
        On Error Resume Next
102     If mobjJLClient Is Nothing Then
104         Set mobjJLClient = CreateObject("JITComVCTKEx.JITVCTKEx.1")
106         If Err.Number <> 0 Then
108             MsgBoxEx "创建签名对象【JITComVCTK_S.dll】失败！请检查该控件是否安装并注册。", vbInformation, gstrSysName
                Exit Function
            End If
        End If
110     Err.Clear
112     If mobjJLServer Is Nothing Then
114         Set mobjJLServer = CreateObject("JITClientCOMAPI.JITClientProc.1")
116         If Err.Number <> 0 Then
118             MsgBoxEx "创建签名对象【JITClientCOMAPI.dll】失败！请检查该控件是否安装并注册。", vbInformation, gstrSysName
                Exit Function
            End If
        End If
    
        On Error GoTo errH
    
120     gstrPara = gobjComLib.zlDatabase.GetPara(90000, glngSys, , "") '读取配置内容
        
122     If gstrPara = "" Then
124         Err.Raise -1, , "当前系统【" & glngSys & "】没有配置电子签名参数,请到启用电子签名接口处设置。"
            Exit Function
        End If
126     If UBound(Split(gstrPara, G_STR_SPLIT)) <> 4 Then
128         MsgBoxEx "电子签名参数值设置有误,请检查。" & vbCrLf & _
                "当前参数值:" & gstrPara & vbCrLf & _
                "正确格式:签名服务器IP&&&签名服务器端口&&&时间戳IP&&&时间戳端口", vbInformation, gstrSysName
            Exit Function
        Else
130         Call NMG_GetPara
        End If
        
        On Error Resume Next
132     If gudtPara.blnSignPic Then
134         If mobjCertInfo Is Nothing Then
136             Set mobjCertInfo = CreateObject("BicengEsealInterface.CEsealInterface")
138             If Err.Number <> 0 Then
140                 MsgBoxEx "创建签章对象【BicengEsealInterface.dll】失败！请检查该控件是否安装并注册。", vbInformation, gstrSysName
                    Exit Function
                End If
            End If
        End If
142     Err.Clear: On Error GoTo 0
        On Error GoTo errH
144     lngRet = mobjJLClient.Initialize(M_STR_PARA_NM)

146     If Not GetErrorInfo("Initialize") Then Exit Function
148     mblnInit = True
150     NMG_InitObj = True
        Exit Function
errH:
152  MsgBoxEx "创建接口部件失败！" & vbNewLine & "第" & CStr(Erl()) & "行 " & Err.Description, vbExclamation, gstrSysName
End Function

Public Function NMG_Sign(ByVal strSource As String, ByRef strSignData As String, ByRef strTimeStamp As String, _
    ByRef strTimeStampCode As String, Optional ByRef blnReDo As Boolean, Optional ByVal blnCheck As Boolean) As Boolean
        '签名
    Dim lngRet As Long
    Dim strErr As String
    Dim strHash As String
    

        On Error GoTo errH
        If Not blnCheck Then
100         blnCheck = NMG_CheckCert(blnReDo)
102         If blnReDo Then Exit Function
        End If
104     If blnCheck Then                 '验证当前USB是否是签名用户的，并获取签名证书
            '证书ID进行签名
106         strSignData = mobjJLClient.DetachSignStr("", strSource)      'DetachSignStr-不带原文签名;AttachSignStr-带原文签名
108         If Not GetErrorInfo("DetachSignStr") Then Exit Function
110         If strSignData <> "" Then
112             If Not ConnectToTsaServer() Then Exit Function
                Call mobjJLServer.SetAlgorithmEx("SM3", "")
114             strHash = StringSHA1(strSource)
116             strTimeStampCode = mobjJLServer.TsaSign("", 1, strHash)            '申请时间戳 传入签名值过长，签名时比较耗时
118             strTimeStamp = mobjJLServer.VerifyTsaSign(strTimeStampCode)
                lngRet = mobjJLServer.GetErrorCodeEx()
                WriteLog "时间戳错误码:" & lngRet
120             Call mobjJLServer.FinalizeServerConnectEx    '断开时间戳服务器连接
122             If strTimeStampCode = "" Then MsgBoxEx "获取时间戳失败！", vbInformation, gstrSysName: Exit Function
                '日期格式化
124             strTimeStamp = Mid(strTimeStamp, 1, 14)
126             strTimeStamp = String14ToDate(strTimeStamp, strErr)
128             If strErr <> "" Then MsgBoxEx strErr, vbInformation, gstrSysName: Exit Function
                '转东八区时间
130             strTimeStamp = Format(DateAdd("h", 8, strTimeStamp), "YYYY-MM-DD HH:MM:SS")
            Else
132             MsgBoxEx "签名失败！", vbInformation, gstrSysName
                Exit Function
            End If

        Else
134         MsgBoxEx "签名失败！", vbInformation, gstrSysName
            Exit Function
        End If
136     NMG_Sign = True
        Exit Function
errH:
138     MsgBoxEx "签名失败！" & vbNewLine & "第" & CStr(Erl()) & "行 " & Err.Description, vbExclamation, gstrSysName
End Function

Public Function NMG_VerifySign(ByVal strSource As String, ByVal strSignData As String, ByVal strTimeStampCode As String) As Boolean
    '功能;验证签名
    '参数:strSignData -签名值
     Dim blnRet As Boolean
     Dim lngRet As Long
     Dim strTS As String
     On Error GoTo errH

    '服务器验签
100 If Not ConnectToSignServer() Then Exit Function
102 lngRet = mobjJLServer.VerifyDetachedSign(strSignData, strSource)  '服务器验证数据 不带原文签名:VerifyDetachedSign(string, string);带原文签名  VerifyAttachedSign
104 If lngRet <> 0 Then
106     MsgBoxEx "签名验证失败:" & mobjJLServer.GetErrorMessage(lngRet), vbInformation, gstrSysName
108     Call mobjJLServer.FinalizeServerConnectEx   '断开签名服务器链接
        Exit Function
    End If
110 Call mobjJLServer.FinalizeServerConnectEx   '断开签名服务器链接


    '连接时间戳服务器
112 If Not ConnectToTsaServer() Then Exit Function
114 strTS = mobjJLServer.VerifyTsaSign(strTimeStampCode)
116 If strTS = "" Then
118       MsgBoxEx "时间戳验证失败！", vbInformation, gstrSysName
120       Call mobjJLServer.FinalizeServerConnectEx   '断开时间戳服务器
          Exit Function
    End If
122 Call mobjJLServer.FinalizeServerConnectEx   '断开时间戳服务器
 
124 MsgBoxEx "验证成功，该电子签名数据有效!", vbInformation, gstrSysName
    
126  NMG_VerifySign = True
     Exit Function
errH:
128     MsgBoxEx "验证签名失败！" & vbNewLine & "第" & CStr(Erl()) & "行 " & Err.Description, vbExclamation, gstrSysName
End Function


Public Function NMG_CheckCert(ByRef blnReDo As Boolean) As Boolean
        '功能：读取USB进行设备初始化并登录
        Dim strKeySN As String, strUserID As String, strUserName As String, strCertDN As String
        Dim strDate As String
        Dim arrDN As Variant
        Dim udtUser As USER_INFO
        Dim blnRet As Boolean
        Dim i As Long
    
        On Error GoTo errH
100     If Not GetCertList(strKeySN, strUserName, strCertDN) Then Exit Function
102     If mUserInfo.strCertSn <> strKeySN Then
104         MsgBoxEx "该证书未注册在您的名下，不能使用！" & vbCrLf & _
                "用户注册证书唯一标识:" & mUserInfo.strCertSn & vbCrLf & _
                "当前所选证书唯一标识:" & strKeySN, vbInformation + vbOKOnly, gstrSysName
            Exit Function
        End If
    
        '判断是否需要更新注册证书
106     udtUser.strName = strUserName
108     udtUser.strSignName = strUserName
110     udtUser.strUserID = strUserID '身份证号
112     udtUser.strCertSn = strKeySN
114     udtUser.strCertDN = strCertDN
116     udtUser.strCert = ""
118     udtUser.strEncCert = ""
120     udtUser.strCertID = ""
122     udtUser.strPicPath = ""
124     arrDN = Split(mUserInfo.strCertDN, ",")     'CN=王二小U3294, O=七台河人民医院, L=七台河市, S=黑龙江省, C=CN, 有效日期=
126     For i = 0 To UBound(arrDN)
128         If Trim(arrDN(i)) Like "有效日期*" Then
130             strDate = Trim(Split(arrDN(i), "=")(1))
                Exit For
            End If
        Next
132     If IsUpdateRegCert(udtUser, strDate, blnReDo) Then
134         blnRet = True
        Else
136         blnRet = False
        End If

138     NMG_CheckCert = blnRet
        Exit Function
errH:
140      MsgBoxEx "检查USBKEY失败！" & vbNewLine & "第" & CStr(Erl()) & "行 " & Err.Description, vbExclamation, gstrSysName
End Function


Public Function NMG_RegCert(arrCertInfo As Variant) As Boolean
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

100         For i = LBound(arrCertInfo) To UBound(arrCertInfo)
102              arrCertInfo(i) = ""
            Next

104         If GetCertList(strKeyId, strCertUserName, strCertDN, strPicPath) Then
106             arrCertInfo(0) = strCertUserName
108             arrCertInfo(1) = strCertDN
110             arrCertInfo(2) = strKeyId
112             arrCertInfo(5) = strPicPath
114             NMG_RegCert = True
            End If

            Exit Function
errH:
116      MsgBoxEx "获取证书信息失败！" & vbNewLine & "第" & CStr(Erl()) & "行 " & Err.Description, vbExclamation, gstrSysName
End Function

Private Function GetCertList(Optional ByRef strUniqueID As String = "-1", Optional ByRef strName As String = "-1", Optional ByRef strCertDN As String = "-1", _
    Optional ByRef strPicPath As String = "-1", Optional ByRef strUserID As String = "-1", Optional ByRef strDate As String = "-1") As Boolean
    '功能:获取证书详情
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
        On Error GoTo errH
        
100     lngRet = mobjJLClient.SetCertChooseType(1)
102     lngRet = mobjJLClient.SetCert("SC", "", "", "", "", "")
104     If Not GetErrorInfo("SetCert") Then Exit Function
106     strDate = mobjJLClient.GetCertInfo("SC", 6, "")     '有效日期
108     If IsDate(strDate) Then
            '检查证书是否过期
110         lngDay = CheckValidaty(strDate)
112         If (lngDay <= 30 And lngDay > 0 And Not gblnShow) Then
114             MsgBoxEx "您的证书还有" & lngDay & "天过期", vbInformation, gstrSysName
116             gblnShow = True
118         ElseIf (lngDay <= 0) Then
120             MsgBoxEx "您的证书已过期 " & Abs(lngDay) & " 天"
                Exit Function
            End If
        End If
    
122     If strUniqueID <> "-1" Then strUniqueID = mobjJLClient.GetCertInfo("SC", 2, "")     '证书序列号
124     If strCertDN <> "-1" Or strName <> "-1" Then
126         strCertDN = mobjJLClient.GetCertInfo("SC", 0, "") 'CN=王二小U3294, O=七台河人民医院, L=七台河市, S=黑龙江省, C=CN, 有效日期=
128         strName = mobjJLClient.GetCertInfo("SC", 9, "")     '用户名称
        End If
130     If gudtPara.blnSignPic Then
132         strPic = mobjCertInfo.SignSeal("数据", "测试123")
134         If Trim(strPic) = "" Then MsgBoxEx "获取签章失败!", vbInformation, gstrSysName: Exit Function
136         strPicPath = App.Path & "\" & Format(Now, "yyyyMMdd") & "_" & strUniqueID & ".bmp"
138         mobjCertInfo.getSealPicturePath strPic, strPicPath
140         If Dir(strPicPath) = "" Then MsgBoxEx "获取签章失败!", vbInformation, gstrSysName: Exit Function
        End If
142     GetCertList = True
        Exit Function
errH:
144     MsgBoxEx "获取证书信息失败！" & vbNewLine & "第" & CStr(Erl()) & "行 " & Err.Description, vbExclamation, gstrSysName
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

Public Function NMG_GetPara() As Boolean
        Dim arrList As Variant
    
        On Error GoTo errH
100     If gstrPara = "" Then gstrPara = gobjComLib.zlDatabase.GetPara(90000, glngSys, , "") '读取URLs 固定读取ZLHIS 系统默认100
102     If gstrPara = "" Then gstrPara = "110&&&175.17.252.155&&&8000&&&175.17.252.156&&&8000&&&0"   '签名服务器IP&&&签名服务器端口&&&时间戳IP&&&时间戳端口&&&启用签章
104     arrList = Split(gstrPara, "&&&")
106     If UBound(arrList) >= 4 Then
108         gudtPara.strSIGNIP = arrList(0)
110         gudtPara.strSignPort = arrList(1)
112         gudtPara.strTSIP = arrList(2)
114         gudtPara.strTSPort = arrList(3)
116         gudtPara.blnSignPic = Val(arrList(4) & "") = 1
        Else
118         gudtPara.strSIGNIP = "175.17.252.155"
120         gudtPara.strSignPort = "8000"
122         gudtPara.strTSIP = "175.17.252.156"
124         gudtPara.strTSPort = "8000"
126         gudtPara.blnSignPic = False
        End If
        Exit Function
errH:
128     MsgBoxEx "读取参数失败！" & vbNewLine & "第" & CStr(Erl()) & "行 " & Err.Description, vbInformation, gstrSysName
End Function

Public Function NMG_SetParaStr() As String
    With gudtPara
        NMG_SetParaStr = IIf(Trim(.strSIGNIP) = "", "175.17.252.155", .strSIGNIP) & G_STR_SPLIT & IIf(Trim(.strSignPort) = "", "8000", .strSignPort) & _
                G_STR_SPLIT & IIf(Trim(.strTSIP) = "", "175.17.252.156", .strTSIP) & G_STR_SPLIT & IIf(Trim(.strTSPort) = "", "8000", .strTSPort) & G_STR_SPLIT & IIf(.blnSignPic, 1, 0)
    End With
End Function

Public Sub NMG_UnLoadObj()
    On Error Resume Next
    Set mobjJLServer = Nothing
    Set mobjCertInfo = Nothing
    Call mobjJLClient.Finalize
    Set mobjJLClient = Nothing
End Sub





