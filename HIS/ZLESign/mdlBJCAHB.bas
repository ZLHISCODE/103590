Attribute VB_Name = "mdlBJCAHB"
Option Explicit

'北京CA中心功能模块(新湖北版)
Private mblnInit As Boolean         '是否已初始化成功

Private BJCA_Pic As Object
Private BJCAHB_Client As Object       '客户端证书部件
Private BJCAHB_svs As Object          '签名验证控件
Private BJCAHB_TS As Object           '时间戳控件

Public Function BJCAHB_InitObj() As Boolean
    '证书部件初始化
        Dim progID As String
        On Error GoTo errH
     BJCAHB_InitObj = mblnInit

     If mblnInit Then Exit Function
100     Set BJCAHB_Client = CreateObject("XTXAppCOM.XTXApp.1")
101     Set BJCA_Pic = CreateObject("GetKeyPic.GetPic")
102     Set BJCAHB_svs = CreateObject("BJCA_SVS_ClientCOM.BJCASVSEngine.1")   '创建证书验证控件
103     Set BJCAHB_TS = CreateObject("BJCA_TS_ClientCom.BJCATSEngine")    '创建时间戳控件
     BJCAHB_InitObj = True
     mblnInit = BJCAHB_InitObj
    Exit Function
errH:
     MsgBoxEx "创建接口部件失败！" & vbNewLine & Err.Description, vbQuestion, gstrSysName
End Function
Public Function BJCAHB_RegCert(arrCertInfo As Variant) As Boolean
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
        Dim strPicData As String
        On Error GoTo errH
    
        For i = LBound(arrCertInfo) To UBound(arrCertInfo)
             arrCertInfo(i) = ""
        Next
        
        If GetCertList(strCertUserName, strKeyId, strSigCert) Then
            arrCertInfo(0) = strCertUserName
            arrCertInfo(1) = BJCAHB_Client.SOF_GetCertInfoByOid(strSigCert, "1.2.156.112562.2.1.1.1") '2.获取证书唯一标识（一般为身份证号）
            'arrCertInfo(1) = BJCAHB_Client.SOF_GetCertInfo(strSigCert, 33)
            arrCertInfo(2) = BJCAHB_Client.SOF_GetCertInfo(strSigCert, 2)
            arrCertInfo(3) = strSigCert
            arrCertInfo(5) = SaveBase64ToFile("gif", strKeyId, BJCA_Pic.getpic())
            BJCAHB_RegCert = True
        End If

        Exit Function
errH:
     MsgBoxEx "获取证书信息失败！" & vbNewLine & "第" & CStr(Erl()) & "行 " & Err.Description, vbQuestion, gstrSysName
End Function

Public Function BJCAHB_CheckCert(ByVal strCurrCertSn As String, Optional ByRef strSigCert As String, Optional ByRef CertID As String) As Boolean
    '功能：读取USB进行设备初始化并登录
     Dim strKey As String, strPIN As String, strUserName As String
     Dim strWebUrl As String, intDate   As Integer
     Dim random
     Dim strClientSignedData, strUsbkeyList As String, strUniqueID As String
     Dim arrUserList() As String
     On Error GoTo errH
     If Not BJCAHB_InitObj Then
        MsgBoxEx "检查部件是否初始化！", vbInformation + vbOKOnly, gstrSysName
        Exit Function
     End If
      '获取证书
     strUsbkeyList = BJCAHB_Client.SOF_GetUserList()
     If (strUsbkeyList = "") Then
        MsgBoxEx "请插入证书Key！"
        BJCAHB_CheckCert = False
        Exit Function
     Else
        arrUserList = Split(strUsbkeyList, "&&&")
        arrUserList = Split(arrUserList(0), "||")
        CertID = arrUserList(1)
        strSigCert = BJCAHB_Client.SOF_ExportUserCert(arrUserList(1)) '3.导出签名证书。
        strUniqueID = BJCAHB_Client.SOF_GetCertInfoByOid(strSigCert, "1.2.156.112562.2.1.1.1") '2.获取证书唯一标识（一般为身份证号）
     End If
     If strCurrCertSn <> BJCAHB_Client.SOF_GetCertInfo(strSigCert, 2) Then
        MsgBoxEx "该证书未注册在您的名下，不能使用！"
        Exit Function
     End If
     random = BJCAHB_Client.SOF_GenRandom(24)
     strClientSignedData = BJCAHB_Client.SOF_SignData(CertID, random)
     If Not GetCertLogin(strUniqueID, strClientSignedData, strSigCert, intDate, strWebUrl) Then
         BJCAHB_CheckCert = False
     Else
         BJCAHB_CheckCert = True
     End If
    Exit Function
errH:
     MsgBoxEx "检查USBKEY失败！" & vbNewLine & "第" & CStr(Erl()) & "行 " & Err.Description, vbQuestion, gstrSysName
End Function

Public Function BJCAHB_Sign(ByVal strCurrCertSn As String, ByVal strSource As String, ByRef strSignData As String, _
    ByRef strTimeStamp As String, ByRef strTimeStampCode As String) As Boolean
        '签名
    Dim strSigCert As String
    Dim CertID As String
    Dim strRequest As String    '时间戳请求
    Dim strErr As String
    Dim strDate As String
    
    If BJCAHB_TS Is Nothing Then Set BJCAHB_TS = CreateObject("BJCA_TS_ClientCom.BJCATSEngine")  '创建时间戳控件
    If Err.Number <> 0 Then
        MsgBoxEx "时间戳控件没有安装！", vbExclamation, gstrSysName
        Exit Function
    End If
    On Error GoTo errH
    If BJCAHB_CheckCert(strCurrCertSn, strSigCert, CertID) Then                '验证当前USB是否是签名用户的，并获取签名证书
        strSignData = BJCAHB_Client.SOF_SignData(CertID, strSource) '产生签名数据
        If strSignData <> "" Then
            strRequest = BJCAHB_TS.CreateTimeStampRequest(strSignData) '产生时间戳请求
            If strRequest <> "" Then
                strTimeStampCode = BJCAHB_TS.CreateTimeStamp(strRequest)  '产生时间戳（带证书）
                If strTimeStampCode = "" Then
                    strErr = "时间戳不能为空！"
                Else
                    strDate = BJCAHB_TS.gettimestampinfo(strTimeStampCode, 1)
                    strTimeStamp = String14ToDate(strDate, strErr)   '取得时间戳时间
                End If
            Else
                strErr = "时间戳请求失败！"
            End If
        Else
            strErr = "签名失败！"
        End If
    Else
        strErr = "验证证书失败！"
    End If
    
    If strErr <> "" Then
        MsgBoxEx strErr, vbOKOnly + vbInformation, gstrSysName
        BJCAHB_Sign = False
        Exit Function
    End If
    BJCAHB_Sign = True
    Exit Function
errH:
     MsgBoxEx "签名失败！" & vbNewLine & "第" & CStr(Erl()) & "行 " & Err.Description, vbQuestion, gstrSysName
End Function
'''验证签名函数
Public Function BJCAHB_VerifySign(ByVal strCert As String, _
    ByVal strSignData As String, ByVal strSource As String, ByVal strStampCode As String) As Boolean
'验证签名
    Dim blnRet As Boolean
    Dim strMsg As String
    Dim lngRuslt As Long
    
    On Error GoTo errH
    
    If strStampCode <> "" Then
        lngRuslt = BJCAHB_TS.verifyTimeStamp(strStampCode)
        If lngRuslt <> 0 Then
            MsgBoxEx "验证时间戳失败！" & GetReturnInfo(lngRuslt), vbInformation + vbOKOnly, gstrSysName
            Exit Function
        End If
    End If
    blnRet = BJCAHB_Client.SOF_VerifySignedData(strCert, strSource, strSignData)
    If blnRet Then
        strMsg = "验证签名成功！"
    Else
        strMsg = "验证签名失败！"
    End If
    MsgBoxEx strMsg, vbOKOnly + vbInformation, gstrSysName
    BJCAHB_VerifySign = blnRet
    Exit Function
errH:
104     MsgBoxEx "验证签名失败！" & vbNewLine & "第" & CStr(Erl()) & "行 " & Err.Description, vbQuestion, gstrSysName
End Function
'销毁对象
Public Sub BJCAHB_UloadObj()
    Set BJCAHB_Client = Nothing
    Set BJCAHB_svs = Nothing
    Set BJCAHB_TS = Nothing
    mblnInit = False
End Sub

'----- 以下是内部函数
Private Function GetCertLogin(ByVal strUniqueID As String, ByVal strClientSignedData As String, ByVal strCert As String, ByRef dDate As Integer, ByRef strWebserviceUrl As String) As Boolean
    '北京CA湖北版数字证书登录函数
    '- 入参
    'strUniqueID            :证书唯一标识
    'strClientSignedData    :签名数据
    'strWebserviceUrl       :签名服务器地址，即为证书验证
    '- 出参
    'dDate       : 返回证书有效时间

    If BJCAHB_svs Is Nothing Then Set BJCAHB_svs = CreateObject("BJCA_SVS_ClientCOM.BJCASVSEngine.1")
    '证书安全登录
    'strClientSignedData:非空:成功
    'strClientSignedData:空:不成功
    If (strClientSignedData <> "") Then
        '服务器端验证证书
        '从组件中导出证书
        Dim retValidateCert As Long
        retValidateCert = ValidateCert(strCert)
        '验证证书结果信息表示
        If retValidateCert <> 0 Then
            Call ValidateCertView(retValidateCert)
            Exit Function
        ElseIf (retValidateCert = 0) Then
            Dim s As String
            '获取客户端证书有效期截止时间
            s = BJCAHB_Client.SOF_GetCertInfo(strCert, 12)
            s = String14ToDate(s)
            If s <> "" Then
            '验证客户端证书有效期剩余天数
                dDate = CheckValidaty(CDate(s))
            
                If (dDate <= 30 And dDate > 0 And Not gblnShow) Then
                    MsgBoxEx "您的证书还有" & dDate & "天过期"
                    gblnShow = True
                    GetCertLogin = True
                ElseIf (dDate <= 0) Then
                    MsgBoxEx "您的证书已过期 " & Abs(dDate) & " 天"
                    GetCertLogin = False
                Else
                    GetCertLogin = True
                End If
            End If
        End If
        
    End If
End Function
Private Function ValidateCert(ByRef userCert As String) As Integer
    '服务器端验证证书
    If BJCAHB_svs Is Nothing Then Set BJCAHB_svs = CreateObject("BJCA_SVS_ClientCOM.BJCASVSEngine.1")
    ValidateCert = BJCAHB_svs.ValidateCertificate(userCert)
 
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
        Case Else
            MsgBoxEx "验证服务器证书失败！验证返回值:" & retValidateCert
    End Select
End Sub

''' 获取客户端证书列表
''' 返回boolean
Private Function GetCertList(ByRef strName As String, ByRef strUniqueID As String, ByRef strCert As String) As Boolean
    '北京CA湖北版获取数字证书列表函数
    '-入参:无
    '-出参
    'strName :      保存接口返回的证书所有者姓名
    'strUniqueID:   保存接口返回的证书所有者唯一标识
    'strCert:       保存接口返回的签名证书
    Dim strUsbkeyList As String
    Dim arrUserListLength As Integer
    Dim arrUserList() As String
      '获取证书
    strUsbkeyList = BJCAHB_Client.SOF_GetUserList()
    If (strUsbkeyList = "") Then
        strName = ""
        MsgBoxEx "请插入证书Key！"
        GetCertList = False
        Exit Function
    Else
        arrUserList = Split(strUsbkeyList, "&&&")
        arrUserList = Split(arrUserList(0), "||")
        strName = arrUserList(0)
        strCert = BJCAHB_Client.SOF_ExportUserCert(arrUserList(1)) '3.导出签名证书。
        strUniqueID = BJCAHB_Client.SOF_GetCertInfoByOid(strCert, "1.2.156.112562.2.1.1.1") '2.获取证书唯一标识（一般为身份证号）
    End If
    GetCertList = True
End Function

''' 检查证书有效性
''' 返回证书有效期天数
Private Function CheckValidaty(ByVal endDate As Date) As Integer
    '北京CA湖北版检查证书有效性接口
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
        GetReturnInfo = "未知错误"
    End If
    If GetReturnInfo <> "" Then
        GetReturnInfo = "时间戳接口返回提示：" & GetReturnInfo
    End If
End Function



