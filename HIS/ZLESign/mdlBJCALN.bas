Attribute VB_Name = "mdlBJCALN"
Option Explicit

'北京CA中心(辽宁省版)
'有时间戳，有签名图片
Private mobjBJCALN As Object     '北京CA辽宁部件
Private mblnInit As Boolean      '是否已初始化
Private mintLogin As Integer     '输入密码次数
Private mstrLastPwd As String    '缓存上次输入密码
Private Const BASE64CHR As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/="

Public Function BJCALN_InitObj() As Boolean
    '证书部件初始化
    
        On Error GoTo errH
100     mstrLastPwd = ""
102     BJCALN_InitObj = mblnInit
104     If mblnInit Then Exit Function
    
106     Set mobjBJCALN = CreateObject("BJCALN.DLFY")
114     BJCALN_InitObj = True
    
116     mblnInit = BJCALN_InitObj
        mintLogin = 0
        Exit Function
errH:
118     MsgBoxEx CStr(Erl()) & "行，创建接口部件失败！" & vbNewLine & Err.Description, vbQuestion, gstrSysName
End Function
Public Function BJCALN_UloadObj()
    Set mobjBJCALN = Nothing
End Function
Public Function BJCALN_CheckCert(ByVal strCurrCertSn As String, Optional ByRef strSigCert As String) As Boolean
        '功能：读取USB进行设备初始化并登录
        Dim strKey As String, strPIN As String, strUserName As String
        Dim strWebUrl As String, intDate   As Integer
        On Error GoTo errH
100     If Not mblnInit Then
102         MsgBoxEx "部件未初始化！"
            Exit Function
        End If
    
104     Call GetCertList(strUserName, strKey, strSigCert)
106     If strCurrCertSn <> strKey Then
108         MsgBoxEx "该证书未注册在您的名下，不能使用！"
            Exit Function
        End If
110     If mstrLastPwd <> "" Then strPIN = mstrLastPwd
112     If strPIN = "" Then
114         If Not frmPassword.ShowMe(strPIN) Then Exit Function
        End If
        
116     If Not GetCertLogin(strKey, strPIN, strSigCert, intDate, strWebUrl) Then
118         strPIN = ""
        Else
            BJCALN_CheckCert = True
        End If
     
120     mstrLastPwd = strPIN
122
    
        Exit Function
errH:
124     MsgBoxEx "检查USBKEY失败！" & vbNewLine & "第" & CStr(Erl()) & "行 " & Err.Description, vbQuestion, gstrSysName
End Function

Public Function BJCALN_RegCert(arrCertInfo As Variant) As Boolean
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
    
100     For i = LBound(arrCertInfo) To UBound(arrCertInfo)
102         arrCertInfo(i) = ""
        Next
    
104     If GetCertList(strCertUserName, strKeyId, strSigCert) Then
            
            strCertDN = mobjBJCALN.GetCertInfo(strSigCert, 2)
106         arrCertInfo(0) = strCertUserName
108         arrCertInfo(1) = strCertDN
110         arrCertInfo(2) = strKeyId
112         arrCertInfo(3) = strSigCert

114
116             If UBound(arrCertInfo) >= 5 Then
118                 strPicData = mobjBJCALN.getpic()
120                 If strPicData <> "" Then
122                     arrCertInfo(5) = SaveBase64ToFile("gif", strCertDN, strPicData) '图片路径
                    End If
                End If
            
124         BJCALN_RegCert = True
        End If

        Exit Function
errH:
126     MsgBoxEx "获取证书信息失败！" & vbNewLine & "第" & CStr(Erl()) & "行 " & Err.Description, vbQuestion, gstrSysName

End Function

Public Function BJCALN_Sign(ByVal strCurrCertSn As String, ByVal strSource As String, ByRef strSignData As String, ByRef strTimeStamp As String) As Boolean
        '签名
        Dim strSigCert As String

        On Error GoTo errH
100     If BJCALN_CheckCert(strCurrCertSn, strSigCert) Then               '验证当前USB是否是签名用户的，并获取签名证书
110         strSignData = mobjBJCALN.SignData(strCurrCertSn, strSource)
112         BJCALN_Sign = True
        Else
            MsgBoxEx "签名失败！"
        End If
        Exit Function
errH:
114     MsgBoxEx "签名失败！" & vbNewLine & "第" & CStr(Erl()) & "行 " & Err.Description, vbQuestion, gstrSysName
End Function

Public Function BJCALN_VerifySign(ByVal strCurrCertSn As String, ByVal strSignData As String, ByVal strSource As String) As Boolean
        '验证签名
        Dim strSigCert As String, strTmp As String
        On Error GoTo errH
100     If BJCALN_CheckCert(strCurrCertSn, strSigCert) Then           '验证当前USB是否是签名用户的，并获取签名证书
102         BJCALN_VerifySign = GetCertVerifySign(strSignData, strSigCert, strSource, strTmp)
        Else
            MsgBoxEx "验证签名失败！"
        End If
        Exit Function
errH:
104     MsgBoxEx "验证签名失败！" & vbNewLine & "第" & CStr(Erl()) & "行 " & Err.Description, vbQuestion, gstrSysName
End Function

'Public Function BJCALN_ShowCert(ByVal strCurrCertSn As String)
'    '功能：显示证书
'    On Error GoTo errH
'    Call mobjBJCALN.ShowCert(strCurrCertSn)
'    Exit Function
'errH:
'    MsgboxEx "证书显示失败！" & vbNewLine & "第" & CStr(Erl()) & "行 " & Err.Description, vbQuestion, gstrSysName
'End Function

'------------------------------
'以下是私有函数

Private Function GetCertList(ByRef strName As String, ByRef strUniqueID As String, ByRef strCert As String) As Boolean
    '大医第一附属医院获取数字证书列表函数
    '-入参:无
    '-出参
    'strName :      保存接口返回的证书所有者姓名
    'strUniqueID:   保存接口返回的证书所有者唯一标识
    'strCert:       保存接口返回的签名证书
      
    Dim strUsbkeyList As String
    Dim arrUserListLength As Integer
    Dim arrUserList() As String
     
    strUsbkeyList = mobjBJCALN.getUserList()
    arrUserList = Split(strUsbkeyList, "&&&")
    arrUserListLength = UBound(arrUserList)
    If (arrUserListLength = -1) Then
        MsgBoxEx "请您插入Key！"
        Exit Function
    End If
    If (arrUserListLength <> 0) Then
        Dim i As Integer
        For i = 0 To arrUserListLength - 1
            Dim strOption As String
            strOption = arrUserList(i)
            strName = Split(strOption, "||")(0)
            strUniqueID = Split(strOption, "||")(1)
            strCert = mobjBJCALN.ExportUserCert(strUniqueID)
        Next
    End If
    GetCertList = True
End Function


Private Function GetCertLogin(ByVal strUniqueID As String, ByVal strPassword As String, ByVal strCert As String, ByRef dDate As Integer, ByRef strWebserviceUrl As String) As Boolean
    '大医第一附属医院数字证书登录函数
    '- 入参
    'strUniqueID : 证书唯一标识
    'strPassword : 证书密码
    'strWebserviceUrl:签名服务器地址，即为证书验证
    '- 出参
    'dDate       : 返回证书有效时间

    Dim result As Boolean
    If mobjBJCALN Is Nothing Then BJCALN_InitObj
    If (strPassword = "") Then
        MsgBoxEx "请输入证书密码！"
    Else
        '证书安全登录
        'result:0:成功
        'result:非0:不成功
        If mintLogin >= 8 Then
            MsgBoxEx "已经输入了" & mintLogin & "次错误密码，超过了最大输入次数！"
            Exit Function
        End If
        result = mobjBJCALN.userLogin(strUniqueID, strPassword)
        If (result) Then
            mintLogin = 0
'            Dim strExtLib As String
'            strExtLib = mobjBJCALN.GetUserInfo(strUniqueID, 15) '后面没有用
            Dim intFlg As Integer
            
            '服务器端验证证书
            '从组件中导出证书
            Dim retValidateCert As Long
            retValidateCert = 100
            retValidateCert = mobjBJCALN.ValidateCertificate(strCert)
             
            '验证证书结果信息表示
            If retValidateCert <> 0 Then Call ValidateCertView(retValidateCert)
            retValidateCert = 0 '先跳过服务端验证
            If (retValidateCert = 0) Then
                Dim uniqueIdStr As String
                Dim oid As String
                oid = "2.16.840.1.113732.2" '辽宁例子中用的是 "1.2.156.112562.2.1.1.4"
                Dim s As String
                '获取客户端证书有效期截止时间
                s = mobjBJCALN.GetCertInfo(strCert, 12)  '辽宁例子中用的是 GetCertInfo(strCert, 2)
                '验证客户端证书有效期剩余天数
                If s Like "##############" Then s = Mid(s, 1, 4) & "-" & Mid(s, 5, 2) & "-" & Mid(s, 7, 2) & " " & Mid(s, 9, 2) & ":" & Mid(s, 11, 2) & ":" & Mid(s, 13, 2)
                dDate = CheckValidaty(s)
            
                If (dDate <= 30 And dDate > 0 And Not gblnShow) Then
                    MsgBoxEx "您的证书还有" & dDate & "天过期"
                    uniqueIdStr = mobjBJCALN.GetCertInfoByOid(strCert, oid)
                    gblnShow = True
                    GetCertLogin = True
                ElseIf (dDate <= 0) Then
                    MsgBoxEx "您的证书已过期 " & Abs(dDate) & " 天"
                    GetCertLogin = False
                Else
                    uniqueIdStr = mobjBJCALN.GetCertInfoByOid(strCert, oid)
                    
                    GetCertLogin = True
                End If
            Else
                GetCertLogin = False
            End If
        Else
            mintLogin = mintLogin + 1
            MsgBoxEx "证书密码可能不正确，您已经输入了" & mintLogin & "次密码，还可以输入" & 8 - mintLogin & "次!"
            GetCertLogin = False
            
        End If
    End If

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
    End Select
End Sub

''' 客户端验证签名函数
''' 返回boolean值
Private Function GetCertVerifySign(ByVal strInData As String, ByVal strCert As String, ByRef strData As String, ByRef strOut As String) As Boolean
    '大医第一附属医院数字证书签名验证函数
    '- 入参
    'strInData     : 签名结果
    'strCert       : 签名证书
    'strData       : 签名原文
    '- 出参
    'strOut       : 返回验签结果

    'result:true:  成功
    'result:false: 失败
    Dim verifySignResult As Boolean
     
    verifySignResult = mobjBJCALN.VerifySignedData(strCert, strData, strInData)
    If (verifySignResult) Then
        strOut = "验证签名成功！"
        GetCertVerifySign = True
    Else
        strOut = "验证签名失败！"
        GetCertVerifySign = False
    End If
End Function

''' 检查证书有效性
''' 返回证书有效期天数
Private Function CheckValidaty(ByVal endDate As Date) As Integer
    '大医第一附属医院检查证书有效性接口
    '-入参: 证书有效截止日期
    '-出参：有效天数
        Dim dblAllSp    As Double
        Dim result      As Integer
        dblAllSp = CDbl(CDate(endDate)) - CDbl(Now)
        result = Int(dblAllSp)
        CheckValidaty = result
End Function





