Attribute VB_Name = "mdlGDCA"
Option Explicit
Private mblnInit As Boolean         '是否已初始化成功
Private mobjGDCA As Object          '海南CA 控件
Private mLastPIN As String          '缓存输入的密码

Public Function GDCA_initObj() As Boolean
    
        '功能： 创建接口部件
        On Error GoTo errH
100     mLastPIN = ""
102     GDCA_initObj = mblnInit
104     If mblnInit Then Exit Function
    
        '请现场调试时如对象名称不对，请修改
106     Set mobjGDCA = CreateObject("Atl_com.Gdca.1")    '因为文档中没有说明，所以名称不确定
108     GDCA_initObj = True
    
110     mblnInit = GDCA_initObj
        Exit Function
errH:
112     MsgBoxEx "调用初始化接口失败！" & vbNewLine & Err.Description, vbQuestion, gstrSysName
End Function


Public Function GDCA_RegCert(arrCertInfo As Variant) As Boolean
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
        On Error GoTo errH
    
100     For i = LBound(arrCertInfo) To UBound(arrCertInfo)
102         arrCertInfo(i) = ""
        Next
    
104     If ReadUSBKeyLogin Then
106         If VerifyCert(strKeyId, strCertDN, strCertUserName, strSigCert) Then
108             arrCertInfo(0) = strCertUserName
110             arrCertInfo(1) = strCertDN
112             arrCertInfo(2) = strKeyId
114             arrCertInfo(3) = strSigCert
116             GDCA_RegCert = True
            End If
        End If

        Exit Function
errH:
118     MsgBoxEx "获取证书信息失败！" & vbNewLine & "第" & CStr(Erl()) & "行 " & Err.Description, vbQuestion, gstrSysName
End Function

Public Function GDCA_Sign(ByVal strCurrCertSn As String, ByVal strSource As String, ByRef strSignData As String, ByRef strTimeStamp As String) As Boolean
        '签名
        Dim strTmp As String, strEndData As String, strSigCert As String
        On Error GoTo errH
100     If GDCA_CheckCert(strCurrCertSn, strSigCert) Then       '验证当前USB是否是签名用户的，并获取签名证书
102         'strTmp = mobjGDCA.AppGetTime()      '获取时间戳,格式未知
104         'If IsDate(strTmp) Then strTimeStamp = strTmp        '是日期格式才返回
        
106         strEndData = strSource '& strTmp                        '暂时不加上时间戳
108         strEndData = mobjGDCA.GDCA_Base64Encode(strEndData)     '对原始数据编码
        
110         strSignData = mobjGDCA.GDCA_Pkcs7Sign("LAB_USERCERT_SIG", 4, strSigCert, strEndData)    '签名
        
112         GDCA_Sign = True
        End If
        Exit Function
errH:
114     MsgBoxEx "签名失败！" & vbNewLine & "第" & CStr(Erl()) & "行 " & Err.Description, vbQuestion, gstrSysName
End Function

Public Function GDCA_VerifySign(ByVal strCurrCertSn As String, ByVal strSignData As String, ByVal strSource As String) As Boolean
        '验证签名
        Dim strSigCert As String, strTmp As String, strEndData
        On Error GoTo errH
100     If GDCA_CheckCert(strCurrCertSn, strSigCert) Then         '验证当前USB是否是签名用户的，并获取签名证书
            strEndData = mobjGDCA.GDCA_Base64Encode(strSource)
102         strTmp = mobjGDCA.GDCA_Pkcs7Verify(strSigCert, strSignData)
104         If strTmp = strEndData Then GDCA_VerifySign = True
        End If
        Exit Function
errH:
106     MsgBoxEx "验证签名失败！" & vbNewLine & "第" & CStr(Erl()) & "行 " & Err.Description, vbQuestion, gstrSysName
End Function

Public Function GDCA_CheckCert(ByVal strCurrCertSn As String, Optional ByRef strSigCert As String) As Boolean
    
        '验证当前插入的USBKey是否是当前用户的,如果成功，则返回签名证书
        Dim strKeyId As String, strCertTime As String, strCertUserName As String, strCertDN As String
        Dim i As Integer, strCACert As String, lngOk As Long
    
        On Error GoTo errH
    
100     GDCA_CheckCert = False

102     If ReadUSBKeyLogin Then
104         If VerifyCert(strKeyId, strCertDN, strCertUserName, strSigCert) Then
106             If strCurrCertSn = strKeyId Then GDCA_CheckCert = True
            End If
        End If

        Exit Function
errH:
108     MsgBoxEx "获取证书信息失败！" & vbNewLine & "第" & CStr(Erl()) & "行 " & Err.Description, vbQuestion, gstrSysName
End Function
Public Sub GDCA_UnloadObj()
    Set mobjGDCA = Nothing
End Sub
'------------------------------------------------------------------------------------
'-- 以下是GDCA模块内部调用过程
'------------------------------------------------------------------------------------
Private Function ReadUSBKeyLogin() As Boolean
        '功能：读取USB进行设备初始化并登录
        Dim strKey As String, strPIN As String, lngOk As Long
        Dim blnOk As Boolean, lngCount As Long
        On Error GoTo errH
100     If Not mblnInit Then
102         MsgBoxEx "GDCA部件未初始化！"
            Exit Function
        End If
    
104     strKey = mobjGDCA.GDCA_GetDevicType             '取插入本机的设备类型
106     Call mobjGDCA.GDCA_SetDeviceType(strKey)        '根据KEY设置设备
108     Call mobjGDCA.GDCA_Initialize                   '初始化
    
110     If mLastPIN <> "" Then strPIN = mLastPIN
112     If strPIN = "" Then
114         If Not frmPassword.ShowMe(strPIN) Then Exit Function
        End If
        
        '用户PIN为1-8个字节，是用户输入的密码
116     lngCount = 0: blnOk = False: lngOk = -1
    
118     Do While lngCount <= 3 And Not blnOk
120         lngOk = mobjGDCA.GDCA_Login(2, strPIN)                                    '登录
122         If lngOk = 0 Then
124             blnOk = True
            Else
126             If Not frmPassword.ShowMe(strPIN) Then Exit Function
            End If
128         lngCount = lngCount + 1
        Loop
    
130     mLastPIN = strPIN
132     ReadUSBKeyLogin = True
    
        Exit Function
errH:
134     MsgBoxEx "初始化KEY失败！" & vbNewLine & "第" & CStr(Erl()) & "行 " & Err.Description, vbQuestion, gstrSysName
End Function


Private Function VerifyCert(Optional CertSN As String, Optional CertDn As String, Optional CertUN As String, Optional SigCert As String) As Boolean
        '验证签名，及获取证书数据
    
        Dim strKeyId As String, strCertTime As String, strCertUserName As String, strCertDN As String
        Dim strSigCert As String, i As Integer, strCACert As String, lngOk As Long
    
100     strKeyId = mobjGDCA.GDCA_ReadLabel("LAB_DISAID", 3)            ' 获取证书唯一标识
102     strKeyId = mobjGDCA.GDCA_Base64Decode(strKeyId)
        
104     strCACert = mobjGDCA.GDCA_ReadLabel("CA_CERT", 9)                   '9-获取CA证书
106     strSigCert = mobjGDCA.GDCA_ReadLabel("LAB_USERCERT_SIG", 7)         '7-签名证书
        
        
108     strCertTime = mobjGDCA.GDCA_GetCertificateInfo(strSigCert, 6)       '6-取证书有效期
110     strCertUserName = mobjGDCA.GDCA_GetCertificateInfo(strSigCert, 7)   '7-取证书持有者名称
112     strCertDN = mobjGDCA.GDCA_GetCertificateInfo(strSigCert, 3)         '3-取证书序列号
        
114     lngOk = -1
116     lngOk = mobjGDCA.GDCA_VerifyCert(strSigCert, strCACert)                 '验证证书
118     If lngOk = 0 Then
120         CertUN = strCertUserName
122         CertDn = strCertDN
124         CertSN = strKeyId
126         SigCert = strSigCert
128         VerifyCert = True
        Else
130         MsgBoxEx "证书验证失败！证书效期" & strCertTime, vbQuestion, gstrSysName
132         VerifyCert = False
        End If
        Exit Function
errH:
134     MsgBoxEx "验证证书失败！" & vbNewLine & "第" & CStr(Erl()) & "行 " & Err.Description, vbQuestion, gstrSysName

End Function

