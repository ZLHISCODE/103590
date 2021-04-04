Attribute VB_Name = "mdlGDWZT"
Option Explicit
'此接口签名值长度超过4000个字符,单独建表:电子签名记录(ID,签章信息)  ID字段对应各个业务表的签章信息中。不在ZLHIS中提交脚步
'Private mobjUtil As NetcaPkiLib.Utilities
'Private mobjSign As NetcaPkiLib.SignedData
'本接口不支持签章图片
Private mobjUtil As Object          '字节数组类型转换成可变类型传入，否则动态创建NetcaPkiLib库对象调用方法抛错
Private mobjSign As Object
Private mobjCert As Object
'签章对象
Private mobjPDFSign As Object
Private mobjPDFUtilTool As Object

Private mblnInit As Boolean

'定制3：默认的证书筛选条件，特定项目需定制 多CA支持时需定制，如 { "NETCA", "GDCA", "SZCA","BJCA" }
Private mstrKeyType As String
Private Const NETCAPKI_CERTFROM As String = "Device"
Private Const CATITLE   As Integer = 21
Private Const NETCAPKI_ALGORITHM_RSA As Integer = 1
Private Const NETCAPKI_CMS_ENCODE_BASE64 As Integer = 2
Private Const NETCAPKI_ALGORITHM_RSASIGN As Integer = 4        '定制6：RSA签名算法，一般无需定制 2017-3-7定制由SHA1改为SHA256
Private Const NETCAPKI_ALGORITHM_SHA1WITHRSA As Integer = 2
Private Const NETCAPKI_ALGORITHM_SM2SIGN As Integer = 25    ''定制7：SM2签名算法，一般无需定制
Private Const NETCAPKI_ALGORITHM_SM3WITHSM2 As Integer = 25
Private Const NETCAPKI_CERT_PURPOSE_SIGN As Integer = 2
Private Const NETCAPKI_CERT_PURPOSE_ENCRYPT As Integer = 1
Private Const NETCAPKI_CP_UTF8 = 65001  '(&HFDE9)
Private Const NETCAPKI_SIGNEDDATA_INCLUDE_CERT_OPTION As Integer = 2
Private Const NETCAPKI_ALGORITHM_HASH As Integer = 8192
Private Const NETCAPKI_BASE64_ENCODE_NO_NL  As Integer = 1
'定制4：NETCA证书实体唯一标识，特定项目需定制
'NETCA证书唯一实体标识OID：1.3.6.1.4.1.18760.1.12.11 NETCA证书绑定值OID：1.3.6.1.4.1.18760.1.12.14；
Private Const NETCAPKI_UUID As String = "1.3.6.1.4.1.18760.1.12.11"
Private mlngSeq As Long
'测试网关地址
Private Const M_STR_WGCert As String = "MIIFKzCCBBOgAwIBAgILEKwZ+Y6IvmAddaMwDQYJKoZIhvcNAQELBQAwUjELMAkGA1UEBhMCQ04xJDAiBgNVBAoMG05FVENBIENlcnRpZmljYXRlIEF1dGhvcml0eTEdMBsGA1UEAwwUQ0NTIE5FVENBIFQyIFN1YjEgQ0EwHhcNMTcwODE1MDcyMTE5WhcN" & vbNewLine & _
            "MjIwODE1MDcyMTE5WjBXMQswCQYDVQQGEwJDTjEwMC4GA1UECgwn5bm/5Lic55yB55S15a2Q5ZWG5Yqh6K6k6K+B5pyJ6ZmQ5YWs5Y+4MRYwFAYDVQQDDA0xNC4xOC4xNTguMTQ3MIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEAuz5rn+GXUqELj1ZlmGUiu39JEy5aW3BDIvnrmEO1y38ERhsH8iBK+nqWkcQpf6fxvBpdPHJGtWk9" & vbNewLine & _
            "gZSTJcVaMq0bfkr3R3u4Xvr+lUSWjt3z1Y3m+wcRP/X8FQ0GMdrhLV31rKBk00XdlHI80ZN/rAbQ+1H5yGqmxOZeIWY2cRdqapc9w+g+VjnZ2XtrBnpxIBFysSXZSg6v" & vbNewLine & _
            "RwT6i5uGdmJFH8/OccmFQHYpG3knJGRjfXbo2Y36qEtVhyyzq5mcHJe6Jcktj6k7iTltyydshEs79Dn3+IW6ZS7XWDloiqppkFrXlCXfViSnrdtxn2p/4qHVXfx5u+bfWRwAXe/+KwIDAQABo4IB+zCCAfcwHwYDVR0jBBgwFoAUBrkf1Z2yDPHEhHRCwjAl" & vbNewLine & _
            "4gnaVtgwHQYDVR0OBBYEFETmsNMDHg342ZJs/vFBkx5gQWQIMIGGBggrBgEFBQcBAQR6MHgwQAYIKwYBBQUHMAKGNGh0dHA6Ly8xNC4xNTIuMTIwLjE3MC90ZXN0Y2FjZXJ0cy9ORVRDQVQyU3ViMUNBLmNydCAwNAYIKwYBBQUHMAGGKGh0dHA6Ly8xOTIuMTY4LjAuNjEvb2NzcGNvbnNvbGUvY2hlY2suZG8wOwYDVR0fBDQwMjAwoC6gLIYqaHR0cDovL3Rlc3QuY25jYS5uZXQvY3JsL05FVENBVDJTdWIxQ0EuY3JsMGsGA1UdIARkMGIwYAYKKwYBBAGBkkgOCjBSMFAGCCsGAQUFBwIBFkRodHRwOi8vd3d3LmNuY2EubmV0L2NzL2tub3dsZWRnZS93aGl0ZXBhcGVyL2Nwcy9uZXRDQXRlc3RjZXJ0" & vbNewLine & _
            "Y3BzLnBkZjAMBgNVHRMBAf8EAjAAMDQGCisGAQQBgZJIAQ4EJgwkNDkwOTFlZTEzMWFhNTRmOGY5YWI3NmZmZWM2MWQzMmNATDIxMA4GA1UdDwEB/wQEAwIEsDAdBgNVHSUEFjAUBggrBgEFBQcDAgYIKwYBBQUHAwEwDwYDVR0RBAgwBocEDhKekzANBgkq" & vbNewLine & _
            "hkiG9w0BAQsFAAOCAQEAxZG7MDyTufKuo9VImkyl7Zxq2JnzvqBC5CBVJjGkJE+DuEvhOKz80isBPOXA4Gbjco0pHdIhBg8uBkyQPNbQwlMB2h2Kxi8+dCt9aGvZ7QU04vHuXIjrMZ0utZJJbiXn0EojaDyrDDiGxtfyv5Cftqrn1jhOKPKYKel2buL7U5lO" & vbNewLine & _
            "fAA1TRdJP5CWwqQf7N8+MfFCmLBugFGYTiQ9LXDOwFK4sTCw2EJMLs8MaioObd+ETkjSkx/39X158kCoW2Ey+XTWdZx1jl8sZ7UEUZRHdfR/oNuTptyWcV8YdGGhg+YA" & vbNewLine & _
            "3dQNRO0LA8MoxHFmXAzqwFamq2wDdtUnGbdcIPI+Gg=="

Public Function SignedDataWithTSA(ByVal strSource As String, ByVal strTsaUrl As String, Optional ByVal blnNoHasSource As Boolean, _
        Optional ByRef strErr As String) As String
'参数:blnNoHasSource=false 带原文签名(缺省) ;Ture=不带原文签名
    Dim arrByte() As Byte
    Dim varRet As Variant
    
        On Error GoTo errH
    
100        If Trim(strSource) = "" Then
102             strErr = "原文内容为空": Exit Function
            End If
104         If Trim(strTsaUrl) = "" Then
106             strErr = "时间戳URL为空": Exit Function
            End If
            '字符串转字节数组
            If (mobjSign.SetSignCertificate(mobjCert, "", False) = False) Then
                strErr = "设置签名证书失败"
                Exit Function
            End If
118         arrByte = ConvertByte(strSource)
126         Call mobjSign.SetSignAlgorithm(-1, IIf(GetX509CertificateInfo(mobjCert, 8) = NETCAPKI_ALGORITHM_RSA & "", NETCAPKI_ALGORITHM_RSASIGN, NETCAPKI_ALGORITHM_SM3WITHSM2))
128         mobjSign.Detached = blnNoHasSource
            varRet = arrByte
            On Error Resume Next
130         varRet = mobjSign.SignWithTSATimeStamp(strTsaUrl, varRet, NETCAPKI_CMS_ENCODE_BASE64)
            If Err.Number <> 0 Then
                strErr = Err.Description: Exit Function
            End If
            Err.Clear: On Error GoTo errH
            SignedDataWithTSA = varRet
    Exit Function
errH:
    MsgBoxEx Err.Description & vbCrLf & "signedDataWithTSA" & "行 " & Erl, vbExclamation + vbOKOnly, gstrSysName
End Function

Private Function VerifySignedData(ByVal strSource As String, ByVal strSignature As String, Optional ByRef objSign As Object, Optional ByRef objCert As Object) As Boolean
'功能:PKCS7签名验证并获取签名证书
    Dim bytSrc() As Byte
    Dim bytSignature() As Byte
    Dim bytRet() As Byte
    Dim blnSignFormat As Boolean
    Dim blnDetached As Boolean
    Dim blnRet As Boolean
    Dim varTemp As Variant
    Dim varRet As Variant
    Dim varSrc As Variant
    On Error GoTo errH

    bytSrc = ConvertByte(strSource)
    bytSignature = Base64Decode(strSignature)
    varTemp = bytSignature
    Set objSign = CreateObject("NetcaPki.SignedData.1")
    If objSign Is Nothing Then
        MsgBoxEx "签名对象创建失败！", vbInformation + vbOKOnly, gstrSysName: Exit Function
    End If
    blnSignFormat = objSign.IsSign(varTemp)
    If Not blnSignFormat Then
        MsgBoxEx "签名信息验签未通过:签名数据格式不正确!", vbInformation + vbOKOnly, gstrSysName: Exit Function
    End If
    
    blnDetached = objSign.IsDetachedSign(varTemp)
    If blnDetached Then
    '不带原文 mobjSign.Detached = true
        varTemp = bytSrc
        blnRet = objSign.DetachedVerify(varTemp, strSignature, False)
        If Not blnRet Then
             MsgBoxEx "签名信息验签未通过!", vbInformation + vbOKOnly, gstrSysName: Exit Function
        End If
    
    Else '带原文
        'mobjSign.Detached = False
        varTemp = strSignature
        bytRet = objSign.Verify(varTemp, True)
        varRet = bytRet: varSrc = bytSrc
        blnRet = mobjUtil.ByteArrayCompare(varRet, varSrc)
        If Not blnRet Then
            MsgBoxEx "签名信息验证未通过:原文与签名信息不一致!", vbInformation + vbOKOnly, gstrSysName: Exit Function
        End If
    End If

    Set objCert = objSign.GetSignCertificate(-1)
    VerifySignedData = True
    Exit Function
errH:
    MsgBoxEx Err.Description & vbCrLf & "VerifySignedData" & "行 " & Erl, vbExclamation + vbOKOnly, gstrSysName
End Function

Private Function VerifySignedDataWithTSA(ByVal strSource As String, ByVal strSignature As String, ByRef strSignTime As String) As Boolean
'功能: 5.4.5 PKCS7时间戳签名验证并获取证书
    Dim blnRet As Boolean
    Dim varTemp As Variant
    Dim objSign As Object
    Dim i As Integer
    
    On Error GoTo errH
          
    If Not VerifySignedData(strSource, strSignature, objSign) Then Exit Function
    i = objSign.GetSignerCount()
    '获取签名时间
    If i >= 1 Then
        blnRet = objSign.HasTSATimeStamp(i)
        If blnRet Then
            strSignTime = objSign.GetTSATimeStamp(i)
            strSignTime = Format(strSignTime, "yyyy-MM-dd HH:mm:ss")
        End If
    End If
    VerifySignedDataWithTSA = True
    Exit Function
errH:
    MsgBoxEx Err.Description & vbCrLf & "VerifySignedDataWithTSA" & "行 " & Erl, vbExclamation + vbOKOnly, gstrSysName
End Function

Private Function VerifyCert(objCert As Object, ByVal strUrl As String, ByVal datVerifytime As Date, _
        ByVal strku As String, ByVal strGWSeverCert As String) As Collection
'功能:  网关数字证书验证函数 网关验证证书方式一：服务接口【推荐使用】
    Dim lngSeq As Long
    Dim strHexDigest As String
    Dim strReqParam As String
    Dim bytReqParam() As Byte
    Dim varData As Variant, varSign As Variant
    Dim strRet As String
    Dim colList As Collection
    Dim strSignatureb64 As String, strDigest As String
    Dim objWGCert As Object
    Dim blnRet As Boolean
    Dim varTemp As Variant
    Dim strTemp As String
    Dim lngRet As Long
    
    On Error GoTo errH
    If strGWSeverCert = "" Then
        MsgBoxEx "公网网关证书内容为空!", vbInformation + vbOKOnly, gstrSysName: Exit Function
    End If

    '封装请求码,http发送到电子认证网关
    strHexDigest = UCase(ConvertHex(objCert.ThumbPrint(NETCAPKI_ALGORITHM_HASH)))
    varTemp = objCert.Encode(1)
    strTemp = Base64Encode(varTemp)
    
    strReqParam = "verifytime=" & datVerifytime & "&b64cert=" & URLEncode(strTemp) & "&ku=" & strku
    bytReqParam = StrConv(strReqParam, vbFromUnicode)
    strRet = HttpPost(strUrl, strReqParam, responseText, "application/x-www-form-urlencoded; charset=utf-8")
    Set colList = ParseXML(strRet)

    '验证服务器签名
    If colList Is Nothing Then
        MsgBoxEx "服务端返回数据包有误或服务端返回数据为空！", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If

'    varData = ConvertByte(colList("data"))
'    strSignatureb64 = Base64Encode(mobjUtil.HexToBinary(colList("signature")))
'    Set objWGCert = GetX509CertificateByBase(strGWSeverCert)
'
'    varSign = Base64Decode(strSignatureb64)
'    lngRet = IIf(GetX509CertificateInfo(objWGCert, 8) = NETCAPKI_ALGORITHM_RSA & "", NETCAPKI_ALGORITHM_SHA1WITHRSA, NETCAPKI_ALGORITHM_SM2SIGN)
'    blnRet = objWGCert.Verify(lngRet, varData, varSign)
'    If Not blnRet Then
'        MsgboxEx "服务端签名无效！", vbInformation + vbOKOnly, gstrSysName
'        Exit Function
'    End If

    '验证证书摘要是否匹配
    strDigest = UCase(colList("certsha1hex"))
    If strDigest <> strHexDigest Then
        MsgBoxEx "被验证的证书摘要不匹配,可能遭到恶意攻击！", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    Set VerifyCert = colList
    Exit Function
errH:
    MsgBoxEx Err.Description & vbCrLf & "VerifyCert" & "行 " & Erl, vbExclamation + vbOKOnly, gstrSysName
End Function
        
Public Function GetX509Certificate(ByVal intPURPOSE As Integer) As Object
'<summary>5.3.2 [常用]获取证书对象
'使用频率：较常用；
'使用场景：
'选择证书通常采用此函数。1）证书绑定时，2）证书登录时；
'根据全局变量定制项2、3，可通过该函数支持多CA支持；
'2016-07-22 luhanmin 修订入参
'</summary>
'<param name="NETCAPKI_CERT_PURPOSE">证书用途,参见Constants.NETCAPKI_CERT_PURPOSE定义；0：所有证书;NETCAPKI_CERT_PURPOSE_SIGN=2;NETCAPKI_CERT_PURPOSE_ENCRYPT= 1;</param>
'<returns></returns>
    Dim objCert  As Object
    Dim strType As String
    Dim strFilter As String
    
     On Error GoTo errH
            
100     Set objCert = CreateObject("NetcaPki.Certificate.1")
        
        strType = "{"
102     strType = strType & """UIFlag"":""default"", ""InValidity"":true,"
104     If (intPURPOSE = NETCAPKI_CERT_PURPOSE_SIGN) Then
106         strType = strType & """Type"":""signature"" , "
108     ElseIf (intPURPOSE = NETCAPKI_CERT_PURPOSE_ENCRYPT) Then
110         strType = strType & """Type"":""encrypt"" , "
        End If
112     If (NETCAPKI_CERTFROM = "Device") Then
114         strType = strType & """Method"":""device"", ""Value"":""any"""
        Else
116         strType = strType & """Method"":""store"", ""Value"":""Type"":""current user"",""Value"":""my"""
        End If
        strType = strType & "}"
118     strFilter = "InValidity='True'"
120     If GetCAFilter() <> "" Then
122        strFilter = strFilter & "&&" & GetCAFilter()
        End If
124     If (intPURPOSE = NETCAPKI_CERT_PURPOSE_SIGN) Then
126         strFilter = strFilter & "&&CertType='Signature'&&CheckPrivKey='True'"
128     ElseIf (intPURPOSE = NETCAPKI_CERT_PURPOSE_ENCRYPT) Then
130         strFilter = strFilter & "&&CertType='Encrypt'&&CheckPrivKey='True'"
        End If
        On Error Resume Next
132     Call objCert.Select(strType, strFilter)
        If Err.Number <> 0 Then Set objCert = Nothing
        Err.Clear: On Error GoTo errH
134     Set GetX509Certificate = objCert
            
      Exit Function
errH:
    MsgBoxEx Err.Description & vbCrLf & "GetX509Certificate" & "行 " & Erl, vbExclamation + vbOKOnly, gstrSysName
End Function

Public Function GetX509CertificateByBase(strBase64 As String) As Object
'功能:[常用]获取证书对象（从证书BASE64编码信息中）
    Dim objCert As Object
    Dim varCert As Variant
    
    On Error GoTo errH
    
    varCert = ClearCertHeader(strBase64)
    Set objCert = CreateObject("NetcaPki.Certificate")
    Call objCert.Decode(varCert)
    Set GetX509CertificateByBase = objCert
    Exit Function
errH:
    MsgBoxEx Err.Description & vbCrLf & "GetX509CertificateByBase" & "行 " & Erl, vbExclamation + vbOKOnly, gstrSysName
End Function

Public Function GetX509CertificateInfo(objCert As Object, ByVal intInfoType As Integer) As String
'功能:[常用]获取证书属性信息
    Dim strRet As String
    Dim strTmp As String
    Dim strCN As String, strO As String
    Dim i As Integer
    Dim arrTmp As Variant
    Dim strCaType As String
    Dim bytData() As Byte
    
    On Error GoTo errH
    
    If objCert Is Nothing Then
        MsgBoxEx "证书信息为空!", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    Select Case intInfoType
    
        Case 0  '获取证书BASE64格式编码字符串 2012-12-03
            strRet = objCert.Encode(2)
            strRet = ClearCertHeader(strRet)
        Case 1 '证书姆印
            strRet = UCase(ConvertHex(objCert.ThumbPrint(NETCAPKI_ALGORITHM_HASH)))
        Case 2 '证书序列号
            strRet = objCert.SerialNumber
        Case 3  '证书Subject
            strRet = objCert.Subject
        Case 4  '证书颁发者Subject
            strRet = objCert.Issuer
        Case 5  '证书有效期起
            strRet = objCert.ValidFromDate
        Case 6  '证书有效期止
            strRet = objCert.ValidToDate
        Case 7  'KeyUsage 密钥用法
            strRet = objCert.KeyUsage
        Case 8  '证书的公钥的算法
            strRet = objCert.PublicKeyAlgorithm
        Case 9
            '获取顺序：1）证书唯一标识2）证书客户服务号，3）证书姆印
            ' UsrCertNO：证书绑定值(系统改造时,建议采用该值)
            '1) 取证书证件号码扩展域信息
            On Error Resume Next
            strRet = GetX509CertificateInfo(objCert, 36)
            If strRet <> "" Then GoTo ReturnLine
            '2) 获取证书客户服务号
            strRet = GetX509CertificateInfo(objCert, 23)
            If strRet <> "" Then GoTo ReturnLine
            '3）证书姆印
            strRet = GetX509CertificateInfo(objCert, 1)
            If strRet <> "" Then GoTo ReturnLine
            strRet = ""
            Err.Clear: On Error GoTo 0
        Case 10  ''OldUsrCertNo：旧的用户证书绑定值(证书更新后的原有9的取值)
            If GetX509CertificateInfo(objCert, CATITLE) = "NETCA" Then
                strRet = GetX509CertificateInfo(objCert, 31) '取证书旧姆印
            End If
        Case 11 '证书主题名称有CN项取CN项值无CN项，取O的值
            If GetX509CertificateInfo(objCert, 12) = "" Then
                strRet = GetX509CertificateInfo(objCert, 13)
            Else
                strRet = GetX509CertificateInfo(objCert, 12)
            End If
        Case 12 'Subject中的CN项（人名）
            On Error Resume Next
            strRet = objCert.GetStringInfo(20)
            Err.Clear: On Error GoTo 0
         Case 13    'Subject中的O项（人名）
            On Error Resume Next
            strRet = objCert.GetStringInfo(18)
            Err.Clear: On Error GoTo 0
         Case 14    'Subject中的地址（L项）
            On Error Resume Next
            strRet = objCert.GetStringInfo(40)
            Err.Clear: On Error GoTo 0
         Case 15    '证书颁发者的Email
            On Error Resume Next
            strRet = objCert.GetStringInfo(21)
            Err.Clear: On Error GoTo 0
         Case 16    'Subject中的部门名（OU项）
            On Error Resume Next
            strRet = objCert.GetStringInfo(19)
            Err.Clear: On Error GoTo 0
         Case 17    '用户国家名（C项）
            On Error Resume Next
            strRet = objCert.GetStringInfo(17)
            Err.Clear: On Error GoTo 0
         Case 18    '用户省州名（S项）
            On Error Resume Next
            strRet = objCert.GetStringInfo(39)
            Err.Clear: On Error GoTo 0
         Case 21  'CATITLE
            strTmp = GetX509CertificateInfo(objCert, 4)
            arrTmp = Split(mstrKeyType, ",")
            For i = LBound(arrTmp) To UBound(arrTmp)
                If InStr(strTmp, arrTmp(i)) > 0 Then
                    strRet = arrTmp(i): GoTo ReturnLine
                End If
            Next
            strRet = ""
         Case 22
             If GetX509CertificateInfo(objCert, CATITLE) = "NETCA" Then
                On Error GoTo ErrHandle
                'netca证书类型扩展OID:NETCA OID(1.3.6.1.4.1.18760.1.12.12.2)
                '1：服务器证书2：个人证书3: 机构证书4：机构员工证书5：机构业务证书(注：该类型国密标准待定)0：其他证书
                 strCaType = mobjUtil.DecodeASN1String(1, objCert.GetExtension("1.3.6.1.4.1.18760.1.12.12.2"))
    
                If strCaType = "001" Then
                    strRet = "3"
                ElseIf strCaType = "002" Then
                    strRet = "5"
                ElseIf strCaType = "003" Then
                    strRet = "4"
                ElseIf strCaType = "004" Then
                    strRet = "2"
                End If
ErrHandle:
                'region 根据CN项和O项判断
                 strCN = GetX509CertificateInfo(objCert, 12)
                 strO = GetX509CertificateInfo(objCert, 13)
                 If strCN <> "" And strO = "" Then
                     strRet = "2"
                 ElseIf ((strO <> "" And strCN = "") Or (strO <> "" And strCN <> "" And strO = strCN)) Then
                     strRet = "3"
                 ElseIf (strO <> "" And strCN <> "" And strO <> strCN) Then
                     strRet = "4"
                 End If
                 'endregion
                strRet = "0"
            Else
                strRet = "0"
            End If
         Case 23    '用户证书客服号
            strTmp = GetX509CertificateInfo(objCert, CATITLE)
            If strTmp = "NETCA" Then
                On Error Resume Next
                'netca 用户证书服务号OID
                strRet = mobjUtil.DecodeASN1String(1, objCert.GetExtension("1.3.6.1.4.1.18760.1.14"))
                Err.Clear: On Error GoTo 0
            ElseIf strTmp = "GDCA" Then
                strRet = GetX509CertificateInfo(objCert, 51)
            Else
                strRet = ""
            End If
         Case 24    '深圳地标
            On Error Resume Next
            '深圳地标
            strRet = mobjUtil.DecodeASN1String(1, objCert.GetExtension("2.16.156.112548"))
            Err.Clear: On Error GoTo 0
        Case 31 '证书旧姆印
            On Error Resume Next
            strTmp = objCert.GetStringInfo(29)
            If strTmp <> "" Then strRet = ConvertHex(strTmp)
            Err.Clear: On Error GoTo 0
        Case 36 'netca证书类型扩展OID
            On Error Resume Next
            strRet = mobjUtil.DecodeASN1String(1, objCert.GetExtension(NETCAPKI_UUID))
            Err.Clear: On Error GoTo 0
        Case 37     '证件号码解码
            strRet = GetX509CertificateInfo(objCert, 36)
            If Len(strRet) > 13 Then
             '00011@0006PO1MTIzNDU2Nzg5MDEyMzQ1Njc4
                i = InStr(strRet, "@")
                
                If i = 0 Then
                    strRet = "": GoTo ReturnLine
                End If
                strTmp = Mid(strRet, i + 7, 1) '获取编码标志位
                If strTmp = "1" Then
                    strTmp = Mid(strRet, i + 8)
                    strRet = mobjUtil.Decode(Base64Decode(strTmp), 65001)
                    GoTo ReturnLine
                ElseIf strTmp = "0" Then
                    strRet = Mid(strRet, i + 8)
                    GoTo ReturnLine
                End If
                strRet = ""
            Else
                strRet = ""
            End If
        Case 51 'GDCA的特定扩展域 51
            strTmp = "1.2.86.21.1.3"
            If GetX509CertificateInfo(objCert, CATITLE) = "GDCA" Then
                On Error Resume Next
                strRet = mobjUtil.Decode(objCert.GetExtension(strTmp), 65001)
                Err.Clear: On Error GoTo 0
            End If
        Case Else
            strRet = ""
    End Select
    
ReturnLine:
    GetX509CertificateInfo = strRet
            
    Exit Function
errH:
    MsgBoxEx Err.Description & vbCrLf & "GetX509CertificateInfo" & "行 " & Erl, vbExclamation + vbOKOnly, gstrSysName
End Function

Private Function ClearCertHeader(ByVal strCertBase As String) As String
'功能: 去除证书Base64编码的头尾部分
    Dim strcertHeader As String
    Dim strcertEnd As String
    
    On Error GoTo errH
    
    strcertHeader = "-----BEGIN CERTIFICATE-----"
    strcertEnd = "-----END CERTIFICATE-----"
    If InStr(strCertBase, strcertHeader) > 0 Then
        strCertBase = Mid(strCertBase, Len(strcertHeader) + 1)
        strCertBase = Mid(strCertBase, 1, Len(strCertBase) - Len(strcertEnd))
    End If
    ClearCertHeader = strCertBase
    Exit Function
errH:
    MsgBoxEx Err.Description & vbCrLf & "signedDataWithTSA" & "行 " & Erl, vbExclamation + vbOKOnly, gstrSysName
End Function

 Private Function GetCAFilter() As String
    Dim strFilter As String
    Dim arrTmp As Variant
    Dim i As Integer
    
    arrTmp = Split(mstrKeyType, ",")
    If (UBound(arrTmp) >= 0) Then
        strFilter = strFilter & "("
        For i = LBound(arrTmp) To UBound(arrTmp)
            If i = 0 Then
                strFilter = strFilter & "IssuerCN~'" & arrTmp(i) & "'"
            Else
                strFilter = strFilter & "||IssuerCN~'" & arrTmp(i) & "'"
            End If
        Next
        strFilter = strFilter & ")"
    End If
    GetCAFilter = strFilter
 End Function

Private Function ConvertHex(objData As Variant) As String
'功能:字节数组转Hex编码字符串
    On Error GoTo errH
    ConvertHex = mobjUtil.BinaryToHex(objData, True)
    Exit Function
errH:
    MsgBoxEx Err.Description & vbCrLf & "ConvertHex" & "行 " & Erl, vbExclamation + vbOKOnly, gstrSysName
 End Function
        
Private Function Base64Decode(ByVal strData As String) As Byte()
'功能:字符串Base64解码为字节数组
    On Error GoTo errH
    Base64Decode = mobjUtil.Base64Decode(strData, NETCAPKI_BASE64_ENCODE_NO_NL)
    Exit Function
errH:
    MsgBoxEx Err.Description & vbCrLf & "Base64Decode" & "行 " & Erl, vbExclamation + vbOKOnly, gstrSysName
 End Function
 
 Private Function Base64Encode(bytData As Variant) As String
'功能:字节数组Base64编码为字符串
    
    On Error GoTo errH
    Base64Encode = mobjUtil.Base64Encode(bytData, NETCAPKI_BASE64_ENCODE_NO_NL)
    Exit Function
errH:
    MsgBoxEx Err.Description & vbCrLf & "Base64Encode" & "行 " & Erl, vbExclamation + vbOKOnly, gstrSysName
 End Function
        
Private Function ConvertByte(ByVal strData As String) As Byte()
'功能:字符串转字节数组
    
    On Error GoTo errH

    ConvertByte = mobjUtil.Encode(strData, NETCAPKI_CP_UTF8)
                
    Exit Function
errH:
    MsgBoxEx Err.Description & vbCrLf & "ConvertByte" & "行 " & Erl, vbExclamation + vbOKOnly, gstrSysName
End Function

Private Function GetRandom(ByVal intLength As Integer) As String
'功能:获取随机数
    Dim objDevice As Object
    Dim bytRandom() As Byte
    Dim varTemp As Variant
    
    On Error GoTo errH
    Set objDevice = CreateObject("NetcaPki.Device.1")
    If objDevice Is Nothing Then
        MsgBoxEx "创建网证通对象【Device】失败！", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    bytRandom = objDevice.GenerateRandom(intLength)
    varTemp = bytRandom
    GetRandom = ConvertHex(varTemp)
    Exit Function
errH:
    MsgBoxEx Err.Description & vbCrLf & "GetRandom" & "行 " & Erl, vbExclamation + vbOKOnly, gstrSysName
End Function

Private Function SignedDataByPwd(ByVal strSource As String, ByVal blnNoHasSource As Boolean, ByRef strErr As String) As String
'功能: 带PIN码PKCS7签名
    Dim bytArr() As Byte
    
    On Error GoTo errH

    bytArr = ConvertByte(strSource)
    
    SignedDataByPwd = SignedDataByCertificate(mobjCert, bytArr, blnNoHasSource, strErr)

    Exit Function
errH:
    MsgBoxEx Err.Description & vbCrLf & "SignedDataByPwd" & "行 " & Erl, vbExclamation + vbOKOnly, gstrSysName
End Function

Private Function SignedDataByCertificate(ByVal objCert As Object, ByRef bytSource() As Byte, ByVal blnNoHasSource As Boolean, ByRef strErr As String) As String
'功能:使用证书进行PKCS7签名[具体代码实现]
'参数:blnNoHasSource=false 带原文签名(缺省) ;Ture=不带原文签名
    Dim varRet As Variant
    
    On Error GoTo errH
    If (mobjSign.SetSignCertificate(mobjCert, "", False) = False) Then
        strErr = "设置签名证书失败"
        Exit Function
    End If
    Call mobjSign.SetSignAlgorithm(-1, IIf(GetX509CertificateInfo(objCert, 8) = NETCAPKI_ALGORITHM_RSA & "", NETCAPKI_ALGORITHM_RSASIGN, NETCAPKI_ALGORITHM_SM2SIGN))
    Call mobjSign.SetIncludeCertificateOption(NETCAPKI_SIGNEDDATA_INCLUDE_CERT_OPTION)
    mobjSign.Detached = blnNoHasSource  'true：不带原文；false：带原文
    varRet = bytSource
    varRet = mobjSign.Sign(varRet, NETCAPKI_CMS_ENCODE_BASE64)

    SignedDataByCertificate = varRet
    Exit Function
errH:
    MsgBoxEx Err.Description & vbCrLf & "SignedDataByCertificate" & "行 " & Erl, vbExclamation + vbOKOnly, gstrSysName
End Function
        
Private Function ParseXML(ByVal strData As String) As Collection
'功能:解析XML字符串
' <summary>解析XML字符串
' <verifytime>20170824110405</verifytime>
'<certsha1hex>d52dbe11084d68f94fe57c13436883f128039f99</certsha1hex>
'<certname>测试员44&&网证通</certname>
'<status>3</status>
' </summary>
' <param name="sResponse"></param>
' <returns></returns>
    Dim colList As New Collection
    Dim xmlDoc As New DOMDocument
    Dim xmlList As IXMLDOMNodeList
    Dim strValue As String
    
    On Error GoTo errH

    '读取网关响应数据（XML格式）
    xmlDoc.loadXML (strData)
    Set xmlList = xmlDoc.selectNodes("VerifyCertResp")
    '证书验证时间
    strValue = xmlList(0).selectSingleNode(".//verifytime").Text
    colList.Add strValue, "verifytime"

    '证书摘要
    strValue = xmlList(0).selectSingleNode(".//certsha1hex").Text
    colList.Add strValue, "certsha1hex"

    '证书主题
    strValue = xmlList(0).selectSingleNode(".//certname").Text
    colList.Add strValue, "certname"

    '读取证书状态码
    strValue = xmlList(0).selectSingleNode(".//status").Text
    colList.Add strValue, "status"

    '读取服务器签名值 与签名数据
    strValue = xmlDoc.selectSingleNode(".//signature").Text
    colList.Add strValue, "signature"
    strValue = "<?xml version=""1.0"" encoding=""UTF-8""?>" & xmlDoc.selectSingleNode(".//data").xml
    colList.Add strValue, "data"
    
    If colList.Count > 0 Then
        Set ParseXML = colList
    End If
    Exit Function
errH:
    MsgBoxEx Err.Description & vbCrLf & "ParseXML" & "行 " & Erl, vbExclamation + vbOKOnly, gstrSysName
End Function
        
Private Function ParseCertStatus(ByVal lngCertCode As Long) As String
'功能:解析证书状态码
    Dim arrCertCode As Variant
    Dim strRet As String
    
    On Error GoTo errH

    arrCertCode = Array("证书有效", "验证处理失败", "证书格式有误", _
                        "证书不在有效期内", "证书不能用于电子签名", _
                        "证书名字不合", "证书策略不合", "证书扩展不合", _
                        "不受证书链信任", "证书被注销", "注销状态未知", _
                        "用户证书未授权", "用户状态被锁定")

    If lngCertCode >= 0 And lngCertCode < 13 Then
        strRet = arrCertCode(lngCertCode)
    Else
        strRet = "其他错误"
    End If
    ParseCertStatus = strRet
    Exit Function
errH:
    MsgBoxEx Err.Description & vbCrLf & "ParseCertStatus" & "行 " & Erl, vbExclamation + vbOKOnly, gstrSysName
End Function
        
Private Function EnCodeURL(ByVal strUrl As String) As String
'功能:将传人字符串按UTF编码方式转换成十六进制的转义序列
'说明：encodeURI-javaScript方法不会对 ASCII 字母和数字进行编码，也不会对这些 ASCII 标点符号进行编码： - _ . ! ~ * ' ( )
    Dim i As Long
    Dim strChar As String
    Dim intAsc As Integer
    Dim strRet As String
    Dim objMSScriptCtl As Object
    
    On Error GoTo errH
    Set objMSScriptCtl = CreateObject("MSScriptControl.ScriptControl.1")
    objMSScriptCtl.Language = "JavaScript"
            
    For i = 1 To Len(strUrl)
        strChar = Mid(strUrl, i, 1)
        intAsc = Asc(strChar)
        If intAsc >= 0 And intAsc <= 127 Then
           strChar = "%" & Hex(intAsc)
        Else
            strChar = objMSScriptCtl.Eval("encodeURI(""" & strChar & """)")
        End If
        strRet = strRet & strChar
    Next
    
    EnCodeURL = strRet
errH:
    MsgBoxEx Err.Description & vbCrLf & "EnCodeURL" & "行 " & Erl, vbExclamation + vbOKOnly, gstrSysName
End Function

Private Function URLEncode(ByVal strParameter As String) As String
    Dim strTemp As String
    Dim i As Integer
    Dim intValue As Integer

    Dim bytData() As Byte

    strTemp = ""
    bytData = StrConv(strParameter, vbFromUnicode)
    For i = 0 To UBound(bytData)
        intValue = bytData(i)
        If (intValue >= 48 And intValue <= 57) Or _
            (intValue >= 65 And intValue <= 90) Or _
            (intValue >= 97 And intValue <= 122) Then
            strTemp = strTemp & Chr(intValue)
        ElseIf intValue = 32 Then
            strTemp = strTemp & "+"
        Else
            strTemp = strTemp & "%" & LCase(Hex(intValue))
        End If
    Next
    URLEncode = strTemp
End Function


Public Function WZT_InitObj() As Boolean
     '证书部件初始化
    Dim lngRet As Long
    Dim strTSAIP As String
    Dim strPara As String
    Dim varTmp As Variant
    Dim i As Long
    

100 If glngSign > 1 Then WZT_InitObj = True: Exit Function

    On Error GoTo errH
102 Set mobjSign = CreateObject("NetcaPki.SignedData.1")
104 Set mobjUtil = CreateObject("NetcaPki.Utilities.1")
106 Set mobjPDFSign = CreateObject("Netca.PDFSign")
108 Set mobjPDFUtilTool = CreateObject("Netca.UtilTool")
    
110 If mobjSign Is Nothing Or mobjUtil Is Nothing Then GoTo errMsg
        
112 mstrKeyType = "NETCA"
        
114 gstrPara = gobjComLib.zlDatabase.GetPara(90000, glngSys, , "") '读取配置内容
116 LogWrite "WZT_InitObj", "CA参数:" & gstrPara
118 If gstrPara = "" Then
120     Err.Raise -1, , "当前系统【" & glngSys & "】没有配置电子签名参数,请到启用电子签名接口处设置。"
        Exit Function
    End If
122 Call WZT_GetPara
124 mblnInit = True
126 WZT_InitObj = True
   
    Exit Function
errMsg:
128 MsgBoxEx "创建网证通对象失败！请确认【网证通安全客户端】已安装且该部件【NetcaPkiCom.dll】已注册成功。" & vbCrLf & "错误行:" & Erl, vbExclamation, gstrSysName
    Exit Function
errH:
130  MsgBoxEx "创建接口部件失败！" & vbNewLine & "第" & CStr(Erl()) & "行 " & Err.Description, vbExclamation, gstrSysName
End Function

Public Function WZT_Sign(ByVal strSource As String, ByRef strSignData As String, ByRef strTimeStamp As String, ByRef strTimeStampCode As String) As Boolean
'功能:签名签名值包含签名值信息和时间戳信息;带原文签名;原文越大签名返回值越大
    Dim strUrl As String
    Dim blnCheck As Boolean
    Dim lngSignID As Long
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim strErr As String
    
    On Error GoTo errH
    blnCheck = WZT_CheckCert()
    If blnCheck Then                 '验证当前USB是否是签名用户的，并获取签名证书
        '证书ID进行签名
        If gudtPara.blnISTS Then
            strUrl = "http://" & gudtPara.strTSIP & IIf(gudtPara.strTSPort = "", "", ":" & gudtPara.strTSPort) & "/NETCATimeStampServer/TSAServer.jsp"
            LogWrite "WZT_Sign", "时间戳地址:" & strUrl
            strSignData = SignedDataWithTSA(strSource, strUrl, , strErr)  '带原文签名
            LogWrite "WZT_Sign", "签名返回值:" & strSignData
            If strSignData <> "" Then
                If Not VerifySignedDataWithTSA(strSource, strSignData, strTimeStamp) Then
                    Exit Function
                End If
            Else
                MsgBoxEx "签名失败!" & IIf(strErr <> "", vbCrLf & "原因:" & vbCrLf & strErr & "!", ""), vbInformation, gstrSysName
                Exit Function
            End If
        Else
            strSignData = SignedDataByPwd(strSource, False, strErr)
            LogWrite "WZT_Sign", "签名返回值:" & strSignData
            If strSignData = "" Then
                MsgBoxEx "签名失败!" & IIf(strErr <> "", vbCrLf & "原因:" & vbCrLf & strErr & "!", ""), vbInformation, gstrSysName
                Exit Function
            End If
            strTimeStamp = Format(gobjComLib.zlDatabase.Currentdate & "", "yyyy-MM-dd HH:mm:ss")
        End If
    Else
        MsgBoxEx "签名失败！", vbInformation, gstrSysName
        Exit Function
    End If

    strSQL = "Select 电子签名记录_ID.Nextval as ID From Dual"
    Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, "WZT_Sign")
    lngSignID = Val(rsTmp!id)
    
    If WZT_SaveLob(lngSignID, strSignData) Then
        strSignData = lngSignID
    Else
        strSignData = ""
        MsgBoxEx "签名失败！", vbInformation, gstrSysName
        Exit Function
    End If
    WZT_Sign = True
    Exit Function
errH:
114     MsgBoxEx "签名失败！" & vbNewLine & "第" & CStr(Erl()) & "行 " & Err.Description, vbExclamation, gstrSysName
End Function

Public Function WZT_VerifySign(ByVal strSource As String, ByVal strSignData As String, Optional ByVal strTimeStampCode As String) As Boolean
    '功能;验证签名
    '参数:strSignData -签名值
    Dim blnRet As Boolean
    
    On Error GoTo errH
    LogWrite "WZT_VerifySign", "验证签名原文:" & strSource & vbCrLf & "验证签名值:" & strSignData & vbCrLf & "签名时间戳信息:" & strTimeStampCode
    strSignData = WZT_ReadLob(CLng(strSignData))
    If gudtPara.blnISTS Then
        blnRet = VerifySignedDataWithTSA(strSource, strSignData, "")
    Else
        blnRet = VerifySignedData(strSource, strSignData)
    End If
    If blnRet Then
        MsgBoxEx "验证成功，该电子签名数据有效!", vbInformation, gstrSysName
    Else
         Exit Function
    End If
    WZT_VerifySign = True
    Exit Function
errH:
140     MsgBoxEx "验证签名失败！" & vbNewLine & "第" & CStr(Erl()) & "行 " & Err.Description, vbExclamation, gstrSysName

End Function

Public Function WZT_CheckCert() As Boolean
    '功能：读取USB进行设备初始化并登录
    Dim strKeySN As String, strUserName As String
    Dim colList As Collection
    Dim datBegin As Date
    Dim strUrl As String
    Dim lngRet As Long
    Dim strMsg As String
    
    On Error GoTo errH
    
    If Not GetCertList(strKeySN, strUserName, , "") Then Exit Function
    If mUserInfo.strCertSn <> strKeySN Then
        MsgBoxEx "绑定证书：" & _
           vbCrLf & vbTab & "【" & mUserInfo.strCertSn & "】" & vbCrLf & _
           "当前证书:" & vbCrLf & vbTab & "【" & strKeySN & "】" & vbCrLf & _
           "当前证书用户名:" & vbCrLf & vbTab & "【" & strUserName & "】" & vbCrLf & _
           "绑定证书与当前证书不是同一个证书,不能使用！！", vbInformation, gstrSysName
        Set mobjCert = Nothing
        Exit Function
    End If
    datBegin = gobjComLib.zlDatabase.Currentdate
    '绑定拉姆值，证书更换不影响拉姆值
    strUrl = "http://" & gudtPara.strSIGNIP & ":" & gudtPara.strSignPort & "/NetcaCertAA/appintf/verifyusercert"
    Set colList = VerifyCert(mobjCert, strUrl, Format(datBegin, "HH:mm:ss"), "c0", gudtPara.strOption)
    If colList Is Nothing Then
        MsgBoxEx "验证失败!", vbInformation + vbOKOnly, gstrSysName: Exit Function
        Exit Function
    End If
    lngRet = Val(colList("status"))
    If lngRet <> 0 Then
        strMsg = ParseCertStatus(lngRet)
        MsgBoxEx "网关验证失败!" & IIf(strMsg <> "", "原因:" & strMsg, ""), vbInformation + vbOKOnly, gstrSysName: Exit Function
        Exit Function
    End If
    WZT_CheckCert = True
    Exit Function
errH:
     MsgBoxEx "检查USBKEY失败！" & vbNewLine & "第" & CStr(Erl()) & "行 " & Err.Description, vbExclamation, gstrSysName
End Function


Public Function WZT_RegCert(arrCertInfo As Variant) As Boolean
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

        For i = LBound(arrCertInfo) To UBound(arrCertInfo)
             arrCertInfo(i) = ""
        Next

        If GetCertList(strKeyId, strCertUserName, strCertDN, , strPicPath) Then
            arrCertInfo(0) = strCertUserName
            arrCertInfo(1) = strCertDN
            arrCertInfo(2) = strKeyId
            arrCertInfo(5) = strPicPath
            WZT_RegCert = True
        End If

        Exit Function
errH:
     MsgBoxEx "获取证书信息失败！" & vbNewLine & "第" & CStr(Erl()) & "行 " & Err.Description, vbExclamation, gstrSysName

End Function

Private Function GetCertList(Optional ByRef strUniqueID As String = "-1", Optional ByRef strName As String = "-1", _
    Optional ByRef strCertDN As String = "-1", Optional ByRef strDate As String = "-1", Optional ByRef strPic As String = "-1") As Boolean
    '功能:获取网证通证书详情
        Dim varTemp As Variant
        Dim lngDay As Long
        Dim varRet  As Variant
        Dim strBase64 As String
        
        On Error GoTo errH
    '        If mobjCert Is Nothing Then
100     Set mobjCert = GetX509Certificate(NETCAPKI_CERT_PURPOSE_SIGN) '签名证书
    '        End If
        
102     If mobjCert Is Nothing Then
104         MsgBoxEx "没有符合的证书供选择,请检查是否插入KEY!", vbInformation + vbOKOnly, gstrSysName
            Exit Function
        End If
    '        varTemp = Array("证书PEM编码", "证书姆印", "证书序列号", "证书主题", "证书颁发者主题", _
    '                "证书有效期起", "证书有效期止", "密钥用法", "证书公钥算法", "用户证书绑定值", _
    '                "旧的用户证书绑定值", "证书主题名称", "Subject的人名(CN)", "Subject中的单位(O)", "Subject中的地址项(L)", _
    '                "Subject中的Email(E)", "Subject中的部门名(OU)", "国家名(C)", "省州名(S)", "", _
    '                "", "CA标志", "证书类型", "用户证书服务号", "深圳地标OID", _
    '                "", "", "", "", "", _
    '                "", "NETCA:旧姆印信息", "纳税人编码", "企业法人代码", "税务登记号", _
    '                "证书来源地", "证件号码信息(NETCA 证书唯一标识)", "证件号码明文", "", "", _
    '                "", "", "", "", "", _
    '                "", "", "", "", "", _
    '                "", "GDCA信任号TrustID")
            
106     If strUniqueID = "" Then strUniqueID = GetX509CertificateInfo(mobjCert, 1) '证书姆印
108     If strCertDN = "" Then strCertDN = GetX509CertificateInfo(mobjCert, 3) '证书主题DN
110     If strName = "" Then strName = GetX509CertificateInfo(mobjCert, 12) 'Subject的人名(CN)
112     If strDate = "" Then strDate = GetX509CertificateInfo(mobjCert, 6) '证书有效期止
  
114     If strPic = "" Then
            
116         varRet = mobjPDFSign.SelectCert("", 1)
118         varRet = mobjPDFUtilTool.GetImageFromDevicByCert(mobjPDFSign.SignCertBase64Encode)
120         strBase64 = Base64Encode(varRet)
122         If strBase64 <> "" Then
124             strPic = SaveBase64ToFile("gif", strUniqueID, strBase64)
            End If
        End If
126     If IsDate(strDate) Then
            '检查证书是否过期
128         lngDay = CheckValidaty(strDate)
130         If (lngDay <= 30 And lngDay > 0 And Not gblnShow) Then
132             MsgBoxEx "您的证书还有" & lngDay & "天过期", vbInformation, gstrSysName
134             gblnShow = True
136         ElseIf (lngDay <= 0) Then
138             MsgBoxEx "您的证书已过期 " & Abs(lngDay) & " 天"
                Exit Function
            End If
        End If
        
140     GetCertList = True
        Exit Function
errH:
142     MsgBoxEx "获取证书信息失败！" & vbNewLine & "第" & CStr(Erl()) & "行 " & Err.Description, vbExclamation, gstrSysName
End Function

Public Sub WZT_GetPara()
        Dim arrList As Variant

        On Error GoTo errH
100     gstrPara = gobjComLib.zlDatabase.GetPara(90000, glngSys, , "")
        If gstrPara = "" Then
            '签名服务器IP&&&签名服务器端口号&&&时间戳IP&&&时间戳端口号&&&网关证书&&&启用时间戳（0/1）
            gstrPara = "14.18.158.147" & G_STR_SPLIT & "10980" & _
                G_STR_SPLIT & "tsa.cnca.net" & G_STR_SPLIT & "" & G_STR_SPLIT & M_STR_WGCert & G_STR_SPLIT & "1"
        End If
        arrList = Split(gstrPara, G_STR_SPLIT)
        If UBound(arrList) = 5 Then
            gudtPara.strSIGNIP = arrList(0)
            gudtPara.strSignPort = arrList(1)
            gudtPara.strTSIP = arrList(2)
            gudtPara.strTSPort = arrList(3)
            gudtPara.strOption = arrList(4)
            gudtPara.blnISTS = Val(arrList(5)) = 1
        Else
            gudtPara.strSIGNIP = "14.18.158.147"
            gudtPara.strSignPort = "10980"
            gudtPara.strTSIP = "tsa.cnca.net"
            gudtPara.strTSPort = ""
            gudtPara.strOption = M_STR_WGCert
            gudtPara.blnISTS = False
        End If
        Exit Sub
errH:
146     MsgBoxEx "读取参数失败！" & vbNewLine & "第" & CStr(Erl()) & "行 " & Err.Description, vbInformation, gstrSysName
End Sub

Public Function WZT_SetParaStr() As String
    With gudtPara
        WZT_SetParaStr = IIf(Trim(.strSIGNIP) = "", "14.18.158.147", .strSIGNIP) & G_STR_SPLIT & IIf(Trim(.strSignPort) = "", "10980", .strSignPort) & _
            G_STR_SPLIT & IIf(Trim(.strTSIP) = "", "tsa.cnca.net", .strTSIP) & G_STR_SPLIT & IIf(Trim(.strTSPort) = "", "", .strTSPort) & _
            G_STR_SPLIT & IIf(Trim(.strOption) = "", M_STR_WGCert, .strOption) & G_STR_SPLIT & IIf(.blnISTS, "1", "0")

    End With
End Function

Public Sub WZT_UnLoadObj()
    Set mobjUtil = Nothing
    Set mobjSign = Nothing
End Sub

Public Function WZT_ReadLob(ByVal lngID As Long) As String
'功能：读取电子签章信息
    Dim rsLob As ADODB.Recordset
    Dim lngCount As Long
    Dim strText As String
    Dim strSQL As String
    Dim strFile As String
    
    Err = 0: On Error GoTo Errhand
    strSQL = "Select ZL_Readlob_电子签名记录([1],[2]) as 片段 From Dual"
    If strSQL = "" Then strFile = "": Exit Function
    lngCount = 0
    strFile = ""
    Do
        Set rsLob = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, "ZL_Readlob_电子签名记录", lngID, lngCount)
        If rsLob.EOF Then Exit Do
        If IsNull(rsLob.Fields(0).Value) Then Exit Do
        strText = rsLob.Fields(0).Value
        strFile = strFile & strText
        lngCount = lngCount + 1
    Loop
    
    WZT_ReadLob = strFile
    Exit Function
Errhand:
    Err.Clear
End Function

Public Function WZT_SaveLob(ByVal KeyWord As String, ByVal strFile As String) As Boolean
'功能：保存签章信息
'参数：
    Dim arrSQL() As String
    Dim i As Long
    
    If WZT_GetLobSql(KeyWord, strFile, arrSQL) Then
        Call gobjComLib.zlDatabase.ExecuteProcedureBeach(arrSQL, "WZT_SaveLob", False, False)
    Else
        WZT_SaveLob = False
    End If
    WZT_SaveLob = True
    Exit Function
Errhand:
    Err.Clear
    WZT_SaveLob = False
End Function

Public Function WZT_GetLobSql(ByVal lngKeyWord As Long, ByVal strFile As String, ByRef arySql() As String) As Boolean
'功能：产生保存指定的文件到指定表记录BLOB/CLOB字段的SQL语句
'参数：
'      KeyWord:确定数据记录的关键字，复合关键字以逗号分隔(仅5-电子病历格式为复合)
'      strFile: CLOB时,需要存储的文本内容
'      arySql():在该数据的基础上扩展增加保存的SQL语句
'返回：成功返回True，失败返回False
    Dim conChunkSize As Integer
    Dim lngCount As Long, lngBound As Long
    Dim aryChunk() As Byte, aryHex() As String, strText As String
    Dim strSQLRoot As String
    Dim strSubTxt As String
    
    On Error GoTo Errhand

    strSQLRoot = "ZL_AppendLob_电子签名记录(" & lngKeyWord
    If strSQLRoot = "" Then WZT_GetLobSql = False: Exit Function
    
    conChunkSize = 2000
    strText = strFile
    lngCount = 0
    Do
        strSubTxt = left(strText, conChunkSize)
        strText = Mid(strText, conChunkSize + 1)
        ReDim Preserve arySql(lngCount)
        arySql(lngCount) = strSQLRoot & ",'" & strSubTxt & "'," & IIf(lngCount = 0, 1, 0) & ")"
        lngCount = lngCount + 1
    Loop While Len(strText) > 0
    
    WZT_GetLobSql = True
    Exit Function
Errhand:
    Err.Clear
    WZT_GetLobSql = False
End Function


