Attribute VB_Name = "mdlGDWZT"
Option Explicit
'�˽ӿ�ǩ��ֵ���ȳ���4000���ַ�,��������:����ǩ����¼(ID,ǩ����Ϣ)  ID�ֶζ�Ӧ����ҵ����ǩ����Ϣ�С�����ZLHIS���ύ�Ų�
'Private mobjUtil As NetcaPkiLib.Utilities
'Private mobjSign As NetcaPkiLib.SignedData
'���ӿڲ�֧��ǩ��ͼƬ
Private mobjUtil As Object          '�ֽ���������ת���ɿɱ����ʹ��룬����̬����NetcaPkiLib�������÷����״�
Private mobjSign As Object
Private mobjCert As Object
'ǩ�¶���
Private mobjPDFSign As Object
Private mobjPDFUtilTool As Object

Private mblnInit As Boolean

'����3��Ĭ�ϵ�֤��ɸѡ�������ض���Ŀ�趨�� ��CA֧��ʱ�趨�ƣ��� { "NETCA", "GDCA", "SZCA","BJCA" }
Private mstrKeyType As String
Private Const NETCAPKI_CERTFROM As String = "Device"
Private Const CATITLE   As Integer = 21
Private Const NETCAPKI_ALGORITHM_RSA As Integer = 1
Private Const NETCAPKI_CMS_ENCODE_BASE64 As Integer = 2
Private Const NETCAPKI_ALGORITHM_RSASIGN As Integer = 4        '����6��RSAǩ���㷨��һ�����趨�� 2017-3-7������SHA1��ΪSHA256
Private Const NETCAPKI_ALGORITHM_SHA1WITHRSA As Integer = 2
Private Const NETCAPKI_ALGORITHM_SM2SIGN As Integer = 25    ''����7��SM2ǩ���㷨��һ�����趨��
Private Const NETCAPKI_ALGORITHM_SM3WITHSM2 As Integer = 25
Private Const NETCAPKI_CERT_PURPOSE_SIGN As Integer = 2
Private Const NETCAPKI_CERT_PURPOSE_ENCRYPT As Integer = 1
Private Const NETCAPKI_CP_UTF8 = 65001  '(&HFDE9)
Private Const NETCAPKI_SIGNEDDATA_INCLUDE_CERT_OPTION As Integer = 2
Private Const NETCAPKI_ALGORITHM_HASH As Integer = 8192
Private Const NETCAPKI_BASE64_ENCODE_NO_NL  As Integer = 1
'����4��NETCA֤��ʵ��Ψһ��ʶ���ض���Ŀ�趨��
'NETCA֤��Ψһʵ���ʶOID��1.3.6.1.4.1.18760.1.12.11 NETCA֤���ֵOID��1.3.6.1.4.1.18760.1.12.14��
Private Const NETCAPKI_UUID As String = "1.3.6.1.4.1.18760.1.12.11"
Private mlngSeq As Long
'�������ص�ַ
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
'����:blnNoHasSource=false ��ԭ��ǩ��(ȱʡ) ;Ture=����ԭ��ǩ��
    Dim arrByte() As Byte
    Dim varRet As Variant
    
        On Error GoTo errH
    
100        If Trim(strSource) = "" Then
102             strErr = "ԭ������Ϊ��": Exit Function
            End If
104         If Trim(strTsaUrl) = "" Then
106             strErr = "ʱ���URLΪ��": Exit Function
            End If
            '�ַ���ת�ֽ�����
            If (mobjSign.SetSignCertificate(mobjCert, "", False) = False) Then
                strErr = "����ǩ��֤��ʧ��"
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
    MsgBoxEx Err.Description & vbCrLf & "signedDataWithTSA" & "�� " & Erl, vbExclamation + vbOKOnly, gstrSysName
End Function

Private Function VerifySignedData(ByVal strSource As String, ByVal strSignature As String, Optional ByRef objSign As Object, Optional ByRef objCert As Object) As Boolean
'����:PKCS7ǩ����֤����ȡǩ��֤��
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
        MsgBoxEx "ǩ�����󴴽�ʧ�ܣ�", vbInformation + vbOKOnly, gstrSysName: Exit Function
    End If
    blnSignFormat = objSign.IsSign(varTemp)
    If Not blnSignFormat Then
        MsgBoxEx "ǩ����Ϣ��ǩδͨ��:ǩ�����ݸ�ʽ����ȷ!", vbInformation + vbOKOnly, gstrSysName: Exit Function
    End If
    
    blnDetached = objSign.IsDetachedSign(varTemp)
    If blnDetached Then
    '����ԭ�� mobjSign.Detached = true
        varTemp = bytSrc
        blnRet = objSign.DetachedVerify(varTemp, strSignature, False)
        If Not blnRet Then
             MsgBoxEx "ǩ����Ϣ��ǩδͨ��!", vbInformation + vbOKOnly, gstrSysName: Exit Function
        End If
    
    Else '��ԭ��
        'mobjSign.Detached = False
        varTemp = strSignature
        bytRet = objSign.Verify(varTemp, True)
        varRet = bytRet: varSrc = bytSrc
        blnRet = mobjUtil.ByteArrayCompare(varRet, varSrc)
        If Not blnRet Then
            MsgBoxEx "ǩ����Ϣ��֤δͨ��:ԭ����ǩ����Ϣ��һ��!", vbInformation + vbOKOnly, gstrSysName: Exit Function
        End If
    End If

    Set objCert = objSign.GetSignCertificate(-1)
    VerifySignedData = True
    Exit Function
errH:
    MsgBoxEx Err.Description & vbCrLf & "VerifySignedData" & "�� " & Erl, vbExclamation + vbOKOnly, gstrSysName
End Function

Private Function VerifySignedDataWithTSA(ByVal strSource As String, ByVal strSignature As String, ByRef strSignTime As String) As Boolean
'����: 5.4.5 PKCS7ʱ���ǩ����֤����ȡ֤��
    Dim blnRet As Boolean
    Dim varTemp As Variant
    Dim objSign As Object
    Dim i As Integer
    
    On Error GoTo errH
          
    If Not VerifySignedData(strSource, strSignature, objSign) Then Exit Function
    i = objSign.GetSignerCount()
    '��ȡǩ��ʱ��
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
    MsgBoxEx Err.Description & vbCrLf & "VerifySignedDataWithTSA" & "�� " & Erl, vbExclamation + vbOKOnly, gstrSysName
End Function

Private Function VerifyCert(objCert As Object, ByVal strUrl As String, ByVal datVerifytime As Date, _
        ByVal strku As String, ByVal strGWSeverCert As String) As Collection
'����:  ��������֤����֤���� ������֤֤�鷽ʽһ������ӿڡ��Ƽ�ʹ�á�
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
        MsgBoxEx "��������֤������Ϊ��!", vbInformation + vbOKOnly, gstrSysName: Exit Function
    End If

    '��װ������,http���͵�������֤����
    strHexDigest = UCase(ConvertHex(objCert.ThumbPrint(NETCAPKI_ALGORITHM_HASH)))
    varTemp = objCert.Encode(1)
    strTemp = Base64Encode(varTemp)
    
    strReqParam = "verifytime=" & datVerifytime & "&b64cert=" & URLEncode(strTemp) & "&ku=" & strku
    bytReqParam = StrConv(strReqParam, vbFromUnicode)
    strRet = HttpPost(strUrl, strReqParam, responseText, "application/x-www-form-urlencoded; charset=utf-8")
    Set colList = ParseXML(strRet)

    '��֤������ǩ��
    If colList Is Nothing Then
        MsgBoxEx "����˷������ݰ���������˷�������Ϊ�գ�", vbInformation + vbOKOnly, gstrSysName
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
'        MsgboxEx "�����ǩ����Ч��", vbInformation + vbOKOnly, gstrSysName
'        Exit Function
'    End If

    '��֤֤��ժҪ�Ƿ�ƥ��
    strDigest = UCase(colList("certsha1hex"))
    If strDigest <> strHexDigest Then
        MsgBoxEx "����֤��֤��ժҪ��ƥ��,�����⵽���⹥����", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    Set VerifyCert = colList
    Exit Function
errH:
    MsgBoxEx Err.Description & vbCrLf & "VerifyCert" & "�� " & Erl, vbExclamation + vbOKOnly, gstrSysName
End Function
        
Public Function GetX509Certificate(ByVal intPURPOSE As Integer) As Object
'<summary>5.3.2 [����]��ȡ֤�����
'ʹ��Ƶ�ʣ��ϳ��ã�
'ʹ�ó�����
'ѡ��֤��ͨ�����ô˺�����1��֤���ʱ��2��֤���¼ʱ��
'����ȫ�ֱ���������2��3����ͨ���ú���֧�ֶ�CA֧�֣�
'2016-07-22 luhanmin �޶����
'</summary>
'<param name="NETCAPKI_CERT_PURPOSE">֤����;,�μ�Constants.NETCAPKI_CERT_PURPOSE���壻0������֤��;NETCAPKI_CERT_PURPOSE_SIGN=2;NETCAPKI_CERT_PURPOSE_ENCRYPT= 1;</param>
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
    MsgBoxEx Err.Description & vbCrLf & "GetX509Certificate" & "�� " & Erl, vbExclamation + vbOKOnly, gstrSysName
End Function

Public Function GetX509CertificateByBase(strBase64 As String) As Object
'����:[����]��ȡ֤����󣨴�֤��BASE64������Ϣ�У�
    Dim objCert As Object
    Dim varCert As Variant
    
    On Error GoTo errH
    
    varCert = ClearCertHeader(strBase64)
    Set objCert = CreateObject("NetcaPki.Certificate")
    Call objCert.Decode(varCert)
    Set GetX509CertificateByBase = objCert
    Exit Function
errH:
    MsgBoxEx Err.Description & vbCrLf & "GetX509CertificateByBase" & "�� " & Erl, vbExclamation + vbOKOnly, gstrSysName
End Function

Public Function GetX509CertificateInfo(objCert As Object, ByVal intInfoType As Integer) As String
'����:[����]��ȡ֤��������Ϣ
    Dim strRet As String
    Dim strTmp As String
    Dim strCN As String, strO As String
    Dim i As Integer
    Dim arrTmp As Variant
    Dim strCaType As String
    Dim bytData() As Byte
    
    On Error GoTo errH
    
    If objCert Is Nothing Then
        MsgBoxEx "֤����ϢΪ��!", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    Select Case intInfoType
    
        Case 0  '��ȡ֤��BASE64��ʽ�����ַ��� 2012-12-03
            strRet = objCert.Encode(2)
            strRet = ClearCertHeader(strRet)
        Case 1 '֤��ķӡ
            strRet = UCase(ConvertHex(objCert.ThumbPrint(NETCAPKI_ALGORITHM_HASH)))
        Case 2 '֤�����к�
            strRet = objCert.SerialNumber
        Case 3  '֤��Subject
            strRet = objCert.Subject
        Case 4  '֤��䷢��Subject
            strRet = objCert.Issuer
        Case 5  '֤����Ч����
            strRet = objCert.ValidFromDate
        Case 6  '֤����Ч��ֹ
            strRet = objCert.ValidToDate
        Case 7  'KeyUsage ��Կ�÷�
            strRet = objCert.KeyUsage
        Case 8  '֤��Ĺ�Կ���㷨
            strRet = objCert.PublicKeyAlgorithm
        Case 9
            '��ȡ˳��1��֤��Ψһ��ʶ2��֤��ͻ�����ţ�3��֤��ķӡ
            ' UsrCertNO��֤���ֵ(ϵͳ����ʱ,������ø�ֵ)
            '1) ȡ֤��֤��������չ����Ϣ
            On Error Resume Next
            strRet = GetX509CertificateInfo(objCert, 36)
            If strRet <> "" Then GoTo ReturnLine
            '2) ��ȡ֤��ͻ������
            strRet = GetX509CertificateInfo(objCert, 23)
            If strRet <> "" Then GoTo ReturnLine
            '3��֤��ķӡ
            strRet = GetX509CertificateInfo(objCert, 1)
            If strRet <> "" Then GoTo ReturnLine
            strRet = ""
            Err.Clear: On Error GoTo 0
        Case 10  ''OldUsrCertNo���ɵ��û�֤���ֵ(֤����º��ԭ��9��ȡֵ)
            If GetX509CertificateInfo(objCert, CATITLE) = "NETCA" Then
                strRet = GetX509CertificateInfo(objCert, 31) 'ȡ֤���ķӡ
            End If
        Case 11 '֤������������CN��ȡCN��ֵ��CN�ȡO��ֵ
            If GetX509CertificateInfo(objCert, 12) = "" Then
                strRet = GetX509CertificateInfo(objCert, 13)
            Else
                strRet = GetX509CertificateInfo(objCert, 12)
            End If
        Case 12 'Subject�е�CN�������
            On Error Resume Next
            strRet = objCert.GetStringInfo(20)
            Err.Clear: On Error GoTo 0
         Case 13    'Subject�е�O�������
            On Error Resume Next
            strRet = objCert.GetStringInfo(18)
            Err.Clear: On Error GoTo 0
         Case 14    'Subject�еĵ�ַ��L�
            On Error Resume Next
            strRet = objCert.GetStringInfo(40)
            Err.Clear: On Error GoTo 0
         Case 15    '֤��䷢�ߵ�Email
            On Error Resume Next
            strRet = objCert.GetStringInfo(21)
            Err.Clear: On Error GoTo 0
         Case 16    'Subject�еĲ�������OU�
            On Error Resume Next
            strRet = objCert.GetStringInfo(19)
            Err.Clear: On Error GoTo 0
         Case 17    '�û���������C�
            On Error Resume Next
            strRet = objCert.GetStringInfo(17)
            Err.Clear: On Error GoTo 0
         Case 18    '�û�ʡ������S�
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
                'netca֤��������չOID:NETCA OID(1.3.6.1.4.1.18760.1.12.12.2)
                '1��������֤��2������֤��3: ����֤��4������Ա��֤��5������ҵ��֤��(ע�������͹��ܱ�׼����)0������֤��
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
                'region ����CN���O���ж�
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
         Case 23    '�û�֤��ͷ���
            strTmp = GetX509CertificateInfo(objCert, CATITLE)
            If strTmp = "NETCA" Then
                On Error Resume Next
                'netca �û�֤������OID
                strRet = mobjUtil.DecodeASN1String(1, objCert.GetExtension("1.3.6.1.4.1.18760.1.14"))
                Err.Clear: On Error GoTo 0
            ElseIf strTmp = "GDCA" Then
                strRet = GetX509CertificateInfo(objCert, 51)
            Else
                strRet = ""
            End If
         Case 24    '���ڵر�
            On Error Resume Next
            '���ڵر�
            strRet = mobjUtil.DecodeASN1String(1, objCert.GetExtension("2.16.156.112548"))
            Err.Clear: On Error GoTo 0
        Case 31 '֤���ķӡ
            On Error Resume Next
            strTmp = objCert.GetStringInfo(29)
            If strTmp <> "" Then strRet = ConvertHex(strTmp)
            Err.Clear: On Error GoTo 0
        Case 36 'netca֤��������չOID
            On Error Resume Next
            strRet = mobjUtil.DecodeASN1String(1, objCert.GetExtension(NETCAPKI_UUID))
            Err.Clear: On Error GoTo 0
        Case 37     '֤���������
            strRet = GetX509CertificateInfo(objCert, 36)
            If Len(strRet) > 13 Then
             '00011@0006PO1MTIzNDU2Nzg5MDEyMzQ1Njc4
                i = InStr(strRet, "@")
                
                If i = 0 Then
                    strRet = "": GoTo ReturnLine
                End If
                strTmp = Mid(strRet, i + 7, 1) '��ȡ�����־λ
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
        Case 51 'GDCA���ض���չ�� 51
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
    MsgBoxEx Err.Description & vbCrLf & "GetX509CertificateInfo" & "�� " & Erl, vbExclamation + vbOKOnly, gstrSysName
End Function

Private Function ClearCertHeader(ByVal strCertBase As String) As String
'����: ȥ��֤��Base64�����ͷβ����
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
    MsgBoxEx Err.Description & vbCrLf & "signedDataWithTSA" & "�� " & Erl, vbExclamation + vbOKOnly, gstrSysName
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
'����:�ֽ�����תHex�����ַ���
    On Error GoTo errH
    ConvertHex = mobjUtil.BinaryToHex(objData, True)
    Exit Function
errH:
    MsgBoxEx Err.Description & vbCrLf & "ConvertHex" & "�� " & Erl, vbExclamation + vbOKOnly, gstrSysName
 End Function
        
Private Function Base64Decode(ByVal strData As String) As Byte()
'����:�ַ���Base64����Ϊ�ֽ�����
    On Error GoTo errH
    Base64Decode = mobjUtil.Base64Decode(strData, NETCAPKI_BASE64_ENCODE_NO_NL)
    Exit Function
errH:
    MsgBoxEx Err.Description & vbCrLf & "Base64Decode" & "�� " & Erl, vbExclamation + vbOKOnly, gstrSysName
 End Function
 
 Private Function Base64Encode(bytData As Variant) As String
'����:�ֽ�����Base64����Ϊ�ַ���
    
    On Error GoTo errH
    Base64Encode = mobjUtil.Base64Encode(bytData, NETCAPKI_BASE64_ENCODE_NO_NL)
    Exit Function
errH:
    MsgBoxEx Err.Description & vbCrLf & "Base64Encode" & "�� " & Erl, vbExclamation + vbOKOnly, gstrSysName
 End Function
        
Private Function ConvertByte(ByVal strData As String) As Byte()
'����:�ַ���ת�ֽ�����
    
    On Error GoTo errH

    ConvertByte = mobjUtil.Encode(strData, NETCAPKI_CP_UTF8)
                
    Exit Function
errH:
    MsgBoxEx Err.Description & vbCrLf & "ConvertByte" & "�� " & Erl, vbExclamation + vbOKOnly, gstrSysName
End Function

Private Function GetRandom(ByVal intLength As Integer) As String
'����:��ȡ�����
    Dim objDevice As Object
    Dim bytRandom() As Byte
    Dim varTemp As Variant
    
    On Error GoTo errH
    Set objDevice = CreateObject("NetcaPki.Device.1")
    If objDevice Is Nothing Then
        MsgBoxEx "������֤ͨ����Device��ʧ�ܣ�", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    bytRandom = objDevice.GenerateRandom(intLength)
    varTemp = bytRandom
    GetRandom = ConvertHex(varTemp)
    Exit Function
errH:
    MsgBoxEx Err.Description & vbCrLf & "GetRandom" & "�� " & Erl, vbExclamation + vbOKOnly, gstrSysName
End Function

Private Function SignedDataByPwd(ByVal strSource As String, ByVal blnNoHasSource As Boolean, ByRef strErr As String) As String
'����: ��PIN��PKCS7ǩ��
    Dim bytArr() As Byte
    
    On Error GoTo errH

    bytArr = ConvertByte(strSource)
    
    SignedDataByPwd = SignedDataByCertificate(mobjCert, bytArr, blnNoHasSource, strErr)

    Exit Function
errH:
    MsgBoxEx Err.Description & vbCrLf & "SignedDataByPwd" & "�� " & Erl, vbExclamation + vbOKOnly, gstrSysName
End Function

Private Function SignedDataByCertificate(ByVal objCert As Object, ByRef bytSource() As Byte, ByVal blnNoHasSource As Boolean, ByRef strErr As String) As String
'����:ʹ��֤�����PKCS7ǩ��[�������ʵ��]
'����:blnNoHasSource=false ��ԭ��ǩ��(ȱʡ) ;Ture=����ԭ��ǩ��
    Dim varRet As Variant
    
    On Error GoTo errH
    If (mobjSign.SetSignCertificate(mobjCert, "", False) = False) Then
        strErr = "����ǩ��֤��ʧ��"
        Exit Function
    End If
    Call mobjSign.SetSignAlgorithm(-1, IIf(GetX509CertificateInfo(objCert, 8) = NETCAPKI_ALGORITHM_RSA & "", NETCAPKI_ALGORITHM_RSASIGN, NETCAPKI_ALGORITHM_SM2SIGN))
    Call mobjSign.SetIncludeCertificateOption(NETCAPKI_SIGNEDDATA_INCLUDE_CERT_OPTION)
    mobjSign.Detached = blnNoHasSource  'true������ԭ�ģ�false����ԭ��
    varRet = bytSource
    varRet = mobjSign.Sign(varRet, NETCAPKI_CMS_ENCODE_BASE64)

    SignedDataByCertificate = varRet
    Exit Function
errH:
    MsgBoxEx Err.Description & vbCrLf & "SignedDataByCertificate" & "�� " & Erl, vbExclamation + vbOKOnly, gstrSysName
End Function
        
Private Function ParseXML(ByVal strData As String) As Collection
'����:����XML�ַ���
' <summary>����XML�ַ���
' <verifytime>20170824110405</verifytime>
'<certsha1hex>d52dbe11084d68f94fe57c13436883f128039f99</certsha1hex>
'<certname>����Ա44&&��֤ͨ</certname>
'<status>3</status>
' </summary>
' <param name="sResponse"></param>
' <returns></returns>
    Dim colList As New Collection
    Dim xmlDoc As New DOMDocument
    Dim xmlList As IXMLDOMNodeList
    Dim strValue As String
    
    On Error GoTo errH

    '��ȡ������Ӧ���ݣ�XML��ʽ��
    xmlDoc.loadXML (strData)
    Set xmlList = xmlDoc.selectNodes("VerifyCertResp")
    '֤����֤ʱ��
    strValue = xmlList(0).selectSingleNode(".//verifytime").Text
    colList.Add strValue, "verifytime"

    '֤��ժҪ
    strValue = xmlList(0).selectSingleNode(".//certsha1hex").Text
    colList.Add strValue, "certsha1hex"

    '֤������
    strValue = xmlList(0).selectSingleNode(".//certname").Text
    colList.Add strValue, "certname"

    '��ȡ֤��״̬��
    strValue = xmlList(0).selectSingleNode(".//status").Text
    colList.Add strValue, "status"

    '��ȡ������ǩ��ֵ ��ǩ������
    strValue = xmlDoc.selectSingleNode(".//signature").Text
    colList.Add strValue, "signature"
    strValue = "<?xml version=""1.0"" encoding=""UTF-8""?>" & xmlDoc.selectSingleNode(".//data").xml
    colList.Add strValue, "data"
    
    If colList.Count > 0 Then
        Set ParseXML = colList
    End If
    Exit Function
errH:
    MsgBoxEx Err.Description & vbCrLf & "ParseXML" & "�� " & Erl, vbExclamation + vbOKOnly, gstrSysName
End Function
        
Private Function ParseCertStatus(ByVal lngCertCode As Long) As String
'����:����֤��״̬��
    Dim arrCertCode As Variant
    Dim strRet As String
    
    On Error GoTo errH

    arrCertCode = Array("֤����Ч", "��֤����ʧ��", "֤���ʽ����", _
                        "֤�鲻����Ч����", "֤�鲻�����ڵ���ǩ��", _
                        "֤�����ֲ���", "֤����Բ���", "֤����չ����", _
                        "����֤��������", "֤�鱻ע��", "ע��״̬δ֪", _
                        "�û�֤��δ��Ȩ", "�û�״̬������")

    If lngCertCode >= 0 And lngCertCode < 13 Then
        strRet = arrCertCode(lngCertCode)
    Else
        strRet = "��������"
    End If
    ParseCertStatus = strRet
    Exit Function
errH:
    MsgBoxEx Err.Description & vbCrLf & "ParseCertStatus" & "�� " & Erl, vbExclamation + vbOKOnly, gstrSysName
End Function
        
Private Function EnCodeURL(ByVal strUrl As String) As String
'����:�������ַ�����UTF���뷽ʽת����ʮ�����Ƶ�ת������
'˵����encodeURI-javaScript��������� ASCII ��ĸ�����ֽ��б��룬Ҳ�������Щ ASCII �����Ž��б��룺 - _ . ! ~ * ' ( )
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
    MsgBoxEx Err.Description & vbCrLf & "EnCodeURL" & "�� " & Erl, vbExclamation + vbOKOnly, gstrSysName
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
     '֤�鲿����ʼ��
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
        
114 gstrPara = gobjComLib.zlDatabase.GetPara(90000, glngSys, , "") '��ȡ��������
116 LogWrite "WZT_InitObj", "CA����:" & gstrPara
118 If gstrPara = "" Then
120     Err.Raise -1, , "��ǰϵͳ��" & glngSys & "��û�����õ���ǩ������,�뵽���õ���ǩ���ӿڴ����á�"
        Exit Function
    End If
122 Call WZT_GetPara
124 mblnInit = True
126 WZT_InitObj = True
   
    Exit Function
errMsg:
128 MsgBoxEx "������֤ͨ����ʧ�ܣ���ȷ�ϡ���֤ͨ��ȫ�ͻ��ˡ��Ѱ�װ�Ҹò�����NetcaPkiCom.dll����ע��ɹ���" & vbCrLf & "������:" & Erl, vbExclamation, gstrSysName
    Exit Function
errH:
130  MsgBoxEx "�����ӿڲ���ʧ�ܣ�" & vbNewLine & "��" & CStr(Erl()) & "�� " & Err.Description, vbExclamation, gstrSysName
End Function

Public Function WZT_Sign(ByVal strSource As String, ByRef strSignData As String, ByRef strTimeStamp As String, ByRef strTimeStampCode As String) As Boolean
'����:ǩ��ǩ��ֵ����ǩ��ֵ��Ϣ��ʱ�����Ϣ;��ԭ��ǩ��;ԭ��Խ��ǩ������ֵԽ��
    Dim strUrl As String
    Dim blnCheck As Boolean
    Dim lngSignID As Long
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim strErr As String
    
    On Error GoTo errH
    blnCheck = WZT_CheckCert()
    If blnCheck Then                 '��֤��ǰUSB�Ƿ���ǩ���û��ģ�����ȡǩ��֤��
        '֤��ID����ǩ��
        If gudtPara.blnISTS Then
            strUrl = "http://" & gudtPara.strTSIP & IIf(gudtPara.strTSPort = "", "", ":" & gudtPara.strTSPort) & "/NETCATimeStampServer/TSAServer.jsp"
            LogWrite "WZT_Sign", "ʱ�����ַ:" & strUrl
            strSignData = SignedDataWithTSA(strSource, strUrl, , strErr)  '��ԭ��ǩ��
            LogWrite "WZT_Sign", "ǩ������ֵ:" & strSignData
            If strSignData <> "" Then
                If Not VerifySignedDataWithTSA(strSource, strSignData, strTimeStamp) Then
                    Exit Function
                End If
            Else
                MsgBoxEx "ǩ��ʧ��!" & IIf(strErr <> "", vbCrLf & "ԭ��:" & vbCrLf & strErr & "!", ""), vbInformation, gstrSysName
                Exit Function
            End If
        Else
            strSignData = SignedDataByPwd(strSource, False, strErr)
            LogWrite "WZT_Sign", "ǩ������ֵ:" & strSignData
            If strSignData = "" Then
                MsgBoxEx "ǩ��ʧ��!" & IIf(strErr <> "", vbCrLf & "ԭ��:" & vbCrLf & strErr & "!", ""), vbInformation, gstrSysName
                Exit Function
            End If
            strTimeStamp = Format(gobjComLib.zlDatabase.Currentdate & "", "yyyy-MM-dd HH:mm:ss")
        End If
    Else
        MsgBoxEx "ǩ��ʧ�ܣ�", vbInformation, gstrSysName
        Exit Function
    End If

    strSQL = "Select ����ǩ����¼_ID.Nextval as ID From Dual"
    Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, "WZT_Sign")
    lngSignID = Val(rsTmp!id)
    
    If WZT_SaveLob(lngSignID, strSignData) Then
        strSignData = lngSignID
    Else
        strSignData = ""
        MsgBoxEx "ǩ��ʧ�ܣ�", vbInformation, gstrSysName
        Exit Function
    End If
    WZT_Sign = True
    Exit Function
errH:
114     MsgBoxEx "ǩ��ʧ�ܣ�" & vbNewLine & "��" & CStr(Erl()) & "�� " & Err.Description, vbExclamation, gstrSysName
End Function

Public Function WZT_VerifySign(ByVal strSource As String, ByVal strSignData As String, Optional ByVal strTimeStampCode As String) As Boolean
    '����;��֤ǩ��
    '����:strSignData -ǩ��ֵ
    Dim blnRet As Boolean
    
    On Error GoTo errH
    LogWrite "WZT_VerifySign", "��֤ǩ��ԭ��:" & strSource & vbCrLf & "��֤ǩ��ֵ:" & strSignData & vbCrLf & "ǩ��ʱ�����Ϣ:" & strTimeStampCode
    strSignData = WZT_ReadLob(CLng(strSignData))
    If gudtPara.blnISTS Then
        blnRet = VerifySignedDataWithTSA(strSource, strSignData, "")
    Else
        blnRet = VerifySignedData(strSource, strSignData)
    End If
    If blnRet Then
        MsgBoxEx "��֤�ɹ����õ���ǩ��������Ч!", vbInformation, gstrSysName
    Else
         Exit Function
    End If
    WZT_VerifySign = True
    Exit Function
errH:
140     MsgBoxEx "��֤ǩ��ʧ�ܣ�" & vbNewLine & "��" & CStr(Erl()) & "�� " & Err.Description, vbExclamation, gstrSysName

End Function

Public Function WZT_CheckCert() As Boolean
    '���ܣ���ȡUSB�����豸��ʼ������¼
    Dim strKeySN As String, strUserName As String
    Dim colList As Collection
    Dim datBegin As Date
    Dim strUrl As String
    Dim lngRet As Long
    Dim strMsg As String
    
    On Error GoTo errH
    
    If Not GetCertList(strKeySN, strUserName, , "") Then Exit Function
    If mUserInfo.strCertSn <> strKeySN Then
        MsgBoxEx "��֤�飺" & _
           vbCrLf & vbTab & "��" & mUserInfo.strCertSn & "��" & vbCrLf & _
           "��ǰ֤��:" & vbCrLf & vbTab & "��" & strKeySN & "��" & vbCrLf & _
           "��ǰ֤���û���:" & vbCrLf & vbTab & "��" & strUserName & "��" & vbCrLf & _
           "��֤���뵱ǰ֤�鲻��ͬһ��֤��,����ʹ�ã���", vbInformation, gstrSysName
        Set mobjCert = Nothing
        Exit Function
    End If
    datBegin = gobjComLib.zlDatabase.Currentdate
    '����ķֵ��֤�������Ӱ����ķֵ
    strUrl = "http://" & gudtPara.strSIGNIP & ":" & gudtPara.strSignPort & "/NetcaCertAA/appintf/verifyusercert"
    Set colList = VerifyCert(mobjCert, strUrl, Format(datBegin, "HH:mm:ss"), "c0", gudtPara.strOption)
    If colList Is Nothing Then
        MsgBoxEx "��֤ʧ��!", vbInformation + vbOKOnly, gstrSysName: Exit Function
        Exit Function
    End If
    lngRet = Val(colList("status"))
    If lngRet <> 0 Then
        strMsg = ParseCertStatus(lngRet)
        MsgBoxEx "������֤ʧ��!" & IIf(strMsg <> "", "ԭ��:" & strMsg, ""), vbInformation + vbOKOnly, gstrSysName: Exit Function
        Exit Function
    End If
    WZT_CheckCert = True
    Exit Function
errH:
     MsgBoxEx "���USBKEYʧ�ܣ�" & vbNewLine & "��" & CStr(Erl()) & "�� " & Err.Description, vbExclamation, gstrSysName
End Function


Public Function WZT_RegCert(arrCertInfo As Variant) As Boolean
    '���ܣ��ṩ��HIS���ݿ���ע�����֤��ı�Ҫ��Ϣ,�����·��Ż����֤��,,��Ҫ����USB-Key
    '���أ�arrCertInfo��Ϊ���鷵��֤�������Ϣ
'      0-ClientSignCertCN:�ͻ���ǩ��֤�鹫������(����),����ע��֤��ʱ������֤���
'      1-ClientSignCertDN:�ͻ���ǩ��֤������(ÿ��Ψһ)
'      2-ClientSignCertSN:�ͻ���ǩ��֤�����к�(ÿ֤��Ψһ)
'      3-ClientSignCert:�ͻ���ǩ��֤������
'      4-ClientEncCert:�ͻ��˼���֤������
'      5-ǩ��ͼƬ�ļ���,�մ���ʾû��ǩ��ͼƬ

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
     MsgBoxEx "��ȡ֤����Ϣʧ�ܣ�" & vbNewLine & "��" & CStr(Erl()) & "�� " & Err.Description, vbExclamation, gstrSysName

End Function

Private Function GetCertList(Optional ByRef strUniqueID As String = "-1", Optional ByRef strName As String = "-1", _
    Optional ByRef strCertDN As String = "-1", Optional ByRef strDate As String = "-1", Optional ByRef strPic As String = "-1") As Boolean
    '����:��ȡ��֤֤ͨ������
        Dim varTemp As Variant
        Dim lngDay As Long
        Dim varRet  As Variant
        Dim strBase64 As String
        
        On Error GoTo errH
    '        If mobjCert Is Nothing Then
100     Set mobjCert = GetX509Certificate(NETCAPKI_CERT_PURPOSE_SIGN) 'ǩ��֤��
    '        End If
        
102     If mobjCert Is Nothing Then
104         MsgBoxEx "û�з��ϵ�֤�鹩ѡ��,�����Ƿ����KEY!", vbInformation + vbOKOnly, gstrSysName
            Exit Function
        End If
    '        varTemp = Array("֤��PEM����", "֤��ķӡ", "֤�����к�", "֤������", "֤��䷢������", _
    '                "֤����Ч����", "֤����Ч��ֹ", "��Կ�÷�", "֤�鹫Կ�㷨", "�û�֤���ֵ", _
    '                "�ɵ��û�֤���ֵ", "֤����������", "Subject������(CN)", "Subject�еĵ�λ(O)", "Subject�еĵ�ַ��(L)", _
    '                "Subject�е�Email(E)", "Subject�еĲ�����(OU)", "������(C)", "ʡ����(S)", "", _
    '                "", "CA��־", "֤������", "�û�֤������", "���ڵر�OID", _
    '                "", "", "", "", "", _
    '                "", "NETCA:��ķӡ��Ϣ", "��˰�˱���", "��ҵ���˴���", "˰��ǼǺ�", _
    '                "֤����Դ��", "֤��������Ϣ(NETCA ֤��Ψһ��ʶ)", "֤����������", "", "", _
    '                "", "", "", "", "", _
    '                "", "", "", "", "", _
    '                "", "GDCA���κ�TrustID")
            
106     If strUniqueID = "" Then strUniqueID = GetX509CertificateInfo(mobjCert, 1) '֤��ķӡ
108     If strCertDN = "" Then strCertDN = GetX509CertificateInfo(mobjCert, 3) '֤������DN
110     If strName = "" Then strName = GetX509CertificateInfo(mobjCert, 12) 'Subject������(CN)
112     If strDate = "" Then strDate = GetX509CertificateInfo(mobjCert, 6) '֤����Ч��ֹ
  
114     If strPic = "" Then
            
116         varRet = mobjPDFSign.SelectCert("", 1)
118         varRet = mobjPDFUtilTool.GetImageFromDevicByCert(mobjPDFSign.SignCertBase64Encode)
120         strBase64 = Base64Encode(varRet)
122         If strBase64 <> "" Then
124             strPic = SaveBase64ToFile("gif", strUniqueID, strBase64)
            End If
        End If
126     If IsDate(strDate) Then
            '���֤���Ƿ����
128         lngDay = CheckValidaty(strDate)
130         If (lngDay <= 30 And lngDay > 0 And Not gblnShow) Then
132             MsgBoxEx "����֤�黹��" & lngDay & "�����", vbInformation, gstrSysName
134             gblnShow = True
136         ElseIf (lngDay <= 0) Then
138             MsgBoxEx "����֤���ѹ��� " & Abs(lngDay) & " ��"
                Exit Function
            End If
        End If
        
140     GetCertList = True
        Exit Function
errH:
142     MsgBoxEx "��ȡ֤����Ϣʧ�ܣ�" & vbNewLine & "��" & CStr(Erl()) & "�� " & Err.Description, vbExclamation, gstrSysName
End Function

Public Sub WZT_GetPara()
        Dim arrList As Variant

        On Error GoTo errH
100     gstrPara = gobjComLib.zlDatabase.GetPara(90000, glngSys, , "")
        If gstrPara = "" Then
            'ǩ��������IP&&&ǩ���������˿ں�&&&ʱ���IP&&&ʱ����˿ں�&&&����֤��&&&����ʱ�����0/1��
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
146     MsgBoxEx "��ȡ����ʧ�ܣ�" & vbNewLine & "��" & CStr(Erl()) & "�� " & Err.Description, vbInformation, gstrSysName
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
'���ܣ���ȡ����ǩ����Ϣ
    Dim rsLob As ADODB.Recordset
    Dim lngCount As Long
    Dim strText As String
    Dim strSQL As String
    Dim strFile As String
    
    Err = 0: On Error GoTo Errhand
    strSQL = "Select ZL_Readlob_����ǩ����¼([1],[2]) as Ƭ�� From Dual"
    If strSQL = "" Then strFile = "": Exit Function
    lngCount = 0
    strFile = ""
    Do
        Set rsLob = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, "ZL_Readlob_����ǩ����¼", lngID, lngCount)
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
'���ܣ�����ǩ����Ϣ
'������
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
'���ܣ���������ָ�����ļ���ָ�����¼BLOB/CLOB�ֶε�SQL���
'������
'      KeyWord:ȷ�����ݼ�¼�Ĺؼ��֣����Ϲؼ����Զ��ŷָ�(��5-���Ӳ�����ʽΪ����)
'      strFile: CLOBʱ,��Ҫ�洢���ı�����
'      arySql():�ڸ����ݵĻ�������չ���ӱ����SQL���
'���أ��ɹ�����True��ʧ�ܷ���False
    Dim conChunkSize As Integer
    Dim lngCount As Long, lngBound As Long
    Dim aryChunk() As Byte, aryHex() As String, strText As String
    Dim strSQLRoot As String
    Dim strSubTxt As String
    
    On Error GoTo Errhand

    strSQLRoot = "ZL_AppendLob_����ǩ����¼(" & lngKeyWord
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


