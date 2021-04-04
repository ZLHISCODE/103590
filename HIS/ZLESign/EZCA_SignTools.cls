VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EZCA_SignTools"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private EZCAClientTools As Object
Private CapicomCert As Object
Private SoapClient As Object

Private SoapAction As String
'wsdl地址
Private END_POINT_URL As String

'返回用户标识
Public Function WF_SelectCert() As String
    Dim ls_ret, user_id As String
    Dim lstr() As String
    Dim users() As String
    Dim Cert_User As String
    
     '1检测控件是否支持
    ls_ret = EZCAClientTools.SOF_CheckSupport()
    If ls_ret <> 0 Then Exit Function

    '2取用户列表及选择用户
    ls_ret = EZCAClientTools.SOF_GetUserList()
    If ls_ret = "" Then
        MsgBoxEx "请插入ＵＳＢＫＥＹ！"
        Exit Function
    Else
        If InStr(ls_ret, "&&&") > 0 Then    '多个用户，
            users = Split(ls_ret, "&&&")
            '选择用户
            Cert_User = frmSelectUser.ShowMe(ls_ret)
        Else
            Cert_User = ls_ret
        End If
    End If
    
    '3再检测用户是否正确选择，并取用户唯一标示
    If Cert_User = "" Then Exit Function

    lstr = Split(Cert_User, "||")
    If UBound(lstr) < 1 Then Exit Function
    If lstr(1) = "" Then Exit Function
    
    user_id = lstr(1)
    
    WF_SelectCert = user_id
End Function

'签名并保存
Public Function WF_SignData(bussinessid As String, sdata As String, CertID As String) As String
'参数：bussinessId 业务唯一标识，sdata 被签名原文，timeStampState 是否需要时间戳(1需要0不需要），signType签名数据类型(0:字符串，1：文件,2:XML)，personid病人ID号
'返回：整型，0表示成功，其他失败
sdata = EncodeBase64String(sdata)  '首先对原文进行base64编码
Const timestampstate = 0
Const signtype = 0
Const personid = ""
 
    Dim li_result As Integer
    Dim ls_ret, signed_data As String
    Dim user_id, cert_id As String
    Dim lstr() As String
    Dim users() As String
    Dim Cert_User As String
    Dim strEdate As String
    
     '1检测控件是否支持
    ls_ret = EZCAClientTools.SOF_CheckSupport()
    If ls_ret = 0 Then
        'MsgboxEx "test::" & "支持"
    Else
        'MsgboxEx "test::" & "不支持" 不支持签名方法，退出
        WF_SignData = 1000
        Exit Function
    End If
    '2取用户列表及选择用户
     ls_ret = EZCAClientTools.SOF_GetUserList()
    'MsgboxEx "test::" & ls_ret
    
    If ls_ret = "" Then
        MsgBoxEx "请插入ＵＳＢＫＥＹ！"
        WF_SignData = 1001
        Exit Function
    Else
        If InStr(ls_ret, "&&&") > 0 Then    '多个用户，
            users = Split(ls_ret, "&&&")
            '选择用户
            Cert_User = frmSelectUser.ShowMe(ls_ret)
        Else
            Cert_User = ls_ret
        End If
    End If
     '3再检测用户是否正确选择，并取用户唯一标示
    If Cert_User = "" Then
        WF_SignData = 1002
        Exit Function
    Else
        lstr = Split(Cert_User, "||")
        If UBound(lstr) < 1 Then
            WF_SignData = 1003
            Exit Function
        End If
        If lstr(1) <> "" Then
            user_id = lstr(1)
        Else
            WF_SignData = 1004
            Exit Function
        End If
    End If
    
    '用户与Key匹配性检查
    If user_id <> CertID Then
        MsgBoxEx "ＵＳＢＫＥＹ与人员不符！"
        WF_SignData = 1001
        Exit Function
    End If
    '2018-4-13 重庆修改过期时间提醒,一次登录只需要提醒一次
    If Not gblnShow Then
        strEdate = EZCAClientTools.SOF_GetCertInfo(EZCAClientTools.SOF_ExportUserCert(user_id), 18)
        If DateDiff("d", Date, Format(strEdate, "yyyy-mm-dd")) <= 30 Then
            If MsgBoxEx("你得证书还有" & DateDiff("d", Date, Format(strEdate, "yyyy-mm-dd")) & "天过期,请尽快到信息科更新!本次登录是否不在提醒？", vbYesNo) = vbYes Then
                gblnShow = True
            End If
        End If
    End If
    '取证书序列号
    cert_id = get_Cert_Id(user_id)
    'MsgboxEx cert_id
    '4调用签名方法
     signed_data = EZCAClientTools.SOF_SignData(user_id, sdata)
     If signed_data = "" Then
        MsgBoxEx "签名失败"
        WF_SignData = 1005
        Exit Function
    End If
    'MsgboxEx signed_data
     '5调用数据保存方法
    'WF_DataSave(业务唯一ID，签名原文，签名值，是否需要时间戳(1需要0不需要），签名数据类型(0:字符串，1：文件,2:XML)，签名方法类型(没有更改默认为空)，签名证书序列号，用户的唯一ID号，签名值类型 （0:P7,1:P1），病人id）
    ls_ret = WF_DataSave(bussinessid, sdata, signed_data, "0", "0", "", cert_id, user_id, "1", Val(personid))
    
    If ls_ret = 10000 Then
        '保存成功
        WF_SignData = signed_data
    Else
        Select Case ls_ret
            Case "10001"
                MsgBoxEx "签名存储失败,错误码:" & ls_ret
            Case "10002"
                MsgBoxEx "签名数据验证失败,错误码:" & ls_ret
            Case "10003"
                MsgBoxEx "证书无效,错误码:" & ls_ret
            Case "10004"
                MsgBoxEx "证书过期,错误码:" & ls_ret
        End Select
        WF_SignData = ls_ret
    End If
 End Function
 
 '取证书序列号
 Private Function get_Cert_Id(ByVal userid As String) As String
 '参数：用户ID  返回：用户证书的序列号
   
    Dim pubcert As String
    pubcert = EZCAClientTools.SOF_ExportUserCert(userid)
    get_Cert_Id = EZCAClientTools.SOF_GetCertInfo(pubcert, 2)

 
 End Function
'调用数据保存的WEBSERVICE方法
Public Function WF_DataSave(ByVal bussinessid As String, ByVal InData As String, ByVal signed_data As String, ByVal timestampstate As String, ByVal signtype As String, ByVal signMethodType As String, ByVal SignCertSerialNo As String, ByVal appid As String, ByVal signdatatype As String, ByVal personid As String) As String
    WF_DataSave = SoapClient.VD_DataSave(bussinessid, InData, signed_data, timestampstate, signtype, signMethodType, SignCertSerialNo, appid, signdatatype, personid)
End Function
'验证存储数据库的电子签名
Public Function WF_VerifyServerSigneData(businessId As String, SignCertSerialNo As String, signdatatype As String) As String
'BusinessId 业务系统的唯一ID号,SignCertSerialNo签名证书序列号,signDataType签名值类型 （0:P7,1:P1）
    WF_VerifyServerSigneData = SoapClient.VD_VerifyServerSigneData(businessId, SignCertSerialNo, signdatatype)

End Function

'调用服务器来验证电子签名
Public Function WF_VerifySigneDataByServer(ByVal InData As String, ByVal SignData As String, ByVal signMethodType As String, ByVal signdatatype As String, ByVal signCert As String, ByVal signtype As String) As String
'BusinessId 业务系统的唯一ID号,SignCertSerialNo签名证书序列号,signDataType签名值类型 （0:P7,1:P1）

    WF_VerifySigneDataByServer = SoapClient.VD_VerifySigneData(InData, SignData, signMethodType, signdatatype, signCert, signtype)

End Function

'调用客户端控件进行签名验证
Public Function WF_VerifySigneData(ByVal cert As String, ByVal InData As String, ByVal sSignValue As String)
    InData = EncodeBase64String(InData)  '首先对原文进行base64编码
    WF_VerifySigneData = EZCAClientTools.SOF_VerifySignedData(cert, InData, sSignValue)
End Function

'通过证书序列号或用户唯一编号，获取用户对应的公钥证书。两个参数一次只能使用期中一种
Public Function WF_GetUserCert(UID As String, IDType As String) As String
'UID 序号编码,IDType编码类型（01：证书序列号；02：用户ID号）
   WF_GetUserCert = SoapClient.VD_GetUserCert(UID, IDType)
End Function

'传入需要签名的原文数据或HASH值，返回Base64编码的签名值
Public Function WF_SignService(InData As String, signtype As String) As String
'inDate[IN]  原文数据或HASH值（传入参数均转为Base64编码）
'SignType    原数据类型，1：原文；2：HASH值
    WF_SignService = SoapClient.VD_SignService(InData, signtype)
End Function

'服务器端的签名验证
Public Function WF_VerifyService(InData As String, SignData As String, certbase64 As String, signtype As String) As String
'inDate 原文数据或HASH值（传入参数均转为Base64编码）,SignData 签名值,certbase64 签名证书BASE64编码,signtype 签名值类型P7为0，P1为1
    WF_VerifyService = SoapClient.VD_VerifyService(InData, SignData, certbase64, signtype)

End Function

'测试方法
Public Function Of_test(bussinessid As String, sdata As String, timestampstate As String, signtype As String, personid) As Integer

''参数：bussinessId 业务唯一标识，sdata 被签名原文，timeStampState 是否需要时间戳(1需要0不需要），signType签名数据类型(0:字符串，1：文件,2:XML)，personid病人ID号
''返回：整型，0表示成功，其他失败

 
Dim li_result As Integer
Dim sd_p7 As String
Dim GS_CertID As String
Dim cert_id As String
Dim signed_data As String
Dim ls_ret As String

''直接赋值
GS_CertID = "EZCA@5014990setup"
cert_id = "10d755cfb9ba15023852aabd88887ee9"
'signed_data = "mIyRj3yyQaNZxv/3qhVp8FQQdoTNXAEnmO/IbgwbqSZbGQvsX/IT0cmU2uhH+rMNMhffdHZTRyYcLDMAlcYqSv1B+6ENsqnNgehaSgF9qfI8HC6BpyriPDjUrtxdKtpl3iXKLVnwjAMEbNpb3kVZKIxYbSNNn27hyqX7DfIppg0="
'base64编码后的P1签名值
signed_data = "1mqsQoe/WuZ+ImXU4q5F7UeCAdbgZuAI+HwNm8M9opJ3eXUlQVY8YR0PBZjLv7rUB7/zF/FlZmXMWTazztPViQNyG55SU1KaIiNV8Ggua+bRpYfHoAM7ZnXDzD1H9LyAxvTQ3O+5unmk68dnetf9bagAZrxMbR8ocN7ckRAtu0o="
'sd_p7 = "MIIE1wYJKoZIhvcNAQcCoIIEyDCCBMQCAQExDzANBgkqhkiG9w0BAQUFADAmBgkqhkiG9w0BBwGgGQQXMzQzNDM01tC5+sjLw/GjrNKq09DLqgCgggNQMIIDTDCCArWgAwIBAgIQENdVz7m6FQI4Uqq9iIh+6TANBgkqhkiG9w0BAQUFADBxMQswCQYDVQQGEwJDTjESMBAGA1UECBMJQ0hPTkdRSU5HMTYwNAYDVQQKEy1DaG9uZ3FpbmcgQ2VydGlmaWNhdGUgQXV0aG9yaXR5IENlbnRlciBDTy5MVEQxFjAUBgNVBAMTDUNob25ncWluZyAgQ0EwHhcNMTEwNTE4MDc1MDI1WhcNMTIwNTE3MDc1MDI1WjA+MQswCQYDVQQGEwJDTjEPMA0GA1UECh4GW8xY65AaMQ0wCwYDVQQLHgRfAFPRMQ8wDQYDVQQDHgZbWVwPWh8wgZ8wDQYJKoZIhvcNAQEBBQADgY0AMIGJAoGBAMUDRoCEGMdDVEQDnGsJBZyc/hhoNWNDuvcG8Naq8UZgXNJJ3HUfkWi8Q5FyLWwfh/aiiHnHsq/Ymr3c03oi1ewKxiIwIRzf5CeF0H0CG88CU6yj55p9JD/JCzOjMXjT7yzbIOTFzhoMy" _
& "L6cooDEE7bYYsdXh2i+HUO8bazpm+3pAgMBAAGjggEWMIIBEjBpBggrBgEFBQcBAQRdMFswMAYIKwYBBQUHMAKGJGh0dHA6Ly93d3cuZGZ6eGNhLmNvbS9DQS9jYWlzc3VlLmh0bTAnBggrBgEFBQcwAYYbaHR0cDovL3d3dy5kZnp4Y2EuY29tOjIwNDQzMB8GA1UdIwQYMBaAFHrTXfuu1AQNTkXzgSsiItPP/FIWMAkGA1UdEwQCMAAwLwYDVR0fBCgwJjAkoCKgIIYeaHR0cDovL3d3dy5kZnp4Y2EuY29tL2NybDguY3JsMAsGA1UdDwQEAwIGwDAdBgNVHQ4EFgQUNwzOUFtyh3NIEX5UJQT/2nImNTIwHAYFKlYVAQE" _
& "EEwwRRVpDQUA1MDE0OTkwc2V0dXAwDQYJKoZIhvcNAQEFBQADgYEACh94eDDIJ0fgaf03/sBhI9OFcqnmQmxMuOhQD+/OAH+e9W0LcHKYi6+/ZtOcclWH8IIQezB3dxMcyMHBKY5vwf483mgpgid7sNAk/JZYjlva18xnUyrehiQWXpy49SPhwov41uFfCZEesKt4P7lAa3GGbEgB1p1t8Iulho2xO1ExggEwMIIBLAIBATCBhTBxMQswCQYDVQQGEwJDTjESMBAGA1UECBMJQ0hPTkdRSU5HMTYwNAYDVQQKEy1DaG9uZ3FpbmcgQ2VydGlmaWNhdGUgQXV0aG9yaXR5IENlbnRlciBDTy5MVEQxFjAUBgNVBAMTDUNob25ncWluZyAgQ0ECEBDXVc+5uhUCOFKqvYiIfukwDQYJKoZIhvcNAQEFBQAwDQYJKoZIhvcNAQEBBQAEgYANpinyDfulyuFun00jbViMKFlF3lvabAQDjPBZLcol3mXaKl3crtQ4POIqp4EuHDzyqX0BSlrogc2psg2h+0H9SirGlQAzLBwmR1N2dN8XMg2z+kfo2pTJ0RPyX+wLGVsmqRsMbsjvmCcBXM2EdhBU8GkVqvf/xlmjQbJ8j5GMmA=="
''
'base64编码后的P7签名值
sd_p7 = "MIIE4QYJKoZIhvcNAQcCoIIE0jCCBM4CAQExDzANBgkqhkiG9w0BAQUFADAwBgkqhkiG9w0BBwGgIwQhTXpRek5ETTAxdEM1K3NqTHcvR2pyTktxMDlETHFnPT0AoIIDUDCCA0wwggK1oAMCAQICEBDXVc+5uhUCOFKqvYiIfukwDQYJKoZIhvcNAQEFBQAwcTELMAkGA1UEBhMCQ04xEjAQBgNVBAgTCUNIT05HUUlORzE2MDQGA1UEChMtQ2hvbmdxaW5nIENlcnRpZmljYXRlIEF1dGhvcml0eSBDZW50ZXIgQ08uTFREMRYwFAYDVQQDEw1DaG9uZ3FpbmcgIENBMB4XDTExMDUxODA3NTAyNVoXDTEyMDUxNzA3NTAyNVowPjELMAkGA1UEBhMCQ04xDzANBgNVBAoeBlvMWOuQGjENMAsGA1UECx4EXwBT0TEPMA0GA1UEAx4GW1lcD1ofMIGfMA0GCSqGSIb3DQEBAQUAA4GNADCBiQKBgQDFA0aAhBjHQ1REA5xrCQWcnP4YaDVjQ7r3BvDWqvFGYFzSSdx1H5FovEORci1sH4f2ooh5x7Kv2Jq93NN6ItXsCsYiMCEc3+QnhdB9AhvPAlOso+eafSQ/yQszozF40+8s2yDkxc4a" _
& "DMi+nKKAxBO22GLHV4dovh1DvG2s6Zvt6QIDAQABo4IBFjCCARIwaQYIKwYBBQUHAQEEXTBbMDAGCCsGAQUFBzAChiRodHRwOi8vd3d3LmRmenhjYS5jb20vQ0EvY2Fpc3N1ZS5odG0wJwYIKwYBBQUHMAGGG2h0dHA6Ly93d3cuZGZ6eGNhLmNvbToyMDQ0MzAfBgNVHSMEGDAWgBR60137rtQEDU5F84ErIiLTz/xSFjAJBgNVHRMEAjAAMC8GA1UdHwQoMCYwJKAioCCGHmh0dHA6Ly93d3cuZGZ6eGNhLmNvbS9jcmw4LmNybDALBgNVHQ8EBAMCBsAwHQYDVR0OBBYEFDcMzlBbcodzSBF+VCUE/9pyJjU" _
& "yMBwGBSpWFQEBBBMMEUVaQ0FANTAxNDk5MHNldHVwMA0GCSqGSIb3DQEBBQUAA4GBAAofeHgwyCdH4Gn9N/7AYSPThXKp5kJsTLjoUA/vzgB/nvVtC3BymIuvv2bTnHJVh/CCEHswd3cTHMjBwSmOb8H+PN5oKYIne7DQJPyWWI5b2tfMZ1Mq3oYkFl6cuPUj4cKL+NbhXwmRHrCreD+5QGtxhmxIAdadbfCLpYaNsTtRMYIBMDCCASwCAQEwgYUwcTELMAkGA1UEBhMCQ04xEjAQBgNVBAgTCUNIT05HUUlORzE2MDQGA1UEChMtQ2hvbmdxaW5nIENlcnRpZmljYXRlIEF1dGhvcml0eSBDZW50ZXIgQ08uTFREMRYwFAYDVQQDEw1DaG9uZ3FpbmcgIENBAhAQ11XPuboVAjhSqr2IiH7pMA0GCSqGSIb3DQEBBQUAMA0GCSqGSIb3DQEBAQUABIGASrstEJHc3nAoH21MvGYAqG3913pnx+ukebq579zQ9MaAvPRHPczDdWY7A6DHh6XR5msuaPBVIyKaUlNSnhtyA4nV086zNlnMZWZl8RfzvwfUur/LmAUPHWE8VkEldXl3kqI9w5sNfPgI4Gbg1gGCR+1FruLUZSJ+5lq/h0KsatY="

If signed_data = "" Then
    '1检测控件是否支持
     li_result = EZCAClientTools.SOF_CheckSupport()
     If li_result = 0 Then
          '支持
     Else
          MsgBoxEx "不支持签名方法，退出"
          Of_test = 1000
     End If
    '2取用户列表及选择用户
    If GS_CertID = "" Then
         ls_ret = EZCAClientTools.SOF_GetUserList()
         If ls_ret = "" Then
              MsgBoxEx "请插入ＵＳＢＫＥＹ！"
              Of_test = 1001
         End If
    End If
    
    '3取证书序列号
     cert_id = EZCAClientTools.SOF_GetCertInfo(EZCAClientTools.SOF_ExportUserCert(GS_CertID), 2)
    
    
    ' '4调用签名方法
      signed_data = EZCAClientTools.SOF_SignData(GS_CertID, sdata)
      If signed_data = "" Then
          MsgBoxEx "签名失败"
          Of_test = 1005
     End If
End If
'  '5调用数据保存方法
' 'WF_DataSave(业务唯一ID，签名原文，签名值，是否需要时间戳(1需要0不需要），签名数据类型(0:字符串，1：文件,2:XML)，签名方法类型(没有更改默认为空)，签名证书序列号，用户的唯一ID号，签名值类型 （0:P7,1:P1），病人id）

 ls_ret = WF_DataSave(bussinessid, sdata, signed_data, "0", "0", "", cert_id, GS_CertID, "1", personid)
         'WF_DataSave(bussinessid , inData , signed_data , timestampstate , signtype , signMethodType , SignCertSerialNo , ByVal appid , signdatatype , personid )
 If ls_ret = "10000" Then
'    '保存成功
      Of_test = 0
 Else
      Select Case ls_ret
            Case "10001"
                MsgBoxEx "签名存储失败,错误码:" + ls_ret
            Case "10002"
                MsgBoxEx "签名数据验证失败,错误码:" + ls_ret
            Case "10003"
                 MsgBoxEx "证书无效,错误码:" + ls_ret
            Case "10004"
                MsgBoxEx "证书过期,错误码:" & ls_ret
        End Select
      Of_test = 1006
       MsgBoxEx "错误码:" + ls_ret
 End If
 
'P7签名测试
If sd_p7 = "" Then
    sd_p7 = EZCAClientTools.SOF_SignDataByP7(GS_CertID, sdata)
    'Form1.Text1.Text = sd_p7
End If
If signed_data = "" Then
    MsgBoxEx "签名失败"
    Of_test = 1007
End If
'  '5调用数据保存方法
' 'WF_DataSave(业务唯一ID，签名原文，签名值，是否需要时间戳(1需要0不需要），签名数据类型(0:字符串，1：文件,2:XML)，签名方法类型(没有更改默认为空)，签名证书序列号，用户的唯一ID号，签名值类型 （0:P7,1:P1），病人id）

 ls_ret = WF_DataSave(bussinessid, sdata, sd_p7, "0", "0", "", cert_id, GS_CertID, "0", personid)
         'WF_DataSave(bussinessid , inData , signed_data , timestampstate , signtype , signMethodType , SignCertSerialNo , ByVal appid , signdatatype , personid )
 If ls_ret = "10000" Then
'    '保存成功
      Of_test = 0
 Else
      Select Case ls_ret
            Case "10001"
                MsgBoxEx "签名存储失败,错误码:" + ls_ret
            Case "10002"
                MsgBoxEx "签名数据验证失败,错误码:" + ls_ret
            Case "10003"
                 MsgBoxEx "证书无效,错误码:" + ls_ret
            Case "10004"
                MsgBoxEx "证书过期,错误码:" & ls_ret
        End Select
      Of_test = 1008
       MsgBoxEx "错误码:" + ls_ret
 End If

'服务器端签名测试
signed_data = WF_SignService(sdata, "1")
If signed_data = "" Then
    MsgBoxEx "签名失败"
    Of_test = 1009
End If
'  '5调用数据保存方法
' 'WF_DataSave(业务唯一ID，签名原文，签名值，是否需要时间戳(1需要0不需要），签名数据类型(0:字符串，1：文件,2:XML)，签名方法类型(没有更改默认为空)，签名证书序列号，用户的唯一ID号，签名值类型 （0:P7,1:P1），病人id）

 ls_ret = WF_DataSave(bussinessid, sdata, signed_data, "0", "3", "", "", "", "1", personid)
         'WF_DataSave(bussinessid , inData , signed_data , timestampstate , signtype , signMethodType , SignCertSerialNo , ByVal appid , signdatatype , personid )
 If ls_ret = "10000" Then
'    '保存成功
      Of_test = 0
 Else
      Select Case ls_ret
            Case "10001"
                MsgBoxEx "签名存储失败,错误码:" + ls_ret
            Case "10002"
                MsgBoxEx "签名数据验证失败,错误码:" + ls_ret
            Case "10003"
                 MsgBoxEx "证书无效,错误码:" + ls_ret
            Case "10004"
                MsgBoxEx "证书过期,错误码:" & ls_ret
        End Select
       Of_test = 1010
       MsgBoxEx "错误码:" + ls_ret
 End If


 Of_test = 0
 
 
 
End Function
'设置webservice的URL地址
Public Function Of_SetUrl(ad As String)
    '初始化,服务器地址可放在服务器数据库中或配置文件中，以便于以后更换服务器地址

    'namespace
    SoapAction = ad '"http://127.0.0.1:8090/ezca_signserver/services/ezcawebservice"
    'wsdl地址
    END_POINT_URL = ad '+"?WSDL" ' "http://127.0.0.1:8090/ezca_signserver/services/ezcawebservice?WSDL"
End Function

'webservice连接对象的初始化
Public Function Of_init()
    SoapClient.MSSoapInit (SoapAction)
End Function

'取证书使用者
Public Function WF_GetCertOwner(CertID As String) As String
 '参数：CertId 用户唯一标识，
 'itype 取值类型
' CERT_SUBJECT    1   用户名
'CERT_UNIQUEID   2   证书实体唯一标识
'CERT_DEPT   3   部门
'CERT_ISSUE  4   颁发者DN
'CERT_DEVICETYPE 8   证书介质类型
 '返回：字符串，证书使用者
Dim itype As Integer
itype = 1
WF_GetCertOwner = EZCAClientTools.SOF_GetUserInfo(CertID, itype)
 
End Function

'取证书使用者
Public Function WF_GetCertDN(CertID As String) As String
 '参数：CertId 用户唯一标识，
 'itype 取值类型
'0 x00000001 证书版本
'0 x00000002 证书序列号
'0 x00000003 证书签名算法标识
'0 x00000004 证书颁发者国家(C)
'0 x00000005 证书颁发者组织名(O)
'0 x00000006 证书颁发者部门名(OU)
'0 x00000007 证书颁发者所在的省?自治区?直辖市(S)
'0 x00000008 证书颁发者通用名称(CN)
'0 x00000009 证书颁发者所在的城市?地区(L)
'0 x00000010 证书颁发者Email
'0 x00000011 证书有效期: 起始日期
'0 x00000012 证书有效期: 终止日期
'0 x00000013 证书拥有者国家(C)
'0 x00000014 证书拥有者组织名(O)
'0 x00000015 证书拥有者部门名(OU)
'0 x00000016 证书拥有者所在的省?自治区?直辖市(S)
'0 x00000017 证书拥有者通用名称(CN)
'0 x00000018 证书拥有者所在的城市?地区(L)
'0 x00000019 证书拥有者Email
'0 x00000020 证书颁发者DN
'0 x00000021 证书拥有者DN
'0 x00000022 证书公钥信息
'0 x00000023 证书扩展项信息
'0 x00000024 颁发者密钥标识符
'0 x00000025 证书持有者密钥标识符
'0 x00000026 密钥用途
'0 x00000027 私钥有效期
'0 x00000028 证书策略
'0 x00000029 策略映射
'0 x00000030 基本限制
'0 x00000031 策略限制
'0 x00000032 扩展密钥用途
'0 x00000033 CRL发布点
'0 x00000034 netscape属性
'0 x00000035 证书实体唯一标识
'0 x00000036 个人身份证号码
'0 x00000037 个人社会保险号
'0 x00000038 企业工商注册号
'0 x00000039 企业组织机构代码
'0 x00000040 税务登记证号
 '返回：字符串，证书主题
Dim itype As Integer
itype = 33
Dim cert_str As String
cert_str = EZCAClientTools.SOF_ExportUserCert(CertID)
WF_GetCertDN = EZCAClientTools.SOF_GetCertInfo(cert_str, itype)
End Function

'取证书序列号
Public Function WF_GetCertSN(CertID As String) As String
 '参数：CertId 用户唯一标识，
 'itype 取值类型
'0 x00000001 证书版本
'0 x00000002 证书序列号
'0 x00000003 证书签名算法标识
'0 x00000004 证书颁发者国家(C)
'0 x00000005 证书颁发者组织名(O)
'0 x00000006 证书颁发者部门名(OU)
'0 x00000007 证书颁发者所在的省?自治区?直辖市(S)
'0 x00000008 证书颁发者通用名称(CN)
'0 x00000009 证书颁发者所在的城市?地区(L)
'0 x00000010 证书颁发者Email
'0 x00000011 证书有效期: 起始日期
'0 x00000012 证书有效期: 终止日期
'0 x00000013 证书拥有者国家(C)
'0 x00000014 证书拥有者组织名(O)
'0 x00000015 证书拥有者部门名(OU)
'0 x00000016 证书拥有者所在的省?自治区?直辖市(S)
'0 x00000017 证书拥有者通用名称(CN)
'0 x00000018 证书拥有者所在的城市?地区(L)
'0 x00000019 证书拥有者Email
'0 x00000020 证书颁发者DN
'0 x00000021 证书拥有者DN
'0 x00000022 证书公钥信息
'0 x00000023 证书扩展项信息
'0 x00000024 颁发者密钥标识符
'0 x00000025 证书持有者密钥标识符
'0 x00000026 密钥用途
'0 x00000027 私钥有效期
'0 x00000028 证书策略
'0 x00000029 策略映射
'0 x00000030 基本限制
'0 x00000031 策略限制
'0 x00000032 扩展密钥用途
'0 x00000033 CRL发布点
'0 x00000034 netscape属性
'0 x00000035 证书实体唯一标识
'0 x00000036 个人身份证号码
'0 x00000037 个人社会保险号
'0 x00000038 企业工商注册号
'0 x00000039 企业组织机构代码
'0 x00000040 税务登记证号
 '返回：字符串，证书主题
Dim itype As Integer
itype = 2
Dim cert_str As String
cert_str = EZCAClientTools.SOF_ExportUserCert(CertID)
WF_GetCertSN = EZCAClientTools.SOF_GetCertInfo(cert_str, itype)
End Function

'取证书公钥
Public Function WF_GetPublicCert(CertID As String) As String
'参数：CertId 用户唯一标识，
'返回：字符串，证书公钥
WF_GetPublicCert = EZCAClientTools.SOF_ExportUserCert(CertID)
End Function

'显示证书信息
Public Function WF_ShowCert(CertID As String) As String
'参数：CertId 用户唯一标识，
Dim CertContent As String
CertContent = WF_GetPublicCert(CertID)
If CertContent <> "" Then
    CapicomCert.Import (CertContent)
    CapicomCert.Display
End If

End Function
 
Private Sub Class_Initialize()
Set SoapClient = CreateObject("MSSOAP.SoapClient30")  'SOAP连接对象，调WEBSERVICE用
Set CapicomCert = CreateObject("CAPICOM.Certificate")   '微软CAPICOM证书对象，显示证书信息用
Set EZCAClientTools = CreateObject("SANITATIONSYSTEMCLIENT.EZCASanitationSystemClient")  '东方中讯卫生客户端控件，客户端签名验签用

'webservice对象初始化
Dim strUrl As String
'strUrl = GetSetting("ZLSOFT", "注册信息", "CQCA_URL", "")
strUrl = ReadIni("CA", "URL", App.Path & "\TSA.ini")
'strUrl = "http://192.168.4.40:8090/ezca_signserver/services/ezcawebservice?wsdl"
If strUrl = "" Then
    Err.Raise -1, , "没有配置签名服务器地址，请先配置。"
    Exit Sub
End If
SoapAction = strUrl

SoapClient.MSSoapInit (SoapAction)
End Sub
