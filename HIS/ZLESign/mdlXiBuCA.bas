Attribute VB_Name = "mdlXiBuCA"
'---------------------------------------------------------------------------------------
' Module    : mdlXiBuCA
' Author    : YWJ
' Date      : 2019-06-27 23:32:35
' Purpose   :
'---------------------------------------------------------------------------------------
Option Explicit

Private Const M_STR_URL   As String = "http://127.0.0.1:25386"
Private mstrLastPwd As String          '�������������

Private Function GetCertInfo(Optional ByRef strCertUserName As String = "0", Optional ByRef strCertSn As String = "0", Optional ByRef strCertDN As String = "0") As Boolean
          Dim strJson As String
          Dim strJsonResult As String
          Dim arrItem As Variant
          
1         On Error GoTo errH

2         strJson = "{""function"":""GetCertInfo"",""keyType"":""2""}"
3         LogWrite "GetCertInfo", "URL��" & M_STR_URL & vbTab & "��Σ�" & strJson
4         strJsonResult = HttpPost(M_STR_URL, strJson, responseText, "application/x-www-form-urlencoded")
          ' {
          '   "dn" : "C=CN,S=�Ĵ�,L=�Ű�,OU=�Ű���У,CN=����֤��2",
          '   "sn" : "10000000006D299B",
          '   "time" : 1655458533
          '}
5         LogWrite "GetCertInfo", "����ֵ��" & strJsonResult
          If strJsonResult = "" Then Exit Function
          
6         If strJsonResult <> "" Then
7             If strCertSn <> "0" Then
8                 strCertSn = JSONParse("sn", strJsonResult)
9                 If strCertSn = "" Then
10                    MsgBoxEx "��ȡ֤��SNʧ�ܣ�", vbExclamation, gstrSysName
11                    Exit Function
12                End If
13            End If
14            If strCertSn <> "0" Then
15                strCertDN = JSONParse("dn", strJsonResult)
16                If strCertDN = "" Then
17                    MsgBoxEx "��ȡ֤��DNʧ�ܣ�", vbExclamation, gstrSysName
18                    Exit Function
19                End If
20            End If
21            If strCertUserName <> "0" Then
22                If strCertDN <> "" Then
23                    arrItem = Split(strCertDN, ",")
24                    If UBound(arrItem) >= 4 Then
25                        strCertUserName = Mid(arrItem(4), 4)
26                    Else
27                        MsgBoxEx "��ȡ֤���û���ʧ�ܣ�", vbExclamation, gstrSysName
28                        Exit Function
29                    End If
30                End If
31            End If
32        End If
33        GetCertInfo = True
34        Exit Function

errH:
35        MsgBox "��GetCertInfo�ĵ�" & Erl() & "�г���" & vbCrLf & _
            "�����: " & Err.Number & vbCrLf & _
            "����������" & Err.Description, vbExclamation, gstrSysName
End Function

Private Function GetCertLogin() As Boolean
      '��¼��֤
          '���������
          Dim strSource As String
          Dim strCert As String
          Dim strSign As String
          
1         On Error GoTo errH

2         Randomize
3         strSource = Int((100000 * Rnd) + 1)
          '�����ǩ��
4         If Not SignData(strSource, strCert, strSign) Then Exit Function
          '֤����֤
5         If Not VerifySignData(strCert, strSource, strSign) Then Exit Function
              
6         GetCertLogin = True

7         Exit Function

errH:
8         MsgBox "��GetCertLogin�ĵ�" & Erl() & "�г���" & vbCrLf & _
                  "�����: " & Err.Number & vbCrLf & _
                  "����������" & Err.Description, vbExclamation, gstrSysName
          
End Function

'---------------------------------------------------------------------------------------
' Procedure : JSONParse
' Author    : Administrator
' Date      : 2019-06-27
' Purpose   : ����JSON�ַ���
' Example   :
'    strJson= {"items":[{"dept_ids":"168,143,159,149,151,156,148,158,"}],"first":{"$ref":"http://192.168.0.231:8080/ords/zlrecipe/recipe/getenabledept"}}
'    JSONParse("items[0].dept_ids", strJson)
'    strJson={\"ͨ������\":\"������Ƭ\",\"��Ʒ��\":null} JSONParse("ͨ������", strJson)
'---------------------------------------------------------------------------------------
'
Public Function JSONParse(ByVal strJSONPath As String, ByVal strJSONData As String) As Variant
        Dim objJSON As Object
          
         

1       On Error GoTo errH

2       Set objJSON = CreateObject("MSScriptControl.ScriptControl")
3       objJSON.Language = "JScript"
4       JSONParse = Nvl(objJSON.Eval("JSON=" & strJSONData & ";JSON." & strJSONPath & ";"))
5       Set objJSON = Nothing

6       Exit Function

errH:
7         MsgBox "��JSONParse�ĵ�" & Erl() & "�г���" & vbCrLf & _
            "�����: " & Err.Number & vbCrLf & _
            "����������" & Err.Description, vbExclamation, gstrSysName
          
End Function

Private Function Login(ByVal strPWD As String) As Boolean
          Dim strJson As String
          Dim strJsonResult As String
          Dim strRet As String
          Dim strMsg As String
          
1        On Error GoTo errH
2         If strPWD = "" Then
3             If Not frmPassword.ShowMe(strPWD) Then Exit Function
4         End If
5         strJson = "{""function"":""Login"",""pin"":""" & strPWD & """}"
          
6         LogWrite "Login", "URL��" & M_STR_URL & vbTab & "��Σ�" & strJson
7         strJsonResult = HttpPost(M_STR_URL, strJson, responseText, "application/x-www-form-urlencoded", , strMsg)
8         LogWrite "Login", "����ֵ��" & strJsonResult
          '����ֵ������
          '{"nPinTryCount" : 0}
9         If strJsonResult <> "" Then
10            strRet = JSONParse("nPinTryCount", strJsonResult)
11            If strRet = "0" Then
12                mstrLastPwd = strPWD
13                Login = True
14            ElseIf strRet = "-1" Then
15                MsgBoxEx "���������֤�飡", vbInformation, gstrSysName
16                Exit Function
17            Else
18                MsgBoxEx "�������ʣ�����Դ�����" & strRet & "�Σ�", vbInformation, gstrSysName
19                Exit Function
20            End If
21        Else
22            MsgBoxEx "�������AisinoCertSrv���Ƿ��������С�" & vbCrLf & strMsg, vbExclamation, gstrSysName
23            Exit Function
24        End If

25       On Error GoTo 0
26       Exit Function

errH:
27        MsgBox "��Login�ĵ�" & Erl() & "�г���" & vbCrLf & _
                  "�����: " & Err.Number & vbCrLf & _
                  "����������" & Err.Description, vbExclamation, gstrSysName

End Function

Private Function SignData(ByVal strSource As String, Optional ByRef strCert As String = "0", Optional ByRef strSign As String = "0") As Boolean
      '---------------------------------------------------------------------------------------
      ' Procedure : SignData
      ' Author    : YWJ
      ' Date      : 2019-06-27 22:18:08
      ' Purpose   : Get Cert and sign
      ' PARA_IN   : strSource-ԭ��
      ' PARA_OUT  : strCert-֤������;strSign-ǩ��ֵ
      '{
      '   "cert" : "MIIDijCCAy6gAwIBAgIIEAAAAABtKZswDAYIKoEcz1UBg3UFADBQMQswCQYDVQQGEwJDTjEQMA4GA1UECAwHTmluZ3hpYTERMA8GA
      '      1UEBwwIWWluY2h1YW4xDTALBgNVBAoMBENXQ0ExDTALBgNVBAMMBE5YQ0EwHhcNMTkwNjIwMDkzNTMzWhcNMjIwNjE3MDkzNTMzWjBeMQswCQYDVQQGEwJDTjEPMA0GA1UECAwG5Zub5bedMQ8wDQYDVQQHDAbpm4XlrokxFTATBgNVBAsMDOmbheWuieWNq+agoTEWMBQGA1UEAwwN5rWL6K+V6K+B5LmmMjBZMBMGByqGSM49AgEGCCqBHM9VAYItA0IABMrGf0Po47zJdH72nWVheYUOgDjZl3UtrQ5ozdPt+9KhMfrHxfDzCg9n0ov3RPGdWUN7UnIeW4nNhze9OdUCtQqjggHgMIIB3DAdBgNVHQ4EFgQUpCusStBTkjDZpReXcaWCNQKfE/cwHwYDVR0jBBgwFoAUGlFGEC4+J/2NCQuh4mHdMKE/SK0wCwYDVR0PBAQDAgbAMB0GA1UdJQQWMBQGCCsGAQUFBwMCBggrBgEFBQcDBDCCAV4GA1UdHwSCAVUwggFRMIG7oGOgYYZfbGRhcDovLzIwMi4xMDAuMTA4LjEzOjM4OS9jbj1mdWxsQ3JsLmNybCxDTj1OWENBX0xEQVAsT1U9TlhDQSxPPUNXQ0EsTD1ZaW5jaHVhbixTVD1OaW5neGlhLEM9Q06iVKRSMFAxCzAJBgNVBAYTAkNOMRAwDgYDVQQIDAdOaW5neGlhMREwDwYDVQQHDAhZaW5jaHVhbjENMAsGA1UECgwEQ1dDQTENMAsGA1UEAwwETlhDQTCBkKA4oDaGNGh0dHA6Ly8yMDIuMTAwLjEwOC4xNTo4MTgxL254Y2EvMTAwMDAwMDAwMDZEMjkwMC5jcmyiVKRSMFAxCzAJBgNVBAYTAkNOMRAwDgYDVQQIDAd
      '      OaW5ne
      '      lhMREwDwYDVQQHDAhZaW5jaHVhbjENMAsGA1UECgwEQ1dDQTENMAsGA1UEAwwETlhDQTAMBgNVHRMEBTADAQEAMAwGCCqBHM9VAYN1BQADSAAwRQIhAL+IXTI4CWsdj2GLAyfMjHnzzUvJ4FkcqoDcrX7IQ/8qAiAdSCDFA5AlSnRLDx3mzrVfvX0xHIc3WnVg2YYqzKsUkw==",
      '   "sign" : "GIUWyqDGA1Q+YeyP6I4BkLX5fPSvTfMSovg9Qtul6D02MVAhvnXw1Uck3RRGdil+HKGjuyHtGnaIhlyejEvefQ=="
      '}
      '---------------------------------------------------------------------------------------
          Dim strJson As String
          Dim strJsonResult As String

1         On Error GoTo errH

2         strJson = "{""function"":""SignData"",""InData"":""" & strSource & """,""keyType"":""2""}"
3         LogWrite "SignData", "URL��" & M_STR_URL & vbTab & "��Σ�" & strJson
4         strJsonResult = HttpPost(M_STR_URL, strJson, responseText, "application/x-www-form-urlencoded")
5         LogWrite "SignData", "����ֵ��" & strJsonResult
6         If strJsonResult <> "" Then
7             If strCert <> "0" Then
8                 strCert = JSONParse("cert", strJsonResult)
9                 If strCert = "" Then
10                    MsgBoxEx "��ȡ֤������ʧ�ܣ�", vbExclamation, gstrSysName
11                    Exit Function
12                End If
13            End If
14            If strSign <> "0" Then
15                strSign = JSONParse("sign", strJsonResult)
16                If strSign = "" Then
17                    MsgBoxEx "��ȡǩ������ʧ�ܣ�", vbExclamation, gstrSysName
18                    Exit Function
19                End If
20            End If
21        Else
22            MsgBoxEx "ǩ��ֵ����Ϊ�գ�", vbExclamation, gstrSysName
23            Exit Function
24        End If
25        SignData = True
26        Exit Function

errH:
27        MsgBox "��SignData�ĵ�" & Erl() & "�г���" & vbCrLf & _
                  "�����: " & Err.Number & vbCrLf & _
                  "����������" & Err.Description, vbExclamation, gstrSysName
          
End Function

Private Function TSASign(ByVal strSource As String, ByRef strTimeStamp As String, ByRef strTimeStampCode As String) As Boolean
    Dim objNetOnex As Object
    Dim objTimeStamp As Object
    Dim objTSAResponseX As Object  'TSAResponseX
    Dim varTemp As Variant
    
    On Error Resume Next
    Set objNetOnex = CreateObject("NetONEX.MainX.1")
    If Err.Number <> 0 Then
        MsgBoxEx "����ʱ�������NetONEX.MainX.1��ʧ�ܣ����鲿����NetONEX.dll���Ƿ���ȷ��װ��ע�ᡣ", vbInformation, gstrSysName
        Exit Function
    End If
    On Error GoTo errH
    
    objNetOnex.DEBUG = 1
    '����ʱ�������
    Set objTimeStamp = objNetOnex.CreateTSAClientXInstance()
    If objTimeStamp Is Nothing Then
        MsgBoxEx "����ʱ�������ʧ�ܣ����÷�����CreateTSAClientXInstance����", vbInformation, gstrSysName
        Exit Function
    End If
    objTimeStamp.ServerAddress = gudtPara.strTSIP  '"sh.syan.com.cn"
    objTimeStamp.ServerPort = gudtPara.strTSPort   '"9110"
    Set objTSAResponseX = objTimeStamp.TSACreate(strSource)
    If objTSAResponseX Is Nothing Then
        MsgBoxEx "����ʱ���ǩ��ʧ�ܣ����÷�����TSACreate��", vbInformation, gstrSysName
        Exit Function
    End If
    strTimeStampCode = objTSAResponseX.ToBASE64()
        
    'objTSAResponseX.Timestamp Jul  5 09:48:48.188001 2019 GMT ʱ��ֻȡǰ��λ�ַ�
    strTimeStamp = objTSAResponseX.TimeStamp()
    strTimeStamp = Replace(strTimeStamp, Space(2), Space(1))  '����Ϊһλ��ʱǰ����ܴ��ڿո��½���ʧ��
    varTemp = Split(strTimeStamp, Space(1)) 'Jan 21 06:34:28.865495 2019 GMT ʱ��ֻȡǰ��λ�ַ�
    strTimeStamp = varTemp(3) & "-" & ConvMonth(varTemp(0)) & "-" & varTemp(1) & " " & Mid(varTemp(2), 1, 8) '��������ʱ��
    strTimeStamp = Format(DateAdd("h", 8, strTimeStamp), "YYYY-MM-DD HH:MM:SS")
    
    'ǩ���ɹ��ɹ�����200
    If Not (objTimeStamp.ServerResponseCode = 200 And strTimeStampCode <> "" And strTimeStamp <> "") Then
        MsgBoxEx "��ȡʱ���ʧ�ܣ�", vbInformation, gstrSysName: Exit Function
    End If
                
    TSASign = True

    Exit Function

errH:
    MsgBox "��TSASign�ĵ�" & Erl() & "�г���" & vbCrLf & _
      "�����: " & Err.Number & vbCrLf & _
      "����������" & Err.Description, vbExclamation, gstrSysName
End Function

Private Function TSAVerify(ByVal strTimeStampCode As String) As Boolean
          Dim objNetOnex As Object
          Dim objTimeStamp As Object
          Dim lngResult As Long
          
1         On Error Resume Next
2         Set objNetOnex = CreateObject("NetONEX.MainX.1")
3         If Err.Number <> 0 Then
4             MsgBoxEx "����ʱ�������NetONEX.MainX.1��ʧ�ܣ����鲿����NetONEX.dll���Ƿ���ȷ��װ��ע�ᡣ", vbInformation, gstrSysName
5             Exit Function
6         End If
7         On Error GoTo errH
          
8         objNetOnex.DEBUG = 1
          '����ʱ�������
9         On Error Resume Next
10        Set objTimeStamp = objNetOnex.CreateTSAClientXInstance()
11        On Error GoTo errH
12        If objTimeStamp Is Nothing Then
13            MsgBoxEx "����ʱ�������ʧ�ܣ����÷�����CreateTSAClientXInstance����", vbInformation, gstrSysName
14            Exit Function
15        End If
16        objTimeStamp.ServerAddress = gudtPara.strTSIP  '"sh.syan.com.cn"
17        objTimeStamp.ServerPort = gudtPara.strTSPort   '"9110"
18        lngResult = objTimeStamp.TSAVerify(strTimeStampCode)
          '��ǩ�ɹ�����200
19        If lngResult <> 200 Then
20            MsgBoxEx "ʱ�����֤ʧ�ܣ�", vbInformation, gstrSysName: Exit Function
21        End If
22        TSAVerify = True

23        Exit Function
errH:
24        MsgBox "��TSAVerify�ĵ�" & Erl() & "�г���" & vbCrLf & _
                  "�����: " & Err.Number & vbCrLf & _
                  "����������" & Err.Description, vbExclamation, gstrSysName
End Function

Private Function VerifySignData(ByVal strCert As String, ByVal strSource As String, ByVal strSignValue As String) As Boolean
      '����˵��
      'strCert:    Base64�����֤���ַ���
      'strSource:    Base64�����ԭ��
      'strSignValue:    Base64�����ǩ��ֵ
      '
      '����ֵ������0X0A000004:֤���ʽ����
      '              0X0A000010:ǩ����֤ʧ��
      '��ȷ:         Subject�е�CN��
          Dim strEnvelope As String
          Dim strResult As String
             
1         On Error GoTo errH
2         strSource = EncodeBase64String(strSource)
3         strEnvelope = "<soapenv:Envelope xmlns:soapenv=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:sig=""http://SignatureServer"">" & vbNewLine & _
                        "   <soapenv:Header/>" & vbNewLine & _
                        "   <soapenv:Body>" & vbNewLine & _
                        "      <sig:SOF_verifySignedData>" & vbNewLine & _
                        "         <sig:cert>" & strCert & "</sig:cert>" & vbNewLine & _
                        "         <sig:Indate>" & strSource & "</sig:Indate>" & vbNewLine & _
                        "         <sig:SignValue>" & strSignValue & "</sig:SignValue>" & vbNewLine & _
                        "      </sig:SOF_verifySignedData>" & vbNewLine & _
                        "   </soapenv:Body>" & vbNewLine & _
                        "</soapenv:Envelope>"

4         LogWrite "VerifySignData", "XML:" & strEnvelope
5         strResult = httpPostSOAP(gudtPara.strSignURL, strEnvelope, ".//SOF_verifySignedDataReturn", , "Content-Type[:]text/xml;charset=UTF-8[;]SOAPAction[:]application/soap+xml;charset=utf-8")
6         LogWrite "VerifySignData", "����ֵ:" & strResult
7         If strResult = "" Then
8             MsgBoxEx "ǩ����֤ʧ��(������:" & strResult & ")��", vbInformation, gstrSysName
9             Exit Function
10        End If
11        If InStr(strResult, "0X") > 0 Then
12            If "0X0A000004" = strResult Then
13                MsgBoxEx "ǩ����֤ʧ��:֤���ʽ����", vbInformation, gstrSysName
14                Exit Function
15            Else
16                MsgBoxEx "ǩ����֤ʧ��(������:" & strResult & ")��", vbInformation, gstrSysName
17                Exit Function
18            End If
19        End If
20        VerifySignData = True

21        Exit Function

errH:
22        MsgBox "��VerifySignData�ĵ�" & Erl() & "�г���" & vbCrLf & _
                  "�����: " & Err.Number & vbCrLf & _
                  "����������" & Err.Description, vbExclamation, gstrSysName
End Function

Public Function XiBuCA_CheckCert() As Boolean
      '--------------------------------------------------------------------------------------------------------------------------
      '���ܣ���ȡUSB�����豸��ʼ������¼
      '����:
      '����:
      '--------------------------------------------------------------------------------------------------------------------------
          Dim strCertSn As String
          
1         On Error GoTo errH
2         If mUserInfo.strCertSn <> gstrLogins Then
3            mstrLastPwd = "": gstrLogins = ""
4         End If
5         If Not Login(mstrLastPwd) Then Exit Function
6         If Not GetCertInfo(, strCertSn) Then Exit Function
7         If mUserInfo.strCertSn <> strCertSn Then
              '֤��Ψһ��ʶǰ׺���Ȳ��̶�,��ȡ18λ���֤��
8             MsgBoxEx "����֤�����кţ�" & _
                     vbCrLf & vbTab & "��" & mUserInfo.strCertSn & "��" & vbCrLf & _
                     "��ǰ֤�����к�" & _
                     vbCrLf & vbTab & "��" & strCertSn & "��" & vbCrLf & _
                     "�û���֤�����к��뵱ǰ֤�����кŲ����,����ʹ�ã�", vbInformation, gstrSysName
9             Exit Function
10        End If
          '��¼��֤
11        If gstrLogins <> strCertSn Then '�л�KEY����Ҫ���µ�¼��֤
12            If Not GetCertLogin() Then
13                gstrLogins = ""
14                Exit Function
15            Else
16                gstrLogins = strCertSn '�����һ��֤ͨ����KEY
17            End If
18        End If
19        XiBuCA_CheckCert = True
20     Exit Function

errH:
21        MsgBox "��XiBuCA_CheckCert�ĵ�" & Erl() & "�г���" & vbCrLf & _
            "�����: " & Err.Number & vbCrLf & _
            "����������" & Err.Description, vbExclamation, gstrSysName
        
End Function

Public Function XiBuCA_GetPara() As Boolean
          '���÷�������ַ
          
1         On Error GoTo errH
2         gudtPara.strSignURL = GetThirdPara(CON_PAR_����, "ǩ������WSDL")
          '�������Ե�ַ:sh.syan.com.cn Port:9110
3         gudtPara.strTSIP = GetThirdPara(CON_PAR_����, "ʱ���IP")
4         gudtPara.strTSPort = GetThirdPara(CON_PAR_����, "ʱ����˿�")
5         If gudtPara.strSignURL = "" Then gudtPara.strSignURL = "http://113.204.104.142:8082/SignatureServer/services/SignatureService?wsdl"
6         If gudtPara.strTSIP = "" Then gudtPara.strTSIP = "sh.syan.com.cn"
7         If gudtPara.strTSPort = "" Then gudtPara.strTSPort = "9110"
8         XiBuCA_GetPara = True
9         Exit Function
errH:
10        MsgBoxEx "��ȡ����ʧ�ܣ�" & vbNewLine & "��" & CStr(Erl()) & "�� " & Err.Description, vbInformation, gstrSysName
End Function

Public Function XiBuCA_InitObj() As Boolean
          Dim objNetOnex As Object
1         On Error GoTo errH
          
2         Call XiBuCA_GetPara
3         On Error Resume Next
4         Set objNetOnex = CreateObject("NetONEX.MainX.1")
5         If Err.Number <> 0 Then
6             LogWrite "XiBuCA_InitObj", "����ʱ�������ʧ�ܣ����鲿����NetONEX.dll���Ƿ���ȷ��װ��ע�ᡣ"
7             Exit Function
8         End If
9         On Error GoTo errH
          
10        XiBuCA_InitObj = True
           
11        Exit Function
errH:
12        GetErrMsg Erl()
End Function

Public Function XiBuCA_RegCert(arrCertInfo As Variant) As Boolean
      '���ܣ��ṩ��HIS���ݿ���ע�����֤��ı�Ҫ��Ϣ,�����·��Ż����֤��,,��Ҫ����USB-Key
      '����:strUserID-���֤��
      '���أ�arrCertInfo��Ϊ���鷵��֤�������Ϣ
      '      0-ClientSignCertCN:�ͻ���ǩ��֤�鹫������(����),����ע��֤��ʱ������֤���
      '      1-ClientSignCertDN:�ͻ���ǩ��֤������(ÿ��Ψһ)
      '      2-ClientSignCertSN:�ͻ���ǩ��֤�����к�(ÿ֤��Ψһ)
      '      3-ClientSignCert:�ͻ���ǩ��֤������
      '      4-ClientEncCert:�ͻ��˼���֤������
      '      5-ǩ��ͼƬ�ļ���,�մ���ʾû��ǩ��ͼƬ
              
          Dim strCertUserName As String, strCertDN As String
          Dim strCert As String, i As Integer
          Dim strCertSn As String
       
1         On Error GoTo errH

2         For i = LBound(arrCertInfo) To UBound(arrCertInfo)
3             arrCertInfo(i) = ""
4         Next
5         If Not Login("") Then Exit Function
6         If GetCertInfo(strCertUserName, strCertSn, strCertDN) Then
7             arrCertInfo(0) = strCertUserName
8             arrCertInfo(1) = strCertDN '֤��DN
9             arrCertInfo(2) = strCertSn '֤�����к� ǩ��ʱҪ��
10            If Not SignData("123", strCert) Then Exit Function
11            arrCertInfo(3) = strCert
12            arrCertInfo(4) = ""
13            arrCertInfo(5) = ""
14            XiBuCA_RegCert = True
15        End If
16        Exit Function

errH:
17        MsgBox "��XiBuCA_RegCert�ĵ�" & Erl() & "�г���" & vbCrLf & _
                  "�����: " & Err.Number & vbCrLf & _
                  "����������" & Err.Description, vbExclamation, gstrSysName
       
End Function

Public Function XiBuCA_Sign(ByVal strSource As String, ByRef strSignData As String, ByRef strTimeStamp As String, ByRef strTimeStampCode As String) As Boolean
      '---------------------------------------------------------------------------------------
      ' Procedure : JSCA_Sign
      ' Author    : YWJ
      ' Date      : 2019-06-27 23:00:26
      ' Parameter :
      ' Purpose   :
      '---------------------------------------------------------------------------------------
      '
      Dim strDigest As String
      
1         On Error GoTo errH
        
2         If XiBuCA_CheckCert() Then
3             strDigest = StringSHA1(strSource)
4             If Not SignData(strDigest, , strSignData) Then Exit Function
5             If Not TSASign(strDigest, strTimeStamp, strTimeStampCode) Then Exit Function
6         End If
7         XiBuCA_Sign = True

8         Exit Function

errH:
9         MsgBox "��XiBuCA_Sign�ĵ�" & Erl() & "�г���" & vbCrLf & _
                  "�����: " & Err.Number & vbCrLf & _
                  "����������" & Err.Description, vbExclamation, gstrSysName
End Function

Public Function XiBuCA_VerifySign(ByVal strCert As String, ByVal strSignData As String, ByVal strSource As String, ByVal strTimeStampCode As String) As Boolean
          '����:
          Dim strDigest As String
          
1         On Error GoTo errH
2         strDigest = StringSHA1(strSource)
3         If Not VerifySignData(strCert, strDigest, strSignData) Then Exit Function
4         If Not TSAVerify(strTimeStampCode) Then Exit Function
          
5         MsgBoxEx "��֤�ɹ����õ���ǩ��������Ч!", vbInformation, gstrSysName
6         XiBuCA_VerifySign = True
7         Exit Function
errH:
8         MsgBoxEx "��֤ǩ��ʧ�ܣ�" & vbNewLine & "��" & CStr(Erl()) & "�� " & Err.Description, vbExclamation, gstrSysName
End Function

Public Sub XiBuCA_UnloadObj()
     mstrLastPwd = ""
End Sub

