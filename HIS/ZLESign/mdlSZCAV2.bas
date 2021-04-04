Attribute VB_Name = "mdlSZCAV2"
Option Explicit
'���ڵ��ӵ���ǩ���ӿ�
'��������ҽԺ����2019/07/15
Private mSZCAClient As Object           '֤�鲿��
Private mobjTSA As Object               'ȷ��ʱ�������
Private mstrAlg As String               '�㷨:RSA;SM2
Private mblnInit As Boolean
Private mstrWSDL As String

Private Const M_STR_OID As String = "1.2.156.1002"

Public Function SZCAV2_InitObj() As Boolean
      '����:����ǩ�������ʼ��
          Dim strMsg As String
          
1         On Error GoTo errH
          
2         SZCAV2_InitObj = mblnInit
3         If mblnInit Then Exit Function
          
4         On Error Resume Next
5         Set mSZCAClient = CreateObject("SZCAPKI.SZCAPKICtrl.1")
6         If Err.Number > 0 Then
7             strMsg = "����ǩ������SZCAPKI.SZCAPKICtrl.1��ʧ�ܡ�"
8             GoTo errH
9         End If
10        On Error GoTo errH
11        If Not SZCA_GetPara Then
12            strMsg = "��ȡ����ʧ�ܣ���������Ƿ����á�"
13            GoTo errH
14        End If
                    
          '"http://202.103.144.98:7006/SZCAJavaCAS/services/szcaCAValidate.wsdl"
15        mstrWSDL = gudtPara.strSignURL
16        LogWrite "SZCAV2_InitObj", "�����ַ:" & mstrWSDL
17        If gudtPara.strTSIP <> "" Then
18            On Error Resume Next
19            Set mobjTSA = CreateObject("SuresecTsaClass.tsa.1")
20            On Error GoTo errH
21            If Err.Number > 0 Then
22                strMsg = "����ʱ�������SuresecTsaClass.tsa.1��ʧ�ܡ�"
23                Set mobjTSA = Nothing
24                GoTo errH
25            Else
26                mobjTSA.ISetTcpServerInfo gudtPara.strTSIP, gudtPara.strTSPort, 20
27            End If
28        End If
          
29        mblnInit = True
30        SZCAV2_InitObj = True
          
31        Exit Function
errH:
32        Call GetErrMsg(Erl(), strMsg)
End Function

Public Function SZCAV2_RegCert(arrCertInfo As Variant) As Boolean
        '���ܣ��ṩ��HIS���ݿ���ע�����֤��ı�Ҫ��Ϣ,�����·��Ż����֤��,,��Ҫ����USB-Key
        '���أ�arrCertInfo��Ϊ���鷵��֤�������Ϣ
        '      0-ClientSignCertCN:�ͻ���ǩ��֤�鹫������(����),����ע��֤��ʱ������֤���
        '      1-ClientSignCertDN:�ͻ���ǩ��֤������(ÿ��Ψһ)
        '      2-ClientSignCertSN:�ͻ���ǩ��֤�����к�(ÿ֤��Ψһ)
        '      3-ClientSignCert:�ͻ���ǩ��֤������
        '      4-ClientEncCert:�ͻ��˼���֤������
        '      5-ǩ��ͼƬ�ļ���,�մ���ʾû��ǩ��ͼƬ
        '      6-ʱ���֤��
        Dim strKeyId As String, strCertUserName As String, strCertDN As String
        Dim strSigCert As String, strTSCert As String
        Dim i As Long
        On Error GoTo errH
    
100     For i = LBound(arrCertInfo) To UBound(arrCertInfo)
101         arrCertInfo(i) = ""
102     Next
        
108     If SZCAV2_GetCertList(strCertUserName, strKeyId, strCertDN, strSigCert, strTSCert) Then
200         arrCertInfo(0) = strCertUserName
201         arrCertInfo(1) = strCertDN
202         arrCertInfo(2) = strKeyId
203         arrCertInfo(3) = strSigCert
205         arrCertInfo(4) = ""
206         arrCertInfo(5) = ""
            arrCertInfo(6) = strTSCert
            SZCAV2_RegCert = True
        End If
        
300     Exit Function

errH:
    MsgBoxEx "֤��ע��ʧ�ܣ�" & vbNewLine & "��" & CStr(Erl()) & "�� " & Err.Description, vbInformation, gstrSysName
End Function

Private Function SZCAV2_GetCertList(Optional ByRef strName As String, Optional ByRef strSN As String, _
    Optional ByRef strDn As String, Optional ByRef strCert As String = "-1", Optional ByRef strTSCert As String = "-1") As Boolean
          '����:��ȡ֤����Ϣ
          Dim arrList As Variant
          
          Dim strSource As String
          Dim lngRet As Long, lngTSALen As Long, lngCertLen As Long
          Dim arrTSData(2048) As Byte
          Dim arrCertData(2048) As Byte
          Dim bytSource() As Byte
          Dim blnRet As Boolean
          
1         On Error GoTo errH
          'KEY�㷨ȷ��
2         If Not GetKeyCertAlg() Then Exit Function
          
3         Call mSZCAClient.AxSetCertFilterStr("SC;szca;#;#;#;")
4         blnRet = mSZCAClient.AxSetKeyStore()
5         strCert = mSZCAClient.AxGetCertData()
6         LogWrite "SZCAV2_GetCertList", "���á�AxGetCertData�� ����ֵ(֤������):" & strCert
7         If strCert <> "" Then
8             strDn = mSZCAClient.AxGetB64CertInfo(strCert, 7) 'CN=����������ҽԺ���˲���8,OU=912837346471111,O=����������ҽԺ���˲���8,O=��Ϣ��,L=������,ST=�㶫ʡ,C=CN
9             LogWrite "SZCAV2_GetCertList", "���á�AxGetB64CertInfo������:����1=֤������,����2=7;" & vbTab & "����ֵ:" & strDn
10            arrList = Split(strDn, ",")
11            strName = Mid(arrList(0), 4)
12            strSN = mSZCAClient.AxGetCertInfoByOid(strCert, M_STR_OID)   '֤��Ψһ��ʶ���� ǰ׺��1@7025SF1��+BASE64��������֤��
13            LogWrite "SZCAV2_GetCertList", "���á�AxGetCertInfoByOid������:����1=֤������,����2=1.2.156.1002;" & vbTab & "����ֵ:" & strSN
14        Else
15            MsgBoxEx "�����֤��Key��", vbOKOnly + vbInformation, gstrSysName
16            Exit Function
17        End If
          'ʱ�����Ϣ----------------------------------------------------------
18        If strTSCert <> "-1" And Not mobjTSA Is Nothing Then
19            strSource = "����ABCabc123"
20            ReDim bytSource(LenB(StrConv(strSource, vbFromUnicode)))
21            bytSource = StrConv(strSource, vbFromUnicode)
22            lngRet = mobjTSA.IGenTokenByPlain("SHA1", bytSource(0), UBound(bytSource) + 1, 1, arrTSData(0), lngTSALen)
23            LogWrite "SZCAV2_GetCertList", "���á�IGenTokenByPlain������ֵ:" & lngRet
24            If lngRet = 0 Then
25                lngRet = mobjTSA.IGetTokenCertificate(arrTSData(0), lngTSALen, arrCertData(0), lngCertLen)
26                strTSCert = FuncEncodeBase64Byte(arrCertData, lngCertLen)
27            Else
28                MsgBoxEx "ʱ���֤���ȡʧ�ܣ�", vbOKOnly + vbInformation, gstrSysName
29                Exit Function
30            End If
31        End If
          
32        SZCAV2_GetCertList = True

33        Exit Function
errH:
34        MsgBox "��SZCAV2_GetCertList�ĵ�" & Erl() & "�г���" & vbCrLf & _
                  "�����: " & Err.Number & vbCrLf & _
                  "����������" & Err.Description, vbExclamation, gstrSysName
       
End Function

Private Function GetCertLogin(ByVal strCert As String) As Boolean
      '����:֤���¼��֤
          Dim strSource As String, strBASE64 As String
          Dim strSign As String, strMsg As String
          Dim strRet As String
          
1         On Error GoTo errH

2         Randomize
3         strSource = Int((100000 * Rnd) + 1)
4         strSign = mSZCAClient.AxSignMessage(strSource, False)
5         If strSign <> "" Then
6             strBASE64 = VerifySign(strSign, mstrAlg)
7             If strBASE64 = "1" Then 'MQ== ����1
              '1 ֤����Ч -1 ֤����Ч�����������εĸ� -2 ֤����Ч��������Ч�� -3 ֤����Ч���Ѽ��������
8                 strRet = CertificateValidate(strCert)
9                 Select Case strRet
                  Case "-1"
10                    strMsg = "֤����Ч�����������εĸ�"
11                Case "-2"
12                    strMsg = "֤����Ч��������Ч��"
13                Case "-3"
14                    strMsg = "֤����Ч���Ѽ��������"
15                Case ""
16                    strMsg = "֤����֤ʧ�ܣ�"
17                End Select
18            Else
19                strMsg = "��¼��֤ʧ�ܣ���֤��Ϣ���¼��Ϣ����"
20            End If
21        Else
22            strMsg = "��¼ʧ�ܣ�"
23        End If
24        If strMsg <> "" Then
25            MsgBoxEx strMsg, vbOKOnly + vbInformation, gstrSysName
26            Exit Function
27        End If
          
28        GetCertLogin = True

29        Exit Function

errH:
30        MsgBox "��GetCertLogin�ĵ�" & Erl() & "�г���" & vbCrLf & _
                  "�����: " & Err.Number & vbCrLf & _
                  "����������" & Err.Description, vbExclamation, gstrSysName
End Function

Public Function SZCAV2_CheckCert(ByVal strCurrCertSn As String) As Boolean
      '���ܣ���ȡUSB�����豸��ʼ������¼
      '����ֵ:
      '  strSigCert -ǩ��֤������

           Dim strSN As String, strSigCert As String
          
1         On Error GoTo errH

2          If Not SZCAV2_InitObj() Then
3              MsgBoxEx "����δ��ʼ����"
4              Exit Function
5          End If
           
6          If Not SZCAV2_GetCertList(, strSN, , strSigCert) Then Exit Function
7          If strCurrCertSn <> strSN Then
8              MsgBoxEx "��֤��δע�����������£�����ʹ�ã�", vbInformation + vbOKOnly, gstrSysName
9              Exit Function
10         End If
              
           '��¼��֤
11         If Not InStr(gstrLogins & "|", "|" & strCurrCertSn & "|") > 0 Then '�״���֤ͨ�����´β��ڼ�����֤
12             If Not GetCertLogin(strSigCert) Then
13                 Exit Function
14             Else
15                 If InStr(gstrLogins & "|", "|" & strCurrCertSn & "|") = 0 Then gstrLogins = gstrLogins & "|" & strCurrCertSn
16             End If
17         End If
          
18         SZCAV2_CheckCert = True

19        Exit Function

errH:
20        MsgBox "��SZCAV2_CheckCert�ĵ�" & Erl() & "�г���" & vbCrLf & _
                  "�����: " & Err.Number & vbCrLf & _
                  "����������" & Err.Description, vbExclamation, gstrSysName

End Function

Public Function SZCAV2_Sign(ByVal strSN As String, ByVal strSource As String, ByRef strSignData As String, ByRef strTimeStamp As String, ByRef strTimeStampCode As String)
      '����:ǩ��
          Dim strMsg As String
          Dim strDigest As String
          Dim strTime As String
          Dim strRet As String

          Dim bytSource() As Byte
          Dim arrTSData(2048) As Byte
          Dim lngTSALen As Long
          Dim lngRet As Long
          
1         On Error GoTo errH
          
2         If SZCAV2_CheckCert(strSN) Then
3             strDigest = StringSHA1(strSource)
4             strSignData = mSZCAClient.AxSignMessage(strDigest, False)
                
5             If strSignData = "" Then
6                 MsgBoxEx "ǩ��ʧ�ܣ�����ǩ��ֵΪ�ա�", vbInformation + vbOKOnly, gstrSysName
7                 Exit Function
8             Else
9                 strTimeStamp = Format(gobjComLib.zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
10                strRet = VerifySign(strSignData, mstrAlg)
11                If strRet <> "1" Then
12                    MsgBoxEx "ǩ��ʧ�ܣ���֤ǩ������ֵ:" & strRet, vbInformation + vbOKOnly, gstrSysName
13                    Exit Function
14                End If
15            End If
16            If Not mobjTSA Is Nothing Then
                  '��ȡʱ���
17                ReDim bytSource(LenB(StrConv(strSource, vbFromUnicode)))
18                bytSource = StrConv(strSource, vbFromUnicode)
                  'ͨ�����Ļ�ȡʱ���
19                Call mobjTSA.IGenTokenByPlain("SHA1", bytSource(0), UBound(bytSource) + 1, 0, arrTSData(0), lngTSALen)
                  '��ȡʱ�����ʱ��
20                lngRet = mobjTSA.IGetTokenGenerateTime(arrTSData(0), lngTSALen, strTime)
                  
21                If lngRet = 0 Then
22                    strTimeStamp = String14ToDate(strTime, strMsg)
23                    strTimeStampCode = FuncEncodeBase64Byte(arrTSData, lngTSALen)
24                Else
25                    MsgBoxEx "ʱ���ʧ�ܣ���ȡʱ�����ʱ��ʧ�ܡ�����ֵ:" & lngRet, vbInformation + vbOKOnly, gstrSysName
26                    Exit Function
27                End If
28            End If
29        Else
30            MsgBoxEx "ǩ��ʧ�ܣ�", vbInformation + vbOKOnly, gstrSysName
31            Exit Function
32        End If

33        SZCAV2_Sign = True
          
34        Exit Function

errH:
35        MsgBox "��SZCAV2_Sign�ĵ�" & Erl() & "�г���" & vbCrLf & _
                  "�����: " & Err.Number & vbCrLf & _
                  "����������" & Err.Description, vbExclamation, gstrSysName

End Function

Public Function SZCAV2_VerifySign(ByVal strSource As String, ByVal strSign As String, ByVal strTSCert As String, ByVal strTStampCode As String) As Boolean
      '����:��֤ǩ��
          Dim strRet As String
          Dim strMsg As String
          Dim blnRet As Boolean
          Dim lngRet As Long
          Dim bytSource() As Byte, bytTStamp() As Byte, bytTSCert() As Byte
          
1         On Error GoTo errH

          '֧�ֲ���KEY��֤ǩ��
2         If gudtPara.bytSignVersion = 1 Then
3             strRet = VerifySign(strSign, "SM2")
4             If strRet <> "1" Then
5                 strRet = VerifySign(strSign, "RSA")
6             End If
7         Else
8             strRet = VerifySign(strSign, "RSA")
9         End If
10        blnRet = False
11        If strRet = "1" Then
12            strMsg = "��֤�ɹ���ǩ��������֤ͨ����"
13            blnRet = True
14        Else
15            strMsg = "��֤ʧ�ܣ�ǩ��������֤ʧ�ܣ�"
16        End If
          
17        If blnRet And Not mobjTSA Is Nothing Then
18            ReDim bytSource(LenB(StrConv(strSource, vbFromUnicode)))
19            bytSource = StrConv(strSource, vbFromUnicode)
20            bytTSCert = DecodeBase64Byte(strTSCert)
21            bytTStamp = DecodeBase64Byte(strTStampCode)
          
22            lngRet = mobjTSA.IVerifyTimeStampTokenByPlain("SHA1", bytSource(0), UBound(bytSource) + 1, bytTStamp(0), UBound(bytTStamp) + 1, bytTSCert(0), UBound(bytTSCert) + 1)
23            If lngRet = 0 Then
24                blnRet = True
25                strMsg = strMsg & vbCrLf & "ʱ�����Ϣ��֤ͨ����"
26            Else
27                strMsg = strMsg & vbCrLf & "ʱ�����Ϣ��֤ʧ�ܣ�"
28                blnRet = False
29            End If
30        End If
          
31        If strMsg <> "" Then
32            MsgBoxEx strMsg, vbInformation + vbOKOnly, gstrSysName
33        End If
34        SZCAV2_VerifySign = blnRet

35        Exit Function

errH:
36        MsgBox "��SZCAV2_VerifySign�ĵ�" & Erl() & "�г���" & vbCrLf & _
                  "�����: " & Err.Number & vbCrLf & _
                  "����������" & Err.Description, vbExclamation, gstrSysName

End Function

Public Sub SZCAV2_UnloadObj()
    Set mSZCAClient = Nothing
    Set mobjTSA = Nothing
    mblnInit = False
End Sub

Private Function FuncEncodeBase64Byte(bytArr() As Byte, ByVal lngLength As Long) As String
'����:��һ���ֽ��������Base64���룬�������ַ���
    Dim strRet As String
    Dim i As Long
    Dim bytBuffer() As Byte
    
    ReDim bytBuffer(lngLength)
    
    For i = 0 To lngLength - 1
        bytBuffer(i) = bytArr(i)
    Next
    strRet = EncodeBase64Byte(bytBuffer)
    FuncEncodeBase64Byte = strRet
End Function

Private Function GetKeyCertAlg() As Boolean
          Dim lngRet As Long
          
1         On Error GoTo errH

2         lngRet = mSZCAClient.AxGetKeyCertAlg             '��ʼ���ӿ�
3         LogWrite "GetKeyCertAlg", "���á�AxGetKeyCertAlg������ֵ:" & lngRet
4         If lngRet = 1 Then
              'RSA
5             mstrAlg = "RSA"
6             lngRet = mSZCAClient.AxSetAlgorithm(0)
7             LogWrite "GetKeyCertAlg", "���á�AxSetAlgorithm������:0;����ֵ:" & lngRet
8         ElseIf lngRet = 2 Then
              'SM2
9             mstrAlg = "SM2"
10            lngRet = mSZCAClient.AxSetAlgorithm(1)
11            LogWrite "GetKeyCertAlg", "���á�AxSetAlgorithm������:1;����ֵ:" & lngRet
12        ElseIf lngRet = 3 Then
13            mstrAlg = "SM2"
              '˫�㷨
14            lngRet = mSZCAClient.AxSetAlgorithm(1)
15            LogWrite "GetKeyCertAlg", "���á�AxSetAlgorithm������:1;����ֵ:" & lngRet
16        Else
17            mstrAlg = ""
              'û�з���֤��
18            MsgBoxEx "û�з���֤�飬�����֤��Key��", vbInformation, gstrSysName
19            Exit Function
20        End If
21        GetKeyCertAlg = True

22        Exit Function

errH:
23        MsgBox "��GetKeyCertAlg�ĵ�" & Erl() & "�г���" & vbCrLf & _
                  "�����: " & Err.Number & vbCrLf & _
                  "����������" & Err.Description, vbExclamation, gstrSysName
          
End Function

Private Function VerifySign(ByVal strSign As String, ByVal strAlg As String) As String
'����:��֤ǩ��
'strSign-ǩ��ֵ(��ԭ�ļ�֤������)
'strAlg= SM2,RSA
          Dim strBASE64 As String
          Dim strEnvelope As String
          
1         On Error GoTo errH

2         If strAlg = "SM2" Then
                '���ýӿ�: szcaWSSignatureValidatePkcs7SM2(strSign)
3             strEnvelope = "<soapenv:Envelope xmlns:soapenv=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:ser=""http://service.webservice.caau.szca.com/"">" & vbNewLine & _
                          "   <soapenv:Header/>" & vbNewLine & _
                          "   <soapenv:Body>" & vbNewLine & _
                          "      <ser:szcaWSSignatureValidatePkcs7SM2>" & vbNewLine & _
                          "      <signdata>" & strSign & "</signdata></ser:szcaWSSignatureValidatePkcs7SM2>" & vbNewLine & _
                          "   </soapenv:Body>" & vbNewLine & _
                          "</soapenv:Envelope>"
4         Else
5             strEnvelope = "<soapenv:Envelope xmlns:soapenv=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:ser=""http://service.webservice.caau.szca.com/"">" & vbNewLine & _
                          "   <soapenv:Header/>" & vbNewLine & _
                          "   <soapenv:Body>" & vbNewLine & _
                          "      <ser:szcaWSSignatureValidatePkcs7String>" & vbNewLine & _
                          "      <signdata>" & strSign & "</signdata></ser:szcaWSSignatureValidatePkcs7String>" & vbNewLine & _
                          "   </soapenv:Body>" & vbNewLine & _
                          "</soapenv:Envelope>"
6         End If
7         LogWrite "VerifySign", "���á�VerifySign������ֵ:" & strEnvelope
8         strBASE64 = httpPostSOAP(mstrWSDL, strEnvelope, ".//return")
9         LogWrite "VerifySign", "���á�VerifySign������ֵ:" & strBASE64
10        VerifySign = DecodeBase64String(strBASE64)

11        Exit Function

errH:
12        MsgBox "��VerifySign�ĵ�" & Erl() & "�г���" & vbCrLf & _
                  "�����: " & Err.Number & vbCrLf & _
                  "����������" & Err.Description, vbExclamation, gstrSysName
End Function

Private Function CertificateValidate(ByVal strCert As String) As String
          Dim strEnvelope As String
          Dim strBASE64 As String
          
1         On Error GoTo errH

2         If mstrAlg = "SM2" Then
3             strEnvelope = "<soapenv:Envelope xmlns:soapenv=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:ser=""http://service.webservice.caau.szca.com/"">" & vbNewLine & _
                        "   <soapenv:Header/>" & vbNewLine & _
                        "   <soapenv:Body>" & vbNewLine & _
                        "      <ser:szcaWSCertValidateSM2>" & vbNewLine & _
                        "      <certBase64>" & strCert & "</certBase64></ser:szcaWSCertValidateSM2>" & vbNewLine & _
                        "   </soapenv:Body>" & vbNewLine & _
                        "</soapenv:Envelope>"
4         Else
5             strEnvelope = "<soapenv:Envelope xmlns:soapenv=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:ser=""http://service.webservice.caau.szca.com/"">" & vbNewLine & _
                          "   <soapenv:Header/>" & vbNewLine & _
                          "   <soapenv:Body>" & vbNewLine & _
                          "      <ser:szcaWSCertificateValidateString>" & vbNewLine & _
                          "      <certBase64>" & strCert & "</certBase64></ser:szcaWSCertificateValidateString>" & vbNewLine & _
                          "   </soapenv:Body>" & vbNewLine & _
                          "</soapenv:Envelope>"
6         End If
7         LogWrite "CertificateValidate", "���á�CertificateValidate������ֵ:" & strEnvelope
8         strBASE64 = httpPostSOAP(mstrWSDL, strEnvelope, ".//return")
9         LogWrite "CertificateValidate", "���á�CertificateValidate������ֵ:" & strBASE64
10        CertificateValidate = DecodeBase64String(strBASE64)

11        Exit Function

errH:
12        MsgBox "��CertificateValidate�ĵ�" & Erl() & "�г���" & vbCrLf & _
                  "�����: " & Err.Number & vbCrLf & _
                  "����������" & Err.Description, vbExclamation, gstrSysName
          
End Function
