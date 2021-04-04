Attribute VB_Name = "mdlShanXiCA"
Option Explicit

'Private mobjCertMgr As HebcaP11XLibCtl.CertMgr
Private mobjCertMgr As Object          'HebcaP11XLib.certMgr  ֤�����

Private Const M_STR_LICENCE As String = "amViY55oZWKcZmhlnWxhaGViY2GXGmJjYWhlYnGH1QQ5GcNqnW6z3vohVnE+nTJr"
Private Const M_STR_SPLIT As String = "<SPLIT>"
Private mstrWSDL As String

Private Function CheckP7(ByVal strSignData As String) As Boolean
          Dim strEnvelope As String
          Dim strResult As String
          
1         On Error GoTo ErrH

2         strEnvelope = "<soap:Envelope xmlns:soap=""http://www.w3.org/2003/05/soap-envelope"" xmlns:snca=""http://snca.CertificateAuthorityServices/"">" & vbNewLine & _
                  "   <soap:Header/>" & vbNewLine & _
                  "   <soap:Body>" & vbNewLine & _
                  "      <snca:checkSNCAPKCS7Certificate>" & vbNewLine & _
                  "         <snca:PKCS7Info>" & strSignData & "</snca:PKCS7Info>" & vbNewLine & _
                  "      </snca:checkSNCAPKCS7Certificate>" & vbNewLine & _
                  "   </soap:Body>" & vbNewLine & _
                  "</soap:Envelope>"
3         LogWrite "CheckP7", "MXL:" & strEnvelope
4         strResult = httpPostSOAP(mstrWSDL, strEnvelope, ".//ns:return", "application/soap+xml;charset=UTF-8")
5         CheckP7 = IIf(strResult = "true", True, False)

6         Exit Function

ErrH:
7         MsgBox "��CheckP7�ĵ�" & Erl() & "�г���" & vbCrLf & _
                  "�����: " & Err.Number & vbCrLf & _
                  "����������" & Err.Description, vbExclamation, gstrSysName
End Function

'Private Function GetTrustNumber(ByVal strSignData As String) As String
'    Dim strEnvelope As String
'    Dim strResult As String
'
'    strEnvelope = "<soap:Envelope xmlns:soap=""http://www.w3.org/2003/05/soap-envelope"" xmlns:snca=""http://snca.CertificateAuthorityServices/"">" & vbNewLine & _
'                "   <soap:Header/>" & vbNewLine & _
'                "   <soap:Body>" & vbNewLine & _
'                "      <snca:getSNCATrustNumber>" & vbNewLine & _
'                "         <snca:in>" & strSignData & "</snca:in>" & vbNewLine & _
'                "         <snca:type>PKCS7</snca:type>" & vbNewLine & _
'                "         <snca:expendingItemKey></snca:expendingItemKey>" & vbNewLine & _
'                "      </snca:getSNCATrustNumber>" & vbNewLine & _
'                "   </soap:Body>" & vbNewLine & _
'                "</soap:Envelope>"
'    strResult = httpPostSOAP(mstrWSDL, strEnvelope, ".//ns:return", "application/soap+xml;charset=UTF-8")
'    GetTrustNumber = strResult
'End Function

Private Function GetCertList(ByRef strName As String, ByRef strCertSn As String, Optional ByRef strCertDN As String, Optional ByRef objCert As Object) As Boolean
          '----------------------------------------------------------------------------------------------------------------------------------
          '����:��ȡ֤����Ϣ
          '����:strName-֤���û���
          '     strCertSn-�ͷ����κ�
          '     strCertDN-DN
          '----------------------------------------------------------------------------------------------------------------------------------
          Dim intCount As Integer
                          
1         On Error GoTo ErrH
2         intCount = mobjCertMgr.GetDeviceCount
3         If intCount < 1 Then
4             MsgBoxEx "δ����KEY,��������Key��", vbInformation, gstrSysName
5             Exit Function
6         ElseIf intCount > 1 Then
7             MsgBoxEx "���ĵ����ϲ����˶������CA����֤�飬�뽫�����֤���Ƴ�!", vbInformation, gstrSysName
8             Exit Function
9         End If

      '    intCount = mobjCertMgr.GetSignCertCount
      '    If intCount > 1 Then
      '        MsgBoxEx "���ĵ����ϲ����˶������CA����֤�飬�뽫�����֤���Ƴ�!", vbInformation, gstrSysName
      '        Exit Function
      '    End If
      '     CN = ����������
      '    For i = 0 To intCount
      '        Set objCert = mobjCertMgr.GetCert(i)
      '        If Not objCert Is Nothing Then
      '            strDn = objCert.GetSubject()
      '        End If
      '    Next
10        Set objCert = mobjCertMgr.SelectSignCert
11        If objCert Is Nothing Then
12            MsgBoxEx "��ȡ֤��ʧ�ܣ�", vbInformation, gstrSysName
13            Exit Function
14        End If
          
15        strName = objCert.GetSubjectItem("cn")
16        If strName = "" Then
17            MsgBoxEx "��ȡ֤��CNʧ�ܣ�", vbInformation, gstrSysName
18            Exit Function
19        End If
20        strCertDN = objCert.GetSubject()
21        strCertSn = objCert.GetCertExtensionByOid("1.2.86.11.7.11")
22        If strCertSn = "" Then
23            MsgBoxEx "��ȡ�ͷ����κ�ʧ�ܣ�", vbInformation, gstrSysName
24            Exit Function
25        End If
26        LogWrite "GetCertList", "CN:" & strName & vbCrLf & "SN:" & strCertSn & vbCrLf & "DN:" & strCertDN
27        GetCertList = True
28        Exit Function

ErrH:
29        MsgBox "��GetCertList�ĵ�" & Erl() & "�г���" & vbCrLf & _
                  "�����: " & Err.Number & vbCrLf & _
                  "����������" & Err.Description, vbExclamation, gstrSysName
       
End Function

 
Private Function GetCertLogin(ByVal objCert As Object) As Boolean
      '����:֤���¼��֤
          Dim strServiceTime As String
          Dim objDevice As Object
          Dim strRan As String
          
          Dim strSignData As String
          Dim strSource As String
          Dim blnResutl As Boolean
          
          '��ȡ������ʱ��
1         On Error GoTo ErrH

2         strServiceTime = GetCurrentTime()
          '���������
3         Set objDevice = mobjCertMgr.GetDevice(0)
4         strRan = objDevice.GenRandom(128)
5         strSource = "TIME" & strServiceTime & "TIME" & strRan
          '��֤֤��
6         strSignData = SignedData(strSource, objCert)
7         If strSignData = "" Then Exit Function
8         blnResutl = CheckP7(strSignData)
9         If Not blnResutl Then
10            MsgBoxEx "��������֤ʧ�ܣ�", vbInformation, gstrSysName
11            Exit Function
12        End If
13        GetCertLogin = True

14        Exit Function

ErrH:
15        MsgBox "��GetCertLogin�ĵ�" & Erl() & "�г���" & vbCrLf & _
                  "�����: " & Err.Number & vbCrLf & _
                  "����������" & Err.Description, vbExclamation, gstrSysName
End Function

'End Function
'
Private Function GetCurrentTime() As String
          Dim strEnvelope As String
          
1         On Error GoTo ErrH

2         strEnvelope = "<soap:Envelope xmlns:soap=""http://www.w3.org/2003/05/soap-envelope"" xmlns:snca=""http://snca.CertificateAuthorityServices/"">" & vbNewLine & _
                  "   <soap:Header/>" & vbNewLine & _
                  "   <soap:Body>" & vbNewLine & _
                  "      <snca:getCurrentTime>" & vbNewLine & _
                  "         <snca:time xsi:nil=""true"" xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance""/>" & vbNewLine & _
                  "      </snca:getCurrentTime>" & vbNewLine & _
                  "   </soap:Body>" & vbNewLine & _
                  "</soap:Envelope>"
          '2019-07-31-15-02-05
3         LogWrite "GetCurrentTime", "MXL:" & strEnvelope
4         GetCurrentTime = httpPostSOAP(mstrWSDL, strEnvelope, ".//ns:return", "application/soap+xml;charset=UTF-8")

5         Exit Function

ErrH:
6         MsgBox "��GetCurrentTime�ĵ�" & Erl() & "�г���" & vbCrLf & _
                  "�����: " & Err.Number & vbCrLf & _
                  "����������" & Err.Description, vbExclamation, gstrSysName
End Function

Private Function SaveToDB(ByVal strSignData As String, ByVal strServerID As String, ByVal strAppID As String, ByVal strExtendId As String) As Boolean
      '����:�ϴ�ǩ��ֵ
      '    strSignData=
      '    strServerID = ϵͳID
      '    strAppID = ��GUID����(��Ϊ�޷���ȡҵ��ϵͳID�Լ�ǩ����¼ID,�ʽ���ֵ����ΪGUID,�������ȡ֤��ʱ��ͨ����ֵ��ȡ�Է�������ǩ��ֵ)
      '    strExtendId =Ĭ��Ϊ"00"
          Dim strEnvelope As String
          Dim strResult As String
          
1         On Error GoTo ErrH

2         strEnvelope = "<soap:Envelope xmlns:soap=""http://www.w3.org/2003/05/soap-envelope"" xmlns:snca=""http://snca.CertificateAuthorityServices/"">" & vbNewLine & _
              "   <soap:Header/>" & vbNewLine & _
              "   <soap:Body>" & vbNewLine & _
              "      <snca:checkSNCAPKCS7SignAndSaveToDB>" & vbNewLine & _
              "         <snca:PKCS7Info>" & strSignData & "</snca:PKCS7Info>" & vbNewLine & _
              "         <snca:Service_id>" & strServerID & "</snca:Service_id>" & vbNewLine & _
              "         <snca:app_id>" & strAppID & "</snca:app_id>" & vbNewLine & _
              "         <snca:extend_id>" & strExtendId & "</snca:extend_id>" & vbNewLine & _
              "      </snca:checkSNCAPKCS7SignAndSaveToDB>" & vbNewLine & _
              "   </soap:Body>" & vbNewLine & _
              "</soap:Envelope>"
          LogWrite "SaveToDB", "MXL:" & strEnvelope
3         strResult = httpPostSOAP(mstrWSDL, strEnvelope, ".//ns:return", "application/soap+xml;charset=UTF-8")
4         SaveToDB = IIf(strResult = "true", True, False)

5         Exit Function

ErrH:
6         MsgBox "��SaveToDB�ĵ�" & Erl() & "�г���" & vbCrLf & _
                  "�����: " & Err.Number & vbCrLf & _
                  "����������" & Err.Description, vbExclamation, gstrSysName
End Function

Public Function ShanXi_CheckCert(Optional ByRef objCert As Object) As Boolean
          '���ܣ���ȡUSB�����豸��ʼ������¼
          '����ֵ:
          '  strSigCert -ǩ��֤������

          Dim strSN As String
          Dim strName As String
          
1         On Error GoTo ErrH

2         If Not GetCertList(strName, strSN, , objCert) Then Exit Function
3         If mUserInfo.strCertSn <> strSN Then
4             MsgBoxEx "��֤��δע�����������£�����ʹ�ã�", vbInformation + vbOKOnly, gstrSysName
5             gstrLogins = ""
6             Exit Function
7         End If

          '��¼��֤
8         If gstrLogins <> strSN Then '�л�KEY����Ҫ���µ�¼��֤
9             If Not GetCertLogin(objCert) Then
10                gstrLogins = ""
11                Exit Function
12            Else
13                gstrLogins = strSN '�����һ��֤ͨ����KEY
14            End If
15        End If
          
16        ShanXi_CheckCert = True
17        Exit Function

ErrH:
18        MsgBox "��ShanXi_CheckCert�ĵ�" & Erl() & "�г���" & vbCrLf & _
                  "�����: " & Err.Number & vbCrLf & _
                  "����������" & Err.Description, vbExclamation, gstrSysName

End Function

 
Public Function ShanXi_GetPara() As Boolean
      '��������CA��������ַ
          
1         On Error GoTo ErrH

2         On Error GoTo ErrH

3         With gudtPara
4             .strSignURL = GetThirdPara(CON_PAR_����, "ǩ������WSDL")
5             .strOption = GetThirdPara(CON_PAR_����, "ϵͳ��ʶ") 'ҽԺȫ��
6             .strTSIP = GetThirdPara(CON_PAR_����, "ʱ�������WSDL")  'ʱ�������WSDL
7             If .strSignURL = "" Or .strOption = "" Or .strTSIP = "" Then
8                 Exit Function
9             End If
10        End With

11        ShanXi_GetPara = True
12        Exit Function

ErrH:
13        MsgBox "��ShanXi_GetPara�ĵ�" & Erl() & "�г���" & vbCrLf & _
                  "�����: " & Err.Number & vbCrLf & _
                  "����������" & Err.Description, vbExclamation, gstrSysName
End Function

Public Function ShanXi_InitObj() As Boolean
          
          Dim strMsg As String
          
1         On Error Resume Next
2         Set mobjCertMgr = CreateObject("HebcaP11X.CertMgr.1")
3         If Err.Number <> 0 Then
              'C:\Windows\System32\HebcaP11X.dll
4             strMsg = "����֤��������ʧ�ܣ����鲿����HebcaP11X.dll���Ƿ���ȷ��װ��ע�ᡣ"
5             GoTo ErrH
6         End If
12        On Error GoTo ErrH
13        mobjCertMgr.Licence = M_STR_LICENCE
14        If Not ShanXi_GetPara() Then
15            strMsg = "û�����õ���ǩ�����������ȵ��������������á����á�"
16            GoTo ErrH
17        End If
      '    gudtPara.strSignURL = "http://111.20.164.185:8771/SNCA_CertificateAuthorityPlatform/services/CertificateAuthorityServices?wsdl"
18        mstrWSDL = gudtPara.strSignURL
19        ShanXi_InitObj = True
20        Exit Function
ErrH:
21       Call GetErrMsg(Erl(), strMsg)
End Function

Public Function ShanXi_RegCert(arrCertInfo As Variant) As Boolean
      '���ܣ��ṩ��HIS���ݿ���ע�����֤��ı�Ҫ��Ϣ,�����·��Ż����֤��,,��Ҫ����USB-Key
      '���أ�arrCertInfo��Ϊ���鷵��֤�������Ϣ
      '      0-ClientSignCertCN:�ͻ���ǩ��֤�鹫������(����),����ע��֤��ʱ������֤���
      '      1-ClientSignCertDN:�ͻ���ǩ��֤������(ÿ��Ψһ)
      '      2-ClientSignCertSN:�ͻ���ǩ��֤�����к�(ÿ֤��Ψһ)
      '      3-ClientSignCert:�ͻ���ǩ��֤������
      '      4-ClientEncCert:�ͻ��˼���֤������
      '      5-ǩ��ͼƬ�ļ���,�մ���ʾû��ǩ��ͼƬ
      '      6-ʱ���֤��
          Dim strCertSn As String, strCertUserName As String, strCertDN As String
          Dim strSigCert As String, strTSCert As String
          Dim objSeal As Object
          Dim strBase64 As String
          Dim strFile As String
          
          Dim i As Long
1         On Error GoTo ErrH
            
2         For i = LBound(arrCertInfo) To UBound(arrCertInfo)
3             arrCertInfo(i) = ""
4         Next
          'ȡǩ��ͼƬ
          'ESEALREAD.ESealReadCtrl 0.1
5         On Error Resume Next
6         Set objSeal = CreateObject("ESEALREAD.ESealReadCtrl.1")
7         If Err.Number <> 0 Then
              'C:\Windows\System32\HebcaP11X.dll
8             MsgBoxEx "����ǩ�¶���ʧ�ܣ����鲿����ESealRead.ocx���Ƿ���ȷ��װ��ע�ᡣ", vbInformation, gstrSysName
9             Exit Function
10        End If
11        On Error GoTo ErrH
12        strBase64 = objSeal.ReadESeal(-3)
13        If strBase64 = "" Then
14            MsgBoxEx "��ȡǩ��BASE64ʧ�ܣ�", vbInformation, gstrSysName
15            Exit Function
16        End If
          
17        If GetCertList(strCertUserName, strCertSn, strCertDN) Then
18            strFile = FormatPic("bmp", strCertSn, strBase64)
19            If strFile = "" Then
20                MsgBoxEx "����ǩ��ͼƬʧ�ܣ�", vbInformation, gstrSysName
21                Exit Function
22            End If
23            arrCertInfo(0) = strCertUserName
24            arrCertInfo(1) = strCertDN
25            arrCertInfo(2) = strCertSn
26            arrCertInfo(3) = strSigCert
27            arrCertInfo(4) = ""
28            arrCertInfo(5) = strFile
29            arrCertInfo(6) = strTSCert
30            ShanXi_RegCert = True
31        End If
          
32        Exit Function

ErrH:
33        MsgBoxEx "֤��ע��ʧ�ܣ�" & vbNewLine & "��" & CStr(Erl()) & "�� " & Err.Description, vbInformation, gstrSysName
End Function

Public Sub ShanXi_SetPara(ByVal strURL As String, ByVal strHosName As String, ByVal strTSURL As String)
1         On Error GoTo ErrH

2         With gudtPara
3             gudtPara.strSignURL = strURL
4             gudtPara.strOption = strHosName
5             gudtPara.strTSIP = strTSURL
6             Call UpdateThirdPara(CON_PAR_����, 1, "ǩ������WSDL", .strSignURL, "ǩ������WSDL")
7             Call UpdateThirdPara(CON_PAR_����, 2, "ϵͳ��ʶ", .strOption, "ϵͳΨһ��ʶ")
8             Call UpdateThirdPara(CON_PAR_����, 3, "ʱ�������WSDL", .strTSIP, "ʱ�������WSDL")
9         End With

10        Exit Sub

ErrH:
11        MsgBox "��ShanXi_SetPara�ĵ�" & Erl() & "�г���" & vbCrLf & _
                  "�����: " & Err.Number & vbCrLf & _
                  "����������" & Err.Description, vbExclamation, gstrSysName
End Sub

Public Function ShanXi_Sign(ByVal strSource As String, ByRef strSignData As String, ByRef strTimeStamp As String, ByRef strTimeStampCode As String)
      '����:ǩ��
          Dim strDigest As String
          Dim strGUID As String
          Dim strBase64 As String
          Dim strGUIDTemp As String
          
          Dim objCert As Object
          
1         On Error GoTo ErrH
          
2         If ShanXi_CheckCert(objCert) Then
              '��ԭ�����ݲ���hashֵ
3             If objCert.IsRSACert Then
4                 strDigest = mobjCertMgr.util.HashText(strSource, 1)
5             ElseIf objCert.IsSM2Cert Then
6                 strDigest = mobjCertMgr.util.HashText(strSource, 2)
7             End If
              '��hashֵ��ǩ��
8             strSignData = SignedData(strDigest, objCert)
9             If strSignData = "" Then Exit Function
              
              '���ýӿ�checkSNCAPKCS7SignAndSaveToDB �ϴ�����֤������
10            strGUID = GUID()
11            If strGUID = "" Then
12                MsgBoxEx "��ȡGUIDʧ�ܣ�", vbInformation, gstrSysName
13                Exit Function
14            End If
15            If Not SaveToDB(strSignData, gudtPara.strOption, strGUID, "00") Then
16                MsgBoxEx "�ϴ�������ʧ�ܣ�", vbInformation, gstrSysName
17                Exit Function
18            End If
19            strSignData = strSignData & M_STR_SPLIT & strGUID
              '��ȡʱ���
20            strBase64 = EncodeBase64String(strSource)
21            strDigest = getHashValue(strBase64)
22            strGUIDTemp = left(strGUID, 1) & "," & Mid(strGUID, 2, Len(strGUID) - 2) & "," & Right(strGUID, 1)
23            strTimeStampCode = SignByTSA(strGUIDTemp, strDigest)
24            If strTimeStampCode = "" Then
25                MsgBoxEx "��ȡʱ�����Ϣʧ�ܣ�", vbInformation, gstrSysName
26                Exit Function
27            End If
28            strTimeStamp = getSignTime(strTimeStampCode)
29            If strTimeStamp = "" Then
30                MsgBoxEx "��ʱ���ǩ��ֵ�л�ȡǩ��ʱ��ʧ�ܣ�", vbInformation, gstrSysName
31                Exit Function
32            End If
33            strTimeStamp = Format(strTimeStamp, "yyyy-MM-dd HH:mm:ss")
34        End If
35        ShanXi_Sign = True
          
36        Exit Function

ErrH:
37        MsgBox "��ShanXi_Sign�ĵ�" & Erl() & "�г���" & vbCrLf & _
                  "�����: " & Err.Number & vbCrLf & _
                  "����������" & Err.Description, vbExclamation, gstrSysName

End Function

Public Sub ShanXi_UnloadObj()
    Set mobjCertMgr = Nothing
End Sub

 
Public Function ShanXi_VerifySign(ByVal strSign As String, ByVal strSource As String) As Boolean
      '����:��֤ǩ��
      'ʱ�����֤ǩ��
          Dim arrSign As Variant
          Dim strGUID As String
          Dim strGUIDTemp As String
          Dim strBase64 As String
          Dim strDigest As String
          Dim blnRet As Boolean
          Dim strMsg As String
1         On Error GoTo ErrH

2         arrSign = Split(strSign, M_STR_SPLIT) 'ǩ��ֵ<SPLIT>GUID
3         If CheckP7(arrSign(0)) Then
4             strMsg = "��֤�ɹ���ǩ��������Ч��"
5         Else
6             MsgBoxEx "��֤ʧ�ܣ�ǩ��������Ч��", vbInformation, gstrSysName
7             Exit Function
8         End If
          
9         If UBound(arrSign) >= 1 Then
10            strGUID = arrSign(1)
11            strBase64 = EncodeBase64String(strSource)
12            strDigest = getHashValue(strBase64)
13            strGUIDTemp = left(strGUID, 1) & "," & Mid(strGUID, 2, Len(strGUID) - 2) & "," & Right(strGUID, 1)
14            blnRet = verifyContentByTSA(strGUIDTemp, strDigest)
15            If blnRet Then
16                strMsg = strMsg & vbCrLf & "ʱ�����֤�ɹ���"
17            Else
18                MsgBoxEx "ʱ�����֤ʧ�ܣ�", vbInformation, gstrSysName
19                Exit Function
20            End If
21        End If
22        If strMsg <> "" Then
23           MsgBoxEx strMsg, vbInformation, gstrSysName
24        End If
25        ShanXi_VerifySign = True

26        Exit Function

ErrH:
27        MsgBox "��ShanXi_VerifySign�ĵ�" & Erl() & "�г���" & vbCrLf & _
                  "�����: " & Err.Number & vbCrLf & _
                  "����������" & Err.Description, vbExclamation, gstrSysName
End Function

Private Function SignedData(ByVal strSource As String, ByVal objCert As Object) As String
          Dim objPkcs7 As Object
          Dim strSignData As String
          Dim strCertB64 As String
          
1         On Error GoTo ErrH

2         strCertB64 = objCert.GetCertB64()
3         Set objPkcs7 = mobjCertMgr.CreatePkcs7()
4         objPkcs7.AddRecipientCert (strCertB64)
5         If objCert.IsRSACert Then
6             On Error Resume Next
7             strSignData = objPkcs7.SignText(0, strSource, 1)
8             If Err.Number = -536145911 Then Exit Function 'ȡ�����봰��
9             If Err.Number > 0 Then GoTo ErrH
10            On Error GoTo ErrH
11        ElseIf objCert.IsSM2Cert Then
12            On Error Resume Next
13            strSignData = objPkcs7.SignText(0, strSource, 2)
14            If Err.Number = -536145911 Then Exit Function  '�û�ȡ������
15            If Err.Number > 0 Then GoTo ErrH
16            On Error GoTo ErrH
17        Else
18            MsgBoxEx "֤�鲻֧��RSA/SM2�㷨��", vbInformation, gstrSysName
19        End If
20        If strSignData <> "" Then
21            If Not objPkcs7.VerifyB64(strSignData) Then
22                MsgBoxEx "��֤ǩ��ʧ��(RSA)��", vbInformation, gstrSysName
23                Exit Function
24            End If
25        Else
26            MsgBoxEx "ǩ��ʧ�ܣ�ǩ��ֵΪ�գ�", vbInformation, gstrSysName
27        End If
28        SignedData = strSignData
29        Exit Function

ErrH:
30        MsgBox "��SignedData�ĵ�" & Erl() & "�г���" & vbCrLf & _
                  "�����: " & Err.Number & vbCrLf & _
                  "����������" & Err.Description, vbExclamation, gstrSysName
End Function

'Private Function SNCAGetCertForPwd(ByVal strPWD As String) As Object
'      ' ����ͻ����û�֤��PIN���ȡ֤����Ϣ
'      ' ����1:֤������Ĭ��Ϊ0
'      ' ����2:�û�֤��PIN��
'      ' ����:�û��ͻ���ǩ��֤�����
'          Dim intCount As Integer
'          Dim i As Long
'
'          Dim objCert As Object
'
'1         On Error GoTo ErrH
'
'2         intCount = mobjCertMgr.GetCertCount
'3         For i = 0 To intCount
'4             Set objCert = mobjCertMgr.GetSignCert(i)
'5             On Error Resume Next
'6             Call objCert.Login(strPWD)
'7             If Err.Number <> 0 Then
'8               MsgBox Err.Description
'9             End If
'10            On Error GoTo 0
'11            If Not objCert Is Nothing Then
'12                Exit For
'13            End If
'14        Next
'15        Set SNCAGetCertForPwd = objCert
'
'16        Exit Function
'
'ErrH:
'17        MsgBox "��SNCAGetCertForPwd�ĵ�" & Erl() & "�г���" & vbCrLf & _
'                  "�����: " & Err.Number & vbCrLf & _
'                  "����������" & Err.Description, vbExclamation, gstrSysName

Public Function VerifyB64(ByVal strSignData As String) As Boolean
      '����:��֤ǩ��
          Dim objCert As Object
          Dim objPkcs7 As Object
          
          Dim strCertB64 As String
          Dim blnResult As Boolean
          
1         On Error GoTo ErrH

2         Set objCert = mobjCertMgr.SelectSignCert
3         strCertB64 = objCert.GetCertB64
4         Set objPkcs7 = mobjCertMgr.CreatePkcs7()
5         Call objPkcs7.AddRecipientCert(strCertB64)
6         blnResult = objPkcs7.VerifyB64(strSignData)
          VerifyB64 = blnResult
7         Exit Function

ErrH:
8         MsgBox "��VerifyB64�ĵ�" & Erl() & "�г���" & vbCrLf & _
                  "�����: " & Err.Number & vbCrLf & _
                  "����������" & Err.Description, vbExclamation, gstrSysName
End Function

Public Function SignByTSA(ByVal strBusinessID As String, ByVal strHashSource As String) As String
 '����:ʱ���ǩ��
 '������strBusinessID ��λ���","ĩλǰ��"," ���磺6,1234567,8
    Dim strEnvelope As String
    Dim strBase64 As String
   
    On Error GoTo ErrH

    strEnvelope = "<soapenv:Envelope xmlns:soapenv=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:web=""http://webservice.client.tsp.snca.com/"">" & vbNewLine & _
                "   <soapenv:Header/>" & vbNewLine & _
                "   <soapenv:Body>" & vbNewLine & _
                "      <web:signByTSA>" & vbNewLine & _
                "         <arg0>HASH</arg0>" & vbNewLine & _
                "         <arg1>" & strBusinessID & "</arg1>" & vbNewLine & _
                "         <arg2>" & strHashSource & "</arg2>" & vbNewLine & _
                "         <arg3>SHA</arg3>" & vbNewLine & _
                "      </web:signByTSA>" & vbNewLine & _
                "   </soapenv:Body>" & vbNewLine & _
                "</soapenv:Envelope>"
    LogWrite "SignByTSA", "���á�SignByTSA������ֵ:" & strEnvelope
    strBase64 = httpPostSOAP(gudtPara.strTSIP, strEnvelope, ".//return")
    LogWrite "SignByTSA", "���á�SignByTSA������ֵ:" & strBase64
    SignByTSA = strBase64

    Exit Function

ErrH:
    MsgBox "��SignByTSA�ĵ�" & Erl() & "�г���" & vbCrLf & _
            "�����: " & Err.Number & vbCrLf & _
            "����������" & Err.Description, vbExclamation, gstrSysName
End Function

Public Function verifyContentByTSA(ByVal strBusinessID As String, ByVal strHash As String) As Boolean
'����:��֤ʱ���ǩ��
'������strBusinessID ��λ���","ĩλǰ��"," ���磺6,1234567,8
'      strHash-ԭ��ժҪ
    Dim strEnvelope As String
    Dim strResult As String
    
    On Error GoTo ErrH

    strEnvelope = "<soapenv:Envelope xmlns:soapenv=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:web=""http://webservice.client.tsp.snca.com/"">" & vbNewLine & _
                "   <soapenv:Header/>" & vbNewLine & _
                "   <soapenv:Body>" & vbNewLine & _
                "      <web:verifyContentByTSA>" & vbNewLine & _
                "         <arg0>HASH</arg0>" & vbNewLine & _
                "         <arg1>" & strBusinessID & "</arg1>" & vbNewLine & _
                "         <arg2>" & strHash & "</arg2>" & vbNewLine & _
                "         <arg3>SHA</arg3>" & vbNewLine & _
                "      </web:verifyContentByTSA>" & vbNewLine & _
                "   </soapenv:Body>" & vbNewLine & _
                "</soapenv:Envelope>"
    LogWrite "verifyContentByTSA", "���á�verifyContentByTSA������ֵ:" & strEnvelope
    strResult = httpPostSOAP(gudtPara.strTSIP, strEnvelope, ".//return")
    LogWrite "verifyContentByTSA", "���á�verifyContentByTSA������ֵ:" & strResult
    verifyContentByTSA = IIf(UCase(strResult) = UCase("True"), True, False)
    Exit Function

ErrH:
    MsgBox "��verifyContentByTSA�ĵ�" & Erl() & "�г���" & vbCrLf & _
            "�����: " & Err.Number & vbCrLf & _
            "����������" & Err.Description, vbExclamation, gstrSysName
End Function


Public Function getHashValue(ByVal strBase64Source As String) As String
'����:���ָ���㷨�����ĵ�ժҪֵ��֧��SHA1 ժҪ�㷨
'����:strBase64Source-ԭ��תBase64�ַ���

    Dim strEnvelope As String
    Dim strBase64 As String

    On Error GoTo ErrH

    strEnvelope = "<soapenv:Envelope xmlns:soapenv=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:web=""http://webservice.client.tsp.snca.com/"">" & vbNewLine & _
                    "   <soapenv:Header/>" & vbNewLine & _
                    "   <soapenv:Body>" & vbNewLine & _
                    "      <web:getHashValue>" & vbNewLine & _
                    "         <arg0>" & strBase64Source & "</arg0>" & vbNewLine & _
                    "         <arg1>SHA</arg1>" & vbNewLine & _
                    "      </web:getHashValue>" & vbNewLine & _
                    "   </soapenv:Body>" & vbNewLine & _
                    "</soapenv:Envelope>"

    LogWrite "getHashValue", "���á�getHashValue������ֵ:" & strEnvelope
    strBase64 = httpPostSOAP(gudtPara.strTSIP, strEnvelope, ".//return")
    LogWrite "getHashValue", "���á�getHashValue������ֵ:" & strBase64
    getHashValue = strBase64

    Exit Function

ErrH:
    MsgBox "��getHashValue�ĵ�" & Erl() & "�г���" & vbCrLf & _
            "�����: " & Err.Number & vbCrLf & _
            "����������" & Err.Description, vbExclamation, gstrSysName
End Function

Public Function getSignTime(ByVal strTimeStampCode As String) As String
'����:��ʱ���ǩ��ֵ�л�ȡǩ��ʱ��
'����:strTimeStampCode-ʱ���ǩ��ֵ

    Dim strEnvelope As String
    Dim strResult As String

    On Error GoTo ErrH

    strEnvelope = "<soapenv:Envelope xmlns:soapenv=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:web=""http://webservice.client.tsp.snca.com/"">" & vbNewLine & _
                    "   <soapenv:Header/>" & vbNewLine & _
                    "   <soapenv:Body>" & vbNewLine & _
                    "      <web:getSignTime>" & vbNewLine & _
                    "         <arg0>" & strTimeStampCode & "</arg0>" & vbNewLine & _
                    "      </web:getSignTime>" & vbNewLine & _
                    "   </soapenv:Body>" & vbNewLine & _
                    "</soapenv:Envelope>"

    LogWrite "getSignTime", "���á�getSignTime������ֵ:" & strEnvelope
    strResult = httpPostSOAP(gudtPara.strTSIP, strEnvelope, ".//return")
    LogWrite "getSignTime", "���á�getSignTime������ֵ:" & strResult
    getSignTime = strResult

    Exit Function

ErrH:
    MsgBox "��getSignTime�ĵ�" & Erl() & "�г���" & vbCrLf & _
            "�����: " & Err.Number & vbCrLf & _
            "����������" & Err.Description, vbExclamation, gstrSysName
End Function


