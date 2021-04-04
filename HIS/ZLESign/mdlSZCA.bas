Attribute VB_Name = "mdlSZCA"
Option Explicit
'���ڵ��ӵ���ǩ���ӿ�
'
Private mSZCAClient As Object           '֤�鲿��
Private mobjSignPic As Object           'ǩ��ͼƬ�ؼ�
Private mobjSoapClient As Object       'soap���Ӷ���
Private mobjTSA As Object               'ȷ��ʱ�������
Private mblnInit As Boolean

Private Const M_STR_SN As String = "SN"
Private Const M_STR_DN As String = "DN"
Private Const M_STR_TB As String = "TIMEB"
Private Const M_STR_TE As String = "TIMEE"
Private Const M_STR_VER As String = "VER"
Private Const M_STR_OID As String = "1.2.156.1002"

Public Function SZCA_InitObj() As Boolean
'����:����ǩ�������ʼ��
'     SOAP���Ӷ����ʼ��
        Dim strUrl As String
        Dim arrPara As Variant
        
        On Error GoTo errH

1000    SZCA_InitObj = mblnInit
1001    If mblnInit Then Exit Function
        
1002    gstrPara = gobjComLib.zlDatabase.GetPara(90000, glngSys)   '��ȡURLs �̶���ȡZLHIS ϵͳĬ��100
        '"http://127.0.0.1:8080/SZCAJavaCAS/service/SZCASafeService?wsdl&&&ʱ���IP&&&ʱ����˿ں�"
        'gstrPara = "127.0.0.1&&&8080&&&124.133.51.13&&&8888"
        If gstrPara = "" Then
            Err.Raise -1, , "��ǰϵͳ��" & glngSys & "��û�����õ���ǩ������,�뵽���õ���ǩ���ӿڴ����á�"
            Exit Function
        End If
        arrPara = Split(gstrPara, G_STR_SPLIT)
        If UBound(arrPara) <> 3 Then
            Err.Raise -2, , "��ǰϵͳ��" & glngSys & "�����õ���ǩ����������,�뵽���õ���ǩ���ӿڴ����á�"
            Exit Function
        End If
        
1005    Set mSZCAClient = CreateObject("SZCAPKI.SZCAPKICtrl.1")
1006    Set mobjSignPic = CreateObject("SZCAPDFSIGNCTRL.SZCAPdfSignCtrlCtrl.1")
1007    Set mobjTSA = CreateObject("SuresecTsaClass.tsa.1")
1008    Set mobjSoapClient = CreateObject("MSSOAP.SoapClient30")  'SOAP���Ӷ���
1009    mobjSoapClient.ClientProperty("ServerHTTPRequest") = True
        strUrl = "http://" & arrPara(0) & ":" & arrPara(1) & "/SZCAJavaCAS/services/szcaCAValidate?wsdl" '��ʽ������ַ ���⣺112774
1010    mobjSoapClient.MSSoapInit (strUrl)
         
        mobjTSA.ISetTcpServerInfo arrPara(2), arrPara(3), 20
1030    mblnInit = True
1031    SZCA_InitObj = True

1090    Exit Function

errH:
118     MsgBoxEx "��ʼ������ǩ������ʧ�ܣ�" & vbNewLine & "��" & CStr(Erl()) & "�У�" & Err.Description, vbInformation, gstrSysName
    
End Function

Public Function SZCA_RegCert(arrCertInfo As Variant) As Boolean
        '���ܣ��ṩ��HIS���ݿ���ע�����֤��ı�Ҫ��Ϣ,�����·��Ż����֤��,,��Ҫ����USB-Key
        '���أ�arrCertInfo��Ϊ���鷵��֤�������Ϣ
        '      0-ClientSignCertCN:�ͻ���ǩ��֤�鹫������(����),����ע��֤��ʱ������֤���
        '      1-ClientSignCertDN:�ͻ���ǩ��֤������(ÿ��Ψһ)
        '      2-ClientSignCertSN:�ͻ���ǩ��֤�����к�(ÿ֤��Ψһ)
        '      3-ClientSignCert:�ͻ���ǩ��֤������
        '      4-ClientEncCert:�ͻ��˼���֤������
        '      5-ǩ��ͼƬ�ļ���,�մ���ʾû��ǩ��ͼƬ
        '      6-ʱ���֤��
        Dim strKeyId As String, strCertTime As String, strCertUserName As String, strCertDN As String
        Dim strSigCert As String, strUserKey As String, strTSCert As String
        Dim strFile As String
        Dim i As Long
        On Error GoTo errH
    
100     For i = LBound(arrCertInfo) To UBound(arrCertInfo)
101         arrCertInfo(i) = ""
102     Next
        
108     If SZCA_GetCertList(strCertUserName, strKeyId, strCertDN, strSigCert, strUserKey, strFile, strTSCert) Then
200         arrCertInfo(0) = strCertUserName
201         arrCertInfo(1) = strCertDN
202         arrCertInfo(2) = strKeyId
203         arrCertInfo(3) = strSigCert
205         arrCertInfo(4) = strUserKey
206         arrCertInfo(5) = strFile
            arrCertInfo(6) = strTSCert
            SZCA_RegCert = True
        End If
        
300     Exit Function

errH:
    MsgBoxEx "֤��ע��ʧ�ܣ�" & vbNewLine & "��" & CStr(Erl()) & "�� " & Err.Description, vbInformation, gstrSysName
End Function

Private Function SZCA_GetCertList(Optional ByRef strName As String, Optional ByRef strSN As String, _
    Optional ByRef strDn As String, Optional ByRef strCert As String = "-1", _
    Optional ByRef strUserUnigueID As String = "-1", Optional ByRef strFile As String = "-1", _
    Optional ByRef strTSCert As String = "-1") As Boolean
'����:��ȡ֤����Ϣ
    Dim blnRet As Boolean
    Dim arrList As Variant
    Dim strPic As String
    Dim strPath As String, strSource As String
    Dim lngRet As Long, lngTSALen As Long, lngCertLen As Long
    Dim arrTSData(2048) As Byte
    Dim arrCertData(2048) As Byte
    Dim bytSource() As Byte
    
    On Error GoTo errH
1000    Call mSZCAClient.AxInit          '��ʼ���ӿ�
1001    Call mSZCAClient.AxSetCertFilterStr("SC;SZCA;#;#;#;")   '֤�����
1002    blnRet = mSZCAClient.AxSetKeyStore()   ' ����ǩ��,���ܵ�֤��
1003    If blnRet Then
1004        strSN = mSZCAClient.AxGetCertInfo(M_STR_SN)
1005        strDn = mSZCAClient.AxGetCertInfo(M_STR_DN)  'CN=֧̫��,OU=429320496,O=����ˮ������ҽԺ,O=�����,L=����ˮ��,ST=����ʡ,C=CN
            arrList = Split(strDn, ",")
            strName = Mid(arrList(0), 4)
1006        If strCert <> "-1" Then strCert = mSZCAClient.AxGetCertData() '֤������
1007        If strUserUnigueID <> "-1" Then strUserUnigueID = mSZCAClient.AxGetCertExt(M_STR_OID)   '��չ�� �û�ҽ������Ψһ��ʶ
            
            If strFile <> "-1" Then
1008            strPic = mobjSignPic.SZCA_GetSealDataFromKey() 'PNG��ʽ������ �������ݸ�ʽΪ 1-@@@ӡ��ͼ��1��Base64����2-@@@ӡ��ͼ��2��Base64����......n-@@@ӡ��ͼ��n��Base64����
                strPic = Split(strPic, "@@@")(1)
1020            strFile = SaveBase64ToFile("BMP", strSN, strPic)
1030            Call SaveStdPicToFile(LoadPictureGDIPlus(strFile), strFile, BMP, 100)
            End If
            If strTSCert <> "-1" Then
1040            strSource = "����ABCabc123"
1041            ReDim bytSource(LenB(StrConv(strSource, vbFromUnicode)))
1042            bytSource = StrConv(strSource, vbFromUnicode)

1050            lngRet = mobjTSA.IGenTokenByPlain("SHA1", bytSource(0), UBound(bytSource) + 1, 1, arrTSData(0), lngTSALen)
                If lngRet = 0 Then
1051                lngRet = mobjTSA.IGetTokenCertificate(arrTSData(0), lngTSALen, arrCertData(0), lngCertLen)
1052                strTSCert = FuncEncodeBase64Byte(arrCertData, lngCertLen)
                Else
                    MsgBoxEx "ʱ���֤���ȡʧ�ܣ�", vbOKOnly + vbInformation, gstrSysName
                    Exit Function
                End If
            End If
        Else
            MsgBoxEx "�����֤��Key��", vbOKOnly + vbInformation, gstrSysName
            Exit Function
        End If
1090
        SZCA_GetCertList = True
    Exit Function
errH:
    MsgBoxEx "��ȡ֤����Ϣʧ�ܣ�" & vbNewLine & "��" & CStr(Erl()) & "�� " & Err.Description, vbInformation, gstrSysName
End Function

Private Function GetCertLogin(ByVal strCert As String) As Boolean
'����:֤���¼��֤
    Dim strSource As String, strBase64 As String
    Dim strSign As String, strMsg As String
    Dim strRet As String
    
1000    On Error GoTo errH
1001
1002    Randomize
1003    strSource = Int((100000 * Rnd) + 1)
1004    strSign = mSZCAClient.AxSign(strSource)
1005    If strSign <> "" Then
1006        strBase64 = mobjSoapClient.szcaWSSignatureValidatePkcs7String(strSign)
1007        strBase64 = DecodeBase64String(strBase64)
1008
1009        If strBase64 = "1" Then
            '1 ֤����Ч -1 ֤����Ч�����������εĸ� -2 ֤����Ч��������Ч�� -3 ֤����Ч���Ѽ��������
1010            strRet = mobjSoapClient.szcaWSCertificateValidateString(strCert)
1011            strRet = DecodeBase64String(strRet)
1012            Select Case strRet
                Case "-1"
                    strMsg = "֤����Ч�����������εĸ�"
                Case "-2"
                    strMsg = "֤����Ч��������Ч��"
                Case "-3"
                    strMsg = "֤����Ч���Ѽ��������"
1020            End Select
            Else
                strMsg = "��¼��֤ʧ�ܣ���֤��Ϣ���¼��Ϣ����"
            End If
        Else
            strMsg = "��¼ʧ�ܣ�"
        End If
        If strMsg <> "" Then
            MsgBoxEx strMsg, vbOKOnly + vbInformation, gstrSysName
        End If
1050
    GetCertLogin = True
    Exit Function
errH:
    MsgBoxEx "֤���¼ʧ�ܣ�" & vbNewLine & "��" & CStr(Erl()) & "�� " & Err.Description, vbInformation, gstrSysName
End Function

Public Function SZCA_CheckCert(ByVal strCurrCertSn As String) As Boolean
'���ܣ���ȡUSB�����豸��ʼ������¼
'����ֵ:
'  strSigCert -ǩ��֤������

        Dim strSN As String, strSigCert As String

        On Error GoTo errH
1000    If Not SZCA_InitObj() Then
1002        MsgBoxEx "����δ��ʼ����"
            Exit Function
        End If
        
1004    If Not SZCA_GetCertList(, strSN, , strSigCert) Then Exit Function
1006    If strCurrCertSn <> strSN Then
1008        MsgBoxEx "��֤��δע�����������£�����ʹ�ã�", vbInformation + vbOKOnly, gstrSysName
            Exit Function
        End If
       
1010
        '��¼��֤
        If Not InStr(gstrLogins & "|", "|" & strCurrCertSn & "|") > 0 Then '�״���֤ͨ�����´β��ڼ�����֤
            If Not GetCertLogin(strSigCert) Then
                Exit Function
            Else
                If InStr(gstrLogins & "|", "|" & strCurrCertSn & "|") = 0 Then gstrLogins = gstrLogins & "|" & strCurrCertSn
            End If
        End If
1016
        SZCA_CheckCert = True
        Exit Function
errH:
124     MsgBoxEx "���USBKEYʧ�ܣ�" & vbNewLine & "��" & CStr(Erl()) & "�� " & Err.Description, vbInformation, gstrSysName
End Function

Public Function SZCA_Sign(ByVal strSN As String, ByVal strSource As String, ByRef strSignData As String, ByRef strTimeStamp As String, ByRef strTimeStampCode As String)
'����:ǩ��
    Dim blnRet As Boolean
    Dim strMsg As String
    Dim strDigest As String
    Dim strTime As String
    Dim bytSource() As Byte
    Dim arrTSData(2048) As Byte
    Dim arrCertData(2048) As Byte
    Dim lngTSALen As Long
    Dim lngRet As Long
    
    On Error GoTo errH
    
1000    If SZCA_CheckCert(strSN) Then
            strDigest = StringSHA1(strSource)
1005        strSignData = mSZCAClient.AxSign(strDigest)
            If strSignData = "" Then
                strMsg = "ǩ��ʧ�ܣ�����ǩ��ֵΪ�ա�"
            Else
1008            strTimeStamp = Format(gobjComLib.zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
            End If
1010
1020    End If

        If strMsg <> "" Then
            MsgBoxEx strMsg, vbInformation + vbOKOnly, gstrSysName
            Exit Function
        End If
1030
        '��ȡʱ���
        ReDim bytSource(LenB(StrConv(strSource, vbFromUnicode)))
        bytSource = StrConv(strSource, vbFromUnicode)
        'ͨ�����Ļ�ȡʱ���
        Call mobjTSA.IGenTokenByPlain("SHA1", bytSource(0), UBound(bytSource) + 1, 0, arrTSData(0), lngTSALen)
        '��ȡʱ�����ʱ��
        lngRet = mobjTSA.IGetTokenGenerateTime(arrTSData(0), lngTSALen, strTime)
1050
        If lngRet = 0 Then
            strTimeStamp = String14ToDate(strTime, strMsg)
            strTimeStampCode = FuncEncodeBase64Byte(arrTSData, lngTSALen)
        Else
            strMsg = "ʱ���ʧ�ܣ���ȡʱ�����ʱ��ʧ�ܡ�"
        End If
        If strMsg <> "" Then
            MsgBoxEx strMsg, vbInformation + vbOKOnly, gstrSysName
            Exit Function
        End If
1080
        SZCA_Sign = True
        Exit Function
errH:
    MsgBoxEx "ǩ��ʧ�ܣ�" & vbNewLine & "��" & CStr(Erl()) & "�� " & Err.Description, vbInformation, gstrSysName
End Function

Public Function SZCA_VerifySign(ByVal strCert As String, ByVal strSign As String, ByVal strSource As String, _
    ByVal strTSCert As String, ByVal strTStampCode As String) As Boolean
'����:��֤ǩ��
    Dim strBase64 As String
    Dim strRet As String, strDigest As String
    Dim strMsg As String
    Dim blnRet As Boolean
    Dim lngRet As Long
    Dim bytSource() As Byte, bytTStamp() As Byte, bytTSCert() As Byte
    
    On Error GoTo errH
    
1000    strBase64 = mobjSoapClient.szcaWSSignatureValidatePkcs7String(strSign)
1001    strRet = DecodeBase64String(strBase64)
1002    blnRet = False
        If strRet = "1" Then
            '1 ֤����Ч -1 ֤����Ч�����������εĸ� -2 ֤����Ч��������Ч�� -3 ֤����Ч���Ѽ��������
            strBase64 = ""
1005        strBase64 = mobjSoapClient.szcaWSCertificateValidateString(strCert)
1006        strRet = DecodeBase64String(strBase64)
1007        Select Case strRet
                Case "1":
                    strMsg = "��ǩ�ɹ�"
                    blnRet = True
                Case "-1":
                    strMsg = "֤����Ч�����������εĸ�"
                Case "-2":
                    strMsg = "֤����Ч��������Ч��"
                Case "-3":
                    strMsg = "֤����Ч���Ѽ��������"
            End Select
        Else
            strMsg = "��¼��֤ʧ�ܣ���֤��Ϣ���¼��Ϣ����"
        End If
        If blnRet Then
1010
            ReDim bytSource(LenB(StrConv(strSource, vbFromUnicode)))
            bytSource = StrConv(strSource, vbFromUnicode)
            bytTSCert = DecodeBase64Byte(strTSCert)
            bytTStamp = DecodeBase64Byte(strTStampCode)
1020
            lngRet = mobjTSA.IVerifyTimeStampTokenByPlain("SHA1", bytSource(0), UBound(bytSource) + 1, bytTStamp(0), UBound(bytTStamp) + 1, bytTSCert(0), UBound(bytTSCert) + 1)
            If lngRet = 0 Then
                blnRet = True
            Else
                strMsg = "ʱ�����֤ʧ��"
                blnRet = False
            End If
        End If
1050
        If strMsg <> "" Then
            MsgBoxEx strMsg, vbInformation + vbOKOnly, gstrSysName
        End If
        SZCA_VerifySign = blnRet
1090
    Exit Function
errH:
    MsgBoxEx "��ǩʧ�ܣ�" & vbNewLine & "��" & CStr(Erl()) & "�� " & Err.Description, vbInformation, gstrSysName
End Function

Public Function SZCA_GetPara() As Boolean
'��������CA��������ַ
    Dim arrList As Variant
    
    On Error GoTo errH
    gstrPara = gobjComLib.zlDatabase.GetPara(90000, glngSys)
    If gstrPara = "" Then gstrPara = "127.0.0.1&&&8080&&&124.133.51.13&&&8888"
    If gstrPara <> "" Then
        arrList = Split(gstrPara, G_STR_SPLIT)
        If UBound(arrList) = 3 Then
             gudtPara.strSIGNIP = Trim(arrList(0))
             gudtPara.strSignPort = Trim(arrList(1))
             gudtPara.strTSIP = Trim(arrList(2))
             gudtPara.strTSPort = Trim(arrList(3))
        End If
    End If
    Exit Function
errH:
    MsgBoxEx "��ȡ����ʧ�ܣ�" & vbNewLine & "��" & CStr(Erl()) & "�� " & Err.Description, vbInformation, gstrSysName
End Function

Public Function SZCA_SetParaStr() As String
    SZCA_SetParaStr = gudtPara.strSIGNIP & G_STR_SPLIT & gudtPara.strSignPort & G_STR_SPLIT & gudtPara.strTSIP & G_STR_SPLIT & gudtPara.strTSPort
End Function

Public Sub SZCA_UnloadObj()
    Set mSZCAClient = Nothing
    Set mobjSignPic = Nothing
    Set mobjSoapClient = Nothing
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


