Attribute VB_Name = "mdlHuBeiCA"
Option Explicit

Private HUBEI_Client As Object  'CSVS_C_SDK.1
Private HUBEI_SVS As Object  'SVS_S_SDK.1
Private HUBEI_TS As Object  'SVS_S_SDK.1
Private HUBEI_PIC As Object  'HBCA_SOFSeal.Seal.1

Private mblnInit As Boolean     '��Ƕ����Ƿ�ж��
Private mintLogin As Integer
Private mstrMethod As String        'RSA-��������ҽԺ;SM2-���ҽԺ

Public Function HUBEI_InitObj() As Boolean
    '֤�鲿����ʼ��
        On Error GoTo errH
        Dim strSIGNIP As String, intSignPort As Integer, strSignURL As String
        Dim strTSIP As String, intTSPort As Integer, strTSURL As String
        Dim arrList As Variant
        Dim strTmp As String
        Dim lngRet As Long
        
        If mblnInit Then HUBEI_InitObj = True: Exit Function
        
        On Error GoTo 0
    
1000    Set HUBEI_Client = CreateObject("CSVS_C_SDK")
2000    Set HUBEI_SVS = CreateObject("SVS_S_SDK")   '����֤����֤�ؼ�
3000    Set HUBEI_TS = CreateObject("SVS_S_SDK")   '����֤����֤�ؼ�
4000    Set HUBEI_PIC = CreateObject("HBCA_SOFSeal.Seal")
        
        gstrPara = gobjComLib.zlDatabase.GetPara(90000, glngSys)
         'gstrPara = "221.232.224.75&&&8082&&&221.232.224.75&&&8084&&&RSA|SM2"   '��ʽ:ǩ��������&&&ʱ��������� ������IP&&&�˿ں�&&&IP&&&�˿ں�&&&KYE�㷨����
        
        If gstrPara = "" Then
            Err.Raise -1, , "��ǰϵͳ��" & 100 & "��û�����õ���ǩ������,�뵽�������������á����á�"
            Exit Function
        End If
        'ǩ��������URL:/hbcaDSS/hbusiness
        'ʱ���������URL:/hbcaTSS/hbusiness
        arrList = Split(gstrPara, "&&&")
        If UBound(arrList) < 3 Then
            MsgBoxEx "ǩ����������ַ���ø�ʽ����,�뵽�������������������á�", vbOKOnly + vbInformation, gstrSysName
            Exit Function
        End If
        lngRet = -1
        strSIGNIP = arrList(0)
        intSignPort = CInt(arrList(1))
        strSignURL = "/hbcaDSS/hbusiness"
        lngRet = HUBEI_SVS.SOF_SetServerInfo(strSIGNIP, intSignPort, strSignURL, 80)
        If lngRet <> 0 Then
            MsgBoxEx "ǩ����������ʼ��ʧ�ܣ�", vbOKOnly + vbInformation, gstrSysName
            Exit Function
        End If
        lngRet = -1
        strTSIP = arrList(2)
        intTSPort = CInt(arrList(3))
        strTSURL = "/hbcaTSS/hbusiness"
        lngRet = HUBEI_TS.SOF_SetServerInfo(strTSIP, intTSPort, strTSURL, 80)
        If lngRet <> 0 Then
            MsgBoxEx "ʱ�����������ʼ��ʧ�ܣ�", vbOKOnly + vbInformation, gstrSysName
            Exit Function
        End If
        If UBound(arrList) >= 4 Then
            mstrMethod = arrList(4)  'SM2-���ҽԺ
        Else
            mstrMethod = "RSA"   'RSA-��������ҽԺ
        End If
        mintLogin = 0
        gstrLogins = ""
        mblnInit = True
        HUBEI_InitObj = True
    
    Exit Function
errH:
     MsgBoxEx "�����ӿڲ���ʧ�ܣ�" & vbNewLine & Err.Description, vbQuestion, gstrSysName
End Function

Public Function HUBEI_RegCert(arrCertInfo As Variant) As Boolean
'���ܣ��ṩ��HIS���ݿ���ע�����֤��ı�Ҫ��Ϣ,�����·��Ż����֤��,,��Ҫ����USB-Key
'���أ�arrCertInfo��Ϊ���鷵��֤�������Ϣ
'      0-ClientSignCertCN:�ͻ���ǩ��֤�鹫������(����),����ע��֤��ʱ������֤���
'      1-ClientSignCertDN:�ͻ���ǩ��֤������(ÿ��Ψһ)
'      2-ClientSignCertSN:�ͻ���ǩ��֤�����к�(ÿ֤��Ψһ)
'      3-ClientSignCert:�ͻ���ǩ��֤������
'      4-ClientEncCert:�ͻ��˼���֤������
'      5-ǩ��ͼƬ�ļ���,�մ���ʾû��ǩ��ͼƬ
        
        Dim strCertID As String, strCertUserName As String, strPicPath As String
        Dim strCert As String, i As Integer
        Dim strCertSn As String
        On Error GoTo errH
        
        For i = LBound(arrCertInfo) To UBound(arrCertInfo)
             arrCertInfo(i) = ""
        Next
        
        If HUBEI_GetCertList(strCertUserName, strCertSn, strCertID, strPicPath) Then
            arrCertInfo(0) = strCertUserName
            arrCertInfo(1) = strCertSn '֤��DN
            arrCertInfo(2) = strCertSn '֤�����к�(֤��������ΪΨһֵ)
            arrCertInfo(3) = ""
            arrCertInfo(4) = ""
            arrCertInfo(5) = strPicPath
                
            HUBEI_RegCert = True
        End If

        Exit Function
errH:
     MsgBoxEx "��ȡ֤����Ϣʧ�ܣ�" & vbNewLine & "��" & CStr(Erl()) & "�� " & Err.Description, vbQuestion, gstrSysName

End Function


Public Function HUBEI_GetCertList(ByRef strName As String, Optional ByRef strCertSn As String, Optional ByRef strCertID As String, Optional ByRef strPicPath As String = "0") As Boolean
'�������ȡ����֤���б���
'-���:��
'-����
'strName :      ����ӿڷ��ص�֤������������
'strCertSN      ����ӿڷ��ص�֤��SN(ʵ�ʷ��ص���֤������DN,CA������Ψһ��,���֤�Ų��Ǳ�����
'strCertID      ����֤��ID ǩ��ʱ��Ҫ��
'strPicPath     ȱʡ����ȡ

    Dim strUsbkeyList As String
    Dim arrUserList() As String
    Dim strUser As String
    Dim strUserID As String
    Dim strTmp As String
    Dim strCertBase As String
    Dim strBase64 As String
    
    Dim i As Integer
    On Error GoTo errH:
    '��ȡ֤��
    Call HUBEI_Client.SOF_SetCertAppPolicy("SIGN")
    If mstrMethod = "SM2" Then
        HUBEI_Client.SOF_SetHashMethod ("SM3")
    End If
    strUsbkeyList = HUBEI_Client.SOF_GetUserList()
    
    If (strUsbkeyList = "") Then
        strName = ""
        MsgBoxEx "�����֤��Key��", vbOKOnly + vbInformation, gstrSysName
        HUBEI_GetCertList = False
        Exit Function
    Else
        '�û�1(CertID||Subject||IssuerSubject||CertBase64)&&&�û�(CertID||Subject||IssuerSubject||CertBase64)&&&�û�&&&�û�
        '1419118795628E856CD1B3C0DD607693||CN=����֤��2, OU=����, O=������ҽԺ, L=�人, S=����, C=CN||CN=HBCA, O=Hubei Digital Certificate Authority Center CO Ltd., L=Wuhan, S=Hubei, C=CN||MIIEkzCCA3ugAwIBAg
        arrUserList = Split(strUsbkeyList, "&&&")

        If UBound(arrUserList) > 1 Then  '���KEY
            For i = LBound(arrUserList) To UBound(arrUserList) - 1
                strTmp = Split(arrUserList(i), "||")(1)
                strTmp = Split(strTmp, ",")(0)
                strTmp = Mid(strTmp, 4)
                strUser = strUser & "&&&" & strTmp
            Next
            If strUser <> "" Then strUser = Mid(strUser, 4)
            strName = frmSelectUser.ShowMe(strUser)
            
            For i = LBound(arrUserList) To UBound(arrUserList) - 1
                strTmp = Split(arrUserList(i), "||")(1)
                strTmp = Split(strTmp, ",")(0)
                strTmp = Mid(strTmp, 4)
                If strName = strTmp Then
                     strCertSn = Split(arrUserList(i), "||")(1)
                     strCertID = Split(arrUserList(i), "||")(0)
                     strCertBase = Split(arrUserList(i), "||")(3)
                     Exit For
                End If
            Next
        Else
            arrUserList = Split(arrUserList(0), "||")
            strCertSn = arrUserList(1)      '֤��DN
            strCertID = arrUserList(0)    '֤��ID
            strCertBase = arrUserList(3)   '֤������
            strName = Mid(Split(arrUserList(1), ",")(0), 4)
        End If
        
    End If
    If strPicPath = "" Then
        strUserID = HUBEI_Client.SOF_GetCertInfoByOidEx(strCertBase, "2.4.16.11.7.3")
        strBase64 = HUBEI_PIC.SOF_GetKeyPictureEx(strCertID, strUserID)
        strPicPath = SaveBase64ToFile("gif", strCertID, strBase64)
    End If
    
    HUBEI_GetCertList = True
    
    Exit Function
errH:
     MsgBoxEx "��ȡ֤����Ϣʧ�ܣ�" & vbNewLine & "��" & CStr(Erl()) & "�� " & Err.Description, vbQuestion, gstrSysName
End Function

Public Function HUBEI_CheckCert(ByVal strCurrCertSn As String, Optional ByRef strCertID As String) As Boolean
'---------------------------------------------------------------------------------------------------------------------
'���ܣ���ȡUSB�����豸��ʼ������¼
'---------------------------------------------------------------------------------------------------------------------
    Dim strKey As String, strPIN As String, strUserName As String
    Dim strCertName As String, strCertDN As String
    Dim strCertSn As String
    Dim strCertUserID As String    '�������֤����Ϣ
    Dim strDate As String
    Dim strCert As String
    Dim blnOk As Boolean
    Dim blnRet As Boolean
    Dim lngRet As Long
    
    On Error GoTo errH
    

     '��ȡ֤����Ϣͬʱ���Key���Ƿ����
    If Not HUBEI_GetCertList(strCertName, strCertSn, strCertID) Then
        HUBEI_CheckCert = False: Exit Function
    End If
    'δע���ڵ�ǰ�û����µ�Key
    If strCurrCertSn <> strCertSn Then
        MsgBoxEx "��֤��:" & vbCrLf & _
                vbTab & "��" & strCertSn & "��" & vbCrLf _
                & "δע�����������£�����ʹ�ã�", vbInformation, gstrSysName
        Exit Function
    End If
    '������֤���������,�״ε���ǩ���ӿ�ʱ�ᴥ��CA�����봰��
    
    '��¼��֤
    If InStr(gstrLogins & "|", "|" & strCertSn & "|") > 0 Then '�״���֤ͨ�����´β��ڼ�����֤
        blnOk = True
    Else
        If Not GetCertLogin(strCertID) Then
            strPIN = ""
            blnOk = False
        Else
            If InStr(gstrLogins & "|", "|" & strCertSn & "|") = 0 Then gstrLogins = gstrLogins & "|" & strCertSn
            blnOk = True
        End If
    End If
    HUBEI_CheckCert = blnOk
    Exit Function
errH:
     MsgBoxEx "���USBKEYʧ�ܣ�" & vbNewLine & "��" & CStr(Erl()) & "�� " & Err.Description, vbInformation, gstrSysName
End Function

Public Function HUBEI_Sign(ByVal strCurrCertSn As String, ByVal strSource As String, ByRef strSignData As String, _
            ByRef strTimeStamp As String, ByRef strTimeStampCodeID As String) As Boolean
    'ǩ��
    '������
    '   strPID --�û���ݱ�ʶ��һ��Ϊ���֤�ţ�
        Dim strCertID As String
        Dim strPicPath As String
        Dim CertID As String
        Dim strTimeStampCode As String    'ʱ�������
        Dim strDate As String
        Dim strMsg As String
        Dim blnCheck As Boolean
        Dim lngRet As Long
    
        On Error GoTo errH

100     blnCheck = HUBEI_CheckCert(strCurrCertSn, strCertID)
    
102     If blnCheck Then                '��֤��ǰUSB�Ƿ���ǩ���û��ģ�����ȡǩ��֤��
104         strSource = HUBEI_Client.SOF_HashData(strSource)   'ԭ��תHASH
            'detach
106         Call HUBEI_Client.SOF_SetP7SignMode(1)
108         strSignData = HUBEI_Client.SOF_SignDataByP7(strCertID, strSource)
110         If strSignData <> "" Then
112             lngRet = -1
                'detach
114             Call HUBEI_SVS.SOF_SetP7SignMode(1)
116             lngRet = HUBEI_SVS.SOF_VerifyDetachSignedData(strSignData, strSource)
118             If lngRet = 0 Then
120                 strTimeStampCodeID = HUBEI_TS.SOF_CreateTimeStampResponse(strSignData)  'ʱ���ID�������ݿ�
122                 If strTimeStampCodeID = "" Then
124                     MsgBoxEx "����ʱ���IDʧ�ܣ�" & HUBEI_TS.SOF_GetErrorMsg(), vbOKOnly + vbInformation, gstrSysName
                        Exit Function
                    End If
126                 strTimeStamp = Format(gobjComLib.zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
                Else
128                 MsgBoxEx "��֤ǩ��ʧ�ܣ�" & HUBEI_SVS.SOF_GetErrorMsg(), vbOKOnly + vbInformation, gstrSysName
                    Exit Function
                End If
            Else
130             MsgBoxEx "ǩ��ʧ�ܣ�" & HUBEI_Client.SOF_GetErrorMsg(), vbOKOnly + vbInformation, gstrSysName
                Exit Function
            End If
        Else
            Exit Function
        End If
    
132     HUBEI_Sign = True
        Exit Function
errH:
134      MsgBoxEx "ǩ��ʧ�ܣ�" & vbNewLine & "��" & CStr(Erl()) & "�� " & Err.Description, vbInformation, gstrSysName
End Function


Public Function HUBEI_VerifySign(ByVal strSignData As String, ByVal strSource As String, ByVal strTimeStampID As String) As Boolean
    '��֤ǩ��
    Dim lngRet As Long
    Dim strTmp As String
    Dim blnRet As Boolean
    Dim strDate As String
    Dim strTimeStamp As String
    
    On Error GoTo errH
    
    lngRet = -1
    strSource = HUBEI_Client.SOF_HashData(strSource)   'ԭ��תHASH
    'detach
    Call HUBEI_SVS.SOF_SetP7SignMode(1)
    lngRet = HUBEI_SVS.SOF_VerifyDetachSignedData(strSignData, strSource)
    
    If lngRet = 0 Then
        strTmp = "��֤����ǩ���ɹ���"
        blnRet = True
    Else
        strTmp = "��֤����ǩ��ʧ�ܣ�" & HUBEI_SVS.SOF_GetErrorMsg()
        blnRet = False
    End If
    'ʱ��� ֻ�ǳ���ҽ���¹�֮����ҽ�ƻ���������֤ ��CA������ṩ�ù��ܣ�
    If strTmp <> "" Then
        MsgBoxEx strTmp, vbOKOnly + vbInformation, gstrSysName
    End If
    
    HUBEI_VerifySign = blnRet
    
    Exit Function
errH:
104     MsgBoxEx "��֤ǩ��ʧ�ܣ�" & vbNewLine & "��" & CStr(Erl()) & "�� " & Err.Description, vbInformation, gstrSysName
End Function

Private Function GetCertLogin(ByVal strCertID As String) As Boolean
    '����CA����֤���¼����
    '- ���
    'strCertID            :֤��ID
    'strPin              ������
    Dim strRandom As String
    Dim strSignValue As String
    Dim lngRet As Long
    Dim strDate As String
    Dim intDay As Integer
    
    On Error GoTo errH

    strRandom = HUBEI_Client.SOF_GenRandom(10)
    Call HUBEI_Client.SOF_SetP7SignMode(1)
    strSignValue = HUBEI_Client.SOF_SignDataByP7(strCertID, strRandom)

    Call HUBEI_SVS.SOF_SetP7SignMode(1)
    lngRet = -1
    lngRet = HUBEI_SVS.SOF_VerifyDetachSignedData(strSignValue, strRandom)

    If lngRet <> 0 Then
        MsgBoxEx "��½ʧ�ܣ�" & HUBEI_Client.SOF_GetErrorMsg(), vbOKOnly + vbInformation, gstrSysName
        Exit Function
    End If
    
    GetCertLogin = True
    Exit Function
errH:
    MsgBoxEx "��¼��֤ʧ�ܣ�" & vbNewLine & "��" & CStr(Erl()) & "�� " & Err.Description, vbInformation, gstrSysName
End Function

Public Function HUBEI_GetPara() As Boolean
'���ú���CA��������ַ
    Dim arrList As Variant
    
    On Error GoTo errH
    gstrPara = gobjComLib.zlDatabase.GetPara(90000, glngSys)
    If gstrPara = "" Then gstrPara = "221.232.224.75&&&8082&&&221.232.224.75&&&8084&&&SM2"
    If gstrPara <> "" Then
        arrList = Split(gstrPara, G_STR_SPLIT)
        If UBound(arrList) >= 3 Then
             gudtPara.strSIGNIP = Trim(arrList(0))
             gudtPara.strSignPort = Trim(arrList(1))
             gudtPara.strTSIP = Trim(arrList(2))
             gudtPara.strTSPort = Trim(arrList(3))
             If UBound(arrList) >= 4 Then
                gudtPara.bytSignVersion = IIf(Trim(arrList(4)) = "RSA", 0, 1)
            Else
                gudtPara.bytSignVersion = 0
             End If
        End If
    Else
        gudtPara.strSIGNIP = "221.232.224.75"
        gudtPara.strSignPort = 8082
        gudtPara.strTSIP = "221.232.224.75"
        gudtPara.strTSPort = 8084
        gudtPara.bytSignVersion = 0
    End If
    
    Exit Function
errH:
    MsgBoxEx "��ȡ����ʧ�ܣ�" & vbNewLine & "��" & CStr(Erl()) & "�� " & Err.Description, vbInformation, gstrSysName
End Function

Public Function HUBEI_SetParaStr() As String
    HUBEI_SetParaStr = gudtPara.strSIGNIP & G_STR_SPLIT & gudtPara.strSignPort & G_STR_SPLIT & gudtPara.strTSIP & G_STR_SPLIT & gudtPara.strTSPort & G_STR_SPLIT & IIf(gudtPara.bytSignVersion = 0, "RSA", "SM2")
End Function
'���ٶ���
Public Sub HUBEI_Unload()
    Set HUBEI_Client = Nothing
    Set HUBEI_SVS = Nothing
    Set HUBEI_TS = Nothing
    Set HUBEI_PIC = Nothing
    mblnInit = False
End Sub
