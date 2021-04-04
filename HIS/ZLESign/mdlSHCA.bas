Attribute VB_Name = "mdlSHCA"
Option Explicit
'�Ϻ�CA���Ĺ���ģ��
Private mblnInit As Boolean         '�Ƿ��ѳ�ʼ���ɹ�
Private mLastPWD As String          '�������������

Private SHCA_Client As Object       '֤�鲿��
Private mLogin As Long              '��������������

Public Enum SH_Version
    V_SEH = 0
    V_ESE = 1
End Enum

Public Function SHCA_InitObj() As Boolean
    '֤�鲿����ʼ��
        Dim progID As String
        
        On Error GoTo errH
100
102     SHCA_InitObj = mblnInit
104     If mblnInit Then Exit Function
105     mLastPWD = ""
        If Not SHCA_GetPar(1) Then Exit Function
108     Set SHCA_Client = CreateObject("SafeEngineCOM.SafeEngineCtl")
        If gudtPara.bytSignVersion = V_SEH Then
            Call SHCA_Client.SEH_InitialSession(2, "", "", 0, 2, "", "") '��ʼ��CA�ӿ�
        Else
            Call SHCA_Client.ESE_InitialSession(2, "", "", 0, 2, "", "") '��ʼ��CA�ӿ�
        End If
        If SHCA_Client.errorCode <> 0 Then
            GoTo errH
        End If
114     SHCA_InitObj = True
    
116     mblnInit = SHCA_InitObj
        mLogin = 0
        Exit Function
errH:
118     MsgBoxEx "�����ӿڲ���ʧ�ܣ�" & vbNewLine & Err.Description, vbQuestion, gstrSysName
End Function

Public Function SHCA_RegCert(arrCertInfo As Variant) As Boolean
        '���ܣ��ṩ��HIS���ݿ���ע�����֤��ı�Ҫ��Ϣ,�����·��Ż����֤��,,��Ҫ����USB-Key
        '���أ�arrCertInfo��Ϊ���鷵��֤�������Ϣ
        '      0-ClientSignCertCN:�ͻ���ǩ��֤�鹫������(����),����ע��֤��ʱ������֤���
        '      1-ClientSignCertDN:�ͻ���ǩ��֤������(ÿ��Ψһ)
        '      2-ClientSignCertSN:�ͻ���ǩ��֤�����к�(ÿ֤��Ψһ)
        '      3-ClientSignCert:�ͻ���ǩ��֤������
        '      4-ClientEncCert:�ͻ��˼���֤������
        '      5-ǩ��ͼƬ�ļ���,�մ���ʾû��ǩ��ͼƬ
        
        Dim strKeyId As String, strCertTime As String, strCertUserName As String, strCertDN As String, strCertSn As String
        Dim strSigCert As String, i As Integer, strCACert As String, lngOk As Long
        Dim strPicData As String
        On Error GoTo errH
    
100     For i = LBound(arrCertInfo) To UBound(arrCertInfo)
102         arrCertInfo(i) = ""
        Next
    
104     If GetCertList(strCertUserName, strKeyId, strSigCert, strCertSn) Then
106         arrCertInfo(0) = strCertUserName
108         arrCertInfo(1) = GetCertDN(strSigCert)
110         arrCertInfo(2) = strCertSn
112         arrCertInfo(3) = strSigCert
124         SHCA_RegCert = True
        End If

        Exit Function
errH:
126     MsgBoxEx "��ȡ֤����Ϣʧ�ܣ�" & vbNewLine & "��" & CStr(Erl()) & "�� " & Err.Description, vbQuestion, gstrSysName

End Function

Public Function GetCertDN(strCert As String) As String
    Dim strCertDN As String
    Dim strCN As String, strO As String, strOU As String, strS As String, strL As String, strC As String, strE As String
    If gudtPara.bytSignVersion = V_SEH Then
        strC = SHCA_Client.SEH_GetCertDetail(strCert, 13)
        strO = SHCA_Client.SEH_GetCertDetail(strCert, 14)
        strOU = SHCA_Client.SEH_GetCertDetail(strCert, 15)
        strS = SHCA_Client.SEH_GetCertDetail(strCert, 16)
        strCN = SHCA_Client.SEH_GetCertDetail(strCert, 17)
        strL = SHCA_Client.SEH_GetCertDetail(strCert, 18)
    Else
        strC = SHCA_Client.ESE_GetCertDetail(strCert, 13)
        strO = SHCA_Client.ESE_GetCertDetail(strCert, 14)
        strOU = SHCA_Client.ESE_GetCertDetail(strCert, 15)
        strS = SHCA_Client.ESE_GetCertDetail(strCert, 16)
        strCN = SHCA_Client.ESE_GetCertDetail(strCert, 17)
        strL = SHCA_Client.ESE_GetCertDetail(strCert, 18)
        strE = SHCA_Client.ESE_GetCertDetail(strCert, 19)
    End If
    strCertDN = IIf(strS = "", "", "S=" & strS & ",") & IIf(strL = "", "", "L=" & strL & ",") & IIf(strO = "", "", "O=" & strO & ",") _
    & IIf(strOU = "", "", "OU=" & strOU & ",") & IIf(strCN = "", "", "CN=" & strCN & ",") & IIf(strE = "", "", "E=" & strE)
    GetCertDN = strCertDN
End Function
Public Function SHCA_Sign(ByVal strCurrCertSn As String, ByVal strSource As String, ByRef strSignData As String, ByRef strTimeStamp As String, ByRef blnReDo As Boolean) As Boolean
        'ǩ��
        Dim strSigCert As String
        Dim blnCheck As Boolean
        Dim datTime As Date
        Dim strDate As String

        On Error GoTo errH
        blnCheck = SHCA_CheckCert("", "", blnReDo)
        If blnReDo Then Exit Function
        If blnCheck Then
            '֤��ID����ǩ��
            datTime = gobjComLib.zlDatabase.Currentdate()
            strDate = Format(datTime, "yyyyMMddhhmmss")
            strTimeStamp = Format(datTime, "yyyy-MM-dd HH:mm:ss")
            If gudtPara.bytSignVersion = V_SEH Then
                strSignData = SHCA_Client.SEH_SignData(strSource, 3)
            Else
                strSignData = SHCA_Client.ESE_SignData(strSource, "")
            End If
            If strSignData <> "" And SHCA_Client.errorCode = 0 Then
                 SHCA_Sign = True
            Else
                MsgBoxEx "ǩ��ʧ�ܣ�" & ValidateCertView(SHCA_Client.errorCode)
            End If
        Else
            If mLastPWD = "" Then
                Exit Function
            Else
                MsgBoxEx "ǩ��ʧ�ܣ�", vbInformation, "����ǩ������"
            End If
        End If
        Exit Function
errH:
114     MsgBoxEx "ǩ��ʧ�ܣ�" & vbNewLine & "��" & CStr(Erl()) & "�� " & Err.Description, vbQuestion, gstrSysName
End Function

Public Function SHCA_VerifySign(ByVal strCurrCertSn As String, ByVal strSignData As String, ByVal strSource As String, ByVal strSignCert As String) As Boolean
        '��֤ǩ��
        Dim strTmp As String
        On Error GoTo errH
102     If gudtPara.bytSignVersion = V_SEH Then
            Call SHCA_Client.SEH_InitialSession(2, "", "", 0, 2, "", "") '��ʼ��CA�ӿ�
            Call SHCA_Client.SEH_VerifySignData(strSource, 3, strSignData, strSignCert)
            If SHCA_Client.errorCode <> 0 Then
                '�����ϰ�
                Call SHCA_Client.ESE_InitialSession(2, "", "", 0, 2, "", "") '��ʼ��CA�ӿ�
                Call SHCA_Client.ESE_VerifySignData(strSource, "", strSignData, strSignCert)
            End If
        Else
            Call SHCA_Client.ESE_InitialSession(2, "", "", 0, 2, "", "") '��ʼ��CA�ӿ�
            Call SHCA_Client.ESE_VerifySignData(strSource, "", strSignData, strSignCert)
            If SHCA_Client.errorCode <> 0 Then
                '�����°�
                Call SHCA_Client.SEH_InitialSession(2, "", "", 0, 2, "", "") '��ʼ��CA�ӿ�
                Call SHCA_Client.SEH_VerifySignData(strSource, 3, strSignData, strSignCert)
            End If
        End If
        If SHCA_Client.errorCode = 0 Then
             MsgBoxEx "��֤ǩ���ɹ���"
        Else
             MsgBoxEx "��֤ǩ��ʧ�ܣ�" & ValidateCertView(SHCA_Client.errorCode)
        End If
        Exit Function
errH:
104     MsgBoxEx "��֤ǩ��ʧ�ܣ�" & vbNewLine & "��" & CStr(Erl()) & "�� " & Err.Description, vbQuestion, gstrSysName
End Function

Public Function SHCA_GetPar(Optional ByVal bytFunc As Byte)
    Dim arrList As Variant
    
    On Error GoTo errH
    gstrPara = gobjComLib.zlDatabase.GetPara(90000, glngSys)  '��ȡURLs �̶���ȡZLHIS'gstrPara = "0&&&0"   '0-SEH:1-ESE
    If Val(gstrPara) = 1 Then
        gudtPara.bytSignVersion = V_ESE
    ElseIf Val(gstrPara) = 0 Then
        gudtPara.bytSignVersion = V_SEH
    End If
    SHCA_GetPar = True
    Exit Function
errH:
    MsgBoxEx "��ȡ����ʧ�ܣ�" & vbNewLine & "��" & CStr(Erl()) & "�� " & Err.Description, vbInformation, gstrSysName
End Function

Public Function SHCA_SetParaStr() As String
    SHCA_SetParaStr = IIf(gudtPara.bytSignVersion = 0, "0", "1")
End Function

Public Function SHCA_CheckCert(ByVal strCurrCertSn As String, Optional ByRef strSigCert As String, Optional ByRef blnReDo As Boolean) As Boolean
        '���ܣ���ȡUSB�����豸��ʼ������¼
        Dim strKey As String, strPIN As String, strUserName As String, strCertSn As String, strDate As String
        Dim strWebUrl As String, intDate   As Integer
        Dim blnRet As Boolean
        Dim udtUser As USER_INFO
        Dim intPoint As Integer
        On Error GoTo errH
        If Not SHCA_InitObj() Then
102         MsgBoxEx "����δ��ʼ����"
            Exit Function
        End If
104     If Not GetCertList(strUserName, strKey, strSigCert, strCertSn) Then Exit Function
        intPoint = InStr(strKey, "F")
        If mUserInfo.strUserID = "" Then
            MsgBoxEx "�������֤��Ϊ��,����ϵ����Ա����Ա������¼�룡", vbOKOnly + vbInformation, gstrSysName
            Exit Function
        ElseIf Mid(strKey, intPoint + 2) <> mUserInfo.strUserID Then
            MsgBoxEx "�������֤�ţ�" & _
                       vbCrLf & vbTab & "��" & mUserInfo.strUserID & "��" & vbCrLf & _
                       "��ǰ֤��Ψһ��ʶ:" & _
                       vbCrLf & vbTab & "��" & Mid(strKey, intPoint + 2) & "��" & vbCrLf & _
                       "�û����֤���뵱ǰ֤��Ψһ��ʶ�����,����ʹ�ã�", vbInformation, gstrSysName
            Exit Function
        End If
110     If mLastPWD <> "" Then strPIN = mLastPWD

112     If strPIN = "" Then
114         If Not frmPassword.ShowMe(strPIN) Then Exit Function
        End If
        
116     If Not GetCertLogin(strKey, strPIN, strSigCert, intDate, strWebUrl) Then
118         strPIN = ""
            blnRet = False
        Else
            blnRet = True
        End If
        
        If blnRet Then
            '�ж��Ƿ���Ҫ����ע��֤��
            udtUser.strName = strUserName
            udtUser.strSignName = strUserName
            udtUser.strUserID = Mid(strKey, intPoint + 2) 'SF+���֤��
            udtUser.strCertSn = strCertSn
            udtUser.strCertDN = GetCertDN(strSigCert)
            udtUser.strCert = strSigCert
            udtUser.strEncCert = ""
            udtUser.strCertID = strKey
            '��ȡ�Ѿ�ע��֤�����Ч�������� ���ڸ�ʽ:axBJCASecCOMV21 ����汾���������Ķ���2015/09/15
            If gudtPara.bytSignVersion = V_SEH Then
                strDate = SHCA_Client.SEH_GetCertValidDate(mUserInfo.strCert)
            Else
                strDate = SHCA_Client.ESE_GetCertValidDate(mUserInfo.strCert)
            End If
            If IsUpdateRegCert(udtUser, strDate, blnReDo) Then
                blnRet = True
            Else
                blnRet = False
            End If
        End If
     
        mLastPWD = strPIN
        SHCA_CheckCert = blnRet
        Exit Function
errH:
124     MsgBoxEx "���USBKEYʧ�ܣ�" & vbNewLine & "��" & CStr(Erl()) & "�� " & Err.Description, vbQuestion, gstrSysName
End Function

Public Sub SHCA_UloadObj()
    Set SHCA_Client = Nothing
    mblnInit = False
End Sub
'----- �������ڲ�����

''' ��ȡ�ͻ���֤���б�
''' ����boolean
Private Function GetCertList(ByRef strName As String, ByRef strUniqueID As String, ByRef strCert As String, Optional ByRef strCertSn As String) As Boolean
    '-���:��
    '-����
    'strName :      ����ӿڷ��ص�֤������������
    'strUniqueID:   ����ӿڷ��ص�֤��������Ψһ��ʶ
    'strCert:       ����ӿڷ��ص�ǩ��֤��
    Dim strPassas As String
    On Error GoTo errH
    If gudtPara.bytSignVersion = V_SEH Then
        SHCA_Client.SEH_InitialSession 2, "", "", 0, 2, "", "" '��ʼ��CA�ӿ�
        strCert = SHCA_Client.SEH_GetSelfCertificate(10, "com1", "")
        If SHCA_Client.errorCode <> 0 Then
            '����SM2��RSA���õ���� ��RSA�л���SM2
            SHCA_Client.ESE_InitialSession 2, "", "", 0, 2, "", ""
            strCert = SHCA_Client.ESE_GetSelfCertificate(36, "com1")
            If SHCA_Client.errorCode = 0 Then
                gudtPara.bytSignVersion = V_ESE
                GoTo LineESE
            Else
                MsgBoxEx ValidateCertView(SHCA_Client.errorCode)
            End If
            Exit Function
        End If
LineSEH:
        strName = SHCA_Client.SEH_GetCertDetail(strCert, 17)
        If SHCA_Client.errorCode <> 0 Then
            MsgBoxEx ValidateCertView(SHCA_Client.errorCode)
            Exit Function
        End If
        
        strUniqueID = SHCA_Client.SEH_GetCertInfoByOID(strCert, "1.2.156.112570.148")
        If SHCA_Client.errorCode <> 0 Then
            MsgBoxEx ValidateCertView(SHCA_Client.errorCode)
            Exit Function
        End If
        
        strCertSn = SHCA_Client.SEH_GetCertDetail(strCert, 2)
        If SHCA_Client.errorCode <> 0 Then
            MsgBoxEx ValidateCertView(SHCA_Client.errorCode)
            Exit Function
        End If
    ElseIf gudtPara.bytSignVersion = V_ESE Then
        SHCA_Client.ESE_InitialSession 2, "", "", 0, 2, "", ""
        strCert = SHCA_Client.ESE_GetSelfCertificate(36, "com1")
        If SHCA_Client.errorCode <> 0 Then
            '����SM2��RSA���õ����
            SHCA_Client.SEH_InitialSession 2, "", "", 0, 2, "", "" '��ʼ��CA�ӿ�
            strCert = SHCA_Client.SEH_GetSelfCertificate(10, "com1", "")
            If SHCA_Client.errorCode = 0 Then
                gudtPara.bytSignVersion = V_SEH
                GoTo LineSEH
            Else
                MsgBoxEx ValidateCertView(SHCA_Client.errorCode)
            End If
            Exit Function
        End If
LineESE:
        strName = SHCA_Client.ESE_GetCertDetail(strCert, 17)
        If SHCA_Client.errorCode <> 0 Then
            MsgBoxEx ValidateCertView(SHCA_Client.errorCode)
            Exit Function
        End If
        
        strUniqueID = SHCA_Client.ESE_GetCertInfoByOID(strCert, "1.2.156.112570.148")
        If SHCA_Client.errorCode <> 0 Then
            MsgBoxEx ValidateCertView(SHCA_Client.errorCode)
            Exit Function
        End If
        
        strCertSn = SHCA_Client.ESE_GetCertDetail(strCert, 2)
        If SHCA_Client.errorCode <> 0 Then
            MsgBoxEx ValidateCertView(SHCA_Client.errorCode)
            Exit Function
        End If
    End If
    GetCertList = True
    Exit Function
errH:
    GetCertList = False
End Function

Private Function GetCertLogin(ByVal strUniqueID As String, ByVal strPassword As String, ByVal strCert As String, ByRef dDate As Integer, ByRef strWebserviceUrl As String) As Boolean
    '- ���
    'strUniqueID : ֤��Ψһ��ʶ
    'strPassword : ֤������
    'strWebserviceUrl:ǩ����������ַ����Ϊ֤����֤
    '- ����
    'dDate       : ����֤����Чʱ��
    On Error GoTo errH
    Dim result As Boolean
    If SHCA_Client Is Nothing Then Set SHCA_Client = CreateObject("SafeEngineCOM.SafeEngineCtl")
    If (strPassword = "") Then
        MsgBoxEx "������֤�����룡"
    Else
        '֤�鰲ȫ��¼
        'result:0:�ɹ�
        'result:��0:���ɹ�
        If mLogin >= 8 Then
            MsgBoxEx "�Ѿ�������" & mLogin & "�δ������룬������������������"
            Exit Function
        End If
        If gudtPara.bytSignVersion = V_SEH Then
            Call SHCA_Client.SEH_InitialSession(27, "com1", strPassword, 0, 27, "com1", "") '��ʼ��CA�ӿ�(����)
        Else
            Call SHCA_Client.ESE_InitialSession(36, "com1", strPassword, 0, 36, "com1", "") '��ʼ��CA�ӿ�(����)
        End If
        If SHCA_Client.errorCode = 0 Then
             '��֤֤������Ϣ��ʾ
            If gudtPara.bytSignVersion = V_SEH Then
                Call SHCA_Client.SEH_VerifyCertificate(strCert)
            Else
                Call SHCA_Client.ESE_VerifyCertificate(strCert)
            End If
            If SHCA_Client.errorCode = 0 Then
                
                '��ȡ�ͻ���֤����Ч�ڽ�ֹʱ��
                If gudtPara.bytSignVersion = V_SEH Then
                    dDate = SHCA_Client.SEH_GetCertValidDate(strCert)
                Else
                    dDate = SHCA_Client.ESE_GetCertValidDate(strCert)
                End If
                If (dDate <= 30 And dDate > 0) And Not gblnShow Then
                    MsgBoxEx "����֤�黹��" & dDate & "�����"
                    gblnShow = True
                    GetCertLogin = True
                ElseIf (dDate <= 0) Then
                    MsgBoxEx "����֤���ѹ��� " & Abs(dDate) & " ��"
                    GetCertLogin = False
                Else
                    GetCertLogin = True
                End If
            Else
               MsgBoxEx "��֤֤�����" & ValidateCertView(SHCA_Client.errorCode)
            End If
        Else
            mLogin = mLogin + 1
            MsgBoxEx "��ʼ��½����" & ValidateCertView(SHCA_Client.errorCode)
        End If
       
    End If
    Exit Function
errH:
    mLogin = mLogin + 1
    MsgBoxEx "֤��������ܲ���ȷ�����Ѿ�������" & mLogin & "�����룬����������" & 8 - mLogin & "��!"
    GetCertLogin = False
End Function

''' <summary>
''' ��֤֤������Ϣ��ʾ
''' </summary>
''' <remarks></remarks>
Private Function ValidateCertView(retValidateCert) As String
    Dim strErrorMsg As String
    Select Case retValidateCert
        Case 0
            strErrorMsg = ""
        Case -2113667072:
            strErrorMsg = "װ�ض�̬�����(-2113667072)"
            
        Case -2113667071:
            strErrorMsg = "�ڴ�������(-2113667071)"
            
        Case -2113601536:
            strErrorMsg = "��ȡ�ļ�����(-2113601536)"
            
        Case -2113601535:
            strErrorMsg = "�������(-2113601535)"
            
        Case -2113601534:
            strErrorMsg = "�Ƿ��������(-2113601534)"
            
        Case -2113601533:
            strErrorMsg = "ȱ��ECC KEY����(-2113601533)"
        
        Case -2113601532:
            strErrorMsg = "ECC KEY ���㷨��ƥ�����(-2113601532)"
            
        Case -2113601531:
            strErrorMsg = "�Ƿ��㷨����(-2113601531)"
            
        Case -2113601530:
            strErrorMsg = "����ǩ������(-2113601530)"
            
        Case -2113601529:
            strErrorMsg = "ժҪ����(-2113601529)"
            
        Case -2113601528:
            strErrorMsg = "������̫С(-2113601528)"
            
        Case -2113601527:
            strErrorMsg = "֤���ʽ����(-2113601527)"
            
        Case -2113601526:
            strErrorMsg = "ȱ�ٹ�Կ����(-2113601526)"
            
        Case -2113601525:
            strErrorMsg = "��֤ǩ������(-2113601525)"
            
        Case -2113601524:
            strErrorMsg = "������˽Կ�Դ���(-2113601524)"
            
        Case -2113601523:
            strErrorMsg = "PKCS12�������(-2113601523)"
            
        Case -2113601522:
            strErrorMsg = "PKCS12��ʽ����(-2113601522)"
            
        Case -2113601520:
            strErrorMsg = "SE_ECC_ERROR_LOAD_BUILTIN_EC(-2113601520)"
            
        Case -2113601519:
            strErrorMsg = "��˽Կ��ƥ�����(-2113601519)"
            
        Case -2113601518:
            strErrorMsg = "PKCS10�������(-2113601518)"
            
        Case -2113601517:
            strErrorMsg = "PKCS10�������(-2113601517)"
            
        Case -2113601516:
            strErrorMsg = "��Կ��ʽ����(-2113601516)"
            
        Case -2113601515:
            strErrorMsg = "PKCS10��ʽ����(-2113601515)"
            
        Case -2113601514:
            strErrorMsg = "��֤PKCS10����(-2113601514)"
            
        Case -2113601513:
            strErrorMsg = "ECC KEY��ʽ����(-2113601513)"
            
        Case -2113601512:
            strErrorMsg = "��Կ�������(-2113601512)"
            
        Case -2113601511:
            strErrorMsg = "ǩ����ʽ����(-2113601511)"
            
        Case -2113601510:
            strErrorMsg = "EC��ʽ����(-2113601510)"
            
        Case -2113601509:
            strErrorMsg = "ECC KEY�������(-2113601509)"
            
        Case -2113601508:
            strErrorMsg = "д���ļ�����(-2113601508)"
            
        Case -2113601507:
            strErrorMsg = "֤�����Ƿ�����(-2113601507)"
            
        Case -2113601506:
            strErrorMsg = "�ڴ�������(-2113601506)"
            
        Case -2113601505:
            strErrorMsg = "��ʼ����������(-2113601505)"
            
        Case -2113601504:
            strErrorMsg = "��ȡ�����ļ�����(-2113601504)"
            
        Case -2113601503:
            strErrorMsg = "���豸����(-2113601503)"
            
        Case -2113601502:
            strErrorMsg = "�򿪻Ự����(-2113601502)"
            
        Case -2113601501:
            strErrorMsg = "װ�ض�̬�����(-2113601501)"
            
        Case -2113601500:
            strErrorMsg = "�豸���ʹ���(-2113601500)"
         
        Case -2113601499:
            strErrorMsg = "�㷨��֧�ִ���(-2113601499)"
            
        Case -2113601498:
            strErrorMsg = "����PKCS10����(-2113601498)"
            
        Case -2113601497:
            strErrorMsg = "������Կ����(-2113601497)"
            
        Case -2113601496:
            strErrorMsg = "EC_POINT�Ƿ�����(-2113601496)"
    
        Case -2113601495:
            strErrorMsg = "�ԳƼ��ܴ���(-2113601495)"
            
        Case -2113601494:
            strErrorMsg = "�Գƽ��ܴ���(-2113601494)"
            
        Case -2113601493:
            strErrorMsg = "PEM�������(-2113601493)"
            
        Case -2113601492:
            strErrorMsg = "��ȡ֤��ϸĿ����(-2113601492)"
            
        Case -2113601491:
            strErrorMsg = "PEM�������(-2113601491)"
            
        Case -2113601490:
            strErrorMsg = "��ȡ֤����չ�����(-2113601490)"
            
        Case -2113601489:
            strErrorMsg = "�Ƿ��ӿ����ʹ���(-2113601489)"
            
        Case -2113601488:
            strErrorMsg = "�Ƿ���������(-2113601488)"
            
        Case -2113601487:
            strErrorMsg = "ö���豸����(-2113601487)"
            
        Case -2113601486:
            strErrorMsg = "û���豸(-2113601486)"
            
        Case -2113601485:
            strErrorMsg = "�豸���Ӵ���(-2113601485)"
            
        Case -2113601484:
            strErrorMsg = "�������������(-2113601484)"
            
        Case -2113601483:
            strErrorMsg = "SE_ECC_ERROR_SKF_SET_SYMKEY(-2113601483)"
            
        Case -2113601482:
            strErrorMsg = "�ԳƼ��ܳ�ʼ������(-2113601482)"
            
        Case -2113601481:
            strErrorMsg = "�ԳƼ��ܴ���(-2113601481)"
            
        Case -2113601480:
            strErrorMsg = "�豸����Ա�������(-2113601480)"
            
        Case -2113601479:
            strErrorMsg = "��Ӧ�ô���(-2113601479)"
            
        Case -2113601478:
            strErrorMsg = "�豸����(-2113601478)"
            
        Case -2113601477:
            strErrorMsg = "�豸�������(-2113601477)"
            
        Case -2113601476:
            strErrorMsg = "ö��Ӧ�ô���(-2113601476)"
            
        Case -2113601475:
            strErrorMsg = "ɾ��Ӧ�ô���(-2113601475)"
            
        Case -2113601474:
            strErrorMsg = "����Ӧ�ô���(-2113601474)"
            
        Case -2113601473:
            strErrorMsg = "������������(-2113601473)"
            
        Case -2113601472:
            strErrorMsg = "�豸��֧�ִ���(-2113601472)"
            
        Case -2113601471:
            strErrorMsg = "����������(-2113601471)"
            
        Case -2113601470:
            strErrorMsg = "������Կ����(-2113601470)"
            
        Case -2113601466:
            strErrorMsg = "�ԳƼ��ܴ���(-2113601466)"
            
        Case -2113601465:
            strErrorMsg = "������Կ�Դ���(-2113601465)"
            
        Case -2113601464:
            strErrorMsg = "�޸��豸�������(-2113601464)"
            
        Case -2113601463:
            strErrorMsg = "����֤�����(-2113601463)"
            
        Case -2113601462:
            strErrorMsg = "����֤�����(-2113601462)"
            
        Case -2113601461:
            strErrorMsg = "�����ļ�����(-2113601461)"
            
        Case -2113601460:
            strErrorMsg = "д���ļ�����(-2113601460)"
            
        Case -2113601459:
            strErrorMsg = "��ȡ�ļ���Ϣ����(-2113601459)"
            
        Case -2113601458:
            strErrorMsg = "��ȡ�ļ�����(-2113601458)"
            
        Case -2113601457:
            strErrorMsg = "��ȡ��Կ����(-2113601457)"
            
        Case -2113601454:
            strErrorMsg = "������Կ�Դ���(-2113601454)"
            
        Case -2113601453:
            strErrorMsg = "֤���ѹ���(-2113601453)"
            
        Case -2113601452:
            strErrorMsg = "����豸����(-2113601452)"
            
        Case -2113601451:
            strErrorMsg = "û���豸(-2113601451)"
            
        Case -2113601450:
            strErrorMsg = "�Զ�����豸����(-2113601450)"
            
        Case -2113601449:
            strErrorMsg = "�豸�޷�ʶ��(-2113601449)"
            
        Case -2113601448:
            strErrorMsg = "��ȡ�Ự��Կ��(-2113601448)"
            
        Case -2113601447:
            strErrorMsg = "����Ự��Կ��(-2113601447)"
            
        Case -2113601446:
            strErrorMsg = "��ʼ��ժҪ����(-2113601446)"
            
        Case -2113601445:
            strErrorMsg = "����ժҪ����(-2113601445)"
            
        Case -2113601444:
            strErrorMsg = "���ɻỰ��Կ��(-2113601444)"
            
        Case -2113601442:
            strErrorMsg = "����Ự��Կ��(-2113601442)"
            
        Case -2113601441:
            strErrorMsg = "������̫С(-2113601441)"
            
        Case -2113601440:
            strErrorMsg = "P7ǩ�����ݳ�ʼ������(-2113601440)"
            
        Case -2113601439:
            strErrorMsg = "�������������(-2113601439)"
            
        Case -2113601438:
            strErrorMsg = "�ԳƼ��ܴ���(-2113601438)"
            
        Case -2113601437:
            strErrorMsg = "�Գƽ��ܴ���(-2113601437)"
            
        Case -2113601436:
            strErrorMsg = "������Կ����(-2113601436)"
            
        Case -2113601435:
            strErrorMsg = "���p7�㷨����(-2113601435)"
            
        Case -2113601434:
            strErrorMsg = "P7���ݴ������(-2113601434)"
            
        Case -2113601433:
            strErrorMsg = "SE_ECC_ERROR_ENVELOPE_ADD_RECIP(-2113601433)"
            
        Case -2113601432:
            strErrorMsg = "ǩ�����ݴ���(-2113601432)"
            
        Case -2113601431:
            strErrorMsg = "ժҪ���ݴ������(-2113601431)"
            
        Case -2113601430:
            strErrorMsg = "���ܸ��´���(-2113601430)"
            
        Case -2113601429:
            strErrorMsg = "���ܴ������(-2113601429)"
            
        Case -2113601428:
            strErrorMsg = "���ܳ�ʼ������(-2113601428)"
            
        Case -2113601427:
            strErrorMsg = "���ܸ��´���(-2113601427)"
            
        Case -2113601426:
            strErrorMsg = "���ܴ������(-2113601426)"
            
        Case -2113601425:
            strErrorMsg = "p7��ʽ����(-2113601425)"
            
        Case -2113601424:
            strErrorMsg = "SE_ECC_ERROR_P7_NO_RECIP(-2113601424)"
            
        Case -2113601423:
            strErrorMsg = "�㷨�Ƿ�(-2113601423)"
            
        Case -2113601422:
            strErrorMsg = "˽Կ���ȴ���(-2113601422)"
            
        Case -2113601421:
            strErrorMsg = "P7ǩ������(-2113601421)"
            
        Case -2113601420:
            strErrorMsg = "��֤P7ǩ������(-2113601420)"
            
        Case -2113601419:
            strErrorMsg = "P7ǩ�����ð汾����(-2113601419)"
            
        Case -2113601418:
            strErrorMsg = "���豸����(-2113601418)"
            
        Case -2113601417:
            strErrorMsg = "������̫С(-2113601417)"
            
        Case -2113601416:
            strErrorMsg = "��LDAP��ȡ֤�����(-2113601416)"
            
        Case -2113601415:
            strErrorMsg = "����OCSP����������(-2113601415)"
            
        Case -2113601414:
            strErrorMsg = "��������(-2113601414)"
            
        Case -2113601413:
            strErrorMsg = "CRL��ʽ����(-2113601413)"
            
        Case -2113601412:
            strErrorMsg = "֤��ϳ�(-2113601412)"
            
        Case -2113601411:
            strErrorMsg = "֤������ʽ����(-2113601411)"
            
        Case -2113601410:
            strErrorMsg = "��֤֤��������(-2113601410)"
            
        Case -2113601409:
            strErrorMsg = "����Ա�������(-2113601409)"
            
        Case -2113601408:
            strErrorMsg = "�豸��ǩ��ʽ����(-2113601408)"
            
        Case -2113601407:
            strErrorMsg = "ɾ����������(-2113601407)"
            
        Case -2113601406:
            strErrorMsg = "ö���ļ�����(-2113601406)"
            
        Case -2113601405:
            strErrorMsg = "ɾ���ļ�����(-2113601405)"
            
        Case -2113601404:
            strErrorMsg = "ö����������(-2113601404)"
            
        Case -2113601403:
            strErrorMsg = "�ر�Ӧ�ô���(-2113601403)"
        
        Case -2113568768:
            strErrorMsg = "SE_ECC_ERROR_FUNC_LOCAL(-2113568768)"
            
        Case -2113667070:
            strErrorMsg = "��˽Կ�豸����(-2113667070)"
            
        Case -2113667069:
            strErrorMsg = "˽Կ�������(-2113667069)"
            
        Case -2113667068:
            strErrorMsg = "��֤�����豸����(-2113667068)"
            
        Case -2113667067:
            strErrorMsg = "֤�����������(-2113667067)"
            
        Case -2113667066:
            strErrorMsg = "��֤���豸����(-2113667066)"
            
        Case -2113667065:
            strErrorMsg = "֤���������(-2113667065)"
            
        Case -2113667064:
            strErrorMsg = "˽Կ��ʱ(-2113667064)"
            
        Case -2113667063:
            strErrorMsg = "������̫С(-2113667063)"
            
        Case -2113667062:
            strErrorMsg = "��ʼ����������(-2113667062)"
            
        Case -2113667061:
            strErrorMsg = "�����������(-2113667061)"
            
        Case -2113667060:
            strErrorMsg = "����ǩ������(-2113667060)"
            
        Case -2113667059:
            strErrorMsg = "��֤ǩ������(-2113667059)"
            
        Case -2113667058:
            strErrorMsg = "ժҪ����(-2113667058)"
            
        Case -2113667057:
            strErrorMsg = "֤���ʽ����(-2113667057)"
            
        Case -2113667056:
            strErrorMsg = "�����ŷ����(-2113667056)"
            
        Case -2113667055:
            strErrorMsg = "��LDAP��ȡ֤�����(-2113667055)"
            
        Case -2113667054:
            strErrorMsg = "֤���ѹ���(-2113667054)"
            
        Case -2113667053:
            strErrorMsg = "��ȡ֤��������(-2113667053)"
            
        Case -2113667052:
            strErrorMsg = "֤������ʽ����(-2113667052)"
            
        Case -2113667051:
            strErrorMsg = "��֤֤��������(-2113667051)"
            
        Case -2113667050:
            strErrorMsg = "֤���ѷϳ�(-2113667050)"
            
        Case -2113667049:
            strErrorMsg = "CRL��ʽ����(-2113667049)"
            
        Case -2113667048:
            strErrorMsg = "����OCSP����������(-2113667048)"
            
        Case -2113667047:
            strErrorMsg = "OCSP����������(-2113667047)"
            
        Case -2113667046:
            strErrorMsg = "OCSP�ذ�����(-2113667046)"
            
        Case -2113667045:
            strErrorMsg = "OCSP�ذ���ʽ����(-2113667045)"
            
        Case -2113667044:
            strErrorMsg = "OCSP�ذ�����(-2113667044)"
            
        Case -2113667043:
            strErrorMsg = "OCSP�ذ���֤ǩ������(-2113667043)"
            
        Case -2113667042:
            strErrorMsg = "֤��״̬δ֪(-2113667042)"
            
        Case -2113667041:
            strErrorMsg = "�ԳƼӽ��ܴ���(-2113667041)"
            
        Case -2113667040:
            strErrorMsg = "��ȡ֤����Ϣ����(-2113667040)"
            
        Case -2113667039:
            strErrorMsg = "��ȡ֤��ϸĿ����(-2113667039)"
            
        Case -2113667038:
            strErrorMsg = "��ȡ֤��Ψһ��ʶ����(-2113667038)"
            
        Case -2113667037:
            strErrorMsg = "��ȡ֤����չ�����(-2113667037)"
            
        Case -2113667036:
            strErrorMsg = "PEM�������(-2113667036)"
            
        Case -2113667035:
            strErrorMsg = "PEM�������(-2113667035)"
            
        Case -2113667034:
            strErrorMsg = "�������������(-2113667034)"
            
        Case -2113667033:
            strErrorMsg = "PKCS12��������(-2113667033)"
            
        Case -2113667032:
            strErrorMsg = "˽Կ��ʽ����(-2113667032)"
            
        Case -2113667031:
            strErrorMsg = "��˽Կ��ƥ��(-2113667031)"
            
        Case -2113667030:
            strErrorMsg = "PKCS12�������(-2113667030)"
            
        Case -2113667029:
            strErrorMsg = "PKCS12��ʽ����(-2113667029)"
            
        Case -2113667028:
            strErrorMsg = "PKCS12�������(-2113667028)"
            
        Case -2113667027:
            strErrorMsg = "�ǶԳƼӽ��ܴ���(-2113667027)"
            
        Case -2113667026:
            strErrorMsg = "OID��ʽ����(-2113667026)"
            
        Case -2113667025:
            strErrorMsg = "LDAP��ַ��ʽ����(-2113667025)"
            
        Case -2113667024:
            strErrorMsg = "LDAP��ַ����(-2113667024)"
            
        Case -2113667023:
            strErrorMsg = "����LDAP����������(-2113667023)"

        Case -2113667022:
            strErrorMsg = "LDAP�󶨴���(-2113667022)"
            
        Case -2113667021:
            strErrorMsg = "û��OID��Ӧ����չ��(-2113667021)"
            
        Case -2113667020:
            strErrorMsg = "��ȡ֤�鼶�����(-2113667020)"
            
        Case -2113667019:
            strErrorMsg = "��ȡ�����ļ�����(-2113667019)"
            
        Case -2113667018:
            strErrorMsg = "˽Կδ����(-2113667018)"
            
  ' ���´������ڵ�¼
        Case -2113666824:
            strErrorMsg = "��Ч�ĵ�¼ƾ֤(-2113666824)"
            
        Case -2113666823:
            strErrorMsg = "��������(-2113666823)"
            
        Case -2113666822:
            strErrorMsg = "���Ƿ�����֤��(-2113666822)"
            
        Case -2113666821:
            strErrorMsg = "��¼����(-2113666821)"
            
        Case -2113666820:
            strErrorMsg = "֤����֤��ʽ����(-2113666820)"
            
        Case -2113666819:
            strErrorMsg = "�������֤����(-2113666819)"
            
        Case -2113666818:
            strErrorMsg = "�뵥���¼�ͻ��˴���ͨ��(-2113666818)"
    End Select
    ValidateCertView = strErrorMsg
End Function





