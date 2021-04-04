Attribute VB_Name = "mdlXJCA"
Option Explicit
'�½�CA���Ĺ���ģ��

Private mblnInit As Boolean         '�Ƿ��ѳ�ʼ���ɹ�
Private mobjXJCA_Client As Object
Private mstrPara  As String
Private mobjGseal As Object
Private mbytType As Byte '1-����������ҽԺʹ�õ��Ǻ�̩KEY
                         '2-��̨ҽԺʹ�õ��ǻ���KEY
'
Private Const M_STR_CSP  As String = "HaiTai Cryptographic Service Provider for xjca"
Private Const M_STR_CSP_HD  As String = "CIDC Cryptographic Service Provider v1.0.0"

'����ǩ�½ӿ�����(C++)
'Private Declare Function XJCA_SignSeal Lib "XJCA_HOS.dll" (ByVal strSrc As String, ByVal lngxml As Long, ByVal lngLen As Long) As Boolean
'˵����bool   XJCA_SignSeal(char* src,char* signxml, DWORD* len)
'����˵��:strSrc-����Դ,lngXml--����ַ����byte��������ַ���,��chr����ת����,lngLen-����ַ ��Ϊ���˵��ǵ�ַ,������long��
'����ֵ��True\false
Private Declare Function XJCA_GetSealBMPB Lib "XJCA_HOS.dll" (ByVal strFilePath As String, ByVal lngTimes As Long) As Boolean
'����:��ȡǩ��ͼƬ
'����:
Private Declare Function XJCA_VerifySeal Lib "XJCA_HOS.dll" (ByVal strSrc As String, ByVal strxml As String, ByVal strPic As String, ByVal strCert As String) As Boolean
'����:��֤ǩ������.�ӿ�ԭ��  bool   XJCA_VerifySeal((char* src,char* xml,char* pct,char* cert)
'������
'���أ�True\false


Public Function XJCA_InitObj() As Boolean
'����:֤�鲿����ʼ��
    Dim strUrl As String

102     XJCA_InitObj = mblnInit
104     If mblnInit Then Exit Function
        On Error Resume Next
        mstrPara = gobjComLib.zlDatabase.GetPara(90000, glngSys)  '��ȡ��������
        If mstrPara = "" Then
            Err.Raise -1, , "�����ļ���ȡʧ�ܣ��������á�"
            Exit Function
        End If
        On Error GoTo 0: Err.Clear
        
        On Error GoTo errH
        If InStr(mstrPara, "����") = 0 Then
            mbytType = 1
            'ȱʡ��̩ KEY
108         Set mobjXJCA_Client = CreateObject("xjcaTechATL.xjcaTechATLLib.1")
            Set mobjGseal = CreateObject("Signature.SignatureForm")      '����SignatureForm�ؼ�,����ʹ��XJCA_HOS.dll�е�ǩ�º��� �����ڴ�����
            If mobjXJCA_Client Is Nothing Or mobjGseal Is Nothing Then
                MsgBoxEx "CA���󴴽�ʧ�ܣ�", vbOKOnly, gstrSysName
                Exit Function
            End If
        Else
            mbytType = 2
            '����KEY
            Set mobjXJCA_Client = CreateObject("XjcaFgwATL.XjcaFgwATLLib.1")
            Set mobjGseal = CreateObject("XJFormSeal.XJFormSealX")
            If mobjXJCA_Client Is Nothing Or mobjGseal Is Nothing Then
                MsgBoxEx "CA���󴴽�ʧ�ܣ�", vbOKOnly, gstrSysName
                Exit Function
            End If
        End If
114     XJCA_InitObj = True
        
116     mblnInit = XJCA_InitObj
        Exit Function
errH:
118     MsgBoxEx "�����½�CA�ӿڲ���ʧ�ܣ�" & vbNewLine & Err.Description, vbQuestion, gstrSysName

End Function

Public Function XJCA_RegCert(arrCertInfo As Variant) As Boolean
        '���ܣ��ṩ��HIS���ݿ���ע�����֤��ı�Ҫ��Ϣ,�����·��Ż����֤��,,��Ҫ����USB-Key
        '���أ�arrCertInfo��Ϊ���鷵��֤�������Ϣ
        '      0-ClientSignCertCN:�ͻ���ǩ��֤�鹫������(����),����ע��֤��ʱ������֤���
        '      1-ClientSignCertDN:�ͻ���ǩ��֤������(ÿ��Ψһ)
        '      2-ClientSignCertSN:�ͻ���ǩ��֤�����к�(ÿ֤��Ψһ)
        '      3-ClientSignCert:�ͻ���ǩ��֤������
        '      4-ClientEncCert:�ͻ��˼���֤������
        '      5-ǩ��ͼƬ�ļ���,�մ���ʾû��ǩ��ͼƬ
        Dim strKeyId As String, strCertTime As String, strCertUserName As String, strCertDN As String
        Dim strSigCert As String
        Dim strFile As String
        Dim blnRet As Boolean
        Dim i As Long
        On Error GoTo errH
        If Not CheckIsXJCA Then Exit Function
100     For i = LBound(arrCertInfo) To UBound(arrCertInfo)
101         arrCertInfo(i) = ""
102     Next
       
108     If XJCA_GetCertList(strCertUserName, strKeyId, strSigCert, strCertDN) Then
200         arrCertInfo(0) = strCertUserName
201         arrCertInfo(1) = strCertDN
202         arrCertInfo(2) = strKeyId
203         arrCertInfo(3) = strSigCert
205         arrCertInfo(4) = ""
            strFile = App.Path & "\" & Format(Now, "yyyyMMdd") & "_" & strKeyId & ".BMP"
            blnRet = XJCA_GetSealBMPB(strFile, 2)
            If blnRet = False Then Exit Function
206         arrCertInfo(5) = strFile
            XJCA_RegCert = True
        End If

300     Exit Function

errH:
    MsgBoxEx "��ȡ֤����Ϣʧ�ܣ�" & vbNewLine & "��" & CStr(Erl()) & "�� " & Err.Description, vbQuestion, gstrSysName
End Function

Public Function XJCA_CheckCert(ByVal strCurrCertSn As String, Optional ByRef strSigCert As String) As Boolean
'���ܣ���ȡUSB�����豸��ʼ������¼
'����ֵ:
'  strSigCert -ǩ��֤������

        Dim strKey As String
        Dim strUserName As String
        Dim strCertDN As String
        
        Dim lngRet As Long
        
        On Error GoTo errH
        
        If Not XJCA_InitObj() Then
             MsgBoxEx "����δ��ʼ����", vbInformation + vbOKOnly, gstrSysName
            Exit Function
        End If
        
100     If Not CheckIsXJCA Then Exit Function
   
104     If XJCA_GetCertList(strUserName, strKey) Then
106        If strCurrCertSn <> strKey Then
108            MsgBoxEx "��֤��δע�����������£�����ʹ�ã�"
               Exit Function
           End If
110
116        If Not GetCertLogin(strKey, strUserName) Then
                Exit Function
           End If
122     End If
        
        XJCA_CheckCert = True
        Exit Function
errH:
124     MsgBoxEx "���USBKEYʧ�ܣ�" & vbNewLine & "��" & CStr(Erl()) & "�� " & Err.Description, vbQuestion, gstrSysName
End Function

Public Function XJCA_GetCertList(Optional ByRef strName As String = "1", Optional ByRef strUniqueID As String = "1", Optional ByRef strCert As String = "1", Optional ByRef strCertDN As String = "1") As Boolean
    '-���:��
    '-����
    'strName :      ����ӿڷ��ص�֤������������
    'strUniqueID:   ����ӿڷ��ص�֤��������Ψһ��ʶ
    'strCert:       ����ӿڷ��ص�ǩ��֤��
    On Error GoTo errH
    Dim strSrc As String
    Dim strTmp As String
    Dim blnRet As Boolean
    Dim arrTmp As Variant
    Dim i As Long
    
    On Error GoTo errH
    If mbytType = 1 Then
        If strUniqueID <> "1" Then
            frmXJCA.txtValue.Text = CStr(mobjXJCA_Client.XJCA_GetCertSN)   '֤�����
            strUniqueID = Trim(frmXJCA.txtValue.Text)
        End If
        If strName <> "1" Or strCertDN <> "1" Then
            arrTmp = Array()
            ReDim Preserve arrTmp(UBound(arrTmp) + 1)
            arrTmp(UBound(arrTmp)) = CStr(mobjXJCA_Client.XJCA_GetCertDN())     'C=CN, S=650105197001010026, L=0026, O=�½�CA, OU=CA����, E=xjcaxmss@xjca.com.cn, CN=�½�CA�ӿڲ���0026
            strTmp = arrTmp(UBound(arrTmp))
            strCertDN = strTmp
            strName = Mid(strCertDN, InStr(strCertDN, "CN=") + 3)   '��ȡ֤�����������
        End If
        
        If strCert <> "1" Then
            strSrc = "1234567890"
            Call mobjGseal.XJCASetFieldByName("IsNeedCert", "true") '����ǩ�½ӿ�ǰ�ȵ�����,���򱨴�,XJCASowSignInSvr
            Call mobjGseal.XJCASowSignInSvr(strSrc, strTmp)
            If strTmp <> "" Then
                strCert = Split(strTmp, ",")(1) & "," & Split(strTmp, ",")(2) '֤����Ϣ '֤��ID
            Else
                Exit Function
            End If
        End If
    Else
        '����KEY
        If strUniqueID <> "1" Then
            strUniqueID = CStr(mobjXJCA_Client.XJCA_GetCertSN(M_STR_CSP_HD))   '֤�����
        End If
        If strName <> "1" Or strCertDN <> "1" Then
            arrTmp = Array()
            ReDim Preserve arrTmp(UBound(arrTmp) + 1)
            arrTmp(UBound(arrTmp)) = CStr(mobjXJCA_Client.XJCA_GetCertDN(M_STR_CSP_HD))      'C=CN, S=650105197001010026, L=0026, O=�½�CA, OU=CA����, E=xjcaxmss@xjca.com.cn, CN=�½�CA�ӿڲ���0026
            strTmp = arrTmp(UBound(arrTmp))
            strCertDN = strTmp
            strName = Mid(strCertDN, InStr(strCertDN, "CN=") + 3)   '��ȡ֤�����������
        End If
        
        If strCert <> "1" Then
            strSrc = "1234567890"
            strTmp = ""
            Call mobjGseal.XJCASowSignInSvr(strSrc, strTmp)
            If strTmp <> "" Then
                strCert = Split(strTmp, ",")(1) & "," & Split(strTmp, ",")(2) '֤����Ϣ '֤��ID
            Else
                Exit Function
            End If
        End If
    End If
    XJCA_GetCertList = True
    Exit Function
errH:
    MsgBoxEx "��ȡ֤����Ϣʧ�ܣ�" & vbNewLine & "��" & CStr(Erl()) & "�� " & Err.Description, vbQuestion, gstrSysName

End Function

Public Function XJCA_Sign(ByVal strCurrCertSn As String, ByVal strSource As String, ByRef strSignData As String) As Boolean
'����:�½�CA����ǩ��
'������strCurrCertSn  -֤��ID(Ψһ����)
'     strSource-��Ҫǩ����Դ����
'     strTimeStamp-ʱ���
'     strTimeStampCode-ʱ�����Ϣ
'����ֵ��true �ɹ�,False -ʧ��
'       strSignData-ǩ���󷵻ص�ǩ������
'       strTimeStamp-���ص�ʱ���

        Dim strTmp As String
        Dim bytXml(40000)  As Byte    'ǩ����Ϣ
        Dim lngLen As Long
        Dim i As Long
        Dim J As Long
        
        Dim blnRet As Boolean
        
        On Error GoTo errH
        
100     If XJCA_CheckCert(strCurrCertSn) Then                '��֤��ǰUSB�Ƿ���ǩ���û��ģ�����ȡǩ��֤��
            '�ո�vbTAb,vbCrLF ����ǩ���ӿ�ʱ��ͳһ����Դ���ص�ǩ��ֵ�п��ܲ�һ��
'            strSource = Replace(strSource, " ", "")
'            strSource = Replace(strSource, vbTab, "|")
'            strSource = Replace(strSource, vbCrLf, "||")
            If mbytType = 1 Then
                Call mobjGseal.XJCASetFieldByName("IsNeedCert", "true") '����ǩ�½ӿ�ǰ�ȵ�����,���򱨴�,XJCASowSignInSvr
                Call mobjGseal.XJCASowSignInSvr(strSource, strTmp)
            Else
                Call mobjGseal.XJCASowSignInSvr(strSource, strTmp)
            End If
            If strTmp <> "" Then
                strSignData = Split(strTmp, ",")(0) 'ǩ������
            Else
                MsgBoxEx "ǩ��ʧ�ܣ�": Exit Function
            End If
        End If
 
        XJCA_Sign = True
        Exit Function
errH:
114     MsgBoxEx "ǩ��ʧ�ܣ�" & vbNewLine & "��" & CStr(Erl()) & "�� " & Err.Description, vbQuestion, gstrSysName
End Function


Private Function GetCertLogin(ByVal strUniqueID As String, ByVal strName As String) As Boolean
    'strUniqueID : ֤��Ψһ��ʶ
    'strWebURL��
    
    Dim strTmp As String
    Dim strWebUrl As String
    Dim strAppId As String
         
'    lngResult = mobjXJCA_Client.XJCA_VerifyPin(strPassword, Len(strPassword))

    '����������֤֤��
'   mstrPara = "http://124.117.245.71:18080/webServices/authService|4028e48a39dd529a0139dd5c383d0010"
'   mstrPara =http://124.117.245.71:48080/webServices/ssoService|4028f6d24a2d7182014a2d83333e001a|����   ����KEY
    On Error Resume Next
    strWebUrl = Split(mstrPara, "|")(0)
    strAppId = Split(mstrPara, "|")(1)
    Err.Clear: On Error GoTo 0

    strTmp = mobjXJCA_Client.XJCA_CertAuth(strWebUrl, strAppId, strName)
    
    '��֤֤������Ϣ��ʾ
    If strTmp <> "" Then
        strTmp = UCase(strTmp)
        strTmp = Mid(strTmp, InStr(strTmp, UCase("<success>")) + 9)
        strTmp = Mid(strTmp, 1, InStr(strTmp, UCase("</success>")) - 1)
        If strTmp = "FALSE" Then
            MsgBoxEx "��¼��֤ʧ��!", vbInformation + vbOKOnly, gstrSysName
            GetCertLogin = False: Exit Function
        End If
    Else
        MsgBoxEx "֤����֤����ֵΪ�գ�"
        GetCertLogin = False: Exit Function
    End If
    GetCertLogin = True
End Function

Public Function XJCA_VerifySign(ByVal strCert As String, ByVal strSignData As String, ByVal strSource As String) As Boolean
'����;��֤ǩ��
'����:strCurrCertSn -֤��ID(Ψһ����)
'     strCert -֤����Ϣ������Կ��Ϣ��
'     strSignData-ǩ��ֵ
'     strSource-����֤Դ��

        Dim blnRet As Boolean
        
        On Error GoTo errH
'        '�ո�vbTAb,vbCrLF ����ǩ���ӿ�ʱ��ͳһ����Դ���ص�ǩ��ֵ�п��ܲ�һ��
'        strSource = Replace(strSource, " ", "")
'        strSource = Replace(strSource, vbTab, "|")
'        strSource = Replace(strSource, vbCrLf, "||")
'
        blnRet = XJCA_VerifySeal(strSource, strSignData & "," & strCert, "", "")
        If blnRet Then
            MsgBoxEx "��֤�ɹ����õ���ǩ��������Ч!", vbInformation, gstrSysName
        Else
            MsgBoxEx "��֤ǩ��ʧ�ܣ�", vbExclamation, gstrSysName
            Exit Function
        End If
       
        XJCA_VerifySign = True
        Exit Function
errH:
104     MsgBoxEx "��֤ǩ��ʧ�ܣ�" & vbNewLine & "��" & CStr(Erl()) & "�� " & Err.Description, vbQuestion, gstrSysName
End Function

Private Function CheckIsXJCA() As Boolean
'���ܣ����CA����
    Dim lngRet As Long
    Dim strTmp As String
    
    If mbytType = 1 Then
        strTmp = M_STR_CSP
    Else
        strTmp = M_STR_CSP_HD
    End If
    '1-�ж�֤�������Ƿ�װ
    lngRet = mobjXJCA_Client.XJCA_CspInstalled(strTmp)
    If lngRet <> 10000 Then
        MsgBoxEx "֤������δ��װ��", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    '2-�ж�֤���Ƿ����
    lngRet = mobjXJCA_Client.XJCA_KeyInsert(strTmp)
    If lngRet <> 10000 Then
        MsgBoxEx "֤��KEYδ���룡", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    CheckIsXJCA = True
End Function

Public Sub XJCA_UnLoadObj()
    Set mobjXJCA_Client = Nothing
    Set mobjGseal = Nothing
    mblnInit = False
End Sub

Public Function XJCA_GetPara() As Boolean
'���ú���CA��������ַ
    
    On Error GoTo errH
    gstrPara = gobjComLib.zlDatabase.GetPara(90000, glngSys)  '��ȡURLs �̶���ȡZLHIS ϵͳĬ��100
    If gstrPara = "" Then gstrPara = "http://124.117.245.71:48080/webServices/ssoService|4028f6d24a2d7182014a2d83333e001a|����"
    If gstrPara <> "" Then
        gudtPara.strSignURL = gstrPara
    End If
    Exit Function
errH:
    MsgBoxEx "��ȡ����ʧ�ܣ�" & vbNewLine & "��" & CStr(Erl()) & "�� " & Err.Description, vbInformation, gstrSysName
End Function

Public Function XJCA_SetParaStr() As String
    XJCA_SetParaStr = gudtPara.strSignURL
End Function
