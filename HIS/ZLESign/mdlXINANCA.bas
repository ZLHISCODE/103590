Attribute VB_Name = "mdlXINANCA"
Option Explicit
Private mobjSecCtrl As Object      'npmobjSecCtrl.dll
'Private mobjSecCtrl As New SecCtrlLib.CACtrlCom
Private Const mstrTitle As String = "�Ű�CA"
Private mblnInit As Boolean
Private mblnLogin As Boolean    '�Ƿ���Ҫ��¼��֤

Public Function XINANCA_InitObj() As Boolean
    '֤�鲿����ʼ��
    'ʱ����������Ե�ַ 218.29.120.82 port:9198
        Dim lngRet As Long
        
        On Error GoTo ErrH
    
100     If mblnInit Then XINANCA_InitObj = True: Exit Function
        Call XINANCA_GetPara
        If gudtPara.strTSIP <> "" Then gudtPara.blnISTS = True   '����ʱ���
102     Set mobjSecCtrl = CreateObject("SecCtrl.CACtrlCom") '��̬����

106     XINANCA_InitObj = True
108     mblnInit = True
        Exit Function
ErrH:
    GetErrMsg Erl()
End Function

Public Function XINANCA_RegCert(arrCertInfo As Variant, Optional ByVal strUserID As String) As Boolean
'���ܣ��ṩ��HIS���ݿ���ע�����֤��ı�Ҫ��Ϣ,�����·��Ż����֤��,,��Ҫ����USB-Key
'����:strUserID-���֤��
'���أ�arrCertInfo��Ϊ���鷵��֤�������Ϣ
'      0-ClientSignCertCN:�ͻ���ǩ��֤�鹫������(����),����ע��֤��ʱ������֤���
'      1-ClientSignCertDN:�ͻ���ǩ��֤������(ÿ��Ψһ)
'      2-ClientSignCertSN:�ͻ���ǩ��֤�����к�(ÿ֤��Ψһ)
'      3-ClientSignCert:�ͻ���ǩ��֤������
'      4-ClientEncCert:�ͻ��˼���֤������
'      5-ǩ��ͼƬ�ļ���,�մ���ʾû��ǩ��ͼƬ
        
    Dim strCertUserID As String, strCertUserName As String, strCertDN As String
    Dim strCert As String, i As Integer
    Dim strCertSn As String
    Dim strPicFile As String
    On Error GoTo ErrH
    
    For i = LBound(arrCertInfo) To UBound(arrCertInfo)
        arrCertInfo(i) = ""
    Next
    
    If GetCertList(strCertUserName, strCertSn, strCertDN, strCertUserID, strCert, strPicFile) Then
        arrCertInfo(0) = strCertUserName
        arrCertInfo(1) = strCertDN '֤��DN
        arrCertInfo(2) = strCertSn '֤�����к� ǩ��ʱҪ��
        arrCertInfo(3) = strCert
        arrCertInfo(4) = ""
        arrCertInfo(5) = strPicFile
        XINANCA_RegCert = True
    End If
    Exit Function
ErrH:
    MsgBoxEx "��ȡ֤����Ϣʧ�ܣ�" & vbNewLine & _
        "��" & CStr(Erl()) & "�� " & Err.Description, vbExclamation, mstrTitle
End Function

Private Function GetCertList(ByRef strName As String, Optional ByRef strCertSn As String = "0", Optional ByRef strCertDN As String = "0", _
           Optional ByRef strCertUserID As String = "0", Optional ByRef strCert As String = "0", Optional ByRef strPicFile As String = "0") As Boolean
'����:�Ű�CA��ȡ֤������
'-���:��
'-����
'strName :      ����ӿڷ��ص�֤������������
'strCertSN      ����ӿڷ��ص�֤��SN
'strCertDN:     ����ӿڷ��ص�֤��DN
'strCertUserID:  ����ӿڷ��ص�֤��������Ψһ��ʶ
'strCert:       ����ӿڷ��ص�ǩ��֤��

        On Error GoTo ErrH
   
        Dim lngRet As Long
        Dim strPic As String
        
100     lngRet = mobjSecCtrl.KS_SetProv("XACA", 0, "") '��ʼ�� δ��KEYʱ�ᵯ��������ʾ
102     If mobjSecCtrl.KS_GetLastErrorCode() <> 0 Then
104         MsgBoxEx mobjSecCtrl.KS_GetLastErrorMsg(), vbExclamation + vbOKOnly, mstrTitle
            Exit Function
        End If
106     strCert = GetCert(2) 'type: 1-����֤�飬2-ǩ��֤��
108     If strCert = "" Then MsgBoxEx "��ȡ֤����Ϣʧ�ܣ�", vbExclamation, mstrTitle: Exit Function
        '15-֤��ӵ���߲�����(OU): С������
110      If Not GetCertInfo(strCert, 15, strName) Then Exit Function
        '17-֤��ӵ����ͨ������(CN):4127589685665
112      If strCertSn <> "0" Then If Not GetCertInfo(strCert, 17, strCertSn) Then Exit Function
    
        '21-֤��ӵ����DN:C=CN,S=����ʡ,L=֣����,O=����ʡ�ط�˰���,OU=С������,CN=4127589685665
114     If strCertDN <> "0" Then
116          If Not GetCertInfo(strCert, 21, strCertDN) Then Exit Function
        End If
        
        If strPicFile <> "0" Then
            If Not XINANCA_GetSeal(strPic) Then Exit Function
            strPicFile = FormatPic("gif", strCertSn, strPic)
        End If
        
        GetCertList = True
        Exit Function
ErrH:
        MsgBoxEx Err.Description & vbCrLf & _
                "��GetCertList ������: " & Erl, _
                    vbExclamation + vbOKOnly, mstrTitle
         
End Function
 
Private Function XINANCA_GetSeal(ByRef strSeal As String) As Boolean
      Dim strFileName As String
      Dim strTemp As String


10       On Error GoTo ErrH
      'mobjSecCtrl.KS_SetProv("XACA", 0, "")��ʼ���ɹ����ٵ���
20    strFileName = mobjSecCtrl.KS_GetSealList()
30    If mobjSecCtrl.KS_GetLastErrorCode() Then
40        MsgBoxEx "�õ�ӡ���б����:" & mobjSecCtrl.KS_GetLastErrorMsg(), vbExclamation + vbOKOnly, mstrTitle
50        Exit Function
60    End If
70    strTemp = mobjSecCtrl.KS_GetSeal(strFileName)
80    If mobjSecCtrl.KS_GetLastErrorCode() Then
90        MsgBoxEx "�õ�ӡ������ʧ��:" & mobjSecCtrl.KS_GetLastErrorMsg(), vbExclamation + vbOKOnly, mstrTitle
100       Exit Function
110   End If
120   strSeal = mobjSecCtrl.KS_GetInfoFromSeal(strTemp, 1)
130   If mobjSecCtrl.KS_GetLastErrorCode() Then
140       MsgBoxEx "��ȡͼƬ����:" & mobjSecCtrl.KS_GetLastErrorMsg(), vbExclamation + vbOKOnly, mstrTitle
150       Exit Function
160   End If
      strSeal = Replace(strSeal, vbLf, "")  '�Է������ַ��������з�
      WriteLog "ͼƬ��Ϣ:" & strSeal
170   XINANCA_GetSeal = True
180   Exit Function
ErrH:
190           MsgBoxEx Err.Description & vbCrLf & _
                      "��XINANCA_GetSeal ������: " & Erl, _
                          vbExclamation + vbOKOnly, mstrTitle
End Function

Public Function XINANCA_CheckCert() As Boolean
    '����:
    '   1-���֤���Ƿ����
    '   2-��鵱ǰ֤���Ƿ�ע���ڵ�ǰ�û�����
        Dim strName As String
        Dim strCertSn As String
        Dim strCert As String
        Dim strPIN As String
        Dim lngResult As Long
        
        On Error GoTo ErrH
100     If Not GetCertList(strName, strCertSn) Then XINANCA_CheckCert = False: Exit Function
102     If strCertSn <> mUserInfo.strCertSn Then
104         MsgBoxEx "��֤��δע�����������£�����ʹ�ã�" & vbCrLf & _
                    "�û�ע��֤��Ψһ��ʶ:" & mUserInfo.strCertSn & vbCrLf & _
                    "��ǰ��ѡ֤��Ψһ��ʶ:" & strCertSn, vbInformation + vbOKOnly, gstrSysName
            Exit Function
        End If
106     If Not mblnLogin Then
108         If Not Login() Then
                Exit Function
            Else
110             mblnLogin = True
            End If
        End If
112     XINANCA_CheckCert = True
        Exit Function
ErrH:
114     MsgBoxEx Err.Description & vbCrLf & _
                "XINANCA_CheckCert ������: " & Erl, _
                    vbExclamation + vbOKOnly, mstrTitle
End Function

Private Function Login() As Boolean
    '����:�Ű�CA����֤���¼����
    '- ���
    'strCertID            :֤��ID
    'strCert              ֤������BASE64����
    Dim strRandom As String, strSignVal As String
    Dim strDate As String
    Dim intDay As Integer
 
    Dim lngRet As Long
    
        On Error GoTo ErrH
         
100     strRandom = GenRandom(16)  '��ȡ�����
102     If strRandom = "" Then Exit Function
104     strSignVal = SignDataByP7(strRandom, 0)
106     If strSignVal = "" Then Exit Function
108     lngRet = VerifySignData(strRandom, strSignVal)
110     If mobjSecCtrl.KS_GetLastErrorCode() <> 0 Then
112         MsgBoxEx "�������ǩʧ�ܣ�" & vbNewLine & mobjSecCtrl.KS_GetLastErrorMsg(), vbExclamation + vbOKOnly, mstrTitle
            Exit Function
        End If
114     Login = True
        Exit Function
ErrH:
116     MsgBoxEx "��¼��֤ʧ�ܣ�" & vbNewLine & "��" & CStr(Erl()) & "�� " & Err.Description, vbExclamation, gstrSysName
End Function

Public Function XINANCA_Sign(ByVal strSource As String, ByRef strSignData As String, _
            ByRef strTimeStamp As String, ByRef strTimeStampCode As String, Optional ByVal blnCheck As Boolean) As Boolean
    '����:
    Dim strURL As String
    Dim strParameter As String
    Dim bytRet() As Byte
    Dim varTemp As Variant
    
        On Error GoTo ErrH
100     If Not blnCheck Then
102         blnCheck = XINANCA_CheckCert()
        End If
    
104     If blnCheck Then
106         strSource = StringSHA1(strSource)
108         strSignData = SignDataByP7(strSource, 0)
110         If strSignData = "" Then
112             MsgBoxEx "ǩ��ʧ�ܣ�", vbExclamation, mstrTitle
                Exit Function
            End If
        Else
114         MsgBoxEx "ǩ��ʧ�ܣ�", vbExclamation, mstrTitle
            Exit Function
        End If
        '����ʱ���
116     If gudtPara.blnISTS Then
118         strURL = "http://" & gudtPara.strTSIP & ":" & gudtPara.strTSPort & "/tsac.svr"
120         strParameter = "digest=" & strSource
122         bytRet = HttpPost(strURL, strParameter, responseBody)
124         strTimeStampCode = EncodeBase64Byte(bytRet)
126         If strTimeStampCode = "" Then
128             MsgBoxEx "��ȡʱ�����Ϣʧ�ܣ�", vbExclamation, mstrTitle
                Exit Function
            End If
130         strURL = "http://" & gudtPara.strTSIP & ":" & gudtPara.strTSPort & "/tsav.svr"
132         strParameter = "tsr=" & Replace(strTimeStampCode, "+", "%2B")
134         strTimeStamp = HttpPost(strURL, strParameter, responseText)
            LogWrite "XINANCA_Sign", "ʱ�������ֵ��" & strTimeStamp
136         If strTimeStamp = "" Then
138             MsgBoxEx "��ȡʱ���ʧ�ܣ�", vbExclamation, mstrTitle
                Exit Function
            Else
140             strTimeStamp = Mid(strTimeStamp, InStr(strTimeStamp, "<timestamp>") + Len("<timestamp>"))
142             strTimeStamp = Mid(strTimeStamp, 1, InStr(strTimeStamp, "</timestamp>") - 1)
                strTimeStamp = Replace(strTimeStamp, Space(2), Space(1))  '����Ϊһλ��ʱǰ����ܴ��ڿո��½���ʧ��
144             varTemp = Split(strTimeStamp, Space(1)) 'Jan 21 06:34:28.865495 2019 GMT ʱ��ֻȡǰ��λ�ַ�
146             strTimeStamp = varTemp(3) & "-" & ConvMonth(varTemp(0)) & "-" & varTemp(1) & " " & Mid(varTemp(2), 1, 8) '��������ʱ��
148             strTimeStamp = Format(DateAdd("h", 8, strTimeStamp), "YYYY-MM-DD HH:MM:SS")
            End If
        Else
150         strTimeStamp = Format(gobjComLib.zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
        End If
152     XINANCA_Sign = True
        Exit Function
ErrH:
154       MsgBoxEx "ǩ��ʧ�ܣ�" & vbNewLine & "��" & CStr(Erl()) & "�� " & Err.Description, vbExclamation, mstrTitle
End Function

Public Function XINANCA_VerifySign(ByVal strSignData As String, ByVal strSource As String, ByVal strTimeStampCode As String) As Boolean
    '����:
        Dim lngRet As Long
        Dim strURL As String
        Dim strParameter As String
        
        On Error GoTo ErrH
100     strSource = StringSHA1(strSource)
102     lngRet = VerifySignData(strSource, strSignData)
104     If lngRet <> 0 Then
106         MsgBoxEx "��֤ʧ�ܣ��õ���ǩ��������Ч!", vbInformation, mstrTitle
            Exit Function
        End If
108     If gudtPara.blnISTS Then
110         strURL = "http://" & gudtPara.strTSIP & ":" & gudtPara.strTSPort & "/tsav.svr"
112         strParameter = "tsr=" & Replace(strTimeStampCode, "+", "%2B")
114         strTimeStampCode = HttpPost(strURL, strParameter, responseText)
116         If strTimeStampCode = "" Then
118             MsgBoxEx "��֤ʱ���ʧ�ܣ�", vbExclamation, mstrTitle
                Exit Function
            End If
        End If
120     MsgBoxEx "��֤�ɹ����õ���ǩ��������Ч!", vbInformation, mstrTitle
122     XINANCA_VerifySign = True
    Exit Function
ErrH:
124  MsgBoxEx "��֤ǩ��ʧ�ܣ�" & vbNewLine & "��" & CStr(Erl()) & "�� " & Err.Description, vbExclamation, mstrTitle
End Function


Public Function XINANCA_UnLoad()
    Set mobjSecCtrl = Nothing
    mblnInit = False
End Function
'/**
' * ��ȡBASE64����֤��
' * type: 1-����֤�飬2-ǩ��֤��
' */
Private Function GetCert(ByVal lngType As Long) As String
    Dim strResult As String
    
    strResult = mobjSecCtrl.KS_GetCert(lngType)
    If mobjSecCtrl.KS_GetLastErrorCode() <> 0 Then
        MsgBoxEx mobjSecCtrl.KS_GetLastErrorMsg(), vbExclamation + vbOKOnly, mstrTitle
        Exit Function
    End If
    GetCert = strResult
End Function

'/**
' * ��ȡ֤����Ϣ
' * cert: Base64����֤��
' * item: �����
' * 1��֤��汾 2��֤�����к� 3��֤��ǩ���㷨��ʶ 4��֤��䷢�߹���(C)  5��֤��䷢����֯��(O)
' * 6��֤��䷢�߲�����(OU)  7��֤��䷢�����ڵ�ʡ����������ֱϽ��(S)  8��֤��䷢��ͨ������(CN)  9��֤��䷢�����ڵĳ��С�����(L)
' * 10��֤��䷢��Email  11��֤����Ч�ڣ���ʼ����:180410101818  12��֤����Ч�ڣ���ֹ����:190410101818  13��֤��ӵ���߹���(C )  14��֤��ӵ������֯��(O)
' * 15��֤��ӵ���߲�����(OU)  16��֤��ӵ�������ڵ�ʡ����������ֱϽ��(S)  17��֤��ӵ����ͨ������(CN)  18��֤��ӵ�������ڵĳ��С�����(L)
' * 19��֤��ӵ����Email  20��֤��䷢��DN  21��֤��ӵ����DN  22��֤�鹫Կ��Ϣ  23��CRL������.
' */
Private Function GetCertInfo(ByVal strCert As String, ByVal lngItem As Long, ByRef strResult As String) As Boolean
     
    strResult = mobjSecCtrl.KS_GetCertInfo(strCert, lngItem)
    If mobjSecCtrl.KS_GetLastErrorCode() <> 0 Then
        MsgBoxEx mobjSecCtrl.KS_GetLastErrorMsg(), vbExclamation + vbOKOnly, mstrTitle
        Exit Function
    End If
    GetCertInfo = True
End Function
 
'/**
' * ��ȡ֤����չ��Ϣ
' * cert: Base64����֤��
' * oid: oidֵ
' */
Private Function GetCertInfoByOid(ByVal strCert As String, ByVal strOid As String) As String
    Dim strResult As String
    strResult = mobjSecCtrl.KS_GetCertInfoByOid(strCert, strOid)
    If mobjSecCtrl.KS_GetLastErrorCode() <> 0 Then
        MsgBoxEx mobjSecCtrl.KS_GetLastErrorMsg(), vbExclamation + vbOKOnly, mstrTitle
    End If
    GetCertInfoByOid = strResult
End Function

'/**
' * ���������
' * len: ���������
' */
Private Function GenRandom(ByVal lngLen As Long) As String
    Dim strResult As String
    strResult = mobjSecCtrl.KS_GenRandom(lngLen)
    If mobjSecCtrl.KS_GetLastErrorCode() <> 0 Then
        MsgBoxEx mobjSecCtrl.KS_GetLastErrorMsg(), vbExclamation + vbOKOnly, mstrTitle
        Exit Function
    End If
    GenRandom = strResult
End Function

Private Function VerifySignData(ByVal strSource As String, ByVal strSignData As String) As Long
'����:��������֤ǩ��
    Dim lngResult As Long

    lngResult = mobjSecCtrl.KS_P7RemoteVerify(1, strSignData, strSource)  '���صĽ�������Σ�0Ϊ�ɹ�����0Ϊʧ��
    If mobjSecCtrl.KS_GetLastErrorCode() <> 0 Then
        MsgBoxEx mobjSecCtrl.KS_GetLastErrorMsg(), vbExclamation + vbOKOnly, mstrTitle
    End If
    VerifySignData = lngResult
End Function

'/**
' * ����ǩ��P7
' * indata����������
' * hashAlg:0. AUTO(�Զ�ѡ�񣬵�RSAʱΪSHA1, SM2ʱΪSM3), 1-SHA1, 2-SHA256, 3-SHA512, 4-MD5, 5-MD4, 6-SM3
' * return��ǩ������
' */
Private Function SignDataByP7(ByVal strSource As String, ByVal lngHashAlg As Long) As String
    Dim strResult As String

    strResult = mobjSecCtrl.KS_SetParam("signtype", "pksc7")
    strResult = mobjSecCtrl.KS_SignData(strSource, lngHashAlg)
    If mobjSecCtrl.KS_GetLastErrorCode() <> 0 Then
        MsgBoxEx mobjSecCtrl.KS_GetLastErrorMsg(), vbExclamation + vbOKOnly, mstrTitle
        Exit Function
    End If
    SignDataByP7 = strResult
End Function

 Public Function XINANCA_GetPara() As Boolean
    '���÷�������ַ
    
    On Error GoTo ErrH
     
    'If gstrPara = "" Then gstrPara = "192.168.20.203" & G_STR_SPLIT & "9198"
    '�������Ե�ַ 218.29.120.82 port:9198
    'gudtPara.strTSIP="" ��������ʱ�������
    gudtPara.strTSIP = GetThirdPara(CON_PAR_�Ű�, "ʱ���IP")
    gudtPara.strTSPort = GetThirdPara(CON_PAR_�Ű�, "ʱ����˿�")
   
    Exit Function
ErrH:
    MsgBoxEx "��ȡ����ʧ�ܣ�" & vbNewLine & "��" & CStr(Erl()) & "�� " & Err.Description, vbInformation, gstrSysName
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
