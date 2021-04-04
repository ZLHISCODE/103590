Attribute VB_Name = "mdlBJCAZX"
Option Explicit
'����CA���Ĺ���ģ��
Private mblnInit As Boolean         '�Ƿ��ѳ�ʼ���ɹ�
Private mLastPWD As String          '�������������

Private BJCA_Client As Object       '֤�鲿��
Private BJCA_svs As Object          '
Private BJCA_Pic As Object          '��ȡ֤��ͼƬ����
Private BJCA_TS  As Object          'ʱ�������
Private mblnTs As Boolean           '����ʱ���
Private mbytTSVer As Byte           'BJCA_TS_CLIENTCOMLib.BJCATSEngine/BJCA_TS_ClientCom.BJCATSEngine.1
                                    '"BJCA_TS_ClientCom.BJCATSEngine.1"-פ��꾫��ҽԺʱ�������;����Ϣ������ҽԺ
                                    '"BJCA_TS_CLIENTCOMLib.BJCATSEngine" -������ͯҽԺ
Private mLogin As Long              '��������������

Private Const BASE64CHR As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/="
Private Const STR_TS_VER_0 As String = "BJCA_TS_CLIENTCOMLIB.BJCATSENGINE"
Private Const STR_TS_VER_1 As String = "BJCA_TS_CLIENTCOM.BJCATSENGINE.1"


Public Function BJCA_InitObj() As Boolean
    '֤�鲿����ʼ��
        Dim progID As String
        Dim strVer As String
        
        On Error GoTo errH
100
102     BJCA_InitObj = mblnInit
104     If mblnInit Then Exit Function
        mLastPWD = ""
106     Set BJCA_svs = CreateObject("BJCA_SVS_ClientCOM.BJCASVSEngine.1")
108     Set BJCA_Client = CreateObject("BJCASecCOM.BJCASecCOMV2.1")
110     Set BJCA_Pic = CreateObject("GetKeyPic.GetPic")

112     gstrPara = gobjComLib.zlDatabase.GetPara(90000, glngSys)  '��ȡURLs
        'gstrPara = "0&&&1"   '��ʽ:����ʱ���&&&ʱ����汾
        '"BJCA_TS_ClientCom.BJCATSEngine.1"�°�\"BJCA_TS_CLIENTCOMLib.BJCATSEngine"�ϰ�
        If gstrPara = "" Then
            Err.Raise -1, , "��ǰϵͳ��" & glngSys & "��û�����õ���ǩ����������ʽ:����ʱ���|����ʱ����汾�����������á�"
            Exit Function
        End If
        mblnTs = (Val(Split(gstrPara, G_STR_SPLIT)(0)) = 1)
        mbytTSVer = Val(Split(gstrPara, G_STR_SPLIT)(1))
        
        If mblnTs Then
            If mbytTSVer = 0 Then
                strVer = STR_TS_VER_0
            ElseIf mbytTSVer = 1 Then
                strVer = STR_TS_VER_1
            End If
113         Set BJCA_TS = CreateObject(strVer)
        End If
        
114     BJCA_InitObj = True
    
116     mblnInit = BJCA_InitObj
        mLogin = 0
        Exit Function
errH:
118     MsgBoxEx "�����ӿڲ���ʧ�ܣ�" & vbNewLine & Err.Description, vbQuestion, gstrSysName
    
End Function

Public Function BJCA_RegCert(arrCertInfo As Variant) As Boolean
        '���ܣ��ṩ��HIS���ݿ���ע�����֤��ı�Ҫ��Ϣ,�����·��Ż����֤��,,��Ҫ����USB-Key
        '���أ�arrCertInfo��Ϊ���鷵��֤�������Ϣ
        '      0-ClientSignCertCN:�ͻ���ǩ��֤�鹫������(����),����ע��֤��ʱ������֤���
        '      1-ClientSignCertDN:�ͻ���ǩ��֤������(ÿ��Ψһ)
        '      2-ClientSignCertSN:�ͻ���ǩ��֤�����к�(ÿ֤��Ψһ)
        '      3-ClientSignCert:�ͻ���ǩ��֤������
        '      4-ClientEncCert:�ͻ��˼���֤������
        '      5-ǩ��ͼƬ�ļ���,�մ���ʾû��ǩ��ͼƬ
        
        Dim strKeyId As String, strCertTime As String, strCertUserName As String, strCertDN As String
        Dim strSigCert As String, i As Integer, strCACert As String, lngOk As Long
        Dim strPicData As String
        
        On Error GoTo errH
    
100     For i = LBound(arrCertInfo) To UBound(arrCertInfo)
102         arrCertInfo(i) = ""
        Next

104     If GetCertList(strCertUserName, strKeyId, strSigCert) Then
106         arrCertInfo(0) = strCertUserName
108         arrCertInfo(1) = strCertDN
110         arrCertInfo(2) = strKeyId
112         arrCertInfo(3) = strSigCert

114         If Not BJCA_Pic Is Nothing Then
116             If UBound(arrCertInfo) >= 5 Then
118                 strPicData = BJCA_Pic.getpic()
120                 If strPicData <> "" Then
                        '�°���Ӳ���gif��ʽǩ������,Ҫ��ĳ�bmp
122                     arrCertInfo(5) = SaveBase64ToFile("bmp", strKeyId, strPicData) 'ͼƬ·��
                    End If
                End If
            End If
124         BJCA_RegCert = True
        End If

        Exit Function
errH:
126     MsgBoxEx "��ȡ֤����Ϣʧ�ܣ�" & vbNewLine & "��" & CStr(Erl()) & "�� " & Err.Description, vbQuestion, gstrSysName

End Function

Public Function BJCA_Sign(ByVal strSource As String, ByRef strSignData As String, ByRef strTimeStamp As String, ByRef strTimeStampCode As String, ByRef blnReDo As Boolean) As Boolean
        'ǩ��
        Dim strTiemRequest As String
        Dim strTmp As String
        Dim blnCheck As Boolean
        
        On Error GoTo errH
        blnCheck = BJCA_CheckCert(blnReDo)
        If blnReDo Then Exit Function
        
100     If blnCheck Then               '��֤��ǰUSB�Ƿ���ǩ���û��ģ�����ȡǩ��֤��
            If mblnTs Then
                If mbytTSVer = 0 Then
                    '�ɹ����ؾ�base64�����ʱ�������ʧ�ܷ��ؿ�ֵ
                    strTiemRequest = BJCA_TS.CreateTimeStampRequest(strSource)
                    '�ɹ����ؾ�base64�����ʱ���������֤�飩��ʧ�ܷ��ؿ�ֵ
                    strTimeStampCode = BJCA_TS.CreateTimeStampNoCert(strTiemRequest)
                    If strTimeStampCode = "" Then
                        MsgBoxEx "��ȡʱ�����Ϣʧ�ܣ�"
                        Exit Function
                    Else
                        strTmp = BJCA_TS.gettimestampinfo(strTimeStampCode, 1) '����ʱ��
                        'ʱ������ظ�ʽ��20140911192555��ת���� 2014-09-11 19:25:55
                        strTimeStamp = Mid(strTmp, 1, 4) & "-" & Mid(strTmp, 5, 2) & "-" & Mid(strTmp, 7, 2) & " " & Mid(strTmp, 9, 2) & ":" & Mid(strTmp, 11, 2) & ":" & Mid(strTmp, 13, 2)
                    End If
                ElseIf mbytTSVer = 1 Then
                    strTiemRequest = BJCA_TS.CreateTSRequest(strSource, 0)   '����֤��
                    strTimeStampCode = BJCA_TS.CreateTS(strTiemRequest)
                    If strTimeStampCode = "" Then
                        MsgBoxEx "��ȡʱ�����Ϣʧ�ܣ�"
                        Exit Function
                    Else
                        strTmp = BJCA_TS.GetTSInfo(strTimeStampCode, 1) '����ʱ��
                        'ʱ������ظ�ʽ��20140911192555��ת���� 2014-09-11 19:25:55
                        strTimeStamp = Mid(strTmp, 1, 4) & "-" & Mid(strTmp, 5, 2) & "-" & Mid(strTmp, 7, 2) & " " & Mid(strTmp, 9, 2) & ":" & Mid(strTmp, 11, 2) & ":" & Mid(strTmp, 13, 2)
                    End If
                End If
            End If
            '֤��ID����ǩ��
110         strSignData = BJCA_Client.SignData(mUserInfo.strCertSn, strSource)
112
        Else
            MsgBoxEx "ǩ��ʧ�ܣ�", vbInformation, "����ǩ������"
            Exit Function
        End If
        BJCA_Sign = True
        Exit Function
errH:
114     MsgBoxEx "ǩ��ʧ�ܣ�" & vbNewLine & "��" & CStr(Erl()) & "�� " & Err.Description, vbQuestion, gstrSysName
End Function

Public Function BJCA_CheckCert(ByRef blnReDo As Boolean) As Boolean
'--------------------------------------------------------------------------------------------------------------------------
'���ܣ���ȡUSB�����豸��ʼ������¼
'����:
'   ����:blnRedo-֤�������Ҫ���¼��
'����:
'--------------------------------------------------------------------------------------------------------------------------
        Dim strPIN As String, strUserName As String
        Dim strWebUrl As String, intDate   As Integer
        Dim strCertSn As String
        Dim strPicData As String, strSigCert As String
        Dim strTmp As String, strFileName As String
        Dim blnRet As Boolean
        Dim udtUser As USER_INFO
        Dim strDate As String
        Dim strCertID As String
        
1       On Error GoTo errH
2       If Not BJCA_InitObj() Then
3           MsgBoxEx "����δ��ʼ����"
4           Exit Function
5       End If
    
6       Call GetCertList(strUserName, strCertID, strSigCert, strCertSn)
7       If mUserInfo.strUserID = "" Then
8           MsgBoxEx "�������֤��Ϊ��,����ϵ����Ա����Ա������¼�룡", vbOKOnly + vbInformation, gstrSysName
9           Exit Function
10      ElseIf mUserInfo.strUserID <> Mid(strCertSn, 3) Then
11          MsgBoxEx "��֤��δע�����������£�����ʹ�ã�"
12          Exit Function
13      End If
        
14      If mLastPWD <> "" Then strPIN = mLastPWD
'        strPIN = ""  'CA����ʦ��� ÿ��ǩ����Ҫ����������
15      If strPIN = "" Then
16          If Not frmPassword.ShowMe(strPIN) Then Exit Function
17      End If
        
18      If Not GetCertLogin(strCertID, strPIN, strSigCert, intDate, strWebUrl) Then
19          strPIN = ""
20          blnRet = False
21      Else
22          blnRet = True
23      End If
24      LogWrite "BJCA_CheckCert", "GetCertLogin����ֵ blnRet=" & blnRet
25      If blnRet Then
            '�ж��Ƿ���Ҫ����ע��֤��
26          udtUser.strName = strUserName
27          udtUser.strSignName = strUserName
28          udtUser.strUserID = Mid(strCertSn, 3) 'SF+���֤��
29          udtUser.strCertSn = strCertID
30          udtUser.strCertDN = ""
31          udtUser.strCert = strSigCert
32          udtUser.strEncCert = ""
33          udtUser.strCertID = strCertID
34          udtUser.strPicCode = BJCA_Pic.getpic()
            '��ȡ�Ѿ�ע��֤�����Ч�������� ���ڸ�ʽ:axBJCASecCOMV21 ����汾���������Ķ���2015/09/15
35          strDate = BJCA_Client.GetCertInfo(mUserInfo.strCert, 12)
36          If IsUpdateRegCert(udtUser, strDate, blnReDo) Then
37              blnRet = True
38          Else
39              blnRet = False
40          End If
41          LogWrite "BJCA_CheckCert", "IsUpdateRegCert����ֵ blnRet=" & blnRet
42      End If
        
43      mLastPWD = strPIN
44      BJCA_CheckCert = blnRet
    
45      Exit Function
errH:
46      MsgBoxEx "���USBKEYʧ�ܣ�" & vbNewLine & "��" & CStr(Erl()) & "�� " & Err.Description, vbQuestion, gstrSysName
End Function

Public Sub BJCA_UloadObj()
    Set BJCA_Client = Nothing
    Set BJCA_svs = Nothing
    Set BJCA_Pic = Nothing
    Set BJCA_TS = Nothing
    mblnInit = False
End Sub
'----- �������ڲ�����

''' ��ȡ�ͻ���֤���б�
''' ����boolean
Private Function GetCertList(ByRef strName As String, ByRef strUniqueID As String, ByRef strCert As String, Optional ByRef strCertSn As String) As Boolean
    '�ӿƴ��һ����ҽԺ��ȡ����֤���б���
    '-���:��
    '-����
    'strName :      ����ӿڷ��ص�֤������������
    'strUniqueID:   ����ӿڷ��ص�֤��������Ψһ��ʶ
    'strCert:       ����ӿڷ��ص�ǩ��֤��
      
    Dim strUsbkeyList As String
    Dim arrUserListLength As Integer
    Dim arrUserList() As String
    If BJCA_Client Is Nothing Then Set BJCA_Client = CreateObject("BJCASecCOM.BJCASecCOMV2.1")
    If BJCA_Pic Is Nothing Then Set BJCA_Pic = CreateObject("GetKeyPic.GetPic")
    
    strUsbkeyList = BJCA_Client.getUserList()
    arrUserList = Split(strUsbkeyList, "&&&")
    arrUserListLength = UBound(arrUserList)
    If (arrUserListLength = -1) Then
        MsgBoxEx "��������Key��", vbInformation, gstrSysName
        Exit Function
    End If
    If (arrUserListLength <> 0) Then
        Dim i As Integer
        For i = 0 To arrUserListLength - 1
            Dim strOption As String
            strOption = arrUserList(i)
            strName = Split(strOption, "||")(0)
            strUniqueID = Split(strOption, "||")(1)
            strCert = BJCA_Client.ExportUserCert(strUniqueID)
            strCertSn = BJCA_Client.GetCertInfoByOid(strCert, "1.2.156.112562.2.1.1.1")
            If strCertSn = "" Then
                'value="1.2.156.112562.2.1.1.1" �ñ�ʶΪ����CA SM2֤����Ψһ��ʶ
                'value="2.16.840.1.113732.2" �ñ�ʶΪ����CA RSA֤����Ψһ��ʶ
                strCertSn = BJCA_Client.GetCertInfoByOid(strCert, "2.16.840.1.113732.2") '����һ�ַ�ʽȡ����ʱȱʡ���ڶ��ַ�ʽȡ
            End If
        Next
    End If
    GetCertList = True
End Function

Private Function GetCertLogin(ByVal strUniqueID As String, ByVal strPassword As String, ByVal strCert As String, ByRef dDate As Integer, ByRef strWebserviceUrl As String) As Boolean
    '�ӿƴ��һ����ҽԺ����֤���¼����
    '- ���
    'strUniqueID : ֤��Ψһ��ʶ
    'strPassword : ֤������
    'strWebserviceUrl:ǩ����������ַ����Ϊ֤����֤
    '- ����
    'dDate       : ����֤����Чʱ��

    Dim result As Boolean
    If BJCA_Client Is Nothing Then Set BJCA_Client = CreateObject("BJCASecCOM.BJCASecCOMV2.1")
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
        result = BJCA_Client.userLogin(strUniqueID, strPassword)
        LogWrite "GetCertLogin", "�ӿ�userLogin" & "    ����ֵ:" & result
        If (result) Then
            mLogin = 0
            Dim strExtLib As String
            strExtLib = BJCA_Client.GetUserInfo(strUniqueID, 15)
            Dim intFlg As Integer
            
            '����������֤֤��
            '������е���֤��
            Dim retValidateCert As Long
            retValidateCert = 100
            retValidateCert = ValidateCert(strCert, strWebserviceUrl)
            LogWrite "GetCertLogin", "ValidateCert" & vbCrLf & _
                        "����1=" & strCert & vbCrLf & _
                        "����2=" & strWebserviceUrl & vbCrLf & _
                        "����ֵ:" & retValidateCert
            '��֤֤������Ϣ��ʾ
            If retValidateCert <> 0 Then Call ValidateCertView(retValidateCert)

            If (retValidateCert = 0) Then
                Dim uniqueIdStr As String
                Dim oid As String
                oid = "2.16.840.1.113732.2"
                Dim s As String
                '��ȡ�ͻ���֤����Ч�ڽ�ֹʱ��
                s = BJCA_Client.GetCertInfo(strCert, 12)
                '��֤�ͻ���֤����Ч��ʣ������
                dDate = CheckValidaty(s)
            
                If (dDate <= 30 And dDate > 0 And Not gblnShow) Then
                    MsgBoxEx "����֤�黹��" & dDate & "�����"
                    uniqueIdStr = BJCA_Client.GetCertInfoByOid(strCert, oid)
                    gblnShow = True
                    GetCertLogin = True
                ElseIf (dDate <= 0) Then
                    MsgBoxEx "����֤���ѹ��� " & Abs(dDate) & " ��"
                    GetCertLogin = False
                Else
                    uniqueIdStr = BJCA_Client.GetCertInfoByOid(strCert, oid)
                    GetCertLogin = True
                End If
            Else
                GetCertLogin = False
            End If
        Else
            mLogin = mLogin + 1
            MsgBoxEx "֤��������ܲ���ȷ�����Ѿ�������" & mLogin & "�����룬����������" & 8 - mLogin & "��!"
            GetCertLogin = False
        End If
    End If

End Function

Private Function ValidateCert(ByRef userCert As String, Optional webserviceUrl As String) As Integer
    '����������֤֤��
 
    If BJCA_svs Is Nothing Then Set BJCA_svs = CreateObject("BJCA_SVS_ClientCOM.BJCASVSEngine.1")
    ValidateCert = BJCA_svs.ValidateCertificate(userCert)
 
End Function

''' <summary>
''' ��֤֤������Ϣ��ʾ
''' </summary>
''' <remarks></remarks>
Private Sub ValidateCertView(retValidateCert)
    Select Case retValidateCert
        Case 0
            MsgBoxEx "֤����Ч��"
        Case -1
            MsgBoxEx "���������εĸ���"
        Case -2
            MsgBoxEx "������Ч�ڣ�"
        Case -3
            MsgBoxEx "����֤�飡"
        Case -4
            MsgBoxEx "�Ѽ����������"
    End Select
End Sub

''' �ͻ�����֤ǩ������
''' ����booleanֵ
Public Function BJCA_VerifySign(ByVal strCert As String, ByVal strInData As String, ByRef strData As String, ByVal strTimeStampCode As String) As Boolean
    '�ӿƴ��һ����ҽԺ����֤��ǩ����֤����
    '- ���
    'strInData     : ǩ�����
    'strCert       : ǩ��֤��
    'strData       : ǩ��ԭ��
    'strTimeStampCode :ʱ�����Ϣ
    '-����ֵ
    'result:true:  �ɹ�
    'result:false: ʧ��
        Dim intVerifyRet As Integer
        Dim lngResult As Long
        Dim strInfo As String
        Dim blnRet As Boolean
        
        On Error GoTo errH
        '����ֵ  �ɹ�����0��ʧ�ܷ�������ֵ��
        '-1Ϊʱ�����֤��ͨ��
        '-2Ϊԭ����֤��ͨ��
        '-3Ϊ���������εĸ�
        '-4֤��δ��Ч
        '-5��ѯ������֤��
        '-6Ϊǩ��ʱ���ʱ������֤�����
        If mblnTs Then
            If mbytTSVer = 0 Then
                lngResult = BJCA_TS.VerifyTimeStampData(strTimeStampCode, "") 'ֻ��֤ʱ���,����֤Դ��
            ElseIf mbytTSVer = 1 Then
                lngResult = BJCA_TS.VerifyTS(strTimeStampCode, strData)
            End If
            If lngResult <> 0 Then
                strInfo = "��֤ʱ���ʧ�ܣ�����:"
                Select Case lngResult
                Case -1
                    MsgBoxEx strInfo & "ʱ�����֤��ͨ����"
                Case -2
                    MsgBoxEx strInfo & "ԭ����֤��ͨ����"
                Case -3
                    MsgBoxEx strInfo & "���������εĸ���"
                Case -4
                    MsgBoxEx strInfo & "֤��δ��Ч��"
                Case -5
                    MsgBoxEx strInfo & "��ѯ������֤�飡"
                Case -6
                    MsgBoxEx strInfo & "ǩ��ʱ���������֤����ڣ�"
                End Select
                Exit Function
            End If
        End If

'100     If BJCA_Client Is Nothing Then Set BJCA_Client = CreateObject("BJCASecCOM.BJCASecCOMV2.1")
'101     verifySignResult = BJCA_Client.verifySignedData(strCert, strData, strInData)
'����ע�ʹ������������ͯҽԺ���ñ���CAʱ������CA����ʦ�������� ָ������֤ǩ��ʱ���÷�������֤�ķ�ʽ,
'Ӧ��ʹ�� �˶���"BJCA_SVS_ClientCOM.BJCASVSEngine.1"�� ������֤ǩ��
        intVerifyRet = -1
100        If BJCA_svs Is Nothing Then Set BJCA_svs = CreateObject("BJCA_SVS_ClientCOM.BJCASVSEngine.1")
101        intVerifyRet = BJCA_svs.VerifySignedData(strCert, strData, strInData)

        If intVerifyRet = 0 Then
            MsgBoxEx "��֤ǩ���ɹ���", vbInformation, gstrSysName
            blnRet = True
        Else
            MsgBoxEx "��֤ǩ��ʧ�ܣ�", vbInformation, gstrSysName
            blnRet = False
        End If
        BJCA_VerifySign = blnRet
    Exit Function
errH:
     MsgBoxEx "��֤ǩ��ʧ�ܣ�" & vbNewLine & "��" & CStr(Erl()) & "�� " & Err.Description, vbQuestion, gstrSysName
End Function

''' ���֤����Ч��
''' ����֤����Ч������
Private Function CheckValidaty(ByVal endDate As Date) As Integer
    '�ӿƴ��һ����ҽԺ���֤����Ч�Խӿ�
    '-���: ֤����Ч��ֹ����
    '-���Σ���Ч����
        Dim dblAllSp    As Double
        Dim result      As Integer
        dblAllSp = CDbl(CDate(endDate)) - CDbl(Now)
        result = Int(dblAllSp)
        CheckValidaty = result
End Function

Public Function BJCA_GetPara() As Boolean
'���ú���CA��������ַ
    Dim arrList As Variant
    
    On Error GoTo errH
    gstrPara = gobjComLib.zlDatabase.GetPara(90000, glngSys)
    If gstrPara = "" Then gstrPara = "0&&&1"   '��ʽ:����ʱ���&&&ʱ����汾
    arrList = Split(gstrPara, "&&&")
    If UBound(arrList) = 1 Then
        gudtPara.blnISTS = Val(arrList(0)) = 1
        gudtPara.strTSVersion = arrList(1)
    End If
    Exit Function
errH:
    MsgBoxEx "��ȡ����ʧ�ܣ�" & vbNewLine & "��" & CStr(Erl()) & "�� " & Err.Description, vbInformation, gstrSysName
End Function

Public Function BJCA_SetParaStr() As String
    BJCA_SetParaStr = IIf(gudtPara.blnISTS, 1, 0) & G_STR_SPLIT & gudtPara.strTSVersion
End Function






