VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsHNCA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����������Ŀ
Private mobjHNCACTRL As Object
Private mobjHNCATS As Object
Private mblnTs As Boolean
Private mbytFunc As Byte   '0-RSA�㷨,1-SM2�㷨

Private Declare Function doPostEx Lib "HNCACertTools6.dll" (ByVal strUrl As String, ByVal strAPPCode As String, ByVal strSource As String, _
        ByVal strSignData As String, ByVal strCertPUB As String, ByVal strGBK As String) As Integer

Public Function HNCA_CheckCert(ByRef blnReDo As Boolean) As Boolean
'���ܣ���ȡUSB�����豸��ʼ������¼
    Dim strKey As String, strPIN As String, strUserName As String
    Dim strCertName As String, strCertDN As String, strPicPath As String
    Dim strCertSn As String
    Dim strCertUserID As String    '��������֤����Ϣ
    Dim strDate As String
    Dim udtUser As USER_INFO
    Dim strCert As String, strCertID As String
    Dim blnOk As Boolean
    Dim blnRet As Boolean
    
    On Error GoTo errH
1000
     '��ȡ֤����Ϣͬʱ���Key���Ƿ����
    If Not GetCertList(strCertName, strCertSn, strCert, strCertDN, strPicPath, strCertUserID) Then
        HNCA_CheckCert = False: Exit Function
    End If
    'δע���ڵ�ǰ�û����µ�Key
    If mbytFunc = 0 Then
        If mUserInfo.strUserID = "" Then
            MsgBoxEx "��������֤��Ϊ��,����ϵ����Ա����Ա������¼�룡", vbOKOnly + vbInformation, gstrSysName
            Exit Function
        ElseIf strCertUserID <> mUserInfo.strUserID Then
            MsgBoxEx "��������֤�ţ�" & _
                       vbCrLf & vbTab & "��" & mUserInfo.strUserID & "��" & vbCrLf & _
                       "��ǰ֤��Ψһ��ʶ:" & _
                       vbCrLf & vbTab & "��" & strCertUserID & "��" & vbCrLf & _
                       "�û�����֤���뵱ǰ֤��Ψһ��ʶ�����,����ʹ�ã�", vbInformation, gstrSysName
            Exit Function
        End If
    Else
        If mUserInfo.strCertSn <> strCertSn Then
            MsgBoxEx "��֤��δע�����������£�����ʹ�ã�", vbInformation + vbOKOnly, gstrSysName
            Exit Function
        End If
    End If
1100
    '��¼��֤
    If InStr(gstrLogins & "|", "|" & strCertSn & "|") > 0 Then '�״���֤ͨ�����´β��ڼ�����֤
        blnOk = True
    Else
        If Not GetCertLogin(strCert, strCertDN) Then
            blnOk = False
        Else
            blnOk = True
            If InStr(gstrLogins & "|", "|" & strCertSn & "|") = 0 Then gstrLogins = gstrLogins & "|" & strCertSn
        End If
    End If
1200
    If blnOk And mbytFunc = 0 Then
        '�ж��Ƿ���Ҫ����ע��֤��
        udtUser.strName = strCertName
        udtUser.strSignName = strCertName
        udtUser.strUserID = strCertUserID
        udtUser.strCertSn = strCertSn
        udtUser.strCertDN = strCertDN
        udtUser.strCert = strCert
        udtUser.strPicPath = strPicPath
        udtUser.strPicCode = ""
        '��ȡ�Ѿ�ע��֤�����Ч�������� ���ڸ�ʽ:2005��12��23�� 18:09:31
        strDate = mobjHNCACTRL.GetCertInfo(6, "")
        strDate = Format(strDate, "YYYY-MM-DD HH:MM:SS")
        If IsUpdateRegCert(udtUser, strDate, blnReDo) Then
            HNCA_CheckCert = True
        Else
            HNCA_CheckCert = False
        End If
    Else
        HNCA_CheckCert = blnOk
    End If
1300
    
    Exit Function
errH:
     MsgBoxEx "���USBKEYʧ�ܣ�" & vbNewLine & "��" & CStr(Erl()) & "�� " & Err.Description, vbInformation, gstrSysName
End Function

Public Function HNCA_RegCert(arrCertInfo As Variant) As Boolean
        '���ܣ��ṩ��HIS���ݿ���ע�����֤��ı�Ҫ��Ϣ,�����·��Ż����֤��,,��Ҫ����USB-Key
        '���أ�arrCertInfo��Ϊ���鷵��֤�������Ϣ
        '      0-ClientSignCertCN:�ͻ���ǩ��֤�鹫������(����),����ע��֤��ʱ������֤����
        '      1-ClientSignCertDN:�ͻ���ǩ��֤������(ÿ��Ψһ)
        '      2-ClientSignCertSN:�ͻ���ǩ��֤�����к�(ÿ֤��Ψһ)
        '      3-ClientSignCert:�ͻ���ǩ��֤������
        '      4-ClientEncCert:�ͻ��˼���֤������
        '      5-ǩ��ͼƬ�ļ���,�մ���ʾû��ǩ��ͼƬ
        
        Dim strCertSn As String, strCertTime As String, strCertUserName As String, strCertDN As String
        Dim strSigCert As String, i As Integer, strCACert As String, lngOk As Long
        Dim strPicPath As String
        On Error GoTo errH
    
100     For i = LBound(arrCertInfo) To UBound(arrCertInfo)
102         arrCertInfo(i) = ""
        Next
    
104     If GetCertList(strCertUserName, strCertSn, strSigCert, strCertDN, strPicPath) Then
106         arrCertInfo(0) = strCertUserName
108         arrCertInfo(1) = strCertDN
110         arrCertInfo(2) = strCertSn
112         arrCertInfo(3) = strSigCert
            arrCertInfo(4) = ""
113         arrCertInfo(5) = strPicPath

            
124         HNCA_RegCert = True
        End If

        Exit Function
errH:
126     MsgBoxEx "��ȡ֤����Ϣʧ�ܣ�" & vbNewLine & "��" & CStr(Erl()) & "�� " & Err.Description, vbQuestion, gstrSysName

End Function

Public Function HNCA_Sign(ByVal strCurrCertSn As String, ByVal strSource As String, ByRef strSignData As String, _
        ByRef strTimeStamp As String, ByRef strTimeStampCode As String, ByRef blnReDo As Boolean) As Boolean
        'ǩ��
        Dim strTSURL As String, strUserName As String, strUserPWD As String
        Dim strRequest As String
        Dim intRet As Integer, intPort As Integer, intSSLPort As Integer
        Dim blnCheck As Boolean, blnRet As Boolean
        Dim strMsg As String
        Dim arrTmp As Variant
        
        On Error GoTo errH
        blnCheck = HNCA_CheckCert(blnReDo)
        If blnReDo Then Exit Function
1100    If blnCheck Then
            '��֤��ǰUSB�Ƿ���ǩ���û��ģ�����ȡǩ��֤��
            If mbytFunc = 0 Then
1110            strSignData = mobjHNCACTRL.Sign(strCurrCertSn, strSource)    'ǩ��ֵ��������Դ
            Else
1111            strSignData = mobjHNCACTRL.SM2Sign("", strSource)    'ǩ��ֵ��������Դ
            End If
            If strSignData <> "" Then
                If mblnTs Then
                    '��ԭ������ʱ���
                    intRet = -1
                    strRequest = mobjHNCATS.SOF_CreateTimeStampRequestByPlain(strSource)
                    intRet = mobjHNCATS.SOF_GetLastError()
                    If (intRet <> 0) Then
                        strMsg = "����ʱ�������ʧ�ܣ�ʧ�ܴ����ǣ�" & intRet
                        blnRet = False
                    Else
                        intRet = -1
                        strTimeStampCode = mobjHNCATS.SOF_CreateTimeStampResponse(strRequest)
                        intRet = mobjHNCATS.SOF_GetLastError()
                        If (intRet <> 0) Then
                            strMsg = "����ʱ�����Ӧʧ�ܣ�ʧ�ܴ����ǣ�" & intRet
                            blnRet = False
                        Else
                            strTimeStamp = mobjHNCATS.SOF_GetTimeStampInfo(strTimeStampCode, 1) '��ʽ;2015��10��20�� 02:18:45.534  GMT
                            strTimeStamp = Format(Mid(strTimeStamp, 1, 20), "YYYY-MM-DD HH:MM:SS")
                            blnRet = True
                        End If
                    End If
                Else
                    blnRet = True
                    strTimeStamp = Format(gobjComLib.zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
                End If
            Else
                strMsg = "ǩ��ʧ�ܣ�"
                blnRet = False
1112        End If
        Else
            strMsg = "ǩ��ʧ�ܣ�"
            blnRet = False
        End If
        If strMsg <> "" Then
            MsgBoxEx strMsg, vbInformation, gstrSysName
        End If
                
        HNCA_Sign = blnRet
        Exit Function
errH:
114     MsgBoxEx "ǩ��ʧ�ܣ�" & vbNewLine & "��" & CStr(Erl()) & "�� " & Err.Description, vbQuestion, gstrSysName
End Function

Public Function HNCA_VerifySign(ByVal strCert As String, ByVal strSignData As String, ByVal strSource As String, ByVal strTimeStampCode As String) As Boolean
'��֤ǩ��
'
        Dim strSigCert As String, strMsg As String
        Dim intRet As Integer
        Dim blnRet As Boolean
        On Error GoTo errH
     
        '��ȡ�ͻ���ǩ��(��Կ)֤��Base64���� �ͻ�����֤ǩ��,��������Դ����֤ǩ��
        If mbytFunc = 0 Then
1000        intRet = mobjHNCACTRL.VerifySign(strSignData, strSource, strCert)
        Else
1001        intRet = mobjHNCACTRL.SM2Verify(strSignData, strSource, strCert)
        End If
        If intRet = 0 Then   '��֤ǩ���ɹ�
            If mblnTs And strTimeStampCode <> "" Then
                '��֤ʱ���
                intRet = -1
                intRet = mobjHNCATS.SOF_VerifyTimeStampByPlain(strSource, strTimeStampCode)
            End If
            If (intRet = 0) Then
                strMsg = "��ǩ�ɹ���"
                blnRet = True
            Else
                strMsg = "ʱ�����֤ʧ�ܣ�������룺" & intRet
                blnRet = False
            End If
        Else
            strMsg = "��ǩʧ�ܣ�"
            blnRet = False
        End If
            
        If strMsg <> "" Then
            MsgBoxEx strMsg, vbInformation, gstrSysName
        End If
        HNCA_VerifySign = blnRet
        
2000    Exit Function
errH:
104     MsgBoxEx "��֤ǩ��ʧ�ܣ�" & vbNewLine & "��" & CStr(Erl()) & "�� " & Err.Description, vbQuestion, gstrSysName
End Function


Private Function GetCertList(ByRef strName As String, ByRef strCertSn As String, ByRef strCertPUB As String, _
                ByRef strCertDN As String, Optional strPicPath As String, Optional strUserID As String) As Boolean
'����:��ȡ֤����Ϣ
'-����
'    strName ֤�����������
'   strCertSN ֤��Ψһ��ʶ
'   strCertPUB ֤�鹫Կ  ��֤ǩ���õ�
'   strCertDN ֤��������Ϣ  ֤��ע���õ�
'   strPicPath ֤��ͼƬ����λ��

    Dim strCertCN As String, strPic As String
    Dim intRet As Integer, intCnIndex As Integer
    Dim strPara As String, strSign As String
    
    On Error GoTo errH

10    If mobjHNCACTRL Is Nothing Then Set mobjHNCACTRL = CreateObject("HNCACTRL.HNCACtrlCtrl.1")
      If mobjHNCATS Is Nothing Then Set mobjHNCATS = CreateObject("TSAACTIVEXVER4.TsaActiveXver4Ctrl.1")    'ʱ�������
      
    '0 ��ʾֻ��һ��֤��ʱҲ����֤��ѡ���1 ��ʾֻ��һ��֤��ʱ��������֤��ѡ���Ĭ��ֵΪ0
20    mobjHNCACTRL.SetCertChooseType (1)
    '����HNCA�䷢��ǩ��֤���б�
    If mbytFunc = 0 Then
30      intRet = mobjHNCACTRL.SetCert("SC", "", "", "", "CN=HNCA,O=HNCA,L=ZZ,S=HN,C=CN", "")
    Else
31      intRet = mobjHNCACTRL.SetCertEx("SC", "", "", "", "", "", "SM2")
    End If
    '�ж��Ƿ��ʼ���ɹ�
    If intRet <> 0 Then
        MsgBoxEx "��ʼ���ͻ���֤�����" & mobjHNCACTRL.GetErrorMessage(), vbOKOnly, "����CA"
        Exit Function
    End If

    '����Ϊҵ��ϵͳ�д�����(DN)��
40     strCertDN = mobjHNCACTRL.GetCertInfo(0, "")
            
    '����Ϊҵ��ϵͳ�д�����(DN)������ȡCN��

50      intCnIndex = InStr(strCertDN, "CN=") '��ȡ"CN="��λ��
51      strCertCN = Mid(strCertDN, intCnIndex) '��CN��ǰ����Ϣȥ��
52      intCnIndex = InStr(strCertCN, ",") '��1����������λ��
53      If (intCnIndex <> -1) Then
54          strCertCN = Mid(strCertCN, 1, intCnIndex - 1)
55      End If
56      strCertCN = Mid(strCertCN, 4)
        '֤��CN�� ����
57      strName = Trim(strCertCN)

        '��ȡ֤�����к�,SN��Ŀ��ÿ��֤��Ψһ��ʶ
60      strCertSn = mobjHNCACTRL.GetCertInfo(2, "")
        '��ȡ�ͻ���ǩ��(��Կ)֤��Base64����
80      strCertPUB = mobjHNCACTRL.GetCertInfo(8, "")
        '��ȡ��չ�������Ŀȷ����չ��Ĳ�����
81      strPara = IIf(mbytFunc = 0, "1.2.86.11.7.12", "1.2.86.11.7.11")
82      strUserID = mobjHNCACTRL.GetCertInfo(7, strPara)  '����֤��
        If mbytFunc = 0 Then
90          strPic = mobjHNCACTRL.LoadSealFromKey(strCertSn, "111111")   '����ӡ�����ݵ�BASE64
        Else
91          strPara = "����"
92          strSign = mobjHNCACTRL.SignSealForSM2(strCertSn, strPara)
93          intRet = mobjHNCACTRL.VerifySealForSM2(strSign, strPara)
94          strPic = mobjHNCACTRL.GetSealForSM2(strSign, intRet)
        End If
        If strPic <> "" Then
100        If mbytFunc = 0 Then strPic = mobjHNCACTRL.GetBmpFromSealData(strPic, 2) '����ӡ��ͼƬ���ݵ�BASE64
120        strPicPath = SaveBase64ToFile("bmp", strCertSn, strPic) '����ӡ��ͼƬ���ݵ�BASE64ת����ͼƬ�ļ�������ͼƬλ��
        Else
            MsgBoxEx "��ȡͼƬʧ�ܣ������Ƿ����key�̣�", vbInformation, gstrSysName
            Exit Function
        End If
    
130     GetCertList = True
        Exit Function
errH:
500     MsgBoxEx "��ȡ֤����Ϣʧ�ܣ�" & vbNewLine & "��" & CStr(Erl()) & "�� " & Err.Description, vbInformation, gstrSysName
End Function

Private Function GetCertLogin(ByVal strCertPUB As String, ByVal strCertDN As String) As Boolean
    Dim blnRet As Boolean
    Dim strSignData As String
    Dim strUrl As String, strAPPCode As String, strRand As String, strMsg As String
    Dim strTSIP As String, strUserName As String, strUserPWD As String
    Dim intRet As Integer, intPort As Integer, intSSLPort As Integer
    Dim arrTmp As Variant
        On Error GoTo errH
1000
         arrTmp = Split(gstrPara, "|")
        'gstrPara= http://218.28.16.104:7080/CAServer/servlet/CertChechServlet|218.28.16.104|8080|443|ʱ�����0-������;1-���ã�|ǩ���㷨(0-RSA\1-SM2)"
        intRet = -1
        strUrl = CStr(Trim(arrTmp(0)))
        strAPPCode = "1101"
        strTSIP = CStr(Trim(arrTmp(1)))
        intPort = CInt(arrTmp(2))
        intSSLPort = CInt(arrTmp(3))
        strUserName = "abc"
        strUserPWD = "123"
        
1100
        '��ȡ�����
        Randomize
        strRand = CStr(Int((9001) * Rnd + 1000))


        '�����ǩ��
        If mbytFunc = 0 Then
1101        strSignData = mobjHNCACTRL.AttachSign(strCertDN, strRand)
        Else
1102        strSignData = mobjHNCACTRL.SM2Sign("", strRand)
        End If
        
        If strSignData = "" Then
            strMsg = mobjHNCACTRL.GetErrorMessage()  ' ��ȡ������Ϣ
            MsgBoxEx "�����ǩ��ʧ��:" + strMsg, vbOKOnly, gstrSysName
            Exit Function
        End If
1200
        strMsg = ""
        '���ͷ�����������֤
        intRet = doPostEx(strUrl, strAPPCode, strRand, strSignData, strCertPUB, "GBK")

        strMsg = ""

        If (intRet = 0) Then
            blnRet = True
        ElseIf (intRet < 0) Then
            '�������ǩʧ��
            blnRet = False
            strMsg = "��¼��������ǩʧ�ܣ�"
        Else
            blnRet = True
            strMsg = "��¼��������ǩ�ɹ�������֤�黹ʣ��" & intRet & " ����ڣ�"
        End If
        If strMsg <> "" Then
           MsgBoxEx strMsg, vbInformation, gstrSysName
        End If
        '��֤ʧ��ֱ���˳�
        If Not blnRet Then Exit Function
1300
        'ʱ�������ϵͳ����
        If mblnTs Then
            strMsg = ""
            intRet = mobjHNCATS.SetSystemParameter(strTSIP, intPort, intSSLPort)
            If (intRet = 0) Then        'MessageBox.Show("���óɹ���")
                
                '2.���ù�ϣ�㷨�����������Ĭ����sha1
                Call mobjHNCATS.SetMD(1)
                
                '3.�û���¼
                If (mobjHNCATS.ISLogin = 1) Then
                    blnRet = True
                Else
                    intRet = -1
                    intRet = mobjHNCATS.LoginByUsername(strUserName, strUserPWD)    '
                    If (intRet = 0) Then
                        blnRet = True
                    Else
                        blnRet = False
                        strMsg = "��¼ʱ���������ʧ�ܣ��������Ϊ��" & intRet
                    End If
                End If
            Else
                strMsg = "����ʱ�����������ϵͳ����ʧ�ܣ�������룺" & intRet
                blnRet = False
            End If
                
            If strMsg <> "" Then
               MsgBoxEx strMsg, vbInformation, gstrSysName
            End If
        End If
1600
    GetCertLogin = blnRet
    Exit Function
errH:
    MsgBoxEx "��¼��������֤ʧ�ܣ�" & vbNewLine & "��" & CStr(Erl()) & "�� " & Err.Description, vbInformation, gstrSysName
End Function

Public Function HNCA_Initialize() As Boolean
    '����:��������CA�ؼ�����
    
    Dim intRet As Integer
    
    On Error GoTo errH
    
1000
        Set mobjHNCACTRL = CreateObject("HNCACTRL.HNCACtrlCtrl.1") 'ǩ������
1001
        Set mobjHNCATS = CreateObject("TSAACTIVEXVER4.TsaActiveXver4Ctrl.1")  'ʱ�������
1100
        gstrPara = gobjComLib.zlDatabase.GetPara(90000, glngSys) '��ȡURL ������
        'gstrPara = "http://218.28.16.104:7080/CAServer/servlet/CertChechServlet|218.28.16.104|8080|443|ʱ���(0-������/1-����)|ǩ���㷨(0-RSA\1-SM2)"
        If gstrPara = "" Then
            MsgBoxEx "û������ǩ����������ַ���뵽���õ���ǩ���ӿڴ����á�", vbInformation, gstrSysName
            Exit Function
        Else
            If UBound(Split(gstrPara, "|")) <> 5 Then
               MsgBoxEx "ǩ����������ַ���������뵽���õ���ǩ���ӿڴ����á�" & _
                    vbCrLf & vbTab & "����ֵ��ʽ""������URL|TSIP|TSPORT|SSLPORT|�Ƿ�����ʱ���(0-������/1-����)|ǩ���㷨(0-RSA\1-SM2)""", vbInformation, gstrSysName
                Exit Function
            End If
        End If
        mblnTs = Val(Split(gstrPara, "|")(4)) = 1
        mbytFunc = CByte(Split(gstrPara, "|")(5))
1500
        HNCA_Initialize = True
        Exit Function

errH:
     MsgBoxEx "�����ӿڲ���ʧ�ܣ�" & vbNewLine & "��" & CStr(Erl()) & "�� " & Err.Description, vbQuestion, gstrSysName
End Function

Private Sub Class_Terminate()
    Set mobjHNCACTRL = Nothing
    Set mobjHNCATS = Nothing
End Sub

Public Function HNCA_GetPara() As Boolean
'���÷�������ַ
    
    On Error GoTo errH
    gstrPara = gobjComLib.zlDatabase.GetPara(90000, glngSys)  '��ȡURLs �̶���ȡZLHIS ϵͳĬ��100
    'gstrPara = "http://218.28.16.104:7080/CAServer/servlet/CertChechServlet|218.28.16.104|8080|443|ʱ���(0-������/1-����)|ǩ���㷨(0-RSA\1-SM2)"
    If gstrPara = "" Then gstrPara = "http://218.28.16.104:7080/CAServer/servlet/CertChechServlet|218.28.16.104|8080|443|0|0"
    If gstrPara <> "" Then
        gudtPara.strSignURL = gstrPara
    End If
    Exit Function
errH:
    MsgBoxEx "��ȡ����ʧ�ܣ�" & vbNewLine & "��" & CStr(Erl()) & "�� " & Err.Description, vbInformation, gstrSysName
End Function

Public Function HNCA_SetParaStr() As String
    HNCA_SetParaStr = gudtPara.strSignURL
End Function

