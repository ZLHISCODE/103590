Attribute VB_Name = "mdlLNCASY"
Option Explicit

Private mobjUSBKEY As Object     '����ʡ����ǩ�� ���� KEYUSBKEYACTIVE.USBKeyActiveCtrl.1
Private mobjMSScriptCtl As Object    'MSScriptControl.ScriptControl.1 ΢���ṩ�ű��ؼ� �õ�javaScript��encodeURI������ȡURL��
Private mblnInit As Boolean
Private mstrLastPwd As String          '�������������
Private mintLogin As Integer
Public gobjLNCAPenSign As Object '����CA��дǩ������
'20170817 SM2�㷨����  �����κ�
Private mbytModel           As Byte             '0-RSA�㷨;1-SM2�㷨
Private mobjKeyManager      As Object           '֤��������
Private mobjCert            As Object           '֤�����
Private mobjKeyStore        As Object           'UKey������KeyStore
Private mobjKeySealArray    As Object
Private mobjKeySeal         As Object           'ǩ����
Private mobjKeyGateOper     As Object
Private mobjKeyDetector     As Object           'JHKey.KeyDetector.1.1
Private Enum E_Model
    E_RSA = 0
    E_SM2 = 1
End Enum

Public Function LNCA_Initialize() As Boolean
        '����:��������CA�ؼ�����
    
        Dim varTmp As Variant
    
        On Error GoTo errH
   
100         If mblnInit Then LNCA_Initialize = True: Exit Function
        
102         gstrPara = gobjComLib.zlDatabase.GetPara(90000, glngSys) '��ȡURL ������
            'gstrPara = "http://218.25.86.214:2010/ssoworker"  '���Ե�ַ
104         If gstrPara = "" Then
106             MsgBoxEx "û��������֤��������ַ���뵽�������������á�������:" & vbCrLf & vbTab & "ϵͳ��100,������90000" & _
                        vbCrLf & vbTab & "����ֵ��ʽ""http://218.25.86.214:2010/ssoworker""", vbInformation, gstrSysName
                Exit Function
            End If
108         varTmp = Split(gstrPara, G_STR_SPLIT)
110         gudtPara.strSignURL = varTmp(0)
112         If UBound(varTmp) >= 1 Then
114             mbytModel = Val(varTmp(1) & "")
            Else
116             mbytModel = E_RSA
            End If
        
118         If UBound(varTmp) >= 2 Then
120             gudtPara.strSIGNIP = varTmp(2)
            Else
122             gudtPara.strSIGNIP = ""
            End If
        
124         If mbytModel = E_RSA Then   'RSA
126             Set mobjUSBKEY = CreateObject("USBKEYACTIVE.USBKeyActiveCtrl.1") 'ǩ������
128             Set mobjMSScriptCtl = CreateObject("MSScriptControl.ScriptControl.1")
130             mobjMSScriptCtl.Language = "JavaScript"
            Else                    'SM2
132             Set mobjKeyManager = CreateObject("JHKey.KeyManager.1")
134             Set mobjCert = CreateObject("JHKey.Cert.1")
136             Set mobjKeyStore = CreateObject("JHKey.KeyStore.1")
138             Set mobjKeySealArray = CreateObject("JHKey.SealArray.1")
140             Set mobjKeySeal = CreateObject("JHKey.Seal.1")
142             Set mobjKeyGateOper = CreateObject("JHKey.GateOper.1")
144             Set mobjKeyDetector = CreateObject("JHKey.KeyDetector.1")
146             Call mobjKeyGateOper.SetTimeout(10)
148             Call mobjKeyGateOper.SetURL(gudtPara.strSignURL)
            
            End If
        
150         gstrLogins = ""
152         mblnInit = True
154         LNCA_Initialize = True
            Exit Function
errH:
156         LogWrite "LNCA_Initialize", "�����ӿڲ���ʧ�ܣ�" & vbNewLine & "��" & CStr(Erl()) & "�� " & Err.Description
End Function

Public Function LNCA_RegCert(arrCertInfo As Variant, Optional ByVal strUserID As String) As Boolean
'���ܣ��ṩ��HIS���ݿ���ע�����֤��ı�Ҫ��Ϣ,�����·��Ż����֤��,,��Ҫ����USB-Key
'����:strUserID-���֤��
'���أ�arrCertInfo��Ϊ���鷵��֤�������Ϣ
'      0-ClientSignCertCN:�ͻ���ǩ��֤�鹫������(����),����ע��֤��ʱ������֤���
'      1-ClientSignCertDN:�ͻ���ǩ��֤������(ÿ��Ψһ)
'      2-ClientSignCertSN:�ͻ���ǩ��֤�����к�(ÿ֤��Ψһ)
'      3-ClientSignCert:�ͻ���ǩ��֤������
'      4-ClientEncCert:�ͻ��˼���֤������
'      5-ǩ��ͼƬ�ļ���,�մ���ʾû��ǩ��ͼƬ
        
        Dim strCertSn As String, strCertUserName As String, strCertDN As String
        Dim strSigCert As String, i As Integer
        Dim strPic As String
        Dim strCertUserID As String
10      On Error GoTo errH
    
20      For i = LBound(arrCertInfo) To UBound(arrCertInfo)
30          arrCertInfo(i) = ""
40      Next
    
50      If GetCertList(strCertUserName, strCertSn, strSigCert, strCertDN, strPic, strCertUserID) Then
60          If UCase(strCertUserID) <> UCase(strUserID) And strUserID <> "" Then
70              MsgBoxEx "�û����֤�ţ�" & _
                        vbCrLf & vbTab & "��" & UCase(strUserID) & "��" & vbCrLf & _
                        "��ǰ֤��Ψһ��ʶ:" & _
                        vbCrLf & vbTab & "��" & UCase(strCertUserID) & "��" & vbCrLf & _
                        "�û����֤���뵱ǰ֤��Ψһ��ʶ�����,����ע�ᣡ", vbInformation, gstrSysName
80              Exit Function
90          End If
100         arrCertInfo(0) = strCertUserName
110         arrCertInfo(1) = strCertDN
120         arrCertInfo(2) = strCertSn
130         arrCertInfo(3) = strSigCert
140         If strPic <> "" Then
150             arrCertInfo(5) = SaveBase64ToFile("gif", strCertSn, strPic)
160         End If
170         LNCA_RegCert = True
180     End If

190     Exit Function
errH:
200     MsgBoxEx "��ȡ֤����Ϣʧ�ܣ�" & vbNewLine & "��" & CStr(Erl()) & "�� " & Err.Description, vbExclamation, gstrSysName

End Function

Public Function LNCA_CheckCert(ByRef blnReDo As Boolean) As Boolean
'���ܣ���ȡUSB�����豸��ʼ������¼
    Dim strCertName As String
    Dim strCertSn As String
    Dim strCertUserID As String    '�������֤����Ϣ
    Dim strDate As String
    Dim udtUser As USER_INFO
    Dim strCertID As String
    Dim strCert As String
    Dim blnOk As Boolean
    
    On Error GoTo errH

     '��ȡ֤����Ϣͬʱ���Key���Ƿ����
    If Not GetCertList(, strCertSn, , , , strCertUserID) Then Exit Function
        
    'δע���ڵ�ǰ�û����µ�Key
    If mUserInfo.strUserID = "" Then
        MsgBoxEx "�������֤��Ϊ��,����ϵ����Ա����Ա������¼�룡", vbOKOnly + vbInformation, gstrSysName
        Exit Function
    ElseIf UCase(strCertUserID) <> UCase(mUserInfo.strUserID) Then
        MsgBoxEx "�������֤�ţ�" & _
                   vbCrLf & vbTab & "��" & mUserInfo.strUserID & "��" & vbCrLf & _
                   "��ǰ֤��Ψһ��ʶ:" & _
                   vbCrLf & vbTab & "��" & strCertUserID & "��" & vbCrLf & _
                   "�û����֤���뵱ǰ֤��Ψһ��ʶ�����,����ʹ�ã�", vbInformation, gstrSysName
        Exit Function
    End If
    'CA�״�ǩ��ʱ���Զ����������
    '��¼��֤
    If InStr(gstrLogins & "|", "|" & strCertSn & "|") > 0 Then '�״���֤ͨ�����´β��ڼ�����֤
        blnOk = True
    Else
        If Not GetCertList(, , strCert, , , , strDate, 1) Then Exit Function
        If Not GetCertLogin(strCert, strDate) Then
            blnOk = False
        Else
            blnOk = True
            If InStr(gstrLogins & "|", "|" & strCertSn & "|") = 0 Then gstrLogins = gstrLogins & "|" & strCertSn
        End If
    End If

    If blnOk And mUserInfo.strCertSn <> strCertSn Then
        '�ж��Ƿ���Ҫ����ע��֤��
        '��ʱ�����ŵ�ע��ʱ��ȡ��ǩ��ͼƬ�Ĵ���Ĵ���
        If Not GetCertList(strCertName, , udtUser.strCert, udtUser.strCertDN, udtUser.strPicCode, , strDate, 1) Then Exit Function
        udtUser.strName = strCertName
        udtUser.strSignName = strCertName
        udtUser.strUserID = strCertUserID
        udtUser.strCertSn = strCertSn
        If IsUpdateRegCert(udtUser, strDate, blnReDo) Then
            blnOk = True
        Else
            blnOk = False
        End If
    End If
    LNCA_CheckCert = blnOk
    Exit Function
errH:
     MsgBoxEx "���USBKEYʧ�ܣ�" & vbNewLine & "��" & CStr(Erl()) & "�� " & Err.Description, vbInformation, gstrSysName
End Function

Private Function GetCertList(Optional ByRef strName As String = "-1", Optional ByRef strCertSn As String = "-1", Optional ByRef strCert As String = "-1", _
                Optional ByRef strCertDN As String = "-1", Optional strPic As String = "-1", _
                Optional strUserID As String = "-1", Optional strEndDate As String = "-1", Optional ByVal bytMode As Byte) As Boolean
          '����:��ȡ֤����Ϣ
          '-����
          '    strName ֤�����������
          '   strCertSN ֤��Ψһ��ʶ
          '   strCert ǩ��֤��
          '   strCertDN ֤��������Ϣ  ֤��ע���õ�
          '   strPic      ֤��ͼƬ
          '   bytMode =0 ��ʼ������;1-����ʼ��
          Dim strMsg As String
          Dim lngRet As Long
          Dim strPIN As String
          Dim i As Integer
          Dim indexS As Integer
          
10        On Error GoTo errH
20        If Not LNCA_Initialize() Then Exit Function
          
30        If mbytModel = E_RSA Then
              '��������
40            If bytMode = 0 Then
50                If mstrLastPwd <> "" Then strPIN = mstrLastPwd
60                If strPIN = "" Then
70                    If Not frmPassword.ShowMe(strPIN) Then Exit Function
80                End If
                  
90                If strPIN = "" Then
100                  MsgBoxEx "������֤�����룡", vbOKOnly + vbInformation, gstrSysName
110                  Exit Function
120               Else
130                   If mintLogin >= 8 Then
140                       MsgBoxEx "�Ѿ�������" & mintLogin & "�δ������룬������������������", vbOKOnly + vbInformation, gstrSysName
150                       Exit Function
160                   End If
170                   On Error Resume Next
180                   lngRet = mobjUSBKEY.MNGInit(strPIN)
190                   If Err.Number <> 0 Then
200                       MsgBoxEx "��������KEY�̣�", vbOKOnly + vbInformation, gstrSysName
210                       Exit Function
220                   End If
230                   Err.Clear: On Error GoTo 0
                      
240                   If lngRet = 0 Then
250                      mstrLastPwd = strPIN
260                   Else
270                       mintLogin = mintLogin + 1
280                       MsgBoxEx "֤��������ܲ���ȷ�����Ѿ�������" & mintLogin & "�����룬����������" & 8 - mintLogin & "��!", vbOKOnly + vbInformation, gstrSysName
290                       mstrLastPwd = ""
300                       Exit Function
310                   End If
320               End If
330           End If
340           On Error GoTo errH
              
350           If bytMode = 0 Then Call mobjUSBKEY.MNGLogin
                
360           If strCertSn <> "-1" Then strCertSn = mobjUSBKEY.MNGGetSignCertSN 'Ψһ��ʶ��
370           If strCert <> "-1" Then strCert = mobjUSBKEY.MNGGetSignCert()    '��ȡǩ��֤��
380           If strName <> "-1" Then strName = mobjUSBKEY.MNGGetSignCertCN()          ''��ȡ����
390           If strCertDN <> "-1" Then strCertDN = mobjUSBKEY.MNGGetSignCertDN      '����
              
400           If strPic <> "-1" Then
410               strPic = mobjUSBKEY.MNGGetSESCount '����ǩ��|������
420               If strPic = "" Then
430                   strMsg = "��ȡǩ��ͼƬʧ�ܣ�"
440                   GoTo msgINFO
450               End If
460               strPic = Split(strPic, "|")(0)
470               strPic = mobjUSBKEY.MNGReadSESealByLabelEx(strPic)         '��ȡǩ��ͼƬBASE64
480               If strPic = "" Then
490                   strMsg = "��ȡǩ��ͼƬʧ�ܣ�"
500                   GoTo msgINFO
510               End If
520           End If
530           If strUserID <> "-1" Then
540               strUserID = mobjUSBKEY.MNGGetSignCertDN_OUa   '���֤��
550               If strUserID <> "" And Len(strUserID) >= 18 Then
560                   strUserID = Right(strUserID, 18)
570               Else
580                   strMsg = "��ȡ���֤����ʧ�ܣ�"
590                   GoTo msgINFO
600               End If
610           End If
              '��ȡ�ͻ���֤����Ч�ڽ�ֹʱ��
620           If strEndDate <> "-1" Then
630               strEndDate = mobjUSBKEY.MNGGetSignCertEndValidityTime()
640               strEndDate = CDate(Format(strEndDate, "YYYY-MM-DD HH:MM:SS"))
650           End If
660       ElseIf mbytModel = E_SM2 Then
670           lngRet = mobjKeyDetector.EnumUKey()
680           If lngRet = 0 Then
690               MsgBoxEx "��������KEY�̣�", vbOKOnly + vbInformation, gstrSysName
700               Exit Function
710           ElseIf lngRet = 1 Then
720               Call mobjKeyManager.EnumKeyStore               '����EnumKeyStoreö�ٳ�����֤��
                  'Ĭ��ѡȡSM2��keyǩ��
730               lngRet = mobjKeyManager.GetCertCount()
740               For i = 0 To lngRet - 1 '����GetCertCount�õ�ȫ��֤��ĸ���
750                   mobjCert.SetCert (mobjKeyManager.GetCert(i)) '��ѯ������е�֤��
760                   If mobjCert.CertUsage = 1 Then 'CertUsage == 1:ǩ��֤��,2:����֤�顣CertType == 1:RSA,2:SM2
770                       Call mobjKeyManager.InitKeyStoreByIndex(i, mobjKeyStore)
780                       indexS = i
790                       Exit For
800                   End If
810               Next
820           Else
830               Call mobjKeyManager.EnumKeyStore               '����EnumKeyStoreö�ٳ�����֤��
                  '��ʾ֤���б���������û�ѡȡ֤�顣
                  '��һ����������ʾ��ʾ��֤�����ͣ�1��RSA��2��SM2��3��RSA��SM2����ʾ;
                  '�ڶ���������ʾ֤����;��1��ǩ����2�����ܣ�3��ǩ�����ܾ���
                  '0-�û�ѡ����֤��;1-δѡ��֤��
840               lngRet = mobjKeyManager.ShowCertsDlg(3, 1)
850               If lngRet <> 0 Then
860                   strMsg = "δѡ���κ�֤�飡"
870                   GoTo msgINFO
880               End If
890               Call mobjKeyManager.GetSelectedCert(mobjCert)    '�����û���ѡ��õ�ָ����֤��
                  '�����û�ѡ���֤�飬��ʼ��UKey������KeyStore ǩ��ʱֱ�ӵ���ǩ���ӿ�
900               Call mobjKeyManager.InitKeyStore(mobjKeyStore)
                  'mobjCert.CertType֤�����ͣ�1:RSA��2:SM2
910               indexS = mobjKeyManager.getDlgSelId()
920           End If
930           strName = mobjCert.CertCN
940           If strCertSn <> "-1" Then strCertSn = mobjCert.CertSN
950           If strCertDN <> "-1" Then strCertDN = mobjCert.CertSubject
960           If strCert <> "-1" Then strCert = mobjCert.Body
970           If strUserID <> "-1" Then strUserID = Mid(mobjCert.CertOuA, 2)
980           If strEndDate <> "-1" Then
990               strEndDate = mobjCert.CertNotAfter
1000          End If
              
1010          If mobjCert.CertType = 2 Then  'SM2 ����������
                  '��������
1020              If mstrLastPwd <> "" Then strPIN = mstrLastPwd
1030              If strPIN = "" Then
1040                  If Not frmPassword.ShowMe(strPIN) Then Exit Function
1050              End If
                      
1060              If strPIN = "" Then
1070                 MsgBoxEx "������֤�����룡", vbOKOnly + vbInformation, gstrSysName
1080                 Exit Function
1090              Else
1100                  Call mobjKeyStore.SetWorkPin(strPIN)  '���Ĭ��PING�룬ʹPIN����ٵ������ﵽ��Ĭ������Ŀ�ģ�ֻ��SM2Key��Ч��RSA��PIN����ɸ��������Ҹ���ʵ�֣��޷����ƣ�
1110                  lngRet = mobjKeyStore.SignData("123")    '0-�ɹ�;��0-ʧ��
1120                  If lngRet = 0 Then
1130                      mstrLastPwd = strPIN
1140                  Else
1150                      mstrLastPwd = ""
1160                      Exit Function
1170                  End If
1180              End If
1190          End If
            
              'ȡӡ�£�������ѡ֤�飬��ѯ��֤������Key�������ӡ�£�����ӡ������SealArray
1200          If strPic <> "-1" Then
                  'Call mobjKeyManager.InitSealStore(mobjKeySealArray)
1210              Call mobjKeyManager.InitSealArrByIndex(indexS, mobjKeySealArray) '���ҽԺ��������CAʱ����
1220              lngRet = mobjKeySealArray.GetSealCount()         '�õ�ӡ����������ӡ�¸���
1230              If lngRet = 0 Then
1240                  strMsg = "��ѡ֤���޶�Ӧ��ӡ�£�"
1250                  GoTo msgINFO
1260              End If
1270              For i = 0 To lngRet - 1
1280                  Call mobjKeySealArray.GetSeal(i, mobjKeySeal)     '��ӡ��������ȡ��ӡ��
1290                  strPic = mobjKeySeal.getpic()              '�õ�ӡ��ͼƬ��base64����
1300                  If strPic <> "" Then
1310                      Exit For
1320                  Else
1330                      strMsg = "��ȡǩ��ͼƬʧ�ܣ�"
1340                      GoTo msgINFO
1350                  End If
1360              Next
1370          End If
1380      End If
1390      GetCertList = True
1400      Exit Function

msgINFO:
1410    If strMsg <> "" Then
1420        MsgBoxEx strMsg, vbInformation, gstrSysName
1430        Exit Function
1440    End If
errH:

1450  MsgBoxEx "��ȡ֤����Ϣʧ�ܣ�" & vbNewLine & "��" & CStr(Erl()) & "�� " & Err.Description, vbInformation, gstrSysName
End Function

Private Function GetCertLogin(ByVal strSignCert As String, ByVal strEnd As String) As Boolean
    Dim strSignResult As String
    Dim strToken As String, strParameter As String, strRandom As String
    Dim strRet As String
    Dim datEnd As Date
    Dim intDay As Integer
    Dim lngRet As Long
    
    On Error GoTo errH
    If mbytModel = E_RSA Then
        '��ȡ�����
        strRandom = HttpPost(gudtPara.strSignURL, "cmd=getrand", responseText)   '��ȡ���������ֵ: {"ret":1,"errinfo":"","rand":"3XO9JCXVJ6LF05M51165"}
        strRandom = GetSubString(strRandom, "rand")
        
        '�����ǩ��
        strSignResult = mobjUSBKEY.MNGSignData(strRandom, Len(strRandom))      '�ؼ�ǩ��
        If strSignResult = "" Then
            MsgBoxEx "�����ǩ��ʧ�ܣ�", vbInformation + vbOKOnly, gstrSysName
            Exit Function
        End If
    
        '�����ǩ����֤
        strParameter = "cmd=sm2certlogin" & "&rand=" & EnCodeURL(strRandom) & "&cert=" & EnCodeURL(strSignCert) & "&signed=" & EnCodeURL(strSignResult)  '��������֤KEY ǩ�����
        strToken = HttpPost(gudtPara.strSignURL, strParameter, responseText)  '
        strRet = GetSubString(strToken, "ret")
        If strRet <> "1" Then
            MsgBoxEx "֤���¼��֤ʧ�ܣ�", vbInformation + vbOKOnly, gstrSysName
            Exit Function
        End If
    ElseIf mbytModel = E_SM2 Then
'        Call mobjKeyStore.SetWorkPin(mstrLastPWD)
'       �ظ��������뵼�µ�½��֤����ֵ=9 ��ʾʧ�� ��ע��
        lngRet = mobjKeyGateOper.ReqCertLogin(mobjKeyStore, "123")  '֤���½
        If lngRet <> 0 Then
            MsgBoxEx "֤���¼��֤ʧ�ܣ�" & vbCrLf & "��������:" & mobjKeyGateOper.GetLastErrText & "����ֵ:" & lngRet, vbInformation + vbOKOnly, gstrSysName
            Exit Function
        End If
    End If
    
    '��ȡ�ͻ���֤����Ч�ڽ�ֹʱ��
    datEnd = CDate(strEnd)
    '��֤�ͻ���֤����Ч��ʣ������
    intDay = Int(CDbl(datEnd) - CDbl(Now))
    
    If (intDay <= 30 And intDay > 0 And Not gblnShow) Then
        MsgBoxEx "����֤�黹��" & intDay & "����ڡ�", vbInformation + vbOKOnly, gstrSysName
        gblnShow = True
        GetCertLogin = True
    ElseIf (intDay <= 0) Then
        MsgBoxEx "����֤���ѹ��� " & Abs(intDay) & " �졣", vbInformation + vbOKOnly, gstrSysName
        GetCertLogin = False
    End If
        
    GetCertLogin = True
    Exit Function
errH:
    MsgBoxEx "��¼��������֤ʧ�ܣ�" & vbNewLine & "��" & CStr(Erl()) & "�� " & Err.Description, vbInformation, gstrSysName
End Function

Public Function LNCA_Sign(ByVal strSource As String, ByRef strSignData As String, ByRef strTimeStamp As String, _
        ByRef strTimeStampCode As String, ByRef blnReDo As Boolean, ByVal blnCheck As Boolean) As Boolean
'ǩ��
        Dim strParameter As String, strMsg As String, strDate As String, strRet As String
        Dim intRet As Integer
        Dim blnRet As Boolean
        Dim datTime As Date
        
10      On Error GoTo errH
        
20      If Not blnCheck Then
30          blnCheck = LNCA_CheckCert(blnReDo)
40          If blnReDo Then Exit Function
50      End If
60      If blnCheck Then
70          If mbytModel = E_RSA Then
                '��֤��ǰUSB�Ƿ���ǩ���û��ģ�����ȡǩ��֤��
80              strSource = EncodeBase64String(strSource) 'Դ���а��������ַ�����Ҫ����ת��
90              strSignData = mobjUSBKEY.MNGSignData(strSource, Len(strSource))        '�ؼ�ǩ��
100             If strSignData <> "" Then
                    '�洢���ط�����
110                 datTime = gobjComLib.zlDatabase.Currentdate()
120                 strDate = Format(datTime, "yyyyMMddhhmmss")
130                 strTimeStamp = Format(datTime, "yyyy-MM-dd HH:mm:ss")
140                 strParameter = "cmd=insert_sign_record" & "&appid=" & EnCodeURL("100") & "&docid=" & EnCodeURL("100") & "&docname=" & _
                    EnCodeURL("ZLHIS") & "&textinfo=" & EnCodeURL(strSource) & "&signdata=" & EnCodeURL(strSignData) & "&signcert=" & _
                    EnCodeURL(mUserInfo.strCert) & "&signdate=" & EnCodeURL(strDate)
150                 strRet = HttpPost(gudtPara.strSignURL, strParameter, responseText)
160                 blnRet = GetSubString(strRet, "ret") = "1"
170                 If Not blnRet Then strMsg = "ǩ��ʧ�ܣ�"
180             Else
190                 strMsg = "ǩ��ʧ�ܣ�"
200                 blnRet = False
210             End If
220         ElseIf mbytModel = E_SM2 Then
                'mobjKeyStore��LNCA_CheckCert �Ѿ�ʵ����
230             Call mobjKeyStore.SetWorkPin(mstrLastPwd)
240             intRet = mobjKeyStore.SignData(strSource)   '0-�ɹ�;��0-ʧ��
250             If intRet = 0 Then
260                 strSignData = mobjKeyStore.GetSignData()               '�õ�ǩ������
270             Else
280                 strMsg = "ǩ��ʧ�ܣ�"
290                 blnRet = False
300             End If
                'ǩ����洢���ط�����
310             datTime = gobjComLib.zlDatabase.Currentdate()
320             strDate = Format(datTime, "yyyyMMddhhmmss")
330             strTimeStamp = Format(datTime, "yyyy-MM-dd HH:mm:ss")
340             intRet = mobjKeyGateOper.ReqUploadMedRecord("01", "appid", "docid", "docname", strSource, strSignData, mUserInfo.strCert, strDate)
350             blnRet = (intRet = 0)
360             If Not blnRet Then strMsg = "ǩ��ʧ�ܣ�"
370         End If
380     Else
390         strMsg = "ǩ��ʧ�ܣ�"
400         blnRet = False
410     End If
420     If strMsg <> "" Then
430         MsgBoxEx strMsg, vbInformation, gstrSysName
440     End If
                
450     LNCA_Sign = blnRet
460     Exit Function
errH:
470     MsgBoxEx "ǩ��ʧ�ܣ�" & vbNewLine & "��" & CStr(Erl()) & "�� " & Err.Description, vbQuestion, gstrSysName
End Function

Public Function LNCA_VerifySign(ByVal strCert As String, ByVal strSignData As String, ByVal strSource As String) As Boolean
'��֤ǩ��
'
    Dim strParameter As String
    Dim strRet As String
    Dim blnRet As Boolean
    Dim strMsg As String
    Dim lngRet As Long
    
    On Error GoTo errH
    If mbytModel = E_RSA Then
        strSource = EncodeBase64String(strSource)
        strCert = strCert & "123"
        strParameter = "cmd=verifysm2" & "&text=" & EnCodeURL(strSource) & "&cert=" & EnCodeURL(strCert) & "&signed=" & EnCodeURL(strSignData)
        strRet = HttpPost(gudtPara.strSignURL, strParameter, responseText)
        blnRet = GetSubString(strRet, "ret") = "1"    '����ֵ=1��֤ǩ���ɹ�
    ElseIf mbytModel = E_SM2 Then
        '���� ǩ�����;ǩ��ԭ��;Ԥ��;������֤��
        lngRet = mobjKeyGateOper.ReqVerifySig(strSignData, strSource, 0, strCert)
        blnRet = lngRet = 0
    End If
    If blnRet Then    '��֤ǩ��ʧ��
        strMsg = "��֤�ɹ����õ���ǩ��������Ч��"
    Else
        strMsg = "��ǩʧ�ܣ�"
    End If
    
    If strMsg <> "" Then
        MsgBoxEx strMsg, vbInformation, gstrSysName
    End If
        
    LNCA_VerifySign = blnRet
    
    Exit Function
errH:
104     MsgBoxEx "��֤ǩ��ʧ�ܣ�" & vbNewLine & "��" & CStr(Erl()) & "�� " & Err.Description, vbQuestion, gstrSysName
End Function

Public Sub LNCA_UnLoadObj()
    If mbytModel = E_RSA Then
        Set mobjUSBKEY = Nothing
        Set mobjMSScriptCtl = Nothing
    Else
        Set mobjCert = Nothing
        Set mobjKeyGateOper = Nothing
        Set mobjKeyManager = Nothing
        Set mobjKeySeal = Nothing
        Set mobjKeySealArray = Nothing
        Set mobjKeyStore = Nothing
        Set mobjKeyDetector = Nothing
    End If
    
    mblnInit = False
End Sub

Private Function EnCodeURL(ByVal strUrl As String) As String
'����:�������ַ�����UTF���뷽ʽת����ʮ�����Ƶ�ת������
'˵����encodeURI-javaScript��������� ASCII ��ĸ�����ֽ��б��룬Ҳ�������Щ ASCII �����Ž��б��룺 - _ . ! ~ * ' ( )
    Dim i As Long
    Dim strChar As String
    Dim intAsc As Integer
    Dim strRet As String
    
    For i = 1 To Len(strUrl)
        strChar = Mid(strUrl, i, 1)
        intAsc = Asc(strChar)
        If intAsc >= 0 And intAsc <= 127 Then
           strChar = "%" & Hex(intAsc)
        Else
            strChar = mobjMSScriptCtl.Eval("encodeURI(""" & strChar & """)")
        End If
        strRet = strRet & strChar
    Next
    
    EnCodeURL = strRet
End Function

Public Function LNCA_GetPara() As Boolean
'���÷�������ַ
    Dim arrTmp As Variant
    
    On Error GoTo errH
    gstrPara = gobjComLib.zlDatabase.GetPara(90000, glngSys)
    If gstrPara = "" Then gstrPara = "http://218.25.86.214:2010/ssoworker"
    arrTmp = Split(gstrPara, G_STR_SPLIT)
    If UBound(arrTmp) > 0 Then
        gudtPara.strSignURL = arrTmp(0)
        gudtPara.bytSignVersion = Val(arrTmp(1) & "")
        If UBound(arrTmp) > 1 Then
            gudtPara.strSIGNIP = arrTmp(2)
        End If
    Else
        gudtPara.strSignURL = arrTmp(0)  '
        gudtPara.strSIGNIP = ""
    End If

    Exit Function
errH:
    MsgBoxEx "��ȡ����ʧ�ܣ�" & vbNewLine & "��" & CStr(Erl()) & "�� " & Err.Description, vbInformation, gstrSysName
End Function

Public Function LNCA_SetParaStr() As String
    LNCA_SetParaStr = gudtPara.strSignURL & G_STR_SPLIT & gudtPara.bytSignVersion & G_STR_SPLIT & gudtPara.strSIGNIP
End Function

Private Function GetSubString(ByVal strSource As String, ByVal strNode As String) As String
'����:��ȡ�����ַ�����ĳ���ڵ�ֵ
'����:strSource -�����ַ��� {"ret":1,"errinfo":"","rand":"3XO9JCXVJ6LF05M51165"}
'    strNode-��ʶҪ��ȡ�Ľڵ�����
    Dim arrMain As Variant
    Dim arrSub As Variant
    Dim strRet As String
    Dim i As Long
    
    arrMain = Split(strSource, ",")
    For i = LBound(arrMain) To UBound(arrMain)
        Select Case UCase(strNode)
        Case UCase("rand"), UCase("token")
            If InStr(UCase(arrMain(i)), UCase(strNode)) > 0 Then
                arrSub = Split(arrMain(i), ":")
                strRet = Mid(arrSub(1), 2)
                strRet = left(strRet, Len(strRet) - 2)
                Exit For
            End If
        Case UCase("ret")
            If InStr(UCase(arrMain(i)), UCase(strNode)) > 0 Then
                arrSub = Split(arrMain(i), ":")
                strRet = arrSub(1)
                Exit For
            End If
        Case UCase$("errinfo")
            If InStr(UCase(arrMain(i)), UCase(strNode)) > 0 Then
                arrSub = Split(arrMain(i), ":")
                strRet = Mid(arrSub(1), 2)
                strRet = left(strRet, Len(strRet) - 1)
                Exit For
            End If
        End Select
    Next
    GetSubString = strRet
End Function





