Attribute VB_Name = "mdlHBCA"
Option Explicit

'�ӱ������е���ҽԺ   �ӱ�CA
'2018-02-26��ɽ�Ϻ�ҽԺ����ʱ���֤��洢����:
'CAʱ����������Ǹ��ؾ���ģ���̨������֤��;ǩ��ʱ��ȡ�Ŀ�������һ̨��Ҳ����������һ̨
'����Ҫ����ʱ���֤��洢ģʽ�����ֽ��飬һ��������һ��ʱ���֤����Ϣ������ǩ��ʱ��ȡ����ʱ���֤����Ϣȥ���ݿ�飬���ֲ�������;
'�ڶ����ǰ�ÿ��ǩ��ʱ��ȡ����ʱ�����Ϣ�浽ǩ����Ϣ��(���ַ�ʽҪ��ǩ����Ϣֵ+�ָ���("[;]")+ʱ���֤�����ݳ���С��4000���ַ�)

Private mblnInit As Boolean         '�Ƿ��ѳ�ʼ���ɹ�
Private mCertMgr As Object          'HebcaP11XLib.certMgr
Private mSignCert As Object         'HebcaP11XLib.cert
Private mFormSeal As Object         'FormSealCtrl1 ����ǩ�¿ؼ�
Private mSVSClient As Object        'SVS_SOFT_COMLib.SvsVerify '���岢ʵ����SVS�ͻ������
Private mblnTs As Boolean           '�Ƿ�����ʱ���

Private Const M_STR_LICENCE As String = "amViY55oZWKcZmhlnWxhaGViY2GXGmJjYWhlYnGH1QQ5GcNqnW6z3vohVnE+nTJr"
Private Const M_STR_SUMMARY As String = "[SUMMARY]"
Private Const M_STR_SPLIT As String = "[;]"

Public Function HBCA_InitObject() As Boolean
    '----------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ������ǩ����Ҫ�õ�����
    '����:True-��ʼ���ɹ�;False-��ʼ��ʧ��
    '����:��ΰ��
    '����:2015-08-31
    '----------------------------------------------------------------------------------------------------------------------------------
      Dim arrList As Variant
    
100   If mblnInit Then HBCA_InitObject = True: Exit Function
      On Error GoTo errH
      '������Ϣ:IP|�˿ں�|�Ƿ�����ʱ���(0-������/1-����)
102   gstrPara = gobjComLib.zlDatabase.GetPara(90000, glngSys, , "121.28.49.158&&&5000&&&0")  '��ȡ��������
104   If gstrPara = "" Then
106       Err.Raise -1, , "�����ļ���ȡʧ�ܣ��뵽���õ���ǩ���ӿڴ����á�"
          Exit Function
      End If
108   arrList = Split(gstrPara, "&&&")
110   If UBound(arrList) <> 2 Then
112       MsgBoxEx "ǩ����������ַ���ø�ʽ����,�뵽���õ���ǩ���ӿڴ����á�", vbOKOnly + vbInformation, gstrSysName
          Exit Function
      End If
114  mblnTs = (Val(CStr(Split(gstrPara, G_STR_SPLIT)(2))) = 1)
116 If mCertMgr Is Nothing Then Set mCertMgr = CreateObject("HebcaP11X.CertMgr.1")
 
118 If mSignCert Is Nothing Then Set mSignCert = CreateObject("HebcaP11X.Cert.1")
120 If mFormSeal Is Nothing Then Set mFormSeal = CreateObject("HebcaFormSeal.FormSealCtrl.1")
122 If mSVSClient Is Nothing Then Set mSVSClient = CreateObject("Svs_soft_com.SvsVerify.1")
124   mCertMgr.Licence = M_STR_LICENCE
126   gstrLogins = ""
128   mblnInit = True
130   HBCA_InitObject = True
      Exit Function
errH:
132 MsgBoxEx "�����ӱ�CA�ӿڲ���ʧ�ܣ�" & vbNewLine & "��" & CStr(Erl()) & "�� " & vbNewLine & _
              Err.Description, vbQuestion, gstrSysName
End Function

Public Function HBCA_RegCert(arrCertInfo As Variant) As Boolean
'----------------------------------------------------------------------------------------------------------------------------------
'����:�ṩ��HIS���ݿ���ע�����֤��ı�Ҫ��Ϣ,�����·��Ż����֤��,��Ҫ����USB-Key
'���أ�arrCertInfo��Ϊ���鷵��֤�������Ϣ
'      0-ClientSignCertCN:�ͻ���ǩ��֤�鹫������(����),����ע��֤��ʱ������֤���
'      1-ClientSignCertDN:�ͻ���ǩ��֤������(ÿ��Ψһ)
'      2-ClientSignCertSN:�ͻ���ǩ��֤�����к�(ÿ֤��Ψһ)
'      3-ClientSignCert:�ͻ���ǩ��֤������
'      4-ClientEncCert:�ͻ��˼���֤������
'      5-ǩ��ͼƬ�ļ���,�մ���ʾû��ǩ��ͼƬ
'      6-ʱ���֤��
'      7-ǩ����Ϣ
'����:��ΰ��
'����:2015-08-31
'----------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    Dim strCertName As String
    Dim strCertDN As String, strUserID As String, strSealB64 As String
    Dim strCertSn As String, strTSCert As String
    Dim strSignCert As String, strPic As String
    Dim strEncCert As String
    
    On Error GoTo errH
    
    If mFormSeal Is Nothing Then Set mFormSeal = CreateObject("HebcaFormSeal.FormSealCtrl.1")
    
100     For i = LBound(arrCertInfo) To UBound(arrCertInfo)
102         arrCertInfo(i) = ""
        Next

104     If GetCertList(strCertName, strCertSn, strSignCert, strUserID, strSealB64, strTSCert, strPic) Then
106         arrCertInfo(0) = strCertName
108         arrCertInfo(1) = strCertDN
110         arrCertInfo(2) = strCertSn
112         arrCertInfo(3) = strSignCert
113         arrCertInfo(4) = strEncCert
            arrCertInfo(5) = strPic
            arrCertInfo(6) = strTSCert
            arrCertInfo(7) = strSealB64
114
124         HBCA_RegCert = True
        End If

        Exit Function
errH:
126     MsgBoxEx "��ȡ֤����Ϣʧ�ܣ�" & vbNewLine & "��" & CStr(Erl()) & "�� " & Err.Description, vbQuestion, gstrSysName
   
        
End Function

Private Function GetCertList(ByRef strName As String, ByRef strCertSn As String, ByRef strSignCert As String, _
                ByRef strUserID As String, Optional ByRef strSealBase64 As String, Optional ByRef strTSCert As String, _
                Optional ByRef strPicFile As String = "0", Optional ByRef strPic As String) As Boolean
    '----------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ֤����Ϣ
    '����:strName-֤���û���
    '     strCertSn-֤�����к�
    '     strSignCert-֤������
    '     strUserID-�û�Ψһ��ʶ  ��ȡ���֤��
    '     strSealBase64-ǩ��BASE64
    '     strTSCert-ʱ���֤��BASE64
    '���أ�
    '����:��ΰ��
    '����:2015-08-31
    '----------------------------------------------------------------------------------------------------------------------------------
        Dim intCount As Integer
        Dim strSignData As String
        Dim strSealSn As String
        Dim objPic As Picture
        
        On Error GoTo errH
        mCertMgr.Licence = M_STR_LICENCE
100     intCount = mCertMgr.GetDeviceCount
102      If intCount < 1 Then
104          MsgBoxEx "δ����KEY,��������Key��", vbInformation, gstrSysName
             Exit Function
         End If
106     Set mSignCert = mCertMgr.SelectSignCert
108     intCount = mFormSeal.GetSealCount()
110     If intCount < 1 Then
112          MsgBoxEx "��ǰ�豸û��ǩ�£�", vbInformation, gstrSysName
             Exit Function
         End If
         'CN=����������
114     strName = mSignCert.GetSubjectItem("cn")
         'strSignCert = mSignCert.GetCertB64        '�õ�ǩ��֤������
     '        strCertDN = mSignCert.GetSubjectItem("DN")
        '��ȡ����֤���Ψһ��ʶ,���ں��û�������
116     strUserID = mSignCert.GetCertExtensionByOid("1.2.156.112586.1.4") '"2@6021SF0130637201507090001"
118     strUserID = Mid(strUserID, 10) '��ȡ���֤��
        '��ȡ֤����Ϣ  ��֤ǩ����Ҫ֤����Ϣ��
        'ǩ��BASE64����,ǩ��֤��,ʱ���֤��BASE64����
120     strSignData = mFormSeal.SignAndSealWithoutTimeStampCert("����20150901", "", 0, True, mblnTs)
121     If strSignData = "" Then MsgBoxEx "�����ǩ��ʧ�ܣ�", vbInformation, gstrSysName: Exit Function
        '��ȡǩ�µ�SN\ǩ��Bases64
122     strSealSn = mFormSeal.GetSelectedSeal()
124     strSealBase64 = mFormSeal.GetSeal(strSealSn) '��ȡ�µ�B64
126     strPic = mFormSeal.GetSealPicFromB64(strSealBase64)
128     If strPicFile <> "0" Then
130            strPicFile = SaveBase64ToFile("gif", strSealSn, strPic)
        End If
         '��ͼƬת����ָ��bmp��ʽ
    '        Set objPic = LoadPicture(strPicFile)
    '        SavePicture objPic, strPicFile
     '
         '��ȡ֤�������Ϣ��֤���SN��֤�����Ч����Ҫ�������ݿ�
132     strSignCert = mFormSeal.GetCert(strSealSn) '��ȡ֤��
134     Set mSignCert = mCertMgr.CreateCertFromB64(strSignCert)
136     strCertSn = mSignCert.GetSerialNumber '֤��SN
         'dCertDate:=mSignCert.NotAfter    ��Ч��
          'ʱ���
138     If mblnTs Then
140         strTSCert = mFormSeal.GetTimeStampCert '��ȡʱ���֤������
         End If

        
142     GetCertList = True
        Exit Function
errH:
144     MsgBoxEx "��ȡ֤����Ϣʧ�ܣ�" & vbNewLine & "��" & CStr(Erl()) & "�� " & Err.Description, vbQuestion, gstrSysName
End Function

Public Function HBCA_Sign( _
    ByVal strSource As String, ByRef strSignData As String, ByRef strTimeStamp As String, _
    ByRef strTimeStampCode As String, ByRef blnReDo As Boolean) As Boolean
    '----------------------------------------------------------------------------------------------------------------------------------
    '����:ǩ��
    '����:strSource-����Դ
    '     strSignData-ǩ��ֵ
    '     strTimeStamp-ʱ���ֵ
    '     strTimeStampCode-ʱ�����Ϣ
    '���أ�True/false
    '����:��ΰ��
    '����:2015-08-31
    '----------------------------------------------------------------------------------------------------------------------------------
        'ǩ��
        Dim strTiemRequest As String
        Dim strTmp As String
        Dim strSealSn As String
        Dim strSealBase64 As String
        Dim strTSCert As String
        Dim blnCheck As Boolean
        
        On Error GoTo errH
100     blnCheck = HBCA_CheckCert(blnReDo)
102     If blnReDo Then Exit Function
104     If blnCheck Then                '��֤��ǰUSB�Ƿ���ǩ���û��ģ�����ȡǩ��֤��
            '����SignAndSealWithoutTimeStampCert��ԭ�Ľ��и��£����Զ�ԭ�����ݽ�����֯��
106         strSource = mCertMgr.util.HashText(strSource, 1)
108         strSignData = mFormSeal.SignAndSealWithoutTimeStampCert(strSource, "", 0, True, mblnTs)
109         If strSignData = "" Then MsgBoxEx "ǩ��ʧ��,ǩ��ֵΪ�գ�", vbInformation, gstrSysName: Exit Function
110         If mblnTs Then
112             strTimeStampCode = mFormSeal.GetTimeStamp() '��ȡʱ�����Ϣ
114             strTimeStamp = mFormSeal.GetTimeStampInfoByB64(strTimeStampCode, "time")
116             strTSCert = mFormSeal.GetTimeStampCert '��ȡʱ���֤������
118             If strTimeStampCode = "" Then
120                 MsgBoxEx "ǩ��ʧ��,ʱ���B64Ϊ�գ�", vbInformation, gstrSysName
                    Exit Function
122             ElseIf strTSCert = "" Then
124                 MsgBoxEx "ǩ��ʧ��,ʱ���֤������Ϊ�գ�", vbInformation, gstrSysName
                    Exit Function
                End If
            Else
126             strTimeStamp = CStr(gobjComLib.zlDatabase.Currentdate)
            End If
        Else
128         MsgBoxEx "ǩ��ʧ�ܣ�", vbInformation, gstrSysName
            Exit Function
        End If
130     If Trim(strSignData) <> "" Then strSignData = M_STR_SUMMARY & strSignData  '�˱�ʶ[SUMMARY]������֤ǩ��ʱ���ְ�ԭ����֤ǩ�����ǰ�ժҪ��֤ǩ��
132     If strTSCert <> "" Then strSignData = strSignData & M_STR_SPLIT & strTSCert     'M_STR_SPLIT �ָ���[;]
134     HBCA_Sign = True
        Exit Function
errH:
136     MsgBoxEx "ǩ��ʧ�ܣ�" & vbNewLine & "��" & CStr(Erl()) & "�� " & Err.Description, vbQuestion, gstrSysName
End Function

Public Function HBCA_VerifySign(ByVal strSignData As String, ByVal strSource As String, ByVal strTimeStampCode As String, _
                            ByVal strCert As String, ByVal strTSCert As String, ByVal strSealCert As String) As Boolean
    '----------------------------------------------------------------------------------------------------------------------------------
    '����:��֤ǩ��
    '����:
    '   strSignData-ǩ��ֵ
    '   strSource-Դ��
    '   strTimeStampCode-ʱ�����Ϣ
    '   strCert-֤������
    '   strTSCert-ʱ���֤������
    '   strSealCert-ǩ��֤������
    '���أ�True/false
    '����:��ΰ��
    '����:2015-08-31
    '----------------------------------------------------------------------------------------------------------------------------------
      Dim strTmp As String
      Dim varTmp As Variant
      Dim lngRet As Long
      Dim blnOk As Boolean
      On Error GoTo errH
    
100   If mFormSeal Is Nothing Then Set mFormSeal = CreateObject("HebcaFormSeal.FormSealCtrl.1")
102   strTmp = ""

      '����ǩ����֤
104   If UCase(left(strSignData, Len(M_STR_SUMMARY))) = M_STR_SUMMARY Then
          '��ժҪǩ��ʱ��֤ǩ��ʱ��ҪȡժҪ
106       strSignData = Mid(strSignData, Len(M_STR_SUMMARY) + 1)
107       mCertMgr.Licence = M_STR_LICENCE
108       strSource = mCertMgr.util.HashText(strSource, 1)
          '��ǩ����Ϣ�н���ʱ���֤�� M_STR_SPLIT �ָ���[;]
110       varTmp = Split(strSignData, M_STR_SPLIT)
112       If UBound(varTmp) = 1 Then
114           strSignData = varTmp(0)
116           strTSCert = varTmp(1)
          End If
      End If

118    Call mFormSeal.VerifyAndShowSeal(strSealCert, strCert, strSource, 1, strSignData, IIf(mblnTs, 0, -1), strTimeStampCode, strTSCert, 0)
120    lngRet = mFormSeal.GetVerifyResult()

122   If lngRet = 0 Then
124       strTmp = "ǩ����֤�ɹ���"
126       blnOk = True
      Else
128       strTmp = "ǩ����֤ʧ�ܣ�"
130       blnOk = False
      End If

132   If strTmp <> "" Then
134       MsgBoxEx strTmp, vbOKOnly + vbInformation, gstrSysName
      End If
    
136    HBCA_VerifySign = blnOk
      Exit Function
errH:
138     MsgBoxEx "��֤ǩ��ʧ�ܣ�" & vbNewLine & "��" & CStr(Erl()) & "�� " & Err.Description, vbInformation, gstrSysName
End Function

Public Function HBCA_CheckCert(ByRef blnReDo As Boolean) As Boolean
    '----------------------------------------------------------------------------------------------------------------------------------
    '���ܣ���ȡUSB�����豸��ʼ������¼
    '����:
    '   ����:blnRedo-True������ע��֤��ɹ�,False-δ����ע��֤��
    '���أ�True/false
    '����:��ΰ��
    '����:2015-08-31
    '----------------------------------------------------------------------------------------------------------------------------------
        Dim strCertUserID As String, strPIN As String, strUserName As String
        Dim strWebUrl As String, intDate   As Integer
        Dim strCertName As String, strCertSn As String, strCert As String, strCertDN As String
        Dim strTSCert As String, strSealCode As String, strPic As String, strDate As String
        Dim blnOk As Boolean
        Dim udtUser As USER_INFO
        
        On Error GoTo errH
100     If Not mblnInit Then
102         Call HBCA_InitObject
104         If Not mblnInit Then
106             MsgBoxEx "����δ��ʼ����"
                Exit Function
            End If
        End If
         '��ȡ֤����Ϣͬʱ���Key���Ƿ����
108     If Not GetCertList(strCertName, strCertSn, strCert, strCertUserID, strSealCode, strTSCert, , strPic) Then
110         HBCA_CheckCert = False: Exit Function
        End If
        'δע���ڵ�ǰ�û����µ�Key
112     If mUserInfo.strUserID = "" Then
114         MsgBoxEx "�������֤��Ϊ��,����ϵ����Ա����Ա������¼�룡", vbOKOnly + vbInformation, gstrSysName
            Exit Function
116     ElseIf strCertUserID <> mUserInfo.strUserID Then
118         MsgBoxEx "�������֤�ţ�" & _
                   vbCrLf & vbTab & "��" & mUserInfo.strUserID & "��" & vbCrLf & _
                   "��ǰ֤��Ψһ��ʶ:" & _
                   vbCrLf & vbTab & "��" & strCertUserID & "��" & vbCrLf & _
                   "�û����֤���뵱ǰ֤��Ψһ��ʶ�����,����ʹ�ã�", vbInformation, gstrSysName
            Exit Function
        End If
        
        '��¼��֤
120     If InStr(gstrLogins & "|", "|" & strCertSn & "|") > 0 Then '�״���֤ͨ�����´β��ڼ�����֤
122         blnOk = True
        Else
124         If Not GetCertLogin() Then
126             blnOk = False
            Else
128             blnOk = True
130             If InStr(gstrLogins & "|", "|" & strCertSn & "|") = 0 Then gstrLogins = gstrLogins & "|" & strCertSn
            End If
        End If
        
132     If blnOk Then
            '�ж��Ƿ���Ҫ����ע��֤��
134         udtUser.strName = strCertName
136         udtUser.strSignName = strCertName
138         udtUser.strUserID = strCertUserID
140         udtUser.strCertSn = strCertSn
142         udtUser.strCertDN = strCertDN
144         udtUser.strCert = strCert
146         udtUser.strEncCert = ""
148         udtUser.strCertID = ""
150         udtUser.strPicCode = strPic
152         udtUser.strTSCert = strTSCert
154         udtUser.strSealCode = strSealCode
            '��ȡ�Ѿ�ע��֤�����Ч��������
                '��ȡ֤�������Ϣ��֤���SN��֤�����Ч����Ҫ�������ݿ�
156         Set mSignCert = mCertMgr.CreateCertFromB64(mUserInfo.strCert)
158         strDate = mSignCert.NotAfter
160         If IsUpdateRegCert(udtUser, strDate, blnReDo) Then
162             HBCA_CheckCert = True
            Else
164             HBCA_CheckCert = False
            End If
        End If
    
     
    
        Exit Function
errH:
166     MsgBoxEx "���USBKEYʧ�ܣ�" & vbNewLine & "��" & CStr(Erl()) & "�� " & Err.Description, vbQuestion, gstrSysName
End Function

Private Function GetCertLogin() As Boolean
    '----------------------------------------------------------------------------------------------------------------------------------
    '���ܣ���¼��֤
    '����:
    '���أ�True/false
    '����:��ΰ��
    '����:2015-08-31
    '----------------------------------------------------------------------------------------------------------------------------------
        Dim strText As String
        Dim strMsg As String
        Dim lngRetVal As Long
        Dim strSignData As String
        Dim strCertB64 As String
        Dim strSealSn As String
        Dim strDate As String
        Dim intDay As Integer
        Dim strIP As String, lngPort As Long
        Dim arrTmp As Variant
    
        On Error GoTo errH
100     If mCertMgr Is Nothing Then Set mCertMgr = CreateObject("HebcaP11X.CertMgr.1") 'ʵ����P11�����CertMgr��
102     If mSVSClient Is Nothing Then Set mSVSClient = CreateObject("Svs_soft_com.SvsVerify.1")
104     If mFormSeal Is Nothing Then Set mFormSeal = CreateObject("HebcaFormSeal.FormSealCtrl.1")
106     strText = "hebca2013" 'ԭʼ�ַ���

108     mCertMgr.Licence = M_STR_LICENCE

110     Set mSignCert = mCertMgr.SelectSignCert  '�õ�ǩ��֤�����
112     strSignData = mSignCert.SignText(strText, 1)   '��������ǩ��,��ǩ��ֵ��ŵ�signdata
114     strCertB64 = mSignCert.GetCertB64         '�õ�ǩ��֤������
        'gstrPara = "121.28.49.158&&&5000"
116     arrTmp = Split(gstrPara, G_STR_SPLIT)     'IP&&&�˿ں� "121.28.49.158", 5000
118     strIP = arrTmp(0): lngPort = Val(CStr(arrTmp(1)))
120     lngRetVal = mSVSClient.InitialVerify(strIP, lngPort) '��ʼ��SVS�ͻ���

        Dim r As Boolean
    
122     If lngRetVal < 0 Then
124         MsgBoxEx "�޷�����SVS������!", vbInformation, gstrSysName
            Exit Function
        End If
    
126     lngRetVal = mSVSClient.VerifyCertSign(-1, 0, strText, Len(strText), strCertB64, strSignData, 1, lngRetVal)     '��֤
128     Select Case lngRetVal
            Case 0
130             strMsg = "��֤�ɹ�"
132         Case 1
134             strMsg = "����֤��δ��Ч!"
136         Case 2
138             strMsg = "����֤���Ѿ�����!"
140         Case 4
142             strMsg = "����֤��Ǻӱ�CA�䷢!"
144         Case 1002
146             strMsg = "����֤��Ǻӱ�CA�䷢!"
148         Case 7
150             strMsg = "����֤���Ѿ�������!"
152         Case -6406
154             strMsg = "ǩ����֤ʧ��,������!"
156         Case Else
158             strMsg = "ǩ����֤ʧ��!������:" & lngRetVal
        End Select
160     If strMsg <> "��֤�ɹ�" Then
162         MsgBoxEx "������Ϣ:" & strMsg, vbInformation, gstrSysName
            Exit Function
        End If
    '
    '     strSignData = mFormSeal.SignAndSealWithoutTimeStampCert("����20150901", "", 0, True, True)
    '
    '     '��ȡǩ�µ�SN\ǩ��Bases64
    '     strSealSn = mFormSeal.GetSelectedSeal()
    '     strSignCert = mFormSeal.GetCert(strSealSn) '��ȡ֤��
    '     Set mSignCert = mCertMgr.CreateCertFromB64(strSignCert)
164      strDate = mSignCert.NotAfter  '��Ч��
166     If strDate <> "" Then
        '��֤�ͻ���֤����Ч��ʣ������
168         intDay = CheckValidaty(CDate(strDate))
    
170         If (intDay <= 30 And intDay > 0 And Not gblnShow) Then
172             MsgBoxEx "����֤�黹��" & intDay & "����ڡ�", vbOKOnly + vbInformation, gstrSysName
174             gblnShow = True
176         ElseIf (intDay <= 0) Then
178             MsgBoxEx "����֤���ѹ��� " & Abs(intDay) & " �졣", vbOKOnly + vbInformation, gstrSysName
                Exit Function
            End If
        End If
    
180     GetCertLogin = True
        Exit Function
errH:
182     MsgBoxEx "��¼��֤ʧ�ܣ�" & vbNewLine & "��" & CStr(Erl()) & "�� " & Err.Description, vbInformation, gstrSysName

End Function

Public Function HBCA_GetPara() As Boolean
    '���÷�������ַ
        Dim arrList As Variant
    
        On Error GoTo errH
100     gstrPara = gobjComLib.zlDatabase.GetPara(90000, glngSys)
102     If gstrPara = "" Then gstrPara = "121.28.49.158&&&5000&&&0" '������Ϣ:IP&&&�˿ں�&&&�Ƿ�����ʱ���(0-������/1-����)
104     If gstrPara <> "" Then
106         arrList = Split(gstrPara, "&&&")
108         If UBound(arrList) = 2 Then
110              gudtPara.strSIGNIP = Trim(arrList(0))
112              gudtPara.strSignPort = Trim(arrList(1))
114              gudtPara.blnISTS = (Val(arrList(2)) = 1)
            End If
        End If
        Exit Function
errH:
116     MsgBoxEx "��ȡ����ʧ�ܣ�" & vbNewLine & "��" & CStr(Erl()) & "�� " & Err.Description, vbInformation, gstrSysName
End Function

Public Function HBCA_SetParaStr() As String
    HBCA_SetParaStr = gudtPara.strSIGNIP & "&&&" & gudtPara.strSignPort & "&&&" & IIf(gudtPara.blnISTS, "1", "0")
End Function

Public Sub HBCA_UnloadObj()
'----------------------------------------------------------------------------------------------------------------------------------
'����:ж�ض���
'����:��
'����:��ΰ��
'����:2015-08-31
'----------------------------------------------------------------------------------------------------------------------------------
    Set mCertMgr = Nothing
    Set mSVSClient = Nothing
    Set mSignCert = Nothing
    Set mFormSeal = Nothing
    mblnInit = False
End Sub
