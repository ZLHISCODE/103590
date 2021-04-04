Attribute VB_Name = "mdlBJCAJS"
Option Explicit

'����CA���Ĺ���ģ��(�½��հ�_��Ǩ������ҽԺ��Ŀ)
Private mstrLastPwd As String    '�����ϴ���������
Private mintLogin As Integer     '�����������
Private BJCAJS_Pic As Object
Private BJCAJS_Client As Object       '�ͻ���֤�鲿��
Private BJCAJS_svs As Object          'ǩ����֤�ؼ�
Private BJCAJS_TS As Object           'ʱ����ؼ�
Private mstrLogins As String          '����Ѿ�ͨ����¼��֤��key�����к�
Private mblnInit As Boolean

Public Function BJCAJS_InitObj() As Boolean
        '֤�鲿����ʼ��
        On Error GoTo errH
    
100     If mblnInit Then BJCAJS_InitObj = True: Exit Function
    
102     Set BJCAJS_Client = CreateObject("XTXAppCOM.XTXApp.1")
104     Set BJCAJS_Pic = CreateObject("GetKeyPic.GetPic")
106     Set BJCAJS_svs = CreateObject("BJCA_SVS_ClientCOM.BJCASVSEngine.1")   '����֤����֤�ؼ�
108     Set BJCAJS_TS = CreateObject("BJCA_TS_ClientCom.BJCATSEngine")    '����ʱ����ؼ�
    
110     mintLogin = 0
112     mstrLogins = ""
114     BJCAJS_InitObj = True
116     mblnInit = True
        Exit Function
errH:
118      MsgBoxEx "�����ӿڲ���ʧ�ܣ�" & vbNewLine & Err.Description, vbQuestion, gstrSysName
End Function

Public Function BJCAJS_RegCert(arrCertInfo As Variant, Optional ByVal strUserID As String) As Boolean
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
            On Error GoTo errH
        
100         For i = LBound(arrCertInfo) To UBound(arrCertInfo)
102              arrCertInfo(i) = ""
            Next
        
104         If BJCAJS_GetCertList(strCertUserName, strCertSn, strCertDN, strCertUserID, strCert) Then
106             If UCase(strCertUserID) <> UCase(strUserID) Then
108                 MsgBoxEx "�û����֤�ţ�" & _
                               vbCrLf & vbTab & "��" & UCase(strUserID) & "��" & vbCrLf & _
                               "��ǰ֤��Ψһ��ʶ:" & _
                               vbCrLf & vbTab & "��" & UCase(strCertUserID) & "��" & vbCrLf & _
                               "�û����֤���뵱ǰ֤��Ψһ��ʶ�����,����ע�ᣡ", vbInformation, gstrSysName
                    Exit Function
                End If
110             arrCertInfo(0) = strCertUserName
112             arrCertInfo(1) = strCertDN '֤��DN
114             arrCertInfo(2) = strCertSn '֤�����к� ǩ��ʱҪ��
116             arrCertInfo(3) = strCert
118             arrCertInfo(4) = ""
120             arrCertInfo(5) = SaveBase64ToFile("gif", strCertUserID, BJCAJS_Pic.getpic())
122             BJCAJS_RegCert = True
            End If

            Exit Function
errH:
124      MsgBoxEx "��ȡ֤����Ϣʧ�ܣ�" & vbNewLine & "��" & CStr(Erl()) & "�� " & Err.Description, vbQuestion, gstrSysName

End Function

Public Function BJCAJS_CheckCert(ByRef blnReDo As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------
    '���ܣ���ȡUSB�����豸��ʼ������¼
    '������
    '   ���� blnRedo:True-����ע��ɹ�
    '---------------------------------------------------------------------------------------------------------------------
        Dim strKey As String, strPIN As String, strUserName As String
        Dim strCertName As String, strCertDN As String
        Dim strCertSn As String
        Dim strCertUserID As String    '�������֤����Ϣ
        Dim strDate As String
        Dim udtUser As USER_INFO
        Dim strCert As String, strCertID As String
        Dim blnOk As Boolean
        Dim blnRet As Boolean
    
        On Error GoTo errH
    
100     If BJCAJS_Client Is Nothing Then Set BJCAJS_Client = CreateObject("XTXAppCOM.XTXApp.1")
102     If BJCAJS_Pic Is Nothing Then Set BJCAJS_Pic = CreateObject("GetKeyPic.GetPic")
104     If BJCAJS_svs Is Nothing Then Set BJCAJS_svs = CreateObject("BJCA_SVS_ClientCOM.BJCASVSEngine.1")   '����֤����֤�ؼ�
106     If BJCAJS_TS Is Nothing Then Set BJCAJS_TS = CreateObject("BJCA_TS_ClientCom.BJCATSEngine")    '����ʱ����ؼ�
    
         '��ȡ֤����Ϣͬʱ���Key���Ƿ����
108     If Not BJCAJS_GetCertList(strCertName, strCertSn, strCertDN, strCertUserID, strCert, strCertID) Then
110         BJCAJS_CheckCert = False: Exit Function
        End If
        'δע���ڵ�ǰ�û����µ�Key
112     If mUserInfo.strUserID = "" Then
114         MsgBoxEx "�������֤��Ϊ��,����ϵ����Ա����Ա������¼�룡", vbOKOnly + vbInformation, gstrSysName
            Exit Function
116     ElseIf UCase(strCertUserID) <> UCase(mUserInfo.strUserID) Then
118         MsgBoxEx "�������֤�ţ�" & _
                       vbCrLf & vbTab & "��" & UCase(mUserInfo.strUserID) & "��" & vbCrLf & _
                       "��ǰ֤��Ψһ��ʶ:" & _
                       vbCrLf & vbTab & "��" & UCase(strCertUserID) & "��" & vbCrLf & _
                       "�û����֤���뵱ǰ֤��Ψһ��ʶ�����,����ʹ�ã�", vbInformation, gstrSysName
            Exit Function
        End If
        '��������
120     If mstrLastPwd <> "" Then strPIN = mstrLastPwd
122     If strPIN = "" Then
124         If Not frmPassword.ShowMe(strPIN) Then Exit Function
        End If
        '������֤���������,�״ε���ǩ���ӿ�ʱ�ᴥ��CA�����봰��
126     If strPIN = "" Then
128        MsgBoxEx "������֤�����룡", vbOKOnly + vbInformation, gstrSysName
           Exit Function
        Else
130         If mintLogin >= 8 Then
132             MsgBoxEx "�Ѿ�������" & mintLogin & "�δ������룬������������������", vbOKOnly + vbInformation, gstrSysName
                Exit Function
            End If
134         blnRet = BJCAJS_Client.SOF_Login(strCertID, strPIN)
136         If Not blnRet Then
138             mintLogin = mintLogin + 1
140             MsgBoxEx "֤��������ܲ���ȷ�����Ѿ�������" & mintLogin & "�����룬����������" & 8 - mintLogin & "��!", vbOKOnly + vbInformation, gstrSysName
142             mstrLastPwd = ""
                Exit Function
            End If
         End If
     
        '��¼��֤
144     If InStr(mstrLogins & "|", "|" & strCertSn & "|") > 0 Then '�״���֤ͨ�����´β��ڼ�����֤
146         blnOk = True
        Else
148         If Not GetCertLogin(strCertID, strCert) Then
150             blnOk = False
            Else
152             blnOk = True
154             If InStr(mstrLogins & "|", "|" & strCertSn & "|") = 0 Then mstrLogins = mstrLogins & "|" & strCertSn
            End If
        End If
    
156     If blnOk Then
            '�ж��Ƿ���Ҫ����ע��֤��
158         udtUser.strName = strCertName
160         udtUser.strSignName = strCertName
162         udtUser.strUserID = strCertUserID
164         udtUser.strCertSn = strCertSn
166         udtUser.strCertDN = strCertDN
168         udtUser.strCert = strCert
170         udtUser.strEncCert = ""
172         udtUser.strCertID = strCertID
174         udtUser.strPicCode = BJCAJS_Pic.getpic()
            '��ȡ�Ѿ�ע��֤�����Ч��������
176         strDate = BJCAJS_Client.SOF_GetCertInfo(mUserInfo.strCert, 12)
178         strDate = String14ToDate(strDate)
180         If IsUpdateRegCert(udtUser, strDate, blnReDo) Then
182             BJCAJS_CheckCert = True
            Else
184             BJCAJS_CheckCert = False
            End If
        End If
    
186     mUserInfo.strCertID = strCertID
    
188     mstrLastPwd = strPIN
        Exit Function
errH:
190      MsgBoxEx "���USBKEYʧ�ܣ�" & vbNewLine & "��" & CStr(Erl()) & "�� " & Err.Description, vbInformation, gstrSysName
End Function

Public Function BJCAJS_Sign(ByVal strCurrCertSn As String, ByVal strSource As String, ByRef strSignData As String, _
            ByRef strTimeStamp As String, ByRef strTimeStampCode As String, ByRef blnReDo As Boolean) As Boolean
    'ǩ��
    '������
    '   strPID --�û���ݱ�ʶ��һ��Ϊ���֤�ţ�
        Dim strSigCert As String
        Dim CertID As String
        Dim strRequest As String    'ʱ�������
        Dim strDate As String
        Dim strMsg As String
        Dim blnCheck As Boolean
        On Error GoTo errH
    
100     blnCheck = BJCAJS_CheckCert(blnReDo)
102     If blnReDo Then Exit Function
    
104     If blnCheck Then                '��֤��ǰUSB�Ƿ���ǩ���û��ģ�����ȡǩ��֤��
106         strSignData = BJCAJS_Client.SOF_SignData(mUserInfo.strCertID, strSource) '����ǩ������
108         If strSignData <> "" Then
                'Դ�ļ�ǩ��ֵ����ʱ�������(����֤��)
110             strRequest = BJCAJS_TS.CreateTSRequest(strSource & strSignData, 0)
112             If strRequest <> "" Then
114                 strTimeStampCode = BJCAJS_TS.CreateTS(strRequest)  '����ʱ�������֤�飩
116                 If strTimeStampCode = "" Then
118                     MsgBoxEx "����ʱ���ʧ�ܣ�", vbOKOnly + vbInformation, gstrSysName
120                     BJCAJS_Sign = False
                        Exit Function
                    End If
122                 strDate = BJCAJS_TS.gettimestampinfo(strTimeStampCode, 1)  'ȡ��ʱ���ʱ��
124                 strTimeStamp = String14ToDate(strDate)
                Else
126                 MsgBoxEx "ʱ�������ʧ�ܣ�", vbOKOnly + vbInformation, gstrSysName
128                 BJCAJS_Sign = False
                    Exit Function
                End If
          
130             BJCAJS_Sign = True
            Else
132             MsgBoxEx "ǩ��ʧ�ܣ�", vbOKOnly + vbInformation, gstrSysName
            End If
        End If
    
        Exit Function
errH:
134      MsgBoxEx "ǩ��ʧ�ܣ�" & vbNewLine & "��" & CStr(Erl()) & "�� " & Err.Description, vbInformation, gstrSysName
End Function

'''��֤ǩ������
Public Function BJCAJS_VerifySign(ByVal strCertSn As String, ByVal strSignData As String, ByVal strSource As String, ByVal strTimeStampCode As String) As Boolean
    '��֤ǩ��
    Dim strTmp As String
    Dim intRet As Integer
    Dim blnOk As Boolean
    Dim strDate As String
    Dim strTimeStamp As String
    On Error GoTo errH
    
100 If BJCAJS_svs Is Nothing Then Set BJCAJS_svs = CreateObject("BJCA_SVS_ClientCOM.BJCASVSEngine.1")   '����֤����֤�ؼ�
102 If BJCAJS_TS Is Nothing Then Set BJCAJS_TS = CreateObject("BJCA_TS_ClientCom.BJCATSEngine")    '����ʱ����ؼ�
    '��֤ʱ���
104 strTmp = ""
106 intRet = BJCAJS_TS.VerifyTS(strTimeStampCode, "")  '����֤Դ��
108 If intRet <> 0 Then
110     strTmp = "��֤ʱ���ʧ�ܣ�" & GetReturnInfo(intRet)
112     blnOk = False
    Else
114     strDate = BJCAJS_TS.gettimestampinfo(strTimeStampCode, 1)  'ȡ��ʱ���ʱ��
116     strTimeStamp = String14ToDate(strDate)
118     strTmp = "��֤ʱ����ɹ���" & vbTab & "ǩ��ʱ��:" & strTimeStamp
120     blnOk = True
    End If
    
    '��֤ǩ��
122 intRet = BJCAJS_svs.VerifySignatureBySN(strCertSn, strSource, strSignData)
124 If (intRet = 0) Then
126     strTmp = IIf(strTmp <> "", strTmp & vbCrLf, "") & "����ǩ����֤�ɹ���"
128     blnOk = True And blnOk
    Else
130     strTmp = IIf(strTmp <> "", strTmp & vbCrLf, "") & "����ǩ����֤ʧ�ܣ�"
132     blnOk = False
    End If

134 If strTmp <> "" Then
136     MsgBoxEx strTmp, vbOKOnly + vbInformation, gstrSysName
    End If
    
138 BJCAJS_VerifySign = blnOk
    Exit Function
errH:
140     MsgBoxEx "��֤ǩ��ʧ�ܣ�" & vbNewLine & "��" & CStr(Erl()) & "�� " & Err.Description, vbInformation, gstrSysName
End Function

'���ٶ���
Public Sub BJCAJS_UloadObj()
    Set BJCAJS_Client = Nothing
    Set BJCAJS_svs = Nothing
    Set BJCAJS_TS = Nothing
    mblnInit = False
End Sub

'----- �������ڲ�����
Private Function GetCertLogin(ByVal strCertID As String, ByVal strCert As String) As Boolean
        '����CA���հ�����֤���¼����
        '- ���
        'strCertID            :֤��ID
        'strCert              ֤������BASE64����
        Dim random As String
        Dim serverCert As String
        Dim serverSign As String, strSignVal As String
        Dim blnRet As Boolean
        Dim strDate As String
        Dim intDay As Integer, intRetSign As Integer, intRetVal As Integer
        Dim strTmp As String
        Dim lngRet As Long
    
        On Error GoTo errH
        '1)��BJCA_SVS_ClientCOM�����HISϵͳ����CA�ӿڣ���ȡ�������������֤�飬��ͨ��������֤������������ǩ����
100     random = BJCAJS_svs.GenRandom(16) '��ȡ�����
102     serverCert = BJCAJS_svs.GetServerCertificate() '��ȡ������֤��
104     serverSign = BJCAJS_svs.SignData(random) '����˶������ǩ��
    
106     blnRet = BJCAJS_Client.SOF_VerifySignedData(serverCert, random, serverSign) '�ͻ�����֤�����ǩ��
108     If Not blnRet Then
110         MsgBoxEx "�����ǩ����֤ʧ�ܣ�", vbOKOnly + vbInformation, gstrSysName
            Exit Function
        End If
        '��֤֤���Ƿ����
112     strDate = BJCAJS_Client.SOF_GetCertInfo(strCert, 12)
114     strDate = String14ToDate(strDate)
116     If strDate <> "" Then
        '��֤�ͻ���֤����Ч��ʣ������
118         intDay = CheckValidaty(CDate(strDate))
    
120         If (intDay <= 30 And intDay > 0 And Not gblnShow) Then
122             MsgBoxEx "����֤�黹��" & intDay & "����ڡ�", vbOKOnly + vbInformation, gstrSysName
124             gblnShow = True
126         ElseIf (intDay <= 0) Then
128             MsgBoxEx "����֤���ѹ��� " & Abs(intDay) & " �졣", vbOKOnly + vbInformation, gstrSysName
                Exit Function
            End If
        End If
    
130     lngRet = BJCAJS_svs.ValidateCertificate(strCert)
132     If lngRet <> 0 Then
134         If lngRet = -1 Then
136             MsgBoxEx "���������εĸ�!", vbOKOnly + vbInformation, gstrSysName
                Exit Function
138         ElseIf lngRet = -2 Then
140             MsgBoxEx "֤�鳬����Ч�ڣ�", vbOKOnly + vbInformation, gstrSysName
                Exit Function
142         ElseIf lngRet = -3 Then
144             MsgBoxEx "֤���Ѿ����ϣ�", vbOKOnly + vbInformation, gstrSysName
                Exit Function
146         ElseIf lngRet = -4 Then
148             MsgBoxEx "֤�鱻�����������", vbOKOnly + vbInformation, gstrSysName
                Exit Function
150         ElseIf lngRet = -5 Then
152             MsgBoxEx "֤��δ��Ч��", vbOKOnly + vbInformation, gstrSysName
                Exit Function
            End If
        End If
    
        '��֤֤���Ƿ����
154     strSignVal = BJCAJS_Client.SOF_SignData(strCertID, random)  '�ͻ��������ǩ��
156     intRetSign = BJCAJS_svs.VerifySignedData(strCert, random, strSignVal)   '�������֤�ͻ���ǩ��
158     intRetVal = BJCAJS_svs.ValidateAndSaveCertificate(strCert)  '�������֤�ͻ���֤����Ч�Բ�����֤��
    
160     If Not (intRetSign = 0 And (intRetVal = 0 Or intRetVal = 1)) Then
162         MsgBoxEx "�ͻ���֤����ʧ�ܣ�", vbOKOnly + vbInformation, gstrSysName
            Exit Function
        End If
    
164     GetCertLogin = True
        Exit Function
errH:
166     MsgBoxEx "��¼��֤ʧ�ܣ�" & vbNewLine & "��" & CStr(Erl()) & "�� " & Err.Description, vbInformation, gstrSysName
End Function

''' ��ȡ�ͻ���֤���б�
''' ����boolean
Public Function BJCAJS_GetCertList(ByRef strName As String, Optional ByRef strCertSn As String = "0", Optional ByRef strCertDN As String = "0", _
            Optional ByRef strCertUserID As String = "0", Optional ByRef strCert As String, Optional ByRef strCertID As String) As Boolean
        '����CA���հ��ȡ����֤���б���
        '-���:��
        '-����
        'strName :      ����ӿڷ��ص�֤������������
        'strCertSN      ����ӿڷ��ص�֤��SN
        'strCertDN:     ����ӿڷ��ص�֤��DN
        'strCertUserID:  ����ӿڷ��ص�֤��������Ψһ��ʶ
        'strCert:       ����ӿڷ��ص�ǩ��֤��
        'strCertID      ����֤��ID
        Dim strUsbkeyList As String
        Dim arrUserListLength As Integer
        Dim arrUserList() As String
        Dim strUser As String
        Dim i As Integer
        
        On Error GoTo errH
        
100     If BJCAJS_Client Is Nothing Then Set BJCAJS_Client = CreateObject("XTXAppCOM.XTXApp.1")
102     If BJCAJS_Pic Is Nothing Then Set BJCAJS_Pic = CreateObject("GetKeyPic.GetPic")
        '��ȡ֤��
104     strUsbkeyList = BJCAJS_Client.SOF_GetUserList()
106     If (strUsbkeyList = "") Then
108         strName = ""
110         MsgBoxEx "�����֤��Key��", vbOKOnly + vbInformation, gstrSysName
112         BJCAJS_GetCertList = False
            Exit Function
        Else
114         arrUserList = Split(strUsbkeyList, "&&&") '��Ǩ������(����)||999000100089956/6001201312021788&&&��Ǩ�����(����)||999000100089948/6002201309019595&&&
116         If UBound(arrUserList) > 1 Then  '���KEY
118             For i = LBound(arrUserList) To UBound(arrUserList) - 1
120                 strUser = strUser & "&&&" & Split(arrUserList(i), "||")(0)
                Next
122             If strUser <> "" Then strUser = Mid(strUser, 4)
124             strName = frmSelectUser.ShowMe(strUser)
            
126             For i = LBound(arrUserList) To UBound(arrUserList) - 1
128                If strName = Split(arrUserList(i), "||")(0) Then
130                     strCertID = Split(arrUserList(i), "||")(1)
                        Exit For
                   End If
                Next
            Else
132             arrUserList = Split(arrUserList(0), "||")
134             strName = arrUserList(0)      '֤��CNͨ����
136             strCertID = arrUserList(1)    '֤��ID
            End If
        
138         strCert = BJCAJS_Client.SOF_ExportUserCert(strCertID) '3.����ǩ��֤�顣
        
140         If strCertSn <> "0" Then strCertSn = BJCAJS_Client.SOF_GetCertInfo(strCert, 2) '֤�����к� ǩ��ʱҪ��
142         If strCertDN <> "0" Then strCertDN = BJCAJS_Client.SOF_GetCertInfo(strCert, 33) '֤��DN
144         If strCertUserID <> "0" Then
146             strCertUserID = BJCAJS_Client.SOF_GetCertInfoByOid(strCert, "1.2.156.112562.2.1.1.1") '2.��ȡ֤��Ψһ��ʶ��һ��Ϊ���֤�ţ�SF+���֤��
148             strCertUserID = Mid(strCertUserID, 3)
            End If
        End If
150     BJCAJS_GetCertList = True
        Exit Function
errH:
    MsgBoxEx "��ȡ֤����Ϣʧ�ܣ�" & vbNewLine & "��" & CStr(Erl()) & "�� " & Err.Description, vbQuestion, gstrSysName
End Function

Private Function GetReturnInfo(ByVal intErrNum As Integer) As String
    '׼���ʱ���������Ϣת������
    If intErrNum = -1 Then
        GetReturnInfo = "ʱ�����֤��ͨ��"
    ElseIf intErrNum = -2 Then
        GetReturnInfo = "ԭ����֤��ͨ��"
    ElseIf intErrNum = -3 Then
        GetReturnInfo = "���������εĸ�"
    ElseIf intErrNum = -4 Then
        GetReturnInfo = "֤��δ��Ч"
    ElseIf intErrNum = -5 Then
        GetReturnInfo = "��ѯ������֤��"
    ElseIf intErrNum = -6 Then
        GetReturnInfo = "ǩ��ʱ���ʱ������֤�����"
    ElseIf intErrNum = 0 Then
        GetReturnInfo = "��֤�ɹ�"
    Else
        GetReturnInfo = "δ֪����"
    End If
    If GetReturnInfo <> "" Then
        GetReturnInfo = "ʱ����ӿڷ�����ʾ��" & GetReturnInfo
    End If
End Function



