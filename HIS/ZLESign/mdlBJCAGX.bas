Attribute VB_Name = "mdlBJCAGX"
Option Explicit

'����CA���Ĺ���ģ��(�¹�����)
Private mblnInit As Boolean         '�Ƿ��ѳ�ʼ���ɹ�

Private BJCAGX_Client As Object       '�ͻ���֤�鲿��
Private BJCAGX_svs As Object          'ǩ����֤�ؼ�
Private BJCAGX_TS As Object           'ʱ����ؼ�
Private mobjPic As Object             '��ȡǩ��ͼƬ             '
Private mstrLastPwd As String    '�����ϴ���������
Private mintLogin As Integer     '�����������
Private mstrLogins As String          '����Ѿ�ͨ����¼��֤��key�����к�
Public gobjGXCAPenSign As Object

Public Enum Version
    V_RSA = 0
    V_SM2 = 1
End Enum

Public Function BJCAGX_InitObj() As Boolean
        '֤�鲿����ʼ��
        Dim progID As String
    
        On Error GoTo errH
100 BJCAGX_InitObj = mblnInit
102 If mblnInit Then Exit Function
104    If Not BJCAGX_GetPara(1) Then Exit Function

106    If gudtPara.blnSignPic Then
108        Set mobjPic = CreateObject("GetKeyPic.GetPic.1")
       End If

110    Set BJCAGX_svs = CreateObject("BJCA_SVS_ClientCOM.BJCASVSEngine.1")   '����֤����֤�ؼ�

112    Set BJCAGX_TS = CreateObject("BJCA_TS_ClientCom.BJCATSEngine")    '����ʱ����ؼ�

114    If gudtPara.bytSignVersion = V_RSA Then
116        Set BJCAGX_Client = CreateObject("BJCAAPPCTRL.BjcaAppCtrlCtrl.1") '����ǩ���ؼ�
       Else
118        Set BJCAGX_Client = CreateObject("XTXAppCOM.XTXApp.1")
       End If

120 BJCAGX_InitObj = True
122    mblnInit = BJCAGX_InitObj
       Exit Function
errH:
124 MsgBoxEx "�����ӿڲ���ʧ�ܣ�" & vbNewLine & "��" & CStr(Erl()) & "�� " & Err.Description, vbQuestion, gstrSysName
End Function

Public Function BJCAGX_RegCert(arrCertInfo As Variant) As Boolean
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
        Dim strFilePath As String
        On Error GoTo errH
    
        For i = LBound(arrCertInfo) To UBound(arrCertInfo)
             arrCertInfo(i) = ""
        Next
        
        If GetCertList(strCertUserName, strKeyId, strSigCert, strFilePath, strCertDN) Then
            arrCertInfo(0) = strCertUserName
            arrCertInfo(1) = strCertDN
            arrCertInfo(2) = strKeyId
            arrCertInfo(3) = strSigCert
            arrCertInfo(5) = strFilePath
            BJCAGX_RegCert = True
        End If

        Exit Function
errH:
     MsgBoxEx "��ȡ֤����Ϣʧ�ܣ�" & vbNewLine & "��" & CStr(Erl()) & "�� " & Err.Description, vbQuestion, gstrSysName

End Function

Public Function BJCAGX_CheckCert(ByRef blnReDo As Boolean) As Boolean
        '���ܣ���ȡUSB�����豸��ʼ������¼
    Dim strCertSN As String, strPIN As String, strCert As String
    Dim strCertUserID As String, strCertName As String
    Dim strCertID As String, strCertDN As String, strPicCode As String
    Dim blnRet As Boolean
    Dim blnOk As Boolean
    Dim udtUser As USER_INFO
    Dim strDate As String
    
    On Error GoTo errH
    
    If Not BJCAGX_InitObj() Then
        MsgBoxEx "����δ��ʼ����", vbInformation, gstrSysName
        Exit Function
    End If

    '��ȡ֤����Ϣͬʱ���Key���Ƿ����
    If Not GetCertList(strCertName, strCertSN, strCert, 0, strCertDN, strCertUserID, strCertID, strPicCode) Then
        BJCAGX_CheckCert = False: Exit Function
    End If
    
    If gudtPara.bytSignVersion = V_RSA Then
        If mUserInfo.strCertSN <> strCertSN Then
            MsgBoxEx "��֤��δע�����������£�����ʹ�ã�", vbOKOnly + vbInformation, gstrSysName
            Exit Function
        End If
        
        If Not GetCertLogin(strCertSN, strCert) Then
            BJCAGX_CheckCert = False
        Else
            BJCAGX_CheckCert = True
        End If
        blnReDo = False
        
    ElseIf gudtPara.bytSignVersion = V_SM2 Then
        'δע���ڵ�ǰ�û����µ�Key
        If mUserInfo.strUserID = "" Then
            MsgBoxEx "�������֤��Ϊ��,����ϵ����Ա����Ա������¼�룡", vbOKOnly + vbInformation, gstrSysName
            Exit Function
        ElseIf strCertUserID <> mUserInfo.strUserID Then
            MsgBoxEx "�������֤�ţ�" & _
                       vbCrLf & vbTab & "��" & mUserInfo.strUserID & "��" & vbCrLf & _
                       "��ǰ֤��Ψһ��ʶ:" & _
                       vbCrLf & vbTab & "��" & strCertUserID & "��" & vbCrLf & _
                       "�û����֤���뵱ǰ֤��Ψһ��ʶ�����,����ʹ�ã�", vbInformation, gstrSysName
            Exit Function
        End If
        '��������
        If mstrLastPwd <> "" Then strPIN = mstrLastPwd
        If strPIN = "" Then
            If Not frmPassword.ShowMe(strPIN) Then Exit Function
        End If
        '������֤���������,�״ε���ǩ���ӿ�ʱ�ᴥ��CA�����봰��
        If strPIN = "" Then
           MsgBoxEx "������֤�����룡", vbOKOnly + vbInformation, gstrSysName
           Exit Function
        Else
            If mintLogin >= 8 Then
                MsgBoxEx "�Ѿ�������" & mintLogin & "�δ������룬������������������", vbOKOnly + vbInformation, gstrSysName
                Exit Function
            End If
            blnRet = BJCAGX_Client.SOF_Login(strCertID, strPIN)
            If Not blnRet Then
                mintLogin = mintLogin + 1
                MsgBoxEx "֤��������ܲ���ȷ�����Ѿ�������" & mintLogin & "�����룬����������" & 8 - mintLogin & "��!", vbOKOnly + vbInformation, gstrSysName
                mstrLastPwd = ""
                Exit Function
            End If
         End If
         
        '��¼��֤
        If InStr(mstrLogins & "|", "|" & strCertSN & "|") > 0 Then '�״���֤ͨ�����´β��ڼ�����֤
            blnOk = True
        Else
            If Not GetCertLogin(strCertSN, strCert, strCertID) Then
                blnOk = False
            Else
                blnOk = True
                If InStr(mstrLogins & "|", "|" & strCertSN & "|") = 0 Then mstrLogins = mstrLogins & "|" & strCertSN
            End If
        End If
        
        If blnOk Then
            '�ж��Ƿ���Ҫ����ע��֤��
            udtUser.strName = strCertName
            udtUser.strSignName = strCertName
            udtUser.strUserID = strCertUserID
            udtUser.strCertSN = strCertSN
            udtUser.strCertDN = strCertDN
            udtUser.strCert = strCert
            udtUser.strEncCert = ""
            udtUser.strCertID = strCertID
            udtUser.strPicCode = strPicCode
            '��ȡ�Ѿ�ע��֤�����Ч��������
            strDate = BJCAGX_Client.SOF_GetCertInfo(mUserInfo.strCert, 12)
            strDate = String14ToDate(strDate)
            If IsUpdateRegCert(udtUser, strDate, blnReDo) Then
                BJCAGX_CheckCert = True
            Else
                BJCAGX_CheckCert = False
            End If
        End If
        
        mUserInfo.strCertID = strCertID
        
        mstrLastPwd = strPIN
    End If
    Exit Function
errH:
     MsgBoxEx "���USBKEYʧ�ܣ�" & vbNewLine & "��" & CStr(Erl()) & "�� " & Err.Description, vbQuestion, gstrSysName
End Function

Public Function BJCAGX_Sign(ByVal strCurrCertSn As String, ByVal strSource As String, ByRef strSignData As String, _
    ByRef strTimeStamp As String, ByRef strTimeStampCode As String, ByRef blnReDo As Boolean) As Boolean
            'ǩ��
        Dim strSigCert As String
        Dim strRequest As String    'ʱ�������
        Dim strDate As String
        Dim strMsg As String
        Dim blnRet As Boolean, blnCheck As Boolean
        Dim bytTsVer    As Byte  '1-�����ϰ汾ʱ����ӿ�
        
        On Error GoTo errH
100     If gudtPara.bytSignVersion = V_RSA Then
102         If BJCAGX_CheckCert(blnReDo) Then                  '��֤��ǰUSB�Ƿ���ǩ���û��ģ�����ȡǩ��֤��
104             strSignData = BJCAGX_Client.signedData(mUserInfo.strCertSN, strSource)  '����ǩ������
106             If strSignData <> "" Then
108                 blnRet = True
110                 strRequest = BJCAGX_TS.CreateTimeStampRequest(strSource)  '����ʱ�������
112                 If strRequest <> "" Then
114                     strTimeStampCode = BJCAGX_TS.CreateTimeStamp(strRequest)   '����ʱ�������֤�飩
116                     If strTimeStampCode <> "" Then
118                         strDate = BJCAGX_TS.gettimestampinfo(strTimeStampCode, 1)
120                         strTimeStamp = GetTimeStamp(strDate)          'ȡ��ʱ���ʱ��
122                         If strTimeStamp = "" Then
124                             strMsg = "����ʱ���ʧ�ܣ�"
126                             blnRet = False
                            End If
                        Else
128                         strMsg = "����ʱ���ʧ�ܣ�"
130                         blnRet = False
                        End If
                    Else
132                     strMsg = "ʱ�������ʧ�ܣ�"
134                     blnRet = False
                    End If
                Else
136                 strMsg = "ǩ��ʧ�ܣ�"
138                 blnRet = False
                End If
            Else
140             strMsg = "��֤ǩ��ʧ�ܣ�"
142             blnRet = False
            End If
        Else
144         blnCheck = BJCAGX_CheckCert(blnReDo)
146         If blnReDo Then Exit Function
148         If blnCheck Then                '��֤��ǰUSB�Ƿ���ǩ���û��ģ�����ȡǩ��֤��
150             strSignData = BJCAGX_Client.SOF_SignData(mUserInfo.strCertID, strSource) '����ǩ������
152             If strSignData <> "" Then
                    'Դ�ļ�ǩ��ֵ����ʱ�������(����֤��)
154                 blnRet = True
                    On Error Resume Next
156                 strRequest = BJCAGX_TS.CreateTSRequest(strSource & strSignData, 0)
158                 If Err.Number = 438 Or strRequest = "" Then '����֧�ָ����Ի򷽷� �����ϰ汾
160                     strRequest = BJCAGX_TS.CreateTimeStampRequest(strSource & strSignData)  '����ʱ�������
162                     bytTsVer = 1
                    End If
164                 Err.Clear: On Error GoTo errH
166                 If strRequest <> "" Then
168                     If bytTsVer = 0 Then
170                         strTimeStampCode = BJCAGX_TS.CreateTS(strRequest)  '����ʱ�������֤�飩
172                         If strTimeStampCode = "" Then
174                             MsgBoxEx "����ʱ���ʧ�ܣ�", vbOKOnly + vbInformation, gstrSysName
176                             blnRet = False
                            Else
178                             strDate = BJCAGX_TS.gettimestampinfo(strTimeStampCode, 1)  'ȡ��ʱ���ʱ��
180                             strTimeStamp = String14ToDate(strDate)
182                             blnRet = True
                            End If
                        Else
184                         strTimeStampCode = BJCAGX_TS.CreateTimeStamp(strRequest)   '����ʱ�������֤�飩
186                         If strTimeStampCode <> "" Then
188                             strDate = BJCAGX_TS.gettimestampinfo(strTimeStampCode, 1)
190                             strTimeStamp = GetTimeStamp(strDate)          'ȡ��ʱ���ʱ��
192                             If strTimeStamp = "" Then
194                                 strMsg = "����ʱ���ʧ�ܣ�"
196                                 blnRet = False
                                End If
                            Else
198                             strMsg = "����ʱ���ʧ�ܣ�"
200                             blnRet = False
                            End If
                        End If
                    Else
202                     strMsg = "ʱ�������ʧ�ܣ�"
204                     blnRet = False
                    End If

                Else
206                 strMsg = "��֤ǩ��ʧ�ܣ�"
208                 blnRet = False
                End If
            End If
        End If
    
210     If strMsg <> "" Then
212         MsgBoxEx strMsg, vbOKOnly + vbInformation, gstrSysName
        End If
    
214     BJCAGX_Sign = blnRet
    
        Exit Function
errH:
216      MsgBoxEx "ǩ��ʧ�ܣ�" & vbNewLine & "��" & CStr(Erl()) & "�� " & Err.Description, vbQuestion, gstrSysName
End Function

Public Function BJCAGX_VerifySign(ByVal strCertSN As String, ByVal strSigCert As String, ByVal strSignData As String, ByVal strSource As String, ByVal strTStampCode As String) As Boolean
        '��֤ǩ��
        Dim strTmp As String
        Dim blnRet As Boolean
        Dim lngRuslt As Long
        Dim intRet As Integer
        Dim strDate As String
        Dim strTimeStamp As String
        Dim bytVer  As Byte
        
        On Error GoTo errH

100     Call BJCAGX_InitObj
        
102     If gudtPara.bytSignVersion = V_RSA Then
104         blnRet = BJCAGX_Client.VerifySignedData(strSigCert, strSource, strSignData)
106         If blnRet Then
108             strTmp = "��֤ǩ���ɹ���"
            Else
110             strTmp = "��֤ǩ��ʧ�ܣ�"
            End If
112         If blnRet And strTStampCode <> "" Then
114             lngRuslt = BJCAGX_TS.verifyTimeStamp(strTStampCode)
116             If lngRuslt <> 0 Then
118                 strTmp = "��֤ʱ���ʧ�ܣ�" & GetReturnInfo(lngRuslt)
120                 blnRet = False
                End If
            End If
        Else
            '��֤ʱ���
122         If strTStampCode <> "" Then
124             strTmp = ""
                On Error Resume Next
126             lngRuslt = BJCAGX_TS.verifyTimeStamp(strTStampCode): bytVer = 0
128             If lngRuslt <> 0 Then strTmp = "��verifyTimeStamp����֤ʱ���ʧ�ܣ�" & GetReturnInfo(lngRuslt): blnRet = False
130             If Err.Number <> 0 Or lngRuslt <> 0 Then
132                 lngRuslt = BJCAGX_TS.VerifyTS(strTStampCode, "")  '����֤Դ��
134                 bytVer = 1
                End If
136             Err.Clear: On Error GoTo errH
138             If bytVer = 1 Then
140                 If lngRuslt <> 0 Then
142                     strTmp = strTmp & vbCrLf & _
                            "��VerifyTS����֤ʱ���ʧ�ܣ�" & GetReturnInfo(lngRuslt)
144                         blnRet = False
                    Else
146                     strDate = BJCAGX_TS.gettimestampinfo(strTStampCode, 1)  'ȡ��ʱ���ʱ��
148                     strTimeStamp = String14ToDate(strDate)
150                     strTmp = "��֤ʱ����ɹ���" & vbTab & "ǩ��ʱ��:" & strTimeStamp
152                     blnRet = True
                    End If
                End If
            End If

            '��֤ǩ��
154         intRet = BJCAGX_svs.VerifySignatureBySN(strCertSN, strSource, strSignData)
156         If (intRet = 0) Then
158             strTmp = IIf(strTmp <> "", strTmp & vbCrLf, "") & "����ǩ����֤�ɹ���"
160             blnRet = True And blnRet
            Else
162             strTmp = IIf(strTmp <> "", strTmp & vbCrLf, "") & "����ǩ����֤ʧ�ܣ�"
164             blnRet = False
            End If
        End If
        
166     If strTmp <> "" Then
168         MsgBoxEx strTmp, vbOKOnly + vbInformation, gstrSysName
        End If
        
170     BJCAGX_VerifySign = blnRet
        Exit Function
errH:
172     MsgBoxEx "��֤ǩ��ʧ�ܣ�" & vbNewLine & "��" & CStr(Erl()) & "�� " & Err.Description, vbQuestion, gstrSysName
End Function

'���ٶ���
Public Sub BJCAGX_UloadObj()
    Set BJCAGX_Client = Nothing
    Set BJCAGX_svs = Nothing
    Set BJCAGX_TS = Nothing
    Set mobjPic = Nothing
    mblnInit = False
End Sub

'----- �������ڲ�����
Private Function GetCertLogin(ByVal strCertSN As String, ByVal strCert As String, Optional ByVal strCertID As String) As Boolean
        Dim random As String
        Dim serverCert As String
        Dim serverSign As String, strSignVal As String
        Dim blnRet As Boolean
        Dim strDate As String
        Dim intDay As Integer, intRetSign As Integer, intRetVal As Integer
        Dim strTmp As String
        Dim retValidateCert As Long
    
        On Error GoTo errH
    
100     If gudtPara.bytSignVersion = V_RSA Then
102         If BJCAGX_svs Is Nothing Then Set BJCAGX_svs = CreateObject("BJCA_SVS_ClientCOM.BJCASVSEngine.1")
104         random = BJCAGX_Client.GenRandom(24)
106         strSignVal = BJCAGX_Client.signedData(strCertSN, random)
            '֤�鰲ȫ��¼
            'strSignVal:�ǿ�:�ɹ�
            'strSignVal:��:���ɹ�
108         If (strSignVal <> "") Then
                '����������֤֤��
                '������е���֤��
110             retValidateCert = ValidateCert(strCert)
            
                '��֤֤������Ϣ��ʾ
112             If retValidateCert <> 0 Then Call ValidateCertView(retValidateCert)
    
114             If (retValidateCert = 0) Then
                    Dim s As String
                    '��ȡ�ͻ���֤����Ч�ڽ�ֹʱ��
116                 s = BJCAGX_Client.GetUserInfo(strCert, 12)
                    '��֤�ͻ���֤����Ч��ʣ������
118                 intDay = CheckValidaty(s)
            
120                 If (intDay <= 30 And intDay > 0) And Not gblnShow Then
122                     MsgBoxEx "����֤�黹��" & intDay & "�����"
124                     gblnShow = True '������ʾ
126                     GetCertLogin = True
128                 ElseIf (intDay <= 0) Then
130                     MsgBoxEx "����֤���ѹ��� " & Abs(intDay) & " ��"
132                     GetCertLogin = False
                    Else
134                     GetCertLogin = True
                    End If
                Else
136                 GetCertLogin = False
                End If
            Else
138             GetCertLogin = False
            End If
140     ElseIf gudtPara.bytSignVersion = V_SM2 Then
            '1)��BJCA_SVS_ClientCOM�����HISϵͳ����CA�ӿڣ���ȡ�������������֤�飬��ͨ��������֤������������ǩ����
142         random = BJCAGX_svs.GenRandom(16) '��ȡ�����
144         serverCert = BJCAGX_svs.GetServerCertificate() '��ȡ������֤��
146         serverSign = BJCAGX_svs.SignData(random) '����˶������ǩ��
        
148         blnRet = BJCAGX_Client.SOF_VerifySignedData(serverCert, random, serverSign) '�ͻ�����֤�����ǩ��
150         If Not blnRet Then
152             MsgBoxEx "�����ǩ����֤ʧ�ܣ�", vbOKOnly + vbInformation, gstrSysName
                Exit Function
            End If
            '��֤֤���Ƿ����
154         strDate = BJCAGX_Client.SOF_GetCertInfo(strCert, 12)
156         strDate = String14ToDate(strDate)
158         If strDate <> "" Then
            '��֤�ͻ���֤����Ч��ʣ������
160             intDay = CheckValidaty(CDate(strDate))
        
162             If (intDay <= 30 And intDay > 0 And Not gblnShow) Then
164                 MsgBoxEx "����֤�黹��" & intDay & "����ڡ�", vbOKOnly + vbInformation, gstrSysName
166                 gblnShow = True
168             ElseIf (intDay <= 0) Then
170                 MsgBoxEx "����֤���ѹ��� " & Abs(intDay) & " �졣", vbOKOnly + vbInformation, gstrSysName
                    Exit Function
                End If
            End If
            '��֤֤���Ƿ����
172         strSignVal = BJCAGX_Client.SOF_SignData(strCertID, random)  '�ͻ��������ǩ��
174         intRetSign = BJCAGX_svs.VerifySignedData(strCert, random, strSignVal)   '�������֤�ͻ���ǩ��
176         intRetVal = BJCAGX_svs.ValidateAndSaveCertificate(strCert)  '�������֤�ͻ���֤����Ч�Բ�����֤��
        
178         If Not (intRetSign = 0 And (intRetVal = 0 Or intRetVal = 1)) Then
180             MsgBoxEx "�ͻ���֤����ʧ�ܣ�", vbOKOnly + vbInformation, gstrSysName
                Exit Function
            End If
    
182         GetCertLogin = True
        End If
    
        Exit Function
errH:
184     MsgBoxEx "��¼��֤ʧ�ܣ�" & vbNewLine & "��" & CStr(Erl()) & "�� " & Err.Description, vbInformation, gstrSysName
End Function

Private Function ValidateCert(ByRef userCert As String) As Integer
    '����������֤֤��
    If BJCAGX_svs Is Nothing Then Set BJCAGX_svs = CreateObject("BJCA_SVS_ClientCOM.BJCASVSEngine.1")
    ValidateCert = BJCAGX_svs.ValidateCertificate(userCert)
 
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
        Case -5
            MsgBoxEx "֤��δ��Ч��"
    End Select
End Sub

''' ��ȡ�ͻ���֤���б�
''' ����boolean
Private Function GetCertList(ByRef strName As String, ByRef strCertSN As String, ByRef strCert As String, Optional ByRef strFilePath As String = "0", _
        Optional ByRef strCertDN As String = "0", Optional ByRef strCertUserID As String = "0", Optional ByRef strCertID As String = "0", _
        Optional ByRef strPicCode As String = "0") As Boolean
        '����CA�������ȡ����֤���б���
        '-���:��
        '-����
        'strName :      ����ӿڷ��ص�֤������������
        'strCertSn:   ����ӿڷ��ص�֤��������Ψһ��ʶ
        'strCert:       ����ӿڷ��ص�ǩ��֤��
        Dim strUsbkeyList As String
        Dim arrUserListLength As Integer
        Dim arrUserList() As String
        Dim strPic As String, strUser As String
        Dim strPicID As String
    
        On Error GoTo errH
    
100     If gudtPara.bytSignVersion = V_RSA Then

102    strUsbkeyList = BJCAGX_Client.getUserList()
104    arrUserList = Split(strUsbkeyList, "&&&")
106    arrUserListLength = UBound(arrUserList)
108    If (arrUserListLength = -1) Then
110        MsgBoxEx "��������Key��", vbOKOnly + vbInformation, gstrSysName
           Exit Function
       End If
112    If (arrUserListLength <> 0) Then
           Dim i As Integer
114        For i = 0 To arrUserListLength - 1
               Dim strOption As String
116            strOption = arrUserList(i)
118            strName = Split(strOption, "||")(0)
120            strCertSN = Split(strOption, "||")(1)
122            strCert = BJCAGX_Client.ExportUserCert(strCertSN)
124            If strCertDN <> "0" Then strCertDN = BJCAGX_Client.GetUserInfo(strCert, 20)
           Next
       End If

126     ElseIf gudtPara.bytSignVersion = V_SM2 Then

128        strUsbkeyList = BJCAGX_Client.SOF_GetUserList()
130    If (strUsbkeyList = "") Then
132        strName = ""
134        MsgBoxEx "��������Key��", vbOKOnly + vbInformation, gstrSysName
136        GetCertList = False
           Exit Function
       Else
138            arrUserList = Split(strUsbkeyList, "&&&") 'sm2����2||216000000279373/1003201510002131&&&sm2����1||216000000279349/1003201510002370&&&
140        If UBound(arrUserList) > 1 Then  '���KEY
142            For i = LBound(arrUserList) To UBound(arrUserList) - 1
144                strUser = strUser & "&&&" & Split(arrUserList(i), "||")(0)
               Next
146            If strUser <> "" Then strUser = Mid(strUser, 4)
148            strName = frmSelectUser.ShowMe(strUser)
                
150            For i = LBound(arrUserList) To UBound(arrUserList) - 1
152               If strName = Split(arrUserList(i), "||")(0) Then
154                    strCertID = Split(arrUserList(i), "||")(1)
                       Exit For
                  End If
               Next
           Else
156            arrUserList = Split(arrUserList(0), "||")
158            strName = arrUserList(0)      '֤��CNͨ����
160            strCertID = arrUserList(1)    '֤��ID
           End If

           '��ȡͼƬ������
162        strPicID = Split(strCertID, "/")(0)
    
164            strCert = BJCAGX_Client.SOF_ExportUserCert(strCertID) '3.����ǩ��֤�顣
166            If strCertSN <> "0" Then strCertSN = BJCAGX_Client.SOF_GetCertInfo(strCert, 2) '֤�����к� ǩ��ʱҪ��
168            If strCertDN <> "0" Then strCertDN = BJCAGX_Client.SOF_GetCertInfo(strCert, 33) '֤��DN
170        If strCertUserID <> "0" Then
172                strCertUserID = BJCAGX_Client.SOF_GetCertInfoByOid(strCert, "2.16.840.1.113732.2") '2.��ȡ֤��Ψһ��ʶ��һ��Ϊ���֤�ţ�SF+���֤��
174                strCertUserID = Mid(strCertUserID, 3)
           End If
       End If
        End If
176     If gudtPara.blnSignPic Then
178         If strFilePath <> "0" Or strPicCode <> "0" Then
180            strPic = mobjPic.GetPic(strPicID)
182            If strPic <> "" Then
184                 strPicCode = mobjPic.ConvertPicFormat(strPic, 5)
186                 If strPicCode <> "" And strFilePath <> "0" Then
188                     strFilePath = SaveBase64ToFile("BMP", strCertSN, strPicCode)
                    Else
190                     strFilePath = ""
                    End If
                End If
            End If
        End If
192     GetCertList = True
    
        Exit Function
errH:
194 MsgBoxEx "��ȡ֤����Ϣʧ�ܣ�" & vbNewLine & "��" & CStr(Erl()) & "�� " & Err.Description, vbQuestion, gstrSysName
End Function

''' ���֤����Ч��
''' ����֤����Ч������
Private Function CheckValidaty(ByVal endDate As Date) As Integer
    '����CA��������֤����Ч�Խӿ�
    '-���: ֤����Ч��ֹ����
    '-���Σ���Ч����
        Dim dblAllSp    As Double
        Dim result      As Integer
        dblAllSp = CDbl(CDate(endDate)) - CDbl(Now)
        result = Int(dblAllSp)
        CheckValidaty = result
End Function

Private Function GetReturnInfo(ByVal strSign As Long) As String
    '׼���ʱ���������Ϣת������
    If strSign = -1 Then
        GetReturnInfo = "ʱ�����֤��ͨ��"
    ElseIf strSign = -2 Then
        GetReturnInfo = "ԭ����֤��ͨ��"
    ElseIf strSign = -3 Then
        GetReturnInfo = "���������εĸ�"
    ElseIf strSign = -4 Then
        GetReturnInfo = "֤��δ��Ч"
    ElseIf strSign = -5 Then
        GetReturnInfo = "��ѯ������֤��"
    ElseIf strSign = -6 Then
        GetReturnInfo = "ǩ��ʱ���ʱ������֤�����"
    ElseIf strSign = 0 Then
        GetReturnInfo = "��֤�ɹ�"
    Else
        GetReturnInfo = "δ֪����" & "������:" & strSign
    End If
    If GetReturnInfo <> "" Then
        GetReturnInfo = "ʱ����ӿڷ�����ʾ��" & GetReturnInfo
    End If
End Function

Private Function GetTimeStamp(ByVal strData As String) As String
    Dim year As String, mouth As String, day As String, hour As String, mm As String, ss As String
    Dim strTimeStamp As String

    '��ȡʱ���
    If Len(strData) = 14 Then
        year = Mid(strData, 1, 4)
        mouth = Mid(strData, 5, 2)
        day = Mid(strData, 7, 2)
        hour = Mid(strData, 9, 2)
        mm = Mid(strData, 11, 2)
        ss = Mid(strData, 13, 2)
        strTimeStamp = year & "-" & mouth & "-" & day & " " & hour & ":" & mm & ":" & ss
        If Not IsDate(strTimeStamp) Then
            MsgBoxEx "��ȡ��ʱ�������һ�����ڣ�" & strTimeStamp, vbExclamation, gstrSysName
            GetTimeStamp = ""
            Exit Function
        End If
    End If
    GetTimeStamp = strTimeStamp
End Function

Public Function BJCAGX_GetPara(Optional ByVal bytFunc As Byte) As Boolean
    Dim arrList As Variant
    
    On Error GoTo errH
    gstrPara = gobjComLib.zlDatabase.GetPara(90000, glngSys)  '��ȡURLs �̶���ȡZLHIS'gstrPara = "0&&&0"   '0-RSA;1-SM2&&&0-������ǩ��ͼƬ;1-����ǩ��
    arrList = Split(gstrPara, G_STR_SPLIT)
    If bytFunc = 1 Then
        If gstrPara = "" Or UBound(arrList) < 1 Then
            Err.Raise -1, , "��ǰϵͳ��" & glngSys & "��û�����õ���ǩ���������������á�"
            Exit Function
        End If
    End If
    
    If UBound(arrList) = 0 Then
        gudtPara.bytSignVersion = V_RSA
        gudtPara.blnSignPic = False
        gudtPara.strSignURL = "|"
    ElseIf UBound(arrList) = 1 Then
        gudtPara.bytSignVersion = Val(arrList(0))
        gudtPara.blnSignPic = Val(arrList(1)) = 1
        gudtPara.strSignURL = "|"
    ElseIf UBound(arrList) = 2 Then
        gudtPara.bytSignVersion = Val(arrList(0))
        gudtPara.blnSignPic = Val(arrList(1)) = 1
        gudtPara.strSignURL = arrList(2) '��|�ָ���ʽ�����ǩ�ϴ�URL�ͻ�ȡURL
    End If

    BJCAGX_GetPara = True
    Exit Function
errH:
    MsgBoxEx "��ȡ����ʧ�ܣ�" & vbNewLine & "��" & CStr(Erl()) & "�� " & Err.Description, vbInformation, gstrSysName
End Function

Public Function BJCAGX_SetParaStr() As String
    BJCAGX_SetParaStr = IIf(gudtPara.bytSignVersion = 0, "0", "1") & G_STR_SPLIT & IIf(gudtPara.blnSignPic, "1", "0") & G_STR_SPLIT & IIf(gudtPara.strSignURL = "", "|", gudtPara.strSignURL)
End Function



