Attribute VB_Name = "mdlGDCA"
Option Explicit
Private mblnInit As Boolean         '�Ƿ��ѳ�ʼ���ɹ�
Private mobjGDCA As Object          '����CA �ؼ�
Private mLastPIN As String          '�������������

Public Function GDCA_initObj() As Boolean
    
        '���ܣ� �����ӿڲ���
        On Error GoTo errH
100     mLastPIN = ""
102     GDCA_initObj = mblnInit
104     If mblnInit Then Exit Function
    
        '���ֳ�����ʱ��������Ʋ��ԣ����޸�
106     Set mobjGDCA = CreateObject("Atl_com.Gdca.1")    '��Ϊ�ĵ���û��˵�����������Ʋ�ȷ��
108     GDCA_initObj = True
    
110     mblnInit = GDCA_initObj
        Exit Function
errH:
112     MsgBoxEx "���ó�ʼ���ӿ�ʧ�ܣ�" & vbNewLine & Err.Description, vbQuestion, gstrSysName
End Function


Public Function GDCA_RegCert(arrCertInfo As Variant) As Boolean
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
        On Error GoTo errH
    
100     For i = LBound(arrCertInfo) To UBound(arrCertInfo)
102         arrCertInfo(i) = ""
        Next
    
104     If ReadUSBKeyLogin Then
106         If VerifyCert(strKeyId, strCertDN, strCertUserName, strSigCert) Then
108             arrCertInfo(0) = strCertUserName
110             arrCertInfo(1) = strCertDN
112             arrCertInfo(2) = strKeyId
114             arrCertInfo(3) = strSigCert
116             GDCA_RegCert = True
            End If
        End If

        Exit Function
errH:
118     MsgBoxEx "��ȡ֤����Ϣʧ�ܣ�" & vbNewLine & "��" & CStr(Erl()) & "�� " & Err.Description, vbQuestion, gstrSysName
End Function

Public Function GDCA_Sign(ByVal strCurrCertSn As String, ByVal strSource As String, ByRef strSignData As String, ByRef strTimeStamp As String) As Boolean
        'ǩ��
        Dim strTmp As String, strEndData As String, strSigCert As String
        On Error GoTo errH
100     If GDCA_CheckCert(strCurrCertSn, strSigCert) Then       '��֤��ǰUSB�Ƿ���ǩ���û��ģ�����ȡǩ��֤��
102         'strTmp = mobjGDCA.AppGetTime()      '��ȡʱ���,��ʽδ֪
104         'If IsDate(strTmp) Then strTimeStamp = strTmp        '�����ڸ�ʽ�ŷ���
        
106         strEndData = strSource '& strTmp                        '��ʱ������ʱ���
108         strEndData = mobjGDCA.GDCA_Base64Encode(strEndData)     '��ԭʼ���ݱ���
        
110         strSignData = mobjGDCA.GDCA_Pkcs7Sign("LAB_USERCERT_SIG", 4, strSigCert, strEndData)    'ǩ��
        
112         GDCA_Sign = True
        End If
        Exit Function
errH:
114     MsgBoxEx "ǩ��ʧ�ܣ�" & vbNewLine & "��" & CStr(Erl()) & "�� " & Err.Description, vbQuestion, gstrSysName
End Function

Public Function GDCA_VerifySign(ByVal strCurrCertSn As String, ByVal strSignData As String, ByVal strSource As String) As Boolean
        '��֤ǩ��
        Dim strSigCert As String, strTmp As String, strEndData
        On Error GoTo errH
100     If GDCA_CheckCert(strCurrCertSn, strSigCert) Then         '��֤��ǰUSB�Ƿ���ǩ���û��ģ�����ȡǩ��֤��
            strEndData = mobjGDCA.GDCA_Base64Encode(strSource)
102         strTmp = mobjGDCA.GDCA_Pkcs7Verify(strSigCert, strSignData)
104         If strTmp = strEndData Then GDCA_VerifySign = True
        End If
        Exit Function
errH:
106     MsgBoxEx "��֤ǩ��ʧ�ܣ�" & vbNewLine & "��" & CStr(Erl()) & "�� " & Err.Description, vbQuestion, gstrSysName
End Function

Public Function GDCA_CheckCert(ByVal strCurrCertSn As String, Optional ByRef strSigCert As String) As Boolean
    
        '��֤��ǰ�����USBKey�Ƿ��ǵ�ǰ�û���,����ɹ����򷵻�ǩ��֤��
        Dim strKeyId As String, strCertTime As String, strCertUserName As String, strCertDN As String
        Dim i As Integer, strCACert As String, lngOk As Long
    
        On Error GoTo errH
    
100     GDCA_CheckCert = False

102     If ReadUSBKeyLogin Then
104         If VerifyCert(strKeyId, strCertDN, strCertUserName, strSigCert) Then
106             If strCurrCertSn = strKeyId Then GDCA_CheckCert = True
            End If
        End If

        Exit Function
errH:
108     MsgBoxEx "��ȡ֤����Ϣʧ�ܣ�" & vbNewLine & "��" & CStr(Erl()) & "�� " & Err.Description, vbQuestion, gstrSysName
End Function
Public Sub GDCA_UnloadObj()
    Set mobjGDCA = Nothing
End Sub
'------------------------------------------------------------------------------------
'-- ������GDCAģ���ڲ����ù���
'------------------------------------------------------------------------------------
Private Function ReadUSBKeyLogin() As Boolean
        '���ܣ���ȡUSB�����豸��ʼ������¼
        Dim strKey As String, strPIN As String, lngOk As Long
        Dim blnOk As Boolean, lngCount As Long
        On Error GoTo errH
100     If Not mblnInit Then
102         MsgBoxEx "GDCA����δ��ʼ����"
            Exit Function
        End If
    
104     strKey = mobjGDCA.GDCA_GetDevicType             'ȡ���뱾�����豸����
106     Call mobjGDCA.GDCA_SetDeviceType(strKey)        '����KEY�����豸
108     Call mobjGDCA.GDCA_Initialize                   '��ʼ��
    
110     If mLastPIN <> "" Then strPIN = mLastPIN
112     If strPIN = "" Then
114         If Not frmPassword.ShowMe(strPIN) Then Exit Function
        End If
        
        '�û�PINΪ1-8���ֽڣ����û����������
116     lngCount = 0: blnOk = False: lngOk = -1
    
118     Do While lngCount <= 3 And Not blnOk
120         lngOk = mobjGDCA.GDCA_Login(2, strPIN)                                    '��¼
122         If lngOk = 0 Then
124             blnOk = True
            Else
126             If Not frmPassword.ShowMe(strPIN) Then Exit Function
            End If
128         lngCount = lngCount + 1
        Loop
    
130     mLastPIN = strPIN
132     ReadUSBKeyLogin = True
    
        Exit Function
errH:
134     MsgBoxEx "��ʼ��KEYʧ�ܣ�" & vbNewLine & "��" & CStr(Erl()) & "�� " & Err.Description, vbQuestion, gstrSysName
End Function


Private Function VerifyCert(Optional CertSN As String, Optional CertDn As String, Optional CertUN As String, Optional SigCert As String) As Boolean
        '��֤ǩ��������ȡ֤������
    
        Dim strKeyId As String, strCertTime As String, strCertUserName As String, strCertDN As String
        Dim strSigCert As String, i As Integer, strCACert As String, lngOk As Long
    
100     strKeyId = mobjGDCA.GDCA_ReadLabel("LAB_DISAID", 3)            ' ��ȡ֤��Ψһ��ʶ
102     strKeyId = mobjGDCA.GDCA_Base64Decode(strKeyId)
        
104     strCACert = mobjGDCA.GDCA_ReadLabel("CA_CERT", 9)                   '9-��ȡCA֤��
106     strSigCert = mobjGDCA.GDCA_ReadLabel("LAB_USERCERT_SIG", 7)         '7-ǩ��֤��
        
        
108     strCertTime = mobjGDCA.GDCA_GetCertificateInfo(strSigCert, 6)       '6-ȡ֤����Ч��
110     strCertUserName = mobjGDCA.GDCA_GetCertificateInfo(strSigCert, 7)   '7-ȡ֤�����������
112     strCertDN = mobjGDCA.GDCA_GetCertificateInfo(strSigCert, 3)         '3-ȡ֤�����к�
        
114     lngOk = -1
116     lngOk = mobjGDCA.GDCA_VerifyCert(strSigCert, strCACert)                 '��֤֤��
118     If lngOk = 0 Then
120         CertUN = strCertUserName
122         CertDn = strCertDN
124         CertSN = strKeyId
126         SigCert = strSigCert
128         VerifyCert = True
        Else
130         MsgBoxEx "֤����֤ʧ�ܣ�֤��Ч��" & strCertTime, vbQuestion, gstrSysName
132         VerifyCert = False
        End If
        Exit Function
errH:
134     MsgBoxEx "��֤֤��ʧ�ܣ�" & vbNewLine & "��" & CStr(Erl()) & "�� " & Err.Description, vbQuestion, gstrSysName

End Function

