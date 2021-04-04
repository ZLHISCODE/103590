Attribute VB_Name = "mdlBJCASC"
Option Explicit
'����CA���Ĺ���ģ��(�Ĵ���)
Private mblnInit As Boolean         '�Ƿ��ѳ�ʼ���ɹ�
Private mLastPWD As String          '�������������

Private BJCASC_Client As Object       '֤�鲿��
Private BJCASC_svs As Object          '

Public Function BJCASC_InitObj() As Boolean
    '֤�鲿����ʼ��
        Dim progID As String
        
        On Error GoTo errH
100     mLastPWD = ""
102     BJCASC_InitObj = mblnInit
104     If mblnInit Then Exit Function
    
106     Set BJCASC_svs = CreateObject("BJCA_SVS_ClientCOM.BJCASVSEngine.1")
108     Set BJCASC_Client = CreateObject("BJCASecCOM.BJCASecCOMV2.1")
110
114     BJCASC_InitObj = True
    
116     mblnInit = BJCASC_InitObj
        Exit Function
errH:
118     MsgBoxEx "�����ӿڲ���ʧ�ܣ�" & vbNewLine & Err.Description, vbQuestion, gstrSysName
    
End Function

Public Function BJCASC_RegCert(arrCertInfo As Variant) As Boolean
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

124         BJCASC_RegCert = True
        End If

        Exit Function
errH:
126     MsgBoxEx "��ȡ֤����Ϣʧ�ܣ�" & vbNewLine & "��" & CStr(Erl()) & "�� " & Err.Description, vbQuestion, gstrSysName

End Function

Public Function BJCASC_Sign(ByVal strCurrCertSn As String, ByVal strSource As String, ByRef strSignData As String, ByRef strTimeStamp As String) As Boolean
        'ǩ��
        Dim strSigCert As String

        On Error GoTo errH
100     If BJCASC_CheckCert(strCurrCertSn, strSigCert) Then               '��֤��ǰUSB�Ƿ���ǩ���û��ģ�����ȡǩ��֤��
110         strSignData = BJCASC_Client.SignData(strCurrCertSn, strSource)
112         BJCASC_Sign = True
        Else
            MsgBoxEx "ǩ��ʧ�ܣ�"
        End If
        Exit Function
errH:
114     MsgBoxEx "ǩ��ʧ�ܣ�" & vbNewLine & "��" & CStr(Erl()) & "�� " & Err.Description, vbQuestion, gstrSysName
End Function

Public Function BJCASC_VerifySign(ByVal strCurrCertSn As String, ByVal strSignData As String, ByVal strSource As String) As Boolean
        '��֤ǩ��
        Dim strSigCert As String, strTmp As String
        On Error GoTo errH
100     If BJCASC_CheckCert(strCurrCertSn, strSigCert) Then           '��֤��ǰUSB�Ƿ���ǩ���û��ģ�����ȡǩ��֤��
102         BJCASC_VerifySign = GetCertVerifySign(strSignData, strSigCert, strSource, strTmp)
        Else
            MsgBoxEx "��֤ǩ��ʧ�ܣ�"
        End If
        Exit Function
errH:
104     MsgBoxEx "��֤ǩ��ʧ�ܣ�" & vbNewLine & "��" & CStr(Erl()) & "�� " & Err.Description, vbQuestion, gstrSysName
End Function


Public Function BJCASC_CheckCert(ByVal strCurrCertSn As String, Optional ByRef strSigCert As String) As Boolean
        '���ܣ���ȡUSB�����豸��ʼ������¼
        Dim strKey As String, strPIN As String, strUserName As String
        Dim strWebUrl As String, intDate   As Integer
        On Error GoTo errH
100     If Not mblnInit Then
102         MsgBoxEx "����δ��ʼ����"
            Exit Function
        End If
    
104     Call GetCertList(strUserName, strKey, strSigCert)
106     If strCurrCertSn <> strKey Then
108         MsgBoxEx "��֤��δע�����������£�����ʹ�ã�"
            Exit Function
        End If
110     If mLastPWD <> "" Then strPIN = mLastPWD
112     If strPIN = "" Then
114         If Not frmPassword.ShowMe(strPIN) Then Exit Function
        End If
        
116     If Not GetCertLogin(strKey, strPIN, strSigCert, intDate, strWebUrl) Then
118         strPIN = ""
             
        Else
            BJCASC_CheckCert = True
        End If
     
120     mLastPWD = strPIN
122
    
        Exit Function
errH:
124     MsgBoxEx "���USBKEYʧ�ܣ�" & vbNewLine & "��" & CStr(Erl()) & "�� " & Err.Description, vbQuestion, gstrSysName
End Function

Public Sub BJCASC_UloadObj()
    Set BJCASC_Client = Nothing
    Set BJCASC_svs = Nothing

    mblnInit = False
End Sub
'----- �������ڲ�����

''' ��ȡ�ͻ���֤���б�
''' ����boolean
Private Function GetCertList(ByRef strName As String, ByRef strUniqueID As String, ByRef strCert As String) As Boolean
    '����CA�Ĵ����ȡ����֤���б���
    '-���:��
    '-����
    'strName :      ����ӿڷ��ص�֤������������
    'strUniqueID:   ����ӿڷ��ص�֤��������Ψһ��ʶ
    'strCert:       ����ӿڷ��ص�ǩ��֤��
      
    Dim strUsbkeyList As String
    Dim arrUserListLength As Integer
    Dim arrUserList() As String
    If BJCASC_Client Is Nothing Then Set BJCASC_Client = CreateObject("BJCASecCOM.BJCASecCOMV2.1")
    strUsbkeyList = BJCASC_Client.getUserList()
    arrUserList = Split(strUsbkeyList, "&&&")
    arrUserListLength = UBound(arrUserList)
    If (arrUserListLength = -1) Then
        MsgBoxEx "��������Key��"
        Exit Function
    End If
    If (arrUserListLength <> 0) Then
        Dim i As Integer
        For i = 0 To arrUserListLength - 1
            Dim strOption As String
            strOption = arrUserList(i)
            strName = Split(strOption, "||")(0)
            strUniqueID = Split(strOption, "||")(1)
            strCert = BJCASC_Client.ExportUserCert(strUniqueID)
        Next
    End If
    GetCertList = True
End Function

Private Function GetCertLogin(ByVal strUniqueID As String, ByVal strPassword As String, ByVal strCert As String, ByRef dDate As Integer, ByRef strWebserviceUrl As String) As Boolean
    '����CA�Ĵ�������֤���¼����
    '- ���
    'strUniqueID : ֤��Ψһ��ʶ
    'strPassword : ֤������
    'strWebserviceUrl:ǩ����������ַ����Ϊ֤����֤
    '- ����
    'dDate       : ����֤����Чʱ��

    Dim result As Boolean
    If BJCASC_Client Is Nothing Then Set BJCASC_Client = CreateObject("BJCASecCOM.BJCASecCOMV2.1")
    If (strPassword = "") Then
        MsgBoxEx "������֤�����룡"
    Else
        '֤�鰲ȫ��¼
        'result:0:�ɹ�
        'result:��0:���ɹ�
        result = BJCASC_Client.userLogin(strUniqueID, strPassword)
        If (result) Then
            Dim strExtLib As String
            strExtLib = BJCASC_Client.GetUserInfo(strUniqueID, 15)
            Dim intFlg As Integer
            
            '����������֤֤��
            '������е���֤��
            Dim retValidateCert As Long
            retValidateCert = 100
            retValidateCert = ValidateCert(strCert, strWebserviceUrl)
            
            '��֤֤������Ϣ��ʾ
            If retValidateCert <> 0 Then Call ValidateCertView(retValidateCert)

            If (retValidateCert = 0) Then
                Dim uniqueIdStr As String
                Dim oid As String
                oid = "2.16.840.1.113732.2"
                Dim s As String
                '��ȡ�ͻ���֤����Ч�ڽ�ֹʱ��
                s = BJCASC_Client.GetCertInfo(strCert, 12)
                '��֤�ͻ���֤����Ч��ʣ������
                dDate = CheckValidaty(s)
            
                If (dDate <= 30 And dDate > 0 And Not gblnShow) Then
                    MsgBoxEx "����֤�黹��" & dDate & "�����"
                    uniqueIdStr = BJCASC_Client.GetCertInfoByOid(strCert, oid)
                    gblnShow = True
                    GetCertLogin = True
                ElseIf (dDate <= 0) Then
                    MsgBoxEx "����֤���ѹ��� " & Abs(dDate) & " ��"
                    GetCertLogin = False
                Else
                    uniqueIdStr = BJCASC_Client.GetCertInfoByOid(strCert, oid)
                    
                    GetCertLogin = True
                End If
            Else
                GetCertLogin = False
            End If
        Else
            GetCertLogin = False
            
        End If
    End If

End Function

Private Function ValidateCert(ByRef userCert As String, Optional webserviceUrl As String) As Integer
    '����������֤֤��
 
    If BJCASC_svs Is Nothing Then Set BJCASC_svs = CreateObject("BJCA_SVS_ClientCOM.BJCASVSEngine.1")
    ValidateCert = BJCASC_svs.ValidateCertificate(userCert)
 
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
Private Function GetCertVerifySign(ByVal strInData As String, ByVal strCert As String, ByRef strData As String, ByRef strOut As String) As Boolean
    '����CA�Ĵ�������֤��ǩ����֤����
    '- ���
    'strInData     : ǩ�����
    'strCert       : ǩ��֤��
    'strData       : ǩ��ԭ��
    '- ����
    'strOut       : ������ǩ���

    'result:true:  �ɹ�
    'result:false: ʧ��
    Dim verifySignResult As Boolean
    If BJCASC_Client Is Nothing Then Set BJCASC_Client = CreateObject("BJCASecCOM.BJCASecCOMV2.1")
    verifySignResult = BJCASC_Client.VerifySignedData(strCert, strData, strInData)
    If (verifySignResult) Then
        strOut = "��֤ǩ���ɹ���"
        GetCertVerifySign = True
    Else
        strOut = "��֤ǩ��ʧ�ܣ�"
        GetCertVerifySign = False
    End If
End Function

''' ���֤����Ч��
''' ����֤����Ч������
Private Function CheckValidaty(ByVal endDate As Date) As Integer
    '����CA�Ĵ�����֤����Ч�Խӿ�
    '-���: ֤����Ч��ֹ����
    '-���Σ���Ч����
        Dim dblAllSp    As Double
        Dim result      As Integer
        dblAllSp = CDbl(CDate(endDate)) - CDbl(Now)
        result = Int(dblAllSp)
        CheckValidaty = result
End Function


