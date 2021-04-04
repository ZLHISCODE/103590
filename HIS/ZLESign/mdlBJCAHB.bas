Attribute VB_Name = "mdlBJCAHB"
Option Explicit

'����CA���Ĺ���ģ��(�º�����)
Private mblnInit As Boolean         '�Ƿ��ѳ�ʼ���ɹ�

Private BJCA_Pic As Object
Private BJCAHB_Client As Object       '�ͻ���֤�鲿��
Private BJCAHB_svs As Object          'ǩ����֤�ؼ�
Private BJCAHB_TS As Object           'ʱ����ؼ�

Public Function BJCAHB_InitObj() As Boolean
    '֤�鲿����ʼ��
        Dim progID As String
        On Error GoTo errH
     BJCAHB_InitObj = mblnInit

     If mblnInit Then Exit Function
100     Set BJCAHB_Client = CreateObject("XTXAppCOM.XTXApp.1")
101     Set BJCA_Pic = CreateObject("GetKeyPic.GetPic")
102     Set BJCAHB_svs = CreateObject("BJCA_SVS_ClientCOM.BJCASVSEngine.1")   '����֤����֤�ؼ�
103     Set BJCAHB_TS = CreateObject("BJCA_TS_ClientCom.BJCATSEngine")    '����ʱ����ؼ�
     BJCAHB_InitObj = True
     mblnInit = BJCAHB_InitObj
    Exit Function
errH:
     MsgBoxEx "�����ӿڲ���ʧ�ܣ�" & vbNewLine & Err.Description, vbQuestion, gstrSysName
End Function
Public Function BJCAHB_RegCert(arrCertInfo As Variant) As Boolean
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
    
        For i = LBound(arrCertInfo) To UBound(arrCertInfo)
             arrCertInfo(i) = ""
        Next
        
        If GetCertList(strCertUserName, strKeyId, strSigCert) Then
            arrCertInfo(0) = strCertUserName
            arrCertInfo(1) = BJCAHB_Client.SOF_GetCertInfoByOid(strSigCert, "1.2.156.112562.2.1.1.1") '2.��ȡ֤��Ψһ��ʶ��һ��Ϊ���֤�ţ�
            'arrCertInfo(1) = BJCAHB_Client.SOF_GetCertInfo(strSigCert, 33)
            arrCertInfo(2) = BJCAHB_Client.SOF_GetCertInfo(strSigCert, 2)
            arrCertInfo(3) = strSigCert
            arrCertInfo(5) = SaveBase64ToFile("gif", strKeyId, BJCA_Pic.getpic())
            BJCAHB_RegCert = True
        End If

        Exit Function
errH:
     MsgBoxEx "��ȡ֤����Ϣʧ�ܣ�" & vbNewLine & "��" & CStr(Erl()) & "�� " & Err.Description, vbQuestion, gstrSysName
End Function

Public Function BJCAHB_CheckCert(ByVal strCurrCertSn As String, Optional ByRef strSigCert As String, Optional ByRef CertID As String) As Boolean
    '���ܣ���ȡUSB�����豸��ʼ������¼
     Dim strKey As String, strPIN As String, strUserName As String
     Dim strWebUrl As String, intDate   As Integer
     Dim random
     Dim strClientSignedData, strUsbkeyList As String, strUniqueID As String
     Dim arrUserList() As String
     On Error GoTo errH
     If Not BJCAHB_InitObj Then
        MsgBoxEx "��鲿���Ƿ��ʼ����", vbInformation + vbOKOnly, gstrSysName
        Exit Function
     End If
      '��ȡ֤��
     strUsbkeyList = BJCAHB_Client.SOF_GetUserList()
     If (strUsbkeyList = "") Then
        MsgBoxEx "�����֤��Key��"
        BJCAHB_CheckCert = False
        Exit Function
     Else
        arrUserList = Split(strUsbkeyList, "&&&")
        arrUserList = Split(arrUserList(0), "||")
        CertID = arrUserList(1)
        strSigCert = BJCAHB_Client.SOF_ExportUserCert(arrUserList(1)) '3.����ǩ��֤�顣
        strUniqueID = BJCAHB_Client.SOF_GetCertInfoByOid(strSigCert, "1.2.156.112562.2.1.1.1") '2.��ȡ֤��Ψһ��ʶ��һ��Ϊ���֤�ţ�
     End If
     If strCurrCertSn <> BJCAHB_Client.SOF_GetCertInfo(strSigCert, 2) Then
        MsgBoxEx "��֤��δע�����������£�����ʹ�ã�"
        Exit Function
     End If
     random = BJCAHB_Client.SOF_GenRandom(24)
     strClientSignedData = BJCAHB_Client.SOF_SignData(CertID, random)
     If Not GetCertLogin(strUniqueID, strClientSignedData, strSigCert, intDate, strWebUrl) Then
         BJCAHB_CheckCert = False
     Else
         BJCAHB_CheckCert = True
     End If
    Exit Function
errH:
     MsgBoxEx "���USBKEYʧ�ܣ�" & vbNewLine & "��" & CStr(Erl()) & "�� " & Err.Description, vbQuestion, gstrSysName
End Function

Public Function BJCAHB_Sign(ByVal strCurrCertSn As String, ByVal strSource As String, ByRef strSignData As String, _
    ByRef strTimeStamp As String, ByRef strTimeStampCode As String) As Boolean
        'ǩ��
    Dim strSigCert As String
    Dim CertID As String
    Dim strRequest As String    'ʱ�������
    Dim strErr As String
    Dim strDate As String
    
    If BJCAHB_TS Is Nothing Then Set BJCAHB_TS = CreateObject("BJCA_TS_ClientCom.BJCATSEngine")  '����ʱ����ؼ�
    If Err.Number <> 0 Then
        MsgBoxEx "ʱ����ؼ�û�а�װ��", vbExclamation, gstrSysName
        Exit Function
    End If
    On Error GoTo errH
    If BJCAHB_CheckCert(strCurrCertSn, strSigCert, CertID) Then                '��֤��ǰUSB�Ƿ���ǩ���û��ģ�����ȡǩ��֤��
        strSignData = BJCAHB_Client.SOF_SignData(CertID, strSource) '����ǩ������
        If strSignData <> "" Then
            strRequest = BJCAHB_TS.CreateTimeStampRequest(strSignData) '����ʱ�������
            If strRequest <> "" Then
                strTimeStampCode = BJCAHB_TS.CreateTimeStamp(strRequest)  '����ʱ�������֤�飩
                If strTimeStampCode = "" Then
                    strErr = "ʱ�������Ϊ�գ�"
                Else
                    strDate = BJCAHB_TS.gettimestampinfo(strTimeStampCode, 1)
                    strTimeStamp = String14ToDate(strDate, strErr)   'ȡ��ʱ���ʱ��
                End If
            Else
                strErr = "ʱ�������ʧ�ܣ�"
            End If
        Else
            strErr = "ǩ��ʧ�ܣ�"
        End If
    Else
        strErr = "��֤֤��ʧ�ܣ�"
    End If
    
    If strErr <> "" Then
        MsgBoxEx strErr, vbOKOnly + vbInformation, gstrSysName
        BJCAHB_Sign = False
        Exit Function
    End If
    BJCAHB_Sign = True
    Exit Function
errH:
     MsgBoxEx "ǩ��ʧ�ܣ�" & vbNewLine & "��" & CStr(Erl()) & "�� " & Err.Description, vbQuestion, gstrSysName
End Function
'''��֤ǩ������
Public Function BJCAHB_VerifySign(ByVal strCert As String, _
    ByVal strSignData As String, ByVal strSource As String, ByVal strStampCode As String) As Boolean
'��֤ǩ��
    Dim blnRet As Boolean
    Dim strMsg As String
    Dim lngRuslt As Long
    
    On Error GoTo errH
    
    If strStampCode <> "" Then
        lngRuslt = BJCAHB_TS.verifyTimeStamp(strStampCode)
        If lngRuslt <> 0 Then
            MsgBoxEx "��֤ʱ���ʧ�ܣ�" & GetReturnInfo(lngRuslt), vbInformation + vbOKOnly, gstrSysName
            Exit Function
        End If
    End If
    blnRet = BJCAHB_Client.SOF_VerifySignedData(strCert, strSource, strSignData)
    If blnRet Then
        strMsg = "��֤ǩ���ɹ���"
    Else
        strMsg = "��֤ǩ��ʧ�ܣ�"
    End If
    MsgBoxEx strMsg, vbOKOnly + vbInformation, gstrSysName
    BJCAHB_VerifySign = blnRet
    Exit Function
errH:
104     MsgBoxEx "��֤ǩ��ʧ�ܣ�" & vbNewLine & "��" & CStr(Erl()) & "�� " & Err.Description, vbQuestion, gstrSysName
End Function
'���ٶ���
Public Sub BJCAHB_UloadObj()
    Set BJCAHB_Client = Nothing
    Set BJCAHB_svs = Nothing
    Set BJCAHB_TS = Nothing
    mblnInit = False
End Sub

'----- �������ڲ�����
Private Function GetCertLogin(ByVal strUniqueID As String, ByVal strClientSignedData As String, ByVal strCert As String, ByRef dDate As Integer, ByRef strWebserviceUrl As String) As Boolean
    '����CA����������֤���¼����
    '- ���
    'strUniqueID            :֤��Ψһ��ʶ
    'strClientSignedData    :ǩ������
    'strWebserviceUrl       :ǩ����������ַ����Ϊ֤����֤
    '- ����
    'dDate       : ����֤����Чʱ��

    If BJCAHB_svs Is Nothing Then Set BJCAHB_svs = CreateObject("BJCA_SVS_ClientCOM.BJCASVSEngine.1")
    '֤�鰲ȫ��¼
    'strClientSignedData:�ǿ�:�ɹ�
    'strClientSignedData:��:���ɹ�
    If (strClientSignedData <> "") Then
        '����������֤֤��
        '������е���֤��
        Dim retValidateCert As Long
        retValidateCert = ValidateCert(strCert)
        '��֤֤������Ϣ��ʾ
        If retValidateCert <> 0 Then
            Call ValidateCertView(retValidateCert)
            Exit Function
        ElseIf (retValidateCert = 0) Then
            Dim s As String
            '��ȡ�ͻ���֤����Ч�ڽ�ֹʱ��
            s = BJCAHB_Client.SOF_GetCertInfo(strCert, 12)
            s = String14ToDate(s)
            If s <> "" Then
            '��֤�ͻ���֤����Ч��ʣ������
                dDate = CheckValidaty(CDate(s))
            
                If (dDate <= 30 And dDate > 0 And Not gblnShow) Then
                    MsgBoxEx "����֤�黹��" & dDate & "�����"
                    gblnShow = True
                    GetCertLogin = True
                ElseIf (dDate <= 0) Then
                    MsgBoxEx "����֤���ѹ��� " & Abs(dDate) & " ��"
                    GetCertLogin = False
                Else
                    GetCertLogin = True
                End If
            End If
        End If
        
    End If
End Function
Private Function ValidateCert(ByRef userCert As String) As Integer
    '����������֤֤��
    If BJCAHB_svs Is Nothing Then Set BJCAHB_svs = CreateObject("BJCA_SVS_ClientCOM.BJCASVSEngine.1")
    ValidateCert = BJCAHB_svs.ValidateCertificate(userCert)
 
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
        Case Else
            MsgBoxEx "��֤������֤��ʧ�ܣ���֤����ֵ:" & retValidateCert
    End Select
End Sub

''' ��ȡ�ͻ���֤���б�
''' ����boolean
Private Function GetCertList(ByRef strName As String, ByRef strUniqueID As String, ByRef strCert As String) As Boolean
    '����CA�������ȡ����֤���б���
    '-���:��
    '-����
    'strName :      ����ӿڷ��ص�֤������������
    'strUniqueID:   ����ӿڷ��ص�֤��������Ψһ��ʶ
    'strCert:       ����ӿڷ��ص�ǩ��֤��
    Dim strUsbkeyList As String
    Dim arrUserListLength As Integer
    Dim arrUserList() As String
      '��ȡ֤��
    strUsbkeyList = BJCAHB_Client.SOF_GetUserList()
    If (strUsbkeyList = "") Then
        strName = ""
        MsgBoxEx "�����֤��Key��"
        GetCertList = False
        Exit Function
    Else
        arrUserList = Split(strUsbkeyList, "&&&")
        arrUserList = Split(arrUserList(0), "||")
        strName = arrUserList(0)
        strCert = BJCAHB_Client.SOF_ExportUserCert(arrUserList(1)) '3.����ǩ��֤�顣
        strUniqueID = BJCAHB_Client.SOF_GetCertInfoByOid(strCert, "1.2.156.112562.2.1.1.1") '2.��ȡ֤��Ψһ��ʶ��һ��Ϊ���֤�ţ�
    End If
    GetCertList = True
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
        GetReturnInfo = "δ֪����"
    End If
    If GetReturnInfo <> "" Then
        GetReturnInfo = "ʱ����ӿڷ�����ʾ��" & GetReturnInfo
    End If
End Function



