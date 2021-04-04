Attribute VB_Name = "mdlSCCA"
Option Explicit
'�Ĵ�CA����ģ��
Private mblnInit As Boolean         '�Ƿ��ѳ�ʼ���ɹ�
Private SCCA_Client As Object       '�ͻ���֤�鲿��
Private SCCA_Server As Object       '��������֤�鲿��

Public Function SCCA_InitObj() As Boolean
    '֤�鲿����ʼ��
        Dim progID As String
        
        On Error GoTo errH
        SCCA_InitObj = mblnInit
        If mblnInit Then Exit Function
        Set SCCA_Client = CreateObject("wsb.WsbClient")
        Set SCCA_Server = CreateObject("wsbServer.wsbServerClass")
        If SCCA_Client Is Nothing Then
            SCCA_InitObj = False
            MsgBoxEx "��ʼ��CA" & "wsb.WsbClient�ؼ�ʧ��!", vbInformation, gstrSysName
            Exit Function
        End If
        If SCCA_Server Is Nothing Then
            SCCA_InitObj = False
            MsgBoxEx "��ʼ��CA" & "wsbServer.wsbServerClass�ؼ�ʧ��!", vbInformation, gstrSysName
            Exit Function
        End If
        '��ʼ���ɹ���
        SCCA_InitObj = True
    
        mblnInit = SCCA_InitObj
        Exit Function
errH:
    Call GetErrMsg(Erl())
End Function

Public Function SCCA_RegCert(arrCertInfo As Variant) As Boolean
        '���ܣ��ṩ��HIS���ݿ���ע�����֤��ı�Ҫ��Ϣ,�����·��Ż����֤��,,��Ҫ����USB-Key
        '���أ�arrCertInfo��Ϊ���鷵��֤�������Ϣ
        '      0-ClientSignCertCN:�ͻ���ǩ��֤�鹫������(����),����ע��֤��ʱ������֤���
        '      1-ClientSignCertDN:�ͻ���ǩ��֤������(ÿ��Ψһ)
        '      2-ClientSignCertSN:�ͻ���ǩ��֤�����к�(ÿ֤��Ψһ)
        '      3-ClientSignCert:�ͻ���ǩ��֤������
        '      4-ClientEncCert:�ͻ��˼���֤������
        '      5-ǩ��ͼƬ�ļ���,�մ���ʾû��ǩ��ͼƬ
        
        Dim strKeyId As String, strCertUserName As String, strEncCert As String, strCertSn As String
        Dim strSigCert As String, i As Integer, strCACert As String, lngOk As Long
        Dim strPicData As String
        On Error GoTo errH
    
100     For i = LBound(arrCertInfo) To UBound(arrCertInfo)
102         arrCertInfo(i) = ""
        Next
    
104     If GetCertList(strCertUserName, strKeyId, strSigCert, strCertSn) Then
106         arrCertInfo(0) = strCertUserName
108         arrCertInfo(1) = GetCertDN(strSigCert)
110         arrCertInfo(2) = strCertSn
112         arrCertInfo(3) = strSigCert
124         SCCA_RegCert = True
        End If

        Exit Function
errH:
126     MsgBoxEx "��ȡ֤����Ϣʧ�ܣ�" & Err.Description, vbInformation, gstrSysName
End Function

Public Function GetCertDN(strCert As String) As String
    Dim strInfo As String
    Dim i As Long
    strInfo = SCCA_Client.SOF_GetCertInfo(strCert, 33)
    If strInfo <> "" Then
        GetCertDN = strInfo
    Else
        Exit Function
    End If
End Function

''' ��ȡ�ͻ���֤���б�
''' ����boolean
Private Function GetCertList(ByRef strName As String, ByRef strUniqueID As String, ByRef strCert As String, ByRef strCertSn As String) As Boolean
    '-���:��
    '-����
    'strName :      ����ӿڷ��ص�֤������������
    'strUniqueID:   ����ӿڷ��ص�֤��������Ψһ��ʶ
    'strCert:       ����ӿڷ��ص�ǩ��֤��
    'strCertSn      ����֤����Ϣ
    Dim strPassas As String
    Dim strList As String '�Ѱ�װ֤���û��б�
    Dim arrList() As String
    On Error GoTo errH
    strList = SCCA_Client.SOF_GetUserList()
    If Trim(strList) <> "" Then
        strList = Replace(strList, "||", "|")
        strList = Replace(strList, "&&&", "&")
        arrList = Split(strList, "|")
        '֤����Ϣ
        strCert = SCCA_Client.SOF_ExportUserCert(arrList(1)) '֤���ַ���
        If strCert <> "" Then
            strUniqueID = SCCA_Client.SOF_GetCertInfo(strCert, 53) 'Ψһ��ʶ
            strName = SCCA_Client.SOF_GetCertInfo(strCert, 23) '֤��ͨ��������
            strCertSn = SCCA_Client.SOF_GetCertInfo(strCert, 2) '֤�����к�
        End If
        GetCertList = True
    Else
        MsgBoxEx "û���ҵ�Key�̣����飡", vbInformation, gstrSysName
        Exit Function
    End If
    Exit Function
errH:
    GetCertList = False
End Function

Public Function SCCA_CheckCert(ByVal strCurrCertSn As String, Optional ByRef strSigCert As String, Optional ByRef strCertSn As String, Optional ByRef blnReDo As Boolean) As Boolean
    '���ܣ���ȡUSB�����豸��ʼ������¼
    Dim strKey As String, strPIN As String, strUserName As String, strDate As String
    Dim blnRet As Boolean, intDate As Date
    Dim udtUser As USER_INFO
    Dim intPoint As Integer
    Dim strArry() As String
    On Error GoTo errH
    If Not SCCA_InitObj() Then
        MsgBoxEx "����δ��ʼ����"
        Exit Function
    End If
    If Not GetCertList(strUserName, strKey, strSigCert, strCertSn) Then Exit Function
    '֤��Ψһ��ʶ
    intPoint = InStr(strKey, "F")
    If mUserInfo.strUserID = "" Then
        MsgBoxEx "�������֤��Ϊ��,����ϵ����Ա����Ա������¼�룡", vbOKOnly + vbInformation, gstrSysName
        Exit Function
    ElseIf UCase(Mid(strKey, intPoint + 2)) <> mUserInfo.strUserID Then
        MsgBoxEx "�������֤�ţ�" & _
                   vbCrLf & vbTab & "��" & mUserInfo.strUserID & "��" & vbCrLf & _
                   "��ǰ֤��Ψһ��ʶ:" & _
                   vbCrLf & vbTab & "��" & UCase(Mid(strKey, intPoint + 2)) & "��" & vbCrLf & _
                   "�û����֤���뵱ǰ֤��Ψһ��ʶ�����,����ʹ�ã�", vbInformation, gstrSysName
        Exit Function
    End If
    If Not GetCertLogin(strSigCert, strCertSn, intDate) Then
        blnRet = False
    Else
        blnRet = True
    End If
    If blnRet Then
            '�ж��Ƿ���Ҫ����ע��֤��
            udtUser.strName = strUserName
            udtUser.strSignName = strUserName
            udtUser.strUserID = UCase(Mid(strKey, intPoint + 2)) 'SF+���֤��
            udtUser.strCertSn = strCertSn
            udtUser.strCertDN = GetCertDN(strSigCert)
            udtUser.strCert = strSigCert
            udtUser.strEncCert = ""
            udtUser.strCertID = strKey
            Call GetEndDate(strSigCert, strDate)
            If IsUpdateRegCert(udtUser, strDate, blnReDo) Then
                blnRet = True
            Else
                blnRet = False
            End If
        End If
     
        SCCA_CheckCert = blnRet
        Exit Function
errH:
124     MsgBoxEx "���USBKEYʧ�ܣ�" & Err.Description, vbInformation, gstrSysName
End Function

Private Function GetCertLogin(ByVal strCert As String, ByVal strCertSn As String, ByVal dDate As Date) As Boolean
    '- ���
    'strUniqueID : ֤��Ψһ��ʶ
    'strPassword : ֤������
    'strWebserviceUrl:ǩ����������ַ����Ϊ֤����֤
    '- ����
    'dDate       : ����֤����Чʱ��
    On Error GoTo errH
    Dim result As Boolean
    Dim strDate As String '֤����Ч���ַ���
    Dim strArry() As String
    Dim lngʱ�� As Long
    Dim strLogin As String
    If SCCA_Client Is Nothing Then Set SCCA_Client = CreateObject("wsb.WsbClient.1")
    If SCCA_Server Is Nothing Then Set SCCA_Client = CreateObject("wsbServer.wsbServerClass.1")
    '֤�鰲ȫ��¼
    'strLogin:��Ϊ��:�ɹ�
    'strLogin:Ϊ��:���ɹ�
    strLogin = SCCA_Client.SOF_SignDataByP7(strCertSn, 1)
    '��֤֤������Ϣ��ʾ
    If strLogin <> "" Then
        '��֤֤����Ч��
        result = SCCA_Client.SOF_ValidateCert(strCert)
        If result Then
            '��ȡ�ͻ���֤����Ч�ڽ�ֹʱ��
            Call GetEndDate(strCert, strDate)
            dDate = CDate(strDate)
            lngʱ�� = CheckValidaty(dDate)
            If lngʱ�� < 0 Then
                MsgBoxEx "����֤���ѹ���!"
                GetCertLogin = False
            ElseIf (lngʱ�� <= 30 And lngʱ�� > 0) And Not gblnShow Then
                MsgBoxEx "����֤�黹��" & lngʱ�� & "�����"
                gblnShow = True
                GetCertLogin = True
            Else
                GetCertLogin = True
            End If
        Else
            MsgBoxEx "��֤֤��ʧ�ܣ�" & "SCCA_Client.GetCertInfo", vbInformation, gstrSysName
        End If
    Else
        MsgBoxEx "��ʼ��½����" & "SCCA_Client.SOF_Login", vbInformation, gstrSysName
    End If
    Exit Function
errH:
    MsgBoxEx "����֤��ӿڴ���!" & Err.Description, vbInformation, gstrSysName
    GetCertLogin = False
End Function

Private Function GetEndDate(ByVal strCert As String, ByRef strDate As String)
    Dim strArry() As String
    strDate = SCCA_Client.SOF_GetCertInfo(strCert, 18)
    If strDate <> "" Then
        strArry = Split(strDate, " ")
        If InStr(strArry(0), "Jan") > 0 Then
            strDate = strArry(3) & "-01-" & strArry(1) & " " & strArry(2)
        ElseIf InStr(strArry(0), "Feb") > 0 Then
            strDate = strArry(3) & "-02-" & strArry(1) & " " & strArry(2)
        ElseIf InStr(strArry(0), "Mar") > 0 Then
            strDate = strArry(3) & "-03-" & strArry(1) & " " & strArry(2)
        ElseIf InStr(strArry(0), "Apr") > 0 Then
            strDate = strArry(3) & "-04-" & strArry(1) & " " & strArry(2)
        ElseIf InStr(strArry(0), "May") > 0 Then
            strDate = strArry(3) & "-05-" & strArry(1) & " " & strArry(2)
        ElseIf InStr(strArry(0), "Jun") > 0 Then
            strDate = strArry(3) & "-06-" & strArry(1) & " " & strArry(2)
        ElseIf InStr(strArry(0), "Jul") > 0 Then
            strDate = strArry(3) & "-07-" & strArry(1) & " " & strArry(2)
        ElseIf InStr(strArry(0), "Aug") > 0 Then
            strDate = strArry(3) & "-08-" & strArry(1) & " " & strArry(2)
        ElseIf InStr(strArry(0), "Sep") > 0 Then
            strDate = strArry(3) & "-09-" & strArry(1) & " " & strArry(2)
        ElseIf InStr(strArry(0), "Oct") > 0 Then
            strDate = strArry(3) & "-10-" & strArry(1) & " " & strArry(2)
        ElseIf InStr(strArry(0), "Nov") > 0 Then
            strDate = strArry(3) & "-11-" & strArry(1) & " " & strArry(2)
        ElseIf InStr(strArry(0), "Dec") > 0 Then
            strDate = strArry(3) & "-12-" & strArry(1) & " " & strArry(2)
        End If
    Else
        Exit Function
    End If
End Function

Public Function SCCA_Sign(ByVal strCurrCertSn As String, ByVal strSource As String, ByRef strSignData As String, ByRef strTimeStamp As String, ByRef blnReDo As Boolean) As Boolean
        'ǩ��
        Dim strSigCert As String, strCertSn As String
        Dim blnCheck As Boolean
        Dim datTime As Date
        Dim strDate As String
        Dim udtUser As USER_INFO

        On Error GoTo errH
        blnCheck = SCCA_CheckCert("", strSigCert, strCertSn, blnReDo)
        If blnReDo Then Exit Function
        If blnCheck Then
            datTime = gobjComLib.zlDatabase.Currentdate()
            strDate = Format(datTime, "yyyyMMddhhmmss")
            strTimeStamp = Format(datTime, "yyyy-MM-dd HH:mm:ss")
            strSource = StringSHA1(strSource)
            strSignData = SCCA_Client.SOF_SignDataByP7(strCertSn, strSource)
            If strSignData <> "" Then
                 SCCA_Sign = True
            Else
                MsgBoxEx "ǩ��ʧ�ܣ�"
            End If
        Else
            MsgBoxEx "ǩ��ʧ�ܣ�", vbInformation, "����ǩ������"
        End If
        Exit Function
errH:
114     MsgBoxEx "ǩ��ʧ�ܣ�" & Err.Description, vbInformation, gstrSysName
End Function

Public Function SCCA_VerifySign(ByVal strSignData As String, ByVal strSource As String) As Boolean
        '��֤ǩ��
        Dim strTmp As String
        Dim strǩ��ԭ�� As String
        On Error GoTo errH
        strǩ��ԭ�� = SCCA_Server.SOF_GetP7SignDataInfo(strSignData, 1)
        strSource = StringSHA1(strSource)
        If strǩ��ԭ�� = strSource Then
            strTmp = SCCA_Server.SOF_VerifySignedDataByP7(strSignData)
            If Val(strTmp) = 0 Then
                 MsgBoxEx "��֤ǩ���ɹ���", vbInformation, gstrSysName
            Else
                 MsgBoxEx "��֤ǩ��ʧ�ܣ�", vbInformation, gstrSysName
            End If
        Else
            MsgBoxEx "ǩ��ԭ����ǩ��ֵ�е�ԭ�Ĳ�һ�£�����ԭ���Ƿ��޸Ĺ���", vbInformation, gstrSysName
        End If
        Exit Function
errH:
104     MsgBoxEx "��֤ǩ��ʧ�ܣ�" & Err.Description, vbInformation, gstrSysName
End Function

Public Sub SCCA_UnloadObj()
    Set SCCA_Client = Nothing
    Set SCCA_Server = Nothing
    mblnInit = False
End Sub

