Attribute VB_Name = "mdlJSCA"
Option Explicit
'����CA���Ĺ���ģ�飨����ҽԺ��
Private mblnInit As Boolean         '�Ƿ��ѳ�ʼ���ɹ�
Private JSCA_Client As Object       '����ǩ������
Private JSCA_Seal As Object         '����ǩ�²���
Private mobjSoapClient As Object        'soap���Ӷ���
Private mstrKeys  As String             '�����key��ʶ
Private mstrSelectKey As String         '���key�������,��¼ѡ���key
'
Private Const GET_CERT_AFTER As Integer = 18   '֤����Ч�ڣ���ֹ����
Private Const GET_USER_NAME  As Integer = 23            '֤�����������

Public Function JSCA_InitObj() As Boolean
    '֤�鲿����ʼ��
    Dim strUrl As String
    
        On Error GoTo errH

102     JSCA_InitObj = mblnInit
104     If mblnInit Then Exit Function

        On Error Resume Next
106     Set mobjSoapClient = CreateObject("MSSOAP.SoapClient30")  'SOAP���Ӷ���
        mobjSoapClient.ClientProperty("ServerHTTPRequest") = True
        
        If Err.Number <> 0 Then
            MsgBoxEx "ϵͳ��ʼ��ʧ�ܣ�" & vbCrLf & vbCrLf & "�ͻ���δ��װSOAP��" & vbCrLf & vbCrLf & "������Ϣ���£�" & vbCrLf & vbCrLf & Err.Description, vbCritical, "����ǩ������"
            Exit Function
        End If
        Err.Clear: On Error GoTo 0
        strUrl = gobjComLib.zlDatabase.GetPara(90000, glngSys)  '��ȡURL
'        strURL = "http://202.102.85.153:8080/HealthWebService.asmx?WSDL"
        If strUrl = "" Then
            Err.Raise -1, , "û������ǩ����������ַ���������á�"
            Exit Function
        End If
        On Error Resume Next
        mobjSoapClient.MSSoapInit (strUrl)
        If Err.Number <> 0 Then
            MsgBoxEx "ϵͳ��ʼ��ʧ�ܣ�" & vbCrLf & vbCrLf & "��������ַ����" & vbCrLf & vbCrLf & "������Ϣ���£�" & vbCrLf & vbCrLf & Err.Description, vbCritical, "����ǩ������"
            Exit Function
        End If
        Err.Clear: On Error GoTo 0
108     Set JSCA_Client = CreateObject("CACltCore.CltCore")
        Set JSCA_Seal = CreateObject("GSEAL.GSealCtrl.1")
        JSCA_Client.IsShowError = 0     '=0ʱ,����SOF_ShowErrMsg()���Ե���������Ϣ
        
114     JSCA_InitObj = True
    
116     mblnInit = JSCA_InitObj
        Exit Function
errH:
118     MsgBoxEx "��������CA�ӿڲ���ʧ�ܣ�" & vbNewLine & Err.Description, vbQuestion, "����ǩ������"
    
End Function

Public Function JSCA_RegCert(arrCertInfo As Variant) As Boolean
        '���ܣ��ṩ��HIS���ݿ���ע�����֤��ı�Ҫ��Ϣ,�����·��Ż����֤��,,��Ҫ����USB-Key
        '���أ�arrCertInfo��Ϊ���鷵��֤�������Ϣ
        '      0-ClientSignCertCN:�ͻ���ǩ��֤�鹫������(����),����ע��֤��ʱ������֤���
        '      1-ClientSignCertDN:�ͻ���ǩ��֤������(ÿ��Ψһ)
        '      2-ClientSignCertSN:�ͻ���ǩ��֤�����к�(ÿ֤��Ψһ)
        '      3-ClientSignCert:�ͻ���ǩ��֤������
        '      4-ClientEncCert:�ͻ��˼���֤������
        '      5-ǩ��ͼƬ�ļ���,�մ���ʾû��ǩ��ͼƬ
        
        Dim strKeyId As String, strCertTime As String, strCertUserName As String, strCertDN As String
        Dim strSigCert As String
        Dim strFile As String
        Dim i As Long
        On Error GoTo errH
    
100     For i = LBound(arrCertInfo) To UBound(arrCertInfo)
101         arrCertInfo(i) = ""
102     Next
        
108     If JSCA_GetCertList(strCertUserName, strKeyId, strSigCert) Then
200         arrCertInfo(0) = strCertUserName
201         arrCertInfo(1) = JSCA_Client.SOF_GetUserInfo(strKeyId, 4) 'C=CN, S=����ʡ, L=�Ͼ���, O=����ʡ��������֤����֤�����������ι�˾, OU=JSCA, CN=JSCA_CA
202         arrCertInfo(2) = strKeyId
203         arrCertInfo(3) = strSigCert
205         arrCertInfo(4) = ""
            strFile = App.Path & "\" & Format(Now, "yyyyMMdd") & "_" & strKeyId & ".gif"
            Call JSCA_Seal.JSCAGetSealPath(strFile)     '��key�̶�ȡͼƬ
            
206         arrCertInfo(5) = strFile
            JSCA_RegCert = True
        End If
        
300     Exit Function

errH:
    MsgBoxEx "��ȡ֤����Ϣʧ�ܣ�" & vbNewLine & "��" & CStr(Erl()) & "�� " & Err.Description, vbQuestion, "����ǩ������"
End Function

Public Function JSCA_GetCertList(ByRef strName As String, Optional ByRef strUniqueID As String, Optional ByRef strCert As String, Optional ByVal strUserUnigueID As String) As Boolean
    '����CA ����ҽԺ
    '-���:��
    '     strUserUnigueID-��ǰ�û��󶨵�Key����key�������
    '-����
    'strName :      ����ӿڷ��ص�֤������������
    'strUniqueID:   ����ӿڷ��ص�֤��������Ψһ��ʶ
    'strCert:       ����ӿڷ��ص�ǩ��֤��
    Dim strTmp As String
    Dim strkeys As String
    Dim arrKey As Variant
    Dim intCount As Integer
    Dim blnChange As Boolean
    Dim i As Integer
    On Error GoTo errH
    
    If JSCA_Client Is Nothing Then Set JSCA_Client = CreateObject("CACltCore.CltCore")
    If Not JSCA_Client Is Nothing Then
        'ֻ�������һ��key,������key��ε�����Աѡ����
        strTmp = JSCA_Client.SOF_GetUserList() '�������ݸ�ʽ��(�û�1||��ʶ1&&&�û�2||��ʶ2&&&��)һ��key����������ͬ����
        If strTmp <> "" Then
            arrKey = Split(strTmp, "&&&")
            intCount = (UBound(arrKey) + 1) \ 2
            If intCount > 1 Then
                strkeys = ""
                For i = LBound(arrKey) To UBound(arrKey)
                    If InStr(1, strkeys & ",", "," & arrKey(i) & ",") = 0 Then
                        strkeys = strkeys & "," & arrKey(i)    '��¼�µ�ǰ����������key��ʶ
                    End If
                    If InStr(1, mstrKeys, "," & arrKey(i) & ",") = 0 And mstrKeys <> "" Then
                        mstrSelectKey = ""
                        blnChange = True  'key�䶯Ҫ����ѡ��
                    End If
                Next
                
                If strkeys <> "" And (mstrKeys = "" Or blnChange) Then
                    mstrKeys = strkeys & ","
                End If
            Else
                mstrKeys = ""
                mstrSelectKey = ""
            End If
        End If
        
        If intCount > 1 And mstrSelectKey <> "" And mstrSelectKey = strUserUnigueID Then
            strUniqueID = mstrSelectKey '��key�������,����Ա������key��δ�䶯
        Else
            strUniqueID = JSCA_Client.SOF_SelectCert(3)   '֤��ID ��key������»� ��������ѡ����
        End If
        
        If intCount > 1 Then
            mstrSelectKey = strUniqueID    '��¼���״�ѡ��
        End If
        
        If strUniqueID = "" Then
            MsgBoxEx "��������Key��", vbInformation + vbOKOnly, "����ǩ������"
            Exit Function
        End If
        strCert = JSCA_Client.SOF_ExportUserCert(strUniqueID)   '֤������
        strName = JSCA_Client.SOF_GetCertInfo(strCert, GET_USER_NAME)   '��ȡ֤�����������
    Else
        MsgBoxEx "����CA������ʼ��ʧ�ܡ�", vbInformation, "����ǩ������"
        Exit Function
    End If
    JSCA_GetCertList = True
    Exit Function
errH:
    MsgBoxEx "��ȡ֤����Ϣʧ�ܣ�" & vbNewLine & "��" & CStr(Erl()) & "�� " & Err.Description, vbQuestion, "����ǩ������"
    
End Function


Public Function JSCA_Sign(ByVal strCurrCertSn As String, ByVal strSource As String, ByRef strSignData As String, ByRef strTimeStamp As String, ByRef strTimeStampCode As String) As Boolean
        'ǩ��
'����:����CA����ǩ��
'������strCurrCertSn  -֤��ID(Ψһ����)
'     strSource-��Ҫǩ����Դ����
'
'����ֵ��true �ɹ�,False -ʧ��
'       strSignData-ǩ���󷵻ص�ǩ������
'       strTimeStamp-���ص�ʱ���

        Dim strRequest As String    'ʱ�������
        
        On Error GoTo errH
        
100     If JSCA_CheckCert(strCurrCertSn) Then               '��֤��ǰUSB�Ƿ���ǩ���û��ģ�����ȡǩ��֤��
            strSource = JSCA_Client.SOF_EncodeBase64(strSource)   'base64����,��ֹ�ַ������пմ����з���֤ʧЧ
110         strSignData = JSCA_Client.SOF_SignData(strCurrCertSn, strSource)
            If strSignData <> "" Then
                JSCA_Sign = True
'                strRequest = mobjSoapClient.CreateTimeStampRequest(strSource)  '����ʱ�������
'                If strRequest <> "" Then
'112                 strTimeStampCode = mobjSoapClient.CreateTimeStampResponse(strRequest)  '��ȡʱ�����Ӧ��base64�����ʽ��
'                    strTimeStamp = mobjSoapClient.GetTimeStampInfo(strTimeStampCode, 1)    '����ʱ���type =1������ʱ�䣻type ��2������ǩ��ֵ��type ��3������ǩ��֤��
'                    strTimeStamp = GetTimeStamp(strTimeStamp)
'                    JSCA_Sign = True
'                Else
'                    MsgboxEx "ʱ�������ʧ�ܣ�", vbExclamation, "����ǩ������"
'                End If
            Else
                MsgBoxEx "ǩ��ʧ�ܣ�", vbExclamation, "����ǩ������"
            End If
        Else
            MsgBoxEx "ǩ��ʧ�ܣ�", vbExclamation, "����ǩ������"
        End If
        Exit Function
errH:
114     MsgBoxEx "ǩ��ʧ�ܣ�" & vbNewLine & "��" & CStr(Erl()) & "�� " & Err.Description, vbQuestion, "����ǩ������"
End Function

Public Function JSCA_CheckCert(ByVal strCurrCertSn As String) As Boolean
'���ܣ���ȡUSB�����豸��ʼ������¼
'����ֵ:
'  strSigCert -ǩ��֤������

        Dim strKey As String, strUserName As String, strSigCert As String
        Dim strWebUrl As String
        
        On Error GoTo errH
100     If Not mblnInit Then
102         MsgBoxEx "����δ��ʼ����"
            Exit Function
        End If
        
104     If JSCA_GetCertList(strUserName, strKey, strSigCert, strCurrCertSn) Then
106        If strCurrCertSn <> strKey Then
108            MsgBoxEx "��֤��δע�����������£�����ʹ�ã�", vbInformation + vbOKOnly, "����ǩ������"
               Exit Function
           End If
110
116        If Not GetCertLogin(strKey, strSigCert) Then
                Exit Function
           Else
               JSCA_CheckCert = True
           End If

122     End If
    
        Exit Function
errH:
124     MsgBoxEx "���USBKEYʧ�ܣ�" & vbNewLine & "��" & CStr(Erl()) & "�� " & Err.Description, vbQuestion, "����ǩ������"
End Function

Private Function GetCertLogin(ByVal strUniqueID As String, ByVal strCert As String) As Boolean
    '����CA
    '- ���
    'strUniqueID : ֤��Ψһ��ʶ
    'strCert-֤������ 64BASE����

    Dim datEnd  As Date
    Dim intDay  As Integer
    Dim lngTimes As Long
    
    '��֤���� �״ε��ý���CAǩ���ӿ����Զ���������¼�봰�壬ֻ¼��һ�Ρ�

'        lngTimes = JSCA_Client.SOF_Login(strUniqueID, strPassword)   'У��ɹ����� 0,ʧ�ܷ���ʣ�����,-1 ����
'        If lngTimes < -1 Then
'            MsgboxEx "������֤ʧ�ܣ�������ע��CA������CACltCore.dll����", vbInformation + vbOKOnly, "����ǩ������"
'            Exit Function
'        End If

    '��ȡ�ͻ���֤����Ч�ڽ�ֹʱ��
    datEnd = JSCA_Client.SOF_GetCertInfo(strCert, GET_CERT_AFTER)
    '��֤�ͻ���֤����Ч��ʣ������
     intDay = Int(CDbl(datEnd) - CDbl(Now))
    
    If (intDay <= 30 And intDay > 0 And Not gblnShow) Then
        MsgBoxEx "����֤�黹��" & intDay & "����ڡ�", vbInformation + vbOKOnly, "����ǩ������"
        gblnShow = True
        GetCertLogin = True
    ElseIf (intDay <= 0) Then
        MsgBoxEx "����֤���ѹ��� " & Abs(intDay) & " �졣", vbInformation + vbOKOnly, "����ǩ������"
        GetCertLogin = False
    End If
    GetCertLogin = True
End Function

Public Function JSCA_VerifySign(ByVal strSigCert As String, ByVal strSignData As String, ByVal strSource As String, ByVal strTimeStamp As String, ByVal strTimeStampCode As String) As Boolean
'����;��֤ǩ��
'����: strSigCert -֤������
'     strSignData-ǩ��ֵ
'     strSource-����֤Դ��
'     strTimeStampCode -ʱ���BASE64����
        Dim strTmp As String
        
        On Error GoTo errH
100
        strSource = JSCA_Client.SOF_EncodeBase64(strSource) 'base64����,��ֹ�ַ������пմ����з���֤ʧЧ
'        strTmp = mobjSoapClient.VerifyTimeStamp(strSource, strTimeStampCode)
'        If strTmp <> "0" Then
'            MsgboxEx "ʱ�����֤ʧ�ܣ�", vbExclamation, "����ǩ������"
'            Exit Function
'        End If
        strTmp = mobjSoapClient.VerifySignedData(strSigCert, strSource, strSignData)
        If strTmp = "0" Then
            MsgBoxEx "��֤�ɹ����õ���ǩ��������Ч!", vbInformation, "����ǩ������"
        Else
            MsgBoxEx "��֤ǩ��ʧ�ܣ�", vbExclamation, "����ǩ������"
            Exit Function
        End If
       
        JSCA_VerifySign = True
        Exit Function
errH:
104     MsgBoxEx "��֤ǩ��ʧ�ܣ�" & vbNewLine & "��" & CStr(Erl()) & "�� " & Err.Description, vbQuestion, "����ǩ������"
End Function


Private Function GetTimeStamp(ByVal strTimeStamp As String) As String
'���ܣ���ȡʱ����е�ʱ��
    Dim arrTime As Variant
    Dim i As Long
    Dim strTime As String
    
    strTimeStamp = Replace(strTimeStamp, " ", "|")
    strTimeStamp = Replace(strTimeStamp, "||", "|") '������Ϊһλ��ʱ,��ֹ�·ݺ�����֮����������ո�����
    arrTime = Split(strTimeStamp, "|")  '���˸�ʽ��Aug 19 13:07:25 2014 GMT
    strTime = arrTime(0) & " " & arrTime(1) & " " & arrTime(3)  '��/��/��
    strTime = CDate(strTime) & ""  '�� �� ��  2014/8/19
    GetTimeStamp = strTime & " " & arrTime(2)  ' ��-��-�� ʱ:��:��
End Function

Public Function JSCA_GetPara() As Boolean
'���ú���CA��������ַ
    Dim arrList As Variant
    
    On Error GoTo errH
    gstrPara = gobjComLib.zlDatabase.GetPara(90000, glngSys)
    If gstrPara = "" Then gstrPara = "http://202.102.85.153:8080/HealthWebService.asmx?WSDL"
    If gstrPara <> "" Then
        gudtPara.strSignURL = gstrPara
    End If
    Exit Function
errH:
    MsgBoxEx "��ȡ����ʧ�ܣ�" & vbNewLine & "��" & CStr(Erl()) & "�� " & Err.Description, vbInformation, gstrSysName
End Function

Public Function JSCA_SetParaStr() As String
    JSCA_SetParaStr = gudtPara.strSignURL
End Function


Public Sub JSCA_UnloadObj()
    Set JSCA_Client = Nothing
    Set JSCA_Seal = Nothing      '����ǩ�²���
    Set mobjSoapClient = Nothing        'soap���Ӷ���
    mblnInit = False
End Sub
