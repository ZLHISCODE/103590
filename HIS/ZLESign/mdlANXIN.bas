Attribute VB_Name = "mdlANXIN"

Option Explicit
'���ֹ�Ͷ����CA         ��̨������ҽԺ
'KEY����:����epass3000gm auto  ����ANXIN3KGM ����GM3000   ���� ANXINLONGMAIGM
Private mobjJLClient As Object   'JITComVCTKEx.JITVCTKEx.1
Private mobjJLServer As Object  'JITClientCOMAPI.JITClientProc.1
Private mobjCertInfo As Object    'ǩ�� JITCertActiveX.CertInfo.1

'Private mobjJLClient As New JITComVCTKExLib.JITVCTKEx
'Private mobjJLServer As New JITClientCOMAPILib.JITClientProc
'Private mobjCertInfo As New JITCertActiveXLib.CertInfo    'ǩ��
Private mblnInit As Boolean

Private mstrPWD As String          '�������������
Private mstrKey As String

Private Const M_STR_PARA As String = "<?xml version=""1.0"" encoding=""gb2312""?><authinfo><liblist>" & _
        "<lib type=""CSP"" version=""1.0"" dllname="""" ><algid val=""SHA1"" sm2_hashalg=""sm3""/></lib>" & _
        "<lib type=""SKF"" version=""1.1"" dllname=""SERfR01DQUlTLmRsbA=="" ><algid val=""SHA1"" sm2_hashalg=""sm3""/></lib>" & _
        "<lib type=""SKF"" version=""1.1"" dllname=""U2h1dHRsZUNzcDExXzMwMDBHTS5kbGw="" ><algid val=""SHA1"" sm2_hashalg=""sm3""/></lib>" & _
        "<lib type=""SKF"" version=""1.1"" dllname=""QU5YSU5Dc3AxMV8zMDAwR01BLmRsbA=="" ><algid val=""SHA1"" sm2_hashalg=""sm3""/></lib>" & _
        "</liblist><checkkeytimes><item times=""3"" ></item></checkkeytimes></authinfo>"

Private Enum E_KEY_TYPE
    K_���� = 0
    K_���� = 1
End Enum

Public Function ANXIN_InitObj() As Boolean
     '֤�鲿����ʼ��
    Dim lngRet As Long
    Dim strTSAIP As String
    Dim strPara As String
    Dim varTmp As Variant
    Dim i As Long
    
100     If glngSign > 1 Then ANXIN_InitObj = True: Exit Function
        On Error Resume Next
102     If mobjJLClient Is Nothing Then
104         Set mobjJLClient = CreateObject("JITComVCTKEx.JITVCTKEx.1")
106         If Err.Number <> 0 Then
108             MsgBoxEx "��������ǩ������JITComVCTK_S.dll��ʧ�ܣ�����ÿؼ��Ƿ�װ��ע�ᡣ", vbInformation, gstrSysName
                Exit Function
            End If
        End If
110     Err.Clear
112     If mobjJLServer Is Nothing Then
114         Set mobjJLServer = CreateObject("JITClientCOMAPI.JITClientProc.1")
116         If Err.Number <> 0 Then
118             MsgBoxEx "��������ǩ������JITClientCOMAPI.dll��ʧ�ܣ�����ÿؼ��Ƿ�װ��ע�ᡣ", vbInformation, gstrSysName
                Exit Function
            End If
        End If
120     Err.Clear
122     If mobjCertInfo Is Nothing Then
124         Set mobjCertInfo = CreateObject("JITCertActiveX.SZZSCertInfo.1")
126         If Err.Number <> 0 Then
128             MsgBoxEx "��������ǩ�¶���JITCertActiveX.dll��ʧ�ܣ�����ÿؼ��Ƿ�װ��ע�ᡣ", vbInformation, gstrSysName
                Exit Function
            End If
        End If
130     Err.Clear: On Error GoTo 0
        On Error GoTo errH
        '������Ϣ:�Ƿ�����ʱ���[0-������;1-����]&&&ǩ��������IP&&&ǩ���������˿�&&&ʱ���IP&&&ʱ����˿�&&&��ѡ����(dllname1&dllname2)
        '��һλ ǩ�����������ڶ�λʱ���������������λ���ء����ž���3��Ӳ�������û��Ӳ�����ǡ�000����ֻ��ǩ�����������ǡ�100��
        'Ϊ�˼�����ǰ�����һ������=0;�����ֻ��ǩ�����������ǡ�100�� ;=1��������ǩ������������ʱ��������� ����"110"����λ������ʱδ���ã�Ԥ������
        '����ʱ���������
        'gstrPara = "1&&&175.17.252.155&&&8000&&&175.17.252.156&&&8000" '��̨������ҽԺ
        'gstrPara = "000&&&175.17.252.155&&&8000&&&175.17.252.156&&&8000" '��̨�Ӹ���ҽԺ
        
132     gstrPara = gobjComLib.zlDatabase.GetPara(90000, glngSys, , "") '��ȡ��������
        LogWrite "ANXIN_InitObj", "CA����:" & gstrPara
134     If gstrPara = "" Then
136         Err.Raise -1, , "��ǰϵͳ��" & glngSys & "��û�����õ���ǩ������,�뵽���õ���ǩ���ӿڴ����á�"
            Exit Function
        End If
138     If UBound(Split(gstrPara, G_STR_SPLIT)) <> 6 Then
140         MsgBoxEx "����ǩ������ֵ��������,���顣" & vbCrLf & _
                "��ǰ����ֵ:" & gstrPara & vbCrLf & _
                "��ȷ��ʽ:�Ƿ�����ʱ���[0-������;1-����]&&&ǩ��������IP&&&ǩ���������˿�&&&ʱ���IP&&&ʱ����˿�&&&KEY����(0-����;1-����)&&&��ѡ����(dllname1&dllname2)", vbInformation, gstrSysName
            Exit Function
        Else
142         Call ANXIN_GetPara
        End If
        
144     If gudtPara.intKeyType = K_���� Then
146         mstrKey = "ANXIN3KGM"
148     ElseIf gudtPara.intKeyType = K_���� Then
150         mstrKey = "ANXINLONGMAIGM"
        End If
152     If gudtPara.strOption <> "" Then
154         varTmp = Split(gudtPara.strOption, "&")
156         strPara = "<?xml version=""1.0"" encoding=""gb2312""?><authinfo><liblist>" & _
                        "<lib type=""CSP"" version=""1.0"" dllname="""" ><algid val=""SHA1"" sm2_hashalg=""sm3""/></lib>"
158         For i = LBound(varTmp) To UBound(varTmp)
160             strPara = strPara & "<lib type=""SKF"" version=""1.1"" dllname=""" & varTmp(i) & """ ><algid val=""SHA1"" sm2_hashalg=""sm3""/></lib>"
            Next
162         strPara = strPara & "</liblist><checkkeytimes><item times=""3"" ></item></checkkeytimes></authinfo>"
        Else
164         strPara = M_STR_PARA
        End If
166     lngRet = mobjJLClient.Initialize(strPara)
168     If Not GetErrorInfo("Initialize") Then Exit Function
170     mblnInit = True
174     mstrPWD = ""
176     ANXIN_InitObj = True
        
        Exit Function
errH:
178  MsgBoxEx "�����ӿڲ���ʧ�ܣ�" & vbNewLine & "��" & CStr(Erl()) & "�� " & Err.Description, vbExclamation, gstrSysName
End Function

Public Function ANXIN_Sign(ByVal strSource As String, ByRef strSignData As String, ByRef strTimeStamp As String, ByRef strTimeStampCode As String, Optional ByRef blnReDo As Boolean) As Boolean
'ǩ��
    Dim lngRet As Long
    Dim strErr As String
    Dim strHash As String
    
    Dim blnCheck As Boolean
        On Error GoTo errH
        blnCheck = ANXIN_CheckCert(blnReDo)
        If blnReDo Then Exit Function
100     If blnCheck Then                 '��֤��ǰUSB�Ƿ���ǩ���û��ģ�����ȡǩ��֤��
            '֤��ID����ǩ��
            lngRet = mobjJLClient.SetPinCode(mstrPWD)
110         strSignData = mobjJLClient.DetachSignStr("", strSource)      'DetachSignStr-����ԭ��ǩ��;AttachSignStr-��ԭ��ǩ��
            If Not GetErrorInfo("DetachSignStr") Then Exit Function
            If strSignData <> "" Then
                If gudtPara.blnISTS Then
                    If Not ConnectToTsaServer() Then Exit Function
                    strHash = StringSHA1(strSource)
                    strTimeStampCode = mobjJLServer.TsaSign("", 1, strHash)            '����ʱ��� ����ǩ��ֵ������ǩ��ʱ�ȽϺ�ʱ,�ʲ��ù̶�ֵ
                    strTimeStamp = mobjJLServer.VerifyTsaSign(strTimeStampCode)
                    Call mobjJLServer.FinalizeServerConnectEx    '�Ͽ�ʱ�������������
                    If strTimeStampCode = "" Then MsgBoxEx "��ȡʱ���ʧ�ܣ�", vbInformation, gstrSysName: Exit Function
                    '���ڸ�ʽ��
                    strTimeStamp = Mid(strTimeStamp, 1, 14)
                    strTimeStamp = String14ToDate(strTimeStamp, strErr)
                    If strErr <> "" Then MsgBoxEx strErr, vbInformation, gstrSysName: Exit Function
                    'ת������ʱ��
                    strTimeStamp = Format(DateAdd("h", 8, strTimeStamp), "YYYY-MM-DD HH:MM:SS")
                Else
                    strTimeStamp = Format(gobjComLib.zlDatabase.Currentdate & "", "yyyy-MM-dd HH:mm:ss")
                End If
            Else
                MsgBoxEx "ǩ��ʧ�ܣ�", vbInformation, gstrSysName
                Exit Function
            End If
112
        Else
            MsgBoxEx "ǩ��ʧ�ܣ�", vbInformation, gstrSysName
            Exit Function
        End If
        ANXIN_Sign = True
        Exit Function
errH:
114     MsgBoxEx "ǩ��ʧ�ܣ�" & vbNewLine & "��" & CStr(Erl()) & "�� " & Err.Description, vbExclamation, gstrSysName
End Function

Public Function ANXIN_VerifySign(ByVal strSource As String, ByVal strSignData As String, ByVal strTimeStampCode As String) As Boolean
    '����;��֤ǩ��
    '����:strSignData -ǩ��ֵ
     Dim blnRet As Boolean
     Dim lngRet As Long
     Dim strTS As String
     On Error GoTo errH
     LogWrite "ANXIN_VerifySign", "��֤ǩ��ԭ��:" & strSource & vbCrLf & "��֤ǩ��ֵ:" & strSignData & vbCrLf & "ǩ��ʱ�����Ϣ:" & strTimeStampCode
100 If gudtPara.blnIsSign Then
        '��������ǩ
102     If Not ConnectToSignServer() Then Exit Function
104     lngRet = mobjJLServer.VerifyDetachedSign(strSignData, strSource) '��������֤���� ����ԭ��ǩ��:VerifyDetachedSign(string, string);��ԭ��ǩ��  VerifyAttachedSign
106     If lngRet <> 0 Then
108         MsgBoxEx "ǩ����֤ʧ��:" & mobjJLServer.GetErrorMessage(lngRet), vbInformation, gstrSysName
110         Call mobjJLServer.FinalizeServerConnectEx   '�Ͽ�ǩ������������
            Exit Function
        End If
112     Call mobjJLServer.FinalizeServerConnectEx   '�Ͽ�ǩ������������
    Else
114     Call mobjJLClient.VerifyDetachedSignStr(strSignData, strSource)  '�ͻ�����֤ǩ��
116     lngRet = mobjJLClient.GetErrorCode()
118     If lngRet <> 0 Then
120         MsgBoxEx "��֤ǩ��ʧ�ܣ������룺" & lngRet & " ������Ϣ��" & mobjJLClient.GetErrorMessage(lngRet), vbInformation, gstrSysName
            Exit Function
        End If
    End If

    '����ʱ���������
122 If gudtPara.blnISTS Then
124     If Not ConnectToTsaServer() Then Exit Function
126     strTS = mobjJLServer.VerifyTsaSign(strTimeStampCode)
128     If strTS = "" Then
130           MsgBoxEx "ʱ�����֤ʧ�ܣ�", vbInformation, gstrSysName
132           Call mobjJLServer.FinalizeServerConnectEx   '�Ͽ�ʱ���������
              Exit Function
        End If
134     Call mobjJLServer.FinalizeServerConnectEx   '�Ͽ�ʱ���������
    End If
136 MsgBoxEx "��֤�ɹ����õ���ǩ��������Ч!", vbInformation, gstrSysName
    
138  ANXIN_VerifySign = True
     Exit Function
errH:
140     MsgBoxEx "��֤ǩ��ʧ�ܣ�" & vbNewLine & "��" & CStr(Erl()) & "�� " & Err.Description, vbExclamation, gstrSysName

End Function


Public Function ANXIN_CheckCert(ByRef blnReDo As Boolean) As Boolean
    '���ܣ���ȡUSB�����豸��ʼ������¼
    Dim strKeySN As String, strUserID As String, strUserName As String, strCertDN As String
    Dim strDate As String
    Dim arrDN As Variant
    Dim udtUser As USER_INFO
    Dim blnRet As Boolean
    Dim i As Long
    
    On Error GoTo errH
    If Not GetCertList(strKeySN, strUserName, strCertDN, , strUserID) Then Exit Function
    If mUserInfo.strUserID = "" Then
        MsgBoxEx "�������֤��Ϊ��,����ϵ����Ա����Ա������¼�룡", vbOKOnly + vbInformation, gstrSysName
        Exit Function
     ElseIf mUserInfo.strUserID <> strUserID Then
        MsgBoxEx "��֤��δע�����������£�����ʹ�ã�"
        Exit Function
    End If
    
    '�ж��Ƿ���Ҫ����ע��֤��
    udtUser.strName = strUserName
    udtUser.strSignName = strUserName
    udtUser.strUserID = strUserID '���֤��
    udtUser.strCertSn = strKeySN
    udtUser.strCertDN = strCertDN
    udtUser.strCert = ""
    udtUser.strEncCert = ""
    udtUser.strCertID = ""
    udtUser.strPicPath = ""
    arrDN = Split(mUserInfo.strCertDN, ",")     'CN=����СU3294, O=��̨������ҽԺ, L=��̨����, S=������ʡ, C=CN, ��Ч����=
    For i = 0 To UBound(arrDN)
        If Trim(arrDN(i)) Like "��Ч����*" Then
            strDate = Trim(Split(arrDN(i), "=")(1))
            Exit For
        End If
    Next
    If IsUpdateRegCert(udtUser, strDate, blnReDo) Then
        blnRet = True
    Else
        blnRet = False
    End If

    ANXIN_CheckCert = blnRet
    Exit Function
errH:
     MsgBoxEx "���USBKEYʧ�ܣ�" & vbNewLine & "��" & CStr(Erl()) & "�� " & Err.Description, vbExclamation, gstrSysName
End Function


Public Function ANXIN_RegCert(arrCertInfo As Variant) As Boolean
    '���ܣ��ṩ��HIS���ݿ���ע�����֤��ı�Ҫ��Ϣ,�����·��Ż����֤��,,��Ҫ����USB-Key
    '���أ�arrCertInfo��Ϊ���鷵��֤�������Ϣ
'      0-ClientSignCertCN:�ͻ���ǩ��֤�鹫������(����),����ע��֤��ʱ������֤���
'      1-ClientSignCertDN:�ͻ���ǩ��֤������(ÿ��Ψһ)
'      2-ClientSignCertSN:�ͻ���ǩ��֤�����к�(ÿ֤��Ψһ)
'      3-ClientSignCert:�ͻ���ǩ��֤������
'      4-ClientEncCert:�ͻ��˼���֤������
'      5-ǩ��ͼƬ�ļ���,�մ���ʾû��ǩ��ͼƬ

        Dim strKeyId As String, strCertUserName As String, strCertDN As String, strPicPath As String
        Dim i As Integer
        On Error GoTo errH

        For i = LBound(arrCertInfo) To UBound(arrCertInfo)
             arrCertInfo(i) = ""
        Next

        If GetCertList(strKeyId, strCertUserName, strCertDN, strPicPath) Then
            arrCertInfo(0) = strCertUserName
            arrCertInfo(1) = strCertDN
            arrCertInfo(2) = strKeyId
            arrCertInfo(5) = strPicPath
            ANXIN_RegCert = True
        End If

        Exit Function
errH:
     MsgBoxEx "��ȡ֤����Ϣʧ�ܣ�" & vbNewLine & "��" & CStr(Erl()) & "�� " & Err.Description, vbExclamation, gstrSysName

End Function

Private Function GetCertList(Optional ByRef strUniqueID As String = "-1", Optional ByRef strName As String = "-1", Optional ByRef strCertDN As String = "-1", _
    Optional ByRef strPicPath As String = "-1", Optional ByRef strUserID As String = "-1", Optional ByRef strDate As String = "-1") As Boolean
    '����:��ȡ����֤������
    'strUserID-���֤��
        Dim lngRet As Long
        Dim strRet As String
        Dim datCurrent As Date
        Dim arrDN As Variant
        Dim i As Long
        Dim strKeyCount As String
        Dim strPic As String, strPIN As String
        Dim strTmp As String
        Dim lngDay As Long
        Dim colTmp As Collection
        
        On Error GoTo errH
        
100     If Not mblnInit Then
102         lngRet = mobjJLClient.Initialize(M_STR_PARA)
104         If Not GetErrorInfo("Initialize") Then Exit Function
106         mblnInit = True
        End If
108     lngRet = mobjJLClient.SetCertChooseType(1)
110     lngRet = mobjJLClient.SetCert("SC", "", "", "", "CN = AnXin SM2 CA,O = AnXin CA,C = CN", "")
112     If Not GetErrorInfo("SetCert") Then Exit Function
114     strDate = mobjJLClient.GetCertInfo("SC", 6, "")     '��Ч����
116     If IsDate(strDate) Then
            '���֤���Ƿ����
118         lngDay = CheckValidaty(strDate)
120         If (lngDay <= 30 And lngDay > 0 And Not gblnShow) Then
122             MsgBoxEx "����֤�黹��" & lngDay & "�����", vbInformation, gstrSysName
124             gblnShow = True
126         ElseIf (lngDay <= 0) Then
128             MsgBoxEx "����֤���ѹ��� " & Abs(lngDay) & " ��"
                Exit Function
            End If
        End If
130     If strUniqueID <> "-1" Then strUniqueID = mobjJLClient.GetCertInfo("SC", 2, "")     '֤�����к�
132     If strCertDN <> "-1" Or strName <> "-1" Then
134         strCertDN = mobjJLClient.GetCertInfo("SC", 0, "") 'CN=����СU3294, O=��̨������ҽԺ, L=��̨����, S=������ʡ, C=CN, ��Ч����=
136         If strCertDN <> "" Then
138             arrDN = Split(strCertDN, ",")
140             For i = 0 To UBound(arrDN)
142                 If Trim(arrDN(i)) Like "CN*" Then
144                     strName = Trim(Split(arrDN(i), "=")(1))
                        Exit For
                    End If
                Next
            End If
146         strCertDN = strCertDN & ", ��Ч����=" & strDate
        End If
    
148     If strUserID <> "-1" Then
150         strUserID = ""
152         strTmp = mobjJLClient.GetCertInfo("SC", 7, "1.2.86.11.7.1")  '���֤����ҪתASCII :31 16 a0 14 13 12 34 33 32 35 30 33 31 39 38 36 30 31 31 32 36 32 31 35
154         If Not GetErrorInfo("GetCertInfo") Then Exit Function
156         If strTmp <> "" Then
158             arrDN = Split(strTmp, " ")
160             For i = 6 To UBound(arrDN)    'ǰ6���ַ�Ϊǰ׺
162                 strUserID = strUserID & Chr(Val("&H" & arrDN(i)))
                Next
            End If
        End If
    
164     If mstrPWD = "" Then
CheckPWD:
166         If Not frmPassword.ShowMe(mstrPWD, 6, 16) Then Exit Function
            'strRet = mobjCertInfo.VerifyUserPin(mstrKey, mstrPWD)
            'VB���Ե�ʱ�򵥲����ٷ������룻ֱ�����з�����ȷ�ַ���'{"RetryCount":"0","VerifyValue":"1"}
            '�״���֤����,ͨ��ǩ���ӿ�������
168         lngRet = mobjJLClient.SetPinCode(mstrPWD)
170         strRet = mobjJLClient.DetachSignStr("", "123")
172         If Not GetErrorInfo("DetachSignStr") Then
174             mstrPWD = ""
                Exit Function
            End If
        End If

        '���Key����
188     If strPicPath <> "-1" Then
            '��ȡǩ�º�ʱ,ǩ��ʱ����ȡ��ֻ��ע���ʱ���ȡ
            'strKeyCount = [{"KeyName":"���ŵ���Կ�� ","KeyType":"ANXIN3KGM","UsbKeySerialNumber":"AX00010415"},{"KeyName":"���ŵ���Կ��","KeyType":"ANXIN3KGM","UsbKeySerialNumber":"AX00010414"}]
190         strKeyCount = mobjCertInfo.GetKeyCount(mstrKey)
192         strRet = strKeyCount 'VB���Ե�ʱ�򵥲����ٷ������룻ֱ�����з�����ȷ�ַ���
194         If strRet <> "" Then
196             If UBound(Split(strRet, "},{")) = 0 Then
198                 strPic = mobjCertInfo.ReadImageData(mstrKey, mstrPWD)
200                 If Len(strPic) > 1 Then
202                     strPicPath = SaveBase64ToFile("gif", strUniqueID, strPic)
                    Else
204                     strPicPath = ""
                    End If
206             ElseIf Val(strKeyCount) > 0 Then
208                 MsgBoxEx "��ѡ��Ψһ��KEY�̲��룡", vbInformation, gstrSysName
                    Exit Function
                End If
            End If
        End If

210     GetCertList = True
        Exit Function
errH:
212     MsgBoxEx "��ȡ֤����Ϣʧ�ܣ�" & vbNewLine & "��" & CStr(Erl()) & "�� " & Err.Description, vbExclamation, gstrSysName
End Function

Public Function ANXIN_GetSeal() As String
'��ȡǩ��ͼƬ
    Dim strPicPath As String
    Call GetCertList(, , , strPicPath)
    ANXIN_GetSeal = strPicPath
End Function

Private Function GetErrorInfo(ByVal strName As String) As Boolean
        Dim lngRet As Long

        On Error GoTo errH
100     lngRet = mobjJLClient.GetErrorCode  'lngRet -536870826 ���벻��;-536870823  ָ��������̫����̫��
102     If lngRet <> 0 Then
104         MsgBoxEx "���ýӿڣ���" & strName & "�������,��������:" & vbCrLf & mobjJLClient.GetErrorMessage(lngRet), vbInformation, gstrSysName
            Exit Function
        End If
106     GetErrorInfo = True
        Exit Function
errH:
108     MsgBoxEx "��ȡ����������" & vbNewLine & "��" & CStr(Erl()) & "�� " & Err.Description, vbExclamation, gstrSysName
End Function

Private Function ConnectToTsaServer() As Boolean
        Dim lngRet As Long

        On Error GoTo errH

100     lngRet = mobjJLServer.InitServerConnectEx(gudtPara.strTSIP, CInt(gudtPara.strTSPort))
102     If lngRet <> 0 Then
104         MsgBoxEx mobjJLServer.GetErrorMessage(lngRet), vbInformation, gstrSysName
            Exit Function
        End If
106     lngRet = mobjJLServer.SetServerUriEx("/signserver/service/xml")
108     If lngRet <> 0 Then
110         MsgBoxEx mobjJLServer.GetErrorMessage(lngRet), vbInformation, gstrSysName
112         Call mobjJLServer.FinalizeServerConnectEx '��ֹ���������ӣ����ͷ����Ӿ��
            Exit Function
        End If
114     ConnectToTsaServer = True
        Exit Function
errH:
116     MsgBoxEx "����ʱ�����������" & vbNewLine & "��" & CStr(Erl()) & "�� " & Err.Description, vbExclamation, gstrSysName

End Function

Private Function ConnectToSignServer() As Boolean
        Dim lngRet As Long

        On Error GoTo errH
100     lngRet = mobjJLServer.InitServerConnectEx(gudtPara.strSIGNIP, CInt(gudtPara.strSignPort))
102     If lngRet <> 0 Then
104         MsgBoxEx mobjJLServer.GetErrorMessage(lngRet), vbInformation, gstrSysName  '���ӷ�����ʧ��
            Exit Function
        End If
106     lngRet = mobjJLServer.SetServerUriEx("/signserver/service/xml")
108     If lngRet <> 0 Then
110         MsgBoxEx mobjJLServer.GetErrorMessage(lngRet), vbInformation, gstrSysName
112         Call mobjJLServer.FinalizeServerConnectEx '��ֹ���������ӣ����ͷ����Ӿ��
            Exit Function
        End If
114     lngRet = mobjJLServer.SetCertAliasEx("")  '���÷�����ǩ��ʱ��ǩ��֤���ʶ,��ΪĬ��֤��
116     If lngRet <> 0 Then
118         MsgBoxEx mobjJLServer.GetErrorMessage(lngRet), vbInformation, gstrSysName
120         Call mobjJLServer.FinalizeServerConnectEx '��ֹ���������ӣ����ͷ����Ӿ��
            Exit Function
        End If
122     ConnectToSignServer = True
        Exit Function
errH:
124     MsgBoxEx "����ǩ������������" & vbNewLine & "��" & CStr(Erl()) & "�� " & Err.Description, vbExclamation, gstrSysName
End Function


Public Function ANXIN_GetPara() As Boolean
        Dim arrList As Variant
    
        On Error GoTo errH
100     If gstrPara = "" Then gstrPara = gobjComLib.zlDatabase.GetPara(90000, glngSys, , "") '��ȡURLs �̶���ȡZLHIS ϵͳĬ��100
        '��ʽ�Ƿ������豸[000-��������;111-������]&&&ǩ��������IP&&&ǩ���������˿�&&&ʱ���IP&&&ʱ����˿�&&&KEY����(0-����;1-����)&&&��������
102     If gstrPara = "" Then gstrPara = "110&&&175.17.252.155&&&8000&&&175.17.252.156&&&8000&&&0&&&" & _
            "SERfR01DQUlTLmRsbA==&U2h1dHRsZUNzcDExXzMwMDBHTS5kbGw=&U2h1dHRsZUNzcDExXzMwMDBHTS5kbGw="
104     arrList = Split(gstrPara, "&&&")
106     If UBound(arrList) >= 6 Then
108         If Len(arrList(0)) = 3 Then
110             gudtPara.blnIsSign = Mid(arrList(0), 1, 1) = "1"
112             gudtPara.blnISTS = Mid(arrList(0), 2, 1) = "1"
            Else
114             gudtPara.blnISTS = Val(arrList(0)) = 1
116             gudtPara.blnIsSign = True
            End If
118         gudtPara.strSIGNIP = arrList(1)
120         gudtPara.strSignPort = arrList(2)
122         gudtPara.strTSIP = arrList(3)
124         gudtPara.strTSPort = arrList(4)
        
126         gudtPara.intKeyType = arrList(5)
128         gudtPara.strOption = arrList(6)
        Else
130         gudtPara.blnISTS = True
132         gudtPara.blnIsSign = True
134         gudtPara.strSIGNIP = "175.17.252.155"
136         gudtPara.strSignPort = "8000"
138         gudtPara.strTSIP = "175.17.252.156"
140         gudtPara.strTSPort = "8000"
142         gudtPara.intKeyType = K_����    'Ĭ�Ϸ���
144         gudtPara.strOption = "SERfR01DQUlTLmRsbA==&U2h1dHRsZUNzcDExXzMwMDBHTS5kbGw=&U2h1dHRsZUNzcDExXzMwMDBHTS5kbGw="
        End If
        Exit Function
errH:
146     MsgBoxEx "��ȡ����ʧ�ܣ�" & vbNewLine & "��" & CStr(Erl()) & "�� " & Err.Description, vbInformation, gstrSysName
End Function

Public Function ANXIN_SetParaStr() As String
    With gudtPara
        ANXIN_SetParaStr = IIf(.blnIsSign, "1", "0") & IIf(.blnISTS, "1", "0") & "0" & G_STR_SPLIT & _
            IIf(Trim(.strSIGNIP) = "", "175.17.252.155", .strSIGNIP) & G_STR_SPLIT & IIf(Trim(.strSignPort) = "", "8000", .strSignPort) & _
            G_STR_SPLIT & IIf(Trim(.strTSIP) = "", "175.17.252.156", .strTSIP) & G_STR_SPLIT & IIf(Trim(.strTSPort) = "", "8000", .strTSPort) & _
            G_STR_SPLIT & .intKeyType & _
            G_STR_SPLIT & IIf(Trim(.strOption) = "", "SERfR01DQUlTLmRsbA==&U2h1dHRsZUNzcDExXzMwMDBHTS5kbGw=&U2h1dHRsZUNzcDExXzMwMDBHTS5kbGw=", .strOption)

    End With
End Function

Public Sub ANXIN_UnLoadObj()
    On Error Resume Next
    Set mobjJLServer = Nothing
    Set mobjCertInfo = Nothing
    Call mobjJLClient.Finalize
    Set mobjJLClient = Nothing
End Sub




