Attribute VB_Name = "mdlNMGCA"

Option Explicit
'��ͷ������ҽԺ
Private mobjJLClient As Object   'JITComVCTKEx.JITVCTKEx.1
Private mobjJLServer As Object  'JITClientCOMAPI.JITClientProc.1
Private mobjCertInfo As Object    'ǩ�� BicengEsealInterface.CEsealInterface
Private mblnInit As Boolean
'���ַ����������й�
Private Const M_STR_PARA_NM As String = "<?xml version=""1.0"" encoding=""utf-8""?>" & _
        "<authinfo><liblist>" & _
        "<lib type=""SKF"" version=""1.1"" dllname=""bXRva2VuX2dtMzAwMC5kbGw="" ><algid val=""SHA1"" sm2_hashalg=""sm3""/></lib>" & _
        "<lib type=""SKF"" version=""1.1"" dllname=""U21hcnRDVENBUEkuZGxs"" ><algid val=""SHA1"" sm2_hashalg=""sm3""/></lib>" & _
        "<lib type=""SKF"" version=""1.1"" dllname=""U0tGQVBJMjA1NDkuZGxs"" ><algid val=""SHA1"" sm2_hashalg=""sm3""/></lib>" & _
        "</liblist></authinfo>"

Public Function NMG_InitObj() As Boolean
     '֤�鲿����ʼ��
    Dim lngRet As Long
    Dim strTSAIP As String
    Dim varTmp As Variant
    
100     If glngSign > 1 Then NMG_InitObj = True: Exit Function
        On Error Resume Next
102     If mobjJLClient Is Nothing Then
104         Set mobjJLClient = CreateObject("JITComVCTKEx.JITVCTKEx.1")
106         If Err.Number <> 0 Then
108             MsgBoxEx "����ǩ������JITComVCTK_S.dll��ʧ�ܣ�����ÿؼ��Ƿ�װ��ע�ᡣ", vbInformation, gstrSysName
                Exit Function
            End If
        End If
110     Err.Clear
112     If mobjJLServer Is Nothing Then
114         Set mobjJLServer = CreateObject("JITClientCOMAPI.JITClientProc.1")
116         If Err.Number <> 0 Then
118             MsgBoxEx "����ǩ������JITClientCOMAPI.dll��ʧ�ܣ�����ÿؼ��Ƿ�װ��ע�ᡣ", vbInformation, gstrSysName
                Exit Function
            End If
        End If
    
        On Error GoTo errH
    
120     gstrPara = gobjComLib.zlDatabase.GetPara(90000, glngSys, , "") '��ȡ��������
        
122     If gstrPara = "" Then
124         Err.Raise -1, , "��ǰϵͳ��" & glngSys & "��û�����õ���ǩ������,�뵽���õ���ǩ���ӿڴ����á�"
            Exit Function
        End If
126     If UBound(Split(gstrPara, G_STR_SPLIT)) <> 4 Then
128         MsgBoxEx "����ǩ������ֵ��������,���顣" & vbCrLf & _
                "��ǰ����ֵ:" & gstrPara & vbCrLf & _
                "��ȷ��ʽ:ǩ��������IP&&&ǩ���������˿�&&&ʱ���IP&&&ʱ����˿�", vbInformation, gstrSysName
            Exit Function
        Else
130         Call NMG_GetPara
        End If
        
        On Error Resume Next
132     If gudtPara.blnSignPic Then
134         If mobjCertInfo Is Nothing Then
136             Set mobjCertInfo = CreateObject("BicengEsealInterface.CEsealInterface")
138             If Err.Number <> 0 Then
140                 MsgBoxEx "����ǩ�¶���BicengEsealInterface.dll��ʧ�ܣ�����ÿؼ��Ƿ�װ��ע�ᡣ", vbInformation, gstrSysName
                    Exit Function
                End If
            End If
        End If
142     Err.Clear: On Error GoTo 0
        On Error GoTo errH
144     lngRet = mobjJLClient.Initialize(M_STR_PARA_NM)

146     If Not GetErrorInfo("Initialize") Then Exit Function
148     mblnInit = True
150     NMG_InitObj = True
        Exit Function
errH:
152  MsgBoxEx "�����ӿڲ���ʧ�ܣ�" & vbNewLine & "��" & CStr(Erl()) & "�� " & Err.Description, vbExclamation, gstrSysName
End Function

Public Function NMG_Sign(ByVal strSource As String, ByRef strSignData As String, ByRef strTimeStamp As String, _
    ByRef strTimeStampCode As String, Optional ByRef blnReDo As Boolean, Optional ByVal blnCheck As Boolean) As Boolean
        'ǩ��
    Dim lngRet As Long
    Dim strErr As String
    Dim strHash As String
    

        On Error GoTo errH
        If Not blnCheck Then
100         blnCheck = NMG_CheckCert(blnReDo)
102         If blnReDo Then Exit Function
        End If
104     If blnCheck Then                 '��֤��ǰUSB�Ƿ���ǩ���û��ģ�����ȡǩ��֤��
            '֤��ID����ǩ��
106         strSignData = mobjJLClient.DetachSignStr("", strSource)      'DetachSignStr-����ԭ��ǩ��;AttachSignStr-��ԭ��ǩ��
108         If Not GetErrorInfo("DetachSignStr") Then Exit Function
110         If strSignData <> "" Then
112             If Not ConnectToTsaServer() Then Exit Function
                Call mobjJLServer.SetAlgorithmEx("SM3", "")
114             strHash = StringSHA1(strSource)
116             strTimeStampCode = mobjJLServer.TsaSign("", 1, strHash)            '����ʱ��� ����ǩ��ֵ������ǩ��ʱ�ȽϺ�ʱ
118             strTimeStamp = mobjJLServer.VerifyTsaSign(strTimeStampCode)
                lngRet = mobjJLServer.GetErrorCodeEx()
                WriteLog "ʱ���������:" & lngRet
120             Call mobjJLServer.FinalizeServerConnectEx    '�Ͽ�ʱ�������������
122             If strTimeStampCode = "" Then MsgBoxEx "��ȡʱ���ʧ�ܣ�", vbInformation, gstrSysName: Exit Function
                '���ڸ�ʽ��
124             strTimeStamp = Mid(strTimeStamp, 1, 14)
126             strTimeStamp = String14ToDate(strTimeStamp, strErr)
128             If strErr <> "" Then MsgBoxEx strErr, vbInformation, gstrSysName: Exit Function
                'ת������ʱ��
130             strTimeStamp = Format(DateAdd("h", 8, strTimeStamp), "YYYY-MM-DD HH:MM:SS")
            Else
132             MsgBoxEx "ǩ��ʧ�ܣ�", vbInformation, gstrSysName
                Exit Function
            End If

        Else
134         MsgBoxEx "ǩ��ʧ�ܣ�", vbInformation, gstrSysName
            Exit Function
        End If
136     NMG_Sign = True
        Exit Function
errH:
138     MsgBoxEx "ǩ��ʧ�ܣ�" & vbNewLine & "��" & CStr(Erl()) & "�� " & Err.Description, vbExclamation, gstrSysName
End Function

Public Function NMG_VerifySign(ByVal strSource As String, ByVal strSignData As String, ByVal strTimeStampCode As String) As Boolean
    '����;��֤ǩ��
    '����:strSignData -ǩ��ֵ
     Dim blnRet As Boolean
     Dim lngRet As Long
     Dim strTS As String
     On Error GoTo errH

    '��������ǩ
100 If Not ConnectToSignServer() Then Exit Function
102 lngRet = mobjJLServer.VerifyDetachedSign(strSignData, strSource)  '��������֤���� ����ԭ��ǩ��:VerifyDetachedSign(string, string);��ԭ��ǩ��  VerifyAttachedSign
104 If lngRet <> 0 Then
106     MsgBoxEx "ǩ����֤ʧ��:" & mobjJLServer.GetErrorMessage(lngRet), vbInformation, gstrSysName
108     Call mobjJLServer.FinalizeServerConnectEx   '�Ͽ�ǩ������������
        Exit Function
    End If
110 Call mobjJLServer.FinalizeServerConnectEx   '�Ͽ�ǩ������������


    '����ʱ���������
112 If Not ConnectToTsaServer() Then Exit Function
114 strTS = mobjJLServer.VerifyTsaSign(strTimeStampCode)
116 If strTS = "" Then
118       MsgBoxEx "ʱ�����֤ʧ�ܣ�", vbInformation, gstrSysName
120       Call mobjJLServer.FinalizeServerConnectEx   '�Ͽ�ʱ���������
          Exit Function
    End If
122 Call mobjJLServer.FinalizeServerConnectEx   '�Ͽ�ʱ���������
 
124 MsgBoxEx "��֤�ɹ����õ���ǩ��������Ч!", vbInformation, gstrSysName
    
126  NMG_VerifySign = True
     Exit Function
errH:
128     MsgBoxEx "��֤ǩ��ʧ�ܣ�" & vbNewLine & "��" & CStr(Erl()) & "�� " & Err.Description, vbExclamation, gstrSysName
End Function


Public Function NMG_CheckCert(ByRef blnReDo As Boolean) As Boolean
        '���ܣ���ȡUSB�����豸��ʼ������¼
        Dim strKeySN As String, strUserID As String, strUserName As String, strCertDN As String
        Dim strDate As String
        Dim arrDN As Variant
        Dim udtUser As USER_INFO
        Dim blnRet As Boolean
        Dim i As Long
    
        On Error GoTo errH
100     If Not GetCertList(strKeySN, strUserName, strCertDN) Then Exit Function
102     If mUserInfo.strCertSn <> strKeySN Then
104         MsgBoxEx "��֤��δע�����������£�����ʹ�ã�" & vbCrLf & _
                "�û�ע��֤��Ψһ��ʶ:" & mUserInfo.strCertSn & vbCrLf & _
                "��ǰ��ѡ֤��Ψһ��ʶ:" & strKeySN, vbInformation + vbOKOnly, gstrSysName
            Exit Function
        End If
    
        '�ж��Ƿ���Ҫ����ע��֤��
106     udtUser.strName = strUserName
108     udtUser.strSignName = strUserName
110     udtUser.strUserID = strUserID '���֤��
112     udtUser.strCertSn = strKeySN
114     udtUser.strCertDN = strCertDN
116     udtUser.strCert = ""
118     udtUser.strEncCert = ""
120     udtUser.strCertID = ""
122     udtUser.strPicPath = ""
124     arrDN = Split(mUserInfo.strCertDN, ",")     'CN=����СU3294, O=��̨������ҽԺ, L=��̨����, S=������ʡ, C=CN, ��Ч����=
126     For i = 0 To UBound(arrDN)
128         If Trim(arrDN(i)) Like "��Ч����*" Then
130             strDate = Trim(Split(arrDN(i), "=")(1))
                Exit For
            End If
        Next
132     If IsUpdateRegCert(udtUser, strDate, blnReDo) Then
134         blnRet = True
        Else
136         blnRet = False
        End If

138     NMG_CheckCert = blnRet
        Exit Function
errH:
140      MsgBoxEx "���USBKEYʧ�ܣ�" & vbNewLine & "��" & CStr(Erl()) & "�� " & Err.Description, vbExclamation, gstrSysName
End Function


Public Function NMG_RegCert(arrCertInfo As Variant) As Boolean
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

100         For i = LBound(arrCertInfo) To UBound(arrCertInfo)
102              arrCertInfo(i) = ""
            Next

104         If GetCertList(strKeyId, strCertUserName, strCertDN, strPicPath) Then
106             arrCertInfo(0) = strCertUserName
108             arrCertInfo(1) = strCertDN
110             arrCertInfo(2) = strKeyId
112             arrCertInfo(5) = strPicPath
114             NMG_RegCert = True
            End If

            Exit Function
errH:
116      MsgBoxEx "��ȡ֤����Ϣʧ�ܣ�" & vbNewLine & "��" & CStr(Erl()) & "�� " & Err.Description, vbExclamation, gstrSysName
End Function

Private Function GetCertList(Optional ByRef strUniqueID As String = "-1", Optional ByRef strName As String = "-1", Optional ByRef strCertDN As String = "-1", _
    Optional ByRef strPicPath As String = "-1", Optional ByRef strUserID As String = "-1", Optional ByRef strDate As String = "-1") As Boolean
    '����:��ȡ֤������
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
        On Error GoTo errH
        
100     lngRet = mobjJLClient.SetCertChooseType(1)
102     lngRet = mobjJLClient.SetCert("SC", "", "", "", "", "")
104     If Not GetErrorInfo("SetCert") Then Exit Function
106     strDate = mobjJLClient.GetCertInfo("SC", 6, "")     '��Ч����
108     If IsDate(strDate) Then
            '���֤���Ƿ����
110         lngDay = CheckValidaty(strDate)
112         If (lngDay <= 30 And lngDay > 0 And Not gblnShow) Then
114             MsgBoxEx "����֤�黹��" & lngDay & "�����", vbInformation, gstrSysName
116             gblnShow = True
118         ElseIf (lngDay <= 0) Then
120             MsgBoxEx "����֤���ѹ��� " & Abs(lngDay) & " ��"
                Exit Function
            End If
        End If
    
122     If strUniqueID <> "-1" Then strUniqueID = mobjJLClient.GetCertInfo("SC", 2, "")     '֤�����к�
124     If strCertDN <> "-1" Or strName <> "-1" Then
126         strCertDN = mobjJLClient.GetCertInfo("SC", 0, "") 'CN=����СU3294, O=��̨������ҽԺ, L=��̨����, S=������ʡ, C=CN, ��Ч����=
128         strName = mobjJLClient.GetCertInfo("SC", 9, "")     '�û�����
        End If
130     If gudtPara.blnSignPic Then
132         strPic = mobjCertInfo.SignSeal("����", "����123")
134         If Trim(strPic) = "" Then MsgBoxEx "��ȡǩ��ʧ��!", vbInformation, gstrSysName: Exit Function
136         strPicPath = App.Path & "\" & Format(Now, "yyyyMMdd") & "_" & strUniqueID & ".bmp"
138         mobjCertInfo.getSealPicturePath strPic, strPicPath
140         If Dir(strPicPath) = "" Then MsgBoxEx "��ȡǩ��ʧ��!", vbInformation, gstrSysName: Exit Function
        End If
142     GetCertList = True
        Exit Function
errH:
144     MsgBoxEx "��ȡ֤����Ϣʧ�ܣ�" & vbNewLine & "��" & CStr(Erl()) & "�� " & Err.Description, vbExclamation, gstrSysName
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

Public Function NMG_GetPara() As Boolean
        Dim arrList As Variant
    
        On Error GoTo errH
100     If gstrPara = "" Then gstrPara = gobjComLib.zlDatabase.GetPara(90000, glngSys, , "") '��ȡURLs �̶���ȡZLHIS ϵͳĬ��100
102     If gstrPara = "" Then gstrPara = "110&&&175.17.252.155&&&8000&&&175.17.252.156&&&8000&&&0"   'ǩ��������IP&&&ǩ���������˿�&&&ʱ���IP&&&ʱ����˿�&&&����ǩ��
104     arrList = Split(gstrPara, "&&&")
106     If UBound(arrList) >= 4 Then
108         gudtPara.strSIGNIP = arrList(0)
110         gudtPara.strSignPort = arrList(1)
112         gudtPara.strTSIP = arrList(2)
114         gudtPara.strTSPort = arrList(3)
116         gudtPara.blnSignPic = Val(arrList(4) & "") = 1
        Else
118         gudtPara.strSIGNIP = "175.17.252.155"
120         gudtPara.strSignPort = "8000"
122         gudtPara.strTSIP = "175.17.252.156"
124         gudtPara.strTSPort = "8000"
126         gudtPara.blnSignPic = False
        End If
        Exit Function
errH:
128     MsgBoxEx "��ȡ����ʧ�ܣ�" & vbNewLine & "��" & CStr(Erl()) & "�� " & Err.Description, vbInformation, gstrSysName
End Function

Public Function NMG_SetParaStr() As String
    With gudtPara
        NMG_SetParaStr = IIf(Trim(.strSIGNIP) = "", "175.17.252.155", .strSIGNIP) & G_STR_SPLIT & IIf(Trim(.strSignPort) = "", "8000", .strSignPort) & _
                G_STR_SPLIT & IIf(Trim(.strTSIP) = "", "175.17.252.156", .strTSIP) & G_STR_SPLIT & IIf(Trim(.strTSPort) = "", "8000", .strTSPort) & G_STR_SPLIT & IIf(.blnSignPic, 1, 0)
    End With
End Function

Public Sub NMG_UnLoadObj()
    On Error Resume Next
    Set mobjJLServer = Nothing
    Set mobjCertInfo = Nothing
    Call mobjJLClient.Finalize
    Set mobjJLClient = Nothing
End Sub





