Attribute VB_Name = "mdlBJCA"
Option Explicit

Private mobjTSA As Object       '����׼���ҽԺ��ʱ����ӿ�
Private mobjAXCA As Object      '����CA����
Private mLastPWD As String      '��������
Private mobjAXSVR As Object     '���ŷ���˲���

Public Function ZGRYY_initObj() As Boolean
    Err.Clear: On Error Resume Next
    Set mobjAXCA = Nothing
    Set mobjAXCA = CreateObject("AXSECURITY.AXSecurityCtrl.1")
    
    If Err.Number <> 0 Then
        MsgBoxEx "����ǩ���ؼ�û�а�װ��", vbExclamation, gstrSysName
        Exit Function
    End If
    
    Set mobjAXSVR = CreateObject("AnXEgovCom.AnXEgovSOF")
    If Err.Number <> 0 Then
        MsgBoxEx "���ŷ�������û�а�װ��", vbExclamation, gstrSysName
        Exit Function
    End If
    
    Err.Clear: On Error Resume Next
    Set mobjTSA = Nothing
    Set mobjTSA = CreateObject("tsaMiddleware.UtilUdp")
    If Err.Number <> 0 Then
        MsgBoxEx "ʱ����ؼ�û�а�װ��", vbExclamation, gstrSysName
        Exit Function
    End If
    ZGRYY_initObj = True
End Function
Public Function ZGRYY_UnloadObj()
    'ጷŌ���
    If Not mobjTSA Is Nothing Then Set mobjTSA = Nothing
    If Not mobjAXCA Is Nothing Then Set mobjAXCA = Nothing
End Function

Public Function ZGRYY_RegCert(arrCertInfo As Variant) As Boolean
    
        
        Dim bSuccess As Boolean
        Dim strDn As String, strSN As String, strUser As String, i As Integer
        Dim strBase64 As String '���漪��ʡҽԺ�ӿڷ��ص�ͼƬ����(�����34527)
        Dim strImage  As String '���漪��ʡҽԺ�ӿڷ��ص�ͼƬ�ļ�·��
        On Error GoTo errH
100     If mobjAXCA Is Nothing Then Set mobjAXCA = CreateObject("AXSECURITY.AXSecurityCtrl.1")

102     If mobjAXCA Is Nothing Then
104         MsgBoxEx "����ǩ���ؼ�δ��ȷ��װ��", vbExclamation, gstrSysName
            Exit Function
        End If
106     bSuccess = mobjAXCA.OpenCert(0, "", 1)
108     If (bSuccess) Then
110         strDn = mobjAXCA.GetCertInfo(1, "")
112         i = InStr(1, strDn, "O=")
114         strUser = Mid(strDn, 4, i - 6)
116         strSN = mobjAXCA.GetCertInfo(2, "")
        '  ���ǵõ�ͼƬǩ�µ�
118         strBase64 = mobjAXCA.ReadFileFromKey(0, 2)
120         If strBase64 <> "" Then
122             strImage = App.Path & "\" & strSN & ".gif"
124             If Dir(strImage) <> "" Then Kill strImage
126             If Not mobjAXCA.B64DecodeSToFile(strBase64, strImage) Then strImage = ""
            Else
128             strImage = ""
            End If
        Else
130         MsgBoxEx mobjAXCA.GetLastError
            Exit Function
        End If

132     arrCertInfo(0) = strUser
134     arrCertInfo(1) = strDn
136     arrCertInfo(2) = strSN
        
138     arrCertInfo(5) = strImage
140     ZGRYY_RegCert = True
    Exit Function
errH:
142 MsgBoxEx "ע��֤��-��" & CStr(Erl()) & "��," & Err.Description, vbQuestion, "����ǩ��"
End Function

Public Function ZGRYY_CheckCert(ByVal strCurrCertSn As String) As Boolean
        '��֤��ǰ��USB
        Dim strPIN As String, blnClientChk As Boolean
        
        
        On Error GoTo hErr
100     If mLastPWD <> "" Then strPIN = mLastPWD

102     Call mobjAXCA.SetKeyPWD(2, strPIN)  '������û����ģգӣ��ͺ��޸ĵڣ�������,ȡֵΪ�� 0:���� 1����̩ 2:����3000 3:����
104     mLastPWD = strPIN
106     blnClientChk = mobjAXCA.SetSignerCert(2, strCurrCertSn)
108     If Not blnClientChk Then
110         mLastPWD = ""
112         MsgBoxEx "��ǰ����֤����֤ʧ�ܡ�" & mobjAXCA.GetLastError, vbExclamation, gstrSysName
            Exit Function
        End If
        
114     ZGRYY_CheckCert = True
        Exit Function
hErr:
116     MsgBoxEx "���֤��-��" & CStr(Erl()) & "��," & Err.Description, vbExclamation, gstrSysName
End Function
Public Function ZGRYY_Sign(ByVal strCurrCertSn As String, ByVal strSource As String, ByRef strSignData As String, ByRef strTimeStamp As String) As Boolean
        'ǩ��
        
    Dim strClientSignData As String '�ͻ���ǩ���������
    Dim strGetTimeDate As String           'ʱ����ӿڷ�������
    Dim intSvrVerfy    As Integer           '���ŷ����P7��֤
    On Error GoTo errH
100 If mobjAXCA Is Nothing Then Set mobjAXCA = CreateObject("AXSECURITY.AXSecurityCtrl.1")
    
102 If Not ZGRYY_CheckCert(strCurrCertSn) Then Exit Function
    
104 If Not AxServer_VerifyCert Then Exit Function   '�����֤����֤ 2012-12-28
    
106 If Not Times_Tamp(strSource, strTimeStamp) Then Exit Function

108 strClientSignData = mobjAXCA.SignString(strSource, True)
    ' MsgboxEx strTimeStamp '���ԣ�ǩ��ʱûȡ��ʱ�䣬��֤ʱ��ȡ��ʱ�����
110 If Len(strTimeStamp) = 0 Then
            '�ٵõ�ʱ���
112     strGetTimeDate = verify_getTimestamp(strSource)
114     If strGetTimeDate <> "��" Then
116         strTimeStamp = Format(CDate(strGetTimeDate), "yyyy-MM-dd HH:mm:ss")
        End If
    End If
118 If strClientSignData = "" Then
120     MsgBoxEx mobjAXCA.GetLastError, vbExclamation, gstrSysName
        Exit Function
    End If
    
    '�����ǩ����֤����ԭ�Ļ��ǿͻ���ǩ����������ĵ���û����ȷ���˴���ʱ��ǩ��������  2012-12-28
    
122 intSvrVerfy = mobjAXSVR.SOF_VerifySignedDataByP7(strClientSignData)
124 If intSvrVerfy <> 0 Then
126     MsgBoxEx "���ŷ����ǩ����֤ʧ�ܣ�������" & intSvrVerfy
        Exit Function
    End If
    
128 strSignData = strClientSignData
130 ZGRYY_Sign = False
    
132 ZGRYY_Sign = True
    Exit Function
errH:
134 MsgBoxEx "ǩ��-��" & CStr(Erl()) & "��," & Err.Description, vbExclamation, gstrSysName
End Function

Public Function ZGRYY_VerifySign(ByVal strCurrCertSn As String, ByVal strSignData As String, ByVal strSource As String, ByVal strTime As String) As Boolean
    '��֤ǩ��
    Dim aaa As String
    Dim csdate As Date
    If mobjAXCA Is Nothing Then Set mobjAXCA = CreateObject("AXSECURITY.AXSecurityCtrl.1")
    
    If verify_Timestamp(strSource) = False Then
        MsgBoxEx "ʱ�����֤ʧ�ܣ�", vbExclamation, gstrSysName
        Exit Function 'ʱ�����֤
    End If
    If Not ZGRYY_CheckCert(strCurrCertSn) Then Exit Function
    
    ZGRYY_VerifySign = mobjAXCA.VerifyString(strSignData, True, False, strSource)
    If Not ZGRYY_VerifySign Then
        MsgBoxEx "ǩ����֤ʧ�ܣ�" & mobjAXCA.GetLastError, vbExclamation, gstrSysName
        Exit Function
    Else
        MsgBoxEx "ǩ����֤�ɹ���"
    End If
End Function
Private Function GetReturnInfo(ByVal strSign As String) As String
    '׼���ʱ���������Ϣת������
    If strSign = "0001" Then
        GetReturnInfo = "����ͨ���쳣"
    ElseIf strSign = "0002" Then
        GetReturnInfo = "ϵͳ�쳣"
    ElseIf strSign = "0003" Then
        GetReturnInfo = "ϵͳ��æ"
    ElseIf strSign = "0004" Then
        GetReturnInfo = "���ݲ������Ϸ�"
    ElseIf strSign = "0005" Then
        GetReturnInfo = "�û������������"
    ElseIf strSign = "0006" Then
        GetReturnInfo = "���ݿ��쳣"
    ElseIf strSign = "0007" Then
        GetReturnInfo = "DLL�����ļ���ȡ����"
    ElseIf strSign = "1001" Then
        GetReturnInfo = "������Ӧʧ��"
    ElseIf strSign = "1002" Then
        GetReturnInfo = "���������ѼӸǹ�ʱ���"
    ElseIf strSign = "1003" Then
        GetReturnInfo = "�������ݵȴ��Ӹ�ʱ���"
    ElseIf strSign = "2001" Then
        GetReturnInfo = "δ����ʱ���"
    ElseIf strSign = "2002" Then
        GetReturnInfo = "У��ʧ��"
    ElseIf strSign = "2010" Then
        GetReturnInfo = "��֤�ɹ�"
    Else
        GetReturnInfo = strSign
    End If
    If GetReturnInfo <> "" Then
        GetReturnInfo = "ʱ����ӿڷ�����ʾ��" & GetReturnInfo
    End If
End Function

Private Function Times_Tamp(ByVal strSource As String, ByRef strTimeStamp As String) As Boolean
        'ȡʱ���
        Dim intCount As Integer, strSign As String
        On Error GoTo hErr
    
100     strSign = mobjTSA.sendTimestamp(strSource, "sha1")
102     If strSign <> "1000" And strSign <> "1002" And strSign <> "1003" Then
104         strSign = GetReturnInfo(strSign)
106         MsgBoxEx "����ʱ���ʧ�ܣ�" & strSign, vbExclamation, gstrSysName
            Times_Tamp = False
            Exit Function
        Else
108         intCount = 0
110         Do While intCount <= 10
112             strSign = mobjTSA.gettimestampinfo(strSource, "sha1")
                'ǩ���л���ʱ��
114             If InStr(strSign, "#") > 0 Then
116                 strTimeStamp = Split(strSign, "#")(0)
118                 If IsDate(strTimeStamp) Then
120                     strTimeStamp = Format(CDate(strTimeStamp), "yyyy-MM-dd HH:mm:ss")
                        Times_Tamp = True
                        Exit Function
                    Else
122                     MsgBoxEx "��ȡ��ʱ�������һ�����ڣ�" & strTimeStamp, vbExclamation, gstrSysName
                    End If
124             ElseIf strSign <> "1003" Then
126                 strSign = GetReturnInfo(strSign)
128                 MsgBoxEx "��ȡʱ���ʧ�ܣ�" & strSign, vbExclamation, gstrSysName
                    Exit Function
                End If
130             intCount = intCount + 1
            Loop
        End If
132     Times_Tamp = True
        Exit Function
hErr:
134     MsgBoxEx "ȡʱ���-��" & CStr(Erl()) & "��," & Err.Description, vbExclamation, gstrSysName
End Function

Private Function verify_Timestamp(ByVal strSource As String) As Boolean
    '��֤ʱ���
    Dim strData As String
    strData = mobjTSA.verifyTimeStamp(strSource, "sha1")
    If strData <> "2010" Then
        MsgBoxEx "��֤ʱ���ʧ�ܣ�" & GetReturnInfo(strData), vbExclamation, gstrSysName
        Exit Function
    End If
    verify_Timestamp = True
End Function

Private Function verify_getTimestamp(ByVal strSource As String) As String
    '��ȡʱ���  ������Ҽӵġ�
    Dim strData As String
    Dim strTimeStamp As String
    strData = mobjTSA.gettimestampinfo(strSource, "sha1")
    If strData = "2001" Then
        MsgBoxEx "��ȡ��֤ʱ���ʧ�ܣ�" & GetReturnInfo(strData), vbExclamation, gstrSysName
        verify_getTimestamp = "��"
        Exit Function
    End If
    
    If InStr(strData, "#") > 0 Then
        strTimeStamp = Split(strData, "#")(0)
        If IsDate(strTimeStamp) Then
            strTimeStamp = Format(CDate(strTimeStamp), "yyyy-MM-dd HH:mm:ss")
        Else
            MsgBoxEx "��ȡ��ʱ�������һ�����ڣ�" & strTimeStamp, vbExclamation, gstrSysName
            verify_getTimestamp = "��"
            Exit Function
        End If
    End If
    verify_getTimestamp = strTimeStamp
    
End Function

Private Function AxServer_VerifyCert() As Boolean

    '��֤��ǰUSBKey�ķ����֤����֤
        
    Dim strBase64Cert As String, intCheck As Integer
    
    
    AxServer_VerifyCert = False
    If mobjAXSVR Is Nothing Then Set mobjAXSVR = CreateObject("AnXEgovCom.AnXEgovSOF")
    
    '��ȡUSB�е�֤��base64�����ַ���
    strBase64Cert = mobjAXCA.GetSignerCertInfo(5, "")
    
    '���ð��ŷ����֤����֤����
    intCheck = mobjAXSVR.SOF_ValidateCert(strBase64Cert)
    If intCheck <> 0 Then
        mLastPWD = ""
        MsgBoxEx "��ǰ����֤��������֤ʧ�ܣ�������Ϊ" & intCheck
        Exit Function
    End If
    
    AxServer_VerifyCert = True
    
End Function
