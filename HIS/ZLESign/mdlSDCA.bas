Attribute VB_Name = "mdlSDCA"
Option Explicit
Private mobjSDCA As Object  '����ɽ��ǩ��
Private mobjTSA As Object   '����ɽ����ʱ���
Private mobjSVS As Object   '����ɽ������ǩ
Private mobjBase64 As Object   'BASE64����

Public Function SDCA_CheckCert(ByRef blnReDo As Boolean, Optional ByRef strCertID As String) As Boolean
      '--------------------------------------------------------------------------------------------------------------------------
      '���ܣ���ȡUSB�����豸��ʼ������¼
      '����:
      '   ����:blnRedo-֤�������Ҫ���¼��
      '����:
      '--------------------------------------------------------------------------------------------------------------------------
          Dim strUniqueID As String
          Dim strName As String
          Dim strCertSn As String
          Dim strCertDN As String
          Dim strSigCert As String
          Dim strDate As String
          Dim blnRet As Boolean
          
          Dim udtUser As USER_INFO
          

1         On Error GoTo ErrH

2         If Not GetCertList(, strCertSn, , , strUniqueID, strCertID) Then Exit Function
3         If mUserInfo.strUserID = "" Then
4             MsgBoxEx "�������֤��Ϊ��,����ϵ����Ա����Ա������¼�룡", vbOKOnly + vbInformation, gstrSysName
5             Exit Function
6         ElseIf mUserInfo.strUserID <> strUniqueID Then
7             MsgBoxEx "�������֤�ţ�" & _
                     vbCrLf & vbTab & "��" & mUserInfo.strUserID & "��" & vbCrLf & _
                     "֤�����֤��:" & _
                     vbCrLf & vbTab & "��" & strUniqueID & "��" & vbCrLf & _
                     "�û����֤����֤�����֤�Ų����,����ʹ�ã�", vbInformation, gstrSysName
8             Exit Function
9         End If
          
          '��¼��֤
10        blnRet = True
11        If mUserInfo.strCertSn <> strCertSn Then
12            If Not GetCertList(strName, , strSigCert, strCertDN) Then Exit Function
              '�ж��Ƿ���Ҫ����ע��֤��
13            udtUser.strName = strName
14            udtUser.strSignName = strName
15            udtUser.strUserID = strUniqueID
16            udtUser.strCertSn = strCertSn
17            udtUser.strCertDN = strCertDN
18            udtUser.strCert = strSigCert
19            udtUser.strEncCert = ""
20            udtUser.strCertID = strCertID
21            udtUser.strPicCode = "" '֤��ע��ʱ������ͼƬ,ͼƬ���¿�ͨ����Ա��������ǩ��ͼƬ��������ɡ�
22            udtUser.strPicPath = ""
              '��ȡ�Ѿ�ע��֤�����Ч��������
23            strDate = mobjSDCA.SOF_GetCertInfo(mUserInfo.strCert, 18)
24            If IsUpdateRegCert(udtUser, strDate, blnReDo) Then
25                blnRet = True
26            Else
27                blnRet = False
28            End If
29        End If
          
30        SDCA_CheckCert = blnRet

31        Exit Function

ErrH:
32        MsgBox "��zl9ESign.mdlSDCA.SDCA_CheckCert�ĵ�" & Erl() & "�г���" & vbCrLf & _
                  "�����: " & Err.Number & vbCrLf & _
                  "����������" & Err.Description, vbExclamation, gstrSysName


End Function


Public Function SDCA_GetPara() As Boolean
      '��ȡ��������

1         On Error GoTo ErrH

2         With gudtPara
3             .strTSIP = GetThirdPara(CON_PAR_ɽ��, "ʱ�����ַ")
4             .strTSPort = GetThirdPara(CON_PAR_ɽ��, "ʱ����˿�")
5             .strSIGNIP = GetThirdPara(CON_PAR_ɽ��, "ǩ����ַ")
6             .strSignPort = GetThirdPara(CON_PAR_ɽ��, "ǩ���˿�")
7             .bytSignVersion = Val(GetThirdPara(CON_PAR_ɽ��, "�汾"))
8             If .bytSignVersion = 1 Then
9                 If .strSIGNIP = "" Or .strTSIP = "" Then Exit Function
10            End If
11        End With
          
12        SDCA_GetPara = True
13        Exit Function

ErrH:
14        MsgBox "��zl9ESign.mdlSDCA.SDCA_GetPara�ĵ�" & Erl() & "�г���" & vbCrLf & _
                  "�����: " & Err.Number & vbCrLf & _
                  "����������" & Err.Description, vbExclamation, gstrSysName

       
End Function

Public Function SDCA_InitObj() As Boolean
          Dim strMsg As String
          
1         On Error Resume Next
2         Set mobjSDCA = CreateObject("SDCASecurityClient.CASecurityClient.1")
3         If Err.Number <> 0 Then
              'C:\Windows\system32\SDCASecurityClient.dll
4             strMsg = "����ǩ������SDCASecurityClient.CASecurityClient.1��ʧ�ܣ����鲿����SDCASecurityClient.dll���Ƿ���ȷ��װ��ע�ᡣ"
5             GoTo ErrH
6         End If
7         On Error GoTo ErrH
          
8         On Error Resume Next
9         Set mobjSVS = CreateObject("NetONEX.SVSClientX.1")
10        If Err.Number <> 0 Then
              'C:\Windows\system32\NetONEX.dll
11            strMsg = "������ǩ����NetONEX.SVSClientX.1��ʧ�ܣ����鲿����NetONEX.dll���Ƿ���ȷ��װ��ע�ᡣ"
12            GoTo ErrH
13        End If
14        On Error GoTo ErrH
          '������ַ��"60.216.5.244" �˿� 9189
15        mobjSVS.ServerAddress = gudtPara.strSIGNIP
16        mobjSVS.ServerPort = Val(gudtPara.strSignPort)
17        On Error Resume Next
18        Set mobjTSA = CreateObject("NetONEX.TSAClientX.1")
19        If Err.Number <> 0 Then
              'C:\Windows\system32\NetONEX.dll
20            strMsg = "����ʱ�������NetONEX.TSAClientX.1��ʧ�ܣ����鲿����NetONEX.dll���Ƿ���ȷ��װ��ע�ᡣ"
21            GoTo ErrH
22        End If
23        On Error GoTo ErrH
          '������ַ��"60.216.5.244" �˿� 9198
24        mobjTSA.ServerAddress = gudtPara.strTSIP
25        mobjTSA.ServerPort = Val(gudtPara.strTSPort)
          
26        On Error Resume Next
27        Set mobjBase64 = CreateObject("NetONEX.Base64X.1")
28        If Err.Number <> 0 Then
               'C:\Windows\system32\NetONEX.dll
29            strMsg = "������ǩ����NetONEX.Base64X.1��ʧ�ܣ����鲿����NetONEX.dll���Ƿ���ȷ��װ��ע�ᡣ"
30            GoTo ErrH
31        End If
32        On Error GoTo ErrH
          
33        SDCA_InitObj = True
34        Exit Function
ErrH:
35       Call GetErrMsg(Erl(), strMsg)
End Function

Public Function SDCA_RegCert(arrCertInfo As Variant) As Boolean
'����:       �ṩ��HIS���ݿ���ע�����֤��ı�Ҫ��Ϣ , �����·��Ż����֤��, , ��Ҫ����USB - Key
'����:       arrCertInfo��Ϊ���鷵��֤�������Ϣ
'            0-ClientSignCertCN:�ͻ���ǩ��֤�鹫������(����),����ע��֤��ʱ������֤���
'            1-ClientSignCertDN:�ͻ���ǩ��֤������(ÿ��Ψһ)
'            2-ClientSignCertSN:�ͻ���ǩ��֤�����к�(ÿ֤��Ψһ)
'            3-ClientSignCert:�ͻ���ǩ��֤������
'            4-ClientEncCert:�ͻ��˼���֤������
'            5-ǩ��ͼƬ�ļ���,�մ���ʾû��ǩ��ͼƬ
'            6-ʱ���֤��
          Dim strSN As String, strName As String, strDn As String
          Dim strCert As String
          Dim strPic As String
          
          Dim i As Long
            
1         On Error GoTo ErrH

2         For i = LBound(arrCertInfo) To UBound(arrCertInfo)
3             arrCertInfo(i) = ""
4         Next
5         If GetCertList(strName, strSN, strCert, strDn, , , strPic) Then
6             arrCertInfo(0) = strName
7             arrCertInfo(1) = strDn
8             arrCertInfo(2) = strSN
9             arrCertInfo(3) = strCert
10            arrCertInfo(4) = ""
11            arrCertInfo(5) = strPic
12            arrCertInfo(6) = ""
13            SDCA_RegCert = True
14        End If

15        Exit Function

ErrH:
16        MsgBox "��zl9ESign.mdlSDCA.SDCA_RegCert�ĵ�" & Erl() & "�г���" & vbCrLf & _
                  "�����: " & Err.Number & vbCrLf & _
                  "����������" & Err.Description, vbExclamation, gstrSysName

End Function

Public Sub SDCA_SetPara()
          
1         On Error GoTo ErrH

2         With gudtPara
3             Call UpdateThirdPara(CON_PAR_ɽ��, 1, "ǩ����ַ", .strSIGNIP, "ǩ�������ַ")
4             Call UpdateThirdPara(CON_PAR_ɽ��, 2, "ǩ���˿�", .strSignPort, "ǩ������˿�")
5             Call UpdateThirdPara(CON_PAR_ɽ��, 3, "ʱ�����ַ", .strTSIP, "ʱ��������ַ")
6             Call UpdateThirdPara(CON_PAR_ɽ��, 4, "ʱ����˿�", .strTSPort, "ʱ�������˿�")
7             Call UpdateThirdPara(CON_PAR_ɽ��, 5, "�汾", .bytSignVersion, "�ӿڰ汾")
8         End With

9         Exit Sub

ErrH:
10        MsgBox "��zl9ESign.mdlSDCA.SDCA_SetPara�ĵ�" & Erl() & "�г���" & vbCrLf & _
                  "�����: " & Err.Number & vbCrLf & _
                  "����������" & Err.Description, vbExclamation, gstrSysName
End Sub

Public Function SDCA_Sign(ByVal strSource As String, ByRef strSignData As String, ByRef strTimeStamp As String, ByRef strTimeStampCode As String, ByRef blnReDo As Boolean)
      '����:ǩ��
          Dim strCertID As String
          Dim blnCheck As Boolean
          
          Dim objTsaRes As Object
          Dim objCert As Object

1         On Error GoTo ErrH
          
2         If Not SDCA_CheckCert(blnReDo, strCertID) Then Exit Function
3         If blnReDo Then Exit Function
          
          '�����ڹ��ܿ�SM2�㷨֤�飬Ĭ��0 ����У��һ��
4         Call mobjSDCA.SOF_InitialVar("CERTID", "KeyLoginPolicy", "0")
          '��ȡǩ��ֵ
5         strSignData = mobjSDCA.SOF_SignData(strCertID, strSource)
6         If strSignData = "" Then Exit Function
          
          '��ȡʱ�������ʱ��������ʼ����ʱ���Ѿ�ʵ������
7         Set objTsaRes = mobjTSA.TSACreate(strSource)
                  
          '��ȡʱ�������
8         strTimeStampCode = objTsaRes.ToBASE64
          
9         If strTimeStampCode = "" Then
10            MsgBoxEx "��ȡʱ�����Ϣʧ�ܣ�", vbInformation, gstrSysName
11            Exit Function
12        End If
          'Oct 13 10:14:47 2019 GMT
13        strTimeStamp = objTsaRes.TimeStamp
14        LogWrite "SDCA_Sign", "�ӿڡ�TimeStamp������ֵ:" & strTimeStamp
15        If strTimeStamp = "" Then
16            MsgBoxEx "��ʱ���ǩ��ֵ�л�ȡǩ��ʱ��ʧ�ܣ�", vbInformation, gstrSysName
17            Exit Function
18        End If
19        strTimeStamp = GetTimeStamp(strTimeStamp)
20        strTimeStamp = Format(strTimeStamp, "yyyy-MM-dd HH:mm:ss")
       
21        SDCA_Sign = True

22        Exit Function

ErrH:
23        MsgBox "��zl9ESign.mdlSDCA.SDCA_Sign�ĵ�" & Erl() & "�г���" & vbCrLf & _
                  "�����: " & Err.Number & vbCrLf & _
                  "����������" & Err.Description, vbExclamation, gstrSysName


End Function

Public Sub SDCA_UnloadObj()
    Set mobjSDCA = Nothing
    Set mobjSVS = Nothing
    Set mobjTSA = Nothing
    Set mobjBase64 = Nothing
End Sub

Public Function SDCA_VerifySign(ByVal strCert As String, ByVal strSign As String, ByVal strSource As String, ByVal strTStampCode As String) As Boolean
          '����:��֤ǩ��
          'ʱ�����֤ǩ��
          Dim lngRet As Long
          Dim strEncode As String
          
          '��֤ǩ����֤����Ч��,����ֵΪ200 ����ȷ,����Ϊ����
1         On Error GoTo ErrH
2         strEncode = mobjBase64.EncodeString(strSource)
3         lngRet = mobjSVS.SVSVerifyPKCS1(strCert, strSign, strEncode)
4         If lngRet <> 200 Then
5             MsgBoxEx "ǩ����ǩʧ�ܣ�", vbInformation, gstrSysName
6             Exit Function
7         End If
        
8         lngRet = mobjTSA.TSAVerify(strTStampCode)
9         If lngRet <> 200 Then
10            MsgBoxEx "ʱ�����֤ʧ�ܣ�", vbInformation, gstrSysName
11            Exit Function
12        End If
13        MsgBoxEx "��ǩ�ɹ���", vbInformation, gstrSysName
14        SDCA_VerifySign = True
          
15        Exit Function
ErrH:
16        MsgBox "��zl9ESign.mdlSDCA.SDCA_VerifySign�ĵ�" & Erl() & "�г���" & vbCrLf & _
                  "�����: " & Err.Number & vbCrLf & _
                  "����������" & Err.Description, vbExclamation, gstrSysName

End Function

Private Function GetCertList(Optional ByRef strName As String = "1", Optional ByRef strCertSn As String = "1", _
   Optional ByRef strCert As String, Optional ByRef strCertDN As String = "1", Optional ByRef strUniqueID As String = "1", _
   Optional ByRef strCertID As String, Optional ByRef strPic As String = "1") As Boolean
          'strPic-�ļ�·��
          Dim lngRet As Long
          
1         On Error GoTo ErrH

2         strCertID = mobjSDCA.SOF_GetUserList()
3         If strCertID <> "" Then
4             strCert = mobjSDCA.SOF_ExportUserCert(strCertID) '֤���ַ���
5             If strCert <> "" Then
6                 If strCertSn <> "1" Then strCertSn = mobjSDCA.SOF_GetCertInfo(strCert, 2) '֤�����к�
7                 If strName <> "1" Then strName = mobjSDCA.SOF_GetCertInfo(strCert, 23) '֤��ͨ��������
8                 If strCertDN <> "1" Then strCertDN = mobjSDCA.SOF_GetCertInfo(strCert, 33) '֤��ӵ����DN
9                 If strUniqueID <> "1" Then
                    strUniqueID = mobjSDCA.SOF_GetCertInfo(strCert, 53) 'Ψһ��ʶ
                    strUniqueID = Right(strUniqueID, 18)
                  End If
10            End If
11            If strPic <> "1" Then
12                strPic = App.Path & "\pic.bmp"
13                lngRet = mobjSDCA.SOF_ShowSeal(strCertID, 0, strPic, 3)
14                If lngRet = 0 Then strPic = "" '��ȡʧ��
15            End If
16            GetCertList = True
17        Else
18            MsgBoxEx "û���ҵ�Key�̣����飡", vbInformation, gstrSysName
19            Exit Function
20        End If

21        Exit Function

ErrH:
22        MsgBox "��zl9ESign.mdlSDCA.GetCertList�ĵ�" & Erl() & "�г���" & vbCrLf & _
                  "�����: " & Err.Number & vbCrLf & _
                  "����������" & Err.Description, vbExclamation, gstrSysName
End Function

Private Function GetTimeStamp(ByVal strTimeStamp As String) As String
      '���ܣ���ȡʱ����е�ʱ��
          Dim arrTime As Variant
          Dim strTime As String
          
1         On Error GoTo ErrH

2         strTimeStamp = Replace(strTimeStamp, " ", "|")
3         strTimeStamp = Replace(strTimeStamp, "||", "|") '������Ϊһλ��ʱ,��ֹ�·ݺ�����֮����������ո�����
4         arrTime = Split(strTimeStamp, "|")  '���˸�ʽ��Aug 19 13:07:25 2014 GMT
5         strTime = arrTime(0) & " " & arrTime(1) & " " & arrTime(3)  '��/��/��
6         strTime = CDate(strTime) & ""  '�� �� ��  2014/8/19
7         GetTimeStamp = strTime & " " & arrTime(2)  ' ��-��-�� ʱ:��:��

8         Exit Function

ErrH:
9         MsgBox "��zl9ESign.mdlSDCA.GetTimeStamp�ĵ�" & Erl() & "�г���" & vbCrLf & _
                  "�����: " & Err.Number & vbCrLf & _
                  "����������" & Err.Description, vbExclamation, gstrSysName

End Function
