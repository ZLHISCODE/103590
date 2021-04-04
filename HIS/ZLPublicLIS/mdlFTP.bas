Attribute VB_Name = "mdlFTP"
Public Function TestFTP(ByVal strUser As String, ByVal strPassWord As String, _
                            ByVal strDevAdress As String, ByVal strFtpPath As String) As String
                            
    Dim FtpNet As New clsFTP, strPath As String, strTmpPath As String           'FTP��
    Dim lngFileNo As Long
    strPath = Format(Now, "yyyymmddHHMMSS")
    strTmpPath = IIf(Right(App.Path, 1) <> "\", App.Path & "\", App.Path) & "temp.txt"
    lngFileNo = FreeFile
    Open strTmpPath For Output As lngFileNo
    Close lngFileNo
    If FtpNet.FuncFtpConnect(strDevAdress, strUser, strPassWord) > 0 Then
        If FtpNet.FuncFtpMkDir(strFtpPath, "FTP����" & strPath) > 0 Then
            TestFTP = "��FTP�ϲ��ܴ���Ŀ¼��"
        Else
            If FtpNet.FuncUploadFile(strFtpPath, strTmpPath, "temp.txt") > 0 Then
                TestFTP = "�ϴ��ļ�ʧ��"
            Else
                FtpNet.FuncFtpDisConnect '�ȶϿ�����ɾ������Ȼɾ����
                If FtpNet.FuncFtpConnect(strDevAdress, strUser, strPassWord) <= 0 Then
                     TestFTP = "FTP�������ӣ�"
                ElseIf FtpNet.FuncFtpDelDir(strFtpPath, "FTP����" & strPath) > 0 Then
                    TestFTP = "��FTP�ϲ���ɾ��Ŀ¼"
                Else
                    TestFTP = ""
                End If
            End If
        End If
    Else
        TestFTP = "��������FTP��"
    End If
    FtpNet.FuncFtpDisConnect
    Set FtpNet = Nothing
    Kill strTmpPath
End Function

Public Function DownFile(ByVal strUser As String, ByVal strPass As String, ByVal strServer As String, _
                          ByVal strFtpFile As String, ByVal strFile As String) As String
        '��FTP�����������ļ���
        'strUser    :�û���
        'strPass    :����
        'strServer  :������
        'strFtpFile :FTP�ϵ��ļ���
        'strFile    :�����ļ�ȫ·����
        '���أ��մ���ʾ�ɹ�������Ϊ������ʾ��
        Dim objFtp As New clsFTP, lngReturn As Long, strFtpFileName As String, strLocaFile As String
        Dim strFtpDir As String
        On Error GoTo errH
100     If strFtpFile = "" Then
102         DownFile = "��ָ��Ҫ���ص��ļ���"
            objFtp.FuncFtpDisConnect
            Set objFtp = Nothing
            Exit Function
        End If
    
104     strFtpFileName = Split(strFtpFile, "/")(UBound(Split(strFtpFile, "/")))
106     strFtpDir = Replace(strFtpFile, "/" & strFtpFileName, "")
108     strLocaFile = strFile
110     If strLocaFile = "" Then
112         DownFile = "��ָ�����ص��ļ����浽�δ���"
            objFtp.FuncFtpDisConnect
            Set objFtp = Nothing
            Exit Function
        End If
114     If Dir(strLocaFile) <> "" Then
116         DownFile = "Ҫ���ص��ļ��Ѵ��ڣ�"
            objFtp.FuncFtpDisConnect
            Set objFtp = Nothing
            Exit Function
        End If
    
118     If strServer = "" Then
120         DownFile = "��ָ��FTP������"
            objFtp.FuncFtpDisConnect
            Set objFtp = Nothing
            Exit Function
        End If
122     lngReturn = objFtp.FuncFtpConnect(strServer, strUser, strPass)
124     If lngReturn = 0 Then
126         DownFile = "�������ӷ�������"
            objFtp.FuncFtpDisConnect
            Set objFtp = Nothing
            Exit Function
        End If
128     lngReturn = objFtp.FuncChangeDir(strFtpDir)
130     If lngReturn <> 0 Then
132         DownFile = "���ܽ���ָ����Ŀ¼��������Ȩ�޲������������޴�Ŀ¼��"
            objFtp.FuncFtpDisConnect
            Set objFtp = Nothing
            Exit Function
        End If
134     lngReturn = objFtp.FuncDownloadFile(strFtpDir, strLocaFile, strFtpFileName)
136     If lngReturn <> 0 Then
138         DownFile = "����ʧ�ܣ�������Ȩ�޲������������޴��ļ���"
            objFtp.FuncFtpDisConnect
            Set objFtp = Nothing
            Exit Function
        End If
        objFtp.FuncFtpDisConnect
140     Set objFtp = Nothing
        Exit Function
errH:
142     DownFile = CStr(Erl()) & "�У�" & Err.Description
End Function

