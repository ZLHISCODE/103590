Attribute VB_Name = "mdlFTP"
Option Explicit

'-----������ FTP ��غ���
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

Public Function UploadFile(ByVal strUser As String, ByVal strPass As String, ByVal strServer As String, _
                            ByVal strFtpPath As String, ByVal strFile As String, Optional strNewFileName As String) As String
        '�������ļ����ϴ��ļ���FTP��������
        'strUser    :�û���
        'strPass    :����
        'strServer  :������
        'strFtpPath :FTP�ϵ�Ŀ¼����Ŀ¼���Զ�������
        'strFile    :�����ļ�ȫ·����
        'strNewFileName: ����FTP�Ϻ���ļ�����Ϊ���򰴱����ļ�������
        '���أ��մ���ʾ�ɹ�������Ϊ������ʾ��
    
        Dim objFtp As New clsFTP, lngReturn As Long, strFileName As String, strLocaFile As String
        On Error GoTo errH
    
    
100     If Left(strFtpPath, 1) = "/" Then strFtpPath = Mid$(strFtpPath, 2)
    
102     If strServer = "" Then
104         UploadFile = "��ָ��FTP������"
            objFtp.FuncFtpDisConnect
            Set objFtp = Nothing
            Exit Function
        End If
106     strLocaFile = strFile
108     If Dir(strLocaFile) = "" Then
110         UploadFile = "�ļ�" & strLocaFile & "������!"
            objFtp.FuncFtpDisConnect
            Set objFtp = Nothing
            Exit Function
        End If
        If strNewFileName = "" Then
112         strFileName = Split(strLocaFile, "\")(UBound(Split(strLocaFile, "\")))
        Else
            strFileName = strNewFileName
        End If
114     lngReturn = objFtp.FuncFtpConnect(strServer, strUser, strPass)
116     If lngReturn <> 0 Then
            '���Ŀ¼�Ƿ����
118         lngReturn = objFtp.FuncChangeDir(strFtpPath)
120         If lngReturn <> 0 Then
122             lngReturn = objFtp.FuncFtpMkDir("/", strFtpPath)
124             If lngReturn <> 0 Then
126                 UploadFile = "����Ŀ¼ʧ�ܣ�������Ȩ�޲��㣡"
                    objFtp.FuncFtpDisConnect
                    Set objFtp = Nothing
                    Exit Function
                End If
            End If
        
128         lngReturn = objFtp.FuncUploadFile("/" & strFtpPath, strLocaFile, strFileName)
130         If lngReturn <> 0 Then
132             UploadFile = "�ϴ��ļ�ʧ�ܣ�������Ȩ�޲��㣡"
                objFtp.FuncFtpDisConnect
                Set objFtp = Nothing
                Exit Function

            Else
134             UploadFile = ""
            End If
        Else
136         UploadFile = "�������ӷ�������"
        End If
        objFtp.FuncFtpDisConnect
        Set objFtp = Nothing
        Exit Function
errH:
138     UploadFile = CStr(Erl()) & "�У�" & Err.Description
End Function


Public Function DownFiles(ByVal strUser As String, ByVal strPass As String, ByVal strServer As String, _
                          ByVal strFtpFile As String, ByVal strLocaFile As String, strFiles() As String) As String
        '��FTP�����������ļ���
        'strUser    :�û���
        'strPass    :����
        'strServer  :������
        'strFtpFile :FTP�ϵ��ļ���
        'strFile    :�����ļ�ȫ·����
        '���أ��մ���ʾ�ɹ�������Ϊ������ʾ��
        Dim objFtp As New clsFTP, lngReturn As Long, strFtpFileName As String
        Dim strFtpDir As String
        On Error GoTo errH
100     If strFtpFile = "" Then
102         DownFiles = "��ָ��Ҫ���ص��ļ���"
            objFtp.FuncFtpDisConnect
            Set objFtp = Nothing
            Exit Function
        End If
    
104     strFtpFileName = Split(strFtpFile, "/")(UBound(Split(strFtpFile, "/")))
106     strFtpDir = Replace(strFtpFile, "/" & strFtpFileName, "")
110     If strLocaFile = "" Then
112         DownFiles = "��ָ�����ص��ļ����浽�δ���"
            objFtp.FuncFtpDisConnect
            Set objFtp = Nothing
            Exit Function
        End If
114     If Dir(strLocaFile) <> "" Then
116         DownFiles = "Ҫ���ص��ļ��Ѵ��ڣ�"
            objFtp.FuncFtpDisConnect
            Set objFtp = Nothing
            Exit Function
        End If
    
118     If strServer = "" Then
120         DownFiles = "��ָ��FTP������"
            objFtp.FuncFtpDisConnect
            Set objFtp = Nothing
            Exit Function
        End If
122     lngReturn = objFtp.FuncFtpConnect(strServer, strUser, strPass)
124     If lngReturn = 0 Then
126         DownFiles = "�������ӷ�������"
            objFtp.FuncFtpDisConnect
            Set objFtp = Nothing
            Exit Function
        End If
128     lngReturn = objFtp.FuncChangeDir(strFtpDir)
130     If lngReturn <> 0 Then
132         DownFiles = "���ܽ���ָ����Ŀ¼��������Ȩ�޲������������޴�Ŀ¼��"
            objFtp.FuncFtpDisConnect
            Set objFtp = Nothing
            Exit Function
        End If
134     lngReturn = objFtp.FuncDownLoadFiles(strFtpDir, strLocaFile, strFiles)
136     If lngReturn <> 0 Then
138         DownFiles = "����ʧ�ܣ�������Ȩ�޲������������޴��ļ���"
            objFtp.FuncFtpDisConnect
            Set objFtp = Nothing
            Exit Function
        End If
        objFtp.FuncFtpDisConnect
140     Set objFtp = Nothing
        Exit Function
errH:
142     DownFiles = CStr(Erl()) & "�У�" & Err.Description
End Function
