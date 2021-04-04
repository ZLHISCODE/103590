Attribute VB_Name = "mdlFTP"
Option Explicit

'-----以下是 FTP 相关函数
Public Function TestFTP(ByVal strUser As String, ByVal strPassWord As String, _
                            ByVal strDevAdress As String, ByVal strFtpPath As String) As String
                            
    Dim FtpNet As New clsFTP, strPath As String, strTmpPath As String           'FTP类
    Dim lngFileNo As Long
    strPath = Format(Now, "yyyymmddHHMMSS")
    strTmpPath = IIf(Right(App.Path, 1) <> "\", App.Path & "\", App.Path) & "temp.txt"
    lngFileNo = FreeFile
    Open strTmpPath For Output As lngFileNo
    Close lngFileNo
    If FtpNet.FuncFtpConnect(strDevAdress, strUser, strPassWord) > 0 Then
        If FtpNet.FuncFtpMkDir(strFtpPath, "FTP测试" & strPath) > 0 Then
            TestFTP = "在FTP上不能创建目录！"
        Else
            If FtpNet.FuncUploadFile(strFtpPath, strTmpPath, "temp.txt") > 0 Then
                TestFTP = "上传文件失败"
            Else
                FtpNet.FuncFtpDisConnect '先断开，再删除，不然删不掉
                If FtpNet.FuncFtpConnect(strDevAdress, strUser, strPassWord) <= 0 Then
                     TestFTP = "FTP不能连接！"
                ElseIf FtpNet.FuncFtpDelDir(strFtpPath, "FTP测试" & strPath) > 0 Then
                    TestFTP = "在FTP上不能删除目录"
                Else
                    TestFTP = ""
                End If
            End If
        End If
    Else
        TestFTP = "不能连接FTP！"
    End If
    FtpNet.FuncFtpDisConnect
    Set FtpNet = Nothing
    Kill strTmpPath
End Function

Public Function DownFile(ByVal strUser As String, ByVal strPass As String, ByVal strServer As String, _
                          ByVal strFtpFile As String, ByVal strFile As String) As String
        '从FTP服务器下载文件。
        'strUser    :用户名
        'strPass    :密码
        'strServer  :服务器
        'strFtpFile :FTP上的文件。
        'strFile    :本地文件全路径。
        '返回：空串表示成功，否则为错误提示。
        Dim objFtp As New clsFTP, lngReturn As Long, strFtpFileName As String, strLocaFile As String
        Dim strFtpDir As String
        On Error GoTo errH
100     If strFtpFile = "" Then
102         DownFile = "请指定要下载的文件！"
            objFtp.FuncFtpDisConnect
            Set objFtp = Nothing
            Exit Function
        End If
    
104     strFtpFileName = Split(strFtpFile, "/")(UBound(Split(strFtpFile, "/")))
106     strFtpDir = Replace(strFtpFile, "/" & strFtpFileName, "")
108     strLocaFile = strFile
110     If strLocaFile = "" Then
112         DownFile = "请指定下载的文件保存到何处！"
            objFtp.FuncFtpDisConnect
            Set objFtp = Nothing
            Exit Function
        End If
114     If Dir(strLocaFile) <> "" Then
116         DownFile = "要下载的文件已存在！"
            objFtp.FuncFtpDisConnect
            Set objFtp = Nothing
            Exit Function
        End If
    
118     If strServer = "" Then
120         DownFile = "请指定FTP服务器"
            objFtp.FuncFtpDisConnect
            Set objFtp = Nothing
            Exit Function
        End If
122     lngReturn = objFtp.FuncFtpConnect(strServer, strUser, strPass)
124     If lngReturn = 0 Then
126         DownFile = "不能连接服务器！"
            objFtp.FuncFtpDisConnect
            Set objFtp = Nothing
            Exit Function
        End If
128     lngReturn = objFtp.FuncChangeDir(strFtpDir)
130     If lngReturn <> 0 Then
132         DownFile = "不能进入指定的目录，可能是权限不足或服务器上无此目录！"
            objFtp.FuncFtpDisConnect
            Set objFtp = Nothing
            Exit Function
        End If
134     lngReturn = objFtp.FuncDownloadFile(strFtpDir, strLocaFile, strFtpFileName)
136     If lngReturn <> 0 Then
138         DownFile = "下载失败，可能是权限不足或服务器上无此文件！"
            objFtp.FuncFtpDisConnect
            Set objFtp = Nothing
            Exit Function
        End If
        objFtp.FuncFtpDisConnect
140     Set objFtp = Nothing
        Exit Function
errH:
142     DownFile = CStr(Erl()) & "行，" & Err.Description
End Function

Public Function UploadFile(ByVal strUser As String, ByVal strPass As String, ByVal strServer As String, _
                            ByVal strFtpPath As String, ByVal strFile As String, Optional strNewFileName As String) As String
        '按本地文件名上传文件到FTP服务器。
        'strUser    :用户名
        'strPass    :密码
        'strServer  :服务器
        'strFtpPath :FTP上的目录，无目录会自动创建。
        'strFile    :本地文件全路径。
        'strNewFileName: 传到FTP上后的文件名，为空则按本地文件名保存
        '返回：空串表示成功，否则为错误提示。
    
        Dim objFtp As New clsFTP, lngReturn As Long, strFileName As String, strLocaFile As String
        On Error GoTo errH
    
    
100     If Left(strFtpPath, 1) = "/" Then strFtpPath = Mid$(strFtpPath, 2)
    
102     If strServer = "" Then
104         UploadFile = "请指定FTP服务器"
            objFtp.FuncFtpDisConnect
            Set objFtp = Nothing
            Exit Function
        End If
106     strLocaFile = strFile
108     If Dir(strLocaFile) = "" Then
110         UploadFile = "文件" & strLocaFile & "不存在!"
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
            '检查目录是否存在
118         lngReturn = objFtp.FuncChangeDir(strFtpPath)
120         If lngReturn <> 0 Then
122             lngReturn = objFtp.FuncFtpMkDir("/", strFtpPath)
124             If lngReturn <> 0 Then
126                 UploadFile = "创建目录失败！可能是权限不足！"
                    objFtp.FuncFtpDisConnect
                    Set objFtp = Nothing
                    Exit Function
                End If
            End If
        
128         lngReturn = objFtp.FuncUploadFile("/" & strFtpPath, strLocaFile, strFileName)
130         If lngReturn <> 0 Then
132             UploadFile = "上传文件失败，可能是权限不足！"
                objFtp.FuncFtpDisConnect
                Set objFtp = Nothing
                Exit Function

            Else
134             UploadFile = ""
            End If
        Else
136         UploadFile = "不能连接服务器！"
        End If
        objFtp.FuncFtpDisConnect
        Set objFtp = Nothing
        Exit Function
errH:
138     UploadFile = CStr(Erl()) & "行，" & Err.Description
End Function


Public Function DownFiles(ByVal strUser As String, ByVal strPass As String, ByVal strServer As String, _
                          ByVal strFtpFile As String, ByVal strLocaFile As String, strFiles() As String) As String
        '从FTP服务器下载文件。
        'strUser    :用户名
        'strPass    :密码
        'strServer  :服务器
        'strFtpFile :FTP上的文件。
        'strFile    :本地文件全路径。
        '返回：空串表示成功，否则为错误提示。
        Dim objFtp As New clsFTP, lngReturn As Long, strFtpFileName As String
        Dim strFtpDir As String
        On Error GoTo errH
100     If strFtpFile = "" Then
102         DownFiles = "请指定要下载的文件！"
            objFtp.FuncFtpDisConnect
            Set objFtp = Nothing
            Exit Function
        End If
    
104     strFtpFileName = Split(strFtpFile, "/")(UBound(Split(strFtpFile, "/")))
106     strFtpDir = Replace(strFtpFile, "/" & strFtpFileName, "")
110     If strLocaFile = "" Then
112         DownFiles = "请指定下载的文件保存到何处！"
            objFtp.FuncFtpDisConnect
            Set objFtp = Nothing
            Exit Function
        End If
114     If Dir(strLocaFile) <> "" Then
116         DownFiles = "要下载的文件已存在！"
            objFtp.FuncFtpDisConnect
            Set objFtp = Nothing
            Exit Function
        End If
    
118     If strServer = "" Then
120         DownFiles = "请指定FTP服务器"
            objFtp.FuncFtpDisConnect
            Set objFtp = Nothing
            Exit Function
        End If
122     lngReturn = objFtp.FuncFtpConnect(strServer, strUser, strPass)
124     If lngReturn = 0 Then
126         DownFiles = "不能连接服务器！"
            objFtp.FuncFtpDisConnect
            Set objFtp = Nothing
            Exit Function
        End If
128     lngReturn = objFtp.FuncChangeDir(strFtpDir)
130     If lngReturn <> 0 Then
132         DownFiles = "不能进入指定的目录，可能是权限不足或服务器上无此目录！"
            objFtp.FuncFtpDisConnect
            Set objFtp = Nothing
            Exit Function
        End If
134     lngReturn = objFtp.FuncDownLoadFiles(strFtpDir, strLocaFile, strFiles)
136     If lngReturn <> 0 Then
138         DownFiles = "下载失败，可能是权限不足或服务器上无此文件！"
            objFtp.FuncFtpDisConnect
            Set objFtp = Nothing
            Exit Function
        End If
        objFtp.FuncFtpDisConnect
140     Set objFtp = Nothing
        Exit Function
errH:
142     DownFiles = CStr(Erl()) & "行，" & Err.Description
End Function
