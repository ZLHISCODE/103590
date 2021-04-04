Attribute VB_Name = "mdlFtp"
Option Explicit



Private mblnIsPass As Boolean
Private mblnIsForceRead As Boolean
Private mblnIsCompareSize As Boolean


Private mobjFtps As Dictionary


Public Function FtpCreateTag(ByVal strIP As String, _
    ByVal strUser As String, _
    ByVal strPwd As String, _
    ByVal strVirtualPath As String, _
    Optional ByVal lngPort As Long, _
    Optional ByVal strShareDir As String, _
    Optional ByVal strShareUser As String, _
    Optional ByVal strSharePwd As String) As TFtpConTag

    FtpCreateTag.Ip = ""
    If strIP = "" Then Exit Function
    
    FtpCreateTag.Ip = strIP
    FtpCreateTag.Port = lngPort
    FtpCreateTag.User = strUser
    FtpCreateTag.pwd = strPwd
    FtpCreateTag.VirtualPath = strVirtualPath
    FtpCreateTag.ShareDir = strShareDir
    FtpCreateTag.ShareUser = strShareUser
    FtpCreateTag.SharePwd = strSharePwd
End Function


Public Sub FtpParInit()
'ftp相关参数初始化
    If Not mobjFtps Is Nothing Then Exit Sub
    
    Set mobjFtps = New Dictionary
    
    mblnIsPass = IIf(Val(GetSetting("ZLSOFT", "公共模块\Ftp", "启用被动传输", 0)) = 1, True, False)
    Call SaveSetting("ZLSOFT", "公共模块\Ftp", "启用被动传输", IIf(mblnIsPass, 1, 0))
    
    mblnIsForceRead = IIf(Val(GetSetting("ZLSOFT", "公共模块\Ftp", "启用强制读取", 0)) = 1, True, False)
    Call SaveSetting("ZLSOFT", "公共模块\Ftp", "启用强制读取", IIf(mblnIsForceRead, 1, 0))
    
    mblnIsCompareSize = IIf(Val(GetSetting("ZLSOFT", "公共模块\Ftp", "启用FTP文件大小对比", 0)) = 1, True, False)
    Call SaveSetting("ZLSOFT", "公共模块\Ftp", "启用FTP文件大小对比", IIf(mblnIsCompareSize, 1, 0))
    
End Sub

Public Function FtpIsValid(ByRef ftpTag As TFtpConTag) As Boolean
'ftp 是否有效
    Dim objFtp As clsFtp
    Dim lngResult As Long
    
    FtpIsValid = False
    
    Set objFtp = GetFtpInstance(ftpTag, False)
    
    lngResult = objFtp.FuncFtpConnect(ftpTag.Ip, ftpTag.User, ftpTag.pwd, mblnIsForceRead)
    If lngResult = 0 Then Exit Function
    
    FtpIsValid = True
End Function

'传输文件
Private Function FtpFileTransfer(ByRef ftpTag As TFtpConTag, ByVal strFtpFile As String, ByVal strLocalFile As String, _
    ByVal lngTransferWay As Long, Optional ByVal blnIsAutoDiscon As Boolean = True, _
    Optional ByVal blnIsAutoHint As Boolean = True) As FtpResult
    
    Dim objFtp As clsFtp
    Dim objFileSystem As New FileSystemObject
    
    Dim blnIsRetry As Boolean
    Dim lngResult As Long
    Dim lngDestFileSize As Long
    Dim lngFtpFileSize As Long
    Dim strMessage As String
    Dim blnFailed As Boolean
    
    Dim strFtpClassPath As String
    Dim strFtpFileName As String
    
    Dim strFtpMsg As String
    
    
    FtpFileTransfer = frNormal
    
    If mobjFtps Is Nothing Then Call FtpParInit
    
    Set objFtp = GetFtpInstance(ftpTag, blnIsAutoDiscon)
     
    strFtpClassPath = objFileSystem.GetParentFolderName(Replace(ftpTag.VirtualPath & "/" & strFtpFile, "//", "/"))
    strFtpFileName = objFileSystem.GetFileName(strFtpFile)
    

    If lngTransferWay = 0 Then
        '下载文件
        lngResult = objFtp.FuncDownloadFile(strFtpClassPath, strLocalFile, strFtpFileName, mblnIsForceRead)
    Else
        If Trim(strFtpClassPath) <> "" Then Call objFtp.FuncFtpMkDir("/", strFtpClassPath)
        
        '上传文件
        'If lngResult = 0 Then
            lngResult = objFtp.FuncUploadFile(strFtpClassPath, strLocalFile, strFtpFileName)
        'End If
    End If
    
    blnIsRetry = False
    
    If lngResult = 0 Then
        'FuncDownloadFile返回0表示成功
        '文件大小检查
        If mblnIsCompareSize Then
            lngDestFileSize = objFileSystem.GetFile(strLocalFile).Size
            lngFtpFileSize = objFtp.FuncFtpGetFileSize(strFtpClassPath, strFtpFileName)
            
            If lngFtpFileSize <> lngDestFileSize Then
                strFtpMsg = Replace(objFtp.GetFtpMsg(), Chr(0), "")
                
                strMessage = "本地文件大小[" & lngDestFileSize & "]与FTP文件大小[" & lngFtpFileSize & "]不一致" & vbCrLf & _
                             "本地文件：" & strLocalFile & vbCrLf & _
                             "FTP文件：" & strFtpClassPath & "," & strFtpFileName & vbCrLf & _
                             IIf(Trim(strFtpMsg) <> "", "FTP响应消息:" & strFtpMsg & vbCrLf, "") & _
                             "是否需要重新" & IIf(lngTransferWay = 0, "下载?", "上传?")
                             
                lngResult = MsgboxEx(GetForegroundWindow, strMessage, vbAbortRetryIgnore, "提示")
                objFtp.FuncFtpDisConnect
                
                If lngResult = vbRetry Then
                    blnIsRetry = True
                ElseIf lngResult = vbIgnore Then
                    '忽略
                    FtpFileTransfer = frIgnore
                    Exit Function
                Else
                    '终止
                    FtpFileTransfer = frAbort
                    Exit Function
                End If
            Else
                '文件大小相同
                objFtp.FuncFtpDisConnect
                Exit Function
            End If
        Else
            '不进行文件大小检查
            objFtp.FuncFtpDisConnect
            Exit Function
        End If
    Else
        blnIsRetry = True
    End If
    
    
    Do While blnIsRetry
            
        blnFailed = False
        
        objFtp.FuncFtpDisConnect
        lngResult = objFtp.FuncFtpConnect(ftpTag.Ip, ftpTag.User, ftpTag.pwd, mblnIsForceRead)
        
    
        'FuncFtpConnect 返回0表示失败
        If lngResult = 0 Then
            'ftp连接失败
            blnFailed = True
            strFtpMsg = Replace(objFtp.GetFtpMsg(), Chr(0), "")
            
            strMessage = "FTP:" & ftpTag.Ip & " 连接失败,是否重试？" & vbCrLf & _
                            IIf(Trim(strFtpMsg) <> "", "FTP响应消息:" & strFtpMsg, "")
                            

        Else
            If lngTransferWay = 0 Then
                lngResult = objFtp.FuncDownloadFile(strFtpClassPath, strLocalFile, strFtpFileName, mblnIsForceRead)
            Else
                If Trim(strFtpClassPath) <> "" Then Call objFtp.FuncFtpMkDir("/", strFtpClassPath)
                
                'If lngResult = 0 Then
                    lngResult = objFtp.FuncUploadFile(strFtpClassPath, strLocalFile, strFtpFileName)
                'End If
            End If
            
            If lngResult <> 0 Then
                '文件传输失败
                blnFailed = True
                strFtpMsg = Replace(objFtp.GetFtpMsg(), Chr(0), "")
                
                strMessage = "从FTP:" & ftpTag.Ip & IIf(lngTransferWay = 0, " 下载", " 上传") & "文件 [" & ftpTag.VirtualPath & " , " & strFtpFile & "] 失败,是否重试？" & vbCrLf & _
                            IIf(Trim(strFtpMsg) <> "", "FTP响应消息:" & strFtpMsg, "")
            Else
                '文件大小检查
                If mblnIsCompareSize Then
                    lngDestFileSize = objFileSystem.GetFile(strLocalFile).Size
                    lngFtpFileSize = objFtp.FuncFtpGetFileSize(strFtpClassPath, strFtpFileName)
                    
                    If lngFtpFileSize <> lngDestFileSize Then
                        blnFailed = True
                        strFtpMsg = Replace(objFtp.GetFtpMsg(), Chr(0), "")
                        
                        strMessage = "本地文件大小[" & lngDestFileSize & "]与FTP文件大小[" & lngFtpFileSize & "]不一致" & vbCrLf & _
                                     "本地文件：" & strLocalFile & vbCrLf & _
                                     "FTP文件：" & strFtpClassPath & ":" & strFtpFileName & vbCrLf & _
                                     IIf(Trim(strFtpMsg) <> "", "FTP响应消息:" & strFtpMsg & vbCrLf, "") & _
                                     "是否需要重新" & IIf(lngTransferWay = 0, "下载?", "上传?")
                    End If
                End If
            End If
            
        End If
        
        If blnFailed And blnIsAutoHint Then
            lngResult = MsgboxEx(GetForegroundWindow, strMessage, vbAbortRetryIgnore, "提示")
            
            Call objFtp.FuncFtpDisConnect
            
            If lngResult = vbRetry Then
                blnIsRetry = True
            ElseIf lngResult = vbIgnore Then
                '忽略
                FtpFileTransfer = frIgnore
                Exit Function
            Else
                '终止
                FtpFileTransfer = frAbort
                Exit Function
            End If
        Else
            If blnFailed Then FtpFileTransfer = frIgnore
            blnIsRetry = False
        End If
    Loop
        
  
    If blnIsAutoDiscon Then
        '断开FTP连接
        Call objFtp.FuncFtpDisConnect
        
        Call FtpDiscon(ftpTag)
    End If
    
End Function

Private Function GetFtpKey(ByRef ftpTag As TFtpConTag) As String
'获取ftpkey信息
    GetFtpKey = ""
    If Trim(ftpTag.Ip) = "" Then Exit Function
    
    GetFtpKey = "ftp://" & ftpTag.User & ":" & ftpTag.pwd & "@" & ftpTag.Ip & ":" & ftpTag.Port '& "/" & ftpTag.VirtualPath
End Function


Private Function GetFtpInstance(ByRef ftpTag As TFtpConTag, Optional ByVal blnIsAutoDiscon As Boolean = True) As clsFtp
'获取ftp连接实例
    Dim strFtpKey As String
    
    Set GetFtpInstance = Nothing
    
    strFtpKey = GetFtpKey(ftpTag)
    
    If mobjFtps.Exists(strFtpKey) Then
        Set GetFtpInstance = mobjFtps(strFtpKey)
    End If
    
    If GetFtpInstance Is Nothing Then
        Set GetFtpInstance = New clsFtp
        
        Call GetFtpInstance.FuncFtpConnect(ftpTag.Ip, ftpTag.User, ftpTag.pwd, mblnIsForceRead)
    End If
    
    If Not blnIsAutoDiscon Then
        '判断集合中是否存在关键字
        If mobjFtps.Exists(strFtpKey) Then
            Set mobjFtps(strFtpKey) = GetFtpInstance
        Else
            Call mobjFtps.Add(strFtpKey, GetFtpInstance)
        End If
    End If
End Function

'删除文件
Public Function FtpDeleteFile(ByRef ftpTag As TFtpConTag, ByVal strFtpFile As String, _
    Optional ByVal blnIsAutoDiscon As Boolean = True, Optional ByVal blnIsAutoHint As Boolean = True) As Long

    
    Dim objFtp As clsFtp
    Dim objFileSystem As New FileSystemObject
    
    Dim blnIsRetry As Boolean
    Dim lngResult As Long
    Dim strMessage As String
    Dim blnFailed As Boolean
    
    Dim strFtpClassPath As String
    Dim strFtpFileName As String
    
    FtpDeleteFile = frNormal
    
    If mobjFtps Is Nothing Then Call FtpParInit
    
    Set objFtp = GetFtpInstance(ftpTag, blnIsAutoDiscon)
     
    strFtpClassPath = objFileSystem.GetParentFolderName(Replace(ftpTag.VirtualPath & "/" & strFtpFile, "//", "/"))
    strFtpFileName = objFileSystem.GetFileName(strFtpFile)
    
    lngResult = objFtp.FuncDelFile(strFtpClassPath, strFtpFileName)
 
    blnIsRetry = IIf(lngResult <> 0, True, False)
    
    
    Do While blnIsRetry
            
        blnFailed = False
        
        objFtp.FuncFtpDisConnect
        lngResult = objFtp.FuncFtpConnect(ftpTag.Ip, ftpTag.User, ftpTag.pwd, mblnIsForceRead)
        
    
        'FuncFtpConnect 返回0表示失败
        If lngResult = 0 Then
            'ftp连接失败
            blnFailed = True
            strMessage = "FTP:" & ftpTag.Ip & " 连接失败,是否重试？" & vbCrLf & _
                            "FTP响应消息:" & objFtp.GetFtpMsg()
                            

        Else
            lngResult = objFtp.FuncDelFile(strFtpClassPath, strFtpFileName)
               
            If lngResult <> 0 Then
                '文件删除失败
                blnFailed = True
                strMessage = "从FTP:" & ftpTag.Ip & "删除文件 [" & strFtpClassPath & " , " & strFtpFileName & "] 失败,是否重试？" & vbCrLf & _
                            "FTP响应消息:" & objFtp.GetFtpMsg()
            End If
        End If
        
        If blnFailed And blnIsAutoHint Then
            lngResult = MsgboxEx(GetForegroundWindow, strMessage, vbAbortRetryIgnore, "提示")
            
            If lngResult = vbRetry Then
                blnIsRetry = True
            ElseIf lngResult = vbIgnore Then
                '忽略
                FtpDeleteFile = frIgnore
                Exit Function
            Else
                '终止
                FtpDeleteFile = frAbort
                Exit Function
            End If
        Else
            If blnFailed Then FtpDeleteFile = frIgnore
            blnIsRetry = False
        End If
    Loop
    
    '删除为空的目录()
    Call objFtp.FuncFtpDelDir(objFileSystem.GetParentFolderName(strFtpClassPath), _
                            objFileSystem.GetFileName(strFtpClassPath))
  
    If blnIsAutoDiscon Then
        '断开FTP连接
        Call objFtp.FuncFtpDisConnect
        
        Call FtpDiscon(ftpTag)
    End If
    
End Function

'下载文件
Public Function FtpDownloadFile(ByRef ftpTag As TFtpConTag, ByVal strFtpFile As String, ByVal strLocalFile As String, _
    Optional ByVal blnIsAutoDiscon As Boolean = True, Optional ByVal blnIsAutoHint As Boolean = True) As FtpResult

    FtpDownloadFile = FtpFileTransfer(ftpTag, strFtpFile, strLocalFile, 0, blnIsAutoDiscon, blnIsAutoHint)
    
End Function

'上传文件
Public Function FtpUploadFile(ByRef ftpTag As TFtpConTag, ByVal strFtpFile As String, ByVal strLocalFile As String, _
    Optional ByVal blnIsAutoDiscon As Boolean = True, Optional ByVal blnIsAutoHint As Boolean = True) As FtpResult
    
    FtpUploadFile = FtpFileTransfer(ftpTag, strFtpFile, strLocalFile, 1, blnIsAutoDiscon, blnIsAutoHint)
    
End Function


Public Sub FtpDiscon(ByRef ftpTag As TFtpConTag)
    Dim objFtp As clsFtp
    Dim strFtpKey As String
    
    strFtpKey = GetFtpKey(ftpTag)
    
    If mobjFtps.Exists(strFtpKey) Then
        Set objFtp = mobjFtps(strFtpKey)
    End If
    
    If objFtp Is Nothing Then Exit Sub
    
    Call objFtp.FuncFtpDisConnect
    
'    Set mobjFtps(ftpTag) = Nothing
'    Call mobjFtps.Remove(ftpTag)
End Sub

Public Sub FtpFree()
'释放资源
    Dim objFtp As clsFtp
    Dim strFtpTag As Variant
    
    If mobjFtps Is Nothing Then Exit Sub
    
    For Each strFtpTag In mobjFtps.Keys
        Set objFtp = mobjFtps.Item(strFtpTag)
        
        If Not objFtp Is Nothing Then
            Call objFtp.FuncFtpDisConnect
        End If
    Next
    
    Call mobjFtps.RemoveAll
    
    Set mobjFtps = Nothing
End Sub


Private Sub Class_Initialize()
    Call FtpParInit
End Sub

Private Sub Class_Terminate()
    Call FtpFree
End Sub


