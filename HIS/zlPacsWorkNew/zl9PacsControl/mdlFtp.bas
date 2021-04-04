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
'ftp��ز�����ʼ��
    If Not mobjFtps Is Nothing Then Exit Sub
    
    Set mobjFtps = New Dictionary
    
    mblnIsPass = IIf(Val(GetSetting("ZLSOFT", "����ģ��\Ftp", "���ñ�������", 0)) = 1, True, False)
    Call SaveSetting("ZLSOFT", "����ģ��\Ftp", "���ñ�������", IIf(mblnIsPass, 1, 0))
    
    mblnIsForceRead = IIf(Val(GetSetting("ZLSOFT", "����ģ��\Ftp", "����ǿ�ƶ�ȡ", 0)) = 1, True, False)
    Call SaveSetting("ZLSOFT", "����ģ��\Ftp", "����ǿ�ƶ�ȡ", IIf(mblnIsForceRead, 1, 0))
    
    mblnIsCompareSize = IIf(Val(GetSetting("ZLSOFT", "����ģ��\Ftp", "����FTP�ļ���С�Ա�", 0)) = 1, True, False)
    Call SaveSetting("ZLSOFT", "����ģ��\Ftp", "����FTP�ļ���С�Ա�", IIf(mblnIsCompareSize, 1, 0))
    
End Sub

Public Function FtpIsValid(ByRef ftpTag As TFtpConTag) As Boolean
'ftp �Ƿ���Ч
    Dim objFtp As clsFtp
    Dim lngResult As Long
    
    FtpIsValid = False
    
    Set objFtp = GetFtpInstance(ftpTag, False)
    
    lngResult = objFtp.FuncFtpConnect(ftpTag.Ip, ftpTag.User, ftpTag.pwd, mblnIsForceRead)
    If lngResult = 0 Then Exit Function
    
    FtpIsValid = True
End Function

'�����ļ�
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
        '�����ļ�
        lngResult = objFtp.FuncDownloadFile(strFtpClassPath, strLocalFile, strFtpFileName, mblnIsForceRead)
    Else
        If Trim(strFtpClassPath) <> "" Then Call objFtp.FuncFtpMkDir("/", strFtpClassPath)
        
        '�ϴ��ļ�
        'If lngResult = 0 Then
            lngResult = objFtp.FuncUploadFile(strFtpClassPath, strLocalFile, strFtpFileName)
        'End If
    End If
    
    blnIsRetry = False
    
    If lngResult = 0 Then
        'FuncDownloadFile����0��ʾ�ɹ�
        '�ļ���С���
        If mblnIsCompareSize Then
            lngDestFileSize = objFileSystem.GetFile(strLocalFile).Size
            lngFtpFileSize = objFtp.FuncFtpGetFileSize(strFtpClassPath, strFtpFileName)
            
            If lngFtpFileSize <> lngDestFileSize Then
                strFtpMsg = Replace(objFtp.GetFtpMsg(), Chr(0), "")
                
                strMessage = "�����ļ���С[" & lngDestFileSize & "]��FTP�ļ���С[" & lngFtpFileSize & "]��һ��" & vbCrLf & _
                             "�����ļ���" & strLocalFile & vbCrLf & _
                             "FTP�ļ���" & strFtpClassPath & "," & strFtpFileName & vbCrLf & _
                             IIf(Trim(strFtpMsg) <> "", "FTP��Ӧ��Ϣ:" & strFtpMsg & vbCrLf, "") & _
                             "�Ƿ���Ҫ����" & IIf(lngTransferWay = 0, "����?", "�ϴ�?")
                             
                lngResult = MsgboxEx(GetForegroundWindow, strMessage, vbAbortRetryIgnore, "��ʾ")
                objFtp.FuncFtpDisConnect
                
                If lngResult = vbRetry Then
                    blnIsRetry = True
                ElseIf lngResult = vbIgnore Then
                    '����
                    FtpFileTransfer = frIgnore
                    Exit Function
                Else
                    '��ֹ
                    FtpFileTransfer = frAbort
                    Exit Function
                End If
            Else
                '�ļ���С��ͬ
                objFtp.FuncFtpDisConnect
                Exit Function
            End If
        Else
            '�������ļ���С���
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
        
    
        'FuncFtpConnect ����0��ʾʧ��
        If lngResult = 0 Then
            'ftp����ʧ��
            blnFailed = True
            strFtpMsg = Replace(objFtp.GetFtpMsg(), Chr(0), "")
            
            strMessage = "FTP:" & ftpTag.Ip & " ����ʧ��,�Ƿ����ԣ�" & vbCrLf & _
                            IIf(Trim(strFtpMsg) <> "", "FTP��Ӧ��Ϣ:" & strFtpMsg, "")
                            

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
                '�ļ�����ʧ��
                blnFailed = True
                strFtpMsg = Replace(objFtp.GetFtpMsg(), Chr(0), "")
                
                strMessage = "��FTP:" & ftpTag.Ip & IIf(lngTransferWay = 0, " ����", " �ϴ�") & "�ļ� [" & ftpTag.VirtualPath & " , " & strFtpFile & "] ʧ��,�Ƿ����ԣ�" & vbCrLf & _
                            IIf(Trim(strFtpMsg) <> "", "FTP��Ӧ��Ϣ:" & strFtpMsg, "")
            Else
                '�ļ���С���
                If mblnIsCompareSize Then
                    lngDestFileSize = objFileSystem.GetFile(strLocalFile).Size
                    lngFtpFileSize = objFtp.FuncFtpGetFileSize(strFtpClassPath, strFtpFileName)
                    
                    If lngFtpFileSize <> lngDestFileSize Then
                        blnFailed = True
                        strFtpMsg = Replace(objFtp.GetFtpMsg(), Chr(0), "")
                        
                        strMessage = "�����ļ���С[" & lngDestFileSize & "]��FTP�ļ���С[" & lngFtpFileSize & "]��һ��" & vbCrLf & _
                                     "�����ļ���" & strLocalFile & vbCrLf & _
                                     "FTP�ļ���" & strFtpClassPath & ":" & strFtpFileName & vbCrLf & _
                                     IIf(Trim(strFtpMsg) <> "", "FTP��Ӧ��Ϣ:" & strFtpMsg & vbCrLf, "") & _
                                     "�Ƿ���Ҫ����" & IIf(lngTransferWay = 0, "����?", "�ϴ�?")
                    End If
                End If
            End If
            
        End If
        
        If blnFailed And blnIsAutoHint Then
            lngResult = MsgboxEx(GetForegroundWindow, strMessage, vbAbortRetryIgnore, "��ʾ")
            
            Call objFtp.FuncFtpDisConnect
            
            If lngResult = vbRetry Then
                blnIsRetry = True
            ElseIf lngResult = vbIgnore Then
                '����
                FtpFileTransfer = frIgnore
                Exit Function
            Else
                '��ֹ
                FtpFileTransfer = frAbort
                Exit Function
            End If
        Else
            If blnFailed Then FtpFileTransfer = frIgnore
            blnIsRetry = False
        End If
    Loop
        
  
    If blnIsAutoDiscon Then
        '�Ͽ�FTP����
        Call objFtp.FuncFtpDisConnect
        
        Call FtpDiscon(ftpTag)
    End If
    
End Function

Private Function GetFtpKey(ByRef ftpTag As TFtpConTag) As String
'��ȡftpkey��Ϣ
    GetFtpKey = ""
    If Trim(ftpTag.Ip) = "" Then Exit Function
    
    GetFtpKey = "ftp://" & ftpTag.User & ":" & ftpTag.pwd & "@" & ftpTag.Ip & ":" & ftpTag.Port '& "/" & ftpTag.VirtualPath
End Function


Private Function GetFtpInstance(ByRef ftpTag As TFtpConTag, Optional ByVal blnIsAutoDiscon As Boolean = True) As clsFtp
'��ȡftp����ʵ��
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
        '�жϼ������Ƿ���ڹؼ���
        If mobjFtps.Exists(strFtpKey) Then
            Set mobjFtps(strFtpKey) = GetFtpInstance
        Else
            Call mobjFtps.Add(strFtpKey, GetFtpInstance)
        End If
    End If
End Function

'ɾ���ļ�
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
        
    
        'FuncFtpConnect ����0��ʾʧ��
        If lngResult = 0 Then
            'ftp����ʧ��
            blnFailed = True
            strMessage = "FTP:" & ftpTag.Ip & " ����ʧ��,�Ƿ����ԣ�" & vbCrLf & _
                            "FTP��Ӧ��Ϣ:" & objFtp.GetFtpMsg()
                            

        Else
            lngResult = objFtp.FuncDelFile(strFtpClassPath, strFtpFileName)
               
            If lngResult <> 0 Then
                '�ļ�ɾ��ʧ��
                blnFailed = True
                strMessage = "��FTP:" & ftpTag.Ip & "ɾ���ļ� [" & strFtpClassPath & " , " & strFtpFileName & "] ʧ��,�Ƿ����ԣ�" & vbCrLf & _
                            "FTP��Ӧ��Ϣ:" & objFtp.GetFtpMsg()
            End If
        End If
        
        If blnFailed And blnIsAutoHint Then
            lngResult = MsgboxEx(GetForegroundWindow, strMessage, vbAbortRetryIgnore, "��ʾ")
            
            If lngResult = vbRetry Then
                blnIsRetry = True
            ElseIf lngResult = vbIgnore Then
                '����
                FtpDeleteFile = frIgnore
                Exit Function
            Else
                '��ֹ
                FtpDeleteFile = frAbort
                Exit Function
            End If
        Else
            If blnFailed Then FtpDeleteFile = frIgnore
            blnIsRetry = False
        End If
    Loop
    
    'ɾ��Ϊ�յ�Ŀ¼()
    Call objFtp.FuncFtpDelDir(objFileSystem.GetParentFolderName(strFtpClassPath), _
                            objFileSystem.GetFileName(strFtpClassPath))
  
    If blnIsAutoDiscon Then
        '�Ͽ�FTP����
        Call objFtp.FuncFtpDisConnect
        
        Call FtpDiscon(ftpTag)
    End If
    
End Function

'�����ļ�
Public Function FtpDownloadFile(ByRef ftpTag As TFtpConTag, ByVal strFtpFile As String, ByVal strLocalFile As String, _
    Optional ByVal blnIsAutoDiscon As Boolean = True, Optional ByVal blnIsAutoHint As Boolean = True) As FtpResult

    FtpDownloadFile = FtpFileTransfer(ftpTag, strFtpFile, strLocalFile, 0, blnIsAutoDiscon, blnIsAutoHint)
    
End Function

'�ϴ��ļ�
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
'�ͷ���Դ
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


