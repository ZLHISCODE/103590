VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3150
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3150
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer timerProcess 
      Interval        =   500
      Left            =   1560
      Top             =   840
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Type TImgTransInfo
    Key As String
    FileName As String
    FilePath As String
    
    FtpIp As String
    FtpPort As String
    FtpUser As String
    FtpPwd As String
    FtpVirturalPath As String
    FtpShareDir As String
    FtpShareUser As String
    FtpSharePwd As String
    FtpFile As String
    
    JpgConvert As Boolean
    ImgFormat  As Long
    ImgCommand As String
End Type

Private Declare Function SetFileAttributes Lib "kernel32" Alias "SetFileAttributesA" (ByVal lpFileName As String, ByVal dwFileAttributes As Long) As Long

Private mstrCmdPath As String
Private mstrTimeoutFile As String

Private mblnIsPass As Boolean
Private mblnIsForceRead As Boolean
Private mblnIsCompareSize As Boolean
Private mlngMaxRedoCount As Long

Private mstrFailedDir As String


Public Sub Start(ByVal strCmdPath As String, Optional ByVal blnIsSingleProcess As Boolean = False)
    Dim objFile As File
    Dim objFiles As FileSystemObject
    Dim strFile As String
    Dim lngCount As Long
    
    mlngMaxRedoCount = 3
    mstrFailedDir = Replace(gstrCmdPath & "\Failed\", "\\", "\")
    mstrTimeoutFile = Replace(gstrCmdPath & "\TimeOut.dat", "\\", "\")
    
    mstrCmdPath = gstrCmdPath
    
    If blnIsSingleProcess Then
        strFile = Replace(gstrCmdPath & "\" & Split(strCmdPath & ",", ",")(1), "\\", "\")
        
        If FileExists(strFile) Then
            Set objFiles = New FileSystemObject
            Set objFile = objFiles.GetFile(strFile)
            
            lngCount = 0
            While FileExists(strFile)
            
                If lngCount >= mlngMaxRedoCount Then
                    Call SendFileToFailed(objFile)
                    
                    Exit Sub
                End If
                
                If lngCount >= 1 Then
                    gblnIsFailed = True
                    
                    DoEvents
                End If
                
                Call ExecuteCmd(objFile)
                
                lngCount = lngCount + 1
                DoEvents
            Wend
        Else
            MsgBox "未找到待处理文件： " & strCmdPath, vbOKOnly, "提示"
        End If
        
    Else
        If DirExists(mstrFailedDir) = False Then Call MkLocalDir(mstrFailedDir)
        
        Call FtpParInit
        
        Me.Hide
    End If
    
    
End Sub

Private Sub SendFileToFailed(objFile As File)
On Error GoTo errHandle
    Call RemoveFile(mstrFailedDir & objFile.Name)
    objFile.Move mstrFailedDir & objFile.Name
Exit Sub
errHandle:
    
End Sub

Private Sub MkLocalDir(ByVal strDir As String)
'------------------------------------------------
'功能：创建本地目录
'参数： strDir－－本地目录
'返回：无
'------------------------------------------------
    Dim objFile As New Scripting.FileSystemObject
    Dim aNestDirs() As String, i As Integer
    Dim strPath As String
    On Error Resume Next
    
    '读取全部需要创建的目录信息
    ReDim Preserve aNestDirs(0)
    aNestDirs(0) = strDir
    
    strPath = objFile.GetParentFolderName(strDir)
    Do While Len(strPath) > 0
        ReDim Preserve aNestDirs(UBound(aNestDirs) + 1)
        aNestDirs(UBound(aNestDirs)) = strPath
        strPath = objFile.GetParentFolderName(strPath)
    Loop
    '创建全部目录
    For i = UBound(aNestDirs) To 0 Step -1
        MkDir aNestDirs(i)
    Next
End Sub

Private Sub timerProcess_Timer()
On Error GoTo errHandle
    Dim objFileSys As New FileSystemObject
    Dim objFolder As Folder
    
    If Trim(mstrCmdPath) = "" Then Exit Sub
     
    If DirExists(mstrCmdPath) = False Then Call MkLocalDir(mstrCmdPath)
    
    Set objFolder = objFileSys.GetFolder(mstrCmdPath)
    
    If objFolder.Files.Count <= 0 Then
        Call UpdateTimeout
    Else
        Call DispatchCmd(objFolder.Files)
    End If
    
    gblnWorking = False
    
    Set objFolder = Nothing
    Set objFileSys = Nothing
Exit Sub
errHandle:
    gblnWorking = False
    
    Set objFolder = Nothing
    Set objFileSys = Nothing
    
    Call ShowTrayMsg(Err.Description, 3)
End Sub

Private Sub DispatchCmd(objFiles As Files)
'分配命令
    Dim lngStart As Long
    Dim objFile As File
    
    
    For Each objFile In objFiles
        If UCase(objFile.Name) <> "TIMEOUT.DAT" And UCase(objFile.Name) <> "DESKTOP.INI" Then
            If (objFile.Attributes And Hidden) = Hidden Then
            
                gblnWorking = True
                
                lngStart = GetTickCount
                Call ExecuteCmd(objFile)
                
                If (GetTickCount - lngStart) > 1000 Then Call UpdateTimeout
                 
                
                DoEvents
            End If
        End If
    Next
    
    Call UpdateTimeout
End Sub

Private Sub ExecuteCmd(objCmdFile As File)
'执行单个命令
    Dim imgTransInfo As TImgTransInfo
    Dim ftpTag As TFtpConTag
    
On Error GoTo errHandle
    imgTransInfo = ReadTransInfo(objCmdFile.Path)
    
    If imgTransInfo.Key = "" _
        Or imgTransInfo.FileName = "" _
        Or imgTransInfo.FilePath = "" _
        Or imgTransInfo.FtpIp = "" _
        Or imgTransInfo.FtpUser = "" _
        Or imgTransInfo.FtpFile = "" _
        Or imgTransInfo.ImgCommand = "" Then
        Call ReturnErrorInfo(objCmdFile, "检测到传输命令缺失关键信息，请检查.")
        Exit Sub
    End If
    
    If gblnSingle Then
        Call DoTrans(imgTransInfo, objCmdFile)
        Exit Sub
    End If
          
    If imgTransInfo.ImgFormat = 3 Or imgTransInfo.ImgFormat = 4 Then
        If imgTransInfo.ImgCommand = 2 Then
            '上传音视频文件
            '如果是音频或者视频文件，则判断文件大小，如果超过10M则启动独立进程进行传输
            If Abs(FileSystem.FileLen(imgTransInfo.FilePath & imgTransInfo.FileName)) > 10485760 Then '10 * 1024 * 1024 10M
                SetFileAttributes objCmdFile.Path, 0
                Shell App.Path & "\" & App.EXEName & ".Exe " & "[BF]-" & gstrCmdPath & "," & objCmdFile.Name
                
                Exit Sub
            End If
        Else
            '下载音视频文件
            ftpTag = FtpTagInstance(imgTransInfo.FtpIp, _
                                    imgTransInfo.FtpUser, _
                                    imgTransInfo.FtpPwd, _
                                    imgTransInfo.FtpVirturalPath, _
                                    imgTransInfo.FtpPort)
            
            If Abs(GetFtpFileSize(ftpTag, imgTransInfo.FtpFile)) > 10485760 Then
                SetFileAttributes objCmdFile.Path, 0
                Shell App.Path & "\" & App.EXEName & ".Exe " & "[BF]-" & gstrCmdPath & "," & objCmdFile.Name
                
                Exit Sub
            End If
        End If
    End If
    
    Call DoTrans(imgTransInfo, objCmdFile)
Exit Sub
errHandle:
    Call ReturnErrorInfo(objCmdFile, Err.Description)
End Sub

Private Function ConvertTrans(imgTranInfo As TImgTransInfo, ftpTag As TFtpConTag) As Boolean
'TODO...
    ConvertTrans = True
End Function

Private Sub DoTrans(imgTranInfo As TImgTransInfo, objCmdFile As File)
'执行下载
On Error GoTo errHandle
    Dim ftpTag As TFtpConTag
    Dim blnResult As Boolean
    Dim strLocalFile As String
    Dim strError As String
    
    ftpTag = FtpTagInstance(imgTranInfo.FtpIp, _
                            imgTranInfo.FtpUser, _
                            imgTranInfo.FtpPwd, _
                            imgTranInfo.FtpVirturalPath, _
                            imgTranInfo.FtpPort)
    
    strLocalFile = Replace(imgTranInfo.FilePath & "\" & imgTranInfo.FileName, "\\", "\")
    
    If DirExists(imgTranInfo.FilePath) = False Then
        Call MkLocalDir(imgTranInfo.FilePath)
    End If
    
    '开始传输文件
    If Val(imgTranInfo.ImgCommand) = 2 Then
'        Call ResetTrayIcon(tsUp)   '该操作会占用处理时间，降低执行效率
        
        '2表示上传
        blnResult = FtpFileTransfer(ftpTag, imgTranInfo.FtpFile, strLocalFile, 1, strError)
        
        If blnResult And imgTranInfo.JpgConvert Then
            '将dicom图像转换为jpg，然后在进行上传
            blnResult = ConvertTrans(imgTranInfo, ftpTag)
        End If
    Else
'        Call ResetTrayIcon(tsDown) '该操作会占用处理时间，降低执行效率
        
        '1表示下载
        blnResult = FtpFileTransfer(ftpTag, imgTranInfo.FtpFile, strLocalFile, 0, strError)
        
        '判断是否存在jpg文件，如果存在则同步下载
        If blnResult And InStr(UCase(imgTranInfo.FtpFile), ".JPG") <= 0 Then
            If FtpFileExists(ftpTag, imgTranInfo.FtpFile & ".jpg") Then
                imgTranInfo.FtpFile = imgTranInfo.FtpFile & ".jpg"
                
                 strLocalFile = Replace(imgTranInfo.FilePath & "\" & imgTranInfo.FtpFile, "\\", "\")
                 
                Call FtpFileTransfer(ftpTag, imgTranInfo.FtpFile, strLocalFile, 0)
            End If
        End If
    End If
    
    If blnResult = False Then
        Call ReturnErrorInfo(objCmdFile, strError)
        
        Call ShowTrayMsg(strError, 3)
    Else
        Call ReturnSuccessInfo(objCmdFile)
    End If
Exit Sub
errHandle:
    Call ReturnErrorInfo(objCmdFile, Err.Description)
    
    Call ShowTrayMsg(Err.Description, 3)
End Sub
 
Private Sub ReturnSuccessInfo(objCmdFile As File)
    Dim objIni As clsIniFile
    
On Error GoTo errHandle
    Set objIni = New clsIniFile
    Call objIni.SetIniFile(objCmdFile.Path)
    
    objIni.WriteValue "OTHERINFO", "ENDTIME", Now
    objIni.WriteValue "OTHERINFO", "REDO", -1
    objIni.WriteValue "OTHERINFO", "ERRORINFO", ""
    
    Set objIni = Nothing
    
    '删除命令文件
     Call RemoveCmdFile(objCmdFile)
Exit Sub
errHandle:
    Set objIni = Nothing
    Call RemoveCmdFile(objCmdFile)
End Sub

Private Sub RemoveCmdFile(objCmdFile As File)
'强制移除命令文件
On Error GoTo errHandle
    Call objCmdFile.Delete(True)
Exit Sub
errHandle:
    
End Sub

Private Sub RemoveFile(ByVal strFile As String)
    Dim objFileSys As FileSystemObject
    Dim objFile As File
On Error GoTo errHandle
    Set objFileSys = New FileSystemObject
    Set objFile = objFileSys.GetFile(strFile)

    objFile.Delete True
    
    Set objFile = Nothing
    Set objFileSys = Nothing
Exit Sub
errHandle:

End Sub

Private Sub ReturnErrorInfo(objCmdFile As File, ByVal strError As String)
'返回错误信息，如果重试次数超过指定次数，则移入到失败目录
    Dim lngRedoCount As Long
On Error GoTo errHandle
    Debug.Print strError
    lngRedoCount = WriteErrorToFile(objCmdFile.Path, strError)
    Debug.Print strError
   '将多次失败的文件移入到失败目录
    If lngRedoCount >= mlngMaxRedoCount Then
        Call RemoveFile(mstrFailedDir & objCmdFile.Name)
 
'        MoveFile objCmdFile.Path, mstrFailedDir & objCmdFile.Name
        objCmdFile.Move mstrFailedDir & objCmdFile.Name
        Exit Sub
    End If
      
Exit Sub
errHandle:
    Debug.Print Err.Description
End Sub

Private Function WriteErrorToFile(ByVal strFile As String, ByVal strError As String) As Long
    Dim objIni As clsIniFile
    Dim lngRedoCount As Long
    
On Error GoTo errHandle

    WriteErrorToFile = 0
    
    Set objIni = New clsIniFile
    Call objIni.SetIniFile(strFile)
    
    lngRedoCount = Val(objIni.ReadValue("OTHERINFO", "REDO", "0"))

    objIni.WriteValue "OTHERINFO", "ENDTIME", Now
    objIni.WriteValue "OTHERINFO", "REDO", lngRedoCount + 1
    objIni.WriteValue "OTHERINFO", "ERRORINFO", strError
    
    Set objIni = Nothing
    
    WriteErrorToFile = lngRedoCount
Exit Function
errHandle:
    Set objIni = Nothing
End Function


Public Sub FtpParInit()
'ftp相关参数初始化
    
    mblnIsPass = IIf(Val(GetSetting("ZLSOFT", "公共模块\Ftp", "启用被动传输", 0)) = 1, True, False)
    Call SaveSetting("ZLSOFT", "公共模块\Ftp", "启用被动传输", IIf(mblnIsPass, 1, 0))
    
    mblnIsForceRead = IIf(Val(GetSetting("ZLSOFT", "公共模块\Ftp", "启用强制读取", 0)) = 1, True, False)
    Call SaveSetting("ZLSOFT", "公共模块\Ftp", "启用强制读取", IIf(mblnIsForceRead, 1, 0))
    
    mblnIsCompareSize = IIf(Val(GetSetting("ZLSOFT", "公共模块\Ftp", "启用FTP文件大小对比", 0)) = 1, True, False)
    Call SaveSetting("ZLSOFT", "公共模块\Ftp", "启用FTP文件大小对比", IIf(mblnIsCompareSize, 1, 0))
    
    mlngMaxRedoCount = Val(GetSetting("ZLSOFT", "公共模块\Ftp", "后台重试次数", 3))
    If mlngMaxRedoCount <= 0 Then mlngMaxRedoCount = 3
    
    Call SaveSetting("ZLSOFT", "公共模块\Ftp", "后台重试次数", mlngMaxRedoCount)
End Sub

Private Function FtpFileExists(ByRef ftpTag As TFtpConTag, ByVal strFtpFile As String) As Boolean
'判断ftp文件是否存在
    Dim objFtp As New clsFtp
    Dim strFtpClassPath As String
    Dim strFtpFileName As String
    Dim lngResult As Long
    Dim objFileSystem As New FileSystemObject
    
    
    FtpFileExists = False
    
    lngResult = objFtp.FuncFtpConnect(ftpTag.Ip, ftpTag.User, ftpTag.pwd, mblnIsForceRead)
    If lngResult = 0 Then
        Exit Function
    End If
    
    strFtpClassPath = objFileSystem.GetParentFolderName(Replace(ftpTag.VirtualPath & "/" & strFtpFile, "//", "/"))
    strFtpFileName = objFileSystem.GetFileName(strFtpFile)
    
    FtpFileExists = objFtp.FuncFtpFileExists(strFtpClassPath, strFtpFileName)
    
    objFtp.FuncFtpDisConnect
    
    Set objFtp = Nothing
End Function


Private Function GetFtpFileSize(ByRef ftpTag As TFtpConTag, ByVal strFtpFile As String) As Long
    Dim objFileSystem As New FileSystemObject
    Dim strFtpClassPath As String
    Dim objFtp As clsFtp
    
On Error GoTo errHandle
    GetFtpFileSize = 0
    
    strFtpClassPath = objFileSystem.GetParentFolderName(Replace(ftpTag.VirtualPath & "/" & strFtpFile, "//", "/"))
    
    Set objFtp = New clsFtp
    
    Call objFtp.FuncFtpConnect(ftpTag.Ip, ftpTag.User, ftpTag.pwd, mblnIsForceRead)
    
    GetFtpFileSize = objFtp.FuncFtpGetFileSize(strFtpClassPath, strFtpFile)
    
    Call objFtp.FuncFtpDisConnect
Exit Function
errHandle:
    GetFtpFileSize = 0
End Function

Private Function FtpFileTransfer(ByRef ftpTag As TFtpConTag, _
    ByVal strFtpFile As String, ByVal strLocalFile As String, _
    ByVal lngTransferWay As Long, Optional ByRef strError As String) As Boolean
    
    Dim blnIsAutoDiscon As Boolean
    Dim objFtp As clsFtp
    Dim objFileSystem As New FileSystemObject
     
    Dim lngResult As Long
    Dim lngDestFileSize As Long
    Dim lngFtpFileSize As Long
    Dim blnFailed As Boolean
    
    Dim strFtpClassPath As String
    Dim strFtpFileName As String
    
    Dim strFtpMsg As String
    
On Error GoTo errHandle

    strError = ""
    
    blnIsAutoDiscon = True
    FtpFileTransfer = True
    
    Set objFtp = New clsFtp
   
    lngResult = objFtp.FuncFtpConnect(ftpTag.Ip, ftpTag.User, ftpTag.pwd, mblnIsForceRead)
    
    'FuncFtpConnect 返回0表示失败
    If lngResult = 0 Then
        'ftp连接失败
        FtpFileTransfer = False
        
        strFtpMsg = Replace(objFtp.GetFtpMsg(), Chr(0), "")
        
        strError = "FTP:" & ftpTag.Ip & " 连接失败,请检查FTP服务是否正常." & vbCrLf & _
                        IIf(Trim(strFtpMsg) <> "", "FTP响应消息:" & strFtpMsg, "")

    Else
        strFtpClassPath = objFileSystem.GetParentFolderName(Replace(ftpTag.VirtualPath & "/" & strFtpFile, "//", "/"))
        strFtpFileName = objFileSystem.GetFileName(strFtpFile)
    
        If lngTransferWay = 0 Then
            lngResult = objFtp.FuncDownloadFile(strFtpClassPath, strLocalFile, strFtpFileName, mblnIsForceRead)
        Else
            If Trim(strFtpClassPath) <> "" Then Call objFtp.FuncFtpMkDir("/", strFtpClassPath)
            
            lngResult = objFtp.FuncUploadFile(strFtpClassPath, strLocalFile, strFtpFileName)
        End If
        
        If lngResult <> 0 Then
            '文件传输失败
            FtpFileTransfer = False
            
            strFtpMsg = Replace(objFtp.GetFtpMsg(), Chr(0), "")
            
            strError = "从FTP:" & ftpTag.Ip & IIf(lngTransferWay = 0, " 下载", " 上传") & "文件 [" & ftpTag.VirtualPath & " , " & strFtpFile & "] 失败." & vbCrLf & _
                        IIf(Trim(strFtpMsg) <> "", "FTP响应消息:" & strFtpMsg, "")
        Else
            '文件大小检查
            If mblnIsCompareSize Then
                lngDestFileSize = objFileSystem.GetFile(strLocalFile).Size
                lngFtpFileSize = objFtp.FuncFtpGetFileSize(strFtpClassPath, strFtpFileName)
                
                If lngFtpFileSize <> lngDestFileSize Then
                    FtpFileTransfer = False
                    
                    strFtpMsg = Replace(objFtp.GetFtpMsg(), Chr(0), "")
                    
                    strError = "本地文件大小[" & lngDestFileSize & "]与FTP文件大小[" & lngFtpFileSize & "]不一致" & vbCrLf & _
                                 "本地文件：" & strLocalFile & vbCrLf & _
                                 "FTP文件：" & strFtpClassPath & "," & strFtpFileName & vbCrLf & _
                                 IIf(Trim(strFtpMsg) <> "", "FTP响应消息:" & strFtpMsg & vbCrLf, "")
                End If
            End If
        End If
    End If
    
    '断开FTP连接
    Call objFtp.FuncFtpDisConnect
    
    Set objFtp = Nothing
    Set objFileSystem = Nothing
Exit Function
errHandle:
    FtpFileTransfer = False
  Resume
    strError = Err.Description
    
    Set objFtp = Nothing
    Set objFileSystem = Nothing
End Function

Private Function ReadTransInfo(ByVal strFile As String) As TImgTransInfo
    Dim objIni As New clsIniFile
    Dim imgTransInfo As TImgTransInfo
    
On Error GoTo errHandle
    Call objIni.SetIniFile(strFile)
    
    imgTransInfo.Key = objIni.ReadValue("BASEINFO", "KEY", "")
    imgTransInfo.FileName = objIni.ReadValue("BASEINFO", "FILENAME", "")
    imgTransInfo.FilePath = objIni.ReadValue("BASEINFO", "FILEPATH", "")
    imgTransInfo.ImgFormat = Val(objIni.ReadValue("BASEINFO", "FORMAT", ""))
    
    imgTransInfo.FtpIp = objIni.ReadValue("FTPINFO", "FTPIP", "")
    imgTransInfo.FtpPort = objIni.ReadValue("FTPINFO", "FTPPORT", "")
    imgTransInfo.FtpUser = objIni.ReadValue("FTPINFO", "FTPUSER", "")
    imgTransInfo.FtpPwd = Trim(objIni.ReadValue("FTPINFO", "FTPPWD", ""))
    imgTransInfo.FtpVirturalPath = objIni.ReadValue("FTPINFO", "FTPVIRTUALPATH", "")
    imgTransInfo.FtpShareDir = objIni.ReadValue("FTPINFO", "FTPSHDIR", "")
    imgTransInfo.FtpShareUser = objIni.ReadValue("FTPINFO", "FTPSHUSER", "")
    imgTransInfo.FtpSharePwd = objIni.ReadValue("FTPINFO", "FTPSHPWD", "")
    imgTransInfo.FtpFile = objIni.ReadValue("FTPINFO", "FTPFILE", "")
    
    imgTransInfo.JpgConvert = CBool(objIni.ReadValue("OTHERINFO", "JPGCONVERT", "false"))
    imgTransInfo.ImgCommand = objIni.ReadValue("OTHERINFO", "IMGCOMMAND", "")
    
    
    ReadTransInfo = imgTransInfo
    
    Set objIni = Nothing
Exit Function
errHandle:
    Set objIni = Nothing
End Function



Private Sub UpdateTimeout()
'更新超时时间
    Dim objIni As New clsIniFile
On Error GoTo errHandle

    Call objIni.SetIniFile(mstrTimeoutFile)
    Call objIni.WriteValue("TIMEOUT", "value", Now)
    
    Set objIni = Nothing
Exit Sub
errHandle:
    Set objIni = Nothing
End Sub
