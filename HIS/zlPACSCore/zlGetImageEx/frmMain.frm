VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "中联图像下载"
   ClientHeight    =   1500
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3735
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1500
   ScaleWidth      =   3735
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.Timer tmMsg 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   120
      Top             =   1080
   End
   Begin VB.Label Label1 
      Caption         =   "正在下载图像，请勿关闭。                                    图像下载完成后会自动关闭。。。。。。"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   3495
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Sub OutputDebugString Lib "kernel32" Alias "OutputDebugStringA" (ByVal lpOutputString As String)

'共享文件夹
Private Declare Function WNetAddConnection2 Lib "mpr.dll" Alias "WNetAddConnection2A" (lpNetResource As NETRESOURCE, ByVal lpPassword As String, ByVal lpUserName As String, ByVal dwFlags As Long) As Long
Private Type NETRESOURCE ' 网络资源
    dwScope As Long
    dwType As Long
    dwDisplayType As Long
    dwUsage As Long
    lpLocalName As String
    lpRemoteName As String
    lpComment As String
    lpProvider As String
End Type
Private Const RESOURCETYPE_ANY = &H0

Public Enum TMediaType
    imgTag = 0   '图像标记
    MULFRAMETAG = 1 '多侦图
    VIDEOTAG = 2 '视频标记
    AUDIOTAG = 3 '音频标记
End Enum

Private Const G_STR_HINT_TITLE As String = "提示"

'Private WithEvents mobjIcon As clsTaskIcon  '托盘类

Private curMsg As clsImgInfo

'日志相关的参数，从注册表“”中读取日志参数，如果日志路径为空，则使用exe相同目录下的“GetImgLog”作为日志路径
Public mblnLogEnable As Boolean     '是否启用日志
Public mstrLogPath As String        '记录日志的路径,
Public mlngLogLevel As String       '记录日志的级别，分成1,2两级。1级只记录消息级别的日志；2级记录每一次下载的日志

Private mftpConnect As clsFtp        '定义一个常连接的FTP类
Private mftpConnectBak As clsFtp
Private mblnIsUpload As Boolean
Private mlngThreadID As Long

Private mConnectedSharedDir() As String  '记录已经连接过的共享目录
Private mlngRetriesn As Long        '上传或下载后重试次数
Private mblnOpenDebug As Boolean
Public mobjDataQueue As clsDataQueue

Public Event OnComPlete(ByVal curMsg As Object)
Public Event OnState(ByVal blnLoadFinish As Boolean, ByVal blnUpLoad As Boolean, ByVal lngThreadID As Long)

Public Sub DoComPlete(ByVal curMsg As Object)
BUGEX "DoComPlete 0"
    RaiseEvent OnComPlete(curMsg)
BUGEX "DoComPlete 1"
End Sub

Public Sub DoState(ByVal blnLoadFinish As Boolean, ByVal blnUpLoad As Boolean, ByVal lngThreadID As Long)
BUGEX "DoState 0"
    RaiseEvent OnState(blnLoadFinish, blnUpLoad, lngThreadID)
BUGEX "DoState 1"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo err
    '断开FTP连接
    Call mftpConnect.FuncFtpDisConnect
    Call mftpConnectBak.FuncFtpDisConnect
    
    Set mobjDataQueue = Nothing
    Set curMsg = Nothing
    
    Exit Sub
err:
    
End Sub

Public Sub zlInitModule(ByVal blnIsUpload As Boolean, ByVal lngThreadID As Long)
'------------------------------------------------
'功能：主程序，负责启动图像下载程序
'参数：
'返回：无
'-----------------------------------------------
    Dim strRegPath As String
    
    Set mftpConnect = New clsFtp
    Set mftpConnectBak = New clsFtp
    
    mblnIsUpload = blnIsUpload
    mlngThreadID = lngThreadID
    
    '如果本程序已经启动过一次，则不再启动
    If App.PrevInstance Then
        Exit Sub
    End If
    
    On Error Resume Next
    
    Set mobjDataQueue = New clsDataQueue
    
    '从注册表读取日志参数
    strRegPath = "公共模块\zlGetImage"
    mblnLogEnable = CBool(GetSetting("ZLSOFT", strRegPath, "记录日志", "True"))
    mstrLogPath = GetSetting("ZLSOFT", strRegPath, "日志路径", App.Path & "\GetImgLog")
    mlngLogLevel = Val(GetSetting("ZLSOFT", strRegPath, "日志级别", 1))
    mblnOpenDebug = CBool(GetSetting("ZLSOFT", strRegPath, "IsOpenDebug", "True"))
    mlngRetriesn = Val(GetSetting("ZLSOFT", strRegPath, "重试次数", 3))
BUGEX "mblnLogEnable=" & mblnLogEnable & "  mlngLogLevel=" & mlngLogLevel & "   mstrLogPath=" & mstrLogPath
    SaveSetting "ZLSOFT", strRegPath, "记录日志", mblnLogEnable
    SaveSetting "ZLSOFT", strRegPath, "日志路径", mstrLogPath
    SaveSetting "ZLSOFT", strRegPath, "日志级别", mlngLogLevel
    SaveSetting "ZLSOFT", strRegPath, "IsOpenDebug", mblnOpenDebug
    SaveSetting "ZLSOFT", strRegPath, "重试次数", mlngRetriesn
    
    '如果启动了日志，检查日志路径是否存在
    If mblnLogEnable = True Then
        '如果没有设置日志路径，则使用默认路径
        If mstrLogPath = "" Then
            mstrLogPath = App.Path & "\GetImgLog"
        End If
        
        '如果日志路径不存在，则创建
        If Dir(mstrLogPath, vbDirectory) = "" Then
            '默认路径不存在，创建这个目录
            If Dir(mstrLogPath, vbDirectory) = "" Then
                Call MkLocalDir(mstrLogPath)
            End If
        End If
    End If

    '初始化数据
    ReDim mConnectedSharedDir(0) As String
End Sub

Public Function funConnectAndSaveSheardDir(strSharedDir As String, strUser As String, strPswd As String) As Boolean
'------------------------------------------------
'功能：连接共享目录，使用用户名和秘密登录服务器
'参数： strSharedDir -- 需要连接的共享目录名称
'返回：True-- 成功，False -- 失败
'-----------------------------------------------
    Dim i As Integer
    
    funConnectAndSaveSheardDir = False

    On Error GoTo err
    
    '判断共享目录是否已经连接，如果没有连接，则进行连接
    For i = 1 To UBound(mConnectedSharedDir)
        If mConnectedSharedDir(i) = strSharedDir Then
            funConnectAndSaveSheardDir = True
            Exit Function
        End If
    Next i
    
    '连接共享目录
    If strSharedDir <> "" Then
        If funConnectSharedDir(strSharedDir, strUser, strPswd) = True Then
            '连接成功，记录成功的连接串
            ReDim Preserve mConnectedSharedDir(UBound(mConnectedSharedDir) + 1) As String
            mConnectedSharedDir(UBound(mConnectedSharedDir)) = strSharedDir
            funConnectAndSaveSheardDir = True
        End If
    End If
    
    Exit Function
err:
    '暂不处理
End Function

Public Function funConnectSharedDir(strShareRemoteDir As String, strUserName As String, _
    strPassWord As String) As Boolean
'------------------------------------------------
'功能：创建网络资源
'参数： strShareRemoteDir -- 共享目录
'       strUserName -- 共享目录用户名
'       strPassWord -- 共享目录密码
'返回：True--连接成功； False -- 连接失败
'------------------------------------------------
    Dim NetR As NETRESOURCE
    Dim lngResult As Long
    
    On Error GoTo err
    
    funConnectSharedDir = False
    
    NetR.dwType = RESOURCETYPE_ANY
    NetR.lpLocalName = vbNullString
    NetR.lpRemoteName = strShareRemoteDir
    NetR.lpProvider = vbNullString
    lngResult = WNetAddConnection2(NetR, strPassWord, strUserName, 0)
    
   ' If lngResult = 0 Then
        funConnectSharedDir = True
   ' End If
    Exit Function
err:
    '暂不处理
End Function

Public Function funDownLoadSharedDirSingle(strSourceDir As String, strDestDir As String) As Boolean
'------------------------------------------------
'功能：通过共享目录的方式下载一个目录下的一个图像，strSourceDir目录所在的计算机登录，通过另外的过程实现
'参数： strSourceDir -- 需要复制文件的源目录，即远程服务器中的共享目录
'       strDestDir  --  文件复制的目的地，即计算机的本机目录
'返回：True-- 成功，False -- 失败
    Dim objFSO As New Scripting.FileSystemObject
    
    funDownLoadSharedDirSingle = False
    
BUGEX "funDownLoadSharedDirSingle S"

    If Dir(strDestDir, vbDirectory) = "" Then
        Call MkLocalDir(objFSO.GetParentFolderName(strDestDir))
    End If
    
BUGEX "strSourceDir=" & strSourceDir & "  strDestDir=" & strDestDir
    
    Call objFSO.CopyFile(strSourceDir, strDestDir, False)
    
BUGEX "funDownLoadSharedDirSingle=true"

    funDownLoadSharedDirSingle = True
End Function

Public Function funDownLoadSharedDir(strSourceDir As String, strDestDir As String) As Boolean
'------------------------------------------------
'功能：通过共享目录的方式下载一个目录的图像，strSourceDir目录所在的计算机登录，通过另外的过程实现
'参数： strSourceDir -- 需要复制文件的源目录，即远程服务器中的共享目录
'       strDestDir  --  文件复制的目的地，即计算机的本机目录
'返回：True-- 成功，False -- 失败
'-----------------------------------------------
    Dim fs As New Scripting.FileSystemObject
    Dim fsFolder As Folder
    Dim fsFiles As Files
    Dim fsFile As File
    
    On Error Resume Next
    
    funDownLoadSharedDir = False
    
    '如果源目录不存在，则退出
    If Dir(strSourceDir, vbDirectory) = "" Then
        '记录日志
        Call WriteCommLog("funDownLoadSharedDir", "通过共享目录方式下载图像", "源目录不存在" & strSourceDir, 1)
        Exit Function
    End If
    
    '如果目标目录不存在，则创建目录
    If Dir(strDestDir, vbDirectory) = "" Then
        Call MkLocalDir(strDestDir)
    End If
    
    '遍历目录中的所有文件，一个个下载，可以确保如果本机目录中已经有了某个文件，后续的文件还可以正常下载
    Set fsFolder = fs.GetFolder(strSourceDir)
    Set fsFiles = fsFolder.Files
    
BUGEX "funDownLoadSharedDir strSourceDir =" & strSourceDir & "  strDestDir=" & strDestDir
    
    '禁止OverWrite，防止用FTP中的文件覆盖本机目录中已经有了的文件。
    For Each fsFile In fsFiles
        '记录日志
        Call WriteCommLog("funDownLoadSharedDir", "下载图像", "下载图像： " & strDestDir & "\" & fsFile.Name, 2)

BUGEX "Source=" & strSourceDir & "\" & fsFile.Name & "    Destination=" & strDestDir & "\" & fsFile.Name
        
        Call fs.CopyFile(strSourceDir & "\" & fsFile.Name, strDestDir & "\" & fsFile.Name, False)
    Next fsFile
    
    funDownLoadSharedDir = True
End Function

Public Function funRemoveSlash(strDir As String) As String
'------------------------------------------------
'功能：去除路径中的第一个和最后一个斜线，或者反斜线
'参数： strDir  -- 需要处理的路径
'返回：处理之后的路径
'-----------------------------------------------
    Dim strTemp As String
    
    strTemp = strDir
    funRemoveSlash = strTemp
    
    On Error GoTo err
    
    '去除路径后部的斜线
    If Right(strTemp, 1) = "/" Or Right(strTemp, 1) = "\" Then
        strTemp = Left(strTemp, Len(strTemp) - 1)
    End If
    
    '去除路径前的斜线
    If Left(strTemp, 1) = "/" Or Left(strTemp, 1) = "\" Then
        strTemp = Mid(strTemp, 2)
    End If
    
    funRemoveSlash = strTemp
    
    Exit Function
err:
    '暂不处理,出错则直接返回未处理的字符串
End Function

Public Function funDownLoadFTPSingle(ByVal thisMsg As clsImgInfo) As Boolean
'------------------------------------------------
'功能：根据下载图像的消息，从FTP中下载指定目录中的一个图像
'参数： thisMsg  -- 需要下载图像的消息
'返回：True -- 成功； False -- 失败
'-----------------------------------------------
    Dim lngResult As String
    Dim objFile As New FileSystemObject
    Dim strLocalFileName As String
    Dim strVirtualPath As String
BUGEX "funDownLoadFTPSingle 0"
    funDownLoadFTPSingle = False
    On Error GoTo err
    
    '先检查消息是否可用
    If thisMsg.Enable = False Or thisMsg.IP = "" Then
        
        '记录日志
        Call WriteCommLog("funDownLoadFTP", "消息不可用", "FTP方式下载图像，消息不可用或者IP地址为空，无法下载。 IP地址是：" & thisMsg.IP, 1)
        
        Exit Function
    End If

    '连接FTP
    If mftpConnect.hConnection = 0 Then
        lngResult = mftpConnect.FuncFtpConnect(thisMsg.IP, thisMsg.FTPUser, thisMsg.FTPPswd)
BUGEX "连接FTP"
        '如果连接失败，则连接备份设备
        If lngResult = 0 And mftpConnectBak.hConnection = 0 Then
            lngResult = mftpConnectBak.FuncFtpConnect(thisMsg.BakIP, thisMsg.BakFTPUser, thisMsg.BakFTPPswd)
            
            If lngResult = 0 Then
            '记录日志
            Call WriteCommLog("funDownLoadFTP", "FTP连接失败", "FTP连接失败： " & thisMsg.IP, 1)
    BUGEX "FTP连接失败"
                Exit Function
            End If
        End If
    End If
    
    '创建本地路径
    Call MkLocalDir(thisMsg.DestMainDir & objFile.GetParentFolderName(thisMsg.SubDir))
    
    '记录日志
    Call WriteCommLog("funDownLoadFTP", "下载图像", "通过FTP下载单个图像。 ", 1)
    
    strLocalFileName = Replace(thisMsg.DestMainDir & thisMsg.SubDir, "/", "\")
    strVirtualPath = Replace(thisMsg.FTPDir & objFile.GetParentFolderName(thisMsg.SubDir), "\", "/")
    
BUGEX "strVirtualPath=" & strVirtualPath & "  strLocalFileName=" & strLocalFileName & " strRemoteFileName=" & objFile.GetFileName(strLocalFileName)
    '从存储设备下载图像
    If mftpConnect.FuncDownloadFile(strVirtualPath, strLocalFileName, objFile.GetFileName(strLocalFileName)) <> 0 Then
        '下载失败则从备份设备下载图像
        If mftpConnectBak.FuncDownloadFile(strVirtualPath, strLocalFileName, objFile.GetFileName(strLocalFileName)) <> 0 Then
            Exit Function
        End If
    End If
    
    funDownLoadFTPSingle = True
    Exit Function
err:
    '暂不处理
BUGEX "funDownLoadFTPSingle Error"
End Function

Public Function funDownLoadFTP(ByVal thisMsg As clsImgInfo) As Boolean
'------------------------------------------------
'功能：根据下载图像的消息，从FTP中下载指定目录中的全部图像
'参数： thisMsg  -- 需要下载图像的消息
'返回：True -- 成功； False -- 失败
'-----------------------------------------------
    Dim lngResult As String
    
    funDownLoadFTP = False
    On Error GoTo err
    
    '先检查消息是否可用
    If thisMsg.Enable = False Or thisMsg.IP = "" Then
        
        '记录日志
        Call WriteCommLog("funDownLoadFTP", "消息不可用", "FTP方式下载图像，消息不可用或者IP地址为空，无法下载。 IP地址是：" & thisMsg.IP, 1)
        
        Exit Function
    End If
BUGEX "连接FTP"
    '连接FTP
    If mftpConnect.hConnection = 0 Then
        lngResult = mftpConnect.FuncFtpConnect(thisMsg.IP, thisMsg.FTPUser, thisMsg.FTPPswd)
BUGEX "连接FTP"
        '如果连接失败，则连接备份设备
        If lngResult = 0 And mftpConnectBak.hConnection = 0 Then
            lngResult = mftpConnectBak.FuncFtpConnect(thisMsg.BakIP, thisMsg.BakFTPUser, thisMsg.BakFTPPswd)
            
            If lngResult = 0 Then
            '记录日志
            Call WriteCommLog("funDownLoadFTP", "FTP连接失败", "FTP连接失败： " & thisMsg.IP, 1)
    BUGEX "FTP连接失败"
                Exit Function
            End If
        End If
    End If
    
    '创建本地路径
    Call MkLocalDir(thisMsg.DestMainDir & thisMsg.SubDir)
    
    '记录日志
    Call WriteCommLog("funDownLoadFTP", "下载图像", "通过FTP下载全部图像。 ", 1)
BUGEX "thisMsg.SubDir=" & Replace(thisMsg.FTPDir & thisMsg.SubDir, "\", "/") & "  strLocalPath = " & Replace(thisMsg.DestMainDir & thisMsg.SubDir, "/", "\")
    '下载图像
    If mftpConnect.funcDownLoadAllFiles(Replace("\" & thisMsg.FTPDir & thisMsg.SubDir, "\", "/"), Replace(thisMsg.DestMainDir & thisMsg.SubDir, "/", "\"), False) = 0 Then
        If mftpConnectBak.funcDownLoadAllFiles(Replace("\" & thisMsg.FTPDir & thisMsg.SubDir, "\", "/"), Replace(thisMsg.DestMainDir & thisMsg.SubDir, "/", "\"), False) = 0 Then
            Exit Function
        End If
    End If
    
    funDownLoadFTP = True
    Exit Function
err:
    '暂不处理
BUGEX "funDownLoadFTP Error"
End Function

Public Function funMsgProcess() As Boolean
'------------------------------------------------
'功能：自动处理消息
'参数： 无
'返回：True -- 成功； False -- 失败
'-----------------------------------------------
    Dim i As Integer
    Dim blnResult As Boolean
   
    On Error GoTo err

    funMsgProcess = False
    
    Call Me.DoState(False, mblnIsUpload, mlngThreadID)
    
    '消息出队并处理消息
    While funGetAMessage
BUGEX curMsg.SubDir
        If mblnIsUpload Then            '上传
BUGEX "UpLoadImages"
            '记录日志
            Call WriteCommLog("funMsgProcess", "接收到并处理消息", "准备上传此目录下的图像：" & curMsg.SubDir, 1)
                
            '处理消息，成功返回true,失败返回false
            '若上传失败则重新尝试上传
            For i = 0 To mlngRetriesn
                blnResult = funUpLoadImages(curMsg)
                
                '上传成功后退出
                If blnResult Then Exit For
            Next
            
        Else                                '下载
BUGEX "DownLoadImages"
            '记录日志
            Call WriteCommLog("funMsgProcess", "接收到并处理消息", "准备下载此目录下的图像：" & curMsg.SubDir, 1)
    
            '处理消息，成功返回true,失败返回false
            '若下载失败则重新尝试下载
            For i = 0 To mlngRetriesn
                blnResult = funDownLoadImages(curMsg)
                
                '下载成功后退出
                If blnResult Then Exit For
            Next
        End If
        
        If mblnIsUpload Then
BUGEX "UpLoadImages = " & blnResult
            If blnResult Then
                Call DoComPlete(curMsg)
            Else
                MsgBox "文件上传失败，可能由于网络不稳定造成。", vbExclamation, "提示"
            End If
        Else
BUGEX "DownLoadImages = " & blnResult
            If blnResult Then Call DoComPlete(curMsg)
        End If
    Wend
    
    Call Me.DoState(True, mblnIsUpload, mlngThreadID)
    funMsgProcess = True
    Exit Function
err:
    '暂不处理
BUGEX "funMsgProcess Err: " & err.Description
End Function

Public Function funGetAMessage() As Boolean
'------------------------------------------------
'功能：从消息队列提取一个消息，填充消息结构
'参数：
'返回：True -- 成功 ； False -- 失败
'-----------------------------------------------
    Dim objMsg As Object
    Dim objOldMsg As Object
    Dim strMsgArr() As String
    
    funGetAMessage = False
    
    On Error GoTo err
    
    '记录上一个消息，如果出现相同的连续消息，则不处理新消息
    Set objOldMsg = objMsg

    '从消息队列提取一个消息
    Set objMsg = mobjDataQueue.MsgOutQueue

    '如果消息不为空，则处理这个消息
    If Not objMsg Is Nothing And Not objOldMsg Is objMsg Then
        Set curMsg = objMsg
        curMsg.Enable = True
        
        funGetAMessage = True
    End If
    
    Exit Function
err:
    '暂不处理
BUGEX "funGetAMessage err=" & err.Description
End Function

Public Function funDownLoadImages(ByVal thisMsg As clsImgInfo) As Boolean
'------------------------------------------------
'功能：根据下载图像的消息，从共享目录或者FTP下载图像
'参数： thisMsg  -- 需要下载图像的消息
'返回：True -- 成功； False -- 失败
'-----------------------------------------------
    Dim blnResult As Boolean
    
    On Error GoTo err
BUGEX "funDownLoadImages Start"
    '如果当前消息可用，则处理消息
    If thisMsg.Enable = True Then
        '有共享目录，则使用共享目录下载图像，没有则使用FTP下载图像
BUGEX "thisMsg.SDDir=" & thisMsg.SDDir
        If thisMsg.SDDir <> "" Then
BUGEX "funDownLoadImages SheardDir "
            '连接共享目录
            If funConnectAndSaveSheardDir("\\" & thisMsg.IP & "\" & thisMsg.SDDir, thisMsg.SDUser, thisMsg.SDPswd) = True Then
                '记录日志
                Call WriteCommLog("funDownLoadImages", "共享目录方式下载图像", "通过共享目录方式下载图像，从目录： " & "\\" & thisMsg.IP & "\" & thisMsg.SDDir & "\" & thisMsg.SubDir & " ，下载到此目录： " & thisMsg.DestMainDir & thisMsg.SubDir, 1)
                
                '下载图像
                If thisMsg.IsLoadSingleFile Then       '下载目录下的一个文件
                    blnResult = funDownLoadSharedDirSingle("\\" & thisMsg.IP & "\" & thisMsg.SDDir & "\" & thisMsg.SubDir, thisMsg.DestMainDir & thisMsg.SubDir)
                Else                '下载目录下的所有文件
                    blnResult = funDownLoadSharedDir("\\" & thisMsg.IP & "\" & thisMsg.SDDir & "\" & thisMsg.SubDir, thisMsg.DestMainDir & thisMsg.SubDir)
                End If
            End If
        Else
BUGEX "funDownLoadImages FTP"
            '记录日志
            Call WriteCommLog("funDownLoadImages", "FTP方式下载图像", "通过FTP方式下载图像，从目录： \" & thisMsg.FTPDir & "\" & thisMsg.SubDir & " ，下载到此目录： " & thisMsg.DestMainDir & thisMsg.SubDir, 1)
                
            '使用FTP下载
            If thisMsg.IsLoadSingleFile Then           '下载目录下的一个文件
                blnResult = funDownLoadFTPSingle(thisMsg)
            Else                    '下载目录下的所有文件
                blnResult = funDownLoadFTP(thisMsg)
            End If
        End If
    End If
    
    funDownLoadImages = blnResult
    
    Exit Function
err:
    '暂不处理
BUGEX "funDownLoadImages Err"
End Function

Public Function funUpLoadImages(ByVal thisMsg As clsImgInfo) As Boolean
'------------------------------------------------
'功能：根据下载图像的消息，从共享目录或者FTP下载图像
'参数： thisMsg  -- 需要下载图像的消息
'返回：True -- 成功； False -- 失败
'-----------------------------------------------
    Dim blnResult As Boolean
    
    On Error GoTo err
BUGEX "funUpLoadImages Start=" & thisMsg.Enable
BUGEX "funUpLoadImages thisMsg.BakIP=" & thisMsg.BakIP

    '如果当前消息可用，则处理消息
    If thisMsg.Enable = True Then
        '有共享目录，则使用共享目录下载图像，没有则使用FTP下载图像
        If thisMsg.SDDir <> "" Then
BUGEX "funUpLoadImages Sheard"
            '连接共享目录
            If funConnectAndSaveSheardDir("\\" & thisMsg.IP & "\" & thisMsg.SDDir, thisMsg.SDUser, thisMsg.SDPswd) = True Then
                '记录日志
                Call WriteCommLog("funUpLoadImages", "共享目录方式上传图像", "通过共享目录方式上传图像，从目录： " & "\\" & thisMsg.IP & "\" & thisMsg.SDDir & "\" & thisMsg.SubDir & " ，上传到此目录： " & thisMsg.DestMainDir & thisMsg.SubDir, 1)

                '上传图像
                If thisMsg.MediaType = VIDEOTAG Then
                    blnResult = funUpLoadSharedDir(thisMsg.DestMainDir & thisMsg.SubDir & ".avi", "\\" & thisMsg.IP & "\" & thisMsg.SDDir & "\" & thisMsg.SubDir)
                ElseIf thisMsg.MediaType = AUDIOTAG Then
                    blnResult = funUpLoadSharedDir(thisMsg.DestMainDir & thisMsg.SubDir & ".wav", "\\" & thisMsg.IP & "\" & thisMsg.SDDir & "\" & thisMsg.SubDir)
                Else
                    blnResult = funUpLoadSharedDir(thisMsg.DestMainDir & thisMsg.SubDir, "\\" & thisMsg.IP & "\" & thisMsg.SDDir & "\" & thisMsg.SubDir)
                    blnResult = funUpLoadSharedDir(thisMsg.DestMainDir & thisMsg.SubDir & ".jpg", "\\" & thisMsg.IP & "\" & thisMsg.SDDir & "\" & thisMsg.SubDir & ".jpg")
                End If
            End If
        Else
BUGEX "funUpLoadImages FTP thisMsg.IP=" & thisMsg.IP
            '记录日志
            Call WriteCommLog("funUpLoadImages", "FTP方式上传图像", "通过FTP方式上传图像，从目录： \" & thisMsg.FTPDir & "\" & thisMsg.SubDir & " ，上传到此目录： " & thisMsg.DestMainDir & thisMsg.SubDir, 1)
                
            '使用FTP上传
            blnResult = funUpLoadFTP(thisMsg, False)
        End If
        
        If thisMsg.BakIP <> "" Then '备份图像
            '有共享目录，则使用共享目录下载图像，没有则使用FTP下载图像
            If thisMsg.BakSDDir <> "" Then
BUGEX "funUpLoadImages Bak Sheard"
                '连接共享目录
                If funConnectAndSaveSheardDir("\\" & thisMsg.BakIP & "\" & thisMsg.BakSDDir, thisMsg.BakSDUser, thisMsg.BakSDPswd) = True Then
                    '记录日志
                    Call WriteCommLog("funUpLoadImages", "共享目录方式备份图像", "通过共享目录方式上传图像，从目录： " & "\\" & thisMsg.IP & "\" & thisMsg.SDDir & "\" & thisMsg.SubDir & " ，上传到此目录： " & thisMsg.DestMainDir & thisMsg.SubDir, 1)
    
                    '上传图像
                    If thisMsg.MediaType = VIDEOTAG Then
                        blnResult = funUpLoadSharedDir(thisMsg.DestMainDir & thisMsg.SubDir & ".avi", "\\" & thisMsg.BakIP & "\" & thisMsg.BakSDDir & "\" & thisMsg.SubDir)
                    ElseIf thisMsg.MediaType = AUDIOTAG Then
                        blnResult = funUpLoadSharedDir(thisMsg.DestMainDir & thisMsg.SubDir & ".wav", "\\" & thisMsg.BakIP & "\" & thisMsg.BakSDDir & "\" & thisMsg.SubDir)
                    Else
                        blnResult = funUpLoadSharedDir(thisMsg.DestMainDir & thisMsg.SubDir, "\\" & thisMsg.BakIP & "\" & thisMsg.BakSDDir & "\" & thisMsg.SubDir)
                    End If
                End If
            Else
                '记录日志
                Call WriteCommLog("funUpLoadImages", "FTP方式备份图像", "通过FTP方式上传图像，从目录： \" & thisMsg.FTPDir & "\" & thisMsg.SubDir & " ，上传到此目录： " & thisMsg.DestMainDir & thisMsg.SubDir, 1)
BUGEX "funUpLoadImages Bak FTP"
                '使用FTP上传
                blnResult = funUpLoadFTP(thisMsg, True)
            End If
        End If
    End If

BUGEX "funUpLoadImages End"
    funUpLoadImages = blnResult
    
    Exit Function
err:
    '暂不处理
BUGEX "funUpLoadImages Err"
End Function

Public Function funUpLoadSharedDir(strSourceDir As String, strDestDir As String) As Boolean
'------------------------------------------------
'功能：通过共享目录的方式上传图像，strSourceDir目录所在的计算机登录，通过另外的过程实现
'参数： strSourceDir -- 需要复制文件的源目录，即远程服务器中的共享目录
'       strDestDir  --  文件复制的目的地，即计算机的本机目录
'返回：True-- 成功，False -- 失败
    Dim objFSO As New Scripting.FileSystemObject
    
BUGEX "funUpLoadSharedDir strSourceDir=" & strSourceDir & "   strDestDir=" & strDestDir
    funUpLoadSharedDir = False
    
    If Dir(objFSO.GetParentFolderName(strDestDir), vbDirectory) = "" Then
        Call MkLocalDir(objFSO.GetParentFolderName(strDestDir))
    End If

    Call objFSO.CopyFile(strSourceDir, strDestDir, True)
    
    funUpLoadSharedDir = True
End Function

Public Function funUpLoadFTP(ByVal thisMsg As clsImgInfo, Optional ByVal blnBakImg As Boolean = False) As Boolean
'------------------------------------------------
'功能：根据上传图像的消息
'参数： thisMsg  -- 需要下载图像的消息
'       blnBakImg --True:备份图像,False:存储图像
'返回：True -- 成功； False -- 失败
    Dim lngResult As String
    Dim strSrcFileName As String
    Dim strVirtualPath As String
    Dim objFSO As New Scripting.FileSystemObject
    
    funUpLoadFTP = False
    On Error GoTo err
    
BUGEX "funUpLoadFTP blnBakImg=" & blnBakImg & "  thisMsg.BakIP=" & thisMsg.BakIP & "  thisMsg.IP= " & thisMsg.IP & "   thisMsg.Enable=" & thisMsg.Enable
    '先检查消息是否可用
    If thisMsg.Enable = False Or IIf(blnBakImg, thisMsg.BakIP = "", thisMsg.IP = "") Then
        '记录日志
        Call WriteCommLog("funUpLoadFTP", "消息不可用", "FTP方式上传图像，消息不可用或者IP地址为空，无法上传。 IP地址是：" & IIf(blnBakImg, thisMsg.BakIP = "", thisMsg.IP = ""), 1)
BUGEX "funUpLoadFTP Exit"
        Exit Function
    End If
    
    '连接FTP
    If blnBakImg Then   '连接备份设备
        If mftpConnectBak.hConnection = 0 Then
            If mftpConnectBak.FuncFtpConnect(thisMsg.BakIP, thisMsg.BakFTPUser, thisMsg.BakFTPPswd) = 0 Then
                Exit Function
            End If
        End If
    Else                '连接存储设备
        If mftpConnect.hConnection = 0 Then
            If mftpConnect.FuncFtpConnect(thisMsg.IP, thisMsg.FTPUser, thisMsg.FTPPswd) = 0 Then
                Exit Function
            End If
        End If
    End If
    
    If thisMsg.MediaType = VIDEOTAG Then
        strSrcFileName = thisMsg.DestMainDir & thisMsg.SubDir & ".avi"
    ElseIf thisMsg.MediaType = AUDIOTAG Then
        strSrcFileName = thisMsg.DestMainDir & thisMsg.SubDir & ".wav"
    Else
        strSrcFileName = thisMsg.DestMainDir & thisMsg.SubDir
    End If
    
    strSrcFileName = Replace(strSrcFileName, "/", "\")
    strVirtualPath = Replace(IIf(blnBakImg, thisMsg.BakFTPDir, thisMsg.FTPDir) & objFSO.GetParentFolderName(thisMsg.SubDir), "\", "/")
    
    '记录日志
    Call WriteCommLog("funUpLoadFTP", "上传图像", "通过FTP上传图像。 ", 1)
    
    '上传图像
BUGEX "strVirtualPath=" & strVirtualPath
BUGEX "strSrcFileName=" & strSrcFileName
BUGEX "strRemoteFileName=" & objFSO.GetFileName(strSrcFileName)

    If blnBakImg Then
        Call mftpConnectBak.FuncFtpMkDir("/", strVirtualPath)
        Call mftpConnectBak.FuncUploadFile(strVirtualPath, strSrcFileName, objFSO.GetFileName(thisMsg.SubDir))
    Else
        Call mftpConnect.FuncFtpMkDir("/", strVirtualPath)
        Call mftpConnect.FuncUploadFile(strVirtualPath, strSrcFileName, objFSO.GetFileName(thisMsg.SubDir))
        
        If thisMsg.MediaType <> VIDEOTAG And thisMsg.MediaType <> AUDIOTAG Then
            Call mftpConnect.FuncUploadFile(strVirtualPath, strSrcFileName & ".jpg", objFSO.GetFileName(strSrcFileName) & ".jpg")
        End If
    End If
    
    funUpLoadFTP = True
    Exit Function
err:
    '暂不处理
BUGEX "funUpLoadFTP Err " & err.Description
End Function

Public Sub MkLocalDir(ByVal strDir As String)
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

Private Sub WriteCommLog(logSubName As String, logTitle As String, logDesc As String, lngLogLevel As Long)
'------------------------------------------------
'功能：记录通讯日志
'参数： logSubName  --  产生日志的函数名
'       logTitle   -- 日志名称
'       logDesc   --  日志内容
'       lngLogLevel -- 日志级别，通过日志级别确定当前日志是否需要记录
'返回：无
'------------------------------------------------
    Dim strLog As String
    Dim strFileName As String
    Dim intHour As Integer

    On Error GoTo err

    If mblnLogEnable = True Then        '启动了记录日志，才记录当前的日志
        '判断日志级别，确定本次日志是否需要记录
        If mlngLogLevel >= lngLogLevel Then
            '通过当前时间，创建日志文件名，每两个小时产生一个日志文件
            intHour = Hour(Time)
            intHour = intHour / 2
            intHour = intHour * 2
            strFileName = mstrLogPath & "\" & Format(Date, "YYYYMMDD") & "-" & intHour & ".log"
BUGEX "WriteCommLog strFileName=" & strFileName
            '产生日志内容
            strLog = Now() & " 日志级别： " & lngLogLevel & " 标题： " & logTitle & vbCrLf & "      函数： " & logSubName & vbCrLf & "     日志内容：" & logDesc & vbCrLf

            '打开日志文件，记录日志
            Open strFileName For Append As #1
            Print #1, strLog
            Close #1
        End If
    End If
    Exit Sub
err:
    Close #1
BUGEX "WriteCommLog Err=" & err.Description
End Sub

Public Sub BUGEX(ByVal strDebug As String, Optional ByVal blnIsForce As Boolean = False)
    If mblnOpenDebug Or blnIsForce Then
        OutputDebugString Format(Now, "mmddhhmmss") & " |-> " & strDebug
    End If
End Sub

Private Sub tmMsg_Timer()
    If funMsgProcess Then
        Unload Me
    End If
End Sub

Public Sub zlLoadImage()
    tmMsg.Enabled = True
End Sub
