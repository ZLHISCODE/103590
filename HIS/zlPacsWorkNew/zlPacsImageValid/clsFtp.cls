VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFtp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''FTP的API声明'''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                                '''''''''''''''''''''
                                ''''FTP连接操作'''''''
                                '''''''''''''''''''''
'打开一个根据连接类型的Internet连接
Private Declare Function InternetConnect Lib "wininet.dll" Alias "InternetConnectA" (ByVal hInternetSession As Long, ByVal sServerName As String, ByVal nServerPort As Integer, ByVal sUsername As String, ByVal sPassword As String, ByVal lService As Long, ByVal lFlags As Long, ByVal lContext As Long) As Long
'hInternetSession--函数InternetOpen（）打开Internet对话返回的值
'sServerName--要连接的服务器的名称或IP
'nServerPort--该连接的Internet端口
'sUsername--登录的用户帐号
'sPassword--登录的口令
'lService--要连接的服务器类型（这里是连接FTP服务器，连接的类型为常数INTERNET_SERVICE_FTP）

Private Declare Function InternetOpen Lib "wininet.dll" Alias "InternetOpenA" (ByVal sAgent As String, ByVal lAccessType As Long, ByVal sProxyName As String, ByVal sProxyBypass As String, ByVal lFlags As Long) As Long

Private Declare Function InternetCloseHandle Lib "wininet.dll" (ByVal hInet As Long) As Integer

'连接Internet服务的常数
Private Const INTERNET_SERVICE_FTP = 1
Private Const INTERNET_SERVICE_GOPHER = 2
Private Const INTERNET_SERVICE_HTTP = 3
Private Const INTERNET_FLAG_PASSIVE = &H8000000     '启用被动传输
Private Const MAX_FILENAME = 260

Private Type FILETIME
        dwLowDateTime As Long
        dwHighDateTime As Long
End Type

Private Type WIN32_FIND_DATA
        dwFileAttributes As Long
        ftCreationTime As FILETIME
        ftLastAccessTime As FILETIME
        ftLastWriteTime As FILETIME
        nFileSizeHigh As Long
        nFileSizeLow As Long
        dwReserved0 As Long
        dwReserved1 As Long
        cFileName As String * MAX_FILENAME
        cAlternate As String * 14
End Type

                                ''''''''''''''''''''''''
                                '''''''FTP目录操作''''''
                                '''''''''''''''''''''''
'在ftp服务器上创建目录
Private Declare Function FtpCreateDirectory Lib "wininet.dll" Alias "FtpCreateDirectoryA" (ByVal hFtpSession As Long, ByVal lpszDirectory As String) As Boolean
'lpszDirectory--包含要创建目录的字符串，可以是一个相对路径或绝对路径
 '根据打开internet连接的函数internetopen（）返回的句柄，用户名，口令
        '打开与ftp服务器的连接
        
Private Declare Function FtpSetCurrentDirectory Lib "wininet.dll" Alias "FtpSetCurrentDirectoryA" (ByVal hFtpSession As Long, ByVal lpszDirectory As String) As Boolean

Private Declare Function FtpGetCurrentDirectory Lib "wininet.dll" Alias "FtpGetCurrentDirectoryA" (ByVal hFtpSession As Long, ByVal lpszCurrentDirectory As String, ByRef lpdwCurrentDirectory As Long) As Boolean




Private Declare Function FtpRemoveDirectory Lib "wininet.dll" Alias "FtpRemoveDirectoryA" (ByVal hFtpSession As Long, ByVal lpszDirectory As String) As Boolean


                                ''''''''''''''''''''''''
                                '''''''FTP文件操作''''''
                                '''''''''''''''''''''''
    
Private Declare Function FtpGetFile Lib "wininet.dll" Alias "FtpGetFileA" (ByVal hFtpSession As Long, ByVal lpszRemoteFile As String, ByVal lpszNewFile As String, ByVal fFailIfExists As Boolean, ByVal dwFlagsAndAttributes As Long, ByVal dwFlags As Long, ByVal dwContext As Long) As Boolean
    
Private Declare Function FtpPutFile Lib "wininet.dll" Alias "FtpPutFileA" (ByVal hFtpSession As Long, ByVal lpszLocalFile As String, ByVal lpszRemoteFile As String, ByVal dwFlags As Long, ByVal dwContext As Long) As Boolean

Private Declare Function FtpDeleteFile Lib "wininet.dll" Alias "FtpDeleteFileA" (ByVal hFtpSession As Long, ByVal lpszFileName As String) As Boolean

Private Declare Function FtpFindFirstFile Lib "wininet.dll" Alias "FtpFindFirstFileA" (ByVal hFtpSession As Long, ByVal lpszSearchFile As String, lpFindFileData As WIN32_FIND_DATA, ByVal dwFlags As Long, ByVal dwContent As Long) As Long

Private Declare Function FtpRenameFile Lib "wininet.dll" Alias "FtpRenameFileA" (ByVal hFtpSession As Long, ByVal lpszExisting As String, ByVal lpszNew As String) As Boolean

Private Declare Function InternetFindNextFile Lib "wininet.dll" Alias "InternetFindNextFileA" (ByVal hFind As Long, lpvFindData As WIN32_FIND_DATA) As Long
'FtpFindFirstFile
'InternetFindNextFile

   
Private Const FILE_ATTRIBUTE_NORMAL = &H80
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''变量'''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private mstrHost As String
Private mstrUser As String
Private mstrPwd As String
Private mstrFtpRoot As String
Private mstrAbsoluteRoot As String

Public hOpen As Long
Public hConnection As Long


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''函数''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function ReadFtpTransferState() As Boolean
On Error GoTo errHandle
    ReadFtpTransferState = IIf(Val(GetSetting("ZLSOFT", "公共模块\Ftp", "启用被动传输", 0)) = 1, True, False)
    Call SaveSetting("ZLSOFT", "公共模块\Ftp", "启用被动传输", IIf(ReadFtpTransferState, 1, 0))
Exit Function
errHandle:
    ReadFtpTransferState = 0
End Function

Public Function FuncFtpConnect(strHost As String, strUser As String, strPassWord As String) As Long
    Dim strDecryptionPassW As String
    Dim blnPasv As Boolean
    Dim strRootPath As String
    
    blnPasv = ReadFtpTransferState
    hOpen = InternetOpen("WinINet1", 1, vbNullString, vbNullString, 0)
    
    '如果密码已经加密，则需解密加密密码
    If Len(strPassWord) >= 3 Then
        If Mid(strPassWord, 1, 1) & Mid(strPassWord, 3, 1) & Mid(strPassWord, Len(strPassWord), 1) = "★※★" Then
            strDecryptionPassW = Mid(strPassWord, 2)
            strDecryptionPassW = Mid(strDecryptionPassW, 1, Len(strDecryptionPassW) - 1)
            strDecryptionPassW = Mid(strDecryptionPassW, 1, 1) & Mid(strDecryptionPassW, 3)
            strDecryptionPassW = GetDecryptionPassW(strDecryptionPassW)
            
            strPassWord = strDecryptionPassW
        End If
    End If
    
    hConnection = InternetConnect(hOpen, strHost, 0, strUser, strPassWord, INTERNET_SERVICE_FTP, IIf(blnPasv = True, INTERNET_FLAG_PASSIVE, 0), IIf(blnPasv = True, 0, 255))
    
    mstrHost = strHost
    mstrUser = strUser
    mstrPwd = strPassWord
    mstrAbsoluteRoot = ""

    If InStr(UCase(strUser), "ROOT_") > 0 Then
        FtpSetCurrentDirectory hConnection, "//"
        
        strRootPath = GetFtpRootPath
        
        mstrAbsoluteRoot = "/" & Split(strRootPath, "/")(1)
        FtpSetCurrentDirectory hConnection, mstrAbsoluteRoot
        
        FtpSetCurrentDirectory hConnection, "/"
    End If
    
    '获取ftp根目录
    mstrFtpRoot = GetFtpRootPath
    
    FuncFtpConnect = hConnection
End Function

Public Function GetFtpCWD(Optional ByVal strDefault As String = "") As String
'获取ftp当前工作目录
On Error GoTo errHandle
    Dim strCwd As String
    Dim lngSize As Long
    Dim blnResult As Boolean
    
    GetFtpCWD = strDefault
    
    If hConnection = 0 Then Exit Function
    
    lngSize = 255
    strCwd = String(lngSize, 0)
    
    blnResult = FtpGetCurrentDirectory(hConnection, strCwd, lngSize)
    If blnResult = False Then Exit Function
    
    strCwd = Mid$(strCwd, 1, lngSize)
    
    GetFtpCWD = strCwd
Exit Function
errHandle:
    GetFtpCWD = strDefault
End Function

Public Function GetFtpRootPath(Optional blnRestore As Boolean = False) As String
'获取根目录路径
    If blnRestore Then
        Call FuncFtpDisConnect
        Call FuncFtpConnect(mstrHost, mstrUser, mstrPwd)
    End If
    
    GetFtpRootPath = GetFtpCWD("/")
End Function


Public Function ResotreFtpConnect() As Long
    ResotreFtpConnect = FuncFtpConnect(mstrHost, mstrUser, mstrPwd)
End Function

Public Function FuncFtpDisConnect()
    InternetCloseHandle hConnection
    InternetCloseHandle hOpen
End Function

Private Function GetFullPath(ByVal strVirtualPath As String) As String
    GetFullPath = mstrFtpRoot & IIf(strVirtualPath <> "" And (mstrFtpRoot = "" Or Len(mstrFtpRoot) >= 3), "/", "") & strVirtualPath
End Function

Private Function FuncSetVirtualPath(ByVal strVirtualPath As String, Optional ByVal blnIsReStart As Boolean = True) As Boolean
'设置当前工作目录
    Dim blnResult As Boolean
    Dim strFormatVirtualPath As String
    Dim strCurWorkDirectory As String
    Dim strFullPath As String
    
    FuncSetVirtualPath = False
    If hConnection = 0 Then Exit Function
    
    strFormatVirtualPath = FormatVirtualPath(strVirtualPath)
    
    If blnIsReStart Then
        strCurWorkDirectory = GetFtpCWD(mstrFtpRoot)
        
        strFullPath = GetFullPath(strFormatVirtualPath)
        
        If UCase(strCurWorkDirectory) = UCase(strFullPath) Then Exit Function
        
        
        blnResult = True
        While (UCase(strCurWorkDirectory) <> UCase(mstrFtpRoot)) And blnResult And strCurWorkDirectory <> "/"
            blnResult = IIf(FtpSetCurrentDirectory(hConnection, "../"), True, False)
            strCurWorkDirectory = GetFtpCWD(mstrFtpRoot)
        Wend

    End If
    
    blnResult = IIf(FtpSetCurrentDirectory(hConnection, "./" & strFormatVirtualPath), True, False)
    
    If Not blnResult Then
        blnResult = IIf(FtpSetCurrentDirectory(hConnection, "/" & strFormatVirtualPath), True, False)
    End If
    

    If Not (blnResult) Then
        '如果虚拟目录设置设备，则断开连接重新进行虚拟目录设置
        FuncFtpDisConnect
        
        FuncFtpConnect mstrHost, mstrUser, mstrPwd
        
        If hConnection = 0 Then Exit Function
        
        blnResult = IIf(FtpSetCurrentDirectory(hConnection, "./" & strFormatVirtualPath), True, False)
        If Not blnResult Then
            blnResult = IIf(FtpSetCurrentDirectory(hConnection, "/" & strFormatVirtualPath), True, False)
        End If
        
        Debug.Print "Set Ftp Path Result:" & blnResult & " Path Way: // To ./" & strFormatVirtualPath
    End If
    
    FuncSetVirtualPath = IIf(blnResult, True, False)
End Function


Public Function FuncFtpMkDir(strVirtualPath As String, strDir As String) As Long
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''功能：创建FTP目录
    '''参数：strIPAddress-----FTP主机名，或IP地址
    '''参数：strMainPath----FTP主机中的虚拟目录名
    '''参数：strUser---FTP登陆的用户名
    '''参数：strPsw---FTP登陆的密码
    '''参数：strDir---需要创建的目录，使用“/”将各级目录分隔开
    '''返回：成功返回0，连接失败返回1，创建目录失败返回2
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    Dim i As Integer
    Dim aNestDirs() As String               '存储所有需要创建的目录
    Dim bResult As Boolean                  '操作结果
        
    FuncFtpMkDir = 1
    
    '打开FTP连接
    If hConnection = 0 Then
        '与ftp服务器的连接失败，显示连接错误信息
        InternetCloseHandle hConnection
        InternetCloseHandle hOpen
        
        Exit Function
    Else               '与FTP服务器的连接成功
        '设置FTP的虚拟目录
        FuncSetVirtualPath strVirtualPath
        
        '读取全部需要创建的目录信息
        aNestDirs = Split(Replace(strDir, "//", "/"), "/")
        
        '创建全部目录
        For i = 0 To UBound(aNestDirs)
            If aNestDirs(i) <> vbNullString Then
                bResult = FtpCreateDirectory(hConnection, aNestDirs(i))
                FuncSetVirtualPath aNestDirs(i), False
            End If
        Next
    End If
    
    FuncFtpMkDir = IIf(bResult, 0, 2)
End Function


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function FuncFtpDelDir(strVirtualPath As String, strDir As String) As Long
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''功能：删除指定的目录
    '''参数：strIPAddress---FTP服务器的名称或IP地址
    '''参数：strVirtualPath---FTP的虚拟目录名称
    '''参数：strUser---FTP登陆的用户名
    '''参数：strPsw---FTP登陆的密码
    '''参数：strDir---需要删除的目录，使用“/”将各级目录分隔开，只删除最后一级目录
    '''返回：成功返回0，连接失败返回1，删除失败返回2
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    Dim bResult As Boolean      '保存返回值
    
    FuncFtpDelDir = 1
    '打开FTP连接
    If hConnection = 0 Then
        '与ftp服务器的连接失败，显示连接错误信息
        InternetCloseHandle hConnection
        InternetCloseHandle hOpen
        
        Exit Function
    Else               '与FTP服务器的连接成功
         '设置FTP的虚拟目录
        FuncSetVirtualPath strVirtualPath
           
        bResult = FtpRemoveDirectory(hConnection, Replace(strDir, "//", "/"))
    End If
    
    FuncFtpDelDir = IIf(bResult, 0, 2)
End Function

Public Function FuncUploadFile(strVirtualPath As String, strLocalFileName As String, strRemoteFileName As String) As Long
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''功能：向FTP服务器的指定目录上传一个文件
    '''参数：strVirtualPath---FTP虚拟目录
    '''参数：strLocalFileName---本地文件名及路径，如“C:\SSS.TXT”
    '''参数：strRemoteFileName---上传后的文件名
    '''返回：成功返回0，连接失败返回1，上传文件失败返回2
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    Dim bResult As Boolean
    
    FuncUploadFile = 1
    '打开FTP连接
    If hConnection = 0 Then
        '与ftp服务器的连接失败，显示连接错误信息
        InternetCloseHandle hConnection
        InternetCloseHandle hOpen
        
        Exit Function
    Else               '与FTP服务器的连接成功
         '设置FTP的虚拟目录
        FuncSetVirtualPath strVirtualPath
         
        bResult = FtpPutFile(hConnection, strLocalFileName, Replace(strRemoteFileName, "//", "/"), 2, 0)
    End If
    
    FuncUploadFile = IIf(bResult, 0, 2)
End Function

Public Function FuncChangeDir(strVirtualPath As String) As Long
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''功能：FTP服务器的指定目录下载一个文件
    '''参数：strVirtualPath---FTP虚拟目录
    '''返回：成功返回0，转换目录失败返回2
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim bResult As Boolean
    
    FuncChangeDir = 1
    If hConnection = 0 Then
        '与ftp服务器的连接失败，显示连接错误信息
        InternetCloseHandle hConnection
        InternetCloseHandle hOpen
        
        Exit Function
    Else               '与FTP服务器的连接成功
         '设置FTP的虚拟目录
        bResult = FuncSetVirtualPath(strVirtualPath)
    End If
    
    FuncChangeDir = IIf(bResult, 0, 2)
End Function

Public Function FuncDownloadOneFile(strLocalFileName As String, strRemoteFileName As String) As Long
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''功能：FTP服务器的指定目录下载一个文件
    '''参数：strLocalFileName---本地文件名及路径，如“C:\SSS.TXT”
    '''参数：strRemoteFileName---需要下载的文件名
    '''返回：成功返回0，连接失败返回1，下载失败返回2
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim bResult As Boolean
    
    FuncDownloadOneFile = 1
    If hConnection = 0 Then
        '与ftp服务器的连接失败，显示连接错误信息
        InternetCloseHandle hConnection
        InternetCloseHandle hOpen
        
        Exit Function
    Else               '与FTP服务器的连接成功
        bResult = FtpGetFile(hConnection, strRemoteFileName, strLocalFileName, False, FILE_ATTRIBUTE_NORMAL, 2, 0)
    End If
    
    FuncDownloadOneFile = IIf(bResult, 0, 2)
End Function


Public Function DownloadFileValid(strFtpPath As String, ByVal strLocalFile As String, ByVal strFileName As String) As Boolean
'检查上传的文件和本地文件大小是否一致，不一致则返回true
    Dim blnResult As Boolean, blnReUpload As Boolean
    Dim lngFtpFileSzie As Long, lngDestFileSize As Long
    Dim StrMessage As String
    Dim objFileSystem As New FileSystemObject
    
On Error GoTo errHandle
    DownloadFileValid = False
    
    '上传后对比一下大小，判断是否正常上传
    lngDestFileSize = objFileSystem.GetFile(strLocalFile).Size
    lngFtpFileSzie = FuncFtpGetFileSize(strFtpPath, strFileName)

    If lngFtpFileSzie <> lngDestFileSize Then
        '强制重读
        Call FuncDownloadFile(strFtpPath, strLocalFile, strFileName, True)
        
        lngDestFileSize = objFileSystem.GetFile(strLocalFile).Size
    End If

    If lngFtpFileSzie <> lngDestFileSize Then
        StrMessage = "下载后的本地文件大小[" & lngDestFileSize & "]与FTP原文件大小[" & lngFtpFileSzie & "]不一致" & vbCrLf & _
                     "本地文件：" & strLocalFile & vbCrLf & _
                     "FTP原文件：" & strFtpPath & strFileName & vbCrLf & _
                     "是否需要重新下载？"
        
        If MsgBox(StrMessage, vbQuestion + vbYesNo, "提示") = vbYes Then
            DownloadFileValid = True
        End If
    End If
Exit Function
errHandle:
    If MsgBox("图像文件[本地:" & strLocalFile & "  FTP:" & strFtpPath & "/" & strFileName & "]一致性检查错误,错误原因:" & Err.Description & "。" & vbCrLf & "是否重试？", vbQuestion + vbYesNo, "提示") = vbYes Then
        DownloadFileValid = True
    End If
End Function

Public Function FuncDownloadFile(strVirtualPath As String, strLocalFileName As String, strRemoteFileName As String, _
Optional ByVal blnIsForceRead As Boolean = False, Optional ByVal lngHwnd As Long = 0) As Long
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''功能：FTP服务器的指定目录下载一个文件
    '''参数：strVirtualPath---FTP虚拟目录
    '''参数：strLocalFileName---本地文件名及路径，如“C:\SSS.TXT”
    '''参数：strRemoteFileName---需要下载的文件名
    '''返回：成功返回0，连接失败返回1，下载失败返回2
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim bResult As Boolean
    
    FuncDownloadFile = 1
    '打开FTP连接
    If hConnection = 0 Then
        '与ftp服务器的连接失败，显示连接错误信息
        InternetCloseHandle hConnection
        InternetCloseHandle hOpen
        
        Exit Function
    Else               '与FTP服务器的连接成功
         '设置FTP的虚拟目录
        FuncSetVirtualPath strVirtualPath
           
        '不重Ftp缓存读取文件
        bResult = FtpGetFile(hConnection, Replace(strRemoteFileName, "//", "/"), strLocalFileName, False, FILE_ATTRIBUTE_NORMAL, IIf(blnIsForceRead, &H80000000, 2), 0)
    
        If bResult = False Then
            '可能因超时失败进行重新下载
'            Debug.Print "zlRichEpr.cFTP(" & Now & "): DownLoad File Time Out."
            bResult = FtpGetFile(hConnection, Replace(strRemoteFileName, "//", "/"), strLocalFileName, False, FILE_ATTRIBUTE_NORMAL, IIf(blnIsForceRead, &H80000000, 2), 0)
            
'            If bResult Then
'                MsgBox "图像 " & strRemoteFileName & " 下载超时，重试下载成功。", vbInformation, "提示"
'            Else
'                MsgBox "图像 " & strRemoteFileName & " 下载超时，重试下载失败。", vbInformation, "提示"
'            End If
        End If
        
        If bResult = False Then
            '兼容处理
            '以前的数据在未锁定主目录时，图像可能被存储在ftp的根目录下
            FuncSetVirtualPath ""
            bResult = FtpGetFile(hConnection, Replace(strRemoteFileName, "//", "/"), strLocalFileName, False, FILE_ATTRIBUTE_NORMAL, IIf(blnIsForceRead, &H80000000, 2), 0)
        End If
    End If
    
    FuncDownloadFile = IIf(bResult, 0, 2)
End Function

Public Function FuncDownLoadFiles(strVirtualPath As String, strLocalPath As String, strRemoteFileNames() As String) As Long
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''功能：从FTP服务器的指定目录中下载几个文件
    '''参数：strVirtualPath---FTP虚拟目录
    '''参数：strLocalPath---本地路径，如“C:\”
    '''参数：strRemoteFileNames()---需要下载的文件名数组
    '''返回：成功返回0，连接失败返回1，下载文件失败返回2
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim i As Integer
    Dim bResult As Boolean
    
    FuncDownLoadFiles = 1
    
    '打开FTP连接
    If hConnection = 0 Then
        '与ftp服务器的连接失败，显示连接错误信息
        InternetCloseHandle hConnection
        InternetCloseHandle hOpen
        
        Exit Function
    Else               '与FTP服务器的连接成功
         '设置FTP的虚拟目录
        FuncSetVirtualPath strVirtualPath
        
        For i = 1 To UBound(strRemoteFileNames)
            bResult = FtpGetFile(hConnection, Replace(strRemoteFileNames(i), "//", "/"), strLocalPath & "\" & strRemoteFileNames(i), False, FILE_ATTRIBUTE_NORMAL, 2, 0)
            If bResult = False Then
                FuncDownLoadFiles = 2
                Exit Function
            End If
        Next
    End If
    
    FuncDownLoadFiles = 0
End Function


Public Function FuncUpLoadFiles(strVirtualPath As String, strLocalPath As String, strLocalFileNames() As String) As Long
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''功能: 将本地目录中几个文件同时上传到FTP服务器的指定目录中
    '''参数：strVirtualPath---FTP虚拟目录
    '''参数：strLocalPath---本地路径，如“C:\”
    '''参数：strLocalFileNames()---需要上传的文件名数组
    '''返回：成功返回0，连接失败返回1，上传文件失败返回2
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim i As Integer
    Dim bResult As Boolean
    
    FuncUpLoadFiles = 1
    
    '打开FTP连接
    If hConnection = 0 Then
        '与ftp服务器的连接失败，显示连接错误信息
        InternetCloseHandle hConnection
        InternetCloseHandle hOpen
        
        Exit Function
    Else               '与FTP服务器的连接成功
         '设置FTP的虚拟目录
        FuncSetVirtualPath strVirtualPath
 
        For i = 1 To UBound(strLocalFileNames)
            bResult = FtpPutFile(hConnection, strLocalPath & "\" & strLocalFileNames(i), Replace(strLocalFileNames(i), "//", "/"), 2, 0)
            If bResult = False Then
                FuncUpLoadFiles = 2
                Exit Function
            End If
        Next
    End If

    FuncUpLoadFiles = 0
End Function


Public Function FuncDelFile(strVirtualPath As String, strFileName As String) As Long
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''功能: 删除FTP服务器的指定文件
    '''参数：strVirtualPath---FTP虚拟目录
    '''参数：strFileName()---需要删除的文件名
    '''返回：成功返回0，连接失败返回1，删除文件失败返回2
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim bResult As Boolean
    
    FuncDelFile = 1
    '打开FTP连接
    
    If hConnection = 0 Then
        '与ftp服务器的连接失败，显示连接错误信息
        InternetCloseHandle hConnection
        InternetCloseHandle hOpen
        
        Exit Function
    Else               '与FTP服务器的连接成功
         '设置FTP的虚拟目录
        FuncSetVirtualPath strVirtualPath
 
        bResult = FtpDeleteFile(hConnection, Replace(strFileName, "//", "/"))
    End If
    
    FuncDelFile = IIf(bResult, 0, 2)
End Function


Public Function FuncReNameFile(ByVal strSourceFile As String, ByVal strNewFile As String) As Long
    Dim bResult As Boolean
    Dim strCurFtpRoot As String
    Dim strCurSourceFile As String
    Dim strCurNewFile As String
    
    FuncReNameFile = 1
    
    If hConnection = 0 Then
        '与ftp服务器的连接失败，显示连接错误信息
        InternetCloseHandle hConnection
        InternetCloseHandle hOpen
        
        Exit Function
    Else               '与FTP服务器的连接成功
'         '设置FTP的虚拟目录
        Call FuncSetVirtualPath("")
        
        strCurFtpRoot = IIf(mstrFtpRoot = mstrAbsoluteRoot Or mstrFtpRoot = "/", "", mstrFtpRoot)
        strCurSourceFile = Replace(strSourceFile, "//", "/")
        strCurSourceFile = Replace(strCurFtpRoot & "/" & strCurSourceFile, "//", "/")
        
        strCurNewFile = Replace(strNewFile, "//", "/")
        strCurNewFile = Replace(strCurFtpRoot & "/" & strCurNewFile, "//", "/")
        
        bResult = FtpRenameFile(hConnection, strCurSourceFile, strCurNewFile)
    End If
    
    FuncReNameFile = IIf(bResult, 0, 2)
End Function


Public Function FuncDelFiles(strVirtualPath As String, strFileNames() As String) As Long
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''功能: 删除FTP服务器同一目录下的一组文件
    '''参数：strVirtualPath---FTP虚拟目录
    '''参数：strFileNames()---需要删除的文件名数组
    '''返回：成功返回0，连接失败返回1，删除文件失败返回2
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'    Dim hOpen As Long
'    Dim hConnection As Long
    Dim i As Integer
    Dim bResult As Boolean
    
    FuncDelFiles = 1
    '打开FTP连接
    If hConnection = 0 Then
        '与ftp服务器的连接失败，显示连接错误信息
        InternetCloseHandle hConnection
        InternetCloseHandle hOpen
        
        Exit Function
    Else               '与FTP服务器的连接成功
         '设置FTP的虚拟目录
        FuncSetVirtualPath strVirtualPath
        
        For i = 1 To UBound(strFileNames)
            bResult = FtpDeleteFile(hConnection, Replace(strFileNames(i), "//", "/"))
            If bResult = False Then
                FuncDelFiles = 2
                Exit Function
            End If
        Next
    End If
    
    FuncDelFiles = 0
End Function

Public Function FuncDirFiles(ByVal strVirtualPath As String) As String
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''功能: 列出FTP服务器同一目录下的一组文件
    '''参数：strVirtualPath---FTP虚拟目录
    '''返回：文件串，文件名之间以|分隔
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim pData As WIN32_FIND_DATA
    Dim lngHINet As Long ', hOpen As Long, hConnection As Long
    Dim strTemp As String
    Dim blnRC As Boolean
    Dim strFormatVirtualPath As String
    
    '格式化虚拟目录，当虚拟目录前使用了"/\VirtualPath\"等形式时，将不能从ftp上获取文件信息
    Call FuncSetVirtualPath(strVirtualPath)
    
    pData.cFileName = String(MAX_FILENAME, 0)
    FuncDirFiles = ""
    lngHINet = FtpFindFirstFile(hConnection, "*.*", pData, 0, 0)
    
    If lngHINet > 0 Then
        FuncDirFiles = Left(pData.cFileName, InStr(1, pData.cFileName, String(1, 0), vbBinaryCompare) - 1)
        Do While True
            pData.cFileName = String(MAX_FILENAME, 0)
            blnRC = InternetFindNextFile(lngHINet, pData)
            If Not blnRC Then
                Exit Do
            Else
                FuncDirFiles = FuncDirFiles & "|" & Left(pData.cFileName, InStr(1, pData.cFileName, String(1, 0), vbBinaryCompare) - 1)
            End If
            DoEvents
        Loop
        InternetCloseHandle lngHINet
        '删除“.”和“..”目录
        If Left(Trim(FuncDirFiles), 1) = "." Then
            FuncDirFiles = Mid(Trim(FuncDirFiles), 6)
        End If
    End If
End Function

Public Function FuncFtpFileExists(ByVal strVirtualPath As String, ByVal strFile As String) As Boolean
'判断ftp文件是否存在
    Dim pData As WIN32_FIND_DATA
    Dim lngHINet As Long ', hOpen As Long, hConnection As Long
    Dim strTemp As String
    Dim blnRC As Boolean
    
    FuncFtpFileExists = False
    
    Call FuncSetVirtualPath(strVirtualPath)
    
    pData.cFileName = String(MAX_FILENAME, 0)
    lngHINet = FtpFindFirstFile(hConnection, Replace(strFile, "//", "/"), pData, 0, 0)
    
    If lngHINet > 0 Then
        strTemp = Left(pData.cFileName, InStr(1, pData.cFileName, String(1, 0), vbBinaryCompare) - 1)
        
        If strTemp = strFile Then FuncFtpFileExists = True
                
        InternetCloseHandle lngHINet
    End If
End Function

Public Function FuncFtpGetFileSize(ByVal strVirtualPath As String, ByVal strFile As String) As Long
'获取FTP上的文件的大小
    Dim pData As WIN32_FIND_DATA
    Dim lngHINet As Long '
    Dim strTemp As String
    Dim strFtpFile As String
    
    FuncFtpGetFileSize = 0
    
    Call FuncSetVirtualPath(strVirtualPath)

    pData.cFileName = String(MAX_FILENAME, 0)
    lngHINet = FtpFindFirstFile(hConnection, Replace(strFile, "//", "/"), pData, 0, 0)
    
    If lngHINet > 0 Then
        strTemp = Left(pData.cFileName, InStr(1, pData.cFileName, String(1, 0), vbBinaryCompare) - 1)
        
        If strTemp = strFile Then FuncFtpGetFileSize = pData.nFileSizeLow
                
        InternetCloseHandle lngHINet
    End If
End Function


Private Function FormatVirtualPath(ByVal strVirtualPath As String) As String
'如将“//20170222/1.23.5674.104/”格式化后的路径形式为“20170222/1.23.5674.104”
    Dim i As Long
    Dim strPaths() As String
    Dim strTmp As String
    
    FormatVirtualPath = ""
    If Trim$(strVirtualPath) = "" Then Exit Function
    
    strTmp = Trim$(strVirtualPath) & "/"
    
    strTmp = Replace(strTmp, "/\", "/")
    strTmp = Replace(strTmp, "\", "/")
    strTmp = Replace(strTmp, "//", "/")
    strTmp = Replace(strTmp, "//", "/")
    
    strPaths = Split(strTmp, "/")
    strTmp = ""
    For i = 0 To UBound(strPaths)
        If Trim(strPaths(i)) <> "" Then
            If strTmp <> "" Then strTmp = strTmp & "/"
            strTmp = strTmp & strPaths(i)
        End If
    Next i
    
    FormatVirtualPath = strTmp
End Function


Private Function GetRandom(ByVal lngBase As Long) As String
    Dim lngNum As Long
    
    Randomize 99
    
    lngNum = Fix(Rnd * lngBase)
    
    If lngNum <= 0 Then lngNum = 1
    
    GetRandom = Chr(lngNum)
End Function

'获取加密密码
Public Function getEncryptionPassW(ByVal strPassW As String) As String
    Dim i As Integer
    Dim lngAsc  As Long
    Dim strTemp() As String
    Dim lngPassWLength As Integer
    Dim strRandom As String
    Dim strBase As String
        
    i = 0
    
    lngPassWLength = Len(strPassW)
    
    strBase = GetRandom(30)
    strRandom = GetRandom(30)
    
    ReDim intAsc(0 To lngPassWLength - 1), strTemp(0 To lngPassWLength - 1)
     
    Do While i < lngPassWLength
        lngAsc = Asc(Mid(strPassW, i + 1, 1))
        lngAsc = lngAsc Xor Asc(strBase) Xor Asc(strRandom)
        strTemp(i) = Chr(lngAsc)
        i = i + 1
    Loop
    
    getEncryptionPassW = strBase & Join(strTemp, "") & strRandom '加密后的字串
End Function

'获取解密密码
Public Function GetDecryptionPassW(ByVal strPassW As String) As String
    Dim i As Integer
    Dim lngAsc  As Integer
    Dim strTemp() As String
    Dim lngPassWLength As Integer
    Dim lngBase As Long
    Dim strRandom As String
    Dim strPassSouce As String

    i = 0
    
    strPassSouce = Mid(strPassW, 2, Len(strPassW) - 2)
    lngPassWLength = Len(strPassSouce)
    lngBase = Asc(Mid(strPassW, 1, 1))
    
    strRandom = Right(strPassW, 1)
    
    ReDim intAsc(0 To lngPassWLength - 1), strTemp(0 To lngPassWLength - 1)
    
    Do While i < lngPassWLength
        lngAsc = Asc(Mid(strPassSouce, i + 1, 1))
        lngAsc = lngAsc Xor Asc(strRandom) Xor lngBase
        strTemp(i) = Chr(lngAsc)
        i = i + 1
    Loop

    GetDecryptionPassW = Join(strTemp, "") '解密后的字串
End Function


