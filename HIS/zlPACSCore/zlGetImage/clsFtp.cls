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
Attribute VB_Creatable = False
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

Private Declare Function FtpRemoveDirectory Lib "wininet.dll" Alias "FtpRemoveDirectoryA" (ByVal hFtpSession As Long, ByVal lpszDirectory As String) As Boolean


                                ''''''''''''''''''''''''
                                '''''''FTP文件操作''''''
                                '''''''''''''''''''''''
    
Private Declare Function FtpGetFile Lib "wininet.dll" Alias "FtpGetFileA" (ByVal hFtpSession As Long, ByVal lpszRemoteFile As String, ByVal lpszNewFile As String, ByVal fFailIfExists As Boolean, ByVal dwFlagsAndAttributes As Long, ByVal dwFlags As Long, ByVal dwContext As Long) As Boolean
    
Private Declare Function FtpPutFile Lib "wininet.dll" Alias "FtpPutFileA" (ByVal hFtpSession As Long, ByVal lpszLocalFile As String, ByVal lpszRemoteFile As String, ByVal dwFlags As Long, ByVal dwContext As Long) As Boolean

Private Declare Function FtpDeleteFile Lib "wininet.dll" Alias "FtpDeleteFileA" (ByVal hFtpSession As Long, ByVal lpszFileName As String) As Boolean

Private Declare Function FtpFindFirstFile Lib "wininet.dll" Alias "FtpFindFirstFileA" (ByVal hFtpSession As Long, ByVal lpszSearchFile As String, lpFindFileData As WIN32_FIND_DATA, ByVal dwFlags As Long, ByVal dwContent As Long) As Long

Private Declare Function InternetFindNextFile Lib "wininet.dll" Alias "InternetFindNextFileA" (ByVal hFind As Long, lpvFindData As WIN32_FIND_DATA) As Long

   
Private Const FILE_ATTRIBUTE_NORMAL = &H80
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''变量'''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public hOpen As Long
Public hConnection As Long

'记录日志状态的变量
Private mblnLogEnable As Boolean     '是否启动日志
Private mstrLogPath As String        '日志路径
Private mlngLogLevel As Long         '日志级别，1级为普通日志，2级为详细日志

Private mstrIPAddress As String
Private mstrUser As String
Private mstrPswd As String



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''属性'''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Get IPAddress() As String
    IPAddress = mstrIPAddress
End Property

Public Property Get User() As String
    User = mstrUser
End Property

Public Property Get PassWord() As String
    PassWord = mstrPswd
End Property

'mblnLogEnable
Public Property Let blnLogEnable(ByVal vData As Boolean)
    mblnLogEnable = vData
    '如果启动日志，而且日志的路径为空，则设置日志路径为默认路径
    
    On Error Resume Next
    
    If mblnLogEnable = True Then
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
End Property

Public Property Get blnLogEnable() As Boolean
    blnLogEnable = mblnLogEnable
End Property

'mstrLogPath
Public Property Let strLogPath(ByVal vData As String)
    mstrLogPath = vData
    
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
    
End Property

Public Property Get strLogPath() As String
    strLogPath = mstrLogPath
End Property

'mlngLogLevel
Public Property Let lngLogLevel(ByVal vData As Long)
    mlngLogLevel = vData
End Property

Public Property Get lngLogLevel() As Long
    lngLogLevel = mlngLogLevel
End Property


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''函数''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function FuncFtpConnect(strHost As String, strUser As String, strPassWord As String) As Long
    Dim i As Integer
    Dim strDecryptionPassW As String
    
    '如果密码已经加密，则需解密加密密码
    If Len(strPassWord) >= 3 Then
        If Mid(strPassWord, 1, 1) & Mid(strPassWord, 3, 1) & Mid(strPassWord, Len(strPassWord), 1) = "★※★" Then
            strDecryptionPassW = Mid(strPassWord, 2)
            strDecryptionPassW = Mid(strDecryptionPassW, 1, Len(strDecryptionPassW) - 1)
            strDecryptionPassW = Mid(strDecryptionPassW, 1, 1) & Mid(strDecryptionPassW, 3)
            strDecryptionPassW = getDecryptionPassW(strDecryptionPassW)
            
            strPassWord = strDecryptionPassW
        End If
    End If
    
    mstrIPAddress = strHost
    mstrUser = strUser
    mstrPswd = strPassWord
    
    hOpen = InternetOpen("WinINet1", 1, vbNullString, vbNullString, 0)
    hConnection = InternetConnect(hOpen, strHost, 0, strUser, strPassWord, INTERNET_SERVICE_FTP, 0, 255)
    FuncFtpConnect = hConnection
End Function

Public Function FuncFtpDisConnect()
    InternetCloseHandle hConnection
    InternetCloseHandle hOpen
    mstrIPAddress = ""
    mstrUser = ""
    mstrPswd = ""
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
    
'    Dim hOpen As Long
'    Dim hConnection As Long
    Dim i As Integer
    Dim aNestDirs() As String               '存储所有需要创建的目录
    Dim bResult As Boolean                  '操作结果
        
    FuncFtpMkDir = 1
    
    If hConnection = 0 Then
        '与ftp服务器的连接失败，显示连接错误信息
        InternetCloseHandle hConnection
        InternetCloseHandle hOpen
        FuncFtpMkDir = 1
        Exit Function
    Else               '与FTP服务器的连接成功
        '设置FTP的虚拟目录
        If strVirtualPath <> vbNullString Then
           FtpSetCurrentDirectory hConnection, strVirtualPath
        End If
        '读取全部需要创建的目录信息
        aNestDirs = Split(strDir, "/")
        '创建全部目录
        For i = 0 To UBound(aNestDirs)
            If aNestDirs(i) <> vbNullString Then
                bResult = FtpCreateDirectory(hConnection, aNestDirs(i))
                FtpSetCurrentDirectory hConnection, aNestDirs(i)
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
    If hConnection = 0 Then
        '与ftp服务器的连接失败，显示连接错误信息
        InternetCloseHandle hConnection
        InternetCloseHandle hOpen
        FuncFtpDelDir = 1
        Exit Function
    Else               '与FTP服务器的连接成功
         '设置FTP的虚拟目录
        If strVirtualPath <> vbNullString Then
           FtpSetCurrentDirectory hConnection, strVirtualPath
        End If
        bResult = FtpRemoveDirectory(hConnection, strDir)
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
    If hConnection = 0 Then
        '与ftp服务器的连接失败，显示连接错误信息
        InternetCloseHandle hConnection
        InternetCloseHandle hOpen
        FuncUploadFile = 1
        Exit Function
    Else               '与FTP服务器的连接成功
         '设置FTP的虚拟目录
        If strVirtualPath <> vbNullString Then
           FtpSetCurrentDirectory hConnection, strVirtualPath
        End If
        bResult = FtpPutFile(hConnection, strLocalFileName, strRemoteFileName, 2, 0)
    End If
    FuncUploadFile = IIf(bResult, 0, 2)
End Function

Public Function FuncDownloadFile(strVirtualPath As String, strLocalFileName As String, strRemoteFileName As String) As Long
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''功能：FTP服务器的指定目录下载一个文件
    '''参数：strVirtualPath---FTP虚拟目录
    '''参数：strLocalFileName---本地文件名及路径，如“C:\SSS.TXT”
    '''参数：strRemoteFileName---需要下载的文件名
    '''返回：成功返回0，连接失败返回1，下载失败返回2
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim bResult As Boolean
    
    FuncDownloadFile = 1
    If hConnection = 0 Then
        '与ftp服务器的连接失败，显示连接错误信息
        InternetCloseHandle hConnection
        InternetCloseHandle hOpen
        FuncDownloadFile = 1
        Exit Function
    Else               '与FTP服务器的连接成功
         '设置FTP的虚拟目录
        If strVirtualPath <> vbNullString Then
           FtpSetCurrentDirectory hConnection, strVirtualPath
        End If
        bResult = FtpGetFile(hConnection, strRemoteFileName, strLocalFileName, False, FILE_ATTRIBUTE_NORMAL, 2, 0)
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
    If hConnection = 0 Then
        '与ftp服务器的连接失败，显示连接错误信息
        InternetCloseHandle hConnection
        InternetCloseHandle hOpen
        FuncDownLoadFiles = 1
        Exit Function
    Else               '与FTP服务器的连接成功
         '设置FTP的虚拟目录
        If strVirtualPath <> vbNullString Then
           FtpSetCurrentDirectory hConnection, strVirtualPath
        End If
        For i = 1 To UBound(strRemoteFileNames)
            bResult = FtpGetFile(hConnection, strRemoteFileNames(i), strLocalPath & "\" & strRemoteFileNames(i), False, FILE_ATTRIBUTE_NORMAL, 2, 0)
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
    If hConnection = 0 Then
        '与ftp服务器的连接失败，显示连接错误信息
        InternetCloseHandle hConnection
        InternetCloseHandle hOpen
        FuncUpLoadFiles = 1
        Exit Function
    Else               '与FTP服务器的连接成功
         '设置FTP的虚拟目录
        If strVirtualPath <> vbNullString Then
           FtpSetCurrentDirectory hConnection, strVirtualPath
        End If
        For i = 1 To UBound(strLocalFileNames)
            bResult = FtpPutFile(hConnection, strLocalPath & "\" & strLocalFileNames(i), strLocalFileNames(i), 2, 0)
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
    If hConnection = 0 Then
        '与ftp服务器的连接失败，显示连接错误信息
        InternetCloseHandle hConnection
        InternetCloseHandle hOpen
        FuncDelFile = 1
        Exit Function
    Else               '与FTP服务器的连接成功
         '设置FTP的虚拟目录
        If strVirtualPath <> vbNullString Then
           FtpSetCurrentDirectory hConnection, strVirtualPath
        End If
        bResult = FtpDeleteFile(hConnection, strFileName)
    End If
    FuncDelFile = IIf(bResult, 0, 2)
End Function


Public Function FuncDelFiles(strVirtualPath As String, strFileNames() As String) As Long
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''功能: 删除FTP服务器同一目录下的一组文件
    '''参数：strVirtualPath---FTP虚拟目录
    '''参数：strFileNames()---需要删除的文件名数组
    '''返回：成功返回0，连接失败返回1，删除文件失败返回2
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim i As Integer
    Dim bResult As Boolean
    
    FuncDelFiles = 1
    If hConnection = 0 Then
        '与ftp服务器的连接失败，显示连接错误信息
        InternetCloseHandle hConnection
        InternetCloseHandle hOpen
        FuncDelFiles = 1
        Exit Function
    Else               '与FTP服务器的连接成功
         '设置FTP的虚拟目录
        If strVirtualPath <> vbNullString Then
           FtpSetCurrentDirectory hConnection, strVirtualPath
        End If
        For i = 1 To UBound(strFileNames)
            bResult = FtpDeleteFile(hConnection, strFileNames(i))
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
    Dim lngHINet As Long
    Dim strTemp As String
    Dim blnRC As Boolean
    
    pData.cFileName = String(MAX_FILENAME, 0)
    FuncDirFiles = ""
    lngHINet = FtpFindFirstFile(hConnection, strVirtualPath & "/*.*", pData, 0, 0)
    
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

Public Function funcDownLoadAllFiles(strVirtualPath As String, strLocalPath As String, Optional blnOverWrite As Boolean = False) As Long
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''功能: 从指定目录中下载所有文件
'''参数：   strVirtualPath---FTP虚拟目录，需要下载的文件所在的目录
'''         strLocalPath -- 本机目录，下载之后存放的目录
'''         blnOverWrite -- 是否覆盖原有文件。True -- 覆盖；False--不覆盖
'''返回：成功返回0，失败返回错误代号
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim pData As WIN32_FIND_DATA
    Dim lngFindFirstResult As Long
    Dim lngFindNextResult As Long
    Dim blnDownResult As Boolean
    Dim strFileName As String
    
    On Error GoTo err
    
    funcDownLoadAllFiles = 1
    
    '检查FTP连接是否有效,无效则退出
    If hConnection = 0 Then
        '与FTP服务器的连接失败，显示连接错误信息
        InternetCloseHandle hConnection
        InternetCloseHandle hOpen
        
        '记录日志
        Call WriteCommLog("funcDownLoadAllFiles", "FTP连接无效", "FTP连接失败，退出程序", 0)
        
        Exit Function
    End If
    
    '切换到FTP虚拟目录
    If strVirtualPath <> vbNullString Then
        FtpSetCurrentDirectory hConnection, strVirtualPath
    End If
    
    '先查找FTP目录下的文件
    pData.cFileName = String(MAX_FILENAME, 0)
    lngFindFirstResult = FtpFindFirstFile(hConnection, strVirtualPath & "\*.*", pData, 0, 0)
    
    '记录日志
    Call WriteCommLog("funcDownLoadAllFiles", "查找第一个文件", "查找FTP目录下的文件，查找结果是：" & IIf(lngFindFirstResult = 0, "没找到。", "找到，并准备下载。"), 1)
    
    If lngFindFirstResult > 0 Then
        '下载第一个文件
        strFileName = Left(pData.cFileName, InStr(1, pData.cFileName, String(1, 0), vbBinaryCompare) - 1)
        If strFileName <> "." And strFileName <> ".." Then
            '先判断本地文件是否存在，以及是否强行覆盖
            If blnOverWrite = False Then
                If Dir(strLocalPath & "\" & strFileName) = "" Then
                    '本地文件不存在，下载这个文件
                    blnDownResult = FtpGetFile(hConnection, strFileName, strLocalPath & "\" & strFileName, False, FILE_ATTRIBUTE_NORMAL, 2, 0)
                    '记录日志
                    Call WriteCommLog("funcDownLoadAllFiles", "下载第一个文件", "已下载第一个文件：" & strFileName & " ,下载结果：" & IIf(blnDownResult, "成功。", "失败。"), 1)
                Else
                    '记录日志
                    Call WriteCommLog("funcDownLoadAllFiles", "下载第一个文件", "此文件：" & strFileName & " 已经在本地存在，未下载。", 1)
                End If
            Else
                '下载这个文件
                blnDownResult = FtpGetFile(hConnection, strFileName, strLocalPath & "\" & strFileName, False, FILE_ATTRIBUTE_NORMAL, 2, 0)
                '记录日志
                Call WriteCommLog("funcDownLoadAllFiles", "下载第一个文件", "已下载第一个文件：" & strFileName & " ,下载结果：" & IIf(blnDownResult, "成功。", "失败。"), 1)
            End If
        End If
         
        Do While True
             pData.cFileName = String(MAX_FILENAME, 0)
             lngFindNextResult = InternetFindNextFile(lngFindFirstResult, pData)
             
             '记录日志
             Call WriteCommLog("funcDownLoadAllFiles", "查找剩余的文件", "剩余文件查找情况：" & IIf(lngFindNextResult = 0, "没找到。", "找到。"), 2)
             
             If lngFindNextResult = 0 Then
                Exit Do
             Else
                strFileName = Left(pData.cFileName, InStr(1, pData.cFileName, String(1, 0), vbBinaryCompare) - 1)
                If strFileName <> "." And strFileName <> ".." Then
                    '先判断本地文件是否存在，以及是否强行覆盖
                    If blnOverWrite = False Then
                        If Dir(strLocalPath & "\" & strFileName) = "" Then
                            '本地文件不存在，下载这个文件
                            blnDownResult = FtpGetFile(hConnection, strFileName, strLocalPath & "\" & strFileName, False, FILE_ATTRIBUTE_NORMAL, 2, 0)
                            '记录日志
                            Call WriteCommLog("funcDownLoadAllFiles", "下载剩余的文件", "下载剩余文件：" & strFileName & " 下载情况：" & IIf(blnDownResult, "成功。", "失败。"), 2)
                        Else
                            '记录日志
                            Call WriteCommLog("funcDownLoadAllFiles", "下载剩余的文件", "此文件：" & strFileName & " 已经在本地存在，未下载。", 2)
                        End If
                    Else
                        '下载这个文件
                        blnDownResult = FtpGetFile(hConnection, strFileName, strLocalPath & "\" & strFileName, False, FILE_ATTRIBUTE_NORMAL, 2, 0)
                        '记录日志
                        Call WriteCommLog("funcDownLoadAllFiles", "下载剩余的文件", "下载剩余文件：" & strFileName & " 下载情况：" & IIf(blnDownResult, "成功。", "失败。"), 2)
                    End If
                 End If
             End If
             
             DoEvents
         Loop
         
         InternetCloseHandle lngFindFirstResult
    End If

    funcDownLoadAllFiles = 0
    '记录日志
    Call WriteCommLog("funcDownLoadAllFiles", "完成下载", "完成下载。", 1)
    Exit Function
err:
    '记录错误日志
    Call WriteCommLog("funcDownLoadAllFiles", "下载目录中的图像失败", "错误代码：" & err.Number & " 错误信息：" & err.Description & " 出错之前最后下载的图像 ：" & strFileName, 0)
    
End Function

Private Sub Class_Initialize()
    
    '设置初始值
    mblnLogEnable = False
    mstrLogPath = ""
    mlngLogLevel = 1
    
End Sub

Private Sub MkLocalDir(ByVal strDir As String)
'------------------------------------------------
'功能：创建本地目录
'参数： strDir－－本地目录
'返回：无
'------------------------------------------------
    Dim objFile As Object
    Dim aNestDirs() As String, i As Integer
    Dim strPath As String
    On Error Resume Next
    
    '读取全部需要创建的目录信息
    ReDim Preserve aNestDirs(0)
    aNestDirs(0) = strDir
    
    Set objFile = CreateObject("Scripting.FileSystemObject")
    
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
    
    If pblnLogEnable = True Then        '启动了记录日志，才记录当前的日志
        '判断日志级别，确定本次日志是否需要记录
        If plngLogLevel >= lngLogLevel Then
            '通过当前时间，创建日志文件名，每两个小时产生一个日志文件
            intHour = Hour(Time)
            intHour = intHour / 2
            intHour = intHour * 2
            strFileName = pstrLogPath & "\" & Date & "-" & intHour & ".log"
            
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
End Sub

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
Public Function getDecryptionPassW(ByVal strPassW As String) As String
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

    getDecryptionPassW = Join(strTemp, "") '解密后的字串
End Function
