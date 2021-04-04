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
'FtpFindFirstFile
'InternetFindNextFile

   
Private Const FILE_ATTRIBUTE_NORMAL = &H80
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''变量'''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Public strIPAddress As String
'Public strUser As String
'Public strPsw As String
Public hOpen As Long
Public hConnection As Long

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''函数''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function FuncFtpConnect(strHost As String, strUser As String, strPassWord As String) As Long
    hOpen = InternetOpen("WinINet1", 1, vbNullString, vbNullString, 0)
    hConnection = InternetConnect(hOpen, strHost, 0, strUser, strPassWord, INTERNET_SERVICE_FTP, 0, 255)
    FuncFtpConnect = hConnection
End Function

Public Function FuncFtpDisConnect()
    InternetCloseHandle hConnection
    InternetCloseHandle hOpen
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
    '打开FTP连接
'    hOpen = InternetOpen("WinINet1", 1, vbNullString, vbNullString, 0)
'    hConnection = InternetConnect(hOpen, strIPAddress, 0, strUser, strPsw, INTERNET_SERVICE_FTP, 0, 255)
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
'    InternetCloseHandle hConnection
'    InternetCloseHandle hOpen
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
'    Dim hOpen As Long
'    Dim hConnection As Long
    Dim bResult As Boolean      '保存返回值
    
    FuncFtpDelDir = 1
    '打开FTP连接
'    hOpen = InternetOpen("WinINet1", 1, vbNullString, vbNullString, 0)
'    hConnection = InternetConnect(hOpen, strIPAddress, 0, strUser, strPsw, INTERNET_SERVICE_FTP, 0, 255)
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
'    InternetCloseHandle hConnection
'    InternetCloseHandle hOpen
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
'    Dim hOpen As Long
'    Dim hConnection As Long
    Dim bResult As Boolean
    
    FuncUploadFile = 1
    '打开FTP连接
'    hOpen = InternetOpen("WinINet1", 1, vbNullString, vbNullString, 0)
'    hConnection = InternetConnect(hOpen, strIPAddress, 0, strUser, strPsw, INTERNET_SERVICE_FTP, 0, 255)
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
'    InternetCloseHandle hConnection
'    InternetCloseHandle hOpen
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
'    Dim hOpen As Long
'    Dim hConnection As Long
    Dim bResult As Boolean
    
    FuncDownloadFile = 1
    '打开FTP连接
'    hOpen = InternetOpen("WinINet1", 1, vbNullString, vbNullString, 0)
'    hConnection = InternetConnect(hOpen, strIPAddress, 0, strUser, strPsw, INTERNET_SERVICE_FTP, 0, 255)
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
'    InternetCloseHandle hConnection
'    InternetCloseHandle hOpen
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
'    Dim hOpen As Long
'    Dim hConnection As Long
    Dim i As Integer
    Dim bResult As Boolean
    
    FuncDownLoadFiles = 1
    '打开FTP连接
'    hOpen = InternetOpen("WinINet1", 1, vbNullString, vbNullString, 0)
'    hConnection = InternetConnect(hOpen, strIPAddress, 0, strUser, strPsw, INTERNET_SERVICE_FTP, 0, 255)
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
'    InternetCloseHandle hConnection
'    InternetCloseHandle hOpen
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
'    Dim hOpen As Long
'    Dim hConnection As Long
    Dim i As Integer
    Dim bResult As Boolean
    
    FuncUpLoadFiles = 1
    '打开FTP连接
'    hOpen = InternetOpen("WinINet1", 1, vbNullString, vbNullString, 0)
'    hConnection = InternetConnect(hOpen, strIPAddress, 0, strUser, strPsw, INTERNET_SERVICE_FTP, 0, 255)
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
'    InternetCloseHandle hConnection
'    InternetCloseHandle hOpen
    FuncUpLoadFiles = 0
End Function


Public Function FuncDelFile(strVirtualPath As String, strFileName As String) As Long
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''功能: 删除FTP服务器的指定文件
    '''参数：strVirtualPath---FTP虚拟目录
    '''参数：strFileName()---需要删除的文件名
    '''返回：成功返回0，连接失败返回1，删除文件失败返回2
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'    Dim hOpen As Long
'    Dim hConnection As Long
    Dim bResult As Boolean
    
    FuncDelFile = 1
    '打开FTP连接
'    hOpen = InternetOpen("WinINet1", 1, vbNullString, vbNullString, 0)
'    hConnection = InternetConnect(hOpen, strIPAddress, 0, strUser, strPsw, INTERNET_SERVICE_FTP, 0, 255)
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
'    InternetCloseHandle hConnection
'    InternetCloseHandle hOpen
    FuncDelFile = IIf(bResult, 0, 2)
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
'    hOpen = InternetOpen("WinINet1", 1, vbNullString, vbNullString, 0)
'    hConnection = InternetConnect(hOpen, strIPAddress, 0, strUser, strPsw, INTERNET_SERVICE_FTP, 0, 255)
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
'    InternetCloseHandle hConnection
'    InternetCloseHandle hOpen
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

