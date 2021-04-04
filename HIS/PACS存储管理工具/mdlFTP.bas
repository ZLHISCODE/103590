Attribute VB_Name = "mdlFTP"
Option Explicit

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''FTP的API声明'''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                                '''''''''''''''''''''
                                ''''FTP连接操作'''''''
                                '''''''''''''''''''''
'打开一个根据连接类型的Internet连接
Public Declare Function InternetConnect Lib "wininet.dll" Alias "InternetConnectA" (ByVal hInternetSession As Long, ByVal sServerName As String, ByVal nServerPort As Integer, ByVal sUsername As String, ByVal sPassword As String, ByVal lService As Long, ByVal lFlags As Long, ByVal lContext As Long) As Long
'hInternetSession--函数InternetOpen（）打开Internet对话返回的值
'sServerName--要连接的服务器的名称或IP
'nServerPort--该连接的Internet端口
'sUsername--登录的用户帐号
'sPassword--登录的口令
'lService--要连接的服务器类型（这里是连接FTP服务器，连接的类型为常数INTERNET_SERVICE_FTP）

Public Declare Function InternetOpen Lib "wininet.dll" Alias "InternetOpenA" (ByVal sAgent As String, ByVal lAccessType As Long, ByVal sProxyName As String, ByVal sProxyBypass As String, ByVal lFlags As Long) As Long

Public Declare Function InternetCloseHandle Lib "wininet.dll" (ByVal hInet As Long) As Integer

'连接Internet服务的常数
Public Const INTERNET_SERVICE_FTP = 1
Public Const INTERNET_SERVICE_GOPHER = 2
Public Const INTERNET_SERVICE_HTTP = 3

                                ''''''''''''''''''''''''
                                '''''''FTP目录操作''''''
                                '''''''''''''''''''''''
'在ftp服务器上创建目录
Public Declare Function FtpCreateDirectory Lib "wininet.dll" Alias "FtpCreateDirectoryA" (ByVal hFtpSession As Long, ByVal lpszDirectory As String) As Boolean
'lpszDirectory--包含要创建目录的字符串，可以是一个相对路径或绝对路径
 '根据打开internet连接的函数internetopen（）返回的句柄，用户名，口令
        '打开与ftp服务器的连接
        
Public Declare Function FtpSetCurrentDirectory Lib "wininet.dll" Alias "FtpSetCurrentDirectoryA" (ByVal hFtpSession As Long, ByVal lpszDirectory As String) As Boolean

Public Declare Function FtpRemoveDirectory Lib "wininet.dll" Alias "FtpRemoveDirectoryA" (ByVal hFtpSession As Long, ByVal lpszDirectory As String) As Boolean


                                ''''''''''''''''''''''''
                                '''''''FTP文件操作''''''
                                '''''''''''''''''''''''
    
Public Declare Function FtpGetFile Lib "wininet.dll" Alias "FtpGetFileA" (ByVal hFtpSession As Long, ByVal lpszRemoteFile As String, ByVal lpszNewFile As String, ByVal fFailIfExists As Boolean, ByVal dwFlagsAndAttributes As Long, ByVal dwFlags As Long, ByVal dwContext As Long) As Boolean
    
Public Declare Function FtpPutFile Lib "wininet.dll" Alias "FtpPutFileA" (ByVal hFtpSession As Long, ByVal lpszLocalFile As String, ByVal lpszRemoteFile As String, ByVal dwFlags As Long, ByVal dwContext As Long) As Boolean

Public Declare Function FtpDeleteFile Lib "wininet.dll" Alias "FtpDeleteFileA" (ByVal hFtpSession As Long, ByVal lpszFileName As String) As Boolean

Public Declare Function FtpOpenFile Lib "wininet.dll" Alias "FtpOpenFileA" (ByVal hFtpSession As Long, ByVal strFileName As String, ByVal lngAccess As Long, ByVal lngFlags As Long, ByVal lngContext As Long) As Long

Public Declare Function FtpFindFirstFile Lib "wininet.dll" Alias "FtpFindFirstFileA" (ByVal hFtpSession As Long, ByVal lpszSearchFile As String, lpFindFileData As WIN32_FIND_DATA, ByVal dwFlags As Long, ByVal dwContent As Long) As Long

Public Declare Function InternetFindNextFile Lib "wininet.dll" Alias "InternetFindNextFileA" (ByVal hFind As Long, lpvFindData As WIN32_FIND_DATA) As Long
   


Public Const FILE_ATTRIBUTE_NORMAL = &H80
Public Const GENERIC_READ = &H80000000
Public Const FTP_TRANSFER_TYPE_ASCII = &H1
Public Const MAX_FILENAME = 260

Public Const ATTR_DIR = &H10

Public Type FILETIME
        dwLowDateTime As Long
        dwHighDateTime As Long
End Type

Public Type WIN32_FIND_DATA
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
