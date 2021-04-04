Attribute VB_Name = "mdlFTP"
Option Explicit
'**************************
'功能:FTP的处理方式
'编写修改:祝庆
'**************************

Public Const FILE_ATTRIBUTE_ARCHIVE = &H20
Public Const FILE_ATTRIBUTE_COMPRESSED = &H800
Public Const FILE_ATTRIBUTE_DIRECTORY = &H10
Public Const FILE_ATTRIBUTE_HIDDEN = &H2
Public Const FILE_ATTRIBUTE_NORMAL = &H80
Public Const FILE_ATTRIBUTE_READONLY = &H1
Public Const FILE_ATTRIBUTE_SYSTEM = &H4
Public Const FILE_ATTRIBUTE_TEMPORARY = &H100

'自定义数据类型FILETIME和WIN32_FIND_DATA的定义

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
    cFileName As String * MAX_PATH
    cAlternate As String * 14
End Type

'************************************************************************************说明
'''连接
'''InternetOpen初始化WinINet，返回HINTERNET handles
'''InternetConnect建立Internet连接，打开FTP、Gopher或者HTTP会话。返回HINTERNET handles
'''InternetCloseHandle关闭Internet连接
'''
'''
'''目录操作
'''FtpCreateDirectory在FTP服务器上建立目录， 需要InternetConnect返回的会话句柄
'''FtpRemoveDirectory在FTP服务器上删除目录， 需要InternetConnect返回的会话句柄
'''FtpGetCurrentDirectory获取当前在FTP服务器上的工作目录， 需要InternetConnect返回的会话句柄
'''FtpSetCurrentDirectory设置在FTP服务器上的工作目录， 需要InternetConnect返回的会话句柄
'''
'''
'''文件操作
'''FtpFindFirstFile在FTP服务器上查找符合条件的文件或目录， 需要InternetConnect返回的会话句柄
'''InternetFindNextFile在FTP服务器上继续查找下一个符合条件的文件或目录，需要FtpFindFirstFile返回的会话句柄
'''FtpPutFile上传一个文件到FTP服务器上， 需要InternetConnect返回的会话句柄
'''FtpGetFile从FTP服务器上下载一个文件， 需要InternetConnect返回的会话句柄
'''FtpDeleteFile在FTP服务器上删除一个文件， 需要InternetConnect返回的会话句柄
'''FtpRenameFile在FTP服务器上更改一个文件的名字， 需要InternetConnect返回的会话句柄
'''FtpOpenFile在FTP服务器上打开一个文件， 需要InternetConnect返回的会话句柄
'''InternetReadFile直接在FTP服务器上读取文件， 需要FtpOpenFile返回的会话句柄
'''InternetWriteFile直接在FTP服务器上写入文件， 需要FtpOpenFile返回的会话句柄
'************************************************************************************

'************************************************************************************
'打开连接Internet的会话 InternetOpen
'************************************************************************************
Public Declare Function InternetOpen Lib "wininet.dll" Alias "InternetOpenA" _
   (ByVal sAgent As String, ByVal lAccessType As Long, ByVal sProxyName As String, _
   ByVal sProxyBypass As String, ByVal lFlags As Long) As Long
'************************************************************************************
'说明：
'************************************************************************************
'    sAgent--要调用Internet对话的应用程序名
'    lAccessType--请求的网络访问的类型，包括：
'************************************************************************************
'        常量                                                          值         说明
'        INTERNET_OPEN_TYPE_PRECONFIG        0          预配置（缺省）
'        INTERNET_OPEN_TYPE_DIRECT               1          直接连接到Internet
'        INTERNET_OPEN_TYPE_PROXY                3          通过代理服务器连接
'************************************************************************************
'    备注：如果lAccessType设置为INTERNET_OPEN_TYPE_PRECONFIG，连接时就要基于
'    HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Internet Settings
'    注册表路径下的注册表数值ProxyEnable、ProxyServer和 ProxyOverride
'************************************************************************************
'    sProxyName--指定代理服务器的名字，访问类型设置为INTERNET_OPEN_TYPE_PROXY才有效
'    sProxyBypass--指定代理服务器的名字或地址，有设置此项时lpszProxyName指定的将失效
'    lFlags--会话的选项，可包括下列值：
'************************************************************************************
'        常量                                                         值          说明
'        INTERNET_FLAG_DONT_CACHE                           不对数据进行本地缓冲或通过网关服务器缓冲
'        INTERNET_FLAG_ASYNC                                      使用异步连接
'        INTERNET_FLAG_OFFLINE                                   只通过永久缓冲进行下载操作
'************************************************************************************
'函数返回值：如果函数调用失败，lngINet 为0。
'************************************************************************************

'************************************************************************************
'建立Internet连接，打开FTP会话
'************************************************************************************
Public Declare Function InternetConnect Lib "wininet.dll" Alias "InternetConnectA" _
    (ByVal hInternetSession As Long, ByVal sServerName As String, _
    ByVal nServerPort As Integer, ByVal sUsername As String, _
    ByVal sPassword As String, ByVal lService As Long, _
    ByVal lFlags As Long, ByVal lContext As Long) As Long
'************************************************************************************
'说明：
'************************************************************************************
'    hInternetSession--函数InternetOpen返回的Internet会话句柄
'    sServerName--要连接的服务器的名称或IP
'    nServerPort--要连接的Internet端口
'    sUsername--登录的用户帐号
'    sPassword--登录的口令
'    lService--要连接的服务器类型（这里是连接FTP服务器，连接的类型为常数INTERNET_SERVICE_FTP）
'    lFlags--如果传递x8000000，连接将使用被动FTP语义，传递0使用非被动语义
'    lContext--当使用回调函数时使用该参数，不使用回调服务传递0
'************************************************************************************
'函数返回值：如果函数调用失败，lngINetConn 为0
'************************************************************************************

'************************************************************************************
'从FTP服务器上下载一个文件
'************************************************************************************
Public Declare Function FtpGetFile Lib "wininet.dll" Alias "FtpGetFileA" _
    (ByVal hFtpSession As Long, ByVal lpszRemoteFile As String, _
    ByVal lpszNewFile As String, ByVal fFailIfExists As Boolean, _
    ByVal dwFlagsAndAttributes As Long, ByVal dwFlags As Long, _
    ByVal dwContext As Long) As Boolean
Private Const INTERNET_FLAG_RELOAD          As Long = &H80000000  'retrieve the original item
Private Const INTERNET_FLAG_NO_CACHE_WRITE  As Long = &H4000000
Private Const INTERNET_FLAG_DONT_CACHE      As Long = INTERNET_FLAG_NO_CACHE_WRITE
'************************************************************************************
'说明：
'************************************************************************************
'    hFtpSession--函数InternetConnect返回的Internet连接句柄
'    lpszRemoteFile--想要获得的FTP服务器上的文件名
'    lpszNewFile--要保存在本地机器中的文件名
'    fFailIfExists--0（替换本地文件）或1 （如果本地文件已经存在则调用失败）。
'    dwFlagsAndAttributes--用来指定本地文件的文件属性，传递0忽略
'    dwFlags--文件的传输方式可能包括下列值：
'************************************************************************************
'        常量                                                         值          说明
'        FTP_TRANSFER_TYPE_ASCII                   1           用ASCII 传输文件（A类传输方法）
'        FTP_TRANSFER_TYPE_BINARY                 2           用二进制传输文件（B类传输方法）
'************************************************************************************
'    dwContext--要取回的文件的描述表标识符
'************************************************************************************
'函数返回值：如果函数调用失败，blnRC 为FALSE
'************************************************************************************
Public Declare Function FtpPutFile Lib "wininet.dll" Alias "FtpPutFileA" _
(ByVal hConnect As Long, ByVal lpszLocalFile As String, _
ByVal lpszNewRemoteFile As String, ByVal dwFlags As Long, _
ByVal dwContext As Long) As Boolean


'************************************************************************************
'关闭Internet连接
'************************************************************************************
Public Declare Function InternetCloseHandle Lib "wininet.dll" (ByVal hInet As Long) As Integer
'************************************************************************************
'说明：
'************************************************************************************
'hInet--要关闭的会话（InternetOpen）或连接（InternetConnect）句柄
'************************************************************************************
'函数返回值：
'************************************************************************************

'************************************************************************************
'说明：创建目录
'************************************************************************************
Public Declare Function FtpCreateDirectory Lib "wininet.dll" Alias "FtpCreateDirectoryA" _
    (ByVal hFtpSession As Long, ByVal lpszDirectory As String) As Boolean
    
'************************************************************************************
'说明：获得当前目录
'************************************************************************************
Public Declare Function FtpGetCurrentDirectory Lib "wininet.dll" Alias "FtpGetCurrentDirectoryA" _
    (ByVal hFtpSession As Long, ByVal lpszCurrentDirectory As String, ByRef lpdwCurrentDirectory As Long) As Boolean

'************************************************************************************
'说明：获得当前目录
'************************************************************************************
Public Declare Function FtpSetCurrentDirectory Lib "wininet.dll" Alias "FtpSetCurrentDirectoryA" _
    (ByVal hFtpSession As Long, ByVal lpszCurrentDirectory As String) As Boolean

'************************************************************************************
'说明：删除文件
'************************************************************************************
Public Declare Function FtpDeleteFile Lib "wininet.dll" Alias "FtpDeleteFileA" _
    (ByVal hFtpSession As Long, ByVal lpszFileName As String) As Boolean
    
'************************************************************************************
'说明：文件更名
'************************************************************************************
Public Declare Function FtpRenameFile Lib "wininet.dll" Alias "FtpRenameFileA" _
    (ByVal hFtpSession As Long, ByVal lpszExisting As String, ByVal lpszNew As String) As Boolean
    
    
'************************************************************************************
'说明：查找文件
'************************************************************************************
Public Declare Function FtpFindFirstFile Lib "wininet.dll" Alias "FtpFindFirstFileA" _
    (ByVal hFtpSession As Long, ByVal lpszSearchFile As String _
    , ByVal lpFindFileData As WIN32_FIND_DATA, ByVal dwFlags As Long _
    , ByVal dwContent As Long) As Long
    
    
'************************************************************************************
'说明：查找下一个文件
'************************************************************************************
Public Declare Function InternetFindNextFile Lib "wininet.dll" Alias "InternetFindNextFileA" _
    (ByVal hFind As Long, lpvFindData As WIN32_FIND_DATA) As Long

'************************************************************************************
'说明：查找打开文件
'************************************************************************************
Public Declare Function FtpOpenFile Lib "wininet.dll" Alias "FtpOpenFileA" _
    (ByVal hFtpSession As Long, ByVal sFileName As String _
    , ByVal lFlags As Long, ByVal lContext As Long) As Long


'************************************************************************************
'说明：查找打开文件
'************************************************************************************
Public Declare Function InternetReadFile Lib "wininet.dll" Alias "InternetReadFileA" _
    (ByVal hFile As Long, ByVal lpBuffer As String _
    , ByVal dwNumberOfBytesToRead As Long, ByVal lpNumberOfBytes As Long) As Long
    
'************************************************************************************
'说明：获取最后一条返回信息
'************************************************************************************
Public Declare Function InternetGetLastResponseInfo Lib "wininet.dll" Alias "InternetGetLastResponseInfoA" _
    (ByVal lpdwError As Long, ByVal lpszBuffer As String _
    , ByVal lpdwBufferLength As Long) As Boolean
    
    
'************************************************************************************
'常量定义
Public Const INTERNET_OPEN_TYPE_PRECONFIG = 0
Public Const INTERNET_OPEN_TYPE_DIRECT = 1
Public Const INTERNET_OPEN_TYPE_PROXY = 3
Public Const INTERNET_SERVICE_FTP = 1
Public Const FTP_TRANSFER_TYPE_BINARY = &H2
Public Const FTP_TRANSFER_TYPE_ASCII = &H1
'************************************************************************************


Private lngINet As Long
Private lngINetConn As Long

Public Function FtpDownFile(ByVal srcPath As String, ByVal descPath As String) As Boolean
    'INTERNET_FLAG_DONT_CACHE 参数用来控制直接从文件路径获取，而不是缓存，避免一些奇葩问题
    Dim blnOK As Boolean
    blnOK = FtpGetFile(lngINetConn, srcPath, descPath, False, 0, FTP_TRANSFER_TYPE_BINARY Or INTERNET_FLAG_DONT_CACHE, 0) <> 0
    FtpDownFile = blnOK
End Function

Public Function FtpupFile(ByVal srcPath As String, ByVal descPath As String) As Boolean
    FtpupFile = FtpPutFile(lngINetConn, srcPath, descPath, FTP_TRANSFER_TYPE_BINARY, 0)
End Function

Public Function IsFtpServer(ByVal strServerPath As String, ByVal strVisitUser As String, ByVal strVisitPassWord As String, ByVal strVisitPort As String) As Boolean
    '----------------------------------------------------------------------------------------------------------
    '--功能:检查是否能正常连接FTP服务器
    '----------------------------------------------------------------------------------------------------------
'                gstrServerPath
'                gstrVisitUser
'                gstrVisitPassWord
'                gstrVisitPort
        On Error GoTo errH
        If strServerPath = "" Or strVisitUser = "" Or strVisitPassWord = "" Or strVisitPort = "" Then
            IsFtpServer = False
            Exit Function
        End If
        
        lngINet = InternetOpen("FTP Control", INTERNET_OPEN_TYPE_DIRECT, vbNullString, vbNullString, 0)
        If lngINet <= 0 Then
            IsFtpServer = False
            Exit Function
        End If
    
        lngINetConn = InternetConnect(lngINet, strServerPath, strVisitPort, strVisitUser, strVisitPassWord, INTERNET_SERVICE_FTP, 0, 0)
        
        If lngINetConn Then
            IsFtpServer = True
        Else
            IsFtpServer = False
        End If
        Exit Function
errH:
    If err Then
        IsFtpServer = False
    End If
End Function

Public Function CancelFtpServer() As Boolean
    On Error Resume Next
    If lngINetConn <> 0 Then
       InternetCloseHandle lngINetConn
    End If

    If lngINet <> 0 Then
       InternetCloseHandle lngINet
    End If
    lngINetConn = 0
    lngINet = 0
    CancelFtpServer = True
End Function
