VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsConnect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'===========================================================================
'                0、模块变量与公共类型、API
'===========================================================================
Private Declare Function InternetOpen Lib "wininet.dll" Alias "InternetOpenA" (ByVal sAgent As String, ByVal lAccessType As Long, ByVal sProxyName As String, ByVal sProxyBypass As String, ByVal lFlags As Long) As Long
'功能：打开连接Internet的会话
'说明：
'    sAgent--要调用Internet对话的应用程序名
'    lAccessType--请求的网络访问的类型
'备注：如果lAccessType设置为INTERNET_OPEN_TYPE_PRECONFIG，连接时就要基于
'    HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Internet Settings
'    注册表路径下的注册表数值ProxyEnable、ProxyServer和 ProxyOverride
'    sProxyName--指定代理服务器的名字，访问类型设置为INTERNET_OPEN_TYPE_PROXY才有效
'    sProxyBypass--指定代理服务器的名字或地址，有设置此项时lpszProxyName指定的将失效
'函数返回值：如果函数调用失败，lngINet 为0。
'    lFlags-
Private Const INTERNET_FLAG_ASYNC           As Long = &H10000000    '使用异步连接
Private Const INTERNET_FLAG_FROM_CACHE      As Long = &H1000000
Private Const INTERNET_FLAG_OFFLINE         As Long = &H1000000     '只通过永久缓冲进行下载操作
'   lAccessType
Private Const INTERNET_OPEN_TYPE_PRECONFIG  As Long = &H0           ' use registry configuration
Private Const INTERNET_OPEN_TYPE_DIRECT     As Long = &H1           'direct to net

Private Declare Function InternetConnect Lib "wininet.dll" Alias "InternetConnectA" (ByVal hInternetSession As Long, ByVal sServerName As String, ByVal nServerPort As Integer, ByVal sUsername As String, ByVal sPassword As String, ByVal lService As Long, ByVal lFlags As Long, ByVal lContext As Long) As Long
'功能：建立Internet连接，打开FTP会话
'说明：
'    hInternetSession--函数InternetOpen返回的Internet会话句柄
'    sServerName--要连接的服务器的名称或IP
'    nServerPort--要连接的Internet端口
'    sUsername--登录的用户帐号
'    sPassword--登录的口令
'    lService--要连接的服务器类型（这里是连接FTP服务器，连接的类型为常数INTERNET_SERVICE_FTP）
'    lFlags--如果传递x8000000，连接将使用被动FTP语义，传递0使用非被动语义
'    lContext--当使用回调函数时使用该参数，不使用回调服务传递0
'函数返回值：如果函数调用失败，lngINetConn 为0
'    lService
Private Const INTERNET_SERVICE_FTP          As Long = &H1
'    lFlags
Private Const INTERNET_FLAG_KEEP_CONNECTION  As Long = &H400000    ' use keep-alive semantics
Private Const INTERNET_FLAG_PASSIVE         As Long = &H8000000   ' used for FTP connections

Private Declare Function FtpGetFile Lib "wininet.dll" Alias "FtpGetFileA" (ByVal hFtpSession As Long, ByVal lpszRemoteFile As String, ByVal lpszNewFile As String, ByVal fFailIfExists As Boolean, ByVal dwFlagsAndAttributes As Long, ByVal dwFlags As Long, ByVal dwContext As Long) As Long
'功能：从FTP服务器上下载一个文件
'说明：
'    hFtpSession--函数InternetConnect返回的Internet连接句柄
'    lpszRemoteFile--想要获得的FTP服务器上的文件名
'    lpszNewFile--要保存在本地机器中的文件名
'    fFailIfExists--0（替换本地文件）或1 （如果本地文件已经存在则调用失败）。
'    dwFlagsAndAttributes--用来指定本地文件的文件属性，传递0忽略
'    dwFlags--文件的传输方式可能包括下列值：
'        FTP_TRANSFER_TYPE_ASCII                   1           用ASCII 传输文件（A类传输方法）
'        FTP_TRANSFER_TYPE_BINARY                 2           用二进制传输文件（B类传输方法）
'    dwContext--要取回的文件的描述表标识符
'函数返回值：如果函数调用失败，blnRC 为FALSE
Private Const FTP_TRANSFER_TYPE_BINARY      As Long = &H2
Private Const INTERNET_FLAG_RELOAD          As Long = &H80000000  'retrieve the original item
Private Const INTERNET_FLAG_NO_CACHE_WRITE  As Long = &H4000000
Private Const INTERNET_FLAG_DONT_CACHE      As Long = INTERNET_FLAG_NO_CACHE_WRITE

'Private Declare Function FtpPutFile Lib "wininet.dll" Alias "FtpPutFileA" (ByVal hConnect As Long, ByVal lpszLocalFile As String, ByVal lpszNewRemoteFile As String, ByVal dwFlags As Long, ByVal dwContext As Long) As Boolean
Private Declare Function FtpFindFirstFile Lib "wininet.dll" Alias "FtpFindFirstFileA" (ByVal hFtpSession As Long, ByVal lpszSearchFile As String, lpFindFileData As WIN32_FIND_DATA, ByVal dwFlags As Long, ByVal dwContent As Long) As Long
' 检索指定的路径
Private Const MAX_PATH = 260
Private Type FILETIME
    dwLowDateTime                           As Long
    dwHighDateTime                          As Long
End Type
Private Type WIN32_FIND_DATA
    dwFileAttributes                        As Long
    ftCreationTime                          As FILETIME
    ftLastAccessTime                        As FILETIME
    ftLastWriteTime                         As FILETIME
    nFileSizeHigh                           As Long
    nFileSizeLow                            As Long
    dwReserved0                             As Long
    dwReserved1                             As Long
    cFileName                               As String * MAX_PATH
    cAlternate                              As String * 14
End Type

Private Declare Function InternetCloseHandle Lib "wininet.dll" (ByVal hInet As Long) As Integer
'功能：关闭Internet连接
'说明：
'hInet--要关闭的会话（InternetOpen）或连接（InternetConnect）句柄
'函数返回值：
'常量定义
Private Declare Function WNetAddConnection2 Lib "mpr.dll" Alias "WNetAddConnection2A" (lpNetResource As NETRESOURCE_STRING, ByVal lpPassword As String, ByVal lpUserName As String, ByVal dwFlags As Long) As Long
Private Declare Function WNetOpenEnum Lib "mpr.dll" Alias "WNetOpenEnumA" (ByVal dwScope As Long, ByVal dwType As Long, ByVal dwUsage As Long, lpNetResource As Any, lplngEnum As Long) As Long
Private Declare Function WNetEnumResource Lib "mpr.dll" Alias "WNetEnumResourceA" (ByVal lngEnum As Long, lpcCount As Long, lpBuffer As NETRESOURCE, lpBufferSize As Long) As Long
Private Declare Function WNetCloseEnum Lib "mpr.dll" (ByVal lngEnum As Long) As Long
Private Declare Function WNetCancelConnection Lib "mpr.dll" Alias "WNetCancelConnection2A" (ByVal lpName As String, ByVal dwFlags As Long, ByVal fForce As Long) As Long
Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As Any) As Long

Private Type NETRESOURCE
    dwScope                                 As Long
    dwType                                  As Long
    dwDisplayType                           As Long
    dwUsage                                 As Long
    pLocalName                              As Long
    pRemoteName                             As Long
    pComment                                As Long
    pProvider                               As Long
End Type
'string版本的NETRESOURCE，方便传参
Private Type NETRESOURCE_STRING
    dwScope                                 As Long
    dwType                                  As Long
    dwDisplayType                           As Long
    dwUsage                                 As Long
    pLocalName                              As String
    pRemoteName                             As String
    pComment                                As String
    pProvider                               As String
End Type

Private Type NETRESOURCE_REAL
    dwScope                                 As Long
    dwType                                  As Long
    dwDisplayType                           As Long
    dwUsage                                 As Long
    sLocalName                              As String
    sRemoteName                             As String
    sComment                                As String
    sProvider                               As String
End Type
Private Const RESOURCE_CONNECTED            As Long = &H1
Private Const RESOURCE_ENUM_ALL             As Long = &HFFFF
Private Const RESOURCE_GLOBALNET            As Long = &H2
Private Const RESOURCETYPE_DISK             As Long = &H1
Private Const RESOURCEDISPLAYTYPE_SHARE     As Long = &H3
Private Const RESOURCEUSAGE_CONNECTABLE     As Long = &H1
Private Const CONNECT_UPDATE_PROFILE        As Long = &H1
Private Const NO_ERROR                      As Long = 0

Private mlngINet                            As Long                 'FTP方式的网络连接
Private mlngINetConn                        As Long                 'FTP连接
Private mintConnTimes                       As Integer              '重新连接次数。

Public Enum ServerConnectType
    SCT_NONE = -1 '无连接
    SCT_Share = 0 '共享文件连接
    SCT_FTP = 1 'FTP连接
End Enum
'===========================================================================
'                1、保存属性的变量
'===========================================================================
'ConnectType当前连接类型(R/W)
Private msctCur                             As ServerConnectType
'ServerPath当前服务器位置(R/W)
Private mstrCurServerPath                   As String
'ConnUser连接当前服务器的用户名(R/W)
Private mstrCurConnUser                     As String
'ConnPassword连接当前服务器的密码(R/W)
Private mstrCurConnPWD                      As String
'ServerPort连接当前服务器的端口(R/W)，仅对FTP方式
Private mstrCurServerPort                   As String
'CollectType收集的文件类型(R/W)，仅对收集文件
Private mstrCollectType                     As String
'===========================================================================
'                2、属性定义
'===========================================================================
'ConnectType当前连接类型(R/W)
Public Property Let ConnectType(ByVal sctCur As ServerConnectType)
    msctCur = sctCur
End Property

Public Property Get ConnectType() As ServerConnectType
    ConnectType = msctCur
End Property

'ServerPath当前服务器位置(R/W)
Public Property Let ServerPath(ByVal strCurServerPath As String)
    mstrCurServerPath = strCurServerPath
End Property

Public Property Get ServerPath() As String
    ServerPath = mstrCurServerPath
End Property

'ConnUser连接当前服务器的用户名(R/W)
Public Property Let ConnUser(ByVal strCurConnUser As String)
    mstrCurConnUser = strCurConnUser
End Property

Public Property Get ConnUser() As String
    ConnUser = mstrCurConnUser
End Property
'ConnPassword连接当前服务器的密码(R/W)
Public Property Let ConnPassword(ByVal strCurConnPWD As String)
    mstrCurConnPWD = strCurConnPWD
End Property

Public Property Get ConnPassword() As String
    ConnPassword = mstrCurConnPWD
End Property

'ServerPort连接当前服务器的端口(R/W)，仅对FTP方式
Public Property Let ServerPort(ByVal strCurServerPort As String)
    mstrCurServerPort = strCurServerPort
End Property

Public Property Get ServerPort() As String
    ServerPort = mstrCurServerPort
End Property
'CollectType收集的文件类型(R/W)，仅对收集文件
Public Property Let CollectType(ByVal strCollectType As String)
    mstrCollectType = strCollectType
End Property

Public Property Get CollectType() As String
    CollectType = mstrCollectType
End Property
'===========================================================================
'                3、公共方法
'===========================================================================
Public Function ToConnect(ByVal sctConnectType As ServerConnectType, ByVal strPath As String, strUser As String, strUserPwd As String, Optional ByVal strPort As String, Optional ByVal strCollectType As String, Optional ByRef strMsg As String) As Boolean
'功能：创建一个连接
'参数：intConnectType=连接类型，0-网络共享连接，1-Ftp连接
'      strPath=资源位置
'      strUser\strUserPwd=访问资源的用户名与密码
'      strPort=访问资源的端口
'      strCollectType=收集文件类型
'返回：是否连接成功
    Dim nrCur As NETRESOURCE_STRING
    Dim lngRet As Long
    mintConnTimes = 0
    gobjTrace.WriteSection "连接文件服务器", SL_LevelThree
    gobjTrace.WriteInfo "ToConnect", "连接类型", IIf(sctConnectType = SCT_FTP, "FTP", "共享"), "服务器目录", strPath, "收集类型", strCollectType
    If sctConnectType = SCT_Share Then '检查服务器是否正常并连接
        If Not gobjFSO.FolderExists(strPath) Then
            With nrCur
                .dwScope = RESOURCE_GLOBALNET
                .dwType = RESOURCETYPE_DISK
                .dwDisplayType = RESOURCEDISPLAYTYPE_SHARE
                .dwUsage = RESOURCEUSAGE_CONNECTABLE
                .pLocalName = "" '映射的驱动器
                .pRemoteName = strPath  '服务器路径
            End With
            On Error GoTo errH
            lngRet = WNetAddConnection2(nrCur, strUserPwd, strUser, CONNECT_UPDATE_PROFILE)
            If lngRet <> NO_ERROR Then
                strMsg = GetLastDllErr(lngRet)
                '备忘：不允许一个用户使用一个以上用户名与服务器或共享资源的多重连接。中断与此服务器或共享资源的所有连接，然后再试一次。
                gobjTrace.WriteInfo "ToConnect", "连接共享出错", strMsg
                gobjTrace.WriteSection "-", SL_LevelThree
                Exit Function
            End If
        End If
    Else '检查是否能正常连接FTP服务器
        mlngINet = InternetOpen("FTP Control", INTERNET_OPEN_TYPE_DIRECT, vbNullString, vbNullString, 0)
        If mlngINet <= 0 Then
            strMsg = GetLastDllErr()
            gobjTrace.WriteInfo "ToConnect", "连接FTP出错", strMsg
            gobjTrace.WriteSection "-", SL_LevelThree
            Exit Function
        End If
        '新增保持连接特性
        mlngINetConn = InternetConnect(mlngINet, strPath, strPort, strUser, strUserPwd, INTERNET_SERVICE_FTP, INTERNET_FLAG_KEEP_CONNECTION Or INTERNET_FLAG_PASSIVE, 0)
        If mlngINetConn = 0 Then
            strMsg = GetLastDllErr()
            gobjTrace.WriteInfo "ToConnect", "连接FTP出错", strMsg
            gobjTrace.WriteSection "-", SL_LevelThree
            Call InternetCloseHandle(mlngINet)
            Exit Function
        End If
    End If
    Me.ConnectType = sctConnectType
    Me.ServerPath = strPath
    Me.ConnUser = strUser
    Me.ConnPassword = strUserPwd
    Me.ServerPort = strPort
    Me.CollectType = strCollectType
    ToConnect = True
    gobjTrace.WriteSection "-", SL_LevelThree
    Exit Function
errH:
    strMsg = Err.Description
    gobjTrace.WriteSection "-", SL_LevelThree
    If Not ReConnect And mintConnTimes <= 3 Then
        ToConnect = ReConnect
    Else
        ToConnect = False
    End If
End Function

Public Function ReConnect() As Boolean
'功能：创建一个连接
'参数：intConnectType=连接类型，0-网络共享连接，1-Ftp连接
'      strPath=资源位置
'      strUser\strUserPwd=访问资源的用户名与密码
'      strPort=访问资源的端口
'      strCollectType=收集文件类型
'返回：是否连接成功
    Dim nrCur As NETRESOURCE_STRING
    Dim lngRet As Long
    '切换超过3次，则不再切换，直接返回True
    If mintConnTimes > 3 Then
        ReConnect = True: Exit Function
    End If
    gobjTrace.WriteSection "重试连接文件服务器", SL_LevelThree
    mintConnTimes = mintConnTimes + 1
    Call CloseConnect
    If Me.ConnectType = SCT_Share Then '检查服务器是否正常并连接
        If Not gobjFSO.FolderExists(Me.ServerPath) Then
            With nrCur
                .dwScope = RESOURCE_GLOBALNET
                .dwType = RESOURCETYPE_DISK
                .dwDisplayType = RESOURCEDISPLAYTYPE_SHARE
                .dwUsage = RESOURCEUSAGE_CONNECTABLE
                .pLocalName = "" '映射的驱动器
                .pRemoteName = Me.ServerPath  '服务器路径
            End With
            On Error GoTo errH
            lngRet = WNetAddConnection2(nrCur, Me.ConnUser, Me.ConnPassword, CONNECT_UPDATE_PROFILE)
            If lngRet <> NO_ERROR Then
                '备忘：不允许一个用户使用一个以上用户名与服务器或共享资源的多重连接。中断与此服务器或共享资源的所有连接，然后再试一次。
                gobjTrace.WriteInfo "ReConnect", "连接共享出错", GetLastDllErr(lngRet)
                gobjTrace.WriteSection "-", SL_LevelThree
                Exit Function
            End If
        End If
    Else '检查是否能正常连接FTP服务器
        mlngINet = InternetOpen("FTP Control", INTERNET_OPEN_TYPE_DIRECT, vbNullString, vbNullString, 0)
        If mlngINet <= 0 Then
            gobjTrace.WriteInfo "ReConnect", "连接FTP出错", GetLastDllErr()
            gobjTrace.WriteSection "-", SL_LevelThree
            Exit Function
        End If
        '自动切换到被动FTP语义，同时新增保持连接特性
        mlngINetConn = InternetConnect(mlngINet, Me.ServerPath, Me.ServerPort, Me.ConnUser, Me.ConnPassword, INTERNET_SERVICE_FTP, INTERNET_FLAG_KEEP_CONNECTION Or IIf(mintConnTimes Mod 2 = 1, INTERNET_FLAG_PASSIVE, 0), 0)
        If mlngINetConn = 0 Then
            gobjTrace.WriteInfo "ReConnect", "连接FTP出错", GetLastDllErr()
            gobjTrace.WriteSection "-", SL_LevelThree
            Call InternetCloseHandle(mlngINet)
            Exit Function
        End If
    End If
    ReConnect = True
    Exit Function
errH:
    gobjTrace.WriteSection "-", SL_LevelThree
    ReConnect = False
End Function

Public Function CloseConnect() As Boolean
'功能：关闭连接
'      strPath=关闭其他连接的路径
    Dim lngReturn As Long
    
    On Error Resume Next
    If Me.ConnectType = SCT_Share Then
        lngReturn = WNetCancelConnection(Me.ServerPath, CONNECT_UPDATE_PROFILE, True)
        If lngReturn <> NO_ERROR Then
            gobjTrace.WriteInfo "CloseConnect", "关闭其他共享出错", GetLastDllErr(lngReturn)
        Else
            CloseConnect = True
        End If
    ElseIf Me.ConnectType = SCT_FTP Then
        If mlngINetConn <> 0 Then
           InternetCloseHandle mlngINetConn
        End If
    
        If mlngINet <> 0 Then
           InternetCloseHandle mlngINet
        End If
        CloseConnect = True
    End If
End Function

Public Function IsServerFileExists(ByVal strFileName As String) As Boolean
'功能：判断服务器文件是否存在
    Dim udtFindData As WIN32_FIND_DATA, lnghFile    As Long
    Dim strFindFile    As String

    If Me.ConnectType = SCT_Share Then
        IsServerFileExists = gobjFSO.FileExists(Me.ServerPath & "\" & strFileName)
    Else
        If Not gblnSameFTP Then
            lnghFile = FtpFindFirstFile(mlngINetConn, strFileName, udtFindData, INTERNET_FLAG_RELOAD Or INTERNET_FLAG_NO_CACHE_WRITE, 0)
            strFindFile = Left(udtFindData.cFileName, InStr(udtFindData.cFileName, vbNullChar) - 1)
            IsServerFileExists = strFindFile <> ""
            If lnghFile <> 0 Then
                Call InternetCloseHandle(lnghFile)
            End If
        Else
            IsServerFileExists = True
        End If
    End If
End Function

Public Function DownloadFile(ByVal strFileName As String, ByVal strToLoaclPath As String, Optional ByRef strErr As String) As Boolean
'功能：下载文件
'参数：strFileName=该文件的文件名
'      strToLoaclPath=下载到本地的路径
'返回：升级是否成功。
    Dim blnDownload      As Boolean
    Dim lngReturn       As Long
    
    If Me.ConnectType = SCT_FTP Then
        DoEvents
        strErr = ""
        lngReturn = FtpGetFile(mlngINetConn, strFileName, strToLoaclPath & "\" & strFileName, False, 0, FTP_TRANSFER_TYPE_BINARY, 0)
        blnDownload = lngReturn <> 0 '必须这样处理，直接通过返回布尔型，该布尔型的为True时，not 仍为True,原因未知
        If Not blnDownload Then
            strErr = GetLastDllErr
            If ReConnect() Then
                DoEvents
                strErr = ""
                lngReturn = FtpGetFile(mlngINetConn, strFileName, strToLoaclPath & "\" & strFileName, False, 0, FTP_TRANSFER_TYPE_BINARY, 0)
                blnDownload = lngReturn <> 0 '必须这样处理，直接通过返回布尔型，该布尔型的为True时，not 仍为True,原因未知
                If Not blnDownload Then strErr = GetLastDllErr
            End If
        End If
    Else
        DoEvents
        Err.Clear: strErr = ""
        Call gobjFSO.CopyFile(Me.ServerPath & "\" & strFileName, strToLoaclPath & "\" & strFileName, True)
        If Err.Number <> 0 Then
            strErr = Err.Description
            Err.Clear
            blnDownload = False
        Else
            blnDownload = True
        End If
    End If
    DownloadFile = blnDownload
End Function

Public Function UploadFile(ByVal strFileName As String, ByVal strFromLoaclPath As String, Optional ByRef strErr As String) As Boolean
'功能：上传文件
'参数：strFileName=该文件的文件名
'      strFromLoaclPath=下载到本地的路径
'返回：升级是否成功。
    Dim blnUpload      As Boolean
    Dim lngReturn       As Long
    
    If Me.ConnectType = SCT_FTP Then
        DoEvents
        strErr = ""
'        lngReturn = FtpPutFile(mlngINetConn, strFromLoaclPath & "\" & strFileName, strFileName, FTP_TRANSFER_TYPE_BINARY, 0)
        blnUpload = lngReturn <> 0 '必须这样处理，直接通过返回布尔型，该布尔型的为True时，not 仍为True,原因未知
        If Not blnUpload Then
            strErr = GetLastDllErr
            If ReConnect() Then
                DoEvents
                strErr = ""
'                lngReturn = FtpPutFile(mlngINetConn, strFromLoaclPath & "\" & strFileName, strFileName, FTP_TRANSFER_TYPE_BINARY, 0)
                blnUpload = lngReturn <> 0 '必须这样处理，直接通过返回布尔型，该布尔型的为True时，not 仍为True,原因未知
                If Not blnUpload Then strErr = GetLastDllErr
            End If
        End If
    Else
        DoEvents
        Err.Clear: strErr = ""
        Call gobjFSO.CopyFile(strFromLoaclPath & "\" & strFileName, Me.ServerPath & "\" & strFileName, True)
        If Err.Number <> 0 Then
            strErr = Err.Description
            Err.Clear
            blnUpload = False
        Else
            blnUpload = True
        End If
    End If
    UploadFile = blnUpload
End Function
'===========================================================================
'                4、私有方法
'===========================================================================
Private Function CloseNetConn(Optional strIp As String = "", Optional strComputerName As String = "") As Boolean
'功能：终止磁盘网终资源连接
'返回：终止成功,返回true,否则返回False
'说明：只要有一个连接没结成功,则也返回false,否则返回true
    Dim lngReturn           As Long, lngEnum        As Long, lngCount           As Long
    Dim lngLength           As Long, i              As Long, lngBufferSize      As Long
    Dim uNetApi(0 To 256)   As NETRESOURCE
    Dim uNet()              As NETRESOURCE_REAL
    Dim blnReturn           As Boolean
    
    On Error Resume Next
    lngReturn = WNetOpenEnum(RESOURCE_CONNECTED, 0, RESOURCEUSAGE_CONNECTABLE, ByVal 0&, lngEnum)
    If lngReturn = NO_ERROR Then
        lngCount = RESOURCE_ENUM_ALL
        lngBufferSize = UBound(uNetApi) * Len(uNetApi(0)) / 2
        lngReturn = WNetEnumResource(lngEnum, lngCount, uNetApi(0), lngBufferSize)
        If lngReturn <> NO_ERROR Then gobjTrace.WriteInfo "枚举共享资源出错", "信息", GetLastDllErr(lngReturn)
        If lngCount > 0 Then
            ReDim Preserve uNet(0 To lngCount - 1) As NETRESOURCE_REAL
            For i = 0 To lngCount - 1
                'Each   Resource   will   appear   here   as   uNet(i)
                uNet(i).dwScope = uNetApi(i).dwScope
                uNet(i).dwType = uNetApi(i).dwType
                uNet(i).dwDisplayType = uNetApi(i).dwDisplayType
                uNet(i).dwUsage = uNetApi(i).dwUsage
                If uNetApi(i).pLocalName <> 0 Then
                    lngLength = lstrlen(uNetApi(i).pLocalName)
                    uNet(i).sLocalName = Space$(lngLength)
                    CopyMemory ByVal uNet(i).sLocalName, ByVal uNetApi(i).pLocalName, lngLength
                End If
                If uNetApi(i).pRemoteName <> 0 Then
                    lngLength = lstrlen(uNetApi(i).pRemoteName)
                    uNet(i).sRemoteName = Space$(lngLength)
                    CopyMemory ByVal uNet(i).sRemoteName, ByVal uNetApi(i).pRemoteName, lngLength
                End If
                If uNetApi(i).pComment <> 0 Then
                    lngLength = lstrlen(uNetApi(i).pComment)
                    uNet(i).sComment = Space$(lngLength)
                    CopyMemory ByVal uNet(i).sComment, ByVal uNetApi(i).pComment, lngLength
                End If
                If uNetApi(i).pProvider <> 0 Then
                    lngLength = lstrlen(uNetApi(i).pProvider)
                    uNet(i).sProvider = Space$(lngLength)
                    CopyMemory ByVal uNet(i).sProvider, ByVal uNetApi(i).pProvider, lngLength
                End If
            Next
        End If
    Else
        gobjTrace.WriteInfo "枚举共享资源出错", "信息", GetLastDllErr(lngReturn)
    End If
    
    If lngEnum > 0 Then
        lngReturn = WNetCloseEnum(lngEnum)
        If lngReturn <> NO_ERROR Then
            gobjTrace.WriteInfo "枚举共享资源出错", "信息", GetLastDllErr(lngReturn)
        End If
    End If
    If lngCount = 0 Then
        CloseNetConn = True
        Exit Function
    End If
    
    '结束连接
    blnReturn = True
    For i = 0 To UBound(uNet)
        lngReturn = WNetCancelConnection(IIf(uNet(i).sLocalName = "", uNet(i).sRemoteName, uNet(i).sLocalName), CONNECT_UPDATE_PROFILE, True)
        If lngReturn <> NO_ERROR Then
            gobjTrace.WriteInfo "枚举共享资源出错", "信息", GetLastDllErr(lngReturn)
            blnReturn = False
        End If
    Next
    CloseNetConn = blnReturn
End Function
