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
'**************************************************************************
'                一、API与常量
'**************************************************************************
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
'**************************************************************************
'                二、模块级变量
'**************************************************************************
'==========================================================================
'                1、模块变量
'==========================================================================
Private mlngINet                            As Long                 'FTP方式的网络连接
Private mlngINetConn                        As Long                 'FTP连接
'Private mobjAPI As APILoad 'API动态加载累
Private mcllInsPath                         As Collection
Private mintConnTimes                       As Integer              '重新连接次数。
Private mstrComputerName                    As String
Private mobjFSO                             As New FileSystemObject
'==========================================================================
'                2、枚举、结构体
'==========================================================================
Public Enum ServerConnectType
    SCT_NONE = -1 '无连接
    SCT_Share = 0 '共享文件连接
    SCT_FTP = 1 'FTP连接
End Enum

'***************************************************************************
'                三、保存类的各种属性的模块级变量，以及属性定义
'***************************************************************************
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
'ISOpen连接是否打开(R)
Private mblnOpen                            As Boolean
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
'ISOpen连接是否打开(R)
Public Property Get IsOpen() As String
    IsOpen = mblnOpen
End Property

'***************************************************************************
'                四、向外部公布的方法
'***************************************************************************
Public Function GetFileConnect(ByVal strComputerName As String) As Boolean
'功能：获取服务器文件连接
    Dim objConn As New clsConnect
    Dim sctConnType As ServerConnectType
    Dim strServerID As String, strServer As String, strUser As String, strPWD As String, strPort As String, strCollectType As String
    Dim rsTmp As ADODB.Recordset, strSQL As String
    Dim blnDefalut As Boolean, blnConnOK As Boolean
    
    On Error Resume Next
    strSQL = "Select 升级文件服务器 From ZLTools.zlClients Where 工作站=[1]"
    Set rsTmp = OpenSQLRecord(strSQL, "获取升级服务器编号", strComputerName)
    If Err.Number = 0 Then
        If Not rsTmp.EOF Then strServerID = rsTmp!升级文件服务器 & ""
    Else
        Err.Clear
        Exit Function
    End If
    On Error GoTo errH
    If strServerID <> "" Then
        strSQL = "Select 编号,类型, 位置, 用户名, 密码, 端口,Nvl(是否缺省,0) 是否缺省 , 批次 From Zltools.Zlupgradeserver Where 编号 = [1]"
        Set rsTmp = OpenSQLRecord(strSQL, "获取升级服务器", Val(strServerID))
        If Not rsTmp.EOF Then
            strServerID = rsTmp!编号 & ""
            sctConnType = IIf(rsTmp!类型 = 0, SCT_Share, SCT_FTP)
            strServer = rsTmp!位置
            strUser = rsTmp!用户名
            strPWD = DeCipher(rsTmp!密码 & "")
            strPort = rsTmp!端口 & ""
            blnDefalut = rsTmp!是否缺省 = 1
        Else
            strServerID = ""
        End If
    End If
    If strServerID <> "" Then
        blnConnOK = Me.ToConnect(sctConnType, strServer, strUser, strPWD, strPort, strCollectType)
    End If
    '连接不成功，升级服务器自动连接默认服务器
    If Not blnConnOK And Not blnDefalut Then
        strSQL = "Select 编号,类型, 位置, 用户名, 密码, 端口, 批次 From Zltools.Zlupgradeserver Where Nvl(是否缺省,0) = 1"
        Set rsTmp = OpenSQLRecord(strSQL, "获取默认升级服务器")
        If Not rsTmp.EOF Then
            strServerID = rsTmp!编号 & ""
            sctConnType = IIf(rsTmp!类型 = 0, SCT_Share, SCT_FTP)
            strServer = rsTmp!位置
            strUser = rsTmp!用户名
            strPWD = DeCipher(rsTmp!密码 & "")
            strPort = rsTmp!端口 & ""
            blnConnOK = Me.ToConnect(sctConnType, strServer, strUser, strPWD, strPort)
        End If
    End If

    If blnConnOK Then
        GetFileConnect = True
    Else
        GetFileConnect = GetFileConnectOld
    End If
    Exit Function
errH:
    If gblnTimer Then
        If ErrCenter() = 1 Then
            Resume
        End If
    Else
        MsgBox "检查文件服务器配置出错：" & Err.Description, vbInformation, gstrSysName
        Err.Clear
    End If
End Function

Private Function GetFileConnectOld() As Boolean
'功能：获取文件服务器连接，老方式
'参数：blnUpgrade=True-预升级与升级的连接 ，false-文件收集的连接
    Dim rsTmp As ADODB.Recordset, strSQL As String
    Dim sctConnType As ServerConnectType, strServerID As String
    Dim objConn As New clsConnect
    Dim arrParas() As Variant, arrValues(4) As String
    Dim strSQLPars As String, i As Integer
    Dim blnReadOk As Boolean, blnConnOK As Boolean, blnGo As Boolean
    
    On Error GoTo errH
    '获取连接类型
    sctConnType = SCT_Share
    strSQL = "Select 项目,内容 From ZLTools.zlregInfo where 项目=[1]"
    Set rsTmp = OpenSQLRecord(strSQL, "升级类型", "升级类型")
    If Not rsTmp.EOF Then
        If NVL(rsTmp!内容, 0) = 1 Then sctConnType = SCT_FTP
    End If
    '获取服务器ID
    strSQL = "Select 升级服务器,FTP服务器 From ZLTools.zlClients Where 工作站=[1]"
    Set rsTmp = OpenSQLRecord(strSQL, "获取升级服务器编号", mstrComputerName)
    If Not rsTmp.EOF Then strServerID = IIf(sctConnType = SCT_FTP, rsTmp!FTP服务器 & "", rsTmp!升级服务器 & "")
    '获取参数主信息
    If sctConnType = SCT_FTP Then
        arrParas = Array("FTP服务器", "FTP用户", "FTP密码", "FTP端口", "")
    Else
        arrParas = Array("服务器目录", "访问用户", "访问密码", "", "")
    End If
ReGetParas:
    '先获取SQL参数
    strSQLPars = ""
    For i = LBound(arrParas) To UBound(arrParas)
        If arrParas(i) <> "" Then
            strSQLPars = strSQLPars & ",'" & arrParas(i) & IIf(i <> UBound(arrParas), strServerID, "") & "'"
        End If
    Next
    strSQLPars = Mid(strSQLPars, 2)
    strSQL = "Select 项目,内容 From ZLTools.zlregInfo where 项目 in(" & strSQLPars & ")"
    Set rsTmp = OpenSQLRecord(strSQL, "获取服务器")
    If Not rsTmp.EOF Then
        For i = LBound(arrParas) To UBound(arrParas)
            If arrParas(i) <> "" Then
                rsTmp.Filter = "项目='" & arrParas(i) & IIf(i <> UBound(arrParas), strServerID, "") & "'"
                If Not rsTmp.EOF Then arrValues(i) = rsTmp!内容 & ""
            End If
        Next
    End If
    
    blnReadOk = True
    '服务器，用户，密码为空，则不能进行收集或升级
    If arrValues(0) = "" Or arrValues(1) = "" Or arrValues(2) = "" Then
        blnReadOk = False
    'FTP方式需要一个端口
    ElseIf sctConnType = SCT_FTP And arrValues(3) = "" Then
        blnReadOk = False
    End If
    If blnReadOk Then
        blnConnOK = Me.ToConnect(sctConnType, arrValues(0), arrValues(1), arrValues(2), arrValues(3), arrValues(4))
    End If
    If (Not blnConnOK Or Not blnReadOk) Then
        If strServerID <> "" And strServerID <> "0" Then
            strServerID = "0"
            GoTo ReGetParas '重新获取连接服务器的参数
        ElseIf (strServerID = "0" Or strServerID = "") And Not blnGo Then
            blnGo = True '防止循环
            strServerID = IIf(strServerID = "0", "", "0")
            GoTo ReGetParas '重新获取连接服务器的参数
        End If
    End If
    GetFileConnectOld = blnConnOK
    Exit Function
errH:
    If gblnTimer Then
        If ErrCenter() = 1 Then
            Resume
        End If
    Else
        MsgBox "获取以前方式的缺省服务器出错：" & Err.Description, vbInformation, gstrSysName
        Err.Clear
    End If
End Function

Public Function ToConnect(ByVal sctConnectType As ServerConnectType, ByVal strPath As String, strUser As String, strUserPwd As String, Optional ByVal strPort As String, Optional ByVal strCollectType As String) As Boolean
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
    mblnOpen = False
    If sctConnectType = SCT_Share Then '检查服务器是否正常并连接
        If Not mobjFSO.FolderExists(strPath) Then
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
                '备忘：不允许一个用户使用一个以上用户名与服务器或共享资源的多重连接。中断与此服务器或共享资源的所有连接，然后再试一次。
                Exit Function
            End If
        End If
    Else '检查是否能正常连接FTP服务器
        mlngINet = InternetOpen("FTP Control", INTERNET_OPEN_TYPE_DIRECT, vbNullString, vbNullString, 0)
        If mlngINet <= 0 Then
            Exit Function
        End If
        '新增保持连接特性
        mlngINetConn = InternetConnect(mlngINet, strPath, strPort, strUser, strUserPwd, INTERNET_SERVICE_FTP, INTERNET_FLAG_KEEP_CONNECTION Or INTERNET_FLAG_PASSIVE, 0)
        If mlngINetConn = 0 Then
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
    mblnOpen = True
    ToConnect = True
    Exit Function
errH:
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
    mblnOpen = False
    mintConnTimes = mintConnTimes + 1
    Call CloseConnect
    If Me.ConnectType = SCT_Share Then '检查服务器是否正常并连接
        If Not mobjFSO.FolderExists(Me.ServerPath) Then
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
                Exit Function
            End If
        End If
    Else '检查是否能正常连接FTP服务器
        mlngINet = InternetOpen("FTP Control", INTERNET_OPEN_TYPE_DIRECT, vbNullString, vbNullString, 0)
        If mlngINet <= 0 Then
            Exit Function
        End If
        '自动切换到被动FTP语义，同时新增保持连接特性
        mlngINetConn = InternetConnect(mlngINet, Me.ServerPath, Me.ServerPort, Me.ConnUser, Me.ConnPassword, INTERNET_SERVICE_FTP, INTERNET_FLAG_KEEP_CONNECTION Or IIf(mintConnTimes Mod 2 = 1, INTERNET_FLAG_PASSIVE, 0), 0)
        If mlngINetConn = 0 Then
            Call InternetCloseHandle(mlngINet)
            Exit Function
        End If
    End If
    mblnOpen = True
    ReConnect = True
    Exit Function
errH:
    ReConnect = False
End Function

Public Function CloseConnect() As Boolean
'功能：关闭连接
'      strPath=关闭其他连接的路径
    Dim lngReturn As Long
    mblnOpen = False
    On Error Resume Next
    If Me.ConnectType = SCT_Share Then
        lngReturn = WNetCancelConnection(Me.ServerPath, CONNECT_UPDATE_PROFILE, True)
        If lngReturn <> NO_ERROR Then
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

Public Function DownloadFile(ByVal strFileName As String, ByVal strDownPath As String) As Boolean
'功能：升级部件
'参数：strFileName=该文件的文件名,\
'      strDownPath=下载地址
'返回：升级是否成功。
    Dim blnDowload As Boolean
    Dim lngReturn As Long, strErr As String
    '先检测服务器文件是否存在，不存在则退出
    If Not IsServerFileExists(strFileName) Then Exit Function
    '进行服务器文件下载
    On Error Resume Next
    '先清理本地文件
    If mobjFSO.FileExists(strDownPath) Then
        If FileSystem.GetAttr(strDownPath) <> vbNormal Then
             Call FileSystem.SetAttr(strDownPath, vbNormal)
        End If
        Call mobjFSO.DeleteFile(strDownPath)
    End If
    
    If Err.Number <> 0 Then Err.Clear
    If Me.ConnectType = SCT_FTP Then
        DoEvents
        lngReturn = FtpGetFile(mlngINetConn, strFileName, strDownPath, False, 0, FTP_TRANSFER_TYPE_BINARY, 0)
        blnDowload = lngReturn <> 0 '必须这样处理，直接通过返回布尔型，该布尔型的为True时，not 仍为True,原因未知
        If Not blnDowload Then
            strErr = GetLastDllErr
            If ReConnect() Then
                DoEvents
                lngReturn = FtpGetFile(mlngINetConn, strFileName, strDownPath, False, 0, FTP_TRANSFER_TYPE_BINARY, 0)
                blnDowload = lngReturn <> 0 '必须这样处理，直接通过返回布尔型，该布尔型的为True时，not 仍为True,原因未知
                If Not blnDowload Then strErr = GetLastDllErr
            End If
        End If
    Else
        DoEvents
        Call mobjFSO.CopyFile(Me.ServerPath & "\" & strFileName, strDownPath, True)
        If Err.Number <> 0 Then
            strErr = Err.Description
            Err.Clear
            blnDowload = False
        Else
            blnDowload = True
        End If
    End If
    DownloadFile = blnDowload
End Function

Public Function IsServerFileExists(ByVal strFileName As String) As Boolean
'功能：判断服务器文件是否存在
    Dim udtFindData As WIN32_FIND_DATA, lnghFile    As Long
    Dim strFile    As String

    If Me.ConnectType = SCT_Share Then
        IsServerFileExists = mobjFSO.FileExists(Me.ServerPath & "\" & strFileName)
    Else
        lnghFile = FtpFindFirstFile(mlngINetConn, strFileName, udtFindData, INTERNET_FLAG_RELOAD Or INTERNET_FLAG_NO_CACHE_WRITE, 0)
        strFile = Left(udtFindData.cFileName, InStr(udtFindData.cFileName, vbNullChar) - 1)
        IsServerFileExists = strFile <> ""
        If lnghFile <> 0 Then
            Call InternetCloseHandle(lnghFile)
        End If
    End If
End Function

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
    End If
    
    If lngEnum > 0 Then
        lngReturn = WNetCloseEnum(lngEnum)
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
            blnReturn = False
        End If
    Next
    CloseNetConn = blnReturn
End Function
