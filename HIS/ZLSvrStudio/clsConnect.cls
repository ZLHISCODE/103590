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
'                0��ģ������빫�����͡�API
'===========================================================================
Private Declare Function InternetOpen Lib "wininet.dll" Alias "InternetOpenA" (ByVal sAgent As String, ByVal lAccessType As Long, ByVal sProxyName As String, ByVal sProxyBypass As String, ByVal lFlags As Long) As Long
'���ܣ�������Internet�ĻỰ
'˵����
'    sAgent--Ҫ����Internet�Ի���Ӧ�ó�����
'    lAccessType--�����������ʵ�����
'��ע�����lAccessType����ΪINTERNET_OPEN_TYPE_PRECONFIG������ʱ��Ҫ����
'    HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Internet Settings
'    ע���·���µ�ע�����ֵProxyEnable��ProxyServer�� ProxyOverride
'    sProxyName--ָ�����������������֣�������������ΪINTERNET_OPEN_TYPE_PROXY����Ч
'    sProxyBypass--ָ�����������������ֻ��ַ�������ô���ʱlpszProxyNameָ���Ľ�ʧЧ
'��������ֵ�������������ʧ�ܣ�lngINet Ϊ0��
'    lFlags-
Private Const INTERNET_FLAG_ASYNC           As Long = &H10000000    'ʹ���첽����
Private Const INTERNET_FLAG_FROM_CACHE      As Long = &H1000000
Private Const INTERNET_FLAG_OFFLINE         As Long = &H1000000     'ֻͨ�����û���������ز���
'   lAccessType
Private Const INTERNET_OPEN_TYPE_PRECONFIG  As Long = &H0           ' use registry configuration
Private Const INTERNET_OPEN_TYPE_DIRECT     As Long = &H1           'direct to net

Private Declare Function InternetConnect Lib "wininet.dll" Alias "InternetConnectA" (ByVal hInternetSession As Long, ByVal sServerName As String, ByVal nServerPort As Integer, ByVal sUsername As String, ByVal sPassword As String, ByVal lService As Long, ByVal lFlags As Long, ByVal lContext As Long) As Long
'���ܣ�����Internet���ӣ���FTP�Ự
'˵����
'    hInternetSession--����InternetOpen���ص�Internet�Ự���
'    sServerName--Ҫ���ӵķ����������ƻ�IP
'    nServerPort--Ҫ���ӵ�Internet�˿�
'    sUsername--��¼���û��ʺ�
'    sPassword--��¼�Ŀ���
'    lService--Ҫ���ӵķ��������ͣ�����������FTP�����������ӵ�����Ϊ����INTERNET_SERVICE_FTP��
'    lFlags--�������x8000000�����ӽ�ʹ�ñ���FTP���壬����0ʹ�÷Ǳ�������
'    lContext--��ʹ�ûص�����ʱʹ�øò�������ʹ�ûص����񴫵�0
'��������ֵ�������������ʧ�ܣ�lngINetConn Ϊ0
'    lService
Private Const INTERNET_SERVICE_FTP          As Long = &H1
'    lFlags
Private Const INTERNET_FLAG_KEEP_CONNECTION  As Long = &H400000    ' use keep-alive semantics
Private Const INTERNET_FLAG_PASSIVE         As Long = &H8000000   ' used for FTP connections

Private Declare Function FtpGetFile Lib "wininet.dll" Alias "FtpGetFileA" (ByVal hFtpSession As Long, ByVal lpszRemoteFile As String, ByVal lpszNewFile As String, ByVal fFailIfExists As Boolean, ByVal dwFlagsAndAttributes As Long, ByVal dwFlags As Long, ByVal dwContext As Long) As Long
'���ܣ���FTP������������һ���ļ�
'˵����
'    hFtpSession--����InternetConnect���ص�Internet���Ӿ��
'    lpszRemoteFile--��Ҫ��õ�FTP�������ϵ��ļ���
'    lpszNewFile--Ҫ�����ڱ��ػ����е��ļ���
'    fFailIfExists--0���滻�����ļ�����1 ����������ļ��Ѿ����������ʧ�ܣ���
'    dwFlagsAndAttributes--����ָ�������ļ����ļ����ԣ�����0����
'    dwFlags--�ļ��Ĵ��䷽ʽ���ܰ�������ֵ��
'        FTP_TRANSFER_TYPE_ASCII                   1           ��ASCII �����ļ���A�ഫ�䷽����
'        FTP_TRANSFER_TYPE_BINARY                 2           �ö����ƴ����ļ���B�ഫ�䷽����
'    dwContext--Ҫȡ�ص��ļ�����������ʶ��
'��������ֵ�������������ʧ�ܣ�blnRC ΪFALSE
Private Const FTP_TRANSFER_TYPE_BINARY      As Long = &H2
Private Const INTERNET_FLAG_RELOAD          As Long = &H80000000  'retrieve the original item
Private Const INTERNET_FLAG_NO_CACHE_WRITE  As Long = &H4000000
Private Const INTERNET_FLAG_DONT_CACHE      As Long = INTERNET_FLAG_NO_CACHE_WRITE

'Private Declare Function FtpPutFile Lib "wininet.dll" Alias "FtpPutFileA" (ByVal hConnect As Long, ByVal lpszLocalFile As String, ByVal lpszNewRemoteFile As String, ByVal dwFlags As Long, ByVal dwContext As Long) As Boolean
Private Declare Function FtpFindFirstFile Lib "wininet.dll" Alias "FtpFindFirstFileA" (ByVal hFtpSession As Long, ByVal lpszSearchFile As String, lpFindFileData As WIN32_FIND_DATA, ByVal dwFlags As Long, ByVal dwContent As Long) As Long
' ����ָ����·��
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
'���ܣ��ر�Internet����
'˵����
'hInet--Ҫ�رյĻỰ��InternetOpen�������ӣ�InternetConnect�����
'��������ֵ��
'��������
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
'string�汾��NETRESOURCE�����㴫��
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

Private mlngINet                            As Long                 'FTP��ʽ����������
Private mlngINetConn                        As Long                 'FTP����
Private mintConnTimes                       As Integer              '�������Ӵ�����

Public Enum ServerConnectType
    SCT_NONE = -1 '������
    SCT_Share = 0 '�����ļ�����
    SCT_FTP = 1 'FTP����
End Enum
'===========================================================================
'                1���������Եı���
'===========================================================================
'ConnectType��ǰ��������(R/W)
Private msctCur                             As ServerConnectType
'ServerPath��ǰ������λ��(R/W)
Private mstrCurServerPath                   As String
'ConnUser���ӵ�ǰ���������û���(R/W)
Private mstrCurConnUser                     As String
'ConnPassword���ӵ�ǰ������������(R/W)
Private mstrCurConnPWD                      As String
'ServerPort���ӵ�ǰ�������Ķ˿�(R/W)������FTP��ʽ
Private mstrCurServerPort                   As String
'CollectType�ռ����ļ�����(R/W)�������ռ��ļ�
Private mstrCollectType                     As String
'IsSampleFTP�Ƿ��Ǽ�FTP(R/W),��FTP�������ļ����ڵ��ж�
Private mblnSampleFTP                       As Boolean

'===========================================================================
'                2�����Զ���
'===========================================================================
'ConnectType��ǰ��������(R/W)
Public Property Let ConnectType(ByVal sctCur As ServerConnectType)
    msctCur = sctCur
End Property

Public Property Get ConnectType() As ServerConnectType
    ConnectType = msctCur
End Property

'ServerPath��ǰ������λ��(R/W)
Public Property Let ServerPath(ByVal strCurServerPath As String)
    mstrCurServerPath = strCurServerPath
End Property

Public Property Get ServerPath() As String
    ServerPath = mstrCurServerPath
End Property

'ConnUser���ӵ�ǰ���������û���(R/W)
Public Property Let ConnUser(ByVal strCurConnUser As String)
    mstrCurConnUser = strCurConnUser
End Property

Public Property Get ConnUser() As String
    ConnUser = mstrCurConnUser
End Property
'ConnPassword���ӵ�ǰ������������(R/W)
Public Property Let ConnPassword(ByVal strCurConnPWD As String)
    mstrCurConnPWD = strCurConnPWD
End Property

Public Property Get ConnPassword() As String
    ConnPassword = mstrCurConnPWD
End Property

'ServerPort���ӵ�ǰ�������Ķ˿�(R/W)������FTP��ʽ
Public Property Let ServerPort(ByVal strCurServerPort As String)
    mstrCurServerPort = strCurServerPort
End Property

Public Property Get ServerPort() As String
    ServerPort = mstrCurServerPort
End Property
'CollectType�ռ����ļ�����(R/W)�������ռ��ļ�
Public Property Let CollectType(ByVal strCollectType As String)
    mstrCollectType = strCollectType
End Property

Public Property Get CollectType() As String
    CollectType = mstrCollectType
End Property

'IsSampleFTP�Ƿ��Ǽ�FTP(R/W),��FTP�������ļ����ڵ��ж�
Public Property Let IsSampleFTP(ByVal blnSampleFTP As Boolean)
    mblnSampleFTP = blnSampleFTP
End Property

Public Property Get IsSampleFTP() As Boolean
    IsSampleFTP = mblnSampleFTP
End Property
'===========================================================================
'                3����������
'===========================================================================
Public Function ToConnect(ByVal sctConnectType As ServerConnectType, ByVal strPath As String, strUser As String, strUserPwd As String, Optional ByVal strPort As String, Optional ByVal strCollectType As String, Optional ByVal blnSampleFTP As Boolean, Optional ByRef strMsg As String) As Boolean
'���ܣ�����һ������
'������intConnectType=�������ͣ�0-���繲�����ӣ�1-Ftp����
'      strPath=��Դλ��
'      strUser\strUserPwd=������Դ���û���������
'      strPort=������Դ�Ķ˿�
'      strCollectType=�ռ��ļ�����
'���أ��Ƿ����ӳɹ�
    Dim nrCur As NETRESOURCE_STRING
    Dim lngRet As Long
    mintConnTimes = 0
    If sctConnectType = SCT_Share Then '���������Ƿ�����������
        If gobjFSO.FolderExists(strPath) Then
            Call CloseNetConn
        End If
        With nrCur
            .dwScope = RESOURCE_GLOBALNET
            .dwType = RESOURCETYPE_DISK
            .dwDisplayType = RESOURCEDISPLAYTYPE_SHARE
            .dwUsage = RESOURCEUSAGE_CONNECTABLE
            .pLocalName = "" 'ӳ���������
            .pRemoteName = strPath  '������·��
        End With
        On Error GoTo errH
        lngRet = WNetAddConnection2(nrCur, strUserPwd, strUser, CONNECT_UPDATE_PROFILE)
        If lngRet <> NO_ERROR Then
            strMsg = GetLastDllErr(lngRet)
            '������������һ���û�ʹ��һ�������û����������������Դ�Ķ������ӡ��ж���˷�����������Դ���������ӣ�Ȼ������һ�Ρ�
            Exit Function
        End If
    Else '����Ƿ�����������FTP������
        mlngINet = InternetOpen("FTP Control", INTERNET_OPEN_TYPE_DIRECT, vbNullString, vbNullString, 0)
        If mlngINet <= 0 Then
            strMsg = GetLastDllErr()
            Exit Function
        End If
        '����������������
        mlngINetConn = InternetConnect(mlngINet, strPath, strPort, strUser, strUserPwd, INTERNET_SERVICE_FTP, INTERNET_FLAG_KEEP_CONNECTION Or INTERNET_FLAG_PASSIVE, 0)
        If mlngINetConn = 0 Then
            strMsg = GetLastDllErr()
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
    Me.IsSampleFTP = blnSampleFTP
    ToConnect = True
    Exit Function
errH:
    strMsg = err.Description
    If Not ReConnect And mintConnTimes <= 3 Then
        ToConnect = ReConnect
    Else
        ToConnect = False
    End If
End Function

Public Function ReConnect() As Boolean
'���ܣ�����һ������
'������intConnectType=�������ͣ�0-���繲�����ӣ�1-Ftp����
'      strPath=��Դλ��
'      strUser\strUserPwd=������Դ���û���������
'      strPort=������Դ�Ķ˿�
'      strCollectType=�ռ��ļ�����
'���أ��Ƿ����ӳɹ�
    Dim nrCur As NETRESOURCE_STRING
    Dim lngRet As Long
    '�л�����3�Σ������л���ֱ�ӷ���True
    If mintConnTimes > 3 Then
        ReConnect = True: Exit Function
    End If
    mintConnTimes = mintConnTimes + 1
    Call CloseConnect
    If Me.ConnectType = SCT_Share Then '���������Ƿ�����������
        If gobjFSO.FolderExists(Me.ServerPath) Then
            Call CloseNetConn
        End If
        With nrCur
            .dwScope = RESOURCE_GLOBALNET
            .dwType = RESOURCETYPE_DISK
            .dwDisplayType = RESOURCEDISPLAYTYPE_SHARE
            .dwUsage = RESOURCEUSAGE_CONNECTABLE
            .pLocalName = "" 'ӳ���������
            .pRemoteName = Me.ServerPath  '������·��
        End With
        On Error GoTo errH
        lngRet = WNetAddConnection2(nrCur, Me.ConnUser, Me.ConnPassword, CONNECT_UPDATE_PROFILE)
        If lngRet <> NO_ERROR Then
            '������������һ���û�ʹ��һ�������û����������������Դ�Ķ������ӡ��ж���˷�����������Դ���������ӣ�Ȼ������һ�Ρ�
            Exit Function
        End If
    Else '����Ƿ�����������FTP������
        mlngINet = InternetOpen("FTP Control", INTERNET_OPEN_TYPE_DIRECT, vbNullString, vbNullString, 0)
        If mlngINet <= 0 Then
            Exit Function
        End If
        '�Զ��л�������FTP���壬ͬʱ����������������
        mlngINetConn = InternetConnect(mlngINet, Me.ServerPath, Me.ServerPort, Me.ConnUser, Me.ConnPassword, INTERNET_SERVICE_FTP, INTERNET_FLAG_KEEP_CONNECTION Or IIf(mintConnTimes Mod 2 = 1, INTERNET_FLAG_PASSIVE, 0), 0)
        If mlngINetConn = 0 Then
            Call InternetCloseHandle(mlngINet)
            Exit Function
        End If
    End If
    ReConnect = True
    Exit Function
errH:
    ReConnect = False
End Function

Public Function CloseConnect() As Boolean
'���ܣ��ر�����
'      strPath=�ر��������ӵ�·��
    Dim lngReturn As Long
    
    On Error Resume Next
    If Me.ConnectType = SCT_Share Then
        lngReturn = WNetCancelConnection(Me.ServerPath, CONNECT_UPDATE_PROFILE, True)
        If lngReturn = NO_ERROR Then
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
'���ܣ��жϷ������ļ��Ƿ����
    Dim udtFindData As WIN32_FIND_DATA, lnghFile    As Long
    Dim strFindFile    As String

    If Me.ConnectType = SCT_Share Then
        IsServerFileExists = gobjFSO.FileExists(Me.ServerPath & "\" & strFileName)
    Else
        If Not Me.IsSampleFTP Then
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
'���ܣ������ļ�
'������strFileName=���ļ����ļ���
'      strToLoaclPath=���ص����ص�·��
'���أ������Ƿ�ɹ���
    Dim blnDownload      As Boolean
    Dim lngReturn       As Long
    
    If Me.ConnectType = SCT_FTP Then
        DoEvents
        strErr = ""
        lngReturn = FtpGetFile(mlngINetConn, strFileName, strToLoaclPath & "\" & strFileName, False, 0, FTP_TRANSFER_TYPE_BINARY, 0)
        blnDownload = lngReturn <> 0 '��������������ֱ��ͨ�����ز����ͣ��ò����͵�ΪTrueʱ��not ��ΪTrue,ԭ��δ֪
        If Not blnDownload Then
            strErr = GetLastDllErr
            If ReConnect() Then
                DoEvents
                strErr = ""
                lngReturn = FtpGetFile(mlngINetConn, strFileName, strToLoaclPath & "\" & strFileName, False, 0, FTP_TRANSFER_TYPE_BINARY, 0)
                blnDownload = lngReturn <> 0 '��������������ֱ��ͨ�����ز����ͣ��ò����͵�ΪTrueʱ��not ��ΪTrue,ԭ��δ֪
                If Not blnDownload Then strErr = GetLastDllErr
            End If
        End If
    Else
        DoEvents
        err.Clear: strErr = ""
        Call gobjFSO.CopyFile(Me.ServerPath & "\" & strFileName, strToLoaclPath & "\" & strFileName, True)
        If err.Number <> 0 Then
            strErr = err.Description
            err.Clear
            blnDownload = False
        Else
            blnDownload = True
        End If
    End If
    DownloadFile = blnDownload
End Function

Public Function UploadFile(ByVal strFileName As String, ByVal strFromLoaclPath As String, Optional ByRef strErr As String) As Boolean
'���ܣ��ϴ��ļ�
'������strFileName=���ļ����ļ���
'      strFromLoaclPath=���ص����ص�·��
'���أ������Ƿ�ɹ���
    Dim blnUpload      As Boolean
    Dim lngReturn       As Long
    
    If Me.ConnectType = SCT_FTP Then
        DoEvents
        strErr = ""
'        lngReturn = FtpPutFile(mlngINetConn, strFromLoaclPath & "\" & strFileName, strFileName, FTP_TRANSFER_TYPE_BINARY, 0)
        blnUpload = lngReturn <> 0 '��������������ֱ��ͨ�����ز����ͣ��ò����͵�ΪTrueʱ��not ��ΪTrue,ԭ��δ֪
        If Not blnUpload Then
            strErr = GetLastDllErr
            If ReConnect() Then
                DoEvents
                strErr = ""
'                lngReturn = FtpPutFile(mlngINetConn, strFromLoaclPath & "\" & strFileName, strFileName, FTP_TRANSFER_TYPE_BINARY, 0)
                blnUpload = lngReturn <> 0 '��������������ֱ��ͨ�����ز����ͣ��ò����͵�ΪTrueʱ��not ��ΪTrue,ԭ��δ֪
                If Not blnUpload Then strErr = GetLastDllErr
            End If
        End If
    Else
        DoEvents
        err.Clear: strErr = ""
        Call gobjFSO.CopyFile(strFromLoaclPath & "\" & strFileName, Me.ServerPath & "\" & strFileName, True)
        If err.Number <> 0 Then
            strErr = err.Description
            err.Clear
            blnUpload = False
        Else
            blnUpload = True
        End If
    End If
    UploadFile = blnUpload
End Function
'===========================================================================
'                4��˽�з���
'===========================================================================
Private Function CloseNetConn(Optional strIp As String = "", Optional strComputerName As String = "") As Boolean
'���ܣ���ֹ����������Դ����
'���أ���ֹ�ɹ�,����true,���򷵻�False
'˵����ֻҪ��һ������û��ɹ�,��Ҳ����false,���򷵻�true
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
        If lngReturn <> NO_ERROR Then
        End If
    End If
    If lngCount = 0 Then
        CloseNetConn = True
        Exit Function
    End If
    
    '��������
    blnReturn = True
    For i = 0 To UBound(uNet)
        lngReturn = WNetCancelConnection(IIf(uNet(i).sLocalName = "", uNet(i).sRemoteName, uNet(i).sLocalName), CONNECT_UPDATE_PROFILE, True)
        If lngReturn <> NO_ERROR Then
            blnReturn = False
        End If
    Next
    CloseNetConn = blnReturn
End Function