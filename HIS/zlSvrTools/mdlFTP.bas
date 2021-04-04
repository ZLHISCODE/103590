Attribute VB_Name = "mdlFTP"
Option Explicit
'**************************
'����:FTP�Ĵ���ʽ
'��д�޸�:ף��
'**************************

Public Const FILE_ATTRIBUTE_ARCHIVE = &H20
Public Const FILE_ATTRIBUTE_COMPRESSED = &H800
Public Const FILE_ATTRIBUTE_DIRECTORY = &H10
Public Const FILE_ATTRIBUTE_HIDDEN = &H2
Public Const FILE_ATTRIBUTE_NORMAL = &H80
Public Const FILE_ATTRIBUTE_READONLY = &H1
Public Const FILE_ATTRIBUTE_SYSTEM = &H4
Public Const FILE_ATTRIBUTE_TEMPORARY = &H100

'�Զ�����������FILETIME��WIN32_FIND_DATA�Ķ���

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

'************************************************************************************˵��
'''����
'''InternetOpen��ʼ��WinINet������HINTERNET handles
'''InternetConnect����Internet���ӣ���FTP��Gopher����HTTP�Ự������HINTERNET handles
'''InternetCloseHandle�ر�Internet����
'''
'''
'''Ŀ¼����
'''FtpCreateDirectory��FTP�������Ͻ���Ŀ¼�� ��ҪInternetConnect���صĻỰ���
'''FtpRemoveDirectory��FTP��������ɾ��Ŀ¼�� ��ҪInternetConnect���صĻỰ���
'''FtpGetCurrentDirectory��ȡ��ǰ��FTP�������ϵĹ���Ŀ¼�� ��ҪInternetConnect���صĻỰ���
'''FtpSetCurrentDirectory������FTP�������ϵĹ���Ŀ¼�� ��ҪInternetConnect���صĻỰ���
'''
'''
'''�ļ�����
'''FtpFindFirstFile��FTP�������ϲ��ҷ����������ļ���Ŀ¼�� ��ҪInternetConnect���صĻỰ���
'''InternetFindNextFile��FTP�������ϼ���������һ�������������ļ���Ŀ¼����ҪFtpFindFirstFile���صĻỰ���
'''FtpPutFile�ϴ�һ���ļ���FTP�������ϣ� ��ҪInternetConnect���صĻỰ���
'''FtpGetFile��FTP������������һ���ļ��� ��ҪInternetConnect���صĻỰ���
'''FtpDeleteFile��FTP��������ɾ��һ���ļ��� ��ҪInternetConnect���صĻỰ���
'''FtpRenameFile��FTP�������ϸ���һ���ļ������֣� ��ҪInternetConnect���صĻỰ���
'''FtpOpenFile��FTP�������ϴ�һ���ļ��� ��ҪInternetConnect���صĻỰ���
'''InternetReadFileֱ����FTP�������϶�ȡ�ļ��� ��ҪFtpOpenFile���صĻỰ���
'''InternetWriteFileֱ����FTP��������д���ļ��� ��ҪFtpOpenFile���صĻỰ���
'************************************************************************************

'************************************************************************************
'������Internet�ĻỰ InternetOpen
'************************************************************************************
Public Declare Function InternetOpen Lib "wininet.dll" Alias "InternetOpenA" _
   (ByVal sAgent As String, ByVal lAccessType As Long, ByVal sProxyName As String, _
   ByVal sProxyBypass As String, ByVal lFlags As Long) As Long
'************************************************************************************
'˵����
'************************************************************************************
'    sAgent--Ҫ����Internet�Ի���Ӧ�ó�����
'    lAccessType--�����������ʵ����ͣ�������
'************************************************************************************
'        ����                                                          ֵ         ˵��
'        INTERNET_OPEN_TYPE_PRECONFIG        0          Ԥ���ã�ȱʡ��
'        INTERNET_OPEN_TYPE_DIRECT               1          ֱ�����ӵ�Internet
'        INTERNET_OPEN_TYPE_PROXY                3          ͨ���������������
'************************************************************************************
'    ��ע�����lAccessType����ΪINTERNET_OPEN_TYPE_PRECONFIG������ʱ��Ҫ����
'    HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Internet Settings
'    ע���·���µ�ע�����ֵProxyEnable��ProxyServer�� ProxyOverride
'************************************************************************************
'    sProxyName--ָ����������������֣�������������ΪINTERNET_OPEN_TYPE_PROXY����Ч
'    sProxyBypass--ָ����������������ֻ��ַ�������ô���ʱlpszProxyNameָ���Ľ�ʧЧ
'    lFlags--�Ự��ѡ��ɰ�������ֵ��
'************************************************************************************
'        ����                                                         ֵ          ˵��
'        INTERNET_FLAG_DONT_CACHE                           �������ݽ��б��ػ����ͨ�����ط���������
'        INTERNET_FLAG_ASYNC                                      ʹ���첽����
'        INTERNET_FLAG_OFFLINE                                   ֻͨ�����û���������ز���
'************************************************************************************
'��������ֵ�������������ʧ�ܣ�lngINet Ϊ0��
'************************************************************************************

'************************************************************************************
'����Internet���ӣ���FTP�Ự
'************************************************************************************
Public Declare Function InternetConnect Lib "wininet.dll" Alias "InternetConnectA" _
    (ByVal hInternetSession As Long, ByVal sServerName As String, _
    ByVal nServerPort As Integer, ByVal sUsername As String, _
    ByVal sPassword As String, ByVal lService As Long, _
    ByVal lFlags As Long, ByVal lContext As Long) As Long
'************************************************************************************
'˵����
'************************************************************************************
'    hInternetSession--����InternetOpen���ص�Internet�Ự���
'    sServerName--Ҫ���ӵķ����������ƻ�IP
'    nServerPort--Ҫ���ӵ�Internet�˿�
'    sUsername--��¼���û��ʺ�
'    sPassword--��¼�Ŀ���
'    lService--Ҫ���ӵķ��������ͣ�����������FTP�����������ӵ�����Ϊ����INTERNET_SERVICE_FTP��
'    lFlags--�������x8000000�����ӽ�ʹ�ñ���FTP���壬����0ʹ�÷Ǳ�������
'    lContext--��ʹ�ûص�����ʱʹ�øò�������ʹ�ûص����񴫵�0
'************************************************************************************
'��������ֵ�������������ʧ�ܣ�lngINetConn Ϊ0
'************************************************************************************

'************************************************************************************
'��FTP������������һ���ļ�
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
'˵����
'************************************************************************************
'    hFtpSession--����InternetConnect���ص�Internet���Ӿ��
'    lpszRemoteFile--��Ҫ��õ�FTP�������ϵ��ļ���
'    lpszNewFile--Ҫ�����ڱ��ػ����е��ļ���
'    fFailIfExists--0���滻�����ļ�����1 ����������ļ��Ѿ����������ʧ�ܣ���
'    dwFlagsAndAttributes--����ָ�������ļ����ļ����ԣ�����0����
'    dwFlags--�ļ��Ĵ��䷽ʽ���ܰ�������ֵ��
'************************************************************************************
'        ����                                                         ֵ          ˵��
'        FTP_TRANSFER_TYPE_ASCII                   1           ��ASCII �����ļ���A�ഫ�䷽����
'        FTP_TRANSFER_TYPE_BINARY                 2           �ö����ƴ����ļ���B�ഫ�䷽����
'************************************************************************************
'    dwContext--Ҫȡ�ص��ļ����������ʶ��
'************************************************************************************
'��������ֵ�������������ʧ�ܣ�blnRC ΪFALSE
'************************************************************************************
Public Declare Function FtpPutFile Lib "wininet.dll" Alias "FtpPutFileA" _
(ByVal hConnect As Long, ByVal lpszLocalFile As String, _
ByVal lpszNewRemoteFile As String, ByVal dwFlags As Long, _
ByVal dwContext As Long) As Boolean


'************************************************************************************
'�ر�Internet����
'************************************************************************************
Public Declare Function InternetCloseHandle Lib "wininet.dll" (ByVal hInet As Long) As Integer
'************************************************************************************
'˵����
'************************************************************************************
'hInet--Ҫ�رյĻỰ��InternetOpen�������ӣ�InternetConnect�����
'************************************************************************************
'��������ֵ��
'************************************************************************************

'************************************************************************************
'˵��������Ŀ¼
'************************************************************************************
Public Declare Function FtpCreateDirectory Lib "wininet.dll" Alias "FtpCreateDirectoryA" _
    (ByVal hFtpSession As Long, ByVal lpszDirectory As String) As Boolean
    
'************************************************************************************
'˵������õ�ǰĿ¼
'************************************************************************************
Public Declare Function FtpGetCurrentDirectory Lib "wininet.dll" Alias "FtpGetCurrentDirectoryA" _
    (ByVal hFtpSession As Long, ByVal lpszCurrentDirectory As String, ByRef lpdwCurrentDirectory As Long) As Boolean

'************************************************************************************
'˵������õ�ǰĿ¼
'************************************************************************************
Public Declare Function FtpSetCurrentDirectory Lib "wininet.dll" Alias "FtpSetCurrentDirectoryA" _
    (ByVal hFtpSession As Long, ByVal lpszCurrentDirectory As String) As Boolean

'************************************************************************************
'˵����ɾ���ļ�
'************************************************************************************
Public Declare Function FtpDeleteFile Lib "wininet.dll" Alias "FtpDeleteFileA" _
    (ByVal hFtpSession As Long, ByVal lpszFileName As String) As Boolean
    
'************************************************************************************
'˵�����ļ�����
'************************************************************************************
Public Declare Function FtpRenameFile Lib "wininet.dll" Alias "FtpRenameFileA" _
    (ByVal hFtpSession As Long, ByVal lpszExisting As String, ByVal lpszNew As String) As Boolean
    
    
'************************************************************************************
'˵���������ļ�
'************************************************************************************
Public Declare Function FtpFindFirstFile Lib "wininet.dll" Alias "FtpFindFirstFileA" _
    (ByVal hFtpSession As Long, ByVal lpszSearchFile As String _
    , ByVal lpFindFileData As WIN32_FIND_DATA, ByVal dwFlags As Long _
    , ByVal dwContent As Long) As Long
    
    
'************************************************************************************
'˵����������һ���ļ�
'************************************************************************************
Public Declare Function InternetFindNextFile Lib "wininet.dll" Alias "InternetFindNextFileA" _
    (ByVal hFind As Long, lpvFindData As WIN32_FIND_DATA) As Long

'************************************************************************************
'˵�������Ҵ��ļ�
'************************************************************************************
Public Declare Function FtpOpenFile Lib "wininet.dll" Alias "FtpOpenFileA" _
    (ByVal hFtpSession As Long, ByVal sFileName As String _
    , ByVal lFlags As Long, ByVal lContext As Long) As Long


'************************************************************************************
'˵�������Ҵ��ļ�
'************************************************************************************
Public Declare Function InternetReadFile Lib "wininet.dll" Alias "InternetReadFileA" _
    (ByVal hFile As Long, ByVal lpBuffer As String _
    , ByVal dwNumberOfBytesToRead As Long, ByVal lpNumberOfBytes As Long) As Long
    
'************************************************************************************
'˵������ȡ���һ��������Ϣ
'************************************************************************************
Public Declare Function InternetGetLastResponseInfo Lib "wininet.dll" Alias "InternetGetLastResponseInfoA" _
    (ByVal lpdwError As Long, ByVal lpszBuffer As String _
    , ByVal lpdwBufferLength As Long) As Boolean
    
    
'************************************************************************************
'��������
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
    'INTERNET_FLAG_DONT_CACHE ������������ֱ�Ӵ��ļ�·����ȡ�������ǻ��棬����һЩ��������
    Dim blnOK As Boolean
    blnOK = FtpGetFile(lngINetConn, srcPath, descPath, False, 0, FTP_TRANSFER_TYPE_BINARY Or INTERNET_FLAG_DONT_CACHE, 0) <> 0
    FtpDownFile = blnOK
End Function

Public Function FtpupFile(ByVal srcPath As String, ByVal descPath As String) As Boolean
    FtpupFile = FtpPutFile(lngINetConn, srcPath, descPath, FTP_TRANSFER_TYPE_BINARY, 0)
End Function

Public Function IsFtpServer(ByVal strServerPath As String, ByVal strVisitUser As String, ByVal strVisitPassWord As String, ByVal strVisitPort As String) As Boolean
    '----------------------------------------------------------------------------------------------------------
    '--����:����Ƿ�����������FTP������
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
