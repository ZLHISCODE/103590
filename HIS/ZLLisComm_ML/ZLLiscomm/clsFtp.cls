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
''''''''''''''''''''''''''FTP��API����'''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                                '''''''''''''''''''''
                                ''''FTP���Ӳ���'''''''
                                '''''''''''''''''''''
'��һ�������������͵�Internet����
Private Declare Function InternetConnect Lib "wininet.dll" Alias "InternetConnectA" (ByVal hInternetSession As Long, ByVal sServerName As String, ByVal nServerPort As Integer, ByVal sUsername As String, ByVal sPassword As String, ByVal lService As Long, ByVal lFlags As Long, ByVal lContext As Long) As Long
'hInternetSession--����InternetOpen������Internet�Ի����ص�ֵ
'sServerName--Ҫ���ӵķ����������ƻ�IP
'nServerPort--�����ӵ�Internet�˿�
'sUsername--��¼���û��ʺ�
'sPassword--��¼�Ŀ���
'lService--Ҫ���ӵķ��������ͣ�����������FTP�����������ӵ�����Ϊ����INTERNET_SERVICE_FTP��

Private Declare Function InternetOpen Lib "wininet.dll" Alias "InternetOpenA" (ByVal sAgent As String, ByVal lAccessType As Long, ByVal sProxyName As String, ByVal sProxyBypass As String, ByVal lFlags As Long) As Long

Private Declare Function InternetCloseHandle Lib "wininet.dll" (ByVal hInet As Long) As Integer

'����Internet����ĳ���
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
                                '''''''FTPĿ¼����''''''
                                '''''''''''''''''''''''
'��ftp�������ϴ���Ŀ¼
Private Declare Function FtpCreateDirectory Lib "wininet.dll" Alias "FtpCreateDirectoryA" (ByVal hFtpSession As Long, ByVal lpszDirectory As String) As Boolean
'lpszDirectory--����Ҫ����Ŀ¼���ַ�����������һ�����·�������·��
 '���ݴ�internet���ӵĺ���internetopen�������صľ�����û���������
        '����ftp������������
        
Private Declare Function FtpSetCurrentDirectory Lib "wininet.dll" Alias "FtpSetCurrentDirectoryA" (ByVal hFtpSession As Long, ByVal lpszDirectory As String) As Boolean

Private Declare Function FtpRemoveDirectory Lib "wininet.dll" Alias "FtpRemoveDirectoryA" (ByVal hFtpSession As Long, ByVal lpszDirectory As String) As Boolean


                                ''''''''''''''''''''''''
                                '''''''FTP�ļ�����''''''
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
'''''''''''''''''''''''''����'''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Public strIPAddress As String
'Public strUser As String
'Public strPsw As String
Public hOpen As Long
Public hConnection As Long

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''����''''''''''''''''''''''''''''''''''''''''''''''''''
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
    '''���ܣ�����FTPĿ¼
    '''������strIPAddress-----FTP����������IP��ַ
    '''������strMainPath----FTP�����е�����Ŀ¼��
    '''������strUser---FTP��½���û���
    '''������strPsw---FTP��½������
    '''������strDir---��Ҫ������Ŀ¼��ʹ�á�/��������Ŀ¼�ָ���
    '''���أ��ɹ�����0������ʧ�ܷ���1������Ŀ¼ʧ�ܷ���2
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
'    Dim hOpen As Long
'    Dim hConnection As Long
    Dim i As Integer
    Dim aNestDirs() As String               '�洢������Ҫ������Ŀ¼
    Dim bResult As Boolean                  '�������
        
    FuncFtpMkDir = 1
    '��FTP����
'    hOpen = InternetOpen("WinINet1", 1, vbNullString, vbNullString, 0)
'    hConnection = InternetConnect(hOpen, strIPAddress, 0, strUser, strPsw, INTERNET_SERVICE_FTP, 0, 255)
    If hConnection = 0 Then
        '��ftp������������ʧ�ܣ���ʾ���Ӵ�����Ϣ
        InternetCloseHandle hConnection
        InternetCloseHandle hOpen
        FuncFtpMkDir = 1
        Exit Function
    Else               '��FTP�����������ӳɹ�
        '����FTP������Ŀ¼
        If strVirtualPath <> vbNullString Then
           FtpSetCurrentDirectory hConnection, strVirtualPath
        End If
        '��ȡȫ����Ҫ������Ŀ¼��Ϣ
        aNestDirs = Split(strDir, "/")
        '����ȫ��Ŀ¼
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
    '''���ܣ�ɾ��ָ����Ŀ¼
    '''������strIPAddress---FTP�����������ƻ�IP��ַ
    '''������strVirtualPath---FTP������Ŀ¼����
    '''������strUser---FTP��½���û���
    '''������strPsw---FTP��½������
    '''������strDir---��Ҫɾ����Ŀ¼��ʹ�á�/��������Ŀ¼�ָ�����ֻɾ�����һ��Ŀ¼
    '''���أ��ɹ�����0������ʧ�ܷ���1��ɾ��ʧ�ܷ���2
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'    Dim hOpen As Long
'    Dim hConnection As Long
    Dim bResult As Boolean      '���淵��ֵ
    
    FuncFtpDelDir = 1
    '��FTP����
'    hOpen = InternetOpen("WinINet1", 1, vbNullString, vbNullString, 0)
'    hConnection = InternetConnect(hOpen, strIPAddress, 0, strUser, strPsw, INTERNET_SERVICE_FTP, 0, 255)
    If hConnection = 0 Then
        '��ftp������������ʧ�ܣ���ʾ���Ӵ�����Ϣ
        InternetCloseHandle hConnection
        InternetCloseHandle hOpen
        FuncFtpDelDir = 1
        Exit Function
    Else               '��FTP�����������ӳɹ�
         '����FTP������Ŀ¼
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
    '''���ܣ���FTP��������ָ��Ŀ¼�ϴ�һ���ļ�
    '''������strVirtualPath---FTP����Ŀ¼
    '''������strLocalFileName---�����ļ�����·�����硰C:\SSS.TXT��
    '''������strRemoteFileName---�ϴ�����ļ���
    '''���أ��ɹ�����0������ʧ�ܷ���1���ϴ��ļ�ʧ�ܷ���2
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'    Dim hOpen As Long
'    Dim hConnection As Long
    Dim bResult As Boolean
    
    FuncUploadFile = 1
    '��FTP����
'    hOpen = InternetOpen("WinINet1", 1, vbNullString, vbNullString, 0)
'    hConnection = InternetConnect(hOpen, strIPAddress, 0, strUser, strPsw, INTERNET_SERVICE_FTP, 0, 255)
    If hConnection = 0 Then
        '��ftp������������ʧ�ܣ���ʾ���Ӵ�����Ϣ
        InternetCloseHandle hConnection
        InternetCloseHandle hOpen
        FuncUploadFile = 1
        Exit Function
    Else               '��FTP�����������ӳɹ�
         '����FTP������Ŀ¼
        If strVirtualPath <> vbNullString Then
           FtpSetCurrentDirectory hConnection, strVirtualPath
        End If
        bResult = FtpPutFile(hConnection, strLocalFileName, strRemoteFileName, 2, 0)
    End If
'    InternetCloseHandle hConnection
'    InternetCloseHandle hOpen
    FuncUploadFile = IIf(bResult, 0, 2)
End Function

Public Function FuncChangeDir(strVirtualPath As String) As Long
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''���ܣ�FTP��������ָ��Ŀ¼����һ���ļ�
    '''������strVirtualPath---FTP����Ŀ¼
    '''���أ��ɹ�����0��ת��Ŀ¼ʧ�ܷ���2
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim bResult As Boolean
    
    FuncChangeDir = 1
    If hConnection = 0 Then
        '��ftp������������ʧ�ܣ���ʾ���Ӵ�����Ϣ
        InternetCloseHandle hConnection
        InternetCloseHandle hOpen
        FuncChangeDir = 1
        Exit Function
    Else               '��FTP�����������ӳɹ�
         '����FTP������Ŀ¼
        If strVirtualPath <> vbNullString Then
           bResult = FtpSetCurrentDirectory(hConnection, strVirtualPath)
        End If
    End If
    FuncChangeDir = IIf(bResult, 0, 2)
End Function

Public Function FuncDownloadOneFile(strLocalFileName As String, strRemoteFileName As String) As Long
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''���ܣ�FTP��������ָ��Ŀ¼����һ���ļ�
    '''������strLocalFileName---�����ļ�����·�����硰C:\SSS.TXT��
    '''������strRemoteFileName---��Ҫ���ص��ļ���
    '''���أ��ɹ�����0������ʧ�ܷ���1������ʧ�ܷ���2
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim bResult As Boolean
    
    FuncDownloadOneFile = 1
    If hConnection = 0 Then
        '��ftp������������ʧ�ܣ���ʾ���Ӵ�����Ϣ
        InternetCloseHandle hConnection
        InternetCloseHandle hOpen
        FuncDownloadOneFile = 1
        Exit Function
    Else               '��FTP�����������ӳɹ�
        bResult = FtpGetFile(hConnection, strRemoteFileName, strLocalFileName, False, FILE_ATTRIBUTE_NORMAL, 2, 0)
    End If
    FuncDownloadOneFile = IIf(bResult, 0, 2)
End Function

Public Function FuncDownloadFile(strVirtualPath As String, strLocalFileName As String, strRemoteFileName As String) As Long
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''���ܣ�FTP��������ָ��Ŀ¼����һ���ļ�
    '''������strVirtualPath---FTP����Ŀ¼
    '''������strLocalFileName---�����ļ�����·�����硰C:\SSS.TXT��
    '''������strRemoteFileName---��Ҫ���ص��ļ���
    '''���أ��ɹ�����0������ʧ�ܷ���1������ʧ�ܷ���2
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'    Dim hOpen As Long
'    Dim hConnection As Long
    Dim bResult As Boolean
    
    FuncDownloadFile = 1
    '��FTP����
'    hOpen = InternetOpen("WinINet1", 1, vbNullString, vbNullString, 0)
'    hConnection = InternetConnect(hOpen, strIPAddress, 0, strUser, strPsw, INTERNET_SERVICE_FTP, 0, 255)
    If hConnection = 0 Then
        '��ftp������������ʧ�ܣ���ʾ���Ӵ�����Ϣ
        InternetCloseHandle hConnection
        InternetCloseHandle hOpen
        FuncDownloadFile = 1
        Exit Function
    Else               '��FTP�����������ӳɹ�
         '����FTP������Ŀ¼
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
    '''���ܣ���FTP��������ָ��Ŀ¼�����ؼ����ļ�
    '''������strVirtualPath---FTP����Ŀ¼
    '''������strLocalPath---����·�����硰C:\��
    '''������strRemoteFileNames()---��Ҫ���ص��ļ�������
    '''���أ��ɹ�����0������ʧ�ܷ���1�������ļ�ʧ�ܷ���2
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'    Dim hOpen As Long
'    Dim hConnection As Long
    Dim i As Integer
    Dim bResult As Boolean
    
    FuncDownLoadFiles = 1
    '��FTP����
'    hOpen = InternetOpen("WinINet1", 1, vbNullString, vbNullString, 0)
'    hConnection = InternetConnect(hOpen, strIPAddress, 0, strUser, strPsw, INTERNET_SERVICE_FTP, 0, 255)
    If hConnection = 0 Then
        '��ftp������������ʧ�ܣ���ʾ���Ӵ�����Ϣ
        InternetCloseHandle hConnection
        InternetCloseHandle hOpen
        FuncDownLoadFiles = 1
        Exit Function
    Else               '��FTP�����������ӳɹ�
         '����FTP������Ŀ¼
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
    '''����: ������Ŀ¼�м����ļ�ͬʱ�ϴ���FTP��������ָ��Ŀ¼��
    '''������strVirtualPath---FTP����Ŀ¼
    '''������strLocalPath---����·�����硰C:\��
    '''������strLocalFileNames()---��Ҫ�ϴ����ļ�������
    '''���أ��ɹ�����0������ʧ�ܷ���1���ϴ��ļ�ʧ�ܷ���2
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'    Dim hOpen As Long
'    Dim hConnection As Long
    Dim i As Integer
    Dim bResult As Boolean
    
    FuncUpLoadFiles = 1
    '��FTP����
'    hOpen = InternetOpen("WinINet1", 1, vbNullString, vbNullString, 0)
'    hConnection = InternetConnect(hOpen, strIPAddress, 0, strUser, strPsw, INTERNET_SERVICE_FTP, 0, 255)
    If hConnection = 0 Then
        '��ftp������������ʧ�ܣ���ʾ���Ӵ�����Ϣ
        InternetCloseHandle hConnection
        InternetCloseHandle hOpen
        FuncUpLoadFiles = 1
        Exit Function
    Else               '��FTP�����������ӳɹ�
         '����FTP������Ŀ¼
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
    '''����: ɾ��FTP��������ָ���ļ�
    '''������strVirtualPath---FTP����Ŀ¼
    '''������strFileName()---��Ҫɾ�����ļ���
    '''���أ��ɹ�����0������ʧ�ܷ���1��ɾ���ļ�ʧ�ܷ���2
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'    Dim hOpen As Long
'    Dim hConnection As Long
    Dim bResult As Boolean
    
    FuncDelFile = 1
    '��FTP����
'    hOpen = InternetOpen("WinINet1", 1, vbNullString, vbNullString, 0)
'    hConnection = InternetConnect(hOpen, strIPAddress, 0, strUser, strPsw, INTERNET_SERVICE_FTP, 0, 255)
    If hConnection = 0 Then
        '��ftp������������ʧ�ܣ���ʾ���Ӵ�����Ϣ
        InternetCloseHandle hConnection
        InternetCloseHandle hOpen
        FuncDelFile = 1
        Exit Function
    Else               '��FTP�����������ӳɹ�
         '����FTP������Ŀ¼
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
    '''����: ɾ��FTP������ͬһĿ¼�µ�һ���ļ�
    '''������strVirtualPath---FTP����Ŀ¼
    '''������strFileNames()---��Ҫɾ�����ļ�������
    '''���أ��ɹ�����0������ʧ�ܷ���1��ɾ���ļ�ʧ�ܷ���2
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'    Dim hOpen As Long
'    Dim hConnection As Long
    Dim i As Integer
    Dim bResult As Boolean
    
    FuncDelFiles = 1
    '��FTP����
'    hOpen = InternetOpen("WinINet1", 1, vbNullString, vbNullString, 0)
'    hConnection = InternetConnect(hOpen, strIPAddress, 0, strUser, strPsw, INTERNET_SERVICE_FTP, 0, 255)
    If hConnection = 0 Then
        '��ftp������������ʧ�ܣ���ʾ���Ӵ�����Ϣ
        InternetCloseHandle hConnection
        InternetCloseHandle hOpen
        FuncDelFiles = 1
        Exit Function
    Else               '��FTP�����������ӳɹ�
         '����FTP������Ŀ¼
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
    '''����: �г�FTP������ͬһĿ¼�µ�һ���ļ�
    '''������strVirtualPath---FTP����Ŀ¼
    '''���أ��ļ������ļ���֮����|�ָ�
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
        'ɾ����.���͡�..��Ŀ¼
        If Left(Trim(FuncDirFiles), 1) = "." Then
            FuncDirFiles = Mid(Trim(FuncDirFiles), 6)
        End If
    End If
End Function

