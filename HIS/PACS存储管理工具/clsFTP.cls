VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFTP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''变量'''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public strIPAddress As String
Public strUser As String
Public strPsw As String

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''函数''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

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
    
    Dim hOpen As Long
    Dim hConnection As Long
    Dim i As Integer
    Dim aNestDirs() As String               '存储所有需要创建的目录
    Dim bResult As Boolean                  '操作结果
        
    FuncFtpMkDir = 1
    '打开FTP连接
    hOpen = InternetOpen("WinINet1", 1, vbNullString, vbNullString, 0)
    hConnection = InternetConnect(hOpen, strIPAddress, 0, strUser, strPsw, INTERNET_SERVICE_FTP, 0, 255)
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
    InternetCloseHandle hConnection
    InternetCloseHandle hOpen
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
    Dim hOpen As Long
    Dim hConnection As Long
    Dim bResult As Boolean      '保存返回值
    
    FuncFtpDelDir = 1
    '打开FTP连接
    hOpen = InternetOpen("WinINet1", 1, vbNullString, vbNullString, 0)
    hConnection = InternetConnect(hOpen, strIPAddress, 0, strUser, strPsw, INTERNET_SERVICE_FTP, 0, 255)
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
    InternetCloseHandle hConnection
    InternetCloseHandle hOpen
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
    Dim hOpen As Long
    Dim hConnection As Long
    Dim bResult As Boolean
    
    FuncUploadFile = 1
    '打开FTP连接
    hOpen = InternetOpen("WinINet1", 1, vbNullString, vbNullString, 0)
    hConnection = InternetConnect(hOpen, strIPAddress, 0, strUser, strPsw, INTERNET_SERVICE_FTP, 0, 255)
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
    InternetCloseHandle hConnection
    InternetCloseHandle hOpen
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
    Dim hOpen As Long
    Dim hConnection As Long
    Dim bResult As Boolean
    
    FuncDownloadFile = 1
    '打开FTP连接
    hOpen = InternetOpen("WinINet1", 1, vbNullString, vbNullString, 0)
    hConnection = InternetConnect(hOpen, strIPAddress, 0, strUser, strPsw, INTERNET_SERVICE_FTP, 0, 255)
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
    InternetCloseHandle hConnection
    InternetCloseHandle hOpen
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
    Dim hOpen As Long
    Dim hConnection As Long
    Dim i As Integer
    Dim bResult As Boolean
    
    FuncDownLoadFiles = 1
    '打开FTP连接
    hOpen = InternetOpen("WinINet1", 1, vbNullString, vbNullString, 0)
    hConnection = InternetConnect(hOpen, strIPAddress, 0, strUser, strPsw, INTERNET_SERVICE_FTP, 0, 255)
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
                InternetCloseHandle hConnection
                InternetCloseHandle hOpen
                FuncDownLoadFiles = 2
                Exit Function
            End If
        Next
    End If
    InternetCloseHandle hConnection
    InternetCloseHandle hOpen
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
    Dim hOpen As Long
    Dim hConnection As Long
    Dim i As Integer
    Dim bResult As Boolean
    
    FuncUpLoadFiles = 1
    '打开FTP连接
    hOpen = InternetOpen("WinINet1", 1, vbNullString, vbNullString, 0)
    hConnection = InternetConnect(hOpen, strIPAddress, 0, strUser, strPsw, INTERNET_SERVICE_FTP, 0, 255)
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
                InternetCloseHandle hConnection
                InternetCloseHandle hOpen
                FuncUpLoadFiles = 2
                Exit Function
            End If
        Next
    End If
    InternetCloseHandle hConnection
    InternetCloseHandle hOpen
    FuncUpLoadFiles = 0
End Function


Public Function FuncDelFile(strVirtualPath As String, strFileName As String) As Long
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''功能: 删除FTP服务器的指定文件
    '''参数：strVirtualPath---FTP虚拟目录
    '''参数：strFileName()---需要删除的文件名
    '''返回：成功返回0，连接失败返回1，删除文件失败返回2
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim hOpen As Long
    Dim hConnection As Long
    Dim bResult As Boolean
    
    FuncDelFile = 1
    '打开FTP连接
    hOpen = InternetOpen("WinINet1", 1, vbNullString, vbNullString, 0)
    hConnection = InternetConnect(hOpen, strIPAddress, 0, strUser, strPsw, INTERNET_SERVICE_FTP, 0, 255)
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
    InternetCloseHandle hConnection
    InternetCloseHandle hOpen
    FuncDelFile = IIf(bResult, 0, 2)
End Function


Public Function FuncDelFiles(strVirtualPath As String, strFileNames() As String) As Long
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''功能: 删除FTP服务器同一目录下的一组文件
    '''参数：strVirtualPath---FTP虚拟目录
    '''参数：strFileNames()---需要删除的文件名数组
    '''返回：成功返回0，连接失败返回1，删除文件失败返回2
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim hOpen As Long
    Dim hConnection As Long
    Dim i As Integer
    Dim bResult As Boolean
    
    FuncDelFiles = 1
    '打开FTP连接
    hOpen = InternetOpen("WinINet1", 1, vbNullString, vbNullString, 0)
    hConnection = InternetConnect(hOpen, strIPAddress, 0, strUser, strPsw, INTERNET_SERVICE_FTP, 0, 255)
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
                InternetCloseHandle hConnection
                InternetCloseHandle hOpen
                FuncDelFiles = 2
                Exit Function
            End If
        Next
    End If
    InternetCloseHandle hConnection
    InternetCloseHandle hOpen
    FuncDelFiles = 0
End Function
Public Function FuncFileExist(ByVal strVirtualPath As String, ByVal strFileName As String) As Boolean
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''功能：查找FTP服务器上的指定文件是否存在
    '''参数：strVirtualPath---FTP虚拟目录
    '''参数：strFileName---文件名
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim hOpen As Long
    Dim hConnection As Long, hFile As Long
    
    '打开FTP连接
    hOpen = InternetOpen("WinINet1", 1, vbNullString, vbNullString, 0)
    hConnection = InternetConnect(hOpen, strIPAddress, 0, strUser, strPsw, INTERNET_SERVICE_FTP, 0, 255)
    If hConnection = 0 Then
        '与ftp服务器的连接失败，显示连接错误信息
        InternetCloseHandle hConnection
        InternetCloseHandle hOpen
        FuncFileExist = False
        Exit Function
    Else               '与FTP服务器的连接成功
         '设置FTP的虚拟目录
        If strVirtualPath <> vbNullString Then
           Call FtpSetCurrentDirectory(hConnection, strVirtualPath)
        End If
        hFile = FtpOpenFile(hConnection, strFileName, GENERIC_READ, FTP_TRANSFER_TYPE_ASCII, 0)
        
        FuncFileExist = IIf(hFile > 0, True, False)
        If hFile > 0 Then InternetCloseHandle hFile
    End If
    InternetCloseHandle hConnection
    InternetCloseHandle hOpen
End Function
Public Function FuncSearchFiles(ByVal strVirtualPath As String, ByVal strSearchFiles As String) As String
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''功能: 从指定目录下查找文件
    '''参数：strVirtualPath---FTP虚拟目录
    '''返回：文件路径－文件名；路径|文件名；路径...
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim pData As WIN32_FIND_DATA
    Dim lngHINet As Long, hOpen As Long, hConnection As Long
    Dim strTemp As String
    Dim blnRC As Boolean
    Dim aStock() As String, iPoint As Integer, strTmpName As String, strCurrFolder As String
    
    '格式化参数
    strSearchFiles = ";" & strSearchFiles & ";"
    
    pData.cFileName = String(MAX_FILENAME, 0)
    FuncSearchFiles = ""
    
    '初始堆栈和指针
    ReDim Preserve aStock(0) As String
    aStock(0) = strVirtualPath: iPoint = 0
    
    hOpen = InternetOpen("WinINet1", 1, vbNullString, vbNullString, 0)
    hConnection = InternetConnect(hOpen, strIPAddress, 0, strUser, strPsw, INTERNET_SERVICE_FTP, 0, 255)
    Do While iPoint >= 0
        strCurrFolder = aStock(iPoint)
        lngHINet = FtpFindFirstFile(hConnection, strCurrFolder & "/*.*", pData, 0, 0)
        iPoint = iPoint - 1
        
        If lngHINet > 0 Then
            strTmpName = Left(pData.cFileName, InStr(1, pData.cFileName, String(1, 0), vbBinaryCompare) - 1)
            Do While True
                If pData.dwFileAttributes = ATTR_DIR Then
                    '目录，压入堆栈并指向它
                    iPoint = iPoint + 1
                    ReDim Preserve aStock(iPoint) As String
                    aStock(iPoint) = strCurrFolder & "/" & strTmpName
                Else
                    '文件，进行查找
                    If InStr(strSearchFiles, ";" & strTmpName & ";") > 0 Then
                        FuncSearchFiles = FuncSearchFiles & "|" & strTmpName & ";" & strCurrFolder
                        strSearchFiles = Replace(strSearchFiles, strTmpName & ";", "")
                        
                        If Replace(strSearchFiles, ";", "") = "" Then
                            '文件全部搜索完，则退出循环并弹出堆栈
                            iPoint = -1
                            Exit Do
                        End If
                    End If
                End If
                '浏览下一个目录/文件
                pData.cFileName = String(MAX_FILENAME, 0)
                blnRC = InternetFindNextFile(lngHINet, pData)
                If Not blnRC Then
                    Exit Do
                Else
                    strTmpName = Left(pData.cFileName, InStr(1, pData.cFileName, String(1, 0), vbBinaryCompare) - 1)
                End If
            Loop
        
            InternetCloseHandle lngHINet
        End If
    Loop
    
    InternetCloseHandle hConnection
    InternetCloseHandle hOpen
    
    If Len(FuncSearchFiles) > 0 Then FuncSearchFiles = Mid(FuncSearchFiles, 2)
End Function
'将一个目录的文件复制到另外的目录中
Public Function FuncDupDir(ByVal strSrcPath As String, ByVal strDestPath As String, Optional ByVal blnRemoveSrc As Boolean = False) As Boolean
    Dim strTmpFile As String
    Dim aFiles() As String, i As Integer
    
    FuncDupDir = True
    If strSrcPath = strDestPath Then Exit Function
    
    aFiles = Split(FuncDirFiles(strSrcPath), "|")
    For i = 0 To UBound(aFiles)
        strTmpFile = App.Path & "\" & aFiles(i)
        Call FuncDownloadFile(strSrcPath, strTmpFile, aFiles(i))
        Call FuncUploadFile(strDestPath, strTmpFile, aFiles(i))
        
        If Dir(strTmpFile, vbDirectory) <> "" Then
            Kill strTmpFile
        End If
        
        If blnRemoveSrc Then Call FuncDelFile(strSrcPath, aFiles(i))
    Next
End Function
Public Function FuncDirFiles(ByVal strVirtualPath As String) As String
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''功能: 列出FTP服务器同一目录下的一组文件
    '''参数：strVirtualPath---FTP虚拟目录
    '''返回：文件串，文件名之间以|分隔
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim pData As WIN32_FIND_DATA
    Dim lngHINet As Long, hOpen As Long, hConnection As Long
    Dim strTemp As String
    Dim blnRC As Boolean
    
    pData.cFileName = String(MAX_FILENAME, 0)
    FuncDirFiles = ""
    
    hOpen = InternetOpen("WinINet1", 1, vbNullString, vbNullString, 0)
    hConnection = InternetConnect(hOpen, strIPAddress, 0, strUser, strPsw, INTERNET_SERVICE_FTP, 0, 255)
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
        Loop
    
        InternetCloseHandle lngHINet
    End If
    InternetCloseHandle hConnection
    InternetCloseHandle hOpen
End Function
