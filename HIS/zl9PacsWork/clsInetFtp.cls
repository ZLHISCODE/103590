VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsInetFTP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private WithEvents mobjInet As InetCtlsObjects.iNet
Attribute mobjInet.VB_VarHelpID = -1

Private mstrData As String
Private mstrFtpVirtualPath As String
Private mblnExecuteState As Boolean
Private mlngLastState As Long

Private mobjAryInet() As InetCtlsObjects.iNet

'Public Event OnStateChanged(ByVal State As Integer)

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)



'注：如果使用了CERN代理服务器，则不能使用execute的异步执行方式



'最后执行状态
Property Get LastState() As Long
    LastState = mlngLastState
End Property


'inet是否处于空闲状态
Property Get FreeState() As Boolean
    FreeState = Not mobjInet.StillExecuting
End Property


'获取当前ftp的虚拟目录
Property Get FtpVirtualPath() As String
    FtpVirtualPath = mstrFtpVirtualPath
End Property


'获取是否连接inet对象
Property Get IsConnectionInet() As Boolean
    IsConnectionInet = IIf(mobjInet Is Nothing, False, True)
End Property





Private Function ExecuteFtpCMD(ByVal strCMD As String) As Boolean
'执行ftp命令
    ExecuteFtpCMD = False
    
    mblnExecuteState = True
    mstrData = ""
    
    '如果当前操作尚未结束，则直接退出过程
    If mlngLastState = icReceivingResponse Then Exit Function
    If mobjInet.StillExecuting Then Exit Function
    
    Call mobjInet.Execute(, strCMD)
    
    Call WaitExecute
    
    If mblnExecuteState Then ExecuteFtpCMD = True
End Function


Private Sub WaitExecute()
'等待执行
    While mobjInet.StillExecuting
'        Call Sleep(1)
        
        DoEvents
    Wend
End Sub


Public Function FormatPath(ByVal strPath As String, ByVal blnIsFtpDir As Boolean) As String
'格式化路径
    Dim strCurPath As String
    Dim strPathChar As String
    
    FormatPath = ""
    If Trim(strPath) = "" Then Exit Function
    
    strPathChar = "\"
    If blnIsFtpDir Then strPathChar = "/"
    
    strCurPath = strPath
    
    '格式化虚拟目录，当虚拟目录前使用了"/\VirtualPath\"等形式时，将不能从ftp上获取文件信息
    If blnIsFtpDir Then
        strCurPath = Replace(strCurPath, "\", strPathChar)
    Else
        strCurPath = Replace(strCurPath, "/", strPathChar)
    End If
    
    strCurPath = Replace(strCurPath, "\\", strPathChar)
    strCurPath = Replace(strCurPath, "//", strPathChar)

    If InStrRev(strCurPath, strPathChar) <> Len(strCurPath) Then strCurPath = strCurPath & strPathChar
    
    FormatPath = strCurPath
End Function


Public Function GetFilePath(ByVal strFile As String, ByVal blnFtpFile As Boolean) As String
'获取文件的路径
    Dim strCurPath As String
    Dim strPathChar As String
    
    GetFilePath = ""
    If Trim(strFile) = "" Then Exit Function
    
    If blnFtpFile Then
        strPathChar = "/"
        strCurPath = Replace(strFile, "\", "/")
    Else
        strPathChar = "\"
        strCurPath = Replace(strFile, "/", "\")
    End If
    
    '如果不存在路径，则直接退出
    If InStr(strCurPath, strPathChar) <= 0 Then Exit Function
    
    GetFilePath = Mid(strCurPath, 1, InStrRev(strCurPath, strPathChar))
    
End Function


Public Function GetFileName(ByVal strFlie As String, ByVal blnFtpFile As Boolean) As String
'获取文件名
    Dim strCurFile As String
    
    GetFileName = ""
    If Trim(strFlie) = "" Then Exit Function
    
    strCurFile = strFlie
    
    If blnFtpFile Then
        strCurFile = Replace(strCurFile, "\", "/")
    Else
        strCurFile = Replace(strCurFile, "/", "\")
    End If
    
    If InStr(strFlie, "/") > 0 Then
        GetFileName = Mid(strCurFile, InStrRev(strCurFile, "/") + 1, 255)
    ElseIf InStr(strFlie, "\") > 0 Then
        GetFileName = Mid(strCurFile, InStrRev(strCurFile, "\") + 1, 255)
    Else
        GetFileName = strCurFile
    End If
    
End Function











Public Sub ConnectionInet(objInet As Object)
'关联Microsoft Internet Transfer control的Inet对象

    Set mobjInet = objInet
    
    mobjInet.Protocol = icFTP
    
End Sub


Public Sub InitFtp(ByVal strHost As String, ByVal lngPort As Long, _
    ByVal strUser As String, ByVal strPwd As String, Optional ByVal strVirtualPath As String = "/")
'初始化ftp连接信息
    mobjInet.RemoteHost = strHost
    mobjInet.RemotePort = lngPort
    mobjInet.UserName = strUser
    mobjInet.Password = strPwd
    
    mstrFtpVirtualPath = FormatPath(strVirtualPath, True)
    
    '切换到指定的虚拟目录中
    If mstrFtpVirtualPath <> "/" Then Call ChangeDir(mstrFtpVirtualPath)
End Sub


Public Function OpenTest() As Boolean
'打开测试
'使用PWD命令判断是否能够连接
'说明：打开ftp连接
    
    OpenTest = ExecuteFtpCMD("PWD")

End Function


Public Function ExecuteFtpCommand(ByVal strFtpCMD As String, ByRef strResult As String) As Boolean
'执行ftp命令
    Dim blnResult As Boolean
    
    ExecuteFtpCommand = False
    
    blnResult = ExecuteFtpCMD(strFtpCMD)
    
    If Not blnResult Then Exit Function
    
    strResult = mstrData
    ExecuteFtpCommand = blnResult
End Function

'Public Sub OpenFtp()
''打开ftp连接
''OPEN命令
''格式：OPEN host[port]
''说明：打开ftp连接
'
'    Call ExecuteFtpCMD("OPEN " & mobjInet.RemoteHost & " " & mobjInet.RemotePort)
'End Sub



Public Function QuitFtp() As Boolean
'关闭ftp
'QUIT命令
'格式：QUIT
'说明：结束FTP连接,不保留ftp
'Terminates the current user

    QuitFtp = ExecuteFtpCMD("QUIT")
End Function


Public Function CloseFtp() As Boolean
'关闭ftp
'CLOSE命令
'格式：CLOSE
'说明：关闭和远程主机的连接，但保留ftp
'Closes the current FTP connection.

    CloseFtp = ExecuteFtpCMD("CLOSE")
End Function


Public Sub Cancel()
'撤销操作
    Call mobjInet.Cancel
    
    Call WaitExecute
End Sub

Public Function ChangeDir(ByVal strFtpDir As String) As Boolean
'改变ftp目录
'CD命令
'格式：CD file1
'说明：改变ftp当前目录到file1所指定的目录中
    Dim strCurPath As String
    Dim blnResult As Boolean
    
    ChangeDir = True
    strCurPath = FormatPath(strFtpDir, True)
    
    '如果虚拟目录相同，则不进行切换
    If mstrFtpVirtualPath = strCurPath Then Exit Function
    
    blnResult = ExecuteFtpCMD("CD " & strCurPath)
    If blnResult Then
        '执行成功后才重新设置虚拟目录
        mstrFtpVirtualPath = strCurPath
    End If
    
    ChangeDir = blnResult
End Function


Public Function ChangeParentDir() As Boolean
'切换到上级目录
'CDUP命令
'格式：CDUP
'说明：切换到当前ftp目录的上级目录
    Dim blnResult As Boolean
    
    ChangeParentDir = True
    If mstrFtpVirtualPath = "/" Then Exit Function
    
    blnResult = ExecuteFtpCMD("CD ..")
    
    If blnResult Then
        mstrFtpVirtualPath = Mid(mstrFtpVirtualPath, 1, Len(mstrFtpVirtualPath) - 1)
        
        If InStr(mstrFtpVirtualPath, "/") > 0 Then
            mstrFtpVirtualPath = Mid(mstrFtpVirtualPath, 1, InStrRev(mstrFtpVirtualPath, "/"))
        Else
            mstrFtpVirtualPath = "/"
        End If
    End If
    
    ChangeParentDir = blnResult
End Function


Public Function ChangeRootDir() As Boolean
'切换到上级目录
'CDUP命令
'格式：CDUP
'说明：切换到当前ftp目录的上级目录
    Dim blnResult As Boolean
    
    ChangeRootDir = True
    If mstrFtpVirtualPath = "/" Then Exit Function
    
    blnResult = ExecuteFtpCMD("CD /")
    
    If blnResult Then
        mstrFtpVirtualPath = "/"
    End If
    
    ChangeRootDir = blnResult
End Function


Public Function DelFile(ByVal strFtpFile As String) As Boolean
'删除文件
'strFtpFile：要删除的ftp文件，如果指定了文件路径，则删除对应路径下的文件，否则默认删除当前ftp路径下的文件

'DELETE命令
'格式：DELETE file1
'说明：删除由file1指定的文件
    Dim strCurFile As String
    
    strCurFile = GetFilePath(strFtpFile, True)
    strCurFile = FormatPath(strCurFile, True) & GetFileName(strFtpFile, True)
    
    DelFile = ExecuteFtpCMD("DELETE " & strCurFile)
End Function



Public Function UpLoadFile(ByVal strLocalFile As String, ByVal strFtpPath As String, _
    Optional ByVal strFtpFileName As String = "") As Boolean
'上传文件
'strLocalFile：     待上传的文件全路径名
'strFtpPath：       ftp目录
'strFtpFileName：   ftp文件名

'PUT命令
'格式：PUT file1 file2
'说明：将由file1指定的文件，复制到file2指定的远程主机文件中
    Dim strCurFileName As String
    
    strCurFileName = strFtpFileName
    If strCurFileName = "" Then
        '获取本地文件名称
        strCurFileName = GetFileName(strLocalFile, False)
    End If
    
    '组织ftp的存放位置
    strCurFileName = FormatPath(strFtpPath, True) & strCurFileName

    '执行文件上传
    UpLoadFile = ExecuteFtpCMD("PUT " & strLocalFile & " " & strCurFileName)
End Function

Public Function DownLoadFile(ByVal strFtpFile As String, ByVal strLocalPath As String, _
    Optional ByVal strLocalFileName As String = "") As Boolean
'下载文件
'GET命令
'格式：GET file1 file2
'说明：从ftp中下载由file1指定的文件到file2中
    Dim strCurFileName As String
    
    strCurFileName = strLocalFileName
    If strCurFileName = "" Then
        '获取ftp的文件名称
        strCurFileName = GetFileName(strFtpFile, True)
    End If
    
    '组织本地文件的存放位置
    strCurFileName = FormatPath(strLocalPath, False) & strCurFileName
    
    DownLoadFile = ExecuteFtpCMD("GET " & strFtpFile & " " & strCurFileName)
End Function


Public Function BatchDownLoad(ByVal strFtpPath As String, ByVal strLocalPath As String) As Boolean
'批量下载目录中的文件
    Dim strCurPath As String
    Dim strFileData As String
    Dim aryFile() As String
    Dim i As Long
    Dim blnResult As Boolean
    Dim strFormatLocalPath As String
    
    BatchDownLoad = False
    If Trim(strFtpPath) = "" Then Exit Function
    
    strCurPath = FormatPath(strFtpPath, True)
    strFileData = GetFtpFileList(strCurPath)
    If strFileData = "" Then Exit Function
    
    strFormatLocalPath = FormatPath(strLocalPath, False)
    aryFile = Split(strFileData, ";")
    
    blnResult = True
    '循环下载ftp目录中的文件
    For i = LBound(aryFile) To UBound(aryFile)
        strFileData = aryFile(i)
        If Trim(strFileData) <> "" Then
            If Not ExecuteFtpCMD("GET " & strCurPath & strFileData & " " & strFormatLocalPath & GetFileName(strFileData, True)) Then blnResult = False
        End If
    Next i
    
    BatchDownLoad = blnResult
End Function



Public Function GetFtpList(Optional ByVal strFtpPath As String = "") As String
'获取ftp文件列表
'DIR命令
'格式：DIR /
'说明：从file1指定的目录检索文件,返回的数据通过GetChunk方法获取
    Dim strCurPath As String
    
    GetFtpList = ""
    strCurPath = FormatPath(strFtpPath, True)
    
    '如果执行失败，则直接退出
    If Not ExecuteFtpCMD("DIR " & IIf(strCurPath = "", mstrFtpVirtualPath, strCurPath)) Then Exit Function
    
    If mstrData = "" Then Exit Function
    
    GetFtpList = ";" & Replace(mstrData, vbCrLf, ";")
End Function


Public Function GetFtpDirList(Optional ByVal strFtpPath As String = "") As String
'获取ftp目录列表
    Dim strList As String
    Dim aryList() As String
    Dim i As Long
    
    GetFtpDirList = ""
    
    strList = GetFtpList(strFtpPath)
    If strList = "" Then Exit Function
    
    aryList = Split(strList, ";")
    
    For i = LBound(aryList) To UBound(aryList)
        If aryList(i) <> "" Then
            If Mid(aryList(i), Len(aryList(i)), 1) = "/" Then GetFtpDirList = GetFtpDirList & ";" & aryList(i)
        End If
    Next i
    
    If GetFtpList <> "" Then GetFtpDirList = GetFtpDirList & ";"
    
End Function

Public Function GetFtpFileList(Optional ByVal strFtpPath As String = "") As String
'获取ftp文件列表
    Dim strList As String
    Dim aryList() As String
    Dim i As Long
    
    GetFtpFileList = ""
    
    strList = GetFtpList(strFtpPath)
    If strList = "" Then Exit Function
    
    aryList = Split(strList, ";")
    
    For i = LBound(aryList) To UBound(aryList)
        If aryList(i) <> "" Then
            If Mid(aryList(i), Len(aryList(i)), 1) <> "/" Then GetFtpFileList = GetFtpFileList & ";" & aryList(i)
        End If
    Next i
    
    If GetFtpFileList <> "" Then GetFtpFileList = GetFtpFileList & ";"
End Function


Public Function SerachFile(ByVal strFtpFileName As String, Optional ByVal strSerachPath As String = "") As Boolean
'检索文件
'strSerachPath:从指定的strSerachPath 路径中检索文件，如果没有指定，则使用当前ftp路径进行检索

'如果存在，则返回true
    Dim strList As String
    
    SerachFile = False
    
    strList = GetFtpList(strSerachPath)
    If strList = "" Then Exit Function
    
    If InStr(strList, ";" & strFtpFileName & ";") > 0 Then SerachFile = True
End Function



Public Function RenameFile(ByVal strOldFile As String, ByVal strNewName As String) As Boolean
'创建目录
'RENAME命令
'格式：RENAME file1 file2
'说明：文件重命名，命令是否能够成功执行，取决于用户在远程主机上的权限
    Dim strFormatOldFile As String
    Dim strFormatNewFile As String
    
    RenameFile = False
    If Trim(strOldFile) = "" Or Trim(strNewName) = "" Then Exit Function
    
    strFormatOldFile = GetFilePath(strOldFile, True)
    strFormatNewFile = FormatPath(strFormatOldFile, True) & strNewName
    strFormatOldFile = FormatPath(strFormatOldFile, True) & GetFileName(strOldFile, True)

    RenameFile = ExecuteFtpCMD("RENAME " & strFormatOldFile & " " & strFormatNewFile)
End Function


Public Function MoveFile(ByVal strOldFile As String, ByVal strNewFile As String) As Boolean
'移动文件
'移动文件时，使用RENAME命令
    Dim strFormatOldFile As String
    Dim strFormatNewFile As String
    
    MoveFile = False
    If Trim(strOldFile) = "" Or Trim(strNewFile) = "" Then Exit Function
    
    strFormatOldFile = GetFilePath(strOldFile, True)
    strFormatNewFile = GetFilePath(strNewFile, True)
    
    strFormatOldFile = FormatPath(strFormatOldFile, True) & GetFileName(strOldFile, True)
    strFormatNewFile = FormatPath(strFormatNewFile, True) & GetFileName(strNewFile, True)
    
    MoveFile = ExecuteFtpCMD("RENAME " & strFormatOldFile & " " & strFormatNewFile)
End Function



Public Function CreateDir(ByVal strNewDir As String, Optional ByVal strFtpDir As String = "") As Boolean
'创建目录
'MKDIR命令
'格式：MKDIR file1
'说明：文件重命名，命令是否能够成功执行，取决于用户在远程主机上的权限
    Dim strCurPath As String
    Dim aryPath() As String
    Dim strCreatedDir As String
    Dim i As Long
    Dim blnResult As Boolean
    
    CreateDir = False
    If strNewDir = "" Then Exit Function
    
    strCurPath = FormatPath(strNewDir, True)
    aryPath = Split(strCurPath, "/")
    strCreatedDir = FormatPath(strFtpDir, True)
    
    For i = LBound(aryPath) To UBound(aryPath)
        If aryPath(i) <> "" Then
            If Not SerachFile(aryPath(i) & "/", strCreatedDir) Then
                blnResult = ExecuteFtpCMD("MKDIR " & strCreatedDir & aryPath(i) & "/")
                If Not blnResult Then Exit For
            End If
            
            strCreatedDir = strCreatedDir & aryPath(i) & "/"
        End If
    Next i

    CreateDir = blnResult
End Function



Public Function RemoteDir(ByVal strDelDir As String, Optional ByVal strFptDir As String = "") As String
'移除目录
'RMDIR命令
'格式：RMDIR file1
'说明：删除目录，命令是否能够成功执行，取决于用户在远程主机上的权限
    Dim strCurPath As String
    
    RemoteDir = False
    If strDelDir = "" Then Exit Function
    
    strCurPath = FormatPath(strDelDir, True)
    
    RemoteDir = ExecuteFtpCMD("RMDIR " & IIf(strFptDir = "", "", FormatPath(strFptDir, True)) & strCurPath)
End Function


Public Function GetFileSize(ByVal strFtpFile As String) As Double
'获取指定ftp文件大小
'如果成功返回文件的字节数,否则返回0 1k=1024字节

'SIZE命令
'格式：SIZE file1
'说明：返回由file1指定的文件大小
    Dim strCurFile As String
    Dim blnResult As Boolean
    
    GetFileSize = 0
    
    strCurFile = GetFilePath(strFtpFile, True)
    strCurFile = FormatPath(strCurFile, True) & GetFileName(strFtpFile, True)

    blnResult = ExecuteFtpCMD("SIZE " & strCurFile)
    
    If blnResult Then
        GetFileSize = mstrData
    End If
End Function


Public Function GetCodeInf(ByVal lngState As Long) As String
    GetCodeInf = ""

    Select Case lngState
        Case icNone
            GetCodeInf = "无状态可报告。"
        Case icResolvingHost
            GetCodeInf = "该控件正在查询所指定的主机的IP地址。"
        Case icHostResolved
            GetCodeInf = "该控件已成功地找到所指定的主机的IP地址。"
        Case icConnecting
            GetCodeInf = "该控件正在与主机连接。"
        Case icConnected
            GetCodeInf = "该控件已与主机连接成功。"
        Case icRequesting
            GetCodeInf = "该控件正在向主机发送请求。"
        Case icRequestSent
            GetCodeInf = "该控件发送请求已成功。"
        Case icReceivingResponse
            GetCodeInf = "该控件正在接收主机的响应。"
        Case icResponseReceived
            GetCodeInf = "该控件已成功地接收到主机的响应。"
        Case icDisconnecting
            GetCodeInf = "该控件正在解除与主机的连接。"
        Case icDisconnected
            GetCodeInf = "该控件已成功地与主机解除了连接。"
        Case icError
            GetCodeInf = "与主机通讯时出现了错误。"
        Case icResponseCompleted
            GetCodeInf = "该请求已经完成，并且所有数据均已接收到。"
    End Select
End Function




Private Sub Class_Initialize()
    mblnExecuteState = False
    mlngLastState = -1
    
    Set mobjInet = Nothing
End Sub

Private Sub mobjInet_StateChanged(ByVal State As Integer)
On Error GoTo errHandle
    mlngLastState = State
    
    Select Case State
        Case icResponseCompleted
            '获取执行后的数据
            mstrData = mobjInet.GetChunk(0, icString)
        Case icError
            '设置执行成功状态
            mblnExecuteState = False
    End Select
    

'   这里不能执行如下事件，否则在调用GetChunk方法时，容易产生上一请求仍在执行问题而造成异常
'    RaiseEvent OnStateChanged(State)
Exit Sub
errHandle:
    Debug.Print err.Description
End Sub
