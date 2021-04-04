Attribute VB_Name = "mdlGetImage"
Option Explicit

Public plngHookHandle As Long       '记录消息hook的handle
Public dss As COPYDATASTRUCT        '传递字符串消息的内存结构
Public pblnQueueBusy As Boolean     '记录当前操作是否在处理队列，一次只能有一个程序处理队列
Public Const TIME_OUT = 60          '超时时间设置，60秒
Public pintQueueIndex As Integer    '记录队列中第一个消息的索引
Public pConnectedSharedDir() As String  '记录已经连接过的共享目录
Public pftpConnect As clsFtp        '定义一个常连接的FTP类
Public pblnMsgProcessing As Boolean '正在处理消息

'日志相关的参数，从注册表“”中读取日志参数，如果日志路径为空，则使用exe相同目录下的“GetImgLog”作为日志路径
Public pblnLogEnable As Boolean     '是否启用日志
Public pstrLogPath As String        '记录日志的路径,
Public plngLogLevel As String       '记录日志的级别，分成1,2两级。1级只记录消息级别的日志；2级记录每一次下载的日志


'保存消息内容的结构
Public Type TGetImgMsg
    strSubDir As String          '图像所在的子目录
    strDestMainDir As String            '复制图像的目的目录，本机目录
    strIP As String                 '图像服务器的IP地址
    strFTPDir As String             'FTP目录
    strFTPUser As String            'FTP用户名
    strFTPPswd As String            'FTP密码
    strSDDir As String              '共享目录名称
    strSDUser As String             '共享目录用户名
    strSDPswd As String             '共享目录密码
    blnEnable As Boolean            '本消息可用
End Type
Public curMsg As TGetImgMsg


'下载图像的消息队列
Public pstrMsgQueue() As String          '消息内容队列

'消息Hook变量
Public plngPreWndProc As Long       '原来的消息处理程序


Public Function MsgInQueue(strMsg As String) As Boolean
'------------------------------------------------
'功能：消息入队，等候被处理
'参数： strMsg －－需要入队的消息内容，消息内容由存储地址，存储用户名，存储密码，存储目录，本机目录，存储类型组成
'返回：True－－入队成功，False－－入队失败
'-----------------------------------------------
    Dim Timer As Long
    Dim intCount As Integer
    
    On Error GoTo err
    
    MsgInQueue = False
    
    If pblnQueueBusy = True Then
        '如果队列忙，正在进行队列处理，则等待，如果等待超时，则将消息放到备用队列
        
        
    Else
        '如果队列空闲，则进行入队处理，并且标记队列忙
        pblnQueueBusy = True
        
        '处理消息入队
        intCount = UBound(pstrMsgQueue) + 1
        ReDim Preserve pstrMsgQueue(intCount) As String
        pstrMsgQueue(intCount) = strMsg
        
        '队列处理完成，标记队列为闲
        pblnQueueBusy = False
    End If
    MsgInQueue = True
    Exit Function
err:
    '出错就退出，暂时不做处理
End Function

Public Function MsgOutQueue() As String
'------------------------------------------------
'功能：消息出队，调用过程负责处理出队的消息
'参数：
'返回：返回出队的消息内容
'-----------------------------------------------
    Dim iCount As Integer
    Dim strMsg As String
    
    On Error GoTo err
    MsgOutQueue = ""        '初始化为空消息
    
    If pblnQueueBusy = True Then
        '如果队列忙，正在进行队列处理，则等待，如果等待超时，退出出队操作
        
    Else
        '如果队列空闲，则进行出队处理，并且标记队列忙
        pblnQueueBusy = True
        
        '消息出队处理
        iCount = UBound(pstrMsgQueue)
        If iCount = 0 Then
            pblnQueueBusy = False
            Exit Function     '队列为空，不用出队
        End If
        
        '从队列中提取消息
        strMsg = pstrMsgQueue(pintQueueIndex)
        
        '处理队列指针
        pintQueueIndex = pintQueueIndex + 1
        If pintQueueIndex > iCount Then
            '如果当前取出来的是队列中的最后一个消息，则清空消息队列
            ReDim Preserve pstrMsgQueue(0) As String
            pintQueueIndex = 1
        End If
        
        MsgOutQueue = strMsg
        
        '出队处理完成，标记队列闲
        pblnQueueBusy = False
    End If
    
    Exit Function
err:
    '出错就退出，暂时不做处理
End Function


Public Function funConnectAndSaveSheardDir(strSharedDir As String, strUser As String, strPswd As String) As Boolean
'------------------------------------------------
'功能：连接共享目录，使用用户名和秘密登录服务器
'参数： strSharedDir -- 需要连接的共享目录名称

'返回：True-- 成功，False -- 失败
'-----------------------------------------------
    Dim i As Integer
    
    funConnectAndSaveSheardDir = False

    On Error GoTo err
    
    '判断共享目录是否已经连接，如果没有连接，则进行连接
    For i = 1 To UBound(pConnectedSharedDir)
        If pConnectedSharedDir(i) = strSharedDir Then
            funConnectAndSaveSheardDir = True
            Exit Function
        End If
    Next i
    
    '连接共享目录
    If strSharedDir <> "" Then
        If funConnectSharedDir(strSharedDir, strUser, strPswd) = True Then
            '连接成功，记录成功的连接串
            ReDim Preserve pConnectedSharedDir(UBound(pConnectedSharedDir) + 1) As String
            pConnectedSharedDir(UBound(pConnectedSharedDir)) = strSharedDir
            funConnectAndSaveSheardDir = True
        End If
    End If
    
    Exit Function
err:
    '暂不处理
End Function


Public Function funConnectSharedDir(strShareRemoteDir As String, strUserName As String, _
    strPassWord As String) As Boolean
'------------------------------------------------
'功能：创建网络资源
'参数： strShareRemoteDir -- 共享目录
'       strUserName -- 共享目录用户名
'       strPassWord -- 共享目录密码
'返回：True--连接成功； False -- 连接失败
'------------------------------------------------
    
    Dim NetR As NETRESOURCE
    Dim lngResult As Long
    
    On Error GoTo err
    
    funConnectSharedDir = False
    
    NetR.dwType = RESOURCETYPE_ANY
    NetR.lpLocalName = vbNullString
    NetR.lpRemoteName = strShareRemoteDir
    NetR.lpProvider = vbNullString
    lngResult = WNetAddConnection2(NetR, strPassWord, strUserName, 0)
    
   ' If lngResult = 0 Then
        funConnectSharedDir = True
   ' End If
    Exit Function
err:
    '暂不处理
End Function


Public Function funDownLoadSharedDir(strSourceDir As String, strDestDir As String) As Boolean
'------------------------------------------------
'功能：通过共享目录的方式下载一个目录的图像，strSourceDir目录所在的计算机登录，通过另外的过程实现
'参数： strSourceDir -- 需要复制文件的源目录，即远程服务器中的共享目录
'       strDestDir  --  文件复制的目的地，即计算机的本机目录
'返回：True-- 成功，False -- 失败
'-----------------------------------------------
    Dim fs As New Scripting.FileSystemObject
    Dim fsFolder As Folder
    Dim fsFiles As Files
    Dim fsFile As File
    
    On Error Resume Next
    
    funDownLoadSharedDir = False
    
    '如果源目录不存在，则退出
    If Dir(strSourceDir, vbDirectory) = "" Then
        '记录日志
        Call WriteCommLog("funDownLoadSharedDir", "通过共享目录方式下载图像", "源目录不存在" & strSourceDir, 1)
        Exit Function
    End If
    
    '如果目标目录不存在，则创建目录
    If Dir(strDestDir, vbDirectory) = "" Then
        Call MkLocalDir(strDestDir)
    End If
    
    '遍历目录中的所有文件，一个个下载，可以确保如果本机目录中已经有了某个文件，后续的文件还可以正常下载
    Set fsFolder = fs.GetFolder(strSourceDir)
    Set fsFiles = fsFolder.Files
                
    '禁止OverWrite，防止用FTP中的文件覆盖本机目录中已经有了的文件。
    For Each fsFile In fsFiles
        '记录日志
        Call WriteCommLog("funDownLoadSharedDir", "下载图像", "下载图像： " & strDestDir & "\" & fsFile.Name, 2)
        
        Call fs.CopyFile(strSourceDir & "\" & fsFile.Name, strDestDir & "\" & fsFile.Name, False)
    Next fsFile
    
    funDownLoadSharedDir = True

End Function


Public Sub MkLocalDir(ByVal strDir As String)
'------------------------------------------------
'功能：创建本地目录
'参数： strDir－－本地目录
'返回：无
'------------------------------------------------
    Dim objFile As New Scripting.FileSystemObject
    Dim aNestDirs() As String, i As Integer
    Dim strPath As String
    On Error Resume Next
    
    '读取全部需要创建的目录信息
    ReDim Preserve aNestDirs(0)
    aNestDirs(0) = strDir
    
    strPath = objFile.GetParentFolderName(strDir)
    Do While Len(strPath) > 0
        ReDim Preserve aNestDirs(UBound(aNestDirs) + 1)
        aNestDirs(UBound(aNestDirs)) = strPath
        strPath = objFile.GetParentFolderName(strPath)
    Loop
    '创建全部目录
    For i = UBound(aNestDirs) To 0 Step -1
        MkDir aNestDirs(i)
    Next
End Sub

Public Function funGetAMessage() As Boolean
'------------------------------------------------
'功能：从消息队列提取一个消息，填充消息结构
'参数：
'返回：True -- 成功 ； False -- 失败
'-----------------------------------------------
    Dim strMsg As String
    Dim strOldMsg As String
    Dim strMsgArr() As String
    
    funGetAMessage = False
    
    On Error GoTo err
    
    '记录上一个消息，如果出现相同的连续消息，则不处理新消息
    strOldMsg = strMsg
    
    '从消息队列提取一个消息
    strMsg = MsgOutQueue
    
    '如果消息不为空，则处理这个消息
    If strMsg <> "" And strOldMsg <> strMsg Then
        '解析消息
        strMsgArr = Split(strMsg, "||")
        '只处理包含了9个段的消息，有的段可能为空
        If UBound(strMsgArr) = 8 Then
            curMsg.strSubDir = strMsgArr(0)
            curMsg.strDestMainDir = strMsgArr(1)
            curMsg.strIP = strMsgArr(2)
            curMsg.strFTPDir = strMsgArr(3)
            curMsg.strFTPUser = strMsgArr(4)
            curMsg.strFTPPswd = strMsgArr(5)
            curMsg.strSDDir = strMsgArr(6)
            curMsg.strSDUser = strMsgArr(7)
            curMsg.strSDPswd = strMsgArr(8)
            curMsg.blnEnable = True
            
            '整理路径，取出路径中第一个和最后一个“\”或者“/”符号
            curMsg.strSubDir = funRemoveSlash(curMsg.strSubDir)
            curMsg.strDestMainDir = funRemoveSlash(curMsg.strDestMainDir)
            curMsg.strFTPDir = funRemoveSlash(curMsg.strFTPDir)
            
            '消息正常出队并写入消息结构，设置消息接收完成标记
            funGetAMessage = True
        Else
            curMsg.blnEnable = False
        End If
    End If
    
    Exit Function
err:
    '暂不处理
End Function

Public Function funRemoveSlash(strDir As String) As String
'------------------------------------------------
'功能：去除路径中的第一个和最后一个斜线，或者反斜线
'参数： strDir  -- 需要处理的路径
'返回：处理之后的路径
'-----------------------------------------------
    Dim strTemp As String
    
    strTemp = strDir
    funRemoveSlash = strTemp
    
    On Error GoTo err
    
    '去除路径后部的斜线
    If Right(strTemp, 1) = "/" Or Right(strTemp, 1) = "\" Then
        strTemp = Left(strTemp, Len(strTemp) - 1)
    End If
    
    '去除路径前的斜线
    If Left(strTemp, 1) = "/" Or Left(strTemp, 1) = "\" Then
        strTemp = Mid(strTemp, 2)
    End If
    
    funRemoveSlash = strTemp
    
    Exit Function
err:
    '暂不处理,出错则直接返回未处理的字符串
End Function


Public Function funDownLoadFTP(thisMsg As TGetImgMsg) As Boolean
'------------------------------------------------
'功能：根据下载图像的消息，从FTP中下载指定目录中的全部图像
'参数： thisMsg  -- 需要下载图像的消息
'返回：True -- 成功； False -- 失败
'-----------------------------------------------
    Dim lngResult As String
    
    funDownLoadFTP = False
    On Error GoTo err
    
    '先检查消息是否可用
    If thisMsg.blnEnable = False Or thisMsg.strIP = "" Then
        
        '记录日志
        Call WriteCommLog("funDownLoadFTP", "消息不可用", "FTP方式下载图像，消息不可用或者IP地址为空，无法下载。 IP地址是：" & thisMsg.strIP, 1)
        
        Exit Function
    End If
    
    '连接FTP
    '判断当前消息中的FTP连接跟现有的FTP连接是否相同
    If thisMsg.strIP = pftpConnect.IPAddress And thisMsg.strFTPUser = pftpConnect.User _
        And thisMsg.strFTPPswd = pftpConnect.PassWord Then
        '不需要重新连接FTP
    Else
        '重新连接FTP
        Call pftpConnect.FuncFtpDisConnect
        lngResult = pftpConnect.FuncFtpConnect(thisMsg.strIP, thisMsg.strFTPUser, thisMsg.strFTPPswd)
        
        '如果连接失败，退出程序
        If lngResult = 0 Then
            '记录日志
            Call WriteCommLog("funDownLoadFTP", "FTP连接失败", "FTP连接失败： " & thisMsg.strIP, 1)
            
            Exit Function
        End If
    End If
    
    '创建本地路径
    Call MkLocalDir(thisMsg.strDestMainDir & "\" & thisMsg.strSubDir)
    
    '记录日志
    Call WriteCommLog("funDownLoadFTP", "下载图像", "通过FTP下载全部图像。 ", 1)
    
    '下载图像
    Call pftpConnect.funcDownLoadAllFiles("\" & thisMsg.strFTPDir & "\" & thisMsg.strSubDir, thisMsg.strDestMainDir & "\" & thisMsg.strSubDir, False)
    
    funDownLoadFTP = True
    Exit Function
err:
    '暂不处理
End Function

Public Function funMsgProcess() As Boolean
'------------------------------------------------
'功能：自动处理消息
'参数： 无
'返回：True -- 成功； False -- 失败
'-----------------------------------------------
    
    On Error GoTo err
    
    funMsgProcess = False
    
    '如果已经进行消息处理了，则退出
    If pblnMsgProcessing = True Then Exit Function
    
    '设置消息处理标记，防止这个过程被多次调用
    pblnMsgProcessing = True
    
    '消息出队并处理消息
    While funGetAMessage = True
        '记录日志
        Call WriteCommLog("funMsgProcess", "接收到并处理消息", "准备下载此目录下的图像：" & curMsg.strSubDir, 1)

        '处理消息
        Call funDownLoadImages(curMsg)
    Wend
    
    '消息处理完成，退出程序
    pblnMsgProcessing = False
    
    funMsgProcess = True
    Exit Function
err:
    '暂不处理
End Function

Public Function funDownLoadImages(thisMsg As TGetImgMsg) As Boolean
'------------------------------------------------
'功能：根据下载图像的消息，从共享目录或者FTP下载图像
'参数： thisMsg  -- 需要下载图像的消息
'返回：True -- 成功； False -- 失败
'-----------------------------------------------
Dim blnResult As Boolean
    
    On Error GoTo err
    
    '如果当前消息可用，则处理消息
    If thisMsg.blnEnable = True Then
        '有共享目录，则使用共享目录下载图像，没有则使用FTP下载图像
        If thisMsg.strSDDir <> "" Then
            '连接共享目录
            If funConnectAndSaveSheardDir("\\" & thisMsg.strIP & "\" & thisMsg.strSDDir, thisMsg.strSDUser, thisMsg.strSDPswd) = True Then
                '记录日志
                Call WriteCommLog("funDownLoadImages", "共享目录方式下载图像", "通过共享目录方式下载图像，从目录： " & "\\" & thisMsg.strIP & "\" & thisMsg.strSDDir & "\" & thisMsg.strSubDir & " ，下载到此目录： " & thisMsg.strDestMainDir & "\" & thisMsg.strSubDir, 1)
                
                '下载图像
                blnResult = funDownLoadSharedDir("\\" & thisMsg.strIP & "\" & thisMsg.strSDDir & "\" & thisMsg.strSubDir, thisMsg.strDestMainDir & "\" & thisMsg.strSubDir)
            End If
        Else
            '记录日志
            Call WriteCommLog("funDownLoadImages", "FTP方式下载图像", "通过FTP方式下载图像，从目录： \" & thisMsg.strFTPDir & "\" & thisMsg.strSubDir & " ，下载到此目录： " & thisMsg.strDestMainDir & "\" & thisMsg.strSubDir, 1)
                
            '使用FTP下载
            blnResult = funDownLoadFTP(thisMsg)
        End If
    End If
    
    funDownLoadImages = blnResult
    Exit Function
err:
    '暂不处理

End Function

Public Sub Main()
'------------------------------------------------
'功能：主程序，负责启动图像下载程序
'参数：
'返回：无
'-----------------------------------------------
    Dim strRegPath As String
    
    '如果本程序已经启动过一次，则不再启动
    If App.PrevInstance Then
        Exit Sub
    End If
    
    
    
    On Error Resume Next
    
    '从注册表读取日志参数
    strRegPath = "公共模块\zlPacsGetImage"
    pblnLogEnable = (Val(GetSetting("ZLSOFT", strRegPath, "记录日志", 0)) = 1)
    pstrLogPath = GetSetting("ZLSOFT", strRegPath, "日志路径", "")
    plngLogLevel = Val(GetSetting("ZLSOFT", strRegPath, "日志级别", 1))
    '如果启动了日志，检查日志路径是否存在
    If pblnLogEnable = True Then
    
        '如果没有设置日志路径，则使用默认路径
        If pstrLogPath = "" Then
            pstrLogPath = App.Path & "\GetImgLog"
        End If
        
        '如果日志路径不存在，则创建
        If Dir(pstrLogPath, vbDirectory) = "" Then
            '默认路径不存在，创建这个目录
            If Dir(pstrLogPath, vbDirectory) = "" Then
                Call MkLocalDir(pstrLogPath)
            End If
        End If
    End If
    
    '第一次启动程序，加载窗体，并隐藏
    frmMain.Show
    frmMain.WindowState = vbMinimized
    frmMain.Hide
    
End Sub


Private Sub WriteCommLog(logSubName As String, logTitle As String, logDesc As String, lngLogLevel As Long)
'------------------------------------------------
'功能：记录通讯日志
'参数： logSubName  --  产生日志的函数名
'       logTitle   -- 日志名称
'       logDesc   --  日志内容
'       lngLogLevel -- 日志级别，通过日志级别确定当前日志是否需要记录
'返回：无
'------------------------------------------------
    Dim strLog As String
    Dim strFileName As String
    Dim intHour As Integer
    
    On Error GoTo err
    
    If pblnLogEnable = True Then        '启动了记录日志，才记录当前的日志
        '判断日志级别，确定本次日志是否需要记录
        If plngLogLevel >= lngLogLevel Then
            '通过当前时间，创建日志文件名，每两个小时产生一个日志文件
            intHour = Hour(Time)
            intHour = intHour / 2
            intHour = intHour * 2
            strFileName = pstrLogPath & "\" & Date & "-" & intHour & ".log"
            
            '产生日志内容
            strLog = Now() & " 日志级别： " & lngLogLevel & " 标题： " & logTitle & vbCrLf & "      函数： " & logSubName & vbCrLf & "     日志内容：" & logDesc & vbCrLf
            
            '打开日志文件，记录日志
            Open strFileName For Append As #1
            Print #1, strLog
            Close #1
        End If
    End If
    Exit Sub
err:
    Close #1
End Sub
