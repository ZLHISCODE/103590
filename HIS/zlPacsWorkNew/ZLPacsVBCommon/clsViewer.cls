VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsViewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mstrImageShareDir As String                  '老版的影像共享存储目录
Private mblnXWLog As Boolean                        '是否记录通讯日志
Private mlngStudySchemeNo As Long                    '检查方案号
Private mlngSeriesSchemeNo As Long                   '序列方案号
Private mblnUseXinWangView As Boolean
Private mConnectedShardDir() As String   '已经连接过的共享目录的设备号数组
Private mobjPacsCore As Object      '观片站对象

Private mobjGetImage As Object
Private mlngSys As Long

Public Function OpenViewer(ByVal lngSys As Long, cnOracle As ADODB.Connection, ByVal strDbUser As String, ByVal lngViewerType As Long, lngAdviceID As Long, _
        blnAddImage As Boolean, Optional ByVal strSerials As String = "", _
        Optional ByVal blnMoved As Boolean = False, Optional ByVal blnLocalizerBackward As Boolean = False, _
        Optional ByVal intImageInterval As Integer = 0, Optional ByVal strImageString As String = "") As Long
'------------------------------------------------
'功能：根据传入的医嘱ID和发送号，打开objPacsCore指向的观片站
'参数：
'       lngViewerType -- 展现图像的Viewer类型；1-放射科专用Viewer；2-临床浏览用Viewer
'       objPacsCore －－观片站对象
'       lngAdviceID －－医嘱ID
'       blnAddImage--True 在原有图像基础上增加当前图像；False删除原有图像，打开当前图像
'       objParent -- 父窗体
'       strSerials－－可选，序列UID名称串，用逗号分隔，如果不输入，则选择全部序列
'       blnMoved－－可选，是否被转储
'       blnLocalizerBackward--可选，定位像后置,跟strImageString互斥
'       intImageInterval ---可选，打开图像的间隔，比如5，表示每5个图打开一个图,跟strImageString互斥
'       strImageString --- 可选，每个序列中需要打开的图象号组合，跟intImageInterval和blnLocalizerBackward互斥，
'                           以strImageString为主
'                           规则是“序列UID1|1-3;5-27;33-100+序列UID2|全部”,全部表示打开全部图象
'返回：图像文件名串数组
'------------------------------------------------
    Dim strSql As String
    Dim rsTmp As New ADODB.Recordset
    Dim strFTPHost As String, strFtpPath As String, strFTPUser As String, strFTPPswd As String
    Dim strSDPath As String, strSDUser As String, strSDPwd As String
    Dim strDeviceNO As String
    Dim i As Integer
    Dim blnConnectDS As Boolean         '是否连接当前的共享目录
    Dim oneMessage As TGetImgMsg        '预取图像的消息结构
    Dim intImageLocation As Integer
    Dim strXWViewerFilter As String
    
    On Error GoTo DBError
    
    Call InitComLib(cnOracle, strDbUser)
    
    mblnXWLog = (Val(zlDatabase.GetPara("XW记录接口日志", mlngSys, G_LNG_XWPACSVIEW_MODULE, "0")) = 1) '是否记录接口日志
    mstrImageShareDir = zlDatabase.GetPara("XW历史图像共享目录", mlngSys, G_LNG_XWPACSVIEW_MODULE, "DCMSHARE")
    mlngStudySchemeNo = Val(zlDatabase.GetPara("XW检查方案号", mlngSys, G_LNG_XWPACSVIEW_MODULE, "1"))
    mlngSeriesSchemeNo = Val(zlDatabase.GetPara("XW序列方案号", mlngSys, G_LNG_XWPACSVIEW_MODULE, "2"))
    mblnUseXinWangView = IsUseXwViewer
    
    BUGEX "mblnUseXinWangView = " & mblnUseXinWangView, True
    
    mlngSys = lngSys
    
    '查询图像在新网PACS或者是中联PACS
    strSql = "Select 图像位置,影像类别 from 影像检查记录 where 医嘱ID =[1]"
    
    If blnMoved Then
        strSql = Replace(strSql, "影像检查记录", "H影像检查记录")
    End If
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "查询图像所在的位置", lngAdviceID)
    
    If rsTmp.RecordCount <> 0 Then
        intImageLocation = Nvl(rsTmp!图像位置, 0)
    End If
    
    BUGEX "lngAdviceID = " & lngAdviceID, True
    BUGEX "intImageLocation = " & intImageLocation, True
    BUGEX "mblnXWLog = " & mblnXWLog, True
    
    '图像在新网数据库，调用新网DLL显示图像
    If intImageLocation = 1 Then
        strXWViewerFilter = lngAdviceID & IIf(strSerials <> "", "[;]" & strSerials, "")
        
        If mblnXWLog = True Then
            Call WriteCommLog("OpenViewer", "调用XWShowImage接口", "查询过滤参数为：" & strXWViewerFilter)
        End If
        
        Call XWShowImage(lngViewerType, strXWViewerFilter, IIf(strSerials <> "", mlngSeriesSchemeNo, mlngStudySchemeNo))
        
        OpenViewer = -1
        
        Exit Function
    End If
    
    '判断是否启用了新版pacs观片
    If mblnUseXinWangView = True Then
        Call OpenViewerWithInXWPacs(lngAdviceID, Nvl(rsTmp!影像类别), blnMoved)
        
        OpenViewer = -1
        Exit Function
    End If
    
    
    '图像在中联数据库，调用中联zl9PacsCore显示图像
    strFTPHost = ""
           
    '查找需要打开的所有图象信息
    strSql = "Select D.IP地址 As Host1,d.设备号 as 设备号1," & _
        "Decode(C.接收日期,Null,'',to_Char(C.接收日期,'YYYYMMDD')||'/')" & _
        "||C.检查UID||'/' As Path,E.IP地址 As Host2,e.设备号 as 设备号2, " & _
        "D.共享目录 AS 共享目录1, E.共享目录 AS 共享目录2,D.共享目录用户名 as 共享目录用户名1, " & _
        "E.共享目录用户名 AS 共享目录用户名2,D.共享目录密码 AS 共享目录密码1,E.共享目录密码 AS 共享目录密码2, " & _
        "D.FTP目录 as FTP目录1, E.FTP目录 as FTP目录2,D.FTP用户名 as FTP用户名1, E.FTP用户名 AS FTP用户名2,  " & _
        "D.FTP密码 as FTP密码1, E.FTP密码 AS FTP密码2 " & _
        "From 影像检查记录 C,影像设备目录 D,影像设备目录 E " & _
        "Where C.位置一=D.设备号(+) And C.位置二=E.设备号(+) And C.医嘱ID=[1] "
    
    '如果有转储标志，则读取转储的历史表
    If blnMoved Then
        strSql = Replace(strSql, "影像检查图象", "H影像检查图象")
        strSql = Replace(strSql, "影像检查序列", "H影像检查序列")
        strSql = Replace(strSql, "影像检查记录", "H影像检查记录")
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "获取共享目录信息", lngAdviceID)
    
    If rsTmp.RecordCount > 0 Then
        '创建本地的缓存目录，需要在调用观片站之前先创建这个目录，观片站中只是下载，不创建本地缓存目录
        MkLocalDir App.Path & "\TmpImage\" & rsTmp("Path")
        
        '读取FTP参数，包括用户名，密码，IP地址等
        If rsTmp("设备号1") <> "" Then
            strDeviceNO = rsTmp("设备号1")
            strFTPHost = rsTmp("Host1")
            strFtpPath = Nvl(rsTmp("FTP目录1"))
            strFTPUser = Nvl(rsTmp("FTP用户名1"))
            strFTPPswd = Nvl(rsTmp("FTP密码1"))
            strSDPath = Nvl(rsTmp("共享目录1"))
            strSDUser = Nvl(rsTmp("共享目录用户名1"))
            strSDPwd = Nvl(rsTmp("共享目录密码1"))
        ElseIf Nvl(rsTmp("设备号2")) <> "" Then
            strDeviceNO = rsTmp("设备号2")
            strFTPHost = rsTmp("Host2")
            strFtpPath = Nvl(rsTmp("FTP目录2"))
            strFTPUser = Nvl(rsTmp("FTP用户名2"))
            strFTPPswd = Nvl(rsTmp("FTP密码2"))
            strSDPath = Nvl(rsTmp("共享目录2"))
            strSDUser = Nvl(rsTmp("共享目录用户名2"))
            strSDPwd = Nvl(rsTmp("共享目录密码2"))
        End If
        
        '判断共享目录是否已经连接，如果没有连接，则进行连接
        blnConnectDS = True
        For i = 1 To UBound(mConnectedShardDir)
            If mConnectedShardDir(i) = strDeviceNO Then
                blnConnectDS = False
                Exit For
            End If
        Next i
        If blnConnectDS = True And strSDPath <> "" Then
            If funcConnectShardDir(Nothing, "\\" & strFTPHost & "\" & strSDPath, strSDUser, strSDPwd) = 0 Then
                ReDim Preserve mConnectedShardDir(UBound(mConnectedShardDir) + 1) As String
                mConnectedShardDir(UBound(mConnectedShardDir)) = strDeviceNO
            End If
        End If
        
        If mobjPacsCore Is Nothing Then Set mobjPacsCore = DynamicCreate("zl9PacsCore.clsViewer", "观片站")
        
        '打开观片站
        If mobjPacsCore Is Nothing Then
            Exit Function
        Else
            OpenViewer = mobjPacsCore.OpenViewer(strImageString, lngAdviceID, cnOracle, blnMoved, blnAddImage, intImageInterval, lngSys)
        End If
        
        '先打开观片站，再预取
        oneMessage.strSubDir = rsTmp("Path")
        oneMessage.strDestMainDir = App.Path & "\TmpImage\"
        oneMessage.strIP = strFTPHost
        oneMessage.strFtpDir = strFtpPath
        oneMessage.strFTPUser = strFTPUser
        oneMessage.strFTPPswd = strFTPPswd
        oneMessage.strSDDir = strSDPath
        oneMessage.strSDUser = strSDUser
        oneMessage.strSDPswd = strSDPwd
        
        If CBool(GetSetting("ZLSOFT", "公共模块", "UseActiveLoad", "True")) Then
            oneMessage.strFtpDir = "/" & oneMessage.strFtpDir & "/"
            Call funPreDownLoadImagesEx(Me, oneMessage)
        Else
            Call funPreDownLoadImages(oneMessage)
        End If
        
    Else    '没有查找到图象记录，则关闭原来已经打开的观片窗口
        If Not mobjPacsCore Is Nothing Then mobjPacsCore.Closefrom
    End If
    
    Exit Function
DBError:
    If gobjComlib.ErrCenter() = 1 Then Resume
    BUGEX "OpenViewer err= " & err.Description, True
End Function

Private Function funPreDownLoadImages(thisMsg As TGetImgMsg) As Boolean
'------------------------------------------------
'功能：后台下载图像
'参数： thisMsg  -- 要下载的图像信息
'返回：True -- 成功； False -- 失败
'------------------------------------------------
    Dim lngWinHandle As Long        '需要接收消息的“中联图像下载”程序的窗口句柄
    Dim strMsg As String
    Dim wParam As Long
    Dim lResult As Long
    Dim strTemp As String
    Dim buf(1 To 1024) As Byte
    Dim dss As COPYDATASTRUCT
    
    On Error Resume Next
    
    '组织消息
    strMsg = thisMsg.strSubDir & "||" & thisMsg.strDestMainDir & "||" & thisMsg.strIP & "||" _
            & thisMsg.strFtpDir & "||" & thisMsg.strFTPUser & "||" & thisMsg.strFTPPswd & "||" _
            & thisMsg.strSDDir & "||" & thisMsg.strSDUser & "||" & thisMsg.strSDPswd
    
    '发送COPYDATA消息
    
    On Error GoTo err
    
    '使用BUF，或者使用lstrcpy函数都可以正常发送字符消息
   '消息定义：wParam = 123，dss中dwData = 3
    wParam = 123
   
    Call CopyMemory(buf(1), ByVal strMsg, LenB(StrConv(strMsg, vbFromUnicode)))
    dss.dwData = 3               '这个消息不用，3只是双方定义的一个标记而已
    dss.cbData = LenB(StrConv(strMsg, vbFromUnicode)) + 1
    
    dss.lpData = VarPtr(buf(1))                    '使用buf发送，可以控制消息在1024之内
'    dss.lpData = lstrcpy(strMsg, strMsg)            '这个方法发送的消息，也是正确的。
    
    
    '启动图像下载窗口
    Shell App.Path & "\zlGetImage.exe"
        
    '加载窗体的时候，查找图像下载程序
    lngWinHandle = FindWindow(vbNullString, "中联图像下载")
    
    lResult = SendMessage(lngWinHandle, WM_COPYDATA, wParam, dss)
    
    funPreDownLoadImages = True
    Exit Function
err:
    '暂不处理
End Function

Private Function funPreDownLoadImagesEx(ByVal objParent As Object, thisMsg As TGetImgMsg) As Boolean
'------------------------------------------------
'功能：后台下载图像
'参数： thisMsg  -- 要下载的图像信息
'返回：True -- 成功； False -- 失败
'------------------------------------------------
    Dim blnUseActivexLoad As Boolean
    Dim objImgInfo As Object
    
    On Error Resume Next
    
    If mobjGetImage Is Nothing Then
        Set mobjGetImage = DynamicCreate("zlGetImageEx.clsImageTransfer", "zlGetImageEx.exe")
        
        If mobjGetImage Is Nothing Then Exit Function
        
        Call mobjGetImage.RegEventObj(objParent)
        Call mobjGetImage.zlInitModule(False, 0)
    End If
    
    Set objImgInfo = mobjGetImage.ImgInfo
    With objImgInfo
        .SubDir = thisMsg.strSubDir
        .DestMainDir = thisMsg.strDestMainDir
        .IP = thisMsg.strIP
        .FtpDir = thisMsg.strFtpDir
        .FtpUser = thisMsg.strFTPUser
        .FTPPswd = thisMsg.strFTPPswd
        .SDDir = thisMsg.strSDDir
        .SDUser = thisMsg.strSDUser
        .SDPswd = thisMsg.strSDPswd
    End With
    
    Call mobjGetImage.MsgInQueue(objImgInfo)
    Call mobjGetImage.zlLoadImage
    
    Exit Function
err:
    '暂不处理
End Function

Private Function OpenViewerWithInXWPacs(ByVal lngAdviceID As Long, ByVal strModalityType As String, ByVal blnMoved As Boolean)
'在新版pacs中打开观片
    Dim rsTemp As ADODB.Recordset

    Dim strFtpUrl As String
    Dim strImages As String
    
    Set rsTemp = GetStudyImageData(lngAdviceID, blnMoved)
    
    strImages = ""

    While Not rsTemp.EOF
        If Nvl(rsTemp!设备号1) <> "" Then
            strFtpUrl = "\\" & Nvl(rsTemp!Host1) & "\" & mstrImageShareDir & Nvl(rsTemp!Root1) & Nvl(rsTemp!Url)
        Else
            strFtpUrl = "\\" & Nvl(rsTemp!Host2) & "\" & mstrImageShareDir & Nvl(rsTemp!Root2) & Nvl(rsTemp!Url)
        End If
        
        If strImages <> "" Then strImages = strImages & "[;]"
        
        strFtpUrl = Replace(strFtpUrl, "//", "/")
        strImages = strImages & Replace(strFtpUrl, "/", "\")
        
        rsTemp.MoveNext
    Wend
    
    '打开远程目录文件进行对比观片
    Call OEMViewOpen(0, strImages, 0, strModalityType)
End Function

Private Function GetStudyImageData(ByVal lngAdviceID As Long, ByVal blnMoved As Boolean) As ADODB.Recordset
'获取检查图像数据

    Dim strSql As String
        
    strSql = "Select rownum as 顺序号, A.图像号,D.FTP用户名 As User1,D.FTP密码 As Pwd1," & _
        "D.IP地址 As Host1,'/'||D.Ftp目录||'/' As Root1," & _
        "Decode(C.接收日期,Null,'',to_Char(C.接收日期,'YYYYMMDD')||'/')" & _
        "||C.检查UID||'/'||A.图像UID As URL,d.设备号 as 设备号1, " & _
        "E.FTP用户名 As User2,E.FTP密码 As Pwd2," & _
        "E.IP地址 As Host2,'/'||E.Ftp目录||'/' As Root2," & _
        "e.设备号 as 设备号2, A.图像UID,C.检查UID,B.序列UID,A.动态图,A.编码名称,A.采集时间, A.录制长度 " & _
        "From 影像检查图象 A,影像检查序列 B,影像检查记录 C,影像设备目录 D,影像设备目录 E " & _
        "Where A.序列UID=B.序列UID And B.检查UID=C.检查UID And C.位置一=D.设备号(+) And C.位置二=E.设备号(+) and c.医嘱ID=[1] "
        

    If blnMoved Then
        strSql = Replace(strSql, "影像检查图象", "H影像检查图象")
        strSql = Replace(strSql, "影像检查序列", "H影像检查序列")
        strSql = Replace(strSql, "影像检查记录", "H影像检查记录")
    End If
    
    Set GetStudyImageData = zlDatabase.OpenSQLRecord(strSql, "查询图像信息", lngAdviceID)
End Function

'判断是否启用观片
Private Function IsUseXwViewer() As Boolean
On Error GoTo errHandle
    Dim lngPhkResult As Long
    Dim lngKey As Long
    Dim blnResult As Boolean
    Dim strValue As String
    Dim lngLen As Long
    
    blnResult = IIf(RegOpenKeyEx(HKEY_CURRENT_USER, "Software\Silver\Silver Pacs\General", 0, 1, lngPhkResult) = 0&, True, False)
    BUGEX "0 IsUseXwViewer = " & blnResult, True
    '如果安装了xw viewer，则注册表中的安装信息
    If blnResult = False Then Exit Function
    
    strValue = Space(255)
    lngLen = Len(strValue)
    
    Call RegQueryValueEx(lngPhkResult, "CompanyName", 0, 0, strValue, lngLen)
    
    Call RegCloseKey(lngKey)
    
    If InStr(strValue, "中联医学影像系统") <= 0 Then blnResult = False
    BUGEX "1 IsUseXwViewer = " & blnResult, True
    IsUseXwViewer = blnResult
Exit Function
errHandle:
    IsUseXwViewer = False
    BUGEX "2 err IsUseXwViewer = " & blnResult & "   err= " & err.Description, True
End Function

Private Function XWADViewerOpen(ByVal strFilter As String, ByVal lngPlanId As Long) As Long
'--------------------------------------------
'功能： 打开指定图像。条件由参数指定，且必须与配置文件中的配置相符。
'       图像打开时与ADViewer 当前模式有关，如果是单记录模式，则软件会自动关闭原来的图像；如果是对比
'       模式，则会添加到ADViewer中。
'参数：
'       lngOrderID -- 医嘱ID
'返回：
'--------------------------------------------
    Dim strRev As String
    Dim lngFunction As Long
    Dim strXwPrivs As String
    
    'XWViewerOpen 参数说明：
    'lPlanID：  方案ID。该ID 必须与INI 文件中一致，在简单网络的情况下，通常该值为1，建议把该ID 做为一配置项，调用时读取该项并传入。
    'cpFilter： 该代表要打开图像的条件值。例如检查号、申请号等，可以传入多个值，
    '           不同值之间用分隔符[;]隔开，该参数的意义及顺序在INI 文件中配置，并且与lPlanID对应。
    'lFunc：    功能权限。每一位代表一项功能，如果具有多项权限，按位“或”即可，具体功能意义:
    '           0x00000002： 重建图像保存，例如：减影后图像、拼接图像等
    '           0 x00000200: 胶片打印
    '           0 x00040000: 图像导出?另存为其他格式
    '           0 x00080000: GSPS 保存
    'cpReserved：   保留，设为NULL
    
    On Error GoTo err
    
    '记录接口日志
    If mblnXWLog = True Then
        Call WriteCommLog("XWADViewerOpen", "XW接口", "打开ADViewer，并显示图像，医嘱ID= " & strFilter)
    End If
    
    '根据RIS中的权限，组织权限串
    lngFunction = 0
    strXwPrivs = gobjComlib.GetPrivFunc(mlngSys, G_LNG_XWPACSVIEW_MODULE)
    
    If InStr(strXwPrivs, "PACS保存重建图像") <> 0 Then
        lngFunction = lngFunction Or &H2
    End If
    
    If InStr(strXwPrivs, "PACS胶片打印") <> 0 Then
        lngFunction = lngFunction Or &H200
    End If
    
    If InStr(strXwPrivs, "PACS图像导出") <> 0 Then
        lngFunction = lngFunction Or &H40000
    End If
    
    If InStr(strXwPrivs, "PACS GSPS保存") <> 0 Then
        lngFunction = lngFunction Or &H80000
    End If
        
    XWADViewerOpen = OEMViewOpen(lngPlanId, strFilter, lngFunction, "")
    
    If XWADViewerOpen <> 0 Then
        MsgBox "ADViewer打开错误，返回的信息是：" & XWADViewerOpen
    End If
    
    Exit Function
err:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function

'-------------------------------------------------------------------------------------------------------
'ADViewer查看图像应用函数
'-------------------------------------------------------------------------------------------------------
Private Function XWShowImage(ByVal lngViewerType As Long, ByVal strFilter As String, Optional ByVal lngPlanId As Long = 1) As Long
''--------------------------------------------
''功能： 打开新网的ADViewer或者WEB Viewer
''参数：    lngViewerType -- 打开Viewer的方式；1-放射科ADViewer；2-临床WEB Viewer
'           lngOrderID -- 医嘱ID
''返回：0-成功;1-出错
''--------------------------------------------
    On Error GoTo err
    
    '记录接口日志
    If mblnXWLog = True Then
        Call WriteCommLog("XWShowImage", "XW接口", "调用ADViewer或者WEB观片，观片方式是： " & IIf(lngViewerType = 1, "ADViewer", "WEB"))
    End If
    
    If lngViewerType = 1 Then
        Call XWADViewerOpen(strFilter, lngPlanId)
    ElseIf lngViewerType = 2 Then
        Call XWWebViewerOpen(strFilter)
    End If
    
    Exit Function
err:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function XWWebViewerOpen(ByVal lngOrderID As Long) As Long
''--------------------------------------------
''功能： 打开新网的WEB Viewer
'           lngOrderID -- 医嘱ID
''返回：0-成功;1-出错
''--------------------------------------------
    Dim strIP As String
    Dim strURL As String
    
    On Error GoTo err
    
    strIP = zlDatabase.GetPara("XWWEB服务器IP", mlngSys, G_LNG_XWPACSVIEW_MODULE, "")
    
    If strIP <> "" Then
        strURL = "C:\Program Files\Internet Explorer\iexplore.exe http://" & strIP & ":8080/imageweb/imageAction.action?ColID0=22&ColValue0=" & lngOrderID
        
        '记录接口日志
        If mblnXWLog = True Then
            Call WriteCommLog("XWWebViewerOpen", "XW接口", "通过WEB方式观片： " & strURL)
        End If
        
        Shell strURL, vbMaximizedFocus
        XWWebViewerOpen = 0
    Else
        '记录接口日志
        If mblnXWLog = True Then
            Call WriteCommLog("XWWebViewerOpen", "XW接口", "通过WEB方式观片：WEB服务器IP地址为空。")
        End If
        
        MsgBox "WEB服务器IP地址为空，请先设置好WEB服务器。", vbOKOnly, "提示信息"
        
        XWWebViewerOpen = 1
    End If
    
    Exit Function
err:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Sub OnState(ByVal blnLoadFinish As Boolean, ByVal blnUpLoad As Boolean, ByVal lngThreadId As Long)
'blnLoadFinish,图像下载情况，ture:图像下载完成；false:图像正在下载
'blnUpLoad，true:上传；false:下载
'lngThreadId，当前下载图像的进程标记
    
On Error GoTo err
    '图像上传成功后，释放其对象
    If blnLoadFinish Then Set mobjGetImage = Nothing
    
    Exit Sub
err:

End Sub

Private Sub Class_Initialize()
    ReDim mConnectedShardDir(0) As String   '初始化共享目录连接串
End Sub

Private Sub Class_Terminate()
    If Not mobjPacsCore Is Nothing Then
        mobjPacsCore.Closefrom
        Set mobjPacsCore = Nothing
    End If
End Sub
