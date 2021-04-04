Attribute VB_Name = "mdlPublic"
Option Explicit

'调用参数： {+}1405{+}ZLHIS[+]ZLHIS[+]HIS[+]0{+}false{+}false{+}0{+}0{+}false

Public gstrLogPath As String        '日志文件
Public gstrImages As String         '消息参数 strImages
Public glngOrderID As Long          '消息参数 lngOrderID
Public gstrDBConnection As String   '消息参数 strDBConnection
Public gblnMoved As Boolean         '消息参数 blnMoved
Public gbAdd As Boolean             '消息参数 bAdd
Public gintImageInterval As Integer '消息参数 intImageInterval
Public glngSys As Long              '消息参数 lngSys
Public gblnReconnectDB As Boolean   '消息参数 blnReconnectDB
Public gstrDBServer As String       '消息参数 strDBServer
Public gstrDBUser As String         '消息参数 strDBUser
Public gstrDBPassword As String     '消息参数 strDBPassword
Public gblnTransPassword As Boolean '消息参数 blnTransPassword
Public gfrmViewImage As frmViewImage    '消息循环的主窗体
Public gobjPacsCore As Object       '观片对象
Public glngPreWndProc As Long       '原来的消息处理程序
Public glngLog As Long              '是否记录日志；0---参数未赋值；1---记录日志；2---不记录日志

Public Const HIS_CAPTION = "中联影像观片窗口"
Public Const MSG_SPLIT = "{+}"

Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long

Public Enum LogType
    ltError = 0
    ltDebug = 1
End Enum

Public Function errHandle(errSubName As String, errTitle As String, Optional errDesc As String = "") As Long
'------------------------------------------------
'功能：错误处理
'参数： logSubName  --  产生错误的函数名
'       logTitle   -- 错误名称
'       logDesc   --  错误描述
'返回：1-程序继续Resume；0-程序退出
'------------------------------------------------
    
    errHandle = 0
    
    '记录错误日志
    Call WriteCommLog("zlSoftViewImage,错误--" & errSubName, errTitle & "，错误代码= " & err.Number, errDesc & "，错误描述=" & err.Description, ltError)
    
    '提示错误
    MsgBox errTitle & errDesc, vbOKOnly, "观片接口zlSoftViewImage出现错误"
    
    '清除错误
    err.Clear
    
End Function

Public Sub WriteCommLog(logSubName As String, logTitle As String, logDesc As String, ByVal ltLogType As LogType)
'------------------------------------------------
'功能：记录通讯日志
'参数： logSubName  --  产生日志的函数名
'       logTitle   -- 日志名称
'       logDesc   --  日志内容
'       ltLogType --  日志类型
'返回：无
'------------------------------------------------
    Dim strLog As String
    Dim strFileName As String
    
    On Error GoTo err
    
    If glngLog = 0 Then
        glngLog = Val(GetSetting("ZLSOFT", "公共模块\zl9PacsCore\zlSoftViewImage\", "Log", 2))
    End If
    
    'Log=1，才记录日志
    If glngLog <> 1 And ltLogType <> ltError Then Exit Sub
    
    strFileName = gstrLogPath & "\Interface" & Format(Date, "YYYY-MM-DD") & ".log"
    
    strLog = Now() & " 标题： " & logTitle & vbCrLf & "   函数： " & logSubName & vbCrLf & "   日志内容：" & logDesc & vbCrLf
    
    '错误日志增加标记，方便查看分析
    If ltLogType = ltError Then
        strLog = "▲▲错误▲▲：" & strLog
    End If
    
    Open strFileName For Append As #1
    Print #1, strLog
    Close #1
    
    Exit Sub
err:
    Close #1
End Sub

Public Function GetLogDir() As String
'------------------------------------------------
'功能：获取日志目录，如果目录不存在，则创建目录
'参数：无
'返回：日志所在目录
'------------------------------------------------
    Dim strLogPath As String
    Dim strBackupPath As String
    
    On Error GoTo err
    
    strLogPath = App.Path & "\zlViewImageLog"
    
    Call MkLocalDir(strLogPath + "\")
    
    GetLogDir = strLogPath
   
    Exit Function
err:
    GetLogDir = App.Path & "\XWInterfaceLog"
    Call MkLocalDir(GetLogDir + "\")
End Function

Public Function ProcessMessage(strMsg As String) As Long
'------------------------------------------------
'功能：处理接收到的消息
'参数：strMsg -- 调用exe时传入的参数串
'返回：无
'------------------------------------------------
    
    Dim lngPartType As Long
    Dim strDBUser As String
    Dim lngPatientID As Long
    Dim lngClinicID As Long
    Dim lngDeptID As Long
    Dim lngOrderID As Long
    
    On Error GoTo err
    ProcessMessage = 1
    
    '传入的参数定义，参数的连接符是三个字符“{+}”
    '参数格式：strImages{+}lngOrderID{+}strDBConnection{+}blnMoved{+}bAdd{+}intImageInterval{+}lngSys{+}blnReconnectDB
    '参数解释： strImages --- 图象号,规则是“序列UID1|1-3;5-27;33-100+序列UID2|全部”,全部表示打开全部图象
    '           lngOrderID --- 医嘱ID
    '           strDBConnection --- 数据库连接串，包含“服务名[+]用户名[+]密码[+]密码是否转换”，连接符是三个字符“[+]”
    '                          当“密码”是用户登录密码时，“密码是否转换”=1；当“密码”是数据库登录密码时，“密码是否转换”=0
    '           blnMoved --- 数据是否被转储
    '           bAdd --- 可选参数，默认值False，新图像是增加进观片站，还是替换原观片站所有图像，True为增加，Fasle为替换
    '           intImageInterval --- 可选参数，默认值0，打开图像的间隔，只对打开全部序列,且序列中图像数量>100时有效
    '           lngSys --- 可选参数，默认,100，系统序号
    '           blnReconnectDB --- 可选参数，默认值False，是否重新连接数据库。第一次打开观片时自动连接数据库，之后再打开观片，
    '                           由blnReconnectDB参数决定是否重新连接数据库。
    '                           =True，使用strDBConnection参数重新连接数据库；=False，不再重新连接数据库，使用观片部件现在的数据库连接
    '
    
    '先处理固定参数
    If UBound(Split(strMsg, MSG_SPLIT)) >= 3 Then
        gstrImages = Split(strMsg, MSG_SPLIT)(0)
        glngOrderID = Val(Split(strMsg, MSG_SPLIT)(1))
        gstrDBConnection = Split(strMsg, MSG_SPLIT)(2)
        gblnMoved = (UCase(Split(strMsg, MSG_SPLIT)(3)) = "TRUE")
    Else
        Call WriteCommLog("错误--zlSoftShowHisForms.ProcessMessage", "解析参数", "解析参数出错，参数数量不够4个，参数为：" & strMsg, ltError)
        Exit Function
    End If
    
    '再处理可选参数
    If UBound(Split(strMsg, MSG_SPLIT)) >= 4 Then
        gbAdd = (UCase(Split(strMsg, MSG_SPLIT)(4)) = "TRUE")
    Else
        gbAdd = False
    End If
    
    If UBound(Split(strMsg, MSG_SPLIT)) >= 5 Then
        gintImageInterval = Val(Split(strMsg, MSG_SPLIT)(5))
    Else
        gintImageInterval = 0
    End If
    
    If UBound(Split(strMsg, MSG_SPLIT)) >= 6 Then
        glngSys = Val(Split(strMsg, MSG_SPLIT)(6))
    Else
        glngSys = 100
    End If
    
    If UBound(Split(strMsg, MSG_SPLIT)) = 7 Then
        gblnReconnectDB = (UCase(Split(strMsg, MSG_SPLIT)(7)) = "TRUE")
    Else
        gblnReconnectDB = False
    End If
    
    If CreatePacsCore = False Then
        Exit Function
    End If
    
    Call WriteCommLog("zlSoftShowHisForms.ProcessMessage", "调用观片", "观片的参数是：gstrImages=" & gstrImages & ",glngOrderID=" & glngOrderID _
        & ",gstrDBConnection=" & gstrDBConnection & ",gblnMoved=" & gblnMoved & ",gbAdd=" & gbAdd & ",gintImageInterval=" & gintImageInterval _
        & ",glngSys=" & glngSys & ",gblnReconnectDB=" & gblnReconnectDB, ltDebug)
    
    Call gobjPacsCore.CallOpenViewerSimple(gstrImages, glngOrderID, gstrDBConnection, gblnMoved, gbAdd, gintImageInterval, glngSys, gblnReconnectDB)
    
    ProcessMessage = 0
    Exit Function
    
err:
    Call WriteCommLog("错误--zlSoftShowHisForms.ProcessMessage", "处理接收到的消息，出现错误，收到的消息是：" & strMsg & "，错误代码= " & err.Number, "，错误描述=" & err.Description, ltError)
End Function

'******************************************************************************************************************
'功能：创建PACS观片对象
'参数：无
'返回：创建成功,返回true,否则返回False
'说明：
'******************************************************************************************************************
Private Function CreatePacsCore() As Boolean

    err = 0: On Error Resume Next
    If Not gobjPacsCore Is Nothing Then CreatePacsCore = True: Exit Function
    
    Set gobjPacsCore = CreateObject("zl9PacsCore.clsViewer")
    
    If err <> 0 Then
        MsgBox "未找到 zl9PacsCore 部件，可能是程序版本不支持，请检查该站点是否部署了此部件!", vbInformation + vbOKOnly, "提示信息"
        Exit Function
    End If
    
    CreatePacsCore = True
    
End Function

Public Function CloseAllForms() As Boolean

    On Error GoTo err
    
    '关闭消息循环主窗口
    If Not gfrmViewImage Is Nothing Then
        Unload gfrmViewImage
        Set gfrmViewImage = Nothing
    End If
    
    CloseAllForms = True
    
    Exit Function
err:
    Call WriteCommLog("错误--zlSoftViewImage.CloseAllForms", "退出程序，关闭所有窗口，出现错误，错误代码= " & err.Number, "，错误描述=" & err.Description, ltError)
    Resume Next
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
