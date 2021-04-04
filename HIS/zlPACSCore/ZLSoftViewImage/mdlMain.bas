Attribute VB_Name = "mdlMain"
Option Explicit


Public Sub Main()
'------------------------------------------------
'功能：主程序，负责启动观片程序
'       供其他开发程序调用的接口方法，比如C#，显示观片站，保存成本地缓存，并优先从本地缓存中读取图像
'       根据lngOrderID和strImages从数据库中查找图像文件串，打开一个检查的图像，调用一次本接口
'       由于使用了数据库连接串，仅支持 10.35.10之后的HIS版本
'参数：
'返回：无
'-----------------------------------------------
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
    
    Dim strMsgs As String
    
    On Error GoTo err
    
    '先创建日志文件目录
    gstrLogPath = GetLogDir()
    
    '把exe命令的参数，传给strMsgs 参数，等待处理。不接受无参数的exe调用，直接退出
    strMsgs = Command
    If Trim(strMsgs) = "" Then Exit Sub
    
    '如果本程序已经启动过一次，则不再启动，直接刷新界面数据后退出
    If App.PrevInstance Then
        Call WriteCommLog("zlSoftViewImage.Sub Main", "★★将消息发送给已存在的zlSoftViewImage，当前程序退出。版本为：" & App.Major & "." & App.Minor & "." & App.Revision, "参数为：strMsgs = " & strMsgs, ltDebug)
        Call SendMsg(strMsgs)
        Exit Sub
    Else
        Call WriteCommLog("zlSoftViewImage.Sub Main", "★★第一次启动zlSoftViewImage.版本为：" & App.Major & "." & App.Minor & "." & App.Revision, "参数为：strMsgs = " & strMsgs, ltDebug)
    End If
    
    '接收到QUIT，直接退出
    If UCase(Trim(strMsgs)) = "QUIT" Then Exit Sub
    
    '初始化，不需要初始化，直接就在zl9PacsCore中初始化了
    
    '创建消息处理窗体，加载消息hook，然后隐藏窗体
    If gfrmViewImage Is Nothing Then Set gfrmViewImage = New frmViewImage
    Call gfrmViewImage.ShowMe(True)
    gfrmViewImage.Hide
    
    '处理消息
    Call ProcessMessage(strMsgs)
    
    Exit Sub
err:
    If errHandle("exe Main", "显示病案查阅窗口出现错误") = 1 Then Resume
End Sub

Private Sub SendMsg(ByVal strMsg As String)
'------------------------------------------------
'功能：将消息发送给消息循环主窗体
'参数：strMsg -- 调用exe时传入的参数串
'返回：无
'-----------------------------------------------
    Dim lngWinHandle As Long        '需要接收消息的“zlSoftViewImage.exe”程序的窗口句柄
    Dim wParam As Long
    Dim lResult As Long
    Dim strTemp As String
    Dim buf(1 To 1024) As Byte
    
    '消息定义：wParam = 223，dss中dwData = 33 处理消息，dwData = 32 退出
    wParam = 223
   
    Call CopyMemory(buf(1), ByVal strMsg, LenB(StrConv(strMsg, vbFromUnicode)))
    
    'dss.dwData这个消息不用，只是双方定义的一个标记而已
    If UCase(Trim(strMsg)) = "QUIT" Then
        dss.dwData = 32 '标记为关闭所有窗口
    Else
        dss.dwData = 33 '标记为刷新窗口或者打开新窗口
    End If
    
    dss.cbData = LenB(StrConv(strMsg, vbFromUnicode)) + 1
    
    '使用buf发送，可以控制消息在1024之内
    dss.lpData = VarPtr(buf(1))
    
    '查找消息循环主窗体
    lngWinHandle = FindWindow(vbNullString, HIS_CAPTION)
    
    Call WriteCommLog("zlSoftViewImage.SendMsg", "将消息发送给消息循环主窗体", "消息为：" & strMsg & "，窗口句柄为：" & lngWinHandle, ltDebug)
    
    If lngWinHandle <> 0 Then
        lResult = SendMessage(lngWinHandle, WM_COPYDATA, wParam, dss)
    End If
End Sub



    
    
