Attribute VB_Name = "mdlMain"
Option Explicit

Public Const gstrRegPath As String = "公共模块\zlXWInterface\"   '注册表存储路径
Public Const gstrSysName As String = "影像信息系统接口"

Public gcnOracle As ADODB.Connection            '公共数据库连接
Public gzlComLib As Object                      '公共数据库处理模块zlComLib

Public glngSys As Long                          '系统号
Public glngModule As Long                       '模块号
Public gstrDBUser As String                     '当前数据库用户
Public gblnBefore3510 As Boolean                '区分10.35.10前后版本。True=10.35.10之前版本,不使用zlRegister，初始化comlib时需要SetDbUser和RegCheck

Public gstrLogPath As String
Public gstrBackupPath As String
Public gblnUseInterface As Boolean

Public mfrmShowHisForms As frmShowHisForms      '程序的主窗体，负责处理消息

Public mclsArchive As clsArchive                '电子病案查阅类
Public mobjLisInsideComm As Object              'LIS接口部件

Public Const HIS_CAPTION = "中联显示HIS窗口"

Public Sub Main()
'------------------------------------------------
'功能：主程序，负责启动电子病案查看程序
'参数：
'返回：无
'-----------------------------------------------
    Dim strRegPath As String
    Dim strMsgs As String
    Dim blnLis As Boolean
    Dim strTag As String
    Dim strOld As String
On Error GoTo ErrorHand

    
    strMsgs = Command
    
    strOld = strMsgs
    C_LOG = 1
    writeTestLog "入参串：" & strOld
    C_LOG = 0
    
    If Trim(strMsgs) = "" Then Exit Sub
    '根据命今行确实是否开了日志，关键字段::LOG=1::
    strTag = "::LOG=1::"
    If InStr(strMsgs, strTag) > 0 Then
        C_LOG = 1
        strMsgs = Replace(strMsgs, strTag, "")
    Else
        C_LOG = 0
    End If
    
    
    
    '网页打开时去特殊字符
    If InStr(strMsgs, "://") > 0 Then
        strMsgs = Split(strMsgs, "://")(1)
    End If
    If InStr(strMsgs, "/") > 0 Then
        strMsgs = Split(strMsgs, "/")(0)
    End If
    If Trim(strMsgs) = "" Then Exit Sub
    
    '如果本程序已经启动过一次，则不再启动，直接刷新界面数据后退出
    If App.PrevInstance Then
        If SendMsg(strMsgs) Then
            Exit Sub
        End If
    Else

    End If
    
    '接收到QUIT，直接退出
    If UCase(Trim(strMsgs)) = "QUIT" Then Exit Sub
    
    '根据传入的参数判断处理哪个部件，消息格式“部件编号:数据库用户名:病人ID:就诊ID:执行部门ID:医嘱ID”
    '就诊ID ： 门诊为挂号ID 病人挂号记录.ID，住院为主页ID
    
    '初始化comlib和数据库连接
    If UBound(Split(strMsgs, MSG_SPLIT)) = 5 Then
        gstrZLHIS主机字符串 = Split(strMsgs, MSG_SPLIT)(0)
        gstr用户名 = Split(strMsgs, MSG_SPLIT)(1)
        gstr密码 = Split(strMsgs, MSG_SPLIT)(2)
        gbln是否转换密码 = Val(Split(strMsgs, MSG_SPLIT)(3)) = 1
    ElseIf UBound(Split(strMsgs, MSG_SPLIT)) = 6 Then
        '调用LIS报告
        blnLis = Val(Split(strMsgs, MSG_SPLIT)(0)) = 25
        gstrZLHIS主机字符串 = Split(strMsgs, MSG_SPLIT)(1)
        gstr用户名 = Split(strMsgs, MSG_SPLIT)(2)
        gstr密码 = Split(strMsgs, MSG_SPLIT)(3)
        gbln是否转换密码 = Val(Split(strMsgs, MSG_SPLIT)(4)) = 1
    Else
        Exit Sub
    End If
    Call InitInterface(Split(strMsgs, MSG_SPLIT)(1))
    
    '初始化系统参数
    If Not blnLis Then Call InitSysParameter
    
    '创建消息处理窗体，加载消息hook，然后隐藏窗体
    If mfrmShowHisForms Is Nothing Then Set mfrmShowHisForms = New frmShowHisForms
    Call mfrmShowHisForms.ShowMe(True)
    mfrmShowHisForms.Hide
    
    '处理消息
    Call ProcessMessage(strMsgs)
    
    Exit Sub
ErrorHand:
    If errHandle("exe Main", "显示病案查阅窗口出现错误") = 1 Then Resume
End Sub

'将消息发送给消息循环主窗体
Private Function SendMsg(ByVal strmsg As String) As Boolean
    Dim lngWinHandle As Long        '需要接收消息的“zlSoftShowHisForms.exe”程序的窗口句柄
    Dim wParam As Long
    Dim lResult As Long
    Dim strTemp As String
    Dim buf(1 To 1024) As Byte
    
    '消息定义：wParam = 223，dss中dwData = 33 处理消息，dwData = 32 退出
    wParam = 223
   
    Call CopyMemory(buf(1), ByVal strmsg, LenB(StrConv(strmsg, vbFromUnicode)))
    
    'dss.dwData这个消息不用，只是双方定义的一个标记而已
    If UCase(Trim(strmsg)) = "QUIT" Then
        dss.dwData = 32 '标记为关闭所有窗口
    Else
        dss.dwData = 33 '标记为刷新窗口或者打开新窗口
    End If
    
    dss.cbData = LenB(StrConv(strmsg, vbFromUnicode)) + 1
    
    '使用buf发送，可以控制消息在1024之内
    dss.lpData = VarPtr(buf(1))
    
    '查找消息循环主窗体
    lngWinHandle = FindWindow(vbNullString, HIS_CAPTION)
    

    If lngWinHandle <> 0 Then
        lResult = SendMessage(lngWinHandle, WM_COPYDATA, wParam, dss)
        SendMsg = True
    End If
End Function


