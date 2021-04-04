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
Public mclsOrder As clsOrder                  '医嘱处理类
Public mclsFee  As clsFee                     '收费处理类

Public mobjLisInsideComm As Object              'LIS接口部件

Public gclsReport As Object
Public gobjRegister As Object

Public Const HIS_CAPTION = "中联显示HIS窗口NEW"

Public Sub Main()
'------------------------------------------------
'功能：主程序，负责启动电子病案查看程序
'参数：
'       Command：(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=127.0.0.1)(PORT=1521))(CONNECT_DATA=(SERVICE_NAME=testbase))):ZLHIS:AQA:1:2:1:1
'       Command参数含义：Oracle连接字符串:用户名:密码:是否密码转换(0或1):调用功能号(-1 功能初始化,0-病案查阅,1-浏览检查报告,2-医嘱处理;3-执行端付费配置;4-执行端付费;5-打印单据;99-打开自定义报表,999-功能初始化):...
'              功能号不同，后续参数的格式与含义也不同
'              功能=0,1,2时:功能号后参数：病人ID,主页ID
'              功能=3,999时，功能后无参数
'              功能=4时,功能后为:病人ID:医嘱信息:NOs                               其中医嘱信息或NOs，任传一个即可,医嘱信息：执行科室|医嘱IDs(多个用逗号分隔);NOs: 多个用逗号分隔
'              功能=5时,功能后为：打印类别(0=含打印及预览,1=直接到预览,2=直接打印,3-输出到Excel,4-输出到PDF,99-打印设置):(格式：报表编号,单据号(par)报表编号,单据号)功能后为:  报表编号,单据号(par)报表编号,单据号(par)报表编号,单据号
'              功能=99时,功能后为：系统号:报表编号:打印类别(0=含打印及预览,1=直接到预览,2=直接打印,3-输出到Excel,4-输出到PDF,99-打印设置):报表参数(可为空 示例格式：病人id=1<par>PDF=C:\1.PDF<par>ExcelFile=C:\1.xls)
    
'返回：无
'-----------------------------------------------
    Dim strMsgs As String
    Dim blnLis As Boolean
    

    Dim varTmp As Variant
    
On Error GoTo ErrorHand

    
    strMsgs = Command
    
        '接收到QUIT，直接退出
    If UCase(Trim(strMsgs)) = "QUIT" Then
            '如果本程序已经启动过一次，则不再启动，直接刷新界面数据后退出
        If App.PrevInstance Then
            If SendMsg(strMsgs) Then
                Exit Sub
            End If
        Else
            Exit Sub
        End If
    End If

    
    varTmp = Split(strMsgs, ":")
    
    glng功能号 = Val(varTmp(4)) '0-病案查阅,1-LIS调用,2-医嘱处理;3-执行端付费配置;4-执行端付费;5-单据打印;99-打开自定义报表
    glngFunID = IIf(glng功能号 = 2, 3001, 0)

    If Trim(strMsgs) = "" Then Exit Sub
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
    End If
    
    
    '根据传入的参数判断处理哪个部件，消息格式“部件编号:数据库用户名:病人ID:就诊ID:执行部门ID:医嘱ID”
    '就诊ID ： 门诊为挂号ID 病人挂号记录.ID，住院为主页ID
    
    '初始化comlib和数据库连接
    If UBound(Split(strMsgs, MSG_SPLIT)) < 4 Then Exit Sub
    gstrZLHIS主机字符串 = Split(strMsgs, MSG_SPLIT)(0)
    gstr用户名 = Split(strMsgs, MSG_SPLIT)(1)
    gstr密码 = Split(strMsgs, MSG_SPLIT)(2)
    gbln是否转换密码 = Val(Split(strMsgs, MSG_SPLIT)(3)) = 1
    blnLis = glng功能号 = 2
    
    
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
    Dim lngWinHandle As Long        '需要接收消息的“zlSoftCISInterface.exe”程序的窗口句柄
    Dim wParam As Long
    Dim lResult As Long
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


