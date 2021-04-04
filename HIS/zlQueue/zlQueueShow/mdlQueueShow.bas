Attribute VB_Name = "mdlQueueShow"
Option Explicit

'消息标记定义
Public Const G_STR_MSG_QUEUE_001 As String = "ZLHIS_QUEUE_001" '入队消息
Public Const G_STR_MSG_QUEUE_002 As String = "ZLHIS_QUEUE_002" '完成消息
Public Const G_STR_MSG_QUEUE_003 As String = "ZLHIS_QUEUE_003" '状态同步
Public Const G_STR_MSG_QUEUE_004 As String = "ZLHIS_QUEUE_004" '语音呼叫

Public Enum TBusinessType
'业务类型定义
    btClinical = 0  '临床排队业务
    btPacs = 1      'Pacs排队业务
    btPeis = 2       '体检排队业务
    'bt...          '如果有其他业务，则在后面进行扩展
End Enum

Public Enum TShowStyle
'排队叫号的显示样式
    ssSingleMan = 0     '单病人样式
    ssSingleQueue = 1   '单队列（已确认诊室或执行间），即按一个诊室或者执行间显示
    ssMultiQueue = 2    '多队列（多个已确认诊室或执行间或分组或按科室排队的队列）,
    ssOld = 3           '老版显示
End Enum

Public Type TRect
'显示位置区域
    lngLeft As Long         '左坐标
    lngTop As Long          '顶点坐标
    lngWidth As Long        '宽
    lngHeight As Long       '高
    lngMonitorIndex As Long '显示器索引
End Type

Public Type TLcdCommonParameter
'LCD通用参数结构
    ssShowStyle As TShowStyle           '窗口显示样式

    lngCurDeptID As Long                '当前科室ID
    strCurDiagnoseRoom As String        '当前诊室名称
    
    strQueryQueueNames As String        '队列名称，多个队列使用“,”逗号分隔
    blnShowAdvertise As Boolean         '是否显示广告
    
    strFilter As String                 '显示数据过滤条件
    lngCallingRows As Long
    lngQueueRows As Long
    
    blnConvertQueueName As Boolean      '转换成老版存储规则下的队列名称
    
    blnScrollDisplay As Boolean         '滚动显示
    blnFontAutoSizeToList As Boolean    '字体自动适应列表
    recPos  As TRect                    '窗口显示区域
End Type

'注册路径
Public Const G_STR_REGPATH = "公共模块\zl9QueueShow"

Public gcnOracle As New ADODB.Connection    '公共数据库连接

Public gobjStyleWindow() As Object
Private mobjIcon As clsTaskIcon

Public gobjComLib As Object            'zl9ComLib.clsComLib
Public gobjQueueShow As Object         'zl9LCDShow.clsLCDShow
Public gstrUserName As String
Public glngBusinessType As Long                     'LCD显示所属业务类型
Public gstrSysName As String
Public gstrSystems As String
Public gstrStation As String
Public glngSys As String
Public gstrCompareVersion As String     '当前版本

Public gobjFile As New FileSystemObject

Public Sub Main()
    Dim objLogin As Object  'zlLogin.clsLogin
    Dim strCommand As String, strUserName As String, strPassword As String, strServer As String
    Dim blnAutoLogin As Boolean
    
    Set objLogin = DynamicCreate("zlLogin.clsLogin", "zlLogin.dll")
    If objLogin Is Nothing Then Exit Sub
    
    Set gobjComLib = DynamicGet("zl9ComLib.clsComLib", "zl9ComLib.dll")
    If gobjComLib Is Nothing Then Exit Sub
    
    If App.PrevInstance Then
        MsgBox "队列显示服务已经启动，不能再次运行。", vbInformation, "警告"
        Exit Sub
    End If
    
    '为实现XP风格，在显示窗体前必须执行该函数
    Call InitCommonControls
    
    '打开登陆界面，如需将用户名保存在注册表中，则需要传入注册路径
    
    blnAutoLogin = Val(GetSetting("ZLSOFT", G_STR_REGPATH, "自动登录", 0)) = 1
    
    If blnAutoLogin Then
        '从注册表中获取登录信息并解密，以便自动登陆
        strUserName = getDecryptionPassW(GetSetting("ZLSOFT", G_STR_REGPATH, "用户名", ""))
        strPassword = getDecryptionPassW(GetSetting("ZLSOFT", G_STR_REGPATH, "密码", ""))
        strServer = getDecryptionPassW(GetSetting("ZLSOFT", G_STR_REGPATH, "服务器", ""))
        
        If strUserName = "" Or strPassword = "" Or strServer = "" Then
            Set gcnOracle = objLogin.Login
        Else
            strCommand = "USER=" & strUserName & " PASS=" & strPassword & " SERVER=" & strServer
            Set gcnOracle = objLogin.Login(0, strCommand)
        End If
    Else
        Set gcnOracle = objLogin.Login
    End If
    
    If gcnOracle Is Nothing Then Exit Sub
    If gcnOracle.ConnectionString = "" Then Exit Sub
    
    '保存登陆信息，加密保存
    SaveSetting "ZLSOFT", G_STR_REGPATH, "用户名", getEncryptionPassW(objLogin.InputUser)
    SaveSetting "ZLSOFT", G_STR_REGPATH, "密码", getEncryptionPassW(objLogin.InputPwd)
    SaveSetting "ZLSOFT", G_STR_REGPATH, "服务器", getEncryptionPassW(objLogin.ServerName)
    
    gstrSysName = "提示"
    gstrUserName = objLogin.DBUser

    '初始化zlcomlib对象
    gobjComLib.InitCommon gcnOracle
    
    gstrSystems = " (系统 =100 Or 系统 Is NULL)"
    glngSys = 100
    
    '调整到指定的显示页面
    Call ShowWindow(blnAutoLogin)
End Sub

Private Sub ShowWindow(ByVal blnAutoLogin As Boolean)
'打开队列页面
'根据参数显示对应的窗口,显示规则如下
'1.如果启用了自动登录，则直接显窗口样式页面，并加载数据
'2.如果没有启用自动登录，则显示配置窗口
'不论已那种方式显示后，都需要显示托盘图标

On Error GoTo ErrorHand
    
    gstrCompareVersion = getCompareVersion
    
    Call InitOldLCDShow
    
    glngBusinessType = Val(GetSetting("ZLSOFT", G_STR_REGPATH, "所属业务", 1))
    
    If blnAutoLogin Then
        '根据参数显示样式窗口
        Call OpenStyleWindow
    Else
        '显示配置窗口
        Call OpenMainCfg
    End If
    
    '打开托盘图标
    Call OpenTrayIcon
    
    Exit Sub
ErrorHand:
    MsgBox Err.Description, vbExclamation, gstrSysName
    Err.Clear
End Sub

Private Function getCompareVersion() As String
'获取当前系统版本
    Dim strSql As String
    Dim rsRecord As ADODB.Recordset
    
    getCompareVersion = ""
    
    strSql = "Select nvl(主版本,1) 主版本,nvl(次版本,0) 次版本,nvl(附版本,0) 附版本,名称 " & _
             "From ZlComponent Where Upper(Rtrim(部件))=upper('zl9PacsWork') And 系统=100"
             
    Set rsRecord = gobjComLib.zlDatabase.OpenSQLRecord(strSql, "获取版本信息")
    
    If rsRecord.RecordCount > 0 Then
        '组装版本号为三位主版本、三位次版本及三位附版本
        getCompareVersion = String(3 - Len(rsRecord!主版本), "0") & rsRecord!主版本 & "." & _
                            String(3 - Len(rsRecord!次版本), "0") & rsRecord!次版本 & "." & _
                            String(3 - Len(rsRecord!附版本), "0") & rsRecord!附版本
    End If
End Function

Private Sub OpenOldLcd(ByVal lngShowNum As Long)
'打开老版本的LcdShow进行显示
    Dim i As Integer
    Dim strQueueNames As String
    Dim str队列名称() As String     '队列名称需按老版本的格式传入：如PACS:/*64:CT1,64:CT2....*/
    Dim blnConvertQueueName As Boolean  '是否转换成老版本格式的队列名称

    blnConvertQueueName = Val(GetSetting("ZLSOFT", G_STR_REGPATH & "\" & lngShowNum, "转换队列名称", 0)) = 1

    '根据业务类型获取对应格式的队列名称
    strQueueNames = ConvertFormat(GetSetting("ZLSOFT", G_STR_REGPATH & "\" & lngShowNum, "显示队列"), blnConvertQueueName)

    If strQueueNames = "" Then Exit Sub

    str队列名称 = Split(strQueueNames, ",")

    Call InitOldLCDShow

    Call gobjQueueShow.zlShow(gcnOracle, str队列名称, "", "", "", 0, False)
End Sub

Private Function ConvertFormat(ByVal strQueueName As String, ByVal blnConvertQueueName As Boolean) As String
'按格式转换：把队列名称转换成数据库中存储的格式
    Dim i As Integer
    Dim str队列名称() As String, strQueueNames As String
    Dim strSql As String
    Dim rsRecord As ADODB.Recordset
    Dim lngPreDeptID As Long
    Dim lngCurDeptID As Long
    Dim blnQueueStyle As Boolean
    
    If strQueueName = "" Then Exit Function
    
    strQueueNames = ""
    lngPreDeptID = 0
    lngCurDeptID = 0
    
    str队列名称 = Split(strQueueName, ",")
    
    If blnConvertQueueName Then    '转换成老版本格式的队列名称
        For i = 0 To UBound(str队列名称)
            lngCurDeptID = Split(Split(str队列名称(i), "|")(1), "_")(0)
            
            Select Case glngBusinessType
                Case TBusinessType.btClinical
                    If lngPreDeptID <> lngCurDeptID Then
                        strQueueNames = strQueueNames & "," & lngCurDeptID
                    End If
                    
                Case TBusinessType.btPacs
                    If InStr(str队列名称(i), "科室队列") Then
                        strQueueNames = strQueueNames & "," & Split(Split(str队列名称(i), "_")(1), ":")(0) & "-" & Split(str队列名称(i), "|")(0)
                    Else
                        strQueueNames = strQueueNames & "," & lngCurDeptID & ":" & Split(str队列名称(i), ":")(1)
                    End If
                    
                Case TBusinessType.btPeis
                    strSql = "select 站点名称 from 体检站点分布 where 执行科室id=[1]"
                    Set rsRecord = gobjComLib.zlDatabase.OpenSQLRecord(strSql, "获取站点名称", lngCurDeptID)
                    
                    If rsRecord.RecordCount > 0 Then
                        strQueueNames = strQueueNames & "," & Nvl(rsRecord!站点名称) & ":" & Split(Split(str队列名称(i), "|")(1), ":")(1)
                    End If
                    
                'Case "" '''''
                '.
            End Select
            
            lngPreDeptID = lngCurDeptID
        Next
    Else
        For i = 0 To UBound(str队列名称)
            lngCurDeptID = Split(Split(str队列名称(i), "|")(1), "_")(0)
            
            Select Case glngBusinessType
                Case TBusinessType.btClinical
                    If lngPreDeptID <> lngCurDeptID Then
                        strQueueNames = strQueueNames & "," & lngCurDeptID
                    End If
                    
                Case TBusinessType.btPacs
                    strQueueNames = strQueueNames & "," & Split(str队列名称(i), "|")(0) & "-" & Split(str队列名称(i), ":")(1)
                    
                Case TBusinessType.btPeis
                    strSql = "select 站点名称 from 体检站点分布 where 执行科室id=[1]"
                    Set rsRecord = gobjComLib.zlDatabase.OpenSQLRecord(strSql, "获取站点名称", lngCurDeptID)
                    
                    If rsRecord.RecordCount > 0 Then
                        strQueueNames = strQueueNames & "," & Nvl(rsRecord!站点名称) & ":" & Split(str队列名称(i), ":")(1)
                    End If
                    
                'Case "" '''''
                '.
            End Select
            
            lngPreDeptID = lngCurDeptID
        Next
    End If
    
    ConvertFormat = strQueueNames
End Function

Private Sub OpenMainCfg()
'打开主配置窗口
    Call frmMain.zlShowMe
End Sub

Public Sub OpenStyleWindow()
'根据配置创建样式窗口对象，并打开对应的样式窗口显示
'blnOpenOldLcd,是否用老版的LCD模式显示
    Dim i As Integer
    Dim lngShowNum As Long
    Dim strShowStyle As String

    lngShowNum = Val(GetSetting("ZLSOFT", G_STR_REGPATH, "窗口数量", 1))
    
    ReDim gobjStyleWindow(lngShowNum) As Object
    
    For i = 1 To lngShowNum
        strShowStyle = GetSetting("ZLSOFT", G_STR_REGPATH & "\" & i, "显示样式", "1-单队列样式")
        
        Select Case Split(strShowStyle, "-")(0)
            Case TShowStyle.ssSingleMan
                Set gobjStyleWindow(i) = New frmStyle_SingleMan

            Case TShowStyle.ssSingleQueue
                Set gobjStyleWindow(i) = New frmStyle_SingleQueue
                
            Case TShowStyle.ssMultiQueue
                Set gobjStyleWindow(i) = New frmStyle_MultiQueue

            Case TShowStyle.ssOld
                Set gobjStyleWindow(i) = Nothing
            
            'Case TShowStyle.ssOther...
            '    Set gobjStyleWindow(i) = New frmStyle_Other...
            '
            '...
        End Select
        
        If Split(strShowStyle, "-")(0) = TShowStyle.ssOld Then
            Call OpenOldLcd(i)
        Else
            Call gobjStyleWindow(i).ISty_Show(i)
        End If
    Next
End Sub

Public Sub CloseStyleWindow()
'关闭所有样式窗口
    Dim i As Integer
    '如果启用的是老版的LCDSHOW则关闭老版的LCDSHOW
    If Not gobjQueueShow Is Nothing Then
        gobjQueueShow.zlclose
        Set gobjQueueShow = Nothing
    End If
    
    If SafeArrayGetDim(gobjStyleWindow) <= 0 Then Exit Sub
    
    For i = 1 To UBound(gobjStyleWindow)
        If Not gobjStyleWindow(i) Is Nothing Then Unload gobjStyleWindow(i)
    Next
End Sub

Private Sub OpenTrayIcon()
'打开托盘图标
    frmTrayIcon.Show
    frmTrayIcon.Hide
End Sub

Public Sub InitOldLCDShow()
'初始化老版本LCD显示部件
    If gobjFile.FileExists("C:\APPSOFT\Apply\zl9LCDShow.dll") Then
        Set gobjQueueShow = DynamicCreate("zl9LCDShow.clsLCDShow", "zl9LCDShow.dll")
    End If
End Sub

