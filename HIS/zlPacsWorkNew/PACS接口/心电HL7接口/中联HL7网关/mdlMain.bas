Attribute VB_Name = "mdlMain"
Option Explicit

Public gcnOracle As New ADODB.Connection    '公共数据库连接

Public gstrSysName As String                '系统名称
Public gstrVersion As String                '系统版本
Public gstrAviPath As String                'AVI文件的存放目录

Public gstrDbUser As String                 '当前数据库用户
Public gstrDbUserPwd As String              '当前数据库密码
Public gstrServerName As String             '当前数据库服务名

Public glngUserId As Long                   '当前用户id
Public gstrUserCode As String               '当前用户编码
Public gstrUserName As String               '当前用户姓名
Public gstrUserAbbr As String               '当前用户简码

Public glngDeptId As Long                   '当前用户部门id
Public gstrDeptCode As String               '当前用户部门编码
Public gstrDeptName As String               '当前用户部门名称

Public gstrPrivs As String                   '当前用户具有的当前模块的功能
Public gstr单位名称 As String
Public glngSys As Long
'-----------------------------------------
'发行码、注册码、发行码解析串、注册码解析串
Public gstrParsePublish As String
'-----------------------------------------

Public gstrSystems As String

Public gcnAccess As New ADODB.Connection
Public gstrAccessPath As String         'Access数据库的文件路径和文件名前缀，无“.mdb”
Public gstrAccessName As String         'Access数据库的文件路径和文件名

Public gstrLocalIP As String             '存储本机IP地址
Public gblnServiceStart As Boolean         '启动服务

'记录日志
Public gblnProcessLog As Boolean        '是否记录日志，方便分析查找网关的错误
Public glngProcessLogLevel As Long      '记录操作日志的级别，分成3级。1级只记通讯息级别的日志；2级记录通讯和处理的日志；3级记录通讯和处理的详细日志

'接收到的HL7消息的队列
Public gstrMsgQueue() As String         '消息内容队列
Public gblnQueueBusy As Boolean         '记录当前操作是否在处理队列，一次只能有一个程序处理队列
Public gintQueueIndex As Integer        '记录队列中第一个消息的索引
Public gblnMsgProcessing As Boolean     '正在处理消息
Public gintTimeOut() As Integer         '记录每一个消息连接的时间
Public gintTimeOutMax As Integer        '超时的最大时间

'消息接收参数
Public gintInputDataType As Integer     '接收消息的方式，默认是0：0-socket方式，1-文件方式；
Public gstrFileDir As String            '文件方式接收消息的路径
Public gstrFileSuffix As String         '文件方式接收消息的后缀
Public gstrFileBackupDir As String      '文件方式接收消息的备份路径

Public gstrRegPath As String            '程序参数注册表路径

Public gzlDatabase As Object           '公共数据库模块 zlComLib的zlDatabase
Public gzlComLib As Object             '公共数据库模块 zlComLib的zlDatabase

'-------------兼容10.35.10之前的版本参数--------------
Public gblnBefore3510 As Boolean       '区分10.35.10前后版本。True=10.35.10之前版本,不使用zlRegister，初始化comlib时需要SetDbUser和RegCheck
Public SplashObj As New frmSplash
Public gstrStation As String           '本工作站名称
Public gstrParseRegCode As String
'-------------兼容10.35.10之前的版本参数--------------


'---------------------------------------------------------------
'   授权、菜单、试用版本
'---------------------------------------------------------------
Public Sub Main()
    Dim objRegister As Object         '10.35.10之后的注册对象
    
    Set gzlDatabase = CreateObject("zl9ComLib.clsDatabase")
    Set gzlComLib = CreateObject("zl9ComLib.clsComLib")
    
    If App.PrevInstance Then
        MsgBox "HL7网关服务已经启动，不能再次运行。", vbInformation, "警告"
        Exit Sub
    End If
    
    On Error Resume Next
    '先通过zlRegister部件判断是不是10.35.10之后的版本，这个版本之后，登录数据库的密码不一样了
    Set objRegister = GetObject("", "zlRegister.clsRegister")
    If objRegister Is Nothing Then
        gblnBefore3510 = True   '35.10之前的版本
    Else
        gblnBefore3510 = False
        Set objRegister = Nothing
    End If
    
    err.Clear
    On Error GoTo err

    If gblnBefore3510 Then
        If LoginBefore3510 = False Then Exit Sub
    Else
        If LoginAfter3510 = False Then Exit Sub
    End If
   
    CodeMan 2000
    Exit Sub
err:
    MsgBox "启动网关出现错误，错误描述是：" & err.Description, vbOKOnly
End Sub

Public Sub CodeMan(ByVal lngModul As Long)
    '------------------------------------------------
    '功能： 根据主程序指定功能，调用执行相关程序
    '参数：
    '   lngModul:需要执行的功能序号
    '返回：
    '------------------------------------------------
    Dim rsUser As ADODB.Recordset
    
    gstrAviPath = GetSetting(appName:="ZLSOFT", Section:="注册信息", Key:=UCase("gstrAviPath"), Default:="")
    gstrSysName = GetSetting(appName:="ZLSOFT", Section:="注册信息", Key:=UCase("gstrSysName"), Default:="")
    gstrVersion = GetSetting(appName:="ZLSOFT", Section:="注册信息", Key:=UCase("gstrVersion"), Default:="")
    gstr单位名称 = gzlComLib.GetUnitName()
    
    '提取用户的信息
    Set rsUser = gzlDatabase.GetUserInfo
    If rsUser.RecordCount <> 0 Then
        glngUserId = Nvl(rsUser!ID)
        gstrUserCode = Nvl(rsUser!编号)
        gstrUserName = Nvl(rsUser!姓名)
        gstrUserAbbr = Nvl(rsUser!简码)
        glngDeptId = Nvl(rsUser!部门ID)
        gstrDeptCode = Nvl(rsUser!部门码)
        gstrDeptName = Nvl(rsUser!部门名)
    Else
        glngUserId = 0
        gstrUserCode = ""
        gstrUserName = ""
        gstrUserAbbr = ""
        glngDeptId = 0
        gstrDeptCode = ""
        gstrDeptName = ""
    End If
    
    gstrPrivs = gzlComLib.GetPrivFunc(glngSys, lngModul)
    '-------------------------------------------------
    
    Select Case lngModul
        Case 2000
            frmHL7Main.Show
    End Select
End Sub

Public Function TranPasswd(strOld As String) As String
    '------------------------------------------------
    '功能： 密码转换函数
    '参数：
    '   strOld：原密码
    '返回： 加密生成的密码
    '------------------------------------------------
    Dim iBit As Integer, strBit As String
    Dim strNew As String
    If Len(Trim(strOld)) = 0 Then TranPasswd = "": Exit Function
    strNew = ""
    For iBit = 1 To Len(Trim(strOld))
        strBit = UCase(Mid(Trim(strOld), iBit, 1))
        Select Case (iBit Mod 3)
        Case 1
            strNew = strNew & _
                Switch(strBit = "0", "W", strBit = "1", "I", strBit = "2", "N", strBit = "3", "T", strBit = "4", "E", strBit = "5", "R", strBit = "6", "P", strBit = "7", "L", strBit = "8", "U", strBit = "9", "M", _
                   strBit = "A", "H", strBit = "B", "T", strBit = "C", "I", strBit = "D", "O", strBit = "E", "K", strBit = "F", "V", strBit = "G", "A", strBit = "H", "N", strBit = "I", "F", strBit = "J", "J", _
                   strBit = "K", "B", strBit = "L", "U", strBit = "M", "Y", strBit = "N", "G", strBit = "O", "P", strBit = "P", "W", strBit = "Q", "R", strBit = "R", "M", strBit = "S", "E", strBit = "T", "S", _
                   strBit = "U", "T", strBit = "V", "Q", strBit = "W", "L", strBit = "X", "Z", strBit = "Y", "C", strBit = "Z", "X", True, strBit)
        Case 2
            strNew = strNew & _
                Switch(strBit = "0", "7", strBit = "1", "M", strBit = "2", "3", strBit = "3", "A", strBit = "4", "N", strBit = "5", "F", strBit = "6", "O", strBit = "7", "4", strBit = "8", "K", strBit = "9", "Y", _
                   strBit = "A", "6", strBit = "B", "J", strBit = "C", "H", strBit = "D", "9", strBit = "E", "G", strBit = "F", "E", strBit = "G", "Q", strBit = "H", "1", strBit = "I", "T", strBit = "J", "C", _
                   strBit = "K", "U", strBit = "L", "P", strBit = "M", "B", strBit = "N", "Z", strBit = "O", "0", strBit = "P", "V", strBit = "Q", "I", strBit = "R", "W", strBit = "S", "X", strBit = "T", "L", _
                   strBit = "U", "5", strBit = "V", "R", strBit = "W", "D", strBit = "X", "2", strBit = "Y", "S", strBit = "Z", "8", True, strBit)
        Case 0
            strNew = strNew & _
                Switch(strBit = "0", "6", strBit = "1", "J", strBit = "2", "H", strBit = "3", "9", strBit = "4", "G", strBit = "5", "E", strBit = "6", "Q", strBit = "7", "1", strBit = "8", "X", strBit = "9", "L", _
                   strBit = "A", "S", strBit = "B", "8", strBit = "C", "5", strBit = "D", "R", strBit = "E", "7", strBit = "F", "M", strBit = "G", "3", strBit = "H", "A", strBit = "I", "N", strBit = "J", "F", _
                   strBit = "K", "O", strBit = "L", "4", strBit = "M", "K", strBit = "N", "Y", strBit = "O", "D", strBit = "P", "2", strBit = "Q", "T", strBit = "R", "C", strBit = "S", "U", strBit = "T", "P", _
                   strBit = "U", "B", strBit = "V", "Z", strBit = "W", "0", strBit = "X", "V", strBit = "Y", "I", strBit = "Z", "W", True, strBit)
        End Select
    Next
    TranPasswd = strNew

End Function

Public Function OraDataOpen(ByVal strServerName As String, ByVal strUserName As String, ByVal strUserPwd As String) As Boolean
    '------------------------------------------------
    '功能： 打开指定的数据库
    '参数：
    '   strServerName：主机字符串
    '   strUserName：用户名
    '   strUserPwd：密码
    '返回： 数据库打开成功，返回true；失败，返回false
    '------------------------------------------------
    Dim strSQL As String
    Dim strError As String
    
    On Error Resume Next
    err = 0
    DoEvents
    With gcnOracle
        If .State = adStateOpen Then .Close
        .Provider = "MSDataShape"
        .Open "Driver={Microsoft ODBC for Oracle};Server=" & strServerName, strUserName, strUserPwd
        If err <> 0 Then
            '保存错误信息
            strError = err.Description
            If InStr(strError, "自动化错误") > 0 Then
                MsgBox "连接串无法创建，请检查数据访问部件是否正常安装。", vbInformation, gstrSysName
            ElseIf InStr(strError, "ORA-12154") > 0 Then
                MsgBox "无法分析服务器名，" & vbCrLf & "请检查在Oracle配置中是否存在该本地网络服务名（主机字符串）。", vbInformation, gstrSysName
            ElseIf InStr(strError, "ORA-12541") > 0 Then
                MsgBox "无法连接，请检查服务器上的Oracle监听器服务是否启动。", vbInformation, gstrSysName
            ElseIf InStr(strError, "ORA-01033") > 0 Then
                MsgBox "ORACLE正在初始化或在关闭，请稍候再试。", vbInformation, gstrSysName
            Else
                MsgBox "由于用户、口令或服务器指定错误，无法注册。", vbInformation, gstrSysName
            End If
            
            OraDataOpen = False
            Exit Function
        End If
    End With
    
    err = 0
    On Error GoTo errHand
    
    gstrServerName = strServerName
    gstrDbUserPwd = strUserPwd
    gstrDbUser = UCase(strUserName)
    gzlComLib.SetDbUser gstrDbUser
    OraDataOpen = True
    Exit Function
    
errHand:
    If gzlComLib.ErrCenter() = 1 Then Resume
    OraDataOpen = False
    err = 0
End Function

'-------------------------------------------------------------------------------------------------
'sample usage - Debug.Print GetKeyValue(HKEY_CLASSES_ROOT, "COMCTL.ListviewCtrl.1\CLSID", "")
'-------------------------------------------------------------------------------------------------
Public Function GetKeyValue(KeyRoot As Long, KeyName As String, SubKeyRef As String) As String
'功能：读注册表
    Dim i As Long                                           ' 循环计数器
    Dim rc As Long                                          ' 返回代码
    Dim hKey As Long                                        ' 处理打开的注册表关键字
    Dim hDepth As Long                                      '
    Dim sKeyVal As String
    Dim lKeyValType As Long                                 ' 注册表关键字数据类型
    Dim tmpVal As String                                    ' 注册表关键字的临时存储器
    Dim KeyValSize As Long                                  ' 注册表关键字变量尺寸
    
    ' 在 KeyRoot {HKEY_LOCAL_MACHINE...} 下打开注册表关键字
    '------------------------------------------------------------
    rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey) ' 打开注册表关键字
    
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' 处理错误...
    
    tmpVal = String$(1024, 0)                             ' 分配变量空间
    KeyValSize = 1024                                       ' 标记变量尺寸
    
    '------------------------------------------------------------
    ' 检索注册表关键字的值...
    '------------------------------------------------------------
    rc = RegQueryValueEx(hKey, SubKeyRef, 0, _
                         lKeyValType, tmpVal, KeyValSize)    ' 获得/创建关键字的值
                        
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' 错误处理
      
    tmpVal = Left$(tmpVal, InStr(tmpVal, Chr(0)) - 1)

    '------------------------------------------------------------
    ' 决定关键字值的转换类型...
    '------------------------------------------------------------
    Select Case lKeyValType                                  ' 搜索数据类型...
    Case REG_SZ, REG_EXPAND_SZ                              ' 字符串注册表关键字数据类型
        sKeyVal = tmpVal                                     ' 复制字符串的值
    Case REG_DWORD                                          ' 四字节注册表关键字数据类型
        For i = Len(tmpVal) To 1 Step -1                    ' 转换每一位
            sKeyVal = sKeyVal + Hex(Asc(Mid(tmpVal, i, 1)))   ' 一个字符一个字符地生成值。
        Next
        sKeyVal = Format$("&h" + sKeyVal)                     ' 转换四字节为字符串
    End Select
    
    GetKeyValue = sKeyVal                                   ' 返回值
    rc = RegCloseKey(hKey)                                  ' 关闭注册表关键字
    Exit Function                                           ' 退出
    
GetKeyError:    ' 错误发生过后进行清除...
    GetKeyValue = vbNullString                              ' 设置返回值为空字符串
    rc = RegCloseKey(hKey)                                  ' 关闭注册表关键字
End Function

Public Sub CheckDBConnect()
    '如果数据库断开，则重新连接数据库
    On Error GoTo ConnErr
    If gcnOracle.State <> 1 Then
        gcnOracle.Provider = "MSDataShape"
        gcnOracle.Open "Driver={Microsoft ODBC for Oracle};Server=" & gstrServerName, gstrDbUser, gstrDbUserPwd
    End If
    gzlDatabase.OpenSQLRecord "select '测试'  from dual", "测试连接是否成功"
    Exit Sub
ConnErr:
    On Error Resume Next
    If gcnOracle.State = 1 Then
        gcnOracle.Close
    End If
End Sub

Private Function LoginBefore3510() As Boolean
'10.35.10之前的登录方法

    Dim BlnShowFlash As Boolean
    Dim StrUnitName As String
    Dim intCount As Integer
    Dim lngReturn As Long
    Dim strCode As String
    
    
    LoginBefore3510 = False
    
    '为实现XP风格，在显示窗体前必须执行该函数
    Call InitCommonControls
    
    BlnShowFlash = False
    Load SplashObj
    '由注册表中获取用户注册相关信息,如果用户单位名称不为空,则显示闪现窗体
    StrUnitName = GetSetting("ZLSOFT", "注册信息", "单位名称", "")
    gstrSysName = GetSetting("ZLSOFT", "注册信息", "提示", "")
    If StrUnitName <> "" Then
        With SplashObj
            '有两处需要处理
            Call gzlComLib.ApplyOEM_Picture(.ImgIndicate, "Picture")
            Call gzlComLib.ApplyOEM_Picture(.imgPic, "PictureB")
            .Show
            .lblGrant = StrUnitName
            StrUnitName = GetSetting("ZLSOFT", "注册信息", "开发商", "")
            If Trim(StrUnitName) = "" Then
                .Label3.Visible = False
                .lbl开发商.Visible = False
            Else
                .lbl开发商.Caption = ""
                For intCount = 0 To UBound(Split(StrUnitName, ";"))
                    .lbl开发商.Caption = .lbl开发商.Caption & Split(StrUnitName, ";")(intCount) & vbCrLf
                Next
            End If
            .LblProductName = GetSetting("ZLSOFT", "注册信息", "产品全称", "")
            .lbl技术支持商 = GetSetting("ZLSOFT", "注册信息", "技术支持商", "")
        End With
        
        BlnShowFlash = True
        DoEvents
    End If
    
    gstrStation = Space(200)
    lngReturn = GetComputerName(gstrStation, 200)
    gstrStation = Trim(gstrStation)
    If Len(gstrStation) > 1 Then
        gstrStation = Left(gstrStation, Len(gstrStation) - 1)
    Else
        gstrStation = "..."
    End If
    
    '用户注册
    frmUserLogin.Show 1
    If gcnOracle.State <> adStateOpen Then
        Unload frmUserLogin
        Unload SplashObj
        Exit Function
    End If
    
    '初始化公共部件
    gzlComLib.InitCommon gcnOracle
    If gzlComLib.RegCheck = False Then
        Unload SplashObj
        Exit Function
    End If
    
    '如果发行码无效（为空或为"-"），则退出
    gstrParsePublish = gzlComLib.zlRegInfo("产品简名")
    gstrParseRegCode = gzlComLib.zlRegInfo("单位名称", , -1)
    
    gstrSysName = gstrParsePublish & "软件"
    SaveSetting "ZLSOFT", "注册信息", "提示", gstrSysName
    SaveSetting "ZLSOFT", "注册信息", UCase("gstrSysName"), gstrSysName
    gstrVersion = App.Major & "." & App.Minor & "." & App.Revision
    SaveSetting "ZLSOFT", "注册信息", UCase("gstrVersion"), gstrVersion
    gstrAviPath = App.Path & "\附加文件"
    SaveSetting "ZLSOFT", "注册信息", UCase("gstrAviPath"), gstrAviPath
    
    With SplashObj
        If BlnShowFlash = False Then
            .lblGrant = gstrParseRegCode
            .lbl技术支持商.Caption = gzlComLib.zlRegInfo("技术支持商", , -1)
            .LblProductName = gzlComLib.zlRegInfo("产品标题")
            
            strCode = gzlComLib.zlRegInfo("产品开发商", , -1)
            .lbl开发商.Caption = ""
            For intCount = 0 To UBound(Split(strCode, ";"))
                .lbl开发商.Caption = .lbl开发商.Caption & Split(strCode, ";")(intCount) & vbCrLf
            Next
            Call gzlComLib.ApplyOEM_Picture(.ImgIndicate, "Picture")
            .Show
            BlnShowFlash = True
        End If
        DoEvents
    End With
    
    '将用户注册相关信息写入注册表,供下次启动时显示
    SaveSetting "ZLSOFT", "注册信息", "单位名称", gstrParseRegCode
    SaveSetting "ZLSOFT", "注册信息", "产品全称", gzlComLib.zlRegInfo("产品标题")
    SaveSetting "ZLSOFT", "注册信息", "产品名称", gzlComLib.zlRegInfo("产品简名")
    SaveSetting "ZLSOFT", "注册信息", "技术支持商", gzlComLib.zlRegInfo("技术支持商", , -1)
    SaveSetting "ZLSOFT", "注册信息", "开发商", gzlComLib.zlRegInfo("产品开发商", , -1)
    SaveSetting "ZLSOFT", "注册信息", "WEB支持商简名", gzlComLib.zlRegInfo("支持商简名")
    SaveSetting "ZLSOFT", "注册信息", "WEB支持EMAIL", gzlComLib.zlRegInfo("支持商MAIL")
    SaveSetting "ZLSOFT", "注册信息", "WEB支持URL", gzlComLib.zlRegInfo("支持商URL")

    gstrSystems = " (系统 =100 Or 系统 Is NULL)"
    glngSys = 100
    '-------------------------------------------------------------
    '创建同义词
    '-------------------------------------------------------------
    gzlDatabase.ExecuteProcedure "Zl_Createsynonyms(" & glngSys & ")", "创建同义词"
    
    '-------------------------------------------------------------
    '选择调用不同风格导航台
    '-------------------------------------------------------------
    On Error Resume Next
    err = 0
    
    Unload SplashObj
    
    LoginBefore3510 = True
End Function

Private Function LoginAfter3510() As Boolean
'10.35.10之后的登录方法

    Dim objLogin As Object
    Dim strCommand As String
    
    LoginAfter3510 = False
    
    Set objLogin = DynamicCreate("zlLogin.clsLogin", "zlLogin.dll")
    If objLogin Is Nothing Then Exit Function
    
    '为实现XP风格，在显示窗体前必须执行该函数
    Call InitCommonControls
    
    Set gcnOracle = objLogin.Login(0, CStr(Command()))
    
    If gcnOracle Is Nothing Then Exit Function
    If gcnOracle.ConnectionString = "" Then Exit Function

    gstrUserName = objLogin.DBUser
    
    '初始化公共部件
    gzlComLib.InitCommon gcnOracle
    
    gstrParsePublish = gzlComLib.zlRegInfo("产品简名")
    gstrSysName = gstrParsePublish & "软件"
    
    gstrSystems = " (系统 =100 Or 系统 Is NULL)"
    glngSys = 100
    
    LoginAfter3510 = True
End Function
