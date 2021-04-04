Attribute VB_Name = "mdlMain"
Option Explicit
'Public gobjDemand As Object                '导航台
Public SplashObj As New frmSplash
Public gcnOracle As ADODB.Connection     '公共数据库连接

Public gstrSysName As String                '系统名称
Public gstrVersion As String                '系统版本
Public gstrAviPath As String                'AVI文件的存放目录

Public gstrUserFlag As String               '当前用户标志(两位表示)，第1位：是否DBA；第2位：系统所有者

Public gstrDbUser As String                 '当前数据库用户
Public gstrStation As String                '本工作站名称
Public gstrMenuSys As String                '系统菜单
Public gobjLogin As Object
Public gobjRegister As Object

'-----------------------------------------
'发行码、注册码、发行码解析串、注册码解析串
Public gstrRegCode As String
Public gstrPublish As String
Public gstrParseRegCode As String
Public gstrParsePublish As String
'-----------------------------------------

Public gstrSystems As String

Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Const GWL_EXSTYLE = (-20)
Public Const WinStyle = &H40000
Public glngOld As Long, glngFormW As Long, glngFormH As Long

'---------------------------------------------------------------
'   授权、菜单、试用版本
'---------------------------------------------------------------
Public Sub Main()
    Dim lngReturn As Long
    Dim StrUnitName As String
    Dim BlnShowFlash As Boolean
    Dim strCode As String
    Dim IntCount As Integer
    Dim StrStyle As String
    Dim rsMenu As ADODB.Recordset
    Dim StrHaveSys As String
    
    If gobjLogin Is Nothing Then
        Set gobjLogin = CreateObject("zlLogin.clsLogin")
    End If
    If gobjLogin Is Nothing Then
        Err = 0: On Error GoTo 0
        MsgBox "创建zlLogin部件失败,可能zlLogin文件丢失或检查是否正确注册！", vbInformation + vbOKOnly, "登陆验证"
        Exit Sub
    End If
    Set gcnOracle = gobjLogin.Login(0, CStr(Command()), , , App.Path & "\" & App.EXEName & ".exe", App.hInstance)
    If gcnOracle Is Nothing Then
        Set gobjLogin = Nothing
        Exit Sub
    End If
    gstrSystems = gobjLogin.Systems
    gstrServerName = gobjLogin.ServerName
    gstrDbUser = gobjLogin.DBUser

    '初始化公共部件
    InitCommon gcnOracle
    '如果发行码无效（为空或为"-"），则退出
    gstrParsePublish = zlRegInfo("产品简名")
    gstrParseRegCode = zlRegInfo("单位名称", , -1)
    gstrSysName = gstrParsePublish & "软件"
    
    StrHaveSys = gstrSystems
    If gstrSystems = "REPORT" Then
        gstrSystems = ""
    Else
        gstrSystems = " (系统 in (" & gstrSystems & ") Or 系统 Is NULL)"
    End If
    If gstrSystems = "" Then
        MsgBox "您没有操作任何系统的权限，程序被迫退出！", vbInformation, gstrSysName
        Exit Sub
    End If
    
    '-------------------------------------------------------------
    '分析菜单及部件
    '-------------------------------------------------------------
    gstrSQL = "SELECT 系统 FROM zlPrograms WHERE 序号=1536 AND 系统 IN (" & StrHaveSys & ")"
    Set gRs = zlDatabase.OpenSQLRecord(gstrSQL, "mdlMain")
    
    If gRs.BOF Then
        MsgBox "您没有操作任何系统的权限,程序被迫退出！", vbInformation, gstrSysName
        Exit Sub
    End If
    
    glngSys = gRs("系统").Value
    If InStr(1, GetPrivFunc(glngSys, 1536), "基本") <= 0 Then
        MsgBox "您没有操作任何系统的权限,程序被迫退出！", vbInformation, gstrSysName
        Exit Sub
    End If
    
    '-------------------------------------------------------------
    '创建同义词
    '-------------------------------------------------------------
    Call CreateSynonyms(glngSys, 1536)
    
    Call GetUserInfo
    Call CodeMan(glngSys, 1536)
    
End Sub

Private Function CreateSynonyms(ByVal lngSys As Long, ByVal lngModul As Long)
    Dim strSQL As String
    '创建模块所需对象的同义词(如果已创建则不会再创建)
    On Error Resume Next
    strSQL = "Zl_Createsynonyms(" & lngSys & ")"
    zlDatabase.ExecuteProcedure strSQL, "创建同义词"
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
    Err = 0
    DoEvents
    With gcnOracle
        If .State = adStateOpen Then .Close
        .Provider = "MSDataShape"
        .Open "Driver={Microsoft ODBC for Oracle};Server=" & strServerName, strUserName, strUserPwd
        If Err <> 0 Then
            '保存错误信息
            strError = Err.Description
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
    
    Err = 0
    On Error GoTo errHand
    
    gstrDbUser = UCase(strUserName)
    gstrServerName = strServerName
    SetDbUser gstrDbUser
    
    gstrConnect = strServerName & ";" & strUserName & ";" & strUserPwd
    
    OraDataOpen = True
    Exit Function
    
errHand:
    If ErrCenter() = 1 Then Resume
    OraDataOpen = False
    Err = 0
End Function


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

Public Function GetUserInfo() As Boolean
'功能：获取登陆用户信息
    Dim rsTmp As New ADODB.Recordset
    
    Set rsTmp = zlDatabase.GetUserInfo
    
    UserInfo.用户名 = gstrDbUser
    UserInfo.姓名 = gstrDbUser
    If Not rsTmp.EOF Then
        UserInfo.ID = rsTmp!ID
        UserInfo.编号 = rsTmp!编号
        UserInfo.简码 = IIf(IsNull(rsTmp!简码), "", rsTmp!简码)
        UserInfo.姓名 = IIf(IsNull(rsTmp!姓名), "", rsTmp!姓名)
        UserInfo.用户名 = IIf(IsNull(rsTmp!用户名), "", rsTmp!用户名)
        UserInfo.部门ID = IIf(IsNull(rsTmp!部门ID), 0, rsTmp!部门ID)
        UserInfo.部门码 = IIf(IsNull(rsTmp!部门码), "", rsTmp!部门码)
        UserInfo.部门 = IIf(IsNull(rsTmp!部门名), "", rsTmp!部门名)
        GetUserInfo = True
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub CodeMan(lngSys As Long, ByVal lngModul As Long)
'功能： 根据主程序指定功能，调用执行相关程序
'参数：
'   lngModul:需要执行的功能序号

    
    If Not GetUserInfo Then
        MsgBox "当前用户未设置对应的人员信息,请与系统管理员联系,先到用户授权管理中设置！", vbInformation, gstrSysName
        Exit Sub
    End If
       
    gstrPrivs = GetPrivFunc(lngSys, lngModul)
    
    glngSys = lngSys

    gstrUnitName = GetUnitName
    gblnInsure = True
    Call gclsInsure.InitOracle(gcnOracle)
    '-------------------------------------------------
        
    frmMainQuery.Show
    
End Sub

Public Sub InitData()

End Sub

Public Function CloseChildWindows(ByVal frmMain As Object, ByVal FrmSon As Object) As Boolean
    '功能:关闭所有子窗口
    
    Dim FrmThis As Form
    
    On Error Resume Next

    CloseChildWindows = True
    
    For Each FrmThis In Forms
        If FrmThis.Caption <> frmMain.Caption And FrmThis.Caption <> FrmSon.Caption Then Unload FrmThis
    Next
    
    '关闭公共部件的窗体
    If CloseChildWindows Then CloseChildWindows = CloseWindows

End Function

Public Sub RunMudal(ByVal lngNO As Long)
    Select Case lngNO
    Case 1
        frmDefTable.Show , gfrmMain
    Case 2
        frmPicture.Show , gfrmMain
    Case 3
        frmDoctor.Show , gfrmMain
    Case 4
        frmAdvice.Show , gfrmMain
    Case 5
        frmDefQuery.Show , gfrmMain
    Case 6
        frmDefTree.Show , gfrmMain
    Case 7
'        If gblnInsure Then
'            If Not gclsInsure.InitInsure(gcnOracle) Then gblnInsure = False
'        End If
        
        Call gclsInsure.InitOracle(gcnOracle)
        
        frmMainQuery.Show , gfrmMain
    Case 8
        Call InitLocPar
        Call InitSysPar
        
        On Error Resume Next
        
        frmselectinfo.Show , gfrmMain
    Case 9
        frmLisPrinterSetup.Show , gfrmMain
    End Select
End Sub


