Attribute VB_Name = "mdlMain"
Option Explicit

Public ZlBrowerDll As Object                '导航台
Public gcnOracle As ADODB.Connection     '公共数据库连接
Public gobjRelogin As New clsRelogin  '重新启动的类的对象实例
Public gobjWait As Object '展示非模态窗体后可以使程序不退出的对象

Public gstrSysName As String                '系统名称
Public gstrVersion As String                '系统版本
Public gstrAviPath As String                'AVI文件的存放目录
Public gstrUserFlag As String               '当前用户标志(两位表示)，第1位：是否DBA(由于访问DBA_ROLE_PRIVS视图产生的IO较高，暂时没有判断,有需要使用时再判断)；第2位：系统所有者

Public gstrDbUser As String                 '当前数据库用户
Public glngUserId As Long                   '当前用户id
Public gstrUserCode As String               '当前用户编码
Public gstrUserName As String               '当前用户姓名
Public gstrUserAbbr As String               '当前用户简码

Public glngDeptId As Long                   '当前用户部门id
Public gstrDeptCode As String               '当前用户部门编码
Public gstrDeptName As String               '当前用户部门名称

Public gstrStation As String                '本工作站名称
Public gstrMenuSys As String                '系统菜单

Public gstrSystems As String

'---------------------------------------------------------------
'-注册表 API 声明...
'---------------------------------------------------------------
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Public Const GWL_EXSTYLE = (-20)
Public Const WinStyle = &H40000
Public Const SWP_NOSIZE = &H1
Public Const SWP_SHOWWINDOW = &H40
Public Const HWND_TOPMOST = -1

Public Enum Register
    注册信息
    私有模块
    私有全局
    公共模块
    公共全局
End Enum

'---------------------------------------------------------------
'启动时间，用以判断闪现屏幕的等待时间
'---------------------------------------------------------------
Public gdtStart As Long

'---------------------------------------------------------------
'   授权、菜单、试用版本
'---------------------------------------------------------------

'----------------------------------------------------------------------------------------------------------------------------------------------------------------
'功能:设置相关进程处理的API声明:2008-10-30 11:34:11:刘兴宏
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function CreateToolhelpSnapshot Lib "kernel32" Alias "CreateToolhelp32Snapshot" (ByVal lFlags As Long, ByVal lProcessId As Long) As Long
Private Declare Function ProcessFirst Lib "kernel32" Alias "Process32First" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function ProcessNext Lib "kernel32" Alias "Process32Next" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Public Declare Function EnumWindows Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Public Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Type PROCESSENTRY32
      lSize             As Long
      lUsage            As Long
      lProcessId        As Long
      lDefaultHeapId    As Long
      lModuleId         As Long
      lThreads          As Long
      lParentProcessId  As Long
      lPriClassBase     As Long
      lFlags            As Long
      sExeFile          As String * 1024
End Type
Private Const PROCESS_TERMINATE = &H1
Public gcll_His_PId As Collection        '存储相关的进程信息:array(进程名称,PID,窗口个数),"K"+进程数

#Const SYS_TRYUSE = "正式" '正式/试用
Private Sub SetAppBusyState()
'当其他进程对象未创建完成时，替换在执行主进程功能时弹出的“部件被挂起”对话框
On Error Resume Next
    App.OleServerBusyMsgTitle = App.ProductName
    App.OleRequestPendingMsgTitle = App.ProductName
    
    App.OleServerBusyMsgText = "相关组件正在创建，请耐心等待。"
    App.OleRequestPendingMsgText = "相关组件正创建，请耐心等待。"
    
    App.OleServerBusyTimeout = 3000
    App.OleRequestPendingTimeout = 10000
    Err.Clear
End Sub

Public Sub Main()
    Dim rsMenu As ADODB.Recordset
    Dim objRIS As Object
    Dim strStyle As String
    
    '功能:杀异常进程:2008-10-30(BUG:14365)
    Call zlKillHISPID
    Set gcnOracle = gobjRelogin.Login(0, CStr(Command()))
    If gcnOracle Is Nothing Then
        Set gobjRelogin = Nothing
        Exit Sub
    End If
    gstrDeptName = gobjRelogin.DeptName
    gstrDbUser = gobjRelogin.DBUser
    
    '写入本次启动程序的EXE文件名
    Call SaveSetting("ZLSOFT", "公共全局", "执行文件", App.EXEName & ".exe")
    gstrVersion = App.Major & "." & App.Minor & "." & App.Revision
    SaveSetting "ZLSOFT", "注册信息", UCase("gstrVersion"), gstrVersion
    gstrAviPath = App.Path & "\附加文件"
    SaveSetting "ZLSOFT", "注册信息", UCase("gstrAviPath"), gstrAviPath
    SaveSetting "ZLSOFT", "公共全局", "程序路径", App.Path & "\" & App.EXEName & ".exe"
    
    
    On Error Resume Next
    Set objRIS = CreateObject("zl9XWInterface.clsHISInner")
    Err.Clear: On Error GoTo 0
    If Not objRIS Is Nothing Then
        Call objRIS.SaveDBConnectInfo(gobjRelogin.InputUser, gobjRelogin.InputPwd, gobjRelogin.ServerName, gobjRelogin.IsTransPwd)
    End If
    gstrSystems = gobjRelogin.Systems
    Call GetUserInfo(IIf(gobjRelogin.Systems = "REPORT", 0, Replace(gobjRelogin.Systems, "'", "")))
    
    '读取登录变量
    gstrUserFlag = IIf(gobjRelogin.IsSysOwner, "01", "00")
    gstrStation = OS.ComputerName
    If gstrStation = "" Then
        gstrStation = "..."
    End If
    
    '-------------------------------------------------------------
    '分析菜单及部件
    '-------------------------------------------------------------
    Set rsMenu = MenuGranted(gobjRelogin.MenuGroup)
    If rsMenu.EOF Then
        MsgBox "您没有操作任何系统的权限,程序被迫退出！", vbInformation, gstrSysName
        Set gobjRelogin = Nothing
        Exit Sub
    End If
    '-------------------------------------------------------------
    '不用再创建公共同义词，公共的在安装和升级时创建，私有的在进入模块时调用
    '-------------------------------------------------------------
    '-------------------------------------------------------------
    '选择调用不同风格导航台
    '-------------------------------------------------------------
    On Error Resume Next
    Err = 0
    
    strStyle = zlDatabase.GetPara("导航台", , , "zlBrw")
    Set ZlBrowerDll = CreateObject(strStyle & ".Cls" & Mid(strStyle, 3))
    If Err <> 0 Then
        If strStyle = "ZLBRW" Then
            MsgBox "启动失败，主程序的相关文件丢失，请重新安装！", vbInformation, gstrSysName
            Set gobjRelogin = Nothing
            Exit Sub
        Else
            Err = 0
            Set ZlBrowerDll = CreateObject("ZLBRW.ClsBrw")
            If Err <> 0 Then
                MsgBox "启动失败，主程序的相关文件丢失，请重新安装！", vbInformation, gstrSysName
                Set gobjRelogin = Nothing
                Exit Sub
            End If
        End If
    End If
    '升级本地注册表参数值
    Call UpdateParameters
    '以下两句防止程序终止
    Set gobjWait = frmSelClient
    Load gobjWait
    Call ZlBrowerDll.SetEnvironment(gstrSysName, gstrVersion, gstrAviPath, _
                          gstrUserFlag, gstrDbUser, glngUserId, _
                          gstrUserCode, gstrUserName, gstrUserAbbr, _
                          glngDeptId, gstrDeptCode, gstrDeptName, _
                          gstrStation, gstrMenuSys, CStr(Command()))
    Call ZlBrowerDll.InitBrower(gobjRelogin, gcnOracle, rsMenu)
End Sub

Public Function MenuGranted(ByVal strMenuGroup As String) As ADODB.Recordset
    '-------------------------------------------------------------
    '功能：分析授权使用并安装的部件，进而产生授权使用的菜单集合
    '参数：注册码
    '-------------------------------------------------------------
    Dim ArrCommand
    Dim strSQL As String
    Dim rsTemp As New ADODB.Recordset
    Dim strCodes As String
    Dim strObjs As String
    Dim intCount As Integer
    Dim strSystems As String
    Dim BlnOnlySys As Boolean '只有报表系统
    Dim strSYS As String
    
    On Error GoTo errH
    BlnOnlySys = (gstrSystems = "REPORT")
    If BlnOnlySys Then
        strSystems = "'0'"
        strSYS = "0"
    Else
        strSystems = Replace(gstrSystems, "','", ",")
        strSYS = Replace(gstrSystems, "'", "")
    End If
    
    If strMenuGroup <> "" Then gstrMenuSys = strMenuGroup
    strObjs = GetSetting("ZLSOFT", "注册信息", "本机部件", "")
    If strObjs = "" Then strObjs = "'Zl9Common'"
    strObjs = Replace(strObjs, "','", ",")
    If OS.IsDesinMode Then
        strSQL = "Select 层次, ID As 编号, Nvl(上级id, 0) As 上级, 标题, Decode(Nvl(短标题,'空'),'空',标题,短标题) as 短标题, 快键, 说明, Nvl(模块, 0) As 模块, Nvl(系统, 0) As 系统, " & _
                 "        Nvl(图标, 0) As 图标, 部件, Decode(Upper(RTrim(部件)), 'ZL9REPORT', 1, 0) As 报表 " & _
                 " From Table(Cast(ZLTOOLS.f_Reg_Menu([1], [2], [3]) As ZLTOOLS.t_Menu_Rowset)) " & _
                 " Union " & _
                 " Select A.层次, A.ID, Nvl(上级id, 0) As 上级, A.标题, Decode(Nvl(A.短标题,'空'),'空',A.标题,A.短标题) As 短标题, A.快键, A.说明, Nvl(A.模块, 0) As 模块, " & _
                 "        Nvl(A.系统, 0) As 系统, Nvl(图标, 0) As 图标, C.部件, Decode(C.部件, 'ZL9REPORT', 1, 0) As 报表 " & _
                 " From (Select Level As 层次, ID, 上级id, 标题, 短标题, 快键, 说明, Nvl(模块,0) 模块, 系统, 图标 " & _
                 "        From zlMenus " & _
                 "        Where 组别 = [1] And Nvl(系统, 0) IN(" & strSYS & ") " & _
                 "        Start With 上级id Is Null " & _
                 "        Connect By Prior ID = 上级id) A, " & _
                 "      (Select 系统, Nvl(模块,0) 模块 " & _
                 "        From zlMenus A " & _
                 "        Where 组别 = [1] And Nvl(系统, 0) IN (" & strSYS & ") " & _
                 "        Minus " & _
                 "        Select 系统 * 100, 序号 From Zlregfunc Where 系统 * 100 IN (" & strSYS & ")) B," & _
                 "      (select 系统, Upper(RTrim(部件)) as 部件,序号 From zlPrograms ) C " & _
                 " Where A.系统 = B.系统 And A.模块 = B.模块 And A.模块 = C.序号(+) and A.系统 = C.系统"

    Else
        strSQL = "SELECT 层次, Id AS 编号, Nvl(上级id, 0) AS 上级, 标题, Decode(Nvl(短标题,'空'),'空',标题,短标题) As 短标题, 快键, 说明, Nvl(模块, 0) AS 模块, Nvl(系统, 0) AS 系统, " & _
                 "        Nvl(图标, 0) AS 图标, 部件, Decode(Upper(Rtrim(部件)), 'ZL9REPORT', 1, 0) AS 报表 " & _
                 " FROM TABLE(CAST(Zltools.f_Reg_Menu([1], [2], [3]) As " & _
                 " Zltools.t_Menu_Rowset)) "
    End If
    '实现报表按编号排序,模块号可能是zlReports.程序id,也可能是zlRPTGroups.程序id,优先zlReports
    '只获取报表发布到模块的报表
    strSQL = "Select 层次, 编号, 上级, 标题, 短标题, 快键, 说明, 模块, 系统, 图标, 部件, 报表, 报表编号, 是否停用" & vbNewLine & _
                    "From (Select a.*, Decode(a.报表, 0, Null, Nvl(b.编号, c.编号)) 报表编号, Nvl(b.是否停用, 0) + Nvl(c.是否停用, 0) 是否停用" & vbNewLine & _
                    "       From (" & strSQL & ")  a," & vbNewLine & _
                    "            (Select Nvl(b.系统, 0) 系统, b.程序id, b.编号, b.是否停用" & vbNewLine & _
                    "              From Zlprograms a, Zlreports b" & vbNewLine & _
                    "              Where Nvl(a.系统, 0) = Nvl(b.系统, 0) And a.序号 = Nvl(b.程序id, 0) And Upper(a.部件) = 'ZL9REPORT') b, (Select 编号, Nvl(系统, 0) 系统, 程序id, 是否停用 From Zlrptgroups) c" & vbNewLine & _
                    "       Where a.系统 = b.系统(+) And a.模块 = b.程序id(+) And a.系统 = c.系统(+) And a.模块 = c.程序id(+))" & vbNewLine & _
                    "Order By 层次, 报表, 系统, 模块, 编号, 报表编号"

    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, gstrSysName, gstrMenuSys, Replace(strSystems, "'", ""), Replace(strObjs, "'", ""))

    Set MenuGranted = rsTemp
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function GetUserInfo(ByVal strSystems As String)
    Dim rsTmp As ADODB.Recordset, rsUser As ADODB.Recordset
    Dim strSQL As String, i As Integer
    On Error GoTo errH
    '读用户信息赋予公共，便于其他程序使用
    strSQL = "Select S.所有者" & _
            " From zlSystems S,(Select Distinct owner From All_Tables Where Table_Name='部门表') D" & _
            " Where S.所有者=D.Owner And S.编号 In (" & strSystems & ") Order by S.编号"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "读取所有者")
    
    With rsTmp
        If Not .EOF Then
            '因为可能该用户具有多个系统的身份，所以循环取身份
            glngUserId = 0 '当前用户id
            gstrUserCode = "" '当前用户编码
            gstrUserName = "" '当前用户姓名
            gstrUserAbbr = "" '当前用户简码
            glngDeptId = 0 '当前用户部门id
            gstrDeptCode = "" '当前用户
            gstrDeptName = "" '当前用户
            
            For i = 1 To .RecordCount
                strSQL = "Select R.人员ID,R.部门ID,D.编码 as 部门编码,D.名称 as 部门名称,P.编号,P.姓名,P.简码" & _
                        " From " & !所有者 & ".上机人员表 U," & !所有者 & ".人员表 P," & !所有者 & ".部门表 D," & !所有者 & ".部门人员 R" & _
                        " Where U.人员ID = P.ID And R.部门ID = D.ID And P.ID=R.人员ID and U.用户名=[1] And (P.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or P.撤档时间 Is Null) and R.缺省=1"
                Set rsUser = zlDatabase.OpenSQLRecord(strSQL, "读取人员信息", gstrDbUser)
                                
                If Not rsUser.EOF Then
                    glngUserId = rsUser!人员ID '当前用户id
                    gstrUserCode = rsUser!编号 '当前用户编码
                    gstrUserName = IIf(IsNull(rsUser!姓名), "", rsUser!姓名) '当前用户姓名
                    gstrUserAbbr = IIf(IsNull(rsUser!简码), "", rsUser!简码) '当前用户简码
                    glngDeptId = rsUser!部门ID '当前用户部门id
                    gstrDeptCode = rsUser!部门编码 '当前用户
                    gstrDeptName = rsUser!部门名称 '当前用户
                    Exit For
                End If
                .MoveNext
            Next
        End If
    End With
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

'**********************************************************************************************************************
'功能:以下处理相关进程的函数
'编制:刘兴洪
'日期:2008-10-30 11:38:58
Public Function zlKillHISPID() As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '功能:杀死所有HIS启动程序的移常进程(杀的条件是:所有ZLHIS+.exe的进程中无任何窗口)
    '入参:
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2008-10-30 11:06:16
    '-----------------------------------------------------------------------------------------------------------
    Dim lngProcess As Long, i As Long
    
    zlKillHISPID = False
    Err = 0: On Error GoTo errHand:
    '第一步:需要处理相关的ZLHIS的相关进程
    Set gcll_His_PId = New Collection
    If zlHISPidToCollect(gcll_His_PId) = False Then zlKillHISPID = True: Exit Function  '如果存在相关的错误，就直接返回了
    If gcll_His_PId Is Nothing Then zlKillHISPID = True: Exit Function
    If gcll_His_PId.Count = 0 Then zlKillHISPID = True: Exit Function
    
    '第二步:需要处理相关ZLHIS的相关进程的相关窗口个数,这样才好判断出相关的进程是否存在异常,出现异常的，就得杀掉
    Call EnumWindows(AddressOf EnumWindowsProc, 0&)
    For i = 1 To gcll_His_PId.Count
        If Val(gcll_His_PId(i)(2)) <= 1 Then
            '肯定窗口数小于1或零,那么肯定有异常，需要杀死他
            If Val(gcll_His_PId(i)(1)) <> 0 Then
                '可能未成功，暂无处理此种情况
                Call TerminatePID(Val(gcll_His_PId(i)(1)))
            End If
        End If
    Next
    zlKillHISPID = True
errHand:
End Function

Private Function EnumWindowsProc(ByVal hwnd As Long, ByVal lParam As Long) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '功能:获取所有窗口符合HIS的进程的窗口
    '入参:
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2008-10-30 10:26:02
    '-----------------------------------------------------------------------------------------------------------
    Dim strTittle As String, lngPID As Long, strName As String
    Dim lngCount As Long
    
    If GetParent(hwnd) = 0 Then
        '读取 hWnd 的视窗标题
        strTittle = String(80, 0)
        Call GetWindowText(hwnd, strTittle, 80)
        strTittle = Left(strTittle, InStr(strTittle, Chr(0)) - 1)
        If Trim(strTittle) <> "" Then
            Call GetWindowThreadProcessId(hwnd, lngPID)
            If IsWindowVisible(hwnd) Then
                Err = 0: On Error Resume Next
                strName = gcll_His_PId("K" & lngPID)(0)
                If Err = 0 Then
                    lngCount = Val(gcll_His_PId("K" & lngPID)(2)) + 1
                    gcll_His_PId.Remove "K" & lngPID
                    gcll_His_PId.Add Array(strName, lngPID, lngCount), "K" & lngPID
                End If
                Err.Clear: On Error GoTo 0
            End If
        End If
    End If
    EnumWindowsProc = True ' 表示继续列举 hWnd
    Exit Function
End Function

Private Function TerminatePID(ByVal lngPID As Long) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '功能:结束指定的进程
    '入参:
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2008-10-30 11:06:16
    '-----------------------------------------------------------------------------------------------------------
    Dim lngProcess As Long
    TerminatePID = False
    
    Err = 0: On Error GoTo errHand:
    lngProcess = OpenProcess(PROCESS_TERMINATE, 0&, lngPID)
    Call TerminateProcess(lngProcess, 1&)
    
    TerminatePID = True
errHand:
End Function

Private Function zlHISPidToCollect(ByRef cll_His_Pid As Collection) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '功能:获取ZLHIS的进程给相关的集合(gcll_HIS_Pid)
    '入参:
    '出参:cll_His_Pid-将符合HIS.exe的程序，装载该集合中
    '返回:
    '编制:刘兴洪
    '日期:2008-10-30 10:07:38
    '-----------------------------------------------------------------------------------------------------------
    Dim strExeName  As String, lngSnapShot As Long, lngProcess As Long, lngCount  As Long
    Dim strCurExeName As String, lngCurPid As Long
    Dim uProcess   As PROCESSENTRY32
    Dim StrSessionID As String '当前会话ID
    Dim StrHISSessionID As String '其他ZLHIS进程会话ID
    Const TH32CS_SNAPPROCESS = &H2
    
    
    Err = 0: On Error GoTo errHand:
    strCurExeName = "*" & UCase(App.EXEName) & "*"
    
    lngCurPid = GetCurrentProcessId '获取当前应用程序进程
    lngSnapShot = CreateToolhelpSnapshot(TH32CS_SNAPPROCESS, 0&)
    
    StrSessionID = GetCurSessionID(lngCurPid)
    
    
    If lngSnapShot <> 0 Then
        uProcess.lSize = Len(uProcess)
        lngProcess = ProcessFirst(lngSnapShot, uProcess)
        lngCount = 0
        Do While lngProcess
            '不等于当前进程的才处理
            If lngCurPid <> uProcess.lProcessId Then
                strExeName = UCase(Left(uProcess.sExeFile, InStr(1, uProcess.sExeFile, vbNullChar) - 1))
                If strExeName Like strCurExeName Then '"ZLHIS+.EXE"
                    StrHISSessionID = GetCurSessionID(uProcess.lProcessId)
                    '如果当前zlhis+的进程会话ID与启动的会话ID相同,才进行关闭处理
                    If StrSessionID = StrHISSessionID Then
                        cll_His_Pid.Add Array(strExeName, uProcess.lProcessId, 0), "K" & uProcess.lProcessId
                    End If
                End If
            End If
            lngProcess = ProcessNext(lngSnapShot, uProcess)
        Loop
        CloseHandle (lngSnapShot)
    End If
    zlHISPidToCollect = True
    Exit Function
errHand:
End Function

Private Function GetCurSessionID(ByVal lngCurPid As Long) As String
    '-----------------------------------------------------------------------------------------------------------
    '功能:获取当前进程的会话ID
    '入参:当前进程PID
    '出参:
    '返回:会话ID
    '编制:祝庆
    '日期:2012-06-06 10:15:00
    '-----------------------------------------------------------------------------------------------------------
    On Error Resume Next
    Dim WMI, objProcess, colProcessList As Object
    Set WMI = GetObject("WinMgmts:")
    Set colProcessList = WMI.InstancesOf("Win32_Process")
    For Each objProcess In colProcessList
        If objProcess.Handle = lngCurPid Then
            GetCurSessionID = objProcess.SessionId
            Exit Function
        End If
    Next
    GetCurSessionID = "-1"
End Function

