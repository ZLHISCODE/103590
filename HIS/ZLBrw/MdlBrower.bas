Attribute VB_Name = "MdlBrower"
Option Explicit
'MDI必须
Public Type Menu_Type
    功能菜单 As Long
    窗口菜单 As Long
    其它功能菜单 As Long
    分隔菜单 As Long
End Type
Public 菜单基准 As Menu_Type
Public Enum 工具清单
    导航功能清单 = 10
    字典管理工具 = 11
    消息收发工具 = 12
    系统选项设置 = 13
    EXCEL报表工具 = 14
    本地参数管理 = 15
End Enum
'外挂功能
Public gobjPlugIn As Object

Public gobjRelogin As Object                   '重启类对象
Public FrmMainface As Form
Public gcnOracle As ADODB.Connection

Public gstrSysName As String                '系统名称
Public gstrVersion As String                '系统版本
Public gstrAviPath As String                'AVI文件的存放目录
Public gstrUserFlag As String               '当前用户标志(两位表示)，第1位：是否DBA；第2位：系统所有者
Public gstrDbUser As String                 '当前数据库用户
Public glngUserId As Long                   '当前用户id
Public gstrUserCode As String               '当前用户编码
Public gstrUserName As String               '当前用户姓名
Public gstrUserAbbr As String               '当前用户简码
Public glngDeptId As Long                   '当前用户部门id
Public gstrDeptCode As String               '当前用户部门编码
Public gstrDeptName As String               '当前用户部门名称
Public gstrStation As String                '本工作站名称
Public gblnShutDown As Boolean              '在锁屏情况下是否可以退出导航台

Public gstrObj() As String
Public gobjCls() As Object
Public grsMenus As New ADODB.Recordset       '菜单记录集
Public gstrMenuSys As String                '菜单名称
Public gstrCommand As String                '命令行参数 陈东 2010-12-06
Private mlngSysPre As Long                  '上次调用私有同义词检查创建时的系统号

Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
'关闭系统相关的变量及API函数
'----------------------------------------------------------------------------------------------------
Public Type LUID
    UsedPart As Long
    IgnoredForNowHigh32BitPart As Long
End Type

Public Type TOKEN_PRIVILEGES
    PrivilegeCount As Long
    TheLuid As LUID
    Attributes As Long
End Type
Public Declare Function ExitWindowsEx Lib "user32" (ByVal uFlags As Long, ByVal dwReserved As Long) As Long
Public Declare Function GetVersion Lib "kernel32" () As Long
Public Declare Function GetCurrentProcess Lib "kernel32" () As Long 'The GetCurrentProcess function returns a pseudohandle for the current process.
Public Declare Function OpenProcessToken Lib "advapi32" (ByVal ProcessHandle As Long, ByVal DesiredAccess As Long, TokenHandle As Long) As Long 'The OpenProcessToken function opens the access token associated with a process.
Public Declare Function LookupPrivilegeValue Lib "advapi32" Alias "LookupPrivilegeValueA" (ByVal lpSystemName As String, ByVal lpName As String, lpLuid As LUID) As Long 'The LookupPrivilegeValue function retrieves the locally unique identifier (LUID) used on a specified system to locally represent the specified privilege name.
Public Declare Function AdjustTokenPrivileges Lib "advapi32" (ByVal TokenHandle As Long, ByVal DisableAllPrivileges As Long, NewState As TOKEN_PRIVILEGES, ByVal BufferLength As Long, PreviousState As TOKEN_PRIVILEGES, ReturnLength As Long) As Long 'The AdjustTokenPrivileges function enables or disables privileges in the specified access token. Enabling or disabling privileges in an access token requires TOKEN_ADJUST_PRIVILEGES access.
Public Declare Sub SetLastError Lib "kernel32" (ByVal dwErrCode As Long)
Public Declare Function GetLastError Lib "kernel32" () As Long
Public Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Boolean
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
'用于ExitWindowsEx
Private Const mlng关闭计算机及电源 As Long = 8
Public Const EWX_FORCE = 4 '强行关闭程序并注销
Public Const WINDOWS95 = 0
Public Const WINDOWSNT = 1
Private mlngWin32 As Long
Private mbln注销 As Boolean
'下列语句用于检测是否合法调用
Public Declare Function GlobalAddAtom Lib "kernel32" Alias "GlobalAddAtomA" (ByVal lpString As String) As Integer
Public Declare Function GlobalDeleteAtom Lib "kernel32" (ByVal nAtom As Integer) As Integer
'在窗口结构中为指定的窗口设置信息
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Const GWL_EXSTYLE        As Long = (-20)
Private Const GWL_STYLE          As Long = (-16)
Private Const WS_EX_TOOLWINDOW   As Long = &H80
Private Const WS_EX_CONTEXTHELP  As Long = &H400
Private Const WS_MAXIMIZEBOX     As Long = &H10000
Private Const WS_MINIMIZEBOX     As Long = &H20000
Private Const WS_SYSMENU         As Long = &H80000
Private Const WS_THICKFRAME      As Long = &H40000
Private Const WS_CAPTION = &HC00000
'从指定窗口的结构中取得信息
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
'返回值是菜单的句柄。如果给定的窗口没有菜单，则返回NULL。如果窗口是一个子窗口，返回值无定义
Public Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
'运行指定的进程
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'向系统注册一个指定的热键
Public Declare Function RegisterHotKey Lib "user32" (ByVal hwnd As Long, ByVal id As Long, ByVal fsModifiers As Long, ByVal vk As Long) As Long
'取消热键并释放占用的资源
Public Declare Function UnregisterHotKey Lib "user32" (ByVal hwnd As Long, ByVal id As Long) As Long
  '热键标志常数,用来判断当键盘按键被按下时是否命中了我们设定的热键
Public Const WM_HOTKEY = &H312
Public Const MOD_ALT = &H1
Public Const MOD_CONTROL = &H2
Public Const MOD_SHIFT = &H4
Public Const GWL_WNDPROC = (-4)    '窗口函数的地址
Public Const SW_HIDE = 0 '隐藏窗口，激活另一个窗口
Public Const SW_SHOWNORMAL = 1 '激活并显示指定窗口，如果该窗口被最大化或最小化，将还原其原本的大小和位置。
Public Const SW_SHOWMINIMIZED = 2 '激活并最小化指定窗口
Public Const SW_SHOWMAXIMIZED = 3 '激活并最大化指定窗口
Public Const SW_MAXIMIZE = 3 '将指定的窗口最大化
Public Const SW_SHOWNOACTIVATE = 4 '以其最近的大小和位置显示指定窗口，当前窗口保持激活
Public Const SW_SHOW = 5 '以当前位置和大小激活窗口
Public Const SW_MINIMIZE = 6 ' 将指定的窗口最小化
Public Const SW_SHOWMINNOACTIVE = 7 '以最小化方式显示指定窗口，当窗口保持激活
Public Const SW_SHOWNA = 8 '以当前状态显示指定窗口，当前窗口保持激活
Public Const SW_RESTORE = 9 '还原指定的窗口
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Const VK_LBUTTON = &H1
'函数实际是在整个系统的范围内工作的
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Long
'获取当前进程一个唯一的标识符
Public Declare Function GetCurrentProcessId Lib "kernel32" () As Long
'找出某个窗口的创建者(线程或进程)，返回创建者的标志符。
Public Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
'该函数确定给定的窗口句柄是否标识一个已存在的窗口。
'返回值：如果窗口句柄标识了一个已存在的窗口，返回值为非零；如果窗口句柄未标识一个已存在窗口，返回值为零
Public Declare Function isWindow Lib "user32" Alias "IsWindow" (ByVal hwnd As Long) As Long
'返回值：如果指定的窗口及其父窗口具有WS_VISIBLE风格，返回值为非零；如果指定的窗口及其父窗口不具有WS_VISIBLE风格，返回值为零。由于返回值表明了窗口是否具有Ws_VISIBLE风格，因此，即使该窗口被其他窗口遮盖，函数返回值也为非零。
Public Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long
'该函数确定给定窗口是否是最小化(图标化)的窗口。
Public Declare Function IsIconic Lib "user32" (ByVal hwnd As Long) As Long
'该函数确定给定窗口是否是最大化的窗口。
Public Declare Function IsZoomed Lib "user32" (ByVal hwnd As Long) As Long
' 该函数枚举所有屏幕上的顶层窗口，并将窗口句柄传送给应用程序定义的回调函数
Public Declare Function EnumWindows Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
'获取当前进程的活动窗体，若当前进程没有激活，则返回0
Public Declare Function GetActiveWindow Lib "user32" () As Long
'该函数获得一个指定子窗口的父窗口句柄。
Public Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
'该函数获得指定窗口所属的类的类名?
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
'取得一个窗体的标题（caption）文字，或者一个控件的内容（在vb里使用：使用vb窗体或控件的caption或text属性）
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long

'窗口Z序处理
'启动构建一系列新窗口位置的过程（以便同时更新）。该函数会向一个内部结果返回一个句柄，这个结构容纳了与窗口位置有关的信息。随后，该结构会由对DeferWindowPos函数的调用填充。准备好更新所有窗口位置以后，对EndDeferWindowPos的一个调用可同时更新结构内所有窗口的位置
Private Declare Function BeginDeferWindowPos Lib "user32" (ByVal nNumWindows As Long) As Long
'该函数为特定的窗口指定一个新窗口位置，并将其输入由BeginDeferWindowPos创建的结构，以便在EndDeferWindowPos函数执行期间更新
'返回一个新句柄，它指向的结构包含了位置更新信息。这个句柄应在对DeferWindowPos的后续调用以及对EndDeferWindowPos的结束调用中用到。如出错，则返回零值
Private Declare Function DeferWindowPos Lib "user32" (ByVal hWinPosInfo As Long, ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
'同时更新DeferWindowPos调用时指定的所有窗口的位置及状态
Private Declare Function EndDeferWindowPos Lib "user32" (ByVal hWinPosInfo As Long) As Long

Private Type LASTINPUTINFO
    cbSize As Long
    dwTime As Long
End Type
'获取上次输入操作的时间。
Private Declare Function GetLastInputInfo Lib "user32" (plii As LASTINPUTINFO) As Boolean
'返回（retrieve）从操作系统启动所经过（elapsed）的毫秒数
Private Declare Function GetTickCount Lib "kernel32" () As Long

Private Const GW_CHILD = &H5
Public Const GW_OWNER = &H4
'获得一个窗口的句柄，该窗口与某源窗口有特定的关系
Public Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
'该函数用于判断指定的窗口是否允许接受键盘或鼠标输入。
Private Declare Function IsWindowEnabled Lib "user32" (ByVal hwnd As Long) As Long
'强制立即更新窗口，窗口中以前屏蔽的所有区域都会重画（在vb里使用：如vb窗体或控件的任何部分需要更新，可考虑直接使用refresh方法
Private Declare Function UpdateWindow Lib "user32" (ByVal hwnd As Long) As Long
'lpRect:指向一个RECT结构的指针，该结构接收窗口的左上角和右下角的屏幕坐标。
Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, ByRef lpRect As RECT) As Long
Private Const SM_CXSIZE = 30
Private Const SM_CYSIZE = 31
Private Const SM_CYCAPTION = 4
Private Const SM_CXBORDER = 5
Private Const SM_CYBORDER = 6
Private Const SM_CXSMSIZE        As Long = 52
Private Const SM_CYSMSIZE        As Long = 53
Private Const SM_CXFRAME         As Long = 32
Private Const SM_CXSIZEFRAME     As Long = SM_CXFRAME
Private Const SM_CYFRAME         As Long = 33
Private Const SM_CYSIZEFRAME     As Long = SM_CYFRAME
Private Const SM_CXDLGFRAME      As Long = 7
Private Const SM_CXFIXEDFRAME    As Long = SM_CXDLGFRAME
Private Const SM_CYDLGFRAME      As Long = 8
Private Const SM_CYFIXEDFRAME    As Long = SM_CYDLGFRAME

Private Const B_EDGE As Long = 2
Public Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long

Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOZORDER = &H4
Public Const SWP_NOACTIVATE = &H10
Private Const SWP_NOOWNERZORDER = &H200
Private Const SWP_SHOWWINDOW = &H40

Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
Public Const HWND_TOP = 0
Public Const HWND_BOTTOM = 1
'在z序中的位于被置位的窗口前的窗口句柄
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Private mlngPid As Long '当前进程
Private mcllHideFrms As Collection '锁屏后隐藏的窗口
Private mcllHideFrmsEx As Collection '锁屏后隐藏的无标题栏窗口
Public gobjButton As frmButton  '锁屏窗口对象
Public gobjLock As frmLock '解锁界面
Public grecButton As RECT '标题栏按钮的上下左右坐标
Private glngPreHwnd As Long '前一个非按钮的活动窗口
Private gblnPreZoomed As Boolean '前一个锁屏图标定位窗体是否是最大化
Public gblnWin10 As Boolean '是否版本为WIn10
Public gintCurTheme As Integer '主题风格0-经典主题,1-AERO主题，WIn8,WIN 10主题,2-BASIC主题
Public gblnHideBtn As Boolean '锁定按钮是否隐藏状态
Public glngLockTime As Long
Public glngMain As Long 'ThunderRT6Main窗体
Public gblnLock As Boolean '是否处于锁屏状态
'管道获取CMD输出
'用来创建一个新的进程和它的主线程，这个新进程运行指定的可执行文件。如果函数执行成功，返回非零值
Private Declare Function CreateProcess Lib "kernel32" Alias "CreateProcessA" (ByVal lpApplicationName As String, ByVal lpCommandLine As String, lpProcessAttributes As Any, lpThreadAttributes As Any, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, lpEnvironment As Any, ByVal lpCurrentDriectory As String, lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long
'创建一个匿名管道，并从中得到读写管道的句柄。如果函数执行成功，返回非零值
Private Declare Function CreatePipe Lib "kernel32" (phReadPipe As Long, phWritePipe As Long, lpPipeAttributes As SECURITY_ATTRIBUTES, ByVal nSize As Long) As Long
'从文件指针指向的位置开始将数据读出到一个文件中， 且支持同步和异步操作
Private Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, lpOverlapped As Any) As Long
'当等待仍在挂起状态时，句柄被关闭，那么函数行为是未定义的。该句柄必须具有 SYNCHRONIZE 访问权限。
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Type STARTUPINFO
    cb                              As Long
    lpReserved                      As String
    lpDesktop                       As String
    lpTitle                         As String
    dwX                             As Long
    dwY                             As Long
    dwXSize                         As Long
    dwYSize                         As Long
    dwXCountChars                   As Long
    dwYCountChars                   As Long
    dwFillAttribute                 As Long
    dwFlags                         As Long
    wShowWindow                     As Integer
    cbReserved2                     As Integer
    lpReserved2                     As Long
    hStdInput                       As Long
    hStdOutput                      As Long
    hStdError                       As Long
End Type
Private Type PROCESS_INFORMATION
    hProcess                        As Long
    hThread                         As Long
    dwProcessId                     As Long
    dwThreadId                      As Long
End Type
Private Type SECURITY_ATTRIBUTES
    nLength                         As Long
    lpSecurityDescriptor            As Long
    bInheritHandle                  As Long
End Type
Private Const NORMAL_PRIORITY_CLASS  As Long = &H20&
Private Const STARTF_USESTDHANDLES   As Long = &H100&
Private Const STARTF_USESHOWWINDOW   As Long = &H1&
Private Const INFINITE               As Long = &HFFFF&

Public Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
Public Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
Public Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal nCode As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Public Const WH_KEYBOARD As Integer = 2      '普通键盘钩子
Public glngHook As Long                      '键盘消息句柄

Public Sub ExecuteFunc(lngSys As Long, Components As String, Modul As Long, Optional ByVal strPara As String) ', Identity As Byte
    '-------------------------------------------------------------
    '功能：调用执行指定部件的功能程序
    '参数：
    '   frmbrower：主窗体
    '   Components：部件
    '   Modul：模块编号
    '   Identity：可执行者身份要求
    '-------------------------------------------------------------
    Dim rsCheck As New ADODB.Recordset                  '检测版本是否符合系统需求
    Dim IntCount As Integer, intClients As Integer
    Dim objNow As Object                                '创建的部件对象
    Dim BlnExecute As Boolean                           '是否存在该部件
    Dim StrVersion As String, StrCompareVersion As String
    Dim ArrayVersion
    '合法性检查
    Dim intAtom As Integer, strCommon As String
    Dim strSQL  As String
    
    Err = 0: On Error Resume Next
    FrmMainface.MousePointer = 11
    
    IntCount = UBound(gstrObj)
    If Err <> 0 Then IntCount = -1
    Err = 0
    
    BlnExecute = False
    If IntCount >= 0 Then
        For IntCount = 0 To UBound(gstrObj)
            If gstrObj(IntCount) = Components Then
                BlnExecute = True
                Exit For
            End If
        Next
    End If
    
    '使用新病历部件
    If UCase(Components) = UCase("zl9EmrInterface") And BlnExecute = False Then
        IntCount = UBound(gstrObj)
        IntCount = IntCount + 1
        ReDim Preserve gstrObj(IntCount)
        gstrObj(IntCount) = Components
        If FrmMainface.mobjEmr Is Nothing Then
            MsgBox "病历组件创建失败！请检查并重新登录。", vbInformation, gstrSysName
            Exit Sub
        ElseIf FrmMainface.mobjEmr.IsInited = False Then
            MsgBox "病历组件未能初始化," & FrmMainface.mobjEmr.GetError(), vbInformation, gstrSysName
            Exit Sub
        End If
        If Not gobjRelogin.IsEMRProxy Then '使用代理用户登录，则不检查权限
            Dim strSpecify As String '片段，范文权限固定在调用前传递
            If Not FrmMainface.mobjEmr.HasInjectAuthorization(2201) Then
                strSpecify = GetPrivFunc(lngSys, 2201)
                Call FrmMainface.mobjEmr.InjectAuthorization(2201, strSpecify)
            End If
            If Not FrmMainface.mobjEmr.HasInjectAuthorization(2203) Then
                strSpecify = GetPrivFunc(lngSys, 2203)
                Call FrmMainface.mobjEmr.InjectAuthorization(2203, strSpecify)
            End If
        End If
        BlnExecute = True
    End If
    '--如果没有该部件,则创建--
    If BlnExecute = False Then
        Set objNow = CreateObject(Components & ".Cls" & Mid(Components, 4))
    
        If Err = 0 Then
            On Error GoTo errH
            '--检查该部件的版本是否满足系统需求(主版本-3;次版本-3;附版本-3)[自定义报表部件除外]--
            If Not (UCase(Components) = "ZL9REPORT") And Not (UCase(Components) = "ZL9DOC") And Not OS.IsDesinMode Then
                strSQL = " Select nvl(主版本,1) 主版本,nvl(次版本,0) 次版本,nvl(附版本,0) 附版本,名称 " & _
                          " From ZlComponent Where Upper(Rtrim(部件))=[1] And 系统=[2]"
                Set rsCheck = zlDatabase.OpenSQLRecord(strSQL, "部件版本检查", UCase(Components), lngSys)
                With rsCheck
                    If .EOF Then
                        MsgBox "系统表部件表ZlComponent数据不完整，请与软件开发商联系！", vbInformation, gstrSysName
                        FrmMainface.MousePointer = 0
                        Exit Sub
                    End If
                    
                    '组装版本号为三位主版本、三位次版本及三位附版本
                    StrCompareVersion = String(3 - Len(!主版本), "0") & !主版本 & "." & _
                                        String(3 - Len(!次版本), "0") & !次版本 & "." & _
                                        String(3 - Len(!附版本), "0") & !附版本
                    ArrayVersion = Split(objNow.Version, ".")
                    StrVersion = String(3 - Len(ArrayVersion(0)), "0") & ArrayVersion(0) & "." & _
                                 String(3 - Len(ArrayVersion(1)), "0") & ArrayVersion(1) & "." & _
                                 String(3 - Len(ArrayVersion(2)), "0") & ArrayVersion(2)
                    
                    If StrVersion < StrCompareVersion Then
                        MsgBox "该部件的版本已不能满足系统的需求，请与软件开发商联系！（" & !名称 & "）", vbInformation, gstrSysName
                        FrmMainface.MousePointer = 0
                        Exit Sub
                    End If
                End With
            End If
        
            IntCount = 0
            On Error Resume Next
            IntCount = UBound(gstrObj)
            IntCount = IntCount + 1
            If Err.Number <> 0 Then Err.Clear
            On Error GoTo errH
            ReDim Preserve gobjCls(IntCount)
            Set gobjCls(IntCount) = objNow
            ReDim Preserve gstrObj(IntCount)
            gstrObj(IntCount) = Components
        '创建部件失败，应该提示
        Else
            Screen.MousePointer = 0
            MsgBox "部件 " & Components & ".Cls" & Mid(Components, 4) & " 不能正常创建，请检查安装是否正确！信息：" & vbNewLine & Err.Description, vbExclamation, gstrSysName
            Err.Clear
            Exit Sub
        End If
    End If
    
    Err = 0: On Error GoTo errH
    '--执行该功能--
    If gstrObj(IntCount) = Components Then
        If UCase(Components) = "ZL9REPORT" Then
            If Modul = 菜单基准.其它功能菜单 Then
                gobjCls(IntCount).ReportMan gcnOracle, FrmMainface
            Else
                
'                strPara = "开始日期=2013-01-01"
                If strPara <> "" Then
                    Dim varPara As Variant
                                        
                    varPara = Split(strPara, "|")
'                    varPara(0) = "开始日期=2013-01-01"
'                    varPara(1) = "结束日期=2014-05-01"
                    
                    '最多支持10个参数，超过10个的不管
                    Select Case UBound(varPara)
                    Case 0
                        gobjCls(IntCount).ReportOpen gcnOracle, lngSys, Modul, FrmMainface, CStr(varPara(0))
                    Case 1
                        gobjCls(IntCount).ReportOpen gcnOracle, lngSys, Modul, FrmMainface, CStr(varPara(0)), CStr(varPara(1))
                    Case 2
                        gobjCls(IntCount).ReportOpen gcnOracle, lngSys, Modul, FrmMainface, CStr(varPara(0)), CStr(varPara(1)), CStr(varPara(2))
                    Case 3
                        gobjCls(IntCount).ReportOpen gcnOracle, lngSys, Modul, FrmMainface, CStr(varPara(0)), CStr(varPara(1)), CStr(varPara(2)), CStr(varPara(3))
                    Case 4
                        gobjCls(IntCount).ReportOpen gcnOracle, lngSys, Modul, FrmMainface, CStr(varPara(0)), CStr(varPara(1)), CStr(varPara(2)), CStr(varPara(3)), CStr(varPara(4))
                    Case 5
                        gobjCls(IntCount).ReportOpen gcnOracle, lngSys, Modul, FrmMainface, CStr(varPara(0)), CStr(varPara(1)), CStr(varPara(2)), CStr(varPara(3)), CStr(varPara(4)), CStr(varPara(5))
                    Case 6
                        gobjCls(IntCount).ReportOpen gcnOracle, lngSys, Modul, FrmMainface, CStr(varPara(0)), CStr(varPara(1)), CStr(varPara(2)), CStr(varPara(3)), CStr(varPara(4)), CStr(varPara(5)), CStr(varPara(6))
                    Case 7
                        gobjCls(IntCount).ReportOpen gcnOracle, lngSys, Modul, FrmMainface, CStr(varPara(0)), CStr(varPara(1)), CStr(varPara(2)), CStr(varPara(3)), CStr(varPara(4)), CStr(varPara(5)), CStr(varPara(6)), CStr(varPara(7))
                    Case 8
                        gobjCls(IntCount).ReportOpen gcnOracle, lngSys, Modul, FrmMainface, CStr(varPara(0)), CStr(varPara(1)), CStr(varPara(2)), CStr(varPara(3)), CStr(varPara(4)), CStr(varPara(5)), CStr(varPara(6)), CStr(varPara(7)), CStr(varPara(8))
                    Case 9
                        gobjCls(IntCount).ReportOpen gcnOracle, lngSys, Modul, FrmMainface, CStr(varPara(0)), CStr(varPara(1)), CStr(varPara(2)), CStr(varPara(3)), CStr(varPara(4)), CStr(varPara(5)), CStr(varPara(6)), CStr(varPara(7)), CStr(varPara(8)), CStr(varPara(9))
                    Case Else
                        gobjCls(IntCount).ReportOpen gcnOracle, lngSys, Modul, FrmMainface, CStr(varPara(0)), CStr(varPara(1)), CStr(varPara(2)), CStr(varPara(3)), CStr(varPara(4)), CStr(varPara(5)), CStr(varPara(6)), CStr(varPara(7)), CStr(varPara(8)), CStr(varPara(9))
                    End Select
                    
                Else
                    gobjCls(IntCount).ReportOpen gcnOracle, lngSys, Modul, FrmMainface
                End If
                
            End If
        ElseIf UCase(Components) = UCase("zl9EmrInterface") Then
            Dim strFuncs As String, strModul As String
            
            strSQL = " Select 标题　From zlPrograms Where 序号=[1] "
            Set rsCheck = zlDatabase.OpenSQLRecord(strSQL, "系统模块检查", Modul)
            With rsCheck
                    If .EOF Then
                        MsgBox "系统表数据不完整，请与软件开发商联系！", vbInformation, gstrSysName
                        FrmMainface.MousePointer = 0
                        Exit Sub
                    Else
                        strModul = !标题
                    End If
            End With
            strFuncs = GetPrivFunc(lngSys, Modul)
            Call FrmMainface.mobjEmr.CodeMain(Modul, strModul, FrmMainface.hwnd, gobjRelogin.EMRUser, gobjRelogin.EMRPwd, strFuncs)
        Else
            Call CreateSynonyms(lngSys, Modul)
            
            '用户站点数检测 (正式版及试用版)
            intClients = Val(zlRegInfo("授权站点"))
            If intClients > 0 Then
                If GetCurStates > intClients Then
                    MsgBox "当前用户登录数超过了最大授权数" & intClients & ",系统将自动结束运行！", vbInformation, gstrSysName
                    Exit Sub
                End If
            End If

            
            '为通讯原子赋值
            strCommon = Format(Now, "yyyyMMddHHmm")
            strCommon = TranPasswd(strCommon) & "||" & OS.ComputerName
            '加入通讯原子
            intAtom = GlobalAddAtom(strCommon)
            Call SaveSetting("ZLSOFT", "公共全局", "公共", intAtom)
            gobjCls(IntCount).CodeMan lngSys, Modul, gcnOracle, FrmMainface, gstrDbUser
            Call GlobalDeleteAtom(intAtom)
            
            '因医保部件只有CodeMan()才能获取系统号，在读取参数时必须知道系统号，特写入注册表，如果医保读不到默认为 100
            Call SaveSetting("ZLSOFT", "公共全局", "系统号", lngSys)
        End If
    End If
    FrmMainface.MousePointer = 0
    Exit Sub
errH:
    FrmMainface.MousePointer = 0
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Public Sub ReLogin()
    '功能:完成重新重录
    mbln注销 = True
    
    Call gobjRelogin.ReLogin(FrmMainface)
End Sub

Public Function OwnerUser(ByVal strUserName As String) As Boolean
    Dim RecUser As New ADODB.Recordset
    Dim strSQL As String
    OwnerUser = True
    On Error GoTo errH
'        If .State = 1 Then .Close
        strSQL = "Select Count(*) 所有者 From ZlSystems Where 所有者='" & strUserName & "'"
         Set RecUser = zlDatabase.OpenSQLRecord(strSQL, "所有者")
'        .Open "Select Count(*) 所有者 From ZlSystems Where 所有者='" & strUserName & "'", gcnOracle By zq
        
        If RecUser.EOF Then
            If Not IsNull(RecUser!所有者) Then
                If RecUser!所有者 = 0 Then OwnerUser = False
            End If
        End If
'    End With
    Exit Function
errH:
    OwnerUser = False
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function CreateSynonyms(ByVal lngSys As Long, ByVal LngModul As Long)
    Dim strSQL As String
    '创建模块所需对象的同义词(如果已创建则不会再创建)
    On Error Resume Next
    If mlngSysPre <> lngSys Then
        strSQL = "Zl_Createsynonyms(" & lngSys & ")"
        zlDatabase.ExecuteProcedure strSQL, "创建同义词"
        mlngSysPre = lngSys
        If Err.Number <> 0 Then Err.Clear
    End If
End Function

Public Sub AddHistory(ByVal strModul As String)
    Dim str系统 As String, str序号 As String, intMax As Integer
    Dim arr系统 As Variant, arr序号 As Variant, strValue As String
    Dim int系统_Cur As Integer, int序号_Cur As Integer
    Dim int系统_Max As Integer, int序号_Max As Integer
    '最近运行的程序，始终在第一个位置；如果已存在于历史记录中，则将其置于第一个位置
    'strModul:系统 & "," & 模块
    
    intMax = 6
    
    strValue = zlDatabase.GetPara("最近使用模块")
    If UBound(Split(strValue, "|")) >= 1 Then
        str系统 = Trim(Split(strValue, "|")(0))
        str序号 = Trim(Split(strValue, "|")(1))
    End If
    If str系统 = "" Or str序号 = "" Then
        str系统 = Split(strModul, ",")(0)
        str序号 = Split(strModul, ",")(1)
        Call zlDatabase.SetPara("最近使用模块", str系统 & "|" & str序号)
        Exit Sub
    End If
    
    arr系统 = Split(str系统, ",")
    arr序号 = Split(str序号, ",")
    int系统_Max = UBound(arr系统)
    int序号_Max = UBound(arr序号)
    str系统 = Split(strModul, ",")(0): str序号 = Split(strModul, ",")(1)
    If int系统_Max > intMax Then int系统_Max = intMax
    
    For int系统_Cur = 0 To int系统_Max
        int序号_Cur = int系统_Cur
        If int序号_Cur > int序号_Max Then Exit For
        If Not (arr系统(int系统_Cur) = Split(strModul, ",")(0) And arr序号(int序号_Cur) = Split(strModul, ",")(1)) Then
            str系统 = str系统 & "," & arr系统(int系统_Cur)
            str序号 = str序号 & "," & arr序号(int序号_Cur)
        End If
    Next
    Call zlDatabase.SetPara("最近使用模块", str系统 & "|" & str序号)
End Sub

Public Sub CheckWinVersion()
    Dim lngVersion As Long
    
    mbln注销 = False
    lngVersion = GetVersion()
    If ((lngVersion And &H80000000) = 0) Then
        mlngWin32 = WINDOWSNT
    Else
        mlngWin32 = WINDOWS95
    End If
End Sub

Public Sub ShutDown(ByVal blnCloseWin As Boolean)
    If mbln注销 Then Exit Sub
    If Not blnCloseWin Then Exit Sub
    If mlngWin32 = WINDOWSNT Then
        Call AdjustToken
        Call ExitWindowsEx(mlng关闭计算机及电源 Or EWX_FORCE, 0)
    Else
        Call ExitWindowsEx(mlng关闭计算机及电源 Or EWX_FORCE, 0)
    End If
End Sub

Public Function AdjustToken() As Boolean
    Const TOKEN_ADJUST_PRIVILEGES = &H20
    Const TOKEN_QUERY = &H8
    Const SE_PRIVILEGE_ENABLED = &H2
    Dim hdlProcessHandle As Long
    Dim hdlTokenHandle As Long
    Dim tmpLuid As LUID
    Dim tkp As TOKEN_PRIVILEGES
    Dim tkpNewButIgnored As TOKEN_PRIVILEGES
    Dim lBufferNeeded As Long
    
    'Set the error code of the last thread to zero using the'SetLast Error function
    SetLastError 0
    
    '得到当前进程的句柄
    hdlProcessHandle = GetCurrentProcess()
    If GetLastError <> 0 Then Exit Function
    
    '得到当前进程的权限句柄
    OpenProcessToken hdlProcessHandle, (TOKEN_ADJUST_PRIVILEGES Or TOKEN_QUERY), hdlTokenHandle
    If GetLastError <> 0 Then Exit Function
     
    '找到关闭权限并赋给LUID
    'SE_REMOTE_SHUTDOWN_NAME = "SeRemoteShutdownPrivilege
    'SE_SHUTDOWN_NAME = "SeShutdownPrivilege"
    LookupPrivilegeValue "", "SeShutdownPrivilege", tmpLuid
    
    tkp.PrivilegeCount = 1    ' One privilege to set
    tkp.TheLuid = tmpLuid
    tkp.Attributes = SE_PRIVILEGE_ENABLED
    
    'Enable the shutdown privilege in the access token of this process
    AdjustTokenPrivileges hdlTokenHandle, False, tkp, Len(tkpNewButIgnored), tkpNewButIgnored, lBufferNeeded
    If GetLastError <> 0 Then Exit Function
    
    AdjustToken = True
End Function

Public Function TranPasswd(strOld As String) As String
    '------------------------------------------------
    '功能： 密码转换函数
    '参数：
    '   strOld：原密码
    '返回： 加密生成的密码
    '------------------------------------------------
    Dim intDo As Integer
    Dim StrPass As String, strReturn As String, strSource As String, strTarget As String
    
    StrPass = "WriteByZybZL"
    strReturn = ""
    
    For intDo = 1 To 12
        strSource = Mid(strOld, intDo, 1)
        strTarget = Mid(StrPass, intDo, 1)
        strReturn = strReturn & Chr(Asc(strSource) Xor Asc(strTarget))
    Next
    TranPasswd = strReturn
End Function

Public Sub HideForm(ByVal lnghWnd As Long, Optional ByVal blnHide As Boolean = True)
'隐藏指定窗口,隐藏时不再任务栏上展示
    On Error Resume Next
    '恢复前一个有效的按钮窗体的最大化状态
    Call ShowWindow(lnghWnd, IIf(Not blnHide, IIf(gblnPreZoomed And lnghWnd = glngPreHwnd, SW_SHOWMAXIMIZED, SW_SHOW), SW_HIDE))
    If Err.Number <> 0 Then Err.Clear
End Sub

Public Function EnumWindowsProc(ByVal hwnd As Long, ByVal lParam As Long) As Long
'功能： 通过PID枚举所属的句柄,查找需要的窗口
    Dim lngPid As Long
    Dim strText As String * 255
    GetWindowThreadProcessId hwnd, lngPid
    If mlngPid = lngPid Then
        If isWindow(hwnd) <> 0 Then
            If IsWindowVisible(hwnd) <> 0 Then
                If isNormalWindow(hwnd) Then
                    mcllHideFrms.Add hwnd
                Else
                    If (IsWindowEnabled(GetWindow(hwnd, GW_OWNER)) = 0) Then '无标题栏的模态窗体需要隐藏
                        mcllHideFrms.Add hwnd
                    Else
                        mcllHideFrmsEx.Add hwnd
                    End If
                End If
            End If
        End If
    End If
    EnumWindowsProc = True
End Function

Public Sub GetAllVisibleWindow(ByVal lngPid As Long)
    mlngPid = lngPid
    Set mcllHideFrms = New Collection
    Set mcllHideFrmsEx = New Collection
    EnumWindows AddressOf EnumWindowsProc, 0
End Sub

Public Sub LockProg(ByVal blnLock As Boolean)
    Dim varItm As Variant
    Dim lnghWnd As Long
    Dim lngPre As Long
    
    If blnLock Then
        '获取所有的可见窗体
        Call GetAllVisibleWindow(GetCurrentProcessId)
        '获取前一个窗体是否最大化，因为这个窗体初始进入最大化，锁定解锁后会出现屏幕绘图错误。需要先恢复再最大化。
        glngPreHwnd = GetActiveWindow
        gblnPreZoomed = False
        If glngPreHwnd <> frmBrower.hwnd Then
            If isWindow(glngPreHwnd) <> 0 Then
                If GetMenu(glngPreHwnd) <> 0 Then '存在窗口自带菜单，则会有TOOLBar
                    If IsZoomed(glngPreHwnd) <> 0 Then
                        gblnPreZoomed = True
                        Call ShowWindow(glngPreHwnd, SW_RESTORE)
                    End If
                End If
            End If
        End If
    End If
    '讲所有窗口隐藏
    gblnLock = blnLock
    For Each varItm In mcllHideFrms
        If varItm = frmBrower.hwnd Then
        Else
            Call HideForm(varItm, blnLock)
        End If
    Next

    If blnLock Then
        Set gobjLock = New frmLock
        gobjLock.Show vbModal, frmBrower
    Else
        Set gobjLock = Nothing
    End If
End Sub

Private Function isNormalWindow(ByVal lnghWnd As Long) As Boolean
'排除特殊控件的窗体干扰
'排除DTPicker弹出的日期选择界面，该界面通过API判断是有标题栏的，通过SPY++跟踪是没有的，暂时排除该窗口
    Dim strText As String * 256
    Dim strTmp As String
    On Error Resume Next
    If GetWindowLong(lnghWnd, GWL_STYLE) And WS_CAPTION Then
        Call GetWindowText(lnghWnd, strText, 255)
        strTmp = zlStr.TruncZero(strText)
        isNormalWindow = strTmp <> ""
    Else
        isNormalWindow = False
    End If
End Function


Private Function GetButtonRect(ByVal lnghWnd As Long) As RECT
'功能：计算锁频按钮位置
    Dim lngcxBut   As Long, lngcyBut   As Long
    Dim uRect    As RECT
    Dim lngButSize    As Long, lngSysButSize As Long
    Dim lngRightEdgeOffset As Long
    Dim lngStyle      As Long, lngExStyle   As Long
    
    '获取窗体样式
    lngStyle = GetWindowLong(lnghWnd, GWL_STYLE)
    lngExStyle = GetWindowLong(lnghWnd, GWL_EXSTYLE)
    '获取右边按钮位置偏移，即右边的最大化，最小化，关闭按钮
    If (lngExStyle And WS_EX_TOOLWINDOW) Then
        lngSysButSize = GetSystemMetrics(SM_CXSMSIZE) - B_EDGE
        If (lngStyle And WS_SYSMENU) Then
            lngButSize = lngSysButSize + B_EDGE
        End If
        If (lngStyle And WS_THICKFRAME) Then
            lngRightEdgeOffset = lngButSize + GetSystemMetrics(SM_CXSIZEFRAME)
        Else
            lngRightEdgeOffset = lngButSize + GetSystemMetrics(SM_CXFIXEDFRAME)
        End If
    Else
        lngSysButSize = GetSystemMetrics(SM_CXSIZE) - B_EDGE
        '系统菜单按钮
        If (lngStyle And WS_SYSMENU) Then
            lngButSize = lngButSize + lngSysButSize + B_EDGE
        End If
        '最大化最小化按钮
        If (lngStyle And (WS_MINIMIZEBOX Or WS_MAXIMIZEBOX)) Then
            lngButSize = lngButSize + B_EDGE + lngSysButSize * 2
        ElseIf (lngExStyle And WS_EX_CONTEXTHELP) Then '帮助按钮
            lngButSize = lngButSize + B_EDGE + lngSysButSize
        End If
        If (lngStyle And WS_THICKFRAME) Then
            lngRightEdgeOffset = lngButSize + GetSystemMetrics(SM_CXSIZEFRAME)
        Else
            lngRightEdgeOffset = lngButSize + GetSystemMetrics(SM_CXFIXEDFRAME)
        End If
    End If
    '获取按钮大小
    If (lngExStyle And WS_EX_TOOLWINDOW) Then
        lngcxBut = GetSystemMetrics(SM_CXSMSIZE)
        lngcyBut = GetSystemMetrics(SM_CYSMSIZE)
      Else
        lngcxBut = GetSystemMetrics(SM_CXSIZE)
        lngcyBut = GetSystemMetrics(SM_CYSIZE)
    End If
    '获取窗体原点位置
    Call GetWindowRect(lnghWnd, uRect)
    With uRect
        If gintCurTheme <> 1 Then
            'Win10,位置和最小化窗口按钮重合，所以减去一个按钮宽度
            .Right = .Right - lngRightEdgeOffset - IIf(gblnWin10, lngcxBut, 0) - B_EDGE * IIf(gintCurTheme = 0, 1, 0)
            If (lngStyle And WS_THICKFRAME) Then
                .Top = .Top + GetSystemMetrics(SM_CYSIZEFRAME)
              Else
                .Top = .Top + GetSystemMetrics(SM_CYFIXEDFRAME)
            End If
            .Top = .Top + (lngcyBut - 16) / 2
        Else
            If IsZoomed(lnghWnd) Then
                'Win10,位置和最小化窗口按钮重合，所以减去一个按钮宽度
                .Right = .Right - lngRightEdgeOffset - IIf(gblnWin10, lngcxBut, 0) - B_EDGE * 4
                If (lngStyle And WS_THICKFRAME) Then
                    .Top = .Top + GetSystemMetrics(SM_CYSIZEFRAME) + B_EDGE
                  Else
                    .Top = .Top + GetSystemMetrics(SM_CYFIXEDFRAME) + B_EDGE
                End If
            Else
                .Right = .Right - lngRightEdgeOffset - IIf(gblnWin10, lngcxBut, 0) - B_EDGE
                .Top = .Top + B_EDGE
            End If
        End If
        .Left = .Right - 16
        .Bottom = .Top + 16
    End With
    GetButtonRect = uRect
End Function

Public Function GetCurTheme() As Integer
'功能：获取当前主题,0-WIndows 经典,1-win7 AERO,win8,win10 ,2-WIn7 BASIC ,
    Dim lngValue As Long, strValue As String
    Dim intCurTheme As Integer
    '获取当前主题
    If OS.GetRegValue("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\ThemeManager", "ThemeActive", strValue) Then
        intCurTheme = Val(strValue)
    Else
        intCurTheme = 0
    End If
    '当是AERO效果时 ，获取DWM组件状况，启用则为AERO
    If intCurTheme = 1 Then
        If OS.GetRegValue("HKEY_CURRENT_USER\Software\Microsoft\Windows\DWM", "Composition", lngValue) Then
            intCurTheme = IIf(lngValue = 0, 2, 1)
        Else
            intCurTheme = 2
        End If
    End If
    GetCurTheme = intCurTheme
End Function

Private Function GetCmdVer()
    Dim strLine As String
    Dim arrTmp As Variant
    strLine = RunCommand("cmd /c ""Ver " & Chr(13) & """")
    
    strLine = Trim(Split(strLine & "]", "]")(0))
    arrTmp = Split(" " & strLine, " ")
    GetCmdVer = arrTmp(UBound(arrTmp))
End Function

Public Function RunCommand(commandline As String) As String
    Dim si As STARTUPINFO                                                       'used to send info the CreateProcess
    Dim pi As PROCESS_INFORMATION                                               'used to receive info about the created process
    Dim retval As Long                                                          'return value
    Dim hRead As Long                                                           'the handle to the read end of the pipe
    Dim hWrite As Long                                                          'the handle to the write end of the pipe
    Dim sBuffer(0 To 63) As Byte                                                'the buffer to store data as we read it from the pipe
    Dim lgSize As Long                                                          'returned number of bytes read by readfile
    Dim sa As SECURITY_ATTRIBUTES
    Dim strResult As String                                                     'returned results of the command line
    
    'set up security attributes structure
    With sa
        .nLength = Len(sa)
        .bInheritHandle = 1&                                                    'inherit, needed for this to work
        .lpSecurityDescriptor = 0&
    End With
    'create our anonymous pipe an check for success
    ' note we use the default buffer size
    ' this could cause problems if the process tries to write more than this buffer size
    retval = CreatePipe(hRead, hWrite, sa, 0&)
    If retval = 0 Then
'        MsgBox "错误提示:创建管道失败!"
        RunCommand = ""
        Exit Function
    End If
    'set up startup info
    With si
        .cb = Len(si)
        .dwFlags = STARTF_USESTDHANDLES Or STARTF_USESHOWWINDOW                 'tell it to use (not ignore) the values below
        .wShowWindow = SW_HIDE
        .hStdOutput = hWrite                                                    'pass the write end of the pipe as the processes standard output
    End With
    'run the command line and check for success
    retval = CreateProcess(vbNullString, commandline & vbNullChar, sa, sa, 1&, NORMAL_PRIORITY_CLASS, ByVal 0&, vbNullString, si, pi)
    If retval Then
        'wait until the command line finishes
        ' trouble if the app doesn't end, or waits for user input, etc
        WaitForSingleObject pi.hProcess, INFINITE
        'read from the pipe until there's no more (bytes actually read is less than what we told it to)
        Do While ReadFile(hRead, sBuffer(0), 64, lgSize, ByVal 0&)
            'convert byte array to string and append to our result
            strResult = strResult & StrConv(sBuffer(), vbUnicode)
            'TODO = what's in the tail end of the byte array when lgSize is less than 64???
            Erase sBuffer()
            If lgSize <> 64 Then Exit Do
            DoEvents
        Loop
        'close the handles of the process
        CloseHandle pi.hProcess
        CloseHandle pi.hThread
    Else
'        MsgBox "错误提示:创建进程失败!" & vbCrLf
        RunCommand = ""
        Exit Function
    End If
    'close pipe handles
    CloseHandle hRead
    CloseHandle hWrite
    'return the command line output
    RunCommand = Replace(strResult, vbNullChar, "")
End Function

Public Function TimeToLock() As Boolean
'是否该锁定系统了
    Dim lii As LASTINPUTINFO
    lii.cbSize = Len(lii)
    GetLastInputInfo lii
    If GetTickCount - lii.dwTime > glngLockTime Then
        TimeToLock = True
    End If
End Function

'用于监控本进程中的键盘消息，
Public Function MyKBHook(ByVal nCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    If nCode >= 0 Then
        MyKBHook = 0 '表示要处理这个消息
        If wParam = vbKeyL And (GetKeyState(vbKeyMenu) And &HFF80) And (GetKeyState(vbKeyControl) And &HFF80) Then
            If gblnLock = False Then
                Call LockProg(True)
                MyKBHook = 1
            End If
        End If
    End If
    Call CallNextHookEx(glngHook, nCode, wParam, lParam) '将消息传给下一个钩子
End Function


