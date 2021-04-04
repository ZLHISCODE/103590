Attribute VB_Name = "mdlPublic"
Option Explicit

Public gcnOracle As ADODB.Connection
Public gobjComlib As Object
Public gobjCommFun As Object
Public gobjControl As Object
Public gobjDatabase As Object
Public gobjGrid As Object

Public gstrDBUser As String
Public gstrNodeNo As String          '当前站点编号；如果未设置启用站点，则为"-"
Public gstrSysName As String
Public gblnAlone As Boolean '是否独立调用

Public gstrIntergrateIP As String    '移动护理ID地址
Public gobjScriptControl  As MSScriptControl.ScriptControl
Public gstrRelatedUserID As String  '整体护理人员ID
Public gstrRelatedUnitID As String  '整体护理病区ID
Public gstrRelatedPatientID As String  '整体护理病人ID

Public glngPid As Long '当前程序的进程ID
Public gcllHideFrmsEx As Collection '所有窗体集合

Public Type TYPE_INTERGRATE_USER_INFO  '整体护理中人员信息(接口UserLogin产生)
    id As String
    UserName As String
    Name As String
    Sex As String
    Cookie As String
End Type
Public IntergrateUserInfo As TYPE_INTERGRATE_USER_INFO

Public Type TYPE_USER_INFO
    id As Long
    编号 As String
    姓名 As String
    简码 As String
    用户名 As String
    性质 As String
    部门ID As Long
    部门码 As String
    部门名 As String
    专业技术职务 As String
    用药级别 As Long
End Type
Public UserInfo As TYPE_USER_INFO

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public glnglpPrevWndProc As Long
Public glngSCMIZE As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Const GWL_WNDPROC = -4&
Public Const SC_MAXIMIZE = &HF030& '最大化
Public Const SC_MINIMIZE = &HF020& '最小化
Public Const SC_RESTORE = &HF120& '还原

Public Const GW_CHILD = &H5
Public Const GW_OWNER = &H4
Public Const GW_HWNDNEXT = 2
'获得一个窗口的句柄，该窗口与某源窗口有特定的关系
Public Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
'找出某个窗口的创建者(线程或进程)，返回创建者的标志符。
Public Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
'该函数确定给定的窗口句柄是否标识一个已存在的窗口。
'返回值：如果窗口句柄标识了一个已存在的窗口，返回值为非零；如果窗口句柄未标识一个已存在窗口，返回值为零
Public Declare Function isWindow Lib "user32" Alias "IsWindow" (ByVal hwnd As Long) As Long
'返回值：如果指定的窗口及其父窗口具有WS_VISIBLE风格，返回值为非零；如果指定的窗口及其父窗口不具有WS_VISIBLE风格，返回值为零。由于返回值表明了窗口是否具有Ws_VISIBLE风格，因此，即使该窗口被其他窗口遮盖，函数返回值也为非零。
Public Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long
'该函数用于判断指定的窗口是否允许接受键盘或鼠标输入?
Public Declare Function IsWindowEnabled Lib "user32" (ByVal hwnd As Long) As Long
'从指定窗口的结构中取得信息
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
'取得一个窗体的标题（caption）文字，或者一个控件的内容（在vb里使用：使用vb窗体或控件的caption或text属性）
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
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
'获取当前进程一个唯一的标识符
Public Declare Function GetCurrentProcessId Lib "kernel32" () As Long
' 该函数枚举所有屏幕上的顶层窗口，并将窗口句柄传送给应用程序定义的回调函数
Public Declare Function EnumWindows Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long

'是否是64位进程（Is64bit）
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Public Declare Function GetCurrentProcess Lib "kernel32" () As Long
Public Declare Function SetProcessWorkingSetSize Lib "kernel32" (ByVal hProcess As Long, ByVal dwMinimumWorkingSetSize As Long, ByVal dwMaximumWorkingSetSize As Long) As Long
Private Declare Function IsWow64Process Lib "kernel32" (ByVal hProc As Long, bWow64Process As Boolean) As Long

'注册表操作**********************************
Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_CONFIG = &H80000005
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_DYN_DATA = &H80000006
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_USERS = &H80000003
Public Const REG_OPTION_NON_VOLATILE = 0
Public Const KEY_ALL_ACCESS = (&H20000 Or &H1 Or &H2 Or &H4 Or &H8 Or &H10 Or &H20) And (Not &H100000)
Public Const REG_SZ = 1
Public Const REG_DWORD = 4
'*****************************************************************
'*****下面声明注册表操作中用到的API函数****************************
Public Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Public Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal uloptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Public Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Public Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Public Declare Function RegSetValueExLong Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpValue As Long, ByVal cbData As Long) As Long
Public Declare Function RegSetValueExString Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpValue As String, ByVal cbData As Long) As Long
Public Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Public Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long         ' Note that if you declare the lpData parameter as String, you must pass it By Value.

'程序调用进程名称获取
Public gstrExeName As String '执行程序的名称
Private Type PROCESSENTRY32
    dwSize As Long
    cntUsage As Long
    th32ProcessID As Long
    th32DefaultHeapID As Long
    th32ModuleID As Long
    cntThreads As Long
    th32ParentProcessID As Long
    pcPriClassBase As Long
    dwFlags As Long
    szExeFile As String * 1024
End Type
Const TH32CS_SNAPHEAPLIST = &H1
Const TH32CS_SNAPPROCESS = &H2
Const TH32CS_SNAPTHREAD = &H4
Const TH32CS_SNAPMODULE = &H8
Const TH32CS_SNAPALL = (TH32CS_SNAPHEAPLIST Or TH32CS_SNAPPROCESS Or TH32CS_SNAPTHREAD Or TH32CS_SNAPMODULE)
Const TH32CS_INHERIT = &H80000000
Private Declare Function CreateToolhelp32Snapshot Lib "kernel32" (ByVal dwFlags As Long, ByVal th32ProcessID As Long) As Long
Private Declare Function Process32First Lib "kernel32" (ByVal hSnapshot As Long, lppe As PROCESSENTRY32) As Long
Private Declare Function Process32Next Lib "kernel32" (ByVal hSnapshot As Long, lppe As PROCESSENTRY32) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long

Public Function WindowProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
     '如果最大化(SC_MAXIMIZE)则屏蔽 (根据自己需要更改成最小化或还原等…)
    If wParam = glngSCMIZE Then Exit Function
    WindowProc = CallWindowProc(glnglpPrevWndProc, hwnd, uMsg, wParam, lParam)
End Function

Public Function zlGetComLib() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取公共部件相关对象
    '返回:获取成功,返回true,否则返回False
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If Not gobjComlib Is Nothing Then zlGetComLib = True: Exit Function
    
    
    Err = 0: On Error Resume Next
    
    Set gobjComlib = GetObject("", "zl9Comlib.clsComlib")
    Set gobjCommFun = GetObject("", "zl9Comlib.clsCommfun")
    Set gobjControl = GetObject("", "zl9Comlib.clsControl")
    Set gobjDatabase = GetObject("", "zl9Comlib.clsDatabase")
    Set gobjGrid = GetObject("", "zl9Comlib.clsGrid")
    gstrNodeNo = ""
    If Not gobjComlib Is Nothing Then gstrNodeNo = gobjComlib.gstrNodeNo
    Err = 0: On Error GoTo 0
    If Not gobjComlib Is Nothing Then
        If gblnAlone = True Then Call gobjComlib.InitCommon(gcnOracle)
        zlGetComLib = True: Exit Function
    End If
    Err = 0: On Error Resume Next
    Set gobjComlib = CreateObject("zl9Comlib.clsComlib")
    If gblnAlone = True Then
        Call gobjComlib.InitCommon(gcnOracle)
    End If
    Set gobjCommFun = gobjComlib.zlCommFun
    Set gobjControl = gobjComlib.zlControl
    Set gobjDatabase = gobjComlib.zlDatabase
    
    If Not gobjComlib Is Nothing Then
        zlGetComLib = True
        gstrNodeNo = gobjComlib.gstrNodeNo
    End If
    Err = 0: On Error GoTo 0
End Function

Public Function Get人员性质(Optional ByVal str姓名 As String) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:读取当前登录人员或指定人员的人员性质
    '返回:返回人员性质,多个用逗号分离
    '---------------------------------------------------------------------------------------------------------------------------------------------
 
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String
    
    On Error GoTo errHandle
    If str姓名 <> "" Then
        strSql = "Select B.人员性质 From 人员表 A,人员性质说明 B Where A.ID=B.人员ID And A.姓名=[1]"
        Set rsTmp = gobjDatabase.OpenSQLRecord(strSql, "获取人员性质", str姓名)
    Else
        strSql = "Select 人员性质 From 人员性质说明 Where 人员ID = [1]"
        Set rsTmp = gobjDatabase.OpenSQLRecord(strSql, "获取人员性质", UserInfo.id)
    End If
    Do While Not rsTmp.EOF
        Get人员性质 = Get人员性质 & "," & rsTmp!人员性质
        rsTmp.MoveNext
    Loop
    Get人员性质 = Mid(Get人员性质, 2)
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Public Function GetUserInfo() As Boolean
'功能：获取登陆用户信息
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo ErrH
    Set rsTmp = gobjDatabase.GetUserInfo
    If Not rsTmp Is Nothing Then
        If Not rsTmp.EOF Then
            UserInfo.id = rsTmp!id
            UserInfo.用户名 = rsTmp!User
            UserInfo.编号 = rsTmp!编号
            UserInfo.简码 = "" & rsTmp!简码
            UserInfo.姓名 = "" & rsTmp!姓名
            UserInfo.部门ID = Val("" & rsTmp!部门ID)
            UserInfo.部门码 = "" & rsTmp!部门码
            UserInfo.部门名 = "" & rsTmp!部门名
            UserInfo.性质 = Get人员性质
            UserInfo.专业技术职务 = "" & rsTmp!专业技术职务
            GetUserInfo = True
        End If
    End If
    gstrDBUser = UserInfo.用户名
    Exit Function
ErrH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function

Public Function SubB(ByVal strInfor As String, ByVal lngStart As Long, ByVal lngLen As Long) As String
'功能:读取指定字串的值,字串中可以包含汉字
 '入参:strInfor-原串
 '         lngStart-直始位置
'         lngLen-长度
'返回:子串
    Dim strTmp As String, i As Long
    Err = 0: On Error GoTo ErrH:
    SubB = StrConv(MidB(StrConv(strInfor, vbFromUnicode), lngStart, lngLen), vbUnicode)
    SubB = Replace(SubB, Chr(0), "")
    Exit Function
ErrH:
    Err.Clear
    SubB = ""
End Function

'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function SendPostUrl(ByVal strUrl As String, ByVal strParams As String, objHttp As XMLHTTP, strErrMsg As String, Optional blnCookie As Boolean = False, Optional ByVal strIP As String = "") As Boolean
'功能：根据URL地址及参数，连接URL
'参数：
'       strUrl-URL地址；strParams 参数（JSON格式）
'       objHttp：XMLHTTP对象
    Dim oXmlHttp As XMLHTTP
    Dim intPos As Integer  '连接次数计数变量
    If strIP = "" Then strIP = gstrIntergrateIP
    If blnCookie = True Then
'        UserLogin
    End If
    Set oXmlHttp = New MSXML2.ServerXMLHTTP
    oXmlHttp.open "POST", strUrl, False   '初始化HTTP请求
    oXmlHttp.setRequestHeader "Content-Type", "application/json; charset=utf-8"
    If blnCookie = True Then
        oXmlHttp.setRequestHeader "Cookie", IntergrateUserInfo.Cookie
    End If
    oXmlHttp.setTimeouts 5000, 10000, 10000, 10000
    '第一个数值: 解析DNS名字的超时时间
    '第二个数值: 建立Winsock连接的超时时间
    '第三个数值: 发送数据的超时时间
    '第四个数值: 接收response的超时时间
    
    On Error Resume Next
RestartSend:  '重新连接，由于连接超时问题再一次进行连接（最多连接两次）
    Err.Clear
    Call oXmlHttp.send(strParams)
    If Err <> 0 Then
        If Err.Number = -2147012894 Then  'IP地址不存在或连接超时
            If intPos > 0 Then
                strErrMsg = "连接整体护理服务器失败，可能是以下原因导致：" & vbCrLf & _
                    "1、可能设置的IP地址(" & strIP & ")无法连接，请联系管理员重新进行设置" & vbCrLf & _
                    "2、可能由于网络问题连接服务器超时，请重新刷新或再次操作" & vbCrLf & "详细信息：" & Err.Description
            Else
                intPos = intPos + 1
                GoTo RestartSend
            End If
        Else 'IP地址正确，但是不是服务器
            strErrMsg = Err.Description & "，请检查设置的IP地址是否是移动服务器地址！" & vbCrLf & "IP地址：" & strIP
        End If
        Err.Clear
        Exit Function
    End If
    
    'ReadyState
    'HTTP 请求的状态.当一个 XMLHttpRequest 初次创建时，这个属性的值从 0 开始，直到接收到完整的 HTTP 响应，这个值增加到 4。
    '0   Uninitialized   初始化状态。XMLHttpRequest 对象已创建或已被 abort() 方法重置。
    '1   Open    open() 方法已调用，但是 send() 方法未调用。请求还没有被发送。
    '2   Sent    Send() 方法已调用，HTTP 请求已发送到 Web 服务器。未接收到响应。
    '3   Receiving 所有响应头部都已经接收到?响应体开始接收但未完成?
    '4   Loaded  HTTP 响应已经完全接收。
    
    If oXmlHttp.readyState = 4 Then '数据接收成功
        If oXmlHttp.Status = "200" Then
            Set objHttp = oXmlHttp
        Else
            If oXmlHttp.Status = "404" Then 'IP地址正确，当时后面的地址不正确
                strErrMsg = "Http地址不正确，请联系软件开发商！" & vbCrLf & "Http地址：" & strUrl
            Else
                strErrMsg = oXmlHttp.statusText & vbCrLf & "Http地址：" & strUrl
            End If
            Exit Function
        End If
    Else
        strErrMsg = "HTTP 响应数据未能正常接收，请检查服务器或网络情况！"
        Exit Function
    End If
    
    SendPostUrl = True
End Function


Public Function encodeURI(ByVal strValue As String) As String
'功能：字符被转换成 UTF-8 编码（URL特殊字符转换）
    Dim strRetrun As String
    If gobjScriptControl Is Nothing Then Set gobjScriptControl = New MSScriptControl.ScriptControl
    gobjScriptControl.Language = "javascript"
    strRetrun = gobjScriptControl.Eval("encodeURI('" & strValue & "')")
    encodeURI = strRetrun
End Function

Public Function decodeURI(ByVal strValue As String) As String
'功能：将 UTF-8 编码转换为URL特殊字符（URL特殊字符转换）
    Dim strRetrun As String
    If gobjScriptControl Is Nothing Then Set gobjScriptControl = New MSScriptControl.ScriptControl
    gobjScriptControl.Language = "javascript"
    strRetrun = gobjScriptControl.Eval("decodeURI('" & strValue & "')")
    decodeURI = strRetrun
End Function

Public Function encodeURIComponent(ByVal strValue As String) As String
'功能：字符被转换成 UTF-8 编码（URL特殊字符转换）
    Dim strRetrun As String
    If gobjScriptControl Is Nothing Then Set gobjScriptControl = New MSScriptControl.ScriptControl
    gobjScriptControl.Language = "javascript"
    strRetrun = gobjScriptControl.Eval("encodeURIComponent('" & strValue & "')")
    encodeURIComponent = strRetrun
End Function

Public Function decodeURIComponent(ByVal strValue As String) As String
'功能：字符被转换成 UTF-8 编码（URL特殊字符转换）
    Dim strRetrun As String
    If gobjScriptControl Is Nothing Then Set gobjScriptControl = New MSScriptControl.ScriptControl
    gobjScriptControl.Language = "javascript"
    strRetrun = gobjScriptControl.Eval("decodeURIComponent('" & strValue & "')")
    decodeURIComponent = strRetrun
End Function

Public Function AnalysisJavaScriptEvent(ByVal strParam As String, objPopup As clsPopup) As Boolean
'功能：
'strParam格式：
'{
'  type: "CloseDialog" || "ShowDialog", // CloseDialog 关闭弹窗  ShowDialog 打开弹窗
'  moduleUrl: "/shiftReport", //功能Url
'  title: "交班报告",
'  width: "100" || null,
'  height: "100" || null,
'  minimal: true,  //最大化
'  max: false,     //最小化
'  isRefresh: true  //是否刷新父窗体
'  data: "xxxxxxxxxxxx"  //打开弹窗是需要带上的参数
'}
    If gobjScriptControl Is Nothing Then Set gobjScriptControl = New MSScriptControl.ScriptControl

    On Error GoTo ErrHand
    With gobjScriptControl
        .Language = "javascript"
        .AddCode "var json = " & strParam & ";"
        objPopup.PopupParams = strParam
        objPopup.PopupType = "" & .Eval("json.type")
        objPopup.PopupModuleUrl = "" & .Eval("json.moduleUrl")
        objPopup.PopupTitle = "" & .Eval("json.title")
        objPopup.PopupWidth = Val("" & .Eval("json.width"))
        objPopup.PopupHeight = Val("" & .Eval("json.height"))
        objPopup.PopupMinimal = .Eval("json.minimal") 'IIf(UCase("" & .Eval("json.minimal")) = "TRUE", True, False)
        objPopup.PopupMax = .Eval("json.max") 'IIf(UCase("" & .Eval("json.max")) = "TRUE", True, False)
        objPopup.PopupIsRefresh = .Eval("json.isRefresh") ' IIf(UCase("" & .Eval("json.isRefresh")) = "TRUE", True, False)
        objPopup.PopupData = "" & encodeURIComponent(.Eval("json.data"))
        objPopup.PopupParentUrl = decodeURIComponent("" & .Eval("json.ParentUrl"))
        objPopup.PopupParentParam = "" & encodeURIComponent(.Eval("json.originParams"))
        objPopup.PopupPatientID = "" & .Eval("json.PatientID")
        objPopup.PopupUnitID = "" & .Eval("json.LessionID")
        objPopup.PopupUserID = "" & .Eval("json.UserID")
    End With
    AnalysisJavaScriptEvent = True
    Exit Function
ErrHand:
    MsgBox Err.Description & vbCrLf & "Json：" & strParam, vbInformation, gstrSysName
End Function

Public Sub WriteBusinessLOG(ByVal strFunc As String, ByVal strInput As String, ByVal strParams As String, Optional ByVal strOutput As String = "")
    '功能：记录日志文件，主要用于接口调试
    '以下用于记录调用接口的入参
    Dim strDate As String
    Dim strFileName As String
    Dim objStream As TextStream
    Dim objFileSystem As New FileSystemObject
    Dim strLog As String, blnBeginWrite As Boolean
    
    '说明：该部件为整体护理接口专用接口，通用性强，故输出日志线调用公共方法，如果用于低版本则自行输出日志
    strLog = "URL：" & strInput & Space(6) & "入参：" & strParams & Space(6) & IIf(strOutput <> "", "出参：" & strOutput, "")
    On Error Resume Next
    Call gobjComlib.LogWrite("新版护士工作站调用移动整体护理接口跟踪日志", "新版护士工作站", strFunc, strLog)
    If Err <> 0 Then
        blnBeginWrite = True
        Err.Clear
    End If
    If blnBeginWrite Then
        If Not objFileSystem.FolderExists("C:\整体护理日志") Then Call objFileSystem.CreateFolder("C:\整体护理日志")
        strFileName = "C:\整体护理日志\" & Format(Date, "yyyyMMdd") & ".LOG"
        If Not objFileSystem.FileExists(strFileName) Then Call objFileSystem.CreateTextFile(strFileName)
        Set objStream = objFileSystem.OpenTextFile(strFileName, ForAppending)
        strDate = Format(gobjDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
        objStream.WriteLine (String(50, "-"))
        objStream.WriteLine ("  名称:" & strFunc)
        objStream.WriteLine ("  URL:" & strInput)
        objStream.WriteLine ("  入参:" & strParams)
        objStream.WriteLine ("  出参:" & strOutput)
        objStream.WriteLine (String(50, "-"))
        objStream.Close
        Set objStream = Nothing
        If Err <> 0 Then Err.Clear
    End If
    On Error GoTo 0
End Sub

Public Function EnumWindowsProc(ByVal hwnd As Long, ByVal lParam As Long) As Long
'功能： 通过PID枚举所属的句柄,查找需要的窗口（网页控件不能用，只要刷新后就获取不到这个窗体了）
    Dim lngPid As Long
    Dim strText As String * 255
    '网页界面每次刷新一个页面都是一个新进程
    GetWindowThreadProcessId hwnd, lngPid
    If glngPid = lngPid Then
        If isWindow(hwnd) <> 0 Then
            If IsWindowVisible(hwnd) <> 0 Then
                If isNormalWindow(hwnd) Then
                Else
                    gcllHideFrmsEx.Add hwnd
                End If
            End If
        End If
    End If
    EnumWindowsProc = True
End Function

Public Function isNormalWindow(ByVal lngHwnd As Long) As Boolean
'排除特殊控件的窗体干扰
'排除DTPicker弹出的日期选择界面，该界面通过API判断是有标题栏的，通过SPY++跟踪是没有的，暂时排除该窗口
    Dim strText As String * 256
    Dim strTmp As String
    On Error Resume Next
    If GetWindowLong(lngHwnd, GWL_STYLE) And WS_CAPTION Then
        Call GetWindowText(lngHwnd, strText, 255)
        strTmp = gobjComlib.zlStr.TruncZero(strText)
        isNormalWindow = strTmp <> ""
    Else
        isNormalWindow = False
    End If
End Function

Public Sub GetAllVisibleWindow(ByVal lngPid As Long)
    glngPid = lngPid
    Set gcllHideFrmsEx = New Collection
    EnumWindows AddressOf EnumWindowsProc, 0
End Sub

 Public Function Is64bit() As Boolean
    '******************************************************************************************************************
    '功能：是否是64位系统
    '返回：
    '******************************************************************************************************************
    Dim handle As Long
    Dim bolFunc As Boolean
        
    bolFunc = False
    handle = GetProcAddress(GetModuleHandle("kernel32"), "IsWow64Process")
    If handle > 0 Then
        IsWow64Process GetCurrentProcess(), bolFunc
    End If
    Is64bit = bolFunc
End Function

'-------------------------------------------------------------------------------------------------------------------------------------------------------------
'注册表操作
'-------------------------------------------------------------------------------------------------------------------------------------------------------------
'*****新键注册表项
Public Sub Createnewkey(ip As Long, snewkeyname As String)
    Dim hnewkey As Long
    Dim retval As Long
    retval = RegCreateKey(ip, snewkeyname, hnewkey)
    If retval = 0 Then
        RegCloseKey (hnewkey) '关闭上面建立或打开的项
    End If
End Sub
'实例：在HKEY_CURRENT_USER下建立项"xiaopeng"
'代码为 createnewkey HKEY_CURRENT_USER ,"xiaopeng"
'******************************************************************
'*******删除注册表项***********************************************
Public Function Deletekey(ip As Long, skeyname As String)
    Dim hKey As Long
    Dim retval As Long
    retval = RegOpenKeyEx(ip, skeyname, 0, KEY_ALL_ACCESS, hKey)
    If retval = 0 Then
        RegDeleteKey ip, skeyname
    End If
End Function
'实例：删除上面建立的HKEY_CURRENT_USER下的项"xiaopeng"
'代码为 deletekey HKEY_CURRENT_USER ,"xiaopeng"
'******************************************************************
'********新建,设置数值名称*****************************************
Public Sub Setkeyvalue(ByVal ip As Long, ByVal keyname As String, ByVal valuename As String, ByVal valuesetting As Variant, ByVal valuetype As Long)
    Dim retval As Long
    Dim hKey As Long
    If RegOpenKeyEx(ip, keyname, 0, KEY_ALL_ACCESS, hKey) > 0 Then Exit Sub
    Select Case valuetype
        Case REG_SZ
             RegSetValueExString hKey, valuename, 0&, REG_SZ, valuesetting, Len(valuesetting)
        Case REG_DWORD
             RegSetValueExLong hKey, valuename, 0, valuetype, valuesetting, 4
    End Select
    RegCloseKey (hKey)
End Sub
'实例：在HKEY_CURRENT_USER下的项"xiaopeng"中建立名为"redice",键值为"is xiaopeng",类型为REG_SZ的新键
'代码为 setkeyvalue HKEY_CURRENT_USER ,"xiaopeng" ,"redice","is xiaopeng",REG_SZ
'又如:在HKEY_CURRENT_USER下的项"xiaopeng"中建立名为"ceshi",键值为2,类型为REG_DWORD的新键
'代码为"setkeyvalue HKEY_CURRENT_USER,"xiaopeng","ceshi",2,REG_DWORD
'********************************************************************************
'*********删除数值名称*********************************************************
Public Sub Deletevalue(ByVal ip As Long, ByVal keyname As String, ByVal valuename As String)
    Dim retval As Long
    Dim hKey As Long
    retval = RegOpenKeyEx(ip, keyname, 0, KEY_ALL_ACCESS, hKey)
    If retval > 0 Then
        Exit Sub
    End If
    RegDeleteValue hKey, valuename
    RegCloseKey hKey
End Sub
'实例：删除HKEY_CURRENT_USER下的项"xiaopeng"中名为"redice"的新键
'代码为 deletevalue HKEY_CURRENT_USER ,"xiaopeng","redice"
'******************************************************************
'**********查询已存在的数值内容************************************
Public Function getvalue(ByVal ip As Long, keyname As String, valuename As String, ByVal valuetype As Long) As String
    Dim retval As Long
    Dim hKey As Long
    Dim valuesetting As Variant
    Dim cddata As Long
    Dim lvalue As Long
    Dim svalue As String
    
    retval = RegOpenKeyEx(ip, keyname, 0, KEY_ALL_ACCESS, hKey)
    If retval > 0 Then
        getvalue = ""
        Exit Function
    End If
    retval = RegQueryValueEx(hKey, valuename, 0, valuetype, ByVal vbNullString, cddata)
    If retval <> 0 Then
        RegCloseKey hKey
        Exit Function
    End If
    Select Case valuetype
        Case REG_SZ
            svalue = String(cddata, Chr(0))
            RegQueryValueEx hKey, valuename, 0, valuetype, ByVal svalue, cddata
            valuesetting = Left$(svalue, cddata)
            getvalue = CStr(valuesetting)
        Case REG_DWORD
            RegQueryValueEx hKey, valuename, 0, valuetype, lvalue, cddata
            valuesetting = lvalue
            getvalue = CStr(valuesetting)
    End Select
End Function
'实例：获取HKEY_CURRENT_USER下的项"xiaopeng"中名为"redice"的新键的键值
'代码为 getvalue HKEY_CURRENT_USER ,"xiaopeng","redice"

'--------------------------------------------------------------------------------------------------------------------------------------------------------
'功能：获取程序调用的进程名
'--------------------------------------------------------------------------------------------------------------------------------------------------------
Public Sub GetProcessName(ByVal processID As Long, szExeName As String, szPathName As String)
'szExeName 返回的完整文件名,没找到 返回 ""
'szPathName 返回的完整路径名,以"\"  结束,没找到 返回 ""
    Dim my As PROCESSENTRY32
    Dim hProcessHandle As Long
    Dim success As Long
    Dim l As Long

    l = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0)
    If l Then
        my.dwSize = 1060
        If (Process32First(l, my)) Then
            Do
            If my.th32ProcessID = processID Then
               CloseHandle l
               szExeName = Left$(my.szExeFile, InStr(1, my.szExeFile, Chr$(0)) - 1)
               For l = Len(szExeName) To 1 Step -1
                   If Mid$(szExeName, l, 1) = "\" Then Exit For
               Next l
               szPathName = Left$(szExeName, l)
               Exit Sub
            End If
            Loop Until (Process32Next(l, my) < 1)
        End If
        CloseHandle l
    End If
End Sub

Public Function SetWBIEVerSion(ByVal strExeName As String, strMsg As String) As Boolean
    '******************************************************************************************************************
    '功能：设置浏览器版本IE11,只有在非IDE环境下设置才能起作用
    '返回：
    '******************************************************************************************************************
    Dim strLocal As String
    If Is64bit Then
        strLocal = "SOFTWARE\Wow6432Node\Microsoft\Internet Explorer\MAIN\FeatureControl\FEATURE_BROWSER_EMULATION"
    Else
        strLocal = "SOFTWARE\Microsoft\Internet Explorer\MAIN\FeatureControl\FEATURE_BROWSER_EMULATION"
    End If
    If getvalue(HKEY_LOCAL_MACHINE, strLocal, strExeName, REG_DWORD) <> "11000" Then
        '新建注册表项
        Call Setkeyvalue(HKEY_LOCAL_MACHINE, strLocal, strExeName, "11000", REG_DWORD)
        '检查注册表是否写入成功
        If getvalue(HKEY_LOCAL_MACHINE, strLocal, strExeName, REG_DWORD) <> "11000" Then
            strMsg = "修改WebBrowser网页控件默认使用IE11浏览器失败，将无法加载整体护理页面数据，请手工添加注册表！" & vbCrLf & _
                "32位系统：HKEY_LOCAL_MACHINE" & strLocal & vbCrLf & _
                "64位系统：HKEY_LOCAL_MACHINE" & strLocal & vbCrLf & _
                "插入内容：[名称]-exe程序名　[类型]-REG_DWORD    [数值]-11000"
            Exit Function
        End If
    End If
    
    SetWBIEVerSion = True
End Function

