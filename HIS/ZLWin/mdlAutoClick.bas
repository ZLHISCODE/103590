Attribute VB_Name = "mdlAutoClick"
Option Explicit
'==================================================================================================
'编写           lshuo
'日期           2019/3/1
'模块           mdlAutoClick
'说明           实现自动点击模块功能
'==================================================================================================
'该函数用于判断指定的窗口是否允许接受键盘或鼠标输入。
Public Declare Function IsWindowEnabled Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long
'返回值：如果窗口句柄标识了一个已存在的窗口，返回值为非零；如果窗口句柄未标识一个已存在窗口，返回值为零
Public Declare Function isWindow Lib "user32" Alias "IsWindow" (ByVal hwnd As Long) As Long
Private Declare Function IsHungAppWindow Lib "user32" (ByVal hwnd As Long) As Long
'该函数确定给定窗口是否是最大化的窗口。
Public Declare Function IsZoomed Lib "user32" (ByVal hwnd As Long) As Long
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
Public Const GW_HWNDFIRST      As Long = 0
Public Const GW_HWNDLAST       As Long = 1
Public Const GW_HWNDNEXT       As Long = 2
Public Const GW_HWNDPREV       As Long = 3
Public Const GW_OWNER          As Long = 4
Public Const GW_CHILD          As Long = 5
Public Const GW_ENABLEDPOPUP   As Long = 6
Public Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long

'取得一个窗体的标题（caption）文字，或者一个控件的内容（在vb里使用：使用vb窗体或控件的caption或text属性）
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare Function GetModuleFileName Lib "kernel32" Alias "GetModuleFileNameA" (ByVal hModule As Long, ByVal lpFileName As String, ByVal nSize As Long) As Long

Private Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Private Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
'返回值是菜单的句柄。如果给定的窗口没有菜单，则返回NULL。如果窗口是一个子窗口，返回值无定义
Public Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long

Public Declare Function GetActiveWindow Lib "user32" () As Long
Public Declare Function SetActiveWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetForegroundWindow Lib "user32" () As Long
Private Declare Function SetFocus Lib "user32" (ByVal hwnd As Long) As Long

Private Declare Function GetProp Lib "user32" Alias "GetPropA" (ByVal hwnd As Long, ByVal lpString As String) As Long

Private Const SMTO_NORMAL = &H0
Private Const SMTO_BLOCK = &H1
Private Const SMTO_ABORTIFHUNG = &H2
Private Const SMTO_NOTIMEOUTIFNOTHUNG = &H8
Private Declare Function SendMessageTimeout Lib "user32.dll" Alias "SendMessageTimeoutA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any, ByVal fuFlags As Long, ByVal uTimeout As Long, lpdwResult As Long) As Long
'Private Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const WM_COMMAND As Long = &H111
Private Const WM_SYSCOMMAND = &H112
Private Const WM_CLOSE = &H10
Private Const SC_CLOSE = &HF060& '关闭窗体
Private Const SC_MINIMIZE = &HF020& '最小化窗体
Private Const SC_MAXIMIZE = &HF030& '最大化窗体
Private Const SC_RESTORE = &HF120& '恢复窗体大小
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const WM_LBUTTONDOWN = &H201 '左键按下
Private Const WM_LBUTTONUP = &H202 '左键弹起
Private Const MK_LBUTTON = &H1
Private Const BM_CLICK = &HF5 '单击
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)
Private Const MOUSEEVENTF_ABSOLUTE = &H8000 '指定鼠标使用绝对坐标系，此时，屏幕在水平和垂直方向上均匀分割成65535×65535个单元
Private Const MOUSEEVENTF_MOVE = &H1 '移动鼠标
Private Const MOUSEEVENTF_LEFTDOWN = &H2 '模拟鼠标左键按下
Private Const MOUSEEVENTF_LEFTUP = &H4 '模拟鼠标左键抬起
Private Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function CloseWindow Lib "user32" (ByVal hwnd As Long) As Long
'返回（retrieve）从操作系统启动所经过（elapsed）的毫秒数
Public Declare Function GetTickCount Lib "kernel32" () As Long


Public Declare Function GetCurrentProcessId Lib "kernel32" () As Long
'找出某个窗口的创建者(线程或进程)，返回创建者的标志符。
Public Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long

Private Type RECT
        Left As Long
        ToP As Long
        Right As Long
        Bottom As Long
End Type
Private Type POINTAPI
        x As Long
        Y As Long
End Type



'字符串用UTF-8编码
Private Const CP_UTF8 = 65001
Private Declare Function MultiByteToWideChar Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, lpMultiByteStr As Any, ByVal cchMultiByte As Long, lpWideCharStr As Any, ByVal cchWideChar As Long) As Long
Private Declare Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, lpWideCharStr As Any, ByVal cchWideChar As Long, lpMultiByteStr As Any, ByVal cchMultiByte As Long, lpDefaultChar As Any, ByVal lpUsedDefaultChar As Long) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function GetQueueStatus Lib "user32" (ByVal flags As Long) As Long
'@原型
'    DWORD WINAPI GetQueueStatus(
'      _In_ UINT flags
'    );
'@功能
'    检索在调用线程的消息队列中找到的消息类型。
'@参数
'    flags
'    要检查的消息类型?此参数可以是以下值中的一个或多个?
Public Const QS_ALLPOSTMESSAGE As Long = &H100
'    队列中有一条已发布的消息(这里列出的消息除外)。
Public Const QS_HOTKEY As Long = &H80
'    队列中有一条WM_HOTKEY消息?
Public Const QS_KEY     As Long = &H1
'    队列中有WM_KEYUP?WM_KEYDOWN?WM_SYSKEYUP或WM_SYSKEYDOWN消息?
Public Const QS_MOUSEBUTTON As Long = &H4
'    一个鼠标按钮消息(WM_LBUTTONUP, WM_RBUTTONDOWN，等等)。
Public Const QS_MOUSEMOVE As Long = &H2
'    队列中有一条WM_MOUSEMOVE消息?
Public Const QS_PAINT As Long = &H20
'    队列中有一条WM_PAINT消息?
Public Const QS_POSTMESSAGE As Long = &H8
'    队列中有一条已发布的消息(这里列出的消息除外)。
Public Const QS_RAWINPUT    As Long = &H400
'    队列中有一条原始输入消息。有关更多信息，请参见原始输入。
'    Windows 2000: 不支持此标志?
Public Const QS_SENDMESSAGE As Long = &H40
'    队列中有另一个线程或应用程序发送的消息?
Public Const QS_TIMER   As Long = 10
'    队列中有一条WM_TIMER消息?
Public Const QS_MOUSE   As Long = (QS_MOUSEMOVE Or QS_MOUSEBUTTON)
'    一个WM_MOUSEMOVE消息或鼠标按钮消息(WM_LBUTTONUP, WM_RBUTTONDOWN，等等)。
Public Const QS_INPUT  As Long = (QS_MOUSE Or QS_KEY Or QS_RAWINPUT)
'    队列中有一条输入消息?
Public Const QS_ALLEVENTS As Long = (QS_INPUT Or QS_POSTMESSAGE Or QS_TIMER Or QS_PAINT Or QS_HOTKEY)
'    输入、WM_TIMER、WM_PAINT、WM_HOTKEY或已发布的消息都在队列中?
Public Const QS_ALLINPUT As Long = (QS_INPUT Or QS_POSTMESSAGE Or QS_TIMER Or QS_PAINT Or QS_HOTKEY Or QS_SENDMESSAGE)
'    任何消息都在队列中?

'@返回值
'    返回值的高字表示队列中当前消息的类型。低字表示已添加到队列中的消息类型，以及自上次调用GetQueueStatus、GetMessage或PeekMessage函数以来仍在队列中的消息类型。
'@备注
'    返回值中出现QS_标志并不保证随后对GetMessage或PeekMessage函数的调用将返回一条消息。GetMessage和PeekMessage执行一些内部过滤，这可能导致消息在内部处理。因此，应该只考虑GetQueueStatus的返回值作为调用GetMessage还是PeekMessage的提示。
'    QS_ALLPOSTMESSAGE和QS_POSTMESSAGE标记在清除时有所不同。无论是否过滤消息，当您调用GetMessage或PeekMessage时，都会清除QS_POSTMESSAGE。当您在没有过滤消息的情况下调用GetMessage或PeekMessage时，QS_ALLPOSTMESSAGE将被清除(wMsgFilterMin和wMsgFilterMax为0)。

Public glngLastVertify      As Long         '上次开始启动模块验证的时间
Public glngLastHwnd         As Long         '上次点击的窗口句柄
Public gblnFinish           As Boolean
Private mlngVBMain          As Long         'VB进程主窗体句柄
Private mlngPid             As Long
Private mstrVBControlHeader As String
Public gobjLog              As New clsLog
Public gstrCurSID           As String        '当前操作系统当前会话SID
Public Const G_CURRENT_SESSION_SID As String = "DB130B49-92AA-488D-9D58-C1671CD21673"          '存储SID
Private Enum WinInfo
    WI_Hwnd = 0
    WI_OwnerHwnd = 1
    WI_PopHwnd = 2
    WI_Enable = 3
    WI_Class = 4
    WI_Text = 5
End Enum
'--------------------------------------------------------------------------------------------------
'方法           IsAllWindowClosed
'功能           判断是否所有窗体已经关闭
'返回值         Boolean
'入参列表:
'参数名         类型                    说明
'lngBrwHwnd     导航台主窗体
'-------------------------------------------------------------------------------------------------
Public Function IsAllWindowClosed(ByVal lngBrwHwnd As Long) As Boolean
    Dim lngTopHwnd          As Long
    On Error GoTo ErrH
    Call gobjLog.PushMethod(RLL_AllLog, "ZlWin.mdlAutoClick.IsAllWindowClosed", lngBrwHwnd)
    If mlngVBMain = 0 Then mlngVBMain = FindVBMain(lngBrwHwnd)
    If mlngPid = 0 Then mlngPid = GetCurrentProcessId
    lngTopHwnd = GetWindow(mlngVBMain, GW_ENABLEDPOPUP)
    If lngTopHwnd <> 0 And lngTopHwnd = lngBrwHwnd Then
        IsAllWindowClosed = True
        If gobjLog.CurrentLogLevel > RLL_NoneLog Then Call GetAllWindows(mlngPid, mlngVBMain)
    End If
    Call gobjLog.PopMethod(RLL_AllLog, "ZlWin.mdlAutoClick.IsAllWindowClosed", IsAllWindowClosed)
    Exit Function
ErrH:
    If gobjLog.ErrCenter(RLL_RunError, "ZlWin.mdlAutoClick.IsAllWindowClosed") = 1 Then
        Resume
    End If
End Function
'--------------------------------------------------------------------------------------------------
'方法           ClickCurrentWindow
'功能           点击当前的业务窗口，直到只有导航台主窗口
'返回值         Boolean
'入参列表:
'参数名         类型                    说明
'blnCloseAll    Boolean                 是否关闭所有窗口
'-------------------------------------------------------------------------------------------------
Public Function ClickCurrentWindow(ByVal lngBrwHwnd As Long, ByRef strMsg As String) As Boolean
    Dim lngTopHwnd          As Long
    Dim lngTimes            As Long
    Dim strTmp              As String
    On Error GoTo ErrH
    Call gobjLog.PushMethod(RLL_AllLog, "ZlWin.mdlAutoClick.ClickCurrentWindow", lngBrwHwnd, strMsg)
    If mstrVBControlHeader = "" Then
        If IsDesinMode Then
            mstrVBControlHeader = "Thunder"
        Else
            mstrVBControlHeader = "ThunderRT6"
        End If
    End If
    If mlngVBMain = 0 Then mlngVBMain = FindVBMain(lngBrwHwnd)
    If mlngPid = 0 Then mlngPid = GetCurrentProcessId
    lngTopHwnd = GetWindow(mlngVBMain, GW_ENABLEDPOPUP)
    '查找到确定按钮
    If lngTopHwnd <> 0 And lngTopHwnd <> lngBrwHwnd Then
        lngTimes = 1
        Do While lngTopHwnd <> 0 And lngTopHwnd <> lngBrwHwnd And IsCanDoMessage(lngTopHwnd)
            strTmp = CloseOneWindown(lngTopHwnd)
            If glngLastHwnd <> lngTopHwnd Then
                If strMsg <> "" Then
                    strMsg = strMsg & " " & strTmp
                Else
                    strMsg = strTmp
                End If
                strMsg = SubB(strMsg, 1, 450)
            End If
            glngLastHwnd = lngTopHwnd
            gobjLog.LogInfo RLL_LogInfo, "ClickCurrentWindow", "blnCloseAll=True", "Times=" & lngTimes, "VBMain =" & mlngVBMain, "BrwHwnd=" & lngBrwHwnd, "TopHwnd=" & lngTopHwnd, "strMsg=" & strMsg
            lngTimes = lngTimes - 1
            If lngTimes <= 0 Then
                Exit Do
            End If
            lngTopHwnd = GetWindow(mlngVBMain, GW_ENABLEDPOPUP)
        Loop
        gobjLog.LogInfo RLL_LogInfo, "ClickCurrentWindow", "blnCloseAll=True", "Times=" & lngTimes, "VBMain =" & mlngVBMain, "BrwHwnd=" & lngBrwHwnd, "TopHwnd=" & lngTopHwnd, "isWindow=" & isWindow(lngTopHwnd)
    End If
    If gobjLog.CurrentLogLevel > RLL_NoneLog Then Call GetAllWindows(mlngPid, mlngVBMain)
    ClickCurrentWindow = lngTopHwnd = lngBrwHwnd
    gobjLog.LogInfo RLL_LogInfo, "ClickCurrentWindow", "VBMain =" & mlngVBMain, "TopHwnd=" & lngTopHwnd, "BrwHwnd=" & lngBrwHwnd, "isWindow=" & isWindow(lngTopHwnd)
    Call gobjLog.PopMethod(RLL_AllLog, "ZlWin.mdlAutoClick.ClickCurrentWindow", ClickCurrentWindow)
    Exit Function
ErrH:
    If gobjLog.ErrCenter(RLL_RunError, "ZlWin.mdlAutoClick.ClickCurrentWindow") = 1 Then
        Resume
    End If
End Function

'--------------------------------------------------------------------------------------------------
'方法           CloseOneWindown
'功能           关闭一个窗口
'返回值         String                  窗口的错误提示
'入参列表:
'参数名         类型                    说明
'lngHwnd        Long                    要关闭的窗口
'-------------------------------------------------------------------------------------------------
Private Function CloseOneWindown(ByVal lngHwnd As Long) As String
    Dim strError        As String
    Dim strWinClass     As String
    Dim lngOkBtnWnd     As Long
    Dim lngToolbar      As Long
    Dim lngToolbarBtn   As Long, lngToolBarBtnIdx       As Long
    On Error GoTo ErrH
    Call gobjLog.PushMethod(RLL_AllLog, "ZlWin.mdlAutoClick.CloseOneWindown", lngHwnd)
    strWinClass = GetWinClass(lngHwnd)
    If strWinClass = mstrVBControlHeader & "FormDC" Then
        If GetProp(lngHwnd, "IsErrCenter") = 1 Then       '是错误处理中心
            '错误中心可能回导致程序崩溃
            lngOkBtnWnd = FindExitButtonErrCenter(lngHwnd, strError)
            Call ClickButton(lngHwnd, lngOkBtnWnd)
        Else
            lngOkBtnWnd = FindExitButtonNormal(lngHwnd)
            If lngOkBtnWnd <> 0 Then
                Call ClickButton(lngHwnd, lngOkBtnWnd)
            Else
                '通过窗口关闭按钮处理
                Call ClickMenuExit(lngHwnd)
            End If
        End If
    ElseIf Mid(strWinClass, 1, 1) = "#" And IsNumeric(Mid(strWinClass, 2)) Then      '是Msgbox
        lngOkBtnWnd = FindExitButtonMsgBox(lngHwnd, strError)
        Call ClickButton(lngHwnd, lngOkBtnWnd)
    Else
        Call ClickMenuExit(lngHwnd)
    End If
    CloseOneWindown = strError
    Call gobjLog.PopMethod(RLL_AllLog, "ZlWin.mdlAutoClick.CloseOneWindown", CloseOneWindown)
    Exit Function
ErrH:
    If gobjLog.ErrCenter(RLL_RunError, "ZlWin.mdlAutoClick.CloseOneWindown") = 1 Then
        Resume
    End If
End Function
'--------------------------------------------------------------------------------------------------
'方法           GetAllWindows
'功能           获取所有的VB窗体
'返回值         Variant                 返回窗体数组，元素[Hwnd,OwnerHwnd,PopHwnd,IsEnable,class,Text,Module]
'入参列表:
'参数名         类型                    说明
'lngCurPID      Long                    当前进程ID
'lngVBMain      Long                    VB主窗体
'-------------------------------------------------------------------------------------------------
Public Function GetAllWindows(ByVal lngCurPID As Long, ByVal lngVBMain As Long) As Long
    Dim lngOwner        As Long, lngNext        As Long, lngPop         As Long
    Dim lngPid          As Long
'    Dim arrRet()        As Variant
    Dim strClass        As String
    Dim strWinText      As String
    Dim lngCount        As Long
    On Error GoTo ErrH
    Call gobjLog.PushMethod(RLL_AllLog, "ZlWin.mdlAutoClick.GetAllWindows", lngCurPID, lngVBMain)
    If gobjLog.CurrentLogLevel > RLL_NoneLog Then
        gobjLog.LogInfo RLL_LogInfo, "GetAllWindows", "VBMainHwnd =" & lngVBMain, "VBMainIsEnable=" & IsWindowEnabled(lngVBMain), "VBMainClass=" & GetWinClass(lngVBMain), "VBMainText=" & GetWinText(lngVBMain)
    End If
'    ReDim Preserve arrRet(0)
    lngNext = GetWindow(lngVBMain, GW_HWNDPREV)
    Do While lngNext <> 0
        lngPid = 0
        GetWindowThreadProcessId lngNext, lngPid
        If lngCurPID = lngPid Then
'            ReDim Preserve arrRet(UBound(arrRet) + 1)
            lngCount = lngCount + 1
            strClass = GetWinClass(lngNext)
            strWinText = GetWinText(lngNext)
            lngOwner = GetWindow(lngNext, GW_OWNER)
            lngPop = GetWindow(lngNext, GW_ENABLEDPOPUP)
'            arrRet(UBound(arrRet)) = Array(lngNext, lngOwner, lngPop, IsWindowEnabled(lngNext), strClass, strWinText)
            If gobjLog.CurrentLogLevel > RLL_NoneLog Then
                gobjLog.LogInfo RLL_LogInfo, "GetAllWindows", "NextHwnd =" & lngNext, "NextIsEnable=" & IsWindowEnabled(lngNext), "NextClass=" & strClass, "NextText=" & strWinText, _
                                        "OwnerHwnd =" & lngOwner, "OwnerIsEnable=" & IsWindowEnabled(lngOwner), "OwnerClass=" & GetWinClass(lngOwner), "OwnerText=" & GetWinText(lngOwner), _
                                        "PopHwnd =" & lngPop, "PopIsEnable=" & IsWindowEnabled(lngPop), "PopClass=" & GetWinClass(lngPop), "PopText=" & GetWinText(lngPop)
            End If
        End If
        lngNext = GetWindow(lngNext, GW_HWNDPREV)
    Loop
    GetAllWindows = lngCount
    Call gobjLog.PopMethod(RLL_AllLog, "ZlWin.mdlAutoClick.GetAllWindows", GetAllWindows)
    Exit Function
ErrH:
    If gobjLog.ErrCenter(RLL_RunError, "ZlWin.mdlAutoClick.GetAllWindows") = 1 Then
        Resume
    End If
End Function
'--------------------------------------------------------------------------------------------------
'方法           GetAllChild
'功能           获取所有的VB子窗体，包含子窗口的窗口。
'返回值         Variant                 返回窗体数组，元素[Hwnd,OwnerHwnd,IsEnable,class,Text,Module]
'入参列表:
'参数名         类型                    说明
'lngFrmHwnd     Long                    枚举的窗口
'blnNormalWin   Boolean                 是否是常规窗口
'-------------------------------------------------------------------------------------------------
Private Function GetAllChild(ByVal lngFrmHwnd As Long, Optional ByVal blnNormalWin As Boolean) As Variant
    Dim arrRet()        As Variant
    Dim strClass        As String
    Dim strWinText      As String
    Dim lngChildHwnd    As Long
    Dim blnAdd          As Boolean

    On Error GoTo ErrH
    Call gobjLog.PushMethod(RLL_AllLog, "ZlWin.mdlAutoClick.GetAllChild", lngFrmHwnd)
    ReDim Preserve arrRet(0)
    lngChildHwnd = GetWindow(lngFrmHwnd, GW_CHILD)
    gobjLog.LogInfo RLL_LogInfo, "GetAllChild", "CHILDHwnd =" & lngChildHwnd
    Do While lngChildHwnd <> 0
        strClass = GetWinClass(lngChildHwnd)
        blnAdd = False
        strWinText = ""
        If blnNormalWin Then
            If strClass = mstrVBControlHeader & "CommandButton" And IsWindowEnabled(lngChildHwnd) Then
                strWinText = GetWinText(lngChildHwnd)
                blnAdd = True
            End If
        Else
            blnAdd = True
        End If
        
        If blnAdd Then
            strWinText = GetWinText(lngChildHwnd)
            ReDim Preserve arrRet(UBound(arrRet) + 1)
            arrRet(UBound(arrRet)) = Array(lngChildHwnd, lngFrmHwnd, 0, IsWindowEnabled(lngChildHwnd), strClass, strWinText)
        End If
        gobjLog.LogInfo RLL_LogInfo, "GetAllChild", "NextChildHwnd =" & lngChildHwnd, "NextParent=" & lngFrmHwnd, "NextIsEnable=" & IsWindowEnabled(lngChildHwnd), "NextClass=" & strClass, "NextText=" & strWinText
        If IsWindowEnabled(lngChildHwnd) <> 0 And IsWindowVisible(lngChildHwnd) <> 0 Then
            Call GetAllChildSub(lngChildHwnd, blnNormalWin, arrRet)
        End If
        lngChildHwnd = GetWindow(lngChildHwnd, GW_HWNDNEXT)
    Loop
    GetAllChild = arrRet
    Call gobjLog.PopMethod(RLL_AllLog, "ZlWin.mdlAutoClick.GetAllChild", UBound(GetAllChild))
    Exit Function
ErrH:
    If gobjLog.ErrCenter(RLL_RunError, "ZlWin.mdlAutoClick.GetAllChild") = 1 Then
        Resume
    End If
End Function

'--------------------------------------------------------------------------------------------------
'方法           GetAllChildSub
'功能           GetAllChildSub的递归方法
'返回值
'入参列表:
'参数名         类型                    说明
'lngParentHwnd  Long                    父窗口
'arrRet         Variant                 已经存在的窗口数据，用来追加
'-------------------------------------------------------------------------------------------------
Private Sub GetAllChildSub(ByVal lngParentHwnd As Long, Optional ByVal blnNormalWin As Boolean, Optional ByRef arrRet As Variant, Optional ByVal lngLevel As Long = 0)
    Dim strClass        As String
    Dim strWinText      As String
    Dim lngChildHwnd    As Long
    Dim blnAdd          As Boolean
    
    On Error GoTo ErrH
    Call gobjLog.PushMethod(RLL_AllLog, "ZlWin.mdlAutoClick.GetAllChildSub", lngParentHwnd, UBound(arrRet), lngLevel)
    If lngLevel < 2 Then
        lngChildHwnd = GetWindow(lngParentHwnd, GW_CHILD)
        gobjLog.LogInfo RLL_LogInfo, "GetAllChildSub", "CHILDHwnd =" & lngChildHwnd
        Do While lngChildHwnd <> 0
            strClass = GetWinClass(lngChildHwnd)
            blnAdd = False
            strWinText = ""
            If blnNormalWin Then
                If strClass = mstrVBControlHeader & "CommandButton" And IsWindowEnabled(lngChildHwnd) Then
                    strWinText = GetWinText(lngChildHwnd)
                    blnAdd = True
                End If
            Else
                blnAdd = True
            End If
            If blnAdd Then
                strWinText = GetWinText(lngChildHwnd)
                ReDim Preserve arrRet(UBound(arrRet) + 1)
                arrRet(UBound(arrRet)) = Array(lngChildHwnd, lngParentHwnd, 0, IsWindowEnabled(lngChildHwnd), strClass, strWinText)
            End If
            gobjLog.LogInfo RLL_LogInfo, "GetAllChild", "NextChildHwnd =" & lngChildHwnd, "NextParent=" & lngParentHwnd, "NextIsEnable=" & IsWindowEnabled(lngChildHwnd), "NextClass=" & strClass, "NextText=" & strWinText
            If IsWindowEnabled(lngChildHwnd) <> 0 And IsWindowVisible(lngChildHwnd) <> 0 Then
                Call GetAllChildSub(lngChildHwnd, blnNormalWin, arrRet, lngLevel + 1)
            End If
            lngChildHwnd = GetWindow(lngChildHwnd, GW_HWNDNEXT)
        Loop
    End If
    Call gobjLog.PopMethod(RLL_AllLog, "ZlWin.mdlAutoClick.GetAllChildSub", UBound(arrRet))
    Exit Sub
ErrH:
    If gobjLog.ErrCenter(RLL_RunError, "ZlWin.mdlAutoClick.GetAllChildSub") = 1 Then
        Resume
    End If
End Sub

'方法           IsTopModal
'功能           判断当前进程的最顶层窗体是否是模态窗体。对VB起作用。通过判断导航台是禁用，顶层窗体是启用，则当前是模态
'返回值         Boolean
'入参列表:
'参数名         类型                    说明
'lngBrwHwnd     Long                    导航台句柄
'lngTopWin      Long                    最顶层窗体句柄
'-------------------------------------------------------------------------------------------------
Public Function IsTopModal(ByVal lngBrwHwnd As Long, Optional ByVal lngTopWin As Long) As Boolean
    On Error GoTo ErrH
    Call gobjLog.PushMethod(RLL_AllLog, "ZlWin.mdlAutoClick.IsTopModal", lngBrwHwnd, lngTopWin)
    If mlngVBMain = 0 Then mlngVBMain = FindVBMain(lngBrwHwnd)
    If mlngPid = 0 Then mlngPid = GetCurrentProcessId
    If lngTopWin = 0 Then
        lngTopWin = GetWindow(mlngVBMain, GW_ENABLEDPOPUP)
    End If
    If lngTopWin <> 0 And lngTopWin <> lngBrwHwnd Then
        IsTopModal = IsWindowEnabled(lngBrwHwnd) = 0 And IsWindowEnabled(lngTopWin) <> 0
    Else
        IsTopModal = True
    End If
    Call gobjLog.PopMethod(RLL_AllLog, "ZlWin.mdlAutoClick.IsTopModal", lngTopWin, lngBrwHwnd, IsTopModal)
    Exit Function
ErrH:
    If gobjLog.ErrCenter(RLL_RunError, "ZlWin.mdlAutoClick.IsTopModal") = 1 Then
        Resume
    End If
End Function
'--------------------------------------------------------------------------------------------------
'方法           IsCanDoMessage
'功能           判断窗口能否接受消息
'返回值         Boolean
'入参列表:
'参数名         类型                    说明
'lngTopWin      Long                    要判断的窗口
'-------------------------------------------------------------------------------------------------
Private Function IsCanDoMessage(ByVal lngTopWin As Long) As Boolean
    On Error GoTo ErrH
    Call gobjLog.PushMethod(RLL_AllLog, "ZlWin.mdlAutoClick.IsCanDoMessage", lngTopWin)
    IsCanDoMessage = True
    If IsHungAppWindow(lngTopWin) <> 0 Then
        gobjLog.LogInfo RLL_LogInfo, "IsCanDoMessage", "IsHungAppWindow=True"
        IsCanDoMessage = False
    End If
    If IsCanDoMessage Then
        If isWindow(lngTopWin) = 0 Or IsWindowVisible(lngTopWin) = 0 Or IsWindowEnabled(lngTopWin) = 0 Then
            gobjLog.LogInfo RLL_LogInfo, "IsCanDoMessage", "isWindow=" & isWindow(lngTopWin), "IsWindowVisible=" & IsWindowVisible(lngTopWin), "IsWindowEnabled=" & IsWindowEnabled(lngTopWin)
            IsCanDoMessage = False
        End If
    End If
    Call gobjLog.PopMethod(RLL_AllLog, "ZlWin.mdlAutoClick.IsCanDoMessage", IsCanDoMessage)
    Exit Function
ErrH:
    If gobjLog.ErrCenter(RLL_RunError, "ZlWin.mdlAutoClick.IsCanDoMessage") = 1 Then
        Resume
    End If
End Function
'--------------------------------------------------------------------------------------------------
'方法           FindVBMain
'功能           获取VB进程的主窗体
'返回值         Long
'入参列表:
'参数名         类型                    说明
'lngBrwHwnd     Long                    导航台句柄
'-------------------------------------------------------------------------------------------------
Private Function FindVBMain(ByVal lngBrwHwnd As Long) As Long
    Dim lngOwner        As Long
    Dim strClass        As String
    
    On Error GoTo ErrH
    Call gobjLog.PushMethod(RLL_AllLog, "ZlWin.mdlAutoClick.FindVBMain", lngBrwHwnd)
    lngOwner = GetWindow(lngBrwHwnd, GW_OWNER)
    Do While lngOwner <> 0
        strClass = GetWinClass(lngOwner)
        If strClass = mstrVBControlHeader & "Main" Then
            FindVBMain = lngOwner
            Exit Do
        Else
            lngOwner = GetWindow(lngOwner, GW_OWNER)
        End If
    Loop
    Call gobjLog.PopMethod(RLL_AllLog, "ZlWin.mdlAutoClick.FindVBMain", FindVBMain)
    Exit Function
ErrH:
    If gobjLog.ErrCenter(RLL_RunError, "ZlWin.mdlAutoClick.FindVBMain") = 1 Then
        Resume
    End If
End Function
'--------------------------------------------------------------------------------------------------
'方法           FindExitButtonNormal
'功能           寻找一个常规窗体界面的确定按钮。
'返回值         Long                    返回确定或取消按钮句柄
'入参列表:
'参数名         类型                    说明
'lngMainFrm     Long                    窗口句柄
'lngToolBarHwnd Long                    ToolBar的句柄（若窗口存在TOOLBar,且存在确定取消按钮，此时函数返回确定取消按钮ID）
'-------------------------------------------------------------------------------------------------
Private Function FindExitButtonNormal(ByVal lngMainFrm As Long) As Long
    Dim arrChild()      As Variant
    Dim i               As Long
    Dim strClass        As String, strWinText   As String
    Dim lngCancelBtn    As Long, lngOKBtn       As Long
    
    On Error GoTo ErrH
    Call gobjLog.PushMethod(RLL_AllLog, "ZlWin.mdlAutoClick.FindExitButtonNormal", lngMainFrm)
    arrChild = GetAllChild(lngMainFrm, True)
    For i = 1 To UBound(arrChild)
        strClass = arrChild(i)(WI_Class)
        strWinText = arrChild(i)(WI_Text)
        If strClass = mstrVBControlHeader & "CommandButton" And arrChild(i)(WI_Enable) <> 0 Then
            If strWinText = "确定" Or strWinText Like "确定(*)" Or strWinText = "保存" Or strWinText Like "保存(*)" Then
                gobjLog.LogInfo RLL_LogInfo, "FindExitButtonNormal", "OKBtn =" & arrChild(i)(WI_Hwnd), "Parent=" & arrChild(i)(WI_OwnerHwnd), "IsEnable=" & arrChild(i)(WI_Enable), "Class=" & arrChild(i)(WI_Class), "Text=" & arrChild(i)(WI_Text)
                lngOKBtn = arrChild(i)(WI_Hwnd)
                If lngCancelBtn <> 0 Then Exit For
            ElseIf strWinText = "退出" Or strWinText Like "退出(*)" Or strWinText = "关闭" Or strWinText Like "关闭(*)" Or strWinText = "取消" Or strWinText Like "取消(*)" Then
                gobjLog.LogInfo RLL_LogInfo, "FindExitButtonNormal", "CancelBtn =" & arrChild(i)(WI_Hwnd), "Parent=" & arrChild(i)(WI_OwnerHwnd), "IsEnable=" & arrChild(i)(WI_Enable), "Class=" & arrChild(i)(WI_Class), "Text=" & arrChild(i)(WI_Text)
                lngCancelBtn = arrChild(i)(WI_Hwnd)
                Exit For
            End If
        End If
    Next
    If lngCancelBtn <> 0 Then
        FindExitButtonNormal = lngCancelBtn
    Else
        FindExitButtonNormal = lngOKBtn
    End If
    Call gobjLog.PopMethod(RLL_AllLog, "ZlWin.mdlAutoClick.FindExitButtonNormal", FindExitButtonNormal, FindExitButtonNormal)
    Exit Function
ErrH:
    If gobjLog.ErrCenter(RLL_RunError, "ZlWin.mdlAutoClick.FindExitButtonNormal") = 1 Then
        Resume
    End If
End Function
'--------------------------------------------------------------------------------------------------
'方法           FindExitButtonMsgBox
'功能           寻找一个Msgbox窗体界面的确定按钮。
'返回值         Long                    返回一个按钮句柄
'入参列表:
'参数名         类型                    说明
'lngMainFrm     Long                    MsgBox的窗口句柄
'strMsg         String                  MsgBox的错误提示
'-------------------------------------------------------------------------------------------------
Private Function FindExitButtonMsgBox(ByVal lngMainFrm As Long, ByRef strMsg As String) As Long
    Dim arrChild()      As Variant
    Dim i               As Long
    Dim strErr          As String
    Dim strClass        As String, strWinText   As String
    Dim lngCancelBtn    As Long, lngOKBtn       As Long, lngDefault As Long
    Dim lngToolBarHwnd  As Long
    
    On Error GoTo ErrH
    Call gobjLog.PushMethod(RLL_AllLog, "ZlWin.mdlAutoClick.FindExitButtonMsgBox", lngMainFrm, strMsg)
    arrChild = GetAllChild(lngMainFrm)
    For i = 1 To UBound(arrChild)
        strClass = arrChild(i)(WI_Class)
        strWinText = arrChild(i)(WI_Text)
        If strClass = "Button" Then
            If strWinText Like "确定*" Or strWinText Like "是*" Then
                gobjLog.LogInfo RLL_LogInfo, "FindExitButtonMsgBox", "OKBtn =" & arrChild(i)(WI_Hwnd), "Parent=" & arrChild(i)(WI_OwnerHwnd), "IsEnable=" & arrChild(i)(WI_Enable), "Class=" & arrChild(i)(WI_Class), "Text=" & arrChild(i)(WI_Text)
                lngOKBtn = arrChild(i)(WI_Hwnd)
            ElseIf strWinText Like "否*" Or strWinText Like "取消*" Or strWinText Like "忽略*" Then
                lngCancelBtn = arrChild(i)(WI_Hwnd)
                gobjLog.LogInfo RLL_LogInfo, "FindExitButtonMsgBox", "CancelBtn =" & arrChild(i)(WI_Hwnd), "Parent=" & arrChild(i)(WI_OwnerHwnd), "IsEnable=" & arrChild(i)(WI_Enable), "Class=" & arrChild(i)(WI_Class), "Text=" & arrChild(i)(WI_Text)
            ElseIf lngDefault = 0 Then
                lngDefault = arrChild(i)(WI_Hwnd)
                gobjLog.LogInfo RLL_LogInfo, "FindExitButtonMsgBox", "DefaultBtn =" & arrChild(i)(WI_Hwnd), "Parent=" & arrChild(i)(WI_OwnerHwnd), "IsEnable=" & arrChild(i)(WI_Enable), "Class=" & arrChild(i)(WI_Class), "Text=" & arrChild(i)(WI_Text)
            End If
        ElseIf strClass = "Static" Then
            If Trim(strWinText) <> "" Then
                strErr = strErr & Trim(strWinText)
                gobjLog.LogInfo RLL_LogInfo, "FindExitButtonMsgBox", "InfoHwnd =" & arrChild(i)(WI_Hwnd), "Parent=" & arrChild(i)(WI_OwnerHwnd), "IsEnable=" & arrChild(i)(WI_Enable), "Class=" & arrChild(i)(WI_Class), "Text=" & arrChild(i)(WI_Text)
            End If
        End If
    Next
    If strErr <> "" Then
        If strMsg <> "" Then
            strMsg = strMsg & vbNewLine & "【消息提示】" & strErr
        Else
            strMsg = "【消息提示】" & strErr
        End If
    End If
    FindExitButtonMsgBox = lngDefault
    If lngCancelBtn <> 0 Then
        FindExitButtonMsgBox = lngCancelBtn
    ElseIf lngOKBtn <> 0 Then
        FindExitButtonMsgBox = lngOKBtn
    End If
    Call gobjLog.PopMethod(RLL_AllLog, "ZlWin.mdlAutoClick.FindExitButtonMsgBox", FindExitButtonMsgBox, strErr)
    Exit Function
ErrH:
    If gobjLog.ErrCenter(RLL_RunError, "ZlWin.mdlAutoClick.FindExitButtonMsgBox") = 1 Then
        Resume
    End If
End Function
'--------------------------------------------------------------------------------------------------
'方法           FindExitButtonErrCenter
'功能           寻找一个错误处理中心窗体界面的确定按钮。
'返回值         Long                    返回确定按钮句柄
'入参列表:
'参数名         类型                    说明
'lngMainFrm     Long                    ErrCenter的窗口句柄
'strMsg         String                  ErrCenter的错误提示
'-------------------------------------------------------------------------------------------------
Private Function FindExitButtonErrCenter(ByVal lngMainFrm As Long, ByRef strMsg As String) As Long
    Dim i               As Long
    Dim arrChild()      As Variant
    Dim strErrText      As String, strError As String
    Dim strClass        As String, strWinText   As String, lngWnd       As Long
    Dim lngCancelBtn    As Long, lngOKBtn       As Long
    Dim lngToolBarHwnd  As Long
    On Error GoTo ErrH
    Call gobjLog.PushMethod(RLL_AllLog, "ZlWin.mdlAutoClick.FindExitButtonErrCenter", lngMainFrm, strMsg)
    arrChild = GetAllChild(lngMainFrm)
    For i = 1 To UBound(arrChild)
        strClass = arrChild(i)(WI_Class)
        strWinText = arrChild(i)(WI_Text)
        If strClass = mstrVBControlHeader & "CommandButton" Then
            If strWinText Like "确定(*)" Then
                gobjLog.LogInfo RLL_LogInfo, "FindExitButtonErrCenter", "OKBtn =" & arrChild(i)(WI_Hwnd), "Parent=" & arrChild(i)(WI_OwnerHwnd), "IsEnable=" & arrChild(i)(WI_Enable), "Class=" & arrChild(i)(WI_Class), "Text=" & arrChild(i)(WI_Text)
                lngOKBtn = arrChild(i)(WI_Hwnd)
            ElseIf strWinText Like "取消(*)" Then
                gobjLog.LogInfo RLL_LogInfo, "FindExitButtonErrCenter", "CancelBtn =" & arrChild(i)(WI_Hwnd), "Parent=" & arrChild(i)(WI_OwnerHwnd), "IsEnable=" & arrChild(i)(WI_Enable), "Class=" & arrChild(i)(WI_Class), "Text=" & arrChild(i)(WI_Text)
                lngCancelBtn = arrChild(i)(WI_Hwnd)
            End If
        ElseIf strClass = mstrVBControlHeader & "TextBox" Then
            If Trim(strWinText) <> "" Then
                strErrText = strErrText & Trim(strWinText)
                gobjLog.LogInfo RLL_LogInfo, "FindExitButtonErrCenter", "InfoHwnd =" & arrChild(i)(WI_Hwnd), "Parent=" & arrChild(i)(WI_OwnerHwnd), "IsEnable=" & arrChild(i)(WI_Enable), "Class=" & arrChild(i)(WI_Class), "Text=" & arrChild(i)(WI_Text)
            End If
        ElseIf strClass = "Static" Then
            If strWinText <> "说明：" And Not strWinText Like "序号：*" Or strWinText <> "错误编号：" And Not strWinText Like "*秒后将自动截图并关闭此界面" Or strWinText <> "倒计时" Then
                gobjLog.LogInfo RLL_LogInfo, "FindExitButtonErrCenter", "InfoHwndStatic =" & arrChild(i)(WI_Hwnd), "Parent=" & arrChild(i)(WI_OwnerHwnd), "IsEnable=" & arrChild(i)(WI_Enable), "Class=" & arrChild(i)(WI_Class), "Text=" & arrChild(i)(WI_Text)
                strError = strWinText
            End If
        End If
    Next
    If strError <> "" Then
        If strErrText <> "" Then
            strError = strError & strErrText
        End If
    Else
        strError = strErrText
    End If
    If strError <> "" Then
        If strMsg <> "" Then
            strMsg = strMsg & vbNewLine & "【错误处理】" & strError
        Else
            strMsg = "【错误处理】" & strError
        End If
    End If
    If lngCancelBtn <> 0 Then
        FindExitButtonErrCenter = lngCancelBtn
    Else
        FindExitButtonErrCenter = lngOKBtn
    End If
    Call gobjLog.PopMethod(RLL_AllLog, "ZlWin.mdlAutoClick.FindExitButtonErrCenter", FindExitButtonErrCenter, strError)
    Exit Function
ErrH:
    If gobjLog.ErrCenter(RLL_RunError, "ZlWin.mdlAutoClick.FindExitButtonErrCenter") = 1 Then
        Resume
    End If
End Function

'--------------------------------------------------------------------------------------------------
'方法           GetWinModule
'功能           获取窗体对应任务文件模块，只对最上层模块有效
'返回值         String
'入参列表:
'参数名         类型                    说明
'lngHwnd        Long                    窗口句柄
'-------------------------------------------------------------------------------------------------
'Private Function GetWinModule(ByVal lngHwnd As Long) As String
'    Dim lngLen       As Long
'    Dim strTMp       As String * 256
'
'    If lngHwnd <> 0 Then
'        lngLen = GetModuleFileName(lngHwnd, strTMp, Len(strTMp) - 1)
'        GetWinModule = Left(strTMp, InStr(strTMp, vbNullChar) - 1)
'    End If
'End Function
'--------------------------------------------------------------------------------------------------
'方法           GetWinClass
'功能           获取窗口类名
'返回值         String
'入参列表:
'参数名         类型                    说明
'lngHwnd        Long                    窗口句柄
'-------------------------------------------------------------------------------------------------
Private Function GetWinClass(ByVal lngHwnd As Long) As String
    Dim lngLen       As Long
    Dim strTmp       As String * 256
    If lngHwnd <> 0 Then
        lngLen = GetClassName(lngHwnd, strTmp, Len(strTmp) - 1)
        GetWinClass = Left(strTmp, InStr(strTmp, vbNullChar) - 1)
    End If
End Function

'--------------------------------------------------------------------------------------------------
'方法           GetWinText
'功能           获取窗口文本内容
'返回值         String
'入参列表:
'参数名         类型                    说明
'lngHwnd        Long                    窗口句柄
'-------------------------------------------------------------------------------------------------
Private Function GetWinText(ByVal lngHwnd As Long) As String  '支持星号密码
    Dim lngLen       As Long
    Dim strTmp       As String * 1024
    If lngHwnd <> 0 Then
        Call GetWindowText(lngHwnd, strTmp, Len(strTmp) - 1)
        GetWinText = Left(strTmp, InStr(strTmp, vbNullChar) - 1)
    End If
End Function

'--------------------------------------------------------------------------------------------------
'方法           ClickButton
'功能           点击按钮
'返回值
'入参列表:
'参数名         类型                    说明
'lngMainFrm     Long                    按钮所在的窗体
'lngBtnHwnd     Long                    按钮句柄
'-------------------------------------------------------------------------------------------------
Public Sub ClickButton(ByVal lngMainFrm As Long, ByVal lngBtnHwnd As Long)
'    Dim vRect       As RECT, pCenter As POINTAPI
    Dim lngResult       As Long
    Dim lngReturnValue  As Long
    
    
    On Error GoTo ErrH
    Call gobjLog.PushMethod(RLL_AllLog, "ZlWin.mdlAutoClick.ClickButton", lngMainFrm, lngBtnHwnd)
    '防止多次触发Active时间
    If GetForegroundWindow() <> lngMainFrm Then
        Call SetForegroundWindow(lngMainFrm)
    End If
    If GetActiveWindow() <> lngMainFrm Then
        Call SetActiveWindow(lngMainFrm)
    End If
    Call SetFocus(lngBtnHwnd)
    '单击按钮方法1
'    Call GetWindowRect(lngBtnHwnd, vRect)
'    pCenter.x = (vRect.Left + vRect.Right) / 2
'    pCenter.Y = (vRect.ToP + vRect.Bottom) / 2
'    pCenter.x = pCenter.x / (Screen.Width / Screen.TwipsPerPixelX) * 65535
'    pCenter.Y = pCenter.Y / (Screen.Height / Screen.TwipsPerPixelY) * 65535
'    mouse_event MOUSEEVENTF_ABSOLUTE Or MOUSEEVENTF_MOVE, pCenter.x, pCenter.Y, 0, 0
'    mouse_event MOUSEEVENTF_ABSOLUTE Or MOUSEEVENTF_LEFTDOWN, pCenter.x, pCenter.Y, 0, 0
'    mouse_event MOUSEEVENTF_ABSOLUTE Or MOUSEEVENTF_LEFTUP, pCenter.x, pCenter.Y, 0, 0
    '单击按钮方法2
'    PostMessage lngBtnHwnd, WM_LBUTTONDOWN, 0, ByVal 0 '鼠标在按钮按下
'    PostMessage lngBtnHwnd, WM_LBUTTONUP, 0, ByVal 0 '鼠标在按钮弹起
    '单击按钮方法3
'    PostMessage lngBtnHwnd, BM_CLICK, 0, 0 '单击按钮
    lngReturnValue = SendMessageTimeout(lngBtnHwnd, BM_CLICK, 0, ByVal 0, SMTO_ABORTIFHUNG Or SMTO_BLOCK, 1000, lngResult)
    Call gobjLog.PopMethod(RLL_AllLog, "ZlWin.mdlAutoClick.ClickButton")
    Exit Sub
ErrH:
    If gobjLog.ErrCenter(RLL_RunError, "ZlWin.mdlAutoClick.ClickButton") = 1 Then
        Resume
    End If
End Sub

'--------------------------------------------------------------------------------------------------
'方法           ClickMenuExit
'功能           点击窗口的关闭按钮
'返回值
'入参列表:
'参数名         类型                    说明
'lngMainFrm     Long                    按钮所在的窗体
'-------------------------------------------------------------------------------------------------
Public Sub ClickMenuExit(ByVal lngMainFrm As Long)
    Dim lngResult       As Long
    Dim lngReturnValue  As Long
    
    On Error GoTo ErrH
    Call gobjLog.PushMethod(RLL_AllLog, "ZlWin.mdlAutoClick.ClickMenuExit", lngMainFrm)
    '防止多次触发Active时间
    If GetForegroundWindow() <> lngMainFrm Then
        Call SetForegroundWindow(lngMainFrm)
    End If
    If GetActiveWindow() <> lngMainFrm Then
        Call SetActiveWindow(lngMainFrm)
    End If
    lngReturnValue = SendMessageTimeout(lngMainFrm, WM_CLOSE, 0, ByVal 0, SMTO_ABORTIFHUNG Or SMTO_BLOCK, 1000, lngResult)
    If lngReturnValue <> 0 Then
        If lngResult = 0 Then
            gobjLog.LogInfo RLL_LogInfo, "ClickMenuExit", "窗口消息处理成功=" & lngMainFrm
        End If
    End If
    Call gobjLog.PopMethod(RLL_AllLog, "ZlWin.mdlAutoClick.ClickMenuExit", lngReturnValue, lngResult, Err.LastDllError)
    Exit Sub
ErrH:
    If gobjLog.ErrCenter(RLL_RunError, "ZlWin.mdlAutoClick.ClickMenuExit") = 1 Then
        Resume
    End If
End Sub
'--------------------------------------------------------------------------------------------------
'方法           GetTickCountDiff
'功能           计算GetTickCcout的差值。由于 GetTickCountVB会产生负值以及归零现象因此需要单独处理
'返回值         Double
'入参列表:
'参数名         类型                    说明
'lngStart       Long                    起始时间
'-------------------------------------------------------------------------------------------------
Public Function GetTickCountDiff(ByVal lngStart As Long) As Double
    Dim lngCur          As Long
    Const M_OFFSET_4    As Double = 4294967296#         '无符号整形的最大值
    
    lngCur = GetTickCount
    If lngCur < lngStart Then
        GetTickCountDiff = M_OFFSET_4 - LongToUnsigned(lngStart) + LongToUnsigned(lngCur)
    Else
        GetTickCountDiff = lngCur - lngStart
    End If
End Function

Private Function LongToUnsigned(value As Long) As Double
    Const M_OFFSET_4    As Double = 4294967296#         '无符号整形的最大值
    If value < 0 Then LongToUnsigned = value + M_OFFSET_4 Else LongToUnsigned = value
End Function

Public Sub ClearObjects(ByVal blnTest As Boolean, Optional ByVal objNow As Object)
    Dim lngCount    As Long
    '清理当前未缓存的对象
    On Error Resume Next
    If Not objNow Is Nothing Then
        objNow.CloseWindows
    End If
    If Err.Number <> 0 Then Err.Clear
    If blnTest Then
        lngCount = UBound(gstrObj)
        If Err.Number <> 0 Then lngCount = -1: Err.Clear
        On Error GoTo ErrH
        If lngCount >= 0 Then
            For lngCount = 0 To UBound(gstrObj)
                On Error Resume Next
                gobjCls(lngCount).CloseWindows
                If Err.Number <> 0 Then Err.Clear
                On Error GoTo ErrH
            Next
            For lngCount = 0 To UBound(gstrObj)
                On Error Resume Next
                Set gobjCls(lngCount) = Nothing
                If Err.Number <> 0 Then Err.Clear
                On Error GoTo ErrH
            Next
            ReDim Preserve gobjCls(0)
            ReDim Preserve gstrObj(0)
        End If
    End If
    Exit Sub
ErrH:
    Err.Clear
End Sub

Public Function IsDesinMode() As Boolean
'功能： 确定当前模式为设计模式
     Err = 0: On Error Resume Next
     Debug.Print 1 / 0
     If Err <> 0 Then
        IsDesinMode = True
     Else
        IsDesinMode = False
     End If
     Err.Clear: Err = 0
End Function

Public Function TruncZero(ByVal strInput As String) As String
'功能：去掉字符串中\0以后的字符
    Dim lngPos As Long

    lngPos = InStr(strInput, Chr(0))
    If lngPos > 0 Then
        TruncZero = Mid(strInput, 1, lngPos - 1)
    Else
        TruncZero = strInput
    End If
End Function

Public Function ActualLen(ByVal strAsk As String) As Long
'功能：求取指定字符串的实际长度，用于判断实际包含双字节字符串的
'       实际数据存储长度
    ActualLen = LenB(StrConv(strAsk, vbFromUnicode))
End Function

Public Function SubB(ByVal strInfor As String, ByVal lngStart As Long, ByVal lngLen As Long) As String
'功能:读取指定字串的值,字串中可以包含汉字
 '入参:strInfor-原串
 '         lngStart-直始位置
'         lngLen-长度
'返回:子串
    Err = 0: On Error GoTo ErrH:
    SubB = StrConv(MidB(StrConv(strInfor, vbFromUnicode), lngStart, lngLen), vbUnicode)
    SubB = Replace(SubB, Chr(0), "")
    Exit Function
ErrH:
    Err.Clear
    SubB = ""
End Function
