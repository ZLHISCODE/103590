Attribute VB_Name = "mdlAutoClick"
Option Explicit
'==================================================================================================
'��д           lshuo
'����           2019/3/1
'ģ��           mdlAutoClick
'˵��           ʵ���Զ����ģ�鹦��
'==================================================================================================
'�ú��������ж�ָ���Ĵ����Ƿ�������ܼ��̻�������롣
Public Declare Function IsWindowEnabled Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long
'����ֵ��������ھ����ʶ��һ���Ѵ��ڵĴ��ڣ�����ֵΪ���㣻������ھ��δ��ʶһ���Ѵ��ڴ��ڣ�����ֵΪ��
Public Declare Function isWindow Lib "user32" Alias "IsWindow" (ByVal hwnd As Long) As Long
Private Declare Function IsHungAppWindow Lib "user32" (ByVal hwnd As Long) As Long
'�ú���ȷ�����������Ƿ�����󻯵Ĵ��ڡ�
Public Declare Function IsZoomed Lib "user32" (ByVal hwnd As Long) As Long
Public Const SW_HIDE = 0 '���ش��ڣ�������һ������
Public Const SW_SHOWNORMAL = 1 '�����ʾָ�����ڣ�����ô��ڱ���󻯻���С��������ԭ��ԭ���Ĵ�С��λ�á�
Public Const SW_SHOWMINIMIZED = 2 '�����С��ָ������
Public Const SW_SHOWMAXIMIZED = 3 '������ָ������
Public Const SW_MAXIMIZE = 3 '��ָ���Ĵ������
Public Const SW_SHOWNOACTIVATE = 4 '��������Ĵ�С��λ����ʾָ�����ڣ���ǰ���ڱ��ּ���
Public Const SW_SHOW = 5 '�Ե�ǰλ�úʹ�С�����
Public Const SW_MINIMIZE = 6 ' ��ָ���Ĵ�����С��
Public Const SW_SHOWMINNOACTIVE = 7 '����С����ʽ��ʾָ�����ڣ������ڱ��ּ���
Public Const SW_SHOWNA = 8 '�Ե�ǰ״̬��ʾָ�����ڣ���ǰ���ڱ��ּ���
Public Const SW_RESTORE = 9 '��ԭָ���Ĵ���
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Const GW_HWNDFIRST      As Long = 0
Public Const GW_HWNDLAST       As Long = 1
Public Const GW_HWNDNEXT       As Long = 2
Public Const GW_HWNDPREV       As Long = 3
Public Const GW_OWNER          As Long = 4
Public Const GW_CHILD          As Long = 5
Public Const GW_ENABLEDPOPUP   As Long = 6
Public Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long

'ȡ��һ������ı��⣨caption�����֣�����һ���ؼ������ݣ���vb��ʹ�ã�ʹ��vb�����ؼ���caption��text���ԣ�
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare Function GetModuleFileName Lib "kernel32" Alias "GetModuleFileNameA" (ByVal hModule As Long, ByVal lpFileName As String, ByVal nSize As Long) As Long

Private Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Private Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
'����ֵ�ǲ˵��ľ������������Ĵ���û�в˵����򷵻�NULL�����������һ���Ӵ��ڣ�����ֵ�޶���
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
Private Const SC_CLOSE = &HF060& '�رմ���
Private Const SC_MINIMIZE = &HF020& '��С������
Private Const SC_MAXIMIZE = &HF030& '��󻯴���
Private Const SC_RESTORE = &HF120& '�ָ������С
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const WM_LBUTTONDOWN = &H201 '�������
Private Const WM_LBUTTONUP = &H202 '�������
Private Const MK_LBUTTON = &H1
Private Const BM_CLICK = &HF5 '����
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)
Private Const MOUSEEVENTF_ABSOLUTE = &H8000 'ָ�����ʹ�þ�������ϵ����ʱ����Ļ��ˮƽ�ʹ�ֱ�����Ͼ��ȷָ��65535��65535����Ԫ
Private Const MOUSEEVENTF_MOVE = &H1 '�ƶ����
Private Const MOUSEEVENTF_LEFTDOWN = &H2 'ģ������������
Private Const MOUSEEVENTF_LEFTUP = &H4 'ģ��������̧��
Private Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function CloseWindow Lib "user32" (ByVal hwnd As Long) As Long
'���أ�retrieve���Ӳ���ϵͳ������������elapsed���ĺ�����
Public Declare Function GetTickCount Lib "kernel32" () As Long


Public Declare Function GetCurrentProcessId Lib "kernel32" () As Long
'�ҳ�ĳ�����ڵĴ�����(�̻߳����)�����ش����ߵı�־����
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



'�ַ�����UTF-8����
Private Const CP_UTF8 = 65001
Private Declare Function MultiByteToWideChar Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, lpMultiByteStr As Any, ByVal cchMultiByte As Long, lpWideCharStr As Any, ByVal cchWideChar As Long) As Long
Private Declare Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, lpWideCharStr As Any, ByVal cchWideChar As Long, lpMultiByteStr As Any, ByVal cchMultiByte As Long, lpDefaultChar As Any, ByVal lpUsedDefaultChar As Long) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function GetQueueStatus Lib "user32" (ByVal flags As Long) As Long
'@ԭ��
'    DWORD WINAPI GetQueueStatus(
'      _In_ UINT flags
'    );
'@����
'    �����ڵ����̵߳���Ϣ�������ҵ�����Ϣ���͡�
'@����
'    flags
'    Ҫ������Ϣ����?�˲�������������ֵ�е�һ������?
Public Const QS_ALLPOSTMESSAGE As Long = &H100
'    ��������һ���ѷ�������Ϣ(�����г�����Ϣ����)��
Public Const QS_HOTKEY As Long = &H80
'    ��������һ��WM_HOTKEY��Ϣ?
Public Const QS_KEY     As Long = &H1
'    ��������WM_KEYUP?WM_KEYDOWN?WM_SYSKEYUP��WM_SYSKEYDOWN��Ϣ?
Public Const QS_MOUSEBUTTON As Long = &H4
'    һ����갴ť��Ϣ(WM_LBUTTONUP, WM_RBUTTONDOWN���ȵ�)��
Public Const QS_MOUSEMOVE As Long = &H2
'    ��������һ��WM_MOUSEMOVE��Ϣ?
Public Const QS_PAINT As Long = &H20
'    ��������һ��WM_PAINT��Ϣ?
Public Const QS_POSTMESSAGE As Long = &H8
'    ��������һ���ѷ�������Ϣ(�����г�����Ϣ����)��
Public Const QS_RAWINPUT    As Long = &H400
'    ��������һ��ԭʼ������Ϣ���йظ�����Ϣ����μ�ԭʼ���롣
'    Windows 2000: ��֧�ִ˱�־?
Public Const QS_SENDMESSAGE As Long = &H40
'    ����������һ���̻߳�Ӧ�ó����͵���Ϣ?
Public Const QS_TIMER   As Long = 10
'    ��������һ��WM_TIMER��Ϣ?
Public Const QS_MOUSE   As Long = (QS_MOUSEMOVE Or QS_MOUSEBUTTON)
'    һ��WM_MOUSEMOVE��Ϣ����갴ť��Ϣ(WM_LBUTTONUP, WM_RBUTTONDOWN���ȵ�)��
Public Const QS_INPUT  As Long = (QS_MOUSE Or QS_KEY Or QS_RAWINPUT)
'    ��������һ��������Ϣ?
Public Const QS_ALLEVENTS As Long = (QS_INPUT Or QS_POSTMESSAGE Or QS_TIMER Or QS_PAINT Or QS_HOTKEY)
'    ���롢WM_TIMER��WM_PAINT��WM_HOTKEY���ѷ�������Ϣ���ڶ�����?
Public Const QS_ALLINPUT As Long = (QS_INPUT Or QS_POSTMESSAGE Or QS_TIMER Or QS_PAINT Or QS_HOTKEY Or QS_SENDMESSAGE)
'    �κ���Ϣ���ڶ�����?

'@����ֵ
'    ����ֵ�ĸ��ֱ�ʾ�����е�ǰ��Ϣ�����͡����ֱ�ʾ����ӵ������е���Ϣ���ͣ��Լ����ϴε���GetQueueStatus��GetMessage��PeekMessage�����������ڶ����е���Ϣ���͡�
'@��ע
'    ����ֵ�г���QS_��־������֤����GetMessage��PeekMessage�����ĵ��ý�����һ����Ϣ��GetMessage��PeekMessageִ��һЩ�ڲ����ˣ�����ܵ�����Ϣ���ڲ�������ˣ�Ӧ��ֻ����GetQueueStatus�ķ���ֵ��Ϊ����GetMessage����PeekMessage����ʾ��
'    QS_ALLPOSTMESSAGE��QS_POSTMESSAGE��������ʱ������ͬ�������Ƿ������Ϣ����������GetMessage��PeekMessageʱ���������QS_POSTMESSAGE��������û�й�����Ϣ������µ���GetMessage��PeekMessageʱ��QS_ALLPOSTMESSAGE�������(wMsgFilterMin��wMsgFilterMaxΪ0)��

Public glngLastVertify      As Long         '�ϴο�ʼ����ģ����֤��ʱ��
Public glngLastHwnd         As Long         '�ϴε���Ĵ��ھ��
Public gblnFinish           As Boolean
Private mlngVBMain          As Long         'VB������������
Private mlngPid             As Long
Private mstrVBControlHeader As String
Public gobjLog              As New clsLog
Public gstrCurSID           As String        '��ǰ����ϵͳ��ǰ�ỰSID
Public Const G_CURRENT_SESSION_SID As String = "DB130B49-92AA-488D-9D58-C1671CD21673"          '�洢SID
Private Enum WinInfo
    WI_Hwnd = 0
    WI_OwnerHwnd = 1
    WI_PopHwnd = 2
    WI_Enable = 3
    WI_Class = 4
    WI_Text = 5
End Enum
'--------------------------------------------------------------------------------------------------
'����           IsAllWindowClosed
'����           �ж��Ƿ����д����Ѿ��ر�
'����ֵ         Boolean
'����б�:
'������         ����                    ˵��
'lngBrwHwnd     ����̨������
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
'����           ClickCurrentWindow
'����           �����ǰ��ҵ�񴰿ڣ�ֱ��ֻ�е���̨������
'����ֵ         Boolean
'����б�:
'������         ����                    ˵��
'blnCloseAll    Boolean                 �Ƿ�ر����д���
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
    '���ҵ�ȷ����ť
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
'����           CloseOneWindown
'����           �ر�һ������
'����ֵ         String                  ���ڵĴ�����ʾ
'����б�:
'������         ����                    ˵��
'lngHwnd        Long                    Ҫ�رյĴ���
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
        If GetProp(lngHwnd, "IsErrCenter") = 1 Then       '�Ǵ���������
            '�������Ŀ��ܻص��³������
            lngOkBtnWnd = FindExitButtonErrCenter(lngHwnd, strError)
            Call ClickButton(lngHwnd, lngOkBtnWnd)
        Else
            lngOkBtnWnd = FindExitButtonNormal(lngHwnd)
            If lngOkBtnWnd <> 0 Then
                Call ClickButton(lngHwnd, lngOkBtnWnd)
            Else
                'ͨ�����ڹرհ�ť����
                Call ClickMenuExit(lngHwnd)
            End If
        End If
    ElseIf Mid(strWinClass, 1, 1) = "#" And IsNumeric(Mid(strWinClass, 2)) Then      '��Msgbox
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
'����           GetAllWindows
'����           ��ȡ���е�VB����
'����ֵ         Variant                 ���ش������飬Ԫ��[Hwnd,OwnerHwnd,PopHwnd,IsEnable,class,Text,Module]
'����б�:
'������         ����                    ˵��
'lngCurPID      Long                    ��ǰ����ID
'lngVBMain      Long                    VB������
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
'����           GetAllChild
'����           ��ȡ���е�VB�Ӵ��壬�����Ӵ��ڵĴ��ڡ�
'����ֵ         Variant                 ���ش������飬Ԫ��[Hwnd,OwnerHwnd,IsEnable,class,Text,Module]
'����б�:
'������         ����                    ˵��
'lngFrmHwnd     Long                    ö�ٵĴ���
'blnNormalWin   Boolean                 �Ƿ��ǳ��洰��
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
'����           GetAllChildSub
'����           GetAllChildSub�ĵݹ鷽��
'����ֵ
'����б�:
'������         ����                    ˵��
'lngParentHwnd  Long                    ������
'arrRet         Variant                 �Ѿ����ڵĴ������ݣ�����׷��
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

'����           IsTopModal
'����           �жϵ�ǰ���̵���㴰���Ƿ���ģ̬���塣��VB�����á�ͨ���жϵ���̨�ǽ��ã����㴰�������ã���ǰ��ģ̬
'����ֵ         Boolean
'����б�:
'������         ����                    ˵��
'lngBrwHwnd     Long                    ����̨���
'lngTopWin      Long                    ��㴰����
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
'����           IsCanDoMessage
'����           �жϴ����ܷ������Ϣ
'����ֵ         Boolean
'����б�:
'������         ����                    ˵��
'lngTopWin      Long                    Ҫ�жϵĴ���
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
'����           FindVBMain
'����           ��ȡVB���̵�������
'����ֵ         Long
'����б�:
'������         ����                    ˵��
'lngBrwHwnd     Long                    ����̨���
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
'����           FindExitButtonNormal
'����           Ѱ��һ�����洰������ȷ����ť��
'����ֵ         Long                    ����ȷ����ȡ����ť���
'����б�:
'������         ����                    ˵��
'lngMainFrm     Long                    ���ھ��
'lngToolBarHwnd Long                    ToolBar�ľ���������ڴ���TOOLBar,�Ҵ���ȷ��ȡ����ť����ʱ��������ȷ��ȡ����ťID��
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
            If strWinText = "ȷ��" Or strWinText Like "ȷ��(*)" Or strWinText = "����" Or strWinText Like "����(*)" Then
                gobjLog.LogInfo RLL_LogInfo, "FindExitButtonNormal", "OKBtn =" & arrChild(i)(WI_Hwnd), "Parent=" & arrChild(i)(WI_OwnerHwnd), "IsEnable=" & arrChild(i)(WI_Enable), "Class=" & arrChild(i)(WI_Class), "Text=" & arrChild(i)(WI_Text)
                lngOKBtn = arrChild(i)(WI_Hwnd)
                If lngCancelBtn <> 0 Then Exit For
            ElseIf strWinText = "�˳�" Or strWinText Like "�˳�(*)" Or strWinText = "�ر�" Or strWinText Like "�ر�(*)" Or strWinText = "ȡ��" Or strWinText Like "ȡ��(*)" Then
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
'����           FindExitButtonMsgBox
'����           Ѱ��һ��Msgbox��������ȷ����ť��
'����ֵ         Long                    ����һ����ť���
'����б�:
'������         ����                    ˵��
'lngMainFrm     Long                    MsgBox�Ĵ��ھ��
'strMsg         String                  MsgBox�Ĵ�����ʾ
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
            If strWinText Like "ȷ��*" Or strWinText Like "��*" Then
                gobjLog.LogInfo RLL_LogInfo, "FindExitButtonMsgBox", "OKBtn =" & arrChild(i)(WI_Hwnd), "Parent=" & arrChild(i)(WI_OwnerHwnd), "IsEnable=" & arrChild(i)(WI_Enable), "Class=" & arrChild(i)(WI_Class), "Text=" & arrChild(i)(WI_Text)
                lngOKBtn = arrChild(i)(WI_Hwnd)
            ElseIf strWinText Like "��*" Or strWinText Like "ȡ��*" Or strWinText Like "����*" Then
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
            strMsg = strMsg & vbNewLine & "����Ϣ��ʾ��" & strErr
        Else
            strMsg = "����Ϣ��ʾ��" & strErr
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
'����           FindExitButtonErrCenter
'����           Ѱ��һ�����������Ĵ�������ȷ����ť��
'����ֵ         Long                    ����ȷ����ť���
'����б�:
'������         ����                    ˵��
'lngMainFrm     Long                    ErrCenter�Ĵ��ھ��
'strMsg         String                  ErrCenter�Ĵ�����ʾ
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
            If strWinText Like "ȷ��(*)" Then
                gobjLog.LogInfo RLL_LogInfo, "FindExitButtonErrCenter", "OKBtn =" & arrChild(i)(WI_Hwnd), "Parent=" & arrChild(i)(WI_OwnerHwnd), "IsEnable=" & arrChild(i)(WI_Enable), "Class=" & arrChild(i)(WI_Class), "Text=" & arrChild(i)(WI_Text)
                lngOKBtn = arrChild(i)(WI_Hwnd)
            ElseIf strWinText Like "ȡ��(*)" Then
                gobjLog.LogInfo RLL_LogInfo, "FindExitButtonErrCenter", "CancelBtn =" & arrChild(i)(WI_Hwnd), "Parent=" & arrChild(i)(WI_OwnerHwnd), "IsEnable=" & arrChild(i)(WI_Enable), "Class=" & arrChild(i)(WI_Class), "Text=" & arrChild(i)(WI_Text)
                lngCancelBtn = arrChild(i)(WI_Hwnd)
            End If
        ElseIf strClass = mstrVBControlHeader & "TextBox" Then
            If Trim(strWinText) <> "" Then
                strErrText = strErrText & Trim(strWinText)
                gobjLog.LogInfo RLL_LogInfo, "FindExitButtonErrCenter", "InfoHwnd =" & arrChild(i)(WI_Hwnd), "Parent=" & arrChild(i)(WI_OwnerHwnd), "IsEnable=" & arrChild(i)(WI_Enable), "Class=" & arrChild(i)(WI_Class), "Text=" & arrChild(i)(WI_Text)
            End If
        ElseIf strClass = "Static" Then
            If strWinText <> "˵����" And Not strWinText Like "��ţ�*" Or strWinText <> "�����ţ�" And Not strWinText Like "*����Զ���ͼ���رմ˽���" Or strWinText <> "����ʱ" Then
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
            strMsg = strMsg & vbNewLine & "��������" & strError
        Else
            strMsg = "��������" & strError
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
'����           GetWinModule
'����           ��ȡ�����Ӧ�����ļ�ģ�飬ֻ�����ϲ�ģ����Ч
'����ֵ         String
'����б�:
'������         ����                    ˵��
'lngHwnd        Long                    ���ھ��
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
'����           GetWinClass
'����           ��ȡ��������
'����ֵ         String
'����б�:
'������         ����                    ˵��
'lngHwnd        Long                    ���ھ��
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
'����           GetWinText
'����           ��ȡ�����ı�����
'����ֵ         String
'����б�:
'������         ����                    ˵��
'lngHwnd        Long                    ���ھ��
'-------------------------------------------------------------------------------------------------
Private Function GetWinText(ByVal lngHwnd As Long) As String  '֧���Ǻ�����
    Dim lngLen       As Long
    Dim strTmp       As String * 1024
    If lngHwnd <> 0 Then
        Call GetWindowText(lngHwnd, strTmp, Len(strTmp) - 1)
        GetWinText = Left(strTmp, InStr(strTmp, vbNullChar) - 1)
    End If
End Function

'--------------------------------------------------------------------------------------------------
'����           ClickButton
'����           �����ť
'����ֵ
'����б�:
'������         ����                    ˵��
'lngMainFrm     Long                    ��ť���ڵĴ���
'lngBtnHwnd     Long                    ��ť���
'-------------------------------------------------------------------------------------------------
Public Sub ClickButton(ByVal lngMainFrm As Long, ByVal lngBtnHwnd As Long)
'    Dim vRect       As RECT, pCenter As POINTAPI
    Dim lngResult       As Long
    Dim lngReturnValue  As Long
    
    
    On Error GoTo ErrH
    Call gobjLog.PushMethod(RLL_AllLog, "ZlWin.mdlAutoClick.ClickButton", lngMainFrm, lngBtnHwnd)
    '��ֹ��δ���Activeʱ��
    If GetForegroundWindow() <> lngMainFrm Then
        Call SetForegroundWindow(lngMainFrm)
    End If
    If GetActiveWindow() <> lngMainFrm Then
        Call SetActiveWindow(lngMainFrm)
    End If
    Call SetFocus(lngBtnHwnd)
    '������ť����1
'    Call GetWindowRect(lngBtnHwnd, vRect)
'    pCenter.x = (vRect.Left + vRect.Right) / 2
'    pCenter.Y = (vRect.ToP + vRect.Bottom) / 2
'    pCenter.x = pCenter.x / (Screen.Width / Screen.TwipsPerPixelX) * 65535
'    pCenter.Y = pCenter.Y / (Screen.Height / Screen.TwipsPerPixelY) * 65535
'    mouse_event MOUSEEVENTF_ABSOLUTE Or MOUSEEVENTF_MOVE, pCenter.x, pCenter.Y, 0, 0
'    mouse_event MOUSEEVENTF_ABSOLUTE Or MOUSEEVENTF_LEFTDOWN, pCenter.x, pCenter.Y, 0, 0
'    mouse_event MOUSEEVENTF_ABSOLUTE Or MOUSEEVENTF_LEFTUP, pCenter.x, pCenter.Y, 0, 0
    '������ť����2
'    PostMessage lngBtnHwnd, WM_LBUTTONDOWN, 0, ByVal 0 '����ڰ�ť����
'    PostMessage lngBtnHwnd, WM_LBUTTONUP, 0, ByVal 0 '����ڰ�ť����
    '������ť����3
'    PostMessage lngBtnHwnd, BM_CLICK, 0, 0 '������ť
    lngReturnValue = SendMessageTimeout(lngBtnHwnd, BM_CLICK, 0, ByVal 0, SMTO_ABORTIFHUNG Or SMTO_BLOCK, 1000, lngResult)
    Call gobjLog.PopMethod(RLL_AllLog, "ZlWin.mdlAutoClick.ClickButton")
    Exit Sub
ErrH:
    If gobjLog.ErrCenter(RLL_RunError, "ZlWin.mdlAutoClick.ClickButton") = 1 Then
        Resume
    End If
End Sub

'--------------------------------------------------------------------------------------------------
'����           ClickMenuExit
'����           ������ڵĹرհ�ť
'����ֵ
'����б�:
'������         ����                    ˵��
'lngMainFrm     Long                    ��ť���ڵĴ���
'-------------------------------------------------------------------------------------------------
Public Sub ClickMenuExit(ByVal lngMainFrm As Long)
    Dim lngResult       As Long
    Dim lngReturnValue  As Long
    
    On Error GoTo ErrH
    Call gobjLog.PushMethod(RLL_AllLog, "ZlWin.mdlAutoClick.ClickMenuExit", lngMainFrm)
    '��ֹ��δ���Activeʱ��
    If GetForegroundWindow() <> lngMainFrm Then
        Call SetForegroundWindow(lngMainFrm)
    End If
    If GetActiveWindow() <> lngMainFrm Then
        Call SetActiveWindow(lngMainFrm)
    End If
    lngReturnValue = SendMessageTimeout(lngMainFrm, WM_CLOSE, 0, ByVal 0, SMTO_ABORTIFHUNG Or SMTO_BLOCK, 1000, lngResult)
    If lngReturnValue <> 0 Then
        If lngResult = 0 Then
            gobjLog.LogInfo RLL_LogInfo, "ClickMenuExit", "������Ϣ����ɹ�=" & lngMainFrm
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
'����           GetTickCountDiff
'����           ����GetTickCcout�Ĳ�ֵ������ GetTickCountVB�������ֵ�Լ��������������Ҫ��������
'����ֵ         Double
'����б�:
'������         ����                    ˵��
'lngStart       Long                    ��ʼʱ��
'-------------------------------------------------------------------------------------------------
Public Function GetTickCountDiff(ByVal lngStart As Long) As Double
    Dim lngCur          As Long
    Const M_OFFSET_4    As Double = 4294967296#         '�޷������ε����ֵ
    
    lngCur = GetTickCount
    If lngCur < lngStart Then
        GetTickCountDiff = M_OFFSET_4 - LongToUnsigned(lngStart) + LongToUnsigned(lngCur)
    Else
        GetTickCountDiff = lngCur - lngStart
    End If
End Function

Private Function LongToUnsigned(value As Long) As Double
    Const M_OFFSET_4    As Double = 4294967296#         '�޷������ε����ֵ
    If value < 0 Then LongToUnsigned = value + M_OFFSET_4 Else LongToUnsigned = value
End Function

Public Sub ClearObjects(ByVal blnTest As Boolean, Optional ByVal objNow As Object)
    Dim lngCount    As Long
    '����ǰδ����Ķ���
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
'���ܣ� ȷ����ǰģʽΪ���ģʽ
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
'���ܣ�ȥ���ַ�����\0�Ժ���ַ�
    Dim lngPos As Long

    lngPos = InStr(strInput, Chr(0))
    If lngPos > 0 Then
        TruncZero = Mid(strInput, 1, lngPos - 1)
    Else
        TruncZero = strInput
    End If
End Function

Public Function ActualLen(ByVal strAsk As String) As Long
'���ܣ���ȡָ���ַ�����ʵ�ʳ��ȣ������ж�ʵ�ʰ���˫�ֽ��ַ�����
'       ʵ�����ݴ洢����
    ActualLen = LenB(StrConv(strAsk, vbFromUnicode))
End Function

Public Function SubB(ByVal strInfor As String, ByVal lngStart As Long, ByVal lngLen As Long) As String
'����:��ȡָ���ִ���ֵ,�ִ��п��԰�������
 '���:strInfor-ԭ��
 '         lngStart-ֱʼλ��
'         lngLen-����
'����:�Ӵ�
    Err = 0: On Error GoTo ErrH:
    SubB = StrConv(MidB(StrConv(strInfor, vbFromUnicode), lngStart, lngLen), vbUnicode)
    SubB = Replace(SubB, Chr(0), "")
    Exit Function
ErrH:
    Err.Clear
    SubB = ""
End Function
