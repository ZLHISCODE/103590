Attribute VB_Name = "MdlBrower"
Option Explicit
'MDI����
Public Type Menu_Type
    ���ܲ˵� As Long
    ���ڲ˵� As Long
    �������ܲ˵� As Long
    �ָ��˵� As Long
End Type
Public �˵���׼ As Menu_Type
Public Enum �����嵥
    ���������嵥 = 10
    �ֵ������ = 11
    ��Ϣ�շ����� = 12
    ϵͳѡ������ = 13
    EXCEL������ = 14
    ���ز������� = 15
End Enum
'��ҹ���
Public gobjPlugIn As Object

Public gobjRelogin As Object                   '���������
Public FrmMainface As Form
Public gcnOracle As ADODB.Connection

Public gstrSysName As String                'ϵͳ����
Public gstrVersion As String                'ϵͳ�汾
Public gstrAviPath As String                'AVI�ļ��Ĵ��Ŀ¼
Public gstrUserFlag As String               '��ǰ�û���־(��λ��ʾ)����1λ���Ƿ�DBA����2λ��ϵͳ������
Public gstrDbUser As String                 '��ǰ���ݿ��û�
Public glngUserId As Long                   '��ǰ�û�id
Public gstrUserCode As String               '��ǰ�û�����
Public gstrUserName As String               '��ǰ�û�����
Public gstrUserAbbr As String               '��ǰ�û�����
Public glngDeptId As Long                   '��ǰ�û�����id
Public gstrDeptCode As String               '��ǰ�û����ű���
Public gstrDeptName As String               '��ǰ�û���������
Public gstrStation As String                '������վ����
Public gblnShutDown As Boolean              '������������Ƿ�����˳�����̨

Public gstrObj() As String
Public gobjCls() As Object
Public grsMenus As New ADODB.Recordset       '�˵���¼��
Public gstrMenuSys As String                '�˵�����
Public gstrCommand As String                '�����в��� �¶� 2010-12-06
Private mlngSysPre As Long                  '�ϴε���˽��ͬ��ʼ�鴴��ʱ��ϵͳ��

Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
'�ر�ϵͳ��صı�����API����
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
'����ExitWindowsEx
Private Const mlng�رռ��������Դ As Long = 8
Public Const EWX_FORCE = 4 'ǿ�йرճ���ע��
Public Const WINDOWS95 = 0
Public Const WINDOWSNT = 1
Private mlngWin32 As Long
Private mblnע�� As Boolean
'����������ڼ���Ƿ�Ϸ�����
Public Declare Function GlobalAddAtom Lib "kernel32" Alias "GlobalAddAtomA" (ByVal lpString As String) As Integer
Public Declare Function GlobalDeleteAtom Lib "kernel32" (ByVal nAtom As Integer) As Integer
'�ڴ��ڽṹ��Ϊָ���Ĵ���������Ϣ
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
'��ָ�����ڵĽṹ��ȡ����Ϣ
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
'����ֵ�ǲ˵��ľ������������Ĵ���û�в˵����򷵻�NULL�����������һ���Ӵ��ڣ�����ֵ�޶���
Public Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
'����ָ���Ľ���
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'��ϵͳע��һ��ָ�����ȼ�
Public Declare Function RegisterHotKey Lib "user32" (ByVal hwnd As Long, ByVal id As Long, ByVal fsModifiers As Long, ByVal vk As Long) As Long
'ȡ���ȼ����ͷ�ռ�õ���Դ
Public Declare Function UnregisterHotKey Lib "user32" (ByVal hwnd As Long, ByVal id As Long) As Long
  '�ȼ���־����,�����жϵ����̰���������ʱ�Ƿ������������趨���ȼ�
Public Const WM_HOTKEY = &H312
Public Const MOD_ALT = &H1
Public Const MOD_CONTROL = &H2
Public Const MOD_SHIFT = &H4
Public Const GWL_WNDPROC = (-4)    '���ں����ĵ�ַ
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
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Const VK_LBUTTON = &H1
'����ʵ����������ϵͳ�ķ�Χ�ڹ�����
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Long
'��ȡ��ǰ����һ��Ψһ�ı�ʶ��
Public Declare Function GetCurrentProcessId Lib "kernel32" () As Long
'�ҳ�ĳ�����ڵĴ�����(�̻߳����)�����ش����ߵı�־����
Public Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
'�ú���ȷ�������Ĵ��ھ���Ƿ��ʶһ���Ѵ��ڵĴ��ڡ�
'����ֵ��������ھ����ʶ��һ���Ѵ��ڵĴ��ڣ�����ֵΪ���㣻������ھ��δ��ʶһ���Ѵ��ڴ��ڣ�����ֵΪ��
Public Declare Function isWindow Lib "user32" Alias "IsWindow" (ByVal hwnd As Long) As Long
'����ֵ�����ָ���Ĵ��ڼ��丸���ھ���WS_VISIBLE��񣬷���ֵΪ���㣻���ָ���Ĵ��ڼ��丸���ڲ�����WS_VISIBLE��񣬷���ֵΪ�㡣���ڷ���ֵ�����˴����Ƿ����Ws_VISIBLE�����ˣ���ʹ�ô��ڱ����������ڸǣ���������ֵҲΪ���㡣
Public Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long
'�ú���ȷ�����������Ƿ�����С��(ͼ�껯)�Ĵ��ڡ�
Public Declare Function IsIconic Lib "user32" (ByVal hwnd As Long) As Long
'�ú���ȷ�����������Ƿ�����󻯵Ĵ��ڡ�
Public Declare Function IsZoomed Lib "user32" (ByVal hwnd As Long) As Long
' �ú���ö��������Ļ�ϵĶ��㴰�ڣ��������ھ�����͸�Ӧ�ó�����Ļص�����
Public Declare Function EnumWindows Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
'��ȡ��ǰ���̵Ļ���壬����ǰ����û�м���򷵻�0
Public Declare Function GetActiveWindow Lib "user32" () As Long
'�ú������һ��ָ���Ӵ��ڵĸ����ھ����
Public Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
'�ú������ָ�������������������?
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
'ȡ��һ������ı��⣨caption�����֣�����һ���ؼ������ݣ���vb��ʹ�ã�ʹ��vb�����ؼ���caption��text���ԣ�
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long

'����Z����
'��������һϵ���´���λ�õĹ��̣��Ա�ͬʱ���£����ú�������һ���ڲ��������һ�����������ṹ�������봰��λ���йص���Ϣ����󣬸ýṹ���ɶ�DeferWindowPos�����ĵ�����䡣׼���ø������д���λ���Ժ󣬶�EndDeferWindowPos��һ�����ÿ�ͬʱ���½ṹ�����д��ڵ�λ��
Private Declare Function BeginDeferWindowPos Lib "user32" (ByVal nNumWindows As Long) As Long
'�ú���Ϊ�ض��Ĵ���ָ��һ���´���λ�ã�������������BeginDeferWindowPos�����Ľṹ���Ա���EndDeferWindowPos����ִ���ڼ����
'����һ���¾������ָ��Ľṹ������λ�ø�����Ϣ��������Ӧ�ڶ�DeferWindowPos�ĺ��������Լ���EndDeferWindowPos�Ľ����������õ���������򷵻���ֵ
Private Declare Function DeferWindowPos Lib "user32" (ByVal hWinPosInfo As Long, ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
'ͬʱ����DeferWindowPos����ʱָ�������д��ڵ�λ�ü�״̬
Private Declare Function EndDeferWindowPos Lib "user32" (ByVal hWinPosInfo As Long) As Long

Private Type LASTINPUTINFO
    cbSize As Long
    dwTime As Long
End Type
'��ȡ�ϴ����������ʱ�䡣
Private Declare Function GetLastInputInfo Lib "user32" (plii As LASTINPUTINFO) As Boolean
'���أ�retrieve���Ӳ���ϵͳ������������elapsed���ĺ�����
Private Declare Function GetTickCount Lib "kernel32" () As Long

Private Const GW_CHILD = &H5
Public Const GW_OWNER = &H4
'���һ�����ڵľ�����ô�����ĳԴ�������ض��Ĺ�ϵ
Public Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
'�ú��������ж�ָ���Ĵ����Ƿ�������ܼ��̻�������롣
Private Declare Function IsWindowEnabled Lib "user32" (ByVal hwnd As Long) As Long
'ǿ���������´��ڣ���������ǰ���ε��������򶼻��ػ�����vb��ʹ�ã���vb�����ؼ����κβ�����Ҫ���£��ɿ���ֱ��ʹ��refresh����
Private Declare Function UpdateWindow Lib "user32" (ByVal hwnd As Long) As Long
'lpRect:ָ��һ��RECT�ṹ��ָ�룬�ýṹ���մ��ڵ����ϽǺ����½ǵ���Ļ���ꡣ
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
'��z���е�λ�ڱ���λ�Ĵ���ǰ�Ĵ��ھ��
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Private mlngPid As Long '��ǰ����
Private mcllHideFrms As Collection '���������صĴ���
Private mcllHideFrmsEx As Collection '���������ص��ޱ���������
Public gobjButton As frmButton  '�������ڶ���
Public gobjLock As frmLock '��������
Public grecButton As RECT '��������ť��������������
Private glngPreHwnd As Long 'ǰһ���ǰ�ť�Ļ����
Private gblnPreZoomed As Boolean 'ǰһ������ͼ�궨λ�����Ƿ������
Public gblnWin10 As Boolean '�Ƿ�汾ΪWIn10
Public gintCurTheme As Integer '������0-��������,1-AERO���⣬WIn8,WIN 10����,2-BASIC����
Public gblnHideBtn As Boolean '������ť�Ƿ�����״̬
Public glngLockTime As Long
Public glngMain As Long 'ThunderRT6Main����
Public gblnLock As Boolean '�Ƿ�������״̬
'�ܵ���ȡCMD���
'��������һ���µĽ��̺��������̣߳�����½�������ָ���Ŀ�ִ���ļ����������ִ�гɹ������ط���ֵ
Private Declare Function CreateProcess Lib "kernel32" Alias "CreateProcessA" (ByVal lpApplicationName As String, ByVal lpCommandLine As String, lpProcessAttributes As Any, lpThreadAttributes As Any, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, lpEnvironment As Any, ByVal lpCurrentDriectory As String, lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long
'����һ�������ܵ��������еõ���д�ܵ��ľ�����������ִ�гɹ������ط���ֵ
Private Declare Function CreatePipe Lib "kernel32" (phReadPipe As Long, phWritePipe As Long, lpPipeAttributes As SECURITY_ATTRIBUTES, ByVal nSize As Long) As Long
'���ļ�ָ��ָ���λ�ÿ�ʼ�����ݶ�����һ���ļ��У� ��֧��ͬ�����첽����
Private Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, lpOverlapped As Any) As Long
'���ȴ����ڹ���״̬ʱ��������رգ���ô������Ϊ��δ����ġ��þ��������� SYNCHRONIZE ����Ȩ�ޡ�
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
Public Const WH_KEYBOARD As Integer = 2      '��ͨ���̹���
Public glngHook As Long                      '������Ϣ���

Public Sub ExecuteFunc(lngSys As Long, Components As String, Modul As Long, Optional ByVal strPara As String) ', Identity As Byte
    '-------------------------------------------------------------
    '���ܣ�����ִ��ָ�������Ĺ��ܳ���
    '������
    '   frmbrower��������
    '   Components������
    '   Modul��ģ����
    '   Identity����ִ�������Ҫ��
    '-------------------------------------------------------------
    Dim rsCheck As New ADODB.Recordset                  '���汾�Ƿ����ϵͳ����
    Dim IntCount As Integer, intClients As Integer
    Dim objNow As Object                                '�����Ĳ�������
    Dim BlnExecute As Boolean                           '�Ƿ���ڸò���
    Dim StrVersion As String, StrCompareVersion As String
    Dim ArrayVersion
    '�Ϸ��Լ��
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
    
    'ʹ���²�������
    If UCase(Components) = UCase("zl9EmrInterface") And BlnExecute = False Then
        IntCount = UBound(gstrObj)
        IntCount = IntCount + 1
        ReDim Preserve gstrObj(IntCount)
        gstrObj(IntCount) = Components
        If FrmMainface.mobjEmr Is Nothing Then
            MsgBox "�����������ʧ�ܣ����鲢���µ�¼��", vbInformation, gstrSysName
            Exit Sub
        ElseIf FrmMainface.mobjEmr.IsInited = False Then
            MsgBox "�������δ�ܳ�ʼ��," & FrmMainface.mobjEmr.GetError(), vbInformation, gstrSysName
            Exit Sub
        End If
        If Not gobjRelogin.IsEMRProxy Then 'ʹ�ô����û���¼���򲻼��Ȩ��
            Dim strSpecify As String 'Ƭ�Σ�����Ȩ�޹̶��ڵ���ǰ����
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
    '--���û�иò���,�򴴽�--
    If BlnExecute = False Then
        Set objNow = CreateObject(Components & ".Cls" & Mid(Components, 4))
    
        If Err = 0 Then
            On Error GoTo errH
            '--���ò����İ汾�Ƿ�����ϵͳ����(���汾-3;�ΰ汾-3;���汾-3)[�Զ��屨��������]--
            If Not (UCase(Components) = "ZL9REPORT") And Not (UCase(Components) = "ZL9DOC") And Not OS.IsDesinMode Then
                strSQL = " Select nvl(���汾,1) ���汾,nvl(�ΰ汾,0) �ΰ汾,nvl(���汾,0) ���汾,���� " & _
                          " From ZlComponent Where Upper(Rtrim(����))=[1] And ϵͳ=[2]"
                Set rsCheck = zlDatabase.OpenSQLRecord(strSQL, "�����汾���", UCase(Components), lngSys)
                With rsCheck
                    If .EOF Then
                        MsgBox "ϵͳ������ZlComponent���ݲ����������������������ϵ��", vbInformation, gstrSysName
                        FrmMainface.MousePointer = 0
                        Exit Sub
                    End If
                    
                    '��װ�汾��Ϊ��λ���汾����λ�ΰ汾����λ���汾
                    StrCompareVersion = String(3 - Len(!���汾), "0") & !���汾 & "." & _
                                        String(3 - Len(!�ΰ汾), "0") & !�ΰ汾 & "." & _
                                        String(3 - Len(!���汾), "0") & !���汾
                    ArrayVersion = Split(objNow.Version, ".")
                    StrVersion = String(3 - Len(ArrayVersion(0)), "0") & ArrayVersion(0) & "." & _
                                 String(3 - Len(ArrayVersion(1)), "0") & ArrayVersion(1) & "." & _
                                 String(3 - Len(ArrayVersion(2)), "0") & ArrayVersion(2)
                    
                    If StrVersion < StrCompareVersion Then
                        MsgBox "�ò����İ汾�Ѳ�������ϵͳ���������������������ϵ����" & !���� & "��", vbInformation, gstrSysName
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
        '��������ʧ�ܣ�Ӧ����ʾ
        Else
            Screen.MousePointer = 0
            MsgBox "���� " & Components & ".Cls" & Mid(Components, 4) & " �����������������鰲װ�Ƿ���ȷ����Ϣ��" & vbNewLine & Err.Description, vbExclamation, gstrSysName
            Err.Clear
            Exit Sub
        End If
    End If
    
    Err = 0: On Error GoTo errH
    '--ִ�иù���--
    If gstrObj(IntCount) = Components Then
        If UCase(Components) = "ZL9REPORT" Then
            If Modul = �˵���׼.�������ܲ˵� Then
                gobjCls(IntCount).ReportMan gcnOracle, FrmMainface
            Else
                
'                strPara = "��ʼ����=2013-01-01"
                If strPara <> "" Then
                    Dim varPara As Variant
                                        
                    varPara = Split(strPara, "|")
'                    varPara(0) = "��ʼ����=2013-01-01"
'                    varPara(1) = "��������=2014-05-01"
                    
                    '���֧��10������������10���Ĳ���
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
            
            strSQL = " Select ���⡡From zlPrograms Where ���=[1] "
            Set rsCheck = zlDatabase.OpenSQLRecord(strSQL, "ϵͳģ����", Modul)
            With rsCheck
                    If .EOF Then
                        MsgBox "ϵͳ�����ݲ����������������������ϵ��", vbInformation, gstrSysName
                        FrmMainface.MousePointer = 0
                        Exit Sub
                    Else
                        strModul = !����
                    End If
            End With
            strFuncs = GetPrivFunc(lngSys, Modul)
            Call FrmMainface.mobjEmr.CodeMain(Modul, strModul, FrmMainface.hwnd, gobjRelogin.EMRUser, gobjRelogin.EMRPwd, strFuncs)
        Else
            Call CreateSynonyms(lngSys, Modul)
            
            '�û�վ������� (��ʽ�漰���ð�)
            intClients = Val(zlRegInfo("��Ȩվ��"))
            If intClients > 0 Then
                If GetCurStates > intClients Then
                    MsgBox "��ǰ�û���¼�������������Ȩ��" & intClients & ",ϵͳ���Զ��������У�", vbInformation, gstrSysName
                    Exit Sub
                End If
            End If

            
            'ΪͨѶԭ�Ӹ�ֵ
            strCommon = Format(Now, "yyyyMMddHHmm")
            strCommon = TranPasswd(strCommon) & "||" & OS.ComputerName
            '����ͨѶԭ��
            intAtom = GlobalAddAtom(strCommon)
            Call SaveSetting("ZLSOFT", "����ȫ��", "����", intAtom)
            gobjCls(IntCount).CodeMan lngSys, Modul, gcnOracle, FrmMainface, gstrDbUser
            Call GlobalDeleteAtom(intAtom)
            
            '��ҽ������ֻ��CodeMan()���ܻ�ȡϵͳ�ţ��ڶ�ȡ����ʱ����֪��ϵͳ�ţ���д��ע������ҽ��������Ĭ��Ϊ 100
            Call SaveSetting("ZLSOFT", "����ȫ��", "ϵͳ��", lngSys)
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
    '����:���������¼
    mblnע�� = True
    
    Call gobjRelogin.ReLogin(FrmMainface)
End Sub

Public Function OwnerUser(ByVal strUserName As String) As Boolean
    Dim RecUser As New ADODB.Recordset
    Dim strSQL As String
    OwnerUser = True
    On Error GoTo errH
'        If .State = 1 Then .Close
        strSQL = "Select Count(*) ������ From ZlSystems Where ������='" & strUserName & "'"
         Set RecUser = zlDatabase.OpenSQLRecord(strSQL, "������")
'        .Open "Select Count(*) ������ From ZlSystems Where ������='" & strUserName & "'", gcnOracle By zq
        
        If RecUser.EOF Then
            If Not IsNull(RecUser!������) Then
                If RecUser!������ = 0 Then OwnerUser = False
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
    '����ģ����������ͬ���(����Ѵ����򲻻��ٴ���)
    On Error Resume Next
    If mlngSysPre <> lngSys Then
        strSQL = "Zl_Createsynonyms(" & lngSys & ")"
        zlDatabase.ExecuteProcedure strSQL, "����ͬ���"
        mlngSysPre = lngSys
        If Err.Number <> 0 Then Err.Clear
    End If
End Function

Public Sub AddHistory(ByVal strModul As String)
    Dim strϵͳ As String, str��� As String, intMax As Integer
    Dim arrϵͳ As Variant, arr��� As Variant, strValue As String
    Dim intϵͳ_Cur As Integer, int���_Cur As Integer
    Dim intϵͳ_Max As Integer, int���_Max As Integer
    '������еĳ���ʼ���ڵ�һ��λ�ã�����Ѵ�������ʷ��¼�У��������ڵ�һ��λ��
    'strModul:ϵͳ & "," & ģ��
    
    intMax = 6
    
    strValue = zlDatabase.GetPara("���ʹ��ģ��")
    If UBound(Split(strValue, "|")) >= 1 Then
        strϵͳ = Trim(Split(strValue, "|")(0))
        str��� = Trim(Split(strValue, "|")(1))
    End If
    If strϵͳ = "" Or str��� = "" Then
        strϵͳ = Split(strModul, ",")(0)
        str��� = Split(strModul, ",")(1)
        Call zlDatabase.SetPara("���ʹ��ģ��", strϵͳ & "|" & str���)
        Exit Sub
    End If
    
    arrϵͳ = Split(strϵͳ, ",")
    arr��� = Split(str���, ",")
    intϵͳ_Max = UBound(arrϵͳ)
    int���_Max = UBound(arr���)
    strϵͳ = Split(strModul, ",")(0): str��� = Split(strModul, ",")(1)
    If intϵͳ_Max > intMax Then intϵͳ_Max = intMax
    
    For intϵͳ_Cur = 0 To intϵͳ_Max
        int���_Cur = intϵͳ_Cur
        If int���_Cur > int���_Max Then Exit For
        If Not (arrϵͳ(intϵͳ_Cur) = Split(strModul, ",")(0) And arr���(int���_Cur) = Split(strModul, ",")(1)) Then
            strϵͳ = strϵͳ & "," & arrϵͳ(intϵͳ_Cur)
            str��� = str��� & "," & arr���(int���_Cur)
        End If
    Next
    Call zlDatabase.SetPara("���ʹ��ģ��", strϵͳ & "|" & str���)
End Sub

Public Sub CheckWinVersion()
    Dim lngVersion As Long
    
    mblnע�� = False
    lngVersion = GetVersion()
    If ((lngVersion And &H80000000) = 0) Then
        mlngWin32 = WINDOWSNT
    Else
        mlngWin32 = WINDOWS95
    End If
End Sub

Public Sub ShutDown(ByVal blnCloseWin As Boolean)
    If mblnע�� Then Exit Sub
    If Not blnCloseWin Then Exit Sub
    If mlngWin32 = WINDOWSNT Then
        Call AdjustToken
        Call ExitWindowsEx(mlng�رռ��������Դ Or EWX_FORCE, 0)
    Else
        Call ExitWindowsEx(mlng�رռ��������Դ Or EWX_FORCE, 0)
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
    
    '�õ���ǰ���̵ľ��
    hdlProcessHandle = GetCurrentProcess()
    If GetLastError <> 0 Then Exit Function
    
    '�õ���ǰ���̵�Ȩ�޾��
    OpenProcessToken hdlProcessHandle, (TOKEN_ADJUST_PRIVILEGES Or TOKEN_QUERY), hdlTokenHandle
    If GetLastError <> 0 Then Exit Function
     
    '�ҵ��ر�Ȩ�޲�����LUID
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
    '���ܣ� ����ת������
    '������
    '   strOld��ԭ����
    '���أ� �������ɵ�����
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
'����ָ������,����ʱ������������չʾ
    On Error Resume Next
    '�ָ�ǰһ����Ч�İ�ť��������״̬
    Call ShowWindow(lnghWnd, IIf(Not blnHide, IIf(gblnPreZoomed And lnghWnd = glngPreHwnd, SW_SHOWMAXIMIZED, SW_SHOW), SW_HIDE))
    If Err.Number <> 0 Then Err.Clear
End Sub

Public Function EnumWindowsProc(ByVal hwnd As Long, ByVal lParam As Long) As Long
'���ܣ� ͨ��PIDö�������ľ��,������Ҫ�Ĵ���
    Dim lngPid As Long
    Dim strText As String * 255
    GetWindowThreadProcessId hwnd, lngPid
    If mlngPid = lngPid Then
        If isWindow(hwnd) <> 0 Then
            If IsWindowVisible(hwnd) <> 0 Then
                If isNormalWindow(hwnd) Then
                    mcllHideFrms.Add hwnd
                Else
                    If (IsWindowEnabled(GetWindow(hwnd, GW_OWNER)) = 0) Then '�ޱ�������ģ̬������Ҫ����
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
        '��ȡ���еĿɼ�����
        Call GetAllVisibleWindow(GetCurrentProcessId)
        '��ȡǰһ�������Ƿ���󻯣���Ϊ��������ʼ������󻯣�����������������Ļ��ͼ������Ҫ�Ȼָ�����󻯡�
        glngPreHwnd = GetActiveWindow
        gblnPreZoomed = False
        If glngPreHwnd <> frmBrower.hwnd Then
            If isWindow(glngPreHwnd) <> 0 Then
                If GetMenu(glngPreHwnd) <> 0 Then '���ڴ����Դ��˵��������TOOLBar
                    If IsZoomed(glngPreHwnd) <> 0 Then
                        gblnPreZoomed = True
                        Call ShowWindow(glngPreHwnd, SW_RESTORE)
                    End If
                End If
            End If
        End If
    End If
    '�����д�������
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
'�ų�����ؼ��Ĵ������
'�ų�DTPicker����������ѡ����棬�ý���ͨ��API�ж����б������ģ�ͨ��SPY++������û�еģ���ʱ�ų��ô���
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
'���ܣ�������Ƶ��ťλ��
    Dim lngcxBut   As Long, lngcyBut   As Long
    Dim uRect    As RECT
    Dim lngButSize    As Long, lngSysButSize As Long
    Dim lngRightEdgeOffset As Long
    Dim lngStyle      As Long, lngExStyle   As Long
    
    '��ȡ������ʽ
    lngStyle = GetWindowLong(lnghWnd, GWL_STYLE)
    lngExStyle = GetWindowLong(lnghWnd, GWL_EXSTYLE)
    '��ȡ�ұ߰�ťλ��ƫ�ƣ����ұߵ���󻯣���С�����رհ�ť
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
        'ϵͳ�˵���ť
        If (lngStyle And WS_SYSMENU) Then
            lngButSize = lngButSize + lngSysButSize + B_EDGE
        End If
        '�����С����ť
        If (lngStyle And (WS_MINIMIZEBOX Or WS_MAXIMIZEBOX)) Then
            lngButSize = lngButSize + B_EDGE + lngSysButSize * 2
        ElseIf (lngExStyle And WS_EX_CONTEXTHELP) Then '������ť
            lngButSize = lngButSize + B_EDGE + lngSysButSize
        End If
        If (lngStyle And WS_THICKFRAME) Then
            lngRightEdgeOffset = lngButSize + GetSystemMetrics(SM_CXSIZEFRAME)
        Else
            lngRightEdgeOffset = lngButSize + GetSystemMetrics(SM_CXFIXEDFRAME)
        End If
    End If
    '��ȡ��ť��С
    If (lngExStyle And WS_EX_TOOLWINDOW) Then
        lngcxBut = GetSystemMetrics(SM_CXSMSIZE)
        lngcyBut = GetSystemMetrics(SM_CYSMSIZE)
      Else
        lngcxBut = GetSystemMetrics(SM_CXSIZE)
        lngcyBut = GetSystemMetrics(SM_CYSIZE)
    End If
    '��ȡ����ԭ��λ��
    Call GetWindowRect(lnghWnd, uRect)
    With uRect
        If gintCurTheme <> 1 Then
            'Win10,λ�ú���С�����ڰ�ť�غϣ����Լ�ȥһ����ť���
            .Right = .Right - lngRightEdgeOffset - IIf(gblnWin10, lngcxBut, 0) - B_EDGE * IIf(gintCurTheme = 0, 1, 0)
            If (lngStyle And WS_THICKFRAME) Then
                .Top = .Top + GetSystemMetrics(SM_CYSIZEFRAME)
              Else
                .Top = .Top + GetSystemMetrics(SM_CYFIXEDFRAME)
            End If
            .Top = .Top + (lngcyBut - 16) / 2
        Else
            If IsZoomed(lnghWnd) Then
                'Win10,λ�ú���С�����ڰ�ť�غϣ����Լ�ȥһ����ť���
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
'���ܣ���ȡ��ǰ����,0-WIndows ����,1-win7 AERO,win8,win10 ,2-WIn7 BASIC ,
    Dim lngValue As Long, strValue As String
    Dim intCurTheme As Integer
    '��ȡ��ǰ����
    If OS.GetRegValue("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\ThemeManager", "ThemeActive", strValue) Then
        intCurTheme = Val(strValue)
    Else
        intCurTheme = 0
    End If
    '����AEROЧ��ʱ ����ȡDWM���״����������ΪAERO
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
'        MsgBox "������ʾ:�����ܵ�ʧ��!"
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
'        MsgBox "������ʾ:��������ʧ��!" & vbCrLf
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
'�Ƿ������ϵͳ��
    Dim lii As LASTINPUTINFO
    lii.cbSize = Len(lii)
    GetLastInputInfo lii
    If GetTickCount - lii.dwTime > glngLockTime Then
        TimeToLock = True
    End If
End Function

'���ڼ�ر������еļ�����Ϣ��
Public Function MyKBHook(ByVal nCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    If nCode >= 0 Then
        MyKBHook = 0 '��ʾҪ���������Ϣ
        If wParam = vbKeyL And (GetKeyState(vbKeyMenu) And &HFF80) And (GetKeyState(vbKeyControl) And &HFF80) Then
            If gblnLock = False Then
                Call LockProg(True)
                MyKBHook = 1
            End If
        End If
    End If
    Call CallNextHookEx(glngHook, nCode, wParam, lParam) '����Ϣ������һ������
End Function


