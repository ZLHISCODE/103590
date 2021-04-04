Attribute VB_Name = "mdlRunas"
Option Explicit
'==================================================================================================
'��д           lshuo
'����           2019/4/18
'ģ��           mdlRunas
'˵��           ��������ģ�飬��������ͨȨ�����У�Ҳ�����Թ���ԱȨ�����С�
'==================================================================================================
Private Type STARTUPINFO
    cb                  As Long
    lpReserved          As Long
    lpDesktop           As Long
    lpTitle             As Long
    dwX                 As Long
    dwY                 As Long
    dwXSize             As Long
    dwYSize             As Long
    dwXCountChars       As Long
    dwYCountChars       As Long
    dwFillAttribute     As Long
    dwFlags             As Long
    wShowWindow         As Integer
    cbReserved2         As Integer
    lpReserved2         As Long
    hStdInput           As Long
    hStdOutput          As Long
    hStdError           As Long
End Type
'���ܣ�
'    ָ�������ڴ���ʱ�Ĵ���վ�����档��׼����������ڵ���ۡ�
'���壺
'    typedef struct _STARTUPINFO {
'      DWORD  cb;
'      LPTSTR lpReserved;
'      LPTSTR lpDesktop;
'      LPTSTR lpTitle;
'      DWORD  dwX;
'      DWORD  dwY;
'      DWORD  dwXSize;
'      DWORD  dwYSize;
'      DWORD  dwXCountChars;
'      DWORD  dwYCountChars;
'      DWORD  dwFillAttribute;
'      DWORD  dwFlags;
'      WORD   wShowWindow;
'      WORD   cbReserved2;
'      LPBYTE lpReserved2;
'      HANDLE hStdInput;
'      HANDLE hStdOutput;
'      HANDLE hStdError;
'    } STARTUPINFO, *LPSTARTUPINFO;
'��Ա��
'    cb
'       �ṹ�Ĵ�С�����ֽ�Ϊ��λ��
'    lpReserved
'       ����;����Ϊ�ա�
'    lpDesktop
'       ��������ƣ���˽��̵�����ʹ���վ�����ơ��ַ����еķ�б�ܱ�ʾ���ַ���ͬʱ��������ʹ���վ���ơ��йظ�����Ϣ����μ���������߳����ӡ�
'    lpTitle
'       ���ڿ���̨���̣�����������µĿ���̨���ڣ����ڱ���������ʾ������⡣���Ϊ�գ���ʹ�ÿ�ִ���ļ���������Ϊ���ڱ��⡣���ڲ������¿���̨���ڵ�GUI�����̨���̣��˲�������ΪNULL��
'    dwX
'       ���dwFlagsָ��STARTF_USEPOSITION����ó�Ա�Ǵ����´���ʱ�������Ͻǵ�xƫ����(������Ϊ��λ)�����򣬴˳�Ա�������ԡ�
'       ƫ����������Ļ�����Ͻǡ�����GUI���̣����CreateWindow��x������CW_USEDEFAULT�����½��̵�һ�ε���CreateWindow�����ص�����ʱ��ʹ��ָ����λ�á�
'    dwY
'       ���dwFlagsָ��STARTF_USEPOSITION����ó�Ա�Ǵ����´���ʱ�������Ͻǵ�yƫ��������λΪ���ء����򣬴˳�Ա�������ԡ�
'       ƫ����������Ļ�����Ͻǡ�����GUI���̣����CreateWindow��y������CW_USEDEFAULT�����½��̵�һ�ε���CreateWindowʱ��ʹ��ָ����λ���������ص��Ĵ��ڡ�
'    dwXSize
'       ���dwFlagsָ��STARTF_USESIZE����˳�Ա�Ǵ����´���ʱ���ڵĿ��(������Ϊ��λ)�����򣬴˳�Ա�������ԡ�
'       ����GUI���̣����CreateWindow��nWidth������CW_USEDEFAULT�����½��̽��ڵ�һ�ε���CreateWindow�����ص�����ʱ��ʹ�ô˷�����
'    dwYSize
'       ���dwFlagsָ��STARTF_USESIZE����˳�Ա�Ǵ����´���ʱ���ڵĸ߶�(������Ϊ��λ)�����򣬴˳�Ա�������ԡ�
'       ����GUI���̣����CreateWindow��nHeight������CW_USEDEFAULT�����½��̽��ڵ�һ�ε���CreateWindow�����ص�����ʱ��ʹ�ô˷�����
'    dwXCountChars
'       ���dwFlagsָ��STARTF_USECOUNTCHARS������ڿ���̨�����д�����һ���µĿ���̨���ڣ���ó�Աָ����Ļ��������ȣ���λΪ�ַ��С����򣬴˳�Ա�������ԡ�
'    dwYCountChars
'       ���dwFlagsָ��STARTF_USECOUNTCHARS������ڿ���̨�����д���һ���µĿ���̨���ڣ������Աָ����Ļ�������ĸ߶ȣ����ַ���Ϊ��λ�����򣬴˳�Ա�������ԡ�
'    dwFillAttribute
'       ���dwFlagsָ��STARTF_USEFILLATTRIBUTE��������ڿ���̨Ӧ�ó����д�����һ���µĿ���̨���ڣ���ó�Ա�ǳ�ʼ�ı��ͱ�����ɫ�����򣬴˳�Ա�������ԡ�
'       ���ֵ����������ֵ���������:FOREGROUND_BLUE��FOREGROUND_GREEN��FOREGROUND_RED��FOREGROUND_INTENSITY��BACKGROUND_BLUE��BACKGROUND_GREEN��BACKGROUND_RED��BACKGROUND_INTENSITY�����磬�����ֵ����ڰ�ɫ���������ɺ�ɫ�ı�:
'           FOREGROUND_RED| BACKGROUND_RED| BACKGROUND_GREEN| BACKGROUND_BLUE
'    dwFlags
'       ȷ�����̴�������ʱ�Ƿ�ʹ��ĳЩSTARTUPINFO��Ա��λ�ֶΡ��˳�Ա����������ֵ�е�һ������
Private Const STARTF_FORCEONFEEDBACK            As Long = &H40
'    ָʾ����CreateProcess���괦�ڷ���ģʽ���롣����ʾ���ڹ����ı������(��������������ʵ�ó����е�Pointersѡ�)��
'    ��������������ڽ��̷�����һ��GUI���ã���ôϵͳ�������5���ӵ�ʱ�䡣��������������ڽ�����ʾ��һ�����ڣ�ϵͳ������������ӵ�ʱ������ɴ��ڵĻ��ơ�
'    ϵͳ�ڵ�һ�ε���GetMessage֮��رշ�����꣬���ܽ����Ƿ����ڻ��ơ�
Private Const STARTF_FORCEOFFFEEDBACK           As Long = &H80
'    ָʾ�ڽ�������ʱǿ�ƹرշ����αꡣ����ʾ������ѡ���ꡣ
Private Const STARTF_PREVENTPINNING             As Long = &H2000
'    ָʾ���̴������κδ��ڲ��̶ܹ����������ϡ�
'    �����־������STARTF_TITLEISAPPID����
Private Const STARTF_RUNFULLSCREEN              As Long = &H20
'    ָʾ����Ӧ��ȫ��ģʽ�����У��������ڴ���ģʽ�����С�
'    �˱�־��������������x86������ϵĿ���̨Ӧ�ó���
Private Const STARTF_TITLEISAPPID               As Long = &H1000
'    lpTitle��Ա����һ��AppUserModelID���˱�ʶ�������������Ϳ�ʼ�˵������ʾӦ�ó��򣬲�ʹ������ȷ�Ŀ�ݷ�ʽ����ת�б��������ͨ����Ӧ�ó���ʹ��SetCurrentProcessExplicitAppUserModelID��GetCurrentProcessExplicitAppUserModelID�������������ô˱�־���йظ�����Ϣ����μ�Ӧ�ó����û�ģ��id��
'    ���ʹ��startf_preventpins�����޷���Ӧ�ó��򴰿ڹ̶����������ϡ�Ӧ�ó���ʹ���κ���appusermodelid��صĴ�������ֻ�Ḳ�Ǹô��ڵĴ����á�
'    �˱�־������STARTF_TITLEISLINKNAMEһ��ʹ��
Private Const STARTF_TITLEISLINKNAME            As Long = &H800
'    lpTitle��Ա�����û�Ϊ�����˽��̶����õĿ�ݷ�ʽ�ļ�(.lnk)��·������ͨ����shell�ڵ���ָ��������Ӧ�ó����.lnk�ļ�ʱ���á������Ӧ�ó�����Ҫ�������ֵ��
'    �˱�־������STARTF_TITLEISAPPIDһ��ʹ��
Private Const STARTF_UNTRUSTEDSOURCE            As Long = &H8000
'    ����������һ�������ŵ�Դ���йظ�����Ϣ����μ���ע��
Private Const STARTF_USECOUNTCHARS              As Long = &H8
'    dwXCountChars��dwYCountChars��Ա����������Ϣ��
Private Const STARTF_USEFILLATTRIBUTE           As Long = &H10
'    dwFillAttribute��Ա����������Ϣ
Private Const STARTF_USEHOTKEY                  As Long = &H200
'    hStdInput��Ա����������Ϣ
'    �˱�־������startf_usestdhandleһ��ʹ��
Private Const STARTF_USEPOSITION                As Long = &H4
'    dwX��dwY��Ա����������Ϣ
Private Const STARTF_USESHOWWINDOW              As Long = &H1
'    wShowWindow��Ա����������Ϣ
Private Const STARTF_USESIZE                    As Long = &H2
'    dwXSize��dwYSize��Ա����������Ϣ
Private Const STARTF_USESTDHANDLES              As Long = &H100
'    hStdInput��hStdOutput��hStdError��Ա����������Ϣ
'    ����ڵ��ý��̴�������ʱָ���˴˱�־�����������ǿɼ̳еģ�������bInheritHandles������������ΪTRUE���йظ�����Ϣ����μ�����̳С�
'    ����ڵ���GetStartupInfo����ʱָ���������־����ô��Щ��ԱҪô�ǽ��̴����ڼ�ָ���ľ��ֵ��Ҫô��INVALID_HANDLE_VALUE��
'    ��������Ҫ�ֱ�ʱ��������close�ֱ��ر��ֱ���
'    �˱�־������STARTF_USEHOTKEһ��ʹ��
'
'    wShowWindow
'        ���dwFlagsָ��STARTF_USESHOWWINDOW�������Ա������ShowWindow������nCmdShow������ָ�����κ�ֵ��SW_SHOWDEFAULT���⡣���򣬴˳�Ա�������ԡ�
'        ����GUI���̣���һ�ε���ShowWindowʱ������nCmdShow���������ԣ�wShowWindowָ����Ĭ��ֵ���ڶ�ShowWindow�ĺ��������У������ShowWindow��nCmdShow��������ΪSW_SHOWDEFAULT����ʹ��wShowWindow��Ա��
'    cbReserved2
'        Ԥ����C����ʱʹ��;�������㡣
'    lpReserved2
'        Ԥ����C����ʱʹ��;����Ϊ�ա�
'    hStdInput
'        ���dwFlagsָ��startf_usestdhandle����˳�Ա�����̵ı�׼�����������û��ָ��startf_usestdhandle����׼�����Ĭ��ֵ�Ǽ��̻�������
'        ���dwFlagsָ��STARTF_USEHOTKEY����˳�Աָ��һ���ȼ�ֵ�����ȼ�ֵ��ΪWM_SETHOTKEY��Ϣ��wParam�������͵�ӵ�и����̵�Ӧ�ó��򴴽��ĵ�һ�����������Ķ������ڡ������������WS_POPUP������ʽ�����ģ�����ǻ�������WS_EX_APPWINDOW��չ������ʽ�����򲻷����������й���ϸ��Ϣ����μ�CreateWindowEx��
'        ���򣬴˳�Ա�������ԡ�
'    hStdOutput
'        ���dwFlagsָ��startf_usestdhandle����˳�Ա�����̵ı�׼�����������򣬸ó�Ա�������ԣ���׼�����Ĭ��ֵ�ǿ���̨���ڵĻ�������
'        ���������������ת�б��������̣�ϵͳ��hStdOutput����Ϊ�������ľ�����þ�����������������̵�����������ת�б��йظ�����Ϣ����μ���ע��
'        Windows 7��Windows Server 2008 R2��Windows Vista��Windows Server 2008��Windows XP��Windows Server 2003:��һ��Ϊ��Windows 8��Windows Server 2012�����롣
'    hStdError
'        ���dwFlagsָ��startf_usestdhandle����˳�Ա�����̵ı�׼�����������򣬸ó�Ա�������ԣ���׼�����Ĭ��ֵ�ǿ���̨���ڵĻ�������
'��ע��
'    ����ͼ���û�����(GUI)���̣�����ϢӰ����CreateWindow������������ShowWindow������ʾ�ĵ�һ�����ڡ����ڿ���̨���̣����Ϊ�ý��̴������¿���̨�������Ϣ��Ӱ�����̨���ڡ����̿���ʹ��GetStartupInfo�����������ڴ�������ʱָ����STARTUPINFO�ṹ��
'    �����������GUI���̣�����û��ָ��STARTF_FORCEONFEEDBACK��STARTF_FORCEOFFFEEDBACK����ʹ��process feedback�αꡣGUI���̵���ϵͳָ��Ϊ��windows����
'    ���������������ת�б��������̣�ϵͳ��hStdOutput����Ϊ�������ľ�����þ�����������������̵�����������ת�б�Ҫ������������ʹ��GetStartupInfo������STARTUPINFO�ṹ�������hStdOutput�Ƿ������á�Ȼ�󣬽��̿���ʹ�þ������λ���Ĵ��ڡ�
'    ���STARTF_UNTRUSTEDSOURCE��־����GetStartupInfo�������ص�STARTUPINFO�ṹ�����õģ���ôӦ�ó���Ӧ��֪���������ǲ������εġ���������˴˱�־��Ӧ�ó���Ӧ�ý���Ǳ�ڵ�Σ�����ԣ���ꡣ�������ݺ��Զ���ӡ�������־�ǿ�ѡ�ġ���ʹ�ò������ε���������������ʱ����������CreateProcess��Ӧ�ó������ô˱�־���Ա㴴���Ľ��̿���Ӧ���ʵ��Ĳ��ԡ�
'    STARTF_UNTRUSTEDSOURCE��־��Windows Vista��֧������������Windows 10 SDK֮ǰû����SDKͷ�ļ��ж��塣Ҫ��Windows 10֮ǰ�İ汾��ʹ�øñ�־�������ڳ������ֶ���������
'֧��
'    ���֧�ֿͻ�
'       Windows XP [ֻ����������Ӧ�ó�ʽ]
'    ���֧�ַ�����
'       Windows Server 2003[ֻ����������Ӧ�ó���]
'    Header
'       WinBase.h on Windows XP, Windows Server 2003, Windows Vista, Windows 7, Windows Server 2008��Windows Server 2008 R2(����Windows.h);
'       Windows 8��Windows Server 2012�ϵ�Processthreadsapi.h
'    Unicode��ANSI����
'       STARTUPINFOW (Unicode)��STARTUPINFOA (ANSI)
Private Type PROCESS_INFORMATION
    hProcess            As Long
    hThread             As Long
    dwProcessId         As Long
    dwThreadId          As Long
End Type
'���ܣ�
'   ���������´����Ľ��̼������̵߳���Ϣ������CreateProcess��CreateProcessAsUser��CreateProcessWithLogonW��CreateProcessWithTokenW����һ��ʹ�á�
'����
'    typedef struct _PROCESS_INFORMATION {
'      HANDLE hProcess;
'      HANDLE hThread;
'      DWORD  dwProcessId;
'      DWORD  dwThreadId;
'    } PROCESS_INFORMATION, *LPPROCESS_INFORMATION;
'��Ա
'    hProcess
'       �´������̵ľ������������ڶԽ��̶���ִ�в��������к�����ָ������
'    hThread
'       �´������̵����̵߳ľ������������ڶ�thread����ִ�в��������к�����ָ���̡߳�
'    dwProcessId
'       �����ڱ�ʶ���̵�ֵ����ֵ�Ӵ�������ʱ����Ч��ֱ���رյ����̵����о�����ͷŽ��̶���Ϊֹ;��ʱ���������ñ�ʶ����
'    dwThreadId
'       �����ڱ�ʶ�̵߳�ֵ����ֵ���̴߳���ʱ����Ч��ֱ���̵߳����о�����رղ��ͷ��̶߳���Ϊֹ;��ʱ���������ñ�ʶ����
'��ע
'    ��������ɹ�����ȷ������close����������ر�hProcess��hThread��������򣬵��ӽ����˳�ʱ��ϵͳ�޷������ӽ��̵Ľ��̽ṹ����Ϊ��������Ȼӵ���ӽ��̵Ĵ򿪾�������ǣ�����������ֹʱ��ϵͳ���ر���Щ�����������ӽ��̶�����صĽṹ���ڴ�ʱ�������
'֧��
'    ���֧�ֿͻ�
'       Windows XP [ֻ����������Ӧ�ó�ʽ]
'    ���֧�ַ�����
'       Windows Server 2003[ֻ����������Ӧ�ó���]
'    Header
'       WinBase.h on Windows XP, Windows Server 2003, Windows Vista, Windows 7, Windows Server 2008 and Windows Server 2008 R2 (include Windows.h);
'       Processthreadsapi.h on Windows 8 and Windows Server 2012
Private Declare Function CreateProcessWithLogon Lib "advapi32" Alias "CreateProcessWithLogonW" (ByVal lpUsername As Long, ByVal lpDomain As Long, ByVal lpPassword As Long, ByVal dwLogonFlags As Long, ByVal lpApplicationName As Long, ByVal lpCommandLine As Long, ByVal dwCreationFlags As Long, ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As Long, lpStartupInfo As STARTUPINFO, lpProcessInfo As PROCESS_INFORMATION) As Long
'���壺
'    BOOL WINAPI CreateProcessWithLogonW(
'      _In_        LPCWSTR               lpUsername,
'      _In_opt_    LPCWSTR               lpDomain,
'      _In_        LPCWSTR               lpPassword,
'      _In_        DWORD                 dwLogonFlags,
'      _In_opt_    LPCWSTR               lpApplicationName,
'      _Inout_opt_ LPWSTR                lpCommandLine,
'      _In_        DWORD                 dwCreationFlags,
'      _In_opt_    LPVOID                lpEnvironment,
'      _In_opt_    LPCWSTR               lpCurrentDirectory,
'      _In_        LPSTARTUPINFOW        lpStartupInfo,
'      _Out_       LPPROCESS_INFORMATION lpProcessInfo
'    );
'���ܣ�
'    ����һ���½��̼������̡߳�Ȼ���½�����ָ��ƾ��(�û����������)�İ�ȫ������������ָ���Ŀ�ִ���ļ���������ѡ�����ָ���û����û������ļ���
'    �˺���������CreateProcessAsUser��CreateProcessWithTokenW������ֻ�ǵ����߲���Ҫ����LogonUser��������֤�û���ݲ���ȡ���ơ�
'������
'    lpUsername _In_
'       �û���������Ҫ��¼���û��ʻ������ơ����ʹ��UPN��ʽuser@ DNS_domain_name, lpDomain��������ΪNULL
'       �û��ʻ�������б��ؼ�����ϵı��ص�¼Ȩ�ޡ���Ȩ�����蹤��վ�ͷ������ϵ������û�������������������ϵĹ���Ա��
'    lpDomain _In_opt_
'       �ʻ����ݿ����lpUsername�ʻ����������������ơ�����ò���Ϊ�գ��������UPN��ʽָ���û�����
'    lpPassword _In_
'       lpUsername�ʻ�����������
'    dwLogonFlags _In_
'       ��¼ѡ��������������0(0)��Ҳ����������ֵ֮һ��
Private Const LOGON_WITH_PROFILE                As Long = &H1
'    ��¼��Ȼ����HKEY_USERSע������м����û���Ҫ�ļ��������ڼ��ظ�Ҫ�ļ��󷵻ء����ظ�Ҫ�ļ����ܺܺ�ʱ���������ֻ�ڱ������HKEY_CURRENT_USERע������е���Ϣʱ��ʹ�����ֵ��
'    Windows Server 2003:���½�����ֹ��ж�ظ�Ҫ�ļ����������Ƿ񴴽����ӽ��̡�
'    Windows XP: ���½��̼��䴴���������ӽ�����ֹ��ж�ظ�Ҫ�ļ�
Private Const LOGON_NETCREDENTIALS_ONLY         As Long = &H2
'    ��¼��������������ʹ��ָ����ƾ�ݡ��½���ʹ�����������ͬ�����ƣ�����ϵͳ��LSA�д���һ���µĵ�¼�Ự�����Ҹý���ʹ��ָ����ƾ����Ϊȱʡƾ�ݡ�
'    ��ֵ�����ڴ���һ�����̣��ý����ڱ���ʹ�õ�ƾ�ݼ���Զ��ʹ�õ�ƾ�ݼ���ͬ������û�����ι�ϵ����䳡���зǳ����á�
'    ϵͳ����ָ֤����ƾ�ݡ���ˣ����̿������������������޷�����������Դ��
'    lpApplicationName _In_opt_
'        Ҫִ�е�ģ������ơ����ģ�������һ������windows��Ӧ�ó���������ؼ�������к��ʵ���ϵͳ�����������������͵�ģ��(����MS-DOS��OS/2)��
'        �ַ�������ָ��Ҫִ�е�ģ�������·�����ļ�����Ҳ����ָ���������ơ�����ǲ������ƣ�����ʹ�õ�ǰ�������͵�ǰĿ¼����ɹ淶���ú�����ʹ������·�����˲�����������ļ�����չ��;û��Ĭ����չ��
'        lpApplicationName��������ΪNULL��ģ�����Ʊ�����lpCommandLine�ַ����е�һ���Կո�ָ��İ����ơ����ʹ�ð����ո�ĳ��ļ�������ʹ�ô����ŵ��ַ���ָʾ�ļ����Ľ����Ͳ����Ŀ�ʼλ��;�����ļ�������ģ���ġ�
'        ���磬������ַ��������ò�ͬ�ķ�ʽ����:
'        "c:\program files\sub dir\program name"
'        ϵͳ��ͼ��������˳�������Щ������:
'        c:\program.exe files\sub dir\program name
'        c:\program files\sub.exe dir\program name
'        c:\program files\sub dir\program.exe name
'        c:\program files\sub dir\program name.exe
'        �����ִ��ģ����16λ��Ӧ�ó���lpApplicationNameӦ��ΪNULL, lpCommandLineָ����ַ���Ӧ��ָ����ִ��ģ�鼰�������
'    lpCommandLine _Inout_opt_
'        Ҫִ�е������С�����ַ�������󳤶���1024���ַ������lpApplicationNameΪNULL����lpCommandLine��ģ������������ΪMAX_PATH�ַ���
'        ���������޸�����ַ��������ݡ���ˣ��ò���������ָ��ֻ���ڴ��ָ��(����const�����������ַ���)������ò����ǳ����ַ�������ú������ܻᵼ�·��ʳ�ͻ��
'        lpCommandLine��������ΪNULL������ʹ��lpApplicationNameָ����ַ�����Ϊ�����С�
'        ���lpApplicationName��lpCommandLine���Ƿǿյģ���*lpApplicationNameָ��Ҫִ�е�ģ�飬*lpCommandLineָ�������С��½��̿���ʹ��GetCommandLine�������������С���C��д�Ŀ���̨���̿���ʹ��argc��argv�������������С���Ϊargv[0]��ģ������C����Աͨ����ģ������Ϊ�������еĵ�һ�������ظ���
'        ���lpApplicationNameΪNULL�����������е�һ���Կո�ָ��ı��ָ��ģ�����ơ����ʹ�ð����ո�ĳ��ļ�������ʹ�ô����ŵ��ַ�����ָʾ�ļ����Ľ����Ͳ����Ŀ�ʼλ��(�����lpApplicationName������˵��)������ļ�����������չ������׷��.exe����ˣ�����ļ�����չ���ǡ�com������������������com��չ��������ļ�����û����չ���ľ������������ļ�������·�����򲻸���.exe������ļ���������Ŀ¼·����ϵͳ������˳��������ִ���ļ�:
'            ����Ӧ�ó����Ŀ¼
'            �����̵ĵ�ǰĿ¼
'            32 λWindowsϵͳĿ¼��ʹ��GetSystemDirectory������ȡ��Ŀ¼��·��
'            16λWindowsϵͳĿ¼��û�к������Ի�����Ŀ¼��·�������ǿ�����������
'            WindowsĿ¼��ʹ��GetWindowsDirectory������ȡ��Ŀ¼��·����
'            PATH�����������г���Ŀ¼��ע�⣬�˺�����������App Paths registry��ָ����ÿ��Ӧ�ó���·����Ҫ�����������а���ÿ��Ӧ�ó����·������ʹ��ShellExecute����
'        ϵͳ��һ�����ַ���ӵ��������ַ����У��Խ��ļ���������ֿ����⽫ԭʼ�ַ����ֳ������ַ��������ڲ�����
'    dwCreationFlags _In_
'           ������δ������̵ı�־��Ĭ������£�CREATE_DEFAULT_ERROR_MODE��CREATE_NEW_CONSOLE��CREATE_NEW_PROCESS_GROUP��־�����õġ�����ʹ��û�����øñ�־��ϵͳ�Ĺ���Ҳ������һ����
Private Const CREATE_DEFAULT_ERROR_MODE         As Long = &H4000000
'    �½��̲��̳е��ý��̵Ĵ���ģʽ���෴��CreateProcessWithLogonWΪ�½����ṩ��ǰĬ�ϴ���ģʽ��Ӧ�ó���ͨ������SetErrorMode���õ�ǰĬ�ϴ���ģʽ��
'    Ĭ����������ô˱�־
Private Const CREATE_NEW_CONSOLE                As Long = &H10
'    �½�����һ���¿���̨�������Ǽ̳и����̵Ŀ���̨���˱�־������DETACHED_PROCESS��־һ��ʹ�á�
'    Ĭ����������ô˱�־
Private Const CREATE_NEW_PROCESS_GROUP          As Long = &H200
'    �½������½�����ĸ����̡�����������˸����̵����к�����̡��½�����Ľ��̱�ʶ������lpProcessInfo�����з��صĽ��̱�ʶ����ͬ��GenerateConsoleCtrlEvent����ʹ�ý���������һ�����̨���̷���CTRL C��CTRL + BREAK�źš�
'    Ĭ����������ô˱�־
Private Const CREATE_SEPARATE_WOW_VDM           As Long = &H800
'    �˱�־������������16λwindows��Ӧ�ó���ʱ��Ч��������úã��½��̽���˽������DOS����(VDM)�����С�Ĭ������£����л���windows��16λӦ�ó���������һ�������VDM�С��������еĺô��Ǳ���ֻ����ֹ����VDM;�ڲ�ͬVDMs�����е��κ��������򶼿����������С����⣬�����ڵ���VDMs�е�16λ����windows��Ӧ�ó�����е�����������У�����ζ�����һ��Ӧ�ó�����ʱֹͣ��Ӧ������VDMs�е�Ӧ�ó��򽫼����������롣
Private Const CREATE_SUSPENDED                  As Long = &H4
'    �½��̵����߳����ڹ���״̬�´����ģ��ڵ���ResumeThread����֮ǰ�������С�
Private Const CREATE_UNICODE_ENVIRONMENT        As Long = &H400
'    ָʾlpEnvironment�����ĸ�ʽ����������˴˱�־��lpEnvironmentָ��Ļ����齫ʹ��Unicode�ַ������򣬻�����ʹ��ANSI�ַ���
'        �˲����������½��̵����ȼ��࣬��������ȷ�������̵߳ĵ������ȼ����й�ֵ�б���μ�GetPriorityClass�����û��ָ���κ����ȼ����־�������ȼ���Ĭ��ΪNORMAL_PRIORITY_CLASS�����Ǵ������̵����ȼ�����IDLE_PRIORITY_CLASS�����idle_normal_priority_class���ڱ����У��ӽ��̽��յ��ý��̵�Ĭ�����ȼ��ࡣ
'    lpEnvironment _In_opt_
'        ָ���½��̵Ļ������ָ�롣����ò���ΪNULL���������̽�ʹ����lpUsernameָ�����û���Ҫ�ļ������Ļ�����
'        ����������null��β���ַ�����ɵ���null��β�Ŀ顣ÿ���ַ�������ʽ����:
'        ���� = ֵ
'        ��Ϊ�Ⱥ�(=)�����ָ��������Բ����ڻ���������������ʹ������
'        ��������԰���Unicode��ANSI�ַ������lpEnvironmentָ��Ļ��������Unicode�ַ�����ȷ��dwCreationFlags����CREATE_UNICODE_ENVIRONMENT������ò���ΪNULL�����Ҹ����̵Ļ��������Unicode�ַ����򻹱���ȷ��dwCreationFlags����CREATE_UNICODE_ENVIRONMENT��
'        ANSI������������0(0)�ֽ���ֹ:һ���������һ���ַ�������һ��������ֹ�ÿ顣Unicode��������ֹΪ4�����ֽ�:���һ���ַ�����ֹΪ2���ֽڣ����һ���ַ�����ֹΪ2���ֽڡ�
'        ҪΪ�ض��û�����������ĸ�������ʹ��CreateEnvironmentBlock������
'    lpCurrentDirectory  _In_opt_
'        ���̵�ǰĿ¼������·�����ַ���������ָ��UNC·����
'        ����ò���ΪNULL�����½��̾�������ý�����ͬ�ĵ�ǰ��������Ŀ¼����������Ҫ��Ϊ��Ҫ����Ӧ�ó���ָ�����ʼ�������͹���Ŀ¼��shell�ṩ�ġ�
'    lpStartupInfo _In_
'        ָ��STARTUPINFO�ṹ��ָ�롣Ӧ�ó�����뽫ָ���û��ʻ���Ȩ����ӵ�ָ���Ĵ���վ�����棬������WinSta0\Default��
'        ���lpDesktop��ԱΪNULL����ַ��������½��̼̳��丸���̵�����ʹ���վ��Ӧ�ó�����뽫ָ���û��ʻ���Ȩ����ӵ��̳еĴ���վ�����档
'        CreateProcessWithLogonW��ָ���û��ʻ���Ȩ����ӵ��̳еĴ���վ�����档
'        ��������Ҫ���ʱ��������close����ر�STARTUPINFO�еľ����
'        ��Ҫ���ǣ����STARTUPINFO�ṹ��dwFlags��Աָ��startf_usestdhandle�����׼����ֶν����Ӹ��ĵظ��Ƶ��ӽ��̣�����������֤�������߸���ȷ����Щ�ֶΰ�����Ч�ľ��ֵ������ȷ��ֵ���ܵ����ӽ�����Ϊ�����������ʹ��Ӧ�ó�����֤��������ʱ��֤���߼����Ч�����
'    lpProcessInfo _Out_
'        ָ��PROCESS_INFORMATION�ṹ��ָ�룬�ýṹ�����½��̵ı�ʶ��Ϣ���������̵ľ����
'        PROCESS_INFORMATION�еľ���ڲ���Ҫʱ����ʹ��close��������رա�
'����ֵ
'    ��������ɹ�������ֵΪ���㡣
'    �������ʧ�ܣ�����ֵΪ0(0)��Ҫ��ȡ��չ�Ĵ�����Ϣ�������GetLastError��
'    ע�⣬�����ڽ�����ɳ�ʼ��֮ǰ���ء�����޷��ҵ������DLL���ʼ��ʧ�ܣ����̽���ֹ��Ҫ��ȡ���̵���ֹ״̬�������GetExitCodeProcess��
'��ע
'    Ĭ������£�CreateProcessWithLogonW���Ὣָ�����û������ļ����ص�HKEY_USERSע������С�����ζ�Ŷ�HKEY_CURRENT_USERע������е���Ϣ�ķ��ʿ��ܲ������������������¼һ�µĽ�������������ڵ���CreateProcessWithLogonW֮ǰ��ͨ��ʹ��LOGON_WITH_PROFILE��ͨ������LoadUserProfile���������û�ע���hive���ص�HKEY_USERS�С�
'    ���lpEnvironment����ΪNULL���������̽�ʹ����lpUserNameָ�����û���Ҫ�ļ������Ļ����顣���û������HOMEDRIVE��HOMEPATH������CreateProcessWithLogonW���޸Ļ����飬��ʹ���û�����Ŀ¼����������·����
'    ����ʱ���½��̺��߳̾�������������ķ���Ȩ��(PROCESS_ALL_ACCESS��THREAD_ALL_ACCESS)�������κ�һ����������û���ṩ��ȫ�����������������Ҫ�����Ͷ��������κκ�����ʹ�øþ�������ṩ��ȫ������ʱ�������������Ȩ֮ǰ�Ծ�������к���ʹ��ִ�з��ʼ�顣������ʱ��ܾ���������̲���ʹ�þ�����ʽ��̻��̡߳�
'    Ҫ������ȫ���ƣ��뽫PROCESS_INFORMATION�ṹ�е����̾�����ݸ�OpenProcessToken������
'    ���̱�����һ�����̱�ʶ������ʶ���ڽ�����ֹ֮ǰ��Ч��������������ʶ���̣�Ҳ������OpenProcess������ָ�������򿪽��̵ľ���������еĳ�ʼ�߳�Ҳ������һ���̱߳�ʶ����������OpenThread������ָ���������̵߳ľ������ʶ�����߳���ֹ֮ǰ����Ч�ģ����ҿ�������Ωһ�ر�ʶϵͳ�е��̡߳���Щ��ʶ����PROCESS_INFORMATION�з��ء�
'    �����߳̿���ʹ��WaitForInputIdle�����ȴ���ֱ���½�����ɳ�ʼ�����������ڵȴ��û������û�������������ڸ����̺��ӽ���֮���ͬ���ǳ����ã���ΪCreateProcessWithLogonW�ڲ��ȴ��½�����ɳ�ʼ��������·��ء����磬�������̽��ڳ��Բ����������̹����Ĵ���֮ǰʹ��WaitForInputIdle��
'    �رս��̵���ѡ������ʹ��ExitProcess��������Ϊ�ú����򸽼ӵ����̵�����dll������ֹ֪ͨ�������رս��̵ķ�����֪ͨ���ӵ�dll��ע�⣬��һ���̵߳���ExitProcessʱ�����̵������߳̽�����ֹ����û�л���ִ���κ���������(��������dll���߳���ֹ����)���йظ�����Ϣ����μ���ֹ���̡�
'��ȫ˵��
'    lpApplicationName��������ΪNULL�����ҿ�ִ�����Ʊ�����lpCommandLine�е�һ���ո�ָ����ַ����������ִ���ļ���·���������пո������ں��������ո�ķ�ʽ�����ܻ����в�ͬ�Ŀ�ִ���ļ�����������ʾ������Ϊ������ͼ���С�Program��������������ڣ������ǡ�MyApp.exe����
'    LPTSTR szCmdline[]=_tcsdup(TEXT("C:\\Program Files\\MyApp"));
'    CreateProcessWithLogonW (��,szCmdline����)
'    ��������û�������һ����Ϊ��Program����Ӧ�ó�����ϵͳ�ϣ��κ�ʹ�ó����ļ�Ŀ¼�������CreateProcessWithLogonW�ĳ��򶼻����ж����û�Ӧ�ó��򣬶�����Ԥ�ڵ�Ӧ�ó���
'    Ϊ�˱���������⣬��ҪΪlpApplicationName����NULL�����ΪlpApplicationName����NULL������lpCommandLine�еĿ�ִ��·����Χʹ�����ţ��������ʾ����ʾ:
'    LPTSTR szCmdline[]=_tcsdup(TEXT("\"C:\\Program Files\\MyApp\""));
'    CreateProcessWithLogonW(..., szCmdline, ...)
'Requirements
'    ���֧�ֿͻ���
'       Windows XP [ֻ�������]
'    ���֧�ַ�����
'       Windows Server 2003 [ֻ�������]
'    Header
'       WinBase.h (include Windows.h)
'    Library
'       advapi32.lib
'    dll
'       advapi32.dll
Private Declare Function LogonUser Lib "advapi32.dll" Alias "LogonUserA" (ByVal lpszUsername As String, ByVal lpszDomain As String, _
                        ByVal lpszPassword As String, ByVal dwLogonType As Long, ByVal dwLogonProvider As Long, phToken As Long) As Long
'���ܣ�
'   LogonUser�������Խ��û���¼�����ؼ���������ؼ�����ǵ���LogonUser�ļ������������ʹ��LogonUser��¼��Զ�̼������������ʹ���û�������ָ���û�����ʹ������������û����������֤����������ɹ�������ձ�ʾ��¼�û������ƾ����Ȼ��������ʹ�ô����ƾ��ģ��ָ�����û��������ڴ��������£�������ָ���û��������������е����̡�
'���壺
'    BOOL LogonUser(
'      _In_     LPTSTR  lpszUsername,
'      _In_opt_ LPTSTR  lpszDomain,
'      _In_opt_ LPTSTR  lpszPassword,
'      _In_     DWORD   dwLogonType,
'      _In_     DWORD   dwLogonProvider,
'      _Out_    PHANDLE phToken
'    );
'��Ա
'    lpszUsername _In_
'       ָ����null��β���ַ�����ָ�룬���ַ���ָ���û���������Ҫ��¼���û��ʻ������ơ����ʹ���û���������(UPN)��ʽUser@DNSDomainName, lpszDomain��������ΪNULL��
'    lpszDomain _In_opt_
'       ָ����null��β���ַ�����ָ�룬���ַ���ָ���ʻ����ݿ����lpszUsername�ʻ����������������ơ�����ò���Ϊ�գ��������UPN��ʽָ���û����������������ǡ������ú���ֻʹ�ñ����ʻ����ݿ���֤�ʻ���
'    lpszPassword _In_opt_
'       ָ����null��β���ַ�����ָ�룬���ַ���ָ��lpszUsernameָ�����û��ʻ����������롣����ʹ���������ͨ������SecureZeroMemory�������ڴ���������롣�йر�������ĸ�����Ϣ����μ��������롣
'    dwLogonType _In_
'       Ҫִ�еĵ�¼���������͡��ò���������Winbase.h�ж��������ֵ֮һ
Private Const LOGON32_LOGON_BATCH               As Long = &H4
'    �˵�¼�������������������������������������ϣ����̿��Դ����û�ִ�У��������û���ֱ�Ӹ�Ԥ����������Ҳ���������ܸ��ߵķ���������Щ������һ�δ�����ി�ı������֤���ԣ������ʼ���web��������
Private Const LOGON32_LOGON_INTERACTIVE         As Long = &H2
'    �˵�¼���������ڽ���ʹ�ü�������û����������ն˷�������Զ��shell�����ƽ��̵�¼���û������ֵ�¼���ͻ��ж���Ŀ�������Ϊ�Ͽ����ӵĲ��������¼��Ϣ;��ˣ�����������ĳЩ�ͻ���/������Ӧ�ó��򣬱����ʼ���������
Private Const LOGON32_LOGON_NETWORK             As Long = &H3
'    �˵�¼�������ڸ����ܷ�������֤�������롣LogonUser����������˵�¼���͵�ƾ��
Private Const LOGON32_LOGON_NETWORK_CLEARTEXT   As Long = &H8
'    �˵�¼�����������֤���б������ƺ����룬�������������ģ��ͻ���ʱ���ӵ�������������������������Խ������Կͻ����Ĵ��ı�ƾ֤������LogonUser����֤�û�����ͨ���������ϵͳ��������Ȼ����������������ͨ�š�
Private Const LOGON32_LOGON_NEW_CREDENTIALS     As Long = &H9
'    �˵�¼����������÷���¡�䵱ǰ���Ʋ�Ϊ��վ����ָ���µ�ƾ�ݡ��µĵ�¼�Ự������ͬ�ı��ر�ʶ��������������������ʹ�ò�ͬ��ƾ�ݡ�
'    �˵�¼���ͽ���LOGON32_PROVIDER_WINNT50��¼�ṩ����֧��
Private Const LOGON32_LOGON_SERVICE             As Long = &H5
'    ָʾ�������͵�¼�����ṩ���ʻ��������÷�����Ȩ��
Private Const LOGON32_LOGON_UNLOCK              As Long = &H7
'    ����֧��GINAs
'    Windows Server 2003��Windows XP:���ֵ�¼������ΪGINA dll�ṩ�ģ����ڵ�¼������ʽʹ�ü�������û����˵�¼���Ϳ�������һ��Ωһ����Ƽ�¼����ʾ����վ��ʱ������
'
'    dwLogonProvider _In_
'       ָ����¼�ṩ���򡣴˲�������������ֵ֮һ
Private Const LOGON32_PROVIDER_DEFAULT          As Long = &H0
'    ʹ��ϵͳ�ı�׼��¼�ṩ����Ĭ�ϵİ�ȫ�ṩ������Э�̵ģ�������Ϊ��������NULL�����û�������UPN��ʽ���ڱ����У�Ĭ���ṩ������NTLM��
Private Const LOGON32_PROVIDER_WINNT50          As Long = &H3
'    ʹ��Э�̵�¼�ṩ����
Private Const LOGON32_PROVIDER_WINNT40          As Long = &H2
'    ʹ��NTLM��¼�ṩ����
'
'    phToken _Out_
'        ָ����������ָ�룬�ñ������ձ�ʾָ���û������Ƶľ����
'        �������ڵ���ImpersonateLoggedOnUser����ʱʹ�÷��صľ����
'        �ڴ��������£����صľ����һ����Ҫ���ƣ��������ڵ���CreateProcessAsUser����ʱʹ���������ǣ����ָ��LOGON32_LOGON_NETWORK��־��LogonUser������һ��ģ�����ƣ����ǵ���DuplicateTokenEx����ת��Ϊһ�������ƣ�������������CreateProcessAsUser��ʹ��������ơ�
'        ����������Ҫ������ʱ��ͨ������close����������ر�����
'����ֵ
'    ��������ɹ����������ط��㡣
'    �������ʧ�ܣ��������㡣Ҫ��ȡ��չ�Ĵ�����Ϣ�������GetLastError��
'��ע
'    LOGON32_LOGON_NETWORK��¼���������ģ�����������������:
'    ��������ģ�����ƣ������������ơ���������CreateProcessAsUser������ֱ��ʹ��������ơ����ǣ������Ե���DuplicateTokenEx����������ת��Ϊ�����ƣ�Ȼ����CreateProcessAsUser��ʹ������
'    ���������ת��Ϊ�����Ʋ���CreateProcessAsUser��ʹ�������������̣����½��̲���ͨ���ض����������������Դ������Զ�̷��������ӡ����һ�������ǣ����������Դ���ܷ��ʿ��ƣ���ô�½��̽��ܹ���������
'    �����������ҪSE_TCB_NAME��Ȩ�����������ڵ�¼Passport�ʻ���
'    ��lpszUsernameָ�����ʻ�������б�Ҫ���ʻ�Ȩ�ޡ����磬Ҫʹ��LOGON32_LOGON_INTERACTIVE��־��¼�û����û�(���û���������)����ӵ��SE_INTERACTIVE_LOGON_NAME�ʻ����й�Ӱ����ֵ�¼�������ʻ�Ȩ���б���μ��ʻ�Ȩ�޳�����
'    �����������һ�����ƣ�����Ϊ�û��ѵ�¼�����������CreateProcessAsUser���ر����ƣ�ϵͳ����Ϊ���û���Ȼ��¼��ֱ������(�Լ������ӽ���)������
'    ���LogonUser���óɹ���ϵͳ��ͨ�������ṩ�ߵ�NPLogonNotify��ڵ㺯��֪ͨ�����ṩ�߷����˵�¼��
'֧��
'    ���֧�ֿͻ�
'       Windows XP [ֻ����������Ӧ�ó�ʽ]
'    ���֧�ַ�����
'       Windows Server 2003[ֻ����������Ӧ�ó���]
'    Header
'       Winbase.h (����Windows.h)
'    Library
'       advapi32.lib
'    dll
'       advapi32.dll
'    Unicode��ANSI����
'       LogonUserW (Unicode)��LogonUserA(ANSI)
Private Declare Function CreateProcessAsUser Lib "advapi32.dll" Alias "CreateProcessAsUserA" (ByVal hToken As Long, ByVal lpApplicationName As String, _
                    ByVal lpCommandLine As String, ByVal lpProcessAttributes As Long, ByVal lpThreadAttributes As Long, ByVal bInheritHandles As Long, _
                    ByVal dwCreationFlags As Long, ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As String, lpStartupInfo As STARTUPINFO, _
                    lpProcessInformation As PROCESS_INFORMATION) As Long
'���ܣ�
'    ����һ���½��̼������̡߳��½�����ָ�����Ʊ�ʾ���û��İ�ȫ������������
'    ͨ��������CreateProcessAsUser�����Ľ��̱������se_incree_quota_name��Ȩ��������Ʋ��ɷ��䣬�������ҪSE_ASSIGNPRIMARYTOKEN_NAME��Ȩ������ú�����ERROR_PRIVILEGE_NOT_HELD(1314)��ʧ�ܣ���ʹ��CreateProcessWithLogonW������CreateProcessWithLogonW����Ҫ��Ȩ�����Ǳ�������ָ�����û��ʻ�����ʽ�ص�¼��ͨ�������ʹ��CreateProcessWithLogonW�������б���ƾ֤�Ľ��̡�
'����
'    BOOL WINAPI CreateProcessAsUser(
'      _In_opt_    HANDLE                hToken,
'      _In_opt_    LPCTSTR               lpApplicationName,
'      _Inout_opt_ LPTSTR                lpCommandLine,
'      _In_opt_    LPSECURITY_ATTRIBUTES lpProcessAttributes,
'      _In_opt_    LPSECURITY_ATTRIBUTES lpThreadAttributes,
'      _In_        BOOL                  bInheritHandles,
'      _In_        DWORD                 dwCreationFlags,
'      _In_opt_    LPVOID                lpEnvironment,
'      _In_opt_    LPCTSTR               lpCurrentDirectory,
'      _In_        LPSTARTUPINFO         lpStartupInfo,
'      _Out_       LPPROCESS_INFORMATION lpProcessInformation
'    );
'��Ա��
'    hToken _In_opt_
'        ��ʾ�û��������Ƶľ��������������TOKEN_QUERY��TOKEN_DUPLICATE��TOKEN_ASSIGN_PRIMARY����Ȩ�ޡ��йظ�����Ϣ����μ��������ƶ���ķ���Ȩ�ޡ����Ʊ�ʾ���û�������ж�lpApplicationName��lpCommandLine����ָ����Ӧ�ó���Ķ�ȡ��ִ�з���Ȩ��
'        Ҫ��ñ�ʾָ���û��������ƣ������LogonUser���������ߣ������Ե���DuplicateTokenEx������ģ������ת��Ϊ�����ơ�������ģ��ͻ����ķ�����Ӧ�ó��򴴽����пͻ�����ȫ�����ĵ����̡�
'        ���hToken�ǵ��÷��������Ƶ����ް汾������ҪSE_ASSIGNPRIMARYTOKEN_NAME��Ȩ�������û�����ñ�Ҫ����Ȩ��CreateProcessAsUser���ڵ����ڼ�������Щ��Ȩ���йظ�����Ϣ����μ�ʹ����Ȩ���С�
'        �ն˷���:������������ָ���ĻỰ�����С�Ĭ������£�������LogonUser��ͬ�ĻỰ��Ҫ���ĻỰ����ʹ��SetTokenInformation������
'    lpApplicationName _In_opt_
'        Ҫִ�е�ģ������ơ����ģ�������һ������windows��Ӧ�ó���������ؼ�������к��ʵ���ϵͳ�����������������͵�ģ��(����MS-DOS��OS/2)��
'        �ַ�������ָ��Ҫִ�е�ģ�������·�����ļ�����Ҳ����ָ���������ơ����ڲ������ƣ�����ʹ�õ�ǰ�������͵�ǰĿ¼����ɹ淶���ú�������ʹ������·�����˲�����������ļ�����չ��;û��Ĭ����չ��
'        lpApplicationName��������ΪNULL������������£�ģ����������lpCommandLine�ַ����е�һ���Կո�ָ��ı�ǡ����ʹ�ð����ո�ĳ��ļ�������ʹ�ô����ŵ��ַ���ָʾ�ļ����Ľ����Ͳ����Ŀ�ʼλ��;�����ļ�������ģ���ġ����磬�����ַ�����c:\program files\sub dir\program name��������ַ��������ö��ַ�ʽ���͡�ϵͳ��ͼ��������˳�������Щ������:
'        c:\program.exe files\sub dir\program name
'        c:\program files\sub.exe dir\program name
'        c:\program files\sub dir\program.exe name
'        c:\program files\sub dir\program name.exe
'        �����ִ��ģ����16λ��Ӧ�ó���lpApplicationNameӦ��ΪNULL, lpCommandLineָ����ַ���Ӧ��ָ����ִ��ģ�鼰�������Ĭ������£�CreateProcessAsUser����������16λ����windows��Ӧ�ó���������һ��������VDM��(�൱��CreateProcess�е�CREATE_SEPARATE_WOW_VDM)��
'    lpCommandLine _Inout_opt_
'        Ҫִ�е������С�����ַ�������󳤶���32K���ַ������lpApplicationNameΪNULL����lpCommandLine��ģ������������ΪMAX_PATH�ַ���
'        ���������Unicode�汾CreateProcessAsUserW�����޸�����ַ��������ݡ���ˣ��ò���������ָ��ֻ���ڴ��ָ��(����const�����������ַ���)������ò����ǳ����ַ�������ú������ܻᵼ�·��ʳ�ͻ��
'        lpCommandLine��������Ϊ�ա�����������£�����ʹ��lpApplicationNameָ����ַ�����Ϊ�����С�
'        ���lpApplicationName��lpCommandLine���Ƿǿյģ���*lpApplicationNameָ��Ҫִ�е�ģ�飬*lpCommandLineָ�������С��½��̿���ʹ��GetCommandLine�������������С���C��д�Ŀ���̨���̿���ʹ��argc��argv�������������С���Ϊargv[0]��ģ������C����Աͨ����ģ������Ϊ�������еĵ�һ�������ظ���
'        ���lpApplicationNameΪNULL�����������е�һ���Կո�ָ��ı��ָ��ģ�����ơ����ʹ�ð����ո�ĳ��ļ�������ʹ�ô����ŵ��ַ�����ָʾ�ļ����Ľ����Ͳ����Ŀ�ʼλ��(�����lpApplicationName������˵��)������ļ�����������չ������׷��.exe����ˣ�����ļ�����չ���ǡ�com������������������com��չ��������ļ�����û����չ���ľ��(.)��β�������ļ�������·�����򲻸���.exe������ļ���������Ŀ¼·����ϵͳ������˳��������ִ���ļ�:
'            ����Ӧ�ó����Ŀ¼
'            �����̵ĵ�ǰĿ¼
'            32 λWindowsϵͳĿ¼��ʹ��GetSystemDirectory������ȡ��Ŀ¼��·��
'            16λWindowsϵͳĿ¼��û�к������Ի�����Ŀ¼��·�������ǿ�����������
'            WindowsĿ¼��ʹ��GetWindowsDirectory������ȡ��Ŀ¼��·��
'            PATH�����������г���Ŀ¼��ע�⣬�˺�����������App Paths registry��ָ����ÿ��Ӧ�ó���·����Ҫ�����������а���ÿ��Ӧ�ó����·������ʹ��ShellExecute������
'        ϵͳ��һ�����ַ���ӵ��������ַ����У��Խ��ļ���������ֿ����⽫ԭʼ�ַ����ֳ������ַ��������ڲ�����
'    lpProcessAttributes _In_opt_
'        ָ��SECURITY_ATTRIBUTES�ṹ��ָ�룬�ýṹΪ�½��̶���ָ����ȫ����������ȷ���ӽ����Ƿ���Լ̳з��ظ����̵ľ�������lpProcessAttributesΪNULL��lpSecurityDescriptorΪNULL�������̽����Ĭ�ϵİ�ȫ�����������Ҳ��ܼ̳о����Ĭ�ϵİ�ȫ����������hToken���������õ��û��İ�ȫ���������˰�ȫ���������ܲ�������÷����ʣ�����������£��������к���ܲ����ٴδ򿪡����̾������Ч�ģ����ҽ�����ӵ����ȫ�ķ���Ȩ�ޡ�
'    lpThreadAttributes _In_opt_
'        ָ��SECURITY_ATTRIBUTES�ṹ��ָ�룬�ýṹΪ���̶߳���ָ����ȫ����������ȷ���ӽ����Ƿ���Խ����صľ���̳и��̡߳����lpThreadAttributesΪNULL��lpSecurityDescriptorΪNULL���߳̽����һ��Ĭ�ϵİ�ȫ�����������Ҳ��ܼ̳о����Ĭ�ϵİ�ȫ����������hToken���������õ��û��İ�ȫ���������˰�ȫ���������ܲ�������÷����ʡ�
'    bInheritHandles  _In_
'        ����˲���Ϊ�棬����������е�ÿ���ɼ̳о�����������̼̳С��������ΪFALSE���򲻻�̳о����ע�⣬�̳еľ��������ԭʼ�����ͬ��ֵ�ͷ���Ȩ�ޡ�
'        �ն˷���:���ܿ�Ự�̳о�������⣬����ò���Ϊ�棬������ڵ��������ڵĻỰ�д������̡�
'        �ܱ����Ľ�����(PPL)����:��PPL���̴�����PPL����ʱ������PROCESS_DUP_HANDLE������ӷ�PPL���̵�PPL���̣�����һ�����̳б��������μ����̰�ȫ�ͷ���Ȩ��
'    dwCreationFlags _In_
'        �������ȼ���ͽ��̴����ı�־���й�ֵ�б���μ����̴�����־��
'        �˲����������½��̵����ȼ��࣬��������ȷ�������̵߳ĵ������ȼ����й�ֵ�б���μ�GetPriorityClass�����û��ָ���κ����ȼ����־�������ȼ���Ĭ��ΪNORMAL_PRIORITY_CLASS�����Ǵ������̵����ȼ�����IDLE_PRIORITY_CLASS�����idle_normal_priority_class���ڱ����У��ӽ��̽��յ��ý��̵�Ĭ�����ȼ��ࡣ
'    lpEnvironment _In_opt_
'        ָ���½��̵Ļ������ָ�롣����ò���ΪNULL�����½��̽�ʹ�õ��ý��̵Ļ�����
'        ����������null��β���ַ�����ɵ���null��β�Ŀ顣ÿ���ַ�������ʽ����:
'            Name = ֵ \ 0
'        ��Ϊ�Ⱥ������ָ��������Բ����ڻ���������������ʹ������
'        ��������԰���Unicode��ANSI�ַ������lpEnvironmentָ��Ļ��������Unicode�ַ�����ȷ��dwCreationFlags����CREATE_UNICODE_ENVIRONMENT������ò���ΪNULL�����Ҹ����̵Ļ��������Unicode�ַ����򻹱���ȷ��dwCreationFlags����CREATE_UNICODE_ENVIRONMENT��
'        ������̵Ļ�������ܴ�С����32,767���ַ�����˺�����ANSI�汾CreateProcessAsUserA��ʧ�ܡ�
'        ע�⣬ANSI���������������ֽ���ֹ:һ���������һ���ַ�������һ��������ֹ�ÿ顣Unicode��������ֹΪ4�����ֽ�:���һ���ַ�����ֹΪ2���ֽڣ����һ���ַ�����ֹΪ2���ֽڡ�
'        Windows Server 2003��Windows XP:����ϲ����û���ϵͳ���������Ĵ�С����8192�ֽڣ�CreateProcessAsUser�����Ľ��̽��������и����̴��ݸ������Ļ����顣�෴���ӽ���ʹ��CreateEnvironmentBlock�������صĻ��������С�
'        ҪΪ�����û�����������ĸ�������ʹ��CreateEnvironmentBlock������
'    lpCurrentDirectory _In_opt_
'        ���̵�ǰĿ¼������·�����ַ���������ָ��UNC·��
'        ����ò���ΪNULL�����½��̽���������ý�����ͬ�ĵ�ǰ��������Ŀ¼��(��������Ҫ������Ҫ����Ӧ�ó���ָ�����ʼ�������͹���Ŀ¼��shell��)
'    lpStartupInfo _In_
'        ָ��STARTUPINFO��STARTUPINFOEX�ṹ��ָ�롣
'        �û�������ȫ����ָ���Ĵ���վ�����档�����ϣ�������ǽ���ʽ�ģ���ָ��winsta0\default�����lpDesktop��ԱΪ�գ����½��̼̳��丸���̵�����ʹ���վ������ó�Ա�ǿ��ַ������������½��̽�ʹ�á��������ӵ�����վ���������Ĺ������ӵ�����վ��
'        Ҫ������չ���ԣ���ʹ��STARTUPINFOEX�ṹ������dwCreationFlags������ָ��EXTENDED_STARTUPINFO_PRESENT��
'        ��������ҪSTARTUPINFO��STARTUPINFOEX�еľ��ʱ��������close����ر����ǡ�
'        ��Ҫ���ǣ������߸���ȷ��STARTUPINFO�еı�׼����ֶΰ�����Ч�ľ��ֵ����ʹ��dwFlags��Աָ��startf_usestdhandleʱ����Щ�ֶ�Ҳ�᲻����֤�ظ��Ƶ��ӽ����С�����ȷ��ֵ���ܵ����ӽ�����Ϊ�����������ʹ��Ӧ�ó�����֤��������ʱ��֤���߼����Ч�����
'    lpProcessInformation _Out_
'        ָ��PROCESS_INFORMATION�ṹ��ָ�룬�ýṹ���չ����½��̵ı�ʶ��Ϣ��
'        PROCESS_INFORMATION�еľ�������ڲ�����Ҫʱ��close����ر�
'����ֵ
'    ��������ɹ�������ֵΪ���㡣
'    �������ʧ�ܣ�����ֵΪ�㡣Ҫ��ȡ��չ�Ĵ�����Ϣ�������GetLastError��
'    ע�⣬�����ڽ�����ɳ�ʼ��֮ǰ���ء�����޷��ҵ������DLL���ʼ��ʧ�ܣ����̽���ֹ��Ҫ��ȡ���̵���ֹ״̬�������GetExitCodeProcess��
'��ע��
'    CreateProcessAsUser�����ܹ�ʹ��TOKEN_DUPLICATE��TOKEN_IMPERSONATE����Ȩ�޴򿪵��ý��̵������ơ�
'    Ĭ������£�CreateProcessAsUser�ڷǽ���ʽ����վ�ϴ����½��̣����治�ɼ���Ҳ���ܽ����û����롣Ҫ�����û����½��̵Ľ�����������STARTUPINFO�ṹ��lpDesktop��Ա��ָ��Ĭ�ϵĽ�������վ����������ơ�winsta0\default�������⣬�ڵ���CreateProcessAsUser֮ǰ���������Ĭ�Ͻ�������վ��Ĭ������Ŀ�����֧����ʿ����б�(discretionary access control list, DACL)������վ�������DACLs����������û�����hToken������ʾ�ĵ�¼�Ự�ķ���Ȩ��
'    CreateProcessAsUser����ָ���û��ĸ�Ҫ�ļ����ص�HKEY_USERSע������С���ˣ�Ҫ����HKEY_CURRENT_USERע������е���Ϣ���ڵ���CreateProcessAsUser֮ǰ������ʹ��LoadUserProfile�������û��ĸ�Ҫ��Ϣ���ص�HKEY_USERS�С�ȷ�����½����˳������UnloadUserProfile��
'    ���lpEnvironment����ΪNULL�����½��̽��̳е��ý��̵Ļ�����CreateProcessAsUser�����Զ��޸Ļ������԰����ض�����hToken��ʾ���û��Ļ������������磬���lpEnvironmentΪ�գ���ӵ��ý��̼̳�USERNAME��USERDOMAIN����������ְ����Ϊ������׼�������鲢��lpEnvironment��ָ������
'    CreateProcessWithLogonW��CreateProcessWithTokenW����������CreateProcessAsUser��ֻ�ǵ����߲���Ҫ����LogonUser���������û����������֤����ȡ���ơ�
'    CreateProcessAsUser���������ʵ����߻�Ŀ���û��İ�ȫ��������ָ����Ŀ¼�Ϳ�ִ��ӳ��Ĭ������£�CreateProcessAsUser���ʵ����߰�ȫ�������е�Ŀ¼�Ϳ�ִ��ӳ������������£����������û�з���Ŀ¼�Ϳ�ִ��ӳ���Ȩ�ޣ�������ʧ�ܡ�Ҫʹ��Ŀ���û��İ�ȫ�����ķ���Ŀ¼�Ϳ�ִ��ӳ�����ڵ���CreateProcessAsUser֮ǰ���ڵ���ImpersonateLoggedOnUser����ʱָ��hToken��
'    ���̱�����һ�����̱�ʶ������ʶ���ڽ�����ֹ֮ǰ��Ч��������������ʶ���̣�������OpenProcess������ָ���򿪽��̾���������еĳ�ʼ�߳�Ҳ������һ���̱߳�ʶ����������OpenThread������ָ���������̵߳ľ������ʶ�����߳���ֹ֮ǰ����Ч�ģ����ҿ�������Ωһ�ر�ʶϵͳ�е��̡߳���Щ��ʶ����PROCESS_INFORMATION�ṹ�з��ء�
'    �����߳̿���ʹ��WaitForInputIdle�����ȴ���ֱ���½�����ɳ�ʼ�����������ڵȴ��û������û�������������ڸ����̺��ӽ���֮���ͬ���ǳ����ã���ΪCreateProcessAsUser����ʱ����Ҫ�ȴ��½�����ɳ�ʼ�������磬�������̽��ڳ��Բ����������̹����Ĵ���֮ǰʹ��WaitForInputIdle��
'    �رս��̵���ѡ������ʹ��ExitProcess��������Ϊ�ú����򸽼ӵ����̵�����dll������ֹ֪ͨ�������رս��̵ķ�����֪ͨ���ӵ�dll��ע�⣬��һ���̵߳���ExitProcessʱ�����̵������߳̽�����ֹ����û�л���ִ���κ���������(��������dll���߳���ֹ����)���йظ�����Ϣ����μ���ֹ���̡�
'��ȫ��ע
'    lpApplicationName��������ΪNULL������������£���ִ�����Ʊ�����lpCommandLine�е�һ���ո�ָ����ַ����������ִ���ļ���·���������пո������ں��������ո�ķ�ʽ�����ܻ����в�ͬ�Ŀ�ִ���ļ�����������Ӻ�Σ�գ���Ϊ�������������С�Program��������������ڣ������ǡ�MyApp.exe����
'       LPTSTR szCmdline[] = _tcsdup(TEXT("C:\\Program Files\\MyApp"));
'       CreateProcessAsUser(hToken, NULL, szCmdline�� /*��* /);
'    ��������û�Ҫ����һ����Ϊ�����򡱵�Ӧ�ó�����ϵͳ�ϣ��κ�ʹ�ó����ļ�Ŀ¼�������CreateProcessAsUser�ĳ��򶼽����д�Ӧ�ó��򣬶�����Ԥ�ڵ�Ӧ�ó���
'    Ϊ�˱���������⣬��ҪΪlpApplicationName����NULL�����ȷʵΪlpApplicationName����NULL������lpCommandLine�еĿ�ִ��·����Χʹ�����ţ��������ʾ����ʾ��
'       LPTSTR szCmdline[] = _tcsdup(TEXT("\"C:\\Program Files\\MyApp\""));
'       CreateProcessAsUser(hToken, NULL, szCmdline�� /*��*/);
'    PowerShell:����PowerShell 2.0�汾��ʹ��CreateProcessAsUser����ʵ��cmdletʱ��cmdlet����������ȳ�Զ�̻Ự��������ȷ�����С����ǣ�����ĳЩ��ȫ������ʹ��CreateProcessAsUserʵ�ֵ�cmdletֻ����PowerShell version 3.0��Ϊ����Զ�̻Ự��ȷ����;�ȳ�Զ�̻Ự�����ڿͻ�����ȫ��Ȩ�����ʧ�ܡ�Ҫ��PowerShell 3.0�汾��ʵ��һ��ͬʱ������������ȳ�Զ�̻Ự��cmdlet����ʹ��CreateProcess������
'֧��
'    ���֧�ֿͻ�
'       Windows XP [ֻ����������Ӧ�ó�ʽ]
'    ���֧�ַ�����
'       Windows Server 2003[ֻ����������Ӧ�ó���]
'    Header
'       Winbase.h (����Windows.h)
'    Library
'       advapi32.lib
'    dll
'       advapi32.dll
'    Unicode��ANSI����
'       CreateProcessAsUserW (Unicode) and CreateProcessAsUserA (ANSI)
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
'@ԭ��
'    BOOL WINAPI CloseHandle(
'      _In_ HANDLE hObject
'    );
'@����
'    �ر�һ���򿪵Ķ�������
'@����
'    *hObject:
'    һ����Ч�Ĵ򿪵Ķ�������
'@����ֵ
'    ��������ɹ����򷵻�ֵ�Ƿ�0��
'    �������ʧ�ܣ��򷵻�ֵΪ0��Ҫ�����չ������Ϣ�������GetLastError��
'    ���Ӧ�ó����ڵ����������У���ô�ú������׳�һ���쳣����������յ��Ĳ�����Ч�ľ��ֵ��α���ֵ������ر�һ��������Σ�����CloseHandle�رյ���FindFirstFile�������صľ���������ǵ���FindClose�������ͻᷢ�����������
'@��ע
'    CloseHandle�ر����¶�����:
'        Access token ��������
'        Communications device ͨѶ�豸
'        Console input ����̨����
'        Console screen buffer ����̨��Ļ������
'        Event �¼�()
'        File �ļ�
'        File mapping �ļ�ӳ��
'        I/O completion port I / O��ɶ˿�
'        Job ����
'        Mailslot �ʲ�
'        Memory resource notification �ڴ���Դ��֪ͨ
'        Mutex ������
'        Named Pipe �����ܵ�
'        Pipe �ܵ�
'        Process ����
'        Semaphore �ź���
'        Thread �߳�
'        Transaction ����
'        Waitable Timer �ɵȴ���ʱ��
'    ������Щ����ĺ������ĵ�������������ɸö���ʱ��Ӧ��ʹ��CloseHandle���Լ��ڸþ���رպ�Զ���Ĵ���������ᷢ��ʲô�����
'        ͨ���� CloseHandle���ָ���Ķ�����ʧЧ���Զ���ľ���������еݼ�����ִ�ж�������顣
'        ����������һ��������رպ󣬶��󽫱���ϵͳ��ɾ�����й���Щ����Ĵ����ߺ�����ժҪ�������Kernel Objects.��
'    ͨ����Ӧ�ó���Ӧ��Ϊ���򿪵�ÿ���������һ�� CloseHandle��
'        ���ʹ�þ���ĺ���ʧ�ܲ�����ERROR_INVALID_HANDLE����ôͨ��û�б�Ҫ����CloseHandle����Ϊ�������ͨ����������Ѿ�ʧЧ��
'        Ȼ����һЩ����ʹ��ERROR_INVALID_HANDLE��ָʾ����������Ч��
'        ���磬����������ӱ��жϣ���ôһ����ͼ��������ʹ�þ���ĺ���ʧ�ܲ�����ERROR_INVALID_HANDLE ����Ϊ���ļ������ٿ��á�����������£�Ӧ�ó���Ӧ�ùرվ����
'    ���һ�������������ô�������ύ֮ǰ�����а󶨵�����ľ����Ӧ�ùرա�
'        ���һ��������ͨ��ʹ��FILE_FLAG_DELETE_ON_CLOSE��־����CreateFileTransacted �������򿪣���ô��Ӧ�ó���رվ���͵��� CommitTransaction֮ǰ�����ļ����ᱻɾ����
'        �й��������ĸ�����Ϣ����μ�Working With Transactions.��
'    �ر�һ���߳̾����������ֹ��ص��̣߳�Ҳ����ɾ���̶߳��󡣹ر�һ�����̾����������ֹ��صĽ��̣�Ҳ����ɾ�����̶���
'        Ҫɾ��һ���̶߳�����������ֹ�̣߳�Ȼ��ر��߳������еľ����Ҫ��ø�����Ϣ����μ�Terminating a Thread��
'        Ҫɾ�����̶�����������ֹ���̣�Ȼ��رս��̵����о����Ҫ�˽������Ϣ����μ�Terminating a Process��
'    ��ʹ��file mapping��Ȼ�Ǵ򿪵ģ��ر�һ���ļ�ӳ��ľ��Ҳ���Գɹ���Ҫ�˽������Ϣ�������Closing a File Mapping Object.��
'    ��Ҫʹ��CloseHandle�ر�һ���׽��֡��෴��ʹ��closesocket�����������ͷ����׽��ֹ�����������Դ�������׽��ֶ���ľ����Ҫ�˽������Ϣ�������Socket Closure��
'    ��Ҫʹ��CloseHandle�ر�һ���򿪵�ע�����ľ�����෴��ʹ��RegCloseKey ������CloseHandle ����رն�ע�����ľ�������ǲ��᷵��һ����������ʾ���ʧ�ܡ�
'@Ҫ��
'    Minimum supported client   Windows 2000 Professional [desktop apps | UWP apps]
'    Minimum supported server   Windows 2000 Server [desktop apps | UWP apps]
'    Minimum supported phone    Windows Phone 8
'    Header                     Winbase.h (include Windows.h)
'    Library                    kernel32.lib
'    dll                        kernel32.dll
Private Declare Function GetVersionExA Lib "kernel32.dll" (lpVersionInformation As OSVERSIONINFOEX) As Long
'@ԭ��
'    BOOL WINAPI GetVersionEx(
'      _Inout_ LPOSVERSIONINFO lpVersionInfo
'    );
'@����
'    [GetVersionEx������Windows 8.1֮��İ汾�б��޸Ļ򲻿��á��෴��ʹ�ð汾��������]
'    ����Windows 8.1�ķ�����GetVersionEx API����Ϊ�����˱仯���������ز���ϵͳ�汾��ֵ��GetVersionEx�������ص�ֵ����ȡ����Ӧ�ó������ʾ��ʽ��
'    δ��Windows 8.1��Windows 10����ʾ��Ӧ�ó��򽫷���Windows 8 OS�汾ֵ(6.2)��һ��Ϊ�����Ĳ���ϵͳ�汾��ʾ��Ӧ�ó���GetVersionEx��ʼ�շ���Ӧ�ó�����δ���汾����ʾ�İ汾��Ҫ��ʾWindows 8.1��Windows 10��Ӧ�ó�����ο����Windows��Ӧ�ó���
'@����
'    lpVersionInfo _Inout_
'    ���ղ���ϵͳ��Ϣ��OSVERSIONINFO��OSVERSIONINFOEX�ṹ
'    �ڵ���GetVersionEx����֮ǰ�������ýṹ��dwOSVersionInfoSize��Ա����ָʾ���ݸ��ú��������ݽṹ��
'@����ֵ
'    ��������ɹ�������ֵΪ����ֵ��
'    �������ʧ�ܣ�����ֵΪ�㡣Ҫ��ȡ��չ�Ĵ�����Ϣ�������GetLastError�����ΪOSVERSIONINFO��OSVERSIONINFOEX�ṹ��dwOSVersionInfoSize��Աָ����Чֵ����ú�����ʧ�ܡ�
'@��ע
'    ȷ����ǰ����ϵͳͨ������ȷ���Ƿ�����ض�����ϵͳ���Ե���ѷ�����������Ϊ����ϵͳ�����ڿ����·ַ���DLL������������ԡ�����ʹ��GetVersionEx��ȷ������ϵͳƽ̨��汾�ţ�����������Ա���Ĵ����ԡ��йظ�����Ϣ����μ�����ϵͳ�汾��
'    GetSystemMetrics�����ṩ���ڵ�ǰ����ϵͳ�ĸ�����Ϣ
'    ��Ʒ   ����
'    Windows XP Media Center Edition    SM_MEDIACENTER
'    Windows XP Starter Edition         SM_STARTER
'    Windows XP Tablet PC Edition       SM_TABLETPC
'    Windows Server 2003 R2             SM_SERVERR2
'    Ҫ����ض��Ĳ���ϵͳ�����ϵͳ���ԣ���ʹ��IsOS������GetProductInfo����������Ʒ���͡�
'    Ҫ����Զ�̼�����ϲ���ϵͳ����Ϣ����ʹ��NetWkstaGetInfo������Win32_OperatingSystem WMI���IADsComputer�ӿڵ�OperatingSystem���ԡ�
'    Ҫ����ǰϵͳ�汾������汾���бȽϣ���ʹ��VerifyVersionInfo������������ʹ��GetVersionEx�Լ�ִ�бȽϡ�
'    �������ģʽ��Ч��GetVersionEx��������������ʶ����Ĳ���ϵͳ���ò���ϵͳ���ܲ����Ѱ�װ�Ĳ���ϵͳ�����磬���������ģʽ��Ч��GetVersionEx������ΪӦ�ó�������Զ�ѡ��Ĳ���ϵͳ��
'@Ҫ��
'    Minimum supported client   Windows 2000 Professional [desktop apps | UWP apps]
'    Minimum supported server   Windows 2000 Server [desktop apps | UWP apps]
'    Header                     Winbase.h (include Windows.h)
'    Library                    kernel32.lib
'    dll                        kernel32.dll
'   Unicode and ANSI names      GetVersionExW (Unicode) And GetVersionExA(ANSI)
Private Type OSVERSIONINFO 'for GetVersionEx API call
    dwOSVersionInfoSize     As Long
    dwMajorVersion          As Long
    dwMinorVersion          As Long
    dwBuildNumber           As Long
    dwPlatformId            As Long
    szCSDVersion            As String * 128
End Type

Private Type OSVERSIONINFOEX
    dwOSVersionInfoSize     As Long
    dwMajorVersion          As Long
    dwMinorVersion          As Long
    dwBuildNumber           As Long
    dwPlatformId            As Long
    szCSDVersion            As String * 128      '  Maintenance string for PSS usage
    wServicePackMajor       As Integer 'win2000 only
    wServicePackMinor       As Integer 'win2000 only
    wSuiteMask              As Integer 'win2000 only
    wProductType            As Byte 'win2000 only
    wReserved               As Byte
End Type
'@ԭ��
'    typedef struct _OSVERSIONINFOEX {
'      DWORD dwOSVersionInfoSize;
'      DWORD dwMajorVersion;
'      DWORD dwMinorVersion;
'      DWORD dwBuildNumber;
'      DWORD dwPlatformId;
'      TCHAR szCSDVersion[128];
'      WORD  wServicePackMajor;
'      WORD  wServicePackMinor;
'      WORD  wSuiteMask;
'      BYTE  wProductType;
'      BYTE  wReserved;
'    } OSVERSIONINFOEX, *POSVERSIONINFOEX, *LPOSVERSIONINFOEX;
'@����
'    ��������ϵͳ�汾��Ϣ����Щ��Ϣ�������汾�źʹΰ汾�š������š�ƽ̨��ʶ�����Լ����ڲ�Ʒ�׼��Ͱ�װ��ϵͳ�ϵ����·��������Ϣ���˽ṹ����GetVersionEx��VerifyVersionInfo������
'@��Ա
'    dwOSVersionInfoSize
'       �����ݽṹ�Ĵ�С�����ֽ�Ϊ��λ�����˳�Ա����Ϊsizeof(OSVERSIONINFOEX)��
'    dwMajorVersion
'       ����ϵͳ����Ҫ�汾�š��йظ�����Ϣ����μ���ע��
'    dwMinorVersion
'       ����ϵͳ�Ĵ�Ҫ�汾�š��йظ�����Ϣ����μ���ע��
'    dwBuildNumber
'       ����ϵͳ�Ĺ����š�
'    dwPlatformId
'       ����ϵͳƽ̨�������Ա������VER_PLATFORM_WIN32_NT(2)��
'    szCSDVersion
'       һ����null��β���ַ������硰Service Pack 3������ʾϵͳ�ϰ�װ�����·���������û�а�װ����������ַ���Ϊ�ա�
'    wServicePackMajor
'       ϵͳ�ϰ�װ�����·��������Ҫ�汾�š����磬����Service Pack 3����Ҫ�汾����3�����û�а�װ�κη���������ֵΪ�㡣
'    wServicePackMinor
'       ϵͳ�ϰ�װ�����·�����Ĵ�Ҫ�汾�š����磬����Service Pack 3����Ҫ�汾����0��
'    wSuiteMask
'       һ��λ���룬���ڱ�ʶϵͳ�Ͽ��õĲ�Ʒ�׼��������Ա����������ֵ����ϡ�
Private Const VER_SUITE_BACKOFFICE              As Long = &H4
'    ��װ��Microsoft BackOffice�����
Private Const VER_SUITE_BLADE                   As Long = &H400
'    ��װWindows Server 2003, Web Edition��
Private Const VER_SUITE_COMPUTE_SERVER          As Long = &H4000
'    ��װWindows Server 2003���������Ⱥ�档
Private Const VER_SUITE_DATACENTER              As Long = &H80
'    ��װ��Windows Server 2008�������ġ�Windows Server 2003���������İ汾��Windows 2000�������ķ�������
Private Const VER_SUITE_ENTERPRISE              As Long = &H2
'    ��װWindows Server 2008��ҵ�档Windows Server 2003����ҵ���Windows 2000�߼����������йش�λ��־�ĸ�����Ϣ������ı�ע���֡�
Private Const VER_SUITE_EMBEDDEDNT              As Long = &H40
'    ��װWindows XPǶ��ʽ��
Private Const VER_SUITE_PERSONAL                As Long = &H200
'    ��װ��Windows Vista��ͥ�߼��档Windows Vista��ͥ�������Windows XP��ͥ�档
Private Const VER_SUITE_SINGLEUSERTS            As Long = &H100
'    ֧��Զ�����棬��ֻ֧��һ�������Ự������ϵͳ��Ӧ�÷�����ģʽ�����У��������ô�ֵ��
Private Const VER_SUITE_SMALLBUSINESS           As Long = &H1
'    Microsoft Small Business Server������װ��ϵͳ�ϣ��������Ѿ�������Windows����һ���汾���йش�λ��־�ĸ�����Ϣ������ı�ע���֡�
Private Const VER_SUITE_SMALLBUSINESS_RESTRICTED As Long = &H20
'    Microsoft Small Business Server��װʱʹ�����ϸ�Ŀͻ������֤���йش�λ��־�ĸ�����Ϣ������ı�ע���֡�
Private Const VER_SUITE_STORAGE_SERVER          As Long = &H2000
'    ��װWindows Storage Server 2003 R2��Windows Storage Server 2003��
Private Const VER_SUITE_TERMINAL                As Long = &H10
'    ��װ�ն˷���.���ֵ���Ǳ�����
'    ���������VER_SUITE_TERMINAL����û������VER_SUITE_SINGLEUSERTS����ϵͳ��Ӧ�÷�����ģʽ�����С�
Private Const VER_SUITE_WH_SERVER               As Long = &H8000
'    ��װWindows��ͥ��������
Private Const VER_SUITE_MULTIUSERTS             As Long = &H20000
'    ����AppServerģʽ��
'    wProductType
'       ����ϵͳ���κθ�����Ϣ�������Ա����������ֵ֮һ��
Private Const VER_NT_DOMAIN_CONTROLLER          As Long = &H2
'    ϵͳΪ�������������ϵͳΪWindows Server 2012��Windows Server 2008 R2��Windows Server 2008��Windows Server 2003��Windows 2000 Server��
Private Const VER_NT_SERVER                     As Long = &H3
'    ����ϵͳ��Windows Server 2012��Windows Server 2008 R2��Windows Server 2008��Windows Server 2003��Windows 2000 Server��
'    ע�⣬ͬʱҲ����������ķ�����������ΪVER_NT_DOMAIN_CONTROLLER��������VER_NT_SERVER��
Private Const VER_NT_WORKSTATION                As Long = &H1
'    ����ϵͳΪWindows 8��Windows 7��Windows Vista��Windows XP Professional��Windows XP Home Edition��Windows 2000 Professional��
'    wReserved
'       �����Ա�����ʹ�á�
'��ע
'    �����汾��Ϣ���ǲ������Ե���ѷ������෴����ο��ĵ��˽����Ȥ�����ԡ��й�������ⳣ�ü����ĸ�����Ϣ����μ�����ϵͳ�汾��
'    �����������Ҫ�ض��Ĳ���ϵͳ����ȷ��������Ϊ֧�ֵ���Ͱ汾ʹ�ã�������Ϊһ������ϵͳ��Ʋ��ԡ����������ļ����뽫������δ���汾��Windows�Ϲ�����
'    �±��ܽ���֧�ֵ�Windows�汾���ص�ֵ��ʹ�ñ��Ϊ��Other�������е���Ϣ�����־�����ͬ�汾�ŵĲ���ϵͳ��
'        Operating system    Version number  dwMajorVersion  dwMinorVersion  Other
'        Windows 10                 10.0*       10                  0   OSVERSIONINFOEX.wProductType == VER_NT_WORKSTATION
'        Windows Server 2016        10.0*       10                  0   OSVERSIONINFOEX.wProductType != VER_NT_WORKSTATION
'        Windows 8.1                6.3*        6                   3   OSVERSIONINFOEX.wProductType == VER_NT_WORKSTATION
'        Windows Server 2012 R2     6.3*        6                   3   OSVERSIONINFOEX.wProductType != VER_NT_WORKSTATION
'        Windows 8                  6.2         6                   2   OSVERSIONINFOEX.wProductType == VER_NT_WORKSTATION
'        Windows Server 2012        6.2         6                   2   OSVERSIONINFOEX.wProductType != VER_NT_WORKSTATION
'        Windows 7                  6.1         6                   1   OSVERSIONINFOEX.wProductType == VER_NT_WORKSTATION
'        Windows Server 2008 R2     6.1         6                   1   OSVERSIONINFOEX.wProductType != VER_NT_WORKSTATION
'        Windows Server 2008        6.0         6                   0   OSVERSIONINFOEX.wProductType != VER_NT_WORKSTATION
'        Windows Vista              6.0         6                   0   OSVERSIONINFOEX.wProductType == VER_NT_WORKSTATION
'        Windows Server 2003 R2     5.2         5                   2   GetSystemMetrics(SM_SERVERR2) != 0
'        Windows Home Server        5.2         5                   2   OSVERSIONINFOEX.wSuiteMask & VER_SUITE_WH_SERVER
'        Windows Server 2003        5.2         5                   2   GetSystemMetrics(SM_SERVERR2) == 0
'        Windows XP Professional x64 Edition 5.2    5               2   (OSVERSIONINFOEX.wProductType == VER_NT_WORKSTATION) && (SYSTEM_INFO.wProcessorArchitecture==PROCESSOR_ARCHITECTURE_AMD64)
'        Windows XP                 5.1         5                   1   Not applicable
'        Windows 2000               5.0         5                   0   Not applicable
'        *����������Windows 8.1��Windows 10����ʾ��Ӧ�ó���δ��Windows 8.1��Windows 10����ʾ��Ӧ�ó��򽫷���Windows 8 OS�汾ֵ(6.2)��Ҫ��ʾWindows 8.1��Windows 10��Ӧ�ó�����ο����Windows��Ӧ�ó���
'    ����Ӧ�ý�����VER_SUITE_SMALLBUSINESS��־��ȷ��ϵͳ���Ƿ��Ѿ���װ��Small Business Server����Ϊ�ڰ�װ�˲�Ʒ�׼�ʱ�����˴˱�־��VER_SUITE_SMALLBUSINESS_RESTRICTED��־����������˰�װ������Windows Server��׼�棬VER_SUITE_SMALLBUSINESS_RESTRICTED��־������������ǣ�VER_SUITE_SMALLBUSINESS��־���������á�����ð�װ����һ��������Windows Server Enterprise Edition, VER_SUITE_SMALLBUSINESS��־���������á�
'    �������ģʽ��Ч����OSVERSIONINFOEX�ṹ�����й�ΪӦ�ó�������Զ�ѡ��Ĳ���ϵͳ����Ϣ��
'    Ҫȷ������win32��Ӧ�ó����Ƿ���WOW64�����У������IsWow64Process������Ҫȷ��ϵͳ�Ƿ�����64λ�汾��Windows�������GetNativeSystemInfo������
'    GetSystemMetrics�����ṩ�˹��ڵ�ǰ����ϵͳ�����¸�����Ϣ��
'        Product                            Setting
'        Windows Server 2003 R2             SM_SERVERR2
'        Windows XP Media Center Edition    SM_MEDIACENTER
'        Windows XP Starter Edition         SM_STARTER
'        Windows XP Tablet PC Edition       SM_TABLETPC
'@Requirements
'    Minimum supported client Windows 2000 Professional [desktop apps only]
'    Minimum supported server   Windows 2000 Server [desktop apps only]
'    Header                     Winnt.h (include Windows.h)
'    Unicode and ANSI names     OSVERSIONINFOEXW (Unicode) And OSVERSIONINFOEXA(ANSI)
Private Const VER_PLATFORM_WIN32s = 0
Private Const VER_PLATFORM_WIN32_WINDOWS = 1
Private Const VER_PLATFORM_WIN32_NT = 2
'Private Declare Function IsWindowsXPOrGreater Lib "kernel32.dll" () As Long
'@ԭ��
'    BOOL WINAPI IsWindowsXPOrGreater(void);
'@����
'    ָʾ��ǰ����ϵͳ�汾�Ƿ�ƥ������Windows XP�汾��
'@����
'    �������û�в���
'@����ֵ
'    �����ǰ����ϵͳ�汾ƥ������Windows XP�汾����ΪTrue;����,�ٵġ�
'@��ע
'   �˺��������ֿͻ����ͷ������汾�������ǰOS�汾�ŵ��ڻ���ڵ�����ָ���Ŀͻ����汾���򷵻�true�����磬��iswindowsxpsp3����߰汾�ĵ��ý���Windows Server 2008�Ϸ���true����Ҫ����Windows�ķ������Ϳͻ����汾��Ӧ�ó���Ӧ�õ���IsWindowsServer��
'   ����Windows�������汾��û����Windows�ͻ����汾��������������ʹ��iswindowsversionor�����汾����ȷ�ϡ�
'@Requirements
'Minimum supported client        Windows 2000 Professional [desktop apps only]
'Minimum supported server        Windows 2000 Server [desktop apps only]
'Header                          VersionHelpers.h
'Library                         Kernel32.lib;   Ntdll.lib
'dll                             Kernel32.dll    Ntdll.dll
'Private Declare Function IsWindowsXPSP1OrGreater Lib "kernel32.dll" () As Long
'@ԭ��
'    BOOL WINAPI IsWindowsXPSP1OrGreater(void);
'@����
'    ָʾ��ǰ����ϵͳ�汾�Ƿ�ƥ������Windows XP with Service Pack 1 (SP1)�汾��
'@����
'    �������û�в���
'@����ֵ
'    �����ǰ����ϵͳ�汾��SP1�汾ƥ�䣬�����Windows XP�汾����ΪTrue;����,�ٵġ�
'@��ע
'    �˺��������ֿͻ����ͷ������汾�������ǰOS�汾�ŵ��ڻ���ڵ�����ָ���Ŀͻ����汾���򷵻�true�����磬��iswindowsxpsp3����߰汾�ĵ��ý���Windows Server 2008�Ϸ���true����Ҫ����Windows�ķ������Ϳͻ����汾��Ӧ�ó���Ӧ�õ���IsWindowsServer��
'    ����Windows�������汾��û����Windows�ͻ����汾��������������ʹ��iswindowsversionor�����汾����ȷ�ϡ�
'@Requirements
'Minimum supported client        Windows 2000 Professional [desktop apps only]
'Minimum supported server        Windows 2000 Server [desktop apps only]
'Header                          VersionHelpers.h
'Library                         Kernel32.lib;   Ntdll.lib
'dll                             Kernel32.dll    Ntdll.dll
'Private Declare Function IsWindowsXPSP2OrGreater Lib "kernel32.dll" () As Long
'@ԭ��
'    BOOL WINAPI IsWindowsXPSP2OrGreater(void);
'@����
'    ָʾ��ǰ����ϵͳ�汾�Ƿ�ƥ������Windows XP with Service Pack 2 (SP2)�汾��
'@����
'    �������û�в���
'@����ֵ
'    �����ǰ����ϵͳ�汾��SP2�汾ƥ�䣬�����Windows XP�汾����ΪTrue;����,�ٵġ�
'@��ע
'    �˺��������ֿͻ����ͷ������汾�������ǰOS�汾�ŵ��ڻ���ڵ�����ָ���Ŀͻ����汾���򷵻�true�����磬��iswindowsxpsp3����߰汾�ĵ��ý���Windows Server 2008�Ϸ���true����Ҫ����Windows�ķ������Ϳͻ����汾��Ӧ�ó���Ӧ�õ���IsWindowsServer��
'    ����Windows�������汾��û����Windows�ͻ����汾��������������ʹ��iswindowsversionor�����汾����ȷ�ϡ�
'@Requirements
'Minimum supported client        Windows 2000 Professional [desktop apps only]
'Minimum supported server        Windows 2000 Server [desktop apps only]
'Header                          VersionHelpers.h
'Library                         Kernel32.lib;   Ntdll.lib
'dll                             Kernel32.dll    Ntdll.dll
'Private Declare Function IsWindowsXPSP3OrGreater Lib "kernel32.dll" () As Long
'@ԭ��
'    BOOL WINAPI IsWindowsXPSP3OrGreater(void);
'@����
'    ָʾ��ǰ����ϵͳ�汾�Ƿ�ƥ������Windows XP with Service Pack 3 (SP3)�汾��
'@����
'    �������û�в���
'@����ֵ
'    �����ǰ����ϵͳ�汾��SP3�汾ƥ�䣬�����Windows XP�汾����ΪTrue;����,�ٵġ�
'@��ע
'    �˺��������ֿͻ����ͷ������汾�������ǰOS�汾�ŵ��ڻ���ڵ�����ָ���Ŀͻ����汾���򷵻�true�����磬��iswindowsxpsp3����߰汾�ĵ��ý���Windows Server 2008�Ϸ���true����Ҫ����Windows�ķ������Ϳͻ����汾��Ӧ�ó���Ӧ�õ���IsWindowsServer��
'    ����Windows�������汾��û����Windows�ͻ����汾��������������ʹ��iswindowsversionor�����汾����ȷ�ϡ�
'@Requirements
'Minimum supported client        Windows 2000 Professional [desktop apps only]
'Minimum supported server        Windows 2000 Server [desktop apps only]
'Header                          VersionHelpers.h
'Library                         Kernel32.lib;   Ntdll.lib
'dll                             Kernel32.dll    Ntdll.dll
'Private Declare Function IsWindowsVistaOrGreater Lib "kernel32.dll" () As Long
'@ԭ��
'    BOOL WINAPI IsWindowsVistaOrGreater(void);
'@����
'    ָʾ��ǰ����ϵͳ�汾�Ƿ�ƥ������Windows Vista�汾��
'@����
'    �������û�в���
'@����ֵ
'    �����ǰ����ϵͳ�汾��Windows Vista�汾ƥ�䣬�����Windows Vista�汾����ΪTrue;����,�ٵġ�
'@��ע
'    �˺��������ֿͻ����ͷ������汾�������ǰOS�汾�ŵ��ڻ���ڵ�����ָ���Ŀͻ����汾���򷵻�true�����磬��iswindowsxpsp3����߰汾�ĵ��ý���Windows Server 2008�Ϸ���true����Ҫ����Windows�ķ������Ϳͻ����汾��Ӧ�ó���Ӧ�õ���IsWindowsServer��
'    ����Windows�������汾��û����Windows�ͻ����汾��������������ʹ��iswindowsversionor�����汾����ȷ�ϡ�
'@Requirements
'Minimum supported client        Windows 2000 Professional [desktop apps only]
'Minimum supported server        Windows 2000 Server [desktop apps only]
'Header                          VersionHelpers.h
'Library                         Kernel32.lib;   Ntdll.lib
'dll                             Kernel32.dll    Ntdll.dll
'Private Declare Function IsWindowsVistaSP1OrGreater Lib "kernel32.dll" () As Long
'@ԭ��
'    BOOL WINAPI IsWindowsVistaSP1OrGreater(void);
'@����
'    ָʾ��ǰ����ϵͳ�汾�Ƿ�ƥ������Windows Vista with Service Pack 1 (SP1)�汾��
'@����
'    �������û�в���
'@����ֵ
'    �����ǰ����ϵͳ�汾��Windows Vista SP1�汾ƥ�䣬�����Windows Vista SP1�汾����ΪTrue;����,�ٵġ�
'@��ע
'    �˺��������ֿͻ����ͷ������汾�������ǰOS�汾�ŵ��ڻ���ڵ�����ָ���Ŀͻ����汾���򷵻�true�����磬��iswindowsxpsp3����߰汾�ĵ��ý���Windows Server 2008�Ϸ���true����Ҫ����Windows�ķ������Ϳͻ����汾��Ӧ�ó���Ӧ�õ���IsWindowsServer��
'    ����Windows�������汾��û����Windows�ͻ����汾��������������ʹ��iswindowsversionor�����汾����ȷ�ϡ�
'@Requirements
'Minimum supported client        Windows 2000 Professional [desktop apps only]
'Minimum supported server        Windows 2000 Server [desktop apps only]
'Header                          VersionHelpers.h
'Library                         Kernel32.lib;   Ntdll.lib
'dll                             Kernel32.dll    Ntdll.dll
'Private Declare Function IsWindowsVistaSP2OrGreater Lib "kernel32.dll" () As Long
'@ԭ��
'    BOOL WINAPI IsWindowsVistaSP2OrGreater(void);
'@����
'    ָʾ��ǰ����ϵͳ�汾�Ƿ�ƥ������Windows Vista with Service Pack 2 (SP2)�汾��
'@����
'    �������û�в���
'@����ֵ
'    �����ǰ����ϵͳ�汾��Windows Vista SP2�汾ƥ�䣬�����Windows Vista SP2�汾����ΪTrue;����,�ٵġ�
'@��ע
'    �˺��������ֿͻ����ͷ������汾�������ǰOS�汾�ŵ��ڻ���ڵ�����ָ���Ŀͻ����汾���򷵻�true�����磬��iswindowsxpsp3����߰汾�ĵ��ý���Windows Server 2008�Ϸ���true����Ҫ����Windows�ķ������Ϳͻ����汾��Ӧ�ó���Ӧ�õ���IsWindowsServer��
'    ����Windows�������汾��û����Windows�ͻ����汾��������������ʹ��iswindowsversionor�����汾����ȷ�ϡ�
'@Requirements
'Minimum supported client        Windows 2000 Professional [desktop apps only]
'Minimum supported server        Windows 2000 Server [desktop apps only]
'Header                          VersionHelpers.h
'Library                         Kernel32.lib;   Ntdll.lib
'dll                             Kernel32.dll    Ntdll.dll
'Private Declare Function IsWindows7OrGreater Lib "kernel32.dll" () As Long
'@ԭ��
'    BOOL WINAPI IsWindows7OrGreater(void);
'@����
'    ָʾ��ǰ����ϵͳ�汾�Ƿ�ƥ������Windows 7�汾��
'@����
'    �������û�в���
'@����ֵ
'    �����ǰ����ϵͳ�汾��Windows 7�汾ƥ�䣬�����Windows 7�汾����ΪTrue;����,�ٵġ�
'@��ע
'    �˺��������ֿͻ����ͷ������汾�������ǰOS�汾�ŵ��ڻ���ڵ�����ָ���Ŀͻ����汾���򷵻�true�����磬��iswindowsxpsp3����߰汾�ĵ��ý���Windows Server 2008�Ϸ���true����Ҫ����Windows�ķ������Ϳͻ����汾��Ӧ�ó���Ӧ�õ���IsWindowsServer��
'    ����Windows�������汾��û����Windows�ͻ����汾��������������ʹ��iswindowsversionor�����汾����ȷ�ϡ�
'@Requirements
'Minimum supported client        Windows 2000 Professional [desktop apps only]
'Minimum supported server        Windows 2000 Server [desktop apps only]
'Header                          VersionHelpers.h
'Library                         Kernel32.lib;   Ntdll.lib
'dll                             Kernel32.dll    Ntdll.dll
'Private Declare Function IsWindows7SP1OrGreater Lib "kernel32.dll" () As Long
'@ԭ��
'    BOOL WINAPI IsWindows7SP1OrGreater(void);
'@����
'    ָʾ��ǰ����ϵͳ�汾�Ƿ�ƥ������Windows 7 with Service Pack 1 (SP1)�汾��
'@����
'    �������û�в���
'@����ֵ
'    �����ǰ����ϵͳ�汾��Windows 7 SP1�汾ƥ�䣬�����Windows 7 SP1�汾����ΪTrue;����,�ٵġ�
'@��ע
'    �˺��������ֿͻ����ͷ������汾�������ǰOS�汾�ŵ��ڻ���ڵ�����ָ���Ŀͻ����汾���򷵻�true�����磬��iswindowsxpsp3����߰汾�ĵ��ý���Windows Server 2008�Ϸ���true����Ҫ����Windows�ķ������Ϳͻ����汾��Ӧ�ó���Ӧ�õ���IsWindowsServer��
'    ����Windows�������汾��û����Windows�ͻ����汾��������������ʹ��iswindowsversionor�����汾����ȷ�ϡ�
'@Requirements
'Minimum supported client        Windows 2000 Professional [desktop apps only]
'Minimum supported server        Windows 2000 Server [desktop apps only]
'Header                          VersionHelpers.h
'Library                         Kernel32.lib;   Ntdll.lib
'dll                             Kernel32.dll    Ntdll.dll
'Private Declare Function IsWindows8OrGreater Lib "kernel32.dll" () As Long
'@ԭ��
'    BOOL WINAPI IsWindows8OrGreater(void);
'@����
'    ָʾ��ǰ����ϵͳ�汾�Ƿ�ƥ������Windows 8�汾��
'@����
'    �������û�в���
'@����ֵ
'    �����ǰ����ϵͳ�汾��Windows 8�汾ƥ�䣬�����Windows 8�汾����ΪTrue;����,�ٵġ�
'@��ע
'    �˺��������ֿͻ����ͷ������汾�������ǰOS�汾�ŵ��ڻ���ڵ�����ָ���Ŀͻ����汾���򷵻�true�����磬��iswindowsxpsp3����߰汾�ĵ��ý���Windows Server 2008�Ϸ���true����Ҫ����Windows�ķ������Ϳͻ����汾��Ӧ�ó���Ӧ�õ���IsWindowsServer��
'    ����Windows�������汾��û����Windows�ͻ����汾��������������ʹ��iswindowsversionor�����汾����ȷ�ϡ�
'@Requirements
'Minimum supported client        Windows 2000 Professional [desktop apps only]
'Minimum supported server        Windows 2000 Server [desktop apps only]
'Header                          VersionHelpers.h
'Library                         Kernel32.lib;   Ntdll.lib
'dll                             Kernel32.dll    Ntdll.dll
'Private Declare Function IsWindows8Point1OrGreater Lib "kernel32.dll" () As Long
'@ԭ��
'    BOOL WINAPI IsWindows8Point1OrGreater(void);
'@����
'    ָʾ��ǰ����ϵͳ�汾�Ƿ�ƥ������Windows 8.1�汾������Windows 10, IsWindows8Point1OrGreater����false������Ӧ�ó������һ���嵥�����а���һ�������Բ��֣����а���ָ��Windows 8.1��/��Windows 10��guid��
'@����
'    �������û�в���
'@����ֵ
'    �����ǰ����ϵͳ�汾��Windows 8.1�汾ƥ�䣬�����Windows 8.1�汾����ΪTrue;����,�ٵġ�
'@��ע
'    �˺��������ֿͻ����ͷ������汾�������ǰOS�汾�ŵ��ڻ���ڵ�����ָ���Ŀͻ����汾���򷵻�true�����磬��iswindowsxpsp3����߰汾�ĵ��ý���Windows Server 2008�Ϸ���true����Ҫ����Windows�ķ������Ϳͻ����汾��Ӧ�ó���Ӧ�õ���IsWindowsServer��
'    ����Windows�������汾��û����Windows�ͻ����汾��������������ʹ��iswindowsversionor�����汾����ȷ�ϡ�
'@Requirements
'Minimum supported client        Windows 2000 Professional [desktop apps only]
'Minimum supported server        Windows 2000 Server [desktop apps only]
'Header                          VersionHelpers.h
'Library                         Kernel32.lib;   Ntdll.lib
'dll                             Kernel32.dll    Ntdll.dll
'Private Declare Function IsWindows10OrGreater Lib "kernel32.dll" () As Long
'@ԭ��
'    BOOL WINAPI IsWindows10OrGreater(void);
'@����
'    ָʾ��ǰ����ϵͳ�汾�Ƿ�ƥ������Windows 10�汾������Windows10, IsWindows10OrGreater����false������Ӧ�ó������һ���嵥�����а���һ�������Բ��֣����а���ָ��Windows10��GUID��
'@����
'    �������û�в���
'@����ֵ
'    �����ǰ����ϵͳ�汾��Windows 10�汾ƥ�䣬�����Windows 10�汾����ΪTrue;����,�ٵġ�
'@��ע
'    û����ʾWindows 10��Ӧ�ó��򷵻�false����ʹ��ǰ�Ĳ���ϵͳ�汾��Windows 10��Ҫ��ʾ���Windows 10��Ӧ�ó�����μ����Windows��Ӧ�ó���
'    �˺��������ֿͻ����ͷ������汾�������ǰOS�汾�ŵ��ڻ���ڵ�����ָ���Ŀͻ����汾���򷵻�true�����磬��iswindowsxpsp3����߰汾�ĵ��ý���Windows Server 2008�Ϸ���true����Ҫ����Windows�ķ������Ϳͻ����汾��Ӧ�ó���Ӧ�õ���IsWindowsServer��
'    ����Windows�������汾��û����Windows�ͻ����汾��������������ʹ��iswindowsversionor�����汾����ȷ�ϡ�
'@Requirements
'Minimum supported client        Windows 2000 Professional [desktop apps only]
'Minimum supported server        Windows 2000 Server [desktop apps only]
'Header                          VersionHelpers.h
'Library                         Kernel32.lib;   Ntdll.lib
'dll                             Kernel32.dll    Ntdll.dll
'Private Declare Function IsWindowsServer Lib "kernel32.dll" () As Long
'@ԭ��
'    BOOL WINAPI IsWindowsServer(void);
'@����
'    ָʾ��ǰ����ϵͳ�Ƿ���Windows�������汾����Ҫ����Windows�ķ������Ϳͻ����汾��Ӧ�ó���Ӧ�õ������������ ע�⣬ֻ�е������ṩ�İ汾�������������ʺ����ĳ���ʱ����Ӧ��ʹ�ô˺���
'@����
'    �������û�в���
'@����ֵ
'    �����ǰ����ϵͳ��Windows�������汾����ΪTrue;����,�ٵġ���
'@Requirements
'Minimum supported client        Windows 2000 Professional [desktop apps only]
'Minimum supported server        Windows 2000 Server [desktop apps only]
'Header                          VersionHelpers.h
'Library                         Kernel32.lib;   Ntdll.lib
'dll                             Kernel32.dll    Ntdll.dll
'Private Declare Function IsWindowsVersionOrGreater Lib "kernel32.dll" (ByVal wMajorVersion As Integer, ByVal wMinorVersion As Integer, ByVal wServicePackMajor As Integer) As Long
'@ԭ��
'    BOOL WINAPI IsWindowsVersionOrGreater(
'       WORD wMajorVersion,
'       WORD wMinorVersion,
'       WORD wServicePackMajor
'    );
'@����
'    ��Ҫ���ǣ���Ӧ��ֻ�������ṩ�İ汾�������������ʺ����ĳ���ʱ��ʹ�ô˺�����
'    ָʾ��ǰOS�汾�Ƿ�ƥ�������ṩ�İ汾��Ϣ���˺�������ȷ��Windows�������汾�Ƿ���ͻ����汾�Ź���
'@����
'    wMajorVersion
'       ��Ҫ����ϵͳ�汾��
'    wMinorVersion
'       ��ҪOS�汾��
'    wServicePackMajor
'       ��Ҫ������汾��
'@����ֵ
'    ���ָ���İ汾�뵱ǰWindows����ϵͳ�İ汾ƥ�䣬����ڸð汾����ΪTRUE;����,�ٵġ�
Private Declare Function OpenThreadToken Lib "advapi32.dll" (ByVal ThreadHandle As Long, ByVal DesiredAccess As Long, ByVal OpenAsSelf As Long, TokenHandle As Long) As Long
'@ԭ��
'    BOOL WINAPI OpenThreadToken(
'      _In_  HANDLE  ThreadHandle,
'      _In_  DWORD   DesiredAccess,
'      _In_  BOOL    OpenAsSelf,
'      _Out_ PHANDLE TokenHandle
'    );
'@����
'    OpenThreadToken���������̹߳����ķ������ơ�
'@����
'ThreadHandle _In_
'   �򿪷������Ƶ��̵߳ľ����
'DesiredAccess _In_]
'   ָ���������룬������ָ���������Ƶ�����������͡���Щ����ķ������������Ƶ����ɷ��ʿ����б�(discretionary access control list, DACL)��Э������ȷ�������ܾ���Щ���ʡ�
'   �йط������Ƶķ���Ȩ���б���μ��������ƶ���ķ���Ȩ�ޡ�
'OpenAsSelf _In_
'   ���Ҫ�Խ��̼���ȫ�����Ľ��з��ʼ�飬��ΪTRUE��
'   ���Ҫ�Ե���OpenThreadToken�������̵߳ĵ�ǰ��ȫ�����Ľ��з��ʼ�飬��ΪFALSE��
'   OpenAsSelf��������˺����ĵ������ڵ�����ģ�ⰲȫ��ʶ���������ʱ��ָ���̵߳ķ������ơ�û�д˲����������߳��޷���ָ���߳��ϵķ������ƣ���Ϊ�޷�ʹ��SecurityIdentificationģ�⼶���ִ�м�����
'TokenHandle _Out_
'   ָ�������ָ�룬�ñ��������´򿪵ķ������Ƶľ����
'@����ֵ
'   ��������ɹ�������ֵΪ���㡣
'   �������ʧ�ܣ�����ֵΪ�㡣Ҫ��ȡ��չ�Ĵ�����Ϣ�������GetLastError��������ƾ�������ģ�⼶���򲻻�����ƣ�OpenThreadToken��ERROR_CANT_OPEN_ANONYMOUS����Ϊ����
'@��ע
'   �޷��򿪾�������ģ�⼶������ơ�
'   ͨ������Close����ر�ͨ��TokenHandle�������صķ������ƾ��
'@Requirements
'Minimum supported client        Windows XP [desktop apps | UWP apps]
'Minimum supported server        Windows Server 2003 [desktop apps | UWP apps]
'Header                          Winbase.h (include Windows.h)
'Library                         Advapi32.lib
'dll                             Advapi32.dll
Private Declare Function OpenProcessToken Lib "advapi32.dll" (ByVal ProcessHandle As Long, ByVal DesiredAccess As Long, TokenHandle As Long) As Long
'@ԭ��
'    BOOL WINAPI OpenProcessToken(
'      _In_  HANDLE  ProcessHandle,
'      _In_  DWORD   DesiredAccess,
'      _Out_ PHANDLE TokenHandle
'    );
'@����
'    OpenProcessToken����������̹����ķ������ơ�
'@����
'ProcessHandle  _In_
'   ���̵ľ��������������Ѵ򿪡����̱������PROCESS_QUERY_INFORMATION����Ȩ�ޡ�
'DesiredAccess _In_
'   ָ���������룬������ָ���������Ƶ�����������͡�����Щ����ķ������������Ƶ����ɷ��ʿ����б�(discretionary access control list, DACL)���бȽϣ���ȷ�������ܾ���Щ���ʡ�
'   �йط������Ƶķ���Ȩ���б���μ��������ƶ���ķ���Ȩ�ޡ�
'TokenHandle _Out_
'   ָ������ָ�룬�þ���ں�������ʱ��ʶ�´򿪵ķ������ơ�
'@����ֵ
'   ��������ɹ�������ֵΪ���㡣
'   �������ʧ�ܣ�����ֵΪ�㡣Ҫ��ȡ��չ�Ĵ�����Ϣ�������GetLastError��
'@��ע
'   ͨ������Close����ر�ͨ��TokenHandle�������صķ������ƾ����
'@Requirements
'Minimum supported client        Windows XP [desktop apps | UWP apps]
'Minimum supported server        Windows Server 2003 [desktop apps | UWP apps]
'Header                          Winbase.h (include Windows.h)
'Library                         Advapi32.lib
'dll                             Advapi32.dll
Private Declare Function SetTokenInformation Lib "advapi32.dll" (ByVal TokenHandle As Long, ByVal TokenInformationClass As TOKEN_INFORMATION_CLASS, TokenInformation As Long, ByVal TokenInformationLength As Long) As Long
Private Declare Function GetTokenInformation Lib "advapi32.dll" (ByVal TokenHandle As Long, ByVal TokenInformationClass As TOKEN_INFORMATION_CLASS, TokenInformation As Long, ByVal TokenInformationLength As Long, ReturnLength As Long) As Long
'@ԭ��
'    BOOL WINAPI GetTokenInformation(
'      _In_      HANDLE                  TokenHandle,
'      _In_      TOKEN_INFORMATION_CLASS TokenInformationClass,
'      _Out_opt_ LPVOID                  TokenInformation,
'      _In_      DWORD                   TokenInformationLength,
'      _Out_     PDWORD                  ReturnLength
'    );
'@����
'    GetTokenInformation�����������ڷ������Ƶ�ָ�����͵���Ϣ�����ý��̱�������ʵ��ķ���Ȩ�޲��ܻ����Ϣ
'    Ҫȷ���û��Ƿ����ض���ĳ�Ա����ʹ��CheckTokenMembership������Ҫȷ��Ӧ�ó����������Ƶ����Ա��ϵ����ʹ��CheckTokenMembershipEx������
'@����
'TokenHandle _In_
'   ��ȡ��Ϣ�ķ������Ƶľ�������TokenInformationClassָ����TokenSource������������TOKEN_QUERY_SOURCE����Ȩ��������������TokenInformationClassֵ������������TOKEN_QUERY����Ȩ��
'TokenInformationClass _In_
'   ��TOKEN_INFORMATION_CLASSö������ָ��һ��ֵ���Ա�ʶ������������Ϣ�����͡��κμ��TokenIsAppContainer����������0�ĵ����߻�Ӧ����֤���������Ʋ��Ǳ�ʶ����ģ�����ơ������ǰ���Ʋ���Ӧ�ó������������Ǳ�ʶ�������ƣ���Ӧ���ؾܾ����ʡ�
'TokenInformation _Out_opt_
'   ָ�򻺳�����ָ�룬�ú��������������Ϣ��仺����������˻������Ľṹȡ����TokenInformationClass����ָ������Ϣ���͡�
'TokenInformationLength _In_
'   ָ��TokenInformation����ָ��Ļ������Ĵ�С(���ֽ�Ϊ��λ)�����TokenInformationΪ�գ���˲�������Ϊ�㡣
'ReturnLength _Out_
'   ָ��һ��������ָ�룬�ñ�������TokenInformation����ָ��Ļ�����������ֽ����������ֵ����TokenInformationLength������ָ����ֵ��������ʧ�ܣ����ڻ������в��洢�κ����ݡ�
'   ���TokenInformationClass������ֵ��TokenDefaultDacl��������û��Ĭ�ϵ�DACL��������ReturnLengthָ��ı�������Ϊsizeof(TOKEN_DEFAULT_DACL)������TOKEN_DEFAULT_DACL�ṹ��DefaultDacl��Ա����ΪNULL��
'@����ֵ
'    ��������ɹ�������ֵΪ���㡣
'    �������ʧ�ܣ�����ֵΪ�㡣Ҫ��ȡ��չ�Ĵ�����Ϣ�������GetLastError��
'@Requirements
'Minimum supported client        Windows XP [desktop apps | UWP apps]
'Minimum supported server        Windows Server 2003 [desktop apps | UWP apps]
'Header                          Winbase.h (include Windows.h)
'Library                         Advapi32.lib
'dll                             Advapi32.dll
Private Enum TOKEN_INFORMATION_CLASS
  TokenUser = 1
  TokenGroups
  TokenPrivileges
  TokenOwner
  TokenPrimaryGroup
  TokenDefaultDacl
  TokenSource
  tokenType
  TokenImpersonationLevel
  TokenStatistics
  TokenRestrictedSids
  TokenSessionId
  TokenGroupsAndPrivileges
  TokenSessionReference
  TokenSandBoxInert
  TokenAuditPolicy
  TokenOrigin
  TokenElevationType
  TokenLinkedToken
  TokenElevation
  TokenHasRestrictions
  TokenAccessInformation
  TokenVirtualizationAllowed
  TokenVirtualizationEnabled
  TokenIntegrityLevel
  TokenUIAccess
  TokenMandatoryPolicy
  TokenLogonSid
  TokenIsAppContainer
  TokenCapabilities
  TokenAppContainerSid
  TokenAppContainerNumber
  TokenUserClaimAttributes
  TokenDeviceClaimAttributes
  TokenRestrictedUserClaimAttributes
  TokenRestrictedDeviceClaimAttributes
  TokenDeviceGroups
  TokenRestrictedDeviceGroups
  TokenSecurityAttributes
  TokenIsRestricted
  MaxTokenInfoClass
End Enum
'@����
'    TOKEN_INFORMATION_CLASSö�ٰ���ָ��������������ƻ�ӷ������Ƽ�������Ϣ���͵�ֵ��
'    GetTokenInformation����ʹ����Щֵ��ָʾҪ������������Ϣ�����͡�
'    SetTokenInformation����ʹ����Щֵ����������Ϣ
'@����
'TokenUser
'   ����������һ��TOKEN_USER�ṹ���ýṹ�������Ƶ��û��ʻ���
'TokenGroups
'   ����������һ��TOKEN_GROUPS�ṹ�����а��������ƹ��������ʻ���
'TokenPrivileges
'   ����������һ������������Ȩ��TOKEN_PRIVILEGES�ṹ
'TokenOwner
'   ����������һ��TOKEN_OWNER�ṹ���ýṹ�����´��������Ĭ�������߰�ȫ��ʶ��(SID)��
'TokenPrimaryGroup
'   ����������һ��TOKEN_PRIMARY_GROUP�ṹ�����а����´��������Ĭ������SID��
'TokenDefaultDacl
'   ����������һ��TOKEN_DEFAULT_DACL�ṹ���ýṹ�����´��������Ĭ��DACL��
'TokenSource
'   ����������һ����������Դ��TOKEN_SOURCE�ṹ����������Ϣ��Ҫ����TOKEN_QUERY_SOURCE��
'TokenType
'   ����������һ��TOKEN_TYPEֵ����ֵָʾ�����������ƻ���ģ�����ơ�
'TokenImpersonationLevel
'   ����������SECURITY_IMPERSONATION_LEVELֵ����ֵָʾ���Ƶ�ģ�⼶������������Ʋ���ģ�����ƣ�������ʧ�ܡ�
'TokenStatistics
'   ����������һ��TOKEN_STATISTICS�ṹ�����а�����������ͳ����Ϣ��
'TokenRestrictedSids
'   ����������һ��TOKEN_GROUPS�ṹ�����а���һ��SID���б�
'TokenSessionId
'   ����������һ��DWORDֵ����ֵָʾ�����ƹ������ն˷���Ự��ʶ����
'   ����������ն˷������ͻ����Ự���������Ự��ʶ��Ϊ���㡣
'   Windows Server 2003��Windows XP:����������ն˷���������̨�Ự���������Ự��ʶ��Ϊ�㡣
'   �ڷ��ն˷��񻷾��У��Ự��ʶ��Ϊ�㡣
'   ���TokenSessionId����SetTokenInformation���õģ���ôӦ�ó�����뽫����Ϊ��Ϊ����ϵͳ��Ȩ��һ���֣�����Ӧ�ó�������ܹ������������ûỰID��
'TokenGroupsAndPrivileges
'   ����������һ��TOKEN_GROUPS_AND_PRIVILEGES�ṹ���ýṹ�����û�SID�����ʻ��������Ƶ�SID�������ƹ����������֤ID��
'TokenSessionReference
'   ������
'TokenSandBoxInert
'   ������ư��� SANDBOX_INERT��־��������������һ�������DWORDֵ��
'TokenAuditPolicy
'   ������
'TokenOrigin
'   ����������һ��TOKEN_ORIGINֵ��
'   �����������ʹ����ʽƾ֤�ĵ�¼�����罫���ơ�������봫�ݸ�LogonUser��������ôTOKEN_ORIGIN�ṹ�������������ĵ�¼�Ự��ID��
'   ��������������������֤���µģ��������AcceptSecurityContext�����LogonUser (dwLogonType����ΪLOGON32_LOGON_NETWORK��LOGON32_LOGON_NETWORK_CLEARTEXT)�����ֵΪ�㡣
'TokenElevationType
'   ����������һ��TOKEN_ELEVATION_TYPEֵ����ֵָ�����Ƶı�߼���
'TokenLinkedToken
'   ����������һ��TOKEN_LINKED_TOKEN�ṹ���ýṹ������һ�����ӵ������Ƶ����ƾ����
'TokenElevation
'   ����������һ��TOKEN_ELEVATION�ṹ���ýṹָ���Ƿ��������ơ�
'TokenHasRestrictions
'   �����Ǳ����˹���������������һ�������DWORDֵ��
'TokenAccessInformation
'   ����������һ��TOKEN_ACCESS_INFORMATION�ṹ���ýṹָ�������а����İ�ȫ��Ϣ��
'TokenVirtualizationAllowed
'   �����������ƽ������⻯��������������һ�������DWORDֵ��
'TokenVirtualizationEnabled
'   ���Ϊ�������������⻯��������������һ�������DWORDֵ��
'TokenIntegrityLevel
'   ����������һ��TOKEN_MANDATORY_LABEL�ṹ���ýṹָ�����Ƶ������Լ���
'TokenUIAccess
'   �������������UIAccess��־��������������һ�������DWORDֵ��
'TokenMandatoryPolicy
'   ����������һ��TOKEN_MANDATORY_POLICY�ṹ���ýṹָ�����Ƶ�ǿ�������Բ��ԡ�
'TokenLogonSid
'   ����������һ��TOKEN_GROUPS�ṹ���ýṹָ�����Ƶĵ�¼SID��
'TokenIsAppContainer
'   ���������Ӧ�ó����������ƣ�������������һ�������DWORDֵ���κμ��TokenIsAppContainer����������0�ĵ����߻�Ӧ����֤���������Ʋ��Ǳ�ʶ����ģ�����ơ������ǰ���Ʋ���Ӧ�ó������������Ǳ�ʶ�������ƣ���Ӧ���ؾܾ����ʡ�
'TokenCapabilities
'   ����������һ��TOKEN_GROUPS�ṹ�����а��������ƹ����Ĺ��ܡ�
'TokenAppContainerSid
'   ����������һ��TOKEN_APPCONTAINER_INFORMATION�ṹ���ýṹ���������ƹ�����AppContainerSid�����������Ӧ�ó�������û�й�������TOKEN_APPCONTAINER_INFORMATION�ṹ�е�TokenAppContainer��Աָ��NULL��
'TokenAppContainerNumber
'   ����������һ���������Ƶ�Ӧ�ó��������ŵ�DWORDֵ�����ڲ���Ӧ�ó����������Ƶ����ƣ���ֵΪ�㡣
'TokenUserClaimAttributes
'   ����������һ��CLAIM_SECURITY_ATTRIBUTES_INFORMATION�ṹ���ýṹ���������ƹ������û�������
'TokenDeviceClaimAttributes
'   ����������һ��CLAIM_SECURITY_ATTRIBUTES_INFORMATION�ṹ���ýṹ���������ƹ������豸������
'TokenRestrictedUserClaimAttributes
'   ������ֵ��
'TokenRestrictedDeviceClaimAttributes
'   ������ֵ��
'TokenDeviceGroups
'   ����������һ��TOKEN_GROUPS�ṹ�����а��������ƹ������豸�顣
'TokenRestrictedDeviceGroups
'   ����������һ��TOKEN_GROUPS�ṹ���ýṹ���������ƹ����������豸�顣
'TokenSecurityAttributes
'   ������ֵ��
'TokenIsRestricted
'   ������ֵ��
'MaxTokenInfoClass
'   ��ö�ٵ����ֵ��
'@Requirements
'Minimum supported client        Windows XP [desktop apps only]
'Minimum supported server        Windows Server 2003 [desktop apps only]
'Header                          Winnt.h (include Windows.h)
Private Type SID_AND_ATTRIBUTES
    Sid         As Long
    Attributes  As Long
End Type
'@ԭ��
'    typedef struct _SID_AND_ATTRIBUTES {
'      PSID  Sid;
'      DWORD Attributes;
'    } SID_AND_ATTRIBUTES, *PSID_AND_ATTRIBUTES;
'@����
'    SID_AND_ATTRIBUTES�ṹ��ʾ��ȫ��ʶ��(SID)�������ԡ�SId������Ψһ��ȷ���û���Ⱥ�塣
'@��Ա
'Sid
'   ָ��SID�ṹ��ָ��
'Attributes
'   ָ��SID�����ԡ���ֵ������32��1λ��־�����ĺ���ȡ����SID�Ķ����ʹ�á�
'@��ע
'   ����SID��ʾ��SID�������б�������Ŀǰ�����á����û���ǿ��ִ�е����ԡ�СSID��ָ�����ʹ����Щ���ԡ�SID_AND_ATTRIBUTES�ṹ���Ա�ʾ���Ծ����仯��SID�����磬SID_AND_ATTRIBUTES���ڱ�ʾTOKEN_GROUPS�ṹ�е��顣
'@Requirements
'Minimum supported client        Windows XP [desktop apps only]
'Minimum supported server        Windows Server 2003 [desktop apps only]
'Header                          Winnt.h (include Windows.h)
Private Const ANYSIZE_ARRAY                     As Long = 1
Private Type TOKEN_GROUPS
    GroupCount              As Long
    Groups(ANYSIZE_ARRAY)   As SID_AND_ATTRIBUTES
End Type
'@����
'   TOKEN_GROUPS�ṹ�������������й����鰲ȫ��ʶ��(SIDs)����Ϣ��
'@ԭ��
'typedef struct _TOKEN_GROUPS {
'  DWORD              GroupCount;
'  SID_AND_ATTRIBUTES Groups[ANYSIZE_ARRAY];
'} TOKEN_GROUPS, *PTOKEN_GROUPS;
'@��Ա
'GroupCount
'   ָ�����������е�������
'Groups
'   ָ������һ��sid����Ӧ���Ե�SID_AND_ATTRIBUTES�ṹ���顣
'SID_AND_ATTRIBUTES�ṹ�����Գ�Ա���Ծ�������ֵ��
Private Const SE_GROUP_ENABLED                  As Long = &H4
'   ����SID���з��ʼ�顣��ϵͳִ�з��ʼ��ʱ�������Ӧ����SID�ķ�������ͷ��ʾܾ����ʿ�����(ace)��
'   û�д����Ե�SID�ڷ��ʼ���ڼ佫�����ԣ�����������SE_GROUP_USE_FOR_DENY_ONLY���ԡ�
Private Const SE_GROUP_ENABLED_BY_DEFAULT       As Long = &H2
'   ȱʡ���������SID��
Private Const SE_GROUP_INTEGRITY                As Long = &H20
'   SID��ǿ�Ƶ�������SID��
Private Const SE_GROUP_INTEGRITY_ENABLED        As Long = &H40
'   SID֧��ǿ�������Լ�顣
Private Const SE_GROUP_LOGON_ID                 As Long = &HC0000000
'   SID��һ����¼SID������ʶ��������ƹ����ĵ�¼�Ự��
Private Const SE_GROUP_MANDATORY                As Long = &H1
'   SID����ͨ�����õ���tokengroups�������SE_GROUP_ENABLED���ԡ����ǣ�������ʹ��CreateRestrictedToken������ǿ��SIDת��Ϊ���ܾ�SID��
Private Const SE_GROUP_OWNER                    As Long = &H8
'   SID��ʶһ�����ʻ������Ƶ��û��Ǹ���������ߣ����߿��Խ�SIDָ��Ϊ���ƻ����������ߡ�
Private Const SE_GROUP_RESOURCE                 As Long = &H20000000
'   SID��ʶһ���򱾵��顣
Private Const SE_GROUP_USE_FOR_DENY_ONLY         As Long = &H10
'   �����������У�SID��һ��ֻ�з����ߵ�SID����ϵͳִ�з��ʼ��ʱ�������Ӧ����SID�ķ��ʱ��ܾ���ace;������SID������ʵ�ace��
'   ��������˴����ԣ���δ����SE_GROUP_ENABLED����SID�޷��������á�
'@Requirements
'Minimum supported client        Windows XP [desktop apps only]
'Minimum supported server        Windows Server 2003 [desktop apps only]
'Header                          Winnt.h (include Windows.h)
Private Type SID_IDENTIFIER_AUTHORITY
    value(5)        As Byte
End Type
'@���ܣ�
'    SID_IDENTIFIER_AUTHORITY�ṹ��ʾ��ȫ��ʶ��(SID)�Ķ���Ȩ�ޡ�
'@ԭ��
'    typedef struct _SID_IDENTIFIER_AUTHORITY {
'      BYTE Value[6];
'    } SID_IDENTIFIER_AUTHORITY, *PSID_IDENTIFIER_AUTHORITY;
'@��Ա
'Value
'   ָ��SID����Ȩ�޵�6�ֽ����顣
'@��ע
'    ��ʶ��Ȩ��ֵ��ʶ�䷢SID�Ĵ���Ԥ���������±�ʶ��Ȩ�ޡ�
Private Const security_null_sid_authority       As Long = &H0
Private Const security_world_sid_authority      As Long = &H1
Private Const security_local_sid_authority      As Long = &H2
Private Const security_creator_sid_authority    As Long = &H3
Private Const security_non_unique_authority     As Long = &H4
Private Const security_nt_authority             As Long = &H5
Private Const security_resource_manager_authority As Long = &H9
'    SID�����������Ȩ�޺�����һ����Ա�ʶ��(RID)ֵ��
'@Requirements
'Minimum supported client        Windows XP [desktop apps only]
'Minimum supported server        Windows Server 2003 [desktop apps only]
'Header                          Winnt.h (include Windows.h)
Private Type LUID
    LowPart             As Long
    HighPart            As Long
End Type
'@ԭ��
'    typedef struct _LUID {
'      DWORD LowPart;
'      LONG  HighPart;
'    } LUID, *PLUID;
'@����
'    LUID��һ��64λֵ����֤������������ϵͳ����Ωһ�ġ�ֻ������������ϵͳ֮ǰ�����ܱ�֤����Ψһ��ʶ��(LUID)��Ψһ�ԡ�
'    Ӧ�ó������ʹ�ú����ͽṹ������LUIDֵ��
'@��Ա
'LowPart
'   �ͽ�λ��
'HighPart
'   �߽�λ��
'@Requirements
'Minimum supported client       Windows XP [desktop apps only]
'Minimum supported server       Windows Server 2003 [desktop apps only]
'Header                         Winnt.h (include Windows.h)
Private Type LUID_AND_ATTRIBUTES
    PLUID       As LUID
    Attributes  As Long
End Type
'@ԭ��
'    typedef struct _LUID_AND_ATTRIBUTES {
'      LUID  Luid;
'      DWORD Attributes;
'    } LUID_AND_ATTRIBUTES, *PLUID_AND_ATTRIBUTES;
'����
'    LUID_AND_ATTRIBUTES�ṹ��ʾһ������Ωһ��ʶ��(LUID)�������ԡ�
'@��Ա
'LUID
'   ָ��һ��LUIDֵ��
'Attributes
'   ָ��LUID�����ԡ���ֵ������32��1λ��־�����ĺ���ȡ����LUID�Ķ����ʹ�á�
'@��ע
'   LUID_AND_ATTRIBUTES�ṹ���Ա�ʾ���Ծ����仯��LUID�����統LUID���ڱ�ʾPRIVILEGE_SET�ṹ�е���Ȩʱ����Ȩ��luid��ʾ��������ָʾ��ǰ�Ƿ����û������Ȩ�����ԡ�
'@Requirements
'Minimum supported client       Windows XP [desktop apps only]
'Minimum supported server       Windows Server 2003 [desktop apps only]
'Header                         Winnt.h (include Windows.h)

Private Type TOKEN_PRIVILEGES
    PrivilegeCount              As Long
    Privileges(ANYSIZE_ARRAY)   As LUID_AND_ATTRIBUTES
End Type
'@ԭ��
'    typedef struct _TOKEN_PRIVILEGES {
'      DWORD               PrivilegeCount;
'      LUID_AND_ATTRIBUTES Privileges[ANYSIZE_ARRAY];
'    } TOKEN_PRIVILEGES, *PTOKEN_PRIVILEGES;
'@����
'    TOKEN_PRIVILEGES�ṹ�������ڷ������Ƶ�һ����Ȩ����Ϣ��
'@��Ա
'PrivilegeCount
'   ���������ΪPrivileges�����е���Ŀ����
'Privileges
'   ָ��һ��LUID_AND_ATTRIBUTES�ṹ���顣ÿ���ṹ��������Ȩ��LUID�����ԡ�Ҫ��ȡ��LUID��������Ȩ�����ƣ������LookupPrivilegeName��������LUID�ĵ�ַ��ΪlpLuid������ֵ���ݡ�
'   ��Ҫ���ǣ�����ANYSIZE_ARRAY�ڹ���ͷ�ļ�Winnt.h�ж���Ϊ1��Ҫ�����������Ԫ�ص����飬����Ϊ�ṹ�����㹻���ڴ棬�Կ�������Ԫ�ء�
'Privileges�����Կ���������ֵ����ϡ�
Private Const SE_PRIVILEGE_ENABLED              As Long = &H1
'   ��������Ȩ
Private Const SE_PRIVILEGE_ENABLED_BY_DEFAULT   As Long = &H2
'   Ĭ�������������Ȩ��
Private Const SE_PRIVILEGE_REMOVED              As Long = &H4
'   ����ɾ����Ȩ���й���ϸ��Ϣ����μ�AdjustTokenPrivileges��
Private Const SE_PRIVILEGE_USED_FOR_ACCESS      As Long = &H80000000
'   ����Ȩ���ڷ��ʶ������񡣴˱�־���ڱ�ʶ�ͻ���Ӧ�ó��򴫵ݵ�һ���п��ܰ�������Ҫ��Ȩ�������Ȩ��
'@Requirements
'Minimum supported client       Windows XP [desktop apps only]
'Minimum supported server       Windows Server 2003 [desktop apps only]
'Header                         Winnt.h (include Windows.h)
Private Type TOKEN_OWNER
    Owner   As Long
End Type
'@ԭ��
'    typedef struct _TOKEN_OWNER {
'      PSID Owner;
'    } TOKEN_OWNER, *PTOKEN_OWNER;
'@����
'    TOKEN_OWNER�ṹ������Ӧ�����´��������ȱʡ�����߰�ȫ��ʶ��(SID)��
'@��Ա
'Owner
'    һ��ָ��SID�ṹ��ָ�룬�ýṹ��ʾһ���û������û�����Ϊʹ�ô˷������ƴ������κζ���������ߡ�SID�������������Ѿ����ڵ��û�����SID֮һ��
'@Requirements
'Minimum supported client       Windows XP [desktop apps only]
'Minimum supported server       Windows Server 2003 [desktop apps only]
'Header                         Winnt.h (include Windows.h)
Private Declare Function LookupPrivilegeName Lib "advapi32.dll" Alias "LookupPrivilegeNameA" (ByVal lpSystemName As String, ByRef lpLuid As LUID, ByVal lpName As String, ByRef cchName As Long) As Long
'@ԭ��
'    BOOL WINAPI LookupPrivilegeName(
'      _In_opt_  LPCTSTR lpSystemName,
'      _In_      PLUID   lpLuid,
'      _Out_opt_ LPTSTR  lpName,
'      _Inout_   LPDWORD cchName
'    );
'@����
'lpSystemName _In_opt_
'   ָ����null��β���ַ�����ָ�룬���ַ���ָ��������Ȩ���Ƶ�ϵͳ�����ơ����ָ���˿��ַ������ú����������ڱ���ϵͳ�ϲ�����Ȩ���ơ�
'lpLuid _In_
'   һ��ָ��LUID��ָ�룬ͨ����id����֪��Ŀ��ϵͳ�ϵ���Ȩ��
'lpName _Out_opt_
'   ָ�򻺳�����ָ�룬�û��������ձ�ʾ��Ȩ���Ƶ���null��β���ַ��������磬����ַ��������ǡ�SeSecurityPrivilege����
'cchName _Inout_
'   ָ��һ��������ָ�룬�ñ�����һ��TCHARֵ��ָ��lpName�������Ĵ�С������������ʱ���˲���������Ȩ���Ƶĳ��ȣ���������ֹnull�ַ������lpName����ָ��Ļ�����̫С����˱�����������Ĵ�С��
'@����ֵ
'   ��������ɹ����������ط��㡣
'   �������ʧ�ܣ��������㡣Ҫ��ȡ��չ�Ĵ�����Ϣ�������GetLastError��
'@��ע
'    LookupPrivilegeName����ֻ֧��Winnt.h�ж������Ȩ������ָ������Ȩ���й�ֵ�б���μ���Ȩ������
'@Requirements
'Minimum supported client           Windows XP [desktop apps | UWP apps]
'Minimum supported server           Windows Server 2003 [desktop apps | UWP apps]
'Header                             Winbase.h (include Windows.h)
'Library                            advapi32.lib
'dll                                advapi32.dll
'Unicode and ANSI names             LookupPrivilegeNameW (Unicode) And LookupPrivilegeNameA(ANSI)
'��Ȩ����
'    ��Ȩ�����û��ʻ�����ִ�е�ϵͳ���������͡�����Ա����Ȩ������û������ʻ���ÿ���û�����Ȩ���������û����û����������Ȩ��
'    ��ȡ�͵������������е���Ȩ�ĺ���ʹ�ñ���Ωһ��ʶ��(LUID)��������ʶ��Ȩ��ʹ��LookupPrivilegeValue����ȷ������ϵͳ�϶�Ӧ����Ȩ������LUID��ʹ��LookupPrivilegeName������LUIDת��Ϊ��Ӧ���ַ���������
'    ����ϵͳʹ���±��Description���С�User Right��������ַ�����ʾ��Ȩ������ϵͳ�ڱ��ذ�ȫ����Microsoft�������̨(MMC)����Ԫ���û�Ȩ�޷���ڵ�Ĳ���������ʾ�û�Ȩ���ַ�����
Private Const SE_ASSIGNPRIMARYTOKEN_NAME        As String = "SeAssignPrimaryTokenPrivilege"
'   ��Ҫ�������̵���Ҫ���ơ��û�Ȩ��: �滻���̼����ơ�
Private Const SE_AUDIT_NAME                     As String = "SeAuditPrivilege"
'   ��Ҫ���������־��Ŀ��������Ȩ���谲ȫ���������û�Ȩ��: ���ɰ�ȫ��ơ�
Private Const SE_BACKUP_NAME                    As String = "SeBackupPrivilege"
'   ��Ҫִ�б��ݲ�����������Ȩ����ϵͳ�����ж����ʿ��������κ��ļ���������Ϊ���ļ�ָ���ķ��ʿ����б�(ACL)��ʲô������read֮����κη���������Ȼʹ��ACL����������RegSaveKey��RegSaveKeyExfunctions��Ҫ����Ȩ�������������Ȩ�ޣ����������·���Ȩ��:
'       READ_CONTROL
'       ACCESS_SYSTEM_SECURITY
'       FILE_GENERIC_READ
'       FILE_TRAVERSE
'   �û�Ȩ��: �����ļ���Ŀ¼��
Private Const SE_CHANGE_NOTIFY_NAME             As String = "SeChangeNotifyPrivilege"
'   ��Ҫ�����ļ���Ŀ¼���ĵ�֪ͨ������Ȩ���ᵼ��ϵͳ�������б������ʼ�顣��Ĭ��Ϊ�����û����á��û�Ȩ��: ��·������顣
Private Const SE_CREATE_GLOBAL_NAME             As String = "SeCreateGlobalPrivilege"
'   ��Ҫ���ն˷���Ự�ڼ���ȫ�����ƿռ��д��������ļ�ӳ����󡣹���Ա������ͱ���ϵͳ�ʻ�Ĭ�����ô���Ȩ���û�Ȩ��: ����ȫ�ֶ���
Private Const SE_CREATE_PAGEFILE_NAME           As String = "SeCreatePagefilePrivilege"
'   ��Ҫ������ҳ�ļ����û�Ȩ��: ����ҳ���ļ���
Private Const SE_CREATE_PERMANENT_NAME          As String = "SeCreatePermanentPrivilege"
'   ��Ҫ����һ�����ö����û�Ȩ��: �������ù������
Private Const SE_CREATE_SYMBOLIC_LINK_NAME      As String = "SeCreateSymbolicLinkPrivilege"
'   ��Ҫ�����������ӡ��û�Ȩ��: �����������ӡ�
Private Const SE_CREATE_TOKEN_NAME              As String = "SeCreateTokenPrivilege"
'   ��Ҫ����һ�������ơ��û�Ȩ��: ����һ�����ƶ���������ʹ�á��������ƶ��󡱲��Խ�����Ȩ��ӵ��û��ʻ������⣬����ʹ��Windows api������Ȩ��ӵ�ӵ�еĽ��̡�
'   Windows Server 2003�ʹ���SP1������汾��Windows XP: Windows api���Խ�����Ȩ��ӵ���ӵ�еĽ��̡�
Public Const SE_DEBUG_NAME                     As String = "SeDebugPrivilege"
'   ���ڵ��Ժ͵�����һ���ʻ�ӵ�еĽ��̵��ڴ档�û�Ȩ��: ���Գ���
Private Const SE_ENABLE_DELEGATION_NAME         As String = "SeEnableDelegationPrivilege"
'   Ҫ���û��ͼ�����ʻ����Ϊ���ŵ�ί���ʻ����û�Ȩ��: ����ί�����μ�������û��ʻ���
Private Const SE_IMPERSONATE_NAME               As String = "SeImpersonatePrivilege"
'   ��Ҫģ�¡��û�Ȩ��: �����֤��ģ��ͻ�����
Private Const SE_INC_BASE_PRIORITY_NAME         As String = "SeIncreaseBasePriorityPrivilege"
'   ��Ҫ���ӽ��̵Ļ������ȼ����û�Ȩ��: ���ӵ������ȼ���
Private Const SE_INCREASE_QUOTA_NAME            As String = "SeIncreaseQuotaPrivilege"
'   Ҫ�����ӷ�������̵����û�Ȩ��: �������̵��ڴ���
Private Const SE_INC_WORKING_SET_NAME           As String = "SeIncreaseWorkingSetPrivilege"
'   ��ҪΪ���û������������е�Ӧ�ó����������ڴ档�û�Ȩ��: �������̹�������
Private Const SE_LOAD_DRIVER_NAME               As String = "SeLoadDriverPrivilege"
'   ��Ҫ���ػ�ж���豸���������û�Ȩ��: ���غ�ж���豸��������
Private Const SE_LOCK_MEMORY_NAME               As String = "SeLockMemoryPrivilege"
'   ��Ҫ�����ڴ��е�����ҳ���û�Ȩ��: �����ڴ��е�ҳ�档
Private Const SE_MACHINE_ACCOUNT_NAME           As String = "SeMachineAccountPrivilege"
'   ��Ҫ����һ��������ʻ����û�Ȩ��: ������ӹ���վ��
Private Const SE_MANAGE_VOLUME_NAME             As String = "SeManageVolumePrivilege"
'   ��Ҫ���þ������Ȩ���û�Ȩ��: ������ϵ��ļ���
Private Const SE_PROF_SINGLE_PROCESS_NAME       As String = "SeProfileSingleProcessPrivilege"
'   ��ҪΪ���������ռ�������Ϣ���û�Ȩ��: ���õ����̡�
Private Const SE_RELABEL_NAME                   As String = "SeRelabelPrivilege"
'   ��Ҫ�޸Ķ����ǿ�������Լ����û�Ȩ��: �޸Ķ����ǩ��
Private Const SE_REMOTE_SHUTDOWN_NAME           As String = "SeRemoteShutdownPrivilege"
'   ��Ҫʹ����������ر�ϵͳ���û�Ȩ��: ǿ�ƹر�Զ��ϵͳ��
Private Const SE_RESTORE_NAME                   As String = "SeRestorePrivilege"
'   ��Ҫִ�л�ԭ������������Ȩ����ϵͳ������д���ʿ��������κ��ļ���������Ϊ���ļ�ָ����ACL��ʲô������д֮����κη���������Ȼʹ��ACL�������������⣬����Ȩ���������κ���Ч���û�����SID����Ϊ�ļ��������ߡ�RegLoadKey������Ҫ����Ȩ�������������Ȩ�ޣ����������·���Ȩ��:
'       WRITE_DAC
'       WRITE_OWNER
'       ACCESS_SYSTEM_SECURITY
'       FILE_GENERIC_WRITE
'       FILE_ADD_FILE
'       FILE_ADD_SUBDIRECTORY
'       DELTE
'   �û�Ȩ��: ��ԭ�ļ���Ŀ¼��
Private Const SE_SECURITY_NAME                  As String = "SeSecurityPrivilege"
'   ��Ҫִ������밲ȫ����صĹ��ܣ�������ƺͲ鿴�����Ϣ������Ȩ��������߱�ʶΪ��ȫ���������û�Ȩ��: ������ƺͰ�ȫ��־��
Private Const SE_SHUTDOWN_NAME                  As String = "SeShutdownPrivilege"
'   ��Ҫ�رձ���ϵͳ���û�Ȩ��: �ر�ϵͳ��
Private Const SE_SYNC_AGENT_NAME                As String = "SeSyncAgentPrivilege"
'   ���������Ҫʹ��������Ŀ¼����Э��Ŀ¼ͬ�����񡣴���Ȩʹ�������ܹ���ȡĿ¼�е����ж�������ԣ��������Ƕ���������ϵı�����Ĭ������£������������������ϵĹ���Ա�ͱ���ϵͳ�ʻ���
'   �û�Ȩ��: ͬ��Ŀ¼�������ݡ�
Private Const SE_SYSTEM_ENVIRONMENT_NAME        As String = "SeSystemEnvironmentPrivilege"
'   ��Ҫ�޸�ʹ�����������ڴ�洢������Ϣ��ϵͳ�ķ���ʧ��RAM���û�Ȩ��: �޸Ĺ̼�����ֵ��
Private Const SE_SYSTEM_PROFILE_NAME            As String = "SeSystemProfilePrivilege"
'   ��ҪΪ����ϵͳ�ռ�������Ϣ���û�Ȩ��: �����ļ�ϵͳ���ܡ�
Private Const SE_SYSTEMTIME_NAME                As String = "SeSystemtimePrivilege"
'   ��Ҫ�޸�ϵͳʱ�䡣�û�Ȩ��: ����ϵͳʱ�䡣
Private Const SE_TAKE_OWNERSHIP_NAME            As String = "SeTakeOwnershipPrivilege"
'   �ڲ����������֧�����Ȩ������»�ö��������Ȩ������Ȩ�������������ֵ����Ϊ��������Ϊ���������ߺϷ������ֵ���û�Ȩ��: ��ȡ�ļ����������������Ȩ��
Private Const SE_TCB_NAME                       As String = "SeTcbPrivilege"
'   ����Ȩ��������߱�ʶΪ���ż�������һ���֡�һЩ�����ε��ܱ�����ϵͳ���������Ȩ���û�Ȩ��: ��Ϊ����ϵͳ��һ���֡�
Private Const SE_TIME_ZONE_NAME                 As String = "SeTimeZonePrivilege"
'   ��Ҫ�����������ڲ�ʱ����ص�ʱ�����û�Ȩ��: ����ʱ����
Private Const SE_TRUSTED_CREDMAN_ACCESS_NAME    As String = "SeTrustedCredManAccessPrivilege"
'   ��Ҫ�Կ��ŵ����ߵ���ݷ���ƾ�ݹ��������û�Ȩ��: ��Ϊ�����εĵ����߷���ƾ�ݹ�������
Private Const SE_UNDOCK_NAME                    As String = "SeUndockPrivilege"
'   ��Ҫ�򿪱ʼǱ����ԡ��û�Ȩ��: ��������ӶԽӿ��ƿ���
Private Const SE_UNSOLICITED_INPUT_NAME         As String = "SeUnsolicitedInputPrivilege"
'   Ҫ����ն��豸��ȡ���������롣�û�Ȩ��: �����á�
'@Requirements
'Minimum supported client           Windows XP [desktop apps | UWP apps]
'Minimum supported server           Windows Server 2003 [desktop apps | UWP apps]
'Header                             Winbase.h (include Windows.h)
Private Declare Function LookupPrivilegeValue Lib "advapi32.dll" Alias "LookupPrivilegeValueA" (ByVal lpSystemName As String, ByVal lpName As String, lpLuid As LUID) As Long
'@ԭ��
'    BOOL WINAPI LookupPrivilegeValue(
'      _In_opt_ LPCTSTR lpSystemName,
'      _In_     LPCTSTR lpName,
'      _Out_    PLUID   lpLuid
'    );
'@����
'    LookupPrivilegeValue��������ָ��ϵͳ��ʹ�õı���Ωһ��ʶ��(LUID)�����ڱ��ر�ʾָ������Ȩ���ơ�
'@����
'lpSystemName _In_opt_
'   ָ����null��β���ַ�����ָ�룬���ַ���ָ��������Ȩ���Ƶ�ϵͳ�����ơ����ָ���˿��ַ������ú����������ڱ���ϵͳ�ϲ�����Ȩ���ơ�
'lpName  _In_
'   ָ����null��β���ַ�����ָ�룬���ַ���ָ����Ȩ�����ƣ���Winnt.hͷ�ļ��ж�������������磬�����������ָ������SE_SECURITY_NAME����������Ӧ���ַ�����SeSecurityPrivilege����
'lpLuid  _Out_
'   һ��ָ�������ָ�룬�ñ�������LUID, lpSystemName����ָ����ϵͳ�Ͽ���ͨ��LUID֪����Ȩ��
'@����ֵ
'   ��������ɹ����������ط��㡣
'   �������ʧ�ܣ��������㡣Ҫ��ȡ��չ�Ĵ�����Ϣ�������GetLastError��
'@��ע
'LookupPrivilegeValue����ֻ֧��Winnt.h�ж������Ȩ������ָ������Ȩ���й�ֵ�б���μ���Ȩ������
'@Requirements
'Minimum supported client           Windows XP [desktop apps | UWP apps]
'Minimum supported server           Windows Server 2003 [desktop apps | UWP apps]
'Header                             Winbase.h (include Windows.h)
'Library                            advapi32.lib
'dll                                advapi32.dll
'Unicode and ANSI names             LookupPrivilegeValueW  (Unicode) And LookupPrivilegeValueA(ANSI)
Private Declare Function AdjustTokenPrivileges Lib "advapi32.dll" (ByVal TokenHandle As Long, ByVal DisableAllPrivileges As Long, NewState As TOKEN_PRIVILEGES, ByVal BufferLength As Long, PreviousState As TOKEN_PRIVILEGES, ReturnLength As Long) As Long
'@ԭ��
'    BOOL WINAPI AdjustTokenPrivileges(
'      _In_      HANDLE            TokenHandle,
'      _In_      BOOL              DisableAllPrivileges,
'      _In_opt_  PTOKEN_PRIVILEGES NewState,
'      _In_      DWORD             BufferLength,
'      _Out_opt_ PTOKEN_PRIVILEGES PreviousState,
'      _Out_opt_ PDWORD            ReturnLength
'    );
'@����
'    ���û����ָ�����������е���Ȩ���ڷ������������û������Ȩ��Ҫtoken_����_privileges���ʡ�
'@����
'TokenHandle _In_
'   �������Ƶľ�������а���Ҫ�޸ĵ���Ȩ�����������ж����Ƶ�token__privileges����Ȩ�����PreviousState��������NULL��������������TOKEN_QUERY����Ȩ��
'DisableAllPrivileges _In_
'   ָ�������Ƿ�������Ƶ�������Ȩ�������ֵΪ�棬����������������Ȩ������NewState���������ΪFALSE����������NewState����ָ�����Ϣ�޸���Ȩ��
'NewState   _In_opt_
'   ָ��TOKEN_PRIVILEGES�ṹ��ָ�룬�ýṹָ��Ȩ�����鼰�����ԡ����DisableAllPrivileges����ΪFALSE������tokenprivileges���������á����û�ɾ�����Ƶ���Щ��Ȩ���±������˵���tokenprivileges����������Ȩ��������ȡ�Ĳ�����
'   SE_PRIVILEGE_ENABLED
'       �ú���������Ȩ?
'   SE_PRIVILEGE_REMOVED
'       �������е���Ȩ�б���ɾ����Ȩ?�б��е�������Ȩ�����������Ա�������?
'   SE_PRIVILEGE_REMOVEDȡ��SE_PRIVILEGE_ENABLED?
'       ��Ϊ��Ȩ�Ѿ���������ɾ������������������Ȩ�ĳ��Իᵼ�¾���ERROR_NOT_ALL_ASSIGNED���ͺ�����Ȩ��δ���ڹ�һ����
'       ��ͼɾ�������в����ڵ���Ȩ�᷵��ERROR_NOT_ALL_ASSIGNED?
'       ��Ȩ���ɾ������Ȩ������STATUS_PRIVILEGE_NOT_HELD?��Ȩ������ʧ�ܽ���������?
'       ɾ����Ȩ�ǲ�����ģ�����ڵ���AdjustTokenPrivileges֮����ɾ����Ȩ�����Ʋ�������PreviousState�����С�
'       ����SP1��Windows XP: �ú�������ɾ����Ȩ?��֧�ִ�ֵ?
'   None    �ú���������Ȩ?
'   ���DisableAllPrivilegesΪ�棬���������Դ˲�����
'BufferLength _In_
'   ָ����PreviousState����ָ��Ļ������Ĵ�С(���ֽ�Ϊ��λ)�����ǰ״̬����ΪNULL����˲�������Ϊ�㡣
'PreviousState _Out_opt_
'   ָ�򻺳�����ָ�룬�û�������TOKEN_PRIVILEGES�ṹ��䣬�ýṹ�����ú����޸ĵ��κ���Ȩ��ǰһ״̬��Ҳ����˵�����һ����Ȩ�Ѿ�����������޸Ĺ�����ô��Ȩ����֮ǰ��״̬�Ͱ�������PreviousState���õ�TOKEN_PRIVILEGES�ṹ�С����TOKEN_PRIVILEGES�е�ecount��ԱΪ�㣬��˺�����������κ���Ȩ���������������NULL��
'   ���ָ���Ļ�����̫С���޷��������������޸���Ȩ�б�������ʧ�ܣ����Ҳ������κ���Ȩ���ڱ����У�������ReturnLength����ָ��ı�������Ϊ�����������޸���Ȩ�б�������ֽ�����
'ReturnLength _Out_opt_
'   һ��ָ�������ָ�룬�ñ������յ�PreviousState������ָ��Ļ������������С(���ֽ�Ϊ��λ)�����ǰ״̬Ϊ�գ���˲�������Ϊ�ա�
'@����ֵ
'   ��������ɹ�������ֵΪ���㡣Ҫȷ�������Ƿ����������ָ������Ȩ�������GetLastError���������ɹ�ʱ��������������ֵ֮һ:
'   ���ش�������
'ERROR_SUCCESS
'   �ú�������������ָ������Ȩ?
Private Const ERROR_NOT_ALL_ASSIGNED            As Long = &H514
'   ����û��NewState������ָ����һ��������Ȩ����ʹû�е�����Ȩ������Ҳ���ܳɹ���ʹ���������ֵ��PreviousState������ʾ�ѵ�������Ȩ��
'    �������ʧ�ܣ�����ֵΪ�㡣Ҫ��ȡ��չ�Ĵ�����Ϣ�������GetLastError��
'@��ע
'   AdjustTokenPrivileges����������������������Ȩ�ޡ���ֻ�����û�������Ƶ�������Ȩ��Ҫȷ�����Ƶ���Ȩ�������GetTokenInformation������
'   NewState��������ָ������û�е���Ȩ�������ᵼ�º���ʧ�ܡ��ڱ����У�������������ȷʵ���е���Ȩ��������������Ȩ���Ա㺯���ɹ�������GetLastError������ȷ���ú����Ƿ����������ָ������Ȩ��PreviousState������ʾ�ѵ�������Ȩ��
'   PreviousState��������һ��TOKEN_PRIVILEGES�ṹ���ýṹ�����������Ȩ�޵�ԭʼ״̬��Ҫ�ָ�ԭʼ״̬���ں�������AdjustTokenPrivileges����ʱ����PreviousStateָ����ΪNewState�������ݡ�
'@Requirements
'Minimum supported client           Windows XP [desktop apps | UWP apps]
'Minimum supported server           Windows Server 2003 [desktop apps | UWP apps]
'Header                             Winbase.h (include Windows.h)
'Library                            advapi32.lib
'dll                                advapi32.dll
Private Declare Function AllocateAndInitializeSid Lib "advapi32.dll" (pIdentifierAuthority As SID_IDENTIFIER_AUTHORITY, ByVal nSubAuthorityCount As Byte, ByVal dwSubAuthority0 As Long, ByVal dwSubAuthority1 As Long, _
                            ByVal dwSubAuthority2 As Long, ByVal dwSubAuthority3 As Long, ByVal dwSubAuthority4 As Long, ByVal dwSubAuthority5 As Long, ByVal dwSubAuthority6 As Long, ByVal dwSubAuthority7 As Long, lpPSid As Long) As Long
'@ԭ��
'    BOOL WINAPI AllocateAndInitializeSid(
'      _In_  PSID_IDENTIFIER_AUTHORITY pIdentifierAuthority,
'      _In_  BYTE                      nSubAuthorityCount,
'      _In_  DWORD                     dwSubAuthority0,
'      _In_  DWORD                     dwSubAuthority1,
'      _In_  DWORD                     dwSubAuthority2,
'      _In_  DWORD                     dwSubAuthority3,
'      _In_  DWORD                     dwSubAuthority4,
'      _In_  DWORD                     dwSubAuthority5,
'      _In_  DWORD                     dwSubAuthority6,
'      _In_  DWORD                     dwSubAuthority7,
'      _Out_ PSID                      *pSid
'    );
'@����
'    AllocateAndInitializeSid����ʹ�����˸���Ȩ�޷���ͳ�ʼ����ȫ��ʶ��(SID)��
'@����
'pIdentifierAuthority _In_
'   ָ��SID_IDENTIFIER_AUTHORITY�ṹ��ָ�롣�˽ṹ�ṩҪ��SID�����õĶ�����ʶ��Ȩ��ֵ��
'nSubAuthorityCount _In_
'   ָ��Ҫ������SID�е���Ȩ�޵��������˲�������ʶ�ж�����Ȩ�޲��������������ֵ����������������һ����1��8��ֵ��
'   ���磬ֵ3��ʾ��dwSubAuthority0��dwSubAuthority1��dwSubAuthority2����ָ������Ȩ��ֵ�����������ֵ����������ֵ��
'dwSubAuthority0 _In_
'   Ҫ������SID�е���Ȩ��ֵ
'dwSubAuthority1 _In_
'   Ҫ������SID�е���Ȩ��ֵ
'dwSubAuthority2 _In_
'   Ҫ������SID�е���Ȩ��ֵ
'dwSubAuthority3 _In_
'   Ҫ������SID�е���Ȩ��ֵ
'dwSubAuthority4 _In_
'   Ҫ������SID�е���Ȩ��ֵ
'dwSubAuthority5 _In_
'   Ҫ������SID�е���Ȩ��ֵ
'dwSubAuthority6 _In_
'   Ҫ������SID�е���Ȩ��ֵ
'dwSubAuthority7 _In_
'   Ҫ������SID�е���Ȩ��ֵ
'pSid _Out_
'   һ��ָ�������ָ�룬�ñ�������ָ���ѷ���ͳ�ʼ��SID�ṹ��ָ�롣
'@����ֵ
'   ��������ɹ�������ֵΪ���㡣
'   �������ʧ�ܣ�����ֵΪ�㡣Ҫ��ȡ��չ�Ĵ�����Ϣ�������GetLastError��
'@��ע
'   ����ʹ��FreeSid�����ͷŷ����AllocateAndInitializeSid������SID��
'   �ú�������һ������32λRIDֵ��SID��������Ҫ������RIDֵ��Ӧ�ó��򣬿���ʹ��CreateWellKnownSid��
'@Requirements
'Minimum supported client        Windows XP [desktop apps | UWP apps]
'Minimum supported server        Windows Server 2003 [desktop apps | UWP apps]
'Header                          Winbase.h (include Windows.h)
'Library                         Advapi32.lib
'dll                             Advapi32.dll
Private Declare Function IsValidSid Lib "advapi32.dll" (ByVal pSid As Long) As Long
'@ԭ��
'    BOOL WINAPI IsValidSid(
'      _In_ PSID pSid
'    );
'@����
'    IsValidSid������֤��ȫ��ʶ��(SID)����������֤�޶�������֪��Χ�ڣ�������Ȩ�޵�����С�����ֵ��
'@����
'pSid  _In_
'   ָ��Ҫ��֤��SID�ṹ��ָ�롣�˲�������Ϊ�ա�
'@����ֵ
'   ���SID�ṹ��Ч���򷵻�ֵΪ���㡣
'   ���SID�ṹ��Ч���򷵻�ֵΪ�㡣�ú���û����չ�Ĵ�����Ϣ;��Ҫ����GetLastError��
'@��ע
'   ���pSidΪ�գ���Ӧ�ó���ʧ�ܣ������ַ��ʳ�ͻ��
'@Requirements
'Minimum supported client        Windows XP [desktop apps | UWP apps]
'Minimum supported server        Windows Server 2003 [desktop apps | UWP apps]
'Header                          Winbase.h (include Windows.h)
'Library                         Advapi32.lib
'dll                             Advapi32.dll
Private Declare Function EqualSid Lib "advapi32.dll" (ByVal pSid1 As Long, ByVal pSid2 As Long) As Long
'@ԭ��
'    BOOL WINAPI EqualSid(
'      _In_ PSID pSid1,
'      _In_ PSID pSid2
'    );
'@����
'    EqualSid��������������ȫ��ʶ��(SID)ֵ�Ƿ���ȡ�����SID������ȫƥ����ܱ���Ϊ��ƽ�ȵġ�
'@����
'pSid1 _In_
'   ָ��Ҫ�Ƚϵĵ�һ��SID�ṹ��ָ�롣����ṹ����Ϊ����Ч�ġ�
'pSid2 [��]
'   ָ��Ҫ�Ƚϵĵڶ���SID�ṹ��ָ�롣����ṹ����Ϊ����Ч�ġ�
'@����ֵ
'   ���SID�ṹ��ȣ��򷵻�ֵΪ���㡣
'   ���SID�ṹ����ȣ��򷵻�ֵΪ�㡣Ҫ��ȡ��չ�Ĵ�����Ϣ�������GetLastError��
'   �����һSID�ṹ��Ч���򷵻�ֵδ���塣
'@Requirements
'Minimum supported client        Windows XP [desktop apps | UWP apps]
'Minimum supported server        Windows Server 2003 [desktop apps | UWP apps]
'Header                          Winbase.h (include Windows.h)
'Library                         Advapi32.lib
'dll                             Advapi32.dll
Private Declare Function FreeSid Lib "advapi32.dll" (ByVal pSid As Long) As Long
'@ԭ��
'    PVOID WINAPI FreeSid(
'      _In_ PSID pSid
'    );
'@����
'pSid  _In_
'   ָ��SID�ṹ��ָ���ͷš�
'@����ֵ
'   ��������ɹ����򷵻�NULL��
'   �������ʧ�ܣ���������һ��ָ����pSid������ʾ��SID�ṹ��ָ�롣
'����SIDs
'   ������֪�İ�ȫ��ʶ��(SIDs)��ʶͨ�����ͨ���û������磬һЩ���õ�SID����������Ⱥ����û�:
'   ÿ���˻����磬����һ�����������û����顣
'   CREATOR_OWNER���������ɼ̳�ACE�е�ռλ������ACE���̳�ʱ��ϵͳ�ö��󴴽��ߵ�SID�滻CREATOR_OWNER SID��
'   ���ؼ�������������Administrators�顣
'��һЩ������֪��SID�����Ƕ�����ʹ�����ְ�ȫģʽ�İ�ȫϵͳ������Windows����Ĳ���ϵͳ����������ġ����⣬��һЩ����SIDֻ��Windowsϵͳ�����塣
'Windows APIΪ��֪�ı�ʶ��Ȩ�޺���Ա�ʶ��(RID)ֵ������һ�鳣����������ʹ����Щ����������������С���췢չ�й��ҡ���������ӽ����SECURITY_WORLD_SID_AUTHORITY��SECURITY_WORLD_RID��������ʾ�˴��������û�(�����˻�����)���������ͨ��֪��SID:
'S -1 - 1 - 0
'����ʹ��SID���ַ�����ʾ��������S���ַ�����ʶΪSID����һ��1��SID���޶�������������������SECURITY_WORLD_SID_AUTHORITY��SECURITY_WORLD_RID������
'������ʹ��AllocateAndInitializeSid����������SID�������ǽ���ʶ��Ȩ��ֵ�����˸���Ȩ��ֵ������������磬Ҫȷ����¼���û��Ƿ���ĳ���ض�֪����ĳ�Ա�������AllocateAndInitializeSidΪ��֪���鹹��һ��SID����ʹ��EqualSid��������SID���û����������е���SID���бȽϡ��й�ʾ������μ���c++�еķ�������������SID�����������FreeSid�������ͷ���AllocateAndInitializeSid�����SID��
'���ڵ����ಿ�ְ���������֪SID�ı���Լ�������������������֪��SID�ı�ʶ��Ȩ������Ȩ�������ı��
'������һЩ������֪��SID
'ͨ�õ�����sid�ַ���ֵ��ʶ
'Null SID           S -1 - 0 - 0        û�г�Ա�����塣��ͨ����SIDֵδ֪ʱʹ�á�
'World              S - 1 - 1 - 0       ���������û����顣
'Local              S -1 - 2 - 0        ��¼������(������)���ӵ�ϵͳ���ն˵��û���
'Creator Owner ID   S -1 - 3 - 0        Ҫ�ɴ����¶�����û��İ�ȫ��ʶ���滻�İ�ȫ��ʶ������SID���ڿɼ̳�ace��
'Creator Group ID   S -1 - 3 - 1        Ҫ�ɴ����¶�����û�������SID�滻�İ�ȫ��ʶ�����ڿɼ̳�ace��ʹ�ô�SID
'�±��г���Ԥ����ı�ʶ��Ȩ�޳�����ǰ�ĸ�ֵ����������֪��SID;���һ��ֵ����Windows��������SID��
'��ʶ��Ȩ��ֵsid�ַ���ǰ׺
'SECURITY_NULL_SID_AUTHORITY        0       S -1 - 0
'SECURITY_WORLD_SID_AUTHORITY       1       S -1 - 1
'SECURITY_LOCAL_SID_AUTHORITY       2       S -1 - 2
'SECURITY_CREATOR_SID_AUTHORITY     3       S -1 - 3
'SECURITY_NT_AUTHORITY              5       S -1 - 5
'����RIDֵ����������֪��SID����ʶ��Ȩ������ʾ��ʶ��Ȩ�޵�ǰ׺�������Խ�RID��֮����Դ���һ��ͨ�õ�������֪��SID��
'��Ա�ʶ��Ȩ��ֵ��ʶ��Ȩ��
'SECURITY_NULL_RID                  0       S -1 - 0
'SECURITY_WORLD_RID                 0       S -1 - 1
'SECURITY_LOCAL_RID                 0       S -1 - 2
'SECURITY_LOCAL_LOGON_RID           1       S -1 - 2
'SECURITY_CREATOR_OWNER_RID         0       S -1 - 3
'SECURITY_CREATOR_GROUP_RID         1       S -1 - 3
'SECURITY_NT_AUTHORITY (S-1-5)Ԥ����ı�ʶ��Ȩ�����ɵ�sid����ͨ�õģ�������Windows��װ�������塣������ʹ�����´���SECURITY_NT_AUTHORITY��RIDֵ������֪����sid��
'�����ַ���ֵ��ʶ
Private Const SECURITY_DIALUP_RID               As Long = &H1
'   SECURITY_DIALUP_RID        S -1 - 5 - 1        ʹ�ò��ŵ��ƽ������¼�ն˵��û�������һ�����ʶ����
Private Const SECURITY_NETWORK_RID              As Long = &H2
'   SECURITY_NETWORK_RID       S -1 - 5 - 2        ͨ�������¼���û�������һ�����ʶ������ӵ��������¼���̵������С���Ӧ�ĵ�¼������LOGON32_LOGON_NETWORK��
Private Const SECURITY_BATCH_RID                As Long = &H3
'   SECURITY_BATCH_RID         S -1 - 5 - 3        ʹ����������й��ܵ�¼���û�������һ�����ʶ������ӵ���Ϊ��������ҵ��¼�Ľ��̵������С���Ӧ�ĵ�¼������LOGON32_LOGON_BATCH��
Private Const SECURITY_INTERACTIVE_RID          As Long = &H4
'   SECURITY_INTERACTIVE_RID   S -1 - 5 - 4        �û���¼���н�������������һ�����ʶ�����������Խ�����ʽ��¼ʱ��ӵ����̵������С���Ӧ�ĵ�¼������LOGON32_LOGON_INTERACTIVE��
Private Const SECURITY_LOGON_IDS_RID            As Long = &H5
'   SECURITY_LOGON_IDS_RID     S -1 - 5 - 5 - X - y    һ����¼�Ự��������ȷ��ֻ�и�����¼�Ự�еĽ��̲��ܷ��ʸûỰ��window-station���󡣶���ÿ����¼�Ự����ЩSID��X��Yֵ�ǲ�ͬ�ġ�SECURITY_LOGON_IDS_RID_COUNTֵ�������ʶ��(5-X-Y)�е�rid������
Private Const SECURITY_SERVICE_RID              As Long = &H6
'   SECURITY_SERVICE_RID       S -1 - 5 - 6        ��Ȩ��Ϊ�����¼���ʻ�������һ�����ʶ������ӵ���Ϊ�����¼�Ľ��̵������С���Ӧ�ĵ�¼������LOGON32_LOGON_SERVICE��
Private Const SECURITY_ANONYMOUS_LOGON_RID      As Long = &H7
'   SECURITY_ANONYMOUS_LOGON_RID   S -1 - 5 - 7    ������¼����ջỰ��¼��
Private Const SECURITY_PROXY_RID                As Long = &H8
'   SECURITY_PROXY_RID         S -1 - 5 - 8        ����
Private Const SECURITY_ENTERPRISE_CONTROLLERS_RID   As Long = &H9
'   SECURITY_ENTERPRISE_CONTROLLERS_RID    S -1 - 5 - 9    ��ҵ��������
Private Const SECURITY_PRINCIPAL_SELF_RID       As Long = &HA
'   SECURITY_PRINCIPAL_SELF_RID    S -1 - 5 - 10       PRINCIPAL_SELF��ȫ��ʶ���������û���������ACL��ʹ�á��ڷ��ʼ���ڼ䣬ϵͳ�ö����SID�滻SID��PRINCIPAL_SELF SID����ָ���ɼ̳е�ACE����ACEӦ���ڼ̳�ACE���û��������������ģʽ��Ĭ�ϰ�ȫ�������б�ʾ�Ѵ��������SID��Ψһ������
Private Const SECURITY_AUTHENTICATED_USER_RID   As Long = &HB
'   SECURITY_AUTHENTICATED_USER_RID    S -1 - 5 - 11   ͨ�������֤���û���
Private Const SECURITY_RESTRICTED_CODE_RID      As Long = &HC
'   SECURITY_RESTRICTED_CODE_RID   S -1 - 5 - 12       �����ƵĴ��롣
Private Const SECURITY_TERMINAL_SERVER_RID      As Long = &HD
'   SECURITY_TERMINAL_SERVER_RID   S -1 - 5 - 13       �ն˷����Զ���ӵ���¼���ն˷��������û��İ�ȫ�����С�
Private Const SECURITY_LOCAL_SYSTEM_RID         As Long = &H12
'   SECURITY_LOCAL_SYSTEM_RID      S -1 - 5 - 18       ����ϵͳʹ�õ������ʻ���
Private Const SECURITY_NT_NON_UNIQUE            As Long = &H15
'   SECURITY_NT_NON_UNIQUE         S -1 - 5 - 21       SID���Ƕ�һ�޶���
Private Const SECURITY_BUILTIN_DOMAIN_RID       As Long = &H20
'   SECURITY_BUILTIN_DOMAIN_RID    S -1 - 5 - 32       ���õ�ϵͳ��
'����rid��ÿ������ء�
'�����ʶ
'DOMAIN_ALIAS_RID_CERTSVC_DCOM_ACCESS_GROUP     ����ʹ�÷ֲ�ʽ�������ģ��(DCOM)���ӵ���֤�������û��顣
'DOMAIN_USER_RID_ADMIN                          ���еĹ����û��ʻ���
'DOMAIN_USER_RID_GUEST                          ���е������û��ʻ���û���ʻ����û������Զ���¼���ʻ���
'DOMAIN_GROUP_RID_ADMINS                        �����Ա�顣���ʻ������������з���������ϵͳ��ϵͳ�ϡ�
'DOMAIN_GROUP_RID_USERS                         һ�����а��������û��ʻ����顣�����û������Զ���ӵ�������С�
'DOMAIN_GROUP_RID_GUESTS                        ���е��������ʻ���
'DOMAIN_GROUP_RID_COMPUTERS                     �������顣���е����м�������������ĳ�Ա��
'DOMAIN_GROUP_RID_CONTROLLERS                   ����������顣���е�����DCs���������ĳ�Ա��
'DOMAIN_GROUP_RID_CERT_ADMINS                   ֤�鷢�����顣����֤�����ļ�����������ĳ�Ա��
'DOMAIN_GROUP_RID_ENTERPRISE_READONLY_DOMAIN_CONTROLLERS    һ����ҵֻ�����������
'DOMAIN_GROUP_RID_SCHEMA_ADMINS                 ģʽ����Ա�顣�����ĳ�Ա�����޸�Active Directoryģʽ��
'DOMAIN_GROUP_RID_ENTERPRISE_ADMINS             ��ҵ����Ա�顣�����ĳ�Ա������ȫ����Active Directory���е���������ҵ����Ա����Ⱥ��Ĳ�����������ӻ�ɾ������
'DOMAIN_GROUP_RID_POLICY_ADMINS                 ���Թ���Ա�顣
'DOMAIN_GROUP_RID_READONLY_CONTROLLERS          ֻ����������顣
'����rid����ָ��ǿ�������Լ���
Private Const SECURITY_MANDATORY_UNTRUSTED_RID  As Long = &H0
'   �����ŵ�
Private Const SECURITY_MANDATORY_LOW_RID        As Long = &H1000
'   �͵�������
Private Const SECURITY_MANDATORY_MEDIUM_RID     As Long = &H2000
'   ý���������
Private Const SECURITY_MANDATORY_MEDIUM_PLUS_RID As Long = SECURITY_MANDATORY_MEDIUM_RID + &H100
'   �еȸ߶ȵ�������
Private Const SECURITY_MANDATORY_HIGH_RID       As Long = &H3000
'   ��������
Private Const SECURITY_MANDATORY_SYSTEM_RID     As Long = &H4000
'   ϵͳ��������
Private Const SECURITY_MANDATORY_PROTECTED_PROCESS_RID As Long = &H5000
'   �ܱ����Ĺ���
'�±�����һЩ�����rid��ʾ����������ʹ������Ϊ������(����)�γ�������֪��SID���йر��غ�ȫ����ĸ�����Ϣ����μ������麯�����麯����
Private Const DOMAIN_ALIAS_RID_ADMINS           As Long = &H220
'   ���������ı����顣
Private Const DOMAIN_ALIAS_RID_USERS            As Long = &H221
'   ��ʾ���������û��ı����顣
Private Const DOMAIN_ALIAS_RID_GUESTS           As Long = &H222
'   ��ʾ��������ı����顣
Private Const DOMAIN_ALIAS_RID_POWER_USERS      As Long = &H223
'   һ�������飬���ڱ�ʾһ����һ���û�����Щ�û�ϣ����һ��ϵͳ��Ϊ���ǵĸ��˼�����������Ƕ���û��Ĺ���վ��
Private Const DOMAIN_ALIAS_RID_ACCOUNT_OPS      As Long = &H224
'   �����������з���������ϵͳ��ϵͳ�ϵı����顣���������������Ʒǹ���Ա�ʻ���
Private Const DOMAIN_ALIAS_RID_SYSTEM_OPS       As Long = &H225
'   �����������з���������ϵͳ��ϵͳ�ϵı����顣���������ִ��ϵͳ�����ܣ���������ȫ���ܡ����������繲�����ƴ�ӡ������������վ��ִ������������
Private Const DOMAIN_ALIAS_RID_PRINT_OPS        As Long = &H226
'   �����������з���������ϵͳ��ϵͳ�ϵı����顣�����������ƴ�ӡ���ʹ�ӡ���С�
Private Const DOMAIN_ALIAS_RID_BACKUP_OPS       As Long = &H227
'   ���ڿ����ļ����ݺͻָ���Ȩ����ı����顣
Private Const DOMAIN_ALIAS_RID_REPLICATOR       As Long = &H228
'   ���𽫰�ȫ���ݿ��������������Ƶ�������������ı����顣��Щ�ʻ�����ϵͳʹ�á�
Private Const DOMAIN_ALIAS_RID_RAS_SERVERS      As Long = &H229
'   ��ʾRAS��IAS�������ı����顣�������������û�����ĸ������ԡ�
Private Const DOMAIN_ALIAS_RID_PREW2KCOMPACCESS As Long = &H22A
'   ������������Windows 2000��������ϵͳ�ϵı����顣�йظ�����Ϣ����μ������������ʡ�
Private Const DOMAIN_ALIAS_RID_REMOTE_DESKTOP_USERS As Long = &H22B
'   ��ʾ����Զ�������û��ı����顣
Private Const DOMAIN_ALIAS_RID_NETWORK_CONFIGURATION_OPS    As Long = &H22C
'   ��ʾ�������õı����顣
Private Const DOMAIN_ALIAS_RID_INCOMING_FOREST_TRUST_BUILDERS    As Long = &H22D
'   ��ʾ�κ�forest trust�û��ı����顣
Private Const DOMAIN_ALIAS_RID_MONITORING_USERS As Long = &H22E
'   ��ʾ�����ӵ������û��ı����顣
Private Const DOMAIN_ALIAS_RID_LOGGING_USERS    As Long = &H22F
'   �����¼�û���־�ı����顣
Private Const DOMAIN_ALIAS_RID_AUTHORIZATIONACCESS  As Long = &H230
'   ��ʾ������Ȩ���ʵı����顣
Private Const DOMAIN_ALIAS_RID_TS_LICENSE_SERVERS    As Long = &H231
'   �����������з���������ϵͳ(�����ն˷����Զ�̷���)��ϵͳ�ϵı����顣
Private Const DOMAIN_ALIAS_RID_DCOM_USERS       As Long = &H232
'   ��ʾ����ʹ�÷ֲ�ʽ�������ģ��(DCOM)���û��ı����顣
Private Const DOMAIN_ALIAS_RID_IUSERS           As Long = &H238
'   ����Internet�û��ı����顣
Private Const DOMAIN_ALIAS_RID_CRYPTO_OPERATORS     As Long = &H239
'   ��ʾ������������ķ��ʵı����顣
Private Const DOMAIN_ALIAS_RID_CACHEABLE_PRINCIPALS_GROUP   As Long = &H23B
'   ��ʾ���Ի��������ı����顣
Private Const DOMAIN_ALIAS_RID_NON_CACHEABLE_PRINCIPALS_GROUP   As Long = &H23C
'   ��ʾ���ܻ��������ı����顣
Private Const DOMAIN_ALIAS_RID_EVENT_LOG_READERS_GROUP    As Long = &H23D
'   ��ʾ�¼���־��ȡ���ı����顣
Private Const DOMAIN_ALIAS_RID_CERTSVC_DCOM_ACCESS_GROUP    As Long = &H23E
'   ����ʹ�÷ֲ�ʽ�������ģ��(DCOM)���ӵ���֤�����ı����û��顣
Private Declare Function GetCurrentThread Lib "kernel32.dll" () As Long
'@ԭ��
'    HANDLE WINAPI GetCurrentThread(void);
'@����
'    ���������̵߳�α�����
'@����
'   �������û�в�����
'@����ֵ
'   ����ֵ�ǵ�ǰ�̵߳�α�����
'@��ע
'    α�����һ������ĳ�������������Ϊ��ǰ�߳̾�������ۺ�ʱ��Ҫ�߳̾���������̶߳�����ʹ����������ָ���Լ����ӽ��̲���̳�α�����
'    ���������ж�thread�����THREAD_ALL_ACCESS����Ȩ���йظ�����Ϣ����μ��̰߳�ȫ�ͷ���Ȩ�ޡ�
'    Windows Server 2003��Windows XP:�����������̵߳İ�ȫ������������ĶԽ��̵������Ƶ�������Ȩ��
'    һ���̲߳���ʹ�øú�������һ������������߳̿���ʹ�øþ�����õ�һ���̡߳�������Ǳ�����Ϊ����ʹ�������̡߳�ͨ���ڵ���DuplicateHandle����ʱ��α���ָ��ΪԴ������߳̿���Ϊ�Լ�����һ������ʵ������������߳̿���ʹ�øþ���������������̼̳С�
'    ��������Ҫα���ʱ������Ҫ�ر�����ʹ�ô˾������close�������û��Ч���������DuplicateHandle����α����������ر��ظ������
'    ģ�ⰲȫ������ʱ��Ҫ�����̡߳����ý��ɹ��������´������߳��ڵ���GetCurrentThreadʱ�����ٶ�����ķ���Ȩ�ޡ�������̵߳ķ���Ȩ�޽���ģ���û��Խ��̵ķ���Ȩ��������һЩ����Ȩ��(����THREAD_SET_THREAD_TOKEN��THREAD_GET_CONTEXT)���ܲ����ڣ��Ӷ����������ʧ�ܡ�
'@Requirements
'Minimum supported client        Windows XP [desktop apps | UWP apps]
'Minimum supported server        Windows Server 2003 [desktop apps | UWP apps]
'Minimum supported phone         Windows Phone 8
'Header                          WinBase.h on Windows XP, Windows Server 2003, Windows Vista, Windows 7, Windows Server 2008 and Windows Server 2008 R2 (include Windows.h);Processthreadsapi.h on Windows 8 and Windows Server 2012
'Library                         kernel32.lib
'dll                             kernel32.dll

'��׼Ȩ��
Private Const Delete                            As Long = &H10000
Private Const READ_CONTROL                      As Long = &H20000
Private Const WRITE_DAC                         As Long = &H40000
Private Const WRITE_OWNER                       As Long = &H80000
Private Const SYNCHRONIZE                       As Long = &H100000
Private Const STANDARD_RIGHTS_REQUIRED          As Long = &HF0000
Private Const STANDARD_RIGHTS_READ              As Long = READ_CONTROL
Private Const STANDARD_RIGHTS_WRITE             As Long = READ_CONTROL
Private Const STANDARD_RIGHTS_EXECUTE           As Long = READ_CONTROL
Private Const STANDARD_RIGHTS_ALL               As Long = &H1F0000
Private Const SPECIFIC_RIGHTS_ALL               As Long = &HFFFF
'TokenȨ��
Private Const TOKEN_ASSIGN_PRIMARY              As Long = &H1
Private Const TOKEN_DUPLICATE                   As Long = &H2
Private Const TOKEN_IMPERSONATE                 As Long = &H4
Private Const TOKEN_QUERY                       As Long = &H8
Private Const TOKEN_QUERY_SOURCE                As Long = &H10
Private Const TOKEN_ADJUST_PRIVILEGES           As Long = &H20
Private Const TOKEN_ADJUST_GROUPS               As Long = &H40
Private Const TOKEN_ADJUST_DEFAULT              As Long = &H80
Private Const TOKEN_ADJUST_SESSIONID            As Long = &H100
Private Const TOKEN_ALL_ACCESS_P                As Long = STANDARD_RIGHTS_REQUIRED Or TOKEN_ASSIGN_PRIMARY Or TOKEN_DUPLICATE Or _
                                                    TOKEN_IMPERSONATE Or TOKEN_QUERY Or TOKEN_QUERY_SOURCE Or TOKEN_ADJUST_PRIVILEGES Or _
                                                    TOKEN_ADJUST_GROUPS Or TOKEN_ADJUST_DEFAULT
Private Const TOKEN_ALL_ACCESS                  As Long = TOKEN_ALL_ACCESS_P Or TOKEN_ADJUST_SESSIONID
Private Const TOKEN_READ                        As Long = STANDARD_RIGHTS_READ Or TOKEN_QUERY
Private Const TOKEN_WRITE                       As Long = STANDARD_RIGHTS_WRITE Or TOKEN_ADJUST_PRIVILEGES Or TOKEN_ADJUST_GROUPS Or TOKEN_ADJUST_DEFAULT
Private Const TOKEN_EXECUTE                     As Long = STANDARD_RIGHTS_EXECUTE
Private Declare Function GetCurrentProcess Lib "kernel32.dll" () As Long
'@ԭ��
'    HANDLE WINAPI GetCurrentProcess(void);
'@����
'    ������ǰ���̵�α�����
'@����
'   �������û�в�����
'@����ֵ
'   ����ֵ�ǵ�ǰ�̵߳�α�����
'@����ֵ
'   ����ֵ�ǵ�ǰ���̵�α�����
'@��ע
'    α�����һ������ĳ�������ǰ(���)-1����������Ϊ��ǰ���̾����Ϊ����δ���Ĳ���ϵͳ���ݣ���õ���GetCurrentProcess��������Ӳ����������������ۺ�ʱ��Ҫ���̾�����������̶�����ʹ��α�����ָ���Լ������̡��ӽ��̲���̳�α�����
'    �˾������PROCESS_ALL_ACCESS�������̶����Ȩ�ޡ��йظ�����Ϣ����μ����̰�ȫ�ͷ���Ȩ�ޡ�
'    Windows Server 2003��Windows XP:�˾�����н��̰�ȫ����������ĶԽ��������Ƶ�������Ȩ��
'    ͨ���ڵ���DuplicateHandle����ʱ��α���ָ��ΪԴ��������̿���Ϊ�Լ�����һ������ʵ��������þ�����������̵�������������Ч�ģ����߿��Ա��������̼̳С����̻�����ʹ��OpenProcess����Ϊ�Լ���һ��ʵ�����
'    ��������Ҫα���ʱ������Ҫ�ر�����ʹ��α�������close�������û��Ч���������DuplicateHandle����α����������ر��ظ������
'@Requirements
'Minimum supported client        Windows XP [desktop apps | UWP apps]
'Minimum supported server        Windows Server 2003 [desktop apps | UWP apps]
'Minimum supported phone         Windows Phone 8
'Header                          WinBase.h on Windows XP, Windows Server 2003, Windows Vista, Windows 7, Windows Server 2008 and Windows Server 2008 R2 (include Windows.h);Processthreadsapi.h on Windows 8 and Windows Server 2012
'Library                         kernel32.lib
'dll                             kernel32.dll
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
'@ԭ��
'    HINSTANCE ShellExecute(
'      _In_opt_ HWND    hwnd,
'      _In_opt_ LPCTSTR lpOperation,
'      _In_     LPCTSTR lpFile,
'      _In_opt_ LPCTSTR lpParameters,
'      _In_opt_ LPCTSTR lpDirectory,
'      _In_     INT     nShowCmd
'    );
'@����
'    ��ָ�����ļ�ִ�в�����
'@����
'hwnd _In_opt_
'   ����: hwnd
'   �����ڵľ����������ʾUI�������Ϣ������������봰�ڹ��������ֵ����Ϊ�ա�
'lpOperation _In_opt_
'   ����: LPCTSTR
'   ָ����null��β���ַ�����ָ�룬�ڱ����г�Ϊν�ʣ���ָ��Ҫִ�еĲ���������ν�ʼ�ȡ�����ض����ļ����ļ��С�ͨ��������Ŀ�ݲ˵��п��õĲ������ǿ��õ�ν�ʡ����õĶ�����:
'   edit    �����༭�������ĵ����б༭�����lpFile�����ĵ��ļ���������ʧ�ܡ�
'   explore ���lpFileָ�����ļ��С�
'   find    ��lpDirectoryָ����Ŀ¼������������
'   open    ����lpFile����ָ�������Ŀ�������ļ����ļ��С�
'   print   ��ӡlpFileָ�����ļ������lpFile�����ĵ��ļ�����ú�����ʧ�ܡ�
'   NULL    ������ã���ʹ��Ĭ��ν�ʡ����û�У���ʹ�á�open�����ʡ�����������ʶ������ã�ϵͳ��ʹ��ע������г��ĵ�һ��ν�ʡ�
'lpFile _In_
'   ����: LPCTSTR
'   ָ����null��β���ַ�����ָ�룬���ַ���ָ��Ҫ������ִ��ָ��ν�ʵ��ļ������Ҫָ��Shell���ƿռ�����봫����ȫ�޶��Ľ�������ע�⣬�������ж���֧������ν�ʡ����磬���������ĵ����Ͷ�֧�֡�print��ν�ʡ������lpDirectory����ʹ�����·������Ҫ��lpFileʹ�����·����
'lpParameters(,��ѡ)
'   ����: LPCTSTR
'   ���lpFileָ����һ����ִ���ļ��������������ָ��һ����null��β���ַ�����ָ�룬���ַ���ָ��Ҫ���ݸ�Ӧ�ó���Ĳ��������ַ����ĸ�ʽ��Ҫ���õ�ν�ʾ��������lpFileָ����һ���ĵ��ļ�����ôlpParametersӦ��ΪNULL��
'lpDirectory _In_opt_
'   ����: LPCTSTR
'   ָ����null��β���ַ�����ָ�룬���ַ���ָ��������Ĭ��(����)Ŀ¼�������ֵΪ�գ���ʹ�õ�ǰ����Ŀ¼�������lpFile���ṩ�����·������ҪΪlpDirectoryʹ�����·����
'nShowCmd    _In_
'   ����:INT
'   ָ��Ӧ�ó����ʱ�����ʾ�ı�־�����lpFileָ����һ���ĵ��ļ�����ñ�־�����򵥵ش��ݸ���ص�Ӧ�ó�����δ�����ȡ����Ӧ�ó�����Щֵ��Winuser.h�ж��塣
Private Const SW_HIDE                           As Long = &H0
'   ���ش��ڲ�������һ�����ڡ�
Private Const SW_MAXIMIZE                       As Long = &H3
'   ���ָ���Ĵ��ڡ�
Private Const SW_MINIMIZE                       As Long = &H6
'   ��С��ָ���Ĵ��ڲ���z˳�򼤻���һ���������ڡ�
Private Const SW_RESTORE                        As Long = &H9
'   �����ʾ���ڡ�������ڱ���С������󻯣����ڻὫ��ָ���ԭ���Ĵ�С��λ�á�Ӧ�ó���Ӧ���ڻָ���С������ʱָ���˱�־��
Private Const SW_SHOW                           As Long = &H5
'   ����ڲ���ʾ�䵱ǰ��С��λ�á�
Private Const SW_SHOWDEFAULT                    As Long = &HA
'   ��������Ӧ�ó���ĳ��򴫵ݸ�CreateProcess������STARTUPINFO�ṹ��ָ����SW_��־������ʾ״̬��Ӧ�ó���Ӧ��ʹ�ô˱�־����ShowWindow�������������ڵĳ�ʼ��ʾ״̬��
Private Const SW_SHOWMAXIMIZED                  As Long = &H3
'   ����ڲ�������ʾΪ��󻯴��ڡ�
Private Const SW_SHOWMINIMIZED                  As Long = &H2
'   ����ڲ�������ʾΪ��С�����ڡ�
Private Const SW_SHOWMINNOACTIVE                As Long = &H7
'   ��������ʾΪ��С�����ڡ�����ڱ��ֻ״̬��
Private Const SW_SHOWNA                         As Long = &H8
'   ��ʾ���ڵĵ�ǰ״̬������ڱ��ֻ״̬��
Private Const SW_SHOWNOACTIVATE                 As Long = &H4
'   ��ʾ���ڵ����´�С��λ�á�����ڱ��ֻ״̬��
Private Const SW_SHOWNORMAL                     As Long = &H1
'   �����ʾ���ڡ�������ڱ���С������󻯣����ڻὫ��ָ���ԭ���Ĵ�С��λ�á�Ӧ�ó���Ӧ���ڵ�һ����ʾ����ʱָ���˱�־��
'@����ֵ
'   ����: HINSTANCE
'   ��������ɹ����򷵻�һ������32��ֵ���������ʧ�ܣ���������һ������ֵ����ֵָʾʧ�ܵ�ԭ�򡣷���ֵ��ת��Ϊһ��HINSTANCE���Ա���16λWindowsӦ�ó��������ݡ�Ȼ�����ⲻ��һ��������HINSTANCE����ֻ�ܱ�ǿ��ת��Ϊ����������32������Ĵ��������бȽϡ�
'   ���ش�������
'   0           ����ϵͳ�ڴ����Դ���㡣
Private Const ERROR_FILE_NOT_FOUND              As Long = &H2
'   û���ҵ�ָ�����ļ���
Private Const ERROR_PATH_NOT_FOUND              As Long = &H3
'   û���ҵ�ָ����·����
Private Const ERROR_BAD_FORMAT                  As Long = &HB
'   .exe�ļ���Ч(��win32 .exe��.exeӳ���еĴ���)��
Private Const SE_ERR_ACCESSDENIED               As Long = &H5
'   ����ϵͳ�ܾ�����ָ�����ļ���
Private Const SE_ERR_ASSOCINCOMPLETE            As Long = &H1B
'   �ļ�����������������Ч��
Private Const SE_ERR_DDEBUSY                    As Long = &H1E
'   �������ڴ�������DDE��������޷����DDE����
Private Const SE_ERR_DDEFAIL                    As Long = &H1D
'   DDE����ʧ�ܡ�
Private Const SE_ERR_DDETIMEOUT                 As Long = &H1C
'   ��������ʱ���޷����DDE����
Private Const SE_ERR_DLLNOTFOUND                As Long = &H20
'   û���ҵ�ָ����DLL��
Private Const SE_ERR_FNF                        As Long = &H2
'   û���ҵ�ָ�����ļ���
Private Const SE_ERR_NOASSOC                    As Long = &H1F
'   û��������ļ���չ��������Ӧ�ó���������Դ�ӡ���ɴ�ӡ���ļ���Ҳ�����ش˴���
Private Const SE_ERR_OOM                        As Long = &H8
'   û���㹻���ڴ���������������
Private Const SE_ERR_PNF                        As Long = &H3
'   û���ҵ�ָ����·����
Private Const SE_ERR_SHARE                      As Long = &H1A
'   �����˹����ͻ��
'@��ע
'   ��ΪShellExecute���Խ�ִ��ί�и�ʹ���������ģ��(COM)�����Shell��չ(����Դ�������Ĳ˵��������ν��ʵ��)������Ӧ���ڵ���ShellExecute֮ǰ��ʼ��COM��һЩShell��չ��ҪCOM���̹߳�Ԣ(STA)���͡�����������£�COMӦ��������ʾ��ʼ��:
'   CoInitializeEx(NULL, coinit_apartmentthreads | COINIT_DISABLE_OLE1DDE)
'   ��Ȼ����ĳЩʵ���У�ShellExecute��ʹ����Щ���͵�Shell��չ��������Щʵ����������Ҫ��ʼ��COM��������ˣ���ʹ�ô˺���֮ǰ��ʼ�ն�COM���г�ʼ����һ���ܺõ�ʵ����
'   �˷���������ִ���ļ��еĿ�ݲ˵���洢��ע����е��κ����
'   Ҫ���ļ��У���ʹ�����µ���֮һ:
'   ShellExecute(�����NULL�� <fully_qualified_path_to_folder>�� NULL, NULL, SW_SHOWNORMAL);
'   ��
'   ShellExecute(�������open����<fully_qualified_path_to_folder>�� NULL, NULL, SW_SHOWNORMAL);
'   Ҫ�鿴�ļ��У���ʹ�����µ���:
'   ShellExecute(�������explore����<fully_qualified_path_to_folder>�� NULL, NULL, SW_SHOWNORMAL);
'   Ҫ����Shell��Ŀ¼����ʵ�ó�����ʹ�����µ��á�
'   ShellExecute(�������find����<fully_qualified_path_to_folder>�� NULL, NULL, 0);
'   ���lpOperationΪ�գ���������lpFileָ�����ļ������lpOperation�ǡ�open����explore�����ú��������Դ򿪻�����ļ��С�
'   Ҫ��ȡ�������ڵ���ShellExecute��������Ӧ�ó������Ϣ����ʹ��ShellExecuteEx��
'   ע�⣬�����ļ��д������ļ���ѡ���еĵ�������������Ӱ��ShellExecute��������ø�ѡ��(Ĭ������)��ShellExecute��ʹ��һ���򿪵���Դ���������ڣ�����������һ���µ���Դ���������ڡ����û�д���Դ���������ڣ�ShellExecute������һ���´��ڡ�
'@Requirements
'Minimum supported client   Windows XP [desktop apps only]
'Minimum supported server   Windows 2000 Server [desktop apps only]
'Header                     Shellapi.h
'Library                    shell32.lib
'dll                        Shell32.dll (version 3.51 or later)
'Unicode and ANSI names     ShellExecuteW (Unicode) And ShellExecuteA(ANSI)
Private Type SHELLEXECUTEINFO
    cbSize          As Long
    fMask           As Long
    hwnd            As Long
    lpVerb          As String
    lpFile          As String
    lpParameters    As String
    nShow           As Long
    hInstApp        As Long
    lpIDList        As Long
    lpClass         As String
    hkeyClass       As Long
    dwHotKey        As Long
    hMonitor        As Long
    'hIcon          as long
    hProcess        As Long
End Type
'@����
'    ����ShellExecuteExʹ�õ���Ϣ��
'@ԭ��
'    typedef struct _SHELLEXECUTEINFO {
'      DWORD     cbSize;
'      ULONG     fMask;
'      HWND      hwnd;
'      LPCTSTR   lpVerb;
'      LPCTSTR   lpFile;
'      LPCTSTR   lpParameters;
'      LPCTSTR   lpDirectory;
'      int       nShow;
'      HINSTANCE hInstApp;
'      LPVOID    lpIDList;
'      LPCTSTR   lpClass;
'      HKEY      hkeyClass;
'      DWORD     dwHotKey;
'      union {
'        HANDLE hIcon;
'        HANDLE hMonitor;
'      } DUMMYUNIONNAME;
'      HANDLE    hProcess;
'    } SHELLEXECUTEINFO, *LPSHELLEXECUTEINFO;
'@��Ա
'cbSize
'   ����: ˫��
'   ����ġ��˽ṹ�Ĵ�С�����ֽ�Ϊ��λ��
'fMask
'   ����: ULONG
'   ��ʾ�����ṹ��Ա�����ݺ���Ч�Եı�־;����ֵ�����:
Private Const SEE_MASK_DEFAULT                  As Long = &H0
'   ʹ��Ĭ��ֵ��
Private Const SEE_MASK_CLASSNAME                As Long = &H1
'   ʹ��lpClass��Ա���������������������SEE_MASK_CLASSKEY��SEE_MASK_CLASSNAME����ʹ�������
Private Const SEE_MASK_CLASSKEY                 As Long = &H3
'   ʹ����hkeyClass��Ա�ṩ������Կ�����������SEE_MASK_CLASSKEY��SEE_MASK_CLASSNAME����ʹ�������
Private Const SEE_MASK_IDLIST                   As Long = &H4
'   ʹ��lpIDList��Ա�ṩ�����ʶ���б�lpIDList��Ա����ָ��ITEMIDLIST�ṹ��
Private Const SEE_MASK_INVOKEIDLIST             As Long = &HC
'   ʹ����ѡ��Ŀ�Ŀ�ݲ˵���������IContextMenu�ӿڡ�ʹ��lpFile�������ļ�ϵͳ·����ʶ�����ʹ��lpIDList������PIDL��ʶ��˱�־����Ӧ�ó���ʹ��ShellExecuteEx�ӿ�ݲ˵���չ�е���ν�ʣ�������ע������г��ľ�̬ν�ʡ�
'   ע�⣬SEE_MASK_INVOKEIDLIST���ǲ���ʾ��SEE_MASK_IDLIST��
Private Const SEE_MASK_ICON                     As Long = &H10
'   ʹ��hIcon��Ա������ͼ�ꡣ�˱�־������SEE_MASK_HMONITOR��ϡ�
'   ע�⣬�˱�־����Windows XP������汾��ʹ�á�����Windows Vista�б������ˡ�
Private Const SEE_MASK_HOTKEY                   As Long = &H20
'   ʹ��dwHotKey��Ա�ṩ�ļ��̿�ݷ�ʽ��
Private Const SEE_MASK_NOCLOSEPROCESS           As Long = &H40
'   ����ָʾhProcess��Ա���ս��̾�����˾��ͨ����������Ӧ�ó������ʹ��ShellExecuteEx�����Ľ��̺�ʱ��ֹ����ĳЩ����£�����ͨ��DDE�Ի�����ִ��ʱ�����᷵�ؾ��������Ӧ�ó������ڲ�����Ҫ���ʱ�رվ����
Private Const SEE_MASK_CONNECTNETDRV            As Long = &H80
'   ��֤�������ӵ��������š��������������ӶϿ����ӵ�������������lpFile��Ա���������ļ���UNC·����
Private Const SEE_MASK_NOASYNC                  As Long = &H100
'   �ȴ�ִ�в�����ɺ󷵻ء������־Ӧ����ʹ��ShellExecute���ĵ�����ʹ�ã���Щ�����߿��ܻᵼ���첽�������DDE��������һ�������ں�̨�߳������еĽ��̡�(ע��:��������ߵ��߳�ģ�Ͳ��ǵ�Ԫ����Ĭ�������ShellExecuteEx�ں�̨�߳������С�)���Ѿ��ں�̨�߳������еĽ��̵���ShellExecuteExӦ�����Ǵ��������־�����⣬����ShellExecuteEx�������˳���Ӧ�ó���Ӧ��ָ���˱�־��
'   ���ִ�в������ں�̨�߳���ִ�еģ����ҵ�����û��ָ��SEE_MASK_ASYNCOK��־����ô�����߳̽��ȵ��½���������ŷ��ء���ͨ����ζ�ŵ�����CreateProcess, DDEͨ���Ѿ���ɣ������Զ���ִ��ί���Ѿ�֪ͨShellExecuteEx���Ѿ���ɡ����ָ����SEE_MASK_WAITFORINPUTIDLE��־����ôShellExecuteEx������WaitForInputIdle�����ȴ��½��̿��У�Ȼ�󷵻أ����ʱΪ1���ӡ�
'   �йغ�ʱ��Ҫ�˱�־�Ľ�һ�����ۣ���μ���ע���֡�
Private Const SEE_MASK_FLAG_DDEWAIT             As Long = &H100
'   ��Ҫʹ��;ʹ��SEE_MASK_NOASYNC���档
Private Const SEE_MASK_DOENVSUBST               As Long = &H200
'   չ��lpDirectory��lpFile��Ա�������ַ�����ָ�����κλ���������
Private Const SEE_MASK_FLAG_NO_UI               As Long = &H400
'   ����������󣬲�Ҫ��ʾ������Ϣ��
Private Const SEE_MASK_UNICODE                  As Long = &H4000
'   ʹ�ô˱�־ָʾUnicodeӦ�ó���
Private Const SEE_MASK_NO_CONSOLE               As Long = &H8000
'   ���ڼ̳и����̵Ŀ���̨�����������������¿���̨��������CreateProcess��ʹ��CREATE_NEW_CONSOLE��־�෴��
Private Const SEE_MASK_ASYNCOK                  As Long = &H100000
'   ִ�п����ں�̨�߳���ִ�У�����Ӧ���������أ�������Ҫ�ȴ���̨�߳���ɡ�ע�⣬��ĳЩ����£�ShellExecuteEx����Դ˱�־�����ȴ�������ɺ󷵻ء�
Private Const SEE_MASK_NOQUERYCLASSSTORE        As Long = &H1000000
'   ��ʹ��
Private Const SEE_MASK_HMONITOR                 As Long = &H200000
'   �ڶ������ϵͳ��ָ��������ʱʹ�ô˱�־����������hMonitor��Ա��ָ�����˱�־������SEE_MASK_ICON��ϡ�
Private Const SEE_MASK_NOZONECHECKS             As Long = &H800000
'   ��Windows XP�����롣��Ҫִ�������顣�����־����ShellExecuteEx�ƹ�IAttachmentExecute���õ������顣
Private Const SEE_MASK_WAITFORINPUTIDLE         As Long = &H2000000
'   �����½���֮�󣬵ȴ����̿��У�Ȼ�󷵻أ���ʱһ���ӡ�������μ�WaitForInputIdle��
Private Const SEE_MASK_FLAG_LOG_USAGE           As Long = &H4000000
'   ��Windows XP�����롣���ٴ�Ӧ�ó����������Ĵ����������㹻�ߵ�Ӧ�ó�������ڿ�ʼ�˵�����ó����б��С�
Private Const SEE_MASK_FLAG_HINST_IS_SITE       As Long = &H8000000
'   hInstApp��Ա����ָ��ʵ��IServiceProvider�Ķ����IUnknown���˶�������վ��ָ�롣վ��ָ��������ShellExecute�������������󶨹��̺͵��õ�ν�ʴ�������ṩ����
'   Ҫ��Windows 8֮ǰ�Ĳ���ϵͳ��ʹ��SEE_MASK_FLAG_HINST_IS_SITE�����ڳ������ֶ�������:#define SEE_MASK_FLAG_HINST_IS_SITE 0x08000000��
'hwnd
'   ����: hwnd
'   ��ѡ�ġ������ڵľ����������ʾϵͳ��ִ�д˺���ʱ���ܲ������κ���Ϣ�����ֵ����Ϊ�ա�
'lpVerb
'   ����: LPCTSTR
'   һ���ַ�������Ϊ���ʣ�ָ��Ҫִ�еĲ���������ν�ʼ�ȡ�����ض����ļ����ļ��С�ͨ��������Ŀ�ݲ˵��п��õĲ������ǿ��õ�ν�ʡ��˲�������ΪNULL������������£�������ã���ʹ��Ĭ��ν�ʡ����û�У���ʹ�á�open��ν�ʡ��������ν�ʶ������ã�ϵͳ��ʹ��ע������г��ĵ�һ��ν�ʡ����õ�ν����:
'   edit    �����༭�������ĵ����б༭�����lpFile�����ĵ��ļ���������ʧ�ܡ�
'   explore ���lpFileָ�����ļ��С�
'   find    ��lpDirectoryָ����Ŀ¼������������
'   open    ����lpFile����ָ�������Ŀ�������ļ����ļ��С�
'   print   ��ӡlpFileָ�����ļ������lpFile�����ĵ��ļ�����ú�����ʧ�ܡ�
'   properties    ��ʾ�ļ����ļ��е����ԡ�
'lpFile
'   ����: LPCTSTR
'   ��null��β���ַ����ĵ�ַ�����ַ���ָ���ļ����������ƣ�ShellExecuteEx���ڸ��ļ��������ִ��lpVerb����ָ���Ĳ�����ShellExecuteEx����֧�ֵ�ϵͳע���ν�ʰ�����ִ���ļ����ĵ��ļ��ġ�open������ע���ӡ���������ĵ��ļ��ġ�print��������Ӧ�ó�������Ѿ�ͨ��ϵͳע��������Shellν�ʣ�����.avi��.wav�ļ��ġ�play����Ҫָ��Shell���ƿռ�����봫����ȫ�޶��Ľ�����������fMask����������SEE_MASK_INVOKEIDLIST��־��
'   ע�⣬���������SEE_MASK_INVOKEIDLIST��־�������ʹ��lpFile��lpIDList�ֱ�������ļ�ϵͳ·����PIDL��ʶ�������������ֵ֮һ����lpfile��lpidlist��
'   ע�⣬���·��û�а����������У���ٶ���ǰĿ¼��
'lpParameters
'   ����: LPCTSTR
'   ��ѡ�ġ�����Ӧ�ó����������null��β���ַ����ĵ�ַ�����������ÿո�ָ������lpFile��Աָ����һ���ĵ��ļ�����ôlpParametersӦ��ΪNULL��
'lpDirectory
'   ����: LPCTSTR
'   ��ѡ�ġ���null��β���ַ����ĵ�ַ�����ַ���ָ������Ŀ¼�����ơ�����ó�ԱΪ�գ���ʹ�õ�ǰĿ¼��Ϊ����Ŀ¼��
'nShow
'   ����:int
'   ����ġ�ָ��Ӧ�ó����ʱ��ʾ��ʽ�ı�־;ShellExecute�����г���SW_ֵ֮һ�����lpFileָ����һ���ĵ��ļ�����ñ�־�����򵥵ش��ݸ���ص�Ӧ�ó�����δ�����ȡ����Ӧ�ó���
'hInstApp
'   ����: ʵ�����
'   ���������SEE_MASK_NOCLOSEPROCESS������ShellExecuteEx���óɹ����򽫸ó�Ա����Ϊ����32��ֵ���������ʧ�ܣ���������ΪSE_ERR_XXX����ֵ����ֵָʾʧ�ܵ�ԭ�򡣾���Ϊ�˼���16λWindowsӦ�ó���hInstApp������Ϊһ��HINSTANCE����������һ��������HINSTANCE����ֻ�ܱ�ת��Ϊһ��int���ͣ�����32������SE_ERR_XXX���������бȽϡ�
'   SE_ERR_FNF (2)  �ļ�δ�ҵ���
'   SE_ERR_PNF (3)  ·��û���ҵ���
'   SE_ERR_ACCESSDENIED (5) �ܾ����ʡ�
'   SE_ERR_OOM (8)  �ڴ治�㡣
'   SE_ERR_DLLNOTFOUND (32) û���ҵ���̬���ӿ⡣
'   SE_ERR_SHARE (26)   �޷�����򿪵��ļ���
'   SE_ERR_ASSOCINCOMPLETE (27) �ļ�������Ϣ��������
'   SE_ERR_DDETIMEOUT (28)  DDE������ʱ��
'   SE_ERR_DDEFAIL (29) DDE����ʧ�ܡ�
'   SE_ERR_DDEBUSY (30) DDE������æ��
'   SE_ERR_NOASSOC (31) �ļ����������á�
'lpIDList
'   ����: LPVOID
'   ����ITEMIDLIST�ṹ(PCIDLIST_ABSOLUTE)�ĵ�ַ���ýṹ����Ψһ��ʶҪִ�е��ļ������ʶ���б����fMask��Ա������SEE_MASK_IDLIST��SEE_MASK_INVOKEIDLIST������Ըó�Ա��
'lpClass
'   ����: LPCTSTR
'   һ����null��β���ַ����ĵ�ַ�����ַ���ָ����������֮һ:
'   ProgId������,��Paint.Picture����
'   URIЭ�鷽��������,��http����
'   һ���ļ���չ��������," .txt "��
'   HKEY_CLASSES_ROOT�µ�ע���·������Ϊ����һ������Shellν�ʵ��Ӽ����������������һ������Shellν��ע���ģʽ���Ӽ�������
'   shell\verb name
'   ���fMask������SEE_MASK_CLASSNAME������Ըó�Ա��
'hkeyClass
'   �ļ����͵�ע������������ע�����ķ���Ȩ��Ӧ������ΪKEY_READ�����fMask������SEE_MASK_CLASSKEY������Ըó�Ա��
'dwHotKey
'   ��Ӧ�ó�������ļ��̿�ݷ�ʽ���ͽ׵�����������Կ���룬�߽׵��������η���־(HOTKEYF_)���й����η���־���б���μ�WM_SETHOTKEY��Ϣ�����������fMask������SEE_MASK_HOTKEY������Ըó�Ա��
'DUMMYUNIONNAME
'hIcon
'   �ļ�����ͼ��ľ�������fMask������SEE_MASK_ICON������Ըó�Ա����ֵ����Windows XP������汾��ʹ�á�����Windows Vista�б������ˡ�
'hMonitor
'   Ҫ��������ʾ�ĵ��ļ������ľ�������fMask������SEE_MASK_HMONITOR������Ըó�Ա��
'hProcess
'   ������Ӧ�ó���ľ���������Ա�ڷ���ʱ������ΪNULL������fMask������ΪSEE_MASK_NOCLOSEPROCESS����ʹfMask������ΪSEE_MASK_NOCLOSEPROCESS�����û�������κν��̣�hProcessҲ��NULL�����磬���Ҫ�������ĵ���URL������Internet Explorer��ʵ���Ѿ������У���ô������ʾ���ĵ���û�������½��̣�hProcess��ΪNULL��
'   ע�⣬ShellExecuteEx�������Ƿ���hProcess����ʹ���õĽ����������һ�����̡����磬��ʹ��SEE_MASK_INVOKEIDLIST����IContextMenuʱ��hProcess���᷵�ء�
'@��ע
'    �������ShellExecuteEx���߳�û����Ϣѭ���������̻߳���̽���ShellExecuteEx���غ󲻾���ֹ�������ָ��SEE_MASK_NOASYNC��־������������£������߳̽��޷����DDE�Ի�������ڽ�����Ȩ���ظ�����Ӧ�ó���֮ǰ��ShellExecuteEx������ɶԻ���δ����ɶԻ����ܵ����ĵ��������ɹ���
'    ��������߳���һ����Ϣѭ���������ڵ���ShellExecuteEx���غ󽫴���һ��ʱ�䣬��SEE_MASK_NOASYNC��־�ǿ�ѡ�ġ����ʡ�Ըñ�־��������̵߳���Ϣ�ý��������DDE�Ի�������Ӧ�ó�����Ը���ػָ����ƣ���ΪDDE�Ի������ں�̨��ɡ�
'    ��ʹ��fMask�е�SEE_MASK_FLAG_LOG_USAGE��־�����õĳ����б�ʱ����classic��Windows xp���Ŀ�ʼ�˵��ļ����ǲ�ͬ�ġ�������ʽ�˵�ֻ�������˵��п�ݷ�ʽ�ĵ��������Windows xp���Ĳ˵������˳���˵��п�ݷ�ʽ�ĵ�����ͳ���˵����ݷ�ʽ��Ŀ����������ˣ���lpFile����Ϊmyfile.exe��Ӱ��Windows xp��ʽ�˵��ļ��������۸��ļ���ֱ�������Ļ���ͨ����ݷ�ʽ�����ġ�������ʽ(Ҫ��lpFile����.lnk�ļ���)�����ܵ�Ӱ�졣
'    Ҫ��lpParameters�а���˫���ţ��뽫ÿ�������һ���������������������ʾ����ʾ��
'    sei.lpParameters = "An example: \"\"\"quoted text\"\"\"";
'    �ڱ����У�Ӧ�ó��������������:An��example:�� "quoted text"��
'@Requirements
'Minimum supported client    Windows XP [desktop apps only]
'Minimum supported server    Windows 2000 Server [desktop apps only]
'Header                      Shellapi.h
Private Declare Function ShellExecuteEx Lib "kernel32.dll" Alias "ShellExecuteExA" (pExecInfo As SHELLEXECUTEINFO) As Long
'@ԭ��
'    BOOL ShellExecuteEx(
'      _Inout_ SHELLEXECUTEINFO *pExecInfo
'    );
'@����
'    ��ָ�����ļ�ִ�в�����
'@����
'pExecInfo _Inout_
'    ����:SHELLEXECUTEINFO *
'    ָ��SHELLEXECUTEINFO�ṹ��ָ�룬�ýṹ�����������й�����ִ�е�Ӧ�ó������Ϣ��
'@����ֵ
'    ����ɹ�����TRUE;����,�ٵġ�����GetLastError��ȡ��չ�Ĵ�����Ϣ��
'@��ע
'    ��ΪShellExecuteEx���Խ�ִ��ί�и�ʹ���������ģ��(COM)�����Shell��չ(����Դ�������Ĳ˵��������ν��ʵ��)������Ӧ���ڵ���ShellExecuteEx֮ǰ��ʼ��COM��һЩShell��չ��ҪCOM���̹߳�Ԣ(STA)���͡�����������£�COMӦ��������ʾ��ʼ��:
'    CoInitializeEx(NULL, coinit_apartmentthreads | COINIT_DISABLE_OLE1DDE)
'    ����Щʵ���У�ShellExecuteEx��ʹ����Щ���͵�Shell��չ��������Щʵ����������Ҫ��ʼ��COM��������ˣ���ʹ�ô˺���֮ǰ��ʼ�ն�COM���г�ʼ����һ���ܺõ�ʵ����
'    ��dll���ص�������ʱ���������һ����Ϊ��������������DllMain���������ڼ���������ִ�С��������м�������ʱ����Ҫ����ShellExecuteEx����һ�����Ҫ����ΪShellExecuteEx�ǿ���չ�ģ����������Լ����ڼ����������ڵ�����²����������еĴ��룬�Ӷ�ð�������ķ��գ��Ӷ������߳�����Ӧ��
'    ���ڶ�������������ָ��HWND����lpExecInfoָ���SHELLEXECUTEINFO�ṹ��lpVerb��Ա����Ϊ��Properties������ô��ShellExecuteEx�������κδ��ڶ����ܲ����������ȷ��λ�á�
'    ��������ɹ�������SHELLEXECUTEINFO�ṹ��hInstApp��Ա����Ϊ����32��ֵ���������ʧ�ܣ�hInstApp������ΪSE_ERR_XXX����ֵ����ֵ����ָʾʧ�ܵ�ԭ�򡣾���Ϊ�˼���16λWindowsӦ�ó���hInstApp������Ϊһ��HINSTANCE����������һ��������HINSTANCE����ֻ�ܱ�ת��Ϊint������ֻ����ֵ32��SE_ERR_XXX���������бȽϡ�
'    SE_ERR_XXX����ֵ��Ϊ����ShellExecute���ݶ��ṩ�ġ�Ҫ������׼ȷ�Ĵ�����Ϣ����ʹ��GetLastError�������ܷ�������ֵ֮һ��
'    ��������
'    ERROR_FILE_NOT_FOUND       δ�ҵ�ָ�����ļ���
'    ERROR_PATH_NOT_FOUND       δ�ҵ�ָ����·����
Private Const ERROR_DDE_FAIL                    As Long = &H484
'    ��̬���ݽ���(DDE)����ʧ�ܡ�
Private Const ERROR_NO_ASSOCIATION              As Long = &H483
'    û����ָ�����ļ���չ��������Ӧ�ó���
Private Const ERROR_ACCESS_DENIED               As Long = &H5
'    �ܾ���ָ���ļ��ķ��ʡ�
Private Const ERROR_DLL_NOT_FOUND               As Long = &H485
'    �Ҳ�������Ӧ�ó�������Ŀ��ļ�֮һ��
Private Const ERROR_CANCELLED                   As Long = &H4C7
'    l������ʾ�û���ȡ������Ϣ�������û�ȡ��������
Private Const ERROR_NOT_ENOUGH_MEMORY           As Long = &H8
'    û���㹻���ڴ���ִ��ָ���Ĳ�����
Private Const ERROR_SHARING_VIOLATION           As Long = &H20
'    �����˹����ͻ��
'    ��URL����ʱ������ע��Ӧ�ó����Ա��ڴ���URLʱ�����������ָ��Ӧ�ó���֧����ЩЭ�顣������Ϣ��μ�����ע�ᡣ
'    ��Windows 8��ʼ���������ṩָ��ShellExecuteEx������վ����ָ�룬��֧��ʹ�ø�վ��ķ��񼤻���йظ�����Ϣ����μ�����Ӧ�ó���(ShellExecute��ShellExecuteEx��SHELLEXECUTEINFO)��
'@Requirements
'Minimum supported client       Windows XP [desktop apps only]
'Minimum supported server       Windows 2000 Server [desktop apps only]
'Header                         Shellapi.h
'Library                        shell32.lib
'dll                            Shell32.dll (version 3.51 or later)
'Unicode and ANSI names         ShellExecuteExW (Unicode) And ShellExecuteExA(ANSI)
Private Declare Function CoInitialize Lib "ole32.dll" (ByVal pvReserved As Long) As Long
'@ԭ��
'    HRESULT CoInitialize(
'      _In_opt_ LPVOID pvReserved
'    );
'@����
'    �ڵ�ǰ�߳��ϳ�ʼ��COM�⣬��������ģ�ͱ�ʶΪ���̹߳�Ԣ(STA)���µ�Ӧ�ó���Ӧ�õ���CoInitializeEx�������ǳ�ʼ�����������ʹ��Windows����ʱ,�����Windows::Foundation::Initialize�������
'@����
'pvReserved:��������Ǳ����ģ������ǿյġ�
'@���أ��ú������Է��ر�׼����ֵE_INVALIDARG��E_OUTOFMEMORY��E_UNEXPECTEDֵ���Լ�����ֵ��
'        ���ش�������
'        S_OK:COM��������߳��ϳɹ��س�ʼ���ˡ�
'        S_FALSE:COM���Ѿ�������߳��Ͻ����˳�ʼ����
'        RPC_E_CHANGED_MODE:֮ǰ��CoInitializeEx�ĵ���ָ�����̵߳Ĳ���ģ��Ϊ���̹߳�Ԣ(MTA)����Ҳ���ܱ������������߹�Ԣ�����̹߳�Ԣ��ת���Ѿ�������
'@ע������ڵ��ó�CoGetMalloc����֮����κο⺯��֮ǰ������Ҫ�����߳��ϳ�ʼ��COM�⣬�Ի��һ��ָ���׼��������ָ�룬�Լ��ڴ���亯����
'        ���������̵߳Ĳ���ģ��֮�󣬾Ͳ��ܸ�������
'        ����ǰ��ʼ��Ϊ���̵߳ĳ����ж�CoInitialize�ĵ��ý�ʧ�ܣ�������RPC_E_CHANGED_MODE��
'        CoInitializeEx�ṩ����CoInitialize��ͬ�Ĺ��ܣ����ṩ��һ����������ʽ��ָ���̵߳Ĳ���ģ�͡�
'        CoInitialize����CoInitializeEx����������ģ��ָ��Ϊ���̵߳Ĺ�Ԣ��
'        ���쿪����Ӧ�ó���Ӧ�õ���CoInitializeEx��������CoInitialize��
'        ͨ����COM��ֻ���߳��ϳ�ʼ��һ�Ρ�
'        ��ͬһ�߳��϶�CoInitialize��CoInitializeEx�ĺ������ý���ɹ���ֻҪ���ǲ���ͼ�ı䲢��ģ�ͣ����ǻ᷵��S_FALSE.��
'        Ҫ���ŵعر�COM�⣬ÿ���ɹ�����CoInitialize��CoInitializeEx����������S_FALSE.�ĵ��ã�������ͨ����Ӧ�ĵ�����������Ӧ�ĵ��á�
'        ���ǣ�Ӧ�ó����еĵ�һ���̵߳���CoInitialize��0(��ʹ��COINIT_APARTMENTTHREADED��CoInitializeEx)�������ǵ��ÿ��������һ���̡߳�
'        ������STA�϶�CoInitialize�ĵ��ý���ʧ�ܣ�Ӧ�ó����޷�������
'        ��Ϊû�а취���ƽ����ڵķ����������ػ�ж�ص�˳�����Բ�Ҫ��DllMain�����е���CoInitialize��CoInitializeEx�������
'@Requirements
'Minimum supported client       Windows 2000 Professional [desktop apps only]
'Minimum supported server       Windows 2000 Server [desktop apps only]
'Header                         Objbase.h
'Library                        Ole32.lib
'dll                            Ole32.dll
Private Declare Sub CoUninitialize Lib "ole32.dll" ()
'@ԭ��
'    void CoUninitialize(void);
'@����
'    �رյ�ǰ�߳��ϵ�COM�⣬ж���̼߳��ص�����dll���ͷ��߳�ά����������Դ����ǿ�����е�RPC�������߳��Ϲرա�
'ע�����
'    һ���̱߳���Ϊ����CoUninitialize ��CoInitializeEx������ÿ���ɹ����õ���һ��CoUninitialize ����������S_FALSE���κε��á�ֻ�жԳ�ʼ����CoInitialize��CoInitializeEx���ö�Ӧ�ĵ��ò��ܹر�����
'    �� OleInitialize ���ñ���ͨ���� OleUninitialize������ƽ�⡣OleUninitialize ��������CoUninitialize ����˵���OleUninitialize ��Ӧ�ó���Ҳ����Ҫ���� CoUninitialize��
'    Ӧ����Ӧ�ó���ر�ʱ����CoUninitialize ��������Ӧ�ó��������������ڲ�ͨ��������Ϣѭ�����COM����е����һ�ε��á���������ڿ��ŵĻỰ�����������һ��ģʽ��Ϣѭ������Ϊ���COMӦ�ó��������������������κδ��������Ϣ��ͨ��������Ϣ��CoUninitialize ȷ��Ӧ�ó����ڽ��յ����й������Ϣ֮ǰ�����˳���Non-COM��Ϣ������
'    ��Ϊû�а취���ƽ����ڵķ����������ػ�ж�ص�˳�����Բ�Ҫ��DllMain�����е���CoInitialize��CoInitializeEx�������
'@Requirements
'Minimum supported client       Windows 2000 Professional [desktop apps only]
'Minimum supported server       Windows 2000 Server [desktop apps only]
'Header                         Objbase.h
'Library                        Ole32.lib
'dll                            Ole32.dll
Private Declare Function CoInitializeEx Lib "ole32.dll" (ByVal pvReserved As Long, ByVal dwCoInit As Long) As Long
'@ԭ��
'    HRESULT CoInitializeEx(
'      _In_opt_ LPVOID pvReserved,
'      _In_     DWORD  dwCoInit
'    );
'@����
'    ��ʼ��COM�⹩�����߳�ʹ�ã������̵߳Ĳ���ģ�ͣ�������ҪʱΪ�̴߳���һ���µ�Ԫ��
'    �����ʹ��Windows����ʱapi��������ͬʱʹ��COM��Windows����ʱ�����Ӧ�õ���Windows::Foundation::Initialize����ʼ���̣߳�������CoInitializeEx����ʼ������COM�����˵�Ѿ��㹻�ˡ�
'@����
'pvReserved _In_opt_
'   �����˲������ұ���Ϊ�ա�
'dwCoInit   _In_
'   �̵߳Ĳ���ģ�ͺͳ�ʼ��ѡ��˲�����ֵȡ��COINITö�١�����ʹ��COINIT�е��κ�ֵ��ϣ���coinit_apartmentthread��coinit_multithread��־����ͬʱ���á�Ĭ��ֵ��coinit_multithread��
'@���أ��ú������Է��ر�׼����ֵE_INVALIDARG��E_OUTOFMEMORY��E_UNEXPECTEDֵ���Լ�����ֵ��
'        ���ش�������
'        S_OK:������߳��ϳɹ���ʼ����COM��
'        S_FALSE:COM���Ѿ�������߳��ϳ�ʼ����
'        RPC_E_CHANGED_MODE:ǰ���CoInitializeEx�ĵ��ý����̵߳Ĳ���ģ��ָ��Ϊ���̵߳�Ԫ(MTA)����Ҳ���ܱ����Ѿ������˴������̹߳�Ԣ�����̹߳�Ԣ�ĸ��ġ�
'@ע������
'    ����ʹ��COM���ÿ���̣߳�CoInitializeEx�������ٵ���һ�Σ�����ͨ��ֻ����һ�Ρ�ֻҪͨ����ͬ�Ĳ�����־��ͬһ���߳̿��Զ�ε���CoInitializeEx�����Ǻ�������Ч���÷���S_FALSE��Ҫ���߳������ŵعر�COM�⣬ÿ���ɹ���CoInitialize��CoInitializeEx����(��������S_FALSE���κε���)������ͨ����Ӧ��CoUninitialize���ý���ƽ�⡣
'    �ڵ��ó�CoGetMalloc֮����κο⺯��֮ǰ����Ҫ���߳��ϳ�ʼ��COM�⣬�Ի��ָ���׼��������ָ����ڴ���亯��������COM����������CO_E_NOTINITIALIZED��
'    �̵߳Ĳ���ģ�����úú󣬾Ͳ��ܸ������ˡ�����ǰ����ʼ��Ϊ���̵߳ĵ�Ԫ�Ͻ���Э��ʼ���ĵ��ý�ʧ�ܣ�������RPC_E_CHANGED_MODE��
'    �ڵ��̹߳�Ԣ(STA)�д����Ķ�������乫Ԣ���߳̽��շ������ã���˵��ñ����л�������ֻ������Ϣ���б߽�(������PeekMessage��SendMessage����ʱ)���ڶ��̵߳�Ԫ(MTA)��COM�߳��ϴ����Ķ�������ܹ���ʱ�������������̵߳ķ������á���ͨ�����ڶ��̶߳���Ĵ�����ʵ��ĳ����ʽ�Ĳ������ƣ�ʹ��ͬ��ԭ��(���ٽ�Ρ��ź����򻥳���)������������������ݡ�������Ϊ�������̵߳�Ԫ(NTA)�����еĶ���STA��MTA�е��̵߳���ʱ�����߳̽����䵽NTA��������߳�������CoInitializeEx������ý�ʧ�ܲ�����RPC_E_CHANGED_MODE��
'    ����OLE���������̰߳�ȫ�ģ�OleInitialize����ʹ��coinit_apartmentthread��־����CoInitializeEx����ˣ�Ϊ���̶߳��󲢷��Գ�ʼ���ĵ�Ԫ����ʹ��OleInitialize���õ����ԡ�
'    ��Ϊ�޷����ƽ����ڷ��������ػ�ж�ص�˳�����Բ�Ҫ��DllMain��������CoInitialize��CoInitializeEx��CoUninitialize��
'@Requirements
'Minimum supported client       Windows 2000 Professional [desktop apps only]
'Minimum supported server       Windows 2000 Server [desktop apps only]
'Minimum supported phone         Windows Phone 8
'Header                         Objbase.h
'Library                        Ole32.lib
'dll                            Ole32.dll
Private Enum COINIT
    COINIT_APARTMENTTHREADED = &H2
    COINIT_MULTITHREADED = &H0
    COINIT_DISABLE_OLE1DDE = &H4
    COINIT_SPEED_OVER_MEMORY = &H8
End Enum
'@����
'    ȷ�����ڴ��̴߳����Ķ���Ĵ�����õĲ���ģ�͡��������ģ�Ϳ����ǵ��̵߳ģ�Ҳ�����Ƕ��̵߳ġ�
'@����
'COINIT_APARTMENTTHREADED
'   Ϊ���̶߳��󲢷���ʼ���߳�(�����ע��)��
'COINIT_MULTITHREADED
'   ��ʼ�����̶߳��󲢷����߳�(�μ���ע)��
'COINIT_DISABLE_OLE1DDE
'   ΪOLE1֧�ֽ���DDE��
'COINIT_SPEED_OVER_MEMORY
'   �����ڴ�ʹ������������ܡ�
'@��ע
'    ��ͨ������CoInitializeEx��ʼ��һ���߳�ʱ��������ͨ��ָ��COINIT��һ����Ա��Ϊ���ĵڶ�����������ѡ������ʼ��Ϊ���̻߳��Ƕ��̡߳���ָ����δ�����̴߳������κζ���Ĵ�����ã�������Ĳ����ԡ�
'    �����̣߳���Ȼ�������߳�ִ�У����л����д���ĵ��ã�Ҫ����ö���ķ�������������߳�����������ͬһ���߳��ϵĹ�Ԣ/�̴߳������ǡ����⣬����ֻ�ܵ�����Ϣ���б߽硣�����������л���ͨ������Ҫ����������д�����Ĵ����У������ڴ�������б����PeekMessage��SendMessage�ĵ��ã���Щ���ò��ܱ������������û��ͬһ��Ԫ/�߳��е���������ĵ��ô�ϡ�
'    ���߳�(Ҳ��Ϊ�����߳�)���������̴߳����Ķ���ķ����ĵ������κ��߳������С��������ÿ��ܷ�������ͬ�ķ�������ͬ�Ķ����ͬʱ�����̶߳��󲢷���Ϊ���̡߳�����̺Ϳ���������ṩ����ߵ����ܣ�����������˶ദ����Ӳ������Ϊ�Զ���ĵ��ò������κη�ʽ���л���Ȼ��������ζ�Ŷ���Ĵ������ǿ���Լ��Ĳ���ģ�ͣ�ͨ��ͨ��ʹ��ͬ��ԭ������ٽ�Ρ��ź����򻥳��������⣬���ڶ��󲻿��Ʒ��������̵߳������ڣ���˶�����(���̱߳��ش洢��)���ܲ��洢�κ��ض����̵߳�״̬��
'    ע�⣬���̵߳�Ԫ���ڷ�gui�̡߳����̵߳�Ԫ�е��̲߳�Ӧ��ִ��UI������������ΪUI�߳���Ҫ��Ϣ�ã���COM��Ϊ���̵߳�Ԫ�е��̱߳�����Ϣ��
'@Requirements
'Minimum supported client       Windows 2000 Professional [desktop apps only]
'Minimum supported server       Windows 2000 Server [desktop apps only]
'Header                         Objbase.h
Private Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
'@ԭ��
'    void MoveMemory(
'      _In_       PVOID  Destination,
'      _In_ const VOID   *Source,
'      _In_       SIZE_T Length
'    );
'@����
'    ���ڴ���һ��λ���ƶ�����һ��λ�á�
'@����
'Destination
'   ָ���ƶ�Ŀ�����ʼ��ַ��ָ�롣
'Source
'   ָ��Ҫ�ƶ����ڴ�����ʼ��ַ��ָ�롣
'Length
'   Ҫ�ƶ����ڴ��Ĵ�С�����ֽ�Ϊ��λ��
'@����ֵ
'   ��
'@��ע
'    �������������ΪRtlMoveMemory����������ʵ���������ṩ�ġ��йظ�����Ϣ����μ�WinBase.h��Winnt.h��
'    Դ���Ŀ�������ص���
'@��ȫ
'    ��һ������Destination�����㹻��������Դ�ĳ����ֽ�;���򣬿��ܻᷢ�����������������������ʳ�ͻ���������������£��������߽���ִ�д���ע�����Ľ��̣�����ܵ��¶�Ӧ�ó���ľܾ����񹥻������Destination�ǻ��ڶ�ջ�Ļ���������������ˡ���ע�⣬���һ������Length��Ҫ���Ƶ�Ŀ����ֽ�����������Ŀ��Ĵ�С��
'@Requirements
'Minimum supported client       Windows XP [desktop apps only]
'Minimum supported server       Windows Server 2003 [desktop apps only]
'Header                         WinBase.h (include Windows.h)
'dll                            kernel32.dll
Private Declare Function lstrlen Lib "kernel32.dll" Alias "lstrlenW" (ByVal lpString As Long) As Long
Private Declare Function lstrcpy Lib "kernel32" Alias "lstrcpyW" (ByVal lpString1 As Long, ByVal lpString2 As Long) As Long
'@ԭ��
'    LPTSTR WINAPI lstrcpy(
'      _Out_ LPTSTR lpString1,
'      _In_  LPTSTR lpString2
'    );
'@����
'    ���ַ������Ƶ������������� ��Ҫʹ��?����ʹ��StringCchCopy?
'@����
'    lpString1
'    ���ڽ�����lpString2����ָ����ַ������ݵĻ������������������㹻���԰����ַ�����������ֹnull�ַ���
'    lpString2
'    Ҫ���Ƶ���null��β���ַ���?
'@����ֵ
'    ��������ɹ�������ֵ��ָ�򻺳�����ָ�롣
'    �������ʧ�ܣ�����ֵΪNULL, lpString1���ܲ�����NULL��β��
'@��ע
'    ʹ��ϵͳ��˫�ֽ��ַ���(DBCS)�汾������ʹ�ô˺�������DBCS�ַ�����
'    ���Դ��������Ŀ�껺�����ص�����lstrcpy��������δ�������Ϊ��
'@��ȫ����
'    ����ȷ��ʹ�ô˺�������Ӧ�ó���İ�ȫ�ԡ��˺���ʹ�ýṹ���쳣����(SEH)��׽����Υ����������󡣵����������׽��SEH����ʱ��������NULL������ֹ�ַ�����Ҳ��֪ͨ�����ߴ��󡣵��÷����ܰ�ȫ�ؼٶ����������ǿռ䲻�㡣
'    lpString1�����㹻��������lpString2�ͽ���'\0'��������ܻᷢ�������������
'    �����������Ӧ�ó�������లȫ�����ԭ������������ʳ�ͻ�����ܻᵼ�¶�Ӧ�ó���ľܾ����񹥻������������£���������������������߽���ִ�д���ע�����Ľ��̣��ر������lpString1�ǻ��ڶ�ջ�Ļ�������
'    ����ʹ��StringCchCopy;ʹ��StringCchCopy(������,sizeof(����)/ sizeof(����[0]),src);,��ʶ�����������벻��һ��ָ���ʹ��StringCchCopy(������,ARRAYSIZE(����),src);,����ʶ��,������ָ��,�����߸��𴫵����ַ���ָ����ڴ�Ĵ�С��
'@Requirements
'Minimum supported client       Windows 2000 Professional [desktop apps only]
'Minimum supported server       Windows 2000 Server [desktop apps only]
'Header                         Winbase.h (include Windows.h)
'Library                        kernel32.lib
'dll                            kernel32.dll
'Unicode and ANSI names         lstrcpyW (Unicode) And lstrcpyA(ANSI)
Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
'@ԭ��
'    BOOL WINAPI GetUserName(
'      _Out_   LPTSTR  lpBuffer,
'      _Inout_ LPDWORD lpnSize
'    );
'@����
'    �����뵱ǰ�̹߳������û�����
'    ʹ��GetUserNameEx������ָ���ĸ�ʽ�����û�����������Ϣ��IADsADSystemInfo�ӿ��ṩ��
'@����
Private Const UNLEN                             As Long = 256
'lpBuffer _Out_
'   ָ�򻺳�����ָ�룬���ڽ����û��ĵ�¼��������û������������޷����������û�����������ʧ�ܡ���������С(UNLEN + 1)�ַ���������󳤶ȵ��û�����������ֹnull�ַ���UNLEN��lmcon .h�ж��塣
'lpnSize _Inout_
'   ������ʱ�����������TCHARs��ָ��lpBuffer�������Ĵ�С�������ʱ���������ո��Ƶ���������TCHARs��������������ֹnull�ַ���
'   ���lpBuffer̫С�������ͻ�ʧ�ܣ�GetLastError����ERROR_INSUFFICIENT_BUFFER���˲�����������Ļ�������С��������ֹ���ַ���
Private Const ERROR_INSUFFICIENT_BUFFER         As Long = &H7A
'@����ֵ
'   ��������ɹ�������ֵΪ����ֵ��lpnSizeָ��ı����������Ƶ�lpBufferָ���Ļ�������TCHARs��������������ֹnull�ַ���
'   �������ʧ�ܣ�����ֵΪ�㡣Ҫ��ȡ��չ�Ĵ�����Ϣ�������GetLastError��
'@��ע
'   �����ǰ�߳�����ģ����һ���ͻ�����GetUserName�����������߳�����ģ��Ŀͻ������û�����
'   ����������ڡ���������ʻ��µĽ��̵���GetUserName, lpBuffer�з��ص��ַ������ܻ���Windows�汾�Ĳ�ͬ����ͬ����Windows XP�У����ء�NETWORK SERVICE���ַ�������Windows Vista�У����ء�<HOSTNAME>$���ַ�����
'@Requirements
'Minimum supported client       Windows 2000 Professional [desktop apps only]
'Minimum supported server       Windows 2000 Server [desktop apps only]
'Header                         Winbase.h (include Windows.h)
'Library                        advapi32.lib
'dll                            advapi32.dll
'Unicode and ANSI names         GetUserNameW (Unicode) And GetUserNameA(ANSI)
Private Declare Function GetUserNameEx Lib "Secur32.dll" Alias "GetUserNameExA" (ByVal NameFormat As Long, ByVal lpNameBuffer As String, lpnSize As Long) As Long
'@����
'    ����������̹߳������û���������ȫ��������ơ�������ָ���������Ƶĸ�ʽ��
'    ����߳�����ģ��ͻ�����GetUserNameEx�����ؿͻ��������ơ�
'@ԭ��
'    BOOLEAN WINAPI GetUserNameEx(
'      _In_    EXTENDED_NAME_FORMAT NameFormat,
'      _Out_   LPTSTR               lpNameBuffer,
'      _Inout_ PULONG               lpnSize
'    );
'@����
'NameFormat _In_
'   ���Ƶĸ�ʽ���ò�����EXTENDED_NAME_FORMATö�����͵�ֵ���������������ġ�����û��ʻ��������У���ֻ֧��NameSamCompatible��
'lpNameBuffer  _Out_
'   ָ����ָ����ʽ�������ƵĻ�������ָ�롣���������������ֹnull�ַ��Ŀռ䡣
'lpnSize    _Inout_
'   ������ʱ�����������TCHARs��ָ��lpNameBuffer�������Ĵ�С����������ɹ�����������ո��Ƶ���������TCHARs����������������ֹnull�ַ���
'   ���lpNameBuffer̫С�������ͻ�ʧ�ܣ�GetLastError����ERROR_MORE_DATA���ò�����������Ļ�������С(��Unicode�ַ�Ϊ��λ)(�����Ƿ�ʹ��Unicode)��������ֹnull�ַ���
'@����ֵ
'   ��������ɹ�������ֵΪ����ֵ��
'   �������ʧ�ܣ�����ֵΪ�㡣Ҫ��ȡ��չ�Ĵ�����Ϣ�������GetLastError�����ܵ�ֵ�����������ݡ�
'   ���ش�������
Private Const ERROR_MORE_DATA                   As Long = &HEA
'   lpNameBuffer������̫С��lpnSize����������������������ֽ�����
Private Const ERROR_NO_SUCH_DOMAIN              As Long = &H54B
'   ���������������ִ�в���
Private Const ERROR_NONE_MAPPED                 As Long = &H534
'   �û�����ָ����ʽ�в����á�
'@Requirements
'Minimum supported client       Windows 2000 Professional [desktop apps only]
'Minimum supported server       Windows 2000 Server [desktop apps only]
'Header                         Secext.h (include Security.h)
'Library                        Secur32.lib
'dll                            Secur32.dll
'Unicode and ANSI names         GetUserNameExW (Unicode) And GetUserNameExA(ANSI)
Public Enum EXTENDED_NAME_FORMAT
    NameUnknown = 0
    NameFullyQualifiedDN = 1
    NameSamCompatible = 2
    NameDisplay = 3
    NameUniqueId = 6
    NameCanonical = 7
    NameUserPrincipal = 8
    NameCanonicalEx = 9
    NameServicePrincipal = 10
    NameDnsDomain = 12
End Enum
'@����
'    ָ��Ŀ¼����������Ƶĸ�ʽ��
'@����
'NameUnknown
'   δ֪�������͡�
'NameFullyQualifiedDN
'   ��ȫ�޶���ר������(���磬CN=Jeff Smith,OU=Users,DC=Engineering,DC=Microsoft,DC=Com)��
'NameSamCompatible
'   �����ʻ���(���磬Engineering\JSmith)��������İ汾�������÷�б��(\\)��
'NameDisplay
'   һ�����Ѻá�����ʾ����(���磬Jeff Smith)����ʾ���Ʋ�һ���Ƕ�������ר������(RDN)��
'NameUniqueId
'   IIDFromString�������ص�GUID�ַ���(���磬{4fa050f0-f561-11cf-bdd9-00aa003a77b6})��
'NameCanonical
'   �����Ĺ淶����(���磬engineering.microsoft.com/software/someone)��ֻ����İ汾����һ����б��(/)��
'NameUserPrincipal
'   �û���������(���磬someone@example.com)��
'NameCanonicalEx
'   ��NameCanonical��ͬ��ֻ�����ұߵ�б��(/)���滻Ϊһ�������ַ�(\n)����ʹ��ֻ����������Ҳ�����(���磬engineering.microsoft.com/software\nJSmith)��
'NameServicePrincipal
'   ͨ�÷�����������(���磬www/www.microsoft.com@microsoft.com)��
'NameDnsDomain
'   �����б�ܺ�SAM�û�����DNS������
'@Requirements
'Minimum supported client       Windows 2000 Professional [desktop apps only]
'Minimum supported server       Windows 2000 Server [desktop apps only]
'Header                         Secext.h (include Security.h)
Private Declare Function CheckTokenMembership Lib "advapi32.dll" (ByVal TokenHandle As Long, ByVal SidToCheck As Long, IsMember As Long) As Long
'@ԭ��
'    BOOL WINAPI CheckTokenMembership(
'      _In_opt_ HANDLE TokenHandle,
'      _In_     PSID   SidToCheck,
'      _Out_    PBOOL  IsMember
'    );
'@����
'    CheckTokenMembership����ȷ���Ƿ��ڷ���������������ָ���İ�ȫ��ʶ��(SID)�������ȷ��Ӧ�ó����������Ƶ����Ա��ϵ����Ҫʹ��CheckTokenMembershipEx������
'@����
'TokenHandle _In_opt_
'   �������Ƶľ�������������ж����Ƶ�TOKEN_QUERY����Ȩ�����Ʊ�����ģ�����ơ�
'   ���TokenHandleΪ�գ�CheckTokenMembership��ʹ�õ����̵߳�ģ�����ơ�����߳�û��ģ�⣬�ú����������̵߳�������������ģ�����ơ�
'SidToCheck _In_
'   ָ��SID�ṹ��ָ�롣CheckTokenMembership�������������Ƶ��û�����SID���Ƿ���ڴ�SID��
'IsMember  _Out_
'   ָ����ռ�����ı�����ָ�롣���SID���ڲ��Ҿ���SE_GROUP_ENABLED���ԣ���IsMember����TRUE;���򣬷���FALSE��
'@����ֵ
'   ��������ɹ�������ֵΪ���㡣
'   �������ʧ�ܣ�����ֵΪ�㡣Ҫ��ȡ��չ�Ĵ�����Ϣ�������GetLastError��
'@��ע
'   CheckTokenMembership��������ȷ�������������Ƿ�ͬʱ���ں�����SID�Ĺ��̡�
'   ��ʹ�����д���SID��ϵͳҲ���ܲ����ڷ��ʼ����ʹ�ø�SID��SID���ܱ����ã�����ֻ��SE_GROUP_USE_FOR_DENY_ONLY���ԡ���ϵͳֻʹ��SID�ڽ��з��ʼ��ʱ�������Ȩ���йظ�����Ϣ����μ����������е�SID���ԡ�
'   ���TokenHandle�������Ƶ����ƣ��������TokenHandleΪNULL�����ҵ����̵߳�ǰ��Ч�������������Ƶ����ƣ�CheckTokenMembership�������SID�Ƿ�����������Ƶ�SID�б��С�
'@Requirements
'Minimum supported client       Windows XP [desktop apps | UWP apps]
'Minimum supported server       Windows Server 2003 [desktop apps | UWP apps]
'Header                         Winbase.h (include Windows.h)
'Library                        advapi32.lib
'dll                            advapi32.dll
Private Declare Function ProcessIdToSessionId Lib "kernel32.dll" (ByVal dwProcessId As Long, pSessionId As Long) As Long
'@ԭ��
'    BOOL ProcessIdToSessionId(
'      DWORD dwProcessId,
'      DWORD *pSessionId
'    );
'@����
'    ������ָ�����̹�����Զ���������Ự��
'@����
'dwProcessId
'    ָ�����̱�ʶ��?ʹ��GetCurrentProcessId����������ǰ���̵Ľ��̱�ʶ��?
'pSessionId
'    ָ��һ��������ָ�룬�ñ���������������ָ�����̵�Զ���������Ự�ı�ʶ����Ҫ������ǰ���ӵ�����̨�ĻỰ�ı�ʶ������ʹ��WTSGetActiveConsoleSessionId������
'@����ֵ
'    ��������ɹ�������ֵΪ����ֵ��
'    �������ʧ�ܣ�����ֵΪ�㡣Ҫ��ȡ��չ�Ĵ�����Ϣ�������GetLastError��
'@��ע
'    �����߱������ָ�����̵�PROCESS_QUERY_INFORMATION����Ȩ���йظ�����Ϣ����μ����̰�ȫ�ͷ���Ȩ�ޡ�
'@Requirements
'Minimum supported client       Windows XP [desktop apps | UWP apps]
'Minimum supported server       Windows Server 2003 [desktop apps | UWP apps]
'Header                         Winbase.h (include Windows.h)
'Library                        Kernel32.lib
'dll                            Kernel32.dll
Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long
'@ԭ��
'    DWORD GetCurrentProcessId(
'
'    );
'@����
'    �������ý��̵Ľ��̱�ʶ����
'@����
'    �������û�в���?
'@����ֵ
'    ����ֵ�ǵ��ý��̵Ľ��̱�ʶ��?
'@��ע
'    �ڽ�����ֹ֮ǰ�����̱�ʶ��������ϵͳ��Ψһ�ر�ʶ���̡�
'@Requirements
'Minimum supported client       Windows XP [desktop apps | UWP apps]
'Minimum supported server       Windows Server 2003 [desktop apps | UWP apps]
'Header                         processthreadsapi.h (include Windows Server 2003, Windows Vista, Windows 7, Windows Server 2008 Windows Server 2008 R2, Windows.h)
'Library                        Kernel32.lib
'dll                            Kernel32.dll
Private Declare Function CopySid Lib "advapi32" (ByVal nDestinationSidLength As Long, pDestinationSid As Long, ByVal pSourceSid As Long) As Long
'@ԭ��
'    BOOL WINAPI CopySid(
'      _In_  DWORD nDestinationSidLength,
'      _Out_ PSID  pDestinationSid,
'      _In_  PSID  pSourceSid
'    );
'@����
'    CopySid��������ȫ��ʶ��(SID)���Ƶ���������
'@����
'    nDestinationSidLength
'    ָ������SID�����Ļ������ĳ���(���ֽ�Ϊ��λ)��
'    pDestinationSid
'    ָ�򻺳�����ָ�룬�û���������ԴSID�ṹ�ĸ�����
'    pSourceSid
'    ָ��SID�ṹ��ָ�룬�ú������ýṹ���Ƶ���pDestinationSid����ָ��Ļ�������
'@����ֵ
'    ��������ɹ�������ֵΪ���㡣
'    �������ʧ�ܣ�����ֵΪ�㡣Ҫ��ȡ��չ�Ĵ�����Ϣ�������GetLastError��
'@��ע
'    Ӧ�ó������ʹ��CopySid�����ڷ�������(���磬��TOKEN_GROUPS�ṹ��)�и���һ��SID���Ա��ڷ��ʿ�����Ŀ(ACE)��ʹ�á�
'@Requirements
'Minimum supported client       Windows XP [desktop apps | UWP apps]
'Minimum supported server       Windows Server 2003 [desktop apps | UWP apps]
'Header                         Sddl.h
'Library                        advapi32.lib
'dll                            advapi32.dll
Private Declare Function GetLengthSid Lib "advapi32" (pSid As Long) As Long
'@ԭ��
'    DWORD WINAPI GetLengthSid(
'      _In_ PSID pSid
'    );
'@����
'    GetLengthSid����������Ч��ȫ��ʶ��(SID)�ĳ��ȣ����ֽ�Ϊ��λ��
'@����
'pSid
'    ���س���ΪSID�ṹ��ָ��?����ṹ����Ϊ����Ч��?
'����ֵ
'    ���SID�ṹ����Ч�ģ�����ֵ����SID�ṹ�ĳ���(���ֽ�Ϊ��λ)��
'    ���SID�ṹ��Ч���򷵻�ֵδ���塣�ڵ���GetLengthSid֮ǰ����SID���ݸ�IsValidSid����������֤SID�Ƿ���Ч��
'@Requirements
'Minimum supported client       Windows XP [desktop apps | UWP apps]
'Minimum supported server       Windows Server 2003 [desktop apps | UWP apps]
'Header                         Sddl.h
'Library                        advapi32.lib
'dll                            advapi32.dll
Private Declare Function ConvertSidToStringSid Lib "advapi32" Alias "ConvertSidToStringSidW" (pSid As Any, StringSid As Long) As Long
'@ԭ��
'    BOOL ConvertSidToStringSid(
'      _In_  PSID   Sid,
'      _Out_ LPTSTR *StringSid
'    );
'@����
'    ConvertSidToStringSid��������ȫ��ʶ��(SID)ת��Ϊ�ʺ���ʾ���洢������ַ�����ʽ��Ҫ���ַ�����ʽSIDת������Ч�ĺ���SID�������ConvertStringSidToSid������
'@����
'Sid
'    ָ��Ҫת����SID�ṹ��ָ��?
'StringSid
'    ָ�������ָ�룬�ñ�������ָ����null��β��SID�ַ�����ָ�롣Ҫ�ͷŷ��صĻ������������LocalFree������
'@����ֵ
'   ��������ɹ�������ֵΪ���㡣
'   �������ʧ�ܣ�����ֵΪ�㡣Ҫ��ȡ��չ�Ĵ�����Ϣ�������GetLastError��GetLastError�������ܷ������´������֮һ��
'���ش�������
'   ERROR_NOT_ENOUGH_MEMORY    �ڴ治��?
Private Const ERROR_INVALID_SID                 As Long = &H539
'   SID��Ч?
Private Const ERROR_INVALID_PARAMETER           As Long = &H57
'   ����һ����������һ����Ч��ֵ?��ͨ����һ����Ч��ָ��
'@��ע
'    ConvertSidToStringSid������SID�ַ���ʹ�ñ�׼��S-R-I-S-S����ʽ���й�SID�ַ�����ʾ���ĸ�����Ϣ����μ�SID���
'@Requirements
'Minimum supported client       Windows XP [desktop apps | UWP apps]
'Minimum supported server       Windows Server 2003 [desktop apps | UWP apps]
'Header                         Sddl.h
'Library                        advapi32.lib
'dll                            advapi32.dll
'Unicode and ANSI names         ConvertSidToStringSidW (Unicode) And ConvertSidToStringSidA(ANSI)
Private Declare Function LookupAccountSid Lib "advapi32.dll" Alias "LookupAccountSidA" (ByVal lpSystemName As String, ByVal Sid As Long, ByVal name As String, cbName As Long, ByVal ReferencedDomainName As String, cbReferencedDomainName As Long, peUse As Long) As Long
'@ԭ��
'    BOOL WINAPI LookupAccountSid(
'      _In_opt_  LPCTSTR       lpSystemName,
'      _In_      PSID          lpSid,
'      _Out_opt_ LPTSTR        lpName,
'      _Inout_   LPDWORD       cchName,
'      _Out_opt_ LPTSTR        lpReferencedDomainName,
'      _Inout_   LPDWORD       cchReferencedDomainName,
'      _Out_     PSID_NAME_USE peUse
'    );
'@����
'    LookupAccountSid�������ܰ�ȫ��ʶ��(SID)��Ϊ���롣��������SID���ʻ����ƺ��ҵ���SID�ĵ�һ��������ơ�
'@����
'lpSystemName(,��ѡ)
'   ָ��ָ��Ŀ����������null��β���ַ�����ָ�롣����ַ���������Զ�̼���������ơ�����ò���ΪNULL�����ڱ���ϵͳ�Ͽ�ʼ�ʻ�����ת��������޷��ڱ���ϵͳ�Ͻ��������ƣ���˺���������ʹ�ñ���ϵͳ���ε�����������������ơ�ͨ����ֻ�е��ʻ�λ�ڲ������ε������Ҹ����м������������֪ʱ����ΪlpSystemNameָ��һ��ֵ��
'lpSid [��]
'   ָ��Ҫ���ҵ�SID��ָ��?
'lpName(,��ѡ)
'   ָ�򻺳�����ָ�룬�û���������һ����null��β���ַ��������ַ���������lpSid������Ӧ���ʻ�����
'cchName [,]
'   On inputָ��lpName�������Ĵ�С(��TCHARsΪ��λ)�����������Ϊ������̫С��cchNameΪ���ʧ�ܣ���cchName��������Ļ�������С��������ֹnull�ַ���
'lpReferencedDomainName(,��ѡ)
'   ָ�򻺳�����ָ�룬�û���������һ����null��β���ַ��������ַ��������ҵ��ʻ�����������ơ�
'   �ڷ������ϣ�Ϊ���ؼ�����İ�ȫ���ݿ��еĴ�����ʻ����ص������Ƿ�������Ϊ���������������
'   �ڹ���վ�ϣ����ؼ�����İ�ȫ���ݿ���Ϊ������ʻ����ص�������ϵͳ���һ������ʱ�����������(��������б��)���������������Ʒ������ģ��򽫼������ؾ�������Ϊ������ֱ����������ϵͳΪֹ��
'   ��Щ�ʻ�����ϵͳԤ�ȶ����?Ϊ��Щ�ʻ����ص�������BUILTIN?
'cchReferencedDomainName [,]
'   On input����TCHARs��ָ��lpReferencedDomainName�������Ĵ�С�����������Ϊ������̫С��ʧ�ܣ�����cchReferencedDomainNameΪ�㣬��cchReferencedDomainName��������Ļ�������С��������ֹnull�ַ���
'peUse [��]
'   ָ��һ��������ָ�룬�ñ�������һ����ʾ�ʻ����͵�SID_NAME_USEֵ��
'@����ֵ
'    ��������ɹ����������ط��㡣
'    �������ʧ�ܣ��������㡣Ҫ��ȡ��չ�Ĵ�����Ϣ�������GetLastError��
'@��ע
'   LookupAccountSid�������ȼ��һ����֪SID�б���ͼΪָ����SID�ҵ�һ�����ơ�����ṩ��SID����֪��SID����Ӧ���ú�����������õĺ͹����϶���ı����ʻ����������������������������ʶ��İ�ȫ��ʶ������������SIDǰ׺��Ӧ�Ŀ�������м�顣
'   ��������Ҳ���SID���ʻ�����GetLastError������error_none_mapping��������糬ʱ��ֹ�����������ƣ��ͻᷢ���������������û�ж�Ӧ�ʻ�����SIDҲ�ᷢ����������������ʶ��¼�Ự�ĵ�¼SID��
'   ���˲���SID�ĵ����ʻ����������ʻ�����ȷ�����ε����ʻ��⣬LookupAccountSid�����Բ���SID��ɭ�����κ�������κ��ʻ�������ֻ������ɭ���ʻ�SIDhistory�ֶ��е�SID��SIDhistory�ֶδ洢����һ�����ƶ��������ʻ���ǰSID��Ҫ����SID, LookupAccountSid��ѯforest��ȫ��Ŀ¼��
'@Requirements
'Minimum supported client       Windows XP [desktop apps | UWP apps]
'Minimum supported server       Windows Server 2003 [desktop apps | UWP apps]
'Header                         Sddl.h
'Library                        advapi32.lib
'dll                            advapi32.dll
'Unicode and ANSI names         LookupAccountSidW (Unicode) And LookupAccountSidA(ANSI)
Private Declare Function LookupAccountName Lib "advapi32.dll" Alias "LookupAccountNameW" (ByVal lpSystemName As Long, ByVal lpAccountName As Long, ByRef Sid As Any, ByRef cbSid As Long, ByVal ReferencedDomainName As Long, ByRef cbReferencedDomainName As Long, ByRef peUse As Long) As Long
'@ԭ��
'    BOOL WINAPI LookupAccountName(
'      _In_opt_  LPCTSTR       lpSystemName,
'      _In_      LPCTSTR       lpAccountName,
'      _Out_opt_ PSID          Sid,
'      _Inout_   LPDWORD       cbSid,
'      _Out_opt_ LPTSTR        ReferencedDomainName,
'      _Inout_   LPDWORD       cchReferencedDomainName,
'      _Out_     PSID_NAME_USE peUse
'    );
'@����
'    LookupAccountName��������ϵͳ�����ʻ�����Ϊ���롣�������ʻ��İ�ȫ��ʶ��(SID)���ҵ��ʻ�����������ơ�LsaLookupNames���������Լ���������ʻ�?
'@����
'lpSystemName(,��ѡ)
'   ָ����null��β���ַ�����ָ�룬���ַ���ָ��ϵͳ�����ơ�����ַ���������Զ�̼���������ơ�������ַ���Ϊ�գ����ڱ���ϵͳ�Ͽ�ʼ�ʻ�����ת��������޷��ڱ���ϵͳ�Ͻ��������ƣ���˺���������ʹ�ñ���ϵͳ���ε�����������������ơ�ͨ����ֻ�е��ʻ�λ�ڲ������ε������Ҹ����м������������֪ʱ����ΪlpSystemNameָ��һ��ֵ��
'lpAccountName [��]
'   ָ����null��β���ַ�����ָ�룬���ַ���ָ���ʻ�����
'   ʹ��domain_name\user_name��ʽ�е���ȫ�޶��ַ�������ȷ��LookupAccountName�ҵ��������е��ʻ���
'Sid(,��ѡ)
'   һ��ָ�򻺳�����ָ�룬�û�����������lpAccountName������ָ����ʻ������Ӧ��SID�ṹ������ò���ΪNULL����cbSid����Ϊ�㡣
'cbSid [,]
'   ָ�������ָ�롣������ʱ����ֵָ��Sid�������Ĵ�С(���ֽ�Ϊ��λ)�����������Ϊ������̫С��ʧ�ܣ�����cbSidΪ�㣬��˱�������������Ļ�������С��
'ReferencedDomainName(,��ѡ)
'   ָ�򻺳�����ָ�룬�û����������ҵ��ʻ�����������ơ�����û�����ӵ���ļ�������˻��������ռ�������ơ�����ò���ΪNULL��������������Ļ�������С��
'cchReferencedDomainName [,]
'   ָ�������ָ�롣������ʱ����ֵָ��ReferencedDomainName�������Ĵ�С(��TCHARs��)�����������Ϊ������̫С��ʧ�ܣ���˱�������������Ļ�������С��������ֹnull�ַ������ReferencedDomainName����ΪNULL����ò�������Ϊ�㡣
'peUse [��]
'   ָ��SID_NAME_USEö�����͵�ָ�룬������ָʾ��������ʱ�ʻ������͡�
'@����ֵ
'    ��������ɹ����������ط��㡣
'    �������ʧ�ܣ��������㡣Ҫ��ȡ��չ�Ĵ�����Ϣ�������GetLastError��
'@��ע
'    LookupAccountName�������ȼ��һ����֪SID�б���ͼΪָ���������ҵ�SID�������������֪��SID����Ӧ���ú�����������õĺ͹����϶���ı����ʻ�������������������������û���ҵ������ƣ����������
'    ʹ����ȫ�޶����ʻ���(���磬domain_name\user_name)�����ǹ���������(���磬user_name)����ȫ�޶�������ȷ�ģ�������ִ�в���ʱ�ṩ���õ����ܡ��ú�����֧����ȫ�޶���DNS����(���磬example.example.com\user_name)���û���������(UPN)(���磬someone@example.com)��
'    ���˲��ұ����ʻ����������ʻ�����ʽ�����ε����ʻ��⣬LookupAccountName�����Բ��������������е������ʻ������ơ�
'@Requirements
'Minimum supported client       Windows XP [desktop apps | UWP apps]
'Minimum supported server       Windows Server 2003 [desktop apps | UWP apps]
'Header                         Sddl.h
'Library                        advapi32.lib
'dll                            advapi32.dll
'Unicode and ANSI names         LookupAccountNameW (Unicode) And LookupAccountNameA(ANSI)
Private Declare Function LocalFree Lib "kernel32" (hMem As Long) As Long
'@ԭ��
'    HLOCAL WINAPI LocalFree(
'      _In_ HLOCAL hMem
'    );
'@����
'    �ͷ�ָ���ı����ڴ����ʹ������Ч?
'    ע�⣬�������ڴ��������ȣ����غ����Ŀ��������ṩ�����Ը��١���Ӧ�ó���Ӧ��ʹ�öѺ����������ĵ�����Ӧ��ʹ�ñ��غ������йظ�����Ϣ����μ�ȫ�ֺͱ��غ�����
'@����
'hMem
'    �����ڴ����ľ��?��������LocalAlloc��LocalReAlloc��������?ʹ��GlobalAlloc�ͷ��ڴ��ǲ���ȫ��?
'@����ֵ
'    ��������ɹ�������ֵΪNULL��
'    �������ʧ�ܣ�����ֵ���ڱ����ڴ����ľ����Ҫ��ȡ��չ�Ĵ�����Ϣ�������GetLastError��
'@��ע
'    ���������ͼ���ͷ��ڴ�֮������޸��ڴ棬����ܻᷢ�����𻵻����ɷ��ʳ�ͻ�쳣(EXCEPTION_ACCESS_VIOLATION)��
'    ���hMem����ΪNULL , LocalFree�����Ըò���������NULL?
'    LocalFree�������ͷ�һ���������ڴ���󡣱��������ڴ����������������㡣LocalLock��������һ�������ڴ���󣬲�������������1��LocalUnlock���������������������������1��Ҫ��ȡ�����ڴ���������������ʹ��LocalFlags������
'    ���Ӧ�ó�����ϵͳ�ĵ��԰汾�����У�LocalFree������һ����Ϣ���������ͷ���һ�������Ķ���������ڵ���Ӧ�ó���LocalFree�����ͷ���������֮ǰ����һ���ϵ㡣����������֤Ԥ�ڵ���Ϊ��Ȼ�����ִ�С�
'@Requirements
'Minimum supported client       Windows XP [desktop apps | UWP apps]
'Minimum supported server       Windows Server 2003 [desktop apps | UWP apps]
'Header                         WinBase.h (include Windows.h)
'Library                        Kernel32.lib
'dll                            Kernel32.dll
Private Declare Function WTSQueryUserToken Lib "wtsapi32" (ByVal SessionId As Long, phToken As Long) As Long
'@ԭ��
'    BOOL WTSQueryUserToken(
'      ULONG   SessionId,
'      Phandle phToken
'    );
'@����
'    ��ȡ�ỰIDָ���ĵ�¼�û������������ơ�Ҫ�ɹ����ô˺���������Ӧ�ó��������LocalSystem�ʻ������������У�������SE_TCB_NAME��Ȩ��
'    ����:WTSQueryUserToken�����ڸ߶����εķ��񡣷����ṩ�߱���ʹ�þ��棬�����ڵ��ô˺���ʱй©�û����ơ������ṩ�߱�����ʹ�����ƾ��֮��ر����ƾ����
'@����
'SessionId
'    Զ���������Ự��ʶ�����ڷ��������������е��κγ���ĻỰ��ʶ����Ϊ0(0)��������ʹ��WTSEnumerateSessions����������ָ��RD�Ự���������������лỰ�ı�ʶ����
'    Ҫ�ܹ�Ϊ�����û��ĻỰ��ѯ��Ϣ������Ҫ���в�ѯ��ϢȨ�ޡ��йظ�����Ϣ����μ�Զ���������Ȩ�ޡ�Ҫ�޸ĻỰ��Ȩ�ޣ���ʹ��Զ������������ù����ߡ�
'phToken
'    ��������ɹ��������һ��ָ���ѵ�¼�û������ƾ����ָ�롣ע�⣬�������close����������ܹر���������
'@����ֵ
'    ��������ɹ�������ֵΪ����ֵ��phToken����ָ���û��������ơ�
'    �������ʧ�ܣ�����ֵΪ�㡣Ҫ��ȡ��չ�Ĵ�����Ϣ�������GetLastError�������������У�GetLastError���Է������´���֮һ��
'@��ע
'    �й������Ƶ���Ϣ����μ��������ơ��й��ʻ���Ȩ�ĸ�����Ϣ����μ�Զ���������Ȩ�޺���Ȩ������
'    �й�����ʻ���������Ȩ����Ϣ����μ�LocalSystem�ʻ���
'@Requirements
'Minimum supported client       Windows Vista
'Minimum supported server       Windows Server 2008
'Header                         wtsapi32.h
'Library                        Wtsapi32.lib
'dll                            Wtsapi32.dll
Private Declare Function ImpersonateLoggedOnUser Lib "advapi32" (ByVal hToken As Long) As Long
'@ԭ��
'    BOOL WINAPI ImpersonateLoggedOnUser(
'      _In_ HANDLE hToken
'    );
'@����
'    ImpersonateLoggedOnUser������������߳�ģ���¼�û��İ�ȫ�����ġ��û������ƾ����ʾ��
'@����
'    hToken
'    ��ʾ�ѵ�¼�û������������ƻ�ģ��������Ƶľ�����������ͨ������LogonUser��CreateRestrictedToken��DuplicateToken��DuplicateTokenEx��OpenProcessToken��OpenThreadToken�������ص����ƾ�������hToken�������Ƶľ���������Ʊ������TOKEN_QUERY��TOKEN_DUPLICATE����Ȩ�����hToken��ģ�����Ƶľ���������Ʊ������TOKEN_QUERY��TOKEN_IMPERSONATE����Ȩ��
'@����ֵ
'    ��������ɹ�������ֵΪ���㡣
'    �������ʧ�ܣ�����ֵΪ�㡣Ҫ��ȡ��չ�Ĵ�����Ϣ�������GetLastError��
'@��ע
'    ģ�⽫�������߳��˳������RevertToSelfΪֹ?
'    �����̲߳���Ҫ���е���ImpersonateLoggedOnUser���κ��ض���Ȩ?
'    �����ImpersonateLoggedOnUser�ĵ���ʧ�ܣ���ģ��ͻ������ӣ����������̵İ�ȫ�������з����ͻ���������������Ը߶���Ȩ�ʻ�(��LocalSystem)����Ϊ������ĳ�Ա���У����û������ܹ�ִ�в�����ִ�еĲ�������ˣ�ʼ�ռ����õķ���ֵ�Ǻ���Ҫ�ģ��������ʧ�ܣ������������;��Ҫ����ִ�пͻ�������
'    ����ģ�⺯��������ImpersonateLoggedOnUser���������ģ�⣬�������һ��Ϊ��:
'        ���Ƶ�����ģ�⼶��С��SecurityImpersonation������SecurityIdentification��securityannamed��
'        �����߾���SeImpersonatePrivilege��Ȩ?
'        һ������(����÷���¼�Ự�е���һ������)ͨ��LogonUser��LsaLogonUser����ʹ����ʽƾ�ݴ������ơ�
'        ���������֤�ı�ʶ���������ͬ?
'    ����SP1�͸���汾��Windows XP: ��֧��SeImpersonatePrivilege��Ȩ?
'    �й�ģ��ĸ�����Ϣ����μ��ͻ���ģ�⡣
'@Requirements
'Minimum supported client       Windows XP
'Minimum supported server       Windows Server 2003
'Header                         Advapi32.h
'Library                        Advapi32.lib
'dll                            Advapi32.dll
Private Declare Function RevertToSelf Lib "advapi32" () As Long
'@ԭ��
'    BOOL WINAPI RevertToSelf(void);
'@����
'    RevertToSelf������ֹ�Կͻ���Ӧ�ó����ģ�⡣
'@����
'@����ֵ
'    ��������ɹ�������ֵΪ���㡣
'    �������ʧ�ܣ�����ֵΪ�㡣Ҫ��ȡ��չ�Ĵ�����Ϣ�������GetLastError��
'@��ע
'    ����Ӧ����ʹ��DdeImpersonateClient��ImpersonateDdeClientWindow��ImpersonateLoggedOnUser��ImpersonateNamedPipeClient��ImpersonateSelf��ImpersonateAnonymousToken��SetThreadToken��������κ�ģ��֮�����RevertToSelf������
'    ʹ��RpcImpersonateClient����ģ��ͻ�����RPC�������������RpcRevertToSelf��RpcRevertToSelfEx������ģ��?
'    ���RevertToSelfʧ�ܣ�Ӧ�ó��򽫼����ڿͻ��������������У����ǲ����ʵġ����RevertToSelfʧ�ܣ�Ӧ�ùرս��̡�
'@Requirements
'Minimum supported client       Windows XP
'Minimum supported server       Windows Server 2003
'Header                         Advapi32.h
'Library                        Advapi32.lib
'dll                            Advapi32.dll
Private Declare Function DuplicateTokenEx Lib "advapi32" (ByVal hExistingToken As Long, ByVal dwDesiredAcces As Long, ByVal lpTokenAttribute As Long, ByVal ImpersonatonLevel As SECURITY_IMPERSONATION_LEVEL, ByVal tokenType As TOKEN_TYPE, phNewToken As Long) As Long
'@ԭ��
'    BOOL WINAPI DuplicateTokenEx(
'      _In_     HANDLE                       hExistingToken,
'      _In_     DWORD                        dwDesiredAccess,
'      _In_opt_ LPSECURITY_ATTRIBUTES        lpTokenAttributes,
'      _In_     SECURITY_IMPERSONATION_LEVEL ImpersonationLevel,
'      _In_     TOKEN_TYPE                   TokenType,
'      _Out_    PHANDLE                      phNewToken
'    );
'@����
'    DuplicateTokenEx��������һ���µķ������ƣ������Ƹ���һ�����е����ơ��˺������Դ��������ƻ�ģ�����ơ�
'@����
'    hExistingToken
'    �������Ƶľ����ʹ��TOKEN_DUPLICATE���ʴ򿪡�
'    dwDesiredAccess
'    ָ�������Ƶ��������Ȩ�ޡ�DuplicateTokenEx����������ķ���Ȩ�����������Ƶ����ɷ��ʿ����б�(discretionary access control list, DACL)���бȽϣ���ȷ�������ܾ���ЩȨ�ޡ���Ҫ����������������ͬ�ķ���Ȩ�ޣ���ָ���㡣Ҫ����Ե��÷���Ч�����з���Ȩ�ޣ���ָ��MAXIMUM_ALLOWED��
'    �йط������Ƶķ���Ȩ���б���μ��������ƶ���ķ���Ȩ�ޡ�
'    lpTokenAttributes(,��ѡ)
'    ָ��SECURITY_ATTRIBUTES�ṹ��ָ�룬�ýṹΪ������ָ����ȫ����������ȷ���ӽ����Ƿ���Լ̳и����ơ����lpTokenAttributesΪ�գ����ƽ����Ĭ�ϵİ�ȫ�����������Ҳ��ܼ̳о���������ȫ����������һ��ϵͳ���ʿ����б�(SACL)�����ƽ����ACCESS_SYSTEM_SECURITY����Ȩ����ʹ��dwDesiredAccess��û����������
'    Ҫ�������Ƶİ�ȫ�����������������ߣ����÷����������Ʊ������SE_RESTORE_NAME��Ȩ����
'    ImpersonationLevel [��]
'    ��SECURITY_IMPERSONATION_LEVELö����ָ��һ��ֵ����ֵָʾ�����Ƶ�ģ�⼶��
'    tokenType [��]
'    ��TOKEN_TYPEö����ָ������ֵ֮һ?
Private Enum TOKEN_TYPE
    TokenPrimary = 1
'       ����������������CreateProcessAsUser������ʹ�õ���Ҫ����?
    TokenImpersonation = 2
'       �µ�������ģ������?
End Enum
'    phNewToken
'    ָ����������Ƶľ��������ָ��?
'    �������ʹ��������ʱ������close����������ر����ƾ����
'@����ֵ
'    ��������ɹ�����������һ������ֵ��
'    �������ʧ�ܣ��������㡣Ҫ��ȡ��չ�Ĵ�����Ϣ�������GetLastError��
'@��ע
'    DuplicateTokenEx��������������һ��������CreateProcessAsUser������ʹ�õ������ơ�������ģ��ͻ����ķ�����Ӧ�ó��򴴽����пͻ�����ȫ�����ĵ����̡�ע�⣬DuplicateToken����ֻ�ܴ���ģ�����ƣ����CreateProcessAsUser��Ч��
'    ������ʹ��DuplicateTokenEx���������Ƶĵ��ͳ�����������Ӧ�ó��򴴽�һ���̣߳����̵߳���һ��ģ�⺯��(��ImpersonateNamedPipeClient)��ģ��ͻ�����ģ���߳�Ȼ�����OpenThreadToken������������Լ������ƣ�����һ�����пͻ�����ȫ�����ĵ�ģ�����ơ��߳��ڵ���DuplicateTokenExʱָ�����ģ�����ƣ�ָ��TokenPrimary��־��DuplicateTokenEx��������һ�����пͻ��˰�ȫ�����ĵ������ơ�
'@Requirements
'Minimum supported client       Windows XP
'Minimum supported server       Windows Server 2003
'Header                         Advapi32.h
'Library                        Advapi32.lib
'dll                            Advapi32.dll
Private Enum SECURITY_IMPERSONATION_LEVEL
    SecurityAnonymous
    SecurityIdentification
    SecurityImpersonation
    SecurityDelegation
End Enum
'@ԭ��
'    typedef enum _SECURITY_IMPERSONATION_LEVEL {
'      SecurityAnonymous,
'      SecurityIdentification,
'      SecurityImpersonation,
'      SecurityDelegation
'    } SECURITY_IMPERSONATION_LEVEL, *PSECURITY_IMPERSONATION_LEVEL;
'@����
'    SECURITY_IMPERSONATION_LEVELö�ٰ���ָ����ȫģ�⼶���ֵ����ȫģ�⼶����Ʒ��������̿��Դ���ͻ������̽��в����ĳ̶ȡ�
'@����
'    SecurityAnonymous
'    ���������̲��ܻ�ȡ���ڿͻ����ı�ʶ��Ϣ��Ҳ����ģ��ͻ��������Ķ���û�и���ֵ����ˣ�����ANSI C����Ĭ��ֵΪ�㡣
'    SecurityIdentification
'    ���������̿��Ի�ȡ���ڿͻ�������Ϣ�����簲ȫ��ʶ������Ȩ������������ģ��ͻ���������ڵ����Լ��Ķ���ķ������ǳ����ã����磬���������ͼ�����ݿ��Ʒ��ʹ�ü������Ŀͻ�����ȫ��Ϣ�������������ڲ�ʹ������ʹ�ÿͻ�����ȫ�����ĵ�������������������������֤���ߡ�
'    SecurityImpersonation
'    ���������̿������䱾��ϵͳ��ģ��ͻ����İ�ȫ������?������������Զ��ϵͳ��ģ��ͻ���?
'    SecurityDelegation
'    ���������̿�����Զ��ϵͳ��ģ��ͻ����İ�ȫ������?
'--------------------------------------------------------------------------------------------------
'����           RunAsCurrentUser
'����           �ڵ�ǰ�Ự�д�������,����SYSTEM����
'����ֵ         Long
'����б�:
'������         ����                    ˵��
'
'-------------------------------------------------------------------------------------------------
Public Function RunAsCurrentUser(ByVal strApplicationName As String, ByVal strCommandLine As String, Optional ByVal strCurrentDirectory As String) As Long
    Dim hToken      As Long
    Dim hProcessToken   As Long
    Dim hSeToken        As Long
    Dim siInfo      As STARTUPINFO
    Dim piInfo      As PROCESS_INFORMATION
    Dim lngRet      As Long
    Dim hNewToken   As Long
    On Error GoTo ErrH
    Call gobjLog.PushMethod(RLL_AllLog, "ZLHelperMain.mdlRunas.RunAsCurrentUser")
    If glngWinSessionID = 0 Then
        glngWinSessionID = GetCurrentSessionID()
    End If
    
'    If OpenThreadToken(GetCurrentThread(), TOKEN_ALL_ACCESS, True, hProcessToken) = 0 Then
'        gobjLog.LogInfo RLL_LogInfo, "OpenThreadTokenʧ��", "����", GetLastDllErr(Err.LastDllError)
'        If OpenProcessToken(GetCurrentProcess(), TOKEN_ALL_ACCESS, hProcessToken) = 0 Then
'            gobjLog.LogInfo RLL_LogInfo, "OpenProcessTokenʧ��", "����", GetLastDllErr(Err.LastDllError)
'        End If
'    End If
'    If hProcessToken Then
    If WTSQueryUserToken(glngWinSessionID, hProcessToken) <> 0 Then
        If DuplicateTokenEx(hProcessToken, TOKEN_ALL_ACCESS, ByVal 0, ByVal SecurityImpersonation, ByVal TokenPrimary, hNewToken) <> 0 Then
            If SetTokenInformation(hNewToken, TokenSessionId, glngWinSessionID, LenB(glngWinSessionID)) <> 0 Then
                gobjLog.LogInfo RLL_LogInfo, "�Ƿ����Ա�û�", IsAdministrator(), "hToken=", hNewToken
'
'                    If SetTokenInformation(hSeToken, TokenSessionId, glngWinSessionID, LenB(glngWinSessionID)) <> 0 Then
                        If ImpersonateLoggedOnUser(hNewToken) <> 0 Then
                            siInfo.cb = LenB(siInfo)
                            lngRet = CreateProcessAsUser(hNewToken, strApplicationName, strCommandLine, 0&, 0&, False, CREATE_DEFAULT_ERROR_MODE, 0&, strCurrentDirectory, siInfo, piInfo)
                            If lngRet <> 0 Then
                                RunAsCurrentUser = piInfo.dwProcessId
                                CloseHandle hToken
                                CloseHandle piInfo.hThread
                                CloseHandle piInfo.hProcess
                                gobjLog.LogInfo RLL_LogInfo, "CreateProcessAsUser�ɹ�"
                            Else
                                gobjLog.LogInfo RLL_LogInfo, "CreateProcessAsUserʧ��", "����", GetLastDllErr(Err.LastDllError)
                            End If
                        Else
                            gobjLog.LogInfo RLL_LogInfo, "ImpersonateLoggedOnUserʧ��", "����", GetLastDllErr(Err.LastDllError)
                        End If
                        If RevertToSelf() = 0 Then
                            gobjLog.LogInfo RLL_LogInfo, "RevertToSelfʧ��", "����", GetLastDllErr(Err.LastDllError)
                        End If
'                    Else
'                        gobjLog.LogInfo RLL_LogInfo, "SetTokenInformation2ʧ��", "����", GetLastDllErr(Err.LastDllError)
'                    End If
'                Else
'                    gobjLog.LogInfo RLL_LogInfo, "WTSQueryUserTokenʧ��", "����", GetLastDllErr(Err.LastDllError)
'                End If
            Else
                gobjLog.LogInfo RLL_LogInfo, "SetTokenInformation1ʧ��", "����", GetLastDllErr(Err.LastDllError)
            End If
            CloseHandle hNewToken
        Else
            gobjLog.LogInfo RLL_LogInfo, "DuplicateTokenExʧ��", "����", GetLastDllErr(Err.LastDllError)
        End If
        CloseHandle hToken
'    End If
    End If
    Call gobjLog.PopMethod(RLL_AllLog, "ZLHelperMain.mdlRunas.RunAsCurrentUser", RunAsCurrentUser)
    Exit Function
ErrH:
    If gobjLog.ErrCenter(RLL_RunError, "ZLHelperMain.mdlRunas.RunAsCurrentUser") = 1 Then
        Resume
    End If
End Function

'--------------------------------------------------------------------------------------------------
'����           RunAsUser
'����           ִ����ָ������Ա���г���
'����ֵ         Boolean
'����б�:
'������         ����                    ˵��
'strUserName    String                  �û���
'strPassword    String                  ����
'strDomainName  String                  �˻���
'strApplicationName String              ����·��
'strCommandLine String                  ������
'strCurrentDirectory    String          ��ǰĿ¼
'-------------------------------------------------------------------------------------------------
Public Function RunAsUser(ByVal strApplicationName As String, ByVal strCommandLine As String, Optional ByVal strCurrentDirectory As String, Optional ByVal strUserName As String, Optional ByVal strPassword As String, Optional ByVal strDomainName As String) As Boolean
    On Error GoTo ErrH
    Call gobjLog.PushMethod(RLL_AllLog, "ZLHelperMain.clsRunas.RunAsUser", strUserName, Sm4EncryptEcb(strPassword), strDomainName, strApplicationName, Sm4EncryptEcb(strCommandLine), strCurrentDirectory)
    If IsWindows2000OrGreater() Then
'        If IsWindowsVistaOrGreater() Then
'            '��ǰ���̲��ǹ���
'            If Not IsProcessRunAsAdmin() And IsAdministrator() Then
'                RunAsUser = RunAsAdmin(strApplicationName, strCommandLine, strCurrentDirectory)
'            Else
'                RunAsUser = RunAsUserW2K(strUserName, strPassword, strDomainName, strApplicationName, strCommandLine, strCurrentDirectory)
'            End If
'        Else
            RunAsUser = RunAsUserW2K(strUserName, strPassword, strDomainName, strApplicationName, strCommandLine, strCurrentDirectory)
'        End If
    Else
        RunAsUser = RunAsUserNT4(strUserName, strPassword, strDomainName, strApplicationName, strCommandLine, strCurrentDirectory)
    End If
    Call gobjLog.PopMethod(RLL_AllLog, "ZLHelperMain.clsRunas.RunAsUser", RunAsUser)
    Exit Function
ErrH:
    If gobjLog.ErrCenter(RLL_RunError, "ZLHelperMain.clsRunas.RunAsUser") = 1 Then
        Resume
    End If
End Function

'--------------------------------------------------------------------------------------------------
'����           GetUserSID
'����           ��ȡ�û�SID���û���ȡע���
'����ֵ         String
'����б�:
'������         ����                    ˵��
'
'-------------------------------------------------------------------------------------------------
Public Function GetUserSID(ByVal strAccountName As String) As String
    Dim lngRet      As Long, bytSid()   As Byte, cbSid  As Long
    Dim strDom      As String, cbDom    As Long
    Dim lpStrSid    As Long, peUse      As Long
    Dim lngError    As Long
    Dim strRet      As String
    
    On Error GoTo ErrH
    Call gobjLog.PushMethod(RLL_AllLog, "ZLHelperMain.mdlRunas.GetUserSID")
    lngRet = LookupAccountName(0, StrPtr(strAccountName), ByVal 0, cbSid, 0, cbDom, peUse)
    lngError = Err.LastDllError
    If lngError = ERROR_INSUFFICIENT_BUFFER Then
        gobjLog.LogInfo RLL_LogInfo, "���������㣬���仺����", "����", lngRet
        ReDim bytSid(cbSid): strDom = String$(cbDom, Chr$(0))
        lngRet = LookupAccountName(0, StrPtr(strAccountName), bytSid(0), cbSid, StrPtr(strDom), cbDom, peUse)
        lngError = Err.LastDllError
        If lngRet = 0 Then
            gobjLog.LogInfo RLL_LogInfo, "LookupAccountName2ʧ��", "����", GetLastDllErr(Err.LastDllError), "����", lngRet
        End If
        If ConvertSidToStringSid(bytSid(0), lpStrSid) Then
            strRet = String$(lstrlen(lpStrSid), Chr(0))
            If lstrcpy(StrPtr(strRet), lpStrSid) <> 0 Then
                If InStr(strRet, Chr$(0)) > 0 Then
                    strRet = Mid(strRet, 1, InStr(strRet, Chr$(0)) - 1)
                End If
            Else
                strRet = ""
                gobjLog.LogInfo RLL_LogInfo, "lstrcpyʧ��", "����", GetLastDllErr(Err.LastDllError)
            End If
            Call LocalFree(lpStrSid)
        Else
            gobjLog.LogInfo RLL_LogInfo, "ConvertSidToStringSidʧ��", "����", GetLastDllErr(Err.LastDllError)
        End If
    ElseIf lngError = 0 Then
        gobjLog.LogInfo RLL_LogInfo, "�û�������", "����", lngRet
    Else
        gobjLog.LogInfo RLL_LogInfo, "LookupAccountNameʧ��", "����", GetLastDllErr(Err.LastDllError), "����", lngRet
    End If
    GetUserSID = strRet
    Call gobjLog.PopMethod(RLL_AllLog, "ZLHelperMain.mdlRunas.GetUserSID", GetUserSID)
    Exit Function
ErrH:
    If gobjLog.ErrCenter(RLL_RunError, "ZLHelperMain.mdlRunas.GetUserSID") = 1 Then
        Resume
    End If
End Function
''--------------------------------------------------------------------------------------------------
''����           GetCurrentSessionSID
''����           ��ȡ��ǰ�Ự��SID��
''����ֵ         String
''����б�:
''������         ����                    ˵��
''
''-------------------------------------------------------------------------------------------------
'Public Function GetCurrentSessionSID(Optional ByVal lngProcessID As Long) As String
'    Dim hProcess            As Long
'    Dim hProcessToken       As Long
'    Dim BufferSize          As Long
'    Dim lResult             As Long
'    Dim i                   As Integer
''    Dim tpTokens            As TOKEN_GROUPS
'    Dim tpTokenGroups()     As SID_AND_ATTRIBUTES
'    Dim lngStrSID           As Long
'    Dim strSID              As String
'    Dim strUserName         As String
'    Dim strDomain           As String
'    Dim lngTmp              As Long
'
'    On Error GoTo ErrH
'    Call gobjLog.PushMethod(RLL_AllLog, "ZLHelperMain.mdlRunas.GetCurrentSessionSID")
'    If lngProcessID = 0 Then
'        If OpenThreadToken(GetCurrentThread(), TOKEN_QUERY, True, hProcessToken) = 0 Then
'            gobjLog.LogInfo RLL_LogInfo, "OpenThreadTokenʧ��", "����", GetLastDllErr(Err.LastDllError)
'            If OpenProcessToken(GetCurrentProcess(), TOKEN_QUERY, hProcessToken) = 0 Then
'                gobjLog.LogInfo RLL_LogInfo, "OpenProcessTokenʧ��", "����", GetLastDllErr(Err.LastDllError)
'            End If
'        End If
'    Else
'
'    End If
'    If hProcessToken Then
'        If GetTokenInformation(hProcessToken, ByVal TokenGroups, 0, 0, BufferSize) = 0 Then ' Determine required buffer size
'            gobjLog.LogInfo RLL_LogInfo, "GetTokenInformationʧ��", "����", GetLastDllErr(Err.LastDllError), "��Ҫ��������С", BufferSize
'        End If
'        If BufferSize Then
'            ReDim InfoBuffer((BufferSize \ 4) - 1) As Long
'            lResult = GetTokenInformation(hProcessToken, ByVal TokenGroups, InfoBuffer(0), BufferSize, BufferSize)
'            If lResult = 0 Then
'                gobjLog.LogInfo RLL_LogInfo, "GetTokenInformationʧ��", "����", GetLastDllErr(Err.LastDllError)
'                Exit Function
'            End If
'            'TOKEN_GROUPS.GROUPCount��Ա
'            ReDim tpTokenGroups(InfoBuffer(0) - 1)
'            Call MoveMemory(tpTokenGroups(0), InfoBuffer(1), Len(tpTokenGroups(0)) * InfoBuffer(0))
'            For i = 0 To UBound(tpTokenGroups)
'                If IsValidSid(tpTokenGroups(i).Sid) Then
'                    GetCurrentSessionSID = GetCurrentSessionSID & " i=" & i & vbNewLine
''                    If (tpTokenGroups(i).Attributes And SE_GROUP_INTEGRITY) = SE_GROUP_INTEGRITY Then
'                        lngStrSID = 0
'                        strSID = String(256, Chr(0))
'                        If ConvertSidToStringSid(tpTokenGroups(i).Sid, lngStrSID) <> 0 Then
'                            If lstrcpy(strSID, lngStrSID) <> 0 Then
'                                If InStr(strSID, Chr$(0)) > 0 Then
'                                    strSID = Mid(strSID, 1, InStr(strSID, Chr$(0)) - 1)
'                                End If
'                                GetCurrentSessionSID = GetCurrentSessionSID & vbNewLine & strSID
'                            Else
'                                gobjLog.LogInfo RLL_LogInfo, "lstrcpyʧ��", "����", GetLastDllErr(Err.LastDllError)
'                            End If
'                            Call LocalFree(ByVal lngStrSID)
'                        Else
'                            gobjLog.LogInfo RLL_LogInfo, "ConvertSidToStringSidʧ��", "����", GetLastDllErr(Err.LastDllError)
'                        End If
'                        strUserName = String(256, Chr(0))
'                        strDomain = String(256, Chr(0))
'                        If LookupAccountSid(vbNullString, tpTokenGroups(i).Sid, strUserName, 255, strDomain, 255, lngTmp) <> 0 Then
'                            If InStr(strUserName, Chr$(0)) > 0 Then
'                                strUserName = Mid(strUserName, 1, InStr(strUserName, Chr$(0)) - 1)
'                            End If
'                            If InStr(strDomain, Chr$(0)) > 0 Then
'                                strDomain = Mid(strDomain, 1, InStr(strDomain, Chr$(0)) - 1)
'                            End If
'                            GetCurrentSessionSID = GetCurrentSessionSID & "(" & strDomain & "\" & strUserName & ")"
'                        End If
''                    End If
'                End If
'            Next
'        End If
'        Call CloseHandle(hProcessToken)
'    End If
'    Call gobjLog.PopMethod(RLL_AllLog, "ZLHelperMain.mdlRunas.GetCurrentSessionSID", GetCurrentSessionSID)
'    Exit Function
'ErrH:
'    If gobjLog.ErrCenter(RLL_RunError, "ZLHelperMain.mdlRunas.GetCurrentSessionSID") = 1 Then
'        Resume
'    End If
'End Function
'--------------------------------------------------------------------------------------------------
'����           GetCurrentSessionID
'����           ��ȡ��ǰ�ỰID
'����ֵ         Long
'����б�:
'������         ����                    ˵��
'
'-------------------------------------------------------------------------------------------------
Public Function GetCurrentSessionID(Optional ByVal lngCurProcID As Long) As Long
    Dim lngSessionID        As Long
    On Error GoTo ErrH
    Call gobjLog.PushMethod(RLL_AllLog, "ZLHelperMain.clsRunas.GetCurrentSessionID")
    If lngCurProcID = 0 Then
        lngCurProcID = GetCurrentProcessId()
    End If
    If ProcessIdToSessionId(lngCurProcID, lngSessionID) = 0 Then
        gobjLog.LogInfo RLL_LogInfo, "ProcessIdToSessionIdʧ��", "����", GetLastDllErr(Err.LastDllError)
    End If
    GetCurrentSessionID = lngSessionID
    Call gobjLog.PopMethod(RLL_AllLog, "ZLHelperMain.clsRunas.GetCurrentSessionID")
    Exit Function
ErrH:
    If gobjLog.ErrCenter(RLL_RunError, "ZLHelperMain.clsRunas.GetCurrentSessionID") = 1 Then
        Resume
    End If
End Function
'--------------------------------------------------------------------------------------------------
'����           IsProcessRunAsAdmin
'����           ��ǰ�������Թ���ԱȨ�����С�����ԱȨ�޺͹���Ա����ͬһ�����Visita֮�ϣ���׼����Ա������Ȩ�޲��ǹ���Ա������RUnAS
'����ֵ         Boolean
'����б�:
'������         ����                    ˵��
'
'-------------------------------------------------------------------------------------------------
Public Function IsProcessRunAsAdmin() As Boolean
    Dim ntAuthority         As SID_IDENTIFIER_AUTHORITY
    Dim psidAdmin           As Long
    Dim lngRet              As Long
    
    On Error GoTo ErrH
    Call gobjLog.PushMethod(RLL_AllLog, "ZLHelperMain.mdlRunas.IsProcessRunAsAdmin")
    ntAuthority.value(5) = security_nt_authority
    lngRet = AllocateAndInitializeSid(ntAuthority, 2, SECURITY_BUILTIN_DOMAIN_RID, DOMAIN_ALIAS_RID_ADMINS, 0, 0, 0, 0, 0, 0, psidAdmin)
    If lngRet <> 0 Then
        If CheckTokenMembership(0, psidAdmin, lngRet) = 0 Then
            gobjLog.LogInfo RLL_LogInfo, "CheckTokenMembershipʧ��", "����", GetLastDllErr(Err.LastDllError)
        End If
        Call FreeSid(psidAdmin)
    Else
        gobjLog.LogInfo RLL_LogInfo, "AllocateAndInitializeSidʧ��", "����", GetLastDllErr(Err.LastDllError)
    End If
    IsProcessRunAsAdmin = lngRet <> 0
    Call gobjLog.PopMethod(RLL_AllLog, "ZLHelperMain.mdlRunas.IsProcessRunAsAdmin", IsProcessRunAsAdmin)
    Exit Function
ErrH:
    If gobjLog.ErrCenter(RLL_RunError, "ZLHelperMain.mdlRunas.IsProcessRunAsAdmin") = 1 Then
        Resume
    End If
End Function
'--------------------------------------------------------------------------------------------------
'����           IsAdministrator
'����           �жϵ�ǰ�����û��Ƿ��ǹ���Ա���û�,����SYSTEM����
'����ֵ         Boolean
'����б�:
'������         ����                    ˵��
'
'-------------------------------------------------------------------------------------------------
Public Function IsAdministrator() As Boolean
    Dim lnghProcess         As Long
    Dim hProcessToken       As Long
    Dim BufferSize          As Long
    Dim psidAdmin           As Long
    Dim lResult             As Long
    Dim i                   As Integer
'    Dim tpTokens            As TOKEN_GROUPS
    Dim tpTokenGroups()     As SID_AND_ATTRIBUTES
    Dim tpSidAuth           As SID_IDENTIFIER_AUTHORITY

    On Error GoTo ErrH
    Call gobjLog.PushMethod(RLL_AllLog, "ZLHelperMain.clsRunas.IsAdministrator")
    IsAdministrator = False
    tpSidAuth.value(5) = security_nt_authority
    If WTSQueryUserToken(GetCurrentSessionID(), hProcessToken) <> 0 Then
        If hProcessToken Then
            If GetTokenInformation(hProcessToken, ByVal TokenGroups, 0, 0, BufferSize) = 0 Then ' Determine required buffer size
                gobjLog.LogInfo RLL_LogInfo, "GetTokenInformationʧ��", "����", GetLastDllErr(Err.LastDllError), "��Ҫ��������С", BufferSize
            End If
            If BufferSize Then
                ReDim InfoBuffer((BufferSize \ 4) - 1) As Long
                lResult = GetTokenInformation(hProcessToken, ByVal TokenGroups, InfoBuffer(0), BufferSize, BufferSize)
                If lResult = 0 Then
                    gobjLog.LogInfo RLL_LogInfo, "GetTokenInformationʧ��", "����", GetLastDllErr(Err.LastDllError)
                    Exit Function
                End If
                'TOKEN_GROUPS.GROUPCount��Ա
                ReDim tpTokenGroups(InfoBuffer(0) - 1)
                Call MoveMemory(tpTokenGroups(0), InfoBuffer(1), Len(tpTokenGroups(0)) * InfoBuffer(0))
                lResult = AllocateAndInitializeSid(tpSidAuth, 2, SECURITY_BUILTIN_DOMAIN_RID, DOMAIN_ALIAS_RID_ADMINS, 0, 0, 0, 0, 0, 0, psidAdmin)
                If lResult = 0 Then
                    gobjLog.LogInfo RLL_LogInfo, "AllocateAndInitializeSidʧ��", "����", GetLastDllErr(Err.LastDllError)
                    Exit Function
                End If
                If IsValidSid(psidAdmin) Then
                    For i = 0 To UBound(tpTokenGroups)
                        If IsValidSid(tpTokenGroups(i).Sid) Then
                            If EqualSid(ByVal tpTokenGroups(i).Sid, ByVal psidAdmin) Then
                                IsAdministrator = True
                                Exit For
                            End If
                        End If
                    Next
                End If
                If psidAdmin Then Call FreeSid(psidAdmin)
            End If
            Call CloseHandle(hProcessToken)
        End If
    Else
        gobjLog.LogInfo RLL_LogInfo, "WTSQueryUserTokendʧ��", "����", GetLastDllErr(Err.LastDllError)
    End If
    Call gobjLog.PopMethod(RLL_AllLog, "ZLHelperMain.clsRunas.IsAdministrator", IsAdministrator)
    Exit Function
ErrH:
    If gobjLog.ErrCenter(RLL_RunError, "ZLHelperMain.clsRunas.IsAdministrator") = 1 Then
        Resume
    End If
End Function

'--------------------------------------------------------------------------------------------------
'����           IsProcesssAdministrator
'����           �жϵ�ǰ�����û��Ƿ��ǹ���Ա���û�
'����ֵ         Boolean
'����б�:
'������         ����                    ˵��
'
'-------------------------------------------------------------------------------------------------
Public Function IsProcesssAdministrator(Optional ByVal lngProcessID As Long) As Boolean
    Dim lnghProcess         As Long
    Dim hProcessToken       As Long
    Dim BufferSize          As Long
    Dim psidAdmin           As Long
    Dim lResult             As Long
    Dim i                   As Integer
'    Dim tpTokens            As TOKEN_GROUPS
    Dim tpTokenGroups()     As SID_AND_ATTRIBUTES
    Dim tpSidAuth           As SID_IDENTIFIER_AUTHORITY

    On Error GoTo ErrH
    Call gobjLog.PushMethod(RLL_AllLog, "ZLHelperMain.clsRunas.IsProcesssAdministrator")
    IsProcesssAdministrator = False
    tpSidAuth.value(5) = security_nt_authority
    
    If OpenThreadToken(GetCurrentThread(), TOKEN_QUERY, True, hProcessToken) = 0 Then
        gobjLog.LogInfo RLL_LogInfo, "OpenThreadTokenʧ��", "����", GetLastDllErr(Err.LastDllError)
        If lngProcessID = 0 Then
            If OpenProcessToken(GetCurrentProcess(), TOKEN_QUERY, hProcessToken) = 0 Then
                gobjLog.LogInfo RLL_LogInfo, "OpenProcessTokenʧ��", "����", GetLastDllErr(Err.LastDllError)
            End If
        End If
    End If
    If hProcessToken Then
        If GetTokenInformation(hProcessToken, ByVal TokenGroups, 0, 0, BufferSize) = 0 Then ' Determine required buffer size
            gobjLog.LogInfo RLL_LogInfo, "GetTokenInformationʧ��", "����", GetLastDllErr(Err.LastDllError), "��Ҫ��������С", BufferSize
        End If
        If BufferSize Then
            ReDim InfoBuffer((BufferSize \ 4) - 1) As Long
            lResult = GetTokenInformation(hProcessToken, ByVal TokenGroups, InfoBuffer(0), BufferSize, BufferSize)
            If lResult = 0 Then
                gobjLog.LogInfo RLL_LogInfo, "GetTokenInformationʧ��", "����", GetLastDllErr(Err.LastDllError)
                Exit Function
            End If
            'TOKEN_GROUPS.GROUPCount��Ա
            ReDim tpTokenGroups(InfoBuffer(0) - 1)
            Call MoveMemory(tpTokenGroups(0), InfoBuffer(1), Len(tpTokenGroups(0)) * InfoBuffer(0))
            lResult = AllocateAndInitializeSid(tpSidAuth, 2, SECURITY_BUILTIN_DOMAIN_RID, DOMAIN_ALIAS_RID_ADMINS, 0, 0, 0, 0, 0, 0, psidAdmin)
            If lResult = 0 Then
                gobjLog.LogInfo RLL_LogInfo, "AllocateAndInitializeSidʧ��", "����", GetLastDllErr(Err.LastDllError)
                Exit Function
            End If
            If IsValidSid(psidAdmin) Then
                For i = 0 To UBound(tpTokenGroups)
                    If IsValidSid(tpTokenGroups(i).Sid) Then
                        If EqualSid(ByVal tpTokenGroups(i).Sid, ByVal psidAdmin) Then
                            IsProcesssAdministrator = True
                            Exit For
                        End If
                    End If
                Next
            End If
            If psidAdmin Then Call FreeSid(psidAdmin)
        End If
        Call CloseHandle(hProcessToken)
    End If
    Call gobjLog.PopMethod(RLL_AllLog, "ZLHelperMain.clsRunas.IsProcesssAdministrator", IsProcesssAdministrator)
    Exit Function
ErrH:
    If gobjLog.ErrCenter(RLL_RunError, "ZLHelperMain.clsRunas.IsProcesssAdministrator") = 1 Then
        Resume
    End If
End Function

'--------------------------------------------------------------------------------------------------
'����           GetProcessUserName
'����           ��ȡ��ǰ���̵Ĳ���ϵͳ���û���
'����ֵ         String
'����б�:
'������         ����                    ˵��
'lngType        Long                    ��ȡ���Ƶĸ�ʽ
'-------------------------------------------------------------------------------------------------
Public Function GetProcessUserName(Optional lngType As Long = NameSamCompatible) As String
    Dim strTemp     As String
    Dim lngLen      As Long

    On Error GoTo ErrH
    Call gobjLog.PushMethod(RLL_AllLog, "ZLHelperMain.clsRunas.GetProcessUserName", lngType)
    lngLen = UNLEN + 1
    strTemp = String(UNLEN + 1, Chr$(0))
    If GetUserName(strTemp, lngLen) = 0 Then
        gobjLog.LogInfo RLL_LogInfo, "GetUserNameʧ��", "����", GetLastDllErr(Err.LastDllError), "����", lngLen
    End If

'    GetUserNameEx lngType, vbNullString, lngLen
'    strTemp = String(lngLen, Chr$(0))
'    GetUserNameEx NameSamCompatible, strTemp, lngLen
    If InStr(strTemp, Chr$(0)) > 0 Then
        strTemp = Mid(strTemp, 1, InStr(strTemp, Chr$(0)) - 1)
    End If
    GetProcessUserName = strTemp
    Call gobjLog.PopMethod(RLL_AllLog, "ZLHelperMain.clsRunas.GetProcessUserName", GetProcessUserName)
    Exit Function
ErrH:
    If gobjLog.ErrCenter(RLL_RunError, "ZLHelperMain.clsRunas.GetProcessUserName") = 1 Then
        Resume
    End If
End Function

'--------------------------------------------------------------------------------------------------
'����           CheckUserPassword
'����           ������ϵͳ�û��Ƿ���ȷ
'����ֵ         Boolean
'����б�:
'������         ����                    ˵��
'strUserName    String                  �û���
'strPassword    String                  ����
'strDomainName  String                  �˻���
'-------------------------------------------------------------------------------------------------
Public Function CheckUserPassword(ByVal strUserName As String, ByVal strPassword As String, ByVal strDomainName As String) As Boolean
    Dim hToken      As Long
    On Error GoTo ErrH
    Call gobjLog.PushMethod(RLL_AllLog, "ZLHelperMain.clsRunas.CheckUserPassword", strUserName, Sm4EncryptEcb(strPassword), strDomainName)
    CheckUserPassword = LogonUser(strUserName, strDomainName, strPassword, LOGON32_LOGON_INTERACTIVE, LOGON32_PROVIDER_DEFAULT, hToken) <> 0
    If CheckUserPassword Then
        CloseHandle hToken
    Else
        gobjLog.LogInfo RLL_LogInfo, "LogonUserʧ��", "����", GetLastDllErr(Err.LastDllError)
    End If
    Call gobjLog.PopMethod(RLL_AllLog, "ZLHelperMain.clsRunas.CheckUserPassword", CheckUserPassword)
    Exit Function
ErrH:
    If gobjLog.ErrCenter(RLL_RunError, "ZLHelperMain.clsRunas.CheckUserPassword") = 1 Then
        Resume
    End If
End Function

'--------------------------------------------------------------------------------------------------
'����           GetPrivilegeList
'����           ��ȡ���̵�Ȩ���б�
'����ֵ         String
'����б�:
'������         ����                    ˵��
'
'-------------------------------------------------------------------------------------------------
Public Function GetPrivilegeList(Optional ByVal hProc As Long) As String
    Dim hProcessToken       As Long
    Dim BufferSize          As Long
    Dim lngRet              As Long
    Dim i                   As Integer
    Dim tpTokenPrivileges() As LUID_AND_ATTRIBUTES
    Dim strName             As String
    Dim lngCount            As Long
    Dim strReturn           As String
    Dim lpLuid              As LUID
    
    If hProc = 0 Then
        hProc = GetCurrentProcess()
    End If
    lngRet = OpenProcessToken(hProc, TOKEN_QUERY, hProcessToken)
    If hProcessToken <> 0 Then
        If GetTokenInformation(hProcessToken, ByVal TokenPrivileges, 0, 0, BufferSize) = 0 Then ' Determine required buffer size
            gobjLog.LogInfo RLL_LogInfo, "GetTokenInformationʧ��", "����", GetLastDllErr(Err.LastDllError), "��Ҫ��������С", BufferSize
        End If
        If BufferSize Then
            ReDim InfoBuffer((BufferSize \ 4) - 1) As Long
            lngRet = GetTokenInformation(hProcessToken, ByVal TokenPrivileges, InfoBuffer(0), BufferSize, BufferSize)
            If lngRet = 0 Then
                gobjLog.LogInfo RLL_LogInfo, "GetTokenInformationʧ��", "����", GetLastDllErr(Err.LastDllError)
                Exit Function
            End If
            ReDim tpTokenPrivileges(InfoBuffer(0) - 1)
            Call MoveMemory(tpTokenPrivileges(0), InfoBuffer(1), Len(tpTokenPrivileges(0)) * InfoBuffer(0))
            For i = 0 To UBound(tpTokenPrivileges)
                strName = String(256, Chr$(0))
                lngRet = LookupPrivilegeName(vbNullString, tpTokenPrivileges(i).PLUID, strName, Len(strName))
                If InStr(strName, Chr$(0)) > 0 Then
                    strName = Mid(strName, 1, InStr(strName, Chr$(0)) - 1)
                End If
                strReturn = strReturn & i & ":" & strName & "-" & tpTokenPrivileges(i).Attributes & vbNewLine
            Next
        End If
    End If
    GetPrivilegeList = strReturn
End Function

'--------------------------------------------------------------------------------------------------
'����           EnablePrivilege
'����           �������̵�Ȩ��
'����ֵ         Boolean
'����б�:
'������         ����                    ˵��
'
'-------------------------------------------------------------------------------------------------
Public Function EnablePrivilegeTest() As Boolean
    '0: SeIncreaseQuotaPrivilege -0
    If Not EnablePrivilege(, SE_INCREASE_QUOTA_NAME) Then
        gobjLog.LogInfo RLL_LogInfo, SE_INCREASE_QUOTA_NAME & "ʧ��"
    Else
        gobjLog.LogInfo RLL_LogInfo, SE_INCREASE_QUOTA_NAME & "�ɹ�"
    End If
    '1: SeSecurityPrivilege -0
    If Not EnablePrivilege(, SE_SECURITY_NAME) Then
        gobjLog.LogInfo RLL_LogInfo, SE_SECURITY_NAME & "ʧ��"
    Else
        gobjLog.LogInfo RLL_LogInfo, SE_SECURITY_NAME & "�ɹ�"
    End If
    '2: SeTakeOwnershipPrivilege -0
    If Not EnablePrivilege(, SE_TAKE_OWNERSHIP_NAME) Then
        gobjLog.LogInfo RLL_LogInfo, SE_TAKE_OWNERSHIP_NAME & "ʧ��"
    Else
        gobjLog.LogInfo RLL_LogInfo, SE_TAKE_OWNERSHIP_NAME & "�ɹ�"
    End If
    '3: SeLoadDriverPrivilege -0
    If Not EnablePrivilege(, SE_LOAD_DRIVER_NAME) Then
        gobjLog.LogInfo RLL_LogInfo, SE_LOAD_DRIVER_NAME & "ʧ��"
    Else
        gobjLog.LogInfo RLL_LogInfo, SE_LOAD_DRIVER_NAME & "�ɹ�"
    End If
    '4: SeSystemProfilePrivilege -0
    If Not EnablePrivilege(, SE_SYSTEM_PROFILE_NAME) Then
        gobjLog.LogInfo RLL_LogInfo, SE_SYSTEM_PROFILE_NAME & "ʧ��"
    Else
        gobjLog.LogInfo RLL_LogInfo, SE_SYSTEM_PROFILE_NAME & "�ɹ�"
    End If
    '5: SeSystemtimePrivilege -0
    If Not EnablePrivilege(, SE_SYSTEMTIME_NAME) Then
        gobjLog.LogInfo RLL_LogInfo, SE_SYSTEMTIME_NAME & "ʧ��"
    Else
        gobjLog.LogInfo RLL_LogInfo, SE_SYSTEMTIME_NAME & "�ɹ�"
    End If
    '6: SeProfileSingleProcessPrivilege -0
    If Not EnablePrivilege(, SE_PROF_SINGLE_PROCESS_NAME) Then
        gobjLog.LogInfo RLL_LogInfo, SE_PROF_SINGLE_PROCESS_NAME & "ʧ��"
    Else
        gobjLog.LogInfo RLL_LogInfo, SE_PROF_SINGLE_PROCESS_NAME & "�ɹ�"
    End If
    '7: SeIncreaseBasePriorityPrivilege -0
    If Not EnablePrivilege(, SE_INC_BASE_PRIORITY_NAME) Then
        gobjLog.LogInfo RLL_LogInfo, SE_INC_BASE_PRIORITY_NAME & "ʧ��"
    Else
        gobjLog.LogInfo RLL_LogInfo, SE_INC_BASE_PRIORITY_NAME & "�ɹ�"
    End If
    '8: SeCreatePagefilePrivilege -0
    If Not EnablePrivilege(, SE_CREATE_PAGEFILE_NAME) Then
        gobjLog.LogInfo RLL_LogInfo, SE_CREATE_PAGEFILE_NAME & "ʧ��"
    Else
        gobjLog.LogInfo RLL_LogInfo, SE_CREATE_PAGEFILE_NAME & "�ɹ�"
    End If
    '9: SeBackupPrivilege -0
    If Not EnablePrivilege(, SE_BACKUP_NAME) Then
        gobjLog.LogInfo RLL_LogInfo, SE_BACKUP_NAME & "ʧ��"
    Else
        gobjLog.LogInfo RLL_LogInfo, SE_BACKUP_NAME & "�ɹ�"
    End If
    
    '10: SeRestorePrivilege -0
    If Not EnablePrivilege(, SE_RESTORE_NAME) Then
        gobjLog.LogInfo RLL_LogInfo, SE_RESTORE_NAME & "ʧ��"
    Else
        gobjLog.LogInfo RLL_LogInfo, SE_RESTORE_NAME & "�ɹ�"
    End If
    'AAA -11: SeShutdownPrivilege -0
    '12: SeDebugPrivilege -0
    If Not EnablePrivilege(, SE_DEBUG_NAME) Then
        gobjLog.LogInfo RLL_LogInfo, SE_DEBUG_NAME & "ʧ��"
    Else
        gobjLog.LogInfo RLL_LogInfo, SE_DEBUG_NAME & "�ɹ�"
    End If
    '13: SeSystemEnvironmentPrivilege -0
    If Not EnablePrivilege(, SE_SYSTEM_ENVIRONMENT_NAME) Then
        gobjLog.LogInfo RLL_LogInfo, SE_SYSTEM_ENVIRONMENT_NAME & "ʧ��"
    Else
        gobjLog.LogInfo RLL_LogInfo, SE_SYSTEM_ENVIRONMENT_NAME & "�ɹ�"
    End If
    'AAA -14: SeChangeNotifyPrivilege -3
    '15: SeRemoteShutdownPrivilege -0
    If Not EnablePrivilege(, SE_REMOTE_SHUTDOWN_NAME) Then
        gobjLog.LogInfo RLL_LogInfo, SE_REMOTE_SHUTDOWN_NAME & "ʧ��"
    Else
        gobjLog.LogInfo RLL_LogInfo, SE_REMOTE_SHUTDOWN_NAME & "�ɹ�"
    End If
    'AAA -16: SeUndockPrivilege -0
    '17: SeManageVolumePrivilege -0
    If Not EnablePrivilege(, SE_MANAGE_VOLUME_NAME) Then
        gobjLog.LogInfo RLL_LogInfo, SE_MANAGE_VOLUME_NAME & "ʧ��"
    Else
        gobjLog.LogInfo RLL_LogInfo, SE_MANAGE_VOLUME_NAME & "�ɹ�"
    End If
    '18: SeImpersonatePrivilege -3
    If Not EnablePrivilege(, SE_IMPERSONATE_NAME) Then
        gobjLog.LogInfo RLL_LogInfo, SE_IMPERSONATE_NAME & "ʧ��"
    Else
        gobjLog.LogInfo RLL_LogInfo, SE_IMPERSONATE_NAME & "�ɹ�"
    End If
    '19: SeCreateGlobalPrivilege -3
    If Not EnablePrivilege(, SE_CREATE_GLOBAL_NAME) Then
        gobjLog.LogInfo RLL_LogInfo, SE_CREATE_GLOBAL_NAME & "ʧ��"
    Else
        gobjLog.LogInfo RLL_LogInfo, SE_CREATE_GLOBAL_NAME & "�ɹ�"
    End If
    'AAA -20: SeIncreaseWorkingSetPrivilege -0
    'AAA -21: SeTimeZonePrivilege -0
    '22: SeCreateSymbolicLinkPrivilege -0
    If Not EnablePrivilege(, SE_CREATE_SYMBOLIC_LINK_NAME) Then
        gobjLog.LogInfo RLL_LogInfo, SE_CREATE_SYMBOLIC_LINK_NAME & "ʧ��"
    Else
        gobjLog.LogInfo RLL_LogInfo, SE_CREATE_SYMBOLIC_LINK_NAME & "�ɹ�"
    End If
End Function


Public Function EnablePrivilege(Optional ByVal hProc As Long, Optional ByVal strPrivilegeName As String) As Boolean
    Dim hToken As Long
    Dim tmpLuid As LUID
    Dim tkp As TOKEN_PRIVILEGES
    Dim tkpNewButIgnored As TOKEN_PRIVILEGES
    Dim lBufferNeeded As Long
    Dim lngRet As Long
    
    If hProc = 0 Then
        hProc = GetCurrentProcess()
    End If

    lngRet = OpenProcessToken(hProc, TOKEN_ADJUST_PRIVILEGES Or TOKEN_QUERY, hToken)
    If lngRet = 0 Then
        gobjLog.LogInfo RLL_LogInfo, "OpenProcessTokenʧ��", "����", GetLastDllErr(Err.LastDllError)
    End If
    If hToken <> 0 Then
        lngRet = LookupPrivilegeValue(vbNullString, strPrivilegeName, tmpLuid)
        If lngRet = 0 Then
            gobjLog.LogInfo RLL_LogInfo, "LookupPrivilegeValueʧ��", "����", GetLastDllErr(Err.LastDllError)
        End If
        tkp.PrivilegeCount = 1
        tkp.Privileges(0).PLUID = tmpLuid
        tkp.Privileges(0).Attributes = SE_PRIVILEGE_ENABLED
        lngRet = AdjustTokenPrivileges(hToken, 0, tkp, Len(tkp), tkpNewButIgnored, lBufferNeeded)
        If lngRet = 0 Then
            gobjLog.LogInfo RLL_LogInfo, "AdjustTokenPrivilegesʧ��", "����", GetLastDllErr(Err.LastDllError)
        End If
        EnablePrivilege = lngRet <> 0
        CloseHandle hToken
    End If
End Function


Public Function EnablePrivilegeToken(ByVal hToken As Long, Optional ByVal strPrivilegeName As String) As Boolean
    Dim tmpLuid As LUID
    Dim tkp As TOKEN_PRIVILEGES
    Dim tkpNewButIgnored As TOKEN_PRIVILEGES
    Dim lBufferNeeded As Long
    Dim lngRet As Long
    
    If lngRet = 0 Then
        gobjLog.LogInfo RLL_LogInfo, "LookupPrivilegeValueʧ��", "����", GetLastDllErr(Err.LastDllError)
    End If
    tkp.PrivilegeCount = 1
    tkp.Privileges(0).PLUID = tmpLuid
    tkp.Privileges(0).Attributes = SE_PRIVILEGE_ENABLED
    lngRet = AdjustTokenPrivileges(hToken, 0, tkp, Len(tkp), tkpNewButIgnored, lBufferNeeded)
    If lngRet = 0 Then
        gobjLog.LogInfo RLL_LogInfo, "AdjustTokenPrivilegesʧ��", "����", GetLastDllErr(Err.LastDllError)
    End If
    EnablePrivilegeToken = lngRet <> 0
End Function
'--------------------------------------------------------------------------------------------------
'����           RunAsAdmin
'����           �Թ���Ա���г���
'����ֵ         long
'����б�:
'������         ����                    ˵��
'strApplicationName String              ����
'strCommandLine      String              ������
'strDirectory        String              ��ǰ����Ŀ¼
'-------------------------------------------------------------------------------------------------
Private Function RunAsAdmin(ByVal strApplicationName As String, Optional ByVal strCommandLine As String, Optional ByVal strDirectory As String) As Long
    Dim lngRet      As Long
    Dim seInfo      As SHELLEXECUTEINFO
    On Error GoTo ErrH
    Call gobjLog.PushMethod(RLL_AllLog, "ZLHelperMain.clsRunas.RunAsAdmin", strApplicationName, Sm4EncryptEcb(strCommandLine), strDirectory)
    If strApplicationName = "" Then
        If InStr(strCommandLine, " ") > 0 Then
            strApplicationName = Mid(strCommandLine, 1, InStr(strCommandLine, " ") - 1)
            strCommandLine = Mid(strCommandLine, InStr(strCommandLine, " ") + 1)
        End If
    End If
    If strDirectory = "" Then strDirectory = CurDir$()
    strCommandLine = strCommandLine & " -Admin"
    seInfo.cbSize = LenB(seInfo)
'    seInfo.hwnd = 0
'    seInfo.lpVerb = "runas"
'    seInfo.lpFile = strApplicationName
'    seInfo.lpParameters = strCommandLine
'    seInfo.lpDirectory = strDirectory
'    seInfo.nShow = SW_SHOWNORMAL
'    seInfo.hInstApp = 0
'    seInfo.fMask = SEE_MASK_NOCLOSEPROCESS
'    ShellExecuteEx (seInfo)
'a.fMask = SEE_MASK_NOCLOSEPROCESS;
'a.hwnd = NULL;
'a.lpVerb = NULL;
'a.lpFile = "C:\\Program Files\\Internet Explorer\\IEXPLORE.exe";
'a.lpParameters = "";
'a.lpDirectory = NULL;
'a.nShow = SW_SHOW;
'a.hInstApp = NULL;
'
'pid1=GetProcessId(a.hProcess);
    lngRet = ShellExecute(0, "runas", strApplicationName, strCommandLine, strDirectory, SW_SHOWNORMAL)
    If lngRet <= 32 Then
        gobjLog.LogInfo RLL_LogInfo, "ShellExecuteʧ��", "����", GetLastDllErr(Err.LastDllError), "����", lngRet
    End If
    RunAsAdmin = lngRet > 32
    Call gobjLog.PopMethod(RLL_AllLog, "ZLHelperMain.clsRunas.RunAsAdmin")
    Exit Function
ErrH:
    If gobjLog.ErrCenter(RLL_RunError, "ZLHelperMain.clsRunas.RunAsAdmin") = 1 Then
        Resume
    End If
End Function
'--------------------------------------------------------------------------------------------------
'����           IsWindows2000OrGreater
'����           �Ƿ���Window2000֮��汾
'����ֵ         Boolean
'����б�:
'������         ����                    ˵��
'
'-------------------------------------------------------------------------------------------------
Private Function IsWindows2000OrGreater() As Boolean
    Dim osInfo          As OSVERSIONINFOEX
    On Error GoTo ErrH
    Call gobjLog.PushMethod(RLL_AllLog, "ZLHelperMain.clsRunas.IsWindows2000OrGreater")
    osInfo.dwOSVersionInfoSize = Len(osInfo)
    osInfo.szCSDVersion = Space$(128)
    If GetVersionExA(osInfo) = 0 Then
        gobjLog.LogInfo RLL_LogInfo, "GetVersionExAʧ��", "����", GetLastDllErr(Err.LastDllError)
    End If
    IsWindows2000OrGreater = osInfo.dwMajorVersion >= 5
    Call gobjLog.PopMethod(RLL_AllLog, "ZLHelperMain.clsRunas.IsWindows2000OrGreater", IsWindows2000OrGreater)
    Exit Function
ErrH:
    If gobjLog.ErrCenter(RLL_RunError, "ZLHelperMain.clsRunas.IsWindows2000OrGreater") = 1 Then
        Resume
    End If
End Function

'����           IsWindowsVistaOrGreater
'����           �Ƿ���Window Vista֮��汾
'����ֵ         Boolean
'����б�:
'������         ����                    ˵��
'
'-------------------------------------------------------------------------------------------------
Private Function IsWindowsVistaOrGreater() As Boolean
    Dim osInfo          As OSVERSIONINFOEX
    On Error GoTo ErrH
    Call gobjLog.PushMethod(RLL_AllLog, "ZLHelperMain.clsRunas.IsWindowsVistaOrGreater")
    osInfo.dwOSVersionInfoSize = Len(osInfo)
    osInfo.szCSDVersion = Space$(128)
    If GetVersionExA(osInfo) = 0 Then
        gobjLog.LogInfo RLL_LogInfo, "GetVersionExAʧ��", "����", GetLastDllErr(Err.LastDllError)
    End If
    IsWindowsVistaOrGreater = osInfo.dwMajorVersion >= 6
    Call gobjLog.PopMethod(RLL_AllLog, "ZLHelperMain.clsRunas.IsWindowsVistaOrGreater", IsWindowsVistaOrGreater)
    Exit Function
ErrH:
    If gobjLog.ErrCenter(RLL_RunError, "ZLHelperMain.clsRunas.IsWindowsVistaOrGreater") = 1 Then
        Resume
    End If
End Function
'--------------------------------------------------------------------------------------------------
'����           RunAsUserW2K
'����           ��Windows2000����ִ��Runas
'����ֵ         Long                    ���ؽ���ID
'����б�:
'������         ����                    ˵��
'strUserName    String                  �û���
'strPassword    String                  ����
'strDomainName  String                  �˻���
'strApplicationName String              ����·��
'strCommandLine String                  ������
'strCurrentDirectory    String          ��ǰĿ¼
'-------------------------------------------------------------------------------------------------
Private Function RunAsUserW2K(ByVal strUserName As String, ByVal strPassword As String, ByVal strDomainName As String, ByVal strApplicationName As String, Optional ByVal strCommandLine As String, Optional ByVal strCurrentDirectory As String) As Long
    Dim siInfo          As STARTUPINFO
    Dim piInfo          As PROCESS_INFORMATION
    
    Dim strWideUser     As String
    Dim strWideDomain       As String
    Dim strWidePassword     As String
    Dim strWideApp          As String
    Dim strWideCommandLine  As String
    Dim strWideCurrentDir   As String
    Dim lngRet              As Long
    
    On Error GoTo ErrH
    Call gobjLog.PushMethod(RLL_AllLog, "ZLHelperMain.clsRunas.RunAsUserW2K", strUserName, Sm4EncryptEcb(strPassword), strDomainName, strApplicationName, Sm4EncryptEcb(strCommandLine), strCurrentDirectory)
    
    siInfo.cb = LenB(siInfo)
    strWideUser = StrConv(strUserName + Chr$(0), vbUnicode)
    strWideDomain = StrConv(strDomainName + Chr$(0), vbUnicode)
    strWidePassword = StrConv(strPassword + Chr$(0), vbUnicode)
    strWideCommandLine = StrConv(strCommandLine + Chr$(0), vbUnicode)
    If strCurrentDirectory = "" Then strCurrentDirectory = CurDir$()
    strWideCurrentDir = StrConv(strCurrentDirectory + Chr$(0), vbUnicode)
    strWideApp = StrConv(strApplicationName + Chr$(0), vbUnicode)
    
    lngRet = CreateProcessWithLogon(StrPtr(strWideUser), StrPtr(strWideDomain), StrPtr(strWidePassword), LOGON_WITH_PROFILE, StrPtr(strWideApp), StrPtr(strWideCommandLine), CREATE_DEFAULT_ERROR_MODE Or CREATE_NEW_CONSOLE Or CREATE_NEW_PROCESS_GROUP, ByVal 0, StrPtr(strWideCurrentDir), siInfo, piInfo)
    If lngRet = 0 Then
        gobjLog.LogInfo RLL_LogInfo, "CreateProcessWithLogonWʧ��", "����", GetLastDllErr(Err.LastDllError)
    Else
        RunAsUserW2K = piInfo.dwProcessId
        CloseHandle piInfo.hThread
        CloseHandle piInfo.hProcess
    End If
    Call gobjLog.PopMethod(RLL_AllLog, "ZLHelperMain.clsRunas.RunAsUserW2K", RunAsUserW2K)
    Exit Function
ErrH:
    If gobjLog.ErrCenter(RLL_RunError, "ZLHelperMain.clsRunas.RunAsUserW2K") = 1 Then
        Resume
    End If
End Function

'--------------------------------------------------------------------------------------------------
'����           RunAsUserNT4
'����           ��Window2000����ִ��Runas
'����ֵ         Long                    ���ؽ���ID
'����б�:
'������         ����                    ˵��
'strUserName    String                  �û���
'strPassword    String                  ����
'strDomainName  String                  �˻���
'strApplicationName String              ����·��
'strCommandLine String                  ������
'strCurrentDirectory    String          ��ǰĿ¼
'-------------------------------------------------------------------------------------------------
Private Function RunAsUserNT4(ByVal strUserName As String, ByVal strPassword As String, ByVal strDomainName As String, ByVal strApplicationName As String, Optional ByVal strCommandLine As String, Optional ByVal strCurrentDirectory As String) As Long
    Dim hToken      As Long
    Dim siInfo      As STARTUPINFO
    Dim piInfo      As PROCESS_INFORMATION
    Dim lngRet      As Long
    On Error GoTo ErrH
    Call gobjLog.PushMethod(RLL_AllLog, "ZLHelperMain.clsRunas.RunAsUserNT4", strUserName, strPassword, strDomainName, strApplicationName, strCommandLine, strCurrentDirectory)
    If strCurrentDirectory = "" Then strCurrentDirectory = CurDir$()
    lngRet = LogonUser(strUserName, strDomainName, strPassword, LOGON32_LOGON_INTERACTIVE, LOGON32_PROVIDER_DEFAULT, hToken)
    If lngRet <> 0 Then
        siInfo.cb = LenB(siInfo)
        lngRet = CreateProcessAsUser(hToken, strApplicationName, strCommandLine, 0&, 0&, False, CREATE_DEFAULT_ERROR_MODE, 0&, strCurrentDirectory, siInfo, piInfo)
        If lngRet <> 0 Then
            RunAsUserNT4 = piInfo.dwProcessId
            CloseHandle hToken
            CloseHandle piInfo.hThread
            CloseHandle piInfo.hProcess
        Else
            gobjLog.LogInfo RLL_LogInfo, "CreateProcessAsUserʧ��", "����", GetLastDllErr(Err.LastDllError)
            CloseHandle hToken
        End If
        
        CloseHandle hToken
        CloseHandle piInfo.hThread
        CloseHandle piInfo.hProcess
    Else
        gobjLog.LogInfo RLL_LogInfo, "LogonUserʧ��", "����", GetLastDllErr(Err.LastDllError)
    End If
    Call gobjLog.PopMethod(RLL_AllLog, "ZLHelperMain.clsRunas.RunAsUserNT4", RunAsUserNT4)
    Exit Function
ErrH:
    If gobjLog.ErrCenter(RLL_RunError, "ZLHelperMain.clsRunas.RunAsUserNT4") = 1 Then
        Resume
    End If
End Function

