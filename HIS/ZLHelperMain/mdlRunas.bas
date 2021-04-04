Attribute VB_Name = "mdlRunas"
Option Explicit
'==================================================================================================
'编写           lshuo
'日期           2019/4/18
'模块           mdlRunas
'说明           进程运行模块，可以以普通权限运行，也可以以管理员权限运行。
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
'功能：
'    指定进程在创建时的窗口站。桌面。标准句柄和主窗口的外观。
'定义：
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
'成员：
'    cb
'       结构的大小，以字节为单位。
'    lpReserved
'       保留;必须为空。
'    lpDesktop
'       桌面的名称，或此进程的桌面和窗口站的名称。字符串中的反斜杠表示该字符串同时包含桌面和窗口站名称。有关更多信息，请参见到桌面的线程连接。
'    lpTitle
'       对于控制台进程，如果创建了新的控制台窗口，则在标题栏中显示这个标题。如果为空，则使用可执行文件的名称作为窗口标题。对于不创建新控制台窗口的GUI或控制台进程，此参数必须为NULL。
'    dwX
'       如果dwFlags指定STARTF_USEPOSITION，则该成员是创建新窗口时窗口左上角的x偏移量(以像素为单位)。否则，此成员将被忽略。
'       偏移量来自屏幕的左上角。对于GUI进程，如果CreateWindow的x参数是CW_USEDEFAULT，则新进程第一次调用CreateWindow创建重叠窗口时将使用指定的位置。
'    dwY
'       如果dwFlags指定STARTF_USEPOSITION，则该成员是创建新窗口时窗口左上角的y偏移量，单位为像素。否则，此成员将被忽略。
'       偏移量来自屏幕的左上角。对于GUI进程，如果CreateWindow的y参数是CW_USEDEFAULT，则新进程第一次调用CreateWindow时将使用指定的位置来创建重叠的窗口。
'    dwXSize
'       如果dwFlags指定STARTF_USESIZE，则此成员是创建新窗口时窗口的宽度(以像素为单位)。否则，此成员将被忽略。
'       对于GUI进程，如果CreateWindow的nWidth参数是CW_USEDEFAULT，则新进程仅在第一次调用CreateWindow创建重叠窗口时才使用此方法。
'    dwYSize
'       如果dwFlags指定STARTF_USESIZE，则此成员是创建新窗口时窗口的高度(以像素为单位)。否则，此成员将被忽略。
'       对于GUI进程，如果CreateWindow的nHeight参数是CW_USEDEFAULT，则新进程仅在第一次调用CreateWindow创建重叠窗口时才使用此方法。
'    dwXCountChars
'       如果dwFlags指定STARTF_USECOUNTCHARS，如果在控制台进程中创建了一个新的控制台窗口，则该成员指定屏幕缓冲区宽度，单位为字符列。否则，此成员将被忽略。
'    dwYCountChars
'       如果dwFlags指定STARTF_USECOUNTCHARS，如果在控制台进程中创建一个新的控制台窗口，这个成员指定屏幕缓冲区的高度，以字符行为单位。否则，此成员将被忽略。
'    dwFillAttribute
'       如果dwFlags指定STARTF_USEFILLATTRIBUTE，则如果在控制台应用程序中创建了一个新的控制台窗口，则该成员是初始文本和背景颜色。否则，此成员将被忽略。
'       这个值可以是以下值的任意组合:FOREGROUND_BLUE。FOREGROUND_GREEN。FOREGROUND_RED。FOREGROUND_INTENSITY。BACKGROUND_BLUE。BACKGROUND_GREEN。BACKGROUND_RED和BACKGROUND_INTENSITY。例如，下面的值组合在白色背景上生成红色文本:
'           FOREGROUND_RED| BACKGROUND_RED| BACKGROUND_GREEN| BACKGROUND_BLUE
'    dwFlags
'       确定进程创建窗口时是否使用某些STARTUPINFO成员的位字段。此成员可以是以下值中的一个或多个
Private Const STARTF_FORCEONFEEDBACK            As Long = &H40
'    指示调用CreateProcess后光标处于反馈模式两秒。将显示正在工作的背景光标(请参阅鼠标控制面板实用程序中的Pointers选项卡)。
'    如果在这两秒钟内进程发出第一个GUI调用，那么系统会给进程5秒钟的时间。如果在这五秒钟内进程显示了一个窗口，系统会给进程五秒钟的时间来完成窗口的绘制。
'    系统在第一次调用GetMessage之后关闭反馈光标，不管进程是否正在绘制。
Private Const STARTF_FORCEOFFFEEDBACK           As Long = &H80
'    指示在进程启动时强制关闭反馈游标。将显示正常的选择光标。
Private Const STARTF_PREVENTPINNING             As Long = &H2000
'    指示进程创建的任何窗口不能固定在任务栏上。
'    这个标志必须与STARTF_TITLEISAPPID相结合
Private Const STARTF_RUNFULLSCREEN              As Long = &H20
'    指示进程应在全屏模式下运行，而不是在窗口模式下运行。
'    此标志仅适用于运行在x86计算机上的控制台应用程序。
Private Const STARTF_TITLEISAPPID               As Long = &H1000
'    lpTitle成员包含一个AppUserModelID。此标识符控制任务栏和开始菜单如何显示应用程序，并使其与正确的快捷方式和跳转列表相关联。通常，应用程序将使用SetCurrentProcessExplicitAppUserModelID和GetCurrentProcessExplicitAppUserModelID函数来代替设置此标志。有关更多信息，请参见应用程序用户模型id。
'    如果使用startf_preventpins，则无法将应用程序窗口固定在任务栏上。应用程序使用任何与appusermodelid相关的窗口属性只会覆盖该窗口的此设置。
'    此标志不能与STARTF_TITLEISLINKNAME一起使用
Private Const STARTF_TITLEISLINKNAME            As Long = &H800
'    lpTitle成员包含用户为启动此进程而调用的快捷方式文件(.lnk)的路径。这通常由shell在调用指向已启动应用程序的.lnk文件时设置。大多数应用程序不需要设置这个值。
'    此标志不能与STARTF_TITLEISAPPID一起使用
Private Const STARTF_UNTRUSTEDSOURCE            As Long = &H8000
'    命令行来自一个不可信的源。有关更多信息，请参见备注。
Private Const STARTF_USECOUNTCHARS              As Long = &H8
'    dwXCountChars和dwYCountChars成员包含附加信息。
Private Const STARTF_USEFILLATTRIBUTE           As Long = &H10
'    dwFillAttribute成员包含附加信息
Private Const STARTF_USEHOTKEY                  As Long = &H200
'    hStdInput成员包含其他信息
'    此标志不能与startf_usestdhandle一起使用
Private Const STARTF_USEPOSITION                As Long = &H4
'    dwX和dwY成员包含其他信息
Private Const STARTF_USESHOWWINDOW              As Long = &H1
'    wShowWindow成员包含其他信息
Private Const STARTF_USESIZE                    As Long = &H2
'    dwXSize和dwYSize成员包含附加信息
Private Const STARTF_USESTDHANDLES              As Long = &H100
'    hStdInput。hStdOutput和hStdError成员包含其他信息
'    如果在调用进程创建函数时指定了此标志，则句柄必须是可继承的，函数的bInheritHandles参数必须设置为TRUE。有关更多信息，请参见句柄继承。
'    如果在调用GetStartupInfo函数时指定了这个标志，那么这些成员要么是进程创建期间指定的句柄值，要么是INVALID_HANDLE_VALUE。
'    当不再需要手柄时，必须用close手柄关闭手柄。
'    此标志不能与STARTF_USEHOTKE一起使用
'
'    wShowWindow
'        如果dwFlags指定STARTF_USESHOWWINDOW，这个成员可以是ShowWindow函数的nCmdShow参数中指定的任何值，SW_SHOWDEFAULT除外。否则，此成员将被忽略。
'        对于GUI进程，第一次调用ShowWindow时，它的nCmdShow参数被忽略，wShowWindow指定了默认值。在对ShowWindow的后续调用中，如果将ShowWindow的nCmdShow参数设置为SW_SHOWDEFAULT，则使用wShowWindow成员。
'    cbReserved2
'        预留给C运行时使用;必须是零。
'    lpReserved2
'        预留给C运行时使用;必须为空。
'    hStdInput
'        如果dwFlags指定startf_usestdhandle，则此成员是流程的标准输入句柄。如果没有指定startf_usestdhandle，标准输入的默认值是键盘缓冲区。
'        如果dwFlags指定STARTF_USEHOTKEY，则此成员指定一个热键值，该热键值作为WM_SETHOTKEY消息的wParam参数发送到拥有该流程的应用程序创建的第一个符合条件的顶级窗口。如果窗口是用WS_POPUP窗口样式创建的，则除非还设置了WS_EX_APPWINDOW扩展窗口样式，否则不符合条件。有关详细信息，请参见CreateWindowEx。
'        否则，此成员将被忽略。
'    hStdOutput
'        如果dwFlags指定startf_usestdhandle，则此成员是流程的标准输出句柄。否则，该成员将被忽略，标准输出的默认值是控制台窗口的缓冲区。
'        如果从任务栏或跳转列表启动进程，系统将hStdOutput设置为监视器的句柄，该句柄包含用于启动进程的任务栏或跳转列表。有关更多信息，请参见备注。
'        Windows 7。Windows Server 2008 R2。Windows Vista。Windows Server 2008。Windows XP和Windows Server 2003:这一行为在Windows 8和Windows Server 2012中引入。
'    hStdError
'        如果dwFlags指定startf_usestdhandle，则此成员是流程的标准错误句柄。否则，该成员将被忽略，标准错误的默认值是控制台窗口的缓冲区。
'备注：
'    对于图形用户界面(GUI)进程，此信息影响由CreateWindow函数创建并由ShowWindow函数显示的第一个窗口。对于控制台进程，如果为该进程创建了新控制台，则此信息将影响控制台窗口。进程可以使用GetStartupInfo函数来检索在创建进程时指定的STARTUPINFO结构。
'    如果正在启动GUI进程，并且没有指定STARTF_FORCEONFEEDBACK或STARTF_FORCEOFFFEEDBACK，则使用process feedback游标。GUI进程的子系统指定为“windows”。
'    如果从任务栏或跳转列表启动进程，系统将hStdOutput设置为监视器的句柄，该句柄包含用于启动进程的任务栏或跳转列表。要检索这个句柄，使用GetStartupInfo来检索STARTUPINFO结构，并检查hStdOutput是否已设置。然后，进程可以使用句柄来定位它的窗口。
'    如果STARTF_UNTRUSTEDSOURCE标志是在GetStartupInfo函数返回的STARTUPINFO结构中设置的，那么应用程序应该知道命令行是不受信任的。如果设置了此标志，应用程序应该禁用潜在的危险特性，如宏。下载内容和自动打印。这个标志是可选的。当使用不受信任的命令行启动程序时，鼓励调用CreateProcess的应用程序设置此标志，以便创建的进程可以应用适当的策略。
'    STARTF_UNTRUSTEDSOURCE标志在Windows Vista中支持启动，但在Windows 10 SDK之前没有在SDK头文件中定义。要在Windows 10之前的版本中使用该标志，可以在程序中手动定义它。
'支持
'    最低支持客户
'       Windows XP [只适用于桌面应用程式]
'    最低支持服务器
'       Windows Server 2003[只适用于桌面应用程序]
'    Header
'       WinBase.h on Windows XP, Windows Server 2003, Windows Vista, Windows 7, Windows Server 2008和Windows Server 2008 R2(包括Windows.h);
'       Windows 8和Windows Server 2012上的Processthreadsapi.h
'    Unicode和ANSI名称
'       STARTUPINFOW (Unicode)和STARTUPINFOA (ANSI)
Private Type PROCESS_INFORMATION
    hProcess            As Long
    hThread             As Long
    dwProcessId         As Long
    dwThreadId          As Long
End Type
'功能：
'   包含关于新创建的进程及其主线程的信息。它与CreateProcess。CreateProcessAsUser。CreateProcessWithLogonW或CreateProcessWithTokenW函数一起使用。
'定义
'    typedef struct _PROCESS_INFORMATION {
'      HANDLE hProcess;
'      HANDLE hThread;
'      DWORD  dwProcessId;
'      DWORD  dwThreadId;
'    } PROCESS_INFORMATION, *LPPROCESS_INFORMATION;
'成员
'    hProcess
'       新创建进程的句柄。句柄用于在对进程对象执行操作的所有函数中指定进程
'    hThread
'       新创建进程的主线程的句柄。句柄用于在对thread对象执行操作的所有函数中指定线程。
'    dwProcessId
'       可用于标识进程的值。该值从创建进程时起有效，直到关闭到进程的所有句柄并释放进程对象为止;此时，可以重用标识符。
'    dwThreadId
'       可用于标识线程的值。该值从线程创建时起有效，直到线程的所有句柄都关闭并释放线程对象为止;此时，可以重用标识符。
'备注
'    如果函数成功，请确保调用close句柄函数来关闭hProcess和hThread句柄。否则，当子进程退出时，系统无法清理子进程的进程结构，因为父进程仍然拥有子进程的打开句柄。但是，当父进程终止时，系统将关闭这些句柄，因此与子进程对象相关的结构将在此时被清除。
'支持
'    最低支持客户
'       Windows XP [只适用于桌面应用程式]
'    最低支持服务器
'       Windows Server 2003[只适用于桌面应用程序]
'    Header
'       WinBase.h on Windows XP, Windows Server 2003, Windows Vista, Windows 7, Windows Server 2008 and Windows Server 2008 R2 (include Windows.h);
'       Processthreadsapi.h on Windows 8 and Windows Server 2012
Private Declare Function CreateProcessWithLogon Lib "advapi32" Alias "CreateProcessWithLogonW" (ByVal lpUsername As Long, ByVal lpDomain As Long, ByVal lpPassword As Long, ByVal dwLogonFlags As Long, ByVal lpApplicationName As Long, ByVal lpCommandLine As Long, ByVal dwCreationFlags As Long, ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As Long, lpStartupInfo As STARTUPINFO, lpProcessInfo As PROCESS_INFORMATION) As Long
'定义：
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
'功能：
'    创建一个新进程及其主线程。然后，新进程在指定凭据(用户。域和密码)的安全上下文中运行指定的可执行文件。它可以选择加载指定用户的用户配置文件。
'    此函数类似于CreateProcessAsUser和CreateProcessWithTokenW函数，只是调用者不需要调用LogonUser函数来验证用户身份并获取令牌。
'参数：
'    lpUsername _In_
'       用户名。这是要登录的用户帐户的名称。如果使用UPN格式user@ DNS_domain_name, lpDomain参数必须为NULL
'       用户帐户必须具有本地计算机上的本地登录权限。此权限授予工作站和服务器上的所有用户，但仅授予域控制器上的管理员。
'    lpDomain _In_opt_
'       帐户数据库包含lpUsername帐户的域或服务器的名称。如果该参数为空，则必须以UPN格式指定用户名。
'    lpPassword _In_
'       lpUsername帐户的明文密码
'    dwLogonFlags _In_
'       登录选项。这个参数可以是0(0)，也可以是下列值之一。
Private Const LOGON_WITH_PROFILE                As Long = &H1
'    登录，然后在HKEY_USERS注册表项中加载用户概要文件。函数在加载概要文件后返回。加载概要文件可能很耗时，所以最好只在必须访问HKEY_CURRENT_USER注册表项中的信息时才使用这个值。
'    Windows Server 2003:在新进程终止后卸载概要文件，不管它是否创建了子进程。
'    Windows XP: 在新进程及其创建的所有子进程终止后卸载概要文件
Private Const LOGON_NETCREDENTIALS_ONLY         As Long = &H2
'    登录，但仅在网络上使用指定的凭据。新进程使用与调用者相同的令牌，但是系统在LSA中创建一个新的登录会话，并且该进程使用指定的凭据作为缺省凭据。
'    此值可用于创建一个进程，该进程在本地使用的凭据集与远程使用的凭据集不同。这在没有信任关系的域间场景中非常有用。
'    系统不验证指定的凭据。因此，进程可以启动，但它可能无法访问网络资源。
'    lpApplicationName _In_opt_
'        要执行的模块的名称。这个模块可以是一个基于windows的应用程序。如果本地计算机上有合适的子系统，它可以是其他类型的模块(例如MS-DOS或OS/2)。
'        字符串可以指定要执行的模块的完整路径和文件名，也可以指定部分名称。如果是部分名称，则函数使用当前驱动器和当前目录来完成规范。该函数不使用搜索路径。此参数必须包含文件名扩展名;没有默认扩展。
'        lpApplicationName参数可以为NULL，模块名称必须是lpCommandLine字符串中第一个以空格分隔的白令牌。如果使用包含空格的长文件名，请使用带引号的字符串指示文件名的结束和参数的开始位置;否则，文件名称是模糊的。
'        例如，下面的字符串可以用不同的方式解释:
'        "c:\program files\sub dir\program name"
'        系统试图按照以下顺序解释这些可能性:
'        c:\program.exe files\sub dir\program name
'        c:\program files\sub.exe dir\program name
'        c:\program files\sub dir\program.exe name
'        c:\program files\sub dir\program name.exe
'        如果可执行模块是16位的应用程序，lpApplicationName应该为NULL, lpCommandLine指向的字符串应该指定可执行模块及其参数。
'    lpCommandLine _Inout_opt_
'        要执行的命令行。这个字符串的最大长度是1024个字符。如果lpApplicationName为NULL，则lpCommandLine的模块名部分限制为MAX_PATH字符。
'        函数可以修改这个字符串的内容。因此，该参数不能是指向只读内存的指针(例如const变量或文字字符串)。如果该参数是常量字符串，则该函数可能会导致访问冲突。
'        lpCommandLine参数可以为NULL，函数使用lpApplicationName指向的字符串作为命令行。
'        如果lpApplicationName和lpCommandLine都是非空的，则*lpApplicationName指定要执行的模块，*lpCommandLine指定命令行。新进程可以使用GetCommandLine检索整个命令行。用C编写的控制台进程可以使用argc和argv参数解析命令行。因为argv[0]是模块名，C程序员通常将模块名作为命令行中的第一个令牌重复。
'        如果lpApplicationName为NULL，则命令行中第一个以空格分隔的标记指定模块名称。如果使用包含空格的长文件名，请使用带引号的字符串来指示文件名的结束和参数的开始位置(请参阅lpApplicationName参数的说明)。如果文件名不包含扩展名，则追加.exe。因此，如果文件名扩展名是。com，这个参数必须包含。com扩展名。如果文件名以没有扩展名的句点结束，或者文件名包含路径，则不附加.exe。如果文件名不包含目录路径，系统按以下顺序搜索可执行文件:
'            加载应用程序的目录
'            父进程的当前目录
'            32 位Windows系统目录。使用GetSystemDirectory函数获取此目录的路径
'            16位Windows系统目录。没有函数可以获得这个目录的路径，但是可以搜索它。
'            Windows目录。使用GetWindowsDirectory函数获取此目录的路径。
'            PATH环境变量中列出的目录。注意，此函数不搜索由App Paths registry键指定的每个应用程序路径。要在搜索序列中包含每个应用程序的路径，请使用ShellExecute函数
'        系统将一个空字符添加到命令行字符串中，以将文件名与参数分开。这将原始字符串分成两个字符串进行内部处理。
'    dwCreationFlags _In_
'           控制如何创建进程的标志。默认情况下，CREATE_DEFAULT_ERROR_MODE。CREATE_NEW_CONSOLE和CREATE_NEW_PROCESS_GROUP标志是启用的——即使您没有设置该标志，系统的功能也与设置一样。
Private Const CREATE_DEFAULT_ERROR_MODE         As Long = &H4000000
'    新进程不继承调用进程的错误模式。相反，CreateProcessWithLogonW为新进程提供当前默认错误模式。应用程序通过调用SetErrorMode设置当前默认错误模式。
'    默认情况下启用此标志
Private Const CREATE_NEW_CONSOLE                As Long = &H10
'    新进程有一个新控制台，而不是继承父进程的控制台。此标志不能与DETACHED_PROCESS标志一起使用。
'    默认情况下启用此标志
Private Const CREATE_NEW_PROCESS_GROUP          As Long = &H200
'    新进程是新进程组的根进程。进程组包含此根进程的所有后代进程。新进程组的进程标识符与在lpProcessInfo参数中返回的进程标识符相同。GenerateConsoleCtrlEvent函数使用进程组来向一组控制台进程发送CTRL C或CTRL + BREAK信号。
'    默认情况下启用此标志
Private Const CREATE_SEPARATE_WOW_VDM           As Long = &H800
'    此标志仅在启动基于16位windows的应用程序时有效。如果设置好，新进程将在私有虚拟DOS机器(VDM)中运行。默认情况下，所有基于windows的16位应用程序都运行在一个共享的VDM中。单独运行的好处是崩溃只会终止单个VDM;在不同VDMs中运行的任何其他程序都可以正常运行。此外，运行在单独VDMs中的16位基于windows的应用程序具有单独的输入队列，这意味着如果一个应用程序暂时停止响应，单独VDMs中的应用程序将继续接收输入。
Private Const CREATE_SUSPENDED                  As Long = &H4
'    新进程的主线程是在挂起状态下创建的，在调用ResumeThread函数之前不会运行。
Private Const CREATE_UNICODE_ENVIRONMENT        As Long = &H400
'    指示lpEnvironment参数的格式。如果设置了此标志，lpEnvironment指向的环境块将使用Unicode字符。否则，环境块使用ANSI字符。
'        此参数还控制新进程的优先级类，该类用于确定进程线程的调度优先级。有关值列表，请参见GetPriorityClass。如果没有指定任何优先级类标志，则优先级类默认为NORMAL_PRIORITY_CLASS，除非创建过程的优先级类是IDLE_PRIORITY_CLASS或低于idle_normal_priority_class。在本例中，子进程接收调用进程的默认优先级类。
'    lpEnvironment _In_opt_
'        指向新进程的环境块的指针。如果该参数为NULL，则新流程将使用由lpUsername指定的用户概要文件创建的环境。
'        环境块由以null结尾的字符串组成的以null结尾的块。每个字符串的形式如下:
'        名称 = 值
'        因为等号(=)用作分隔符，所以不能在环境变量的名称中使用它。
'        环境块可以包含Unicode或ANSI字符。如果lpEnvironment指向的环境块包含Unicode字符，请确保dwCreationFlags包含CREATE_UNICODE_ENVIRONMENT。如果该参数为NULL，并且父进程的环境块包含Unicode字符，则还必须确保dwCreationFlags包含CREATE_UNICODE_ENVIRONMENT。
'        ANSI环境块由两个0(0)字节终止:一个用于最后一个字符串，另一个用于终止该块。Unicode环境块终止为4个零字节:最后一个字符串终止为2个字节，最后一个字符串终止为2个字节。
'        要为特定用户检索环境块的副本，请使用CreateEnvironmentBlock函数。
'    lpCurrentDirectory  _In_opt_
'        进程当前目录的完整路径。字符串还可以指定UNC路径。
'        如果该参数为NULL，则新进程具有与调用进程相同的当前驱动器和目录。该特性主要是为需要启动应用程序并指定其初始驱动器和工作目录的shell提供的。
'    lpStartupInfo _In_
'        指向STARTUPINFO结构的指针。应用程序必须将指定用户帐户的权限添加到指定的窗口站和桌面，甚至是WinSta0\Default。
'        如果lpDesktop成员为NULL或空字符串，则新进程继承其父进程的桌面和窗口站。应用程序必须将指定用户帐户的权限添加到继承的窗口站和桌面。
'        CreateProcessWithLogonW将指定用户帐户的权限添加到继承的窗口站和桌面。
'        当不再需要句柄时，必须用close句柄关闭STARTUPINFO中的句柄。
'        重要的是，如果STARTUPINFO结构的dwFlags成员指定startf_usestdhandle，则标准句柄字段将不加更改地复制到子进程，而不进行验证。调用者负责确保这些字段包含有效的句柄值。不正确的值可能导致子进程行为不当或崩溃。使用应用程序验证程序运行时验证工具检测无效句柄。
'    lpProcessInfo _Out_
'        指向PROCESS_INFORMATION结构的指针，该结构接收新进程的标识信息，包括进程的句柄。
'        PROCESS_INFORMATION中的句柄在不需要时必须使用close句柄函数关闭。
'返回值
'    如果函数成功，返回值为非零。
'    如果函数失败，返回值为0(0)。要获取扩展的错误信息，请调用GetLastError。
'    注意，函数在进程完成初始化之前返回。如果无法找到所需的DLL或初始化失败，进程将终止。要获取进程的终止状态，请调用GetExitCodeProcess。
'备注
'    默认情况下，CreateProcessWithLogonW不会将指定的用户配置文件加载到HKEY_USERS注册表项中。这意味着对HKEY_CURRENT_USER注册表项中的信息的访问可能不会产生与正常交互登录一致的结果。您有责任在调用CreateProcessWithLogonW之前，通过使用LOGON_WITH_PROFILE或通过调用LoadUserProfile函数，将用户注册表hive加载到HKEY_USERS中。
'    如果lpEnvironment参数为NULL，则新流程将使用由lpUserName指定的用户概要文件创建的环境块。如果没有设置HOMEDRIVE和HOMEPATH变量，CreateProcessWithLogonW将修改环境块，以使用用户工作目录的驱动器和路径。
'    创建时，新进程和线程句柄将接收完整的访问权限(PROCESS_ALL_ACCESS和THREAD_ALL_ACCESS)。对于任何一个句柄，如果没有提供安全描述符，则可以在需要该类型对象句柄的任何函数中使用该句柄。当提供安全描述符时，将在授予访问权之前对句柄的所有后续使用执行访问检查。如果访问被拒绝，请求进程不能使用句柄访问进程或线程。
'    要检索安全令牌，请将PROCESS_INFORMATION结构中的流程句柄传递给OpenProcessToken函数。
'    进程被分配一个进程标识符。标识符在进程终止之前有效。它可以用来标识进程，也可以在OpenProcess函数中指定它来打开进程的句柄。进程中的初始线程也被分配一个线程标识符。可以在OpenThread函数中指定它来打开线程的句柄。标识符在线程终止之前是有效的，并且可以用于惟一地标识系统中的线程。这些标识符在PROCESS_INFORMATION中返回。
'    调用线程可以使用WaitForInputIdle函数等待，直到新进程完成初始化，并且正在等待用户输入而没有输入挂起。这对于父进程和子进程之间的同步非常有用，因为CreateProcessWithLogonW在不等待新进程完成初始化的情况下返回。例如，创建流程将在尝试查找与新流程关联的窗口之前使用WaitForInputIdle。
'    关闭进程的首选方法是使用ExitProcess函数，因为该函数向附加到进程的所有dll发送终止通知。其他关闭进程的方法不通知附加的dll。注意，当一个线程调用ExitProcess时，进程的其他线程将被终止，而没有机会执行任何其他代码(包括附加dll的线程终止代码)。有关更多信息，请参见终止进程。
'安全说明
'    lpApplicationName参数可以为NULL，并且可执行名称必须是lpCommandLine中第一个空格分隔的字符串。如果可执行文件或路径名称中有空格，则由于函数解析空格的方式，可能会运行不同的可执行文件。避免以下示例，因为函数试图运行“Program”。，如果它存在，而不是“MyApp.exe”。
'    LPTSTR szCmdline[]=_tcsdup(TEXT("C:\\Program Files\\MyApp"));
'    CreateProcessWithLogonW (…,szCmdline……)
'    如果恶意用户创建了一个名为“Program”的应用程序。在系统上，任何使用程序文件目录错误调用CreateProcessWithLogonW的程序都会运行恶意用户应用程序，而不是预期的应用程序。
'    为了避免这个问题，不要为lpApplicationName传递NULL。如果为lpApplicationName传递NULL，请在lpCommandLine中的可执行路径周围使用引号，如下面的示例所示:
'    LPTSTR szCmdline[]=_tcsdup(TEXT("\"C:\\Program Files\\MyApp\""));
'    CreateProcessWithLogonW(..., szCmdline, ...)
'Requirements
'    最低支持客户端
'       Windows XP [只有桌面版]
'    最低支持服务器
'       Windows Server 2003 [只有桌面版]
'    Header
'       WinBase.h (include Windows.h)
'    Library
'       advapi32.lib
'    dll
'       advapi32.dll
Private Declare Function LogonUser Lib "advapi32.dll" Alias "LogonUserA" (ByVal lpszUsername As String, ByVal lpszDomain As String, _
                        ByVal lpszPassword As String, ByVal dwLogonType As Long, ByVal dwLogonProvider As Long, phToken As Long) As Long
'功能：
'   LogonUser函数尝试将用户登录到本地计算机。本地计算机是调用LogonUser的计算机。您不能使用LogonUser登录到远程计算机。您可以使用用户名和域指定用户，并使用明文密码对用户进行身份验证。如果函数成功，则接收表示登录用户的令牌句柄。然后，您可以使用此令牌句柄模拟指定的用户，或者在大多数情况下，创建在指定用户的上下文中运行的流程。
'定义：
'    BOOL LogonUser(
'      _In_     LPTSTR  lpszUsername,
'      _In_opt_ LPTSTR  lpszDomain,
'      _In_opt_ LPTSTR  lpszPassword,
'      _In_     DWORD   dwLogonType,
'      _In_     DWORD   dwLogonProvider,
'      _Out_    PHANDLE phToken
'    );
'成员
'    lpszUsername _In_
'       指向以null结尾的字符串的指针，该字符串指定用户名。这是要登录的用户帐户的名称。如果使用用户主体名称(UPN)格式User@DNSDomainName, lpszDomain参数必须为NULL。
'    lpszDomain _In_opt_
'       指向以null结尾的字符串的指针，该字符串指定帐户数据库包含lpszUsername帐户的域或服务器的名称。如果该参数为空，则必须以UPN格式指定用户名。如果这个参数是“。，该函数只使用本地帐户数据库验证帐户。
'    lpszPassword _In_opt_
'       指向以null结尾的字符串的指针，该字符串指定lpszUsername指定的用户帐户的明文密码。当您使用完密码后，通过调用SecureZeroMemory函数从内存中清除密码。有关保护密码的更多信息，请参见处理密码。
'    dwLogonType _In_
'       要执行的登录操作的类型。该参数可以是Winbase.h中定义的以下值之一
Private Const LOGON32_LOGON_BATCH               As Long = &H4
'    此登录类型适用于批处理服务器，在批处理服务器上，进程可以代表用户执行，而无需用户的直接干预。这种类型也适用于性能更高的服务器，这些服务器一次处理许多纯文本身份验证尝试，比如邮件或web服务器。
Private Const LOGON32_LOGON_INTERACTIVE         As Long = &H2
'    此登录类型适用于交互使用计算机的用户，例如由终端服务器。远程shell或类似进程登录的用户。这种登录类型还有额外的开销，即为断开连接的操作缓存登录信息;因此，它不适用于某些客户机/服务器应用程序，比如邮件服务器。
Private Const LOGON32_LOGON_NETWORK             As Long = &H3
'    此登录类型用于高性能服务器验证明文密码。LogonUser函数不缓存此登录类型的凭据
Private Const LOGON32_LOGON_NETWORK_CLEARTEXT   As Long = &H8
'    此登录类型在身份验证包中保留名称和密码，这允许服务器在模拟客户机时连接到其他网络服务器。服务器可以接受来自客户机的纯文本凭证，调用LogonUser，验证用户可以通过网络访问系统，并且仍然可以与其他服务器通信。
Private Const LOGON32_LOGON_NEW_CREDENTIALS     As Long = &H9
'    此登录类型允许调用方克隆其当前令牌并为出站连接指定新的凭据。新的登录会话具有相同的本地标识符，但对其他网络连接使用不同的凭据。
'    此登录类型仅受LOGON32_PROVIDER_WINNT50登录提供程序支持
Private Const LOGON32_LOGON_SERVICE             As Long = &H5
'    指示服务类型登录。所提供的帐户必须启用服务特权。
Private Const LOGON32_LOGON_UNLOCK              As Long = &H7
'    不再支持GINAs
'    Windows Server 2003和Windows XP:这种登录类型是为GINA dll提供的，用于登录将交互式使用计算机的用户。此登录类型可以生成一个惟一的审计记录，显示工作站何时解锁。
'
'    dwLogonProvider _In_
'       指定登录提供程序。此参数可以是以下值之一
Private Const LOGON32_PROVIDER_DEFAULT          As Long = &H0
'    使用系统的标准登录提供程序。默认的安全提供程序是协商的，除非您为域名传递NULL，而用户名不是UPN格式。在本例中，默认提供程序是NTLM。
Private Const LOGON32_PROVIDER_WINNT50          As Long = &H3
'    使用协商登录提供程序
Private Const LOGON32_PROVIDER_WINNT40          As Long = &H2
'    使用NTLM登录提供程序
'
'    phToken _Out_
'        指向句柄变量的指针，该变量接收表示指定用户的令牌的句柄。
'        您可以在调用ImpersonateLoggedOnUser函数时使用返回的句柄。
'        在大多数情况下，返回的句柄是一个主要令牌，您可以在调用CreateProcessAsUser函数时使用它。但是，如果指定LOGON32_LOGON_NETWORK标志，LogonUser将返回一个模拟令牌，除非调用DuplicateTokenEx将其转换为一个主令牌，否则您不能在CreateProcessAsUser中使用这个令牌。
'        当您不再需要这个句柄时，通过调用close句柄函数来关闭它。
'返回值
'    如果函数成功，则函数返回非零。
'    如果函数失败，它返回零。要获取扩展的错误信息，请调用GetLastError。
'备注
'    LOGON32_LOGON_NETWORK登录类型是最快的，但是它有以下限制:
'    函数返回模拟令牌，而不是主令牌。您不能在CreateProcessAsUser函数中直接使用这个令牌。但是，您可以调用DuplicateTokenEx函数将令牌转换为主令牌，然后在CreateProcessAsUser中使用它。
'    如果将令牌转换为主令牌并在CreateProcessAsUser中使用它来启动进程，则新进程不能通过重定向访问其他网络资源，例如远程服务器或打印机。一个例外是，如果网络资源不受访问控制，那么新进程将能够访问它。
'    这个函数不需要SE_TCB_NAME特权，除非您正在登录Passport帐户。
'    由lpszUsername指定的帐户必须具有必要的帐户权限。例如，要使用LOGON32_LOGON_INTERACTIVE标志登录用户，用户(或用户所属的组)必须拥有SE_INTERACTIVE_LOGON_NAME帐户。有关影响各种登录操作的帐户权限列表，请参见帐户权限常量。
'    如果存在至少一个令牌，则认为用户已登录。如果您调用CreateProcessAsUser并关闭令牌，系统将认为该用户仍然登录，直到进程(以及所有子进程)结束。
'    如果LogonUser调用成功，系统将通过调用提供者的NPLogonNotify入口点函数通知网络提供者发生了登录。
'支持
'    最低支持客户
'       Windows XP [只适用于桌面应用程式]
'    最低支持服务器
'       Windows Server 2003[只适用于桌面应用程序]
'    Header
'       Winbase.h (包括Windows.h)
'    Library
'       advapi32.lib
'    dll
'       advapi32.dll
'    Unicode和ANSI名称
'       LogonUserW (Unicode)和LogonUserA(ANSI)
Private Declare Function CreateProcessAsUser Lib "advapi32.dll" Alias "CreateProcessAsUserA" (ByVal hToken As Long, ByVal lpApplicationName As String, _
                    ByVal lpCommandLine As String, ByVal lpProcessAttributes As Long, ByVal lpThreadAttributes As Long, ByVal bInheritHandles As Long, _
                    ByVal dwCreationFlags As Long, ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As String, lpStartupInfo As STARTUPINFO, _
                    lpProcessInformation As PROCESS_INFORMATION) As Long
'功能：
'    创建一个新进程及其主线程。新进程在指定令牌表示的用户的安全上下文中运行
'    通常，调用CreateProcessAsUser函数的进程必须具有se_incree_quota_name特权，如果令牌不可分配，则可能需要SE_ASSIGNPRIMARYTOKEN_NAME特权。如果该函数在ERROR_PRIVILEGE_NOT_HELD(1314)中失败，则使用CreateProcessWithLogonW函数。CreateProcessWithLogonW不需要特权，但是必须允许指定的用户帐户交互式地登录。通常，最好使用CreateProcessWithLogonW创建具有备用凭证的进程。
'定义
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
'成员：
'    hToken _In_opt_
'        表示用户的主令牌的句柄。句柄必须具有TOKEN_QUERY。TOKEN_DUPLICATE和TOKEN_ASSIGN_PRIMARY访问权限。有关更多信息，请参见访问令牌对象的访问权限。令牌表示的用户必须具有对lpApplicationName或lpCommandLine参数指定的应用程序的读取和执行访问权。
'        要获得表示指定用户的主令牌，请调用LogonUser函数。或者，您可以调用DuplicateTokenEx函数将模拟令牌转换为主令牌。这允许模拟客户机的服务器应用程序创建具有客户机安全上下文的流程。
'        如果hToken是调用方的主令牌的受限版本，则不需要SE_ASSIGNPRIMARYTOKEN_NAME特权。如果还没有启用必要的特权，CreateProcessAsUser将在调用期间启用这些特权。有关更多信息，请参见使用特权运行。
'        终端服务:进程在令牌中指定的会话中运行。默认情况下，这是与LogonUser相同的会话。要更改会话，请使用SetTokenInformation函数。
'    lpApplicationName _In_opt_
'        要执行的模块的名称。这个模块可以是一个基于windows的应用程序。如果本地计算机上有合适的子系统，它可以是其他类型的模块(例如MS-DOS或OS/2)。
'        字符串可以指定要执行的模块的完整路径和文件名，也可以指定部分名称。对于部分名称，函数使用当前驱动器和当前目录来完成规范。该函数将不使用搜索路径。此参数必须包含文件名扩展名;没有默认扩展。
'        lpApplicationName参数可以为NULL。在这种情况下，模块名必须是lpCommandLine字符串中第一个以空格分隔的标记。如果使用包含空格的长文件名，请使用带引号的字符串指示文件名的结束和参数的开始位置;否则，文件名称是模糊的。例如，考虑字符串“c:\program files\sub dir\program name”。这个字符串可以用多种方式解释。系统试图按照以下顺序解释这些可能性:
'        c:\program.exe files\sub dir\program name
'        c:\program files\sub.exe dir\program name
'        c:\program files\sub dir\program.exe name
'        c:\program files\sub dir\program name.exe
'        如果可执行模块是16位的应用程序，lpApplicationName应该为NULL, lpCommandLine指向的字符串应该指定可执行模块及其参数。默认情况下，CreateProcessAsUser创建的所有16位基于windows的应用程序都运行在一个单独的VDM中(相当于CreateProcess中的CREATE_SEPARATE_WOW_VDM)。
'    lpCommandLine _Inout_opt_
'        要执行的命令行。这个字符串的最大长度是32K个字符。如果lpApplicationName为NULL，则lpCommandLine的模块名部分限制为MAX_PATH字符。
'        这个函数的Unicode版本CreateProcessAsUserW可以修改这个字符串的内容。因此，该参数不能是指向只读内存的指针(例如const变量或文字字符串)。如果该参数是常量字符串，则该函数可能会导致访问冲突。
'        lpCommandLine参数可以为空。在这种情况下，函数使用lpApplicationName指向的字符串作为命令行。
'        如果lpApplicationName和lpCommandLine都是非空的，则*lpApplicationName指定要执行的模块，*lpCommandLine指定命令行。新进程可以使用GetCommandLine检索整个命令行。用C编写的控制台进程可以使用argc和argv参数解析命令行。因为argv[0]是模块名，C程序员通常将模块名作为命令行中的第一个令牌重复。
'        如果lpApplicationName为NULL，则命令行中第一个以空格分隔的标记指定模块名称。如果使用包含空格的长文件名，请使用带引号的字符串来指示文件名的结束和参数的开始位置(请参阅lpApplicationName参数的说明)。如果文件名不包含扩展名，则追加.exe。因此，如果文件名扩展名是。com，这个参数必须包含。com扩展名。如果文件名以没有扩展名的句点(.)结尾，或者文件名包含路径，则不附加.exe。如果文件名不包含目录路径，系统按以下顺序搜索可执行文件:
'            加载应用程序的目录
'            父进程的当前目录
'            32 位Windows系统目录。使用GetSystemDirectory函数获取此目录的路径
'            16位Windows系统目录。没有函数可以获得这个目录的路径，但是可以搜索它。
'            Windows目录。使用GetWindowsDirectory函数获取此目录的路径
'            PATH环境变量中列出的目录。注意，此函数不搜索由App Paths registry键指定的每个应用程序路径。要在搜索序列中包含每个应用程序的路径，请使用ShellExecute函数。
'        系统将一个空字符添加到命令行字符串中，以将文件名与参数分开。这将原始字符串分成两个字符串进行内部处理。
'    lpProcessAttributes _In_opt_
'        指向SECURITY_ATTRIBUTES结构的指针，该结构为新进程对象指定安全描述符，并确定子进程是否可以继承返回给进程的句柄。如果lpProcessAttributes为NULL或lpSecurityDescriptor为NULL，则流程将获得默认的安全描述符，并且不能继承句柄。默认的安全描述符是在hToken参数中引用的用户的安全描述符。此安全描述符可能不允许调用方访问，在这种情况下，进程运行后可能不会再次打开。进程句柄是有效的，并且将继续拥有完全的访问权限。
'    lpThreadAttributes _In_opt_
'        指向SECURITY_ATTRIBUTES结构的指针，该结构为新线程对象指定安全描述符，并确定子进程是否可以将返回的句柄继承给线程。如果lpThreadAttributes为NULL或lpSecurityDescriptor为NULL，线程将获得一个默认的安全描述符，并且不能继承句柄。默认的安全描述符是在hToken参数中引用的用户的安全描述符。此安全描述符可能不允许调用方访问。
'    bInheritHandles  _In_
'        如果此参数为真，则调用流程中的每个可继承句柄将由新流程继承。如果参数为FALSE，则不会继承句柄。注意，继承的句柄具有与原始句柄相同的值和访问权限。
'        终端服务:不能跨会话继承句柄。此外，如果该参数为真，则必须在调用者所在的会话中创建流程。
'        受保护的进程轻(PPL)进程:当PPL进程创建非PPL进程时，由于PROCESS_DUP_HANDLE不允许从非PPL进程到PPL进程，所以一般句柄继承被阻塞。参见流程安全和访问权限
'    dwCreationFlags _In_
'        控制优先级类和进程创建的标志。有关值列表，请参见进程创建标志。
'        此参数还控制新进程的优先级类，该类用于确定进程线程的调度优先级。有关值列表，请参见GetPriorityClass。如果没有指定任何优先级类标志，则优先级类默认为NORMAL_PRIORITY_CLASS，除非创建过程的优先级类是IDLE_PRIORITY_CLASS或低于idle_normal_priority_class。在本例中，子进程接收调用进程的默认优先级类。
'    lpEnvironment _In_opt_
'        指向新进程的环境块的指针。如果该参数为NULL，则新进程将使用调用进程的环境。
'        环境块由以null结尾的字符串组成的以null结尾的块。每个字符串的形式如下:
'            Name = 值 \ 0
'        因为等号用作分隔符，所以不能在环境变量的名称中使用它。
'        环境块可以包含Unicode或ANSI字符。如果lpEnvironment指向的环境块包含Unicode字符，请确保dwCreationFlags包含CREATE_UNICODE_ENVIRONMENT。如果该参数为NULL，并且父进程的环境块包含Unicode字符，则还必须确保dwCreationFlags包含CREATE_UNICODE_ENVIRONMENT。
'        如果进程的环境块的总大小超过32,767个字符，则此函数的ANSI版本CreateProcessAsUserA将失败。
'        注意，ANSI环境块以两个零字节终止:一个用于最后一个字符串，另一个用于终止该块。Unicode环境块终止为4个零字节:最后一个字符串终止为2个字节，最后一个字符串终止为2个字节。
'        Windows Server 2003和Windows XP:如果合并的用户和系统环境变量的大小超过8192字节，CreateProcessAsUser创建的进程将不再运行父进程传递给函数的环境块。相反，子进程使用CreateEnvironmentBlock函数返回的环境块运行。
'        要为给定用户检索环境块的副本，请使用CreateEnvironmentBlock函数。
'    lpCurrentDirectory _In_opt_
'        进程当前目录的完整路径。字符串还可以指定UNC路径
'        如果该参数为NULL，则新进程将具有与调用进程相同的当前驱动器和目录。(该特性主要用于需要启动应用程序并指定其初始驱动器和工作目录的shell。)
'    lpStartupInfo _In_
'        指向STARTUPINFO或STARTUPINFOEX结构的指针。
'        用户必须完全访问指定的窗口站和桌面。如果您希望进程是交互式的，请指定winsta0\default。如果lpDesktop成员为空，则新进程继承其父进程的桌面和窗口站。如果该成员是空字符串“”，则新进程将使用“进程连接到窗口站”中描述的规则连接到窗口站。
'        要设置扩展属性，请使用STARTUPINFOEX结构，并在dwCreationFlags参数中指定EXTENDED_STARTUPINFO_PRESENT。
'        当不再需要STARTUPINFO或STARTUPINFOEX中的句柄时，必须用close句柄关闭它们。
'        重要的是，调用者负责确保STARTUPINFO中的标准句柄字段包含有效的句柄值。即使在dwFlags成员指定startf_usestdhandle时，这些字段也会不加验证地复制到子进程中。不正确的值可能导致子进程行为不当或崩溃。使用应用程序验证程序运行时验证工具检测无效句柄。
'    lpProcessInformation _Out_
'        指向PROCESS_INFORMATION结构的指针，该结构接收关于新进程的标识信息。
'        PROCESS_INFORMATION中的句柄必须在不再需要时用close句柄关闭
'返回值
'    如果函数成功，返回值为非零。
'    如果函数失败，返回值为零。要获取扩展的错误信息，请调用GetLastError。
'    注意，函数在进程完成初始化之前返回。如果无法找到所需的DLL或初始化失败，进程将终止。要获取进程的终止状态，请调用GetExitCodeProcess。
'备注：
'    CreateProcessAsUser必须能够使用TOKEN_DUPLICATE和TOKEN_IMPERSONATE访问权限打开调用进程的主令牌。
'    默认情况下，CreateProcessAsUser在非交互式窗口站上创建新进程，桌面不可见，也不能接收用户输入。要启用用户与新进程的交互，必须在STARTUPINFO结构的lpDesktop成员中指定默认的交互窗口站和桌面的名称“winsta0\default”。此外，在调用CreateProcessAsUser之前，必须更改默认交互窗口站和默认桌面的可自由支配访问控制列表(discretionary access control list, DACL)。窗口站和桌面的DACLs必须授予对用户或由hToken参数表示的登录会话的访问权。
'    CreateProcessAsUser不将指定用户的概要文件加载到HKEY_USERS注册表项中。因此，要访问HKEY_CURRENT_USER注册表项中的信息，在调用CreateProcessAsUser之前，必须使用LoadUserProfile函数将用户的概要信息加载到HKEY_USERS中。确保在新进程退出后调用UnloadUserProfile。
'    如果lpEnvironment参数为NULL，则新进程将继承调用进程的环境。CreateProcessAsUser不会自动修改环境块以包含特定于由hToken表示的用户的环境变量。例如，如果lpEnvironment为空，则从调用进程继承USERNAME和USERDOMAIN变量。您的职责是为新流程准备环境块并在lpEnvironment中指定它。
'    CreateProcessWithLogonW和CreateProcessWithTokenW函数类似于CreateProcessAsUser，只是调用者不需要调用LogonUser函数来对用户进行身份验证并获取令牌。
'    CreateProcessAsUser允许您访问调用者或目标用户的安全上下文中指定的目录和可执行映像。默认情况下，CreateProcessAsUser访问调用者安全上下文中的目录和可执行映像。在这种情况下，如果调用者没有访问目录和可执行映像的权限，则函数将失败。要使用目标用户的安全上下文访问目录和可执行映像，请在调用CreateProcessAsUser之前，在调用ImpersonateLoggedOnUser函数时指定hToken。
'    流程被分配一个流程标识符。标识符在进程终止之前有效。它可以用来标识进程，或者在OpenProcess函数中指定打开进程句柄。进程中的初始线程也被分配一个线程标识符。可以在OpenThread函数中指定它来打开线程的句柄。标识符在线程终止之前是有效的，并且可以用于惟一地标识系统中的线程。这些标识符在PROCESS_INFORMATION结构中返回。
'    调用线程可以使用WaitForInputIdle函数等待，直到新进程完成初始化，并且正在等待用户输入而没有输入挂起。这对于父进程和子进程之间的同步非常有用，因为CreateProcessAsUser返回时不需要等待新进程完成初始化。例如，创建流程将在尝试查找与新流程关联的窗口之前使用WaitForInputIdle。
'    关闭进程的首选方法是使用ExitProcess函数，因为该函数向附加到进程的所有dll发送终止通知。其他关闭进程的方法不通知附加的dll。注意，当一个线程调用ExitProcess时，进程的其他线程将被终止，而没有机会执行任何其他代码(包括附加dll的线程终止代码)。有关更多信息，请参见终止进程。
'安全备注
'    lpApplicationName参数可以为NULL，在这种情况下，可执行名称必须是lpCommandLine中第一个空格分隔的字符串。如果可执行文件或路径名称中有空格，则由于函数解析空格的方式，可能会运行不同的可执行文件。下面的例子很危险，因为函数将尝试运行“Program”。，如果它存在，而不是“MyApp.exe”。
'       LPTSTR szCmdline[] = _tcsdup(TEXT("C:\\Program Files\\MyApp"));
'       CreateProcessAsUser(hToken, NULL, szCmdline， /*…* /);
'    如果恶意用户要创建一个名为“程序”的应用程序。在系统上，任何使用程序文件目录错误调用CreateProcessAsUser的程序都将运行此应用程序，而不是预期的应用程序。
'    为了避免这个问题，不要为lpApplicationName传递NULL。如果确实为lpApplicationName传递NULL，请在lpCommandLine中的可执行路径周围使用引号，如下面的示例所示。
'       LPTSTR szCmdline[] = _tcsdup(TEXT("\"C:\\Program Files\\MyApp\""));
'       CreateProcessAsUser(hToken, NULL, szCmdline， /*…*/);
'    PowerShell:当在PowerShell 2.0版本中使用CreateProcessAsUser函数实现cmdlet时，cmdlet对于扇入和扇出远程会话都可以正确地运行。但是，由于某些安全场景，使用CreateProcessAsUser实现的cmdlet只能在PowerShell version 3.0中为扇入远程会话正确运行;扇出远程会话将由于客户机安全特权不足而失败。要在PowerShell 3.0版本中实现一个同时适用于扇入和扇出远程会话的cmdlet，请使用CreateProcess函数。
'支持
'    最低支持客户
'       Windows XP [只适用于桌面应用程式]
'    最低支持服务器
'       Windows Server 2003[只适用于桌面应用程序]
'    Header
'       Winbase.h (包括Windows.h)
'    Library
'       advapi32.lib
'    dll
'       advapi32.dll
'    Unicode和ANSI名称
'       CreateProcessAsUserW (Unicode) and CreateProcessAsUserA (ANSI)
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
'@原型
'    BOOL WINAPI CloseHandle(
'      _In_ HANDLE hObject
'    );
'@功能
'    关闭一个打开的对象句柄。
'@参数
'    *hObject:
'    一个有效的打开的对象句柄。
'@返回值
'    如果函数成功，则返回值是非0。
'    如果函数失败，则返回值为0。要获得扩展错误信息，请调用GetLastError。
'    如果应用程序在调试器下运行，那么该函数将抛出一个异常，如果它接收到的不是有效的句柄值或伪句柄值。如果关闭一个句柄两次，或者CloseHandle关闭调用FindFirstFile函数返回的句柄，而不是调用FindClose函数，就会发生这种情况。
'@附注
'    CloseHandle关闭如下对象句柄:
'        Access token 访问令牌
'        Communications device 通讯设备
'        Console input 控制台输入
'        Console screen buffer 控制台屏幕缓冲区
'        Event 事件()
'        File 文件
'        File mapping 文件映射
'        I/O completion port I / O完成端口
'        Job 任务
'        Mailslot 邮槽
'        Memory resource notification 内存资源的通知
'        Mutex 互斥锁
'        Named Pipe 命名管道
'        Pipe 管道
'        Process 进程
'        Semaphore 信号量
'        Thread 线程
'        Transaction 事务
'        Waitable Timer 可等待定时器
'    创建这些对象的函数的文档表明，当您完成该对象时，应该使用CloseHandle，以及在该句柄关闭后对对象的待处理操作会发生什么情况。
'        通常， CloseHandle会对指定的对象句柄失效，对对象的句柄计数进行递减，并执行对象保留检查。
'        当对象的最后一个句柄被关闭后，对象将被从系统中删除。有关这些对象的创建者函数的摘要，请参阅Kernel Objects.。
'    通常，应用程序应该为它打开的每个句柄调用一次 CloseHandle。
'        如果使用句柄的函数失败并返回ERROR_INVALID_HANDLE，那么通常没有必要调用CloseHandle，因为这个错误通常表明句柄已经失效。
'        然而，一些函数使用ERROR_INVALID_HANDLE来指示对象本身不再有效。
'        例如，如果网络连接被切断，那么一个试图在网络上使用句柄的函数失败并返回ERROR_INVALID_HANDLE ，因为该文件对象不再可用。在这种情况下，应用程序应该关闭句柄。
'    如果一个句柄是事务，那么在事务提交之前，所有绑定到事务的句柄都应该关闭。
'        如果一个事务句柄通过使用FILE_FLAG_DELETE_ON_CLOSE标志调用CreateFileTransacted 操作来打开，那么在应用程序关闭句柄和调用 CommitTransaction之前，该文件不会被删除。
'        有关事务对象的更多信息，请参见Working With Transactions.。
'    关闭一个线程句柄并不会终止相关的线程，也不会删除线程对象。关闭一个进程句柄并不会终止相关的进程，也不会删除进程对象。
'        要删除一个线程对象，您必须终止线程，然后关闭线程中所有的句柄。要获得更多信息，请参见Terminating a Thread。
'        要删除进程对象，您必须终止进程，然后关闭进程的所有句柄。要了解更多信息，请参见Terminating a Process。
'    即使有file mapping仍然是打开的，关闭一个文件映射的句柄也可以成功。要了解更多信息，请参阅Closing a File Mapping Object.。
'    不要使用CloseHandle关闭一个套接字。相反，使用closesocket函数，它将释放与套接字关联的所有资源，包括套接字对象的句柄。要了解更多信息，请参阅Socket Closure。
'    不要使用CloseHandle关闭一个打开的注册表键的句柄。相反，使用RegCloseKey 函数。CloseHandle 不会关闭对注册表键的句柄，但是不会返回一个错误来表示这个失败。
'@要求
'    Minimum supported client   Windows 2000 Professional [desktop apps | UWP apps]
'    Minimum supported server   Windows 2000 Server [desktop apps | UWP apps]
'    Minimum supported phone    Windows Phone 8
'    Header                     Winbase.h (include Windows.h)
'    Library                    kernel32.lib
'    dll                        kernel32.dll
Private Declare Function GetVersionExA Lib "kernel32.dll" (lpVersionInformation As OSVERSIONINFOEX) As Long
'@原型
'    BOOL WINAPI GetVersionEx(
'      _Inout_ LPOSVERSIONINFO lpVersionInfo
'    );
'@功能
'    [GetVersionEx可能在Windows 8.1之后的版本中被修改或不可用。相反，使用版本帮助函数]
'    随着Windows 8.1的发布，GetVersionEx API的行为发生了变化，它将返回操作系统版本的值。GetVersionEx函数返回的值现在取决于应用程序的显示方式。
'    未在Windows 8.1或Windows 10中显示的应用程序将返回Windows 8 OS版本值(6.2)。一旦为给定的操作系统版本显示了应用程序，GetVersionEx将始终返回应用程序在未来版本中显示的版本。要显示Windows 8.1或Windows 10的应用程序，请参考针对Windows的应用程序。
'@参数
'    lpVersionInfo _Inout_
'    接收操作系统信息的OSVERSIONINFO或OSVERSIONINFOEX结构
'    在调用GetVersionEx函数之前，请设置结构的dwOSVersionInfoSize成员，以指示传递给该函数的数据结构。
'@返回值
'    如果函数成功，返回值为非零值。
'    如果函数失败，返回值为零。要获取扩展的错误信息，请调用GetLastError。如果为OSVERSIONINFO或OSVERSIONINFOEX结构的dwOSVersionInfoSize成员指定无效值，则该函数将失败。
'@备注
'    确定当前操作系统通常不是确定是否存在特定操作系统特性的最佳方法。这是因为操作系统可能在可重新分发的DLL中添加了新特性。与其使用GetVersionEx来确定操作系统平台或版本号，不如测试特性本身的存在性。有关更多信息，请参见操作系统版本。
'    GetSystemMetrics函数提供关于当前操作系统的附加信息
'    产品   设置
'    Windows XP Media Center Edition    SM_MEDIACENTER
'    Windows XP Starter Edition         SM_STARTER
'    Windows XP Tablet PC Edition       SM_TABLETPC
'    Windows Server 2003 R2             SM_SERVERR2
'    要检查特定的操作系统或操作系统特性，请使用IsOS函数。GetProductInfo函数检索产品类型。
'    要检索远程计算机上操作系统的信息，请使用NetWkstaGetInfo函数。Win32_OperatingSystem WMI类或IADsComputer接口的OperatingSystem属性。
'    要将当前系统版本与所需版本进行比较，请使用VerifyVersionInfo函数，而不是使用GetVersionEx自己执行比较。
'    如果兼容模式生效，GetVersionEx函数将报告它标识自身的操作系统，该操作系统可能不是已安装的操作系统。例如，如果兼容性模式生效，GetVersionEx将报告为应用程序兼容性而选择的操作系统。
'@要求
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
'@原型
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
'@功能
'    包含操作系统版本信息。这些信息包括主版本号和次版本号。构建号。平台标识符，以及关于产品套件和安装在系统上的最新服务包的信息。此结构用于GetVersionEx和VerifyVersionInfo函数。
'@成员
'    dwOSVersionInfoSize
'       此数据结构的大小，以字节为单位。将此成员设置为sizeof(OSVERSIONINFOEX)。
'    dwMajorVersion
'       操作系统的主要版本号。有关更多信息，请参见备注。
'    dwMinorVersion
'       操作系统的次要版本号。有关更多信息，请参见备注。
'    dwBuildNumber
'       操作系统的构建号。
'    dwPlatformId
'       操作系统平台。这个成员可以是VER_PLATFORM_WIN32_NT(2)。
'    szCSDVersion
'       一个以null结尾的字符串，如“Service Pack 3”，表示系统上安装的最新服务包。如果没有安装服务包，则字符串为空。
'    wServicePackMajor
'       系统上安装的最新服务包的主要版本号。例如，对于Service Pack 3，主要版本号是3。如果没有安装任何服务包，则该值为零。
'    wServicePackMinor
'       系统上安装的最新服务包的次要版本号。例如，对于Service Pack 3，次要版本号是0。
'    wSuiteMask
'       一个位掩码，用于标识系统上可用的产品套件。这个成员可以是以下值的组合。
Private Const VER_SUITE_BACKOFFICE              As Long = &H4
'    安装了Microsoft BackOffice组件。
Private Const VER_SUITE_BLADE                   As Long = &H400
'    安装Windows Server 2003, Web Edition。
Private Const VER_SUITE_COMPUTE_SERVER          As Long = &H4000
'    安装Windows Server 2003，计算机集群版。
Private Const VER_SUITE_DATACENTER              As Long = &H80
'    安装了Windows Server 2008数据中心。Windows Server 2003。数据中心版本或Windows 2000数据中心服务器。
Private Const VER_SUITE_ENTERPRISE              As Long = &H2
'    安装Windows Server 2008企业版。Windows Server 2003。企业版或Windows 2000高级服务器。有关此位标志的更多信息，请参阅备注部分。
Private Const VER_SUITE_EMBEDDEDNT              As Long = &H40
'    安装Windows XP嵌入式。
Private Const VER_SUITE_PERSONAL                As Long = &H200
'    安装了Windows Vista家庭高级版。Windows Vista家庭基础版或Windows XP家庭版。
Private Const VER_SUITE_SINGLEUSERTS            As Long = &H100
'    支持远程桌面，但只支持一个交互会话。除非系统在应用服务器模式下运行，否则将设置此值。
Private Const VER_SUITE_SMALLBUSINESS           As Long = &H1
'    Microsoft Small Business Server曾经安装在系统上，但可能已经升级到Windows的另一个版本。有关此位标志的更多信息，请参阅备注部分。
Private Const VER_SUITE_SMALLBUSINESS_RESTRICTED As Long = &H20
'    Microsoft Small Business Server安装时使用了严格的客户端许可证。有关此位标志的更多信息，请参阅备注部分。
Private Const VER_SUITE_STORAGE_SERVER          As Long = &H2000
'    安装Windows Storage Server 2003 R2或Windows Storage Server 2003。
Private Const VER_SUITE_TERMINAL                As Long = &H10
'    安装终端服务.这个值总是被设置
'    如果设置了VER_SUITE_TERMINAL，但没有设置VER_SUITE_SINGLEUSERTS，则系统在应用服务器模式下运行。
Private Const VER_SUITE_WH_SERVER               As Long = &H8000
'    安装Windows家庭服务器。
Private Const VER_SUITE_MULTIUSERTS             As Long = &H20000
'    启用AppServer模式。
'    wProductType
'       关于系统的任何附加信息。这个成员可以是以下值之一。
Private Const VER_NT_DOMAIN_CONTROLLER          As Long = &H2
'    系统为域控制器，操作系统为Windows Server 2012。Windows Server 2008 R2。Windows Server 2008。Windows Server 2003或Windows 2000 Server。
Private Const VER_NT_SERVER                     As Long = &H3
'    操作系统是Windows Server 2012。Windows Server 2008 R2。Windows Server 2008。Windows Server 2003或Windows 2000 Server。
'    注意，同时也是域控制器的服务器被报告为VER_NT_DOMAIN_CONTROLLER，而不是VER_NT_SERVER。
Private Const VER_NT_WORKSTATION                As Long = &H1
'    操作系统为Windows 8。Windows 7。Windows Vista。Windows XP Professional。Windows XP Home Edition或Windows 2000 Professional。
'    wReserved
'       保留以备将来使用。
'备注
'    依赖版本信息不是测试特性的最佳方法。相反，请参考文档了解感兴趣的特性。有关特征检测常用技术的更多信息，请参见操作系统版本。
'    如果您必须需要特定的操作系统，请确保将其作为支持的最低版本使用，而不是为一个操作系统设计测试。这样，您的检测代码将继续在未来版本的Windows上工作。
'    下表总结了支持的Windows版本返回的值。使用标记为“Other”的列中的信息来区分具有相同版本号的操作系统。
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
'        *适用于已在Windows 8.1或Windows 10中显示的应用程序。未在Windows 8.1或Windows 10中显示的应用程序将返回Windows 8 OS版本值(6.2)。要显示Windows 8.1或Windows 10的应用程序，请参考针对Windows的应用程序。
'    您不应该仅依赖VER_SUITE_SMALLBUSINESS标志来确定系统上是否已经安装了Small Business Server，因为在安装此产品套件时设置了此标志和VER_SUITE_SMALLBUSINESS_RESTRICTED标志。如果您将此安装升级到Windows Server标准版，VER_SUITE_SMALLBUSINESS_RESTRICTED标志将被清除—但是，VER_SUITE_SMALLBUSINESS标志将保持设置。如果该安装被进一步升级到Windows Server Enterprise Edition, VER_SUITE_SMALLBUSINESS标志将保持设置。
'    如果兼容模式有效，则OSVERSIONINFOEX结构包含有关为应用程序兼容性而选择的操作系统的信息。
'    要确定基于win32的应用程序是否在WOW64上运行，请调用IsWow64Process函数。要确定系统是否运行64位版本的Windows，请调用GetNativeSystemInfo函数。
'    GetSystemMetrics函数提供了关于当前操作系统的以下附加信息。
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
'@原型
'    BOOL WINAPI IsWindowsXPOrGreater(void);
'@功能
'    指示当前操作系统版本是否匹配或大于Windows XP版本。
'@参数
'    这个函数没有参数
'@返回值
'    如果当前操作系统版本匹配或大于Windows XP版本，则为True;否则,假的。
'@备注
'   此函数不区分客户机和服务器版本。如果当前OS版本号等于或高于调用中指定的客户机版本，则返回true。例如，对iswindowsxpsp3或更高版本的调用将在Windows Server 2008上返回true。需要区分Windows的服务器和客户机版本的应用程序应该调用IsWindowsServer。
'   对于Windows服务器版本号没有与Windows客户机版本共享的情况，可以使用iswindowsversionor或更大版本进行确认。
'@Requirements
'Minimum supported client        Windows 2000 Professional [desktop apps only]
'Minimum supported server        Windows 2000 Server [desktop apps only]
'Header                          VersionHelpers.h
'Library                         Kernel32.lib;   Ntdll.lib
'dll                             Kernel32.dll    Ntdll.dll
'Private Declare Function IsWindowsXPSP1OrGreater Lib "kernel32.dll" () As Long
'@原型
'    BOOL WINAPI IsWindowsXPSP1OrGreater(void);
'@功能
'    指示当前操作系统版本是否匹配或大于Windows XP with Service Pack 1 (SP1)版本。
'@参数
'    这个函数没有参数
'@返回值
'    如果当前操作系统版本与SP1版本匹配，或大于Windows XP版本，则为True;否则,假的。
'@备注
'    此函数不区分客户机和服务器版本。如果当前OS版本号等于或高于调用中指定的客户机版本，则返回true。例如，对iswindowsxpsp3或更高版本的调用将在Windows Server 2008上返回true。需要区分Windows的服务器和客户机版本的应用程序应该调用IsWindowsServer。
'    对于Windows服务器版本号没有与Windows客户机版本共享的情况，可以使用iswindowsversionor或更大版本进行确认。
'@Requirements
'Minimum supported client        Windows 2000 Professional [desktop apps only]
'Minimum supported server        Windows 2000 Server [desktop apps only]
'Header                          VersionHelpers.h
'Library                         Kernel32.lib;   Ntdll.lib
'dll                             Kernel32.dll    Ntdll.dll
'Private Declare Function IsWindowsXPSP2OrGreater Lib "kernel32.dll" () As Long
'@原型
'    BOOL WINAPI IsWindowsXPSP2OrGreater(void);
'@功能
'    指示当前操作系统版本是否匹配或大于Windows XP with Service Pack 2 (SP2)版本。
'@参数
'    这个函数没有参数
'@返回值
'    如果当前操作系统版本与SP2版本匹配，或大于Windows XP版本，则为True;否则,假的。
'@备注
'    此函数不区分客户机和服务器版本。如果当前OS版本号等于或高于调用中指定的客户机版本，则返回true。例如，对iswindowsxpsp3或更高版本的调用将在Windows Server 2008上返回true。需要区分Windows的服务器和客户机版本的应用程序应该调用IsWindowsServer。
'    对于Windows服务器版本号没有与Windows客户机版本共享的情况，可以使用iswindowsversionor或更大版本进行确认。
'@Requirements
'Minimum supported client        Windows 2000 Professional [desktop apps only]
'Minimum supported server        Windows 2000 Server [desktop apps only]
'Header                          VersionHelpers.h
'Library                         Kernel32.lib;   Ntdll.lib
'dll                             Kernel32.dll    Ntdll.dll
'Private Declare Function IsWindowsXPSP3OrGreater Lib "kernel32.dll" () As Long
'@原型
'    BOOL WINAPI IsWindowsXPSP3OrGreater(void);
'@功能
'    指示当前操作系统版本是否匹配或大于Windows XP with Service Pack 3 (SP3)版本。
'@参数
'    这个函数没有参数
'@返回值
'    如果当前操作系统版本与SP3版本匹配，或大于Windows XP版本，则为True;否则,假的。
'@备注
'    此函数不区分客户机和服务器版本。如果当前OS版本号等于或高于调用中指定的客户机版本，则返回true。例如，对iswindowsxpsp3或更高版本的调用将在Windows Server 2008上返回true。需要区分Windows的服务器和客户机版本的应用程序应该调用IsWindowsServer。
'    对于Windows服务器版本号没有与Windows客户机版本共享的情况，可以使用iswindowsversionor或更大版本进行确认。
'@Requirements
'Minimum supported client        Windows 2000 Professional [desktop apps only]
'Minimum supported server        Windows 2000 Server [desktop apps only]
'Header                          VersionHelpers.h
'Library                         Kernel32.lib;   Ntdll.lib
'dll                             Kernel32.dll    Ntdll.dll
'Private Declare Function IsWindowsVistaOrGreater Lib "kernel32.dll" () As Long
'@原型
'    BOOL WINAPI IsWindowsVistaOrGreater(void);
'@功能
'    指示当前操作系统版本是否匹配或大于Windows Vista版本。
'@参数
'    这个函数没有参数
'@返回值
'    如果当前操作系统版本与Windows Vista版本匹配，或大于Windows Vista版本，则为True;否则,假的。
'@备注
'    此函数不区分客户机和服务器版本。如果当前OS版本号等于或高于调用中指定的客户机版本，则返回true。例如，对iswindowsxpsp3或更高版本的调用将在Windows Server 2008上返回true。需要区分Windows的服务器和客户机版本的应用程序应该调用IsWindowsServer。
'    对于Windows服务器版本号没有与Windows客户机版本共享的情况，可以使用iswindowsversionor或更大版本进行确认。
'@Requirements
'Minimum supported client        Windows 2000 Professional [desktop apps only]
'Minimum supported server        Windows 2000 Server [desktop apps only]
'Header                          VersionHelpers.h
'Library                         Kernel32.lib;   Ntdll.lib
'dll                             Kernel32.dll    Ntdll.dll
'Private Declare Function IsWindowsVistaSP1OrGreater Lib "kernel32.dll" () As Long
'@原型
'    BOOL WINAPI IsWindowsVistaSP1OrGreater(void);
'@功能
'    指示当前操作系统版本是否匹配或大于Windows Vista with Service Pack 1 (SP1)版本。
'@参数
'    这个函数没有参数
'@返回值
'    如果当前操作系统版本与Windows Vista SP1版本匹配，或大于Windows Vista SP1版本，则为True;否则,假的。
'@备注
'    此函数不区分客户机和服务器版本。如果当前OS版本号等于或高于调用中指定的客户机版本，则返回true。例如，对iswindowsxpsp3或更高版本的调用将在Windows Server 2008上返回true。需要区分Windows的服务器和客户机版本的应用程序应该调用IsWindowsServer。
'    对于Windows服务器版本号没有与Windows客户机版本共享的情况，可以使用iswindowsversionor或更大版本进行确认。
'@Requirements
'Minimum supported client        Windows 2000 Professional [desktop apps only]
'Minimum supported server        Windows 2000 Server [desktop apps only]
'Header                          VersionHelpers.h
'Library                         Kernel32.lib;   Ntdll.lib
'dll                             Kernel32.dll    Ntdll.dll
'Private Declare Function IsWindowsVistaSP2OrGreater Lib "kernel32.dll" () As Long
'@原型
'    BOOL WINAPI IsWindowsVistaSP2OrGreater(void);
'@功能
'    指示当前操作系统版本是否匹配或大于Windows Vista with Service Pack 2 (SP2)版本。
'@参数
'    这个函数没有参数
'@返回值
'    如果当前操作系统版本与Windows Vista SP2版本匹配，或大于Windows Vista SP2版本，则为True;否则,假的。
'@备注
'    此函数不区分客户机和服务器版本。如果当前OS版本号等于或高于调用中指定的客户机版本，则返回true。例如，对iswindowsxpsp3或更高版本的调用将在Windows Server 2008上返回true。需要区分Windows的服务器和客户机版本的应用程序应该调用IsWindowsServer。
'    对于Windows服务器版本号没有与Windows客户机版本共享的情况，可以使用iswindowsversionor或更大版本进行确认。
'@Requirements
'Minimum supported client        Windows 2000 Professional [desktop apps only]
'Minimum supported server        Windows 2000 Server [desktop apps only]
'Header                          VersionHelpers.h
'Library                         Kernel32.lib;   Ntdll.lib
'dll                             Kernel32.dll    Ntdll.dll
'Private Declare Function IsWindows7OrGreater Lib "kernel32.dll" () As Long
'@原型
'    BOOL WINAPI IsWindows7OrGreater(void);
'@功能
'    指示当前操作系统版本是否匹配或大于Windows 7版本。
'@参数
'    这个函数没有参数
'@返回值
'    如果当前操作系统版本与Windows 7版本匹配，或大于Windows 7版本，则为True;否则,假的。
'@备注
'    此函数不区分客户机和服务器版本。如果当前OS版本号等于或高于调用中指定的客户机版本，则返回true。例如，对iswindowsxpsp3或更高版本的调用将在Windows Server 2008上返回true。需要区分Windows的服务器和客户机版本的应用程序应该调用IsWindowsServer。
'    对于Windows服务器版本号没有与Windows客户机版本共享的情况，可以使用iswindowsversionor或更大版本进行确认。
'@Requirements
'Minimum supported client        Windows 2000 Professional [desktop apps only]
'Minimum supported server        Windows 2000 Server [desktop apps only]
'Header                          VersionHelpers.h
'Library                         Kernel32.lib;   Ntdll.lib
'dll                             Kernel32.dll    Ntdll.dll
'Private Declare Function IsWindows7SP1OrGreater Lib "kernel32.dll" () As Long
'@原型
'    BOOL WINAPI IsWindows7SP1OrGreater(void);
'@功能
'    指示当前操作系统版本是否匹配或大于Windows 7 with Service Pack 1 (SP1)版本。
'@参数
'    这个函数没有参数
'@返回值
'    如果当前操作系统版本与Windows 7 SP1版本匹配，或大于Windows 7 SP1版本，则为True;否则,假的。
'@备注
'    此函数不区分客户机和服务器版本。如果当前OS版本号等于或高于调用中指定的客户机版本，则返回true。例如，对iswindowsxpsp3或更高版本的调用将在Windows Server 2008上返回true。需要区分Windows的服务器和客户机版本的应用程序应该调用IsWindowsServer。
'    对于Windows服务器版本号没有与Windows客户机版本共享的情况，可以使用iswindowsversionor或更大版本进行确认。
'@Requirements
'Minimum supported client        Windows 2000 Professional [desktop apps only]
'Minimum supported server        Windows 2000 Server [desktop apps only]
'Header                          VersionHelpers.h
'Library                         Kernel32.lib;   Ntdll.lib
'dll                             Kernel32.dll    Ntdll.dll
'Private Declare Function IsWindows8OrGreater Lib "kernel32.dll" () As Long
'@原型
'    BOOL WINAPI IsWindows8OrGreater(void);
'@功能
'    指示当前操作系统版本是否匹配或大于Windows 8版本。
'@参数
'    这个函数没有参数
'@返回值
'    如果当前操作系统版本与Windows 8版本匹配，或大于Windows 8版本，则为True;否则,假的。
'@备注
'    此函数不区分客户机和服务器版本。如果当前OS版本号等于或高于调用中指定的客户机版本，则返回true。例如，对iswindowsxpsp3或更高版本的调用将在Windows Server 2008上返回true。需要区分Windows的服务器和客户机版本的应用程序应该调用IsWindowsServer。
'    对于Windows服务器版本号没有与Windows客户机版本共享的情况，可以使用iswindowsversionor或更大版本进行确认。
'@Requirements
'Minimum supported client        Windows 2000 Professional [desktop apps only]
'Minimum supported server        Windows 2000 Server [desktop apps only]
'Header                          VersionHelpers.h
'Library                         Kernel32.lib;   Ntdll.lib
'dll                             Kernel32.dll    Ntdll.dll
'Private Declare Function IsWindows8Point1OrGreater Lib "kernel32.dll" () As Long
'@原型
'    BOOL WINAPI IsWindows8Point1OrGreater(void);
'@功能
'    指示当前操作系统版本是否匹配或大于Windows 8.1版本。对于Windows 10, IsWindows8Point1OrGreater返回false，除非应用程序包含一个清单，其中包含一个兼容性部分，其中包含指定Windows 8.1和/或Windows 10的guid。
'@参数
'    这个函数没有参数
'@返回值
'    如果当前操作系统版本与Windows 8.1版本匹配，或大于Windows 8.1版本，则为True;否则,假的。
'@备注
'    此函数不区分客户机和服务器版本。如果当前OS版本号等于或高于调用中指定的客户机版本，则返回true。例如，对iswindowsxpsp3或更高版本的调用将在Windows Server 2008上返回true。需要区分Windows的服务器和客户机版本的应用程序应该调用IsWindowsServer。
'    对于Windows服务器版本号没有与Windows客户机版本共享的情况，可以使用iswindowsversionor或更大版本进行确认。
'@Requirements
'Minimum supported client        Windows 2000 Professional [desktop apps only]
'Minimum supported server        Windows 2000 Server [desktop apps only]
'Header                          VersionHelpers.h
'Library                         Kernel32.lib;   Ntdll.lib
'dll                             Kernel32.dll    Ntdll.dll
'Private Declare Function IsWindows10OrGreater Lib "kernel32.dll" () As Long
'@原型
'    BOOL WINAPI IsWindows10OrGreater(void);
'@功能
'    指示当前操作系统版本是否匹配或大于Windows 10版本。对于Windows10, IsWindows10OrGreater返回false，除非应用程序包含一个清单，其中包含一个兼容性部分，其中包含指定Windows10的GUID。
'@参数
'    这个函数没有参数
'@返回值
'    如果当前操作系统版本与Windows 10版本匹配，或大于Windows 10版本，则为True;否则,假的。
'@备注
'    没有显示Windows 10的应用程序返回false，即使当前的操作系统版本是Windows 10。要显示针对Windows 10的应用程序，请参见针对Windows的应用程序。
'    此函数不区分客户机和服务器版本。如果当前OS版本号等于或高于调用中指定的客户机版本，则返回true。例如，对iswindowsxpsp3或更高版本的调用将在Windows Server 2008上返回true。需要区分Windows的服务器和客户机版本的应用程序应该调用IsWindowsServer。
'    对于Windows服务器版本号没有与Windows客户机版本共享的情况，可以使用iswindowsversionor或更大版本进行确认。
'@Requirements
'Minimum supported client        Windows 2000 Professional [desktop apps only]
'Minimum supported server        Windows 2000 Server [desktop apps only]
'Header                          VersionHelpers.h
'Library                         Kernel32.lib;   Ntdll.lib
'dll                             Kernel32.dll    Ntdll.dll
'Private Declare Function IsWindowsServer Lib "kernel32.dll" () As Long
'@原型
'    BOOL WINAPI IsWindowsServer(void);
'@功能
'    指示当前操作系统是否是Windows服务器版本。需要区分Windows的服务器和客户机版本的应用程序应该调用这个函数。 注意，只有当其他提供的版本帮助器函数不适合您的场景时，才应该使用此函数
'@参数
'    这个函数没有参数
'@返回值
'    如果当前操作系统是Windows服务器版本，则为True;否则,假的。。
'@Requirements
'Minimum supported client        Windows 2000 Professional [desktop apps only]
'Minimum supported server        Windows 2000 Server [desktop apps only]
'Header                          VersionHelpers.h
'Library                         Kernel32.lib;   Ntdll.lib
'dll                             Kernel32.dll    Ntdll.dll
'Private Declare Function IsWindowsVersionOrGreater Lib "kernel32.dll" (ByVal wMajorVersion As Integer, ByVal wMinorVersion As Integer, ByVal wServicePackMajor As Integer) As Long
'@原型
'    BOOL WINAPI IsWindowsVersionOrGreater(
'       WORD wMajorVersion,
'       WORD wMinorVersion,
'       WORD wServicePackMajor
'    );
'@功能
'    重要的是，您应该只在其他提供的版本帮助器函数不适合您的场景时才使用此函数。
'    指示当前OS版本是否匹配或大于提供的版本信息。此函数用于确认Windows服务器版本是否与客户机版本号共享。
'@参数
'    wMajorVersion
'       主要操作系统版本号
'    wMinorVersion
'       次要OS版本号
'    wServicePackMajor
'       主要服务包版本号
'@返回值
'    如果指定的版本与当前Windows操作系统的版本匹配，或大于该版本，则为TRUE;否则,假的。
Private Declare Function OpenThreadToken Lib "advapi32.dll" (ByVal ThreadHandle As Long, ByVal DesiredAccess As Long, ByVal OpenAsSelf As Long, TokenHandle As Long) As Long
'@原型
'    BOOL WINAPI OpenThreadToken(
'      _In_  HANDLE  ThreadHandle,
'      _In_  DWORD   DesiredAccess,
'      _In_  BOOL    OpenAsSelf,
'      _Out_ PHANDLE TokenHandle
'    );
'@功能
'    OpenThreadToken函数打开与线程关联的访问令牌。
'@参数
'ThreadHandle _In_
'   打开访问令牌的线程的句柄。
'DesiredAccess _In_]
'   指定访问掩码，该掩码指定访问令牌的请求访问类型。这些请求的访问类型与令牌的自由访问控制列表(discretionary access control list, DACL)相协调，以确定授予或拒绝哪些访问。
'   有关访问令牌的访问权限列表，请参见访问令牌对象的访问权限。
'OpenAsSelf _In_
'   如果要对进程级安全上下文进行访问检查，则为TRUE。
'   如果要对调用OpenThreadToken函数的线程的当前安全上下文进行访问检查，则为FALSE。
'   OpenAsSelf参数允许此函数的调用者在调用者模拟安全标识级别的令牌时打开指定线程的访问令牌。没有此参数，调用线程无法打开指定线程上的访问令牌，因为无法使用SecurityIdentification模拟级别打开执行级对象。
'TokenHandle _Out_
'   指向变量的指针，该变量接收新打开的访问令牌的句柄。
'@返回值
'   如果函数成功，返回值为非零。
'   如果函数失败，返回值为零。要获取扩展的错误信息，请调用GetLastError。如果令牌具有匿名模拟级别，则不会打开令牌，OpenThreadToken将ERROR_CANT_OPEN_ANONYMOUS设置为错误。
'@备注
'   无法打开具有匿名模拟级别的令牌。
'   通过调用Close句柄关闭通过TokenHandle参数返回的访问令牌句柄
'@Requirements
'Minimum supported client        Windows XP [desktop apps | UWP apps]
'Minimum supported server        Windows Server 2003 [desktop apps | UWP apps]
'Header                          Winbase.h (include Windows.h)
'Library                         Advapi32.lib
'dll                             Advapi32.dll
Private Declare Function OpenProcessToken Lib "advapi32.dll" (ByVal ProcessHandle As Long, ByVal DesiredAccess As Long, TokenHandle As Long) As Long
'@原型
'    BOOL WINAPI OpenProcessToken(
'      _In_  HANDLE  ProcessHandle,
'      _In_  DWORD   DesiredAccess,
'      _Out_ PHANDLE TokenHandle
'    );
'@功能
'    OpenProcessToken函数打开与进程关联的访问令牌。
'@参数
'ProcessHandle  _In_
'   进程的句柄，其访问令牌已打开。流程必须具有PROCESS_QUERY_INFORMATION访问权限。
'DesiredAccess _In_
'   指定访问掩码，该掩码指定访问令牌的请求访问类型。将这些请求的访问类型与令牌的自由访问控制列表(discretionary access control list, DACL)进行比较，以确定授予或拒绝哪些访问。
'   有关访问令牌的访问权限列表，请参见访问令牌对象的访问权限。
'TokenHandle _Out_
'   指向句柄的指针，该句柄在函数返回时标识新打开的访问令牌。
'@返回值
'   如果函数成功，返回值为非零。
'   如果函数失败，返回值为零。要获取扩展的错误信息，请调用GetLastError。
'@备注
'   通过调用Close句柄关闭通过TokenHandle参数返回的访问令牌句柄。
'@Requirements
'Minimum supported client        Windows XP [desktop apps | UWP apps]
'Minimum supported server        Windows Server 2003 [desktop apps | UWP apps]
'Header                          Winbase.h (include Windows.h)
'Library                         Advapi32.lib
'dll                             Advapi32.dll
Private Declare Function SetTokenInformation Lib "advapi32.dll" (ByVal TokenHandle As Long, ByVal TokenInformationClass As TOKEN_INFORMATION_CLASS, TokenInformation As Long, ByVal TokenInformationLength As Long) As Long
Private Declare Function GetTokenInformation Lib "advapi32.dll" (ByVal TokenHandle As Long, ByVal TokenInformationClass As TOKEN_INFORMATION_CLASS, TokenInformation As Long, ByVal TokenInformationLength As Long, ReturnLength As Long) As Long
'@原型
'    BOOL WINAPI GetTokenInformation(
'      _In_      HANDLE                  TokenHandle,
'      _In_      TOKEN_INFORMATION_CLASS TokenInformationClass,
'      _Out_opt_ LPVOID                  TokenInformation,
'      _In_      DWORD                   TokenInformationLength,
'      _Out_     PDWORD                  ReturnLength
'    );
'@功能
'    GetTokenInformation函数检索关于访问令牌的指定类型的信息。调用进程必须具有适当的访问权限才能获得信息
'    要确定用户是否是特定组的成员，请使用CheckTokenMembership函数。要确定应用程序容器令牌的组成员关系，请使用CheckTokenMembershipEx函数。
'@参数
'TokenHandle _In_
'   获取信息的访问令牌的句柄。如果TokenInformationClass指定了TokenSource，句柄必须具有TOKEN_QUERY_SOURCE访问权。对于所有其他TokenInformationClass值，句柄必须具有TOKEN_QUERY访问权。
'TokenInformationClass _In_
'   从TOKEN_INFORMATION_CLASS枚举类型指定一个值，以标识函数检索的信息的类型。任何检查TokenIsAppContainer并让它返回0的调用者还应该验证调用者令牌不是标识级别模拟令牌。如果当前令牌不是应用程序容器，而是标识级别令牌，则应返回拒绝访问。
'TokenInformation _Out_opt_
'   指向缓冲区的指针，该函数用所请求的信息填充缓冲区。放入此缓冲区的结构取决于TokenInformationClass参数指定的信息类型。
'TokenInformationLength _In_
'   指定TokenInformation参数指向的缓冲区的大小(以字节为单位)。如果TokenInformation为空，则此参数必须为零。
'ReturnLength _Out_
'   指向一个变量的指针，该变量接收TokenInformation参数指向的缓冲区所需的字节数。如果该值大于TokenInformationLength参数中指定的值，则函数将失败，并在缓冲区中不存储任何数据。
'   如果TokenInformationClass参数的值是TokenDefaultDacl，而令牌没有默认的DACL，则函数将ReturnLength指向的变量设置为sizeof(TOKEN_DEFAULT_DACL)，并将TOKEN_DEFAULT_DACL结构的DefaultDacl成员设置为NULL。
'@返回值
'    如果函数成功，返回值为非零。
'    如果函数失败，返回值为零。要获取扩展的错误信息，请调用GetLastError。
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
'@功能
'    TOKEN_INFORMATION_CLASS枚举包含指定分配给访问令牌或从访问令牌检索的信息类型的值。
'    GetTokenInformation函数使用这些值来指示要检索的令牌信息的类型。
'    SetTokenInformation函数使用这些值设置令牌信息
'@常量
'TokenUser
'   缓冲区接收一个TOKEN_USER结构，该结构包含令牌的用户帐户。
'TokenGroups
'   缓冲区接收一个TOKEN_GROUPS结构，其中包含与令牌关联的组帐户。
'TokenPrivileges
'   缓冲区接收一个包含令牌特权的TOKEN_PRIVILEGES结构
'TokenOwner
'   缓冲区接收一个TOKEN_OWNER结构，该结构包含新创建对象的默认所有者安全标识符(SID)。
'TokenPrimaryGroup
'   缓冲区接收一个TOKEN_PRIMARY_GROUP结构，其中包含新创建对象的默认主组SID。
'TokenDefaultDacl
'   缓冲区接收一个TOKEN_DEFAULT_DACL结构，该结构包含新创建对象的默认DACL。
'TokenSource
'   缓冲区接收一个包含令牌源的TOKEN_SOURCE结构。检索此信息需要访问TOKEN_QUERY_SOURCE。
'TokenType
'   缓冲区接收一个TOKEN_TYPE值，该值指示令牌是主令牌还是模拟令牌。
'TokenImpersonationLevel
'   缓冲区接收SECURITY_IMPERSONATION_LEVEL值，该值指示令牌的模拟级别。如果访问令牌不是模拟令牌，则函数将失败。
'TokenStatistics
'   缓冲区接收一个TOKEN_STATISTICS结构，其中包含各种令牌统计信息。
'TokenRestrictedSids
'   缓冲区接收一个TOKEN_GROUPS结构，其中包含一个SID的列表。
'TokenSessionId
'   缓冲区接收一个DWORD值，该值指示与令牌关联的终端服务会话标识符。
'   如果令牌与终端服务器客户机会话相关联，则会话标识符为非零。
'   Windows Server 2003和Windows XP:如果令牌与终端服务器控制台会话相关联，则会话标识符为零。
'   在非终端服务环境中，会话标识符为零。
'   如果TokenSessionId是用SetTokenInformation设置的，那么应用程序必须将该行为作为操作系统特权的一部分，并且应用程序必须能够在令牌中设置会话ID。
'TokenGroupsAndPrivileges
'   缓冲区接收一个TOKEN_GROUPS_AND_PRIVILEGES结构，该结构包含用户SID。组帐户。受限制的SID和与令牌关联的身份验证ID。
'TokenSessionReference
'   保留。
'TokenSandBoxInert
'   如果令牌包含 SANDBOX_INERT标志，缓冲区将接收一个非零的DWORD值。
'TokenAuditPolicy
'   保留。
'TokenOrigin
'   缓冲区接收一个TOKEN_ORIGIN值。
'   如果令牌来自使用显式凭证的登录，例如将名称。域和密码传递给LogonUser函数，那么TOKEN_ORIGIN结构将包含创建它的登录会话的ID。
'   如果令牌是由网络身份验证导致的，例如调用AcceptSecurityContext或调用LogonUser (dwLogonType设置为LOGON32_LOGON_NETWORK或LOGON32_LOGON_NETWORK_CLEARTEXT)，则该值为零。
'TokenElevationType
'   缓冲区接收一个TOKEN_ELEVATION_TYPE值，该值指定令牌的标高级别。
'TokenLinkedToken
'   缓冲区接收一个TOKEN_LINKED_TOKEN结构，该结构包含另一个连接到该令牌的令牌句柄。
'TokenElevation
'   缓冲区接收一个TOKEN_ELEVATION结构，该结构指定是否提升令牌。
'TokenHasRestrictions
'   如果标记被过滤过，缓冲区将接收一个非零的DWORD值。
'TokenAccessInformation
'   缓冲区接收一个TOKEN_ACCESS_INFORMATION结构，该结构指定令牌中包含的安全信息。
'TokenVirtualizationAllowed
'   如果允许对令牌进行虚拟化，缓冲区将接收一个非零的DWORD值。
'TokenVirtualizationEnabled
'   如果为令牌启用了虚拟化，缓冲区将接收一个非零的DWORD值。
'TokenIntegrityLevel
'   缓冲区接收一个TOKEN_MANDATORY_LABEL结构，该结构指定令牌的完整性级别。
'TokenUIAccess
'   如果令牌设置了UIAccess标志，缓冲区将接收一个非零的DWORD值。
'TokenMandatoryPolicy
'   缓冲区接收一个TOKEN_MANDATORY_POLICY结构，该结构指定令牌的强制完整性策略。
'TokenLogonSid
'   缓冲区接收一个TOKEN_GROUPS结构，该结构指定令牌的登录SID。
'TokenIsAppContainer
'   如果令牌是应用程序容器令牌，缓冲区将接收一个非零的DWORD值。任何检查TokenIsAppContainer并让它返回0的调用者还应该验证调用者令牌不是标识级别模拟令牌。如果当前令牌不是应用程序容器，而是标识级别令牌，则应返回拒绝访问。
'TokenCapabilities
'   缓冲区接收一个TOKEN_GROUPS结构，其中包含与令牌关联的功能。
'TokenAppContainerSid
'   缓冲区接收一个TOKEN_APPCONTAINER_INFORMATION结构，该结构包含与令牌关联的AppContainerSid。如果令牌与应用程序容器没有关联，则TOKEN_APPCONTAINER_INFORMATION结构中的TokenAppContainer成员指向NULL。
'TokenAppContainerNumber
'   缓冲区接收一个包含令牌的应用程序容器号的DWORD值。对于不是应用程序容器令牌的令牌，此值为零。
'TokenUserClaimAttributes
'   缓冲区接收一个CLAIM_SECURITY_ATTRIBUTES_INFORMATION结构，该结构包含与令牌关联的用户声明。
'TokenDeviceClaimAttributes
'   缓冲区接收一个CLAIM_SECURITY_ATTRIBUTES_INFORMATION结构，该结构包含与令牌关联的设备声明。
'TokenRestrictedUserClaimAttributes
'   保留此值。
'TokenRestrictedDeviceClaimAttributes
'   保留此值。
'TokenDeviceGroups
'   缓冲区接收一个TOKEN_GROUPS结构，其中包含与令牌关联的设备组。
'TokenRestrictedDeviceGroups
'   缓冲区接收一个TOKEN_GROUPS结构，该结构包含与令牌关联的受限设备组。
'TokenSecurityAttributes
'   保留此值。
'TokenIsRestricted
'   保留此值。
'MaxTokenInfoClass
'   此枚举的最大值。
'@Requirements
'Minimum supported client        Windows XP [desktop apps only]
'Minimum supported server        Windows Server 2003 [desktop apps only]
'Header                          Winnt.h (include Windows.h)
Private Type SID_AND_ATTRIBUTES
    Sid         As Long
    Attributes  As Long
End Type
'@原型
'    typedef struct _SID_AND_ATTRIBUTES {
'      PSID  Sid;
'      DWORD Attributes;
'    } SID_AND_ATTRIBUTES, *PSID_AND_ATTRIBUTES;
'@功能
'    SID_AND_ATTRIBUTES结构表示安全标识符(SID)及其属性。SId被用来唯一地确定用户或群体。
'@成员
'Sid
'   指向SID结构的指针
'Attributes
'   指定SID的属性。此值最多包含32个1位标志。它的含义取决于SID的定义和使用。
'@备注
'   组由SID表示。SID的属性有表明它们目前是启用。禁用还是强制执行的属性。小SID还指出如何使用这些属性。SID_AND_ATTRIBUTES结构可以表示属性经常变化的SID。例如，SID_AND_ATTRIBUTES用于表示TOKEN_GROUPS结构中的组。
'@Requirements
'Minimum supported client        Windows XP [desktop apps only]
'Minimum supported server        Windows Server 2003 [desktop apps only]
'Header                          Winnt.h (include Windows.h)
Private Const ANYSIZE_ARRAY                     As Long = 1
Private Type TOKEN_GROUPS
    GroupCount              As Long
    Groups(ANYSIZE_ARRAY)   As SID_AND_ATTRIBUTES
End Type
'@功能
'   TOKEN_GROUPS结构包含访问令牌中关于组安全标识符(SIDs)的信息。
'@原型
'typedef struct _TOKEN_GROUPS {
'  DWORD              GroupCount;
'  SID_AND_ATTRIBUTES Groups[ANYSIZE_ARRAY];
'} TOKEN_GROUPS, *PTOKEN_GROUPS;
'@成员
'GroupCount
'   指定访问令牌中的组数。
'Groups
'   指定包含一组sid和相应属性的SID_AND_ATTRIBUTES结构数组。
'SID_AND_ATTRIBUTES结构的属性成员可以具有以下值。
Private Const SE_GROUP_ENABLED                  As Long = &H4
'   启用SID进行访问检查。当系统执行访问检查时，它检查应用于SID的访问允许和访问拒绝访问控制项(ace)。
'   没有此属性的SID在访问检查期间将被忽略，除非设置了SE_GROUP_USE_FOR_DENY_ONLY属性。
Private Const SE_GROUP_ENABLED_BY_DEFAULT       As Long = &H2
'   缺省情况下启用SID。
Private Const SE_GROUP_INTEGRITY                As Long = &H20
'   SID是强制的完整性SID。
Private Const SE_GROUP_INTEGRITY_ENABLED        As Long = &H40
'   SID支持强制完整性检查。
Private Const SE_GROUP_LOGON_ID                 As Long = &HC0000000
'   SID是一个登录SID，它标识与访问令牌关联的登录会话。
Private Const SE_GROUP_MANDATORY                As Long = &H1
'   SID不能通过调用调整tokengroups函数清除SE_GROUP_ENABLED属性。但是，您可以使用CreateRestrictedToken函数将强制SID转换为仅拒绝SID。
Private Const SE_GROUP_OWNER                    As Long = &H8
'   SID标识一个组帐户，令牌的用户是该组的所有者，或者可以将SID指定为令牌或对象的所有者。
Private Const SE_GROUP_RESOURCE                 As Long = &H20000000
'   SID标识一个域本地组。
Private Const SE_GROUP_USE_FOR_DENY_ONLY         As Long = &H10
'   在受限令牌中，SID是一个只有否认者的SID。当系统执行访问检查时，它检查应用于SID的访问被拒绝的ace;它忽略SID允许访问的ace。
'   如果设置了此属性，则未设置SE_GROUP_ENABLED，且SID无法重新启用。
'@Requirements
'Minimum supported client        Windows XP [desktop apps only]
'Minimum supported server        Windows Server 2003 [desktop apps only]
'Header                          Winnt.h (include Windows.h)
Private Type SID_IDENTIFIER_AUTHORITY
    value(5)        As Byte
End Type
'@功能：
'    SID_IDENTIFIER_AUTHORITY结构表示安全标识符(SID)的顶级权限。
'@原型
'    typedef struct _SID_IDENTIFIER_AUTHORITY {
'      BYTE Value[6];
'    } SID_IDENTIFIER_AUTHORITY, *PSID_IDENTIFIER_AUTHORITY;
'@成员
'Value
'   指定SID顶级权限的6字节数组。
'@备注
'    标识符权限值标识颁发SID的代理。预定义了以下标识符权限。
Private Const security_null_sid_authority       As Long = &H0
Private Const security_world_sid_authority      As Long = &H1
Private Const security_local_sid_authority      As Long = &H2
Private Const security_creator_sid_authority    As Long = &H3
Private Const security_non_unique_authority     As Long = &H4
Private Const security_nt_authority             As Long = &H5
Private Const security_resource_manager_authority As Long = &H9
'    SID必须包含顶级权限和至少一个相对标识符(RID)值。
'@Requirements
'Minimum supported client        Windows XP [desktop apps only]
'Minimum supported server        Windows Server 2003 [desktop apps only]
'Header                          Winnt.h (include Windows.h)
Private Type LUID
    LowPart             As Long
    HighPart            As Long
End Type
'@原型
'    typedef struct _LUID {
'      DWORD LowPart;
'      LONG  HighPart;
'    } LUID, *PLUID;
'@功能
'    LUID是一个64位值，保证仅在生成它的系统上是惟一的。只有在重新启动系统之前，才能保证本地唯一标识符(LUID)的唯一性。
'    应用程序必须使用函数和结构来操作LUID值。
'@成员
'LowPart
'   低阶位。
'HighPart
'   高阶位。
'@Requirements
'Minimum supported client       Windows XP [desktop apps only]
'Minimum supported server       Windows Server 2003 [desktop apps only]
'Header                         Winnt.h (include Windows.h)
Private Type LUID_AND_ATTRIBUTES
    PLUID       As LUID
    Attributes  As Long
End Type
'@原型
'    typedef struct _LUID_AND_ATTRIBUTES {
'      LUID  Luid;
'      DWORD Attributes;
'    } LUID_AND_ATTRIBUTES, *PLUID_AND_ATTRIBUTES;
'功能
'    LUID_AND_ATTRIBUTES结构表示一个本地惟一标识符(LUID)及其属性。
'@成员
'LUID
'   指定一个LUID值。
'Attributes
'   指定LUID的属性。此值最多包含32个1位标志。它的含义取决于LUID的定义和使用。
'@备注
'   LUID_AND_ATTRIBUTES结构可以表示属性经常变化的LUID，例如当LUID用于表示PRIVILEGE_SET结构中的特权时。特权由luid表示，并具有指示当前是否启用或禁用特权的属性。
'@Requirements
'Minimum supported client       Windows XP [desktop apps only]
'Minimum supported server       Windows Server 2003 [desktop apps only]
'Header                         Winnt.h (include Windows.h)

Private Type TOKEN_PRIVILEGES
    PrivilegeCount              As Long
    Privileges(ANYSIZE_ARRAY)   As LUID_AND_ATTRIBUTES
End Type
'@原型
'    typedef struct _TOKEN_PRIVILEGES {
'      DWORD               PrivilegeCount;
'      LUID_AND_ATTRIBUTES Privileges[ANYSIZE_ARRAY];
'    } TOKEN_PRIVILEGES, *PTOKEN_PRIVILEGES;
'@功能
'    TOKEN_PRIVILEGES结构包含关于访问令牌的一组特权的信息。
'@成员
'PrivilegeCount
'   这必须设置为Privileges数组中的条目数。
'Privileges
'   指定一个LUID_AND_ATTRIBUTES结构数组。每个结构都包含特权的LUID和属性。要获取与LUID关联的特权的名称，请调用LookupPrivilegeName函数，将LUID的地址作为lpLuid参数的值传递。
'   重要的是，常量ANYSIZE_ARRAY在公共头文件Winnt.h中定义为1。要创建包含多个元素的数组，必须为结构分配足够的内存，以考虑其他元素。
'Privileges的属性可以是以下值的组合。
Private Const SE_PRIVILEGE_ENABLED              As Long = &H1
'   启用了特权
Private Const SE_PRIVILEGE_ENABLED_BY_DEFAULT   As Long = &H2
'   默认情况下启用特权。
Private Const SE_PRIVILEGE_REMOVED              As Long = &H4
'   用于删除特权。有关详细信息，请参见AdjustTokenPrivileges。
Private Const SE_PRIVILEGE_USED_FOR_ACCESS      As Long = &H80000000
'   该特权用于访问对象或服务。此标志用于标识客户机应用程序传递的一组中可能包含不必要特权的相关特权。
'@Requirements
'Minimum supported client       Windows XP [desktop apps only]
'Minimum supported server       Windows Server 2003 [desktop apps only]
'Header                         Winnt.h (include Windows.h)
Private Type TOKEN_OWNER
    Owner   As Long
End Type
'@原型
'    typedef struct _TOKEN_OWNER {
'      PSID Owner;
'    } TOKEN_OWNER, *PTOKEN_OWNER;
'@功能
'    TOKEN_OWNER结构包含将应用于新创建对象的缺省所有者安全标识符(SID)。
'@成员
'Owner
'    一个指向SID结构的指针，该结构表示一个用户，该用户将成为使用此访问令牌创建的任何对象的所有者。SID必须是令牌中已经存在的用户或组SID之一。
'@Requirements
'Minimum supported client       Windows XP [desktop apps only]
'Minimum supported server       Windows Server 2003 [desktop apps only]
'Header                         Winnt.h (include Windows.h)
Private Declare Function LookupPrivilegeName Lib "advapi32.dll" Alias "LookupPrivilegeNameA" (ByVal lpSystemName As String, ByRef lpLuid As LUID, ByVal lpName As String, ByRef cchName As Long) As Long
'@原型
'    BOOL WINAPI LookupPrivilegeName(
'      _In_opt_  LPCTSTR lpSystemName,
'      _In_      PLUID   lpLuid,
'      _Out_opt_ LPTSTR  lpName,
'      _Inout_   LPDWORD cchName
'    );
'@参数
'lpSystemName _In_opt_
'   指向以null结尾的字符串的指针，该字符串指定检索特权名称的系统的名称。如果指定了空字符串，该函数将尝试在本地系统上查找特权名称。
'lpLuid _In_
'   一个指向LUID的指针，通过该id可以知道目标系统上的特权。
'lpName _Out_opt_
'   指向缓冲区的指针，该缓冲区接收表示特权名称的以null结尾的字符串。例如，这个字符串可以是“SeSecurityPrivilege”。
'cchName _Inout_
'   指向一个变量的指针，该变量在一个TCHAR值中指定lpName缓冲区的大小。当函数返回时，此参数包含特权名称的长度，不包括终止null字符。如果lpName参数指向的缓冲区太小，则此变量包含所需的大小。
'@返回值
'   如果函数成功，则函数返回非零。
'   如果函数失败，它返回零。要获取扩展的错误信息，请调用GetLastError。
'@备注
'    LookupPrivilegeName函数只支持Winnt.h中定义的特权部分中指定的特权。有关值列表，请参见特权常量。
'@Requirements
'Minimum supported client           Windows XP [desktop apps | UWP apps]
'Minimum supported server           Windows Server 2003 [desktop apps | UWP apps]
'Header                             Winbase.h (include Windows.h)
'Library                            advapi32.lib
'dll                                advapi32.dll
'Unicode and ANSI names             LookupPrivilegeNameW (Unicode) And LookupPrivilegeNameA(ANSI)
'特权常量
'    特权决定用户帐户可以执行的系统操作的类型。管理员将特权分配给用户和组帐户。每个用户的特权包括授予用户和用户所属组的特权。
'    获取和调整访问令牌中的特权的函数使用本地惟一标识符(LUID)类型来标识特权。使用LookupPrivilegeValue函数确定本地系统上对应于特权常量的LUID。使用LookupPrivilegeName函数将LUID转换为相应的字符串常量。
'    操作系统使用下表的Description列中“User Right”后面的字符串表示特权。操作系统在本地安全设置Microsoft管理控制台(MMC)管理单元的用户权限分配节点的策略列中显示用户权限字符串。
Private Const SE_ASSIGNPRIMARYTOKEN_NAME        As String = "SeAssignPrimaryTokenPrivilege"
'   需要分配流程的主要令牌。用户权限: 替换进程级令牌。
Private Const SE_AUDIT_NAME                     As String = "SeAuditPrivilege"
'   需要生成审核日志条目。将此特权授予安全服务器。用户权限: 生成安全审计。
Private Const SE_BACKUP_NAME                    As String = "SeBackupPrivilege"
'   需要执行备份操作。这种特权导致系统将所有读访问控制授予任何文件，而不管为该文件指定的访问控制列表(ACL)是什么。除了read之外的任何访问请求仍然使用ACL进行评估。RegSaveKey和RegSaveKeyExfunctions需要此特权。如果持有以下权限，则授予以下访问权限:
'       READ_CONTROL
'       ACCESS_SYSTEM_SECURITY
'       FILE_GENERIC_READ
'       FILE_TRAVERSE
'   用户权限: 备份文件和目录。
Private Const SE_CHANGE_NOTIFY_NAME             As String = "SeChangeNotifyPrivilege"
'   需要接收文件或目录更改的通知。此特权还会导致系统跳过所有遍历访问检查。它默认为所有用户启用。用户权限: 旁路遍历检查。
Private Const SE_CREATE_GLOBAL_NAME             As String = "SeCreateGlobalPrivilege"
'   需要在终端服务会话期间在全局名称空间中创建命名文件映射对象。管理员。服务和本地系统帐户默认启用此特权。用户权限: 创建全局对象。
Private Const SE_CREATE_PAGEFILE_NAME           As String = "SeCreatePagefilePrivilege"
'   需要创建分页文件。用户权限: 创建页面文件。
Private Const SE_CREATE_PERMANENT_NAME          As String = "SeCreatePermanentPrivilege"
'   需要创建一个永久对象。用户权限: 创建永久共享对象。
Private Const SE_CREATE_SYMBOLIC_LINK_NAME      As String = "SeCreateSymbolicLinkPrivilege"
'   需要创建符号链接。用户权限: 创建符号链接。
Private Const SE_CREATE_TOKEN_NAME              As String = "SeCreateTokenPrivilege"
'   需要创建一个主令牌。用户权限: 创建一个令牌对象。您不能使用“创建令牌对象”策略将此特权添加到用户帐户。此外，不能使用Windows api将此特权添加到拥有的进程。
'   Windows Server 2003和带有SP1及更早版本的Windows XP: Windows api可以将此特权添加到所拥有的进程。
Public Const SE_DEBUG_NAME                     As String = "SeDebugPrivilege"
'   用于调试和调整另一个帐户拥有的进程的内存。用户权限: 调试程序。
Private Const SE_ENABLE_DELEGATION_NAME         As String = "SeEnableDelegationPrivilege"
'   要求将用户和计算机帐户标记为可信的委托帐户。用户权限: 允许委托信任计算机和用户帐户。
Private Const SE_IMPERSONATE_NAME               As String = "SeImpersonatePrivilege"
'   需要模仿。用户权限: 身份验证后模拟客户机。
Private Const SE_INC_BASE_PRIORITY_NAME         As String = "SeIncreaseBasePriorityPrivilege"
'   需要增加进程的基本优先级。用户权限: 增加调度优先级。
Private Const SE_INCREASE_QUOTA_NAME            As String = "SeIncreaseQuotaPrivilege"
'   要求增加分配给进程的配额。用户权限: 调整进程的内存配额。
Private Const SE_INC_WORKING_SET_NAME           As String = "SeIncreaseWorkingSetPrivilege"
'   需要为在用户上下文中运行的应用程序分配更多内存。用户权限: 增加流程工作集。
Private Const SE_LOAD_DRIVER_NAME               As String = "SeLoadDriverPrivilege"
'   需要加载或卸载设备驱动程序。用户权限: 加载和卸载设备驱动程序。
Private Const SE_LOCK_MEMORY_NAME               As String = "SeLockMemoryPrivilege"
'   需要锁定内存中的物理页。用户权限: 锁定内存中的页面。
Private Const SE_MACHINE_ACCOUNT_NAME           As String = "SeMachineAccountPrivilege"
'   需要创建一个计算机帐户。用户权限: 向域添加工作站。
Private Const SE_MANAGE_VOLUME_NAME             As String = "SeManageVolumePrivilege"
'   需要启用卷管理特权。用户权限: 管理卷上的文件。
Private Const SE_PROF_SINGLE_PROCESS_NAME       As String = "SeProfileSingleProcessPrivilege"
'   需要为单个进程收集分析信息。用户权限: 配置单进程。
Private Const SE_RELABEL_NAME                   As String = "SeRelabelPrivilege"
'   需要修改对象的强制完整性级别。用户权限: 修改对象标签。
Private Const SE_REMOTE_SHUTDOWN_NAME           As String = "SeRemoteShutdownPrivilege"
'   需要使用网络请求关闭系统。用户权限: 强制关闭远程系统。
Private Const SE_RESTORE_NAME                   As String = "SeRestorePrivilege"
'   需要执行还原操作。这种特权导致系统将所有写访问控制授予任何文件，而不管为该文件指定的ACL是什么。除了写之外的任何访问请求仍然使用ACL进行评估。此外，此特权允许您将任何有效的用户或组SID设置为文件的所有者。RegLoadKey函数需要此特权。如果持有以下权限，则授予以下访问权限:
'       WRITE_DAC
'       WRITE_OWNER
'       ACCESS_SYSTEM_SECURITY
'       FILE_GENERIC_WRITE
'       FILE_ADD_FILE
'       FILE_ADD_SUBDIRECTORY
'       DELTE
'   用户权限: 还原文件和目录。
Private Const SE_SECURITY_NAME                  As String = "SeSecurityPrivilege"
'   需要执行许多与安全性相关的功能，例如控制和查看审计消息。此特权将其持有者标识为安全操作符。用户权限: 管理审计和安全日志。
Private Const SE_SHUTDOWN_NAME                  As String = "SeShutdownPrivilege"
'   需要关闭本地系统。用户权限: 关闭系统。
Private Const SE_SYNC_AGENT_NAME                As String = "SeSyncAgentPrivilege"
'   域控制器需要使用轻量级目录访问协议目录同步服务。此特权使持有者能够读取目录中的所有对象和属性，而不考虑对象和属性上的保护。默认情况下，它被分配给域控制器上的管理员和本地系统帐户。
'   用户权限: 同步目录服务数据。
Private Const SE_SYSTEM_ENVIRONMENT_NAME        As String = "SeSystemEnvironmentPrivilege"
'   需要修改使用这种类型内存存储配置信息的系统的非易失性RAM。用户权限: 修改固件环境值。
Private Const SE_SYSTEM_PROFILE_NAME            As String = "SeSystemProfilePrivilege"
'   需要为整个系统收集分析信息。用户权限: 配置文件系统性能。
Private Const SE_SYSTEMTIME_NAME                As String = "SeSystemtimePrivilege"
'   需要修改系统时间。用户权利: 更改系统时间。
Private Const SE_TAKE_OWNERSHIP_NAME            As String = "SeTakeOwnershipPrivilege"
'   在不授予可自由支配访问权的情况下获得对象的所有权。此特权允许仅将所有者值设置为持有者作为对象所有者合法分配的值。用户权限: 获取文件或其他对象的所有权。
Private Const SE_TCB_NAME                       As String = "SeTcbPrivilege"
'   此特权将其持有者标识为可信计算机库的一部分。一些受信任的受保护子系统被授予此特权。用户权限: 作为操作系统的一部分。
Private Const SE_TIME_ZONE_NAME                 As String = "SeTimeZonePrivilege"
'   需要调整与计算机内部时钟相关的时区。用户权限: 更改时区。
Private Const SE_TRUSTED_CREDMAN_ACCESS_NAME    As String = "SeTrustedCredManAccessPrivilege"
'   需要以可信调用者的身份访问凭据管理器。用户权限: 作为受信任的调用者访问凭据管理器。
Private Const SE_UNDOCK_NAME                    As String = "SeUndockPrivilege"
'   需要打开笔记本电脑。用户权限: 将计算机从对接口移开。
Private Const SE_UNSOLICITED_INPUT_NAME         As String = "SeUnsolicitedInputPrivilege"
'   要求从终端设备读取非请求输入。用户权利: 不适用。
'@Requirements
'Minimum supported client           Windows XP [desktop apps | UWP apps]
'Minimum supported server           Windows Server 2003 [desktop apps | UWP apps]
'Header                             Winbase.h (include Windows.h)
Private Declare Function LookupPrivilegeValue Lib "advapi32.dll" Alias "LookupPrivilegeValueA" (ByVal lpSystemName As String, ByVal lpName As String, lpLuid As LUID) As Long
'@原型
'    BOOL WINAPI LookupPrivilegeValue(
'      _In_opt_ LPCTSTR lpSystemName,
'      _In_     LPCTSTR lpName,
'      _Out_    PLUID   lpLuid
'    );
'@功能
'    LookupPrivilegeValue函数检索指定系统上使用的本地惟一标识符(LUID)，用于本地表示指定的特权名称。
'@参数
'lpSystemName _In_opt_
'   指向以null结尾的字符串的指针，该字符串指定检索特权名称的系统的名称。如果指定了空字符串，该函数将尝试在本地系统上查找特权名称。
'lpName  _In_
'   指向以null结尾的字符串的指针，该字符串指定特权的名称，如Winnt.h头文件中定义的那样。例如，这个参数可以指定常量SE_SECURITY_NAME，或者它对应的字符串“SeSecurityPrivilege”。
'lpLuid  _Out_
'   一个指向变量的指针，该变量接收LUID, lpSystemName参数指定的系统上可以通过LUID知道特权。
'@返回值
'   如果函数成功，则函数返回非零。
'   如果函数失败，它返回零。要获取扩展的错误信息，请调用GetLastError。
'@备注
'LookupPrivilegeValue函数只支持Winnt.h中定义的特权部分中指定的特权。有关值列表，请参见特权常量。
'@Requirements
'Minimum supported client           Windows XP [desktop apps | UWP apps]
'Minimum supported server           Windows Server 2003 [desktop apps | UWP apps]
'Header                             Winbase.h (include Windows.h)
'Library                            advapi32.lib
'dll                                advapi32.dll
'Unicode and ANSI names             LookupPrivilegeValueW  (Unicode) And LookupPrivilegeValueA(ANSI)
Private Declare Function AdjustTokenPrivileges Lib "advapi32.dll" (ByVal TokenHandle As Long, ByVal DisableAllPrivileges As Long, NewState As TOKEN_PRIVILEGES, ByVal BufferLength As Long, PreviousState As TOKEN_PRIVILEGES, ReturnLength As Long) As Long
'@原型
'    BOOL WINAPI AdjustTokenPrivileges(
'      _In_      HANDLE            TokenHandle,
'      _In_      BOOL              DisableAllPrivileges,
'      _In_opt_  PTOKEN_PRIVILEGES NewState,
'      _In_      DWORD             BufferLength,
'      _Out_opt_ PTOKEN_PRIVILEGES PreviousState,
'      _Out_opt_ PDWORD            ReturnLength
'    );
'@功能
'    启用或禁用指定访问令牌中的特权。在访问令牌中启用或禁用特权需要token_调节_privileges访问。
'@参数
'TokenHandle _In_
'   访问令牌的句柄，其中包含要修改的特权。句柄必须具有对令牌的token__privileges访问权。如果PreviousState参数不是NULL，句柄还必须具有TOKEN_QUERY访问权。
'DisableAllPrivileges _In_
'   指定函数是否禁用令牌的所有特权。如果该值为真，函数将禁用所有特权并忽略NewState参数。如果为FALSE，则函数根据NewState参数指向的信息修改特权。
'NewState   _In_opt_
'   指向TOKEN_PRIVILEGES结构的指针，该结构指定权限数组及其属性。如果DisableAllPrivileges参数为FALSE，调整tokenprivileges函数将启用、禁用或删除令牌的这些特权。下表描述了调整tokenprivileges函数基于特权属性所采取的操作。
'   SE_PRIVILEGE_ENABLED
'       该函数启用特权?
'   SE_PRIVILEGE_REMOVED
'       从令牌中的特权列表中删除特权?列表中的其他特权被重新排序以保持连续?
'   SE_PRIVILEGE_REMOVED取代SE_PRIVILEGE_ENABLED?
'       因为特权已经从令牌中删除，所以重新启用特权的尝试会导致警告ERROR_NOT_ALL_ASSIGNED，就好像特权从未存在过一样。
'       试图删除令牌中不存在的特权会返回ERROR_NOT_ALL_ASSIGNED?
'       特权检查删除的特权将导致STATUS_PRIVILEGE_NOT_HELD?特权检查审核失败将正常发生?
'       删除特权是不可逆的，因此在调用AdjustTokenPrivileges之后，已删除特权的名称不包含在PreviousState参数中。
'       带有SP1的Windows XP: 该函数不能删除特权?不支持此值?
'   None    该函数禁用特权?
'   如果DisableAllPrivileges为真，则函数将忽略此参数。
'BufferLength _In_
'   指定由PreviousState参数指向的缓冲区的大小(以字节为单位)。如果前状态参数为NULL，则此参数可以为零。
'PreviousState _Out_opt_
'   指向缓冲区的指针，该缓冲区由TOKEN_PRIVILEGES结构填充，该结构包含该函数修改的任何特权的前一状态。也就是说，如果一个特权已经被这个函数修改过，那么特权和它之前的状态就包含在由PreviousState引用的TOKEN_PRIVILEGES结构中。如果TOKEN_PRIVILEGES中的ecount成员为零，则此函数不会更改任何特权。这个参数可以是NULL。
'   如果指定的缓冲区太小，无法接收完整的已修改特权列表，则函数将失败，并且不调整任何特权。在本例中，函数将ReturnLength参数指向的变量设置为持有完整的修改特权列表所需的字节数。
'ReturnLength _Out_opt_
'   一个指向变量的指针，该变量接收到PreviousState参数所指向的缓冲区的所需大小(以字节为单位)。如果前状态为空，则此参数可以为空。
'@返回值
'   如果函数成功，返回值为非零。要确定函数是否调整了所有指定的特权，请调用GetLastError，当函数成功时，它将返回以下值之一:
'   返回代码描述
'ERROR_SUCCESS
'   该函数调整了所有指定的特权?
Private Const ERROR_NOT_ALL_ASSIGNED            As Long = &H514
'   令牌没有NewState参数中指定的一个或多个特权。即使没有调整特权，函数也可能成功地使用这个错误值。PreviousState参数表示已调整的特权。
'    如果函数失败，返回值为零。要获取扩展的错误信息，请调用GetLastError。
'@备注
'   AdjustTokenPrivileges函数不能向访问令牌添加新权限。它只能启用或禁用令牌的现有特权。要确定令牌的特权，请调用GetTokenInformation函数。
'   NewState参数可以指定令牌没有的特权，而不会导致函数失败。在本例中，函数调整令牌确实具有的特权，并忽略其他特权，以便函数成功。调用GetLastError函数来确定该函数是否调整了所有指定的特权。PreviousState参数表示已调整的特权。
'   PreviousState参数检索一个TOKEN_PRIVILEGES结构，该结构包含调整后的权限的原始状态。要恢复原始状态，在后续调用AdjustTokenPrivileges函数时，将PreviousState指针作为NewState参数传递。
'@Requirements
'Minimum supported client           Windows XP [desktop apps | UWP apps]
'Minimum supported server           Windows Server 2003 [desktop apps | UWP apps]
'Header                             Winbase.h (include Windows.h)
'Library                            advapi32.lib
'dll                                advapi32.dll
Private Declare Function AllocateAndInitializeSid Lib "advapi32.dll" (pIdentifierAuthority As SID_IDENTIFIER_AUTHORITY, ByVal nSubAuthorityCount As Byte, ByVal dwSubAuthority0 As Long, ByVal dwSubAuthority1 As Long, _
                            ByVal dwSubAuthority2 As Long, ByVal dwSubAuthority3 As Long, ByVal dwSubAuthority4 As Long, ByVal dwSubAuthority5 As Long, ByVal dwSubAuthority6 As Long, ByVal dwSubAuthority7 As Long, lpPSid As Long) As Long
'@原型
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
'@功能
'    AllocateAndInitializeSid函数使用最多八个子权限分配和初始化安全标识符(SID)。
'@参数
'pIdentifierAuthority _In_
'   指向SID_IDENTIFIER_AUTHORITY结构的指针。此结构提供要在SID中设置的顶级标识符权限值。
'nSubAuthorityCount _In_
'   指定要放置在SID中的子权限的数量。此参数还标识有多少子权限参数具有有意义的值。这个参数必须包含一个从1到8的值。
'   例如，值3表示由dwSubAuthority0。dwSubAuthority1和dwSubAuthority2参数指定的子权限值具有有意义的值，忽略其余值。
'dwSubAuthority0 _In_
'   要放置在SID中的子权限值
'dwSubAuthority1 _In_
'   要放置在SID中的子权限值
'dwSubAuthority2 _In_
'   要放置在SID中的子权限值
'dwSubAuthority3 _In_
'   要放置在SID中的子权限值
'dwSubAuthority4 _In_
'   要放置在SID中的子权限值
'dwSubAuthority5 _In_
'   要放置在SID中的子权限值
'dwSubAuthority6 _In_
'   要放置在SID中的子权限值
'dwSubAuthority7 _In_
'   要放置在SID中的子权限值
'pSid _Out_
'   一个指向变量的指针，该变量接收指向已分配和初始化SID结构的指针。
'@返回值
'   如果函数成功，返回值为非零。
'   如果函数失败，返回值为零。要获取扩展的错误信息，请调用GetLastError。
'@备注
'   必须使用FreeSid函数释放分配给AllocateAndInitializeSid函数的SID。
'   该函数创建一个具有32位RID值的SID。对于需要更长的RID值的应用程序，可以使用CreateWellKnownSid。
'@Requirements
'Minimum supported client        Windows XP [desktop apps | UWP apps]
'Minimum supported server        Windows Server 2003 [desktop apps | UWP apps]
'Header                          Winbase.h (include Windows.h)
'Library                         Advapi32.lib
'dll                             Advapi32.dll
Private Declare Function IsValidSid Lib "advapi32.dll" (ByVal pSid As Long) As Long
'@原型
'    BOOL WINAPI IsValidSid(
'      _In_ PSID pSid
'    );
'@功能
'    IsValidSid函数验证安全标识符(SID)，方法是验证修订号在已知范围内，并且子权限的数量小于最大值。
'@参数
'pSid  _In_
'   指向要验证的SID结构的指针。此参数不能为空。
'@返回值
'   如果SID结构有效，则返回值为非零。
'   如果SID结构无效，则返回值为零。该函数没有扩展的错误信息;不要调用GetLastError。
'@备注
'   如果pSid为空，则应用程序将失败，并出现访问冲突。
'@Requirements
'Minimum supported client        Windows XP [desktop apps | UWP apps]
'Minimum supported server        Windows Server 2003 [desktop apps | UWP apps]
'Header                          Winbase.h (include Windows.h)
'Library                         Advapi32.lib
'dll                             Advapi32.dll
Private Declare Function EqualSid Lib "advapi32.dll" (ByVal pSid1 As Long, ByVal pSid2 As Long) As Long
'@原型
'    BOOL WINAPI EqualSid(
'      _In_ PSID pSid1,
'      _In_ PSID pSid2
'    );
'@功能
'    EqualSid函数测试两个安全标识符(SID)值是否相等。两个SID必须完全匹配才能被认为是平等的。
'@参数
'pSid1 _In_
'   指向要比较的第一个SID结构的指针。这个结构被认为是有效的。
'pSid2 [在]
'   指向要比较的第二个SID结构的指针。这个结构被认为是有效的。
'@返回值
'   如果SID结构相等，则返回值为非零。
'   如果SID结构不相等，则返回值为零。要获取扩展的错误信息，请调用GetLastError。
'   如果任一SID结构无效，则返回值未定义。
'@Requirements
'Minimum supported client        Windows XP [desktop apps | UWP apps]
'Minimum supported server        Windows Server 2003 [desktop apps | UWP apps]
'Header                          Winbase.h (include Windows.h)
'Library                         Advapi32.lib
'dll                             Advapi32.dll
Private Declare Function FreeSid Lib "advapi32.dll" (ByVal pSid As Long) As Long
'@原型
'    PVOID WINAPI FreeSid(
'      _In_ PSID pSid
'    );
'@参数
'pSid  _In_
'   指向SID结构的指针释放。
'@返回值
'   如果函数成功，则返回NULL。
'   如果函数失败，它将返回一个指向由pSid参数表示的SID结构的指针。
'常用SIDs
'   众所周知的安全标识符(SIDs)标识通用组和通用用户。例如，一些常用的SID查明了下列群体和用户:
'   每个人或世界，这是一个包含所有用户的组。
'   CREATOR_OWNER，它用作可继承ACE中的占位符。当ACE被继承时，系统用对象创建者的SID替换CREATOR_OWNER SID。
'   本地计算机上内置域的Administrators组。
'有一些众所周知的SID，它们对所有使用这种安全模式的安全系统，包括Windows以外的操作系统都是有意义的。此外，有一些著名SID只对Windows系统有意义。
'Windows API为已知的标识符权限和相对标识符(RID)值定义了一组常量。您可以使用这些常量来创建著名的小岛屿发展中国家。下面的例子结合了SECURITY_WORLD_SID_AUTHORITY和SECURITY_WORLD_RID常量，显示了代表所有用户(所有人或世界)的特殊组的通用知名SID:
'S -1 - 1 - 0
'本例使用SID的字符串表示法，其中S将字符串标识为SID，第一个1是SID的修订级别，其余两个数字是SECURITY_WORLD_SID_AUTHORITY和SECURITY_WORLD_RID常量。
'您可以使用AllocateAndInitializeSid函数来构建SID，方法是将标识符权限值与最多八个子权限值组合起来。例如，要确定登录的用户是否是某个特定知名组的成员，请调用AllocateAndInitializeSid为该知名组构建一个SID，并使用EqualSid函数将该SID与用户访问令牌中的组SID进行比较。有关示例，请参见在c++中的访问令牌中搜索SID。您必须调用FreeSid函数来释放由AllocateAndInitializeSid分配的SID。
'本节的其余部分包括众所周知SID的表格，以及可以用来建立众所周知的SID的标识符权力和子权力常数的表格。
'以下是一些众所周知的SID
'通用的著名sid字符串值标识
'Null SID           S -1 - 0 - 0        没有成员的团体。这通常在SID值未知时使用。
'World              S - 1 - 1 - 0       包含所有用户的组。
'Local              S -1 - 2 - 0        登录到本地(物理上)连接到系统的终端的用户。
'Creator Owner ID   S -1 - 3 - 0        要由创建新对象的用户的安全标识符替换的安全标识符。此SID用于可继承ace。
'Creator Group ID   S -1 - 3 - 1        要由创建新对象的用户的主组SID替换的安全标识符。在可继承ace中使用此SID
'下表列出了预定义的标识符权限常量。前四个值用于众所周知的SID;最后一个值用于Windows中著名的SID。
'标识符权限值sid字符串前缀
'SECURITY_NULL_SID_AUTHORITY        0       S -1 - 0
'SECURITY_WORLD_SID_AUTHORITY       1       S -1 - 1
'SECURITY_LOCAL_SID_AUTHORITY       2       S -1 - 2
'SECURITY_CREATOR_SID_AUTHORITY     3       S -1 - 3
'SECURITY_NT_AUTHORITY              5       S -1 - 5
'下列RID值用于众所周知的SID。标识符权限列显示标识符权限的前缀，您可以将RID与之组合以创建一个通用的众所周知的SID。
'相对标识符权限值标识符权限
'SECURITY_NULL_RID                  0       S -1 - 0
'SECURITY_WORLD_RID                 0       S -1 - 1
'SECURITY_LOCAL_RID                 0       S -1 - 2
'SECURITY_LOCAL_LOGON_RID           1       S -1 - 2
'SECURITY_CREATOR_OWNER_RID         0       S -1 - 3
'SECURITY_CREATOR_GROUP_RID         1       S -1 - 3
'SECURITY_NT_AUTHORITY (S-1-5)预定义的标识符权限生成的sid不是通用的，但仅在Windows安装上有意义。您可以使用以下带有SECURITY_NT_AUTHORITY的RID值来创建知名的sid。
'常量字符串值标识
Private Const SECURITY_DIALUP_RID               As Long = &H1
'   SECURITY_DIALUP_RID        S -1 - 5 - 1        使用拨号调制解调器登录终端的用户。这是一个组标识符。
Private Const SECURITY_NETWORK_RID              As Long = &H2
'   SECURITY_NETWORK_RID       S -1 - 5 - 2        通过网络登录的用户。这是一个组标识符，添加到跨网络登录进程的令牌中。对应的登录类型是LOGON32_LOGON_NETWORK。
Private Const SECURITY_BATCH_RID                As Long = &H3
'   SECURITY_BATCH_RID         S -1 - 5 - 3        使用批处理队列功能登录的用户。这是一个组标识符，添加到作为批处理作业记录的进程的令牌中。对应的登录类型是LOGON32_LOGON_BATCH。
Private Const SECURITY_INTERACTIVE_RID          As Long = &H4
'   SECURITY_INTERACTIVE_RID   S -1 - 5 - 4        用户登录进行交互操作。这是一个组标识符，当进程以交互方式登录时添加到进程的令牌中。对应的登录类型是LOGON32_LOGON_INTERACTIVE。
Private Const SECURITY_LOGON_IDS_RID            As Long = &H5
'   SECURITY_LOGON_IDS_RID     S -1 - 5 - 5 - X - y    一个登录会话。这用于确保只有给定登录会话中的进程才能访问该会话的window-station对象。对于每个登录会话，这些SID的X和Y值是不同的。SECURITY_LOGON_IDS_RID_COUNT值是这个标识符(5-X-Y)中的rid数量。
Private Const SECURITY_SERVICE_RID              As Long = &H6
'   SECURITY_SERVICE_RID       S -1 - 5 - 6        授权作为服务登录的帐户。这是一个组标识符，添加到作为服务记录的进程的令牌中。对应的登录类型是LOGON32_LOGON_SERVICE。
Private Const SECURITY_ANONYMOUS_LOGON_RID      As Long = &H7
'   SECURITY_ANONYMOUS_LOGON_RID   S -1 - 5 - 7    匿名登录，或空会话登录。
Private Const SECURITY_PROXY_RID                As Long = &H8
'   SECURITY_PROXY_RID         S -1 - 5 - 8        代理。
Private Const SECURITY_ENTERPRISE_CONTROLLERS_RID   As Long = &H9
'   SECURITY_ENTERPRISE_CONTROLLERS_RID    S -1 - 5 - 9    企业控制器。
Private Const SECURITY_PRINCIPAL_SELF_RID       As Long = &HA
'   SECURITY_PRINCIPAL_SELF_RID    S -1 - 5 - 10       PRINCIPAL_SELF安全标识符可以在用户或组对象的ACL中使用。在访问检查期间，系统用对象的SID替换SID。PRINCIPAL_SELF SID用于指定可继承的ACE，该ACE应用于继承ACE的用户或组对象。它是在模式的默认安全描述符中表示已创建对象的SID的唯一方法。
Private Const SECURITY_AUTHENTICATED_USER_RID   As Long = &HB
'   SECURITY_AUTHENTICATED_USER_RID    S -1 - 5 - 11   通过身份验证的用户。
Private Const SECURITY_RESTRICTED_CODE_RID      As Long = &HC
'   SECURITY_RESTRICTED_CODE_RID   S -1 - 5 - 12       受限制的代码。
Private Const SECURITY_TERMINAL_SERVER_RID      As Long = &HD
'   SECURITY_TERMINAL_SERVER_RID   S -1 - 5 - 13       终端服务。自动添加到登录到终端服务器的用户的安全令牌中。
Private Const SECURITY_LOCAL_SYSTEM_RID         As Long = &H12
'   SECURITY_LOCAL_SYSTEM_RID      S -1 - 5 - 18       操作系统使用的特殊帐户。
Private Const SECURITY_NT_NON_UNIQUE            As Long = &H15
'   SECURITY_NT_NON_UNIQUE         S -1 - 5 - 21       SID并非独一无二。
Private Const SECURITY_BUILTIN_DOMAIN_RID       As Long = &H20
'   SECURITY_BUILTIN_DOMAIN_RID    S -1 - 5 - 32       内置的系统域。
'以下rid与每个域相关。
'处理标识
'DOMAIN_ALIAS_RID_CERTSVC_DCOM_ACCESS_GROUP     可以使用分布式组件对象模型(DCOM)连接到认证机构的用户组。
'DOMAIN_USER_RID_ADMIN                          域中的管理用户帐户。
'DOMAIN_USER_RID_GUEST                          域中的来宾用户帐户。没有帐户的用户可以自动登录该帐户。
'DOMAIN_GROUP_RID_ADMINS                        域管理员组。此帐户仅存在于运行服务器操作系统的系统上。
'DOMAIN_GROUP_RID_USERS                         一个域中包含所有用户帐户的组。所有用户都会自动添加到这个组中。
'DOMAIN_GROUP_RID_GUESTS                        域中的来宾组帐户。
'DOMAIN_GROUP_RID_COMPUTERS                     域计算机组。域中的所有计算机都是这个组的成员。
'DOMAIN_GROUP_RID_CONTROLLERS                   域控制器的组。域中的所有DCs都是这个组的成员。
'DOMAIN_GROUP_RID_CERT_ADMINS                   证书发布者组。运行证书服务的计算机是这个组的成员。
'DOMAIN_GROUP_RID_ENTERPRISE_READONLY_DOMAIN_CONTROLLERS    一组企业只读域控制器。
'DOMAIN_GROUP_RID_SCHEMA_ADMINS                 模式管理员组。这个组的成员可以修改Active Directory模式。
'DOMAIN_GROUP_RID_ENTERPRISE_ADMINS             企业管理员组。这个组的成员可以完全访问Active Directory林中的所有域。企业管理员负责群体的操作，例如添加或删除新域。
'DOMAIN_GROUP_RID_POLICY_ADMINS                 策略管理员组。
'DOMAIN_GROUP_RID_READONLY_CONTROLLERS          只读域控制器组。
'以下rid用于指定强制完整性级别。
Private Const SECURITY_MANDATORY_UNTRUSTED_RID  As Long = &H0
'   不可信的
Private Const SECURITY_MANDATORY_LOW_RID        As Long = &H1000
'   低的完整性
Private Const SECURITY_MANDATORY_MEDIUM_RID     As Long = &H2000
'   媒介的完整性
Private Const SECURITY_MANDATORY_MEDIUM_PLUS_RID As Long = SECURITY_MANDATORY_MEDIUM_RID + &H100
'   中等高度的完整性
Private Const SECURITY_MANDATORY_HIGH_RID       As Long = &H3000
'   高完整性
Private Const SECURITY_MANDATORY_SYSTEM_RID     As Long = &H4000
'   系统的完整性
Private Const SECURITY_MANDATORY_PROTECTED_PROCESS_RID As Long = &H5000
'   受保护的过程
'下表中有一些域相关rid的示例，您可以使用它们为本地组(别名)形成众所周知的SID。有关本地和全局组的更多信息，请参见本地组函数和组函数。
Private Const DOMAIN_ALIAS_RID_ADMINS           As Long = &H220
'   用于域管理的本地组。
Private Const DOMAIN_ALIAS_RID_USERS            As Long = &H221
'   表示域中所有用户的本地组。
Private Const DOMAIN_ALIAS_RID_GUESTS           As Long = &H222
'   表示域的来宾的本地组。
Private Const DOMAIN_ALIAS_RID_POWER_USERS      As Long = &H223
'   一个本地组，用于表示一个或一组用户，这些用户希望将一个系统视为他们的个人计算机，而不是多个用户的工作站。
Private Const DOMAIN_ALIAS_RID_ACCOUNT_OPS      As Long = &H224
'   仅存在于运行服务器操作系统的系统上的本地组。这个本地组允许控制非管理员帐户。
Private Const DOMAIN_ALIAS_RID_SYSTEM_OPS       As Long = &H225
'   仅存在于运行服务器操作系统的系统上的本地组。这个本地组执行系统管理功能，不包括安全功能。它建立网络共享。控制打印机。解锁工作站和执行其他操作。
Private Const DOMAIN_ALIAS_RID_PRINT_OPS        As Long = &H226
'   仅存在于运行服务器操作系统的系统上的本地组。这个本地组控制打印机和打印队列。
Private Const DOMAIN_ALIAS_RID_BACKUP_OPS       As Long = &H227
'   用于控制文件备份和恢复特权分配的本地组。
Private Const DOMAIN_ALIAS_RID_REPLICATOR       As Long = &H228
'   负责将安全数据库从主域控制器复制到备份域控制器的本地组。这些帐户仅供系统使用。
Private Const DOMAIN_ALIAS_RID_RAS_SERVERS      As Long = &H229
'   表示RAS和IAS服务器的本地组。这个组允许访问用户对象的各种属性。
Private Const DOMAIN_ALIAS_RID_PREW2KCOMPACCESS As Long = &H22A
'   仅存在于运行Windows 2000服务器的系统上的本地组。有关更多信息，请参见允许匿名访问。
Private Const DOMAIN_ALIAS_RID_REMOTE_DESKTOP_USERS As Long = &H22B
'   表示所有远程桌面用户的本地组。
Private Const DOMAIN_ALIAS_RID_NETWORK_CONFIGURATION_OPS    As Long = &H22C
'   表示网络配置的本地组。
Private Const DOMAIN_ALIAS_RID_INCOMING_FOREST_TRUST_BUILDERS    As Long = &H22D
'   表示任何forest trust用户的本地组。
Private Const DOMAIN_ALIAS_RID_MONITORING_USERS As Long = &H22E
'   表示被监视的所有用户的本地组。
Private Const DOMAIN_ALIAS_RID_LOGGING_USERS    As Long = &H22F
'   负责记录用户日志的本地组。
Private Const DOMAIN_ALIAS_RID_AUTHORIZATIONACCESS  As Long = &H230
'   表示所有授权访问的本地组。
Private Const DOMAIN_ALIAS_RID_TS_LICENSE_SERVERS    As Long = &H231
'   仅存在于运行服务器操作系统(允许终端服务和远程访问)的系统上的本地组。
Private Const DOMAIN_ALIAS_RID_DCOM_USERS       As Long = &H232
'   表示可以使用分布式组件对象模型(DCOM)的用户的本地组。
Private Const DOMAIN_ALIAS_RID_IUSERS           As Long = &H238
'   代表Internet用户的本地组。
Private Const DOMAIN_ALIAS_RID_CRYPTO_OPERATORS     As Long = &H239
'   表示对密码操作符的访问的本地组。
Private Const DOMAIN_ALIAS_RID_CACHEABLE_PRINCIPALS_GROUP   As Long = &H23B
'   表示可以缓存的主体的本地组。
Private Const DOMAIN_ALIAS_RID_NON_CACHEABLE_PRINCIPALS_GROUP   As Long = &H23C
'   表示不能缓存的主体的本地组。
Private Const DOMAIN_ALIAS_RID_EVENT_LOG_READERS_GROUP    As Long = &H23D
'   表示事件日志读取器的本地组。
Private Const DOMAIN_ALIAS_RID_CERTSVC_DCOM_ACCESS_GROUP    As Long = &H23E
'   可以使用分布式组件对象模型(DCOM)连接到认证机构的本地用户组。
Private Declare Function GetCurrentThread Lib "kernel32.dll" () As Long
'@原型
'    HANDLE WINAPI GetCurrentThread(void);
'@功能
'    检索调用线程的伪句柄。
'@参数
'   这个函数没有参数。
'@返回值
'   返回值是当前线程的伪句柄。
'@备注
'    伪句柄是一个特殊的常量，它被解释为当前线程句柄。无论何时需要线程句柄，调用线程都可以使用这个句柄来指定自己。子进程不会继承伪句柄。
'    这个句柄具有对thread对象的THREAD_ALL_ACCESS访问权。有关更多信息，请参见线程安全和访问权限。
'    Windows Server 2003和Windows XP:这个句柄具有线程的安全描述符所允许的对进程的主令牌的最大访问权。
'    一个线程不能使用该函数创建一个句柄，其他线程可以使用该句柄引用第一个线程。句柄总是被解释为引用使用它的线程。通过在调用DuplicateHandle函数时将伪句柄指定为源句柄，线程可以为自己创建一个“真实”句柄，其他线程可以使用该句柄，或由其他进程继承。
'    当不再需要伪句柄时，不需要关闭它。使用此句柄调用close句柄函数没有效果。如果用DuplicateHandle复制伪句柄，则必须关闭重复句柄。
'    模拟安全上下文时不要创建线程。调用将成功，但是新创建的线程在调用GetCurrentThread时将减少对自身的访问权限。授予此线程的访问权限将从模拟用户对进程的访问权限派生。一些访问权限(包括THREAD_SET_THREAD_TOKEN和THREAD_GET_CONTEXT)可能不存在，从而导致意外的失败。
'@Requirements
'Minimum supported client        Windows XP [desktop apps | UWP apps]
'Minimum supported server        Windows Server 2003 [desktop apps | UWP apps]
'Minimum supported phone         Windows Phone 8
'Header                          WinBase.h on Windows XP, Windows Server 2003, Windows Vista, Windows 7, Windows Server 2008 and Windows Server 2008 R2 (include Windows.h);Processthreadsapi.h on Windows 8 and Windows Server 2012
'Library                         kernel32.lib
'dll                             kernel32.dll

'标准权限
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
'Token权限
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
'@原型
'    HANDLE WINAPI GetCurrentProcess(void);
'@功能
'    检索当前进程的伪句柄。
'@参数
'   这个函数没有参数。
'@返回值
'   返回值是当前线程的伪句柄。
'@返回值
'   返回值是当前进程的伪句柄。
'@备注
'    伪句柄是一个特殊的常量，当前(句柄)-1，它被解释为当前进程句柄。为了与未来的操作系统兼容，最好调用GetCurrentProcess，而不是硬编码这个常量。无论何时需要流程句柄，调用流程都可以使用伪句柄来指定自己的流程。子进程不会继承伪句柄。
'    此句柄具有PROCESS_ALL_ACCESS访问流程对象的权限。有关更多信息，请参见流程安全和访问权限。
'    Windows Server 2003和Windows XP:此句柄具有进程安全描述符允许的对进程主令牌的最大访问权。
'    通过在调用DuplicateHandle函数时将伪句柄指定为源句柄，流程可以为自己创建一个“真实”句柄，该句柄在其他流程的上下文中是有效的，或者可以被其他流程继承。进程还可以使用OpenProcess函数为自己打开一个实句柄。
'    当不再需要伪句柄时，不需要关闭它。使用伪句柄调用close句柄函数没有效果。如果用DuplicateHandle复制伪句柄，则必须关闭重复句柄。
'@Requirements
'Minimum supported client        Windows XP [desktop apps | UWP apps]
'Minimum supported server        Windows Server 2003 [desktop apps | UWP apps]
'Minimum supported phone         Windows Phone 8
'Header                          WinBase.h on Windows XP, Windows Server 2003, Windows Vista, Windows 7, Windows Server 2008 and Windows Server 2008 R2 (include Windows.h);Processthreadsapi.h on Windows 8 and Windows Server 2012
'Library                         kernel32.lib
'dll                             kernel32.dll
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
'@原型
'    HINSTANCE ShellExecute(
'      _In_opt_ HWND    hwnd,
'      _In_opt_ LPCTSTR lpOperation,
'      _In_     LPCTSTR lpFile,
'      _In_opt_ LPCTSTR lpParameters,
'      _In_opt_ LPCTSTR lpDirectory,
'      _In_     INT     nShowCmd
'    );
'@功能
'    对指定的文件执行操作。
'@参数
'hwnd _In_opt_
'   类型: hwnd
'   父窗口的句柄，用于显示UI或错误消息。如果操作不与窗口关联，则此值可以为空。
'lpOperation _In_opt_
'   类型: LPCTSTR
'   指向以null结尾的字符串的指针，在本例中称为谓词，它指定要执行的操作。可用谓词集取决于特定的文件或文件夹。通常，对象的快捷菜单中可用的操作都是可用的谓词。常用的动词有:
'   edit    启动编辑器并打开文档进行编辑。如果lpFile不是文档文件，函数将失败。
'   explore 浏览lpFile指定的文件夹。
'   find    在lpDirectory指定的目录中启动搜索。
'   open    打开由lpFile参数指定的项。项目可以是文件或文件夹。
'   print   打印lpFile指定的文件。如果lpFile不是文档文件，则该函数将失败。
'   NULL    如果可用，则使用默认谓词。如果没有，则使用“open”动词。如果两个动词都不可用，系统将使用注册表中列出的第一个谓词。
'lpFile _In_
'   类型: LPCTSTR
'   指向以null结尾的字符串的指针，该字符串指定要在其上执行指定谓词的文件或对象。要指定Shell名称空间对象，请传递完全限定的解析名。注意，并非所有对象都支持所有谓词。例如，并非所有文档类型都支持“print”谓词。如果对lpDirectory参数使用相对路径，则不要对lpFile使用相对路径。
'lpParameters(,可选)
'   类型: LPCTSTR
'   如果lpFile指定了一个可执行文件，这个参数就是指向一个以null结尾的字符串的指针，该字符串指定要传递给应用程序的参数。此字符串的格式由要调用的谓词决定。如果lpFile指定了一个文档文件，那么lpParameters应该为NULL。
'lpDirectory _In_opt_
'   类型: LPCTSTR
'   指向以null结尾的字符串的指针，该字符串指定操作的默认(工作)目录。如果该值为空，则使用当前工作目录。如果在lpFile中提供了相对路径，则不要为lpDirectory使用相对路径。
'nShowCmd    _In_
'   类型:INT
'   指定应用程序打开时如何显示的标志。如果lpFile指定了一个文档文件，则该标志将被简单地传递给相关的应用程序。如何处理它取决于应用程序。这些值在Winuser.h中定义。
Private Const SW_HIDE                           As Long = &H0
'   隐藏窗口并激活另一个窗口。
Private Const SW_MAXIMIZE                       As Long = &H3
'   最大化指定的窗口。
Private Const SW_MINIMIZE                       As Long = &H6
'   最小化指定的窗口并按z顺序激活下一个顶级窗口。
Private Const SW_RESTORE                        As Long = &H9
'   激活并显示窗口。如果窗口被最小化或最大化，窗口会将其恢复到原来的大小和位置。应用程序应该在恢复最小化窗口时指定此标志。
Private Const SW_SHOW                           As Long = &H5
'   激活窗口并显示其当前大小和位置。
Private Const SW_SHOWDEFAULT                    As Long = &HA
'   根据启动应用程序的程序传递给CreateProcess函数的STARTUPINFO结构中指定的SW_标志设置显示状态。应用程序应该使用此标志调用ShowWindow来设置其主窗口的初始显示状态。
Private Const SW_SHOWMAXIMIZED                  As Long = &H3
'   激活窗口并将其显示为最大化窗口。
Private Const SW_SHOWMINIMIZED                  As Long = &H2
'   激活窗口并将其显示为最小化窗口。
Private Const SW_SHOWMINNOACTIVE                As Long = &H7
'   将窗口显示为最小化窗口。活动窗口保持活动状态。
Private Const SW_SHOWNA                         As Long = &H8
'   显示窗口的当前状态。活动窗口保持活动状态。
Private Const SW_SHOWNOACTIVATE                 As Long = &H4
'   显示窗口的最新大小和位置。活动窗口保持活动状态。
Private Const SW_SHOWNORMAL                     As Long = &H1
'   激活并显示窗口。如果窗口被最小化或最大化，窗口会将其恢复到原来的大小和位置。应用程序应该在第一次显示窗口时指定此标志。
'@返回值
'   类型: HINSTANCE
'   如果函数成功，则返回一个大于32的值。如果函数失败，它将返回一个错误值，该值指示失败的原因。返回值被转换为一个HINSTANCE，以便与16位Windows应用程序向后兼容。然而，这不是一个真正的HINSTANCE。它只能被强制转换为整数，并与32或下面的错误代码进行比较。
'   返回代码描述
'   0           操作系统内存或资源不足。
Private Const ERROR_FILE_NOT_FOUND              As Long = &H2
'   没有找到指定的文件。
Private Const ERROR_PATH_NOT_FOUND              As Long = &H3
'   没有找到指定的路径。
Private Const ERROR_BAD_FORMAT                  As Long = &HB
'   .exe文件无效(非win32 .exe或.exe映像中的错误)。
Private Const SE_ERR_ACCESSDENIED               As Long = &H5
'   操作系统拒绝访问指定的文件。
Private Const SE_ERR_ASSOCINCOMPLETE            As Long = &H1B
'   文件名关联不完整或无效。
Private Const SE_ERR_DDEBUSY                    As Long = &H1E
'   由于正在处理其他DDE事务，因此无法完成DDE事务。
Private Const SE_ERR_DDEFAIL                    As Long = &H1D
'   DDE事务失败。
Private Const SE_ERR_DDETIMEOUT                 As Long = &H1C
'   由于请求超时，无法完成DDE事务。
Private Const SE_ERR_DLLNOTFOUND                As Long = &H20
'   没有找到指定的DLL。
Private Const SE_ERR_FNF                        As Long = &H2
'   没有找到指定的文件。
Private Const SE_ERR_NOASSOC                    As Long = &H1F
'   没有与给定文件扩展名关联的应用程序。如果尝试打印不可打印的文件，也将返回此错误。
Private Const SE_ERR_OOM                        As Long = &H8
'   没有足够的内存来完成这项操作。
Private Const SE_ERR_PNF                        As Long = &H3
'   没有找到指定的路径。
Private Const SE_ERR_SHARE                      As Long = &H1A
'   发生了共享冲突。
'@备注
'   因为ShellExecute可以将执行委托给使用组件对象模型(COM)激活的Shell扩展(数据源。上下文菜单处理程序。谓词实现)，所以应该在调用ShellExecute之前初始化COM。一些Shell扩展需要COM单线程公寓(STA)类型。在这种情况下，COM应该如下所示初始化:
'   CoInitializeEx(NULL, coinit_apartmentthreads | COINIT_DISABLE_OLE1DDE)
'   当然，在某些实例中，ShellExecute不使用这些类型的Shell扩展，并且这些实例根本不需要初始化COM。尽管如此，在使用此函数之前，始终对COM进行初始化是一个很好的实践。
'   此方法允许您执行文件夹的快捷菜单或存储在注册表中的任何命令。
'   要打开文件夹，请使用以下调用之一:
'   ShellExecute(句柄，NULL， <fully_qualified_path_to_folder>， NULL, NULL, SW_SHOWNORMAL);
'   或
'   ShellExecute(句柄，“open”，<fully_qualified_path_to_folder>， NULL, NULL, SW_SHOWNORMAL);
'   要查看文件夹，请使用以下调用:
'   ShellExecute(句柄，“explore”，<fully_qualified_path_to_folder>， NULL, NULL, SW_SHOWNORMAL);
'   要启动Shell的目录查找实用程序，请使用以下调用。
'   ShellExecute(句柄，“find”，<fully_qualified_path_to_folder>， NULL, NULL, 0);
'   如果lpOperation为空，函数将打开lpFile指定的文件。如果lpOperation是“open”或“explore”，该函数将尝试打开或浏览文件夹。
'   要获取关于由于调用ShellExecute而启动的应用程序的信息，请使用ShellExecuteEx。
'   注意，启动文件夹窗口在文件夹选项中的单独进程设置中影响ShellExecute。如果禁用该选项(默认设置)，ShellExecute将使用一个打开的资源管理器窗口，而不是启动一个新的资源管理器窗口。如果没有打开资源管理器窗口，ShellExecute将启动一个新窗口。
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
'@功能
'    包含ShellExecuteEx使用的信息。
'@原型
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
'@成员
'cbSize
'   类型: 双字
'   必需的。此结构的大小，以字节为单位。
'fMask
'   类型: ULONG
'   表示其他结构成员的内容和有效性的标志;下列值的组合:
Private Const SEE_MASK_DEFAULT                  As Long = &H0
'   使用默认值。
Private Const SEE_MASK_CLASSNAME                As Long = &H1
'   使用lpClass成员给出的类名。如果设置了SEE_MASK_CLASSKEY和SEE_MASK_CLASSNAME，则使用类键。
Private Const SEE_MASK_CLASSKEY                 As Long = &H3
'   使用由hkeyClass成员提供的类密钥。如果设置了SEE_MASK_CLASSKEY和SEE_MASK_CLASSNAME，则使用类键。
Private Const SEE_MASK_IDLIST                   As Long = &H4
'   使用lpIDList成员提供的项标识符列表。lpIDList成员必须指向ITEMIDLIST结构。
Private Const SEE_MASK_INVOKEIDLIST             As Long = &HC
'   使用所选项目的快捷菜单处理程序的IContextMenu接口。使用lpFile根据其文件系统路径标识项，或者使用lpIDList根据其PIDL标识项。此标志允许应用程序使用ShellExecuteEx从快捷菜单扩展中调用谓词，而不是注册表中列出的静态谓词。
'   注意，SEE_MASK_INVOKEIDLIST覆盖并暗示了SEE_MASK_IDLIST。
Private Const SEE_MASK_ICON                     As Long = &H10
'   使用hIcon成员给出的图标。此标志不能与SEE_MASK_HMONITOR组合。
'   注意，此标志仅在Windows XP及更早版本中使用。它在Windows Vista中被忽略了。
Private Const SEE_MASK_HOTKEY                   As Long = &H20
'   使用dwHotKey成员提供的键盘快捷方式。
Private Const SEE_MASK_NOCLOSEPROCESS           As Long = &H40
'   用于指示hProcess成员接收进程句柄。此句柄通常用于允许应用程序查明使用ShellExecuteEx创建的进程何时终止。在某些情况下，例如通过DDE对话满足执行时，不会返回句柄。调用应用程序负责在不再需要句柄时关闭句柄。
Private Const SEE_MASK_CONNECTNETDRV            As Long = &H80
'   验证共享并连接到驱动器信。这允许重新连接断开连接的网络驱动器。lpFile成员是网络上文件的UNC路径。
Private Const SEE_MASK_NOASYNC                  As Long = &H100
'   等待执行操作完成后返回。这个标志应该由使用ShellExecute表单的调用者使用，这些调用者可能会导致异步激活，例如DDE，并创建一个可能在后台线程上运行的进程。(注意:如果调用者的线程模型不是单元，则默认情况下ShellExecuteEx在后台线程上运行。)从已经在后台线程上运行的进程调用ShellExecuteEx应该总是传递这个标志。此外，调用ShellExecuteEx后立即退出的应用程序应该指定此标志。
'   如果执行操作是在后台线程上执行的，并且调用者没有指定SEE_MASK_ASYNCOK标志，那么调用线程将等到新进程启动后才返回。这通常意味着调用了CreateProcess, DDE通信已经完成，或者自定义执行委托已经通知ShellExecuteEx它已经完成。如果指定了SEE_MASK_WAITFORINPUTIDLE标志，那么ShellExecuteEx将调用WaitForInputIdle，并等待新进程空闲，然后返回，最大超时为1分钟。
'   有关何时需要此标志的进一步讨论，请参见备注部分。
Private Const SEE_MASK_FLAG_DDEWAIT             As Long = &H100
'   不要使用;使用SEE_MASK_NOASYNC代替。
Private Const SEE_MASK_DOENVSUBST               As Long = &H200
'   展开lpDirectory或lpFile成员给出的字符串中指定的任何环境变量。
Private Const SEE_MASK_FLAG_NO_UI               As Long = &H400
'   如果发生错误，不要显示错误消息框。
Private Const SEE_MASK_UNICODE                  As Long = &H4000
'   使用此标志指示Unicode应用程序。
Private Const SEE_MASK_NO_CONSOLE               As Long = &H8000
'   用于继承父进程的控制台，而不是让它创建新控制台。它与在CreateProcess中使用CREATE_NEW_CONSOLE标志相反。
Private Const SEE_MASK_ASYNCOK                  As Long = &H100000
'   执行可以在后台线程上执行，调用应该立即返回，而不需要等待后台线程完成。注意，在某些情况下，ShellExecuteEx会忽略此标志，并等待进程完成后返回。
Private Const SEE_MASK_NOQUERYCLASSSTORE        As Long = &H1000000
'   不使用
Private Const SEE_MASK_HMONITOR                 As Long = &H200000
'   在多监视器系统上指定监视器时使用此标志。监视器在hMonitor成员中指定。此标志不能与SEE_MASK_ICON组合。
Private Const SEE_MASK_NOZONECHECKS             As Long = &H800000
'   在Windows XP中引入。不要执行区域检查。这个标志允许ShellExecuteEx绕过IAttachmentExecute放置的区域检查。
Private Const SEE_MASK_WAITFORINPUTIDLE         As Long = &H2000000
'   创建新进程之后，等待进程空闲，然后返回，超时一分钟。详情请参见WaitForInputIdle。
Private Const SEE_MASK_FLAG_LOG_USAGE           As Long = &H4000000
'   在Windows XP中引入。跟踪此应用程序已启动的次数。计数足够高的应用程序出现在开始菜单的最常用程序列表中。
Private Const SEE_MASK_FLAG_HINST_IS_SITE       As Long = &H8000000
'   hInstApp成员用于指定实现IServiceProvider的对象的IUnknown。此对象将用作站点指针。站点指针用于向ShellExecute函数。处理程序绑定过程和调用的谓词处理程序提供服务。
'   要在Windows 8之前的操作系统中使用SEE_MASK_FLAG_HINST_IS_SITE，请在程序中手动定义它:#define SEE_MASK_FLAG_HINST_IS_SITE 0x08000000。
'hwnd
'   类型: hwnd
'   可选的。父窗口的句柄，用于显示系统在执行此函数时可能产生的任何消息框。这个值可以为空。
'lpVerb
'   类型: LPCTSTR
'   一个字符串，作为动词，指定要执行的操作。可用谓词集取决于特定的文件或文件夹。通常，对象的快捷菜单中可用的操作都是可用的谓词。此参数可以为NULL，在这种情况下，如果可用，则使用默认谓词。如果没有，则使用“open”谓词。如果两个谓词都不可用，系统将使用注册表中列出的第一个谓词。常用的谓词有:
'   edit    启动编辑器并打开文档进行编辑。如果lpFile不是文档文件，函数将失败。
'   explore 浏览lpFile指定的文件夹。
'   find    在lpDirectory指定的目录中启动搜索。
'   open    打开由lpFile参数指定的项。项目可以是文件或文件夹。
'   print   打印lpFile指定的文件。如果lpFile不是文档文件，则该函数将失败。
'   properties    显示文件或文件夹的属性。
'lpFile
'   类型: LPCTSTR
'   以null结尾的字符串的地址，该字符串指定文件或对象的名称，ShellExecuteEx将在该文件或对象上执行lpVerb参数指定的操作。ShellExecuteEx函数支持的系统注册表谓词包括可执行文件和文档文件的“open”和已注册打印处理程序的文档文件的“print”。其他应用程序可能已经通过系统注册表添加了Shell谓词，例如.avi和.wav文件的“play”。要指定Shell名称空间对象，请传递完全限定的解析名，并在fMask参数中设置SEE_MASK_INVOKEIDLIST标志。
'   注意，如果设置了SEE_MASK_INVOKEIDLIST标志，则可以使用lpFile或lpIDList分别根据其文件系统路径或PIDL标识项。必须设置两个值之一——lpfile或lpidlist。
'   注意，如果路径没有包含在名称中，则假定当前目录。
'lpParameters
'   类型: LPCTSTR
'   可选的。包含应用程序参数的以null结尾的字符串的地址。参数必须用空格分隔。如果lpFile成员指定了一个文档文件，那么lpParameters应该为NULL。
'lpDirectory
'   类型: LPCTSTR
'   可选的。以null结尾的字符串的地址，该字符串指定工作目录的名称。如果该成员为空，则使用当前目录作为工作目录。
'nShow
'   类型:int
'   必需的。指定应用程序打开时显示方式的标志;ShellExecute函数列出的SW_值之一。如果lpFile指定了一个文档文件，则该标志将被简单地传递给相关的应用程序。如何处理它取决于应用程序。
'hInstApp
'   类型: 实例句柄
'   如果设置了SEE_MASK_NOCLOSEPROCESS，并且ShellExecuteEx调用成功，则将该成员设置为大于32的值。如果函数失败，则将其设置为SE_ERR_XXX错误值，该值指示失败的原因。尽管为了兼容16位Windows应用程序，hInstApp被声明为一个HINSTANCE，但它不是一个真正的HINSTANCE。它只能被转换为一个int类型，并与32或以下SE_ERR_XXX错误代码进行比较。
'   SE_ERR_FNF (2)  文件未找到。
'   SE_ERR_PNF (3)  路径没有找到。
'   SE_ERR_ACCESSDENIED (5) 拒绝访问。
'   SE_ERR_OOM (8)  内存不足。
'   SE_ERR_DLLNOTFOUND (32) 没有找到动态链接库。
'   SE_ERR_SHARE (26)   无法共享打开的文件。
'   SE_ERR_ASSOCINCOMPLETE (27) 文件关联信息不完整。
'   SE_ERR_DDETIMEOUT (28)  DDE操作超时。
'   SE_ERR_DDEFAIL (29) DDE操作失败。
'   SE_ERR_DDEBUSY (30) DDE操作繁忙。
'   SE_ERR_NOASSOC (31) 文件关联不可用。
'lpIDList
'   类型: LPVOID
'   绝对ITEMIDLIST结构(PCIDLIST_ABSOLUTE)的地址，该结构包含唯一标识要执行的文件的项标识符列表。如果fMask成员不包含SEE_MASK_IDLIST或SEE_MASK_INVOKEIDLIST，则忽略该成员。
'lpClass
'   类型: LPCTSTR
'   一个以null结尾的字符串的地址，该字符串指定以下内容之一:
'   ProgId。例如,“Paint.Picture”。
'   URI协议方案。例如,“http”。
'   一个文件扩展名。例如," .txt "。
'   HKEY_CLASSES_ROOT下的注册表路径，它为包含一个或多个Shell谓词的子键命名。这个键将有一个符合Shell谓词注册表模式的子键，例如
'   shell\verb name
'   如果fMask不包含SEE_MASK_CLASSNAME，则忽略该成员。
'hkeyClass
'   文件类型的注册表项句柄。这个注册表项的访问权限应该设置为KEY_READ。如果fMask不包含SEE_MASK_CLASSKEY，则忽略该成员。
'dwHotKey
'   与应用程序关联的键盘快捷方式。低阶单词是虚拟密钥代码，高阶单词是修饰符标志(HOTKEYF_)。有关修饰符标志的列表，请参见WM_SETHOTKEY消息的描述。如果fMask不包含SEE_MASK_HOTKEY，则忽略该成员。
'DUMMYUNIONNAME
'hIcon
'   文件类型图标的句柄。如果fMask不包含SEE_MASK_ICON，则忽略该成员。此值仅在Windows XP及更早版本中使用。它在Windows Vista中被忽略了。
'hMonitor
'   要在其上显示文档的监视器的句柄。如果fMask不包含SEE_MASK_HMONITOR，则忽略该成员。
'hProcess
'   新启动应用程序的句柄。这个成员在返回时被设置为NULL，除非fMask被设置为SEE_MASK_NOCLOSEPROCESS。即使fMask被设置为SEE_MASK_NOCLOSEPROCESS，如果没有启动任何进程，hProcess也是NULL。例如，如果要启动的文档是URL，并且Internet Explorer的实例已经在运行，那么它将显示该文档。没有启动新进程，hProcess将为NULL。
'   注意，ShellExecuteEx并不总是返回hProcess，即使调用的结果是启动了一个进程。例如，当使用SEE_MASK_INVOKEIDLIST调用IContextMenu时，hProcess不会返回。
'@备注
'    如果调用ShellExecuteEx的线程没有消息循环，或者线程或进程将在ShellExecuteEx返回后不久终止，则必须指定SEE_MASK_NOASYNC标志。在这种情况下，调用线程将无法完成DDE对话，因此在将控制权返回给调用应用程序之前，ShellExecuteEx必须完成对话。未能完成对话可能导致文档启动不成功。
'    如果调用线程有一个消息循环，并且在调用ShellExecuteEx返回后将存在一段时间，则SEE_MASK_NOASYNC标志是可选的。如果省略该标志，则调用线程的消息泵将用于完成DDE对话。调用应用程序可以更快地恢复控制，因为DDE对话可以在后台完成。
'    当使用fMask中的SEE_MASK_FLAG_LOG_USAGE标志填充最常用的程序列表时，对classic和Windows xp风格的开始菜单的计数是不同的。经典样式菜单只计算程序菜单中快捷方式的点击次数。Windows xp风格的菜单计算了程序菜单中快捷方式的点击量和程序菜单外快捷方式的目标点击量。因此，将lpFile设置为myfile.exe将影响Windows xp样式菜单的计数，无论该文件是直接启动的还是通过快捷方式启动的。经典样式(要求lpFile包含.lnk文件名)不会受到影响。
'    要在lpParameters中包含双引号，请将每个标记用一对引号括起来，如下面的示例所示。
'    sei.lpParameters = "An example: \"\"\"quoted text\"\"\"";
'    在本例中，应用程序接收三个参数:An。example:和 "quoted text"。
'@Requirements
'Minimum supported client    Windows XP [desktop apps only]
'Minimum supported server    Windows 2000 Server [desktop apps only]
'Header                      Shellapi.h
Private Declare Function ShellExecuteEx Lib "kernel32.dll" Alias "ShellExecuteExA" (pExecInfo As SHELLEXECUTEINFO) As Long
'@原型
'    BOOL ShellExecuteEx(
'      _Inout_ SHELLEXECUTEINFO *pExecInfo
'    );
'@功能
'    对指定的文件执行操作。
'@参数
'pExecInfo _Inout_
'    类型:SHELLEXECUTEINFO *
'    指向SHELLEXECUTEINFO结构的指针，该结构包含并接收有关正在执行的应用程序的信息。
'@返回值
'    如果成功返回TRUE;否则,假的。调用GetLastError获取扩展的错误信息。
'@备注
'    因为ShellExecuteEx可以将执行委托给使用组件对象模型(COM)激活的Shell扩展(数据源。上下文菜单处理程序。谓词实现)，所以应该在调用ShellExecuteEx之前初始化COM。一些Shell扩展需要COM单线程公寓(STA)类型。在这种情况下，COM应该如下所示初始化:
'    CoInitializeEx(NULL, coinit_apartmentthreads | COINIT_DISABLE_OLE1DDE)
'    在有些实例中，ShellExecuteEx不使用这些类型的Shell扩展，并且这些实例根本不需要初始化COM。尽管如此，在使用此函数之前，始终对COM进行初始化是一个很好的实践。
'    当dll加载到进程中时，您将获得一个称为加载器锁的锁。DllMain函数总是在加载器锁下执行。当您持有加载器锁时，不要调用ShellExecuteEx，这一点很重要。因为ShellExecuteEx是可扩展的，所以您可以加载在加载器锁存在的情况下不能正常运行的代码，从而冒着死锁的风险，从而导致线程无响应。
'    对于多个监视器，如果指定HWND并将lpExecInfo指向的SHELLEXECUTEINFO结构的lpVerb成员设置为“Properties”，那么由ShellExecuteEx创建的任何窗口都可能不会出现在正确的位置。
'    如果函数成功，它将SHELLEXECUTEINFO结构的hInstApp成员设置为大于32的值。如果函数失败，hInstApp被设置为SE_ERR_XXX错误值，该值最能指示失败的原因。尽管为了兼容16位Windows应用程序，hInstApp被声明为一个HINSTANCE，但它不是一个真正的HINSTANCE。它只能被转换为int，并且只能与值32或SE_ERR_XXX错误代码进行比较。
'    SE_ERR_XXX错误值是为了与ShellExecute兼容而提供的。要检索更准确的错误信息，请使用GetLastError。它可能返回以下值之一。
'    错误描述
'    ERROR_FILE_NOT_FOUND       未找到指定的文件。
'    ERROR_PATH_NOT_FOUND       未找到指定的路径。
Private Const ERROR_DDE_FAIL                    As Long = &H484
'    动态数据交换(DDE)事务失败。
Private Const ERROR_NO_ASSOCIATION              As Long = &H483
'    没有与指定的文件扩展名关联的应用程序。
Private Const ERROR_ACCESS_DENIED               As Long = &H5
'    拒绝对指定文件的访问。
Private Const ERROR_DLL_NOT_FOUND               As Long = &H485
'    找不到运行应用程序所需的库文件之一。
Private Const ERROR_CANCELLED                   As Long = &H4C7
'    l函数提示用户获取附加信息，但是用户取消了请求。
Private Const ERROR_NOT_ENOUGH_MEMORY           As Long = &H8
'    没有足够的内存来执行指定的操作。
Private Const ERROR_SHARING_VIOLATION           As Long = &H20
'    发生了共享冲突。
'    从URL打开项时，可以注册应用程序，以便在传递URL时激活。您还可以指定应用程序支持哪些协议。更多信息请参见申请注册。
'    从Windows 8开始，您可以提供指向ShellExecuteEx函数的站点链指针，以支持使用该站点的服务激活项。有关更多信息，请参见启动应用程序(ShellExecute。ShellExecuteEx。SHELLEXECUTEINFO)。
'@Requirements
'Minimum supported client       Windows XP [desktop apps only]
'Minimum supported server       Windows 2000 Server [desktop apps only]
'Header                         Shellapi.h
'Library                        shell32.lib
'dll                            Shell32.dll (version 3.51 or later)
'Unicode and ANSI names         ShellExecuteExW (Unicode) And ShellExecuteExA(ANSI)
Private Declare Function CoInitialize Lib "ole32.dll" (ByVal pvReserved As Long) As Long
'@原型
'    HRESULT CoInitialize(
'      _In_opt_ LPVOID pvReserved
'    );
'@功能
'    在当前线程上初始化COM库，并将并发模型标识为单线程公寓(STA)。新的应用程序应该调用CoInitializeEx，而不是初始化。如果你想使用Windows运行时,必须调Windows::Foundation::Initialize来替代。
'@参数
'pvReserved:这个参数是保留的，必须是空的。
'@返回：该函数可以返回标准返回值E_INVALIDARG。E_OUTOFMEMORY和E_UNEXPECTED值，以及以下值。
'        返回代码描述
'        S_OK:COM库在这个线程上成功地初始化了。
'        S_FALSE:COM库已经在这个线程上进行了初始化。
'        RPC_E_CHANGED_MODE:之前对CoInitializeEx的调用指定该线程的并发模型为多线程公寓(MTA)。这也可能表明，从中性线公寓到单线程公寓的转变已经发生。
'@注意事项：在调用除CoGetMalloc函数之外的任何库函数之前，您需要先在线程上初始化COM库，以获得一个指向标准分配器的指针，以及内存分配函数。
'        在设置了线程的并发模型之后，就不能更改它。
'        在以前初始化为多线程的场合中对CoInitialize的调用将失败，并返回RPC_E_CHANGED_MODE。
'        CoInitializeEx提供了与CoInitialize相同的功能，还提供了一个参数来显式地指定线程的并发模型。
'        CoInitialize调用CoInitializeEx，并将并发模型指定为单线程的公寓。
'        今天开发的应用程序应该调用CoInitializeEx，而不是CoInitialize。
'        通常，COM库只在线程上初始化一次。
'        在同一线程上对CoInitialize或CoInitializeEx的后续调用将会成功，只要它们不试图改变并发模型，但是会返回S_FALSE.。
'        要优雅地关闭COM库，每个成功调用CoInitialize或CoInitializeEx，包括返回S_FALSE.的调用，都必须通过相应的调用来进行相应的调用。
'        但是，应用程序中的第一个线程调用CoInitialize与0(或使用COINIT_APARTMENTTHREADED的CoInitializeEx)，必须是调用可数的最后一个线程。
'        否则，在STA上对CoInitialize的调用将会失败，应用程序将无法工作。
'        因为没有办法控制进程内的服务器被加载或卸载的顺序，所以不要从DllMain函数中调用CoInitialize。CoInitializeEx或可数。
'@Requirements
'Minimum supported client       Windows 2000 Professional [desktop apps only]
'Minimum supported server       Windows 2000 Server [desktop apps only]
'Header                         Objbase.h
'Library                        Ole32.lib
'dll                            Ole32.dll
Private Declare Sub CoUninitialize Lib "ole32.dll" ()
'@原型
'    void CoUninitialize(void);
'@功能
'    关闭当前线程上的COM库，卸载线程加载的所有dll，释放线程维护的其他资源，并强制所有的RPC连接在线程上关闭。
'注意事项：
'    一个线程必须为它对CoUninitialize 或CoInitializeEx函数的每个成功调用调用一次CoUninitialize ，包括返回S_FALSE的任何调用。只有对初始化库CoInitialize或CoInitializeEx调用对应的调用才能关闭它。
'    对 OleInitialize 调用必须通过对 OleUninitialize调用来平衡。OleUninitialize 函数调用CoUninitialize ，因此调用OleUninitialize 的应用程序也不需要调用 CoUninitialize。
'    应该在应用程序关闭时调用CoUninitialize ，这是在应用程序隐藏其主窗口并通过其主消息循环后对COM库进行的最后一次调用。如果还存在开放的会话，则可数启动一个模式消息循环，并为这个COM应用程序从容器或服务器发送任何待处理的消息。通过发送消息，CoUninitialize 确保应用程序在接收到所有挂起的消息之前不会退出。Non-COM消息丢弃。
'    因为没有办法控制进程内的服务器被加载或卸载的顺序，所以不要从DllMain函数中调用CoInitialize。CoInitializeEx或可数。
'@Requirements
'Minimum supported client       Windows 2000 Professional [desktop apps only]
'Minimum supported server       Windows 2000 Server [desktop apps only]
'Header                         Objbase.h
'Library                        Ole32.lib
'dll                            Ole32.dll
Private Declare Function CoInitializeEx Lib "ole32.dll" (ByVal pvReserved As Long, ByVal dwCoInit As Long) As Long
'@原型
'    HRESULT CoInitializeEx(
'      _In_opt_ LPVOID pvReserved,
'      _In_     DWORD  dwCoInit
'    );
'@功能
'    初始化COM库供调用线程使用，设置线程的并发模型，并在需要时为线程创建一个新单元。
'    如果想使用Windows运行时api，或者想同时使用COM和Windows运行时组件，应该调用Windows::Foundation::Initialize来初始化线程，而不是CoInitializeEx。初始化对于COM组件来说已经足够了。
'@参数
'pvReserved _In_opt_
'   保留此参数，且必须为空。
'dwCoInit   _In_
'   线程的并发模型和初始化选项。此参数的值取自COINIT枚举。可以使用COINIT中的任何值组合，但coinit_apartmentthread和coinit_multithread标志不能同时设置。默认值是coinit_multithread。
'@返回：该函数可以返回标准返回值E_INVALIDARG。E_OUTOFMEMORY和E_UNEXPECTED值，以及以下值。
'        返回代码描述
'        S_OK:在这个线程上成功初始化了COM库
'        S_FALSE:COM库已经在这个线程上初始化。
'        RPC_E_CHANGED_MODE:前面对CoInitializeEx的调用将此线程的并发模型指定为多线程单元(MTA)。这也可能表明已经发生了从中性线程公寓到单线程公寓的更改。
'@注意事项
'    对于使用COM库的每个线程，CoInitializeEx必须至少调用一次，而且通常只调用一次。只要通过相同的并发标志，同一个线程可以多次调用CoInitializeEx，但是后续的有效调用返回S_FALSE。要在线程上优雅地关闭COM库，每个成功的CoInitialize或CoInitializeEx调用(包括返回S_FALSE的任何调用)都必须通过相应的CoUninitialize调用进行平衡。
'    在调用除CoGetMalloc之外的任何库函数之前，需要在线程上初始化COM库，以获得指向标准分配器的指针和内存分配函数。否则，COM函数将返回CO_E_NOTINITIALIZED。
'    线程的并发模型设置好后，就不能更改它了。在以前被初始化为多线程的单元上进行协初始化的调用将失败，并返回RPC_E_CHANGED_MODE。
'    在单线程公寓(STA)中创建的对象仅从其公寓的线程接收方法调用，因此调用被序列化，并且只到达消息队列边界(当调用PeekMessage或SendMessage函数时)。在多线程单元(MTA)的COM线程上创建的对象必须能够随时接收来自其他线程的方法调用。您通常会在多线程对象的代码中实现某种形式的并发控制，使用同步原语(如临界段。信号量或互斥锁)来帮助保护对象的数据。当配置为在中立线程单元(NTA)中运行的对象被STA或MTA中的线程调用时，该线程将传输到NTA。如果该线程随后调用CoInitializeEx，则调用将失败并返回RPC_E_CHANGED_MODE。
'    由于OLE技术不是线程安全的，OleInitialize函数使用coinit_apartmentthread标志调用CoInitializeEx。因此，为多线程对象并发性初始化的单元不能使用OleInitialize启用的特性。
'    因为无法控制进程内服务器加载或卸载的顺序，所以不要从DllMain函数调用CoInitialize。CoInitializeEx或CoUninitialize。
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
'@功能
'    确定用于此线程创建的对象的传入调用的并发模型。这个并发模型可以是单线程的，也可以是多线程的。
'@常量
'COINIT_APARTMENTTHREADED
'   为单线程对象并发初始化线程(请参阅注释)。
'COINIT_MULTITHREADED
'   初始化多线程对象并发的线程(参见备注)。
'COINIT_DISABLE_OLE1DDE
'   为OLE1支持禁用DDE。
'COINIT_SPEED_OVER_MEMORY
'   增加内存使用量以提高性能。
'@备注
'    当通过调用CoInitializeEx初始化一个线程时，您可以通过指定COINIT的一个成员作为它的第二个参数，来选择将它初始化为单线程还是多线程。它指定如何处理该线程创建的任何对象的传入调用，即对象的并发性。
'    分体线程，虽然允许多个线程执行，序列化所有传入的调用，要求调用对象的方法创建的这个线程总是运行在同一个线程上的公寓/线程创建他们。此外，调用只能到达消息队列边界。由于这种序列化，通常不需要将并发控制写入对象的代码中，除非在处理过程中避免对PeekMessage和SendMessage的调用，这些调用不能被其他方法调用或对同一单元/线程中的其他对象的调用打断。
'    多线程(也称为自由线程)允许对这个线程创建的对象的方法的调用在任何线程上运行。“许多调用可能发生在相同的方法或相同的对象或同时。多线程对象并发性为跨线程。跨进程和跨机器调用提供了最高的性能，并充分利用了多处理器硬件，因为对对象的调用不会以任何方式序列化。然而，这意味着对象的代码必须强制自己的并发模型，通常通过使用同步原语，例如临界段。信号量或互斥量。此外，由于对象不控制访问它的线程的生存期，因此对象中(在线程本地存储中)可能不存储任何特定于线程的状态。
'    注意，多线程单元用于非gui线程。多线程单元中的线程不应该执行UI操作。这是因为UI线程需要消息泵，而COM不为多线程单元中的线程泵送消息。
'@Requirements
'Minimum supported client       Windows 2000 Professional [desktop apps only]
'Minimum supported server       Windows 2000 Server [desktop apps only]
'Header                         Objbase.h
Private Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
'@原型
'    void MoveMemory(
'      _In_       PVOID  Destination,
'      _In_ const VOID   *Source,
'      _In_       SIZE_T Length
'    );
'@功能
'    将内存块从一个位置移动到另一个位置。
'@参数
'Destination
'   指向移动目标的起始地址的指针。
'Source
'   指向要移动的内存块的起始地址的指针。
'Length
'   要移动的内存块的大小，以字节为单位。
'@返回值
'   无
'@备注
'    这个函数被定义为RtlMoveMemory函数。它的实现是内联提供的。有关更多信息，请参见WinBase.h和Winnt.h。
'    源块和目标块可能重叠。
'@安全
'    第一个参数Destination必须足够大，以容纳源的长度字节;否则，可能会发生缓冲区溢出。如果发生访问冲突，或者在最坏的情况下，允许攻击者将可执行代码注入您的进程，则可能导致对应用程序的拒绝服务攻击。如果Destination是基于堆栈的缓冲区，则尤其如此。请注意，最后一个参数Length是要复制到目标的字节数，而不是目标的大小。
'@Requirements
'Minimum supported client       Windows XP [desktop apps only]
'Minimum supported server       Windows Server 2003 [desktop apps only]
'Header                         WinBase.h (include Windows.h)
'dll                            kernel32.dll
Private Declare Function lstrlen Lib "kernel32.dll" Alias "lstrlenW" (ByVal lpString As Long) As Long
Private Declare Function lstrcpy Lib "kernel32" Alias "lstrcpyW" (ByVal lpString1 As Long, ByVal lpString2 As Long) As Long
'@原型
'    LPTSTR WINAPI lstrcpy(
'      _Out_ LPTSTR lpString1,
'      _In_  LPTSTR lpString2
'    );
'@功能
'    将字符串复制到缓冲区。警告 不要使用?考虑使用StringCchCopy?
'@参数
'    lpString1
'    用于接收由lpString2参数指向的字符串内容的缓冲区。缓冲区必须足够大，以包含字符串，包括终止null字符。
'    lpString2
'    要复制的以null结尾的字符串?
'@返回值
'    如果函数成功，返回值是指向缓冲区的指针。
'    如果函数失败，返回值为NULL, lpString1可能不会以NULL结尾。
'@备注
'    使用系统的双字节字符集(DBCS)版本，可以使用此函数复制DBCS字符串。
'    如果源缓冲区和目标缓冲区重叠，则lstrcpy函数具有未定义的行为。
'@安全评价
'    不正确地使用此函数会损害应用程序的安全性。此函数使用结构化异常处理(SEH)捕捉访问违规和其他错误。当这个函数捕捉到SEH错误时，它返回NULL而不终止字符串，也不通知调用者错误。调用方不能安全地假定错误条件是空间不足。
'    lpString1必须足够大，以容纳lpString2和结束'\0'，否则可能会发生缓冲区溢出。
'    缓冲区溢出是应用程序中许多安全问题的原因，如果发生访问冲突，可能会导致对应用程序的拒绝服务攻击。在最坏的情况下，缓冲区溢出可能允许攻击者将可执行代码注入您的进程，特别是如果lpString1是基于堆栈的缓冲区。
'    考虑使用StringCchCopy;使用StringCchCopy(缓冲区,sizeof(缓冲)/ sizeof(缓冲[0]),src);,意识到缓冲区必须不是一个指针或使用StringCchCopy(缓冲区,ARRAYSIZE(缓冲),src);,被意识到,当复制指针,调用者负责传递在字符的指针的内存的大小。
'@Requirements
'Minimum supported client       Windows 2000 Professional [desktop apps only]
'Minimum supported server       Windows 2000 Server [desktop apps only]
'Header                         Winbase.h (include Windows.h)
'Library                        kernel32.lib
'dll                            kernel32.dll
'Unicode and ANSI names         lstrcpyW (Unicode) And lstrcpyA(ANSI)
Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
'@原型
'    BOOL WINAPI GetUserName(
'      _Out_   LPTSTR  lpBuffer,
'      _Inout_ LPDWORD lpnSize
'    );
'@功能
'    检索与当前线程关联的用户名。
'    使用GetUserNameEx函数以指定的格式检索用户名。附加信息由IADsADSystemInfo接口提供。
'@参数
Private Const UNLEN                             As Long = 256
'lpBuffer _Out_
'   指向缓冲区的指针，用于接收用户的登录名。如果该缓冲区不够大，无法包含整个用户名，则函数将失败。缓冲区大小(UNLEN + 1)字符将保存最大长度的用户名，包括终止null字符。UNLEN在lmcon .h中定义。
'lpnSize _Inout_
'   在输入时，这个变量在TCHARs中指定lpBuffer缓冲区的大小。在输出时，变量接收复制到缓冲区的TCHARs的数量，包括终止null字符。
'   如果lpBuffer太小，函数就会失败，GetLastError返回ERROR_INSUFFICIENT_BUFFER。此参数接收所需的缓冲区大小，包括终止空字符。
Private Const ERROR_INSUFFICIENT_BUFFER         As Long = &H7A
'@返回值
'   如果函数成功，返回值为非零值，lpnSize指向的变量包含复制到lpBuffer指定的缓冲区的TCHARs的数量，包括终止null字符。
'   如果函数失败，返回值为零。要获取扩展的错误信息，请调用GetLastError。
'@备注
'   如果当前线程正在模拟另一个客户机，GetUserName函数将返回线程正在模拟的客户机的用户名。
'   如果从运行在“网络服务”帐户下的进程调用GetUserName, lpBuffer中返回的字符串可能会因Windows版本的不同而不同。在Windows XP中，返回“NETWORK SERVICE”字符串。在Windows Vista中，返回“<HOSTNAME>$”字符串。
'@Requirements
'Minimum supported client       Windows 2000 Professional [desktop apps only]
'Minimum supported server       Windows 2000 Server [desktop apps only]
'Header                         Winbase.h (include Windows.h)
'Library                        advapi32.lib
'dll                            advapi32.dll
'Unicode and ANSI names         GetUserNameW (Unicode) And GetUserNameA(ANSI)
Private Declare Function GetUserNameEx Lib "Secur32.dll" Alias "GetUserNameExA" (ByVal NameFormat As Long, ByVal lpNameBuffer As String, lpnSize As Long) As Long
'@功能
'    检索与调用线程关联的用户或其他安全主体的名称。您可以指定返回名称的格式。
'    如果线程正在模拟客户机，GetUserNameEx将返回客户机的名称。
'@原型
'    BOOLEAN WINAPI GetUserNameEx(
'      _In_    EXTENDED_NAME_FORMAT NameFormat,
'      _Out_   LPTSTR               lpNameBuffer,
'      _Inout_ PULONG               lpnSize
'    );
'@参数
'NameFormat _In_
'   名称的格式。该参数是EXTENDED_NAME_FORMAT枚举类型的值。它不能是无名的。如果用户帐户不在域中，则只支持NameSamCompatible。
'lpNameBuffer  _Out_
'   指向以指定格式接收名称的缓冲区的指针。缓冲区必须包含终止null字符的空间。
'lpnSize    _Inout_
'   在输入时，这个变量在TCHARs中指定lpNameBuffer缓冲区的大小。如果函数成功，则变量接收复制到缓冲区的TCHARs的数量，不包括终止null字符。
'   如果lpNameBuffer太小，函数就会失败，GetLastError返回ERROR_MORE_DATA。该参数接收所需的缓冲区大小(以Unicode字符为单位)(无论是否使用Unicode)，包括终止null字符。
'@返回值
'   如果函数成功，返回值为非零值。
'   如果函数失败，返回值为零。要获取扩展的错误信息，请调用GetLastError。可能的值包括以下内容。
'   返回代码描述
Private Const ERROR_MORE_DATA                   As Long = &HEA
'   lpNameBuffer缓冲区太小。lpnSize参数包含接收名称所需的字节数。
Private Const ERROR_NO_SUCH_DOMAIN              As Long = &H54B
'   域控制器不可用于执行查找
Private Const ERROR_NONE_MAPPED                 As Long = &H534
'   用户名在指定格式中不可用。
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
'@功能
'    指定目录服务对象名称的格式。
'@常量
'NameUnknown
'   未知名称类型。
'NameFullyQualifiedDN
'   完全限定的专有名称(例如，CN=Jeff Smith,OU=Users,DC=Engineering,DC=Microsoft,DC=Com)。
'NameSamCompatible
'   遗留帐户名(例如，Engineering\JSmith)。仅限域的版本包括后置反斜杠(\\)。
'NameDisplay
'   一个“友好”的显示名称(例如，Jeff Smith)。显示名称不一定是定义的相对专有名称(RDN)。
'NameUniqueId
'   IIDFromString函数返回的GUID字符串(例如，{4fa050f0-f561-11cf-bdd9-00aa003a77b6})。
'NameCanonical
'   完整的规范名称(例如，engineering.microsoft.com/software/someone)。只有域的版本包含一个后斜杠(/)。
'NameUserPrincipal
'   用户主体名称(例如，someone@example.com)。
'NameCanonicalEx
'   与NameCanonical相同，只是最右边的斜杠(/)被替换为一个新行字符(\n)，即使在只有域的情况下也是如此(例如，engineering.microsoft.com/software\nJSmith)。
'NameServicePrincipal
'   通用服务主体名称(例如，www/www.microsoft.com@microsoft.com)。
'NameDnsDomain
'   后跟反斜杠和SAM用户名的DNS域名。
'@Requirements
'Minimum supported client       Windows 2000 Professional [desktop apps only]
'Minimum supported server       Windows 2000 Server [desktop apps only]
'Header                         Secext.h (include Security.h)
Private Declare Function CheckTokenMembership Lib "advapi32.dll" (ByVal TokenHandle As Long, ByVal SidToCheck As Long, IsMember As Long) As Long
'@原型
'    BOOL WINAPI CheckTokenMembership(
'      _In_opt_ HANDLE TokenHandle,
'      _In_     PSID   SidToCheck,
'      _Out_    PBOOL  IsMember
'    );
'@功能
'    CheckTokenMembership函数确定是否在访问令牌中启用了指定的安全标识符(SID)。如果想确定应用程序容器令牌的组成员关系，需要使用CheckTokenMembershipEx函数。
'@参数
'TokenHandle _In_opt_
'   访问令牌的句柄。句柄必须具有对令牌的TOKEN_QUERY访问权。令牌必须是模拟令牌。
'   如果TokenHandle为空，CheckTokenMembership将使用调用线程的模拟令牌。如果线程没有模拟，该函数将复制线程的主令牌来创建模拟令牌。
'SidToCheck _In_
'   指向SID结构的指针。CheckTokenMembership函数检查访问令牌的用户和组SID中是否存在此SID。
'IsMember  _Out_
'   指向接收检查结果的变量的指针。如果SID存在并且具有SE_GROUP_ENABLED属性，则IsMember返回TRUE;否则，返回FALSE。
'@返回值
'   如果函数成功，返回值为非零。
'   如果函数失败，返回值为零。要获取扩展的错误信息，请调用GetLastError。
'@备注
'   CheckTokenMembership函数简化了确定访问令牌中是否同时存在和启用SID的过程。
'   即使令牌中存在SID，系统也可能不会在访问检查中使用该SID。SID可能被禁用，或者只有SE_GROUP_USE_FOR_DENY_ONLY属性。该系统只使用SID在进行访问检查时授予访问权。有关更多信息，请参见访问令牌中的SID属性。
'   如果TokenHandle是受限制的令牌，或者如果TokenHandle为NULL，并且调用线程当前有效的令牌是受限制的令牌，CheckTokenMembership还将检查SID是否出现在受限制的SID列表中。
'@Requirements
'Minimum supported client       Windows XP [desktop apps | UWP apps]
'Minimum supported server       Windows Server 2003 [desktop apps | UWP apps]
'Header                         Winbase.h (include Windows.h)
'Library                        advapi32.lib
'dll                            advapi32.dll
Private Declare Function ProcessIdToSessionId Lib "kernel32.dll" (ByVal dwProcessId As Long, pSessionId As Long) As Long
'@原型
'    BOOL ProcessIdToSessionId(
'      DWORD dwProcessId,
'      DWORD *pSessionId
'    );
'@功能
'    检索与指定进程关联的远程桌面服务会话。
'@参数
'dwProcessId
'    指定进程标识符?使用GetCurrentProcessId函数检索当前进程的进程标识符?
'pSessionId
'    指向一个变量的指针，该变量接收正在运行指定进程的远程桌面服务会话的标识符。要检索当前附加到控制台的会话的标识符，请使用WTSGetActiveConsoleSessionId函数。
'@返回值
'    如果函数成功，返回值为非零值。
'    如果函数失败，返回值为零。要获取扩展的错误信息，请调用GetLastError。
'@备注
'    调用者必须持有指定进程的PROCESS_QUERY_INFORMATION访问权。有关更多信息，请参见流程安全和访问权限。
'@Requirements
'Minimum supported client       Windows XP [desktop apps | UWP apps]
'Minimum supported server       Windows Server 2003 [desktop apps | UWP apps]
'Header                         Winbase.h (include Windows.h)
'Library                        Kernel32.lib
'dll                            Kernel32.dll
Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long
'@原型
'    DWORD GetCurrentProcessId(
'
'    );
'@功能
'    检索调用进程的进程标识符。
'@参数
'    这个函数没有参数?
'@返回值
'    返回值是调用进程的进程标识符?
'@备注
'    在进程终止之前，进程标识符在整个系统中唯一地标识进程。
'@Requirements
'Minimum supported client       Windows XP [desktop apps | UWP apps]
'Minimum supported server       Windows Server 2003 [desktop apps | UWP apps]
'Header                         processthreadsapi.h (include Windows Server 2003, Windows Vista, Windows 7, Windows Server 2008 Windows Server 2008 R2, Windows.h)
'Library                        Kernel32.lib
'dll                            Kernel32.dll
Private Declare Function CopySid Lib "advapi32" (ByVal nDestinationSidLength As Long, pDestinationSid As Long, ByVal pSourceSid As Long) As Long
'@原型
'    BOOL WINAPI CopySid(
'      _In_  DWORD nDestinationSidLength,
'      _Out_ PSID  pDestinationSid,
'      _In_  PSID  pSourceSid
'    );
'@功能
'    CopySid函数将安全标识符(SID)复制到缓冲区。
'@参数
'    nDestinationSidLength
'    指定接收SID副本的缓冲区的长度(以字节为单位)。
'    pDestinationSid
'    指向缓冲区的指针，该缓冲区接收源SID结构的副本。
'    pSourceSid
'    指向SID结构的指针，该函数将该结构复制到由pDestinationSid参数指向的缓冲区。
'@返回值
'    如果函数成功，返回值为非零。
'    如果函数失败，返回值为零。要获取扩展的错误信息，请调用GetLastError。
'@备注
'    应用程序可以使用CopySid函数在访问令牌(例如，在TOKEN_GROUPS结构中)中复制一个SID，以便在访问控制条目(ACE)中使用。
'@Requirements
'Minimum supported client       Windows XP [desktop apps | UWP apps]
'Minimum supported server       Windows Server 2003 [desktop apps | UWP apps]
'Header                         Sddl.h
'Library                        advapi32.lib
'dll                            advapi32.dll
Private Declare Function GetLengthSid Lib "advapi32" (pSid As Long) As Long
'@原型
'    DWORD WINAPI GetLengthSid(
'      _In_ PSID pSid
'    );
'@功能
'    GetLengthSid函数返回有效安全标识符(SID)的长度，以字节为单位。
'@参数
'pSid
'    返回长度为SID结构的指针?这个结构被认为是有效的?
'返回值
'    如果SID结构是有效的，返回值就是SID结构的长度(以字节为单位)。
'    如果SID结构无效，则返回值未定义。在调用GetLengthSid之前，将SID传递给IsValidSid函数，以验证SID是否有效。
'@Requirements
'Minimum supported client       Windows XP [desktop apps | UWP apps]
'Minimum supported server       Windows Server 2003 [desktop apps | UWP apps]
'Header                         Sddl.h
'Library                        advapi32.lib
'dll                            advapi32.dll
Private Declare Function ConvertSidToStringSid Lib "advapi32" Alias "ConvertSidToStringSidW" (pSid As Any, StringSid As Long) As Long
'@原型
'    BOOL ConvertSidToStringSid(
'      _In_  PSID   Sid,
'      _Out_ LPTSTR *StringSid
'    );
'@功能
'    ConvertSidToStringSid函数将安全标识符(SID)转换为适合显示、存储或传输的字符串格式。要将字符串格式SID转换回有效的函数SID，请调用ConvertStringSidToSid函数。
'@参数
'Sid
'    指向要转换的SID结构的指针?
'StringSid
'    指向变量的指针，该变量接收指向以null结尾的SID字符串的指针。要释放返回的缓冲区，请调用LocalFree函数。
'@返回值
'   如果函数成功，返回值为非零。
'   如果函数失败，返回值为零。要获取扩展的错误信息，请调用GetLastError。GetLastError函数可能返回以下错误代码之一。
'返回代码描述
'   ERROR_NOT_ENOUGH_MEMORY    内存不足?
Private Const ERROR_INVALID_SID                 As Long = &H539
'   SID无效?
Private Const ERROR_INVALID_PARAMETER           As Long = &H57
'   其中一个参数包含一个无效的值?这通常是一个无效的指针
'@备注
'    ConvertSidToStringSid函数对SID字符串使用标准的S-R-I-S-S…格式。有关SID字符串表示法的更多信息，请参见SID组件
'@Requirements
'Minimum supported client       Windows XP [desktop apps | UWP apps]
'Minimum supported server       Windows Server 2003 [desktop apps | UWP apps]
'Header                         Sddl.h
'Library                        advapi32.lib
'dll                            advapi32.dll
'Unicode and ANSI names         ConvertSidToStringSidW (Unicode) And ConvertSidToStringSidA(ANSI)
Private Declare Function LookupAccountSid Lib "advapi32.dll" Alias "LookupAccountSidA" (ByVal lpSystemName As String, ByVal Sid As Long, ByVal name As String, cbName As Long, ByVal ReferencedDomainName As String, cbReferencedDomainName As Long, peUse As Long) As Long
'@原型
'    BOOL WINAPI LookupAccountSid(
'      _In_opt_  LPCTSTR       lpSystemName,
'      _In_      PSID          lpSid,
'      _Out_opt_ LPTSTR        lpName,
'      _Inout_   LPDWORD       cchName,
'      _Out_opt_ LPTSTR        lpReferencedDomainName,
'      _Inout_   LPDWORD       cchReferencedDomainName,
'      _Out_     PSID_NAME_USE peUse
'    );
'@功能
'    LookupAccountSid函数接受安全标识符(SID)作为输入。它检索此SID的帐户名称和找到此SID的第一个域的名称。
'@参数
'lpSystemName(,可选)
'   指向指定目标计算机的以null结尾的字符串的指针。这个字符串可以是远程计算机的名称。如果该参数为NULL，则在本地系统上开始帐户名称转换。如果无法在本地系统上解析该名称，则此函数将尝试使用本地系统信任的域控制器解析该名称。通常，只有当帐户位于不受信任的域中且该域中计算机的名称已知时，才为lpSystemName指定一个值。
'lpSid [在]
'   指向要查找的SID的指针?
'lpName(,可选)
'   指向缓冲区的指针，该缓冲区接收一个以null结尾的字符串，该字符串包含与lpSid参数对应的帐户名。
'cchName [,]
'   On input指定lpName缓冲区的大小(以TCHARs为单位)。如果函数因为缓冲区太小或cchName为零而失败，则cchName接收所需的缓冲区大小，包括终止null字符。
'lpReferencedDomainName(,可选)
'   指向缓冲区的指针，该缓冲区接收一个以null结尾的字符串，该字符串包含找到帐户名的域的名称。
'   在服务器上，为本地计算机的安全数据库中的大多数帐户返回的域名是服务器作为域控制器的域名。
'   在工作站上，本地计算机的安全数据库中为大多数帐户返回的域名是系统最后一次启动时计算机的名称(不包括反斜杠)。如果计算机的名称发生更改，则将继续返回旧名称作为域名，直到重新启动系统为止。
'   有些帐户是由系统预先定义的?为这些帐户返回的域名是BUILTIN?
'cchReferencedDomainName [,]
'   On input，在TCHARs中指定lpReferencedDomainName缓冲区的大小。如果函数因为缓冲区太小而失败，或者cchReferencedDomainName为零，则cchReferencedDomainName接收所需的缓冲区大小，包括终止null字符。
'peUse [出]
'   指向一个变量的指针，该变量接收一个表示帐户类型的SID_NAME_USE值。
'@返回值
'    如果函数成功，则函数返回非零。
'    如果函数失败，它返回零。要获取扩展的错误信息，请调用GetLastError。
'@备注
'   LookupAccountSid函数首先检查一个已知SID列表，试图为指定的SID找到一个名称。如果提供的SID与已知的SID不对应，该函数将检查内置的和管理上定义的本地帐户。接下来，函数检查主域。主域不能识别的安全标识符将根据与其SID前缀对应的可信域进行检查。
'   如果函数找不到SID的帐户名，GetLastError将返回error_none_mapping。如果网络超时阻止函数查找名称，就会发生这种情况。对于没有对应帐户名的SID也会发生这种情况，例如标识登录会话的登录SID。
'   除了查找SID的当地帐户、当地域帐户和明确受信任的域帐户外，LookupAccountSid还可以查找SID在森林中任何领域的任何帐户，包括只出现在森林帐户SIDhistory字段中的SID。SIDhistory字段存储从另一个域移动过来的帐户的前SID。要查找SID, LookupAccountSid查询forest的全局目录。
'@Requirements
'Minimum supported client       Windows XP [desktop apps | UWP apps]
'Minimum supported server       Windows Server 2003 [desktop apps | UWP apps]
'Header                         Sddl.h
'Library                        advapi32.lib
'dll                            advapi32.dll
'Unicode and ANSI names         LookupAccountSidW (Unicode) And LookupAccountSidA(ANSI)
Private Declare Function LookupAccountName Lib "advapi32.dll" Alias "LookupAccountNameW" (ByVal lpSystemName As Long, ByVal lpAccountName As Long, ByRef Sid As Any, ByRef cbSid As Long, ByVal ReferencedDomainName As Long, ByRef cbReferencedDomainName As Long, ByRef peUse As Long) As Long
'@原型
'    BOOL WINAPI LookupAccountName(
'      _In_opt_  LPCTSTR       lpSystemName,
'      _In_      LPCTSTR       lpAccountName,
'      _Out_opt_ PSID          Sid,
'      _Inout_   LPDWORD       cbSid,
'      _Out_opt_ LPTSTR        ReferencedDomainName,
'      _Inout_   LPDWORD       cchReferencedDomainName,
'      _Out_     PSID_NAME_USE peUse
'    );
'@功能
'    LookupAccountName函数接受系统名和帐户名作为输入。它检索帐户的安全标识符(SID)和找到帐户所在域的名称。LsaLookupNames函数还可以检索计算机帐户?
'@参数
'lpSystemName(,可选)
'   指向以null结尾的字符串的指针，该字符串指定系统的名称。这个字符串可以是远程计算机的名称。如果此字符串为空，则在本地系统上开始帐户名称转换。如果无法在本地系统上解析该名称，则此函数将尝试使用本地系统信任的域控制器解析该名称。通常，只有当帐户位于不受信任的域中且该域中计算机的名称已知时，才为lpSystemName指定一个值。
'lpAccountName [在]
'   指向以null结尾的字符串的指针，该字符串指定帐户名。
'   使用domain_name\user_name格式中的完全限定字符串，以确保LookupAccountName找到所需域中的帐户。
'Sid(,可选)
'   一个指向缓冲区的指针，该缓冲区接收与lpAccountName参数所指向的帐户名相对应的SID结构。如果该参数为NULL，则cbSid必须为零。
'cbSid [,]
'   指向变量的指针。在输入时，此值指定Sid缓冲区的大小(以字节为单位)。如果函数因为缓冲区太小而失败，或者cbSid为零，则此变量将接收所需的缓冲区大小。
'ReferencedDomainName(,可选)
'   指向缓冲区的指针，该缓冲区接收找到帐户名的域的名称。对于没有连接到域的计算机，此缓冲区接收计算机名称。如果该参数为NULL，则函数返回所需的缓冲区大小。
'cchReferencedDomainName [,]
'   指向变量的指针。在输入时，此值指定ReferencedDomainName缓冲区的大小(在TCHARs中)。如果函数因为缓冲区太小而失败，则此变量将接收所需的缓冲区大小，包括终止null字符。如果ReferencedDomainName参数为NULL，则该参数必须为零。
'peUse [出]
'   指向SID_NAME_USE枚举类型的指针，该类型指示函数返回时帐户的类型。
'@返回值
'    如果函数成功，则函数返回非零。
'    如果函数失败，它返回零。要获取扩展的错误信息，请调用GetLastError。
'@备注
'    LookupAccountName函数首先检查一个已知SID列表，试图为指定的名称找到SID。如果名称与已知的SID不对应，该函数将检查内置的和管理上定义的本地帐户。接下来，函数检查主域。如果没有找到该名称，则检查可信域。
'    使用完全限定的帐户名(例如，domain_name\user_name)而不是孤立的名称(例如，user_name)。完全限定名是明确的，并且在执行查找时提供更好的性能。该函数还支持完全限定的DNS名称(例如，example.example.com\user_name)和用户主体名称(UPN)(例如，someone@example.com)。
'    除了查找本地帐户、本地域帐户和显式受信任的域帐户外，LookupAccountName还可以查找林中任意域中的任意帐户的名称。
'@Requirements
'Minimum supported client       Windows XP [desktop apps | UWP apps]
'Minimum supported server       Windows Server 2003 [desktop apps | UWP apps]
'Header                         Sddl.h
'Library                        advapi32.lib
'dll                            advapi32.dll
'Unicode and ANSI names         LookupAccountNameW (Unicode) And LookupAccountNameA(ANSI)
Private Declare Function LocalFree Lib "kernel32" (hMem As Long) As Long
'@原型
'    HLOCAL WINAPI LocalFree(
'      _In_ HLOCAL hMem
'    );
'@功能
'    释放指定的本地内存对象并使其句柄无效?
'    注意，与其他内存管理函数相比，本地函数的开销更大，提供的特性更少。新应用程序应该使用堆函数，除非文档声明应该使用本地函数。有关更多信息，请参见全局和本地函数。
'@参数
'hMem
'    本地内存对象的句柄?这个句柄由LocalAlloc或LocalReAlloc函数返回?使用GlobalAlloc释放内存是不安全的?
'@返回值
'    如果函数成功，返回值为NULL。
'    如果函数失败，返回值等于本地内存对象的句柄。要获取扩展的错误信息，请调用GetLastError。
'@备注
'    如果进程试图在释放内存之后检查或修改内存，则可能会发生堆损坏或生成访问冲突异常(EXCEPTION_ACCESS_VIOLATION)。
'    如果hMem参数为NULL , LocalFree将忽略该参数并返回NULL?
'    LocalFree函数将释放一个锁定的内存对象。被锁定的内存对象的锁计数大于零。LocalLock函数锁定一个本地内存对象，并将锁计数增加1。LocalUnlock函数将其解锁，并将锁计数减少1。要获取本地内存对象的锁计数，请使用LocalFlags函数。
'    如果应用程序在系统的调试版本下运行，LocalFree将发出一条消息，告诉您释放了一个锁定的对象。如果正在调试应用程序，LocalFree将在释放锁定对象之前输入一个断点。这允许您验证预期的行为，然后继续执行。
'@Requirements
'Minimum supported client       Windows XP [desktop apps | UWP apps]
'Minimum supported server       Windows Server 2003 [desktop apps | UWP apps]
'Header                         WinBase.h (include Windows.h)
'Library                        Kernel32.lib
'dll                            Kernel32.dll
Private Declare Function WTSQueryUserToken Lib "wtsapi32" (ByVal SessionId As Long, phToken As Long) As Long
'@原型
'    BOOL WTSQueryUserToken(
'      ULONG   SessionId,
'      Phandle phToken
'    );
'@功能
'    获取会话ID指定的登录用户的主访问令牌。要成功调用此函数，调用应用程序必须在LocalSystem帐户上下文中运行，并具有SE_TCB_NAME特权。
'    警告:WTSQueryUserToken适用于高度信任的服务。服务提供者必须使用警告，以免在调用此函数时泄漏用户令牌。服务提供者必须在使用令牌句柄之后关闭令牌句柄。
'@参数
'SessionId
'    远程桌面服务会话标识符。在服务上下文中运行的任何程序的会话标识符都为0(0)。您可以使用WTSEnumerateSessions函数来检索指定RD会话主机服务器上所有会话的标识符。
'    要能够为其他用户的会话查询信息，您需要具有查询信息权限。有关更多信息，请参见远程桌面服务权限。要修改会话的权限，请使用远程桌面服务配置管理工具。
'phToken
'    如果函数成功，则接收一个指向已登录用户的令牌句柄的指针。注意，必须调用close句柄函数才能关闭这个句柄。
'@返回值
'    如果函数成功，返回值为非零值，phToken参数指向用户的主令牌。
'    如果函数失败，返回值为零。要获取扩展的错误信息，请调用GetLastError。在其他错误中，GetLastError可以返回以下错误之一。
'@备注
'    有关主令牌的信息，请参见访问令牌。有关帐户特权的更多信息，请参见远程桌面服务权限和授权常量。
'    有关与该帐户关联的特权的信息，请参见LocalSystem帐户。
'@Requirements
'Minimum supported client       Windows Vista
'Minimum supported server       Windows Server 2008
'Header                         wtsapi32.h
'Library                        Wtsapi32.lib
'dll                            Wtsapi32.dll
Private Declare Function ImpersonateLoggedOnUser Lib "advapi32" (ByVal hToken As Long) As Long
'@原型
'    BOOL WINAPI ImpersonateLoggedOnUser(
'      _In_ HANDLE hToken
'    );
'@功能
'    ImpersonateLoggedOnUser函数允许调用线程模拟登录用户的安全上下文。用户由令牌句柄表示。
'@参数
'    hToken
'    表示已登录用户的主访问令牌或模拟访问令牌的句柄。这可以是通过调用LogonUser、CreateRestrictedToken、DuplicateToken、DuplicateTokenEx、OpenProcessToken或OpenThreadToken函数返回的令牌句柄。如果hToken是主令牌的句柄，则令牌必须具有TOKEN_QUERY和TOKEN_DUPLICATE访问权。如果hToken是模拟令牌的句柄，则令牌必须具有TOKEN_QUERY和TOKEN_IMPERSONATE访问权。
'@返回值
'    如果函数成功，返回值为非零。
'    如果函数失败，返回值为零。要获取扩展的错误信息，请调用GetLastError。
'@备注
'    模拟将持续到线程退出或调用RevertToSelf为止?
'    调用线程不需要具有调用ImpersonateLoggedOnUser的任何特定特权?
'    如果对ImpersonateLoggedOnUser的调用失败，则不模拟客户机连接，并且在流程的安全上下文中发出客户机请求。如果进程以高度特权帐户(如LocalSystem)或作为管理组的成员运行，则用户可能能够执行不允许执行的操作。因此，始终检查调用的返回值是很重要的，如果调用失败，则会引发错误;不要继续执行客户机请求。
'    所有模拟函数，包括ImpersonateLoggedOnUser允许请求的模拟，如果其中一个为真:
'        令牌的请求模拟级别小于SecurityImpersonation，例如SecurityIdentification或securityannamed。
'        调用者具有SeImpersonatePrivilege特权?
'        一个进程(或调用方登录会话中的另一个进程)通过LogonUser或LsaLogonUser函数使用显式凭据创建令牌。
'        经过身份验证的标识与调用者相同?
'    带有SP1和更早版本的Windows XP: 不支持SeImpersonatePrivilege特权?
'    有关模拟的更多信息，请参见客户端模拟。
'@Requirements
'Minimum supported client       Windows XP
'Minimum supported server       Windows Server 2003
'Header                         Advapi32.h
'Library                        Advapi32.lib
'dll                            Advapi32.dll
Private Declare Function RevertToSelf Lib "advapi32" () As Long
'@原型
'    BOOL WINAPI RevertToSelf(void);
'@功能
'    RevertToSelf函数终止对客户机应用程序的模拟。
'@参数
'@返回值
'    如果函数成功，返回值为非零。
'    如果函数失败，返回值为零。要获取扩展的错误信息，请调用GetLastError。
'@备注
'    进程应该在使用DdeImpersonateClient、ImpersonateDdeClientWindow、ImpersonateLoggedOnUser、ImpersonateNamedPipeClient、ImpersonateSelf、ImpersonateAnonymousToken或SetThreadToken函数完成任何模拟之后调用RevertToSelf函数。
'    使用RpcImpersonateClient函数模拟客户机的RPC服务器必须调用RpcRevertToSelf或RpcRevertToSelfEx来结束模拟?
'    如果RevertToSelf失败，应用程序将继续在客户机上下文中运行，这是不合适的。如果RevertToSelf失败，应该关闭进程。
'@Requirements
'Minimum supported client       Windows XP
'Minimum supported server       Windows Server 2003
'Header                         Advapi32.h
'Library                        Advapi32.lib
'dll                            Advapi32.dll
Private Declare Function DuplicateTokenEx Lib "advapi32" (ByVal hExistingToken As Long, ByVal dwDesiredAcces As Long, ByVal lpTokenAttribute As Long, ByVal ImpersonatonLevel As SECURITY_IMPERSONATION_LEVEL, ByVal tokenType As TOKEN_TYPE, phNewToken As Long) As Long
'@原型
'    BOOL WINAPI DuplicateTokenEx(
'      _In_     HANDLE                       hExistingToken,
'      _In_     DWORD                        dwDesiredAccess,
'      _In_opt_ LPSECURITY_ATTRIBUTES        lpTokenAttributes,
'      _In_     SECURITY_IMPERSONATION_LEVEL ImpersonationLevel,
'      _In_     TOKEN_TYPE                   TokenType,
'      _Out_    PHANDLE                      phNewToken
'    );
'@功能
'    DuplicateTokenEx函数创建一个新的访问令牌，该令牌复制一个现有的令牌。此函数可以创建主令牌或模拟令牌。
'@参数
'    hExistingToken
'    访问令牌的句柄，使用TOKEN_DUPLICATE访问打开。
'    dwDesiredAccess
'    指定新令牌的请求访问权限。DuplicateTokenEx函数将请求的访问权限与现有令牌的自由访问控制列表(discretionary access control list, DACL)进行比较，以确定授予或拒绝哪些权限。若要请求与现有令牌相同的访问权限，请指定零。要请求对调用方有效的所有访问权限，请指定MAXIMUM_ALLOWED。
'    有关访问令牌的访问权限列表，请参见访问令牌对象的访问权限。
'    lpTokenAttributes(,可选)
'    指向SECURITY_ATTRIBUTES结构的指针，该结构为新令牌指定安全描述符，并确定子进程是否可以继承该令牌。如果lpTokenAttributes为空，令牌将获得默认的安全描述符，并且不能继承句柄。如果安全描述符包含一个系统访问控制列表(SACL)，令牌将获得ACCESS_SYSTEM_SECURITY访问权，即使在dwDesiredAccess中没有请求它。
'    要在新令牌的安全描述符中设置所有者，调用方的流程令牌必须具有SE_RESTORE_NAME特权集。
'    ImpersonationLevel [在]
'    从SECURITY_IMPERSONATION_LEVEL枚举中指定一个值，该值指示新令牌的模拟级别。
'    tokenType [在]
'    从TOKEN_TYPE枚举中指定下列值之一?
Private Enum TOKEN_TYPE
    TokenPrimary = 1
'       新令牌是您可以在CreateProcessAsUser函数中使用的主要令牌?
    TokenImpersonation = 2
'       新的令牌是模拟令牌?
End Enum
'    phNewToken
'    指向接收新令牌的句柄变量的指针?
'    当您完成使用新令牌时，调用close句柄函数来关闭令牌句柄。
'@返回值
'    如果函数成功，则函数返回一个非零值。
'    如果函数失败，它返回零。要获取扩展的错误信息，请调用GetLastError。
'@备注
'    DuplicateTokenEx函数允许您创建一个可以在CreateProcessAsUser函数中使用的主令牌。这允许模拟客户机的服务器应用程序创建具有客户机安全上下文的流程。注意，DuplicateToken函数只能创建模拟令牌，这对CreateProcessAsUser无效。
'    下面是使用DuplicateTokenEx创建主令牌的典型场景。服务器应用程序创建一个线程，该线程调用一个模拟函数(如ImpersonateNamedPipeClient)来模拟客户机。模拟线程然后调用OpenThreadToken函数来获得它自己的令牌，这是一个具有客户机安全上下文的模拟令牌。线程在调用DuplicateTokenEx时指定这个模拟令牌，指定TokenPrimary标志。DuplicateTokenEx函数创建一个具有客户端安全上下文的主令牌。
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
'@原型
'    typedef enum _SECURITY_IMPERSONATION_LEVEL {
'      SecurityAnonymous,
'      SecurityIdentification,
'      SecurityImpersonation,
'      SecurityDelegation
'    } SECURITY_IMPERSONATION_LEVEL, *PSECURITY_IMPERSONATION_LEVEL;
'@功能
'    SECURITY_IMPERSONATION_LEVEL枚举包含指定安全模拟级别的值。安全模拟级别控制服务器进程可以代表客户机进程进行操作的程度。
'@常量
'    SecurityAnonymous
'    服务器进程不能获取关于客户机的标识信息，也不能模拟客户机。它的定义没有给定值，因此，根据ANSI C规则，默认值为零。
'    SecurityIdentification
'    服务器进程可以获取关于客户机的信息，比如安全标识符和特权，但是它不能模拟客户机。这对于导出自己的对象的服务器非常有用，例如，导出表和视图的数据库产品。使用检索到的客户机安全信息，服务器可以在不使用正在使用客户机安全上下文的其他服务的情况下做出访问验证决策。
'    SecurityImpersonation
'    服务器进程可以在其本地系统上模拟客户机的安全上下文?服务器不能在远程系统上模拟客户机?
'    SecurityDelegation
'    服务器进程可以在远程系统上模拟客户机的安全上下文?
'--------------------------------------------------------------------------------------------------
'方法           RunAsCurrentUser
'功能           在当前会话中创建进程,必须SYSTEM调用
'返回值         Long
'入参列表:
'参数名         类型                    说明
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
'        gobjLog.LogInfo RLL_LogInfo, "OpenThreadToken失败", "错误", GetLastDllErr(Err.LastDllError)
'        If OpenProcessToken(GetCurrentProcess(), TOKEN_ALL_ACCESS, hProcessToken) = 0 Then
'            gobjLog.LogInfo RLL_LogInfo, "OpenProcessToken失败", "错误", GetLastDllErr(Err.LastDllError)
'        End If
'    End If
'    If hProcessToken Then
    If WTSQueryUserToken(glngWinSessionID, hProcessToken) <> 0 Then
        If DuplicateTokenEx(hProcessToken, TOKEN_ALL_ACCESS, ByVal 0, ByVal SecurityImpersonation, ByVal TokenPrimary, hNewToken) <> 0 Then
            If SetTokenInformation(hNewToken, TokenSessionId, glngWinSessionID, LenB(glngWinSessionID)) <> 0 Then
                gobjLog.LogInfo RLL_LogInfo, "是否管理员用户", IsAdministrator(), "hToken=", hNewToken
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
                                gobjLog.LogInfo RLL_LogInfo, "CreateProcessAsUser成功"
                            Else
                                gobjLog.LogInfo RLL_LogInfo, "CreateProcessAsUser失败", "错误", GetLastDllErr(Err.LastDllError)
                            End If
                        Else
                            gobjLog.LogInfo RLL_LogInfo, "ImpersonateLoggedOnUser失败", "错误", GetLastDllErr(Err.LastDllError)
                        End If
                        If RevertToSelf() = 0 Then
                            gobjLog.LogInfo RLL_LogInfo, "RevertToSelf失败", "错误", GetLastDllErr(Err.LastDllError)
                        End If
'                    Else
'                        gobjLog.LogInfo RLL_LogInfo, "SetTokenInformation2失败", "错误", GetLastDllErr(Err.LastDllError)
'                    End If
'                Else
'                    gobjLog.LogInfo RLL_LogInfo, "WTSQueryUserToken失败", "错误", GetLastDllErr(Err.LastDllError)
'                End If
            Else
                gobjLog.LogInfo RLL_LogInfo, "SetTokenInformation1失败", "错误", GetLastDllErr(Err.LastDllError)
            End If
            CloseHandle hNewToken
        Else
            gobjLog.LogInfo RLL_LogInfo, "DuplicateTokenEx失败", "错误", GetLastDllErr(Err.LastDllError)
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
'方法           RunAsUser
'功能           执行以指定管理员运行程序
'返回值         Boolean
'入参列表:
'参数名         类型                    说明
'strUserName    String                  用户名
'strPassword    String                  密码
'strDomainName  String                  账户域
'strApplicationName String              程序路径
'strCommandLine String                  命令行
'strCurrentDirectory    String          当前目录
'-------------------------------------------------------------------------------------------------
Public Function RunAsUser(ByVal strApplicationName As String, ByVal strCommandLine As String, Optional ByVal strCurrentDirectory As String, Optional ByVal strUserName As String, Optional ByVal strPassword As String, Optional ByVal strDomainName As String) As Boolean
    On Error GoTo ErrH
    Call gobjLog.PushMethod(RLL_AllLog, "ZLHelperMain.clsRunas.RunAsUser", strUserName, Sm4EncryptEcb(strPassword), strDomainName, strApplicationName, Sm4EncryptEcb(strCommandLine), strCurrentDirectory)
    If IsWindows2000OrGreater() Then
'        If IsWindowsVistaOrGreater() Then
'            '当前进程不是管理
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
'方法           GetUserSID
'功能           获取用户SID，用户读取注册表
'返回值         String
'入参列表:
'参数名         类型                    说明
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
        gobjLog.LogInfo RLL_LogInfo, "缓冲区不足，分配缓冲区", "返回", lngRet
        ReDim bytSid(cbSid): strDom = String$(cbDom, Chr$(0))
        lngRet = LookupAccountName(0, StrPtr(strAccountName), bytSid(0), cbSid, StrPtr(strDom), cbDom, peUse)
        lngError = Err.LastDllError
        If lngRet = 0 Then
            gobjLog.LogInfo RLL_LogInfo, "LookupAccountName2失败", "错误", GetLastDllErr(Err.LastDllError), "返回", lngRet
        End If
        If ConvertSidToStringSid(bytSid(0), lpStrSid) Then
            strRet = String$(lstrlen(lpStrSid), Chr(0))
            If lstrcpy(StrPtr(strRet), lpStrSid) <> 0 Then
                If InStr(strRet, Chr$(0)) > 0 Then
                    strRet = Mid(strRet, 1, InStr(strRet, Chr$(0)) - 1)
                End If
            Else
                strRet = ""
                gobjLog.LogInfo RLL_LogInfo, "lstrcpy失败", "错误", GetLastDllErr(Err.LastDllError)
            End If
            Call LocalFree(lpStrSid)
        Else
            gobjLog.LogInfo RLL_LogInfo, "ConvertSidToStringSid失败", "错误", GetLastDllErr(Err.LastDllError)
        End If
    ElseIf lngError = 0 Then
        gobjLog.LogInfo RLL_LogInfo, "用户不存在", "返回", lngRet
    Else
        gobjLog.LogInfo RLL_LogInfo, "LookupAccountName失败", "错误", GetLastDllErr(Err.LastDllError), "返回", lngRet
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
''方法           GetCurrentSessionSID
''功能           获取当前会话的SID串
''返回值         String
''入参列表:
''参数名         类型                    说明
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
'            gobjLog.LogInfo RLL_LogInfo, "OpenThreadToken失败", "错误", GetLastDllErr(Err.LastDllError)
'            If OpenProcessToken(GetCurrentProcess(), TOKEN_QUERY, hProcessToken) = 0 Then
'                gobjLog.LogInfo RLL_LogInfo, "OpenProcessToken失败", "错误", GetLastDllErr(Err.LastDllError)
'            End If
'        End If
'    Else
'
'    End If
'    If hProcessToken Then
'        If GetTokenInformation(hProcessToken, ByVal TokenGroups, 0, 0, BufferSize) = 0 Then ' Determine required buffer size
'            gobjLog.LogInfo RLL_LogInfo, "GetTokenInformation失败", "错误", GetLastDllErr(Err.LastDllError), "需要缓冲区大小", BufferSize
'        End If
'        If BufferSize Then
'            ReDim InfoBuffer((BufferSize \ 4) - 1) As Long
'            lResult = GetTokenInformation(hProcessToken, ByVal TokenGroups, InfoBuffer(0), BufferSize, BufferSize)
'            If lResult = 0 Then
'                gobjLog.LogInfo RLL_LogInfo, "GetTokenInformation失败", "错误", GetLastDllErr(Err.LastDllError)
'                Exit Function
'            End If
'            'TOKEN_GROUPS.GROUPCount成员
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
'                                gobjLog.LogInfo RLL_LogInfo, "lstrcpy失败", "错误", GetLastDllErr(Err.LastDllError)
'                            End If
'                            Call LocalFree(ByVal lngStrSID)
'                        Else
'                            gobjLog.LogInfo RLL_LogInfo, "ConvertSidToStringSid失败", "错误", GetLastDllErr(Err.LastDllError)
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
'方法           GetCurrentSessionID
'功能           获取当前会话ID
'返回值         Long
'入参列表:
'参数名         类型                    说明
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
        gobjLog.LogInfo RLL_LogInfo, "ProcessIdToSessionId失败", "错误", GetLastDllErr(Err.LastDllError)
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
'方法           IsProcessRunAsAdmin
'功能           当前进程是以管理员权限运行。管理员权限和管理员不是同一个概念，Visita之上，标准管理员运行是权限不是管理员，必须RUnAS
'返回值         Boolean
'入参列表:
'参数名         类型                    说明
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
            gobjLog.LogInfo RLL_LogInfo, "CheckTokenMembership失败", "错误", GetLastDllErr(Err.LastDllError)
        End If
        Call FreeSid(psidAdmin)
    Else
        gobjLog.LogInfo RLL_LogInfo, "AllocateAndInitializeSid失败", "错误", GetLastDllErr(Err.LastDllError)
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
'方法           IsAdministrator
'功能           判断当前进程用户是否是管理员组用户,必须SYSTEM调用
'返回值         Boolean
'入参列表:
'参数名         类型                    说明
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
                gobjLog.LogInfo RLL_LogInfo, "GetTokenInformation失败", "错误", GetLastDllErr(Err.LastDllError), "需要缓冲区大小", BufferSize
            End If
            If BufferSize Then
                ReDim InfoBuffer((BufferSize \ 4) - 1) As Long
                lResult = GetTokenInformation(hProcessToken, ByVal TokenGroups, InfoBuffer(0), BufferSize, BufferSize)
                If lResult = 0 Then
                    gobjLog.LogInfo RLL_LogInfo, "GetTokenInformation失败", "错误", GetLastDllErr(Err.LastDllError)
                    Exit Function
                End If
                'TOKEN_GROUPS.GROUPCount成员
                ReDim tpTokenGroups(InfoBuffer(0) - 1)
                Call MoveMemory(tpTokenGroups(0), InfoBuffer(1), Len(tpTokenGroups(0)) * InfoBuffer(0))
                lResult = AllocateAndInitializeSid(tpSidAuth, 2, SECURITY_BUILTIN_DOMAIN_RID, DOMAIN_ALIAS_RID_ADMINS, 0, 0, 0, 0, 0, 0, psidAdmin)
                If lResult = 0 Then
                    gobjLog.LogInfo RLL_LogInfo, "AllocateAndInitializeSid失败", "错误", GetLastDllErr(Err.LastDllError)
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
        gobjLog.LogInfo RLL_LogInfo, "WTSQueryUserTokend失败", "错误", GetLastDllErr(Err.LastDllError)
    End If
    Call gobjLog.PopMethod(RLL_AllLog, "ZLHelperMain.clsRunas.IsAdministrator", IsAdministrator)
    Exit Function
ErrH:
    If gobjLog.ErrCenter(RLL_RunError, "ZLHelperMain.clsRunas.IsAdministrator") = 1 Then
        Resume
    End If
End Function

'--------------------------------------------------------------------------------------------------
'方法           IsProcesssAdministrator
'功能           判断当前进程用户是否是管理员组用户
'返回值         Boolean
'入参列表:
'参数名         类型                    说明
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
        gobjLog.LogInfo RLL_LogInfo, "OpenThreadToken失败", "错误", GetLastDllErr(Err.LastDllError)
        If lngProcessID = 0 Then
            If OpenProcessToken(GetCurrentProcess(), TOKEN_QUERY, hProcessToken) = 0 Then
                gobjLog.LogInfo RLL_LogInfo, "OpenProcessToken失败", "错误", GetLastDllErr(Err.LastDllError)
            End If
        End If
    End If
    If hProcessToken Then
        If GetTokenInformation(hProcessToken, ByVal TokenGroups, 0, 0, BufferSize) = 0 Then ' Determine required buffer size
            gobjLog.LogInfo RLL_LogInfo, "GetTokenInformation失败", "错误", GetLastDllErr(Err.LastDllError), "需要缓冲区大小", BufferSize
        End If
        If BufferSize Then
            ReDim InfoBuffer((BufferSize \ 4) - 1) As Long
            lResult = GetTokenInformation(hProcessToken, ByVal TokenGroups, InfoBuffer(0), BufferSize, BufferSize)
            If lResult = 0 Then
                gobjLog.LogInfo RLL_LogInfo, "GetTokenInformation失败", "错误", GetLastDllErr(Err.LastDllError)
                Exit Function
            End If
            'TOKEN_GROUPS.GROUPCount成员
            ReDim tpTokenGroups(InfoBuffer(0) - 1)
            Call MoveMemory(tpTokenGroups(0), InfoBuffer(1), Len(tpTokenGroups(0)) * InfoBuffer(0))
            lResult = AllocateAndInitializeSid(tpSidAuth, 2, SECURITY_BUILTIN_DOMAIN_RID, DOMAIN_ALIAS_RID_ADMINS, 0, 0, 0, 0, 0, 0, psidAdmin)
            If lResult = 0 Then
                gobjLog.LogInfo RLL_LogInfo, "AllocateAndInitializeSid失败", "错误", GetLastDllErr(Err.LastDllError)
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
'方法           GetProcessUserName
'功能           获取当前进程的操作系统的用户名
'返回值         String
'入参列表:
'参数名         类型                    说明
'lngType        Long                    获取名称的格式
'-------------------------------------------------------------------------------------------------
Public Function GetProcessUserName(Optional lngType As Long = NameSamCompatible) As String
    Dim strTemp     As String
    Dim lngLen      As Long

    On Error GoTo ErrH
    Call gobjLog.PushMethod(RLL_AllLog, "ZLHelperMain.clsRunas.GetProcessUserName", lngType)
    lngLen = UNLEN + 1
    strTemp = String(UNLEN + 1, Chr$(0))
    If GetUserName(strTemp, lngLen) = 0 Then
        gobjLog.LogInfo RLL_LogInfo, "GetUserName失败", "错误", GetLastDllErr(Err.LastDllError), "长度", lngLen
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
'方法           CheckUserPassword
'功能           检查操作系统用户是否正确
'返回值         Boolean
'入参列表:
'参数名         类型                    说明
'strUserName    String                  用户名
'strPassword    String                  密码
'strDomainName  String                  账户域
'-------------------------------------------------------------------------------------------------
Public Function CheckUserPassword(ByVal strUserName As String, ByVal strPassword As String, ByVal strDomainName As String) As Boolean
    Dim hToken      As Long
    On Error GoTo ErrH
    Call gobjLog.PushMethod(RLL_AllLog, "ZLHelperMain.clsRunas.CheckUserPassword", strUserName, Sm4EncryptEcb(strPassword), strDomainName)
    CheckUserPassword = LogonUser(strUserName, strDomainName, strPassword, LOGON32_LOGON_INTERACTIVE, LOGON32_PROVIDER_DEFAULT, hToken) <> 0
    If CheckUserPassword Then
        CloseHandle hToken
    Else
        gobjLog.LogInfo RLL_LogInfo, "LogonUser失败", "错误", GetLastDllErr(Err.LastDllError)
    End If
    Call gobjLog.PopMethod(RLL_AllLog, "ZLHelperMain.clsRunas.CheckUserPassword", CheckUserPassword)
    Exit Function
ErrH:
    If gobjLog.ErrCenter(RLL_RunError, "ZLHelperMain.clsRunas.CheckUserPassword") = 1 Then
        Resume
    End If
End Function

'--------------------------------------------------------------------------------------------------
'方法           GetPrivilegeList
'功能           获取进程的权限列表
'返回值         String
'入参列表:
'参数名         类型                    说明
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
            gobjLog.LogInfo RLL_LogInfo, "GetTokenInformation失败", "错误", GetLastDllErr(Err.LastDllError), "需要缓冲区大小", BufferSize
        End If
        If BufferSize Then
            ReDim InfoBuffer((BufferSize \ 4) - 1) As Long
            lngRet = GetTokenInformation(hProcessToken, ByVal TokenPrivileges, InfoBuffer(0), BufferSize, BufferSize)
            If lngRet = 0 Then
                gobjLog.LogInfo RLL_LogInfo, "GetTokenInformation失败", "错误", GetLastDllErr(Err.LastDllError)
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
'方法           EnablePrivilege
'功能           提升进程的权限
'返回值         Boolean
'入参列表:
'参数名         类型                    说明
'
'-------------------------------------------------------------------------------------------------
Public Function EnablePrivilegeTest() As Boolean
    '0: SeIncreaseQuotaPrivilege -0
    If Not EnablePrivilege(, SE_INCREASE_QUOTA_NAME) Then
        gobjLog.LogInfo RLL_LogInfo, SE_INCREASE_QUOTA_NAME & "失败"
    Else
        gobjLog.LogInfo RLL_LogInfo, SE_INCREASE_QUOTA_NAME & "成功"
    End If
    '1: SeSecurityPrivilege -0
    If Not EnablePrivilege(, SE_SECURITY_NAME) Then
        gobjLog.LogInfo RLL_LogInfo, SE_SECURITY_NAME & "失败"
    Else
        gobjLog.LogInfo RLL_LogInfo, SE_SECURITY_NAME & "成功"
    End If
    '2: SeTakeOwnershipPrivilege -0
    If Not EnablePrivilege(, SE_TAKE_OWNERSHIP_NAME) Then
        gobjLog.LogInfo RLL_LogInfo, SE_TAKE_OWNERSHIP_NAME & "失败"
    Else
        gobjLog.LogInfo RLL_LogInfo, SE_TAKE_OWNERSHIP_NAME & "成功"
    End If
    '3: SeLoadDriverPrivilege -0
    If Not EnablePrivilege(, SE_LOAD_DRIVER_NAME) Then
        gobjLog.LogInfo RLL_LogInfo, SE_LOAD_DRIVER_NAME & "失败"
    Else
        gobjLog.LogInfo RLL_LogInfo, SE_LOAD_DRIVER_NAME & "成功"
    End If
    '4: SeSystemProfilePrivilege -0
    If Not EnablePrivilege(, SE_SYSTEM_PROFILE_NAME) Then
        gobjLog.LogInfo RLL_LogInfo, SE_SYSTEM_PROFILE_NAME & "失败"
    Else
        gobjLog.LogInfo RLL_LogInfo, SE_SYSTEM_PROFILE_NAME & "成功"
    End If
    '5: SeSystemtimePrivilege -0
    If Not EnablePrivilege(, SE_SYSTEMTIME_NAME) Then
        gobjLog.LogInfo RLL_LogInfo, SE_SYSTEMTIME_NAME & "失败"
    Else
        gobjLog.LogInfo RLL_LogInfo, SE_SYSTEMTIME_NAME & "成功"
    End If
    '6: SeProfileSingleProcessPrivilege -0
    If Not EnablePrivilege(, SE_PROF_SINGLE_PROCESS_NAME) Then
        gobjLog.LogInfo RLL_LogInfo, SE_PROF_SINGLE_PROCESS_NAME & "失败"
    Else
        gobjLog.LogInfo RLL_LogInfo, SE_PROF_SINGLE_PROCESS_NAME & "成功"
    End If
    '7: SeIncreaseBasePriorityPrivilege -0
    If Not EnablePrivilege(, SE_INC_BASE_PRIORITY_NAME) Then
        gobjLog.LogInfo RLL_LogInfo, SE_INC_BASE_PRIORITY_NAME & "失败"
    Else
        gobjLog.LogInfo RLL_LogInfo, SE_INC_BASE_PRIORITY_NAME & "成功"
    End If
    '8: SeCreatePagefilePrivilege -0
    If Not EnablePrivilege(, SE_CREATE_PAGEFILE_NAME) Then
        gobjLog.LogInfo RLL_LogInfo, SE_CREATE_PAGEFILE_NAME & "失败"
    Else
        gobjLog.LogInfo RLL_LogInfo, SE_CREATE_PAGEFILE_NAME & "成功"
    End If
    '9: SeBackupPrivilege -0
    If Not EnablePrivilege(, SE_BACKUP_NAME) Then
        gobjLog.LogInfo RLL_LogInfo, SE_BACKUP_NAME & "失败"
    Else
        gobjLog.LogInfo RLL_LogInfo, SE_BACKUP_NAME & "成功"
    End If
    
    '10: SeRestorePrivilege -0
    If Not EnablePrivilege(, SE_RESTORE_NAME) Then
        gobjLog.LogInfo RLL_LogInfo, SE_RESTORE_NAME & "失败"
    Else
        gobjLog.LogInfo RLL_LogInfo, SE_RESTORE_NAME & "成功"
    End If
    'AAA -11: SeShutdownPrivilege -0
    '12: SeDebugPrivilege -0
    If Not EnablePrivilege(, SE_DEBUG_NAME) Then
        gobjLog.LogInfo RLL_LogInfo, SE_DEBUG_NAME & "失败"
    Else
        gobjLog.LogInfo RLL_LogInfo, SE_DEBUG_NAME & "成功"
    End If
    '13: SeSystemEnvironmentPrivilege -0
    If Not EnablePrivilege(, SE_SYSTEM_ENVIRONMENT_NAME) Then
        gobjLog.LogInfo RLL_LogInfo, SE_SYSTEM_ENVIRONMENT_NAME & "失败"
    Else
        gobjLog.LogInfo RLL_LogInfo, SE_SYSTEM_ENVIRONMENT_NAME & "成功"
    End If
    'AAA -14: SeChangeNotifyPrivilege -3
    '15: SeRemoteShutdownPrivilege -0
    If Not EnablePrivilege(, SE_REMOTE_SHUTDOWN_NAME) Then
        gobjLog.LogInfo RLL_LogInfo, SE_REMOTE_SHUTDOWN_NAME & "失败"
    Else
        gobjLog.LogInfo RLL_LogInfo, SE_REMOTE_SHUTDOWN_NAME & "成功"
    End If
    'AAA -16: SeUndockPrivilege -0
    '17: SeManageVolumePrivilege -0
    If Not EnablePrivilege(, SE_MANAGE_VOLUME_NAME) Then
        gobjLog.LogInfo RLL_LogInfo, SE_MANAGE_VOLUME_NAME & "失败"
    Else
        gobjLog.LogInfo RLL_LogInfo, SE_MANAGE_VOLUME_NAME & "成功"
    End If
    '18: SeImpersonatePrivilege -3
    If Not EnablePrivilege(, SE_IMPERSONATE_NAME) Then
        gobjLog.LogInfo RLL_LogInfo, SE_IMPERSONATE_NAME & "失败"
    Else
        gobjLog.LogInfo RLL_LogInfo, SE_IMPERSONATE_NAME & "成功"
    End If
    '19: SeCreateGlobalPrivilege -3
    If Not EnablePrivilege(, SE_CREATE_GLOBAL_NAME) Then
        gobjLog.LogInfo RLL_LogInfo, SE_CREATE_GLOBAL_NAME & "失败"
    Else
        gobjLog.LogInfo RLL_LogInfo, SE_CREATE_GLOBAL_NAME & "成功"
    End If
    'AAA -20: SeIncreaseWorkingSetPrivilege -0
    'AAA -21: SeTimeZonePrivilege -0
    '22: SeCreateSymbolicLinkPrivilege -0
    If Not EnablePrivilege(, SE_CREATE_SYMBOLIC_LINK_NAME) Then
        gobjLog.LogInfo RLL_LogInfo, SE_CREATE_SYMBOLIC_LINK_NAME & "失败"
    Else
        gobjLog.LogInfo RLL_LogInfo, SE_CREATE_SYMBOLIC_LINK_NAME & "成功"
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
        gobjLog.LogInfo RLL_LogInfo, "OpenProcessToken失败", "错误", GetLastDllErr(Err.LastDllError)
    End If
    If hToken <> 0 Then
        lngRet = LookupPrivilegeValue(vbNullString, strPrivilegeName, tmpLuid)
        If lngRet = 0 Then
            gobjLog.LogInfo RLL_LogInfo, "LookupPrivilegeValue失败", "错误", GetLastDllErr(Err.LastDllError)
        End If
        tkp.PrivilegeCount = 1
        tkp.Privileges(0).PLUID = tmpLuid
        tkp.Privileges(0).Attributes = SE_PRIVILEGE_ENABLED
        lngRet = AdjustTokenPrivileges(hToken, 0, tkp, Len(tkp), tkpNewButIgnored, lBufferNeeded)
        If lngRet = 0 Then
            gobjLog.LogInfo RLL_LogInfo, "AdjustTokenPrivileges失败", "错误", GetLastDllErr(Err.LastDllError)
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
        gobjLog.LogInfo RLL_LogInfo, "LookupPrivilegeValue失败", "错误", GetLastDllErr(Err.LastDllError)
    End If
    tkp.PrivilegeCount = 1
    tkp.Privileges(0).PLUID = tmpLuid
    tkp.Privileges(0).Attributes = SE_PRIVILEGE_ENABLED
    lngRet = AdjustTokenPrivileges(hToken, 0, tkp, Len(tkp), tkpNewButIgnored, lBufferNeeded)
    If lngRet = 0 Then
        gobjLog.LogInfo RLL_LogInfo, "AdjustTokenPrivileges失败", "错误", GetLastDllErr(Err.LastDllError)
    End If
    EnablePrivilegeToken = lngRet <> 0
End Function
'--------------------------------------------------------------------------------------------------
'方法           RunAsAdmin
'功能           以管理员运行程序
'返回值         long
'入参列表:
'参数名         类型                    说明
'strApplicationName String              程序
'strCommandLine      String              命令行
'strDirectory        String              当前程序目录
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
        gobjLog.LogInfo RLL_LogInfo, "ShellExecute失败", "错误", GetLastDllErr(Err.LastDllError), "返回", lngRet
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
'方法           IsWindows2000OrGreater
'功能           是否是Window2000之后版本
'返回值         Boolean
'入参列表:
'参数名         类型                    说明
'
'-------------------------------------------------------------------------------------------------
Private Function IsWindows2000OrGreater() As Boolean
    Dim osInfo          As OSVERSIONINFOEX
    On Error GoTo ErrH
    Call gobjLog.PushMethod(RLL_AllLog, "ZLHelperMain.clsRunas.IsWindows2000OrGreater")
    osInfo.dwOSVersionInfoSize = Len(osInfo)
    osInfo.szCSDVersion = Space$(128)
    If GetVersionExA(osInfo) = 0 Then
        gobjLog.LogInfo RLL_LogInfo, "GetVersionExA失败", "错误", GetLastDllErr(Err.LastDllError)
    End If
    IsWindows2000OrGreater = osInfo.dwMajorVersion >= 5
    Call gobjLog.PopMethod(RLL_AllLog, "ZLHelperMain.clsRunas.IsWindows2000OrGreater", IsWindows2000OrGreater)
    Exit Function
ErrH:
    If gobjLog.ErrCenter(RLL_RunError, "ZLHelperMain.clsRunas.IsWindows2000OrGreater") = 1 Then
        Resume
    End If
End Function

'方法           IsWindowsVistaOrGreater
'功能           是否是Window Vista之后版本
'返回值         Boolean
'入参列表:
'参数名         类型                    说明
'
'-------------------------------------------------------------------------------------------------
Private Function IsWindowsVistaOrGreater() As Boolean
    Dim osInfo          As OSVERSIONINFOEX
    On Error GoTo ErrH
    Call gobjLog.PushMethod(RLL_AllLog, "ZLHelperMain.clsRunas.IsWindowsVistaOrGreater")
    osInfo.dwOSVersionInfoSize = Len(osInfo)
    osInfo.szCSDVersion = Space$(128)
    If GetVersionExA(osInfo) = 0 Then
        gobjLog.LogInfo RLL_LogInfo, "GetVersionExA失败", "错误", GetLastDllErr(Err.LastDllError)
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
'方法           RunAsUserW2K
'功能           在Windows2000以上执行Runas
'返回值         Long                    返回进程ID
'入参列表:
'参数名         类型                    说明
'strUserName    String                  用户名
'strPassword    String                  密码
'strDomainName  String                  账户域
'strApplicationName String              程序路径
'strCommandLine String                  命令行
'strCurrentDirectory    String          当前目录
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
        gobjLog.LogInfo RLL_LogInfo, "CreateProcessWithLogonW失败", "错误", GetLastDllErr(Err.LastDllError)
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
'方法           RunAsUserNT4
'功能           在Window2000以下执行Runas
'返回值         Long                    返回进程ID
'入参列表:
'参数名         类型                    说明
'strUserName    String                  用户名
'strPassword    String                  密码
'strDomainName  String                  账户域
'strApplicationName String              程序路径
'strCommandLine String                  命令行
'strCurrentDirectory    String          当前目录
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
            gobjLog.LogInfo RLL_LogInfo, "CreateProcessAsUser失败", "错误", GetLastDllErr(Err.LastDllError)
            CloseHandle hToken
        End If
        
        CloseHandle hToken
        CloseHandle piInfo.hThread
        CloseHandle piInfo.hProcess
    Else
        gobjLog.LogInfo RLL_LogInfo, "LogonUser失败", "错误", GetLastDllErr(Err.LastDllError)
    End If
    Call gobjLog.PopMethod(RLL_AllLog, "ZLHelperMain.clsRunas.RunAsUserNT4", RunAsUserNT4)
    Exit Function
ErrH:
    If gobjLog.ErrCenter(RLL_RunError, "ZLHelperMain.clsRunas.RunAsUserNT4") = 1 Then
        Resume
    End If
End Function

