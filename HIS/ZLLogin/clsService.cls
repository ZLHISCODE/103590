VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'==================================================================================================
'编写           lshuo
'日期           2019/3/7
'模块           mdlService
'说明           该模块创建一个服务以及进行服务的一般控制。
'==================================================================================================
'服务启动类型
Public Enum SERVICE_START_TYPE
    SERVICE_BOOT_START = &H0            '用于由系统加载器创建的设备驱动程序。只能用于驱动服务程序.
    SERVICE_SYSTEM_START = &H1          '用于由IoInitSystem函数创建的设备驱动程序。
    SERVICE_AUTO_START = &H2            '系统启动时由服务控制管理器自动启动该服务程序。
    SERVICE_DEMAND_START = &H3          '由服务控制管理器(SCM)启动的服务。
    SERVICE_DISABLED = &H4              '表示该服务不可启动。
End Enum
Private Const Delete As Long = &H10000
Public Enum SERVICE_TYPE
    SERVICE_ADAPTER = &H4               '保留
    SERVICE_FILE_SYSTEM_DRIVER = &H2    '文件系统驱动服务程序
    SERVICE_KERNEL_DRIVER = &H1         '驱动服务程序
    SERVICE_RECOGNIZER_DRIVER = &H8     '保留
    SERVICE_WIN32_OWN_PROCESS = &H10    '运行于独立进程的服务程序
    SERVICE_WIN32_SHARE_PROCESS = &H20  '被多个进程共享的服务程序
    SERVICE_INTERACTIVE_PROCESS = &H100  '该服务可以与桌面程序进行交互操作.若使用了SERVICE_WIN32_OWN_PROCESS 或 SERVICE_WIN32_SHARE_PROCESS且使用LocalSystem帐号来运行该服务程序，则还该值作为附加值
End Enum
Private Const STANDARD_RIGHTS_REQUIRED      As Long = &HF0000 ' masks for the predefined standard access types
Private Declare Function LockServiceDatabase Lib "advapi32.dll" (ByVal hSCManager As Long) As Long
'功能
'    锁定指定的数据库?
'参数
'【hSCManager】
'标识服务控制管理器数据库。OpenSCManager函数返回此句柄，该句柄必须具有SC_MANAGER_LOCK访问权限。
'
'返回值
'如果函数成功，则返回值是指定的服务控制管理器数据库的锁定。
'如果函数失败，返回值为NULL。要获取扩展错误信息，请调用GetLastError.
'
'错误
'以下错误代码可能由服务控制管理器设置?其他错误代码可以由服务控制管理器调用的注册表功能设置?
'值 含义
'ERROR_ACCESS_DENIED 没有使用SC_MANAGER_LOCK访问来打开指定的句柄?
'ERROR_INVALID_HANDLE 指定的句柄无效?
'ERROR_SERVICE_DATABASE_LOCKED 数据库被锁定?
'
'备注
'LockServiceDatabase函数尝试获取指定数据库上的锁?一次只能有一个进程可以锁定数据库?
'锁是安装和配置程序和服务控制管理器使用的协议，用于序列化对注册表中服务树的访问。服务控制管理员唯一获取锁的时间是启动服务时。在使用ChangeServiceConfig或SetServiceObjectSecurity功能重新配置服务之前，预期安装和配置程序将获取锁定。在使用注册表功能重新配置服务之前，他们也应该获取一个锁。该锁可防止服务控制管理器在重新配置时启动服务。
'调用StartService功能以在锁定的数据库中启动服务失败?所有其他服务控制管理器功能不受锁定的影响?
'在随后调用UnlockServiceDatabase功能之前，锁定被保持，直到指定了SC_LOCK句柄。如果持有锁的进程终止，则服务控制管理器将自动清除并释放锁。
Private Declare Function UnlockServiceDatabase Lib "advapi32.dll" (ByVal ScLock As Long) As Long
'功能
'通过释放指定的锁来解锁服务控制管理器数据库?
'
'参数
'【ScLock】
'指定从以前调用LockServiceDatabase函数获取的锁?
'
'
'返回值
'如果函数成功，返回值不为零。
'如果函数失败，返回值为零。要获取扩展错误信息，请调用GetLastError.
'
'错误
'服务控制管理器可以设置以下错误代码?其他错误代码可以由服务控制管理器调用的注册表函数设置?
'
'值 含义
'ERROR_INVALID_SERVICE_LOCK 指定的锁无效?


Private Declare Function OpenSCManager Lib "advapi32.dll" Alias "OpenSCManagerW" (ByVal lpMachineName As String, ByVal lpDatabaseName As String, ByVal dwDesiredAccess As Long) As Long
'返回值：函数执行成功则返回一个指向service control manager database的句柄,失败则返回NULL。
'lpszMachineName:指定计算机名 , 若为空则指定为本机.
'lpszDatabaseName:指定要打开的service control manager database,默认为空。
'dwDesiredAccess: 指定操作的权限, Service Control Manager object specific access types
Public Enum SERVICE_MANAGER_ACCESS_TYPE
    SC_MANAGER_CONNECT = &H1                '允许连接service control manager
    SC_MANAGER_CREATE_SERVICE = &H2         '允许创建服务对象并把它加入service control manager database
    SC_MANAGER_ENUMERATE_SERVICE = &H4      '允许枚举service control manager database中的服务
    SC_MANAGER_LOCK = &H8                   '允许锁住service control manager database
    SC_MANAGER_QUERY_LOCK_STATUS = &H10     '允许获取servicecontrolmanagerdatabase的封锁信息
    SC_MANAGER_MODIFY_BOOT_CONFIG = &H20
    SC_MANAGER_ALL_ACCESS = (STANDARD_RIGHTS_REQUIRED Or SC_MANAGER_CONNECT Or SC_MANAGER_CREATE_SERVICE Or SC_MANAGER_ENUMERATE_SERVICE Or SC_MANAGER_ENUMERATE_SERVICE Or SC_MANAGER_LOCK Or SC_MANAGER_QUERY_LOCK_STATUS Or SC_MANAGER_MODIFY_BOOT_CONFIG)
End Enum
Private Declare Function OpenService Lib "advapi32.dll" Alias "OpenServiceW" (ByVal hSCManager As Long, ByVal lpServiceName As Long, ByVal dwDesiredAccess As Long) As Long
'返回值：打开指定的Service,函数执行成功则返回指向某项服务的句柄,失败则返回NULL。
'schSCManager:指向service control manager database的句柄,由OpenSCManager函数返回。
'lpszServiceName:要打开的服务的名字,注意大小写。
'fdwDesiredAccess:指定操作的权限,Service object specific access type
Public Enum SERVICE_ACCESS_TYPE
    SERVICE_QUERY_CONFIG = &H1              '允许获取服务配置
    SERVICE_CHANGE_CONFIG = &H2             '允许更改服务的配置
    SERVICE_QUERY_STATUS = &H4              '允许通过访问service control manager获取服务状态
    SERVICE_ENUMERATE_DEPENDENTS = &H8      '允许获取依赖于该服务的其他服务
    SERVICE_START = &H10                    '允许启动服务
    SERVICE_STOP = &H20                     '允许停止服务
    SERVICE_PAUSE_CONTINUE = &H40           '允许暂停和唤醒服务
    SERVICE_INTERROGATE = &H80              '允许立即获取服务状态
    SERVICE_USER_DEFINED_CONTROL = &H100    '允许用户指定特殊的服务控制码
    SERVICE_ALL_ACCESS = (STANDARD_RIGHTS_REQUIRED Or SERVICE_QUERY_CONFIG Or SERVICE_CHANGE_CONFIG Or SERVICE_QUERY_STATUS Or SERVICE_ENUMERATE_DEPENDENTS Or SERVICE_START Or SERVICE_STOP Or SERVICE_PAUSE_CONTINUE Or SERVICE_INTERROGATE Or SERVICE_USER_DEFINED_CONTROL)
End Enum
Private Declare Function CreateService Lib "advapi32.dll" Alias "CreateServiceW" (ByVal hSCManager As Long, ByVal lpServiceName As Long, ByVal lpDisplayName As Long, ByVal dwDesiredAccess As Long, ByVal dwServiceType As Long, ByVal dwStartType As Long, ByVal dwErrorControl As Long, ByVal lpBinaryPathName As Long, ByVal lpLoadOrderGroup As String, lpdwTagId As Long, ByVal lpDependencies As String, ByVal lp As String, ByVal lpPassword As String) As Long
'返回值：创建一个服务对象，并将其添加到指定的服务控制管理器数据库.如果函数成功，返回值将是该服务的句柄。如果函数失败，则返回值为 NULL.
'hSCManager:服务控制管理程序维护的登记数据库的句柄，由系统函数OpenSCManager 返回,并且必须 具有 SC_MANAGER_CREATE_SERVICE 的访问权限
'lpServiceName:以NULL 结尾的服务名，用于创建登记数据库中的关键字.要安装该服务的名称。 最大字符串长度为 256 个字符。 服务控制管理器数据库将保留在的字符的大小写，但是服务名称比较总是区分大小写。 正斜杠和一个反斜线不是有效的服务名称字符。
'lpDisplayName:对被用户界面程序用来识别服务的显示名称。 此字符串具有最大长度为 256 个字符。 服务控制管理器中的情况下保留名称。 显示名称比较总是不区分大小写。
'dwDesiredAccess:指定服务返回类型.对服务的访问。 请求的访问之前，系统将检查调用进程的访问令牌。 一个值列表请参阅服务安全和访问权限。
'dwServiceType:服务类型。 SERVICE_TYPE
'dwStartType:服务启动选项。SERVICE_START_TYPE
'dwErrorControl:当该启动服务失败时产生错误的严重程度以及采取的保护措施。Error control type
Public Enum SERVICE_ERROR_CONTROL
    SERVICE_ERROR_IGNORE = &H0              '服务启动程序将把该错误记录到事件日志中。 若最后一次正确配置可用，服务启动程序将以最后一次正确配置重新启动;否则将退出执行。
    SERVICE_ERROR_NORMAL = &H1              '服务启动程序将忽略该错误并返回继续执行。
    SERVICE_ERROR_SEVERE = &H2              '服务启动程序将把该错误记录到事件日志中并返回继续执行。
    SERVICE_ERROR_CRITICAL = &H3            '服务启动程序将把该错误记录到事件日志中。若最后一次正确配置可用，服务启动程序将以最后一次正确配置重新启动;否则将返回继续执行。
End Enum

'lpBinaryPathName:服务程序二进制文件，完全限定路径。 如果路径中包含空格它必须被引用，以便它正确的解析。 例如"d:\myshare\myservice.exe"应指定为""d:\myshare\myservice.exe""。该路径也可以包含一个自动启动服务的参数。 例如"d:\myshare\myservice.exe arg1 arg2"。 这些参数被传递给服务的入口点通常主要作用。
'                 如果另一台计算机上指定一个路径，共享必须可以访问由本地计算机的计算机帐户，因为这是在远程调用中使用的安全上下文。 但是，这项规定会影响本地计算机远程计算机中允许任何潜在的漏洞。 因此，最好使用一个本地文件。
'lpLoadOrderGroup:在加载顺序此服务所属的组的名称。 指定 NULL 或空字符串，如果该服务不属于组。 启动程序使用加载顺序组加载服务相对于其他组的指定顺序组。 加载顺序组列表包含在下面的注册表值:[HKEY_LOCAL_MACHINE\System\CurrentControlSet\Control\ServiceGroupOrder]
'lpdwTagId:指向接收 lpLoadOrderGroup 参数中指定的组中唯一的标记值的变量。 如果不更改现有的标签，请指定 NULL。订购服务启动一个加载顺序通过在以下注册表值中指定标记顺序矢量的组中，您可以使用的标记:[HKEY_LOCAL_MACHINE\System\CurrentControlSet\Control\GroupOrderList]。标记只能用于具有 SERVICE_BOOT_START 或启动类型的 SERVICE_SYSTEM_START 驱动程序服务。
'lpDependencies:空分隔名称的服务或加载顺序组系统必须在这个服务开始之前的双空终止数组的指针。 如果服务没有任何依赖关系，请指定为 NULL 或空字符串。 一组依赖意味着如果至少一个组的成员运行后试图启动组的所有成员，可以运行此服务。当需要区分各服务由于服务和服务组共享相同的命名空间前缀与 SC_GROUP_IDENTIFIER 的组名时必须加上此参数。
'lpServiceStartName:该服务应在其下运行的帐户的名称。 如果SERVICE_WIN32_OWN_PROCESS 服务类型DomainNameUserName窗体中使用一个帐户名。 服务进程将被作为此用户登录。 如果该帐户属于内置域，您可以指定.UserName。
'                   如果为NULL此参数CreateService使用LocalSystem帐户。 如果服务类型指定 SERVICE_INTERACTIVE_PROCESS，该服务必须运行在LocalSystem帐户。
'                   如果NTAUTHORITYLocalService此参数CreateService使用LocalService帐户? 如果NTAUTHORITYNetworkService参数CreateService使用NetworkService帐户共享的处理可以运行于任何用户.
'                   如果服务类型 SERVICE_KERNEL_DRIVER 或 SERVICE_FILE_SYSTEM_DRIVER 名称是系统用来加载设备驱动程序的驱动程序对象名称。 如果该驱动程序将使用由IO系统创建的默认对象名称，请指定NULL。
'                   可以将服务配置为使用一个托管的帐户或虚拟帐户。 如果该服务被配置为使用一个托管的服务帐户，名称是托管的服务帐户名称。 如果该服务被配置为使用一个虚拟帐户，作为NTSERVICEServiceName指定名称。 有关托管的服务帐户和虚拟帐户的详细信息请参阅服务帐户的分步指南。
'                   WindowsServer2008、WindowsVista、WindowsServer2003 和WindowsXP2000:Windows7 和WindowsServer2008 R2 之前不支持托管的服务帐户和虚拟帐户。
'lpPassword:由lpServiceStartName参数指定的帐户名的密码。 如果该帐户没有密码，或者如果在LocalService，NetworkService或LocalSystem帐户中运行该服务，请指定一个空字符串。 更多的信息请参阅服务记录列表。
'           如果托管的服务帐户或虚拟帐户名称的名称由lpServiceStartName参数指定的帐户名称lpPassword参数必须为NULL,密码将被忽略，对驱动程序服务。
'注意事项：CreateService函数创建一个服务对象并将其安装在SCM数据库中，同时在注册表以下位置中创建同名键值:[HKEY_LOCAL_MACHINE\System\CurrentControlSet\Services]
'           由CreateService?ChangeServiceConfig和ChangeServiceConfig2等函数创建服务时提供的服务相关信息就存放于这个键值中.以下是一些键值的描述:
'               DependOnGroup:由lpDependencies参数指定的依赖性服务群组，包含加载顺序。
'               DependOnService:由lpDependencies参数指定的依赖性服务.
'               Description:由ChangeServiceConfig2函数写入的描述信息.
'               DisplayName:由lpDisplayName参数指定的服务显示名称.
'               ErrorControl:由dwErrorControl参数指定的纠错方案.
'               FailureActions:由ChangeServiceConfig2函数写入的异常信息.
'               Group:由lpLoadOrderGroup参数指定的服务群组，包含加载顺序。
'               ImagePath:由lpBinaryPathName参数指定的服务程序文件位置.
'               ObjectName:由lpServiceStartName参数指定的名称.
'               Start:由dwStartType参数指定的服务启动时间.
'               Tag:由lpdwTagId参数指定的识别信息?
'               Type:由dwServiceType参数指定的服务类型.
'       服务安装程序可以创建子键来提供更多额外的信息.
'       CreateService函数返回的句柄只能由调用它的进程使用，可以调用CloseServiceHandle函数来关掉这个句柄。若创建进程共享服务，请不要随意调用可能影响其他进程的函数，如ExitProcess。另外，请不要卸载服务程序的动态链接库(DLL文件)。
Private Declare Function ChangeServiceConfig2 Lib "advapi32" Alias "ChangeServiceConfig2W" (ByVal hService As Long, ByVal dwInfoLevel As Long, lpInfo As Any) As Long
'返回值：如果函数成功，返回值为非零。如果函数失败，返回值为零。要得到扩展的错误信息，请调用GetLastError。
'hService:服务句柄。该句柄由OpenService或CreateService函数返回，并且必须具有SERVICE_CHANGE_CONFIG访问权。有关更多信息，请参见服务安全性和访问权限
'dwInfoLevel:要更改的配置信息。Info levels for ChangeServiceConfig2 and QueryServiceConfig2
Public Enum SERVICE_CONFIG_LEVEL
    SERVICE_CONFIG_DESCRIPTION = &H1                    'lpInfo参数是一个指向SERVICE_DESCRIPTION结构的指针。
    SERVICE_CONFIG_FAILURE_ACTIONS = &H2                'lpInfo参数是一个指向SERVICE_FAILURE_ACTIONS结构的指针。如果指定SC_ACTION_REBOOT，调用者必须具有SE_SHUTDOWN_NAME特权。有关更多信息，请参见具有特权的运行。
    SERVICE_CONFIG_DELAYED_AUTO_START_INFO = &H3
    SERVICE_CONFIG_FAILURE_ACTIONS_FLAG = &H4
    SERVICE_CONFIG_SERVICE_SID_INFO = &H5
    SERVICE_CONFIG_REQUIRED_PRIVILEGES_INFO = &H6
    SERVICE_CONFIG_PRESHUTDOWN_INFO = &H7
    SERVICE_CONFIG_TRIGGER_INFO = &H8
    SERVICE_CONFIG_PREFERRED_NODE = &H9
End Enum
'lpInfo:指向用于配置信息的新值的指针。此数据的格式取决于dwInfoLevel参数的值。如果此值为NULL，则信息将保持不变。
'注意事项：ChangeServiceConfig2函数在服务控制管理器数据库中更改指定服务的可选配置信息。您可以通过使用QueryServiceConfig2函数获得当前可选的配置信息。
'           您不能为共享服务控制管理器进程的服务设置SERVICE_CONFIG_FAILURE_ACTIONS值。这包括所有可执行图像为“services.exe”的服务。
'           您可以分别使用ChangeServiceConfig和QueryServiceConfig函数来更改和查询额外的配置信息.
Private Declare Function CloseServiceHandle Lib "advapi32.dll" (ByVal hSCObject As Long) As Long
'返回值：如果函数成功，返回值为非零。如果函数失败，返回值为零。要得到扩展的错误信息，请调用GetLastError。下面的错误代码可以由服务控制管理器设。其他错误代码可以由服务控制管理器调用的注册表函数来设置。
Private Const ERROR_INVALID_HANDLE      As Long = &H6     '指定的句柄无效。
'hSCObject:处理服务控制管理器对象或服务对象的句柄。对于服务控制管理器对象的句柄是由OpenSCManager函数返回的，而服务对象的句柄则由OpenService或CreateService函数返回。
'注意事项：CloseServiceHandle函数不会破坏由句柄引用的服务控制管理器对象。不能销毁服务控制管理器对象。可以通过调用DeleteService函数来销毁服务对象
Private Declare Function ControlService Lib "advapi32.dll" (ByVal hService As Long, ByVal dwControl As Long, lpServiceStatus As SERVICE_STATUS) As Long
'说明：向服务发送控制代码。要在停止服务时指定附加信息，请使用ControlServiceEx函数。
'返回：如果函数成功，返回值为非零。如果函数失败，返回值为零。要得到扩展的错误信息，请调用GetLastError。下面的错误代码可以由服务控制管理器设置?其他错误代码可以由服务控制管理器调用的注册表函数设置。
Private Const ERROR_ACCESS_DENIED   As Long = &H5                '句柄没有所需的访问权限。
Private Const ERROR_DEPENDENT_SERVICES_RUNNING As Long = &H41B   '服务不能停止，因为其他运行服务依赖于它。
'ERROR_INVALID_HANDLE 使用CreateService或OpenService无法获得指定的句柄，或句柄不再有效
'ERROR_INVALID_PARAMETER 请求的控制代码是未定义的
Private Const RROR_INVALID_SERVICE_CONTROL As Long = &H41C       '请求的控制代码无效，或者服务无法接受。
Private Const ERROR_SERVICE_CANNOT_ACCEPT_CTRL As Long = &H425   '请求的控制代码不能被发送到服务，因为服务的状态是service_stop、SERVICE_START_PENDING或SERVICE_STOP_PENDING。
Private Const ERROR_SERVICE_NOT_ACTIVE As Long = &H426           '服务尚未启动。
Private Const ERROR_SERVICE_REQUEST_TIMEOUT As Long = &H41D      '服务的进程开始了，但它并没有调用StartServiceCtrlDispatcher，或者调用StartServiceCtrlDispatcher的线程可能在控制处理程序函数中被阻塞。
Private Const ERROR_SHUTDOWN_IN_PROGRESS As Long = &H45B         '系统正在关闭。
'hService:服务的句柄。该句柄由OpenService或CreateService函数返回。此句柄所需的访问权限取决于所请求的dwControl代码。
'dwControl:这个参数可以是下列控制码之一。
Private Enum SERVICE_CONTROL_CODE
    SERVICE_CONTROL_STOP = &H1                      '通知服务应该停止。hService句柄必须具有SERVICE_STOP访问权。在将停止请求发送到服务之后，您不应该将其他控件发送到服务。
    SERVICE_CONTROL_PAUSE = &H2                     '通知服务应该暂停。hService句柄必须具有SERVICE_PAUSE_CONTINUE访问权。
    SERVICE_CONTROL_CONTINUE = &H3                  '通知暂停的服务，它应该恢复。hService句柄必须具有SERVICE_PAUSE_CONTINUE访问权。
    SERVICE_CONTROL_INTERROGATE = &H4               '通知服务，它应该向服务控制管理器报告其当前状态信息。hService句柄必须具有service_INTERROGATE。请注意，这种控制通常不太有用，因为SCM知道服务的当前状态。
    SERVICE_CONTROL_SHUTDOWN = &H5
    SERVICE_CONTROL_PARAMCHANGE = &H6               '通知服务，它的启动参数已经更改。hService句柄必须具有SERVICE_PAUSE_CONTINUE访问权。
    SERVICE_CONTROL_NETBINDADD = &H7                '通知网络服务有一个新的绑定组件。hService句柄必须具有SERVICE_PAUSE_CONTINUE访问权。但是，这种控制代码已经被弃用，取而代之的是使用插件和播放功能
    SERVICE_CONTROL_NETBINDREMOVE = &H8             '通知网络服务，绑定的组件已经被删除。hService句柄必须具有SERVICE_PAUSE_CONTINUE访问权。但是，这种控制代码已经被弃用，取而代之的是使用插件和播放功能。
    SERVICE_CONTROL_NETBINDENABLE = &H9             '通知网络服务，禁用了禁用绑定。hService句柄必须具有SERVICE_PAUSE_CONTINUE访问权。但是，这种控制代码已经被弃用，取而代之的是使用插件和播放功能。
    SERVICE_CONTROL_NETBINDDISABLE = &HA            '通知一个网络服务，其中一个绑定已被禁用。hService句柄必须具有SERVICE_PAUSE_CONTINUE访问权。但是，这种控制代码已经被弃用，取而代之的是使用插件和播放功能。
    SERVICE_CONTROL_DEVICEEVENT = &HB
    SERVICE_CONTROL_HARDWAREPROFILECHANGE = &HC
    SERVICE_CONTROL_POWEREVENT = &HD
    SERVICE_CONTROL_SESSIONCHANGE = &HE
    SERVICE_CONTROL_PRESHUTDOWN = &HF
    SERVICE_CONTROL_TIMECHANGE = &H10
    SERVICE_CONTROL_TRIGGEREVENT = &H20
    '这个值也可以是用户定义的控制代码， 128-255 服务定义与控制代码相关联的操作。hService句柄必须具有SERVICE_USER_DEFINED_CONTROL访问权。
End Enum
'lpServiceStatus:一个指向SERVICE_STATUS结构的指针，它接收最新的服务状态信息。返回的信息反映了服务向服务控制管理器报告的最新状态。只有当ControlService返回以下错误代码中的一个时，服务控制管理器才会填充这个结构:NO_ERROR、ERROR_INVALID_SERVICE_CONTROL、ERROR_SERVICE_CANNOT_ACCEPT_CTRL或ERROR_SERVICE_NOT_ACTIVE。否则，结构就没有填充
Private Type SERVICE_STATUS
    dwServiceType               As Long '服务的类型。SERVICE_TYPE。
    dwCurrentState              As Long '服务的当前状态。SERVICE_STATE
    dwControlsAccepted          As Long
    dwWin32ExitCode             As Long '服务用于报告发生在启动或停止时发生的错误的错误代码。返回一个错误代码特定于服务,服务必须设置这个值ERROR_SERVICE_SPECIFIC_ERROR表明dwServiceSpecificExitCode成员包含错误代码。服务应该在运行和正常终止时将该值设置为NO_ERROR。
    dwServiceSpecificExitCode   As Long '服务特定的错误代码，当服务启动或停止时，服务将返回错误。除非将dwWin32ExitCode成员设置为ERROR_SERVICE_SPECIFIC_ERROR，否则该值将被忽略。
    dwCheckPoint                As Long '这个检查点的值是定期的，以在一个漫长的开始、停止、暂停或继续操作的过程中报告它的进展。例如，当服务启动时，它完成了初始化的每一步，该服务应该增加这个值。调用服务上的操作的用户界面程序使用这个值来跟踪在长时间运行期间服务的进程。当服务没有启动、停止、暂停或继续操作时，该值是无效的，应该为零。
    dwWaitHint                  As Long '以毫秒为间隔的开始、停止、暂停或继续操作所需的时间。在指定的时间流逝之前，服务应该使用递增dwCheckPoint 或dwCurrentState的更改来对SetServiceStatus函数进行下一个调用。如果dwWaitHint指定的时间长度，dwCheckPoint没有增加或dwCurrentState没有更改，则服务控制管理器或服务控制程序可以假定发生了错误，应该停止服务。但是，如果服务与其他服务共享一个进程，服务控制管理器不能终止服务应用程序，因为它还必须终止其他共享进程的服务。
End Type
'       dwControlsAccepted:
'        控制代码在其处理程序函数中接受和处理(请参阅处理程序和HandlerEx)。
'        用户接口流程可以通过在ControlService或ControlServiceEx函数中指定控制命令来控制服务?
'        默认情况下，所有服务都接受 SERVICE_CONTROL_INTERROGATE 值。
'        接受SERVICE_CONTROL_DEVICEEVENT值 , 服务必须登记接收设备事件通过RegisterDeviceNotification函数?
'        以下是控制代码?
Public Enum SERVICE_CONTROL_ACCEPTED
    SERVICE_ACCEPT_NETBINDCHANGE = &H10
    '该服务是一个网络组件，它可以接受它的绑定中的更改，而不需要停止和重新启动。
    '该控制代码允许服务接收SERVICE_CONTROL_NETBINDADD, SERVICE_CONTROL_NETBINDREMOVE, SERVICE_CONTROL_NETBINDENABLE,SERVICE_CONTROL_NETBINDDISABLE 通知。
    SERVICE_ACCEPT_PARAMCHANGE = &H8
    '该服务可以重新读取它的启动参数，而不会被停止和重新启动。
    '该控制代码允许服务接收SERVICE_CONTROL_PARAMCHANGE通知。
    SERVICE_ACCEPT_PAUSE_CONTINUE = &H2
    '服务可以暂停并继续。
    '该控制代码允许服务接收SERVICE_CONTROL_PAUSE , SERVICE_CONTROL_CONTINUE 通知?
    SERVICE_ACCEPT_PRESHUTDOWN = &H100
    '该服务可以执行预关闭任务?
    '此控制代码使服务能够接收SERVICE_CONTROL_PRESHUTDOWN通知?
    '注意，ControlService和ControlServiceEx不能发送这个通知;
    '只有系统才能发送?
    'Windows Server 2003和Windows XP:这个值不受支持。
    SERVICE_ACCEPT_SHUTDOWN = &H4
    '当系统关闭发生时，该服务将被通知。
    '该控制代码允许服务接收SERVICE_CONTROL_SHUTDOWN通知?
    '注意，ControlService和ControlServiceEx不能发送这个通知;
    '只有系统才能发送?
    SERVICE_ACCEPT_STOP = &H1
    '服务可以被停止?
    '该控制代码允许服务接收SERVICE_CONTROL_STOP 通知?
'该成员还可以包含以下扩展控制代码，这些代码仅由HandlerEx支持。(请注意，这些控制代码不能由ControlService或ControlServiceEx发送。)
    SERVICE_ACCEPT_HARDWAREPROFILECHANGE = &H20
    '当计算机的硬件配置发生改变时，该服务将被通知。
    '这样，系统就可以将SERVICE_CONTROL_HARDWAREPROFILECHANGE通知发送到服务。
    SERVICE_ACCEPT_POWEREVENT = &H40
    '当计算机的电源状态发生变化时，该服务将被通知。
    '这使系统能够向服务发送SERVICE_CONTROL_POWEREVENT通知?
    SERVICE_ACCEPT_SESSIONCHANGE = &H80
    '当计算机的会话状态发生变化时，该服务将被通知。
    '这使系统能够向服务发送SERVICE_CONTROL_SESSIONCHANGE 通知?
    SERVICE_ACCEPT_TIMECHANGE = &H200
    '当系统时间发生变化时，将通知服务。
    '这使得系统能够将SERVICE_CONTROL_TIMECHANGE通知发送给服务?
    'Windows Server 2008、Windows Vista、Windows Server 2003和Windows XP:不支持这一控制代码。
    SERVICE_ACCEPT_TRIGGEREVENT = &H400
    '当服务已注册的事件发生时，将通知服务。
    '这使得系统能够将SERVICE_CONTROL_TRIGGEREVENT通知发送到服务?
    'Windows Server 2008、Windows Vista、Windows Server 2003和Windows XP:不支持这一控制代码。
    SERVICE_ACCEPT_USERMODEREBOOT = &H800
    '当用户发起重新引导时，将通知服务。
    'Windows Server 2008 R2、Windows 7、Windows Server 2008、Windows Vista、Windows Server 2003和Windows XP:不支持这一控制代码。
End Enum
'    dwWin32ExitCode:

'注意事项：ControlService函数要求服务控制管理器(SCM)将请求的控制代码发送到服务。SCM发送代码，如果服务指定它将接受代码，并且处于一个控制代码可以被发送到它的状态。
'          SCM流程服务控制通知在一个串行的过程中——它将等待一个服务在发送下一个服务之前完成一个服务控制通知。因此，如果任何服务都忙于处理控制代码，那么对ControlService的调用将阻塞30秒。如果当超时过期时，繁忙的服务仍然没有从它的处理程序返回，则ControlService会与ERROR_SERVICE_REQUEST_TIMEOUT发生故障。
'           要停止并启动服务，需要一个允许您这样做的安全描述符。默认的安全描述符允许LocalSystem帐户，以及管理员和Power用户组的成员停止和启动服务。要更改服务的安全描述符，请参见修改用于服务的DACL。
'           QueryServiceStatusEx函数返回一个SERVICE_STATUS_PROCESS结构，它的dwCurrentState和dwcontrols接受的成员指示运行服务接受的当前状态和控件。默认情况下，所有运行服务都接受service_control_审问控制代码。除了SERVICE_CONTROL_STOP和service_control_审问，驱动程序不接受其他的控制代码。每个服务指定当调用SetServiceStatus函数来报告其状态时接受的其他控制代码。服务在运行时应该始终接受这些代码，无论它在做什么。
'           下表显示了SCM在每个可能的服务状态中的操作?
'           Service state       Stop    Other controls
'           STOPPED             (c)         (c)
'           STOP_PENDING        (b)         (b)
'           START_PENDING       (a)         (b)
'           RUNNING             (a)         (a)
'           CONTINUE_PENDING    (a)         (a)
'           PAUSE_PENDING       (a)         (a)
'           PAUSED              (a)         (a)
'           (a)如果服务接受此控制代码，则将请求发送到服务;否则，ControlService返回0,GetLastError返回ERROR_INVALID_SERVICE_CONTROL。
'           (b)该服务不处于控制可以发送到它的状态，因此控制服务返回零，GetLastError返回ERROR_SERVICE_CANNOT_ACCEPT_CTRL。
'           (c)服务不活动，所以ControlService返回0,GetLastError返回ERROR_SERVICE_NOT_ACTIVE。
Private Declare Function DeleteService Lib "advapi32.dll" (ByVal hService As Long) As Long
'说明：从服务控制管理器数据库标记指定的删除服务
'返回：如果函数成功，返回值为非零。如果函数失败，返回值为零。要得到扩展的错误信息，请调用GetLastError。下面的错误代码可以由服务控制管理器设置?其他可能由服务控制管理器调用的注册表函数设置。
'ERROR_ACCESS_DENIED                                            句柄没有删除权限。
'ERROR_INVALID_HANDLE                                           指定的句柄无效。
Private Const ERROR_SERVICE_MARKED_FOR_DELETE As Long = &H430    '指定的服务已经被标记为删除。
'hService:服务的句柄。该句柄由OpenService或CreateService函数返回，它必须具有删除权限。有关更多信息，请参见服务安全性和访问权限
'注意事项：DeleteService函数标志着从服务控制管理器数据库中删除的服务。直到服务的所有打开句柄通过调用CloseServiceHandle函数来关闭，并且服务没有运行时，数据库条目才会被删除。使用SERVICE_CONTROL_STOP控制代码对ControlService函数的调用停止运行服务。如果服务无法停止，则在重新启动系统时删除数据库条目。服务控制管理器从注册表中删除服务密钥及其子键来删除服务?
Private Declare Function ChangeServiceConfig Lib "advapi32.dll" Alias "ChangeServiceConfigA" (ByVal hService As Long, ByVal dwServiceType As Long, ByVal dwStartType As Long, ByVal dwErrorControl As Long, ByVal lpBinaryPathName As String, ByVal lpLoadOrderGroup As String, lpdwTagId As Long, ByVal lpDependencies As String, ByVal lpServiceStartName As String, ByVal lpPassword As String, ByVal lpDisplayName As String) As Long
'说明：更改服务的配置参数。要更改可选的配置参数，使用ChangeServiceConfig2函数
'返回：如果函数成功，返回值为非零。如果函数失败，返回值为零。要得到扩展的错误信息，请调用GetLastError。下面的错误代码可以由服务控制管理器设置?其他错误代码可能由服务控制管理器调用的注册表函数设置。
'ERROR_ACCESS_DENIED 句柄没有SERVICE_CHANGE_CONFIG访问权。
Private Const ERROR_CIRCULAR_DEPENDENCY As Long = &H423          '指定了循环服务依赖项。
Private Const ERROR_DUPLICATE_SERVICE_NAME As Long = &H436       '显示名称已经存在于服务控制器管理器数据库中，无论是作为服务名称还是作为另一个显示名称。
'ERROR_INVALID_HANDLE                                           指定的句柄无效
'ERROR_INVALID_PARAMETER                                        其中一个函数参数无效
Private Const ERROR_INVALID_SERVICE_ACCOUNT As Long = &H421      '帐户名称不存在，或者指定一个服务与已经安装的服务共享相同的二进制文件，但是使用的帐户名称与已安装的服务不相同。
'ERROR_SERVICE_MARKED_FOR_DELETE                                指定的服务已经被标记为删除
'hService:服务的句柄。该句柄由OpenService或CreateService函数返回，并且必须具有SERVICE_CHANGE_CONFIG访问权。有关更多信息，请参见服务安全性和访问权限。
Private Const SERVICE_NO_CHANGE     As Long = &HFFFFFFFF
'dwServiceType:服务的类型。如果不更改现有的服务类型，请指定SERVICE_NO_CHANGE;否则，指定下列服务类型之一。
'               SERVICE_FILE_SYSTEM_DRIVER,SERVICE_KERNEL_DRIVER,SERVICE_WIN32_OWN_PROCESS,SERVICE_WIN32_SHARE_PROCESS
'               如果您指定SERVICE_WIN32_OWN_PROCESS或SERVICE_WIN32_SHARE_PROCESS，并且服务在LocalSystem帐户的上下文中运行，您还可以指定以下类型。
'               SERVICE_INTERACTIVE_PROCESS
'dwStartType:服务启动选项。如果不更改现有的start类型，指定SERVICE_NO_CHANGE;否则，指定下列值之一。
'           SERVICE_AUTO_START,SERVICE_BOOT_START,SERVICE_DEMAND_START,SERVICE_DISABLED,SERVICE_SYSTEM_START
'dwErrorControl:如果服务未能启动，则错误的严重性和所采取的操作。如果不更改现有的错误控制，请指定SERVICE_NO_CHANGE;否则，指定下列值之一。
'               SERVICE_ERROR_CRITICAL,SERVICE_ERROR_IGNORE,SERVICE_ERROR_NORMAL,SERVICE_ERROR_SEVERE
'lpBinaryPathName:服务二进制文件的完全限定路径。如果不更改现有路径，则指定NULL。如果路径包含空格，则必须引用它，以便正确地解释它。例如,“d:\ \ \ \ myservice。exe应该被指定为\ \ \ \ \ \ \ \ \ \ myservice。exe \ "。
'                 该路径还可以包含自动启动服务的参数。例如,“d:\ \ myshare \ \ myservice。exe __arg1最长。”这些参数传递给服务入口点(通常是主函数)。
'               如果指定另一台计算机上的路径，则必须通过本地计算机的计算机帐户访问共享，因为这是远程调用中使用的安全上下文。但是，这一要求允许远程计算机中的任何潜在漏洞影响本地计算机。因此，最好使用本地文件。
'lpLoadOrderGroup:这个服务是成员的负载排序组的名称。如果不更改现有组，指定NULL。如果服务不属于某个组，则指定一个空字符串。
'                   启动程序使用load命令组以指定的顺序将服务组装入到其他组中?负载排序组的列表包含在以下注册表键的ServiceGroupOrder值中:
'                   HKEY_LOCAL_MACHINE\System\CurrentControlSet\Control
'lpdwTagId:一个指向一个变量的指针，该变量在lpLoadOrderGroup参数中指定的组中所接收的标记值是唯一的。如果不更改现有的标记，则指定NULL。
'           您可以通过在以下注册表键的GroupOrderList值中指定一个标记顺序向量来在负载排序组中使用一个标签来订购服务启动:
'           HKEY_LOCAL_MACHINE\System\CurrentControlSet\Control
'           标签只对有SERVICE_BOOT_START或SERVICE_SYSTEM_START类型的驱动服务进行评估?
'lpDependencies:一个指针，指向一个双NULL结束的、被NULL分割的服务或负载排序组名称，该系统必须在此服务启动之前启动。(对一个组的依赖意味着这个服务可以运行，如果组中至少有一个成员在尝试启动所有组的所有成员之后运行。)如果不更改现有的依赖项，指定NULL。如果服务没有依赖关系，指定一个空字符串。
'               您必须用SC_GROUP_IDENTIFIER前缀组名称，这样它们可以与服务名区别开来，因为服务和服务组共享相同的名称空间。
'lpServiceStartName:服务应该运行的帐户的名称。如果不更改现有帐户名称，则指定NULL。如果服务类型是SERVICE_WIN32_OWN_PROCESS，则在form DomainName \ UserName中使用帐户名。服务流程将作为该用户登录。如果帐户属于内建域，则可以指定。\用户名(注意，对应的C / c++字符串是“. \ \用户名”)。有关更多信息，请参阅服务用户帐户和备注部分中的警告。共享进程可以作为任何用户运行.
'                   如果服务类型是SERVICE_KERNEL_DRIVER或SERVICE_FILE_SYSTEM_DRIVER，名称是系统用来加载设备驱动程序的驱动对象名称。如果驱动程序使用由I / O系统创建的默认对象名称，则指定NULL。
'                   服务可以配置为使用托管帐户或虚拟帐户。如果服务配置为使用托管服务帐户，则名称为托管服务帐户名称。如果服务配置为使用虚拟帐户，则指定名称为NT服务\ ServiceName。有关托管服务帐户和虚拟帐户的更多信息，请参阅服务帐户一步一步指南。
'                   Windows Server 2008,Windows Vista,Windows Server 2003和Windows XP:在windows7和Windows Server 2008 R2之前都不支持托管服务帐户和虚拟帐户
'lpPassword:由lpServiceStartName参数指定的帐户名称的密码。如果不更改现有密码，则指定NULL。如果帐户没有密码，或者服务在LocalService、NetworkService或LocalSystem帐户中运行，则指定一个空字符串。有关更多信息，请参见服务记录列表。
'               如果lpServiceStartName参数指定的帐户名称是托管服务帐户或虚拟帐户名称的名称，则lpPassword参数必须为空。
'               在驱动程序服务中，密码被忽略。
'lpDisplayName:用于为其用户标识服务的应用程序使用的显示名称。如果不更改现有的显示名称，则指定NULL;否则，该字符串的最大长度为256个字符。名称在服务控制管理器中保存。显示名称比较总是不区分大小写的。
'               该参数可以使用以下格式指定本地化字符串:
'               @[路径\]dllname -strID
'               带有标识符strID的字符串从dllname加载;路径是可选的。有关更多信息，请参见RegLoadMUIString。
'               Windows Server 2003和Windows XP:在Windows Vista之前不支持本地化字符串。
'注意事项：ChangeServiceConfig函数更改服务控制管理器数据库中指定服务的配置信息。您可以通过使用QueryServiceConfig函数获得当前的配置信息。
'           如果为正在运行的服务更改配置，除lpDisplayName外，更改在停止服务之前不会生效。要更新凭据而无需重新启动服务，请使用LsaCallAuthenticationPackage函数。
'           设置lpServiceStartName参数更改服务的登录帐户。这可能会导致问题。如果您已经注册了服务主名(SPN)，那么它现在将在错误的帐户上注册。类似地，如果您使用ACE来授权访问服务，那么它现在将允许访问错误的帐户。
Private Declare Function QueryServiceConfig Lib "advapi32.dll" Alias "QueryServiceConfigW" (ByVal hService As Long, lpServiceConfig As QUERY_SERVICE_CONFIG, ByVal cbBufSize As Long, pcbBytesNeeded As Long) As Long
'功能: 检索指定服务的配置参数
'参数
'【hService】标识服务。该句柄由OpenService或CreateService函数返回，并且必须具有SERVICE_QUERY_CONFIG访问权限。
'【lpServiceConfig】指向返回服务配置信息的QUERY_SERVICE_CONFIG结构?
'【cbBufSize】指定由【lpServiceConfig】参数指向的缓冲区的大小（以字节为单位）。
'【pcbBytesNeeded】指向一个变量，它接收返回所有配置信息所需的字节数。
'
'返回值
'   如果函数成功，返回值不为零。
'       如果函数失败，返回值为零。要获取扩展错误信息，请调用GetLastError.
'错误
'
'   服务控制管理器可以设置以下错误代码?其他可以由服务控制管理器调用的注册表功能设置?
'
'值                         含义
'ERROR_ACCESS_DENIED        指定的句柄未被打开，具有SERVICE_QUERY_CONFIG访问权限。
'ERROR_INSUFFICIENT_BUFFER 与【lpServiceConfig】缓冲区相比，有更多的服务配置信息。在【pcbBytesNeeded】参数中返回获取所有信息所需的字节数。没有写入【lpServiceConfig】缓冲区。
'ERROR_INVALID_HANDLE       指定的句柄无效?
'备注
'
'   QueryServiceConfig函数返回特定服务在注册表中保存的服务配置信息。此配置信息首先通过CreateService功能设置，并且可能已通过ChangeServiceConfig功能更新。如果服务在上次更改配置信息时运行，则返回的信息将不会反映正在运行的服务的真实配置。相反，它将在下一次运行时反映服务的配置。停止服务后，将使用存储在注册表中的配置更新服务控制管理器数据库。DisplayName键是一个例外。当DisplayName键更改时，无论该服务是否正在运行，它将立即生效。
'【lpServiceConfig】缓冲区必须足够大以容纳QUERY_SERVICE_CONFIG结构的成员指向的字符串?
Private Const ERROR_INSUFFICIENT_BUFFER = &H7A

Private Type QUERY_SERVICE_CONFIG
    dwServiceType As Long
    dwStartType As Long
    dwErrorControl As Long
    lpBinaryPathName As Long
    lpLoadOrderGroup As Long
    dwTagId As Long
    lpDependencies As Long
    lpServiceStartName As Long
    lpDisplayName As Long
End Type

Private Declare Function StartService Lib "advapi32" Alias "StartServiceW" (ByVal hService As Long, ByVal dwNumServiceArgs As Long, ByVal lpServiceArgVectors As Long) As Long
'说明：启动一个服务。
'返回：如果函数成功，返回值为非零。如果函数失败，返回值为零。要得到扩展的错误信息，请调用GetLastError。下面的错误代码可以由服务控制管理器设置?其他可以由服务控制管理器调用的注册表函数设置?
'ERROR_ACCESS_DENIED                    句柄没有SERVICE_START访问权。
'ERROR_INVALID_HANDLE                   指定的句柄无效
'ERROR_PATH_NOT_FOUND                   服务二进制文件不能被找到
'ERROR_SERVICE_ALREADY_RUNNING          服务的实例已经在运行了
'ERROR_SERVICE_DATABASE_LOCKED          数据库是锁着的
'ERROR_SERVICE_DEPENDENCY_DELETED       服务依赖于不存在的服务或已标记为删除的服务
'ERROR_SERVICE_DEPENDENCY_FAIL          该服务依赖于另一个未能启动的服务
'ERROR_SERVICE_DISABLED                 该服务已被禁用
'ERROR_SERVICE_LOGON_FAILED             由于登录失败，服务并没有开始。如果将服务配置为在没有“作为服务的日志”的帐户下运行，则会出现此错误。
'ERROR_SERVICE_MARKED_FOR_DELETE
'该服务已被标记为删除?
'ERROR_SERVICE_NO_THREAD
'不能为该服务创建一个线程?
'ERROR_SERVICE_REQUEST_TIMEOUT
'服务的过程开始 , 但它不叫StartServiceCtrlDispatcher, 或称为StartServiceCtrlDispatcher可能阻塞的线程控制处理器函数?
'hService:服务的句柄。这个句柄由OpenService或CreateService函数返回，它必须具有servicestart访问权。要了解更多信息，请参阅服务安全性和访问权限。
'dwNumServiceArgs:lpServiceArgVectors数组中的字符串数量。如果lpServiceArgVectors为空，则该参数可以为零。
'lpServiceArgVectors:以null结尾的字符串作为参数传递给ServiceMain函数。如果没有参数，这个参数可以是NULL。否则，第一个参数(lpServiceArgVectors[0])是服务的名称，后面是任何附加的参数(lpServiceArgVectors[1]到lpServiceArgVectors[dwnumserviceargs-1])。
'                      驱动程序服务不接收这些参数?
'注意事项：当启动一个驱动程序服务时，StartService函数在设备驱动程序完成初始化之前不会返回。
'            当服务启动时，如果需要，服务控制管理器(SCM)会生成服务进程。如果指定的服务与其他服务共享一个进程，则需要进程已经存在。StartService函数不会等待新服务的第一次状态更新，因为它可能需要一段时间。相反，当SCM收到来自服务控制调度程序的通知时，它将返回，该服务的ServiceMain线程被成功创建。
'            在从StartService返回之前，SCM设置以下默认的状态值
'                   服务的当前状态设置为servicestartpending
'                   被接受的控件被设置为none(0)。
'                   检查点值被设置为0?
'                   这段时间的时间设置为2秒?
'            调用进程可以确定新服务是否完成了初始化，通过调用QueryServiceStatus函数来查询服务的状态。
'            在初始化期间，服务不能调用StartService。原因是SCM在初始化期间锁定了服务控制数据库，因此对StartService的调用将会阻塞。在服务向SCM报告它已经成功启动之后，它可以调用StartService。
'            与ControlService一样，如果任何服务都忙于处理控制代码，StartService将会阻塞30秒。如果繁忙的服务在超时过期时仍然没有返回它的处理程序函数，那么StartService将会出现errorservicerequesttimeout。这是因为SCM每次只处理一个服务控制通知。
Private Declare Function StartServiceCtrlDispatcher Lib "advapi32.dll" Alias "StartServiceCtrlDispatcherW" (lpServiceStartTable As SERVICE_TABLE_ENTRY) As Long
'说明：将服务进程的主线程连接到服务控制管理器，这导致线程成为调用进程的服务控制调度程序线程。
'返回：如果函数成功，则返回值是非0。
'        如果函数失败，则返回值为0。要获得扩展错误信息，请调用GetLastError。
'        下面的错误代码可以由服务控制管理器设置?其他的错误代码可以由服务控制管理器调用的注册中心函数来设置?
Private Const ERROR_FAILED_SERVICE_CONTROLLER_CONNECT   As Long = &H427         '如果程序是作为控制台应用程序而不是作为服务运行，则会返回这个错误。如果程序将作为一个用于调试的控制台应用程序运行，那么在返回这个错误时，将不会调用特定于服务的代码。
Private Const ERROR_INVALID_DATA        As Long = &HD           '指定的分派表包含不属于适当格式的条目
Private Const ERROR_SERVICE_ALREADY_RUNNING             As Long = &H420          '这个过程已经被StartServiceCtrlDispatcher调用。每个进程可以调用StartServiceCtrlDispatcher只有一次。
'注意事项:当服务控制管理器启动一个服务进程,这等待的过程叫StartServiceCtrlDispatcher函数。服务流程的主线程应该在启动后尽快进行(在30秒内)。如果StartServiceCtrlDispatcher成功,它连接调用线程服务控制管理器和不返回,直到所有正在运行的服务过程中进入SERVICE_STOPPED状态。服务控制管理器使用这个连接将控制和服务启动请求发送给服务进程的主线程。主线程通过调用适当的HandlerEx函数来处理控制请求，或者在启动新服务时创建一个新线程来执行适当的ServiceMain函数，从而充当调度程序。
'        lpServiceTable参数包含可以在调用过程中运行的每个服务的条目。每个条目都指定该服务的ServiceMain函数。对于servicewin32shareprocess服务，每个条目必须包含服务的名称。这个名称是在安装服务时由CreateService函数指定的服务名称。对于servicewin32ownprocess服务，表条目中的服务名被忽略。
'        如果一个服务运行在自己的过程,服务进程的主线程应立即叫StartServiceCtrlDispatcher。当服务启动时，所有的初始化任务都在服务的ServiceMain函数中完成。
'        如果多个服务共享一个过程和一些常见的过程差距需要做初始化之前ServiceMain函数被调用时 , 主线程可以做这项工作在调用StartServiceCtrlDispatcher之前, 只要不到30秒?否则, 必须创建另一个线程初始化过程差距, 而主线程调用StartServiceCtrlDispatcher, 成为服务控制分派器?任何服务特定的初始化都应该在单独的服务主函数中完成?
'        服务不应该试图直接显示用户界面。要了解更多信息，请参见交互式服务。
'lpServiceTable:一个指向servicetableentry结构数组的指针，它包含了一个可以在调用过程中执行的每个服务的条目。表中的最后一个条目的成员必须有空值来指定表的结束。
'为可以在调用过程中运行的服务指定ServiceMain函数。它是由StartServiceCtrlDispatcher函数使用。
Private Type SERVICE_TABLE_ENTRY
    lpServiceName       As Long         '在这个服务过程中运行的服务的名称。如果服务是使用SERVICE_WIN32_OWN_PROCESS 服务类型安装的，那么该成员将被忽略，但不能为空。这个成员可以是一个空字符串(“”)。
                                        '如果服务是使用SERVICE_WIN32_SHARE_PROCESS服务类型安装的，则该成员指定使用由lpServiceProc成员指向的ServiceMain函数的服务的名称。
    lpServiceProc       As Long         '一个指向ServiceMain函数的指针。
    lpServiceNameNull   As Long
    lpServiceProcNull   As Long
End Type
Private Declare Function QueryServiceStatus Lib "advapi32.dll" (ByVal hService As Long, lpServiceStatus As SERVICE_STATUS) As Long

Private Declare Function SetServiceStatus Lib "advapi32.dll" (ByVal hServiceStatus As Long, lpServiceStatus As SERVICE_STATUS) As Long
'功能：更新服务控制管理器对调用服务的状态信息。
'返回：如果函数成功，则返回值是非0。
'        如果函数失败，则返回值为0。要获得扩展错误信息，请调用GetLastError。
'        下面的错误代码可以由服务控制管理器设置?其他的错误代码可以由服务控制管理器调用的注册中心函数来设置。
'        ERROR_INVALID_DATA：指定的服务状态结构是无效的?
'        ERROR_INVALID_HANDLE：指定的句柄是无效的?
'注意事项：ServiceMain函数首先调用服务的 RegisterServiceCtrlHandlerEx函数获取SERVICE_STATUS_HANDLE。然后，它立即调用SetServiceStatus函数，
'        通知服务控制管理器，它的状态是SERVICE_STATUS_HANDLE。在初始化期间，服务可以提供更新的状态，以表明它正在取得进展，但它需要更多的时间。
'        一个常见的错误是服务让主线程执行初始化，而另一个线程则继续调用SetServiceStatus，以防止服务控制管理器将其标记为挂起。
'        但是，如果主线程挂起，那么服务开始在无限循环中结束，因为工作线程继续报告主线程正在取得进展。
'        在处理了一个控制请求之后，如果服务状态更改为服务控制管理器报告它的新状态，那么服务的处理程序函数必须调用SetServiceStatus。
'        当服务正在改变状态时，只有当它正在处理停止或关闭控制时，才有必要这样做。服务还可以在任何时候从服务的任何线程中使用此函数来通知状态更改的服务控制管理器，
'        例如服务必须由于可恢复的错误而停止?
'        服务可以调用这个函数之后才叫做RegisterServiceCtrlHandlerEx服务状态处理?
'        如果一个服务调用了带有dwCurrentState成员的SetServiceStatus，并且dwWin32ExitCode成员设置为非零值，那么下面的条目将被写入系统事件日志:
'        Event ID    = 7023
'        Source      = Service Control Manager
'        Type        = Error
'        Description = <ServiceName> terminated with the following error:
'                      <ExitCode>.
'       以下是调用该函数时的最佳实践:
'           在 SERVICE_STATUS 结构中初始化所有字段，确保为等待状态提供有效的检查点和等待提示值。使用合理的等待提示。
'           在状态为 SERVICE_START_PENDING服务崩溃时，不要注册接受控制。完成初始化之后，接受SERVICE_CONTROL_STOP代码。
'           只有当服务在与等待开始、停止、暂停或继续操作相关的任务上取得进展时，才会使用检查点和等待提示值来调用此函数。否则，SCM无法检测您的服务是否挂起。
'           如果ServiceMain失败，就进入停止状态，并使用适当的退出代码。
'           如果状态停止，执行所有必要的清理，并只调用SetServiceStatus。这个函数向SCM发出一个LRPC调用。在service中止状态下的第一个调用函数将关闭RPC上下文句柄，任何后续调用都可能导致进程崩溃。
'           在调用SetServiceStatus之后，不要尝试执行任何额外的工作，因为服务流程可以在任何时候终止。
'hServiceStatus:对当前服务的状态信息结构的句柄。这是用RegisterServiceCtrlHandlerEx函数返回的句柄?
'lpServiceStatus:一个指向SERVICE_STATUS结构的指针，包含了调用服务的最新状态信息。
Public Enum SERVICE_STATE
    SERVICE_STOPPED = &H1           '服务没有运行。
    SERVICE_START_PENDING = &H2     '服务开始。
    SERVICE_STOP_PENDING = &H3      '服务停止。
    SERVICE_RUNNING = &H4           '服务正在运行。
    SERVICE_CONTINUE_PENDING = &H5  '服务仍在等待中
    SERVICE_PAUSE_PENDING = &H6     '服务暂停正在等待。
    SERVICE_PAUSED = &H7            '服务暂停。
End Enum
Private Declare Function RegisterServiceCtrlHandler Lib "advapi32.dll" Alias "RegisterServiceCtrlHandlerW" (ByVal lpServiceName As Long, ByVal lpHandlerProc As Long) As Long
'功能：注册一个函数来处理服务控制请求。这个函数取代了RegisterServiceCtrlHandlerEx函数。一个服务可以使用任何一个函数，但是新的函数支持用户定义的上下文数据，新的处理函数支持额外的扩展控制代码。
'返回：如果函数成功，则返回值是一个服务状态句柄。
'        如果函数失败，则返回值为0。要获得扩展错误信息，请调用GetLastError。
'        下面的错误代码可以由服务控制管理器设置?
'        ERROR_NOT_ENOUGH_MEMORY：没有足够的内存可以将ANSI字符串参数转换为Unicode?这个错误不会出现在Unicode字符串参数中?
'        ERROR_SERVICE_NOT_IN_EXE：指定的服务条目是错误被StartServiceCtrlDispatcher函数带调用过程。
'注意事项：ServiceMain函数的一个新的服务应立即呼叫RegisterServiceCtrlHandler函数来注册一个控制与控制分派器的处理函数。这使控制分派器在接收到该服务的控制请求时调用指定的函数。对于可能的控制代码列表，请参见处理程序。调用进程的线程可以使用该函数返回的服务状态句柄，以在随后对SetServiceStatus函数的调用中识别服务。
'        RegisterServiceCtrlHandler函数前必须被称为第一个SetServiceStatus叫因为RegisterServiceCtrlHandler返回调用者使用的服务状态处理,没有其他服务可以无意中设置这个服务状态。此外，当服务指定通过SetServiceStatus函数所接受的控件时，控制处理程序必须到位来接收控制请求。
'        当使用控制请求调用控制处理程序函数时，服务必须调用SetServiceStatus，仅当服务状态发生更改时才向服务控制管理器报告状态，例如服务处理停止或关闭控制时。如果服务状态没有更改，则服务不应该向服务控制管理器报告状态。
'        服务状态句柄不需要关闭?
'lpServiceName:由调用线程运行的服务的名称。这是创建服务时在CreateService函数中指定的服务控制程序的服务名称。如果服务类型是SERVICE_WIN32_OWN_PROCESS，则该函数不会验证指定的名称是否有效，因为流程中只有一个注册的服务。
'lpHandlerProc:一个指向处理程序函数的指针。要了解更多信息，请参见处理程序。

Private Declare Function GetCurrentThread Lib "kernel32" () As Long
'功能：返回调用线程的伪句柄。
'返回：返回值是当前线程的一个伪句柄。
'注意事项：伪句柄是一个特殊的常量，它被解释为当前的线程句柄。
'           当需要一个线程句柄时，调用线程可以使用这个句柄来指定它自己。
'           伪句柄不是由子进程继承的。
'           这个句柄拥有线程对象的 THREAD_ALL_ACCESS 访问权。要了解更多信息，请参阅线程安全性和访问权限。
'            Windows Server 2003和Windows XP:这个句柄拥有线程的安全描述符允许的最大访问权限，这是进程的主标记。
'            一个线程不能使用该函数创建一个句柄，其他线程可以使用该句柄来引用第一个线程。
'            句柄总是被解释为引用使用它的线程?
'            一个线程可以为自己创建一个“真实的”句柄，它可以被其他线程使用，也可以通过其他进程继承，通过在调用DuplicateHandle函数的过程中指定伪句柄作为源句柄。
'            当不再需要时，伪句柄不需要关闭。
'            用这个句柄调用闭句函数没有效果?
'            如果伪句柄被DuplicateHandle复制，那么复制句柄必须被关闭。
'            在模拟安全上下文时不要创建线程?
'            调用将成功，但是新创建的线程在调用GetCurrentThread时将会减少对自身的访问权。
'            授予此线程的访问权限将来自模拟用户对流程的访问权限?
'            包括THREAD_SET_THREAD_TOKEN 和 THREAD_GET_CONTEXT 在内的一些访问权限可能不会出现，从而导致意外的失败。
Private Declare Function GetCurrentThreadId Lib "kernel32" () As Long
'说明：检索正在调用线程的线程标识符。
'返回：返回值是调用线程的线程标识符。
'注意事项：在线程终止之前，线程标识符惟一地标识整个系统中的线程。

Private Declare Function NetWkstaUserGetInfo Lib "Netapi32" (ByVal Reserved As Any, ByVal Level As Long, lpBuffer As Any) As Long
'说明：NetWkstaUserGetInfo函数返回有关当前登录用户的信息。这个函数必须在登录用户的上下文中调用。
'返回：如果函数成功，返回值是NERR_Success。如果函数失败，返回值可能是下列错误代码之一。
Private Const ERROR_NOT_ENOUGH_MEMORY   As Long = &H8           '系统耗尽了内存资源。要么网络管理器配置不正确，要么程序运行在内存不足的系统上。
Private Const ERROR_INVALID_LEVEL       As Long = &H7C          'Level参数无效。
Private Const ERROR_INVALID_PARAMETER   As Long = &H57          '其中一个函数参数无效。
'reserved:这个参数必须设置为NULL。
'Level:指定数据的信息级别。这个参数可以是下列值之一
'       0-返回当前登录到工作站上的用户名。bufptr参数指向WKSTA_USER_INFO_0结构。
'       1-返回关于工作站的信息，包括当前用户的名称和工作站访问的域。bufptr参数指向WKSTA_USER_INFO_1结构。
'       1101-返回被工作站浏览的域。bufptr参数指向WKSTA_USER_INFO_1101结构。
'lpBuffer:指向接收数据的缓冲区的指针。此数据的格式取决于bufptr参数的值。此缓冲区由系统分配，必须使用NetApiBufferFree函数释放。有关更多信息，请参见网络管理函数缓冲区和网络管理函数缓冲区长度。
Private Declare Function NetApiBufferFree Lib "Netapi32" (ByVal lpBuffer As Long) As Long
'说明：NetApiBufferFree函数释放了netapibuffer分配函数的内存。应用程序还应该调用NetApiBufferFree来释放其他网络管理函数在内部使用的内存来返回信息。
'返回：如果函数成功，返回值是NERR_Success。如果函数失败，返回值是一个系统错误代码。有关错误代码的列表，请参见系统错误代码。
'lpBuffer:通过调用netapibuffer分配函数的另一个网络管理函数或内存来返回的缓冲区的指针。
'注意事项：NetApiBufferFree函数用于网络管理函数使用的内存。这个函数用于两种情况:
'               在不需要内存时，将应用程序中的调用显式地分配给netapibuffer分配函数。
'               通过在应用程序中调用远程网络管理函数来释放内存，将信息返回给调用者。RPC运行时库在内部分配包含返回信息的缓冲区。
'           许多网络管理函数检索信息并将这些信息作为缓冲区返回，这些缓冲区可能包含复杂的结构、结构数组或嵌套结构数组。这些函数使用RPC运行时库在内部分配包含返回信息的缓冲区，无论调用是本地计算机还是远程服务器。例如，NetServerEnum函数检索一个服务器列表，并将此信息作为一个由bufptr参数指向的结构数组返回。当函数成功时，内存由NetServerEnum函数在内部分配，以存储bufptr参数中返回的结构数组。当不再需要这种结构的数组时，NetApiBufferFree函数应该被应用程序调用，并将缓冲区参数设置为NetServerEnum返回的bufptr参数，以释放所使用的内存。在这些情况下，NetApiBufferFree函数释放所有为缓冲区分配的内存，包括用于嵌套结构的内存、指向字符串的指针和其他数据。
'           不需要特殊的组成员来成功地执行NetApiBufferFree函数或任何其他的ApiBuffer函数.
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
'说明：将内存块从一个位置移动到另一个位置
'Destination:指向移动目的地起始地址的指针。
'Source:指向要移动的内存块起始地址的指针。
'Length:内存块的大小以字节为单位移动。
'注意事项：这个函数定义为RtlMoveMemory函数。它的实现是内联的。有关更多信息，请参见WinBase。h和Winnt.h。源和目标块可能会重叠。
'           第一个参数，目的地，必须足够大，以容纳长度字节的源;否则，可能会出现缓冲区溢出。这可能导致拒绝服务攻击，如果有访问违反，或者在最坏的情况下，允许攻击者向您的进程注入可执行代码。如果目的地是一个基于堆栈的缓冲区，则尤其如此。要注意，最后一个参数，长度，是将字节复制到目的地的数量，而不是目的地的大小。
Private Declare Function lstrlenW Lib "kernel32" (ByVal lpString As Long) As Long
'说明：确定指定字符串的长度(不包括终止null字符)。
'返回：函数返回字符串的长度，在字符中。如果lpString是NULL，函数将返回0。
'lpString:要检查的NULL终止字符串。
'注意事项：错误地使用这个函数会损害应用程序的安全性。lstrlen假设lpString是一个空终止字符串，或NULL。如果不是这样，这会导致缓冲区溢出或拒绝对应用程序进行服务攻击。考虑使用下列选项之一: StringCbLength或StringCchLength?

Private Declare Function GetMessage Lib "user32" Alias "GetMessageA" (lpMsg As Msg, ByVal hwnd As Long, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long) As Long
'说明：从调用线程的消息队列中检索一条消息。该函数将传入的消息发送到一个已发布的消息，以供检索。与GetMessage不同的是，PeekMessage函数不会等待消息在返回之前被发布。
'返回：如果函数检索除WM_QUIT之外的消息，则返回值是非0。
'       如果函数检索WM_QUIT消息，则返回值为0。
'       如果存在错误，则返回值为-1。例如，如果hWnd是一个无效的窗口句柄或者lpMsg是一个无效指针，那么函数就会失败。要获得扩展错误信息，请调用GetLastError。
'       因为返回值可以是非0，0，或-1，避免这样的代码:
'       while (GetMessage( lpMsg, hWnd, 0, 0)) ...
'       如果hWnd是一个无效的参数(例如指向已经被破坏的窗口)，那么返回-1返回值的可能性意味着这样的代码可能导致致命的应用程序错误。相反，使用这样的代码:
'       BOOL bRet;
'
'        while( (bRet = GetMessage( &msg, hWnd, 0, 0 )) != 0)
'        {
'            if (bRet == -1)
'            {
'                // handle the error and possibly exit
'            }
'            Else
'            {
'                TranslateMessage(&msg);
'                DispatchMessage(&msg);
'            }
'        }
'注意事项：应用程序通常使用返回值来确定是否结束主消息循环并退出程序。
'        GetMessage函数检索由hWnd参数或它的任何一个子元素所标识的消息，如IsChild函数所指定的，以及wMsgFilterMin和wMsgFilterMax参数所给出的消息值范围内的消息。
'        请注意，应用程序只能在wMsgFilterMin和wMsgFilterMax参数中使用Low WOrd;
'        High WOrd是为系统保留的?
'        请注意，GetMessage总是检索 WM_QUIT消息，不管您为wMsgFilterMin和wMsgFilterMax指定的值是什么。
'        在此调用期间，系统将发送未排队的、非排队的消息，即发送到由调用线程所拥有的、使用SendMessage、SendMessageCallback、SendMessageTimeout或SendNotifyMessage函数的消息。
'        然后，检索与指定过滤器匹配的第一个队列消息。
'        系统还可以处理内部事件?
'        如果没有指定过滤器，则按以下顺序处理消息:
'               发送消息
'               发布的消息
'               输入(硬件)消息和系统内部事件
'               发送消息 (再一次)
'               WM_PAINT消息
'               WM_TIMER消息
'        要在发布消息之前检索输入消息，请使用wMsgFilterMin和wMsgFilterMax参数。
'
'        GetMessage没有从队列中删除 WM_PAINT 消息?
'        消息保留在队列中，直到处理完毕。
'        如果一个顶级窗口停止对消息的响应超过几秒钟，系统会认为窗口没有响应，并使用一个具有相同z顺序、位置、大小和可视属性的幽灵窗口来替换它。
'        这允许用户移动它，调整大小，甚至关闭应用程序。
'        然而，这些是惟一可用的操作，因为应用程序实际上没有响应。
'        当在调试器模式下，系统不会产生一个幽灵窗口。
'lpMsg:一个指向MSG结构的指针，该结构从线程的消息队列接收消息信息。
Private Type POINT
    x           As Long             '点的x坐标。
    y           As Long             '点的Y坐标。
End Type

Private Type Msg
    hwnd        As Long             '窗口处理程序接收消息的窗口的句柄。当消息是一个线程消息时，该成员是NULL。
    message     As Long             '消息标识符。应用程序只能使用LowWOrd;High Word是由系统保留的。
    wParam      As Long             '关于消息的附加信息。确切的含义取决于消息成员的值。
    lParam      As Long             '关于消息的附加信息。确切的含义取决于消息成员的值。
    time        As Long             '消息发布的时间。
    pt          As POINT            '当消息被发布时，光标在屏幕坐标下的位置。
End Type
'hWnd:对将要检索消息的窗口的句柄。窗口必须属于当前线程?
'           如果hWnd为空，GetMessage将检索属于当前线程的任何窗口的消息，以及当前线程的消息队列上的任何消息，该消息队列的hWnd值为NULL(请参见MSG结构)。因此，如果hWnd为空，则处理窗口消息和线程消息。
'           如果hWnd是-1，GetMessage只检索当前线程的消息队列上的消息，该消息队列的hWnd值为NULL，即PostMessage(当hWnd参数为NULL)或PostThreadMessage时所发布的线程消息。
'wMsgFilterMin:检索到的最低消息值的整数值。使用WM_KEYFIRST(0x0100)指定第一个键盘消息或WM_MOUSEFIRST(0x0200)来指定第一个鼠标消息。
'           在这里和wMsgFilterMax中使用WM_INPUT  ，只指定WM_INPUT  消息。
'           如果wMsgFilterMin和wMsgFilterMax都是零，GetMessage将返回所有可用的消息(也就是说，没有执行范围过滤)。
'wMsgFilterMax:要检索的最高消息值的整数值?使用WM_KEYLAST指定最后一个键盘消息或WM_MOUSEFIRST来指定最后的鼠标消息?
'              在这里和wMsgFilterMin中使用WM_INPUT，只指WM_INPUT消息。
'              如果wMsgFilterMin和wMsgFilterMax都是零，GetMessage将返回所有可用的消息(也就是说，没有执行范围过滤)。
Private Declare Function PostThreadMessage Lib "user32" Alias "PostThreadMessageA" (ByVal idThread As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'功能：将消息发送到指定线程的消息队列。它返回时不等待线程处理消息。
'返回：如果函数成功，则返回值是非0。
'        如果函数失败，则返回值为0。要获得扩展错误信息，请调用GetLastError。如果idThread不是一个有效的线程标识符，或者由idThread指定的线程没有一个消息队列，那么GetLastError将返回ERROR_INVALID_THREAD_ID。当消息限制被触发时，GetLastError返回ERROR_NOT_ENOUGH_QUOTA 。
'注意事项：当一个消息被UIPI阻塞时，使用GetLastError检索的最后一个错误将被设置为5(访问被拒绝)。
'        消息被发布的线程必须创建一个消息队列，否则对PostThreadMessage的调用将失败。
'        使用下面的方法来处理这种情况?
'           创建一个事件对象，然后创建线程。
'           在调用PostThreadMessage之前，使用 WaitForSingleObject 函数来等待事件被设置为信号状态。
'           在发送消息的线程中，将调用PeekMessage，以强制系统创建消息队列。
'           PeekMessage(&msg, NULL, WM_USER, WM_USER, PM_NOREMOVE)
'           设置事件，以表明该线程已经准备好接收已发布的消息。
'        消息被发布的线程通过调用GetMessage或PeekMessage函数来检索消息。返回的MSG结构的hwnd成员为空。
'        PostThreadMessage发送的消息与窗口没有关联.通常情况下，与窗口没有关联的消息不能由DispatchMessage函数发送。
'        因此，如果接收线程处于一个模式循环中(如消息框或对话框所使用)，则消息将丢失。
'        要在模式循环中拦截线程消息，请使用特定于线程的钩子。
'        系统只对系统消息进行编组(范围为0到(WM_USER-1))。
'        要向另一个进程发送其他消息(这些消息>=WM_USER)，您必须进行自定义编组。
'        每个消息队列有10,000条消息的限制。这个极限应该足够大。
'        如果您的应用程序超出了限制，那么应该重新设计它，以避免使用过多的系统资源。
'        要调整这个限制，请修改下面的注册表键。
'        HKEY_LOCAL_MACHINE
'           SOFTWARE
'              Microsoft
'                 Windows NT
'                    CurrentVersion
'                       Windows
'                          USERPostMessageLimit
'        最小可接受值是4000
'idThread: 消息要被发布的线程的标识符?
'        如果指定的线程没有消息队列，则该函数将失败。
'        当线程对一个用户或GDI函数进行第一次调用时，系统会创建一个线程的消息队列。
'        要了解更多信息，请参见备注部分。
'        信息发布要服从于UIPI。
'        进程的线程可以只在较小或相等的完整性级别的进程中，将消息发送到消息队列的消息队列。
'        此线程必须具有SE_TCB_NAME特权，以便将消息发布到具有相同本地惟一标识符(LUID)的进程中，但位于不同的桌面。
'        否则，该函数将失败并返回ERROR_INVALID_THREAD_ID。
'        该线程必须属于与调用线程相同的桌面，也必须属于具有相同LUID的进程。
'        否则，该函数将失败并返回ERROR_INVALID_THREAD_ID.
'Msg:要发布的消息类型。
'wParam ：附加于消息的信息。
'lParam ： 附加于消息的信息。
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
'功能：在指定的对象处于信号状态或超时时间间隔的情况下等待。
'       要进入alertable等待状态，请使WaitForSingleObjectex函数。要等待多个对象，请使用 WaitForMultipleObjects.。
'返回：如果函数成功，则返回值表示导致函数返回的事件。它可以是下列值之一。
Private Const WAIT_ABANDONED        As Long = &H80       '指定的对象是一个互斥对象，这个互斥对象是在拥有线程终止之前拥有互斥对象的线程所释放的。
                                                        '互斥对象的所有权被授予给调用线程，而互斥锁状态被设置为非信号。
                                                        '如果互斥锁是在保护持久状态信息，那么您应该检查它是否具有一致性。
Private Const WAIT_OBJECT_0         As Long = &H0        '指定对象的状态是有信号的?
Private Const WAIT_TIMEOUT          As Long = &H102      '超时时间间隔，对象的状态是没有信号的。
Private Const WAIT_FAILED           As Long = &HFFFFFFFF     '这个函数失败了。要获得扩展错误信息，请调用GetLastError
'注意事项：WaitForSingleObject函数检查指定对象的当前状态。
'            如果对象的状态是非信号的，则调用线程进入等待状态，直到对象被发出信号或超时时间间隔。
'            该函数修改了某些类型的同步对象的状态?
'            修改只发生在信号状态导致函数返回的对象上?
'            例如，一个信号量对象的计数减少了1。
'            WaitForSingleObject函数可以等待以下对象:
'               更改通知
'               控制台输入
'               事件
'               内存资源的通知
'               互斥锁
'               过程
'               信号量
'               线程
'               可等待定时器
'            在调用等待函数和直接或间接创建windows的代码时要谨慎?
'            如果一个线程创建任何窗口，它必须处理消息。
'            消息广播被发送到系统中的所有窗口?
'            使用一个没有超时间隔的等待函数的线程可能会导致系统陷入僵局?
'            两个间接创建窗口的代码示例是DDE和CoInitialize函数?
'            因此 , 如果你有一个线程创建窗口, 使用MsgWaitForMultipleObjects或MsgWaitForMultipleObjectsEx, 而不是WaitForSingleObject。
'hHandle:对对象的句柄。
'        对于可以指定句柄的对象类型列表，请参阅下面的备注部分。
'        如果该句柄在等待仍未完成时关闭，则该函数的行为是未定义的。
'        该句柄必须具有同步访问权?
'        要了解更多信息，请参阅标准访问权限。
'dwMilliseconds:时间间隔，以毫秒为间隔。
'            如果指定了非零值，则函数将等待对象的信号或间隔时间的流逝。
'            如果dw毫秒为0，那么如果对象没有信号，函数就不会进入等待状态;
'            它总是立即返回?
'            如果dwMilliseconds是INFINITE，那么函数只有在对象发出信号时才会返回。
'            Windows XP、Windows Server 2003、Windows Vista、Windows 7、Windows Server 2008和Windows Server 2008 R2:dwMilliseconds确实包含了在低功耗状态下的时间。
'            例如，当计算机处于睡眠状态时，超时会继续计算。
'            Windows 8、Windows Server 2012、Windows 8.1、Windows Server 2012 R2、Windows 10和Windows Server 2016:dwMilliseconds不包括低功耗状态下的时间。
'            例如，当计算机处于睡眠状态时，超时并不会继续计数。
Private Declare Function PeekMessage Lib "user32" Alias "PeekMessageA" (lpMsg As Msg, ByVal hwnd As Long, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long, ByVal wRemoveMsg As Long) As Long
'功能：发送传入的消息，检查发送消息的线程消息队列，并检索消息(如果存在的话)。
'返回：如果有消息可用，则返回值是非0的。
'        如果没有可用的消息，则返回值为0。
'注意事项：PeekMessage检索与由hWnd参数指定的窗口相关联的消息，或者由IsChild函数指定的任何其子元素，以及wMsgFilterMin和wMsgFilterMax参数所给出的消息值范围内的消息。
'        请注意，应用程序只能在wMsgFilterMin和wMsgFilterMax参数中使用low Word;
'        High Word是为系统保留的?
'        注意，不管您为wMsgFilterMin和wMsgFilterMax指定的值是什么，PeekMessage总是检索WM_QUIT 消息。
'        在此调用期间，系统将发送未排队的、非排队的消息，即发送到由调用线程所拥有的、使用SendMessage、SendMessageCallback、SendMessageTimeout或SendNotifyMessage函数的消息。
'        然后，检索与指定过滤器匹配的第一个队列消息。
'        系统还可以处理内部事件?
'        如果没有指定过滤器，则按以下顺序处理消息:
'           发送消息
'           发布的消息
'           输入(硬件)消息和系统内部事件
'           发送消息 (再一次)
'            WM_PAINT消息
'           WM_TIMER消息
'        要在发布消息之前检索输入消息，请使用wMsgFilterMin和wMsgFilterMax参数。
'        PeekMessage函数通常不会从队列中删除 WM_PAIN消息?
'         WM_PAIN消息保留在队列中，直到它们被处理。
'        但是，如果 WM_PAIN消息有一个空更新区域，则PeekMessage会将它从队列中删除。
'        如果一个顶级窗口停止对消息的响应超过几秒钟，系统会认为窗口没有响应，并使用一个具有相同z顺序、位置、大小和可视属性的幽灵窗口来替换它。
'        这允许用户移动它，调整大小，甚至关闭应用程序。
'        然而，这些是惟一可用的操作，因为应用程序实际上没有响应。
'        当应用程序被调试时，系统不会生成一个幽灵窗口。
'lpMsg:一个指向一个接收消息信息的MSG结构的指针。
'hWnd:对将要检索消息的窗口的句柄。窗口必须属于当前线程。
'        如果hWnd为空，则PeekMessage将检索属于当前线程的任何窗口的消息，以及当前线程的消息队列上的任何消息，该消息队列的hWnd值为NULL(见MSG结构)。因此，如果hWnd为空，则处理窗口消息和线程消息。
'        如果hWnd是-1，PeekMessage只检索当前线程的消息队列上的消息，该消息队列的hWnd值为NULL，即PostMessage(当hWnd参数为NULL)或PostThreadMessage时所发布的线程消息。
'wMsgFilterMin:在要检查的消息范围内的第一个消息的值。使用WM_KEYFIRST (0x0100)指定第一个键盘消息或WM_MOUSEFIRST (0x0200)来指定第一个鼠标消息。
'                如果wMsgFilterMin和wMsgFilterMax都是0，那么PeekMessage将返回所有可用的消息(也就是说，没有执行范围过滤)。
'wMsgFilterMax: 在要检查的消息范围内的最后一条消息的值?使WM_KEYLAST 指定最后一个键盘消息或WM_MOUSELAST来指定最后的鼠标消息?
'                如果wMsgFilterMin和wMsgFilterMax都是0，那么PeekMessage将返回所有可用的消息(也就是说，没有执行范围过滤)。
'wRemoveMsg: 指定如何处理消息?该参数可以是下列值中的一个或多个?
Public Enum PEEK_REMOVE_TYPE
    PM_NOREMOVE = &H0 '通过PeekMessage处理消息后，不会从队列中删除消息。
    PM_REMOVE = &H1  '通过PeekMessage处理消息后，将从队列中删除消息。
    PM_NOYIELD = &H2 '防止系统释放等待调用者空闲的任何线程(参见WaitForInputIdle)。将该值与PM_NOREMOVE或PM_REMOVE结合起来?
    '默认情况下，所有消息类型都被处理。要指定只处理特定的消息，请指定一个或多个以下值。
'    PM_QS_INPUT=(QS_INPUT < < 16)处理鼠标和键盘消息?
'    PM_QS_PAINT(QS_PAINT < < 16)油漆过程信息?
'    PM_QS_POSTMESSAGE(qspostmessage qshotkey qstimer)处理所有发布的消息，包括计时器和热键。
'    PM_QS_SENDMESSAGE(QS_SENDMESSAGE < < 16)处理所有发送消息?
End Enum
Private Declare Function DispatchMessage Lib "user32" Alias "DispatchMessageA" (lpMsg As Msg) As Long
'功能：向窗口过程发送一条消息。它通常用于分派由GetMessage函数检索的消息。
'返回：返回值指定窗口过程返回的值。尽管它的含义依赖于发送的消息，但是返回值通常被忽略。
'注意事项：MSG结构必须包含有效的消息值。如果lpmsg参数指向一个WM_TIMER消息，并WM_TIMER消息的lParam参数不是空的，lParam将指向一个函数，而不是窗口过程。
'       请注意，应用程序负责检索和向对话框发送输入消息。大多数应用程序使用主消息循环来进行此工作。但是，为了允许用户通过使用键盘来选择和选择控件，应用程序必须调用IsDialogMessage。要了解更多信息，请参见 Dialog Box Keyboard Interface.。
'lpMsg:指向包含消息的结构的指针。

Private Const WM_QUIT As Long = &H12
'说明：指示终止应用程序的请求，并在应用程序调用PostQuitMessage函数时生成。这个消息会导致GetMessage函数返回0。
'注意事项：WM_QUIT消息与窗口没有关联，因此永远不会通过窗口的窗口过程接收。它只能通过GetMessage或PeekMessage函数来检索。
'           不要使用PostMessage函数发布WM_QUIT消息;使用PostQuitMessage。

Private Declare Function CoInitialize Lib "ole32.dll" (ByVal pvReserved As Long) As Long
'功能：在当前线程上初始化COM库，并将并发模型标识为单线程公寓(STA)。新的应用程序应该调用CoInitializeEx，而不是初始化。如果你想使用Windows运行时,必须调Windows::Foundation::Initialize来替代。
'返回：该函数可以返回标准返回值E_INVALIDARG、E_OUTOFMEMORY和E_UNEXPECTED值，以及以下值。
'        返回代码描述
'        S_OK:COM库在这个线程上成功地初始化了?
'        S_FALSE:COM库已经在这个线程上进行了初始化?
'        RPC_E_CHANGED_MODE:之前对CoInitializeEx的调用指定该线程的并发模型为多线程公寓(MTA)。这也可能表明，从中性线公寓到单线程公寓的转变已经发生。
'注意事项：在调用除CoGetMalloc函数之外的任何库函数之前，您需要先在线程上初始化COM库，以获得一个指向标准分配器的指针，以及内存分配函数。
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
'        因为没有办法控制进程内的服务器被加载或卸载的顺序，所以不要从DllMain函数中调用CoInitialize、CoInitializeEx或可数。
'pvReserved:这个参数是保留的，必须是空的。
Private Declare Sub CoUninitialize Lib "ole32.dll" ()
'功能：关闭当前线程上的COM库，卸载线程加载的所有dll，释放线程维护的其他资源，并强制所有的RPC连接在线程上关闭。
'注意事项：一个线程必须为它对CoUninitialize 或CoInitializeEx函数的每个成功调用调用一次CoUninitialize ，包括返回S_FALSE的任何调用。只有对初始化库CoInitialize或CoInitializeEx调用对应的调用才能关闭它。
'            对 OleInitialize 调用必须通过对 OleUninitialize调用来平衡。OleUninitialize 函数调用CoUninitialize ，因此调用OleUninitialize 的应用程序也不需要调用 CoUninitialize。
'            应该在应用程序关闭时调用CoUninitialize ，这是在应用程序隐藏其主窗口并通过其主消息循环后对COM库进行的最后一次调用。如果还存在开放的会话，则可数启动一个模式消息循环，并为这个COM应用程序从容器或服务器发送任何待处理的消息。通过发送消息，CoUninitialize 确保应用程序在接收到所有挂起的消息之前不会退出。Non-COM消息丢弃。
'            因为没有办法控制进程内的服务器被加载或卸载的顺序，所以不要从DllMain函数中调用CoInitialize、CoInitializeEx或可数。

Private Declare Function GetTickCount Lib "kernel32" () As Long
'功能：检索自系统启动以来已运行的毫秒数，最多可达49.7天。
'返回：返回值是自系统启动以来运行的毫秒数。
'注意事项：GetTickCount函数的解析仅限于系统计时器的精度，通常在10毫秒到16毫秒之间。
'        GetTickCount函数的解析不会受到getsystemtime调适函数的调整的影响?
'        经过的时间存储为DWORD的值?
'        因此，如果系统连续运行49.7天，那么时间将会是零。
'        为了避免这个问题，请使用GetTickCount64函数。
'        否则，在比较时检查溢出条件。
'        如果你需要一个更高的分辨率计时器，可以使用多媒体定时器或高分辨率计时器。
'        为了获得计算机启动后的时间，在注册表关键hkeyperformance cedata的性能数据中检索系统的时间计数器。
'        返回的值是一个8字节的值?
'        要了解更多信息，请参见性能计数器。
'        Note：获得时间系统在工作状态自启动 , 使用QueryUnbiasedInterruptTime函数?
'        调试注意QueryUnbiasedInterruptTime函数产生不同的结果(“checked”)构建Windows,因为中断时间计数是由大约49天。
'        这有助于识别在系统运行很长时间之前可能不会发生的错误?
'        通过Microsoft Developer Network(MSDN)Web站点可以对MSDN的用户进行检查。
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
'功能：暂停当前线程的执行，直到超时时间间隔。要进入alertable等待状态，请使用 SleepEx 。
'注意事项：这个函数会导致一个线程放弃剩余的时间片，并在基于dwMilliseconds的时间间隔内变得不可运行。
'        系统时钟以恒定的速率“滴答”。
'        如果dwMilliseconds小于系统时钟的分辨率，那么线程的睡眠时间可能少于指定的时间长度。
'        如果dwMilliseconds大于1，但小于2，那么等待可以在1到2个节拍之间，以此类推。
'        为了提高睡眠时间间隔的准确性，请调用timeGetDevCaps 函数来确定受支持的最小计时器分辨率和timeBeginPeriod 函数，以便将计时器分辨率设置为最小值。
'        当调用timeBeginPeriod时要小心，因为频繁的调用会严重影响系统时钟、系统的使用和调度器。
'        如果您调用timeBeginPeriod，那么在应用程序的早期就调用它，并确保在应用程序的最后调用 timeEndPeriod 函数。
'        在睡眠时间间隔之后，线程就可以运行了。
'        如果指定0毫秒，线程将放弃剩余的时间片，但仍然准备好。
'        注意，一个就绪的线程并不能保证立即运行。
'        因此，线程可能在睡眠时间间隔之后的某一段时间内运行。
'        要了解更多信息，请参见调度优先级。
'        在以下场景中使用睡眠时要小心:
'           直接或间接创建windows的代码(例如，DDE和COM代码初始化)。如果一个线程创建任何窗口，它必须处理消息。消息广播被发送到系统中的所有窗口。如果您的线程使用了无限延迟的睡眠，系统将死锁。
'           在并发控制下的线程。例如，一个输入/输出端口或线程池限制了可以运行的相关线程的数量。如果线程的最大数量已经在运行，那么在运行的线程完成之前，没有其他相关的线程可以运行。
'              如果线程使用睡眠间隔为0，等待其他相关线程完成某些工作，则该进程可能会死锁。
'       对于这些场景 , 使用MsgWaitForMultipleObjects或MsgWaitForMultipleObjectsEx, 而不是Sleep。
'        Windows Phone 8.1:这个功能支持Windows Phone 8.1和以后的Windows Phone Store应用。
'        Windows 8.1和Windows Server 2012 R2:这个功能支持Windows 8.1、Windows Server 2012 R2以及以后的Windows Store应用。
Private Declare Function lstrcpyW Lib "kernel32" (ByVal lpString1 As Any, ByVal lpString2 As Any) As Long

'--------------------------------------------------------------------------------------------------
'方法           IsInstalled
'功能           判断服务是否已安装
'返回值         Boolean
'入参列表:
'参数名         类型                    说明
'strServiceName String                  服务名称
'-------------------------------------------------------------------------------------------------
Public Function IsInstalled(ByVal strServiceName As String) As Boolean
    Dim blnSuccess      As Long
    Dim lnghSCM         As Long
    Dim lnghService     As Long
    Dim qscAll()        As QUERY_SERVICE_CONFIG, lngCount   As Long, i      As Long
    
    On Error GoTo errH
    lnghSCM = OpenSCManager(vbNullString, vbNullString, SC_MANAGER_CONNECT)
    If lnghSCM <> 0 Then
        lnghService = OpenService(lnghSCM, StrPtr(strServiceName), SERVICE_QUERY_CONFIG)
        If lnghService <> 0 Then
            ReDim qscAll(0)
            If QueryServiceConfig(lnghService, qscAll(0), LenB(qscAll(0)), lngCount) = 0 Then
                If Err.LastDllError = ERROR_INSUFFICIENT_BUFFER Then
                    lngCount = (lngCount \ LenB(qscAll(0))) + 1
                    ReDim qscAll(lngCount)
                    If QueryServiceConfig(lnghService, qscAll(0), lngCount * 36, lngCount) <> 0 Then
'                        Debug.Print 0 & "." & qscAll(0).dwErrorControl
'                        Debug.Print 0 & "." & qscAll(0).dwServiceType
'                        Debug.Print 0 & "." & qscAll(0).dwStartType
'                        Debug.Print 0 & "." & qscAll(0).dwTagId
'                        Debug.Print 0 & "." & GetString(qscAll(0).lpBinaryPathName)
'                        Debug.Print 0 & "." & GetString(qscAll(0).lpDependencies)
'                        Debug.Print 0 & "." & GetString(qscAll(0).lpDisplayName)
'                        Debug.Print 0 & "." & GetString(qscAll(0).lpLoadOrderGroup)
'                        Debug.Print 0 & "." & GetString(qscAll(0).lpServiceStartName)
                        blnSuccess = qscAll(0).dwStartType <> SERVICE_DISABLED
                    Else
                        blnSuccess = False
                    End If
                Else
                    blnSuccess = False
                End If
            End If
            Call CloseServiceHandle(lnghService)
        Else
            blnSuccess = False
        End If
        Call CloseServiceHandle(lnghSCM)
    Else
        blnSuccess = False
    End If
    IsInstalled = blnSuccess
    Exit Function
errH:
    Err.Clear
    If lnghService <> 0 Then Call CloseServiceHandle(lnghService)
    If lnghSCM <> 0 Then Call CloseServiceHandle(lnghSCM)
End Function


'--------------------------------------------------------------------------------------------------
'方法           Install
'功能           安装服务
'返回值         Boolean
'入参列表:
'参数名         类型                    说明
'
'-------------------------------------------------------------------------------------------------
Public Function Install(ByVal strServiceName As String, ByVal strDisplayName As String, ByVal strDescription As String, ByVal strServicePath As String, Optional ByVal strAccount As String, Optional ByVal strPassword As String, Optional ByVal sstStartType As SERVICE_START_TYPE = SERVICE_AUTO_START, Optional ByVal sitServiceType As SERVICE_TYPE = SERVICE_WIN32_OWN_PROCESS) As Boolean
    Dim lnghSCM         As Long
    Dim lnghService     As Long
    Dim blnSuccess      As Boolean
    Dim strTmp          As String
    Dim lngSclock       As Long
    
    On Error GoTo errH
    If strAccount = "" Then
        strAccount = "LocalSystem"
    Else
        If strAccount = "LocalSystem" Then
            If InStr(1, strAccount, "\") = 0 Then
                strTmp = GetDomainName()
                If Len(strTmp) = 0 Then strTmp = "."
                strAccount = strTmp & "\" & strAccount
            End If
        End If
    End If
    '打开服务控制管理器数据库
    lnghSCM = OpenSCManager(vbNullString, vbNullString, SC_MANAGER_ALL_ACCESS)
    If lnghSCM <> 0 Then
        lngSclock = LockServiceDatabase(lnghSCM)
        If lngSclock = 0 Then
        End If
        '获得服务对象的句柄,查询方式
        lnghService = OpenService(lnghSCM, StrPtr(strServiceName), SERVICE_ALL_ACCESS)
        If lnghService = 0 Then '服务未安装，则安装
            blnSuccess = True
            lnghService = CreateService(lnghSCM, StrPtr(strServiceName), StrPtr(strDisplayName), SERVICE_ALL_ACCESS, sitServiceType, sstStartType, SERVICE_ERROR_NORMAL, StrPtr(strServicePath), vbNullString, ByVal 0&, vbNullString, vbNullString, vbNullString)
            If lnghService <> 0 Then
                '填冲服务描述
                If ChangeServiceConfig2(lnghService, SERVICE_CONFIG_DESCRIPTION, StrPtr(strDescription)) = 0 Then
                End If
            Else
                blnSuccess = False
            End If
        Else
            '服务安装则退出
            blnSuccess = True
        End If
        '关闭服务对象句柄
        If lnghService <> 0 Then Call CloseServiceHandle(lnghService)
        If lngSclock <> 0 Then Call UnlockServiceDatabase(lngSclock)
        '关闭服务控制管理器
        Call CloseServiceHandle(lnghSCM)
    Else
    End If
    Install = blnSuccess
    Exit Function
errH:
    Err.Clear
    If lnghService <> 0 Then Call CloseServiceHandle(lnghService)
    If lngSclock <> 0 Then Call UnlockServiceDatabase(lngSclock)
    If lnghSCM <> 0 Then Call CloseServiceHandle(lnghSCM)
End Function
'--------------------------------------------------------------------------------------------------
'方法           Uninstall
'功能           卸载服务
'返回值         Boolean
'入参列表:
'参数名         类型                    说明
'strServiceName String                  服务名称
'说明：服务卸载后，程序处于标记删除状态，只有进程终止，才会真正删除。在标记删除状态不能安装。
'-------------------------------------------------------------------------------------------------
Public Function Uninstall(ByVal strServiceName As String) As Boolean
    Dim lnghSCM         As Long
    Dim lnghService     As Long
    Dim blnSuccess      As Boolean
    Dim ssStatus        As SERVICE_STATUS
    Dim lngSclock       As Long
    
    On Error GoTo errH
    '打开服务控制管理器数据库
    lnghSCM = OpenSCManager(vbNullString, vbNullString, SC_MANAGER_ALL_ACCESS)
    If lnghSCM <> 0 Then
        blnSuccess = True
        lngSclock = LockServiceDatabase(lnghSCM)
        If lngSclock = 0 Then
        End If
        '获得服务对象的句柄,查询方式
        lnghService = OpenService(lnghSCM, StrPtr(strServiceName), SERVICE_QUERY_STATUS Or Delete Or SERVICE_STOP)
        If lnghService = 0 Then '服务未安装，则退出
            blnSuccess = True
        Else
            '服务安装则卸载
            If ControlService(lnghService, SERVICE_CONTROL_STOP, ssStatus) = 0 Then
            End If
            If DeleteService(lnghService) = 0 Then
                blnSuccess = False
            End If
            Call CloseServiceHandle(lnghService) '关闭服务对象句柄
        End If
        If lngSclock <> 0 Then Call UnlockServiceDatabase(lngSclock)
        '关闭服务控制管理器
        Call CloseServiceHandle(lnghSCM)
    Else
    End If
    Uninstall = blnSuccess
    Exit Function
errH:
    Err.Clear
    If lnghService <> 0 Then Call CloseServiceHandle(lnghService)
    If lngSclock <> 0 Then Call UnlockServiceDatabase(lngSclock)
    If lnghSCM <> 0 Then Call CloseServiceHandle(lnghSCM)
End Function
'--------------------------------------------------------------------------------------------------
'方法           IsServiceRunning
'功能           服务是否处于运行状态
'返回值         Boolean
'入参列表:
'参数名         类型                    说明
'
'-------------------------------------------------------------------------------------------------
Public Function IsRunning(ByVal strServiceName As String) As Boolean
    On Error GoTo errH
    IsRunning = GetStatus(strServiceName) = SERVICE_RUNNING
    Exit Function
errH:
    Err.Clear
End Function
'--------------------------------------------------------------------------------------------------
'方法           Start
'功能           启动服务
'返回值         Boolean
'入参列表:
'参数名         类型                    说明
'strServiceName String                  服务名称
'-------------------------------------------------------------------------------------------------
Public Function Start(ByVal strServiceName As String) As Boolean
    Dim lnghSCM         As Long
    Dim lnghService     As Long
    Dim blnSuccess      As Boolean
    Dim ssStatus As SERVICE_STATUS
    
    On Error GoTo errH
    '打开服务控制管理器数据库
    lnghSCM = OpenSCManager(vbNullString, vbNullString, SC_MANAGER_CONNECT)
    If lnghSCM <> 0 Then
        blnSuccess = True
        '获得服务对象的句柄,查询方式
        lnghService = OpenService(lnghSCM, StrPtr(strServiceName), SERVICE_START)
        If lnghService = 0 Then '服务未安装，则退出
            blnSuccess = False
        Else
            '服务安装则启动
            If StartService(lnghService, ByVal 0, ByVal 0) = 0 Then
                blnSuccess = False
            End If
            Call CloseServiceHandle(lnghService) '关闭服务对象句柄
        End If
        '关闭服务控制管理器
        Call CloseServiceHandle(lnghSCM)
    Else
    End If
    Start = blnSuccess
    Exit Function
errH:
    Err.Clear
    If lnghService <> 0 Then Call CloseServiceHandle(lnghService)
    If lnghSCM <> 0 Then Call CloseServiceHandle(lnghSCM)
End Function
'--------------------------------------------------------------------------------------------------
'方法           IsStoppinged
'功能           服务是否处于停止状态
'返回值         Boolean
'入参列表:
'参数名         类型                    说明
'
'-------------------------------------------------------------------------------------------------
Public Function IsStopped(ByVal strServiceName As String) As Boolean
    On Error GoTo errH
    IsStopped = GetStatus(strServiceName) = SERVICE_STOPPED
    Exit Function
errH:
    Err.Clear
End Function
'--------------------------------------------------------------------------------------------------
'方法           Stopping
'功能           停止服务
'返回值         Boolean
'入参列表:
'参数名         类型                    说明
'strServiceName String                  服务名称
'-------------------------------------------------------------------------------------------------
Public Function Stopping(ByVal strServiceName As String) As Boolean
    Dim lnghSCM         As Long
    Dim lnghService     As Long
    Dim blnSuccess      As Boolean
    Dim ssStatus        As SERVICE_STATUS
    
    On Error GoTo errH
    '打开服务控制管理器数据库
    lnghSCM = OpenSCManager(vbNullString, vbNullString, SC_MANAGER_CONNECT)
    If lnghSCM <> 0 Then
        blnSuccess = True
        '获得服务对象的句柄,查询方式
        lnghService = OpenService(lnghSCM, StrPtr(strServiceName), SERVICE_STOP)
        If lnghService = 0 Then '服务未安装，则退出
            blnSuccess = False
        Else
            '服务安装则停止
            If ControlService(lnghService, SERVICE_CONTROL_STOP, ssStatus) = 0 Then
                blnSuccess = False
            End If
            Call CloseServiceHandle(lnghService) '关闭服务对象句柄
        End If
        '关闭服务控制管理器
        Call CloseServiceHandle(lnghSCM)
    Else
    End If
    Stopping = blnSuccess
    Exit Function
errH:
    Err.Clear
    If lnghService <> 0 Then Call CloseServiceHandle(lnghService)
    If lnghSCM <> 0 Then Call CloseServiceHandle(lnghSCM)
End Function
'--------------------------------------------------------------------------------------------------
'方法           AutoRun
'功能           设置服务为自启动
'返回值         Boolean
'入参列表:
'参数名         类型                    说明
'strServiceName String                  服务名称
'-------------------------------------------------------------------------------------------------
Public Function AutoRun(ByVal strServiceName As String) As Boolean
    Dim lnghSCM         As Long
    Dim lnghService     As Long
    Dim blnSuccess      As Boolean
    Dim ssStatus        As SERVICE_STATUS
    Dim lngSclock       As Long
    
    On Error GoTo errH
    '打开服务控制管理器数据库
    lnghSCM = OpenSCManager(vbNullString, vbNullString, SC_MANAGER_ALL_ACCESS)
    If lnghSCM <> 0 Then
        blnSuccess = True
        lngSclock = LockServiceDatabase(lnghSCM)
        If lngSclock = 0 Then
        End If
        '获得服务对象的句柄,查询方式
        lnghService = OpenService(lnghSCM, StrPtr(strServiceName), SERVICE_ALL_ACCESS)
        If lnghService = 0 Then '服务未安装，则退出
            blnSuccess = False
        Else
            '服务安装则停止。该地方调用不了，查询资料无果
            If ChangeServiceConfig(lnghService, SERVICE_NO_CHANGE, SERVICE_AUTO_START, SERVICE_ERROR_NORMAL, vbNullString, vbNullString, 0, vbNullString, vbNullString, vbNullString, vbNullString) = 0 Then
                blnSuccess = False
            End If
            Call CloseServiceHandle(lnghService) '关闭服务对象句柄
        End If
        If lngSclock <> 0 Then Call UnlockServiceDatabase(lngSclock)
        '关闭服务控制管理器
        Call CloseServiceHandle(lnghSCM)
    Else
    End If
    AutoRun = blnSuccess
    Exit Function
errH:
    Err.Clear
    If lnghService <> 0 Then Call CloseServiceHandle(lnghService)
    If lngSclock <> 0 Then Call UnlockServiceDatabase(lngSclock)
    If lnghSCM <> 0 Then Call CloseServiceHandle(lnghSCM)
End Function
'--------------------------------------------------------------------------------------------------
'方法           GetStatus
'功能           查询服务状态
'返回值         SERVICE_STATE
'入参列表:
'参数名         类型                    说明
'strServiceName String                  服务名称
'-------------------------------------------------------------------------------------------------
Private Function GetStatus(ByVal strServiceName As String) As SERVICE_STATE
    Dim lnghSCM         As Long
    Dim lnghService     As Long
    Dim ssStatus        As SERVICE_STATUS
    
    On Error GoTo errH
    lnghSCM = OpenSCManager(vbNullString, vbNullString, SC_MANAGER_CONNECT)
    If lnghSCM <> 0 Then
        lnghService = OpenService(lnghSCM, StrPtr(strServiceName), SERVICE_QUERY_STATUS)
        If lnghService <> 0 Then
            If QueryServiceStatus(lnghService, ssStatus) <> 0 Then
                GetStatus = ssStatus.dwCurrentState
            Else
            End If
            Call CloseServiceHandle(lnghService)
        Else
        End If
        Call CloseServiceHandle(lnghSCM)
    Else
    End If
    Exit Function
errH:
    Err.Clear
    If lnghService <> 0 Then Call CloseServiceHandle(lnghService)
    If lnghSCM <> 0 Then Call CloseServiceHandle(lnghSCM)
End Function

'--------------------------------------------------------------------------------------------------
'方法           GetDomainName
'功能           获取域名或工作组
'返回值         String
'入参列表:
'参数名         类型                    说明
'
'-------------------------------------------------------------------------------------------------
Private Function GetDomainName() As String
    Dim lpBuffer As Long, lngLen As Long, p As Long
    
    On Error GoTo errH
    If NetWkstaUserGetInfo(0&, 1&, lpBuffer) = 0 Then
        CopyMemory p, ByVal lpBuffer + 4, 4
        lngLen = lstrlenW(p)
        If lngLen > 0 Then
            GetDomainName = Space$(lngLen)
            CopyMemory ByVal StrPtr(GetDomainName), ByVal p, lngLen * 2
        End If
        NetApiBufferFree lpBuffer
    End If
    Exit Function
errH:
    Err.Clear
    If lpBuffer <> 0 Then NetApiBufferFree lpBuffer
End Function

Private Function GetString(ByVal lngStrPtr As Long) As String
    Dim strTmp  As String
    
    If lngStrPtr <> 0 Then
        strTmp = Space$(lstrlenW(lngStrPtr))
        lstrcpyW StrPtr(strTmp), lngStrPtr
        GetString = strTmp
    End If
End Function

