VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEnvironment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@模块 clsEnvironment-2019/7/3
'@编写 lshuo
'@功能
'   获取运行环境信息
'@引用
'
'@备注
'
Option Explicit
'---------------------------------------------------------------------------
'                0、API和常量声明
'---------------------------------------------------------------------------
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
Private Declare Function EnumProcessModules Lib "psapi.dll" (ByVal hProcess As Long, ByRef lphModule As Long, ByVal cb As Long, ByRef cbNeeded As Long) As Long
'@功能
'    检索指定流程中每个模块的句柄?
'    要控制一个64位应用程序枚举32位模块、64位模块还是这两种类型的模块，可以使用EnumProcessModulesEx函数。
'@原型
'    BOOL EnumProcessModules(
'      HANDLE  hProcess,
'      HMODULE *lphModule,
'      DWORD   cb,
'      LPDWORD lpcbNeeded
'    );
'@参数
'hProcess
'    进程的句柄?
'lphModule
'    接收模块句柄列表的数组?
'cb
'    lphModule数组的大小，以字节为单位。
'lpcbNeeded
'    在lphModule数组中存储所有模块句柄所需的字节数?
'@返回值
'    如果函数成功，返回值为非零。
'    如果函数失败，返回值为零。要获取扩展的错误信息，请调用GetLastError。
'@备注
'    EnumProcessModules函数主要是为调试器和类似的应用程序设计的，这些应用程序必须从另一个进程提取模块信息。如果目标进程中的模块列表已损坏或尚未初始化，或者模块列表在函数调用期间由于加载或卸载dll而发生更改，EnumProcessModules可能会失败或返回不正确的信息。
'    指定大量HMODULE值是一个好主意，因为在调用EnumProcessModules时很难预测流程中有多少个模块。要确定lphModule数组是否太小，无法容纳进程的所有模块句柄，请将lpcbrequired中返回的值与cb中指定的值进行比较。如果lpcbrequired大于cb，则增加数组的大小并再次调用EnumProcessModules。
'    要确定调用EnumProcessModules枚举了多少个模块，请将lpcbrequired参数中的结果值除以sizeof(HMODULE)。
'    EnumProcessModules函数不检索加载了LOAD_LIBRARY_AS_DATAFILE或类似标志的模块的句柄。有关更多信息，请参见LoadLibraryEx。
'    不要对这个函数返回的任何句柄调用close句柄。信息来自快照，因此没有要释放的资源。
'    如果从运行在WOW64上的32位应用程序调用此函数，则它只能枚举32位进程的模块。如果进程是64位进程，这个函数将失败，最后一个错误代码是ERROR_PARTIAL_COPY(299)。
'    要获取指定进程以及这些进程使用的堆、模块和线程的快照，请使用CreateToolhelp32Snapshot函数。
'    从Windows 7和Windows Server 2008 R2开始，PSAPI .h为PSAPI函数建立版本号。PSAPI版本号影响用于调用函数的名称和程序必须加载的库。
'    如果PSAPI_VERSION大于等于2，则在Psapi.h中将此函数定义为K32EnumProcessModules，并在Kernel32中导出。自由和Kernel32.dll。如果PSAPI_VERSION为1，则此函数在Psapi.h中定义为EnumProcessModules，并在Psapi中导出。dll作为调用k32enumprocessmodule的包装器。
'    必须在Windows早期版本以及Windows 7和更高版本上运行的程序应该始终将此函数调用为EnumProcessModules。要确保符号的正确分辨率，请添加Psapi。lib到TARGETLIBS宏，并使用-DPSAPI_VERSION=1编译程序。要使用运行时动态链接，请加载Psapi.dll。
'@Requirements
'    Minimum supported client    Windows XP [desktop apps only]
'    Minimum supported server    Windows Server 2003 [desktop apps only]
'    Target Platform Windows
'    Header psapi.h
'    Library Kernel32.lib on Windows 7 and Windows Server 2008 R2; Psapi.lib (if PSAPI_VERSION=1) on Windows 7 and Windows Server 2008 R2; Psapi.lib on Windows Server 2008, Windows Vista, Windows Server 2003 and Windows XP
'    DLL Kernel32.dll on Windows 7 and Windows Server 2008 R2; Psapi.dll (if PSAPI_VERSION=1) on Windows 7 and Windows Server 2008 R2; Psapi.dll on Windows Server 2008, Windows Vista, Windows Server 2003 and Windows XP

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
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
'@原型
'    BOOL GetComputerNameA(
'      LPSTR   lpBuffer,
'      LPDWORD nSize
'    );
'@功能
'    检索本地计算机的NetBIOS名称。这个名称是在系统启动时，当系统从注册表中读取它时建立的。
'    GetComputerName只检索本地计算机的NetBIOS名称。要检索DNS主机名、DNS域名或完全限定的DNS名称，请调用GetComputerNameEx函数。附加信息由IADsADSystemInfo接口提供。
'    如果本地计算机是集群中的节点，则此函数的行为可能受到影响。有关更多信息，请参见ResUtilGetEnvironmentWithNetName和UseNetworkName。
'@参数
Private Const MAX_COMPUTERNAME_LENGTH           As Long = 31
'lpBuffer
'    指向缓冲区的指针，该缓冲区接收计算机名或集群虚拟服务器名。缓冲区大小应该足够大，可以包含MAX_COMPUTERNAME_LENGTH + 1个字符。
'nSize
'    作为输入，在TCHARs中指定缓冲区的大小。在输出时，复制到目标缓冲区的tchar数，不包括终止空字符。
'    如果缓冲区太小，函数就会失败，GetLastError返回ERROR_BUFFER_OVERFLOW。lpnSize参数指定所需缓冲区的大小，包括终止null字符。
'
'@返回值
'    如果函数成功，返回值为非零值。
'    如果函数失败，返回值为零。要获取扩展的错误信息，请调用GetLastError。
'@备注
'    GetComputerName函数检索在系统启动时建立的NetBIOS名称。在用户重新启动计算机之前，SetComputerName或SetComputerNameEx函数所做的名称更改不会生效。
'    如果调用者在客户端会话下运行，此函数将返回服务器名称。要检索客户端名称，请使用WTSQuerySessionInformation函数。
'@Requirements
'    Minimum supported client    Windows 2000 Professional [desktop apps | UWP apps]
'    Minimum supported server    Windows 2000 Server [desktop apps | UWP apps]
'    Target Platform Windows
'    Header  winbase.h (include Windows.h)
'    Library kernel32.lib
'    dll kernel32.dll
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
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

Private Const ANY_SIZE          As Long = 1
Private Type MIB_IPADDRROW_W2K
    dwAddr          As Long
    dwIndex         As Long
    dwMask          As Long
    dwBCastAddr     As Long
    dwReasmSize     As Long
    unused1         As Integer
    unused2         As Integer
End Type
'@功能
'    mib_ipaddrrowin MIB_IPADDRTABLE结构中指定特定IPv4地址的信息。
'@原型
'    typedef struct _MIB_IPADDRROW_W2K {
'      DWORD          dwAddr;
'      DWORD          dwIndex;
'      DWORD          dwMask;
'      DWORD          dwBCastAddr;
'      DWORD          dwReasmSize;
'      unsigned short unused1;
'      unsigned short unused2;
'    } MIB_IPADDRROW_W2K, *PMIB_IPADDRROW_W2K;
'@成员
'dwAddr
'    按网络字节顺序排列的IPv4地址?
'dwIndex
'    与此IPv4地址关联的接口的索引?
'dwMask
'    IPv4地址的子网掩码，按网络字节顺序排列。
'dwBCastAddr
'    网络字节顺序的广播地址。广播地址通常是IPv4地址，主机部分设置为所有0或所有1。
'    GetIpAddrTable函数不返回此成员的适当值?
'dwReasmSize
'    接收到的数据报的最大重新组装大小?
'unused1
'    此成员已被保留?
'unused2
'    此成员已被保留?
'@备注
'    在Windows XP及以后版本中，MIB_IPADDRROW结构的dwIndex成员的数据类型为IF_INDEX。wType成员只能在Windows XP或更高版本上使用。在Windows 2000或更早的版本中，这个成员被定义为Unused2。
'    GetIpAddrTable函数检索本地计算机上的interto - ipv4地址映射表，并以MIB_IPADDRTABLE结构返回此信息。MIB_IPADDRTABLE结构中的表成员包含一个MIB_IPADDRROW条目数组。
'    在为Windows Vista及以后版本发布的Microsoft Windows软件开发工具包(SDK)上，头文件的组织发生了变化，MIB_IPADDRROW结构是在Ipmib.h头文件中定义的，而不是在Iprtrmib.h头文件中定义的。注意，Ipmib.h头文件自动包含在iprtrmi .h中，iprtrmi .h自动包含在Iphlpapi.h头文件中。不应该直接使用Ipmib.h和Iprtrmib.h头文件。
'@Requirements
'    Minimum supported client    Windows 2000 Professional [desktop apps only]
'    Minimum supported server    Windows 2000 Server [desktop apps only]
'    Header  ipmib.h (include Iphlpapi.h)
Private Type MIB_IPADDRROW_XP
    dwAddr          As Long
    dwIndex         As Long
    dwMask          As Long
    dwBCastAddr     As Long
    dwReasmSize     As Long
    unused1         As Integer
    wType           As Integer
End Type
'@功能
'    mib_ipaddrrowin MIB_IPADDRTABLE结构中指定特定IPv4地址的信息
'@原型
'    typedef struct _MIB_IPADDRROW_XP {
'      DWORD          dwAddr;
'      IF_INDEX       dwIndex;
'      DWORD          dwMask;
'      DWORD          dwBCastAddr;
'      DWORD          dwReasmSize;
'      unsigned short unused1;
'      unsigned short wType;
'    } MIB_IPADDRROW_XP, *PMIB_IPADDRROW_XP;
'@成员
'dwAddr
'    按网络字节顺序排列的IPv4地址?
'dwIndex
'    与此IPv4地址关联的接口的索引?
'dwMask
'    IPv4地址的子网掩码，按网络字节顺序排列。
'dwBCastAddr
'    网络字节顺序的广播地址。广播地址通常是IPv4地址，主机部分设置为所有0或所有1。
'    GetIpAddrTable函数不返回此成员的适当值?
'dwReasmSize
'    接收到的数据报的最大重新组装大小?
'unused1
'    此成员已被保留?
'wType
'    地址类型或状态?这个成员可以是以下值的组合?
Private Const MIB_IPADDR_PRIMARY            As Integer = &H1
'    主要的IP地址
Private Const MIB_IPADDR_DYNAMIC            As Integer = &H4
'    动态IP地址
Private Const MIB_IPADDR_DISCONNECTED       As Integer = &H8
'    地址在断开连接的接口上
Private Const MIB_IPADDR_DELETED            As Integer = &H40
'    地址被删除
Private Const MIB_IPADDR_TRANSIENT          As Integer = &H80
'    临时地址
'@备注
'    在Windows XP及以后版本中，MIB_IPADDRROW结构的dwIndex成员的数据类型为IF_INDEX。wType成员只能在Windows XP或更高版本上使用。在Windows 2000或更早的版本中，这个成员被定义为Unused2。
'    GetIpAddrTable函数检索本地计算机上的interto - ipv4地址映射表，并以MIB_IPADDRTABLE结构返回此信息。MIB_IPADDRTABLE结构中的表成员包含一个MIB_IPADDRROW条目数组。
'    在为Windows Vista及以后版本发布的Microsoft Windows软件开发工具包(SDK)上，头文件的组织发生了变化，MIB_IPADDRROW结构是在Ipmib.h头文件中定义的，而不是在Iprtrmib.h头文件中定义的。注意，Ipmib.h头文件自动包含在iprtrmi .h中，iprtrmi .h自动包含在Iphlpapi.h头文件中。不应该直接使用Ipmib.h和Iprtrmib.h头文件。
'@Requirements
'    Minimum supported client    Windows 2000 Professional [desktop apps only]
'    Minimum supported server    Windows 2000 Server [desktop apps only]
'    Header  ipmib.h (include Iphlpapi.h)
Private Type MIB_IPADDRTABLE
    dwNumEntries            As Long
    table(ANY_SIZE - 1)     As MIB_IPADDRROW_XP
End Type
'@功能
'    MIB_IPADDRTABLE结构包含一个IPv4地址条目表。
'@原型
'    typedef struct _MIB_IPADDRTABLE {
'      DWORD         dwNumEntries;
'      MIB_IPADDRROW table[ANY_SIZE];
'    } MIB_IPADDRTABLE, *PMIB_IPADDRTABLE;
'@成员
'dwNumEntries
'    表中IPv4地址项的数量?
'table
'    指向IPv4地址项表的指针，该表以MIB_IPADDRROW结构数组的形式实现。
'@备注
'    GetIpAddrTable函数检索本地计算机上的interto - ipv4地址映射表，并以MIB_IPADDRTABLE结构返回此信息。
'    MIB_IPADDRTABLE结构可能包含用于dwNumEntries成员和表成员中的第一个MIB_IPADDRROW数组条目之间对齐的填充?表成员中的MIB_IPADDRROW数组条目之间也可能存在用于对齐的填充?对MIB_IPADDRROW数组条目的任何访问都应该假设可能存在填充?
'    在为Windows Vista及其后版本发布的Microsoft Windows软件开发工具包(SDK)上，头文件的组织发生了变化，MIB_IPADDRROW定义在Ipmib.h头文件中，而不是在Iprtrmib.h头文件中。注意，Ipmib.h头文件自动包含在iprtrmi .h中，iprtrmi .h自动包含在Iphlpapi.h头文件中。不应该直接使用Ipmib.h和Iprtrmib.h头文件。

Private Declare Function GetIpAddrTable Lib "IPHlpApi" (pIPAdrTable As Long, pdwSize As Long, ByVal bOrder As Long) As Long 'MD5计算
'@原型
'    IPHLPAPI_DLL_LINKAGE DWORD GetIpAddrTable(
'      PMIB_IPADDRTABLE pIpAddrTable,
'      PULONG           pdwSize,
'      BOOL bOrder
'    );
'@功能
'    GetIpAddrTable函数检索接口到ipv4地址映射表。
'@参数
'pIpAddrTable
'    一个指向缓冲区的指针，该缓冲区接收作为MIB_IPADDRTABLE结构的接口到ipv4地址映射表。
'pdwSize
'    在输入时指定pIpAddrTable参数指向的缓冲区的大小(以字节为单位)。
'    在输出时，如果缓冲区不够大，无法容纳返回的映射表，则函数将此参数设置为所需的缓冲区大小(以字节为单位)。
'bOrder
'    一个布尔值，指定返回的映射表是否应按IPv4地址升序排序。如果该参数为真，则对表进行排序。
'@返回值
'    如果函数成功，返回值为NO_ERROR。
'    如果函数失败，返回值是以下错误代码之一。
'        返回代码描述
'    ERROR_INSUFFICIENT_BUFFER
'        pIpAddrTable参数指向的缓冲区不够大?所需的大小在pdwSize参数指向的DWORD变量中返回?
'    ERROR_INVALID_PARAMETER
'        pdwSize参数为NULL，或者GetIpAddrTable无法写入pdwSize参数指向的内存。
Private Const ERROR_NOT_SUPPORTED       As Long = 50
'        本地系统上使用的操作系统不支持此功能?
'    其他
'        使用FormatMessage获取返回错误的消息字符串?
'@备注
'    GetIpAddrTable函数检索本地计算机上的interto - ipv4地址映射表，并以MIB_IPADDRTABLE结构返回此信息。
'    GetIpAddrTable函数返回的IPv4地址受本地计算机上网络接口状态的影响。手动重置网络接口卡(NIC)和某些PnP事件可能导致IP地址被删除或更改。
'    在Windows Server 2003和Windows XP上，如果通过调用DisableMediaSense函数禁用本地计算机上TCP/IP堆栈的媒体感知能力，GetIpAddrTable函数返回的IPv4地址也会受到影响。当禁用媒体感知时，GetIpAddrTable函数可能返回与断开连接的接口关联的IPv4地址。这些用于断开连接的接口的Ipv4地址不能使用。
'    在Windows Server 2008和Windows Vista上，GetIpAddrTable函数返回的IPv4地址不受本地计算机上TCP/IP堆栈的媒体感知能力的影响。GetIpAddrTable函数只返回有效的IPv4地址。
'    在Windows XP上可用的getadaptersaddress函数可用于检索IPv6和IPv4地址以及接口信息?
'    GetIpAddrTable函数返回的MIB_IPADDRTABLE结构可能包含用于dwNumEntries成员和表成员中的第一个MIB_IPADDRROW数组条目之间对齐的填充?表成员中的MIB_IPADDRROW数组条目之间也可能存在用于对齐的填充?对MIB_IPADDRROW数组条目的任何访问都应该假设可能存在填充?
'    在为Windows Vista及其后版本发布的Microsoft Windows软件开发工具包(SDK)上，头文件的组织发生了变化，MIB_IPADDRROW定义在Ipmib.h头文件中，而不是在Iprtrmib.h头文件中。注意，Ipmib.h头文件自动包含在iprtrmi .h中，iprtrmi .h自动包含在Iphlpapi.h头文件中。不应该直接使用Ipmib.h和Iprtrmib.h头文件。
'@Requirements
'    Minimum supported client    Windows 2000 Professional [desktop apps only]
'    Minimum supported server    Windows 2000 Server [desktop apps only]
'    Target Platform Windows
'    Header IPHlpApi.h
'    Library IPHlpApi.lib
'    dll IPHlpApi.dll
Private Declare Function GetModuleFileName Lib "kernel32" Alias "GetModuleFileNameA" (ByVal hModule As Long, ByVal lpFilename As String, ByVal nSize As Long) As Long
'@原型
'    DWORD GetModuleFileNameA(
'      HMODULE hModule,
'      LPSTR   lpFilename,
'      DWORD nSize
'    );
'@功能
'    检索包含指定模块的文件的完全限定路径?模块必须由当前进程加载?
'    要定位由另一个进程加载的模块的文件，请使用GetModuleFileNameEx函数。
'@参数
'hModule
'    正在请求其路径的已加载模块的句柄?如果该参数为NULL , GetModuleFileName将检索当前进程的可执行文件的路径?
'    GetModuleFileName函数不检索使用LOAD_LIBRARY_AS_DATAFILE标志加载的模块的路径。有关更多信息，请参见LoadLibraryEx。
'lpFileName
'    指向缓冲区的指针，该缓冲区接收模块的完全限定路径。如果路径的长度小于nSize参数指定的大小，则函数成功，并将路径作为以null结尾的字符串返回。
'    如果路径的长度超过了nSize参数指定的大小，则函数成功，字符串被截断为nSize字符，包括终止null字符。
'    Windows XP:字符串被截断为nSize字符，不以null结尾。
'    返回的字符串将使用模块加载时指定的相同格式。因此，路径可以是长文件名，也可以是短文件名，并且可以使用前缀“\?”。有关更多信息，请参见命名文件。
'nSize
'    lpFilename缓冲区的大小，在TCHARs中。
'@返回值
'    如果函数成功，返回值是复制到缓冲区的字符串的长度，以字符为单位，不包括终止null字符。如果缓冲区太小，无法保存模块名，则将字符串截断为nSize字符，包括终止null字符，函数返回nSize，函数将最后一个错误设置为ERROR_INSUFFICIENT_BUFFER。
'    如果缓冲区太小，无法保存模块名，函数返回nSize。最后一个错误代码仍然是ERROR_SUCCESS。如果nSize为零，则返回值为零，最后一个错误代码为ERROR_SUCCESS。
'    如果函数失败，返回值为0(0)。要获取扩展的错误信息，请调用GetLastError。
'@备注
'    如果DLL在两个进程中加载，则一个进程中的文件名可能与另一个进程中的文件名不同。
'    全局变量_pgmptr被自动初始化为可执行文件的完整路径，并可用于检索可执行文件的完整路径名。
'@Requirements
'    Minimum supported client    Windows XP [desktop apps | UWP apps]
'    Minimum supported server    Windows Server 2003 [desktop apps | UWP apps]
'    Target Platform Windows
'    Header  libloaderapi.h (include Windows.h)
'    Library kernel32.lib
'    dll kernel32.dll
Private Declare Function GetModuleFileNameEx Lib "psapi.dll" Alias "GetModuleFileNameExA" (ByVal hProcess As Long, ByVal hModule As Long, ByVal lpFilename As String, ByVal nSize As Long) As Long
'@原型
'    DWORD GetModuleFileNameExA(
'      HANDLE  hProcess,
'      HMODULE hModule,
'      LPSTR   lpFilename,
'      DWORD nSize
'    );
'@功能
'    检索包含指定模块的文件的完全限定路径。
'@参数
'hProcess
'    包含模块的进程的句柄?
'    句柄必须具有PROCESS_QUERY_INFORMATION和PROCESS_VM_READ访问权限。有关更多信息，请参见流程安全和访问权限。
'    GetModuleFileNameEx函数不检索使用LOAD_LIBRARY_AS_DATAFILE标志加载的模块的路径。有关更多信息，请参见LoadLibraryEx。
'hModule
'    模块的句柄?如果该参数为NULL , GetModuleFileNameEx将返回hProcess中指定的进程的可执行文件的路径?
'lpFilename
'    指向缓冲区的指针，该缓冲区接收到模块的完全限定路径。如果文件名的大小大于nSize参数的值，函数将成功执行，但是文件名将被截断并以null结尾。
'nSize
'    lpFilename缓冲区的大小，以字符为单位。
'@返回值
'    如果函数成功，返回值指定复制到缓冲区的字符串的长度。
'    如果函数失败，返回值为零。要获取扩展的错误信息，请调用GetLastError。
'@备注
'    GetModuleFileNameEx函数主要是为调试器和类似的应用程序设计的，这些应用程序必须从另一个进程提取模块信息。如果目标进程中的模块列表已损坏或尚未初始化，或者模块列表在函数调用期间由于加载或卸载dll而发生更改，GetModuleFileNameEx可能会失败或返回不正确的信息。
'    要检索当前进程中模块的名称，请使用GetModuleFileName函数。这比使用当前进程的句柄调用GetModuleFileNameEx更有效、更可靠。
'    要检索远程进程的主可执行模块的名称，请使用GetProcessImageFileName或QueryFullProcessImageName函数。这比使用NULL模块句柄调用GetModuleFileNameEx函数更有效、更可靠。
'    从Windows 7和Windows Server 2008 R2开始，PSAPI .h为PSAPI函数建立版本号。PSAPI版本号影响用于调用函数的名称和程序必须加载的库。
'    如果PSAPI_VERSION大于或等于2，则该函数在Psapi.h中定义为K32GetModuleFileNameEx，并在Kernel32中导出。自由和Kernel32.dll。如果PSAPI_VERSION为1，则此函数在Psapi.h中定义为GetModuleFileNameEx，并在Psapi中导出。库和Psapi.dll作为调用K32GetModuleFileNameEx的包装器。
'    必须运行在Windows早期版本以及Windows 7和更高版本上的程序应该始终将此函数调用为GetModuleFileNameEx。要确保符号的正确分辨率，请添加Psapi。lib到TARGETLIBS宏，并使用-DPSAPI_VERSION=1编译程序。要使用运行时动态链接，请加载Psapi.dll。
'@Requirements
'    Minimum supported client    Windows XP [desktop apps | UWP apps]
'    Minimum supported server    Windows Server 2003 [desktop apps | UWP apps]
'    Target Platform Windows
'    Header psapi.h
'    Library Kernel32.lib on Windows 7 and Windows Server 2008 R2; Psapi.lib (if PSAPI_VERSION=1) on Windows 7 and Windows Server 2008 R2; Psapi.lib on Windows Server 2008, Windows Vista, Windows Server 2003 and Windows XP
'    DLL Kernel32.dll on Windows 7 and Windows Server 2008 R2; Psapi.dll (if PSAPI_VERSION=1) on Windows 7 and Windows Server 2008 R2; Psapi.dll on Windows Server 2008, Windows Vista, Windows Server 2003 and Windows XP

Private Const MAX_PATH          As Long = 260
Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
'@原型
'    HMODULE GetModuleHandleA(
'      LPCSTR lpModuleName
'    );
'@功能
'    检索指定模块的模块句柄?模块必须是由调用进程加载的?
'    要避免备注部分中描述的竞态条件，请使用GetModuleHandleEx函数。
'@参数
'lpModuleName
'    加载模块的名称(.dll或.exe文件)。如果省略文件名扩展名，则添加默认的库扩展名.dll。文件名字符串可以包含一个尾随点字符(.)，以指示模块名称没有扩展名。字符串不需要指定路径。在指定路径时，请确保使用反斜杠()，而不是前斜杠(/)。将名称(独立于大小写)与当前映射到调用进程的地址空间的模块名称进行比较。
'    如果该参数为NULL, GetModuleHandle将返回用于创建调用进程的文件的句柄(.exe文件)。
'    GetModuleHandle函数不检索使用LOAD_LIBRARY_AS_DATAFILE标志加载的模块的句柄。有关更多信息，请参见LoadLibraryEx。
'@返回值
'    如果函数成功，返回值是指定模块的句柄。
'    如果函数失败，返回值为NULL。要获取扩展的错误信息，请调用GetLastError。
'@备注
'    返回的句柄不是全局的或可继承的?它不能被其他进程复制或使用?
'    如果lpModuleName不包含路径，并且有多个加载的模块具有相同的基本名称和扩展名，则无法预测将返回哪个模块句柄。要解决这个问题，可以指定路径，使用并排的程序集，或者使用GetModuleHandleEx来指定内存位置，而不是DLL名称。
'    GetModuleHandle函数返回一个映射模块的句柄，而不增加它的引用计数。但是，如果将此句柄传递给FreeLibrary函数，则映射模块的引用计数将递减。因此，不要将GetModuleHandle返回的句柄传递给FreeLibrary函数。这样做会导致DLL模块过早地未映射。
'    在多线程应用程序中必须谨慎使用此函数。不能保证模块句柄在函数返回句柄和使用它之间保持有效。例如，假设一个线程检索一个模块句柄，但是在使用该句柄之前，第二个线程释放该模块。如果系统加载另一个模块，它可以重用最近释放的模块句柄。因此，第一个线程将拥有一个不同于预期模块的句柄。
'@Requirements
'    Minimum supported client    Windows XP [desktop apps only]
'    Minimum supported server    Windows Server 2003 [desktop apps only]
'    Target Platform Windows
'    Header  libloaderapi.h (include Windows.h)
'    Library kernel32.lib
'    dll kernel32.dll
Private Declare Function GetModuleHandleEx Lib "kernel32" Alias "GetModuleHandleExA" (ByVal lpModuleName As String) As Long
'@原型
'    BOOL GetModuleHandleExA(
'      DWORD   dwFlags,
'      LPCSTR  lpModuleName,
'      HMODULE *phModule
'    );
'@功能
'    检索指定模块的模块句柄并增加模块的引用计数，除非指定GET_MODULE_HANDLE_EX_FLAG_UNCHANGED_REFCOUNT。模块必须是由调用进程加载的
'@参数
'dwFlags
'    此参数可以是以下值中的零个、一个或多个。如果模块的引用计数增加，当不再需要模块句柄时，调用者必须使用FreeLibrary函数来减少引用计数。
Private Const GET_MODULE_HANDLE_EX_FLAG_FROM_ADDRESS                    As Long = &H4
'    lpModuleName参数是模块中的一个地址?
Private Const GET_MODULE_HANDLE_EX_FLAG_PIN                             As Long = &H1
'    无论调用了多少次FreeLibrary，模块都将保持加载状态，直到进程终止。
'    此选项不能与GET_MODULE_HANDLE_EX_FLAG_UNCHANGED_REFCOUNT一起使用?
Private Const GET_MODULE_HANDLE_EX_FLAG_UNCHANGED_REFCOUNT              As Long = &H2
'    模块的引用计数不增加。这个选项等价于GetModuleHandle的行为。不要将检索到的模块句柄传递给FreeLibrary函数;这样做会导致DLL被过早地取消映射。有关更多信息，请参见备注。
'    此选项不能与GET_MODULE_HANDLE_EX_FLAG_PIN一起使用?
'lpModuleName
'    加载模块的名称(.dll或.exe文件)或模块中的地址(如果dwFlags是GET_MODULE_HANDLE_EX_FLAG_FROM_ADDRESS)。
'    对于模块名，如果省略文件名扩展名，则添加默认的库扩展名.dll。文件名字符串可以包含一个尾随点字符(.)，以指示模块名称没有扩展名。字符串不需要指定路径。在指定路径时，请确保使用反斜杠()，而不是前斜杠(/)。将名称(独立于大小写)与当前映射到调用进程的地址空间的模块名称进行比较。
'    如果该参数为NULL，函数将返回用于创建调用进程的文件的句柄(.exe文件)。
'phModule
'    指定模块的句柄。如果函数失败，则此参数为NULL。
'    GetModuleHandleEx函数不检索使用LOAD_LIBRARY_AS_DATAFILE标志加载的模块的句柄。有关更多信息，请参见LoadLibraryEx。
'@返回值
'    如果函数成功，返回值为非零。
'    如果函数失败，返回值为零。要获得扩展的错误信息，请参阅GetLastError。
'@备注
'    返回的句柄不是全局的或可继承的?它不能被其他进程复制或使用?
'    如果lpModuleName不包含路径，并且有多个加载的模块具有相同的基本名称和扩展名，则无法预测将返回哪个模块句柄。要解决这个问题，可以指定路径、使用并行程序集，或者在lpModuleName参数中指定内存位置而不是DLL名称。
'    如果dwFlags包含GET_MODULE_HANDLE_EX_FLAG_UNCHANGED_REFCOUNT，则GetModuleHandleEx函数将向映射模块返回一个句柄，而不增加其引用计数。但是，如果将此句柄传递给FreeLibrary函数，则映射模块的引用计数将递减。因此，不要将get_module_handleex用GET_MODULE_HANDLE_EX_FLAG_UNCHANGED_REFCOUNT返回的句柄传递给FreeLibrary函数。这样做会导致DLL模块过早地未映射。
'    如果dwFlags包含GET_MODULE_HANDLE_EX_UNCHANGED_REFCOUNT，这个函数必须在多线程应用程序中谨慎使用。不能保证模块句柄在函数返回句柄和使用它之间保持有效。例如，一个线程检索一个模块句柄，但是在使用该句柄之前，第二个线程释放模块。如果系统加载另一个模块，它可以重用最近释放的模块句柄。因此，第一个线程将拥有一个与预期不同的模块句柄。
'    要编译使用此函数的应用程序，请将_WIN32_WINNT定义为0x0501或更高。有关更多信息，请参见使用Windows头文件。
'@Requirements
'    Minimum supported client    Windows XP [desktop apps only]
'    Minimum supported server    Windows Server 2003 [desktop apps only]
'    Target Platform Windows
'    Header  libloaderapi.h (include Windows.h)
'    Library kernel32.lib
'    dll kernel32.dll
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
'@原型
'    FARPROC GetProcAddress(
'      HMODULE hModule,
'      LPCSTR lpProcName
'    );
'@功能
'    从指定的动态链接库(DLL)检索导出函数或变量的地址。
'@参数
'hModule
'    DLL模块的句柄，包含函数或变量。LoadLibrary、LoadLibraryEx、LoadPackagedLibrary或GetModuleHandle函数返回这个句柄。
'    GetProcAddress函数不从使用LOAD_LIBRARY_AS_DATAFILE标志加载的模块中检索地址。有关更多信息，请参见LoadLibraryEx。
'lpProcName
'    函数或变量名，或函数的序号值。如果该参数是序数值，则它必须在低阶字中;高阶字必须为零。
'@返回值
'    如果函数成功，返回值是导出函数或变量的地址。
'    如果函数失败，返回值为NULL。要获取扩展的错误信息，请调用GetLastError。
'@备注
'    由lpProcName指向的函数名的拼写和大小写必须与源DLL模块定义(.def)文件的EXPORTS语句中的相同。函数的导出名称可能与在代码中调用这些函数时使用的名称不同。这种差异被SDK头文件中使用的宏所隐藏。有关更多信息，请参见函数原型的约定。
'    lpProcName参数可以通过在export语句中指定与函数关联的序号值来标识DLL函数。GetProcAddress验证指定的序号在范围1到.def文件中导出的最高序号值之间。然后，函数使用序号作为索引，从函数表中读取函数的地址。
'    如果.def文件没有按从1到N的顺序对函数编号(其中N是导出函数的数量)，那么GetProcAddress返回一个无效的非null地址时可能会发生错误，即使没有指定序号的函数。
'    如果函数可能不存在在DLL模块的例子中,如果函数只能在Windows Vista,但是应用程序可能运行在Windows XP-specify函数的名字而不是顺序值和设计您的应用程序在处理这个案子时,功能不可用,如下面代码片段所示。
'@Requirements
'    Minimum supported client    Windows XP [desktop apps | UWP apps]
'    Minimum supported server    Windows Server 2003 [desktop apps | UWP apps]
'    Target Platform Windows
'    Header  libloaderapi.h (include Windows.h)
'    Library kernel32.lib
'    dll kernel32.dll
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
    Value(5)        As Byte
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
Private Type TOKEN_USER
    User    As SID_AND_ATTRIBUTES
End Type
'@原型
'    typedef struct _TOKEN_USER {
'      SID_AND_ATTRIBUTES User;
'    } TOKEN_USER, *PTOKEN_USER;
'@功能
'    TOKEN_USER结构标识与访问令牌关联的用户?
'@成员
'User
'    指定表示与访问令牌关联的用户的SID_AND_ATTRIBUTES结构。目前还没有为用户安全标识符(SIDs)定义属性。
'@Requirements
'    Minimum supported client    Windows XP [desktop apps only]
'    Minimum supported server    Windows Server 2003 [desktop apps only]
'    Header  winnt.h (include Windows.h)
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
Private Enum EXTENDED_NAME_FORMAT
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
Private Declare Function IsWow64Process Lib "kernel32" (ByVal hProc As Long, bWow64Process As Long) As Long
'@原型
'    BOOL IsWow64Process(
'      HANDLE hProcess,
'      PBOOL Wow64Process
'    );
'@功能
'    确定指定的进程是在WOW64下运行，还是在x64处理器的Intel64下运行。
'@参数
'hProcess
'    进程的句柄。句柄必须具有PROCESS_QUERY_INFORMATION或PROCESS_QUERY_LIMITED_INFORMATION访问权限。有关更多信息，请参见流程安全和访问权限。
'    Windows Server 2003和Windows XP:句柄必须具有PROCESS_QUERY_INFORMATION访问权限。
'Wow64Process
'    如果进程在Intel64或x64处理器上以WOW64运行，则指向一个值的指针，该值被设置为TRUE。如果进程在32位窗口下运行，则将该值设置为FALSE。如果进程是运行在ARM上64位Windows 10下的32位应用程序，则将该值设置为FALSE。如果进程是运行在64位Windows下的64位应用程序，则该值也设置为FALSE。
'@返回值
'    如果函数成功，返回值为非零值。
'    如果函数失败，返回值为零。要获取扩展的错误信息，请调用GetLastError。
'@备注
'    应用程序应该使用IsWow64Process2而不是IsWow64Process来确定一个进程是否在WOW下运行。IsWow64Process2通过显式地返回给定进程的主机和客户机体系结构，消除了多个WOW环境固有的模糊性。应用程序可以使用这些信息可靠地识别在ARM64上模拟运行的情况。要编译使用此函数的应用程序，请将_WIN32_WINNT定义为0x0501或更高。有关更多信息，请参见使用Windows头文件。
'@Requirements
'    Minimum supported client    Windows Vista, Windows XP with SP2 [desktop apps | UWP apps]
'    Minimum supported server    Windows Server 2008, Windows Server 2003 with SP1 [desktop apps | UWP apps]
'    Target Platform Windows
'    Header  wow64apiset.h (include Windows Server 2003, Windows Vista, Windows 7, Windows Server 2008 Windows Server 2008 R2, Windows.h)
'    Library kernel32.lib
'    dll kernel32.dll

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
Private Declare Function lstrlen Lib "kernel32.dll" Alias "lstrlenW" (ByVal lpString As Long) As Long
'@原型
'    int lstrlenA(
'      LPCSTR lpString
'    );
'@功能
'    确定指定字符串的长度(不包括终止空字符)。
'@参数
'lpString LPCTSTR
'    要检查的以null结尾的字符串?
'@返回值
'    函数返回字符串的长度，以字符为单位。如果lpString为空，函数返回0。
'@Requirements
'    Minimum supported client    Windows 2000 Professional [desktop apps only]
'    Minimum supported server    Windows 2000 Server [desktop apps only]
'    Target Platform Windows
'    Header  winbase.h (include Windows.h)
'    Library kernel32.lib
'    dll kernel32.dll
Private Enum SID_NAME_USE
  SidTypeUser = 0
  '一个用户SID。
  SidTypeGroup = 1
  '组SID。
  SidTypeDomain = 2
  '域SID
  SidTypeAlias = 3
  '别名SID。
  SidTypeWellKnownGroup = 4
  '表示一个知名组的SID。
  SidTypeDeletedAccount = 5
  '是已删除帐户的SID。
  SidTypeInvalid = 6
  '一个无效的SID。
  SidTypeUnknown = 7
  '类型未知的SID。
  SidTypeComputer = 8
  '为计算机编写一个SID?
  SidTypeLabel = 9
  '强制的完整性标签SID。
  SidTypeLogonSession = 10
End Enum
'@原型
'    typedef enum _SID_NAME_USE {
'      SidTypeUser,
'      SidTypeGroup,
'      SidTypeDomain,
'      SidTypeAlias,
'      SidTypeWellKnownGroup,
'      SidTypeDeletedAccount,
'      SidTypeInvalid,
'      SidTypeUnknown,
'      SidTypeComputer,
'      SidTypeLabel,
'      SidTypeLogonSession
'    } SID_NAME_USE, *PSID_NAME_USE;
'@功能
'    SID_NAME_USE枚举包含指定安全标识符(SID)类型的值。
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
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
'@原型
'    HANDLE OpenProcess(
'      DWORD dwDesiredAccess,
'      BOOL  bInheritHandle,
'      DWORD dwProcessId
'    );
'@功能
'   打开现有的本地进程对象。
'@参数
'dwDesiredAccess
'    对流程对象的访问?此访问权限将根据流程的安全描述符进行检查?该参数可以是一个或多个进程访问权限?
'    如果调用方启用了SeDebugPrivilege特权，则不管安全描述符的内容如何，都将授予请求的访问权。
'bInheritHandle
'    如果该值为真，则由该流程创建的流程将继承该句柄。否则，进程不会继承此句柄。
'dwProcessId
'    要打开的本地进程的标识符?
'    如果指定的进程是系统进程(0x00000000)，则函数失败，最后一个错误代码是ERROR_INVALID_PARAMETER。如果指定的进程是空闲进程或CSRSS进程之一，则此函数失败，最后一个错误代码为ERROR_ACCESS_DENIED，因为它们的访问限制阻止用户级代码打开它们。
'    如果使用GetCurrentProcessId作为该函数的参数，请考虑使用GetCurrentProcess而不是OpenProcess，以提高性能。
'返回值
'    如果函数成功，返回值是指定进程的一个打开句柄。
'    如果函数失败，返回值为NULL。要获取扩展的错误信息，请调用GetLastError。
'@备注
'    要打开另一个本地进程的句柄并获得完全访问权限，必须启用SeDebugPrivilege特权。有关更多信息，请参见更改令牌中的特权。
'    OpenProcess函数返回的句柄可以用于任何需要进程句柄的函数，例如wait函数，只要请求了适当的访问权限。
'    当您处理完句柄时，请确保使用close handle函数关闭它。
'@Requirements
'    Minimum supported client    Windows XP [desktop apps | UWP apps]
'    Minimum supported server    Windows Server 2003 [desktop apps | UWP apps]
'    Target Platform Windows
'    Header  processthreadsapi.h (include Windows Server 2003, Windows Vista, Windows 7, Windows Server 2008 Windows Server 2008 R2, Windows.h)
'    Library kernel32.lib
'    dll kernel32.dll
Private Const Delete                     As Long = &H10000
'    需要删除对象。
Private Const READ_CONTROL               As Long = &H20000
'   需要读取对象的安全描述符中的信息，而不包括SACL中的信息。要读写SACL，必须请求ACCESS_SYSTEM_SECURITY访问权限。有关更多信息，请参见SACL访问权限。
'Private Const SYNCHRONIZE                As Long = &H100000
'   使用对象进行同步的权利。这使线程能够等待对象处于有信号状态。
Private Const WRITE_DAC                  As Long = &H40000
'    需要在对象的安全描述符中修改DACL。
Private Const WRITE_OWNER                As Long = &H80000
'    需要更改对象的安全描述符中的所有者。
'下表列出了特定于进程的访问权限。
Private Const PROCESS_ALL_ACCESS         As Long = &H0
'    进程对象的所有可能访问权限。Windows Server 2003和Windows XP: PROCESS_ALL_ACCESS标志的大小在Windows Server 2008和Windows Vista上增加了。如果为Windows Server 2008和Windows Vista编译的应用程序在Windows Server 2003或Windows XP上运行，PROCESS_ALL_ACCESS标志太大，指定该标志的函数失败，ERROR_ACCESS_DENIED。要避免此问题，请指定操作所需的最小访问权限集。如果必须使用PROCESS_ALL_ACCESS，那么将_WIN32_WINNT设置为应用程序所针对的最小操作系统(例如，#define _WIN32_WINNT _WIN32_WINNT_WINXP)。有关更多信息，请参见使用Windows头文件。
Private Const PROCESS_CREATE_PROCESS     As Long = &H80
'    需要创建进程
Private Const PROCESS_CREATE_THREAD      As Long = &H2
'    需要创建线程。
Private Const PROCESS_DUP_HANDLE         As Long = &H40
'    需要使用DuplicateHandle复制句柄。
Private Const PROCESS_QUERY_INFORMATION  As Long = &H400
'    需要检索关于进程的特定信息，比如它的令牌、退出代码和优先级类(请参阅OpenProcessToken)。
Private Const PROCESS_QUERY_LIMITED_INFORMATION  As Long = &H1000
'    需要检索关于进程的特定信息(参见GetExitCodeProcess、GetPriorityClass、IsProcessInJob、QueryFullProcessImageName)。具有PROCESS_QUERY_INFORMATION访问权限的句柄将自动授予PROCESS_QUERY_LIMITED_INFORMATION。Windows Server 2003和Windows XP:不支持此访问权限。
Private Const PROCESS_SET_INFORMATION    As Long = &H200
'    需要设置进程的某些信息，比如它的优先级类(参见SetPriorityClass)。
Private Const PROCESS_SET_QUOTA          As Long = &H100
'    需要使用SetProcessWorkingSetSize设置内存限制。
Private Const PROCESS_SUSPEND_RESUME     As Long = &H800
'    需要暂停或恢复进程。
Private Const PROCESS_TERMINATE          As Long = &H1
'    使用TerminateProcess来终止一个进程。
Private Const PROCESS_VM_OPERATION       As Long = &H8
'    需要在进程的地址空间上执行操作(请参阅VirtualProtectEx和WriteProcessMemory)。
Private Const PROCESS_VM_READ            As Long = &H10
'    需要使用ReadProcessMemory读取进程中的内存。
Private Const PROCESS_VM_WRITE           As Long = &H20
'    需要使用WriteProcessMemory在进程中写入内存。
Private Const SYNCHRONIZE                As Long = &H100000
'    需要使用wait函数等待进程终止。
'Private Const Delete                            As Long = &H10000
'Private Const READ_CONTROL                      As Long = &H20000
'Private Const WRITE_DAC                         As Long = &H40000
'Private Const WRITE_OWNER                       As Long = &H80000
'Private Const SYNCHRONIZE                       As Long = &H100000
Private Const STANDARD_RIGHTS_REQUIRED          As Long = &HF0000
Private Const STANDARD_RIGHTS_READ              As Long = READ_CONTROL
Private Const STANDARD_RIGHTS_WRITE             As Long = READ_CONTROL
Private Const STANDARD_RIGHTS_EXECUTE           As Long = READ_CONTROL
Private Const STANDARD_RIGHTS_ALL               As Long = &H1F0000
Private Const SPECIFIC_RIGHTS_ALL               As Long = &HFFFF
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
Private Declare Sub RtlMoveMemory Lib "kernel32" (Destination As Any, Source As Any, ByVal Length As Long)
'说明：将内存块从一个位置移动到另一个位置
'Destination:指向移动目的地起始地址的指针。
'Source:指向要移动的内存块起始地址的指针。
'Length:内存块的大小以字节为单位移动。
'注意事项：这个函数定义为RtlMoveMemory函数。它的实现是内联的。有关更多信息，请参见WinBase。h和Winnt.h。源和目标块可能会重叠。
'           第一个参数，目的地，必须足够大，以容纳长度字节的源;否则，可能会出现缓冲区溢出。这可能导致拒绝服务攻击，如果有访问违反，或者在最坏的情况下，允许攻击者向您的进程注入可执行代码。如果目的地是一个基于堆栈的缓冲区，则尤其如此。要注意，最后一个参数，长度，是将字节复制到目的地的数量，而不是目的地的大小。
Private Declare Function SetTokenInformation Lib "advapi32.dll" (ByVal TokenHandle As Long, ByVal TokenInformationClass As TOKEN_INFORMATION_CLASS, TokenInformation As Long, ByVal TokenInformationLength As Long) As Long
'@原型
'    BOOL SetTokenInformation(
'      HANDLE                  TokenHandle,
'      TOKEN_INFORMATION_CLASS TokenInformationClass,
'      LPVOID                  TokenInformation,
'      DWORD TokenInformationLength
'    );
'@功能
'    SetTokenInformation函数为指定的访问令牌设置各种类型的信息。此函数集替换现有信息的信息。调用进程必须具有适当的访问权限来设置信息。
'@参数
'TokenHandle
'    要为其设置信息的访问令牌句柄?
'TokenInformationClass
'    TOKEN_INFORMATION_CLASS枚举类型中的一个值，用于标识函数集的信息类型。TOKEN_INFORMATION_CLASS的有效值在TokenInformation参数中进行了描述。
'TokenInformation
'    指向包含访问令牌中信息集的缓冲区的指针?这个缓冲区的结构取决于TokenInformationClass参数指定的信息类型?
'TokenInformationLength
'    指定令牌信息指向的缓冲区的长度(以字节为单位)。
'@返回值
'    如果函数成功，则函数返回非零。
'    如果函数失败，它返回零。要获取扩展的错误信息，请调用GetLastError。
'@备注
'    要设置权限信息，应用程序可以调用AdjustTokenPrivileges函数。要设置令牌组，应用程序可以调用AdjustTokenGroups函数。
'    只有在创建访问令牌时才能设置令牌类型的信息?
'@Requirements
'    Minimum supported client    Windows XP [desktop apps | UWP apps]
'    Minimum supported server    Windows Server 2003 [desktop apps | UWP apps]
'    Target Platform Windows
'    Header  securitybaseapi.h (include Windows.h)
'    Library advapi32.lib
'    dll advapi32.dll
Private Declare Sub WTSFreeMemory Lib "wtsapi32" (ByVal pMemory As Long)
'@原型
'    void WTSFreeMemory(
'      IN PVOID pMemory
'    );
'@功能
'   释放由远程桌面服务函数分配的内存。
'@参数
'pMemory
'    指针指向内存释放?
'@返回值
'    此函数不返回值?
'@备注
'    几个远程桌面服务函数分配缓冲区来返回信息?使用WTSFreeMemory函数释放这些缓冲区?
'@Requirements
'    Minimum supported client    Windows Vista
'    Minimum supported server    Windows Server 2008
'    Target Platform Windows
'    Header wtsapi32.h
'    Library wtsapi32.lib
'    dll wtsapi32.dll
Private Declare Function WTSQuerySessionInformation Lib "wtsapi32" Alias "WTSQuerySessionInformationW" (ByVal hServer As Long, ByVal SessionID As Long, ByVal WTSInfoClass As Long, ppBuffer As Long, pBytesReturned As Long) As Long
'@原型
'    BOOL WTSQuerySessionInformationW(
'      IN HANDLE         hServer,
'      IN DWORD          SessionId,
'      IN WTS_INFO_CLASS WTSInfoClass,
'      LPWSTR            *ppBuffer,
'      DWORD             *pBytesReturned
'    );
'@功能
'    在指定的远程桌面会话主机(RD会话主机)服务器上检索指定会话的会话信息。它可用于查询本地和远程RD会话主机服务器上的会话信息。
'@参数
'hServer
'    RD会话主机服务器的句柄。指定WTSOpenServer函数打开的句柄，或者指定WTS_CURRENT_SERVER_HANDLE来指示运行应用程序的RD会话主机服务器。
'SessionId
'    远程桌面服务会话标识符。要指示调用应用程序正在运行的会话(或当前会话)，请指定WTS_CURRENT_SESSION。在获取本地服务器上的会话信息时，只指定WTS_CURRENT_SESSION。如果在查询远程服务器上的会话信息时指定了WTS_CURRENT_SESSION，那么返回的会话信息将是不一致的。不要使用返回的数据。
'    您可以使用wtsenumerates函数来检索指定RD会话主机服务器上的所有会话的标识符?
'    要为其他用户的会话查询信息，必须具有查询信息权限。有关更多信息，请参见远程桌面服务权限。要修改会话的权限，请使用远程桌面服务配置管理工具。
'WTSInfoClass
'    WTS_INFO_CLASS枚举的值，该值指示在调用WTSQuerySessionInformation函数时检索的会话信息的类型。
'ppBuffer
'    指向变量的指针，该变量接收指向请求信息的指针。数据的格式和内容取决于WTSInfoClass参数中指定的information类。要释放返回的缓冲区，请调用WTSFreeMemory函数。
'pBytesReturned
'    指向一个变量的指针，该变量接收ppBuffer中返回的数据的大小(以字节为单位)。
'@返回值
'    如果函数成功，返回值为非零值。
'    如果函数失败，返回值为零。要获取扩展的错误信息，请调用GetLastError。
'@备注
'    要在远程桌面服务运行时检索当前会话的会话ID，请调用WTSQuerySessionInformation，并为SessionId参数指定WTS_CURRENT_SESSION，为WTSInfoClass参数指定WTSSessionId。会话ID将在ppBuffer参数中返回。如果远程桌面服务没有运行，对WTSQuerySessionInformation的调用将失败。在这种情况下，您可以通过调用ProcessIdToSessionId函数来检索当前会话ID。
'    要确定应用程序是否在物理控制台上运行，必须为SessionId参数指定WTS_CURRENT_SESSION，并将WTSClientProtocolType指定为WTSInfoClass参数。如果ppBuffer为“0”，则会话被附加到物理控制台。
'@Requirements
'    Minimum supported client    Windows Vista
'    Minimum supported server    Windows Server 2008
'    Target Platform Windows
'    Header wtsapi32.h
'    Library wtsapi32.lib
'    dll wtsapi32.dll
Private Const WTS_CURRENT_SERVER_HANDLE         As Long = 0
Private Const WTS_CURRENT_SESSION               As Long = -1
Private Enum WTS_INFO_CLASS
      WTSInitialProgram = 0
      '一个以null结尾的字符串，包含用户登录时远程桌面服务运行的初始程序的名称。
      WTSApplicationName = 1
      '一个以null结尾的字符串，包含会话正在运行的应用程序的已发布名称。Windows Server 2008 R2、Windows 7、Windows Server 2008和Windows Vista:不支持此值
      WTSWorkingDirectory = 2
      '一个以null结尾的字符串，包含启动初始程序时使用的默认目录。
      WTSOEMId = 3
      '不使用此值。
      WTSSessionId = 4
      '包含会话标识符的ULONG值。
      WTSUserName = 5
      '包含与会话关联的用户名的以null结尾的字符串。
      WTSWinStationName = 6
      '一个以null结尾的字符串，包含远程桌面服务会话的名称。
        '注意，尽管指定了该类型的名称，但它并不返回窗口站名称。相反，它返回远程桌面服务会话的名称。每个远程桌面服务会话都与一个交互式窗口站相关联。由于交互式窗口站惟一支持的窗口站名称是“WinSta0”，因此每个会话都与自己的“WinSta0”窗口站相关联。有关更多信息，请参见窗口站。
      WTSDomainName = 7
      '一个以null结尾的字符串，包含登录用户所属的域的名称。
      WTSConnectState = 8
      '会话的当前连接状态。有关更多信息，请参见WTS_CONNECTSTATE_CLASS。
      WTSClientBuildNumber = 9
      '包含客户端构建号的ULONG值。
      WTSClientName = 10
      '包含客户端名称的以null结尾的字符串。
      WTSClientDirectory = 11
      '一个以null结尾的字符串，包含安装客户机的目录。
      WTSClientProductId = 12
      '一个特定于客户端的产品标识符。
      WTSClientHardwareId = 13
      '包含特定于客户机的硬件标识符的ULONG值。此选项保留为将来使用。WTSQuerySessionInformation总是返回0值。
      WTSClientAddress = 14
      '客户端的网络类型和网络地址。有关更多信息，请参见WTS_CLIENT_ADDRESS。IP地址从WTS_CLIENT_ADDRESS结构的address成员开始偏移两个字节?
      WTSClientDisplay = 15
      '有关客户端显示分辨率的信息。有关更多信息，请参见WTS_CLIENT_DISPLAY。
      WTSClientProtocolType = 16
      '指定会话协议类型信息的USHORT值。这是以下值之一。
        '0 控制台会话?
        '1 此值保留用于遗留用途?
        '2 RDP协议?
      WTSIdleTime = 17
      '此值返回FALSE。如果您调用GetLastError来获取扩展的错误信息，GetLastError将返回ERROR_NOT_SUPPORTED。Windows Server 2008和Windows Vista:不使用此值。
      WTSLogonTime = 18
      '此值返回FALSE。如果您调用GetLastError来获取扩展的错误信息，GetLastError将返回ERROR_NOT_SUPPORTED。Windows Server 2008和Windows Vista:不使用此值。
      WTSIncomingBytes = 19
      '此值返回FALSE。如果您调用GetLastError来获取扩展的错误信息，GetLastError将返回ERROR_NOT_SUPPORTED。Windows Server 2008和Windows Vista:不使用此值。
      WTSOutgoingBytes = 20
      '此值返回FALSE。如果您调用GetLastError来获取扩展的错误信息，GetLastError将返回ERROR_NOT_SUPPORTED。Windows Server 2008和Windows Vista:不使用此值。
      WTSIncomingFrames = 21
      '此值返回FALSE。如果您调用GetLastError来获取扩展的错误信息，GetLastError将返回ERROR_NOT_SUPPORTED。Windows Server 2008和Windows Vista:不使用此值。
      WTSOutgoingFrames = 22
      '此值返回FALSE。如果您调用GetLastError来获取扩展的错误信息，GetLastError将返回ERROR_NOT_SUPPORTED。Windows Server 2008和Windows Vista:不使用此值。
      WTSClientInfo = 24
      '有关远程桌面连接(RDC)客户机的信息。有关更多信息，请参见WTSCLIENT。
      WTSSessionInfo = 25
      '有关RD会话主机服务器上的客户机会话的信息。有关更多信息，请参见WTSINFO。
      WTSSessionInfoEx = 26
      '关于RD会话主机服务器上会话的扩展信息。有关更多信息，请参见WTSINFOEX。Windows Server 2008和Windows Vista:不支持此值。
      WTSConfigInfo = 27
      '包含有关RD会话主机服务器配置信息的WTSCONFIGINFO结构。Windows Server 2008和Windows Vista:不支持此值。
      WTSValidationInfo = 28
      '不支持此值。
      WTSSessionAddressV4 = 29
      '包含分配给会话的IPv4地址的WTS_SESSION_ADDRESS结构。如果会话没有虚拟IP地址，WTSQuerySessionInformation函数将返回ERROR_NOT_SUPPORTED。Windows Server 2008和Windows Vista:不支持此值。
      WTSIsRemoteSession = 30
      '确定当前会话是否为远程会话。WTSQuerySessionInformation函数返回值TRUE，表示当前会话是远程会话，返回值FALSE表示当前会话是本地会话。这个值只能用于本地机器，因此WTSQuerySessionInformation函数的hServer参数必须包含WTS_CURRENT_SERVER_HANDLE。Windows Server 2008和Windows Vista:不支持此值。
End Enum
'原型
'    typedef enum _WTS_INFO_CLASS {
'      WTSInitialProgram,
'      WTSApplicationName,
'      WTSWorkingDirectory,
'      WTSOEMId,
'      WTSSessionId,
'      WTSUserName,
'      WTSWinStationName,
'      WTSDomainName,
'      WTSConnectState,
'      WTSClientBuildNumber,
'      WTSClientName,
'      WTSClientDirectory,
'      WTSClientProductId,
'      WTSClientHardwareId,
'      WTSClientAddress,
'      WTSClientDisplay,
'      WTSClientProtocolType,
'      WTSIdleTime,
'      WTSLogonTime,
'      WTSIncomingBytes,
'      WTSOutgoingBytes,
'      WTSIncomingFrames,
'      WTSOutgoingFrames,
'      WTSClientInfo,
'      WTSSessionInfo,
'      WTSSessionInfoEx,
'      WTSConfigInfo,
'      WTSValidationInfo,
'      WTSSessionAddressV4,
'      WTSIsRemoteSession
'    } WTS_INFO_CLASS;
'@功能
'   包含一些值，这些值指示在调用WTSQuerySessionInformation函数时检索的会话信息的类型。
'@Requirements
'    Minimum supported client    Windows Vista
'    Minimum supported server    Windows Server 2008
'    Header wtsapi32.h
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
Private Const SE_DEBUG_NAME                     As String = "SeDebugPrivilege"
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


'是否已经判断的枚举
Private Enum CacheType
    CT_Is64BitOs = 0
    CT_StartExePath = 1
    CT_SessionID = 2
    CT_IsDesinMode = 3
    CT_ProcessUser = 4
    CT_ProcessUserSID = 5
    CT_SessionUser = 6
    CT_SeesionDomain = 7
    CT_SessionUserSID = 8
End Enum
'---------------------------------------------------------------------------
'                1、常规变量
'---------------------------------------------------------------------------
Private mobjFSO                 As New FileSystemObject

Private marrCahceInfo(2)        As Byte                 '已经缓存的信息标记，按位标记
'缓存环境结果
Private mblnIs64BitOs           As Boolean              '是否是64位操作系统
Private mstrStartExePath        As String               '程序启动路径
Private mlngCurrentSessionID    As Long                 '当前进程所在的会话ID
Private mblnIsDesinMode         As Boolean              '当前是否是源码模式
Private mstrProcessUser         As String               '当前进程的操作系统用户
Private mstrProcessDomian       As String               '当前进程的操作系统的用户域
Private mstrProcessUserFull     As String               '当前进程的操作用户的全名，域\用户名
Private mstrProcessUserSID      As String               '当前进程用户的SID
Private mstrSessonUser          As String               '当前会话的用户名
Private mstrSessionDomain       As String               '当前会话的用户域
Private mstrSessionUserSID      As String               '当前会话用户的SID
Private mstrComputerName        As String               '当前电脑名
'---------------------------------------------------------------------------
'                2、属性变量与定义
'---------------------------------------------------------------------------

'---------------------------------------------------------------------------
'                3、公共方法
'---------------------------------------------------------------------------

'@方法    Is64bitOS
'   是否64位操作系统
'@返回值  Boolean
'
'@参数:
'@备注
'
Public Function Is64bitOS() As Boolean
    Dim Handle  As Long
    Dim blnFunc As Long
        
    On Error GoTo ErrH
    Call Logger.PushMethod("ZLHelperMain.clsEnvironment.Is64bitOS")
    If IsInfoCached(CT_Is64BitOs) Then
        Is64bitOS = mblnIs64BitOs
    Else
        blnFunc = 0
        Handle = GetProcAddress(GetModuleHandle("kernel32"), "IsWow64Process")
        If Handle > 0 Then
            IsWow64Process GetCurrentProcess(), blnFunc
        End If
        Is64bitOS = blnFunc <> 0
        
        mblnIs64BitOs = Is64bitOS
        SetInfoCached CT_Is64BitOs
    End If
    Call Logger.PopMethod("ZLHelperMain.clsEnvironment.Is64bitOS", Is64bitOS)
    Exit Function
ErrH:
    If Logger.ErrCenter("ZLHelperMain.clsEnvironment.Is64bitOS") = 1 Then
        Resume
    End If

    Call Logger.PopMethod("ZLHelperMain.clsEnvironment.Is64bitOS")
End Function

'@方法    StartExePath
'   获取启动的EXE路径
'@返回值  String
'
'@参数:
'Param 方向
'
'@备注
'
Public Function StartExePath(Optional ByVal lngProcessID As Long) As String
    Dim strFile         As String
    Dim lngRet          As Long
    Dim hProcess        As Long
    Dim hModule         As Long
    
    On Error GoTo ErrH
    Call Logger.PushMethod("ZLHelperMain.clsEnvironment.StartExePath")
    If lngProcessID = 0 Then
        If IsInfoCached(CT_StartExePath) Then
            StartExePath = mstrStartExePath
        Else
            strFile = String(MAX_PATH, Chr(0))
            
            lngRet = GetModuleFileName(0, strFile, MAX_PATH)
            If Err.LastDllError <> 0 Then
                Logger.Error "GetModuleFileName", "Error", GetLastDllErr(Err.LastDllError)
            Else
                StartExePath = TruncZero(strFile)
                
                mstrStartExePath = StartExePath
                SetInfoCached CT_StartExePath
            End If
        End If
    Else
        hProcess = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ, 0, lngProcessID)
        If hProcess = 0 Then
            Logger.Error "OpenProcess", "Error", GetLastDllErr(Err.LastDllError)
        Else
            If EnumProcessModules(hProcess, hModule, 4&, 0&) = 0 Then
                Logger.Error "EnumProcessModules", "Error", GetLastDllErr(Err.LastDllError)
            Else
                strFile = String(MAX_PATH, Chr(0))
                lngRet = GetModuleFileNameEx(hProcess, hModule, strFile, MAX_PATH)
                If Err.LastDllError <> 0 Then
                    Logger.Error "GetModuleFileName", "Error", GetLastDllErr(Err.LastDllError)
                Else
                    StartExePath = TruncZero(strFile)
                End If
            End If
            If CloseHandle(hProcess) = 0 Then
                Logger.Error "CloseHandle", "Error", GetLastDllErr(Err.LastDllError)
            Else
                hProcess = 0
            End If
        End If
    End If
    Call Logger.PopMethod("ZLHelperMain.clsEnvironment.StartExePath", StartExePath)
    Exit Function
ErrH:
    If Logger.ErrCenter("ZLHelperMain.clsEnvironment.StartExePath") = 1 Then
        Resume
    End If
    If hProcess <> 0 Then
        If CloseHandle(hProcess) = 0 Then
            Logger.Error "CloseHandle", "Error", GetLastDllErr(Err.LastDllError)
        End If
    End If
    Call Logger.PopMethod("ZLHelperMain.clsEnvironment.StartExePath")
End Function

'@方法    SessionID
'   获取当前会话ID
'@返回值  Long
'
'@参数:
'lngProcessID Long In (Optional,Defualt=0)
'   0-当前进程所在的会话，<>0-指定进程的会话
'@备注
'
Public Function SessionID(Optional ByVal lngProcessID As Long) As Long
    Dim lngSessionID            As Long
    Dim lngTmpProcessID         As Long
    On Error GoTo ErrH
    Call Logger.PushMethod("ZLHelperMain.clsEnvironment.SessionID", lngProcessID)
    If lngProcessID = 0 Then
        If IsInfoCached(CT_SessionID) Then
            SessionID = mlngCurrentSessionID
        Else
            lngTmpProcessID = GetCurrentProcessId
        End If
    Else
        lngTmpProcessID = lngProcessID
    End If
    If lngTmpProcessID <> 0 Then
        SessionID = -1
        If ProcessIdToSessionId(lngTmpProcessID, lngSessionID) = 0 Then
            Logger.Error "ProcessIdToSessionId", "Error", GetLastDllErr(Err.LastDllError)
        Else
            SessionID = lngSessionID
            If lngProcessID = 0 Then
                mlngCurrentSessionID = SessionID
                SetInfoCached CT_SessionID
            End If
        End If
    End If
    Call Logger.PopMethod("ZLHelperMain.clsEnvironment.SessionID", SessionID)
    Exit Function
ErrH:
    If Logger.ErrCenter("ZLHelperMain.clsEnvironment.SessionID") = 1 Then
        Resume
    End If
    
    Call Logger.PopMethod("ZLHelperMain.clsEnvironment.SessionID")
End Function

'@方法    IsDesinMode
'   当前是否是源码环境,该判断可以准确判断，不管该方法是否在源码中
'@返回值  Boolean
Public Function IsDesinMode() As Boolean
    Dim strStartPath        As String
    
    On Error GoTo ErrH
    Call Logger.PushMethod("ZLHelperMain.clsEnvironment.IsDesinMode")
    If IsInfoCached(CT_IsDesinMode) Then
        IsDesinMode = mblnIsDesinMode
    Else
        strStartPath = Me.StartExePath()
        If UCase(Trim(mobjFSO.GetFileName(strStartPath))) = "VB6.EXE" Then
            IsDesinMode = True
        End If
        
        mblnIsDesinMode = IsDesinMode
        SetInfoCached CT_IsDesinMode
    End If
    Call Logger.PopMethod("ZLHelperMain.clsEnvironment.IsDesinMode", IsDesinMode)
    Exit Function
ErrH:
    If Logger.ErrCenter("ZLHelperMain.clsEnvironment.IsDesinMode") = 1 Then
        Resume
    End If

    Call Logger.PopMethod("ZLHelperMain.clsEnvironment.IsDesinMode")
End Function
'@方法    ProcessUser
'   获取当前进程的操作系统的用户名
'@返回值  String
'
'@参数:
'lngProcessID Long In (Optional,Defualt=0)
'   0,当前进程用户。<>0指定进程的用户
'@备注
'
Public Function ProcessUser(Optional ByVal lngProcessID As Long) As String
    Dim strUserName As String, strDomain    As String, strFullUser      As String
    On Error GoTo ErrH
    Call Logger.PushMethod("ZLHelperMain.clsEnvironment.ProcessUser", lngProcessID)
    
    If lngProcessID = 0 Then
        If IsInfoCached(CT_ProcessUser) Then
            ProcessUser = mstrProcessUser
        Else
            If ProcessUserInfo_Sub(lngProcessID, strUserName, strDomain, strFullUser) Then
                ProcessUser = strUserName
            End If
        End If
    Else
        If ProcessUserInfo_Sub(lngProcessID, strUserName, strDomain, strFullUser) Then
            ProcessUser = strUserName
        End If
    End If
    Call Logger.PopMethod("ZLHelperMain.clsEnvironment.ProcessUser", ProcessUser)
    Exit Function
ErrH:
    If Logger.ErrCenter("ZLHelperMain.clsEnvironment.ProcessUser") = 1 Then
        Resume
    End If

    Call Logger.PopMethod("ZLHelperMain.clsEnvironment.ProcessUser")
End Function

'@方法    ProcessDomain
'   获取当前进程的操作系统的域
'@返回值  String
'
'@参数:
'lngProcessID Long In (Optional,Defualt=0)
'   0,当前进程用户。<>0指定进程的用户
'@备注
'
Public Function ProcessDomain(Optional ByVal lngProcessID As Long) As String
    Dim strUserName As String, strDomain    As String, strFullUser      As String
    On Error GoTo ErrH
    Call Logger.PushMethod("ZLHelperMain.clsEnvironment.ProcessDomain", lngProcessID)
    
    If lngProcessID = 0 Then
        If IsInfoCached(CT_ProcessUser) Then
            ProcessDomain = mstrProcessDomian
        Else
            If ProcessUserInfo_Sub(lngProcessID, strUserName, strDomain, strFullUser) Then
                ProcessDomain = strDomain
            End If
        End If
    Else
        If ProcessUserInfo_Sub(lngProcessID, strUserName, strDomain, strFullUser) Then
            ProcessDomain = strDomain
        End If
    End If
    Call Logger.PopMethod("ZLHelperMain.clsEnvironment.ProcessDomain", ProcessDomain)
    Exit Function
ErrH:
    If Logger.ErrCenter("ZLHelperMain.clsEnvironment.ProcessDomain") = 1 Then
        Resume
    End If

    Call Logger.PopMethod("ZLHelperMain.clsEnvironment.ProcessDomain")
End Function

'@方法    ProcessUserFullName
'   获取当前进程的操作系统的用户全名，包含域
'@返回值  String
'
'@参数:
'lngProcessID Long In (Optional,Defualt=0)
'   0,当前进程用户全名。<>0指定进程的用户全名
'@备注
'
Public Function ProcessUserFullName(Optional ByVal lngProcessID As Long) As String
    Dim strUserName As String, strDomain    As String, strFullUser      As String
    On Error GoTo ErrH
    Call Logger.PushMethod("ZLHelperMain.clsEnvironment.ProcessUserFullName", lngProcessID)
    
    If lngProcessID = 0 Then
        If IsInfoCached(CT_ProcessUser) Then
            ProcessUserFullName = mstrProcessUserFull
        Else
            If ProcessUserInfo_Sub(lngProcessID, strUserName, strDomain, strFullUser) Then
                ProcessUserFullName = strFullUser
            End If
        End If
    Else
        If ProcessUserInfo_Sub(lngProcessID, strUserName, strDomain, strFullUser) Then
            ProcessUserFullName = strFullUser
        End If
    End If
    Call Logger.PopMethod("ZLHelperMain.clsEnvironment.ProcessUserFullName", ProcessUserFullName)
    Exit Function
ErrH:
    If Logger.ErrCenter("ZLHelperMain.clsEnvironment.ProcessUserFullName") = 1 Then
        Resume
    End If

    Call Logger.PopMethod("ZLHelperMain.clsEnvironment.ProcessUserFullName")
End Function
'@方法    ProcessUserSID
'   获取进程用户的SID
'@返回值  String
'
'@参数:
'lngProcessID Long In (Optional,Defualt=0)
'   0,当前进程。<>0指定进程
'@备注
'
Public Function ProcessUserSID(Optional ByVal lngProcessID As Long) As String
    If lngProcessID = 0 Then
        If Not IsInfoCached(CT_ProcessUserSID) Then
            mstrProcessUserSID = GetUserSID(ProcessUserFullName(lngProcessID))
            SetInfoCached CT_ProcessUserSID
        End If
        ProcessUserSID = mstrProcessUserSID
    Else
        ProcessUserSID = GetUserSID(ProcessUserFullName(lngProcessID))
    End If
End Function
'@方法    SessionUser
'   获取会话的用户
'@返回值  String
'
'@参数:
'lngSeesionID Long In (Optional,Defualt=-1)
'   -1,当前会话的用户。>-1,指定会话的用户
'@备注
'
Public Function SessionUser(Optional ByVal lngSeesionID As Long = -1) As String
    Dim pBuffer         As Long
    Dim dwBufferLen     As Long
    Dim arrBytRet()     As Byte
    Dim strUserName     As String
    Dim blnNew          As Boolean
    
    On Error GoTo ErrH
    Call Logger.PushMethod("ZLHelperMain.clsEnvironment.SessionUser", lngSeesionID)
    If lngSeesionID = -1 Then
        If IsInfoCached(CT_SessionUser) Then
            SessionUser = mstrSessonUser
        Else
            blnNew = True
        End If
    Else
        blnNew = True
    End If
    If blnNew Then
        If WTSQuerySessionInformation(WTS_CURRENT_SERVER_HANDLE, lngSeesionID, WTSUserName, pBuffer, dwBufferLen) <> 0 Then
            Logger.DebugEx "WTSQuerySessionInformation", "pBuffer", pBuffer, "dwBufferLen", dwBufferLen
            If dwBufferLen <> 0 Then
                ReDim Preserve arrBytRet(dwBufferLen - 1)
                RtlMoveMemory ByVal VarPtr(arrBytRet(0)), ByVal pBuffer, dwBufferLen
                strUserName = arrBytRet
                strUserName = TruncZero(strUserName)
                
                SessionUser = strUserName
                
                If lngSeesionID = -1 Then
                    mstrSessonUser = SessionUser
                    SetInfoCached CT_SessionUser
                End If
            End If
            WTSFreeMemory pBuffer
        Else
            Logger.Error "WTSQuerySessionInformation", "Error", GetLastDllErr(Err.LastDllError)
        End If
    End If
    Call Logger.PopMethod("ZLHelperMain.clsEnvironment.SessionUser", SessionUser)
    Exit Function
ErrH:
    If Logger.ErrCenter("ZLHelperMain.clsEnvironment.SessionUser") = 1 Then
        Resume
    End If

    Call Logger.PopMethod("ZLHelperMain.clsEnvironment.SessionUser")
End Function

'@方法    SessionDomain
'   获取会话的域
'@返回值  String
'
'@参数:
'lngSeesionID Long In (Optional,Defualt=-1)
'   -1,当前会话的域。>-1,指定会话的域
'@备注
'
Public Function SessionDomain(Optional ByVal lngSeesionID As Long = -1) As String
    Dim pBuffer         As Long
    Dim dwBufferLen     As Long
    Dim arrBytRet()     As Byte
    Dim strDomain       As String
    Dim blnNew          As Boolean
    
    On Error GoTo ErrH
    Call Logger.PushMethod("ZLHelperMain.clsEnvironment.SessionDomain", lngSeesionID)
    If lngSeesionID = -1 Then
        If IsInfoCached(CT_SeesionDomain) Then
            SessionDomain = mstrSessionDomain
        Else
            blnNew = True
        End If
    Else
        blnNew = True
    End If
    If blnNew Then
        If WTSQuerySessionInformation(WTS_CURRENT_SERVER_HANDLE, lngSeesionID, WTSDomainName, pBuffer, dwBufferLen) <> 0 Then
            Logger.DebugEx "WTSQuerySessionInformation", "pBuffer", pBuffer, "dwBufferLen", dwBufferLen
            If dwBufferLen <> 0 Then
                ReDim Preserve arrBytRet(dwBufferLen - 1)
                RtlMoveMemory ByVal VarPtr(arrBytRet(0)), ByVal pBuffer, dwBufferLen
                strDomain = arrBytRet
                strDomain = TruncZero(strDomain)
                
                SessionDomain = strDomain
                If lngSeesionID = -1 Then
                    mstrSessionDomain = SessionDomain
                    SetInfoCached CT_SeesionDomain
                End If
            End If
            WTSFreeMemory pBuffer
        Else
            Logger.Error "WTSQuerySessionInformation", "Error", GetLastDllErr(Err.LastDllError)
        End If
    End If
    Call Logger.PopMethod("ZLHelperMain.clsEnvironment.SessionDomain", SessionDomain)
    Exit Function
ErrH:
    If Logger.ErrCenter("ZLHelperMain.clsEnvironment.SessionDomain") = 1 Then
        Resume
    End If

    Call Logger.PopMethod("ZLHelperMain.clsEnvironment.SessionDomain")
End Function
'@方法    SessionUserFullName
'   获取会话的用户全名。包含域
'@返回值  String
'
'@参数:
'lngSeesionID Long In (Optional,Defualt=-1)
'   -1,当前会话的全名。>-1,指定会话的全名
'@备注
'
Public Function SessionUserFullName(Optional ByVal lngSeesionID As Long = -1) As String
    Dim strTmp        As String
    
    If lngSeesionID = -1 Then
        If Not IsInfoCached(CT_SessionUser) Then
            strTmp = SessionUser(lngSeesionID)
        End If
        If Not IsInfoCached(CT_SeesionDomain) Then
            strTmp = SessionDomain(lngSeesionID)
        End If
        If IsInfoCached(CT_SessionUser) And IsInfoCached(CT_SeesionDomain) Then
            SessionUserFullName = mstrSessionDomain & "\" & mstrSessonUser
        End If
    Else
        strTmp = SessionDomain(lngSeesionID)
        If LenB(strTmp) <> 0 Then
            SessionUserFullName = strTmp & "\" & SessionUser(lngSeesionID)
        End If
    End If
End Function

'@方法    SessionUserSID
'   获取当会话用户的SID
'@返回值  String
'
'@参数:
'lngSeesionID Long In (Optional,Defualt=-1)
'   -1,当前会话。>-1,指定会话
'@备注
'
Public Function SessionUserSID(Optional ByVal lngSeesionID As Long = -1) As String
    If lngSeesionID = -1 Then
        If Not IsInfoCached(CT_SessionUserSID) Then
            mstrSessionUserSID = GetUserSID(SessionUserFullName(lngSeesionID))
            SetInfoCached CT_SessionUserSID
        End If
        SessionUserSID = mstrSessionUserSID
    Else
        SessionUserSID = GetUserSID(SessionUserFullName(lngSeesionID))
    End If
End Function

'@方法    IsProcessRunAs
'   当前进程是以管理员权限运行。管理员权限和管理员不是同一个概念，Visita之上，标准管理员运行是权限不是管理员，必须RUnAS
'@返回值  Boolean
'
'@参数:
'lngProcessID Long In (Optional,Defualt=0)
'   0-当前进程，<>0-指定进程
'@备注
'   在判断其他进程前，先保证当前进程IsProcessRunAs=True,否则会没有权限，返回False
Public Function IsProcessRunAs(Optional ByVal lngProcessID As Long) As Boolean
    Dim ntAuthority         As SID_IDENTIFIER_AUTHORITY
    Dim psidAdmin           As Long
    Dim lngIsMember         As Long
    Dim lngProcess          As Long, hToken     As Long, hNewToken      As Long, blnRevertToSelf        As Boolean
    
    On Error GoTo ErrH
    Call Logger.PushMethod("ZLHelperMain.clsEnvironment.IsProcessRunAs")
    ntAuthority.Value(5) = security_nt_authority
    If AllocateAndInitializeSid(ntAuthority, 2, SECURITY_BUILTIN_DOMAIN_RID, DOMAIN_ALIAS_RID_ADMINS, 0, 0, 0, 0, 0, 0, psidAdmin) <> 0 Then
        If lngProcessID <> 0 Then
            If Not Process.EnablePrivilege(SE_DEBUG_NAME) Then
                Logger.DebugEx "EnablePrivilege", "lngProcessID", lngProcessID, "结果", False
            End If
            lngProcess = OpenProcess(PROCESS_QUERY_INFORMATION, 0, lngProcessID)
            If lngProcess = 0 Then
                Logger.Error "OpenProcess", "Error", GetLastDllErr(Err.LastDllError)
            Else
                If OpenProcessToken(lngProcess, TOKEN_QUERY Or TOKEN_DUPLICATE Or TOKEN_ADJUST_PRIVILEGES Or TOKEN_IMPERSONATE, hToken) = 0 Then
                    Logger.Error "OpenProcessToken", "Error", GetLastDllErr(Err.LastDllError)
                Else
                    If DuplicateTokenEx(hToken, TOKEN_ALL_ACCESS, ByVal 0, ByVal SecurityImpersonation, ByVal TokenImpersonation, hNewToken) = 0 Then
                        Logger.Error "DuplicateTokenEx", "Error", GetLastDllErr(Err.LastDllError)
                    Else
'                        If Not Process.EnablePrivilegeToken(hNewToken, SE_DEBUG_NAME) Then
'                            Logger.DebugEx "EnablePrivilege", "lngProcessID", lngProcessID, "结果", False
'                        End If
                        If ImpersonateLoggedOnUser(hNewToken) = 0 Then
                            Logger.Error "ImpersonateLoggedOnUser", "Error", GetLastDllErr(Err.LastDllError)
                        Else
                            blnRevertToSelf = True
                            If CheckTokenMembership(hNewToken, psidAdmin, lngIsMember) = 0 Then
                                Logger.Warn "CheckTokenMembership", "Error", GetLastDllErr(Err.LastDllError)
                            End If
                            If RevertToSelf() = 0 Then
                                Logger.Error "RevertToSelf", "Error", GetLastDllErr(Err.LastDllError)
                            Else
                                blnRevertToSelf = False
                            End If
                            If CloseHandle(hNewToken) = 0 Then
                                Logger.Error "CloseHandle(NewToken)", "Error", GetLastDllErr(Err.LastDllError)
                            Else
                                hNewToken = 0
                            End If
                            If CloseHandle(hToken) = 0 Then
                                Logger.Error "CloseHandle(Token)", "Error", GetLastDllErr(Err.LastDllError)
                            Else
                                hToken = 0
                            End If
                        End If
                    End If
                End If
                If CloseHandle(lngProcess) = 0 Then
                    Logger.Error "CloseHandle(Process)", "Error", GetLastDllErr(Err.LastDllError)
                Else
                    lngProcess = 0
                End If
            End If
        Else
            If CheckTokenMembership(0, psidAdmin, lngIsMember) = 0 Then
                Logger.Warn "CheckTokenMembership", "Error", GetLastDllErr(Err.LastDllError)
            End If
        End If
        If FreeSid(psidAdmin) <> 0 Then
            Logger.Error "FreeSid", "Error", GetLastDllErr(Err.LastDllError)
        Else
            psidAdmin = 0
            IsProcessRunAs = lngIsMember <> 0
        End If
    Else
        Logger.Error "AllocateAndInitializeSid", "Error", GetLastDllErr(Err.LastDllError)
    End If
    Call Logger.PopMethod("ZLHelperMain.clsEnvironment.IsProcessRunAs", IsProcessRunAs)
    Exit Function
ErrH:
    If Logger.ErrCenter("ZLHelperMain.clsEnvironment.IsProcessRunAs") = 1 Then
        Resume
    End If
    
    If blnRevertToSelf Then
        If RevertToSelf() = 0 Then
            Logger.Error "RevertToSelf", "Error", GetLastDllErr(Err.LastDllError)
        End If
    End If
    If hToken <> 0 Then
        If CloseHandle(hToken) = 0 Then
            Logger.Error "CloseHandle(Token)", "Error", GetLastDllErr(Err.LastDllError)
        End If
    End If
    If hNewToken <> 0 Then
        If CloseHandle(hNewToken) = 0 Then
            Logger.Error "CloseHandle(NewToken)", "Error", GetLastDllErr(Err.LastDllError)
        End If
    End If
    If lngProcess <> 0 Then
        If CloseHandle(lngProcess) = 0 Then
            Logger.Error "CloseHandle(Process)", "Error", GetLastDllErr(Err.LastDllError)
        End If
    End If
    If psidAdmin <> 0 Then
        If FreeSid(psidAdmin) <> 0 Then
            Logger.Error "FreeSid", "Error", GetLastDllErr(Err.LastDllError)
        End If
    End If
    Call Logger.PopMethod("ZLHelperMain.clsEnvironment.IsProcessRunAs")
End Function
'@方法    IsWindows2000OrGreater
'   是否是Window2000之后版本
'@返回值  Boolean
'
'@参数:
'@备注
'
Private Function IsWindows2000OrGreater() As Boolean
    Dim osInfo          As OSVERSIONINFOEX
    On Error GoTo ErrH
    Call Logger.PushMethod("ZLHelperMain.clsEnvironment.IsWindows2000OrGreater")
    osInfo.dwOSVersionInfoSize = LenB(osInfo)
    osInfo.szCSDVersion = Space$(128)
    If GetVersionExA(osInfo) = 0 Then
        Logger.Error "GetVersionExA", "Error", GetLastDllErr(Err.LastDllError)
    Else
        IsWindows2000OrGreater = osInfo.dwMajorVersion >= 5
    End If
    Call Logger.PopMethod("ZLHelperMain.clsEnvironment.IsWindows2000OrGreater", IsWindows2000OrGreater)
    Exit Function
ErrH:
    If Logger.ErrCenter("ZLHelperMain.clsEnvironment.IsWindows2000OrGreater") = 1 Then
        Resume
    End If
    
    Call Logger.PopMethod("ZLHelperMain.clsEnvironment.IsWindows2000OrGreater")
End Function
'@方法    IsWindowsXPrGreater
'   是否是WindowXP之后版本
'@返回值  Boolean
'
'@参数:
'@备注
'
Private Function IsWindowsXPrGreater() As Boolean
    Dim osInfo          As OSVERSIONINFOEX
    On Error GoTo ErrH
    Call Logger.PushMethod("ZLHelperMain.clsEnvironment.IsWindowsXPrGreater")
    osInfo.dwOSVersionInfoSize = LenB(osInfo)
    osInfo.szCSDVersion = Space$(128)
    If GetVersionExA(osInfo) = 0 Then
        Logger.Error "GetVersionExA", "Error", GetLastDllErr(Err.LastDllError)
    Else
        IsWindowsXPrGreater = osInfo.dwMajorVersion >= 5 And osInfo.dwMinorVersion > 0
    End If
    Call Logger.PopMethod("ZLHelperMain.clsEnvironment.IsWindowsXPrGreater", IsWindowsXPrGreater)
    Exit Function
ErrH:
    If Logger.ErrCenter("ZLHelperMain.clsEnvironment.IsWindowsXPrGreater") = 1 Then
        Resume
    End If
    
    Call Logger.PopMethod("ZLHelperMain.clsEnvironment.IsWindowsXPrGreater")
End Function
'@方法    IP
'   返回IP地址,优先公网IP,然后局域网IP，首选IP优先
'@返回值  String
'
'@参数:

'@备注
'
Public Function IP() As String
    Dim lngSize         As Long, lngReturn              As Long
    Dim arrlngData()    As Long
    Dim arrIpRows()     As MIB_IPADDRROW_XP
    Dim i               As Long
    Dim blnXP           As Boolean
    Dim strPubIp        As String, strPubPriIP          As String
    Dim strPrivateIp    As String, strPrivatePriIp      As String
    Dim strTmp          As String, strIp                As String, strFMTIp    As String
    
    
    On Error GoTo ErrH
    Call Logger.PushMethod("ZLHelperMain.clsEnvironment.IP")
    blnXP = IsWindowsXPrGreater
    lngReturn = GetIpAddrTable(ByVal 0, lngSize, 1)
    If lngReturn <> ERROR_INSUFFICIENT_BUFFER Then
        Logger.Error "GetIpAddrTable", "Error", GetLastDllErr(Err.LastDllError)
    ElseIf lngSize <= 0 Then
        Logger.DebugEx "GetIpAddrTable", "Error", GetLastDllErr(Err.LastDllError), "dwSize", lngSize
    Else
        Logger.DebugEx "GetIpAddrTable", "Error", GetLastDllErr(Err.LastDllError), "dwSize", lngSize
        ReDim arrlngData(lngSize \ 4 - 1)
        lngReturn = GetIpAddrTable(ByVal VarPtr(arrlngData(0)), lngSize, 0)
        If lngReturn <> 0 Then
            Logger.Error "GetIpAddrTable", "Error", GetLastDllErr(Err.LastDllError)
        Else
            Logger.DebugEx "GetIpAddrTable", "dwNumEntries", arrlngData(0)
            If arrlngData(0) > 0 Then
                ReDim arrIpRows(arrlngData(0) - 1)
                RtlMoveMemory ByVal VarPtr(arrIpRows(0)), ByVal VarPtr(arrlngData(1)), LenB(arrIpRows(0)) * (UBound(arrIpRows) + 1)
                For i = LBound(arrIpRows) To UBound(arrIpRows)
                    strTmp = IP2String(arrIpRows(i).dwAddr)
                    strIp = ""
                    Logger.DebugEx "IpAddrTable", "IP", strTmp, "Mask", IP2String(arrIpRows(i).dwMask), "BroadcastIp", IP2String(arrIpRows(i).dwBCastAddr), "NetInterfaceIndex", arrIpRows(i).dwIndex, "IPType", arrIpRows(i).wType
                    '0.0.0.0-本机地址，255.255.255.255-当前子网的广播地址，这两个地址先排除
                    '不能以127开头，该类地址中127．0．0．1到127．255．255．255用于回路测试。127.0.0.1可以代表本机IP地址
                    'A类地址网络号范围：0.0.0.0---127.0.0.0 ,地址范围0.0.0.0到127.255.255.255,私有10.0.0.0---10.255.255.255
                    'B类地址网络号范围：128.0.0.0---191.255.0.0,地址范围128.0.0.0到191.255.255.255。私有：172.16.0.0---172.31.255.255
                    '   169.254.X.X是保留地址。191.255.255.255是广播地址
                    'C类地址网络号范围：192.0.0.0---223.255.255.0。地址范围 192.0.0.0到223.255.255.255,私有192.168.0.0---192.168.255.255
                    'D类地址。不分网络地址和主机地址，它的第1个字节的前四位固定为1110。地址范围：224.0.0.0---239.255.255.255
                    'E类地址。分网络地址和主机地址，它的第1个字节的前五位固定为11110，地址范围：240.0.0.0---255.255.255.254
                    '组播地址：224.0.0.0到239.255.255.255，224.0.0.1特指所有主机， 224.0.0.2特指所有路由器
                    If Not strTmp Like "127.*" And Not strTmp Like "169.254.*" And strTmp <> "0.0.0.0" And strTmp <> "255.255.255.255" Then
                        If blnXP Then
                            If (arrIpRows(i).wType And MIB_IPADDR_DISCONNECTED) <> MIB_IPADDR_DISCONNECTED And (arrIpRows(i).wType And MIB_IPADDR_DELETED) <> MIB_IPADDR_DELETED Then
                                strFMTIp = FormatIpString(strTmp)
                                If strFMTIp < "224.000.000.000" Then
                                    strIp = strTmp
                                End If
                            End If
                        Else
                            strFMTIp = FormatIpString(strIp)
                            If strFMTIp < "224.000.000.000" Then
                                strIp = strTmp
                            End If
                        End If
                    End If
                    If LenB(strIp) <> 0 Then
                        If strIp Like "10.*" Or strIp Like "192.168.*" Then
                            strPrivateIp = strIp
                            If blnXP And (arrIpRows(i).wType And MIB_IPADDR_PRIMARY) = MIB_IPADDR_PRIMARY Then
                                strPrivatePriIp = strIp
                            End If
                        ElseIf strFMTIp < "172.032.000.000" And strFMTIp >= "172.016.000.000" Then
                            strPrivateIp = strIp
                            If blnXP And (arrIpRows(i).wType And MIB_IPADDR_PRIMARY) = MIB_IPADDR_PRIMARY Then
                                strPrivatePriIp = strIp
                            End If
                        Else
                            strPubIp = strIp
                            If blnXP And (arrIpRows(i).wType And MIB_IPADDR_PRIMARY) = MIB_IPADDR_PRIMARY Then
                                strPubPriIP = strIp
                            End If
                        End If
                    End If
                Next
            End If
        End If
    End If
    If LenB(strPubPriIP) <> 0 Then
        IP = strPubPriIP
    ElseIf LenB(strPubIp) <> 0 Then
        IP = strPubIp
    ElseIf LenB(strPrivatePriIp) <> 0 Then
        IP = strPrivatePriIp
    ElseIf LenB(strPrivateIp) <> 0 Then
        IP = strPrivateIp
    End If
    Call Logger.PopMethod("ZLHelperMain.clsEnvironment.IP", IP)
    Exit Function
ErrH:
    If Logger.ErrCenter("ZLHelperMain.clsEnvironment.IP") = 1 Then
        Resume
    End If

    Call Logger.PopMethod("ZLHelperMain.clsEnvironment.IP")
End Function

'@方法    ComputerName
'   获取计算机名称
'@返回值  String
'
'@参数:
'@备注
'
Public Function ComputerName() As String
'功能：获取电脑名称
    Dim strComputer     As String
    Dim lngLen          As Long
    On Error GoTo ErrH
    Call Logger.PushMethod("ZLHelperMain.clsEnvironment.ComputerName")
    If LenB(mstrComputerName) = 0 Then
        strComputer = String(MAX_COMPUTERNAME_LENGTH + 1, Chr(0))
        lngLen = MAX_COMPUTERNAME_LENGTH + 1
        If GetComputerName(strComputer, lngLen) = 0 Then
            If Err.LastDllError = ERROR_INSUFFICIENT_BUFFER Then
                strComputer = String(lngLen + 1, Chr(0))
                If GetComputerName(strComputer, lngLen) <> 0 Then
                    ComputerName = TruncZero(strComputer)
                Else
                    Logger.Error "GetComputerName(Second)", "Error", GetLastDllErr(Err.LastDllError)
                End If
            Else
                Logger.Error "GetComputerName(First)", "Error", GetLastDllErr(Err.LastDllError)
            End If
        Else
            ComputerName = TruncZero(strComputer)
        End If
        mstrComputerName = ComputerName
    Else
        ComputerName = mstrComputerName
    End If
    Call Logger.PopMethod("ZLHelperMain.clsEnvironment.ComputerName", ComputerName)
    Exit Function
ErrH:
    If Logger.ErrCenter("ZLHelperMain.clsEnvironment.ComputerName") = 1 Then
        Resume
    End If

    Call Logger.PopMethod("ZLHelperMain.clsEnvironment.ComputerName")
End Function
'---------------------------------------------------------------------------
'                4、私有方法
'---------------------------------------------------------------------------

'@方法    IsInfoCached
'   指定的信息是否已经缓存
'@返回值  Boolean
'
'@参数:
'@备注
'
Private Function IsInfoCached(ByVal lngType As Long) As Boolean
    Dim lngIdx      As Long
    Dim lngTimes    As Long
    
    lngIdx = lngType \ 8
    lngTimes = lngType Mod 8
    IsInfoCached = (marrCahceInfo(lngIdx) And (2 ^ lngTimes)) = (2 ^ lngTimes)
End Function

'@方法    SetInfoCached
'   标记指定的信息是已经缓存
'@返回值
'
'@参数:
'@备注
'
Private Sub SetInfoCached(ByVal lngType As Long)
    Dim lngIdx      As Long
    Dim lngTimes    As Long
    
    lngIdx = lngType \ 8
    lngTimes = lngType Mod 8
    
    marrCahceInfo(lngIdx) = marrCahceInfo(lngIdx) Or (2 ^ lngTimes)
End Sub


'@方法    ProcessUserInfo_Sub
'   获取进程的用户相关信息
'@返回值  Boolean
'
'@参数:
'lngProcessID Long In
'   获取的进程ID
'strUserName String Out
'   返回用户名
'strDomainName String Out
'   域名
'strFullUserName String Out
'   用户全名。域\用户名
'@备注
'
Private Function ProcessUserInfo_Sub(ByVal lngProcessID As Long, strUserName As String, strDomainName As String, strFullUserName As String) As Boolean
    Dim lngLen      As Long
    Dim arrTmp      As Variant
    Dim strTmp      As String
    Dim lngProcess  As Long, hToken         As Long
    Dim BufferSize  As Long, InfoBuffer()   As Byte
    Dim tkUser      As TOKEN_USER
    Dim lngUserSize As Long, lngDomainSize      As Long, pUse   As Long
    
    On Error GoTo ErrH
    Call Logger.PushMethod("ZLHelperMain.clsEnvironment.ProcessUserInfo_Sub", lngProcessID, strUserName, strDomainName, strFullUserName)
    strFullUserName = ""
    strUserName = ""
    strDomainName = ""
    If lngProcessID = 0 Then
        strTmp = String(MAX_PATH, Chr$(0))
        lngLen = MAX_PATH
        If GetUserName(strTmp, lngLen) = 0 Then
            Logger.Error "GetUserName", "Error", GetLastDllErr(Err.LastDllError)
        Else
            strTmp = TruncZero(strTmp)
            If LenB(strTmp) <> 0 Then
                strUserName = strTmp
                
                strTmp = String(MAX_PATH, Chr$(0))
                lngLen = MAX_PATH
                '普通用户下，返回，电脑名\用户名，SYSTEM用户下返回工作组\电脑名
                If GetUserNameEx(NameSamCompatible, strTmp, lngLen) = 0 Then
                    Logger.Error "GetUserNameEx(QueryLength)", "Error", GetLastDllErr(Err.LastDllError)
                Else
                    strTmp = TruncZero(strTmp)
                    If LenB(strTmp) <> 0 Then
                        arrTmp = Split(strTmp & "", "\")
                        If arrTmp(1) = strUserName Then
                            strDomainName = arrTmp(0)
                        Else
                            strDomainName = arrTmp(1)
                            If Right(strDomainName, 1) = "$" Then
                                strDomainName = Mid(strDomainName, 1, Len(strDomainName) - 1)
                            End If
                        End If
                    End If
                End If
                strFullUserName = strDomainName & "\" & strUserName
                ProcessUserInfo_Sub = True
                
                mstrProcessUser = strUserName
                mstrProcessDomian = strDomainName
                mstrProcessUserFull = strFullUserName
                SetInfoCached CT_ProcessUser
            End If
        End If
    Else
        If Not Process.EnablePrivilege(SE_DEBUG_NAME) Then
            Logger.DebugEx "EnablePrivilege", "lngProcessID", lngProcessID, "结果", False
        End If
        lngProcess = OpenProcess(PROCESS_QUERY_INFORMATION, 0, lngProcessID)
        If lngProcess = 0 Then
            Logger.Error "OpenProcess", "Error", GetLastDllErr(Err.LastDllError)
        Else
            If OpenProcessToken(lngProcess, TOKEN_QUERY, hToken) = 0 Then
                Logger.Error "OpenProcessToken", "Error", GetLastDllErr(Err.LastDllError)
            Else
                If GetTokenInformation(hToken, ByVal TokenUser, 0, 0, BufferSize) = 0 Then ' Determine required buffer size
                    Logger.DebugEx "GetTokenInformation", "Error", GetLastDllErr(Err.LastDllError), "BufferSize", BufferSize
                End If
                If BufferSize Then
                    ReDim InfoBuffer(BufferSize - 1)
                    If GetTokenInformation(hToken, ByVal TokenUser, ByVal VarPtr(InfoBuffer(0)), BufferSize, BufferSize) = 0 Then
                        Logger.Error "GetTokenInformation", "Error", GetLastDllErr(Err.LastDllError)
                    Else
                        Call RtlMoveMemory(tkUser, InfoBuffer(0), LenB(tkUser))
                        strUserName = String(MAX_PATH, Chr(0))
                        strDomainName = String(MAX_PATH, Chr(0))
                        lngUserSize = MAX_PATH
                        lngDomainSize = MAX_PATH
                        If LookupAccountSid(vbNullString, tkUser.User.Sid, strUserName, lngUserSize, strDomainName, lngDomainSize, pUse) = 0 Then
                            Logger.Error "LookupAccountSid", "Error", GetLastDllErr(Err.LastDllError)
                            strDomainName = ""
                            strUserName = ""
                        Else
                            strUserName = TruncZero(strUserName)
                            strDomainName = TruncZero(strDomainName)
                            strFullUserName = strDomainName & "\" & strUserName
                            ProcessUserInfo_Sub = True
                        End If
                    End If
                End If
                If CloseHandle(hToken) = 0 Then
                    Logger.Error "CloseHandle(Token)", "Error", GetLastDllErr(Err.LastDllError)
                    ProcessUserInfo_Sub = False
                Else
                    hToken = 0
                End If
            End If
            If CloseHandle(lngProcess) = 0 Then
                Logger.Error "CloseHandle(Process)", "Error", GetLastDllErr(Err.LastDllError)
                ProcessUserInfo_Sub = False
            Else
                lngProcess = 0
            End If
        End If
    End If
    Call Logger.PopMethod("ZLHelperMain.clsEnvironment.ProcessUserInfo_Sub", ProcessUserInfo_Sub, strUserName, strDomainName, strFullUserName)
    Exit Function
ErrH:
    If Logger.ErrCenter("ZLHelperMain.clsEnvironment.ProcessUserInfo_Sub") = 1 Then
        Resume
    End If
    If hToken <> 0 Then
        If CloseHandle(hToken) = 0 Then
            Logger.Error "CloseHandle(Token)", "Error", GetLastDllErr(Err.LastDllError)
        End If
    End If
    If lngProcess <> 0 Then
        If CloseHandle(lngProcess) = 0 Then
            Logger.Error "CloseHandle(Process)", "Error", GetLastDllErr(Err.LastDllError)
        End If
    End If
    Call Logger.PopMethod("ZLHelperMain.clsEnvironment.ProcessUserInfo_Sub")
End Function


'@方法    GetUserSID
'   获取用户SID
'@返回值  String
'
'@参数:
'strAccountName String In
'   操作系统用户名，或域名\账户名
'@备注
'   主要用于读取注册表
Private Function GetUserSID(ByVal strAccountName As String) As String
    Dim lngRet      As Long, bytSid()   As Byte, cbSid  As Long
    Dim strDom      As String, cbDom    As Long
    Dim lpStrSid    As Long, peUse      As Long
    Dim lngError    As Long
    Dim strRet      As String
    
    On Error GoTo ErrH
    Call Logger.PushMethod("ZLHelperMain.clsEnvironment.GetUserSID")
    lngRet = LookupAccountName(0, StrPtr(strAccountName), ByVal 0, cbSid, 0, cbDom, peUse)
    lngError = Err.LastDllError
    If lngError = ERROR_INSUFFICIENT_BUFFER Then
        Logger.DebugEx "LookupAccountName", "Error", GetLastDllErr(lngError), "cbSid", cbSid, "cbDom", cbDom
        ReDim bytSid(cbSid): strDom = String$(cbDom, Chr$(0))
        lngError = 0
    End If
    If lngError = 0 Then
        If LookupAccountName(0, StrPtr(strAccountName), bytSid(0), cbSid, StrPtr(strDom), cbDom, peUse) = 0 Then
            Logger.Error "LookupAccountName(Second)", "Error", GetLastDllErr(Err.LastDllError)
        Else
            If ConvertSidToStringSid(bytSid(0), lpStrSid) <> 0 Then
                strRet = String$(lstrlen(lpStrSid), Chr(0))
                If lstrcpy(StrPtr(strRet), lpStrSid) <> 0 Then
                    If InStr(strRet, Chr$(0)) > 0 Then
                        strRet = Mid(strRet, 1, InStr(strRet, Chr$(0)) - 1)
                    End If
                Else
                    strRet = ""
                    Logger.Error "lstrcpy", "Error", GetLastDllErr(Err.LastDllError)
                End If
                If LocalFree(lpStrSid) <> 0 Then
                    '不管是否正常都报句柄无效
                    Logger.DebugEx "LocalFree", "Error", GetLastDllErr(Err.LastDllError)
                Else
                    lpStrSid = 0
                End If
            Else
                Logger.Error "ConvertSidToStringSid", "Error", GetLastDllErr(Err.LastDllError)
            End If
        End If
    Else
        Logger.Warn "LookupAccountName(First)", "Error", GetLastDllErr(lngError)
    End If
    GetUserSID = strRet
    Call Logger.PopMethod("ZLHelperMain.clsEnvironment.GetUserSID", GetUserSID)
    Exit Function
ErrH:
    If Logger.ErrCenter("ZLHelperMain.clsEnvironment.GetUserSID") = 1 Then
        Resume
    End If
    If lpStrSid <> 0 Then
        If LocalFree(lpStrSid) <> 0 Then
            Logger.Error "LocalFree", "Error", GetLastDllErr(Err.LastDllError)
        End If
    End If
    Call Logger.PopMethod("ZLHelperMain.clsEnvironment.GetUserSID")
End Function


'---------------------------------------------------------------------------
'                5、对象方法与事件
'---------------------------------------------------------------------------
Private Sub Class_Terminate()
    Set mobjFSO = Nothing
End Sub
