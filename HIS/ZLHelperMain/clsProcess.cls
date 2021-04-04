VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsProcess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@模块 clsProcess-2019/7/4
'@编写 lshuo
'@功能
'
'@引用
'
'@备注
'
Option Explicit
'---------------------------------------------------------------------------
'                0、API和常量声明
'---------------------------------------------------------------------------
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
Private Declare Function CreateToolhelp32Snapshot Lib "kernel32" (ByVal dwFlags As Long, ByVal th32ProcessID As Long) As Long
'@原型
'    HANDLE CreateToolhelp32Snapshot(
'      DWORD dwFlags,
'      DWORD th32ProcessID
'    );
'@功能
'    获取指定进程的快照，以及这些进程使用的堆、模块和线程。
'@参数
'dwFlags
'    快照中包含的系统部分?此参数可以是以下值中的一个或多个?
Private Const TH32CS_INHERIT        As Long = &H80000000
'    指示快照句柄是可继承的?
Private Const TH32CS_SNAPHEAPLIST   As Long = &H1
'    在快照中包含th32ProcessID中指定的进程的所有堆。要枚举堆，请参见Heap32ListFirst。
Private Const TH32CS_SNAPMODULE     As Long = &H8
'    在快照中包含th32ProcessID中指定的进程的所有模块。要枚举这些模块，请先查看module32。如果函数使用ERROR_BAD_LENGTH失败，则重试该函数，直到成功。
'    64位Windows:在32位进程中使用该标志包括th32ProcessID中指定的进程的32位模块，而在64位进程中使用该标志包括64位模块。要包含来自64位进程的th32ProcessID中指定的32位进程的32位模块，请使用TH32CS_SNAPMODULE32标志。
Private Const TH32CS_SNAPMODULE32   As Long = &H10
'    在快照中包含从64位进程调用时在th32ProcessID中指定的进程的所有32位模块。此标志可以与TH32CS_SNAPMODULE或TH32CS_SNAPALL组合使用。如果函数使用ERROR_BAD_LENGTH失败，则重试该函数，直到成功。
Private Const TH32CS_SNAPPROCESS    As Long = &H2
'    在快照中包含系统中的所有进程。要枚举这些进程，请先查看process32。
Private Const TH32CS_SNAPTHREAD     As Long = &H4
'    在快照中包含系统中的所有线程。要枚举线程，请先查看thread32。
'    要标识属于特定进程的线程，请在枚举线程时将其进程标识符与THREADENTRY32结构的th32OwnerProcessID成员进行比较。
Private Const TH32CS_SNAPALL        As Long = TH32CS_SNAPHEAPLIST Or TH32CS_SNAPMODULE Or TH32CS_SNAPPROCESS Or TH32CS_SNAPTHREAD
'    包括系统中的所有进程和线程，以及th32ProcessID中指定的进程的堆和模块。相当于使用OR操作('|')组合指定th32cs_snaphealthplist、TH32CS_SNAPMODULE、TH32CS_SNAPPROCESS和TH32CS_SNAPTHREAD值。
'th32ProcessID
'    快照中包含的进程的进程标识符。此参数可以为零以指示当前进程。当指定th32cs_snaphealthplist、TH32CS_SNAPMODULE、TH32CS_SNAPMODULE32或TH32CS_SNAPALL值时，将使用此参数。否则，它将被忽略，所有进程都包含在快照中。
'    如果指定的进程是空闲进程或CSRSS进程之一，则此函数失败，最后一个错误代码为ERROR_ACCESS_DENIED，因为它们的访问限制阻止用户级代码打开它们。
'    如果指定的进程是64位进程，而调用者是32位进程，则此函数将失败，最后一个错误代码是ERROR_PARTIAL_COPY(299)。
'@返回值
'    如果函数成功，它将返回指定快照的打开句柄。
'    如果函数失败，它将返回INVALID_HANDLE_VALUE。要获取扩展的错误信息，请调用GetLastError。可能的错误代码包括ERROR_BAD_LENGTH。
'@备注
'    此函数所捕获的快照由其他工具帮助函数检查，以提供其结果。对快照的访问是只读的。快照句柄充当对象句柄，并且受关于它在哪些进程和线程中有效的相同规则的约束。
'    要枚举所有进程的堆或模块状态，请指定TH32CS_SNAPALL并将th32ProcessID设置为零。然后，对于快照中的每个附加进程，再次调用CreateToolhelp32Snapshot，指定其进程标识符和th32cs_snaphealthplist或TH32_SNAPMODULE值。
'    当为当前进程以外的进程获取包含堆和模块的快照时，CreateToolhelp32Snapshot函数可能会由于各种原因失败或返回不正确的信息。例如，如果目标进程中的加载器数据表已损坏或未初始化，或者模块列表在函数调用期间由于加载或卸载dll而发生更改，函数可能会因为ERROR_BAD_LENGTH或其他错误代码而失败。确保目标进程没有在挂起状态下启动，然后再次尝试调用该函数。如果函数在使用TH32CS_SNAPMODULE或TH32CS_SNAPMODULE32调用时错误地使用ERROR_BAD_LENGTH，则再次调用该函数，直到成功为止。
'    TH32CS_SNAPMODULE和TH32CS_SNAPMODULE32标志不会检索装载了LOAD_LIBRARY_AS_DATAFILE或类似标志的模块的句柄。有关更多信息，请参见LoadLibraryEx。
'    要销毁快照，请使用close句柄函数。
'    注意，您可以使用QueryFullProcessImageName函数从32位进程检索32位和64位进程的可执行映像的全名。
'@Requirements
'    Minimum supported client    Windows XP [desktop apps only]
'    Minimum supported server    Windows Server 2003 [desktop apps only]
'    Target Platform Windows
'    Header tlhelp32.h
'    Library kernel32.lib
'    dll kernel32.dll
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
Private Declare Sub ExitProcess Lib "kernel32" (ByVal uExitCode As Long)
'@原型
'    void ExitProcess(
'      UINT uExitCode
'    );
'@功能
'    结束调用进程及其所有线程?
'@返回值
'    此函数不返回值?
'@备注
'    使用GetExitCodeProcess函数检索进程的退出值?使用GetExitCodeThread函数检索线程的退出值?
'    退出进程会导致以下情况:
'        进程中的所有线程(调用线程除外)在没有收到DLL_THREAD_DETACH通知的情况下终止执行。
'        在步骤1中终止的所有线程的状态都有信号?
'        使用DLL_PROCESS_DETACH调用所有加载的动态链接库(dll)的入口点函数。
'        在所有附加的dll执行任何进程终止代码之后，ExitProcess函数将终止当前进程，包括调用线程。
'        调用线程的状态变为有信号状态?
'        进程打开的所有对象句柄都是关闭的?
'        流程的终止状态从STILL_ACTIVE更改为流程的退出值?
'        进程对象的状态变为有信号状态，满足任何等待进程终止的线程。
'    如果进程中的一个终止线程持有锁，而其中一个加载的DLL中的DLL分离代码试图获得相同的锁，那么调用ExitProcess将导致死锁。相反，如果进程通过调用TerminateProcess来终止，则不会通知进程所附加的dll。因此，如果不知道进程中所有线程的状态，最好调用TerminateProcess而不是ExitProcess。注意，从应用程序的主函数返回会导致对ExitProcess的调用。
'    在DLL中调用ExitProcess可能导致意外的应用程序或系统错误。只有在您知道哪些应用程序或系统组件将加载DLL，并且在此上下文中调用ExitProcess是安全的情况下，才能确保从DLL调用ExitProcess。
'    退出进程不会导致子进程终止?
'    退出流程并不一定要从操作系统中删除流程对象。当进程的最后一个句柄关闭时，进程对象将被删除。
'@Requirements
'    Minimum supported client    Windows XP [desktop apps only]
'    Minimum supported server    Windows Server 2003 [desktop apps only]
'    Target Platform Windows
'    Header  processthreadsapi.h (include Windows Server 2003, Windows Vista, Windows 7, Windows Server 2008 Windows Server 2008 R2, Windows.h)
'    Library kernel32.lib
'    dll kernel32.dll
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
'    伪句柄是一个特殊的常量，当前(句柄)-1，它被解释为当前进程句柄。为了与未来的操作系统兼容，最好调用GetCurrentProcess，而不是硬编码这个常量。无论何时需要进程句柄，调用进程都可以使用伪句柄来指定自己的进程。子进程不会继承伪句柄。
'    此句柄具有PROCESS_ALL_ACCESS访问进程对象的权限。有关更多信息，请参见进程安全和访问权限。
'    Windows Server 2003和Windows XP:此句柄具有进程安全描述符允许的对进程主令牌的最大访问权。
'    通过在调用DuplicateHandle函数时将伪句柄指定为源句柄，进程可以为自己创建一个“真实”句柄，该句柄在其他进程的上下文中是有效的，或者可以被其他进程继承。进程还可以使用OpenProcess函数为自己打开一个实句柄。
'    当不再需要伪句柄时，不需要关闭它。使用伪句柄调用close句柄函数没有效果。如果用DuplicateHandle复制伪句柄，则必须关闭重复句柄。
'@Requirements
'Minimum supported client        Windows XP [desktop apps | UWP apps]
'Minimum supported server        Windows Server 2003 [desktop apps | UWP apps]
'Minimum supported phone         Windows Phone 8
'Header                          WinBase.h on Windows XP, Windows Server 2003, Windows Vista, Windows 7, Windows Server 2008 and Windows Server 2008 R2 (include Windows.h);Processthreadsapi.h on Windows 8 and Windows Server 2012
'Library                         kernel32.lib
'dll                             kernel32.dll
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
    GroupCount                  As Long
    Groups(ANYSIZE_ARRAY - 1)   As SID_AND_ATTRIBUTES
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
'   需要分配进程的主要令牌。用户权限: 替换进程级令牌。
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
'   需要为在用户上下文中运行的应用程序分配更多内存。用户权限: 增加进程工作集。
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

Private Const MAX_PATH                          As Long = 260
Private Const MAX_MODULE_NAME32                 As Long = 255

Private Type MODULEENTRY32
    dwSize                                      As Long
    th32ModuleID                                As Long
    th32ProcessID                               As Long
    GlblcntUsage                                As Long
    ProccntUsage                                As Long
    modBaseAddr                                 As Long
    modBaseSize                                 As Long
    hModule                                     As Long
    szModule(MAX_MODULE_NAME32 * 2 - 1)         As Byte
    szExePath(MAX_PATH * 2 - 1)                 As Byte
End Type
'@功能
'    描述属于指定进程的模块列表中的项。
'@原型
'    typedef struct tagMODULEENTRY32 {
'      DWORD   dwSize;
'      DWORD   th32ModuleID;
'      DWORD   th32ProcessID;
'      DWORD   GlblcntUsage;
'      DWORD   ProccntUsage;
'      BYTE    *modBaseAddr;
'      DWORD   modBaseSize;
'      HMODULE hModule;
'      char    szModule[MAX_MODULE_NAME32 + 1];
'      char    szExePath[MAX_PATH];
'    } MODULEENTRY32;
'@成员
'dwSize
'    结构的大小，以字节为单位。在调用Module32First函数之前，将这个成员设置为sizeof(MODULEENTRY32)。如果不初始化dwSize, Module32First将失败。
'th32ModuleID
'    这个成员不再使用，并且总是被设置为1。
'th32ProcessID
'    要检查模块的进程的标识符?
'GlblcntUsage
'    模块的负载计数，通常没有意义，通常等于0xFFFF。
'ProccntUsage
'    模块的负载计数(与GlblcntUsage相同)，通常没有意义，通常等于0xFFFF。
'modBaseAddr
'    模块在所属进程上下文中的基本地址?
'modBaseSize
'    模块的大小，以字节为单位。
'hModule
'    拥有进程上下文中模块的句柄?
'szModule
'    模块的名字?
'szExePath
'    模块路径?
'@备注
'    modBaseAddr和hModule成员仅在th32ProcessID指定的进程上下文中有效?
'@Requirements
'    Minimum supported client    Windows XP [desktop apps only]
'    Minimum supported server    Windows Server 2003 [desktop apps only]
'    Header tlhelp32.h
Private Declare Function Module32First Lib "kernel32" Alias "Module32FirstW" (ByVal hSnapshot As Long, lppe As MODULEENTRY32) As Long
'@原型
'    BOOL Module32First(
'      HANDLE          hSnapshot,
'      LPMODULEENTRY32 lpme
'    );
'@功能
'    检索与进程关联的第一个模块的信息?
'@参数
'hSnapshot
'    从先前对CreateToolhelp32Snapshot函数的调用返回快照句柄?
'lpme
'    指向MODULEENTRY32结构的指针?
'@返回值
'    如果模块列表的第一个条目已复制到缓冲区，则返回TRUE，否则返回FALSE。如果不存在模块或快照不包含模块信息，则GetLastError函数返回ERROR_NO_MORE_FILES错误值。
'@备注
'    调用应用程序必须将MODULEENTRY32的dwSize成员设置为结构的大小(以字节为单位)。
'    要检索与指定进程关联的其他模块的信息，请使用Module32Next函数。
'@Requirements
'    Minimum supported client    Windows XP [desktop apps only]
'    Minimum supported server    Windows Server 2003 [desktop apps only]
'    Target Platform Windows
'    Header tlhelp32.h
'    Library kernel32.lib
'    dll kernel32.dll
Private Declare Function Module32Next Lib "kernel32" Alias "Module32NextW" (ByVal hSnapshot As Long, lppe As MODULEENTRY32) As Long
'@原型
'    BOOL Module32Next(
'      HANDLE          hSnapshot,
'      LPMODULEENTRY32 lpme
'    );
'@功能
'    检索与进程或线程关联的下一个模块的信息?
'@参数
'hSnapshot
'    从先前对CreateToolhelp32Snapshot函数的调用返回快照句柄?
'lpme
'    指向MODULEENTRY32结构的指针?
'@返回值
'    如果模块列表的下一项已复制到缓冲区，则返回TRUE，否则返回FALSE。如果不再存在模块，则GetLastError函数返回ERROR_NO_MORE_FILES错误值。
Private Const ERROR_NO_MORE_FILES       As Long = 18
'@备注
'    要检索与进程关联的第一个模块的信息，请使用Module32First函数。
'@Requirements
'    Minimum supported client    Windows XP [desktop apps only]
'    Minimum supported server    Windows Server 2003 [desktop apps only]
'    Target Platform Windows
'    Header tlhelp32.h
'    Library kernel32.lib
'    dll kernel32.dll
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
'    对进程对象的访问?此访问权限将根据进程的安全描述符进行检查?该参数可以是一个或多个进程访问权限?
'    如果调用方启用了SeDebugPrivilege特权，则不管安全描述符的内容如何，都将授予请求的访问权。
'bInheritHandle
'    如果该值为真，则由该进程创建的进程将继承该句柄。否则，进程不会继承此句柄。
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
'   进程的句柄，其访问令牌已打开。进程必须具有PROCESS_QUERY_INFORMATION访问权限。
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



Private Type PROCESSENTRY32
    dwSize                  As Long
    cntUsage                As Long
    th32ProcessID           As Long
    th32DefaultHeapID       As Long
    th32ModuleID            As Long
    cntThreads              As Long
    th32ParentProcessID     As Long
    pcPriClassBase          As Long
    dwFlags                 As Long
    sExeFile(MAX_PATH * 2 - 1) As Byte
End Type
'@功能
'    描述捕获快照时驻留在系统地址空间中的进程列表中的项?
'@原型
'    typedef struct tagPROCESSENTRY32 {
'      DWORD     dwSize;
'      DWORD     cntUsage;
'      DWORD     th32ProcessID;
'      ULONG_PTR th32DefaultHeapID;
'      DWORD     th32ModuleID;
'      DWORD     cntThreads;
'      DWORD     th32ParentProcessID;
'      LONG      pcPriClassBase;
'      DWORD     dwFlags;
'      CHAR      szExeFile[MAX_PATH];
'    } PROCESSENTRY32;
'@成员
'dwSize
'    结构的大小，以字节为单位。在调用Process32First函数之前，将这个成员设置为sizeof(PROCESSENTRY32)。如果不初始化dwSize, Process32First将失败。
'cntUsage
'    这个成员不再使用，总是被设置为零。
'th32ProcessID
'    进程标识符?
'th32DefaultHeapID
'    这个成员不再使用，总是被设置为零。
'th32ModuleID
'    这个成员不再使用，总是被设置为零。
'cntThreads
'    进程启动的执行线程数?
'th32ParentProcessID
'    创建此进程(其父进程)的进程的标识符。
'pcPriClassBase
'    此进程创建的任何线程的基本优先级?
'dwFlags
'    这个成员不再使用，并且总是被设置为零。
'szExeFile
'    进程的可执行文件的名称。要检索可执行文件的完整路径，请调用Module32First函数并检查返回的MODULEENTRY32结构的szExePath成员。但是，如果调用的进程是32位进程，则必须调用QueryFullProcessImageName函数来检索64位进程的可执行文件的完整路径。
'@Requirements
'    Minimum supported client    Windows XP [desktop apps only]
'    Minimum supported server    Windows Server 2003 [desktop apps only]
'    Header tlhelp32.h
Private Declare Function Process32First Lib "kernel32" Alias "Process32FirstW" (ByVal hSnapshot As Long, lppe As PROCESSENTRY32) As Long
'@原型
'    BOOL Process32First(
'      HANDLE           hSnapshot,
'      LPPROCESSENTRY32 lppe
'    );
'@功能
'    检索系统快照中遇到的第一个进程的信息。
'@参数
'hSnapshot
'    从先前对CreateToolhelp32Snapshot函数的调用返回快照句柄?
'lppe
'    指向PROCESSENTRY32结构的指针。它包含进程信息，例如可执行文件的名称、进程标识符和父进程的进程标识符。
'@返回值
'    如果进程列表的第一个条目已复制到缓冲区，则返回TRUE，否则返回FALSE。如果没有进程存在或快照不包含进程信息，则GetLastError函数返回ERROR_NO_MORE_FILES错误值。
'@备注
'    调用应用程序必须将PROCESSENTRY32的dwSize成员设置为结构的大小(以字节为单位)。
'    要检索关于同一快照中记录的其他进程的信息，请使用Process32Next函数。
'@Requirements
'    Minimum supported client    Windows XP [desktop apps only]
'    Minimum supported server    Windows Server 2003 [desktop apps only]
'    Target Platform Windows
'    Header tlhelp32.h
'    Library kernel32.lib
'    dll kernel32.dll
Private Declare Function Process32Next Lib "kernel32" Alias "Process32NextW" (ByVal hSnapshot As Long, lppe As PROCESSENTRY32) As Long
'@原型
'    BOOL Process32Next(
'      HANDLE           hSnapshot,
'      LPPROCESSENTRY32 lppe
'    );
'@功能
'    检索系统快照中记录的下一个进程的信息。
'@参数
'hSnapshot
'    从先前对CreateToolhelp32Snapshot函数的调用返回快照句柄?
'lppe
'    指向PROCESSENTRY32结构的指针?
'@返回值
'    如果进程列表的下一项已复制到缓冲区，则返回TRUE，否则返回FALSE。如果没有进程存在或快照不包含进程信息，则GetLastError函数返回ERROR_NO_MORE_FILES错误值。
'@备注
'    要检索快照中记录的第一个进程的信息，请使用Process32First函数。
'@Requirements
'    Minimum supported client    Windows XP [desktop apps only]
'    Minimum supported server    Windows Server 2003 [desktop apps only]
'    Target Platform Windows
'    Header tlhelp32.h
'    Library kernel32.lib
'    dll kernel32.dll
Private Declare Sub RtlMoveMemory Lib "kernel32" (Destination As Any, Source As Any, ByVal Length As Long)
'说明：将内存块从一个位置移动到另一个位置
'Destination:指向移动目的地起始地址的指针。
'Source:指向要移动的内存块起始地址的指针。
'Length:内存块的大小以字节为单位移动。
'注意事项：这个函数定义为RtlMoveMemory函数。它的实现是内联的。有关更多信息，请参见WinBase。h和Winnt.h。源和目标块可能会重叠。
'           第一个参数，目的地，必须足够大，以容纳长度字节的源;否则，可能会出现缓冲区溢出。这可能导致拒绝服务攻击，如果有访问违反，或者在最坏的情况下，允许攻击者向您的进程注入可执行代码。如果目的地是一个基于堆栈的缓冲区，则尤其如此。要注意，最后一个参数，长度，是将字节复制到目的地的数量，而不是目的地的大小。
Private Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
'@原型
'    BOOL TerminateProcess(
'      HANDLE hProcess,
'      UINT uExitCode
'    );
'@功能
'    终止指定的进程及其所有线程?
'@参数
'hProcess
'    要终止的进程的句柄?
'    句柄必须具有PROCESS_TERMINATE访问权限。有关更多信息，请参见流程安全和访问权限。
'uExitCode
'    进程和线程将使用的退出代码将因该调用而终止?使用GetExitCodeProcess函数检索进程的退出值?使用GetExitCodeThread函数检索线程的退出值?
'@返回值
'    如果函数成功，返回值为非零。
'    如果函数失败，返回值为零。要获取扩展的错误信息，请调用GetLastError。
'@备注
'    TerminateProcess函数用于无条件地导致进程退出。如果使用TerminateProcess而不是ExitProcess，动态链接库(dll)维护的全局数据的状态可能会受到影响。
'    该函数停止进程中所有线程的执行，并请求取消所有挂起的I/O。在所有挂起的I/O完成或取消之前，终止的进程不能退出。当一个进程终止时，它的内核对象不会被销毁，直到所有对该进程具有打开句柄的进程都释放了这些句柄。
'    TerminateProcess是异步的;它启动终止并立即返回。如果需要确保进程已终止，请使用进程的句柄调用WaitForSingleObject函数。
'    进程不能阻止自己被终止?
'@Requirements
'    Minimum supported client    Windows XP [desktop apps | UWP apps]
'    Minimum supported server    Windows Server 2003 [desktop apps | UWP apps]
'    Target Platform Windows
'    Header  processthreadsapi.h (include Windows Server 2003, Windows Vista, Windows 7, Windows Server 2008 Windows Server 2008 R2, Windows.h)
'    Library kernel32.lib
'    dll kernel32.dll
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
'---------------------------------------------------------------------------
'                1、常规变量
'---------------------------------------------------------------------------

'---------------------------------------------------------------------------
'                2、属性变量与定义
'---------------------------------------------------------------------------

'---------------------------------------------------------------------------
'                3、公共方法
'---------------------------------------------------------------------------
'@方法    EnablePrivilegeToken
'   提升指定权限
'@返回值  Boolean
'
'@参数:
'lnghToken Long In
'   要调整权限的句柄
'strPrivilegeName String In
'   要提升的权限名称
'@备注
'
Public Function EnablePrivilegeToken(ByVal lnghToken As Long, ByVal strPrivilegeName As String, Optional ByVal blnEnable As Boolean) As Boolean
    Dim hToken              As Long
    Dim tmpLuid             As LUID
    Dim tkp                 As TOKEN_PRIVILEGES
    Dim tkpNewButIgnored    As TOKEN_PRIVILEGES
    Dim lBufferNeeded       As Long


    On Error GoTo ErrH
    Call Logger.PushMethod("ZLHelperMain.clsProcess.EnablePrivilegeToken", strPrivilegeName)
    If LookupPrivilegeValue(vbNullString, strPrivilegeName, tmpLuid) = 0 Then
        Logger.Error "LookupPrivilegeValue", "Error", GetLastDllErr(Err.LastDllError)
    Else
        tkp.PrivilegeCount = 1
        tkp.Privileges(0).PLUID = tmpLuid
        tkp.Privileges(0).Attributes = SE_PRIVILEGE_ENABLED
        If AdjustTokenPrivileges(lnghToken, 0, tkp, Len(tkp), tkpNewButIgnored, lBufferNeeded) = 0 Then
            Logger.Error "AdjustTokenPrivileges", "Error", GetLastDllErr(Err.LastDllError)
        Else
            EnablePrivilegeToken = True
        End If
    End If
    Call Logger.PopMethod("ZLHelperMain.clsProcess.EnablePrivilegeToken", EnablePrivilegeToken)
    Exit Function
ErrH:
    If Logger.ErrCenter("ZLHelperMain.clsProcess.EnablePrivilegeToken") = 1 Then
        Resume
    End If
    Call Logger.PopMethod("ZLHelperMain.clsProcess.EnablePrivilegeToken")
End Function
'@方法    EnablePrivilege
'   提升当前进程权限
'@返回值  Boolean
'
'@参数:
'strPrivilegeName String In
'   要提升的权限名称
'@备注
'
Public Function EnablePrivilege(Optional ByVal strPrivilegeName As String, Optional ByVal blnEnable As Boolean) As Boolean
    Dim hToken              As Long
    Dim tmpLuid             As LUID
    Dim tkp                 As TOKEN_PRIVILEGES
    Dim tkpNewButIgnored    As TOKEN_PRIVILEGES
    Dim lBufferNeeded       As Long


    On Error GoTo ErrH
    Call Logger.PushMethod("ZLHelperMain.clsProcess.EnablePrivilege", strPrivilegeName)

    If OpenProcessToken(GetCurrentProcess(), TOKEN_ADJUST_PRIVILEGES Or TOKEN_QUERY, hToken) = 0 Then
        Logger.Error "OpenProcessToken", "Error", GetLastDllErr(Err.LastDllError)
    Else
        If LookupPrivilegeValue(vbNullString, strPrivilegeName, tmpLuid) = 0 Then
            Logger.Error "LookupPrivilegeValue", "Error", GetLastDllErr(Err.LastDllError)
        Else
            tkp.PrivilegeCount = 1
            tkp.Privileges(0).PLUID = tmpLuid
            tkp.Privileges(0).Attributes = SE_PRIVILEGE_ENABLED
            If AdjustTokenPrivileges(hToken, 0, tkp, Len(tkp), tkpNewButIgnored, lBufferNeeded) = 0 Then
                Logger.Error "AdjustTokenPrivileges", "Error", GetLastDllErr(Err.LastDllError)
            Else
                EnablePrivilege = True
            End If
        End If
        If CloseHandle(hToken) = 0 Then
            Logger.Error "CloseHandle", "Error", GetLastDllErr(Err.LastDllError)
        Else
            hToken = 0
        End If
    End If
    Call Logger.PopMethod("ZLHelperMain.clsProcess.EnablePrivilege", EnablePrivilege)
    Exit Function
ErrH:
    If Logger.ErrCenter("ZLHelperMain.clsProcess.EnablePrivilege") = 1 Then
        Resume
    End If
    If hToken <> 0 Then
        If CloseHandle(hToken) = 0 Then
            Logger.Error "CloseHandle", "Error", GetLastDllErr(Err.LastDllError)
        End If
    End If
    Call Logger.PopMethod("ZLHelperMain.clsProcess.EnablePrivilege")
End Function
'@方法    PrivilegeList
'   获取当前进程的权限列表
'@返回值  String
'
'@参数:
'@备注
'
Public Function PrivilegeList() As String
    Dim hProcessToken       As Long
    Dim BufferSize          As Long
    Dim i                   As Integer
    Dim tpTokenPrivileges() As LUID_AND_ATTRIBUTES
    Dim strName             As String
    Dim strReturn           As String
    
    On Error GoTo ErrH
    Call Logger.PushMethod("ZLHelperMain.clsProcess.PrivilegeList")
    If OpenProcessToken(GetCurrentProcess(), TOKEN_QUERY, hProcessToken) <> 0 Then
        If GetTokenInformation(hProcessToken, ByVal TokenPrivileges, 0, 0, BufferSize) = 0 Then ' Determine required buffer size
            Logger.DebugEx "GetTokenInformation(First)", "Error", GetLastDllErr(Err.LastDllError), "BufferSize", BufferSize
        End If
        If BufferSize Then
            ReDim InfoBuffer((BufferSize \ 4) - 1) As Long
            If GetTokenInformation(hProcessToken, ByVal TokenPrivileges, InfoBuffer(0), BufferSize, BufferSize) = 0 Then
                Logger.Error "GetTokenInformation(Second)", "Error", GetLastDllErr(Err.LastDllError)
            Else
                ReDim tpTokenPrivileges(InfoBuffer(0) - 1)
                Call RtlMoveMemory(tpTokenPrivileges(0), InfoBuffer(1), Len(tpTokenPrivileges(0)) * InfoBuffer(0))
                For i = 0 To UBound(tpTokenPrivileges)
                    strName = String(256, Chr$(0))
                    If LookupPrivilegeName(vbNullString, tpTokenPrivileges(i).PLUID, strName, Len(strName)) <> 0 Then
                        strName = TruncZero(strName)
                    Else
                        Logger.Error "LookupPrivilegeName", "Error", GetLastDllErr(Err.LastDllError)
                        strName = ""
                    End If
                    strReturn = strReturn & i & ":" & strName & "-" & tpTokenPrivileges(i).Attributes & vbNewLine
                Next
            End If
        End If
        If CloseHandle(hProcessToken) = 0 Then
            Logger.Error "CloseHandle", "Error", GetLastDllErr(Err.LastDllError)
        Else
            hProcessToken = 0
        End If
    Else
        Logger.Error "OpenProcessToken", "Error", GetLastDllErr(Err.LastDllError)
    End If
    PrivilegeList = strReturn
    Call Logger.PopMethod("ZLHelperMain.clsProcess.PrivilegeList", PrivilegeList)
    Exit Function
ErrH:
    If Logger.ErrCenter("ZLHelperMain.clsProcess.PrivilegeList") = 1 Then
        Resume
    End If
    If hProcessToken <> 0 Then
        If CloseHandle(hProcessToken) = 0 Then
            Logger.Error "CloseHandle", "Error", GetLastDllErr(Err.LastDllError)
        End If
    End If
    Call Logger.PopMethod("ZLHelperMain.clsProcess.PrivilegeList")
End Function
'@方法    ProcessesByProcessName
'   根据进程文件名查找进程，
'@返回值  Variant
'   每个元素均为进程ID
'@参数:
'strProcessName String In
'   进程文件名
'lngExcludeProcID Long In(Optional)
'   排除检查的进程ID
'blnCurrentSession  Long In(Optional，Defualt=True)
'   是否查找当前会话，否则只搜索所有会话（ADMIN用户生效）
'@备注
'   也可以按全路径查找，此时，必须strProcessName传入全路径
Public Function ProcessesByProcessName(ByVal strProcessName As String, Optional ByVal lngExcludeProcID As Long, Optional ByVal blnCurrentSession As Boolean = True) As Variant
    Dim arrProcess()    As Long
    Dim uProcess        As PROCESSENTRY32
    Dim lngProcID       As Long
    Dim lngSnapShot     As Long, lngRet             As Long
    Dim strFindName     As String, strFileName      As String, strFilePath      As String
    Dim lngPid          As Long
    Dim blnFullPath     As Boolean, blnDo           As Boolean
    
    On Error GoTo ErrH
    Call Logger.PushMethod("ZLHelperMain.clsProcess.ProcessesByProcessName", strProcessName, lngExcludeProcID)
    strProcessName = UCase(strProcessName)
    blnFullPath = InStr(strProcessName, "\") > 0
    If blnFullPath Then
        strFileName = Mid(strProcessName, InStrRev(strProcessName, "\") + 1)
        strFilePath = Mid(strProcessName, 1, InStrRev(strProcessName, "\") - 1)
    Else
        strFileName = strProcessName
        strFilePath = ""
    End If
    
    lngSnapShot = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0)
    If lngSnapShot = 0 Then
        Logger.Error "CreateToolhelp32Snapshot", "Error", GetLastDllErr(Err.LastDllError)
    Else
        uProcess.dwSize = LenB(uProcess)
        If (Process32First(lngSnapShot, uProcess)) = 0 Then
            If Err.LastDllError <> ERROR_NO_MORE_FILES Then
                Logger.Error "Process32First", "Error", GetLastDllErr(Err.LastDllError)
            Else
                Logger.DebugEx "Process32First", "Error", "ERROR_NO_MORE_FILES"
            End If
        Else
            Do
                blnDo = False
                If lngExcludeProcID <> uProcess.th32ProcessID Then
                    If blnCurrentSession Then
                        If Environment.SessionID = Environment.SessionID(uProcess.th32ProcessID) Then
                            blnDo = True
                        End If
                    Else
                        blnDo = True
                    End If
                End If
                If blnDo Then
                    strFindName = uProcess.sExeFile
                    strFindName = UCase(TruncZero(strFindName))
                    If strFindName = strFileName Then
                        If blnFullPath Then
                            If UCase(Environment.StartExePath(uProcess.th32ProcessID)) = strProcessName Then
                                ReDim Preserve arrProcess(UboundEx(arrProcess) + 1)
                                arrProcess(UboundEx(arrProcess)) = uProcess.th32ProcessID
                            End If
                        Else
                            ReDim Preserve arrProcess(UboundEx(arrProcess) + 1)
                            arrProcess(UboundEx(arrProcess)) = uProcess.th32ProcessID
                        End If
                    End If
                End If
                lngRet = Process32Next(lngSnapShot, uProcess)
                If lngRet = 0 Then
                    If Err.LastDllError <> ERROR_NO_MORE_FILES Then
                        Logger.Error "Process32Next", "Error", GetLastDllErr(Err.LastDllError)
                    Else
                        Logger.DebugEx "Process32Next", "Error", "ERROR_NO_MORE_FILES"
                    End If
                End If
            Loop Until (lngRet = 0)
        End If
        If CloseHandle(lngSnapShot) = 0 Then
            Logger.Error "CloseHandle", "Error", GetLastDllErr(Err.LastDllError)
        Else
            lngSnapShot = 0
        End If
    End If
    ProcessesByProcessName = arrProcess
    Call Logger.PopMethod("ZLHelperMain.clsProcess.ProcessesByProcessName", ProcessesByProcessName)
    Exit Function
ErrH:
    If Logger.ErrCenter("ZLHelperMain.clsProcess.ProcessesByProcessName") = 1 Then
        Resume
    End If
    If lngSnapShot <> 0 Then
        If CloseHandle(lngSnapShot) = 0 Then
            Logger.Error "CloseHandle", "Error", GetLastDllErr(Err.LastDllError)
        End If
    End If
    Call Logger.PopMethod("ZLHelperMain.clsProcess.ProcessesByProcessName")
End Function

'@方法    ProcessesByModuleName
'   根据进程加载的文件名查找进程
'@返回值  Variant
'   每个元素均为进程ID
'@参数:
'strModuleName String In
'   进程加载的模块文件名
'lngExcludeProcID Long In(Optional)
'   排除检查的进程ID
'blnCurrentSession  Long In(Optional，Defualt=True)
'   是否查找当前会话，否则只搜索所有会话（ADMIN用户生效）
'@备注
'   也可以按全路径查找，此时，必须strModuleName传入全路径
Public Function ProcessesByModuleName(ByVal strModuleName As String, Optional ByVal lngExcludeProcID As Long, Optional ByVal blnCurrentSession As Boolean = True) As Variant
    Dim arrProcess()        As Long
    Dim uProcess            As PROCESSENTRY32, uMdlInfor    As MODULEENTRY32
    Dim lngProcID           As Long
    Dim lngSnapShot         As Long, lngRet                 As Long
    Dim lngMdlSnapShot      As Long
    Dim strFindName         As String, strFileName      As String, strFilePath      As String
    Dim strFindModule       As String, strModulePath    As String
    Dim lngPid              As Long
    Dim blnFullPath         As Boolean, blnDo           As Boolean
    
    On Error GoTo ErrH
    Call Logger.PushMethod("ZLHelperMain.clsProcess.ProcessesByModuleName", strModuleName, lngExcludeProcID)
    strModuleName = UCase(strModuleName)
    blnFullPath = InStr(strModuleName, "\") > 0
    If blnFullPath Then
        strFileName = Mid(strModuleName, InStrRev(strModuleName, "\") + 1)
        strFilePath = Mid(strModuleName, 1, InStrRev(strModuleName, "\") - 1)
    Else
        strFileName = strModuleName
        strFilePath = ""
    End If
    
    lngSnapShot = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0)
    If lngSnapShot = 0 Then
        Logger.Error "CreateToolhelp32Snapshot", "Error", GetLastDllErr(Err.LastDllError)
    Else
        uProcess.dwSize = LenB(uProcess)
        If (Process32First(lngSnapShot, uProcess)) = 0 Then
            If Err.LastDllError <> ERROR_NO_MORE_FILES Then
                Logger.Error "Process32First", "Error", GetLastDllErr(Err.LastDllError)
            Else
                Logger.DebugEx "Process32First", "Error", "ERROR_NO_MORE_FILES"
            End If
        Else
            Do
                blnDo = False
                If lngExcludeProcID <> uProcess.th32ProcessID Then
                    If blnCurrentSession Then
                        If Environment.SessionID = Environment.SessionID(uProcess.th32ProcessID) Then
                            blnDo = True
                        End If
                    Else
                        blnDo = True
                    End If
                End If
                If blnDo Then
                    strFindName = uProcess.sExeFile
                    strFindName = UCase(TruncZero(strFindName))
                    If strFindName = strFileName Then
                        If blnFullPath Then
                            If UCase(Environment.StartExePath(uProcess.th32ProcessID)) = strModuleName Then
                                ReDim Preserve arrProcess(UboundEx(arrProcess) + 1)
                                arrProcess(UboundEx(arrProcess)) = uProcess.th32ProcessID
                            End If
                        Else
                            ReDim Preserve arrProcess(UboundEx(arrProcess) + 1)
                            arrProcess(UboundEx(arrProcess)) = uProcess.th32ProcessID
                        End If
                    Else
                        lngMdlSnapShot = CreateToolhelp32Snapshot(TH32CS_SNAPMODULE, uProcess.th32ProcessID)
                        If lngMdlSnapShot = 0 Then
                            Logger.Error "CreateToolhelp32Snapshot(Module)", "Error", GetLastDllErr(Err.LastDllError)
                        Else
                            uMdlInfor.dwSize = LenB(uMdlInfor)
                            If Module32First(lngMdlSnapShot, uMdlInfor) = 0 Then
                                If Err.LastDllError <> ERROR_NO_MORE_FILES Then
                                    Logger.Error "Module32First", "Error", GetLastDllErr(Err.LastDllError)
                                Else
                                    Logger.DebugEx "Module32First", "Error", "ERROR_NO_MORE_FILES"
                                End If
                            Else
                                Do
                                    strFindModule = uMdlInfor.szModule
                                    '开头部门是一个Chr(0)
                                    strModulePath = uMdlInfor.szExePath
                                    strFindModule = UCase(TruncZero(strFindModule))
                                    strModulePath = UCase(TrimEx(strModulePath, Chr(0)))
                                    If strFindModule = strFileName Then
                                        If blnFullPath Then
                                            If strModulePath = strModuleName Then
                                                ReDim Preserve arrProcess(UboundEx(arrProcess) + 1)
                                                arrProcess(UboundEx(arrProcess)) = uProcess.th32ProcessID
                                                Exit Do
                                            End If
                                        Else
                                            ReDim Preserve arrProcess(UboundEx(arrProcess) + 1)
                                            arrProcess(UboundEx(arrProcess)) = uProcess.th32ProcessID
                                            Exit Do
                                        End If
                                    End If
                                    lngRet = Module32Next(lngMdlSnapShot, uMdlInfor)
                                    If lngRet = 0 Then
                                        If Err.LastDllError <> ERROR_NO_MORE_FILES Then
                                            Logger.Error "Module32Next", "Error", GetLastDllErr(Err.LastDllError)
                                        Else
                                            Logger.DebugEx "Module32Next", "Error", "ERROR_NO_MORE_FILES"
                                        End If
                                    End If
                                Loop Until (lngRet = 0)
                            End If
                            If CloseHandle(lngMdlSnapShot) = 0 Then
                                Logger.Error "CloseHandle(ModuleSnapShot)", "Error", GetLastDllErr(Err.LastDllError)
                            Else
                                lngMdlSnapShot = 0
                            End If
                        End If
                    End If
                End If
                lngRet = Process32Next(lngSnapShot, uProcess)
                If lngRet = 0 Then
                    If Err.LastDllError <> ERROR_NO_MORE_FILES Then
                        Logger.Error "Process32Next", "Error", GetLastDllErr(Err.LastDllError)
                    Else
                        Logger.DebugEx "Process32Next", "Error", "ERROR_NO_MORE_FILES"
                    End If
                End If
            Loop Until (lngRet = 0)
        End If
        
        If CloseHandle(lngSnapShot) = 0 Then
            Logger.Error "CloseHandle", "Error", GetLastDllErr(Err.LastDllError)
        Else
            lngSnapShot = 0
        End If
    End If
    ProcessesByModuleName = arrProcess
    Call Logger.PopMethod("ZLHelperMain.clsProcess.ProcessesByModuleName", ProcessesByModuleName)
    Exit Function
ErrH:
    If Logger.ErrCenter("ZLHelperMain.clsProcess.ProcessesByModuleName") = 1 Then
        Resume
    End If
    If lngSnapShot <> 0 Then
        If CloseHandle(lngSnapShot) = 0 Then
            Logger.Error "CloseHandle", "Error", GetLastDllErr(Err.LastDllError)
        End If
    End If
    If lngMdlSnapShot <> 0 Then
        If CloseHandle(lngMdlSnapShot) = 0 Then
            Logger.Error "CloseHandle(ModuleSnapShot)", "Error", GetLastDllErr(Err.LastDllError)
        End If
    End If
    Call Logger.PopMethod("ZLHelperMain.clsProcess.ProcessesByModuleName")
End Function


'@方法    ProcessTerminate
'   杀掉其他进程
'@返回值  Boolean
'
'@参数:
'lngProcessID Long In
'   进程ID,为0为当前进程
'@备注
'
Public Function ProcessTerminate(ByVal lngProcessID As Long) As Boolean
    Dim lngProcess      As Long, lngRet         As Long
    
    On Error GoTo ErrH
    Call Logger.PushMethod("ZLHelperMain.clsProcess.ProcessTerminate", lngProcessID)
    If lngProcessID = 0 Then
        Call ExitProcess(1)
    Else
        lngProcess = OpenProcess(SYNCHRONIZE Or PROCESS_TERMINATE, 0, lngProcessID)
        If lngProcess = 0 Then
            Logger.Error "OpenProcess", "Error", GetLastDllErr(Err.LastDllError)
        Else
            Call TerminateProcess(lngProcess, 1)
            lngRet = WaitForSingleObject(lngProcess, 5000)
            Select Case lngRet
                Case WAIT_ABANDONED
                    Logger.Error "WaitForSingleObject", "Error", "WAIT_ABANDONED"
                Case WAIT_OBJECT_0
                    Logger.DebugEx "WaitForSingleObject", "Error", "WAIT_OBJECT_0"
                    ProcessTerminate = True
                Case WAIT_TIMEOUT
                    Logger.Error "WaitForSingleObject", "Error", "WAIT_TIMEOUT"
                Case WAIT_FAILED
                    Logger.Error "WaitForSingleObject", "Error", GetLastDllErr(Err.LastDllError)
            End Select
            If CloseHandle(lngProcess) = 0 Then
                Logger.Error "CloseHandle", "Error", GetLastDllErr(Err.LastDllError)
            Else
                lngProcess = 0
            End If
        End If
    End If
    Call Logger.PopMethod("ZLHelperMain.clsProcess.ProcessTerminate", ProcessTerminate)
    Exit Function
ErrH:
    If Logger.ErrCenter("ZLHelperMain.clsProcess.ProcessTerminate") = 1 Then
        Resume
    End If
    If lngProcess <> 0 Then
        If CloseHandle(lngProcess) = 0 Then
            Logger.Error "CloseHandle", "Error", GetLastDllErr(Err.LastDllError)
        End If
    End If
    Call Logger.PopMethod("ZLHelperMain.clsProcess.ProcessTerminate")
End Function


'---------------------------------------------------------------------------
'                4、私有方法
'---------------------------------------------------------------------------

'---------------------------------------------------------------------------
'                5、对象方法与事件
'---------------------------------------------------------------------------



