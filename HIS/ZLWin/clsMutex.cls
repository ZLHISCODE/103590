VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMutex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'==================================================================================================
'编写           lshuo
'日期           2019/1/16
'模块           clsMutex
'说明           互斥体
'==================================================================================================
Private Declare Function CreateMutexA Lib "kernel32.dll" (ByVal lpMutexAttributes As Long, ByVal bInitialOwner As Long, ByVal lpName As String) As Long
'@原型
'    HANDLE WINAPI CreateMutex(
'      _In_opt_ LPSECURITY_ATTRIBUTES lpMutexAttributes,
'      _In_     BOOL                  bInitialOwner,
'      _In_opt_ LPCTSTR               lpName
'    );
'@功能
'    创建或打开已命名或未命名互斥对象。要为对象指定访问掩码，请使用CreateMutexEx函数。
'@参数
'lpMutexAttributes _In_opt_
'   指向SECURITY_ATTRIBUTES结构的指针。如果该参数为空，则子进程不能继承句柄。
'   结构的lpSecurityDescriptor成员为新的互斥对象指定一个安全描述符。如果lpMutexAttributes为空，互斥对象将获得默认的安全描述符。互斥锁的默认安全描述符中的acl来自创建器的主令牌或模拟令牌。有关更多信息，请参见同步对象安全性和访问权限。
'bInitialOwner  _In_
'   如果这个值为真，并且调用者创建了互斥对象，那么调用线程将获得互斥对象的初始所有权。否则，调用线程不会获得互斥对象的所有权。要确定调用者是否创建了互斥锁，请参阅返回值部分。
'lpName _In_opt_
'   互斥对象的名称。名称被限制为MAX_PATH长度。名称比较区分大小写。
'   如果lpName匹配现有已命名互斥对象的名称，则此函数请求MUTEX_ALL_ACCESS访问权限。在本例中，bInitialOwner参数被忽略，因为它已经由创建过程设置。如果lpMutexAttributes参数不是NULL，它将确定是否可以继承句柄，但是忽略其安全描述符成员。
'   如果lpName为空，则创建互斥对象时没有名称。
'   如果lpName匹配现有事件、信号量、可等待计时器、作业或文件映射对象的名称，则该函数将失败，GetLastError函数将返回ERROR_INVALID_HANDLE。这是因为这些对象共享相同的名称空间。
'   名称可以具有 "Global\"或  "Local\"前缀，以显式地在全局或会话名称空间中创建对象。名称的其余部分可以包含除反斜杠字符(\)之外的任何字符。有关更多信息，请参见内核对象名称空间。使用终端服务会话实现快速用户切换。内核对象名称必须遵循终端服务的指导原则，以便应用程序能够支持多个用户。
'   对象可以在私有名称空间中创建。有关更多信息，请参见对象名称空间。
'@返回值
'    如果函数成功，返回值是新创建的互斥对象的句柄。
'    如果函数失败，返回值为NULL。要获取扩展的错误信息，请调用GetLastError。
'    如果互斥对象是一个已命名的互斥对象，并且对象在这个函数调用之前已经存在，那么返回值就是现有对象的句柄，GetLastError返回ERROR_ALREADY_EXISTS, bInitialOwner被忽略，调用线程没有被授予所有权。但是，如果调用者具有有限的访问权限，函数将失败，ERROR_ACCESS_DENIED，调用者应该使用OpenMutex函数。
'@备注
'    CreateMutex返回的句柄具有MUTEX_ALL_ACCESS访问权限;它可以用于任何需要互斥对象句柄的函数，前提是调用者已被授予访问权。如果互斥对象是由一个服务或一个模拟不同用户的线程创建的，您可以在创建互斥对象时向它应用一个安全描述符，或者通过更改其默认DACL来更改创建过程的默认安全描述符。有关更多信息，请参见同步对象安全性和访问权限。
'    如果使用指定的互斥锁将应用程序限制为单个实例，恶意用户可以在您之前创建此互斥锁并阻止应用程序启动。为了防止这种情况，创建一个随机命名的互斥体，并存储这个名称，以便只有授权用户才能获得它。或者，您可以为此目的使用文件。要将应用程序限制为每个用户一个实例，请在用户的配置文件目录中创建一个锁定的文件。
'    调用进程的任何线程都可以在调用某个等待函数时指定互斥对象句柄。当指定对象的状态被发出信号时，单对象等待函数返回。可以指示多对象等待函数在任何一个或所有指定对象都发出信号时返回。当等待函数返回时，释放等待线程继续执行。
'    互斥对象的状态在它不属于任何线程时发出信号。创建线程可以使用bInitialOwner标志请求互斥对象的立即所有权。否则，线程必须使用一个等待函数来请求所有权。当互斥锁的状态有信号时，一个等待线程被授予所有权，互斥锁的状态变为无信号状态，等待函数返回。任何时候，只有一个线程可以拥有互斥量。拥有线程使用ReleaseMutex函数来释放其所有权。
'    拥有互斥锁的线程可以在重复的等待函数调用中指定相同的互斥锁，而不会阻塞它的执行。通常，您不会重复等待同一个互斥锁，但是这种机制可以防止线程在等待它已经拥有的互斥锁时死锁自己。但是，为了释放它的所有权，线程必须每次调用ReleaseMutex一次，每次互斥量满足一个等待。
'    两个或多个进程可以调用CreateMutex来创建相同的命名互斥量。第一个进程实际创建互斥锁，具有足够访问权限的后续进程只需打开现有互斥锁的句柄。这使得多个进程能够获得相同互斥量的句柄，同时减轻了用户确保首先启动创建进程的责任。使用此技术时，应该将bInitialOwner标志设置为FALSE;否则，很难确定哪个流程具有初始所有权。
'    多个进程可以具有相同互斥对象的句柄，从而支持使用该对象进行进程间同步。以下对象共享机制可用:
'       如果CreateMutex的lpMutexAttributes参数启用了继承，那么由CreateProcess函数创建的子进程可以继承互斥对象的句柄。这种机制既适用于已命名互斥对象，也适用于未命名互斥对象。
'       进程可以在调用DuplicateHandle函数时指定互斥对象的句柄，以创建可由另一个进程使用的重复句柄。这种机制既适用于已命名互斥对象，也适用于未命名互斥对象。
'       进程可以在调用OpenMutex或CreateMutex函数时指定一个已命名的互斥对象，以检索互斥对象的句柄。
'    使用closeHandle函数关闭句柄。当进程终止时，系统自动关闭句柄。当互斥对象的最后一个句柄被关闭时，它将被销毁。
'@要求
'    Minimum supported client       Windows XP [desktop apps | UWP apps]
'    Minimum supported server       Windows Server 2003 [desktop apps | UWP apps]
'    Header                         WinBase.h on Windows XP, Windows Server 2003, Windows Vista, Windows 7, Windows Server 2008 and Windows Server 2008 R2 (include Windows.h);Synchapi.h on Windows 8 and Windows Server 2012
'    Library                        kernel32.lib
'    dll                            kernel32.dll
'    Unicode and ANSI names         CreateMutexW (Unicode) And CreateMutexA(ANSI)
Private Declare Function CreateMutexExA Lib "kernel32.dll" (ByVal lpMutexAttributes As Long, ByVal lpName As String, ByVal dwFlags As Long, ByVal dwDesiredAccess As Long) As Long
'@原型
'    HANDLE WINAPI CreateMutexEx(
'      _In_opt_ LPSECURITY_ATTRIBUTES lpMutexAttributes,
'      _In_opt_ LPCTSTR               lpName,
'      _In_     DWORD                 dwFlags,
'      _In_     DWORD                 dwDesiredAccess
'    );
'@功能
'    创建或打开已命名或未命名互斥对象，并返回该对象的句柄。
'@参数
'lpMutexAttributes _In_opt_
'    指向SECURITY_ATTRIBUTES结构的指针。如果该参数为NULL，则子进程不能继承互斥对象句柄。
'    结构的lpSecurityDescriptor成员为新的互斥对象指定一个安全描述符。如果lpMutexAttributes为空，互斥对象将获得默认的安全描述符。互斥锁的默认安全描述符中的acl来自创建器的主令牌或模拟令牌。有关更多信息，请参见同步对象安全性和访问权限。
'lpName _In_opt_
'   互斥对象的名称。名称被限制为MAX_PATH长度。名称比较区分大小写。
'   如果lpName匹配现有已命名互斥对象的名称，则此函数请求MUTEX_ALL_ACCESS访问权限。在本例中，bInitialOwner参数被忽略，因为它已经由创建过程设置。如果lpMutexAttributes参数不是NULL，它将确定是否可以继承句柄，但是忽略其安全描述符成员。
'   如果lpName为空，则创建互斥对象时没有名称。
'   如果lpName匹配现有事件、信号量、可等待计时器、作业或文件映射对象的名称，则该函数将失败，GetLastError函数将返回ERROR_INVALID_HANDLE。这是因为这些对象共享相同的名称空间。
'   名称可以具有 "Global\"或  "Local\"前缀，以显式地在全局或会话名称空间中创建对象。名称的其余部分可以包含除反斜杠字符(\)之外的任何字符。有关更多信息，请参见内核对象名称空间。使用终端服务会话实现快速用户切换。内核对象名称必须遵循终端服务的指导原则，以便应用程序能够支持多个用户。
'   对象可以在私有名称空间中创建。有关更多信息，请参见对象名称空间。dwFlags [在]
'dwFlags _In_
'    这个参数可以是0，也可以是以下值。
Private Const CREATE_MUTEX_INITIAL_OWNER        As Long = &H1
'    对象创建者是互斥对象的初始所有者?
'dwDesiredAccess _In_
'    互斥对象的访问掩码。有关访问权限列表，请参见同步对象安全性和访问权限。
'@返回值
'    如果函数成功，返回值是新创建的互斥对象的句柄。
'    如果函数失败，返回值为NULL。要获取扩展的错误信息，请调用GetLastError。
'    如果互斥对象是一个已命名的互斥对象，并且对象在这个函数调用之前已经存在，那么返回值就是现有对象的句柄，GetLastError返回ERROR_ALREADY_EXISTS, bInitialOwner被忽略，调用线程没有被授予所有权。但是，如果调用者具有有限的访问权限，函数将失败，ERROR_ACCESS_DENIED，调用者应该使用OpenMutex函数。
'@备注
'    CreateMutex返回的句柄具有MUTEX_ALL_ACCESS访问权限;它可以用于任何需要互斥对象句柄的函数，前提是调用者已被授予访问权。如果互斥对象是由一个服务或一个模拟不同用户的线程创建的，您可以在创建互斥对象时向它应用一个安全描述符，或者通过更改其默认DACL来更改创建过程的默认安全描述符。有关更多信息，请参见同步对象安全性和访问权限。
'    如果使用指定的互斥锁将应用程序限制为单个实例，恶意用户可以在您之前创建此互斥锁并阻止应用程序启动。为了防止这种情况，创建一个随机命名的互斥体，并存储这个名称，以便只有授权用户才能获得它。或者，您可以为此目的使用文件。要将应用程序限制为每个用户一个实例，请在用户的配置文件目录中创建一个锁定的文件。
'    调用进程的任何线程都可以在调用某个等待函数时指定互斥对象句柄。当指定对象的状态被发出信号时，单对象等待函数返回。可以指示多对象等待函数在任何一个或所有指定对象都发出信号时返回。当等待函数返回时，释放等待线程继续执行。
'    互斥对象的状态在它不属于任何线程时发出信号。创建线程可以使用bInitialOwner标志请求互斥对象的立即所有权。否则，线程必须使用一个等待函数来请求所有权。当互斥锁的状态有信号时，一个等待线程被授予所有权，互斥锁的状态变为无信号状态，等待函数返回。任何时候，只有一个线程可以拥有互斥量。拥有线程使用ReleaseMutex函数来释放其所有权。
'    拥有互斥锁的线程可以在重复的等待函数调用中指定相同的互斥锁，而不会阻塞它的执行。通常，您不会重复等待同一个互斥锁，但是这种机制可以防止线程在等待它已经拥有的互斥锁时死锁自己。但是，为了释放它的所有权，线程必须每次调用ReleaseMutex一次，每次互斥量满足一个等待。
'    两个或多个进程可以调用CreateMutex来创建相同的命名互斥量。第一个进程实际创建互斥锁，具有足够访问权限的后续进程只需打开现有互斥锁的句柄。这使得多个进程能够获得相同互斥量的句柄，同时减轻了用户确保首先启动创建进程的责任。使用此技术时，不应该使用CREATE_MUTEX_INITIAL_OWNER标志;否则，很难确定哪个流程具有初始所有权。
'    多个进程可以具有相同互斥对象的句柄，从而支持使用该对象进行进程间同步。以下对象共享机制可用:
'       如果CreateMutex的lpMutexAttributes参数启用了继承，那么由CreateProcess函数创建的子进程可以继承互斥对象的句柄。这种机制既适用于已命名互斥对象，也适用于未命名互斥对象。
'       进程可以在调用DuplicateHandle函数时指定互斥对象的句柄，以创建可由另一个进程使用的重复句柄。这种机制既适用于已命名互斥对象，也适用于未命名互斥对象。
'       进程可以在调用OpenMutex或CreateMutex函数时指定一个已命名的互斥对象，以检索互斥对象的句柄。
'    使用closeHandle函数关闭句柄。当进程终止时，系统自动关闭句柄。当互斥对象的最后一个句柄被关闭时，它将被销毁。
'@要求
'    Minimum supported client       Windows XP [desktop apps | UWP apps]
'    Minimum supported server       Windows Server 2003 [desktop apps | UWP apps]
'    Header                         WinBase.h on Windows XP, Windows Server 2003, Windows Vista, Windows 7, Windows Server 2008 and Windows Server 2008 R2 (include Windows.h);Synchapi.h on Windows 8 and Windows Server 2012
'    Library                        kernel32.lib
'    dll                            kernel32.dll
'    Unicode and ANSI names         CreateMutexExW (Unicode) And CreateMutexExA(ANSI)
Private Declare Function OpenMutexA Lib "kernel32.dll" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal lpName As String) As Long
'@原型
'    HANDLE WINAPI OpenMutex(
'      _In_ DWORD   dwDesiredAccess,
'      _In_ BOOL    bInheritHandle,
'      _In_ LPCTSTR lpName
'    );
'@功能
'    打开现有的命名互斥对象。
'@参数
'dwDesiredAccess _In_
'    对互斥对象的访问。使用互斥锁只需要同步访问权;要更改互斥锁的安全性，请指定MUTEX_ALL_ACCESS。如果指定对象的安全描述符不允许调用进程的请求访问，则该函数将失败。有关访问权限列表，请参见同步对象安全性和访问权限。
'bInheritHandle _In_
'    如果该值为真，则由该进程创建的进程将继承该句柄。否则，进程不会继承此句柄。
'lpName _In_opt_
'    要打开的互斥对象的名称?名称比较区分大小写?
'    此函数可以打开私有名称空间中的对象。有关更多信息，请参见对象名称空间。
'    终端服务:名称可以具有 "Global\"或  "Local\"前缀，以显式地在全局或会话名称空间中创建对象。名称的其余部分可以包含除反斜杠字符(\)之外的任何字符。有关更多信息，请参见内核对象名称空间。
'    注意，快速用户切换是使用终端服务会话实现的。第一个登录的用户使用会话0，下一个登录的用户使用会话1，依此类推。使用终端服务会话实现快速用户切换。内核对象名称必须遵循终端服务的指导原则，以便应用程序能够支持多个用户。
'    对象可以在私有名称空间中创建。有关更多信息，请参见对象名称空间。dwFlags [在]
'@返回值
'    如果函数成功，返回值是互斥对象的句柄。
'    如果函数失败，返回值为NULL。要获取扩展的错误信息，请调用GetLastError。
'    如果指定的互斥对象不存在，则函数失败，GetLastError返回ERROR_FILE_NOT_FOUND。
'@备注
'    OpenMutex函数允许多个进程打开同一个mutex对象的句柄。只有当某个进程已经使用CreateMutex函数创建了互斥对象时，该函数才会成功。调用进程可以在任何需要互斥对象句柄的函数中使用返回的句柄，例如wait函数，但受dwDesiredAccess参数中指定的访问限制。
'    可以使用DuplicateHandle函数复制句柄。使用closeHandle函数关闭句柄。当进程终止时，系统自动关闭句柄。当互斥对象的最后一个句柄被关闭时，它将被销毁。
'    如果多线程应用程序必须反复创建、打开和关闭一个指定的互斥对象，则可能会出现竞态条件。在这种情况下，最好使用CreateMutex而不是OpenMutex，因为CreateMutex在互斥锁存在时打开它，在不存在时创建它。
'@要求
'    Minimum supported client       Windows XP [desktop apps | UWP apps]
'    Minimum supported server       Windows Server 2003 [desktop apps | UWP apps]
'    Header                         WinBase.h (include Windows.h)
'    Library                        kernel32.lib
'    dll                            kernel32.dll
'    Unicode and ANSI names         OpenMutexW (Unicode) and OpenMutexA (ANSI)
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
Private Declare Function ReleaseMutex Lib "kernel32.dll" (ByVal mlnghMutex As Long) As Long
'@原型
'    BOOL WINAPI ReleaseMutex(
'      _In_ HANDLE hMutex
'    );
'@功能
'    释放指定互斥对象的所有权。
'@参数
'hMutex _In_
'    互斥对象的句柄?CreateMutex或OpenMutex函数返回这个句柄?
'@返回值
'    如果函数成功，返回值为非零。
'    如果函数失败，返回值为零。要获取扩展的错误信息，请调用GetLastError。
'@备注
'    如果调用线程不拥有互斥对象，则ReleaseMutex函数将失败。
'    线程通过将bInitialOwner参数设置为TRUE来创建互斥对象，或者在调用某个等待函数时指定它的句柄，从而获得互斥对象的所有权。当线程不再需要拥有互斥对象时，它调用ReleaseMutex函数，以便另一个线程可以获得所有权。
'    线程可以在调用某个等待函数时指定它已经拥有的互斥量，而不会阻塞它的执行。这可以防止线程在等待它已经拥有的互斥对象时死锁自己。但是，为了释放它的所有权，线程必须在每次获得所有权时调用ReleaseMutex一次(通过CreateMutex或等待函数)。
'@要求
'    Minimum supported client       Windows XP [desktop apps | UWP apps]
'    Minimum supported server       Windows Server 2003 [desktop apps | UWP apps]
'    Header                         WinBase.h on Windows XP, Windows Server 2003, Windows Vista, Windows 7, Windows Server 2008 and Windows Server 2008 R2 (include Windows.h);Synchapi.h on Windows 8 and Windows Server 2012
'    Library                        kernel32.lib
'    dll                            kernel32.dll


Private Const ERROR_SHARING_VIOLATION         As Long = &H20
Private Const ERROR_INSUFFICIENT_BUFFER       As Long = 122
Private Const ERROR_MORE_DATA                 As Long = 234
Private Const ERROR_NO_MORE_ITEMS             As Long = &H103
Private Const ERROR_KEY_NOT_FOUND             As Long = &H2
Private Const ERROR_ALREADY_EXISTS            As Long = 183&
Private Const ERROR_ACCESS_DENIED             As Long = 5&
Private Const ERROR_SUCCESS                   As Long = 0&

Private mlnghMutex              As Long
Private Const RLL_AllLog        As Long = RunLogLevel.RLL_AllLog + 3
Private Const RLL_LogInfo       As Long = RunLogLevel.RLL_AllLog + 3
'--------------------------------------------------------------------------------------------------
'方法           CheckMutex
'功能           检查互斥体是否存在。不存在则创建
'返回值         Boolean
'入参列表:
'参数名         类型                    说明
'strMutexName   String                  互斥体名称
'-------------------------------------------------------------------------------------------------
Public Function CheckMutex(ByVal strMutexName As String) As Boolean
    Dim lngRet As Long, lnghM As Long
    On Error GoTo ErrH
    Call gobjLog.PushMethod(RLL_AllLog, "ZlWin.clsMutex.CheckMutex", "MUTEX_" & strMutexName)
    
    If mlnghMutex = 0 Then
        mlnghMutex = CreateMutexA(ByVal 0&, 1, "MUTEX_" & strMutexName)
        If (Err.LastDllError = ERROR_ALREADY_EXISTS) Or (Err.LastDllError = ERROR_ACCESS_DENIED) Then
            CheckMutex = True
            Call CloseMutex
        Else
            CheckMutex = False
        End If
    Else
        CloseMutex
    End If
    Call gobjLog.PopMethod(RLL_AllLog, "ZlWin.clsMutex.CheckMutex", CheckMutex)
    Exit Function
ErrH:
    If gobjLog.ErrCenter(RLL_RunError, "ZlWin.clsMutex.CheckMutex") = 1 Then
        Resume
    End If
End Function
'--------------------------------------------------------------------------------------------------
'方法           CloseMutex
'功能           关闭互斥体句柄
'返回值         Boolean
'入参列表:
'参数名         类型                    说明
'
'-------------------------------------------------------------------------------------------------
Private Function CloseMutex() As Boolean
    Dim lngRet As Long
 
    On Error GoTo ErrH
    Call gobjLog.PushMethod(RLL_AllLog, "ZlWin.clsMutex.CloseMutex")
    If (mlnghMutex <> 0) Then
        lngRet = ReleaseMutex(mlnghMutex)
        lngRet = CloseHandle(mlnghMutex)
    End If
    If lngRet = 1 Then
        CloseMutex = False
        mlnghMutex = 0
    Else
        CloseMutex = True
    End If
    Call gobjLog.PopMethod(RLL_AllLog, "ZlWin.clsMutex.CloseMutex", CloseMutex)
    Exit Function
ErrH:
    If gobjLog.ErrCenter(RLL_RunError, "ZlWin.clsMutex.CloseMutex") = 1 Then
        Resume
    End If
End Function

Private Sub Class_Terminate()
    Call CloseMutex
End Sub
