VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMemoryShare"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@模块 clsMemoryShare-2019/6/26
'@编写 lshuo
'@功能
'   共享内存处理类
'@引用
'
'@备注
'
Option Explicit
'---------------------------------------------------------------------------
'                0、API和常量声明
'---------------------------------------------------------------------------
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
Private Declare Function CreateFileMapping Lib "kernel32" Alias "CreateFileMappingA" (ByVal hFile As Long, lpFileMappigAttributes As Long, ByVal flProtect As Long, ByVal dwMaximumSizeHigh As Long, ByVal dwMaximumSizeLow As Long, ByVal lpName As String) As Long
'@功能
'    创建或打开指定文件的命名或未命名的文件映射对象。指定物理内存 NUMA 节点，请参见 CreateFileMappingNuma。
'@原型
'    HANDLE WINAPI CreateFileMapping(
'      _In_     HANDLE                hFile,
'      _In_opt_ LPSECURITY_ATTRIBUTES lpAttributes,
'      _In_     DWORD                 flProtect,
'      _In_     DWORD                 dwMaximumSizeHigh,
'      _In_     DWORD                 dwMaximumSizeLow,
'      _In_opt_ LPCTSTR               lpName
'    );
'@参数
'    hFile
'       用于创建文件映射对象的文件的句柄?
'       必须使用与flProtect参数指定的保护标志兼容的访问权限打开文件。这不是必需的，但建议打开要映射的文件以进行独占访问。有关更多信息，请参见文件安全性和访问权限。
'       如果hFile是INVALID_HANDLE_VALUE，则调用进程还必须在dwMaximumSizeHigh和dwMaximumSizeLow参数中指定文件映射对象的大小。在此场景中，CreateFileMapping创建指定大小的文件映射对象，该对象由系统分页文件支持，而不是由文件系统中的文件支持。
'    lpAttributes(,可选)
'       指向SECURITY_ATTRIBUTES结构的指针，该结构决定子进程是否可以继承返回的句柄。SECURITY_ATTRIBUTES结构的lpSecurityDescriptor成员为新文件映射对象指定了一个安全描述符。
'       如果lpAttributes为空，则不能继承句柄，文件映射对象将获得默认的安全描述符。文件映射对象的默认安全描述符中的访问控制列表(ACL)来自创建器的主令牌或模拟令牌。有关更多信息，请参见文件映射安全性和访问权限。
'    flProtect
'        指定文件映射对象的页保护?对象的所有映射视图都必须与此保护兼容?
'        此参数可以是下列值中的某一个：
'        值 含义
Private Const PAGE_EXECUTE_READ         As Long = &H20
'           允许视图被映射为只读、写时复制或执行访问
'           由 hFile 参数指定的文件句柄必须以 GENERIC_READ 和 GENERIC_EXECUTE 访问权限创建
'           Windows Server 2003 和 Windows XP：Windows XP SP2 和 Windows Server 2003 SP1 之前此值不可用
Private Const PAGE_EXECUTE_READWRITE    As Long = &H40
'           允许视图被映射为只读、写时复制，读/写或者执行访问
'           由 hFile 参数指定的文件句柄必须以 GENERIC_READ，GENERIC_WRITE 和 GENERIC_EXECUTE 访问权限创建
'           Windows Server 2003 和 Windows XP：Windows XP SP2 和 Windows Server 2003 SP1 之前此值不可用
Private Const PAGE_EXECUTE_WRITECOPY    As Long = &H80
'           允许视图被映射为只读、写时复制或执行访问。这个数值等同于PAGE_EXECUTE_READ
'           由 hFile 参数指定的文件句柄必须以 GENERIC_READ 和 GENERIC_EXECUTE访 问权限创建
'           Windows Vista：Windows Vista SP1 之前不支持此值
'           Windows Server 2003 和 Windows XP：不支持此值
Private Const PAGE_READONLY             As Long = &H2
'           允许视图被映射为只读或写时复制访问。尝试写入特定区域会导致非法访问异常
'           由 hFile 参数指定的文件句柄必须以 GENERIC_READ 访问权限创建
Private Const PAGE_READWRITE            As Long = &H4
'           允许视图被映射为只读，写时复制或读/写访问
'           由 hFile 参数指定的文件句柄必须以 GENERIC_READ 和 GENERIC_WRITE 访问权限创建
Private Const PAGE_WRITECOPY            As Long = &H8
'           允许视图被映射为只读、写时复制访问。这个数值等同于 PAGE_READONLY
'           由 hFile 参数指定的文件句柄必须以 GENERIC_READ 访问权限创建
'       应用程序可以为文件映射对象指定一个或多个下列属性，并将它们与前面的一个页面保护值相结合：
Private Const SEC_COMMIT                As Long = &H8000000
'           如果文件映射对象是被操作系统的页文件所支持的（hFile 参数是 INVALID_HANDLE_VALUE），表明当一个文件的视图被映射到进程的地址空间时，页的整个范围将被提交而非保留。系统必须有足够的可提交页面来保存整个映射。否则，CreatFileMapping 调用失败
'           此属性不影响可执行映像文件或数据文件支持的文件映射对象
'           SEC_COMMIT 不能与 SEC_RESERVE 同时使用
'           如果没有指定属性，那么默认会假定为 SEC_COMMIT
Private Const SEC_IMAGE                 As Long = &H1000000
'           指定的 hFile 参数指定的文件是可执行映像文件
'           该 SEC_IMAGE 属性必须与页面保护值，如 PAGE_READONLY 组合。然而，此页面保护值不影响可执行映像文件的视图。可执行映像文件的视图的页面保护是由可执行文件本身决定的
'           和 SEC_IMAGE 属性一同使用的属性都无效
Private Const SEC_IMAGE_NO_EXECUTE      As Long = &H11000000
'           指定的 hFile 参数指定的文件是不会执行的可执行映像文件并且加载的图像文件没有进行强制完整性检查。此外，映射一个以 SEC_IMAGE_NO_EXECUTE 属性创建的文件映射对象不会使用 PsSetLoadImageNotifyRoutine 内核 API 注册驱动程序回调函数
'           SEC_IMAGE_NO_EXECUTE 属性必须结合 PAGE_READONLY 页面保护值。和 SEC_IMAGE_NO_EXECUTE 一同使用的其他属性都是无效的
'           Windows Server 2008 R2, Windows 7, Windows Server 2008, Windows Vista, Windows Server 2003, 和 Windows XP：在Windows Server 2012 和 Windows 8 以前不支持此值
Private Const SEC_LARGE_PAGES           As Long = &H80000000
'           被操作系统页文件所支持的文件映射对象所使用的大页面被允许（hFile 参数是INVALID_HANDLE_VALUE）。这个属性不支持由可执行镜像文件或数据文件所支持的文件映射对象（hFile 参数是一个可执行镜像文件或数据文件的句柄）
'           文件映射对象的最大大小必须是 GetLargePageMinimum 函数返回的一个较大的页面最小大小的倍数。如果不是，那么 CreateFileMapping 调用失败。当映射一个以 SEC_LARGE_PAGES 创建的文件映射对象的视图时，基地址和视图的大小必须是最小的大页面大小的倍数
'           SEC_LARGE_PAGES 要求 SeLockMemoryPrivilege 特权被调用者启用
'           如果指定了 SEC_LARGE_PAGES，也必须指定 SEC_COMMIT
'           Windows Server 2003：Windows Server 2003 SP1 之前不支持此值
'           Windows XP：不支持此值
Private Const SEC_NOCACHE               As Long = &H10000000
'           设置所有页面不能被缓存
'           应用程序不应使用此属性，除非设备的明确要求。对以 SEC_NOCACHE 映射的内存使用互锁函数会导致 EXCEPTION_ILLEGAL_INSTRUCTION 异常
'           SEC_NOCACHE 需要设置 SEC_RESERVE 或 SEC_COMMIT 属性
Private Const SEC_RESERVE               As Long = &H4000000
'           如果文件映射对象是被操作系统的页文件所支持的（hFile 参数是 INVALID_HANDLE_VALUE），表明当一个文件的视图被映射到进程的地址空间时，页的整个范围将被保留用作以后的使用而不是提交
'           保留页可以在随后的 VirtualAlloc 函数调用中被提交。当这些页面被提交后，不能通过 VirtualFree 函数释放或者重新提交
'           SEC_COMMIT 不能与 SEC_RESERVE 结合使用
Private Const SEC_WRITECOMBINE          As Long = &H40000000
'           将所有页面设置为写聚合
'           应用程序不应使用此属性，除非设备的明确要求。对以 SEC_WRITECOMBINE 映射的内存使用互锁函数会导致 EXCEPTION_ILLEGAL_INSTRUCTION 异常
'           SEC_WRITECOMBINE 需要设置 SEC_RESERVE 或 SEC_COMMIT 属性
'           Windows Server 2003 和 Windows XP：在 Windows Vista 之前不支持此标志
'    dwMaximumSizeHigh
'            文件映射对象最大大小的高 32 位
'    dwMaximumSizeLow
'           文件映射对象最大大小的低32位
'           如这个参数和 dwMaximumSizeHigh 都是零，那么文件映射对象的最大大小等于 hFile 指定文件的实际大小
'           映射一个大小为 0 的文件将会引发出错码为 ERROR_FILE_INVALID 的错误。应用程序应该检测大小为 0 的文件，并拒绝这些文件
'    lpName
'           文件映射对象的名称
'           如果这个参数与一个已经存在的文件映射对象的名称相同，那么该函数请求以 flProtect 指定的保护属性访问该对象
'           如果这个参数为 NULL，那么创建的这个文件映射对象将没有名称
'           如果 lpName 与一个现有的事件,信号量,互斥锁,可等待定时器,或工作对象同名，那么函数调用失败，并且 GetLassError 函数返回 ERROR_INVALID_HANDLE。因为这些对象共享相同的命名空间
'           该名称可以有一个 "Global\" 或 "Local\" 前缀来在全局或会话命名空间中显示地创建对象。其余的名称可以包含除反斜杠字符（\）以外的任何字符。在全局命名空间中从一个会话而不是会话0中创建一个文件映射对象需要 SeCreateGlobalPrivilege 特权。更多信息，请参考 Kernel Object Namespaces
'           快速用户切换是通过使用终端服务会话实现的。第一个用户以会话 0 登陆，第二个用户以会话 1 登陆等等。内核对象名称必须遵循为终端服务所列出的指导原则，以便应用程序可以支持多个用户
'@返回值:
'    1. 如果函数执行成功，返回值是新创建的文件映射对象的句柄。。
'    2. 如果在函数调用之前，该对象存在，函数将返回现有对象的句柄（并且是对象现有的大小而不是调用函数时指定的大小），并且 GetLastError 函数返回 ERROR_ALREADY_EXISTS
'    3. 如果函数失败，则返回值为 NULL
'    想要获取更多错误有关信息，请调用 GetLastError
'@备注:
'    1. 创建一个文件映射对象后，文件大小必须不超过该文件映射对象的大小；否则，并不是所有的文件内容都可以共享。
'    2. 如果一个应用程序指定一个文件映射对象的大小大于磁盘上的实际命名文件的大小，并且如果页面保护允许写入访问权限（也就是说，flProtect 参数指定了 PAGE_READWRITE 或 PAGE_EXECUTE_READWRITE），那么磁盘上的文件会增加大小到文件映射对象的大小。如果文件被扩展，文件旧的尾端和文件的新的尾端之间的文件的内容不保证为零，这个行为由文件系统来定义。如果磁盘上的文件不能够被扩展，那么 CreateFileMapping 函数调用失败，并且 GetLastError 函数返回 ERROR_DISK_FULL。
'    3. 操作系统分页文件支持的文件映射对象中的页的初始内容为0。
'    4. CreateFileMapping 函数返回的句柄对那个新的文件映射对象具有完全的访问权限，并且能够被任何需要文件映射对象句柄的函数所使用。
'    5. 多进程可以通过使用单个共享文件映射对象或创建单独的文件映射对象来共享同一文件的视图。一个文件映射对象能够被多个进程共享，通过在创建进程时继承句柄，复制句柄或通过名字打开一个文件映射对象。更多信息，请参考 CreateProcess, DuplicateHandle 和 OpenFileMapping 函数。
'    6. 创建一个文件映射对象并没有将视图映射到一个进程地址空间。MapViewOfFile 和 MapViewOfFileEx 函数才映射一个文件的视图到进程地址空间。
'    7. 在一个重要的例外情况下，来自同一文件的任何文件映射对象的文件视图在特定时间内是一致的或相同的。在一个进程中的视图和被不同进程映射的视图才能保证一致性。
'    8. 唯一的例外是与远程文件相联系。虽然 CreateFileMapping 函数能够处理远程文件，但是并不能保证他们的一致性。例如，如果两台电脑都将一个文件映射为可写的，并且同时改变相同的页面，每台电脑只看到他们自己对那个页面所做的修改。所以当数据在磁盘上更新的时候，并没有合并他们的内容。
'    9. 一个映射的文件和一个通过输入和输出（I/O）函数（ReadFile 和 WriteFile）访问的文件并不需要保持一致性。
'    10. 一个文件映射对象的映射视图维持着对象的内部引用，直到所有的引用被释放，文件映射对象才能够被关闭。因此，要想完全关闭一个文件映射对象，应用程序必须通过调用 UnmapViewOfFile 来解除文件映射对象的所有映射视图，并且通过调用 CloseHandle 来关闭文件映射对象的句柄。这些函数能够以任意的顺序被调用。
'    11. 当通过映射的视图修改文件时，最后的修改时间戳可能不能够自动更新。如果需要自动更新，可以使用 SetFileTime 来设置一个时间戳。
'    12. 在全局命名空间中从一个会话而不是会话 0 中来创建一个文件映射对象需要 SeCreateGlobalPrivilege 特权。注意这个特权被限制在创建文件映射对象的时候使用，而不能应用于打开一个已经存在的对象。例如，如果一个服务或者系统在全局命名空间中创建了一个文件映射对象，在任何会话中运行的任何进程都可以访问该文件映射对象，该对象提供了该调用方具有所需的访问权限。
'    13. Windows XP：在前几段中通过 Windows Server 2003 和 Windows XP 介绍了需求描述。
'    14. 使用结构化异常处理来保护写入或读取文件视图的任何代码。更多信息，请参见 Reading and Writing From a File View。
'    15. 为了有可执行权限的映射，应用程序必须调用 CreateFileMapping 函数，并且设置 PAGE_EXECUTE_READWRITE 或 PAGE_EXECUTE_READ。然后调用 MapViewOfFile 函数，并且设置 FILE_MAP_EXECUTE | FILE_MAP_WRITE 或 FILE_MAP_EXECUTE | FILE_MAP_READ。
'    16. Windows Server 2012 中，这个功能是通过以下技术支持。
'    技术 支持
'    Server Message Block (SMB) 3.0 protocol Yes
'    SMB 3.0 Transparent Failover (TFO)  Yes
'    SMB 3.0 with Scale-out File Shares (SO) Yes
'    Cluster Shared Volume File System (CsvFS)   Yes
'    Resilient File System (ReFS)    Yes
'@需求:
'    Minimum supported client    Windows XP [仅桌面应用程序]
'    Minimum supported server    Windows 2003 [仅桌面应用程序]
'    Header  WinBase.h (include Windows.h)
'    Library Kernel32.lib
'    dll Kernel32.dll
'    Unicode and ANSI names  CreateFileMappingW (Unicode) and CreateFileMappingA (ANSI)

Private Const STANDARD_RIGHTS_REQUIRED          As Long = &HF0000
Private Const SECTION_QUERY                     As Long = &H1
Private Const SECTION_MAP_WRITE                 As Long = &H2
Private Const SECTION_MAP_READ                  As Long = &H4
Private Const SECTION_MAP_EXECUTE               As Long = &H8
Private Const SECTION_EXTEND_SIZE               As Long = &H10
Private Const SECTION_MAP_EXECUTE_EXPLICIT      As Long = &H20       'not included in SECTION_ALL_ACCESS
Private Const SECTION_ALL_ACCESS                As Long = STANDARD_RIGHTS_REQUIRED Or SECTION_QUERY Or SECTION_MAP_WRITE Or SECTION_MAP_READ Or SECTION_MAP_EXECUTE Or SECTION_EXTEND_SIZE
Private Declare Function MapViewOfFile Lib "kernel32" (ByVal hFileMappingObject As Long, ByVal dwDesiredAccess As Long, ByVal dwFileOffsetHigh As Long, ByVal dwFileOffsetLow As Long, ByVal dwNumberOfBytesToMap As Long) As Long
'@原型
'    LPVOID MapViewOfFile(
'      HANDLE hFileMappingObject,
'      DWORD  dwDesiredAccess,
'      DWORD  dwFileOffsetHigh,
'      DWORD  dwFileOffsetLow,
'      SIZE_T dwNumberOfBytesToMap
'    );
'@功能
'    将文件映射的视图映射到调用进程的地址空间?
'    要为视图指定建议的基本地址，请使用MapViewOfFileEx函数。但是，不建议这样做。
'@参数
'hFileMappingObject
'    文件映射对象的句柄?CreateFileMapping和OpenFileMapping函数返回这个句柄?
'dwDesiredAccess
'    访问文件映射对象的类型，该类型决定了页面的页面保护。此参数可以是下列值之一，也可以是多个值的位或组合(如果合适)。
Private Const FILE_MAP_ALL_ACCESS               As Long = SECTION_ALL_ACCESS
'    映射文件的读/写视图。文件映射对象必须使用PAGE_READWRITE或PAGE_EXECUTE_READWRITE保护创建。
'    当与MapViewOfFile函数一起使用时，FILE_MAP_ALL_ACCESS等价于FILE_MAP_WRITE。
Private Const FILE_MAP_READ                     As Long = SECTION_MAP_READ
'    映射文件的只读视图?试图写入文件视图会导致访问冲突?
'    必须使用PAGE_READONLY?PAGE_READWRITE?PAGE_EXECUTE_READ或PAGE_EXECUTE_READWRITE保护创建文件映射对象?
Private Const FILE_MAP_WRITE                    As Long = SECTION_MAP_WRITE
'    映射文件的读/写视图。文件映射对象必须使用PAGE_READWRITE或PAGE_EXECUTE_READWRITE保护创建。
'    当与MapViewOfFile一起使用时，(FILE_MAP_WRITE | FILE_MAP_READ)和FILE_MAP_ALL_ACCESS等价于FILE_MAP_WRITE。
'使用位或者，您可以将上面的值与这些值组合起来。
Private Const FILE_MAP_COPY                     As Long = SECTION_QUERY
'    映射文件的copy-on-write视图。必须使用PAGE_READONLY、PAGE_READ_EXECUTE、PAGE_WRITECOPY、PAGE_EXECUTE_WRITECOPY、PAGE_READWRITE或PAGE_EXECUTE_READWRITE保护创建文件映射对象。
'    当进程写入到“写中复制”页时，系统将原始页复制到进程私有的新页。新页面由分页文件支持。新页面的保护从写时复制更改为读/写。
'    当指定了写时复制访问时，系统和进程提交的费用将用于整个视图，因为调用进程可能会写入视图中的每个页面，从而使所有页面都是私有的。新页面的内容永远不会写回原始文件，并且在视图未映射时丢失。
Private Const FILE_MAP_EXECUTE                  As Long = SECTION_MAP_EXECUTE_EXPLICIT
'    文件的可执行视图被映射(映射内存可以作为代码运行)。文件映射对象必须使用PAGE_EXECUTE_READ、PAGE_EXECUTE_WRITECOPY或PAGE_EXECUTE_READWRITE保护创建。
'    Windows Server 2003和Windows XP:这个值从SP2的Windows XP和SP1的Windows Server 2003开始可用。
'Private Const FILE_MAP_LARGE_PAGES              As Long = ?
'    从Windows 10 1703版本开始，这个标志指定视图应该使用大页面支持进行映射。视图的大小必须是GetLargePageMinimum函数报告的大页面大小的倍数，并且文件映射对象必须使用SEC_LARGE_PAGES选项创建。如果为lpBaseAddress提供非空值，那么该值必须是GetLargePageMinimum的倍数。
'Private Const FILE_MAP_TARGETS_INVALID          As Long = ?
'    将映射文件中的所有位置设置为控制流保护(CFG)的无效目标。这个标志类似于PAGE_TARGETS_INVALID。将此标志与执行访问权限FILE_MAP_EXECUTE组合使用。任何对这些页面中位置的间接调用都将无法通过CFG检查，该过程将被终止。分配给可执行页面的默认行为是为CFG标记有效的调用目标。
'    对于使用SEC_IMAGE属性创建的文件映射对象，dwDesiredAccess参数没有影响，应该设置为任何有效值，比如FILE_MAP_READ。
'    有关访问文件映射对象的更多信息，请参见文件映射安全性和访问权限。
'dwFileOffsetHigh
'    视图开始处文件偏移量的高阶DWORD?
'dwFileOffsetLow
'    视图开始处文件偏移量的低阶DWORD。高偏移量和低偏移量的组合必须在文件映射中指定偏移量。它们还必须匹配系统的内存分配粒度。也就是说，偏移量必须是分配粒度的倍数。要获得系统的内存分配粒度，可以使用GetSystemInfo函数，它填充SYSTEM_INFO结构的成员。
'dwNumberOfBytesToMap
'    要映射到视图的文件映射的字节数。所有字节必须在CreateFileMapping指定的最大大小内。如果该参数为0(0)，映射将从指定的偏移量扩展到文件映射的末尾。
'@返回值
'    如果函数成功，返回值就是映射视图的起始地址。
'    如果函数失败，返回值为NULL。要获取扩展的错误信息，请调用GetLastError。
'@备注
'    映射文件使文件的指定部分在调用进程的地址空间中可见?
'    对于大于地址空间的文件，一次只能映射文件数据的一小部分。当第一个视图完成时，您可以解映射它并映射一个新视图。
'    要获得视图的大小，请使用VirtualQuery函数。
'    如果文件(或文件映射对象及其映射文件)的多个视图在指定时间包含相同的数据，那么它们是一致的。如果文件视图派生自由同一文件支持的任何文件映射对象，则会发生这种情况。一个进程可以使用DuplicateHandle函数将文件映射对象句柄复制到另一个进程中，或者另一个进程可以使用OpenFileMapping函数按名称打开文件映射对象。
'    除了一个重要的例外，从由相同文件支持的任何文件映射对象派生的文件视图在特定时间是一致的或相同的。对于流程中的视图和由不同流程映射的视图，保证了一致性。
'    异常与远程文件有关。尽管MapViewOfFile可以处理远程文件，但它并不能使它们保持一致。例如，如果两台计算机都将一个文件映射为可写的，并且都更改了相同的页面，那么每台计算机只能看到自己对该页面的写操作。当磁盘上的数据得到更新时，它不会被合并。
'    文件的映射视图不能保证与ReadFile或WriteFile函数正在访问的文件一致?
'    不要在内存映射文件中存储指针;从文件映射的基础存储偏移量，以便映射可以在任何地址使用。
'    为了防止EXCEPTION_IN_PAGE_ERROR异常，可以使用结构化异常处理来保护除页面文件外对文件的内存映射视图进行读写的任何代码。有关更多信息，请参见从文件视图中读取和写入。
'    当通过映射视图修改文件时，最后一次修改的时间戳可能不会自动更新。如果需要，调用者应该使用SetFileTime设置时间戳。
'    如果文件映射对象由分页文件支持(调用CreateFileMapping时将hFile参数设置为INVALID_HANDLE_VALUE)，那么分页文件必须足够大，以容纳整个映射。如果不是，则MapViewOfFile失败。分页文件支持的文件映射对象中的页的初始内容为0(0)。
'    当创建由分页文件支持的文件映射对象时，调用者可以指定MapViewOfFile应该同时保留和提交页面(SEC_COMMIT)，还是简单地保留页面(SEC_RESERVE)。映射该文件使整个映射的虚拟地址范围对进程中的其他分配不可用。在提交了来自保留范围的页面之后，不能通过调用VirtualFree释放或解除该页面。当视图未映射且文件映射对象关闭时，将释放保留页和提交页。有关详细信息，请参见UnmapViewOfFile和close句柄函数。
'    要拥有具有可执行权限的文件，应用程序必须使用PAGE_EXECUTE_READWRITE或PAGE_EXECUTE_READ调用CreateFileMapping，然后使用FILE_MAP_EXECUTE | FILE_MAP_WRITE或FILE_MAP_EXECUTE | FILE_MAP_READ调用MapViewOfFile。
'    在Windows Server 2012中，此功能由以下技术支持。
'            服务器消息块(SMB) 3.0协议      是的
'            SMB 3.0透明故障转移(TFO)       是的
'            带扩展文件共享(SO)的SMB 3.0    是的
'            集群共享卷文件系统(CsvFS)      是的
'            弹性文件系统(ReFS)             是的
'    当CsvFs暂停时，此调用可能会失败，并出现一个错误，表明存在锁冲突。
'@Requirements
'    Minimum supported client    Windows XP [desktop apps only]
'    Minimum supported server    Windows Server 2003 [desktop apps only]
'    Target Platform Windows
'    Header  memoryapi.h (include Windows.h, Memoryapi.h)
'    Library kernel32.lib
'    dll kernel32.dll
Private Declare Function OpenFileMapping Lib "kernel32.dll" Alias "OpenFileMappingA" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal lpName As String) As Long
'@原型
'    HANDLE OpenFileMappingA(
'      DWORD  dwDesiredAccess,
'      BOOL   bInheritHandle,
'      LPCSTR lpName
'    );
'@功能
'    打开命名文件映射对象。
'@参数
'dwDesiredAccess
'    访问文件映射对象。此访问将根据目标文件映射对象上的任何安全描述符进行检查。有关值列表，请参见文件映射安全性和访问权限。
'bInheritHandle
'    如果该参数为真，则由CreateProcess函数创建的进程可以继承该句柄;否则，句柄不能继承。
'lpName
'    要打开的文件映射对象的名称。如果文件映射对象有一个按此名称打开的句柄，并且映射对象上的安全描述符不与dwDesiredAccess参数冲突，则打开操作成功。名称可以具有“全局”或“本地”前缀，以显式地打开全局或会话名称空间中的对象。名称的其余部分可以包含除反斜杠字符()之外的任何字符。有关更多信息，请参见内核对象名称空间。使用终端服务会话实现快速用户切换。第一个登录的用户使用会话0，下一个登录的用户使用会话1，依此类推。内核对象名称必须遵循终端服务的指导原则，以便应用程序能够支持多个用户。
'返回值
'    如果函数成功，返回值是指定文件映射对象的打开句柄。
'    如果函数失败，返回值为NULL。要获取扩展的错误信息，请调用GetLastError。
'@备注
'    OpenFileMapping返回的句柄可以用于任何需要文件映射对象句柄的函数?
'    当通过映射视图修改文件时，最后一次修改的时间戳可能不会自动更新。如果需要，调用者应该使用SetFileTime设置时间戳。
'    当不再需要该句柄时，调用方应该调用释放返回的句柄，方法是通过调用OpenFileMapping来关闭句柄。
'    在Windows Server 2012中，此功能由以下技术支持。
'        服务器消息块(SMB) 3.0协议       是的
'        SMB 3.0透明故障转移(TFO)        是的
'        带扩展文件共享(SO)的SMB 3.0     是的
'        集群共享卷文件系统(CsvFS)       是的
'        弹性文件系统(ReFS)              是的
'@Requirements
'    Minimum supported client    Windows XP [desktop apps only]
'    Minimum supported server    Windows Server 2003 [desktop apps only]
'    Target Platform Windows
'    Header  winbase.h (include Windows.h, Memoryapi.h)
'    Library kernel32.lib
'    dll kernel32.dll
Private Declare Sub RtlMoveMemory Lib "kernel32" (Destination As Any, Source As Any, ByVal Length As Long)
'说明：将内存块从一个位置移动到另一个位置
'Destination:指向移动目的地起始地址的指针。
'Source:指向要移动的内存块起始地址的指针。
'Length:内存块的大小以字节为单位移动。
'注意事项：这个函数定义为RtlMoveMemory函数。它的实现是内联的。有关更多信息，请参见WinBase。h和Winnt.h。源和目标块可能会重叠。
'           第一个参数，目的地，必须足够大，以容纳长度字节的源;否则，可能会出现缓冲区溢出。这可能导致拒绝服务攻击，如果有访问违反，或者在最坏的情况下，允许攻击者向您的进程注入可执行代码。如果目的地是一个基于堆栈的缓冲区，则尤其如此。要注意，最后一个参数，长度，是将字节复制到目的地的数量，而不是目的地的大小。
Private Declare Function UnmapViewOfFile Lib "kernel32" (ByVal lpBaseAddress As Long) As Long
'@原型
'    BOOL UnmapViewOfFile(
'      LPCVOID lpBaseAddress
'    );
'@功能
'    从调用进程的地址空间解映射文件的映射视图。
'@参数
'lpBaseAddress
'    指向要取消映射的文件的映射视图的基本地址的指针?这个值必须与之前调用MapViewOfFile或MapViewOfFileEx函数返回的值相同?
'@返回值
'    如果函数成功，返回值为非零。
'    如果函数失败，返回值为零。要获取扩展的错误信息，请调用GetLastError。
'@备注
'    取消映射文件的映射视图将使视图在进程的地址空间中占用的范围无效，并使该范围可用于其他分配。它删除了作为流程工作集一部分的每个未映射虚拟页的工作集条目，并减少了流程的工作集大小。它还减少了相应物理页面的共享计数。
'    未映射视图中的修改页在其共享计数为零之前不会被写入磁盘，换句话说，直到它们从共享这些页的所有进程的工作集中被未映射或删除。即使这样，修改后的页面也被“惰性地”写入磁盘;也就是说，修改可以缓存在内存中，并在稍后写入磁盘。为了在电源故障或系统崩溃时最小化数据丢失的风险，应用程序应该使用FlushViewOfFile函数显式地刷新修改过的页面。
'    虽然应用程序可以关闭用于创建文件映射对象的文件句柄，但是系统会保持相应的文件打开状态，直到文件的最后一个视图未映射为止。未对最后一个视图进行未映射的文件保持打开状态，没有共享限制。
'    在Windows Server 2012中，此功能由以下技术支持。
'        服务器消息块(SMB) 3.0协议          是的
'        SMB 3.0透明故障转移(TFO)           是的
'        带扩展文件共享(SO)的SMB 3.0        是的
'        集群共享卷文件系统(CsvFS)          是的
'        弹性文件系统(ReFS)                 是的
'@Requirements
'    Minimum supported client    Windows XP [desktop apps | UWP apps]
'    Minimum supported server    Windows Server 2003 [desktop apps | UWP apps]
'    Target Platform Windows
'    Header  memoryapi.h (include Windows.h, Memoryapi.h)
'    Library kernel32.lib
'    dll kernel32.dll
Private Const ERROR_ALREADY_EXISTS          As Long = &HB7
Private Const M_INIT_FILEMAPPING_SIZE       As Long = 2 ^ 16 - 1
'---------------------------------------------------------------------------
'                1、常规变量
'---------------------------------------------------------------------------
Private mlngFile                            As Long
Private mlngMapPtr                          As Long             '映射的内存地址
Private mblnCreate                          As Boolean
Private Type MemHeader
    MaxSize             As Long                         '该区域的最大尺寸
    Writed              As Byte                         '该区域的写标志
    Readed              As Byte                         '该区域的读标志
    Reserved            As Integer                      '为了四字节对齐，因此增加该保留区
    ProcessID           As Long                         '进程ID
    Param               As Long                         '状态的参数
    State               As Long                         '状态的值
    DataSize            As Long                         '使用的总长度
End Type
'---------------------------------------------------------------------------
'                2、属性变量与定义
'---------------------------------------------------------------------------
Private mstrShareName                           As String           '共享名称
Private mmhCur                                  As MemHeader
Private mstrData                                As String

'ShareName(R):当前共享区域的全局唯一名称，建议使用GUID
Public Property Get ShareName() As String
    ShareName = mstrShareName
End Property
'MaxSize(R):当前共享区域的最大值
Public Property Get MaxSize() As Long
    MaxSize = mmhCur.MaxSize
End Property
'Writed(R):写入状态。0-未写入。1-已经写入
Public Property Get Writed() As Long
    Writed = CLng(mmhCur.Writed)
End Property
'Readed(R):读取状态。0-未读取。>0-读取次数
Public Property Get Readed() As Long
    Readed = CLng(mmhCur.Readed)
End Property
'CurrentDataSize(R):当前共享域存储的字符串大小
Public Property Get CurrentDataSize() As Long
    CurrentDataSize = mmhCur.DataSize
End Property
'ProcessID(R):当前共享区域的写入进程ID
Public Property Get ProcessID() As String
    ProcessID = mmhCur.ProcessID
End Property
'Param(R):当前共享域的进程交互的参数
Public Property Get Param() As Long
    Param = mmhCur.Param
End Property
'State(R):当前共享域的进程交互的状态值
Public Property Get State() As Long
    State = mmhCur.State
End Property
'Data(R):当前共享区域的内容
Public Property Get Data() As String
    Data = mstrData
End Property
'---------------------------------------------------------------------------
'                3、公共方法
'---------------------------------------------------------------------------
'@方法    CreateMemoryShare
'   建立共享内存
'@返回值  Boolean
'
'@参数:
'strShareName String In
'   共享名，建议使用GUID,自动增加前缀 FILEMAPPING_,防止不同类型对象同名
'lngSize    Long In (Optional, Default=65535)
'   共享内存大小
'@备注
'
Public Function CreateMemoryShare(ByVal strShareName As String, Optional ByVal lngSize As Long = M_INIT_FILEMAPPING_SIZE) As Boolean
    '建立内存共享
    On Error GoTo ErrH
    Call Logger.PushMethod("ZLHelperMain.clsMemoryShare.CreateMemoryShare", "FILEMAPPING_" & strShareName, lngSize)
    Call CloseMemory
    mstrShareName = strShareName
    mlngFile = CreateFileMapping(-1, ByVal 0&, PAGE_READWRITE, 0, lngSize, "FILEMAPPING_" & strShareName)
    If mlngFile = 0 Then
        Logger.Error "CreateFileMapping", "Error", GetLastDllErr(Err.LastDllError)
    Else
        mlngMapPtr = MapViewOfFile(mlngFile, FILE_MAP_READ Or FILE_MAP_WRITE, 0, 0, 0)
        If mlngMapPtr = 0 Then
            Logger.Error "MapViewOfFile", "Error", GetLastDllErr(Err.LastDllError)
        Else
            mmhCur.MaxSize = lngSize
            mmhCur.DataSize = Len(mmhCur)
            '缓存区域头
            RtlMoveMemory ByVal mlngMapPtr, mmhCur, LenB(mmhCur)
            mblnCreate = True
            CreateMemoryShare = True
        End If
    End If
    If Not CreateMemoryShare Then
        Call CloseMemory
    End If
    Call Logger.PopMethod("ZLHelperMain.clsMemoryShare.CreateMemoryShare", CreateMemoryShare)
    Exit Function
ErrH:
    If Logger.ErrCenter("ZLHelperMain.clsMemoryShare.CreateMemoryShare") = 1 Then
        Resume
    End If
    Call CloseMemory
    Call Logger.PopMethod("ZLHelperMain.clsMemoryShare.CreateMemoryShare")
End Function
'@方法    OpenMemoryShare
'   打开已经存在的内存共享
'@返回值  Boolean
'
'@参数:
'strShareName String In
'   共享名，自动增加前缀 FILEMAPPING_
'@备注
'
Public Function OpenMemoryShare(ByVal strShareName As String) As Boolean
    '打开内存共享
    On Error GoTo ErrH
    Call Logger.PushMethod("ZLHelperMain.clsMemoryShare.OpenMemoryShare", "FILEMAPPING_" & strShareName)
    Call CloseMemory
    mstrShareName = strShareName
    mlngFile = OpenFileMapping(FILE_MAP_READ Or FILE_MAP_WRITE, 0, "FILEMAPPING_" & strShareName) '打开命名内存文件映射对象
    If mlngFile = 0 Then
        Logger.Error "OpenFileMapping", "Error", GetLastDllErr(Err.LastDllError)
    Else
        mlngMapPtr = MapViewOfFile(mlngFile, FILE_MAP_READ Or FILE_MAP_WRITE, 0, 0, 0)
        If mlngMapPtr = 0 Then
            Logger.Error "MapViewOfFile", "Error", GetLastDllErr(Err.LastDllError)
        Else
            '缓存区域头
            RtlMoveMemory mmhCur, ByVal mlngMapPtr, LenB(mmhCur)
            OpenMemoryShare = True
        End If
    End If
    If Not OpenMemoryShare Then
        Call CloseMemory
    End If
    Call Logger.PopMethod("ZLHelperMain.clsMemoryShare.OpenMemoryShare", OpenMemoryShare)
    Exit Function
ErrH:
    If Logger.ErrCenter("ZLHelperMain.clsMemoryShare.OpenMemoryShare") = 1 Then
        Resume
    End If
    Call CloseMemory
    Call Logger.PopMethod("ZLHelperMain.clsMemoryShare.OpenMemoryShare")
End Function
'@方法    CloseMemory
'   关闭内存映射。该方法中中断需要重启VB
'@返回值
'
'@参数:
'Param 方向
'
'@备注
'
Private Function CloseMemory() As Boolean
    Dim mhCur       As MemHeader
    On Error GoTo ErrH
    Call Logger.PushMethod("ZLHelperMain.clsMemoryShare.CloseMemory")
    CloseMemory = True
    '清空状态与头部缓存
    If mblnCreate Then
        If mlngMapPtr <> 0 And mlngFile <> 0 Then
            RtlMoveMemory mmhCur, mhCur, LenB(mhCur)
        End If
        mstrData = ""
    End If
    If mlngMapPtr <> 0 Then
        If UnmapViewOfFile(mlngMapPtr) = 0 Then
            Logger.Error "UnmapViewOfFile", "Error", GetLastDllErr(Err.LastDllError)
            CloseMemory = False
        Else
            mlngMapPtr = 0
        End If
    End If
    If mlngFile <> 0 Then
        If CloseHandle(mlngFile) = 0 Then
            Logger.Error "CloseHandle", "Error", GetLastDllErr(Err.LastDllError)
            CloseMemory = False
        Else
            mlngFile = 0
        End If
    End If
    If CloseMemory Then
        mstrShareName = ""
    End If
    Call Logger.PopMethod("ZLHelperMain.clsMemoryShare.CloseMemory", CloseMemory)
    Exit Function
ErrH:
    If Logger.ErrCenter("ZLHelperMain.clsMemoryShare.CloseMemory") = 1 Then
        Resume
    End If
    If mlngMapPtr <> 0 Then
        If UnmapViewOfFile(mlngMapPtr) = 0 Then
            Logger.Error "UnmapViewOfFile", "Error", GetLastDllErr(Err.LastDllError)
        End If
    End If
    If mlngFile <> 0 Then
        If CloseHandle(mlngFile) = 0 Then
            Logger.Error "CloseHandle", "Error", GetLastDllErr(Err.LastDllError)
        End If
    End If
    CloseMemory = False
    Call Logger.PopMethod("ZLHelperMain.clsMemoryShare.CloseMemory")
End Function
'@方法    WriteMemory
'   向内存映射写入数据。所有数据同时写入，不能逐一写入
'@返回值  Boolean
'
'@参数:
'strInfo String In
'   写入的内容
'lngProcessPID Long In
'   写入共享的进程ID。即当前进程ID
'lngParam   Long In
'   共享状态参数。
'lngState Long  In
'   共享状态值
'blnForceWrite Boolean In
'   当共享内存存在尚未读取的数据时，强制写入
'@备注
'
Public Function WriteMemory(ByVal strInfo As String, Optional ByVal lngProcessPID As Long, Optional ByVal lngParam As Long, Optional ByVal lngState As Long, Optional ByVal blnForceWrite As Boolean) As Boolean
    Dim arrData()       As Byte
    Dim lngTotal        As Long
    Dim arrByteTmp()    As Byte
    Dim objMutex        As New clsMutex
    
    On Error GoTo ErrH
    Call Logger.PushMethod("ZLHelperMain.clsMemoryShare.WriteMemory", strInfo, lngProcessPID, lngParam, lngState)
    If mlngMapPtr <> 0 Then
        If Not objMutex.CheckMutex(mstrShareName) Then '没有互斥体，则自动创建，该函数结束时，会自动释放互斥体
            If SubReadMemory(False, True) Then
                '尚未被读取
                If Not (mmhCur.Writed = 1 And mmhCur.Readed = 0) Or blnForceWrite Then
                    mmhCur.Writed = 1
                    mmhCur.Readed = 0
                    mmhCur.ProcessID = lngProcessPID
                    mmhCur.Param = lngParam
                    mmhCur.State = lngState
                    If strInfo <> "" Then
                        arrByteTmp = StrConv(strInfo, vbFromUnicode)
                        lngTotal = UBound(arrByteTmp) + 1
                    End If
                    mmhCur.DataSize = lngTotal + LenB(mmhCur)
                    ReDim Preserve arrData(mmhCur.DataSize - 1)
                    '将头部放在临时区域
                    Call RtlMoveMemory(arrData(0), mmhCur, LenB(mmhCur))
                    '字符串放在临时存储区域
                    If lngTotal > 0 Then
                        Call RtlMoveMemory(arrData(mmhCur.DataSize - lngTotal), arrByteTmp(0), lngTotal)
                    End If
                    '将临时缓存放在共享内存
                    Call RtlMoveMemory(ByVal mlngMapPtr, arrData(0), mmhCur.DataSize)
                    WriteMemory = True
                End If
            End If
        End If
    End If
    Call Logger.PopMethod("ZLHelperMain.clsMemoryShare.WriteMemory", WriteMemory)
    Exit Function
ErrH:
    If Logger.ErrCenter("ZLHelperMain.clsMemoryShare.WriteMemory") = 1 Then
        Resume
    End If
    
    Call Logger.PopMethod("ZLHelperMain.clsMemoryShare.WriteMemory")
End Function
'@方法    ReadMemory
'   从内存映射文件内读数据。该方法中中断需要重启VB
'@返回值  Boolean
'
'@参数:
'@备注
'
Public Function ReadMemory() As Boolean
    Dim objMutex        As New clsMutex
    
    On Error GoTo ErrH
    Call Logger.PushMethod("ZLHelperMain.clsMemoryShare.ReadMemory")
    If mlngMapPtr <> 0 Then
        If Not objMutex.CheckMutex(mstrShareName) Then  '没有互斥体，则自动创建，该函数结束时，会自动释放互斥体
            ReadMemory = SubReadMemory()
        End If
    End If
    Call Logger.PopMethod("ZLHelperMain.clsMemoryShare.ReadMemory", ReadMemory)
    Exit Function
ErrH:
    If Logger.ErrCenter("ZLHelperMain.clsMemoryShare.ReadMemory") = 1 Then
        Resume
    End If
    
    Call Logger.PopMethod("ZLHelperMain.clsMemoryShare.ReadMemory")
End Function
'@方法    ReadMemoryOnce
'   从内存映射文件内读数据。调用时成功后自动释放该内存区域。该方法中中断需要重启VB。
'@返回值  Boolean
'
'@参数:
'Param 方向
'
'@备注
'
Public Function ReadMemoryOnce() As Boolean
    Dim objMutex        As New clsMutex
    On Error GoTo ErrH
    Call Logger.PushMethod("ZLHelperMain.clsMemoryShare.ReadMemoryOnce")
    If mlngMapPtr <> 0 Then
        If Not objMutex.CheckMutex(mstrShareName) Then  '没有互斥体，则自动创建，该函数结束时，会自动释放互斥体
            ReadMemoryOnce = SubReadMemory(True)
        End If
    End If
    Call Logger.PopMethod("ZLHelperMain.clsMemoryShare.ReadMemoryOnce", ReadMemoryOnce)
    Exit Function
ErrH:
    If Logger.ErrCenter("ZLHelperMain.clsMemoryShare.ReadMemoryOnce") = 1 Then
        Resume
    End If
    
    Call Logger.PopMethod("ZLHelperMain.clsMemoryShare.ReadMemoryOnce")
End Function
'@方法    ReleaseMemory
'   释放内存。该方法中中断需要重启VB
'@返回值  Boolean
'
'@参数:
'@备注
'
Public Function ReleaseMemory() As Boolean
    Dim objMutex        As New clsMutex
    On Error GoTo ErrH
    Call Logger.PushMethod("ZLHelperMain.clsMemoryShare.ReleaseMemory")
    If mlngMapPtr <> 0 Then
        If Not objMutex.CheckMutex(mstrShareName) Then
            ReleaseMemory = SubReleaseMemory
        End If
    End If
    Call Logger.PopMethod("ZLHelperMain.clsMemoryShare.ReleaseMemory", ReleaseMemory)
    Exit Function
ErrH:
    If Logger.ErrCenter("ZLHelperMain.clsMemoryShare.ReleaseMemory") = 1 Then
        Resume
    End If
    
    Call Logger.PopMethod("ZLHelperMain.clsMemoryShare.ReleaseMemory")
End Function
'---------------------------------------------------------------------------
'                4、私有方法
'---------------------------------------------------------------------------
'@方法    SubReadMemory
'   从内存映射文件内读数据。SubReadMemory与SubReadMemoryOnce的子函数
'@返回值  Boolean
'
'@参数:
'blnOnceCall Boolean In (Optional)
'   是否只调用一次，若是，调用时成功后自动释放该内存区域
'blnOnlyHeader  Boolean In (Optional)
'   是否只读取头部
'@备注
'
Private Function SubReadMemory(Optional ByVal blnOnceCall As Boolean, Optional ByVal blnOnlyHeader As Boolean) As Boolean
    Dim arrData()       As Byte
    Dim lngTotal        As Long
    Dim arrByteTmp()    As Byte
    Dim mhCur           As MemHeader
    On Error GoTo ErrH
    Call Logger.PushMethod("ZLHelperMain.clsMemoryShare.SubReadMemory", blnOnceCall, blnOnlyHeader)
    '读取共享内存头部，并对头部合法性进行检查
    Call RtlMoveMemory(mhCur, ByVal mlngMapPtr, LenB(mhCur))
    If mhCur.MaxSize <> mmhCur.MaxSize Or mhCur.DataSize < 0 Or mhCur.DataSize > mhCur.MaxSize Then
        Logger.Warn "共享内存头部和初始化不匹配", "ST_MaxSize", mmhCur.MaxSize, "MaxSize", mhCur.MaxSize, "ST_DataSize", mmhCur.DataSize, "DataSize", mhCur.DataSize
    Else
        Call RtlMoveMemory(mmhCur, mhCur, LenB(mhCur))
        Logger.DebugEx "共享内存头部", "MaxSize", mmhCur.MaxSize, "DataSize", mmhCur.DataSize, "Writed", mmhCur.Writed, "Readed", mmhCur.Readed, "ProcessID", mmhCur.ProcessID, "Param", mmhCur.Param, "State", mmhCur.State
        If blnOnlyHeader Then
            SubReadMemory = True
        Else
            ReDim Preserve arrData(mmhCur.DataSize - 1)
            lngTotal = mmhCur.DataSize - LenB(mmhCur)
            Call RtlMoveMemory(arrData(0), ByVal mlngMapPtr, mmhCur.DataSize)
            If mmhCur.Writed <> 0 Then
                If lngTotal > 0 Then
                    ReDim Preserve arrByteTmp(lngTotal - 1)
                    Call RtlMoveMemory(arrByteTmp(0), arrData(mmhCur.DataSize - lngTotal), lngTotal)
                    mstrData = StrConv(arrByteTmp, vbUnicode)
                    Logger.DebugEx "共享内存扩展信息", "Data", mstrData
                Else
                    mstrData = ""
                    Logger.DebugEx "共享内存扩展信息", "Data", mstrData
                End If
                
                mmhCur.Readed = (mmhCur.Readed Mod 255) + 1
                Call RtlMoveMemory(ByVal mlngMapPtr, mmhCur, LenB(mmhCur))         '回写状态
                If blnOnceCall Then Call SubReleaseMemory
                SubReadMemory = True
            End If
        End If
    End If
    Call Logger.PopMethod("ZLHelperMain.clsMemoryShare.SubReadMemory", SubReadMemory)
    Exit Function
ErrH:
    If Logger.ErrCenter("ZLHelperMain.clsMemoryShare.SubReadMemory") = 1 Then
        Resume
    End If
    
    Call Logger.PopMethod("ZLHelperMain.clsMemoryShare.SubReadMemory")
End Function
'--------------------------------------------------------------------------------------------------
'方法           ReleaseMemory
'功能
'返回值         Boolean
'入参列表:
'参数名         类型                    说明
'-------------------------------------------------------------------------------------------------

'@方法    SubReleaseMemory
'   释放内存。SubReadMemory与ReleaseMemory子函数
'@返回值  Boolean
'
'@参数:
'@备注
'
Private Function SubReleaseMemory() As Boolean
    On Error GoTo ErrH
    Call Logger.PushMethod("ZLHelperMain.clsMemoryShare.SubReleaseMemory")
    If mlngMapPtr <> 0 Then
        '将清空后的状态缓存
        mmhCur.Writed = 0
        mmhCur.Readed = 0
        mmhCur.DataSize = LenB(mmhCur)
        '将清空的状态放入内存共享
        Call RtlMoveMemory(ByVal mlngMapPtr, mmhCur, LenB(mmhCur))
    End If
    SubReleaseMemory = True
    Call Logger.PopMethod("ZLHelperMain.clsMemoryShare.SubReleaseMemory", SubReleaseMemory)
    Exit Function
ErrH:
    If Logger.ErrCenter("ZLHelperMain.clsMemoryShare.SubReleaseMemory") = 1 Then
        Resume
    End If
    
    Call Logger.PopMethod("ZLHelperMain.clsMemoryShare.SubReleaseMemory")
End Function
'---------------------------------------------------------------------------
'                5、对象方法与事件
'---------------------------------------------------------------------------
Private Sub Class_Initialize()
    On Error GoTo ErrH
    Call Logger.PushMethod("ZLHelperMain.clsMemoryShare.Class_Initialize")
    Call Logger.PopMethod("ZLHelperMain.clsMemoryShare.Class_Initialize")
    Exit Sub
ErrH:
    If Logger.ErrCenter("ZLHelperMain.clsMemoryShare.Class_Initialize") = 1 Then
        Resume
    End If
    
    Call Logger.PopMethod("ZLHelperMain.clsMemoryShare.Class_Initialize")
End Sub

Private Sub Class_Terminate()
    On Error GoTo ErrH
    Call Logger.PushMethod("ZLHelperMain.clsMemoryShare.Class_Terminate")
    Call CloseMemory
    Call Logger.PopMethod("ZLHelperMain.clsMemoryShare.Class_Terminate")
    Exit Sub
ErrH:
    If Logger.ErrCenter("ZLHelperMain.clsMemoryShare.Class_Terminate") = 1 Then
        Resume
    End If
    
    Call Logger.PopMethod("ZLHelperMain.clsMemoryShare.Class_Terminate")
End Sub


