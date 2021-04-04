Attribute VB_Name = "mdlPublic"
Option Explicit


Private Const CP_UTF8                       As Long = 65001
Private Const WAIT_ABANDONED                As Long = &H80          '指定的对象是一个互斥对象，这个互斥对象是在拥有线程终止之前拥有互斥对象的线程所释放的。
                                                                    '互斥对象的所有权被授予给调用线程，而互斥锁状态被设置为非信号。
                                                                    '如果互斥锁是在保护持久状态信息，那么您应该检查它是否具有一致性。
Private Const WAIT_OBJECT_0                 As Long = &H0           '指定对象的状态是有信号的?
Private Const WAIT_TIMEOUT                  As Long = &H102         '超时时间间隔，对象的状态是没有信号的。
Private Const WAIT_FAILED                   As Long = &HFFFFFFFF    '这个函数失败了。要获得扩展错误信息，请调用GetLastError
Private Const MAX_PATH                      As Long = 260
Private Const TOKEN_QUERY                   As Long = &H8
Private Const PROCESS_QUERY_INFORMATION     As Long = &H400
Private Const PROCESS_VM_READ               As Long = &H10
Private Const WTS_CURRENT_SERVER_HANDLE     As Long = 0
Private Const WTS_CURRENT_SESSION           As Long = -1
Private Const IMAGE_SIZEOF_SHORT_NAME            As Integer = 8
Private Const IMAGE_NUMBEROF_DIRECTORY_ENTRIES   As Integer = 16

Public Const SM4_CRYPT_RANDOMIZE_KEY            As Long = 999                       'sm4加密算法密钥生成器的随机种子
Public Const SM4_CRYPT_RANDOMIZE_IV             As Long = 666                       'sm4加密算法初始向量生成器的随机种子
Public Const IMAGE_DOS_SIGNATURE                As Integer = &H5A4D                 'MZ
Public Const IMAGE_OS2_SIGNATURE                As Integer = &H454E                 'NE
Public Const IMAGE_OS2_SIGNATURE_LE             As Integer = &H454C                 'LE
Public Const IMAGE_NT_SIGNATURE                 As Long = &H4550                    'PE00
Public Const IMAGE_NT_OPTIONAL_HDR32_MAGIC      As Integer = &H10B                  '这是一个32位镜像文件
Public Const IMAGE_NT_OPTIONAL_HDR64_MAGIC      As Integer = &H20B                  '这是一个PE32+可执行文件
Public Const IMAGE_ROM_OPTIONAL_HDR_MAGIC       As Integer = &H107                  '这是一个ROM镜像
Public Const IMAGE_FILE_RELOCS_STRIPPED         As Long = &H1                       ' 仅适用于映像文件，适用于 Windows CE、Microsoft Windows NT. 及其后继操作系统.它表明此文件不包含基址重定位信息，因此必须被加载到其首选基地址上。如果基地址不可用，加载器会报错。链接器默认会移除可执行（EXE）文件中的重定位信息.
Public Const IMAGE_FILE_EXECUTABLE_IMAGE        As Long = &H2                       ' File is executable  (i.e. no unresolved externel references).仅适用于映像文件。它表明此映像文件是合法的，可以被运行。如果未设置此标志，表明出现了链接器错误
Public Const IMAGE_FILE_LINE_NUMS_STRIPPED      As Long = &H4                       ' Line nunbers stripped from file.行号信息已经被移除。不赞成使用此标志，它应该为0。
Public Const IMAGE_FILE_LOCAL_SYMS_STRIPPED     As Long = &H8                       ' Local symbols stripped from file.COFF 符号表中有关局部符号的项已经被移除。不赞成使用此标志，它应该为0。
Public Const IMAGE_FILE_AGGRESIVE_WS_TRIM       As Long = &H10                      ' Agressively trim working set 此标志已经被舍弃。它用于调整工作集。Windows 2000 及其后继操作系统不赞成使用此标志，它应该为0。
Public Const IMAGE_FILE_LARGE_ADDRESS_AWARE     As Long = &H20                      ' App can handle >2gb addresses 应用程序可以处理大于2GB 的地址。
Public Const IMAGE_FILE_BYTES_REVERSED_LO       As Long = &H80                      ' Bytes of machine word are reversed. 小尾：在内存中，最低位（LSB）在最高位（MSB）前面。不赞成使用此标志，它应该为0.
Public Const IMAGE_FILE_32BIT_MACHINE           As Long = &H100                     ' 32 bit word machine.          机器类型基于32 位字体系结构。
Public Const IMAGE_FILE_DEBUG_STRIPPED          As Long = &H200                     ' Debugging info stripped from file in .DBG file 调试信息已经从此映像文件中移除。
Public Const IMAGE_FILE_REMOVABLE_RUN_FROM_SWAP As Long = &H400                     ' If Image is on removable media, copy and run from the swap file. 如果此映像文件在可移动介质上，完全加载它并把它复制到交换文件中
Public Const IMAGE_FILE_NET_RUN_FROM_SWAP       As Long = &H800                     ' If Image is on Net, copy and run from the swap file.如果此映像文件在网络介质上，完全加载它并把它复制到交换文件中
Public Const IMAGE_FILE_SYSTEM                  As Long = &H1000                    ' System File.  此映像文件是系统文件，而不是用户程序
Public Const IMAGE_FILE_DLL                     As Long = &H2000                    ' File is a DLL.此映像文件是动态链接库（DLL）。这样的文件总被认为是可执行文件，尽管它们并不能直接被运行
Public Const IMAGE_FILE_UP_SYSTEM_ONLY          As Long = &H4000                    ' File should only be run on a UP machine   此文件只能运行于单处理器机器上。
Public Const IMAGE_FILE_BYTES_REVERSED_HI       As Long = &H8000                    ' Bytes of machine word are reversed. 大尾：在内存中，MSB 在LSB 前面。不赞成使用此标志，它应该为0。

'IMAGE_OPTIONAL_HEADER Directory Entries
Public Const IMAGE_DIRECTORY_ENTRY_EXPORT       As Integer = 0                      ' Export Directory  导出表的地址和大小
Public Const IMAGE_DIRECTORY_ENTRY_IMPORT       As Integer = 1                      ' Import Directory  导入表的地址和大小
Public Const IMAGE_DIRECTORY_ENTRY_RESOURCE     As Integer = 2                      ' Resource Directory 资源表的地址和大小
Public Const IMAGE_DIRECTORY_ENTRY_EXCEPTION    As Integer = 3                      ' Exception Directory   异常表的地址和大小
Public Const IMAGE_DIRECTORY_ENTRY_SECURITY     As Integer = 4                      ' Security Directory    属性证书表的地址和大小 Certificate Table 域指向属性证书表。这些证书并不作为映像的一部分被加载进内存。此时它的第一个域是一个文件指针，而不是通常的RVA
Public Const IMAGE_DIRECTORY_ENTRY_BASERELOC    As Integer = 5                      ' Base Relocation Table 基址重定位表的地址和大小
Public Const IMAGE_DIRECTORY_ENTRY_DEBUG        As Integer = 6                      ' Debug Directory       调试数据起始地址和大小
'Public Const IMAGE_DIRECTORY_ENTRY_COPYRIGHT    As Integer = 7                      ' (X86 usage)          调试数据起始地址和大小。
Public Const IMAGE_DIRECTORY_ENTRY_ARCHITECTURE As Integer = 7                      ' Architecture Specific Data 保留，必须为0
Public Const IMAGE_DIRECTORY_ENTRY_GLOBALPTR    As Integer = 8                      ' RVA of GP             将被存储在全局指针寄存器中的一个值的RVA。这个结构的Size 域必须为0
Public Const IMAGE_DIRECTORY_ENTRY_TLS          As Integer = 9                      ' TLS Directory         线程局部存储（TLS）表的地址和大小
Public Const IMAGE_DIRECTORY_ENTRY_LOAD_CONFIG  As Integer = 10                     ' Load Configuration Directory  加载配置表的地址和大小
Public Const IMAGE_DIRECTORY_ENTRY_BOUND_IMPORT As Integer = 11                     ' Bound Import Directory in headers 绑定导入表的地址和大小
Public Const IMAGE_DIRECTORY_ENTRY_IAT          As Integer = 12                     ' Import Address Table  绑定导入查找表的地址和大小
Public Const IMAGE_DIRECTORY_ENTRY_DELAY_IMPORT As Integer = 13                     ' Delay Load Import Descriptors 延迟导入描述符的地址和大小
Public Const IMAGE_DIRECTORY_ENTRY_COM_DESCRIPTOR   As Integer = 14                 ' COM Runtime descriptor    CLR 运行时头部的地址和大小
Public Const DATA_DIRECTORY_OTHER               As Integer = 15                     ' COM Runtime descriptor    保留，必须为0

Private Enum CrypeMode
    CM_Encrypt = 1   '加密
    CM_Decrypt = 0   '解密
End Enum

Public Type IMAGE_RESOURCE_DIR_STRING_U
    Length              As Integer              '字符串的长度
    NameString          As Integer              'UNICODE字符串，由于字符串是不定长的，所以这里只能用一个dw表示，实际上当长度为100的时候，这里的数据是NameString dw 100 dup (?)
End Type

Public Type IMAGE_IMPORT_DESCRIPTOR
    'DUMMYUNIONNAME
'    Characteristics     As Long                 'DUMMYUNIONNAME_Characteristics 0 for terminating null import descriptor 导入查找表的RVA。这个表包含了每一个导入符号的名称或序数
    OriginalFirstThunk  As Long                 'DUMMYUNIONNAME_OriginalFirstThunk  RVA to original unbound IAT (PIMAGE_THUNK_DATA)
    TimeDateStamp       As Long                 ' 0 if not bound,
                                                ' -1 if bound, and real date\time stamp
                                                '   in IMAGE_DIRECTORY_ENTRY_BOUND_IMPORT (new BIND) 这个域一直被设置为0，直到映像被绑定。当映像被绑定之后，这个域被设置为这个DLL 的日期/时间戳。
                                                'O.W. date/time stamp of DLL bound to (Old BIND)
    ForwarderChain      As Long                 '-1 if no forwarders 第一个转发项的索引。
    Name                As Long                 ' 包含DLL 名称的ASCII 码字符串相对于映像基址的偏移地址
    FirstThunk          As Long                 'RVA to IAT (if bound this IAT has actual addresses) 导入地址表的RVA。这个表的内容与导入查找表的内容完全一样，直到映像被绑定。
End Type

Public Type IMAGE_DATA_DIRECTORY
    VirtualAddress      As Long                 '数据块的RVA
    Size                As Long                 '数据块大小
End Type

Public Type IMAGE_COR20_HEADER
    'Header versioning
    Cb                  As Long                 '头的字节大小
    MajorRuntimeVersion As Integer              'CLR需要运行的最小版本主板本号
    MinorRuntimeVersion As Integer              'CLR需要运行的最小版本次版本号

    'Symbol table and startup information
    MetaData            As IMAGE_DATA_DIRECTORY 'Rav和元数据的大小
    Flags               As Long                 '二进制标记

    'If COMIMAGE_FLAGS_NATIVE_ENTRYPOINT is not set, EntryPointToken represents a managed entrypoint.
    'If COMIMAGE_FLAGS_NATIVE_ENTRYPOINT is set, EntryPointRVA represents an RVA to a native entrypoint.
    'DUMMYUNIONNAME
    EntryPointToken     As Long                 'DUMMYUNIONNAME_EntryPointToken
'    EntryPointRVA       As Long                 'DUMMYUNIONNAME_EntryPointRVAToken

    'Binding information
    Resources           As IMAGE_DATA_DIRECTORY 'Rav和托管资源的大小
    StrongNameSignature As IMAGE_DATA_DIRECTORY 'Rav和用于这个pe文件的哈希数据的大小

    'Regular fixup and binding information
    CodeManagerTable    As IMAGE_DATA_DIRECTORY 'Rva和代码管理表的大小
    VTableFixups        As IMAGE_DATA_DIRECTORY 'Rav和一个由虚拟表修正组成的数组的字节大小
    ExportAddressTableJumps As IMAGE_DATA_DIRECTORY 'Rav和由jump thunk的地址组成的数组的大小
    'Precompiled image info (internal use only - set to zero)
    ManagedNativeHeader As IMAGE_DATA_DIRECTORY '为预编译而保留的，被设置为0
End Type

'.net Meta Data Structor
Public Type CLR_MetaDataVer
    Signature                                   As Long
    MajorVersion                                As Integer
    MinorVersion                                As Integer
    ExtraData                                   As Long
    Length                                      As Long
    VersionString(15)                           As Byte 'array[0..IMAGE_NUMBEROF_DIRECTORY_ENTRIES-1] of Char;   //.net字符串
    Flags                                       As Byte
    Pading                                      As Byte
    Streams                                     As Integer
End Type

Public Type IMAGE_RESOURCE_DIRECTORY
    Characteristics     As Long                 '理论上为资源的属性，不过事实上总是0
    TimeDateStamp       As Long                 '资源的产生时刻
    MajorVersion        As Integer              '理论上为资源的版本，不过事实上总是0
    MinorVersion        As Integer
    NumberOfNamedEntries    As Integer          '以名称命名的入口数量
    NumberOfIdEntries   As Integer              '以ID命名的入口数量
'    IMAGE_RESOURCE_DIRECTORY_ENTRY DirectoryEntries[];
End Type

Public Type IMAGE_RESOURCE_DIRECTORY_ENTRY
    'DUMMYSTRUCTNAME
    Name                As Long                 'DUMMYSTRUCTNAME_Name,目录项的名称字符串指针或ID
'    NameOffset          As Long                 'DUMMYSTRUCTNAME_RvaBased.Bits31
'    NameIsString        As Long                 'DUMMYSTRUCTNAME_RvaBased.Bits1
'    Id                  As Integer              'DUMMYSTRUCTNAME_Id
    'DUMMYUNIONNAME2
    OffsetToData        As Long                 'DUMMYSTRUCTNAME2_OffsetToData,目录项指针
'    OffsetToDirectory   As Long                 'DUMMYSTRUCTNAME2_OffsetToDirectory.Bits31
'    DataIsDirectory     As Long                 'DUMMYSTRUCTNAME2_DataIsDirectory.Bits1
End Type

Private Type SID_AND_ATTRIBUTES
    Sid         As Long
    Attributes  As Long
End Type

Private Type TOKEN_USER
    User As SID_AND_ATTRIBUTES
End Type

Private Enum REGErr
    ERROR_SUCCESS = &H0
    ERROR_FILE_NOT_FOUND = &H2 'The system cannot find the file specified
    ERROR_BADDB = 1009&
    ERROR_BADKEY = 1010&
    ERROR_CANTOPEN = 1011&
    ERROR_CANTREAD = 1012&
    ERROR_CANTWRITE = 1013&
    ERROR_OUTOFMEMORY = 14&
    ERROR_INVALID_PARAMETER = 87&
    ERROR_ACCESS_DENIED = 5&
    ERROR_NO_MORE_ITEMS = 259&
    ERROR_MORE_DATA = 234&
End Enum

Private Enum REGRights
    KEY_QUERY_VaLUE = &H1
    KEY_SET_VaLUE = &H2
    KEY_CREaTE_Sub_KEY = &H4
    KEY_ENUMERaTE_Sub_KEYS = &H8
    KEY_NOTIFY = &H10
    KEY_CREaTE_LINK = &H20
    KEY_aLL_aCCESS = &H3F
    KEY_READ = &H20019
End Enum

Private Enum REGValueType
    REG_NONE = 0                       ' No value type
    REG_SZ = 1 'Unicode空终结字符串
    REG_EXPAND_SZ = 2 'Unicode空终结字符串
    REG_BINARY = 3 '二进制数值
    REG_DWORD = 4 '32-bit 数字
    REG_DWORD_BIG_ENDIAN = 5
    REG_LINK = 6
    REG_MULTI_SZ = 7 ' 二进制数值串
End Enum

Private Enum REGRoot
    HKEY_CLASSES_ROOT = &H80000000 '记录Windows操作系统中所有数据文件的格式和关联信息，主要记录不同文件的文件名后缀和与之对应的应用程序。其下子键可分为两类，一类是已经注册的各类文件的扩展名，这类子键前面都有一个“。”；另一类是各类文件类型有关信息。
    HKEY_CURRENT_USER = &H80000001 '此根键包含了当前登录用户的用户配置文件信息。这些信息保证不同的用户登录计算机时，使用自己的个性化设置，例如自己定义的墙纸、自己的收件箱、自己的安全访问权限等。
    HKEY_LOCaL_MaCHINE = &H80000002 '此根键包含了当前计算机的配置数据，包括所安装的硬件以及软件的设置。这些信息是为所有的用户登录系统服务的。它是整个注册表中最庞大也是最重要的根键！
    HKEY_USERS = &H80000003 '此根键包括默认用户的信息（Default子键）和所有以前登录用户的信息。
    HKEY_PERFORMANCE_DATA = &H80000004 '在Windows NT/2000/XP注册表中虽然没有HKEY_DYN_DATA键，但是它却隐藏了一个名为“HKEY_ PERFOR MANCE_DATA”键。所有系统中的动态信息都是存放在此子键中。系统自带的注册表编辑器无法看到此键
    HKEY_CURRENT_CONFIG = &H80000005  '此根键实际上是HKEY_LOCAL_MACHINE中的一部分，其中存放的是计算机当前设置，如显示器、打印机等外设的设置信息等。它的子键与HKEY_LOCAL_ MACHINE\ Config\0001分支下的数据完全一样。
    HKEY_DYN_DATA = &H80000006 '此根键中保存每次系统启动时，创建的系统配置和当前性能信息。这个根键只存在于Windows 98中。
End Enum

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

Public Type IMAGE_SECTION_HEADER
    vName(IMAGE_SIZEOF_SHORT_NAME - 1) As Byte  '这是一个8 字节的UTF-8 编码的字符串，不足8 字节时用NULL 填充。如果它正好是8 字节，那就没有最后的NULL 字符。如果名称更长的话，这个域中是一个斜杠（/）后跟一个用ASCII 码表示的十进制数，这个十进制数表示字符串表中的偏移。可行映像不使用字符串表也不支持长度超过8 字节的节名.如果目标文件中有长节名的节最后要出现在可执行文件中，那么相应的长节名会被截断。
    Misc                As Long                 'PhysicalAddress or VirtualSize
    'PhysicalAddress
    'VirtualSize                                '当加载进内存时这个节的总大小。如果此值比SizeOfRawData 大，那么多出的部分用0 填充。这个域仅对可执行映像是合法的，对于目标文件来说，它应该为0。
    VirtualAddress      As Long                 '对于可执行映像来说，这个域的值是这个节被加载进内存之后它的第一个字节相对于映像基址的偏移地址。对于目标文件来说，这个域的值是没有重定位之前其第一个字节的地址；为了简单起见，编译器应该把此值设置为0。否则这个值是个任意值，但是在重定位时应该从偏移地址中减去这个值
    SizeOfRawData       As Long                 '（对于目标文件来说）节的大小或者（对于映像文件来说）磁盘文件中已初始化数据的大小。对于可执行映像来说，它必须是可选文件头中FileAlignment 域的倍数.如果它小于VirtualSize 域的值，余下的部分用0 填充。由于SizeOfRawData 域要向上舍入，但是VirtualSize域并不舍入，因此可能出现SizeOfRawData 域大于VirtualSize 域的情况.当节中仅包含未初始化的数据时，这个域应该为0。
    PointerToRawData    As Long                 '指向COFF 文件中节的第一个页面的文件指针。对于可执行映像来说，它必须是可选文件头中FileAlignment 域的倍数。对于目标文件来说，要获得最好的性能，此值应该按4 字节边界对齐。当节中仅包含未初始化的数据时，这个域应该为0。
    PointerToRelocations    As Long             '指向节中重定位项开头的文件指针。对于可执行文件或者没有重定位项的文件来说，此值应该为0。
    PointerToLinenumbers    As Long             '指向节中行号项开头的文件指针。如果没有COFF行号信息的话，此值应该为0。对于映像来说，此值应该为0，因为已经不赞成使用COFF 调试信息了
    NumberOfRelocations As Integer              '节中重定位项的个数。对于可执行映像来说，此值应该为0
    NumberOfLinenumbers As Integer              '节中行号项的个数。对于映像来说，此值应该为0，因为已经不赞成使用COFF 调试信息了。
    Characteristics     As Long                 '描述节特征的标志。
End Type

Public Type IMAGE_DOS_HEADER        'DOS .EXE header 64B
    e_magic             As Integer              'Magic number  分别为MZ,4Dh和5Ah
    e_cblp              As Integer              'Bytes on last page of file 文件最后一页字节数
    e_cp                As Integer              'Pages in file          文件总页数(512B/页)
    e_crlc              As Integer              'Relocations            重定位项数量
    e_cparhdr           As Integer              'Size of header in paragraphs   文件头总段数(16B/段)
    e_minalloc          As Integer              'Minimum extra paragraphs needed
    e_maxalloc          As Integer              'Maximum extra paragraphs needed
    e_ss                As Integer              'Initial (relative) SS value SS:SP
    e_sp                As Integer              'Initial SP value   SS:SP
    e_csum              As Integer              'Checksum           校验和
    e_ip                As Integer              'Initial IP value   CS:IP
    e_cs                As Integer              'Initial (relative) CS value    CS:IP
    e_lfarlc            As Integer              'File address of relocation table   重定位表偏移地址
    e_ovno              As Integer              'Overlay number
    e_res(3)            As Integer              'Reserved words
    e_oemid             As Integer              'OEM identifier (for e_oeminfo)
    e_oeminfo           As Integer              'OEM information; e_oemid specific
    e_res2(9)           As Integer              'Reserved words
    e_lfanew            As Long                 'File address of new exe header PE头偏移,指向PE文件头
End Type

Public Type IMAGE_FILE_HEADER                   '20B
    Machine             As Integer              '标识目标机器类型的数字
    NumberOfSections    As Integer              '节的数目。它给出了节表的大小，而节表紧跟着文件头
    TimeDateStamp       As Long                 '从UTC 时间1970 年1 月1 日00:00 起的总秒数（一个C 运行时time_t 类型的值）的低32 位，它指出文件何时被创建
    PointerToSymbolTable    As Long             'COFF 符号表的文件偏移。如果不存在COFF 符号表，此值为0。对于映像文件来说，此值应该为0，因为已经不赞成使用COFF 调试信息了
    NumberOfSymbols     As Long                 '符号表中的元素数目。由于字符串表紧跟符号表，所以可以利用这个值来定位字符串表?对于映像文件来说，此值应该为0，因为已经不赞成使用COFF调试信息了
    SizeOfOptionalHeader    As Integer          '可选文件头的大小。可执行文件需要可选文件头而目标文件并不需要。对于目标文件来说，此值应该为0
    Characteristics     As Integer              '指示文件属性的标志。
End Type

Private Type Big_Iint
    Low                 As Long
    High                As Long
End Type

Public Type IMAGE_OPTIONAL_HEADER32
    'Standard fields.标准字段
    Magic               As Integer              '这个无符号整数指出了映像文件的状态。最常用的数字是0x10B，它表明这是一个正常的可执行文件。0x107 表明这是一个ROM 映像，0x20B 表明这是一个PE32 + 可执行文件?
    MajorLinkerVersion  As Byte                 '链接器的主版本号
    MinorLinkerVersion  As Byte                 '链接器的次版本号
    SizeOfCode          As Long                 '代码节（.text）的大小。如果有多个代码节的话，它是所有代码节的和。
    SizeOfInitializedData   As Long             '已初始化数据节的大小。如果有多个这样的数据节的话，它是所有这些数据节的和。
    SizeOfUninitializedData As Long             '未初始化数据节（.bss）的大小。如果有多个.bss 节的话，它是所有这些节的和。
    AddressOfEntryPoint As Long                 '当可执行文件被加载进内存时其入口点相对于映像基址的偏移地址.对于一般程序映像来说，它就是启动地址。对于设备驱动程序来说，它是初始化函数的地址。入口点对于DLL来说是可选的。如果不存在入口点的话，这个域必须为0.
    BaseOfCode          As Long                 '当映像被加载进内存时代码节的开头相对于映像基址的偏移地址。
    BaseOfData          As Long                 '当映像被加载进内存时数据节的开头相对于映像基址的偏移地址。PE32独有
    ' NT additional fields.NT附加字段
    ImageBase           As Long                 '当加载进内存时映像的第一个字节的首选地址.它必须是64K 的倍数.DLL默认是0x10000000。Windows CE EXE默认是0x00010000.Windows NT?Windows 2000、Windows XP、Windows 95、Windows 98 和Windows Me 默认是0x00400000。
    SectionAlignment    As Long                 '当加载进内存时节的对齐值（以字节计）。它必须大于或等于FileAlignment.默认是相应系统的页面大小
    FileAlignment       As Long                 '用来对齐映像文件的节中的原始数据的对齐因子（以字节计）。它应该是界于512 和64K 之间的2 的幂（包括这两个边界值）。默认是512。如果SectionAlignment 小于相应系统的页面大小，那么FileAlignment 必须与SectionAlignment 匹配.
    MajorOperatingSystemVersion As Integer      '所需操作系统的主版本号。
    MinorOperatingSystemVersion As Integer      '所需操作系统的次版本号
    MajorImageVersion   As Integer              '映像的主版本号。
    MinorImageVersion   As Integer              '映像的次版本号。
    MajorSubsystemVersion   As Integer          '子系统的主版本号
    MinorSubsystemVersion   As Integer          '子系统的次版本号
    Win32VersionValue   As Long                 '保留，必须为0
    SizeOfImage         As Long                 '当映像被加载进内存时的大小（以字节计），包括所有的文件头。它必须是SectionAlignment 的倍数.
    SizeOfHeaders       As Long                 'MS-DOS 占位程序、PE 文件头和节头的总大小，向上舍入为FileAlignment的倍数.
    Checksum            As Long                 '映像文件的校验和。计算校验和的算法被合并到了IMAGEHLP.DLL 中。以下程序在加载时被校验以确定其是否合法: 所有的驱动程序、任何在引导时被加载的DLL 以及加载进关键Windows 进程中的DLL.
    Subsystem           As Integer              '运行此映像所需的子系统
    DllCharacteristics  As Integer              'DLL特征
    SizeOfStackReserve  As Long                 '保留的堆栈大小。只有SizeOfStackCommit 指定的部分被提交；其余的每次可用一页，直到到达保留的大小为止.
    SizeOfStackCommit   As Long                 '提交的堆栈大小。
    SizeOfHeapReserve   As Long                 '保留的局部堆空间大小。只有SizeOfHeapCommit 指定的部分被提交；其余的每次可用一页，直到到达保留的大小为止.
    SizeOfHeapCommit    As Long                 '提交的局部堆空间大小。
    LoaderFlags         As Long                 '保留，必须为0。
    NumberOfRvaAndSizes As Long                 '可选文件头其余部分中数据目录项的个数.每个数据目录描述了一个表的位置和大小.
    DataDirectory(IMAGE_NUMBEROF_DIRECTORY_ENTRIES - 1)     As IMAGE_DATA_DIRECTORY
End Type

Public Type IMAGE_OPTIONAL_HEADER64
    'PE32部分
    Magic               As Integer              '这个无符号整数指出了映像文件的状态。最常用的数字是0x10B，它表明这是一个正常的可执行文件。0x107 表明这是一个ROM 映像，0x20B 表明这是一个PE32 + 可执行文件?
    MajorLinkerVersion  As Byte                 '链接器的主版本号
    MinorLinkerVersion  As Byte                 '链接器的次版本号
    SizeOfCode          As Long                 '代码节（.text）的大小。如果有多个代码节的话，它是所有代码节的和。
    SizeOfInitializedData   As Long             '已初始化数据节的大小。如果有多个这样的数据节的话，它是所有这些数据节的和。
    SizeOfUninitializedData As Long             '未初始化数据节（.bss）的大小。如果有多个.bss 节的话，它是所有这些节的和。
    AddressOfEntryPoint As Long                 '当可执行文件被加载进内存时其入口点相对于映像基址的偏移地址.对于一般程序映像来说，它就是启动地址。对于设备驱动程序来说，它是初始化函数的地址。入口点对于DLL来说是可选的。如果不存在入口点的话，这个域必须为0.
    BaseOfCode          As Long                 '当映像被加载进内存时代码节的开头相对于映像基址的偏移地址。

    ImageBase           As Big_Iint             '当加载进内存时映像的第一个字节的首选地址.它必须是64K 的倍数.DLL默认是0x10000000。Windows CE EXE默认是0x00010000.Windows NT?Windows 2000、Windows XP、Windows 95、Windows 98 和Windows Me 默认是0x00400000。
    SectionAlignment    As Long                 '当加载进内存时节的对齐值（以字节计）。它必须大于或等于FileAlignment.默认是相应系统的页面大小
    FileAlignment       As Long                 '用来对齐映像文件的节中的原始数据的对齐因子（以字节计）。它应该是界于512 和64K 之间的2 的幂（包括这两个边界值）。默认是512。如果SectionAlignment 小于相应系统的页面大小，那么FileAlignment 必须与SectionAlignment 匹配.
    MajorOperatingSystemVersion As Integer      '所需操作系统的主版本号。
    MinorOperatingSystemVersion As Integer      '所需操作系统的次版本号
    MajorImageVersion   As Integer              '映像的主版本号。
    MinorImageVersion   As Integer              '映像的次版本号。
    MajorSubsystemVersion   As Integer          '子系统的主版本号
    MinorSubsystemVersion   As Integer          '子系统的次版本号
    Win32VersionValue   As Long                 '保留，必须为0
    SizeOfImage         As Long                 '当映像被加载进内存时的大小（以字节计），包括所有的文件头。它必须是SectionAlignment 的倍数.
    SizeOfHeaders       As Long                 'MS-DOS 占位程序、PE 文件头和节头的总大小，向上舍入为FileAlignment的倍数.
    Checksum            As Long                 '映像文件的校验和。计算校验和的算法被合并到了IMAGEHLP.DLL 中。以下程序在加载时被校验以确定其是否合法: 所有的驱动程序、任何在引导时被加载的DLL 以及加载进关键Windows 进程中的DLL.
    Subsystem           As Integer              '运行此映像所需的子系统
    DllCharacteristics  As Integer              'DLL特征
    SizeOfStackReserve  As Big_Iint             '保留的堆栈大小。只有SizeOfStackCommit 指定的部分被提交；其余的每次可用一页，直到到达保留的大小为止.
    SizeOfStackCommit   As Big_Iint             '提交的堆栈大小。
    SizeOfHeapReserve   As Big_Iint             '保留的局部堆空间大小。只有SizeOfHeapCommit 指定的部分被提交；其余的每次可用一页，直到到达保留的大小为止.
    SizeOfHeapCommit    As Big_Iint             '提交的局部堆空间大小
    LoaderFlags         As Long                 '保留，必须为0。
    NumberOfRvaAndSizes As Long                 '可选文件头其余部分中数据目录项的个数.每个数据目录描述了一个表的位置和大小.
    DataDirectory(IMAGE_NUMBEROF_DIRECTORY_ENTRIES - 1)     As IMAGE_DATA_DIRECTORY
End Type

Public Type IMAGE_NT_HEADERS64
    Signature           As Long
    FileHeader          As IMAGE_FILE_HEADER
    OptionalHeader      As IMAGE_OPTIONAL_HEADER64
End Type

Public Type IMAGE_NT_HEADERS
    Signature           As Long
    FileHeader          As IMAGE_FILE_HEADER
    OptionalHeader      As IMAGE_OPTIONAL_HEADER32
End Type

'SM4加密
'/**
' * \brief          SM4-ECB block encryption/decryption
' * \param mode     SM4_ENCRYPT or SM4_DECRYPT
' * \param length   length of the input data
' * \param input    input block
' * \param output   output block
' */
Private Declare Function sm4_crypt_ecb Lib "zlSm4.dll" (ByVal Mode As Long, ByVal Length As Long, key As Byte, in_put As Byte, out_put As Byte) As Long
'SM4分组密码加密
'/**
' * \brief          SM4-CBC buffer encryption/decryption
' * \param mode     SM4_ENCRYPT or SM4_DECRYPT
' * \param length   length of the input data
' * \param iv       initialization vector (updated after use)
' * \param input    buffer holding the input data
' * \param output   buffer holding the output data
' */
Private Declare Function sm4_crypt_cbc Lib "zlSm4.dll" (ByVal Mode As Long, ByVal Length As Long, iv As Byte, key As Byte, in_put As Byte, out_put As Byte) As Long
'获取字符串的哈希编码
'/**
' * \brief          Output = SM3( input buffer )
' *
' * \param input    buffer holding the  data
' * \param ilen     length of the input data
' * \param output   SM3 checksum result
' */
Private Declare Sub sm3_hash Lib "zlSm4.dll" Alias "sm3" (in_put As Byte, ByVal Length As Long, out_put As Byte)
'获取文件的sm哈希编码
'/**
' * \brief          Output = SM3( file contents )
' *
' * \param path     input file name
' * \param output   SM3 checksum result
' *
' * \return         0 if successful, 1 if fopen failed,
' *                 or 2 if fread failed
' */
Private Declare Function sm3_file_hash Lib "zlSm4.dll" Alias "sm3_file" (in_path As Byte, out_put As Byte) As Long
'HMAC是密钥相关的哈希运算消息认证码，HMAC运算利用哈希算法，以一个密钥和一个消息为输入，生成一个消息摘要作为输出。
'/**
' * \brief          Output = HMAC-SM3( hmac key, input buffer )
' *
' * \param key      HMAC secret key
' * \param keylen   length of the HMAC key
' * \param input    buffer holding the  data
' * \param ilen     length of the input data
' * \param output   HMAC-SM3 result
' */
Private Declare Sub sm3_hmac_hash Lib "zlSm4.dll" Alias "sm3_hmac" (key As Byte, ByVal keylen As Long, in_put As Byte, ByVal InputLen As Long, out_put As Byte)
'获取ZLSM4的修改版本
'1:只支持sm4_crypt_ecb,sm4_crypt_cbc
'2:增加支持sm3，sm3_file，sm3_hmac，sm_version
'/**
' * \brief          Output = zlSM4.DLL Version
' */
Private Declare Function get_sm_version Lib "zlSm4.dll" Alias "sm_version" () As Long

Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Private Declare Function GetModuleFileName Lib "kernel32" Alias "GetModuleFileNameA" (ByVal hModule As Long, ByVal lpFilename As String, ByVal nSize As Long) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function ProcessIdToSessionId Lib "kernel32.dll" (ByVal dwProcessId As Long, pSessionId As Long) As Long
Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetUserNameEx Lib "Secur32.dll" Alias "GetUserNameExA" (ByVal NameFormat As Long, ByVal lpNameBuffer As String, lpnSize As Long) As Long
Private Declare Function OpenProcessToken Lib "advapi32.dll" (ByVal ProcessHandle As Long, ByVal DesiredAccess As Long, TokenHandle As Long) As Long
Private Declare Function GetTokenInformation Lib "advapi32.dll" (ByVal TokenHandle As Long, ByVal TokenInformationClass As TOKEN_INFORMATION_CLASS, TokenInformation As Long, ByVal TokenInformationLength As Long, ReturnLength As Long) As Long
Private Declare Function LookupAccountSid Lib "advapi32.dll" Alias "LookupAccountSidA" (ByVal lpSystemName As String, ByVal Sid As Long, ByVal Name As String, cbName As Long, ByVal ReferencedDomainName As String, cbReferencedDomainName As Long, peUse As Long) As Long
Private Declare Sub RtlMoveMemory Lib "kernel32" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function EnumProcessModules Lib "psapi.dll" (ByVal hProcess As Long, ByRef lphModule As Long, ByVal Cb As Long, ByRef cbNeeded As Long) As Long
Private Declare Function GetModuleFileNameEx Lib "psapi.dll" Alias "GetModuleFileNameExA" (ByVal hProcess As Long, ByVal hModule As Long, ByVal lpFilename As String, ByVal nSize As Long) As Long
Private Declare Function WTSQuerySessionInformation Lib "wtsapi32" Alias "WTSQuerySessionInformationW" (ByVal hServer As Long, ByVal SessionID As Long, ByVal WTSInfoClass As Long, ppBuffer As Long, pBytesReturned As Long) As Long
Private Declare Sub WTSFreeMemory Lib "wtsapi32" (ByVal pMemory As Long)
Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
Private Declare Function IsWow64Process Lib "kernel32" (ByVal hProc As Long, bWow64Process As Long) As Long
Private Declare Function RegEnumKey Lib "advapi32.dll" Alias "RegEnumKeyA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, ByVal cbName As Long) As Long
Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal uloptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegQueryValueEx_Long Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Long, lpcbData As Long) As Long
Private Declare Function RegQueryValueEx_String Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
Private Declare Function RegQueryValueEx_BINARY Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Byte, lpcbData As Long) As Long
Private Declare Function RegSetValueEx_String Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpData As String, ByVal lpcbData As Long) As Long
Private Declare Function RegSetValueEx_Long Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpValue As Long, ByVal cbData As Long) As Long
Private Declare Function RegSetValueEx_BINARY Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpValue As Byte, ByVal cbData As Long) As Long
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Private Declare Function RegQueryValueEx_ValueType Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function ExpandEnvironmentStrings Lib "kernel32" Alias "ExpandEnvironmentStringsA" (ByVal lpSrc As String, ByVal lpDst As String, ByVal nSize As Long) As Long
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, lpWideCharStr As Any, ByVal cchWideChar As Long, lpMultiByteStr As Any, ByVal cchMultiByte As Long, lpDefaultChar As Any, ByVal lpUsedDefaultChar As Long) As Long

Private M_SM4_VERSION As Long

Public Function GetServerInfo(ByVal varServer As Variant) As String
'功能：获取IP:Port/SID信息
    Dim strServerInfo As String, strProp As String, strIp As String, strPort As String, strSID As String
    Dim arrTmp As Variant

    If IsObject(varServer) Then
        If TypeOf varServer Is ADODB.Connection Then
            If IsOLEDBConnection(varServer) Then
                '(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=192.168.0.60)(PORT=1522))(CONNECT_DATA=(SERVICE_NAME=qzyy)))
                'Testbase
                strProp = "Data Source Name"
                GoSub mak_GetProperty
            Else
                'Driver={Microsoft ODBC for Oracle};Server=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=192.168.0.60)(PORT=1522))(CONNECT_DATA=(SERVICE_NAME=qzyy)))
                'Driver={Microsoft ODBC for Oracle};Server=Testbase
                strProp = "Extended Properties"
                GoSub mak_GetProperty
                If strServerInfo <> "" Then
                    strServerInfo = UCase(Trim(Mid(strServerInfo, InStrRev(strServerInfo, "SERVER=") + Len("SERVER="))))
                End If
            End If
        End If
    Else
        strServerInfo = varServer
    End If
    
    If LenB(strServerInfo) > 0 Then
        If InStr(strServerInfo, "=") = 0 Then
            If InStr(strServerInfo, "/") > 0 Then
                arrTmp = Split(strServerInfo, "/")
                strSID = arrTmp(1)
                If InStr(arrTmp(0), ":") > 0 Then
                    arrTmp = Split(arrTmp(0), ":")
                    strIp = arrTmp(0)
                    strPort = arrTmp(1)
                Else
                    strIp = arrTmp(0)
                    strPort = "1521"
                End If
                GetServerInfo = strIp & ":" & strPort & "/" & strSID
            Else
                Call GetServerInfoByFile(strServerInfo, strSID, strIp, strPort)
                If strSID <> "" And strIp <> "" And strPort <> "" Then
                    GetServerInfo = strIp & ":" & strPort & "/" & strSID
                Else
                    GetServerInfo = strServerInfo
                End If
            End If
        Else
            If InStr(strServerInfo, "HOST=") > 0 Then
                strIp = Mid(strServerInfo, InStr(strServerInfo, "HOST=") + Len("HOST="))
                strIp = Trim(Mid(strIp, 1, InStr(strIp, ")") - 1))
            End If
            If InStr(strServerInfo, "PORT=") > 0 Then
                strPort = Mid(strServerInfo, InStr(strServerInfo, "PORT=") + Len("PORT="))
                strPort = Trim(Mid(strPort, 1, InStr(strPort, ")") - 1))
            End If
            If InStr(strServerInfo, "(SID=") > 0 Then
                strSID = Mid(strServerInfo, InStr(strServerInfo, "(SID=") + Len("(SID="))
                strSID = Trim(Mid(strSID, 1, InStr(strSID, ")") - 1))
            ElseIf InStr(strServerInfo, "(SERVICE_NAME=") > 0 Then
                strSID = Mid(strServerInfo, InStr(strServerInfo, "(SERVICE_NAME=") + Len("(SERVICE_NAME="))
                strSID = Trim(Mid(strSID, 1, InStr(strSID, ")") - 1))
            End If
            GetServerInfo = strIp & ":" & strPort & "/" & strSID
        End If
    End If
    Exit Function
    
mak_GetProperty:
    On Error Resume Next
    strServerInfo = UCase(Trim(Replace(varServer.Properties(strProp), " ", "")))
    If Err.Number <> 0 Then strServerInfo = ""
    On Error GoTo 0
    Return
End Function

Public Function CurrentPID() As Long
    CurrentPID = GetCurrentProcessId
End Function

Public Function ProcessUserFullName(Optional ByVal lngProcessID As Long) As String
    Dim strUserName As String, strDomain As String, strFullUser As String
    
    On Error GoTo ErrH
    If ProcessUserInfo_Sub(lngProcessID, strUserName, strDomain, strFullUser) Then
        ProcessUserFullName = strFullUser
    End If
    Exit Function
    
ErrH:
    Err.Clear
End Function

Public Function SessionID(Optional ByVal lngProcessID As Long) As Long
    Dim lngSessionID            As Long
    Dim lngTmpProcessID         As Long
    On Error GoTo ErrH
    If lngProcessID = 0 Then
        lngTmpProcessID = GetCurrentProcessId
    Else
        lngTmpProcessID = lngProcessID
    End If
    If lngTmpProcessID <> 0 Then
        SessionID = -1
        If ProcessIdToSessionId(lngTmpProcessID, lngSessionID) = 0 Then
            '...
        Else
            SessionID = lngSessionID
        End If
    End If
    Exit Function
ErrH:
    Err.Clear
End Function

Public Function SessionUserFullName(Optional ByVal lngSeesionID As Long = -1) As String
    Dim strTmp        As String
    
    strTmp = SessionDomain(lngSeesionID)
    If LenB(strTmp) <> 0 Then
        SessionUserFullName = strTmp & "\" & SessionUser(lngSeesionID)
    End If
End Function

Private Function ProcessUserInfo_Sub(ByVal lngProcessID As Long, strUserName As String, strDomainName As String _
    , strFullUserName As String) As Boolean
    
    Dim lngLen      As Long
    Dim arrTmp      As Variant
    Dim strTmp      As String
    Dim lngProcess  As Long, hToken         As Long
    Dim BufferSize  As Long, InfoBuffer()   As Byte
    Dim tkUser      As TOKEN_USER
    Dim lngUserSize As Long, lngDomainSize      As Long, pUse   As Long
    
    On Error GoTo ErrH
    strFullUserName = ""
    strUserName = ""
    strDomainName = ""
    If lngProcessID = 0 Then
        strTmp = String(MAX_PATH, Chr$(0))
        lngLen = MAX_PATH
        If GetUserName(strTmp, lngLen) = 0 Then
        Else
            strTmp = mdlPublic.TruncZero(strTmp)
            If LenB(strTmp) <> 0 Then
                strUserName = strTmp
                
                strTmp = String(MAX_PATH, Chr$(0))
                lngLen = MAX_PATH
                '普通用户下，返回，电脑名\用户名，SYSTEM用户下返回工作组\电脑名
                If GetUserNameEx(NameSamCompatible, strTmp, lngLen) = 0 Then
                Else
                    strTmp = mdlPublic.TruncZero(strTmp)
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
            End If
        End If
    Else
        lngProcess = OpenProcess(PROCESS_QUERY_INFORMATION, 0, lngProcessID)
        If lngProcess = 0 Then
        Else
            If OpenProcessToken(lngProcess, TOKEN_QUERY, hToken) = 0 Then
            Else
                If GetTokenInformation(hToken, ByVal TokenUser, 0, 0, BufferSize) = 0 Then ' Determine required buffer size
                End If
                If BufferSize Then
                    ReDim InfoBuffer(BufferSize - 1)
                    If GetTokenInformation(hToken, ByVal TokenUser, ByVal VarPtr(InfoBuffer(0)), BufferSize, BufferSize) = 0 Then
                    Else
                        Call RtlMoveMemory(tkUser, InfoBuffer(0), LenB(tkUser))
                        strUserName = String(MAX_PATH, Chr(0))
                        strDomainName = String(MAX_PATH, Chr(0))
                        lngUserSize = MAX_PATH
                        lngDomainSize = MAX_PATH
                        If LookupAccountSid(vbNullString, tkUser.User.Sid, strUserName, lngUserSize, strDomainName, lngDomainSize, pUse) = 0 Then
                            strDomainName = ""
                            strUserName = ""
                        Else
                            strUserName = mdlPublic.TruncZero(strUserName)
                            strDomainName = mdlPublic.TruncZero(strDomainName)
                            strFullUserName = strDomainName & "\" & strUserName
                            ProcessUserInfo_Sub = True
                        End If
                    End If
                End If
                If CloseHandle(hToken) = 0 Then
                    ProcessUserInfo_Sub = False
                Else
                    hToken = 0
                End If
            End If
            If CloseHandle(lngProcess) = 0 Then
                ProcessUserInfo_Sub = False
            Else
                lngProcess = 0
            End If
        End If
    End If
    Exit Function
    
ErrH:
    If hToken <> 0 Then
        If CloseHandle(hToken) = 0 Then
        End If
    End If
    If lngProcess <> 0 Then
        If CloseHandle(lngProcess) = 0 Then
        End If
    End If
    Err.Clear
End Function

Public Function AppsoftPath() As String
    If IsDesinModeEx Then
        AppsoftPath = "C:\APPSOFT"
    Else
        AppsoftPath = Mid(App.Path & "\", 1, InStr(5, App.Path & "\", "\"))
        If Right(AppsoftPath, 1) = "\" Then AppsoftPath = Mid(AppsoftPath, 1, Len(AppsoftPath) - 1)
    End If
End Function

Public Function IsDesinModeEx() As Boolean
    Dim strStartPath As String
    Dim objFSO As FileSystemObject
    
    On Error GoTo ErrH
    Set objFSO = New FileSystemObject
    strStartPath = StartExePath()
    If UCase(Trim(objFSO.GetFileName(strStartPath))) = "VB6.EXE" Then
        IsDesinModeEx = True
    End If
    Exit Function
ErrH:
    Err.Clear
End Function

Public Function StartExePath(Optional ByVal lngProcessID As Long) As String
    Dim strFile         As String
    Dim lngRet          As Long
    Dim hProcess        As Long
    Dim hModule         As Long
    
    On Error GoTo ErrH
    If lngProcessID = 0 Then
        strFile = String(MAX_PATH, Chr(0))
        
        lngRet = GetModuleFileName(0, strFile, MAX_PATH)
        If Err.LastDllError <> 0 Then
        Else
            StartExePath = mdlPublic.TruncZero(strFile)
        End If
    Else
        hProcess = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ, 0, lngProcessID)
        If hProcess = 0 Then
        Else
            If EnumProcessModules(hProcess, hModule, 4&, 0&) = 0 Then
            Else
                strFile = String(MAX_PATH, Chr(0))
                lngRet = GetModuleFileNameEx(hProcess, hModule, strFile, MAX_PATH)
                If Err.LastDllError <> 0 Then
                Else
                    StartExePath = mdlPublic.TruncZero(strFile)
                End If
            End If
            If CloseHandle(hProcess) = 0 Then
            Else
                hProcess = 0
            End If
        End If
    End If
    Exit Function
ErrH:
    If hProcess <> 0 Then
        If CloseHandle(hProcess) = 0 Then
        End If
    End If
    Err.Clear
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

Public Function IsOLEDBConnection(ByVal cnMain As ADODB.Connection) As Boolean
'功能：判断当前连接是否是OraOLEDB连接
'根据Provider来判断，存在两种方式
'方式一：'Provider=OraOLEDB.Oracle.1;Password=HIS;Persist Security Info=True;User ID=ZLHIS;Data Source="DYYY";Extended Properties="PLSQLRSet=1"
'方式二：
'.Provider = "OraOLEDB.Oracle"
'.Open "PLSQLRSet=1;Data Source=" & strServer & strPersist_Security_Info, strUserName, strPassWord
'这两种方式均会自动设置.Provider属性
    '使用Like是因为可能后面增加版本如OraOLEDB.Oracle.1
    If UCase(cnMain.Provider) Like "ORAOLEDB.ORACLE*" Then
        IsOLEDBConnection = True
    End If
End Function

Private Sub GetServerInfoByFile(ByVal strServer As String, ByRef setServiceName As String, strServerIp As String, ByRef strServerPort As String)
'功能:根据tnsname.ora文件获取服务器IP、端口、实例名
'传入参数: strServer=服务名
'传出参数 setServiceName = 实例名  strServerIp = 服务器IP   strServerPort = 服务器端口

    Dim strTxt      As String, strFile As String
    Dim lngTmp      As Long, strTmp As String
    Dim lngIndex    As Long, lngPos As Long, i  As Long
    Dim objFSO As FileSystemObject
    
    On Error Resume Next
    
    Set objFSO = New FileSystemObject
    strFile = GetOracleHome()
    If strFile = "" Then Exit Sub
    strFile = strFile & "\network\ADMIN\tnsnames.ora"
    If Not objFSO.FileExists(strFile) Then Exit Sub
    
    strTxt = objFSO.OpenTextFile(strFile).ReadAll
    strServer = UCase(strServer): strTxt = ConvertStr(strTxt) '格式化字符
    strTxt = Mid(strTxt, InStr(1, strTxt, strServer & "="))
    lngIndex = 0
    lngPos = 1
    lngPos = InStr(lngPos, strTxt, "(")
    If lngPos <> 0 Then
        For i = lngPos To Len(strTxt)
            Select Case Mid(strTxt, i, 1)
                Case "("
                    lngIndex = lngIndex + 1
                Case ")"
                    lngIndex = lngIndex - 1
            End Select
            If lngIndex = 0 Then
                Exit For
            End If
        Next
        If lngIndex = 0 Then
            strTxt = Mid(strTxt, 1, i)
        End If
        '获取IP
        lngTmp = InStr(1, strTxt, "HOST=")
        strTmp = Mid(strTxt, lngTmp + Len("HOST="))
        strServerIp = Mid(strTmp, 1, InStr(1, strTmp, ")") - 1)
        
        '获取端口
        lngTmp = InStr(1, strTxt, "PORT=")
        strTmp = Mid(strTxt, lngTmp + Len("PORT="))
        strServerPort = Mid(strTmp, 1, InStr(1, strTmp, ")") - 1)
        
        '获取服务名
        lngTmp = InStr(1, strTxt, "SERVICE_NAME=")
        If lngTmp > 0 Then
            strTmp = Mid(strTxt, lngTmp + Len("SERVICE_NAME="))
        Else
            lngTmp = InStr(1, strTxt, "SID=")
            strTmp = Mid(strTxt, lngTmp + Len("SID="))
        End If
        
        setServiceName = Mid(strTmp, 1, InStr(1, strTmp, ")") - 1)
    End If
End Sub

Public Function SessionDomain(Optional ByVal lngSeesionID As Long = -1) As String
    Dim pBuffer         As Long
    Dim dwBufferLen     As Long
    Dim arrBytRet()     As Byte
    Dim strDomain       As String
    Dim blnNew          As Boolean
    
    On Error GoTo ErrH

    If WTSQuerySessionInformation(WTS_CURRENT_SERVER_HANDLE, lngSeesionID, WTSDomainName, pBuffer, dwBufferLen) <> 0 Then
        If dwBufferLen <> 0 Then
            ReDim Preserve arrBytRet(dwBufferLen - 1)
            RtlMoveMemory ByVal VarPtr(arrBytRet(0)), ByVal pBuffer, dwBufferLen
            strDomain = arrBytRet
            strDomain = mdlPublic.TruncZero(strDomain)
            
            SessionDomain = strDomain
        End If
        WTSFreeMemory pBuffer
    Else
    End If
    Exit Function
ErrH:
    Err.Clear
End Function

Public Function SessionUser(Optional ByVal lngSeesionID As Long = -1) As String
    Dim pBuffer         As Long
    Dim dwBufferLen     As Long
    Dim arrBytRet()     As Byte
    Dim strUserName     As String
    Dim blnNew          As Boolean
    
    On Error GoTo ErrH
    If WTSQuerySessionInformation(WTS_CURRENT_SERVER_HANDLE, lngSeesionID, WTSUserName, pBuffer, dwBufferLen) <> 0 Then
        If dwBufferLen <> 0 Then
            ReDim Preserve arrBytRet(dwBufferLen - 1)
            RtlMoveMemory ByVal VarPtr(arrBytRet(0)), ByVal pBuffer, dwBufferLen
            strUserName = arrBytRet
            strUserName = mdlPublic.TruncZero(strUserName)
            
            SessionUser = strUserName
        End If
        WTSFreeMemory pBuffer
    Else
    End If
    Exit Function
ErrH:
    Err.Clear
End Function

Public Function GetOracleHome() As String
'功能：获取OracleHome路径
    Dim arrTmp  As Variant, arrSubKey   As Variant
    Dim strHome As String, strDefault   As String, strPath As String
    Dim i       As Integer
    Dim objPE   As New clsPEReader
    Dim blnRead As Boolean
    Dim objFSO As New FileSystemObject
    
    strHome = Environ("PATH")
    '1、PATH变量都没有，操作系统的环境变量存在问题或者非WIn系统，可能为麦金塔系统（MAC）
    If strHome = "" Then Exit Function
    arrTmp = Split(strHome, ";")
    strHome = ""
    For i = LBound(arrTmp) To UBound(arrTmp)
    
        If UCase(arrTmp(i)) Like "*ORA*\BIN" Then
            '判断Oracle的OCI基础部件是否存在
            If objFSO.FileExists(arrTmp(i) & "\oci.dll") Then
                If Not objPE.Is64bit(arrTmp(i) & "\oci.dll") Then
                    strHome = objFSO.GetParentFolderName(arrTmp(i))
                    If objFSO.FileExists(strHome & "\network\ADMIN\tnsnames.ora") Then
                        GetOracleHome = strHome
                        Exit Function
                    End If
                End If
            End If
        End If
    Next
    '2、寻找TNS_ADMIN:ORACLE_HOME & "\network\ADMIN
    strHome = Environ("TNS_ADMIN")
    If strHome <> "" Then
        If InStr(UCase(strHome), "\NETWORK\ADMIN") > 0 Then
            '判断TNSNAME
            If Not objFSO.FileExists(strHome & "\tnsnames.ora") Then
                strHome = ""
            End If
            '获取ORACLE_HOME,判断OCI
            If strHome <> "" Then
                strHome = objFSO.GetParentFolderName(objFSO.GetParentFolderName(strHome))
                If objFSO.FileExists(strHome & "\Bin\oci.dll") Then
                    If Not objPE.Is64bit(strHome & "\Bin\oci.dll") Then
                        GetOracleHome = strHome
                        Exit Function
                    End If
                End If
            End If
        End If
    End If
    '3、ORACLE_HOME环境变量
    strHome = Environ("ORACLE_HOME")
    If strHome <> "" Then
        If objFSO.FileExists(strHome & "\Bin\oci.dll") Then
            If Not objPE.Is64bit(strHome & "\Bin\oci.dll") Then
                If objFSO.FileExists(strHome & "\network\ADMIN\tnsnames.ora") Then
                    GetOracleHome = strHome
                    Exit Function
                End If
            End If
        End If
    End If
    
    '4、注册表判断,读取64位下32目录会自动定位到SOFTWARE\Wow6432Node\Oracle 2：读取32位下32位目录
    '4.1 ALL_HOMES
    '         DEFAULT_HOME"="DEFAULT_HOME"
    '      ALL_HOMES\ID0
    '        "NAME"="DEFAULT_HOME"
    '        "PATH"="F:\\instantclient_11_2_3"
    blnRead = mdlPublic.GetRegValue("HKEY_LOCAL_MACHINE\SOFTWARE\" & IIf(mdlPublic.Is64bit, "WOW6432Node\", "") & "Oracle\ALL_HOMES", "DEFAULT_HOME", strDefault)
    If blnRead And strDefault <> "" Then
        arrSubKey = mdlPublic.GetAllSubKey("HKEY_LOCAL_MACHINE\SOFTWARE\" & IIf(mdlPublic.Is64bit, "WOW6432Node\", "") & "Oracle\ALL_HOMES")
        If TypeName(arrSubKey) <> "Empty" Then
            For i = LBound(arrSubKey) To UBound(arrSubKey)
                strHome = ""
                blnRead = mdlPublic.GetRegValue("HKEY_LOCAL_MACHINE\SOFTWARE\" & IIf(mdlPublic.Is64bit, "WOW6432Node\", "") & "Oracle\ALL_HOMES\" & arrSubKey(i), "NAME", strHome)
                If blnRead And strHome <> "" Then
                    If strHome = strDefault Then
                        blnRead = mdlPublic.GetRegValue("HKEY_LOCAL_MACHINE\SOFTWARE\" & IIf(mdlPublic.Is64bit, "WOW6432Node\", "") & "Oracle\ALL_HOMES\" & arrSubKey(i), "PATH", strPath)
                        If blnRead And strPath <> "" Then
                            If Not objPE.Is64bit(strPath & "\Bin\oci.dll") Then
                                If objFSO.FileExists(strPath & "\network\ADMIN\tnsnames.ora") Then
                                    GetOracleHome = strPath
                                    Exit Function
                                End If
                            End If
                        End If
                    End If
                End If
            Next
        End If
    End If
    '4.2非ALL_Homes方式,只获取第一个符合条件的。
    arrSubKey = Empty
    arrSubKey = mdlPublic.GetAllSubKey("HKEY_LOCAL_MACHINE\SOFTWARE\" & IIf(mdlPublic.Is64bit, "WOW6432Node\", "") & "Oracle")
    If TypeName(arrSubKey) <> "Empty" Then
        For i = LBound(arrSubKey) To UBound(arrSubKey)
            strHome = ""
            blnRead = mdlPublic.GetRegValue("HKEY_LOCAL_MACHINE\SOFTWARE\" & IIf(mdlPublic.Is64bit, "WOW6432Node\", "") & "Oracle\" & arrSubKey(i), "ORACLE_HOME", strHome)
            If blnRead And strHome <> "" Then
                If Not objPE.Is64bit(strHome & "\Bin\oci.dll") Then
                    If objFSO.FileExists(strHome & "\network\ADMIN\tnsnames.ora") Then
                        GetOracleHome = strHome
                        Exit Function
                    End If
                End If
            End If
        Next
    End If
End Function

Private Function ConvertStr(ByVal strSource As String) As String
    '功能:去掉字符串的空格\换行符,并转换为大写
    
    strSource = UCase(strSource)
    strSource = Replace(strSource, " ", "")
    strSource = Replace(strSource, vbNewLine, "")
    strSource = Replace(strSource, vbCr, "")
    strSource = Replace(strSource, vbLf, "")
    strSource = Replace(strSource, vbTab, "")
    strSource = Replace(strSource, vbBack, "")
    ConvertStr = strSource
End Function

Public Function Is64bit() As Boolean
    '******************************************************************************************************************
    '功能：是否是64位系统
    '返回：
    '******************************************************************************************************************
    Dim Handle As Long
    Dim lngFunc As Long
        
    lngFunc = 0
    Handle = GetProcAddress(GetModuleHandle("kernel32"), "IsWow64Process")
    If Handle > 0 Then
        IsWow64Process GetCurrentProcess(), lngFunc
    End If
    Is64bit = lngFunc <> 0
End Function

Public Function GetRegValue(ByVal strKey As String, ByVal strValueName As String, ByRef varValue As Variant, Optional blnOneString As Boolean = False) As Boolean
'功能：获取注册表中指定位置的值
'参数：strKey=注册表键位，如“HKEY_CURRENT_USER\Printers\DevModePerUser"
'          strValueName=变量名
'          strValue=变量值
'          strValueType=变量类型，默认为字符串
'           blnOneString = 对REG_EXPAND_SZ、REG_MULTI_SZ,REG_BINARY有效。-  True 则函数返回单一字符串，且不经任何处理，只去掉字符串尾！
'返回：是否读取成功
'说明：当前只对REG_SZ, REG_EXPAND_SZ, REG_MULTI_SZ，REG_DWORD，REG_BINARY实现了读取。没有查询到可以自动查找键名
    Dim hRootKey As REGRoot, strSubKey As String
    Dim lngReturn As Long
    Dim lngKey As Long, ruType As REGValueType
    Dim lngLength As Long, strBufVar() As String, lngBuf As Long, bytBuf() As Byte, strBuf As String
    Dim i As Long, strReturn As String, strTmp As String
    '不是有效的注册表键位,获取键名类型
    If Not GetKeyValueInfo(strKey, strValueName, hRootKey, strSubKey, ruType) Then Exit Function
    '打开变量
    lngReturn = RegOpenKeyEx(hRootKey, strSubKey, 0, KEY_QUERY_VaLUE, lngKey)
    If lngReturn <> ERROR_SUCCESS Then
        Exit Function
    End If
    On Error GoTo ErrH
    Select Case ruType
        Case REG_SZ, REG_EXPAND_SZ, REG_MULTI_SZ '字符串类型读取
'            lngReturn = RegQueryValueEx(lngKey, strValueName, 0, ruType, 0, lngLength)
'            If lngReturn <> ERROR_SUCCESS Then Err.Clear '可能出错，因此这样处理
            lngLength = 1024: strBuf = Space(lngLength)
            lngReturn = RegQueryValueEx_String(lngKey, strValueName, 0, ruType, strBuf, lngLength)
            If lngReturn <> ERROR_SUCCESS Then: RegCloseKey (lngKey): Exit Function
            Select Case ruType
                Case REG_SZ
                    varValue = mdlPublic.TruncZero(strBuf)
                Case REG_EXPAND_SZ ' 扩充环境字符串，查询环境变量和返回定义值
                    If Not blnOneString Then
                        varValue = mdlPublic.TruncZero(ExpandEnvStr(mdlPublic.TruncZero(strBuf)))
                    Else
                        varValue = mdlPublic.TruncZero(strBuf)
                    End If
                Case REG_MULTI_SZ ' 多行字符串
                    If Not blnOneString Then
                        If Len(strBuf) <> 0 Then ' 读到的是非空字符串，可以分割。
                            strBufVar = Split(Left$(strBuf, Len(strBuf) - 1), Chr$(0))
                        Else ' 若是空字符串，要定义S(0) ，否则出错！
                            ReDim strBufVar(0) As String
                        End If
                        ' 函数返回值，返回一个字符串数组？！
                        varValue = strBufVar()
                    Else
                        varValue = mdlPublic.TruncZero(strBuf)
                    End If
            End Select
        Case REG_DWORD
            lngReturn = RegQueryValueEx_Long(lngKey, strValueName, ByVal 0&, ruType, lngBuf, Len(lngBuf))
            If lngReturn <> ERROR_SUCCESS Then: RegCloseKey (lngKey): varValue = 0: Exit Function
            varValue = lngBuf
        Case REG_BINARY
            lngReturn = RegQueryValueEx_BINARY(lngKey, strValueName, 0, ruType, ByVal 0, lngLength)
            If lngReturn <> ERROR_SUCCESS And lngReturn <> ERROR_MORE_DATA Then
                RegCloseKey lngKey: Exit Function
                If blnOneString Then
                    varValue = "00"
                Else
                    ReDim bytBuf(0)
                    varValue = bytBuf()
                End If
            End If
            ReDim bytBuf(lngLength - 1)
            lngReturn = RegQueryValueEx_BINARY(lngKey, strValueName, 0, ruType, bytBuf(0), lngLength)
            If lngReturn <> ERROR_SUCCESS And lngReturn <> ERROR_MORE_DATA Then
                RegCloseKey lngKey: Exit Function
                If blnOneString Then
                    varValue = "00"
                Else
                    ReDim bytBuf(0)
                    varValue = bytBuf()
                End If
            End If
            If lngLength <> UBound(bytBuf) + 1 Then
               ReDim Preserve bytBuf(0 To lngLength - 1) As Byte
            End If
            ' 返回字符串，注意：要将字节数组进行转化！
            If blnOneString Then
                '循环数据，把字节转换为16进制字符串
                For i = LBound(bytBuf) To UBound(bytBuf)
                   strTmp = CStr(Hex(bytBuf(i)))
                   If (Len(strTmp) = 1) Then strTmp = "0" & strTmp
                   strReturn = strReturn & " " & strTmp
                Next i
                varValue = Trim$(strReturn)
            Else
                varValue = bytBuf()
            End If
    End Select
    RegCloseKey lngKey
    GetRegValue = True
    Exit Function
ErrH:
    If 0 = 1 Then
        Resume
    End If
End Function

Public Function GetAllSubKey(ByVal strKey As String) As Variant
'功能:获取某项的所有子项
'返回：=子项数组
    Dim lnghKey As Long, lngRet As Long, strName As String, lngIdx As Long
    Dim hRootKey As Long, strKeyName As String
    Dim strSubKey As Variant
    strSubKey = Array()
    lngIdx = 0: strName = String(256, Chr(0))
     If Not GetKeyValueInfo(strKey, "", hRootKey, strKeyName) Then Exit Function
    lngRet = RegOpenKey(hRootKey, strKeyName, lnghKey)
    If lngRet = 0 Then
        Do
            lngRet = RegEnumKey(lnghKey, lngIdx, strName, Len(strName))
            If lngRet = 0 Then
                ReDim Preserve strSubKey(UBound(strSubKey) + 1)
                strSubKey(UBound(strSubKey)) = Left(strName, InStr(strName, Chr(0)) - 1)
                lngIdx = lngIdx + 1
            End If
        Loop Until lngRet <> 0
    End If
    RegCloseKey lnghKey
    GetAllSubKey = strSubKey
End Function

Private Function GetKeyValueInfo(ByVal strKey As String, Optional ByVal strValueName As String, Optional ByRef hRootKey As REGRoot, Optional ByRef strSubKey As String, Optional ByRef lngType As Long) As Boolean
'功能：根据键位获取根键值与子健,以及值类型
'参数：strKey=注册表键位，如“HKEY_CURRENT_USER\Printers\DevModePerUser"
'          strValueName=变量名
'出参：
'          hRootKey=根键
'          strSubKey=子健
'          lngType=键类型
'返回：是否获取成功
    Dim strRoot As String, lngPos As String, hKey As Long
    Dim lngReturn As Long, strName As String * 255
    
    On Error GoTo ErrH
    hRootKey = 0: strSubKey = "": lngType = 0
    lngPos = InStr(strKey, "\")
    If lngPos = 0 Then Exit Function
    strRoot = Mid(strKey, 1, lngPos - 1)
    strSubKey = Mid(strKey, lngPos + 1)
    
    hRootKey = mdlPublic.Decode(UCase(strRoot), "HKEY_CLASSES_ROOT", HKEY_CLASSES_ROOT, _
                                                                         "HKEY_CURRENT_USER", HKEY_CURRENT_USER, _
                                                                         "HKEY_LOCAL_MACHINE", HKEY_LOCaL_MaCHINE, _
                                                                         "HKEY_USERS", HKEY_USERS, _
                                                                         "HKEY_PERFORMANCE_DATA", HKEY_PERFORMANCE_DATA, _
                                                                         "HKEY_CURRENT_CONFIG", HKEY_CURRENT_CONFIG, _
                                                                         "HKEY_DYN_DATA", HKEY_DYN_DATA, 0)
    If hRootKey = 0 Then Exit Function
    If lngType <> -1 Then
        '使用查询方式打开，进行键名类型查询
        lngReturn = RegOpenKeyEx(hRootKey, strSubKey, 0, KEY_QUERY_VaLUE, hKey)
        If lngReturn <> ERROR_SUCCESS Then
            Exit Function
        End If
        If strValueName <> "" Then
            lngReturn = RegQueryValueEx_ValueType(hKey, strValueName, ByVal 0&, lngType, ByVal strName, Len(strName))
            'SetRegKey这种情况返回的类型为很大的数，数值不固定,因此设置为0，根据传入数据类型判断
            If lngReturn = ERROR_BADKEY Then
                If lngType < REG_NONE Or lngType > REG_MULTI_SZ Then lngType = REG_NONE
            End If
            '可能字段超长，长度不够，所以出错不退出
            'If lngReturn <> ERROR_SUCCESS Then: RegCloseKey (hKey): Exit Function
        End If
        RegCloseKey (hKey)
    End If
    GetKeyValueInfo = True
    Exit Function
ErrH:
    If 0 = 1 Then
        Resume
    End If
    Err.Clear
End Function

Public Function ExpandEnvStr(ByVal strInput As String) As String
'功能：将字符串中的环境变量替换为常规值
'         strInput=包含环境变量的字符串
'返回：用实际的值替换字符串中的环境变量后的字符串
    '// 如： %PATH% 则返回 "c:\;c:\windows;"
    Dim lngLen As Long, strBuf As String, strOld As String
    strOld = strInput & "  " ' 不知为什么要加两个字符，否则返回值会少最后两个字符！
    strBuf = "" '// 不支持Windows 95
    '// get the length
    lngLen = ExpandEnvironmentStrings(strOld, strBuf, lngLen)
    '// 展开字符串
    strBuf = String$(lngLen - 1, Chr$(0))
    lngLen = ExpandEnvironmentStrings(strOld, strBuf, LenB(strBuf))
    '// 返回环境变量
    ExpandEnvStr = mdlPublic.TruncZero(strBuf)
End Function

Public Function Decode(ParamArray arrPar() As Variant) As Variant
'功能：模拟Oracle的Decode函数
    Dim varValue As Variant, i As Integer
    
    i = 1
    varValue = arrPar(0)
    Do While i <= UBound(arrPar)
        If i = UBound(arrPar) Then
            Decode = arrPar(i): Exit Function
        ElseIf varValue = arrPar(i) Then
            Decode = arrPar(i + 1): Exit Function
        Else
            i = i + 2
        End If
    Loop
End Function

Public Function FromatSQL(ByVal strText As String, Optional ByVal blnCrlf As Boolean) As String
'功能：去掉TAB字符，两边空格，回车，最后只由单空格分隔。
'参数：strText=处理字符
'         blnCrlf=是否去掉换行符
    Dim i As Long
    
    If blnCrlf Then
        strText = Replace(strText, vbCrLf, " ")
        strText = Replace(strText, vbCr, " ")
        strText = Replace(strText, vbLf, " ")
    End If
    strText = Trim(Replace(strText, vbTab, " "))
    
    i = 5
    Do While i > 1
        strText = Replace(strText, String(i, " "), " ")
        If InStr(strText, String(i, " ")) = 0 Then i = i - 1
    Loop
    FromatSQL = strText
End Function

'======================================================================================================================
'方法           Sm4DecryptEcb           SM4解密
'返回值         String                  解密后的值
'入参列表:
'参数名         类型                    说明
'strInput       String                  要解密的字符串（该字符串是Sm4EncryptEcb生成的结果）
'strKey         String(Optional)        加密密钥也就是解密密钥（32位的16进制字符串，可以通过HexStringToByte返回）
'======================================================================================================================
Public Function Sm4DecryptEcb(ByVal strInput As String, Optional ByVal strKey As String) As String
    Dim arrKey()        As Byte
    Dim arrInput()      As Byte
    Dim arrOutPut()     As Byte
    Dim lngVersion      As Long

    If M_SM4_VERSION = 0 Then
        M_SM4_VERSION = sm_version
    End If
    If strInput Like "ZLSV*:*" Then
        lngVersion = Val(Mid(strInput, 5, InStr(strInput, ":") - 5))
        strInput = Mid(strInput, InStr(strInput, ":") + 1)
        '当前客户端的ZLSM4不支持该版本的加密字符串解密，仍旧解密，因为一般来说都能解密出相同的字符串
'        If lngVersion > M_SM4_VERSION Then
'            Exit Function
'        End If
    Else
        Exit Function
    End If
    
    arrKey = GetKey(strKey, 2)
    arrInput = HexStringToByte(strInput)
    ReDim arrOutPut(UBound(arrInput))
    
    Call sm4_crypt_ecb(CM_Decrypt, UBound(arrInput) + 1, arrKey(0), arrInput(0), arrOutPut(0))
    If lngVersion = 1 Then
        Sm4DecryptEcb = Trim(StrConv(arrOutPut(), vbUnicode))
    Else
        Sm4DecryptEcb = TruncZeroInside(StrConv(arrOutPut(), vbUnicode))
    End If
End Function
'======================================================================================================================
'方法           Sm4EncryptCbc           SM4分组加密
'返回值         String                  加密后的值
'入参列表:
'参数名         类型                    说明
'strInput       String                  要加密的字符串
'strKey         String(Optional)        加密密钥（32位的16进制字符串，可以通过HexStringToByte返回）
'strIv          String(Optional)        分组加密密钥（32位的16进制字符串，可以通过HexStringToByte返回）
'======================================================================================================================
Public Function Sm4EncryptCbc(ByVal strInput As String, Optional ByVal strKey As String, Optional ByVal strIv As String) As String
    Dim arrKey()        As Byte
    Dim arrInput()      As Byte
    Dim arrOutPut()     As Byte
    Dim arrIv()         As Byte
    
    If M_SM4_VERSION = 0 Then
        M_SM4_VERSION = sm_version
    End If
    If strInput = "" Then
        Sm4EncryptCbc = ""
    Else
        arrKey = GetKey(strKey, 2)
        arrIv = GetKey(strIv, 1)
        
        arrInput = BytePadding(strInput, M_SM4_VERSION)
        ReDim arrOutPut(UBound(arrInput))
        
        Call sm4_crypt_cbc(CM_Encrypt, UBound(arrInput) + 1, arrIv(0), arrKey(0), arrInput(0), arrOutPut(0))
        Sm4EncryptCbc = "ZLSV" & M_SM4_VERSION & ":" & ByteToHexString(arrOutPut)
    End If
End Function

Public Function HexStringToByte(ByVal strInput As String, Optional ByVal lngRetBytLen As Long) As Byte()
    Dim arrReturn() As Byte
    Dim i           As Long
    Dim lngLen      As Long
    
    lngLen = Len(strInput)
    If lngRetBytLen <> 0 Then
        lngLen = lngLen \ 2
        If lngLen > lngRetBytLen Then
            lngLen = lngRetBytLen
        End If
        ReDim arrReturn(lngRetBytLen - 1)
    Else
        lngLen = lngLen \ 2
        ReDim arrReturn(lngLen - 1)
    End If
    
    For i = 0 To lngLen - 1
        arrReturn(i) = Val("&H" & Mid(strInput, 2 * i + 1, 2))
    Next
    
    HexStringToByte = arrReturn()
End Function

Public Function ByteToHexString(bytInpu() As Byte) As String
    Dim i           As Long
    Dim strReturn   As String
    
    For i = LBound(bytInpu) To UBound(bytInpu)
        If Len("" & Hex(bytInpu(i))) = 1 Then
            strReturn = strReturn & "0" & Hex(bytInpu(i))
        Else
            strReturn = strReturn & Hex(bytInpu(i))
        End If
    Next
    
    ByteToHexString = strReturn
End Function

Public Function sm_version() As Long
    Dim lngVersion As Long
    On Error Resume Next
    lngVersion = get_sm_version
    If Err.Number <> 0 Then
        Err.Clear
        sm_version = 1
    Else
        sm_version = lngVersion
    End If
End Function

Private Function GetKey(ByVal strKey As String, ByVal intType As Integer) As Byte()
    Dim arrReturn() As Byte
    Dim i           As Long
    If strKey <> "" Then
        arrReturn = HexStringToByte(strKey, 16)
    Else
        ReDim arrReturn(15)
        If intType = 0 Then
            For i = 0 To 15
                arrReturn(i) = i * 15
            Next
        ElseIf intType = 1 Then
            Rnd (-1)
            Randomize (SM4_CRYPT_RANDOMIZE_IV)
            For i = 0 To 15
                arrReturn(i) = Int(Rnd() * 256)
            Next
        ElseIf intType = 2 Then
            Rnd (-1)
            Randomize (SM4_CRYPT_RANDOMIZE_KEY)
            For i = 0 To 15
                arrReturn(i) = Int(Rnd() * 256)
            Next
        End If
    End If
    GetKey = arrReturn
End Function

Private Function TruncZeroInside(ByVal strInput As String) As String
'功能：去掉字符串中\0以后的字符,仅用作该工程,可以单独是用clsstring
    Dim lngPos As Long
    
    lngPos = InStr(strInput, Chr(0))
    If lngPos > 0 Then
        TruncZeroInside = Mid(strInput, 1, lngPos - 1)
    Else
        TruncZeroInside = strInput
    End If
End Function

Public Function BytePadding(ByVal strInput As String, Optional ByVal lngVersion As Long = 2, Optional ByVal lngPaddingNum As Long = 16) As Byte()
    Dim arrReturn()     As Byte
    Dim lngLenBef       As Long
    Dim i               As Long
    Dim lngLenAft       As Long
    
    '先将字符串由 Unicode 转成系统的缺省码页
    arrReturn = StrConv(strInput, vbFromUnicode)
    lngLenBef = UBound(arrReturn) + 1
    '判断得到的数组的长度，若不是16的整数倍，则补空格或:Chr(0)
    lngLenAft = -Int(-lngLenBef / lngPaddingNum) * lngPaddingNum
    If lngLenBef <> lngLenAft Then
        ReDim Preserve arrReturn(lngLenAft - 1)
        For i = lngLenBef To lngLenAft - 1
            If lngVersion = 1 Then
                arrReturn(i) = 32
            Else
                arrReturn(i) = 0
            End If
        Next
    End If
    BytePadding = arrReturn()
End Function

Public Function ComputerName() As String
'******************************************************************************************************************
'功能：获取电脑名称
'参数：
'说明：
'******************************************************************************************************************
    Dim strComputer As String * 256
    Call GetComputerName(strComputer, 255)
    ComputerName = strComputer
    ComputerName = Trim(Replace(ComputerName, Chr(0), ""))
End Function

Public Function DisPlayOneValue(valValue As Variant, Optional ByVal blnSerializeObject As Boolean = True) As String
    Dim strTmp As String
    Dim i As Long
    
    If IsArray(valValue) Then
        strTmp = ""
        For i = mdlPublic.LboundEx(valValue) To mdlPublic.UboundEx(valValue)
            strTmp = strTmp & ", " & DisPlayOneValue(valValue(i), blnSerializeObject)
        Next
        If strTmp <> "" Then
            strTmp = "[" & Mid(strTmp, 3) & "]"
        End If
    ElseIf IsNull(valValue) Then
        strTmp = "{NULL}"
    ElseIf IsEmpty(valValue) Then
        strTmp = "{EMPTY}"
    ElseIf IsObject(valValue) Then
        If valValue Is Nothing Then
            strTmp = "{NOTHING}"
        Else
            If blnSerializeObject Then
                strTmp = "{OBJECT(" + TypeName(valValue) + ")=" & Serialize(valValue) & "}"
            Else
                strTmp = "{OBJECT(" + TypeName(valValue) + ")}"
            End If
        End If
    Else
        If VarType(valValue) = vbString Then
            strTmp = """" & valValue & """"
        Else
            strTmp = CStr(valValue)
        End If
    End If
    DisPlayOneValue = strTmp
End Function

Public Function LboundEx(varArray As Variant) As Long
    On Error GoTo ErrH
    LboundEx = LBound(varArray)
    Exit Function
ErrH:
    LboundEx = 0
End Function

Public Function UboundEx(varArray As Variant) As Long
    On Error GoTo ErrH
    UboundEx = UBound(varArray)
    Exit Function
ErrH:
    UboundEx = -1
End Function

Public Function Serialize(ByVal objinfo As Variant) As String
    Const KEY_DEFAULT_NAME = "K0"
    Dim objBag      As New PropertyBag
    Dim bytData()   As Byte
    Dim i           As Long
    On Error Resume Next
    
    If IsArray(objinfo) Then
        objBag.WriteProperty "KL", UBound(objinfo)
        For i = LBound(objinfo) To UBound(objinfo)
            If IsArray(objinfo(i)) Then
                objBag.WriteProperty "A" & i, 1
                objBag.WriteProperty "K" & i, Serialize(objinfo(i))
            Else
                objBag.WriteProperty "K" & i, objinfo(i)
                If Err.Number = 330 Then
                    '非法参数。  因为不支持持久性不能写对象。
                    Err.Clear
                    objBag.WriteProperty "K" & i, Nothing
                End If
            End If
        Next
        bytData = objBag.Contents
        Serialize = EncodeBase64(bytData())
    Else
        objBag.WriteProperty KEY_DEFAULT_NAME, objinfo
        If Err.Number = 330 Then
            '非法参数。  因为不支持持久性不能写对象。
            Serialize = "{NotPersistable}"
            Err.Clear
        Else
            bytData = objBag.Contents
            Serialize = EncodeBase64(bytData())
        End If
    End If
End Function

'======================================================================================================================
'方法           EncodeBase64            进行Base64编码，返回Base64的字符串
'返回值         String                  Base64编码结果
'入参列表:
'参数名         类型                    说明
'varInput       Variant                 需要进行Base64编码的字符串或者字节数组，字符串采取UTF-8编码。Byte()类型前面的数组，元素个数传3的倍数，最后一次传递所有剩下的即可。
'方法说明，Base64是将三个字节，每6位分割为四个字节处理的
'======================================================================================================================
Public Function EncodeBase64(varInput As Variant) As String
    Dim bytInput()  As Byte, lngInputLen    As Long
    Dim bytOut()    As Byte, lngOutLen      As Long
    Dim i           As Long, j              As Long, lngBit     As Long
    
    On Error GoTo ErrH
    
    If VarType(varInput) = vbString Then
        If Len(varInput) = 0 Then Exit Function
        '原始内容,先将原文以UTF-8的方式编码
        bytInput = StringToUTF8Bytes(CStr(varInput))
    ElseIf VarType(varInput) = vbArray + vbByte Then
        If UBound(varInput) < 0 Then Exit Function
        bytInput = varInput
    Else
        Exit Function
    End If
    lngInputLen = UBound(bytInput) + 1
 
    lngOutLen = lngInputLen + (lngInputLen - 1) \ 3 + 1
    ReDim bytOut(lngOutLen - 1)
    '将8-bit字节数组转换为6-bit字节数组
    For i = 0 To lngInputLen - 1
        If lngBit = 0 Then 'bytOut(J)未被写入
            bytOut(j) = (bytInput(i) And &HFC) \ &H4
            j = j + 1
            bytOut(j) = (bytInput(i) And &H3) * &H10
            lngBit = 2 '234567 'NNNN01 'N:Next byte
        ElseIf lngBit = 2 Then 'bytOut(J)已被写入两位
            bytOut(j) = bytOut(j) Or ((bytInput(i) And &HF0) \ &H10)
            j = j + 1
            bytOut(j) = (bytInput(i) And &HF) * &H4
            lngBit = 4 '4567PP 'P:Prev byte 'NN0123 'N:Next byte
        ElseIf lngBit = 4 Then 'bytOut(J)已被写入四位
            bytOut(j) = bytOut(j) Or ((bytInput(i) And &HC0) / &H40)
            j = j + 1
            bytOut(j) = bytInput(i) And &H3F
            j = j + 1
            lngBit = 0 '67PPPP 'P:Prev byte '012345
        End If
    Next

    For i = 0 To lngOutLen - 1
        bytOut(i) = EncBase64Char(bytOut(i)) '转换为Base64字符
    Next
    EncodeBase64 = StrConv(bytOut, vbUnicode) & String(2 - (lngInputLen - 1) Mod 3, "=") '原文剩余内容不足3个字节需要补齐
    Exit Function
ErrH:
    Err.Clear
    If 0 = 1 Then
        Resume
    End If
End Function

'======================================================================================================================
'方法           StringToUTF8Bytes       将字符串转换为UTF-8编码的字节数组
'返回值         Byte()                  16进制字符串转换的字节组
'入参列表:
'参数名         类型                    说明
'strInput      String                  16进制字符串
'======================================================================================================================
Public Function StringToUTF8Bytes(strInput As String) As Byte()
    Dim bytUTF8Bytes() As Byte
    Dim lngBytesRequired As Long
    
    '先计算需求字节数
    lngBytesRequired = WideCharToMultiByte(CP_UTF8, 0, ByVal StrPtr(strInput), Len(strInput), ByVal 0, 0, ByVal 0, ByVal 0)
     
    '然后转换
    ReDim bytUTF8Bytes(lngBytesRequired - 1)
    WideCharToMultiByte CP_UTF8, 0, ByVal StrPtr(strInput), Len(strInput), bytUTF8Bytes(0), lngBytesRequired, ByVal 0, ByVal 0
    
    StringToUTF8Bytes = bytUTF8Bytes
End Function

'======================================================================================================================
'方法           EncBase64Char           将6-bit字节转换为Base64字符
'返回值         Byte                    字符数值
'入参列表:
'参数名         类型                    说明
'方法说明，Base64是将三个字节，每6位分割为四个字节处理的
'======================================================================================================================
Private Function EncBase64Char(ByVal bytValue As Byte) As Byte
    If bytValue < 26 Then '26个大写英文字母
        EncBase64Char = bytValue + &H41
    ElseIf bytValue < 52 Then '26个小写英文字母
        EncBase64Char = bytValue + &H61 - 26
    ElseIf bytValue < 62 Then '10个数字
        EncBase64Char = bytValue + &H30 - 52
    ElseIf bytValue = 62 Then
        EncBase64Char = &H2B '+
    Else
        EncBase64Char = &H2F '/
    End If
End Function


