Attribute VB_Name = "mdlPEDefine"
Option Explicit

' ----------------------------------------------------
'   MS-DOS 2.0 兼容EXE 文件头       |               |映像头基地址
'-----------------------------------|               |
'       未使用                      |               |
'-----------------------------------|               |
'   OEM 标识                        |               |
'   OEM 信息                        |               |
'   PE 文件头偏移                   |               |MS-DOS 2.0 节（仅用于MS-DOS 兼容）
'-----------------------------------|               |
'MS-DOS 2.0 占位程序 和 重定位表    |               |
'-----------------------------------|               |
'   未使用                          |               |
'---------------------------------------------------
'PE 文件头（按8 字节边界对齐）      |
'-----------------------------------
'   节头                            |
'-----------------------------------
'   映像页:                         |
'   导入信息                        |
'   导出信息                        |
'   基址重定位信息                  |
'   资源信息                        |
'-----------------------------------

'    +-------------------+
'    | DOS-stub          |    --DOS-头
'    +-------------------+
'    | file-header       |    --文件头
'    +-------------------+
'    | optional header   |    --可选头
'    |- - - - - - - - - -|
'    |                   |
'    | data directories  |    --数据目录
'    |                   |
'    +-------------------+
'    |                   |
'    | section headers   |     --节头
'    |                   |
'    +-------------------+
'    |                   |
'    | section 1         |     --节1
'    |                   |
'    +-------------------+
'    |                   |
'    | section 2         |     --节2
'    |                   |
'    +-------------------+
'    |                   |
'    | ...               |
'    |                   |
'    +-------------------+
'    |                   |
'    | section n         |     --节n
'    |                   |
'    +-------------------+
'节名       内容                                                        特征
'.bss       未初始化的数据（自由格式）                                  IMAGE_SCN_CNT_UNINITIALIZED_DATA |IMAGE_SCN_MEM_READ |IMAGE_SCN_MEM_WRITE
'.cormeta   CLR 元数据，它表明目标文件中包含托管代码                    IMAGE_SCN_LNK_INFO
'.data      已初始化的数据（自由格式）                                  IMAGE_SCN_CNT_INITIALIZED_DATA |IMAGE_SCN_MEM_READ |IMAGE_SCN_MEM_WRITE
'.debug$F   生成的FPO 调试信息（仅适用于目标文件，仅用于x86 平台，现已被舍弃）  IMAGE_SCN_CNT_INITIALIZED_DATA |IMAGE_SCN_MEM_READ |IMAGE_SCN_MEM_DISCARDABLE
'.debug$P   预编译的调试类型信息（仅适用于目标文件）                    IMAGE_SCN_CNT_INITIALIZED_DATA |IMAGE_SCN_MEM_READ |IMAGE_SCN_MEM_DISCARDABLE
'.debug$S   调试符号信息（仅适用于目标文件）                            IMAGE_SCN_CNT_INITIALIZED_DATA |IMAGE_SCN_MEM_READ |IMAGE_SCN_MEM_DISCARDABLE
'.debug$T   调试类型信息（仅适用于目标文件）                            IMAGE_SCN_CNT_INITIALIZED_DATA |IMAGE_SCN_MEM_READ |IMAGE_SCN_MEM_DISCARDABLE
'.drectve   链接器选项                                                  IMAGE_SCN_LNK_INFO
'.edata     导出表                                                      IMAGE_SCN_CNT_INITIALIZED_DATA |IMAGE_SCN_MEM_READ
'.idata     导入表                                                      IMAGE_SCN_CNT_INITIALIZED_DATA |IMAGE_SCN_MEM_READ |IMAGE_SCN_MEM_WRITE
'.idlsym    包含已注册的SEH（仅适用于映像文件），它们用以支持IDL 属性   IMAGE_SCN_LNK_INFO
'.pdata     异常信息                                                    IMAGE_SCN_CNT_INITIALIZED_DATA |IMAGE_SCN_MEM_READ
'.rdata     只读的已初始化数据                                          IMAGE_SCN_CNT_INITIALIZED_DATA |IMAGE_SCN_MEM_READ
'.reloc     映像文件的重定位信息                                        IMAGE_SCN_CNT_INITIALIZED_DATA |IMAGE_SCN_MEM_READ |IMAGE_SCN_MEM_DISCARDABLE
'.rsrc      资源目录                                                    IMAGE_SCN_CNT_INITIALIZED_DATA |IMAGE_SCN_MEM_READ
'.sbss      与GP 相关的未初始化数据（自由格式）                         IMAGE_SCN_CNT_UNINITIALIZED_DATA |IMAGE_SCN_MEM_READ |IMAGE_SCN_MEM_WRITE |IMAGE _SCN_GPREL 其中IMAGE_SCN_GPREL 标志仅用于IA64 平台，不能用于其它平台。此标志只能用于目标文件。当映像文件中出现这种类型的节时，一定不能设置这个标志
'.sdata     与GP 相关的已初始化数据（自由格式）                         IMAGE_SCN_CNT_INITIALIZED_DATA |IMAGE_SCN_MEM_READ |IMAGE_SCN_MEM_WRITE |IMAGE _SCN_GPREL其中IMAGE_SCN_GPREL 标志仅用于IA64 平台，不能用于其它平台。此标志只能用于目标文件。当映像文件中出现这种类型的节时，一定不能设置这个标志
'.srdata    与GP 相关的只读数据（自由格式）                             IMAGE_SCN_CNT_INITIALIZED_DATA |IMAGE_SCN_MEM_READ |IMAGE _SCN_GPREL其中IMAGE_SCN_GPREL 标志仅用于IA64 平台，不能用于其它平台。此标志只能用于目标文件。当映像文件中出现这种类型的节时，一定不能设置这个标志
'.sxdata    已注册的异常处理程序数据（自由格式，仅适用于目标文件，仅用于x86 平台）  MAGE_SCN_LNK_INFO这个节中包含目标文件中的代码所涉及到的所有异常处理程序在符号表中的索引.这些符号可以是IMAGE_SYM_UNDEFINED 类型的符号，也可以是定义在那个模块中的符号
'.text      可执行代码（自由格式）                                      IMAGE_SCN_CNT_CODE |IMAGE_SCN_MEM_EXECUTE |IIMAGE_SCN_MEM_READ
'.tls       线程局部存储（仅适用于目标文件）                            IMAGE_SCN_CNT_INITIALIZED_DATA |IMAGE_SCN_MEM_READ |IMAGE_SCN_MEM_WRITE
'.tls$      线程局部存储（仅适用于目标文件）                            IMAGE_SCN_CNT_INITIALIZED_DATA |IMAGE_SCN_MEM_READ |IMAGE_SCN_MEM_WRITE
'.vsdata    与GP 相关的已初始化数据（自由格式，仅适用于ARM、SH4 和Thumb 平台）  IMAGE_SCN_CNT_INITIALIZED_DATA |IMAGE_SCN_MEM_READ |IMAGE_SCN_MEM_WRITE
'.xdata     异常信息（自由格式）                                        IMAGE_SCN_CNT_INITIALIZED_DATA |IMAGE_SCN_MEM_READ
'*******************************************************************************************
'   MS-DOS 2.0 兼容EXE 文件头
'*******************************************************************************************
'PE_SIGNATURE(e_magic,ne_magic,e32_magic)
Public Const IMAGE_DOS_SIGNATURE                As Integer = &H5A4D                 'MZ
Public Const IMAGE_OS2_SIGNATURE                As Integer = &H454E                 'NE
Public Const IMAGE_OS2_SIGNATURE_LE             As Integer = &H454C                 'LE
Public Const IMAGE_NT_SIGNATURE                 As Long = &H4550                    'PE00

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

Public Type IMAGE_OS2_HEADER      'OS/2 .EXE header 64B
    ne_magic            As Integer              'Magic number
    ne_ver              As Byte                 'Version number
    ne_rev              As Byte                 'Revision number
    ne_enttab           As Integer              'Offset of Entry Table
    ne_cbenttab         As Integer              'Number of bytes in Entry Table
    ne_crc              As Long                 'Checksum of whole file
    ne_flags            As Integer              'Flag word
    ne_autodata         As Integer              'Automatic data segment number
    ne_heap             As Integer              'Initial heap allocation
    ne_stack            As Integer              'Initial stack allocation
    ne_csip             As Long                 'Initial CS:IP setting
    ne_sssp             As Long                 'Initial SS:SP setting
    ne_cseg             As Integer              'Count of file segments
    ne_cmod             As Integer              'Entries in Module Reference Table
    ne_cbnrestab        As Integer              'Size of non-resident name table
    ne_segtab           As Integer              'Offset of Segment Table
    ne_rsrctab          As Integer              'Offset of Resource Table
    ne_restab           As Integer              'Offset of resident name table
    ne_modtab           As Integer              'Offset of Module Reference Table
    ne_imptab           As Integer              'Offset of Imported Names Table
    ne_nrestab          As Long                 'Offset of Non-resident Names Table
    ne_cmovent          As Integer              'Count of movable entries
    ne_align            As Integer              'Segment alignment shift count
    ne_cres             As Integer              'Count of resource segments
    ne_exetyp           As Byte                 'Target Operating system
    ne_flagsothers      As Byte                 'Other .EXE flags
    ne_pretthunks       As Integer              'offset to return thunks
    ne_psegrefbytes     As Integer              'offset to segment ref. bytes
    ne_swaparea         As Integer              'Minimum code swap area size
    ne_expver           As Integer              'Expected Windows version number
End Type

Public Type IMAGE_VXD_HEADER    'Windows VXD header 196B
    e32_magic           As Integer              'Magic number
    e32_border          As Byte                 'The byte ordering for the VXD
    e32_worder          As Byte                 'The word ordering for the VXD
    e32_level           As Long                 'The EXE format level for now = 0
    e32_cpu             As Integer              'The CPU type
    e32_os              As Integer              'The OS type
    e32_ver             As Long                 'Module version
    e32_mflags          As Long                 'Module flags
    e32_mpages          As Long                 'Module # pages
    e32_startobj        As Long                 'Object # for instruction pointer
    e32_eip             As Long                 'Extended instruction pointer
    e32_stackobj        As Long                 'Object # for stack pointer
    e32_esp             As Long                 'Extended stack pointer
    e32_pagesize        As Long                 'VXD page size
    e32_lastpagesize    As Long                 'Last page size in VXD
    e32_fixupsize       As Long                 'Fixup section size
    e32_fixupsum        As Long                 'Fixup section checksum
    e32_ldrsize         As Long                 'Loader section size
    e32_ldrsum          As Long                 'Loader section checksum
    e32_objtab          As Long                 'Object table offset
    e32_objcnt          As Long                 'Number of objects in module
    e32_objmap          As Long                 'Object page map offset
    e32_itermap         As Long                 'Object iterated data map offset
    e32_rsrctab         As Long                 'Offset of Resource Table
    e32_rsrccnt         As Long                 'Number of resource entries
    e32_restab          As Long                 'Offset of resident name table
    e32_enttab          As Long                 'Offset of Entry Table
    e32_dirtab          As Long                 'Offset of Module Directive Table
    e32_dircnt          As Long                 'Number of module directives
    e32_fpagetab        As Long                 'Offset of Fixup Page Table
    e32_frectab         As Long                 'Offset of Fixup Record Table
    e32_impmod          As Long                 'Offset of Import Module Name Table
    e32_impmodcnt       As Long                 'Number of entries in Import Module Name Table
    e32_impproc         As Long                 'Offset of Import Procedure Name Table
    e32_pagesum         As Long                 'Offset of Per-Page Checksum Table
    e32_datapage        As Long                 'Offset of Enumerated Data Pages
    e32_preload         As Long                 'Number of preload pages
    e32_nrestab         As Long                 'Offset of Non-resident Names Table
    e32_cbnrestab       As Long                 'Size of Non-resident Name Table
    e32_nressum         As Long                 'Non-resident Name Table Checksum
    e32_autodata        As Long                 'Object # for automatic data object
    e32_debuginfo       As Long                 'Offset of the debugging information
    e32_debuglen        As Long                 'The length of the debugging info. in bytes
    e32_instpreload     As Long                 'Number of instance pages in preload section of VXD file
    e32_instdemand      As Long                 'Number of instance pages in demand load section of VXD file
    e32_heapsize        As Long                 'Size of heap - for 16-bit apps
    e32_res3(11)        As Byte                 'Reserved words
    e32_winresoff       As Long
    e32_winreslen       As Long
    e32_devid           As Integer               'Device ID for VxD
    e32_ddkver          As Integer               'DDK version for VxD
End Type
'*******************************************************************************************
'   文件头（3部分）
'*******************************************************************************************
'-------NT头-------------
Public Const IMAGE_SIZEOF_FILE_HEADER       As Integer = 20
Public Type IMAGE_FILE_HEADER                   '20B
    Machine             As Integer              '标识目标机器类型的数字
    NumberOfSections    As Integer              '节的数目。它给出了节表的大小，而节表紧跟着文件头
    TimeDateStamp       As Long                 '从UTC 时间1970 年1 月1 日00:00 起的总秒数（一个C 运行时time_t 类型的值）的低32 位，它指出文件何时被创建
    PointerToSymbolTable    As Long             'COFF 符号表的文件偏移。如果不存在COFF 符号表，此值为0。对于映像文件来说，此值应该为0，因为已经不赞成使用COFF 调试信息了
    NumberOfSymbols     As Long                 '符号表中的元素数目。由于字符串表紧跟符号表，所以可以利用这个值来定位字符串表?对于映像文件来说，此值应该为0，因为已经不赞成使用COFF调试信息了
    SizeOfOptionalHeader    As Integer          '可选文件头的大小。可执行文件需要可选文件头而目标文件并不需要。对于目标文件来说，此值应该为0
    Characteristics     As Integer              '指示文件属性的标志。
End Type

'IMAGE_FILE_MACHINE
Public Const IMAGE_FILE_MACHINE_UNKNOWN         As Integer = &H0                    '适用于任何类型处理器
Public Const IMAGE_FILE_MACHINE_I386            As Integer = &H14C                  'Intel 386.Intel 386 或后继处理器及其兼容处理器
Public Const IMAGE_FILE_MACHINE_R3000           As Integer = &H162                  'MIPS little-endian,  big-endian
Public Const IMAGE_FILE_MACHINE_R4000           As Integer = &H166                  'MIPS little-endian     MIPS 小尾处理器
Public Const IMAGE_FILE_MACHINE_R10000          As Integer = &H168                  'MIPS little-endian
Public Const IMAGE_FILE_MACHINE_WCEMIPSV2       As Integer = &H169                  'MIPS little-endian WCE v2  MIPS 小尾WCE v2 处理器
Public Const IMAGE_FILE_MACHINE_ALPHA           As Integer = &H184                  'Alpha_AXP
Public Const IMAGE_FILE_MACHINE_SH3             As Integer = &H1A2                  'SH3 little-endian  Hitachi SH3 处理器
Public Const IMAGE_FILE_MACHINE_SH3DSP          As Integer = &H1A3                  'Hitachi SH3 DSP 处理器
Public Const IMAGE_FILE_MACHINE_SH3E            As Integer = &H1A4                  'SH3E little-endian
Public Const IMAGE_FILE_MACHINE_SH4             As Integer = &H1A6                  'SH4 little-endian      Hitachi SH4 处理器
Public Const IMAGE_FILE_MACHINE_SH5             As Integer = &H1A8                  'SH5        Hitachi SH5 处理器
Public Const IMAGE_FILE_MACHINE_ARM             As Integer = &H1C0                  'ARM Little-Endian  ARM 小尾处理器
Public Const IMAGE_FILE_MACHINE_THUMB           As Integer = &H1C2                  'ARM Thumb/Thumb-2 Little-Endian    Thumb 处理器
Public Const IMAGE_FILE_MACHINE_ARMNT           As Integer = &H1C4                  'ARM Thumb-2 Little-Endian
Public Const IMAGE_FILE_MACHINE_AM33            As Integer = &H1D3                  'Matsushita AM33 处理器
Public Const IMAGE_FILE_MACHINE_POWERPC         As Integer = &H1F0                  'IBM PowerPC Little-Endian  PowerPC 小尾处理器
Public Const IMAGE_FILE_MACHINE_POWERPCFP       As Integer = &H1F1                  '带符点运算支持的PowerPC 处理器
Public Const IMAGE_FILE_MACHINE_IA64            As Integer = &H200                  'Intel 64 Intel Itanium 处理器家族
Public Const IMAGE_FILE_MACHINE_MIPS16          As Integer = &H266                  'MIPS   MIPS16 处理器
Public Const IMAGE_FILE_MACHINE_ALPHA64         As Integer = &H284                  'ALPHA64
Public Const IMAGE_FILE_MACHINE_MIPSFPU         As Integer = &H366                  'MIPS   带FPU 的MIPS 处理器
Public Const IMAGE_FILE_MACHINE_MIPSFPU16       As Integer = &H466                  'MIPS   带FPU 的MIPS16 处理器
Public Const IMAGE_FILE_MACHINE_AXP64           As Integer = IMAGE_FILE_MACHINE_ALPHA64
Public Const IMAGE_FILE_MACHINE_TRICORE         As Integer = &H520                  'Infineon
Public Const IMAGE_FILE_MACHINE_CEF             As Integer = &HCEF
Public Const IMAGE_FILE_MACHINE_EBC             As Integer = &HEBC                  'EFI Byte Code  EFI 字节码处理器
Public Const IMAGE_FILE_MACHINE_AMD64           As Integer = &H8664                 'AMD64 (K8) x64 处理器
Public Const IMAGE_FILE_MACHINE_M32R            As Integer = &H9041                 'M32R little-endian  Mitsubishi M32R 小尾处理器
Public Const IMAGE_FILE_MACHINE_CEE             As Integer = &HC0EE

'IMAGE_FILE_MACHINE  Characteristics
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

'-------可选头-------------
'目录格式
Public Type IMAGE_DATA_DIRECTORY
    VirtualAddress      As Long                 '数据块的RVA
    Size                As Long                 '数据块大小
End Type
'目录数目
Public Const IMAGE_NUMBEROF_DIRECTORY_ENTRIES   As Integer = 16
'偏移（PE32/PE32+）  大小（PE32/PE32+）   文件头部分        描述
'0                  28/24                 标准域            这些域被所有COFF 实现所定义，其中包括UNIX
'28/24              68/88                 Windows 特定域    支持Windows 特性（例如子系统）的附加域
'96/112             Variable              数据目录          映像文件中的特殊表（例如导入表和导出表）的地址/大小组合，它们供操作系统使用.
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

Public Type IMAGE_ROM_OPTIONAL_HEADER
    Magic               As Integer              '这个无符号整数指出了映像文件的状态。最常用的数字是0x10B，它表明这是一个正常的可执行文件。0x107 表明这是一个ROM 映像，0x20B 表明这是一个PE32 + 可执行文件?
    MajorLinkerVersion  As Byte                 '链接器的主版本号
    MinorLinkerVersion  As Byte                 '链接器的次版本号
    SizeOfCode          As Long                 '代码节（.text）的大小。如果有多个代码节的话，它是所有代码节的和。
    SizeOfInitializedData   As Long             '已初始化数据节的大小。如果有多个这样的数据节的话，它是所有这些数据节的和。
    SizeOfUninitializedData As Long             '未初始化数据节（.bss）的大小。如果有多个.bss 节的话，它是所有这些节的和。
    AddressOfEntryPoint As Long                 '当可执行文件被加载进内存时其入口点相对于映像基址的偏移地址.对于一般程序映像来说，它就是启动地址。对于设备驱动程序来说，它是初始化函数的地址。入口点对于DLL来说是可选的。如果不存在入口点的话，这个域必须为0.
    BaseOfCode          As Long                 '当映像被加载进内存时代码节的开头相对于映像基址的偏移地址。
    BaseOfData          As Long                 '当映像被加载进内存时数据节的开头相对于映像基址的偏移地址。PE32独有
    BaseOfBss           As Long
    GprMask             As Long
    CprMask(3)          As Long
    GpValue             As Long
End Type
'64位整形定义
Private Type Big_Iint
    Low                 As Long
    High                As Long
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

'IMAGE_OPTIONAL_HEADER Magic
Public Const IMAGE_NT_OPTIONAL_HDR32_MAGIC      As Integer = &H10B                  '这是一个32位镜像文件
Public Const IMAGE_NT_OPTIONAL_HDR64_MAGIC      As Integer = &H20B                  '这是一个PE32+可执行文件
Public Const IMAGE_ROM_OPTIONAL_HDR_MAGIC       As Integer = &H107                  '这是一个ROM镜像

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

Public Type IMAGE_ROM_HEADERS
    FileHeader          As IMAGE_FILE_HEADER
    OptionalHeader      As IMAGE_ROM_OPTIONAL_HEADER
End Type

'IMAGE_OPTIONAL_HEADER  Subsystem Values
Public Const IMAGE_SUBSYSTEM_UNKNOWN            As Integer = 0                      ' Unknown subsystem. 未知子系统
Public Const IMAGE_SUBSYSTEM_NATIVE             As Integer = 1                      ' Image doesn't require a subsystem.设备驱动程序和Native Windows 进程
Public Const IMAGE_SUBSYSTEM_WINDOWS_GUI        As Integer = 2                      ' Image runs in the Windows GUI subsystem.Windows 图形用户界面（GUI）子系统
Public Const IMAGE_SUBSYSTEM_WINDOWS_CUI        As Integer = 3                      ' Image runs in the Windows character subsystem.Windows 字符模式（CUI）子系统
Public Const IMAGE_SUBSYSTEM_OS2_CUI            As Integer = 5                      ' image runs in the OS/2 character subsystem.
Public Const IMAGE_SUBSYSTEM_POSIX_CUI          As Integer = 7                      ' image runs in the Posix character subsystem.Posix 字符模式子系统
Public Const IMAGE_SUBSYSTEM_NATIVE_WINDOWS     As Integer = 8                      ' image is a native Win9x driver.
Public Const IMAGE_SUBSYSTEM_WINDOWS_CE_GUI     As Integer = 9                      ' Image runs in the Windows CE subsystem. Windows CE
Public Const IMAGE_SUBSYSTEM_EFI_APPLICATION    As Integer = 10                     ' 可扩展固件接口（EFI）应用程序
Public Const IMAGE_SUBSYSTEM_EFI_BOOT_SERVICE_DRIVER    As Integer = 11             ' 带引导服务的EFI 驱动程序
Public Const IMAGE_SUBSYSTEM_EFI_RUNTIME_DRIVER As Integer = 12                     ' 带运行时服务的EFI 驱动程序
Public Const IMAGE_SUBSYSTEM_EFI_ROM            As Integer = 13                     ' EFI ROM 映像
Public Const IMAGE_SUBSYSTEM_XBOX               As Integer = 14                     ' XBOX
Public Const IMAGE_SUBSYSTEM_WINDOWS_BOOT_APPLICATION   As Integer = 16
'IMAGE_OPTIONAL_HEADER DllCharacteristics Entries

'IMAGE_OPTIONAL_HEADER DllCharacteristics Entries
'Public Const IMAGE_LIBRARY_PROCESS_INIT            0x0001                           ' Reserved.保留，必须为0。
'Public Const IMAGE_LIBRARY_PROCESS_TERM            0x0002                           ' Reserved.保留，必须为0。
'Public Const IMAGE_LIBRARY_THREAD_INIT             0x0004                           ' Reserved.保留，必须为0。
'Public Const IMAGE_LIBRARY_THREAD_TERM             0x0008                           ' Reserved.保留，必须为0。
Public Const IMAGE_DLLCHARACTERISTICS_HIGH_ENTROPY_VA   As Integer = &H20           ' Image can handle a high entropy 64-bit virtual address space.
Public Const IMAGE_DLLCHARACTERISTICS_DYNAMIC_BASE  As Integer = &H40               ' DLL can move.DLL 可以在加载时被重定位。
Public Const IMAGE_DLLCHARACTERISTICS_FORCE_INTEGRITY   As Integer = &H80           ' Code Integrity Image  强制进行代码完整性校验。
Public Const IMAGE_DLLCHARACTERISTICS_NX_COMPAT As Integer = &H100                  ' Image is NX compatible    映像兼容于NX。
Public Const IMAGE_DLLCHARACTERISTICS_NO_ISOLATION As Integer = &H200               ' Image understands isolation and doesn't want it   可以隔离，但并不隔离此映像。
Public Const IMAGE_DLLCHARACTERISTICS_NO_SEH    As Integer = &H400                  ' Image does not use SEH.  No SE handler may reside in this image   不使用结构化异常（SE）处理。在此映像中不能调用SE 处理程序
Public Const IMAGE_DLLCHARACTERISTICS_NO_BIND   As Integer = &H800                  ' Do not bind this image.   不绑定映像。
Public Const IMAGE_DLLCHARACTERISTICS_APPCONTAINER As Integer = &H1000              ' Image should execute in an AppContainer
Public Const IMAGE_DLLCHARACTERISTICS_WDM_DRIVER   As Integer = &H2000              ' Driver uses WDM model WDM 驱动程序。
'                                                As Integer=&H4000                   ' Reserved.    保留，必须为0。
Public Const IMAGE_DLLCHARACTERISTICS_TERMINAL_SERVER_AWARE     As Integer = &H8000 '可以用于终端服务器。

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

Public Type GUID
    Data1               As Long
    Data2               As Integer
    Data3               As Integer
    Data4(7)            As Byte
End Type

Public Type CLSID
    Value               As GUID
End Type

'Non-COFF Object file header
Public Type ANON_OBJECT_HEADER
    Sig1                As Integer              ' Must be IMAGE_FILE_MACHINE_UNKNOWN
    Sig2                As Integer              ' Must be 0xffff
    Version             As Integer              ' >= 1 (implies the CLSID field is present)
    Machine             As Integer
    TimeDateStamp       As Long
    ClassID             As CLSID                ' Used to invoke CoCreateInstance
    SizeOfData          As Long                 ' Size of data that follows the header
End Type

Public Type ANON_OBJECT_HEADER_V2
    Sig1                As Integer              ' Must be IMAGE_FILE_MACHINE_UNKNOWN
    Sig2                As Integer              ' Must be 0xffff
    Version             As Integer              ' >= 2 (implies the Flags field is present - otherwise V1)
    Machine             As Integer
    TimeDateStamp       As Long
    ClassID             As CLSID                ' Used to invoke CoCreateInstance
    SizeOfData          As Long                 ' Size of data that follows the header
    Flags               As Long                 ' 0x1 -> contains metadata
    MetaDataSize        As Long                 ' Size of CLR metadata
    MetaDataOffset      As Long                 ' Offset of CLR metadata
End Type

Public Type ANON_OBJECT_HEADER_BIGOBJ
    'same as ANON_OBJECT_HEADER_V2
    Sig1                As Integer              ' Must be IMAGE_FILE_MACHINE_UNKNOWN
    Sig2                As Integer              ' Must be 0xffff
    Version             As Integer              ' >= 2 (implies the Flags field is present)
    Machine             As Integer              ' Actual machine - IMAGE_FILE_MACHINE_xxx
    TimeDateStamp       As Long
    ClassID             As CLSID                ' {D1BAA1C7-BAEE-4ba9-AF20-FAF66AA4DCB8}
    SizeOfData          As Long                 ' Size of data that follows the header
    Flags               As Long                 ' 0x1 -> contains metadata
    MetaDataSize        As Long                 ' Size of CLR metadata
    MetaDataOffset      As Long                 ' Offset of CLR metadata
    'bigobj specifics
    NumberOfSections    As Long                 ' extended from WORD
    PointerToSymbolTable    As Long
    NumberOfSymbols     As Long
End Type

'Section header format.
Public Const IMAGE_SIZEOF_SHORT_NAME            As Integer = 8
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

'IMAGE_SECTION_HEADER Section characteristics.
'Public Const IMAGE_SCN_TYPE_REG                 As Long=&H00000000                  ' Reserved.保留供将来使用
'Public Const IMAGE_SCN_TYPE_DSECT               As Long=&H00000001                  ' Reserved.保留供将来使用
'Public Const IMAGE_SCN_TYPE_NOLOAD              As Long=&H00000002                  ' Reserved.保留供将来使用
'Public Const IMAGE_SCN_TYPE_GROUP               As Long=&H00000004                  ' Reserved.保留供将来使用
Public Const IMAGE_SCN_TYPE_NO_PAD              As Long = &H8                       ' Reserved.从这个节结尾到下一个边界之间不能填充。此标志被舍弃，它已经被IMAGE_SCN_ALIGN_1BYTES 标志取代.此标志仅对目标文件合法
'Public Const IMAGE_SCN_TYPE_COPY                As Long=&H00000010                  ' Reserved.保留供将来使用。

Public Const IMAGE_SCN_CNT_CODE                 As Long = &H20                      ' Section contains code.此节包含可执行代码。
Public Const IMAGE_SCN_CNT_INITIALIZED_DATA     As Long = &H40                      ' Section contains initialized data.此节包含已初始化的数据。
Public Const IMAGE_SCN_CNT_UNINITIALIZED_DATA   As Long = &H80                      ' Section contains uninitialized data.此节包含未初始化的数据。

Public Const IMAGE_SCN_LNK_OTHER                As Long = &H100                     ' Reserved.保留供将来使用。
Public Const IMAGE_SCN_LNK_INFO                 As Long = &H200                     ' Section contains comments or some other type of information.此节包含注释或者其它信息..drectve 节具有这种属性?此标志仅对目标文件合法
'Public Const IMAGE_SCN_TYPE_OVER                As Long=&H00000400                  ' Reserved.保留供将来使用。
Public Const IMAGE_SCN_LNK_REMOVE               As Long = &H800                     ' Section contents will not become part of image.此节不会成为最终形成的映像文件的一部分.此标志仅对目标文件合法
Public Const IMAGE_SCN_LNK_COMDAT               As Long = &H1000                    ' Section contents comdat.此节包含COMDAT 数据。
'Public Const                                    As Long=&H00002000                  ' Reserved.
'Public Const IMAGE_SCN_MEM_PROTECTED - Obsolete As Long=&H00004000
Public Const IMAGE_SCN_NO_DEFER_SPEC_EXC        As Long = &H4000                    ' Reset speculative exceptions handling bits in the TLB entries for this section.
Public Const IMAGE_SCN_GPREL                    As Long = &H8000                    ' Section content can be accessed relative to GP 此节包含通过全局指针（GP）来引用的数据
Public Const IMAGE_SCN_MEM_FARDATA              As Long = &H8000
'Public Const IMAGE_SCN_MEM_SYSHEAP  - Obsolete  As Long=&H00010000
Public Const IMAGE_SCN_MEM_PURGEABLE            As Long = &H20000                   '保留供将来使用
Public Const IMAGE_SCN_MEM_16BIT                As Long = &H20000                   '保留供将来使用
Public Const IMAGE_SCN_MEM_LOCKED               As Long = &H40000                   '保留供将来使用
Public Const IMAGE_SCN_MEM_PRELOAD              As Long = &H80000                   '保留供将来使用

Public Const IMAGE_SCN_ALIGN_1BYTES             As Long = &H100000                  '按1 字节边界对齐数据。此标志仅对目标文件合法.
Public Const IMAGE_SCN_ALIGN_2BYTES             As Long = &H200000                  '按2字节边界对齐数据。此标志仅对目标文件合法.
Public Const IMAGE_SCN_ALIGN_4BYTES             As Long = &H300000                  '按4字节边界对齐数据。此标志仅对目标文件合法.
Public Const IMAGE_SCN_ALIGN_8BYTES             As Long = &H400000                  '按8字节边界对齐数据。此标志仅对目标文件合法.
Public Const IMAGE_SCN_ALIGN_16BYTES            As Long = &H500000                  ' Default alignment if no others are specified.按16 字节边界对齐数据。此标志仅对目标文件合法
Public Const IMAGE_SCN_ALIGN_32BYTES            As Long = &H600000                  '按32字节边界对齐数据。此标志仅对目标文件合法.
Public Const IMAGE_SCN_ALIGN_64BYTES            As Long = &H700000                  '按64字节边界对齐数据。此标志仅对目标文件合法.
Public Const IMAGE_SCN_ALIGN_128BYTES           As Long = &H800000                  '按128字节边界对齐数据。此标志仅对目标文件合法.
Public Const IMAGE_SCN_ALIGN_256BYTES           As Long = &H900000                  '按256字节边界对齐数据。此标志仅对目标文件合法.
Public Const IMAGE_SCN_ALIGN_512BYTES           As Long = &HA00000                  '按512字节边界对齐数据。此标志仅对目标文件合法.
Public Const IMAGE_SCN_ALIGN_1024BYTES          As Long = &HB00000                  '按1024字节边界对齐数据。此标志仅对目标文件合法.
Public Const IMAGE_SCN_ALIGN_2048BYTES          As Long = &HC00000                  '按2048字节边界对齐数据。此标志仅对目标文件合法.
Public Const IMAGE_SCN_ALIGN_4096BYTES          As Long = &HD00000                  '按4096字节边界对齐数据。此标志仅对目标文件合法.
Public Const IMAGE_SCN_ALIGN_8192BYTES          As Long = &HE00000                  '按8192字节边界对齐数据。此标志仅对目标文件合法.
' Unused                                         As Long=&H00F00000
Public Const IMAGE_SCN_ALIGN_MASK               As Long = &HF00000                  '

Public Const IMAGE_SCN_LNK_NRELOC_OVFL          As Long = &H1000000                 ' Section contains extended relocations. 此节包含扩展的重定位信息。
'                                       IMAGE_SCN_LNK_NRELOC_OVFL 标志表明节中重定位项的个数超出了节头中为每个节保留的16 位所能表示的范围?如果设置了此标志并且节头中的NumberOfRelocations 域的值是0xffff，那么实际的重定位项个数被保存在第一个重
'                                       定位项的VirtualAddress 域（32 位）中。如果设置了IMAGE_SCN_LNK_NRELOC_OVFL标志但节中的重定位项的个数少于0xffff，则表示出现了错误。
Public Const IMAGE_SCN_MEM_DISCARDABLE          As Long = &H2000000                 ' Section can be discarded.此节可以在需要时被丢弃。
Public Const IMAGE_SCN_MEM_NOT_CACHED           As Long = &H4000000                 ' Section is not cachable.此节不能被缓存。
Public Const IMAGE_SCN_MEM_NOT_PAGED            As Long = &H8000000                 ' Section is not pageable.此节不能被交换到页面文件中。
Public Const IMAGE_SCN_MEM_SHARED               As Long = &H10000000                ' Section is shareable.此节可以在内存中共享。
Public Const IMAGE_SCN_MEM_EXECUTE              As Long = &H20000000                ' Section is executable.此节可以作为代码执行。
Public Const IMAGE_SCN_MEM_READ                 As Long = &H40000000                ' Section is readable.此节可读。
Public Const IMAGE_SCN_MEM_WRITE                As Long = &H80000000                ' Section is writeable.此节可写

'TLS Chaacteristic Flags
Public Const IMAGE_SCN_SCALE_INDEX              As Long = &H1                       'Tls index is scaled

'Symbol format.
Public Type IMAGE_SYMBOL
    ShortName(7)        As Byte                 '符号名称，这是一个由三个成员组成的共用体。如果名称的长度不超过8 个字节，那么它就是一个8 字节长的数组。
    'Union
'    Short               As Long                 'if 0, use LongName
'    Long                As Long                 'offset into string table
    'Union
'    LongName(1)         As Long                 'PBYTE [2]
    Value               As Long                 '与符号相关的值。其意义依赖于SectionNumber 和StorageClass 这两个域.它通常表示可重定位的地址
    SectionNumber       As Integer              '这个带符号整数是节表的索引（从1 开始），用以标识定义此符号的节
    vType               As Integer              '一个表示类型的数字。Microsoft 的工具将它设置为0x20（如果是函数）或者0x0（如果不是函数）
    StorageClass        As Byte                 '这是一个表示存储类别的枚举类型值
    NumberOfAuxSymbols  As Byte                 '跟在本记录后面的辅助符号表项的个数。
End Type
Public Const IMAGE_SIZEOF_SYMBOL                As Integer = 18

Public Type IMAGE_SYMBOL_EX
    ShortName(7)        As Byte                 '符号名称，这是一个由三个成员组成的共用体。如果名称的长度不超过8 个字节，那么它就是一个8 字节长的数组。
    'Union
'    Short               As Long                 'if 0, use LongName
'    Long                As Long                 'offset into string table
    'Union
'    LongName(1)         As Long                 'PBYTE [2]
    Value               As Long                 '与符号相关的值。其意义依赖于SectionNumber 和StorageClass 这两个域.它通常表示可重定位的地址
    SectionNumber       As Long                 '这个带符号整数是节表的索引（从1 开始），用以标识定义此符号的节
    vType               As Integer              '一个表示类型的数字。Microsoft 的工具将它设置为0x20（如果是函数）或者0x0（如果不是函数）Type 域占2 个字节，其中的每一个字节都表示类型信息。低位字节（LSB）表示简单（基本）数据类型，高位字节（MSB）表示复杂类型（如果存在）
                                                'MSB:复杂类型：无、指针、函数、数组,LSB:基本类型：整数、浮点数等。
    StorageClass        As Byte                 '这是一个表示存储类别的枚举类型值
    NumberOfAuxSymbols  As Byte                 '跟在本记录后面的辅助符号表项的个数。
End Type
Public Const IMAGE_SIZEOF_SYMBOL_EX             As Integer = 20
' Section values.
' Symbols have a section number of the section in which they are defined. Otherwise, section numbers have the following meanings:
Public Const IMAGE_SYM_UNDEFINED                As Integer = 0                      ' Symbol is undefined or is common.尚未为此符号记录分配一个节。这个零值表明引用了一个定义在其它地方的外部符号；而非零值则表明是一个普通符号，其大小由Value 域给出。
Public Const IMAGE_SYM_ABSOLUTE                 As Integer = -1                     ' Symbol is an absolute value.此符号是个绝对符号（不可重定位），并且不是地址。
Public Const IMAGE_SYM_DEBUG                    As Integer = -2                     ' Symbol is a special debug item.此符号提供普通类型信息或者调试信息，但它并不对应于某一个节。Microsoft 的工具将.file 记录（存储类别为FILE）设置为这个值。
Public Const IMAGE_SYM_SECTION_MAX              As Integer = &HFEFF                 ' Values 0xFF00-0xFFFF are special
Public Const IMAGE_SYM_SECTION_MAX_EX           As Integer = &HFFFF

' IMAGE_SYMBOL Type (fundamental) values.
Public Const IMAGE_SYM_TYPE_NULL                As Integer = &H0                    ' no type.类型信息不存在，或者是未知的基本类型。Microsoft 的工具使用这个值
Public Const IMAGE_SYM_TYPE_VOID                As Integer = &H1                    '不是合法类型；用于void 指针和函数。
Public Const IMAGE_SYM_TYPE_CHAR                As Integer = &H2                    ' type character.字符（带符号的1 个字节）。
Public Const IMAGE_SYM_TYPE_SHORT               As Integer = &H3                    ' type short integer.长度为2 个字节的带符号整数。
Public Const IMAGE_SYM_TYPE_INT                 As Integer = &H4                    '自然的整数类型（在Windows 中通常为4 个字节）。
Public Const IMAGE_SYM_TYPE_LONG                As Integer = &H5                    '长度为4 个字节的带符号整数。
Public Const IMAGE_SYM_TYPE_FLOAT               As Integer = &H6                    '长度为4 个字节的浮点数。
Public Const IMAGE_SYM_TYPE_DOUBLE              As Integer = &H7                    '长度为8 个字节的浮点数。
Public Const IMAGE_SYM_TYPE_STRUCT              As Integer = &H8                    '结构体。
Public Const IMAGE_SYM_TYPE_UNION               As Integer = &H9                    '共用体。
Public Const IMAGE_SYM_TYPE_ENUM                As Integer = &HA                    ' enumeration.枚举类型。
Public Const IMAGE_SYM_TYPE_MOE                 As Integer = &HB                    ' member of enumeration.枚举类型成员（具体值）。
Public Const IMAGE_SYM_TYPE_BYTE                As Integer = &HC                    '字节；长度为1 个字节的无符号整数。
Public Const IMAGE_SYM_TYPE_WORD                As Integer = &HD                    '字；长度两个字节的无符号整数。
Public Const IMAGE_SYM_TYPE_UINT                As Integer = &HE                    '长度为自然尺寸的无符号整数（通常为4 个字节）。
Public Const IMAGE_SYM_TYPE_DWORD               As Integer = &HF                    '长度为4 个字节的无符号整数。
Public Const IMAGE_SYM_TYPE_PCODE               As Integer = &H8000                 '
' Type (derived) values.
Public Const IMAGE_SYM_DTYPE_NULL               As Integer = 0                      ' no derived type.非导出类型；此符号是简单的标量变量。
Public Const IMAGE_SYM_DTYPE_POINTER            As Integer = 1                      ' pointer.此符号是指向基本类型的指针。
Public Const IMAGE_SYM_DTYPE_FUNCTION           As Integer = 2                      ' function.此符号是返回基本类型的函数。
Public Const IMAGE_SYM_DTYPE_ARRAY              As Integer = 3                      ' array.此符号是由基本类型组成的数组。

'IMAGE_SYMBOL Storage classes. 注意StorageClass 域是长度为1 个字节的无符号整数。因此如果这个域的值为-1 的话，实际上应该被看作是与它相等的无符号数，也就是0xFF。尽管传统的COFF 格式使用许多存储类别，但是Microsoft 的工具使用Visual
'                               C++调试信息来表示大部分符号信息，它通常仅使用四种存储类别：EXTERNAL（2）、STATIC（3）、FUNCTION（101）和FILE（103）。
Public Const IMAGE_SYM_CLASS_END_OF_FUNCTION    As Byte = &HFF                      '(0xFF)表示函数结尾的特殊符号，用于调试。
Public Const IMAGE_SYM_CLASS_NULL               As Byte = &H0                       '未被赋予存储类别。
Public Const IMAGE_SYM_CLASS_AUTOMATIC          As Byte = &H1                       '自动（堆栈）变量。Value 域指出此变量在栈帧中的偏移
Public Const IMAGE_SYM_CLASS_EXTERNAL           As Byte = &H2                       'Microsoft 的工具使用此值来表示外部符号.如果SectionNumber 域为0（IMAGE_SYM_UNDEFINED），那么Value 域给出大小；如果SectionNumber 域不为0，那么Value 域给出节中的偏移
Public Const IMAGE_SYM_CLASS_STATIC             As Byte = &H3                       '符号在节中的偏移。如果Value 域为0，那么此符号表示节名
Public Const IMAGE_SYM_CLASS_REGISTER           As Byte = &H4                       '寄存器变量。Value 域给出寄存器编号。
Public Const IMAGE_SYM_CLASS_EXTERNAL_DEF       As Byte = &H5                       '在外部定义的符号。
Public Const IMAGE_SYM_CLASS_LABEL              As Byte = &H6                       '模块中定义的代码标号。Value 域给出此符号在节中的偏移
Public Const IMAGE_SYM_CLASS_UNDEFINED_LABEL    As Byte = &H7                       '引用的未定义的代码标号。
Public Const IMAGE_SYM_CLASS_MEMBER_OF_STRUCT   As Byte = &H8                       '结构体成员。Value 域指出是第几个成员。
Public Const IMAGE_SYM_CLASS_ARGUMENT           As Byte = &H9                       '函数的形式参数（形参）。Value 域指出是第几个参数
Public Const IMAGE_SYM_CLASS_STRUCT_TAG         As Byte = &HA                       '结构体名。
Public Const IMAGE_SYM_CLASS_MEMBER_OF_UNION    As Byte = &HB                       '共用体成员。Value 域指出是第几个成员。
Public Const IMAGE_SYM_CLASS_UNION_TAG          As Byte = &HC                       '共用体名。
Public Const IMAGE_SYM_CLASS_TYPE_DEFINITION    As Byte = &HD                       'Typedef 项。
Public Const IMAGE_SYM_CLASS_UNDEFINED_STATIC   As Byte = &HE                       '静态数据声明。
Public Const IMAGE_SYM_CLASS_ENUM_TAG           As Byte = &HF                       '枚举类型名。
Public Const IMAGE_SYM_CLASS_MEMBER_OF_ENUM     As Byte = &H10                      '枚举类型成员。Value 域指出是第几个成员
Public Const IMAGE_SYM_CLASS_REGISTER_PARAM     As Byte = &H11                      '寄存器参数。
Public Const IMAGE_SYM_CLASS_BIT_FIELD          As Byte = &H12                      '位域。Value 域指出是位域中的第几个位。

Public Const IMAGE_SYM_CLASS_FAR_EXTERNAL       As Byte = &H44                      '

Public Const IMAGE_SYM_CLASS_BLOCK              As Byte = &H64                      '.bb（beginning of block，块开头)或.eb 记录（end of block，块结尾）。Value 域是代码位置，它是一个可重定位的地址
Public Const IMAGE_SYM_CLASS_FUNCTION           As Byte = &H65                      'Microsoft 的工具用此值来表示定义函数范围的符号记录，这些符号记录分别是：.bf（begin function，函数开头）、.ef（endfunction，函数结尾）以及.lf（lines in function，函数中的行）。对于.lf 记录来说，Value 域给出了源代码中此函数所占的行数。对于.ef 记录来说，Value 域给出了函数代码的大小
Public Const IMAGE_SYM_CLASS_END_OF_STRUCT      As Byte = &H66                      '结构体末尾。
Public Const IMAGE_SYM_CLASS_FILE               As Byte = &H67                      'Microsoft 的工具以及传统COFF 格式都使用此值来表示源文件符号记录.这种符号表记录后面跟着给出文件名的辅助符号表记录
'new
Public Const IMAGE_SYM_CLASS_SECTION            As Byte = &H68                      '节的定义（Microsoft 的工具使用STATIC 存储类别代替）
Public Const IMAGE_SYM_CLASS_WEAK_EXTERNAL      As Byte = &H69                      '弱外部符号。

Public Const IMAGE_SYM_CLASS_CLR_TOKEN          As Byte = &H6B                      '表示CLR 记号的符号。它的名称是这个记号的十六进制值的ASCII 码表示?

'CLR 记号定义（仅适用于目标文件）
Public Type IMAGE_AUX_SYMBOL_TOKEN_DEF
    bAuxType            As Byte                 'IMAGE_AUX_SYMBOL_TYPE 必须为IMAGE_AUX_SYMBOL_TYPE_TOKEN_DEF（1）
    bReserved           As Byte                 ' Must be 0
    SymbolTableIndex    As Long                 '此CLR 记号定义涉及的COFF 符号在符号表中的索引。
    rgbReserved(11)     As Byte                 'Must be 0
End Type
'如果一个符号表记录拥有下列属性：存储类别为EXTERNAL（2）、Type 域的值表明它是一个函数（0x20）以及SectionNumber 域的值大于0，它就标志着函数的开头.注意如果一个符号表记录SectionNumber 域的值为IMAGE_SYM_UNDEFINED（0），那么它并不定义一个函数，也没有相应的辅助符号表记录
Public Type IMAGE_AUX_SYMBOL
    'Sym
    TagIndex            As Long                 'struct, union, or enum tag index.相应的.bf（函数开头）记录在符号表中的索引。
    TotalSize           As Long                 'union_Misc 函数经编译后生成的可执行代码的大小。如果此函数单独成节，那么根据对齐值的不同，节头中的SizeOfRawData 域可能大于或等于这个域
                                                '如果这个值为IMAGE_WEAK_EXTERN_SEARCH_NOLIBRARY，表明链接时不在库中查找sym1如果这个值为IMAGE_WEAK_EXTERN_SEARCH_LIBRARY，表明链接时在库中查找sym1.如果这个值为IMAGE_WEAK_EXTERN_SEARCH_ALIAS，表明sym1 是sym2 的别名。
'    Linenumber          As Integer              'union_Misc,declaration line number
'    Size                As Integer              'union_Misc,size of struct, union, or enum
    Dimension(3)        As Integer              'Array ,Union_FcnAry,if ISARY, up to 4 dimen.
'    PointerToLinenumber As Long                 'Function,Union_FcnAry  if ISFCN, tag, or .bb如果此函数存在行号记录，那么这个值表示它的第一个COFF 行号记录的文件偏移；如果不存在，那么这个值为0
'    PointerToNextFunction   As Long             'Function,Union_FcnAry   if ISFCN, tag, or .bb对应于下一个函数的符号表记录在符号表中的索引。如果此函数是符号表中的最后一个函数，那么这个域的值为0
    TvIndex             As Integer              'tv index
    'File
    vName(IMAGE_SIZEOF_SYMBOL - 1)  As Byte     '表示源文件名的ANSI 字符串。如果源文件名的长度小于最大长度，用NULL 填充。
    'Section
    Length              As Long                 'section length节中数据的大小。与节头的SizeOfRawData 域一样
    NumberOfRelocations As Integer              'number of relocation entries 此节中重定位项的数目。
    NumberOfLinenumbers As Integer              'number of line numbers 此节中行号信息项的数目。
    Checksum            As Long                 'checksum for communal 公共数据的校验和。只有节头中设置了IMAGE_SCN_LNK_COMDAT 标志时才使用此域
    Number              As Integer              'section number to associate with 与此节相关的节在节表中的索引（从1 开始）。当COMDAT 的Selection 域为5 时才使用这个域
    Selection           As Byte                 'communal selection type 表示COMDAT 选择方式的数字。这个域只用于COMDAT 节
    bReserved           As Byte                 '
    HighNumber          As Integer              'high bits of the section number

    TokenDef            As IMAGE_AUX_SYMBOL_TOKEN_DEF
    'CRC
    crc                 As Long
    rgbReserved(3)      As Byte
End Type

Public Type IMAGE_AUX_SYMBOL_EX
    'Sym
    WeakDefaultSymIndex As Long                 'the weak extern default symbol index
    WeakSearchType      As Long
    rgbReserved(11)     As Byte                 '
    'File
    vName(IMAGE_SIZEOF_SYMBOL_EX - 1)  As Byte
    'Section
    Length              As Long                 'section length
    NumberOfRelocations As Integer              'number of relocation entries
    NumberOfLinenumbers As Integer              'number of line numbers
    Checksum            As Long                 'checksum for communal
    Number              As Integer              'section number to associate with
    Selection           As Byte                 'communal selection type
    bReserved           As Byte                 '
    HighNumber          As Integer              'high bits of the section number
    rgbReserved1(1)     As Byte                 '

    TokenDef            As IMAGE_AUX_SYMBOL_TOKEN_DEF
    rgbReserved2(2)     As Byte
    'CRC
    crc                 As Long
    rgbReserved3(15)    As Byte
End Type

'IMAGE_AUX_SYMBOL Communal selection types.
Public Const IMAGE_COMDAT_SELECT_NODUPLICATES   As Byte = 1                         '如果此符号已经被定义过，链接器将生成一个“multiply defined symbol（符号多重定义）”错误。
Public Const IMAGE_COMDAT_SELECT_ANY            As Byte = 2                         '链接器从这些定义同一个COMDAT 符号的节中任选一个，其余（未被选中）的节都被移除。
Public Const IMAGE_COMDAT_SELECT_SAME_SIZE      As Byte = 3                         '链接器从定义这个符号的多个节中任选一个。如果所有这些定义大小不等，链接器将生成一个“符号多重定义”错误。
Public Const IMAGE_COMDAT_SELECT_EXACT_MATCH    As Byte = 4                         '链接器从定义这个符号的多个节中任选一个。如果所有这些定义不严格一致，链接器将生成一个“符号多重定义”错误。
Public Const IMAGE_COMDAT_SELECT_ASSOCIATIVE    As Byte = 5                         '如果“其它某个”COMDAT 节被链接的话，此节也要被链接。这里的“其它某个”节由与定义
                                                                                    '此节的符号表记录相关的辅助符号表记录的Number 域给出.这个设置对于那些在多个节中都有其相关部分（例如代码在一个节中而数据在另一个节中）但必须作为一个整体进行链接或丢弃的定义非常有用?与此节关联的这个
                                                                                    '“其它某个”节必须也是COMDAT 节并且它不能再与其它COMDAT 节关联（也就是说，这个“其它某个”节不能将Selection 域设置为IMAGE_COMDAT_SELECT_ASSOCIATIVE）。
Public Const IMAGE_COMDAT_SELECT_LARGEST        As Byte = 6                         '链接器从这个符号的所有定义中选取长度最大的进行链接。如果长度最大的不止一个，那么就在这几个最大的中任选一个
Public Const IMAGE_COMDAT_SELECT_NEWEST         As Byte = 7                         '

Public Const IMAGE_WEAK_EXTERN_SEARCH_NOLIBRARY As Byte = 1                         '
Public Const IMAGE_WEAK_EXTERN_SEARCH_LIBRARY   As Byte = 2                         '
Public Const IMAGE_WEAK_EXTERN_SEARCH_ALIAS     As Byte = 3                         '

Public Type IMAGE_RELOCATION
    'DUMMYUNIONNAME
    VirtualAddress      As Long                 'DUMMYUNIONNAME_VirtualAddress 需要进行重定位的代码或数据的地址。这是从节开头算起的偏移，加上节的RVA/Offset 域的值。
'    RelocCount          As Long                 'DUMMYUNIONNAME_RelocCount,Set to the real count when IMAGE_SCN_LNK_NRELOC_OVFL is set
    SymbolTableIndex    As Long                 '符号表的索引（从0 开始）。这个符号给出了用于重定位的地址。如果这个指定符号的存储类别为节，那么它的地址就是第一个与它同名的节的地址
    Type                As Byte                 '重定位类型。合法的重定位类型依赖于机器类型。
End Type

' I386 relocation types.
Public Const IMAGE_REL_I386_ABSOLUTE            As Byte = &H0                       ' Reference is absolute, no relocation is necessary 重定位被忽略。
Public Const IMAGE_REL_I386_DIR16               As Byte = &H1                       ' Direct 16-bit reference to the symbols virtual address 不支持。
Public Const IMAGE_REL_I386_REL16               As Byte = &H2                       ' PC-relative 16-bit reference to the symbols virtual address 不支持。
Public Const IMAGE_REL_I386_DIR32               As Byte = &H6                       ' Direct 32-bit reference to the symbols virtual address 重定位目标的32 位VA。
Public Const IMAGE_REL_I386_DIR32NB             As Byte = &H7                       ' Direct 32-bit reference to the symbols virtual address, base not included 重定位目标的32 位RVA。
Public Const IMAGE_REL_I386_SEG12               As Byte = &H9                       ' Direct 16-bit reference to the segment-selector bits of a 32-bit virtual address 不支持。
Public Const IMAGE_REL_I386_SECTION             As Byte = &HA                       '包含重定位目标的节的16 位索引。用于支持调试信息
Public Const IMAGE_REL_I386_SECREL              As Byte = &HB                       '重定位目标相对于它所在节开头的32 位偏移。用于支持调试信息和静态线程局部存储
Public Const IMAGE_REL_I386_TOKEN               As Byte = &HC                       ' clr token CLR 记号。
Public Const IMAGE_REL_I386_SECREL7             As Byte = &HD                       ' 7 bit offset from base of section containing target 相对于重定位目标所在节基地址的7 位偏移。
Public Const IMAGE_REL_I386_REL32               As Byte = &H14                      ' PC-relative 32-bit reference to the symbols virtual address 重定位目标的32 位相对偏移。用于支持x86 的相对分支和CALL 指令

' MIPS relocation types.
Public Const IMAGE_REL_MIPS_ABSOLUTE            As Byte = &H0                       ' Reference is absolute, no relocation is necessary 重定位被忽略
Public Const IMAGE_REL_MIPS_REFHALF             As Byte = &H1                       '重定位目标32 位VA 的高16 位。
Public Const IMAGE_REL_MIPS_REFWORD             As Byte = &H2                       '重定位目标的32 位VA。
Public Const IMAGE_REL_MIPS_JMPADDR             As Byte = &H3                       '重定位目标VA 的低26 位。用于支持MIPS 平台的J 和JAL 指令
Public Const IMAGE_REL_MIPS_REFHI               As Byte = &H4                       '重定位目标32 位VA 的高16 位。它用于加载一个完整地址所需的两指令序列中的第一条指令.这种重定位类型后面必须紧跟IMAGE_REL_MIPS_PAIR类型的重定位项，而后者的SymbolTableIndex 域包含的是一个16 位偏移（符号数），这个偏移要被加到重定位目标位置的高16 位
Public Const IMAGE_REL_MIPS_REFLO               As Byte = &H5                       '重定位目标VA 的低16 位。
Public Const IMAGE_REL_MIPS_GPREL               As Byte = &H6                       '重定位目标相对于GP 寄存器的16 位偏移（符号数）。
Public Const IMAGE_REL_MIPS_LITERAL             As Byte = &H7                       '与IMAGE_REL_MIPS_GPREL 相同。
Public Const IMAGE_REL_MIPS_SECTION             As Byte = &HA                       '包含重定位目标的节的16 位索引。用于支持调试信息
Public Const IMAGE_REL_MIPS_SECREL              As Byte = &HB                       '重定位目标相对于它所在节开头的32 位偏移。用于支持调试信息和静态线程局部存储
Public Const IMAGE_REL_MIPS_SECRELLO            As Byte = &HC                       ' Low 16-bit section relative referemce (used for >32k TLS) 重定位目标相对于它所在节开头的32 位偏移的低16 位
Public Const IMAGE_REL_MIPS_SECRELHI            As Byte = &HD                       ' High 16-bit section relative reference (used for >32k TLS) 重定位目标相对于它所在节开头的32 位VA 的高16 位.这种重定位类型后面必须紧跟IMAGE_REL_MIPS_PAIR 类型的重定位项，而后者的SymbolTableIndex 域包含的是一个16 位偏移（符号数），这个偏移要被加到重定位符号位置的高16 位
Public Const IMAGE_REL_MIPS_TOKEN               As Byte = &HE                       ' clr token
Public Const IMAGE_REL_MIPS_JMPADDR16           As Byte = &H10                      '重定位目标VA 的低26 位。用于支持MIPS16 的JAL 指令
Public Const IMAGE_REL_MIPS_REFWORDNB           As Byte = &H22                      '重定位目标的32 位RVA。
Public Const IMAGE_REL_MIPS_PAIR                As Byte = &H25                      '只有紧跟IMAGE_REL_MIPS_REFHI 或IMAGE_REL_MIPS_SECRELHI 类型的重定位时这种重定位类型才是合法的.此重定位项的SymbolTableIndex 域包含的是偏移而不是符号表索引

' Alpha Relocation types.
Public Const IMAGE_REL_ALPHA_ABSOLUTE           As Byte = &H0
Public Const IMAGE_REL_ALPHA_REFLONG            As Byte = &H1
Public Const IMAGE_REL_ALPHA_REFQUAD            As Byte = &H2
Public Const IMAGE_REL_ALPHA_GPREL32            As Byte = &H3
Public Const IMAGE_REL_ALPHA_LITERAL            As Byte = &H4
Public Const IMAGE_REL_ALPHA_LITUSE             As Byte = &H5
Public Const IMAGE_REL_ALPHA_GPDISP             As Byte = &H6
Public Const IMAGE_REL_ALPHA_BRADDR             As Byte = &H7
Public Const IMAGE_REL_ALPHA_HINT               As Byte = &H8
Public Const IMAGE_REL_ALPHA_INLINE_REFLONG     As Byte = &H9
Public Const IMAGE_REL_ALPHA_REFHI              As Byte = &HA
Public Const IMAGE_REL_ALPHA_REFLO              As Byte = &HB
Public Const IMAGE_REL_ALPHA_PAIR               As Byte = &HC
Public Const IMAGE_REL_ALPHA_MATCH              As Byte = &HD
Public Const IMAGE_REL_ALPHA_SECTION            As Byte = &HE
Public Const IMAGE_REL_ALPHA_SECREL             As Byte = &HF
Public Const IMAGE_REL_ALPHA_REFLONGNB          As Byte = &H10
Public Const IMAGE_REL_ALPHA_SECRELLO           As Byte = &H11                      ' Low 16-bit section relative reference
Public Const IMAGE_REL_ALPHA_SECRELHI           As Byte = &H12                      ' High 16-bit section relative reference
Public Const IMAGE_REL_ALPHA_REFQ3              As Byte = &H13                      ' High 16 bits of 48 bit reference
Public Const IMAGE_REL_ALPHA_REFQ2              As Byte = &H14                      ' Middle 16 bits of 48 bit reference
Public Const IMAGE_REL_ALPHA_REFQ1              As Byte = &H15                      ' Low 16 bits of 48 bit reference
Public Const IMAGE_REL_ALPHA_GPRELLO            As Byte = &H16                      ' Low 16-bit GP relative reference
Public Const IMAGE_REL_ALPHA_GPRELHI            As Byte = &H17                      ' High 16-bit GP relative reference

' IBM PowerPC relocation types.
Public Const IMAGE_REL_PPC_ABSOLUTE             As Byte = &H0                       ' NOP 重定位被忽略。
Public Const IMAGE_REL_PPC_ADDR64               As Byte = &H1                       ' 64-bit address 重定位目标的64 位VA。
Public Const IMAGE_REL_PPC_ADDR32               As Byte = &H2                       ' 32-bit address 重定位目标的32 位VA。
Public Const IMAGE_REL_PPC_ADDR24               As Byte = &H3                       ' 26-bit address, shifted left 2 (branch absolute) 重定位目标VA 的低24 位。只有当重定位目标符号是绝对符号且可以按符号扩展到它的原始值时才是合法的
Public Const IMAGE_REL_PPC_ADDR16               As Byte = &H4                       ' 16-bit address 重定位目标VA 的低16 位。
Public Const IMAGE_REL_PPC_ADDR14               As Byte = &H5                       ' 16-bit address, shifted left 2 (load doubleword) 重定位目标VA 的低14 位。只有当重定位目标符号是绝对符号且可以按符号扩展到它的原始值时才是合法的
Public Const IMAGE_REL_PPC_REL24                As Byte = &H6                       ' 26-bit PC-relative offset, shifted left 2 (branch relative) 符号位置相对于PC 的24 位偏移。
Public Const IMAGE_REL_PPC_REL14                As Byte = &H7                       ' 16-bit PC-relative offset, shifted left 2 (br cond relative) 符号位置相对于PC 的14 位偏移。
Public Const IMAGE_REL_PPC_TOCREL16             As Byte = &H8                       ' 16-bit offset from TOC base
Public Const IMAGE_REL_PPC_TOCREL14             As Byte = &H9                       ' 16-bit offset from TOC base, shifted left 2 (load doubleword)

Public Const IMAGE_REL_PPC_ADDR32NB             As Byte = &HA                       ' 32-bit addr w/o image base  重定位目标的32 位RVA。
Public Const IMAGE_REL_PPC_SECREL               As Byte = &HB                       ' va of containing section (as in an image sectionhdr) 重定位目标相对于它所在节开头的32 位偏移。用于支持调试信息和静态线程局部存储
Public Const IMAGE_REL_PPC_SECTION              As Byte = &HC                       ' sectionheader number包含重定位目标的节的16 位索引。用于支持调试信息
Public Const IMAGE_REL_PPC_IFGLUE               As Byte = &HD                       ' substitute TOC restore instruction iff symbol is glue code
Public Const IMAGE_REL_PPC_SECREL16             As Byte = &HF                       ' va of containing section (limited to 16 bits) 重定位目标相对于它所在节开头的16 位偏移。用于支持调试信息和静态线程局部存储
Public Const IMAGE_REL_PPC_IMGLUE               As Byte = &HE                       ' symbol is glue code; virtual address is TOC restore instruction
Public Const IMAGE_REL_PPC_REFHI                As Byte = &H10                      '重定位目标32 位VA 的高16 位。它用于加载一个完整地址所需的两指令序列中的第一条指令?这种重定位类型后面必须紧跟IMAGE_REL_PPC_PAIR类型的重定位项，而后者的SymbolTableIndex 域包含的是一个16 位偏移（符号数），这个偏移要被加到重定位目标位置的高16 位
Public Const IMAGE_REL_PPC_REFLO                As Byte = &H11                      '重定位目标VA 的低16 位。
Public Const IMAGE_REL_PPC_PAIR                 As Byte = &H12                      '只有紧跟IMAGE_REL_PPC_REFHI 或IMAGE_REL_PPC_SECRELHI 类型的重定位时这种重定位类型才是合法的.此重定位项的SymbolTableIndex 域包含的是偏移而不是符号表的索引
Public Const IMAGE_REL_PPC_SECRELLO             As Byte = &H13                      ' Low 16-bit section relative reference (used for >32k TLS) 重定位目标相对于它所在节开头的32 位偏移的低16 位
Public Const IMAGE_REL_PPC_SECRELHI             As Byte = &H14                      ' High 16-bit section relative reference (used for >32k TLS) 重定位目标相对于它所在节开头的32 位偏移的高16 位
Public Const IMAGE_REL_PPC_GPREL                As Byte = &H15                      '重定位目标相对于GP 寄存器的16 位偏移（带符号数）。
Public Const IMAGE_REL_PPC_TOKEN                As Byte = &H16                      ' clr token CLR 记号。

Public Const IMAGE_REL_PPC_TYPEMASK             As Byte = &HFF                      ' mask to isolate above values in IMAGE_RELOCATION.Type

' Flag bits in IMAGE_RELOCATION.TYPE
Public Const IMAGE_REL_PPC_NEG                  As Integer = &H100                  ' subtract reloc value rather than adding it
Public Const IMAGE_REL_PPC_BRTAKEN              As Integer = &H200                  ' fix branch prediction bit to predict branch taken
Public Const IMAGE_REL_PPC_BRNTAKEN             As Integer = &H400                  ' fix branch prediction bit to predict branch not taken
Public Const IMAGE_REL_PPC_TOCDEFN              As Integer = &H800                  ' toc slot defined in file (or, data in toc)


' Hitachi SH3 relocation types.
Public Const IMAGE_REL_SH3_ABSOLUTE             As Byte = &H0                       ' No relocation 重定位被忽略。
Public Const IMAGE_REL_SH3_DIRECT16             As Byte = &H1                       ' 16 bit direct 对包含重定位目标符号VA 的16 位单元的引用。
Public Const IMAGE_REL_SH3_DIRECT32             As Byte = &H2                       ' 32 bit direct 重定位目标符号的32 位VA。
Public Const IMAGE_REL_SH3_DIRECT8              As Byte = &H3                       ' 8 bit direct, -128..255 对包含重定位目标符号VA 的8 位单元的引用。
Public Const IMAGE_REL_SH3_DIRECT8_WORD         As Byte = &H4                       ' 8 bit direct .W (0 ext.)对包含重定位目标符号16 位有效VA 的8 位指令的引用
Public Const IMAGE_REL_SH3_DIRECT8_LONG         As Byte = &H5                       ' 8 bit direct .L (0 ext.)对包含重定位目标符号32 位有效VA 的8 位指令的引用
Public Const IMAGE_REL_SH3_DIRECT4              As Byte = &H6                       ' 4 bit direct (0 ext.)对其低4 位包含重定位目标符号VA 的8 位单元的引用
Public Const IMAGE_REL_SH3_DIRECT4_WORD         As Byte = &H7                       ' 4 bit direct .W (0 ext.)对其低4 位包含重定位目标符号16 位有效VA 的8 位指令的引用
Public Const IMAGE_REL_SH3_DIRECT4_LONG         As Byte = &H8                       ' 4 bit direct .L (0 ext.)对其低4 位包含重定位目标符号32 位有效VA 的8 位指令的引用
Public Const IMAGE_REL_SH3_PCREL8_WORD          As Byte = &H9                       ' 8 bit PC relative .W 对包含重定位目标符号16 位有效相对偏移的8位指令的引用
Public Const IMAGE_REL_SH3_PCREL8_LONG          As Byte = &HA                       ' 8 bit PC relative .L 对包含重定位目标符号32 位有效相对偏移的8位指令的引用
Public Const IMAGE_REL_SH3_PCREL12_WORD         As Byte = &HB                       ' 12 LSB PC relative .W 对其低12 位包含重定位目标符号16 位有效相对偏移的16 位指令的引用
Public Const IMAGE_REL_SH3_STARTOF_SECTION      As Byte = &HC                       ' Start of EXE section 对包含重定位目标符号所在节VA 的32 位单元的引用
Public Const IMAGE_REL_SH3_SIZEOF_SECTION       As Byte = &HD                       ' Size of EXE section 对包含重定位目标符号所在节大小的32 位单元的引用
Public Const IMAGE_REL_SH3_SECTION              As Byte = &HE                       ' Section table index 包含重定位目标的节的16 位索引。用于支持调试信息
Public Const IMAGE_REL_SH3_SECREL               As Byte = &HF                       ' Offset within section 重定位目标相对于它所在节开头的32 位偏移。用于支持调试信息和静态线程局部存储
Public Const IMAGE_REL_SH3_DIRECT32_NB          As Byte = &H10                      ' 32 bit direct not based 重定位目标符号的32 位RVA。
Public Const IMAGE_REL_SH3_GPREL4_LONG          As Byte = &H11                      ' GP-relative addressing    与GP 相关。
Public Const IMAGE_REL_SH3_TOKEN                As Byte = &H12                      ' clr token     CLR 记号。
Public Const IMAGE_REL_SHM_PCRELPT              As Byte = &H13                      ' Offset from current 距当前指令的偏移（长字）。如果没有设置IMAGE_REL_SHM_NOMODE 标志，那么将低位取反插入到第32 位以选择PTA 指令或PTB 指令。
                                                                                    '  instruction in longwords
                                                                                    '  if not NOMODE, insert the
                                                                                    '  inverse of the low bit at
                                                                                    '  bit 32 to select PTA/PTB
Public Const IMAGE_REL_SHM_REFLO                As Byte = &H14                      ' Low bits of 32-bit address  32 位地址的低16 位。
Public Const IMAGE_REL_SHM_REFHALF              As Byte = &H15                      ' High bits of 32-bit address 32 位地址的高16 位。
Public Const IMAGE_REL_SHM_RELLO                As Byte = &H16                      ' Low bits of relative reference 相对地址的低16 位。
Public Const IMAGE_REL_SHM_RELHALF              As Byte = &H17                      ' High bits of relative reference 相对地址的高16 位。
Public Const IMAGE_REL_SHM_PAIR                 As Byte = &H18                      ' offset operand for relocation只有紧跟IMAGE_REL_SHM_REFHALF、IMAGE_REL_SHM_RELLO 或IMAGE_REL_SHM_RELHALF 类型的重定位项时这种重定位类型才是合法的.此重定位项的SymbolTableIndex 域包含的是偏移而不是符号表的索引

Public Const IMAGE_REL_SH_NOMODE                As Integer = &H8000                 ' relocation ignores section mode 重定位忽略节模式。


Public Const IMAGE_REL_ARM_ABSOLUTE             As Byte = &H0                       ' No relocation required 重定位被忽略。
Public Const IMAGE_REL_ARM_ADDR32               As Byte = &H1                       ' 32 bit address 重定位目标的32 位VA。
Public Const IMAGE_REL_ARM_ADDR32NB             As Byte = &H2                       ' 32 bit address w/o image base 重定位目标的32 位RVA。
Public Const IMAGE_REL_ARM_BRANCH24             As Byte = &H3                       ' 24 bit offset << 2 & sign ext. 重定位目标的24 位相对偏移。
Public Const IMAGE_REL_ARM_BRANCH11             As Byte = &H4                       ' Thumb: 2 11 bit offsets 对子程序调用的引用。这个引用由两个16 位指令
Public Const IMAGE_REL_ARM_TOKEN                As Byte = &H5                       ' clr token
Public Const IMAGE_REL_ARM_GPREL12              As Byte = &H6                       ' GP-relative addressing (ARM)
Public Const IMAGE_REL_ARM_GPREL7               As Byte = &H7                       ' GP-relative addressing (Thumb)
Public Const IMAGE_REL_ARM_BLX24                As Byte = &H8
Public Const IMAGE_REL_ARM_BLX11                As Byte = &H9
Public Const IMAGE_REL_ARM_SECTION              As Byte = &HE                       ' Section table index 包含重定位目标的节的16 位索引。用于支持调试信息
Public Const IMAGE_REL_ARM_SECREL               As Byte = &HF                       ' Offset within section重定位目标相对于它所在节开头的32 位偏移。用于支持调试信息和静态线程局部存储
Public Const IMAGE_REL_ARM_MOV32A               As Byte = &H10                      ' ARM: MOVW/MOVT
Public Const IMAGE_REL_ARM_MOV32                As Byte = &H10                      ' ARM: MOVW/MOVT (deprecated)
Public Const IMAGE_REL_ARM_MOV32T               As Byte = &H11                      ' Thumb: MOVW/MOVT
Public Const IMAGE_REL_THUMB_MOV32              As Byte = &H11                      ' Thumb: MOVW/MOVT (deprecated)
Public Const IMAGE_REL_ARM_BRANCH20T            As Byte = &H12                      ' Thumb: 32-bit conditional B
Public Const IMAGE_REL_THUMB_BRANCH20           As Byte = &H12                      ' Thumb: 32-bit conditional B (deprecated)
Public Const IMAGE_REL_ARM_BRANCH24T            As Byte = &H14                      ' Thumb: 32-bit B or BL
Public Const IMAGE_REL_THUMB_BRANCH24           As Byte = &H14                      ' Thumb: 32-bit B or BL (deprecated)
Public Const IMAGE_REL_ARM_BLX23T               As Byte = &H15                      ' Thumb: BLX immediate
Public Const IMAGE_REL_THUMB_BLX23              As Byte = &H15                      ' Thumb: BLX immediate (deprecated)

Public Const IMAGE_REL_AM_ABSOLUTE              As Byte = &H0
Public Const IMAGE_REL_AM_ADDR32                As Byte = &H1
Public Const IMAGE_REL_AM_ADDR32NB              As Byte = &H2
Public Const IMAGE_REL_AM_CALL32                As Byte = &H3
Public Const IMAGE_REL_AM_FUNCINFO              As Byte = &H4
Public Const IMAGE_REL_AM_REL32_1               As Byte = &H5
Public Const IMAGE_REL_AM_REL32_2               As Byte = &H6
Public Const IMAGE_REL_AM_SECREL                As Byte = &H7
Public Const IMAGE_REL_AM_SECTION               As Byte = &H8
Public Const IMAGE_REL_AM_TOKEN                 As Byte = &H9

' x64 relocations
Public Const IMAGE_REL_AMD64_ABSOLUTE           As Byte = &H0                       ' Reference is absolute, no relocation is necessary 重定位被忽略。
Public Const IMAGE_REL_AMD64_ADDR64             As Byte = &H1                       ' 64-bit address (VA). 重定位目标的64 位VA。
Public Const IMAGE_REL_AMD64_ADDR32             As Byte = &H2                       ' 32-bit address (VA). 重定位目标的32 位VA。
Public Const IMAGE_REL_AMD64_ADDR32NB           As Byte = &H3                       ' 32-bit address w/o image base (RVA). 不包含映像基址的32 位地址（RVA）。
Public Const IMAGE_REL_AMD64_REL32              As Byte = &H4                       ' 32-bit relative address from byte following reloc 相对于重定位目标的32 位地址。
Public Const IMAGE_REL_AMD64_REL32_1            As Byte = &H5                       ' 32-bit relative address from byte distance 1 from reloc 相对于距重定位目标1 字节处的32 位地址。
Public Const IMAGE_REL_AMD64_REL32_2            As Byte = &H6                       ' 32-bit relative address from byte distance 2 from reloc 相对于距重定位目标2 字节处的32 位地址。
Public Const IMAGE_REL_AMD64_REL32_3            As Byte = &H7                       ' 32-bit relative address from byte distance 3 from reloc 相对于距重定位目标3 字节处的32 位地址。
Public Const IMAGE_REL_AMD64_REL32_4            As Byte = &H8                       ' 32-bit relative address from byte distance 4 from reloc 相对于距重定位目标4 字节处的32 位地址。
Public Const IMAGE_REL_AMD64_REL32_5            As Byte = &H9                       ' 32-bit relative address from byte distance 5 from reloc 相对于距重定位目标5 字节处的32 位地址。
Public Const IMAGE_REL_AMD64_SECTION            As Byte = &HA                       ' Section index  包含重定位目标的节的16 位索引。用于支持调试信息
Public Const IMAGE_REL_AMD64_SECREL             As Byte = &HB                       ' 32 bit offset from base of section containing target 重定位目标相对于它所在节开头的32 位偏移。用于支持调试信息和静态线程局部存储?
Public Const IMAGE_REL_AMD64_SECREL7            As Byte = &HC                       ' 7 bit unsigned offset from base of section containing target 相对于重定位目标所在节基地址的7 位偏移（无符号数）。
Public Const IMAGE_REL_AMD64_TOKEN              As Byte = &HD                       ' 32 bit metadata token    CLR 记号。
Public Const IMAGE_REL_AMD64_SREL32             As Byte = &HE                       ' 32 bit signed span-dependent value emitted into object 放入目标文件中的32 位跨度依赖值（符号数）
Public Const IMAGE_REL_AMD64_PAIR               As Byte = &HF                       '与跨度依赖值成对出现，它必须紧跟每一个跨度依赖值
Public Const IMAGE_REL_AMD64_SSPAN32            As Byte = &H10                      ' 32 bit signed span-dependent value applied at link time 链接时应用的32 位跨度依赖值（符号数）。

' IA64 relocation types.
Public Const IMAGE_REL_IA64_ABSOLUTE            As Byte = &H0                       '重定位被忽略。
Public Const IMAGE_REL_IA64_IMM14               As Byte = &H1                       '这种指令重定位后面可以跟着IMAGE_REL_IA64_ADDEND 类型的重定位项，而后者的值在被插入到IMM14 指令包的指定的指令槽中之前被加到目标地址上.这种重定位目标必须是绝对符号，否则这个映像必须被修正。
Public Const IMAGE_REL_IA64_IMM22               As Byte = &H2                       '这种指令重定位后面可以跟着IMAGE_REL_IA64_ADDEND 类型的重定位项，而后者的值在被插入到IMM22 指令包的指定的指令槽中之前被加到目标地址上.这种重定位目标必须是绝对符号，否则这个映像必须被修正。
Public Const IMAGE_REL_IA64_IMM64               As Byte = &H3                       '这种重定位项的指令槽编号必须为1。这种重定位后面可以跟着IMAGE_REL_IA64_ADDEND 类型的重定位项，而后者的值在被存储到IMM64 指令包的三个指令槽中之前被加到目标地址上
Public Const IMAGE_REL_IA64_DIR32               As Byte = &H4                       '重定位目标的32 位VA。仅支持使用/LARGEADDRESSAWARE:NO 链接器选项生成的映像。
Public Const IMAGE_REL_IA64_DIR64               As Byte = &H5                       '重定位目标的64 位VA。
Public Const IMAGE_REL_IA64_PCREL21B            As Byte = &H6                       '使用按16 位边界对齐的重定位目标的25 位相对偏移来修正指令。这个偏移的低4 位全为0，因此并没有被存储
Public Const IMAGE_REL_IA64_PCREL21M            As Byte = &H7                       '使用按16 位边界对齐的重定位目标的25 位相对偏移来修正指令。这个偏移的低4 位全为0，因此并没有被存储
Public Const IMAGE_REL_IA64_PCREL21F            As Byte = &H8                       '这种重定位目标偏移的LSB 部分包含的是指令槽编号，其余部分包含的是指令包的地址。使用按16 位边界对齐的重定位目标的25 位相对偏移来修正指令。这个偏移的低4 位全为0，因此并没有被存储
Public Const IMAGE_REL_IA64_GPREL22             As Byte = &H9                       '这种指令重定位后面可以跟着IMAGE_REL_IA64_ADDEND 类型的重定位项，后者的值被加到目标地址上，而后计算GPREL22 指令包相对于GP 的偏移并应用
Public Const IMAGE_REL_IA64_LTOFF22             As Byte = &HA                       '使用重定位目标符号的常量表项相对于GP 的22位偏移来修正指令?链接器根据这个重定位项以及可能跟着它的IMAGE_REL_IA64_ADDEND 类型的重定位项来创建这个常量表项
Public Const IMAGE_REL_IA64_SECTION             As Byte = &HB                       '包含重定位目标的节的16 位索引。用于支持调试信息
Public Const IMAGE_REL_IA64_SECREL22            As Byte = &HC                       '使用重定位目标相对于它所在节开头的22 位偏移来修正指令.这种类型的重定位项后面可以紧跟着IMAGE_REL_IA64_ADDEND 类型的重定位项，后者的Value 域包含重定位目标相对于它所在节开头的32 位偏移（无符号数）。
Public Const IMAGE_REL_IA64_SECREL64I           As Byte = &HD                       '这种重定位项的指令槽编号必须为1。使用重定位目标相对于它所在节开头的64 位偏移来修正指令.这种类型的重定位项后面可以紧跟着IMAGE_REL_IA64_ADDEND 类型的重定位项，后者的Value 域包含重定位目标相对于它所在节开头的32 位偏移（无符号数）。
Public Const IMAGE_REL_IA64_SECREL32            As Byte = &HE                       '使用重定位目标相对于它所在节开头的32 位偏移来修正的数据的地址

Public Const IMAGE_REL_IA64_DIR32NB             As Byte = &H10                      '目标的32 位RVA。
Public Const IMAGE_REL_IA64_SREL14              As Byte = &H11                      '用于包含两个重定位目标之差的14 位立即数（符号数）。对于链接器来说这是一个说明域，表明编译器已经生成了这个值
Public Const IMAGE_REL_IA64_SREL22              As Byte = &H12                      '用于包含两个重定位目标之差的22 位立即数（符号数）。对于链接器来说这是一个说明域，表明编译器已经生成了这个值
Public Const IMAGE_REL_IA64_SREL32              As Byte = &H13                      '用于包含两个重定位目标之差的32 位立即数（符号数）。对于链接器来说这是一个说明域，表明编译器已经生成了这个值
Public Const IMAGE_REL_IA64_UREL32              As Byte = &H14                      '用于包含两个重定位目标之差的32 位立即数（无符号数）。对于链接器来说这是一个说明域，表明编译器已经生成了这个值
Public Const IMAGE_REL_IA64_PCREL60X            As Byte = &H15                      ' This is always a BRL and never converted 相对于PC 的60 位修正，用于MLX 指令包的BRL指令
Public Const IMAGE_REL_IA64_PCREL60B            As Byte = &H16                      ' If possible, convert to MBB bundle with NOP.B in slot 1 相对于PC 的60 位修正。如果重定位目标偏移不超过一个25 位域所能表示的范围（符号数），那么就在1 号指令槽中使用NOP.B 指令、2 号指令槽中使用25 位（最低4 位全为0，舍弃）的BR 指令将整个指令包转换成MBB 指令包
Public Const IMAGE_REL_IA64_PCREL60F            As Byte = &H17                      ' If possible, convert to MFB bundle with NOP.F in slot 1 相对于PC 的60 位修正。如果重定位目标偏移不超过一个25 位域所能表示的范围（符号数），那么就在1 号指令槽中使用NOP.F 指令、2 号指令槽中使用25 位（最低4 位全为0，舍弃）的BR 指令将整个指令包转换成MFB 指令包
Public Const IMAGE_REL_IA64_PCREL60I            As Byte = &H18                      ' If possible, convert to MIB bundle with NOP.I in slot 1 相对于PC 的60 位修正。如果重定位目标偏移不超过一个25 位域所能表示的范围（符号数），那么就在1 号指令槽中使用NOP.I 指令、2 号指令槽中使用25 位（最低4 位全为0，舍弃）的BR 指令将整个指令包转换成MIB 指令包
Public Const IMAGE_REL_IA64_PCREL60M            As Byte = &H19                      ' If possible, convert to MMB bundle with NOP.M in slot 1 相对于PC 的60 位修正。如果重定位目标偏移不超过一个25 位域所能表示的范围（符号数），那么就在1 号指令槽中使用NOP.M 指令、2 号指令槽中使用25 位（最低4 位全为0，舍弃）的BR 指令将整个指令包转换成MMB 指令包
Public Const IMAGE_REL_IA64_IMMGPREL64          As Byte = &H1A                      '相对于GP 的64 位修正。
Public Const IMAGE_REL_IA64_TOKEN               As Byte = &H1B                      ' clr token CLR 记号。
Public Const IMAGE_REL_IA64_GPREL32             As Byte = &H1C                      '相对于GP 的32 位修正。
Public Const IMAGE_REL_IA64_ADDEND              As Byte = &H1F                      '只有紧跟下列类型的重定位时这种重定位类型才是合法的: IMAGE_REL_IA64_IMM14?IMAGE_REL_IA64_IMM22 IMAGE_REL_IA64_IMM64 IMAGE_REL_IA64_GPREL22 IMAGE_REL_IA64_LTOFF22 IMAGE_REL_IA64_LTOFF64 IMAGE_REL_IA64_SECREL22 IMAGE_REL_IA64_SECREL64I 或IMAGE_REL_IA64_SECREL32.它的值是应用到指令包中的指令上的加数，而不是用于数据。

' CEF relocation types.
Public Const IMAGE_REL_CEF_ABSOLUTE             As Byte = &H0                       ' Reference is absolute, no relocation is necessary
Public Const IMAGE_REL_CEF_ADDR32               As Byte = &H1                       ' 32-bit address (VA).
Public Const IMAGE_REL_CEF_ADDR64               As Byte = &H2                       ' 64-bit address (VA).
Public Const IMAGE_REL_CEF_ADDR32NB             As Byte = &H3                       ' 32-bit address w/o image base (RVA).
Public Const IMAGE_REL_CEF_SECTION              As Byte = &H4                       ' Section index
Public Const IMAGE_REL_CEF_SECREL               As Byte = &H5                       ' 32 bit offset from base of section containing target
Public Const IMAGE_REL_CEF_TOKEN                As Byte = &H6                       ' 32 bit metadata token

' clr relocation types.
Public Const IMAGE_REL_CEE_ABSOLUTE             As Byte = &H0                       ' Reference is absolute, no relocation is necessary
Public Const IMAGE_REL_CEE_ADDR32               As Byte = &H1                       ' 32-bit address (VA).
Public Const IMAGE_REL_CEE_ADDR64               As Byte = &H2                       ' 64-bit address (VA).
Public Const IMAGE_REL_CEE_ADDR32NB             As Byte = &H3                       ' 32-bit address w/o image base (RVA).
Public Const IMAGE_REL_CEE_SECTION              As Byte = &H4                       ' Section index
Public Const IMAGE_REL_CEE_SECREL               As Byte = &H5                       ' 32 bit offset from base of section containing target
Public Const IMAGE_REL_CEE_TOKEN                As Byte = &H6                       ' 32 bit metadata token


Public Const IMAGE_REL_M32R_ABSOLUTE            As Byte = &H0                       ' No relocation required  重定位被忽略。
Public Const IMAGE_REL_M32R_ADDR32              As Byte = &H1                       ' 32 bit address 重定位目标的32 位VA。
Public Const IMAGE_REL_M32R_ADDR32NB            As Byte = &H2                       ' 32 bit address w/o image base 重定位目标的32 位RVA。
Public Const IMAGE_REL_M32R_ADDR24              As Byte = &H3                       ' 24 bit address 重定位目标的24 位VA。
Public Const IMAGE_REL_M32R_GPREL16             As Byte = &H4                       ' GP relative addressing 重定位目标相对于GP 寄存器的16 位偏移。
Public Const IMAGE_REL_M32R_PCREL24             As Byte = &H5                       ' 24 bit offset << 2 & sign ext. 重定位目标相对于程序计数器（PC）的24 位偏移，已经左移2 位并按符号扩展。
Public Const IMAGE_REL_M32R_PCREL16             As Byte = &H6                       ' 16 bit offset << 2 & sign ext. 重定位目标相对于PC 的16 位偏移，已经左移2位并按符号扩展
Public Const IMAGE_REL_M32R_PCREL8              As Byte = &H7                       ' 8 bit offset << 2 & sign ext.  重定位目标相对于PC 的8 位偏移，已经左移2 位并按符号扩展
Public Const IMAGE_REL_M32R_REFHALF             As Byte = &H8                       ' 16 MSBs 重定位目标VA 的16 位MSB。
Public Const IMAGE_REL_M32R_REFHI               As Byte = &H9                       ' 16 MSBs adj for LSB sign ext. 重定位目标VA 的16 位MSB，已经按LSB 符号扩展调整.它用于加载一个完整的32 位地址所需的两指令序列中的第一条指令.这种重定位类型后面必须紧跟IMAGE_REL_M32R_PAIR 类型的重定位项，而后者的SymbolTableIndex 域包含的是一个16 位偏移（符号数），这个偏移要被加到重定位符号位置的高16 位
Public Const IMAGE_REL_M32R_REFLO               As Byte = &HA                       ' 16 LSBs 重定位目标VA 的16 位LSB。
Public Const IMAGE_REL_M32R_PAIR                As Byte = &HB                       ' Link HI and LO 这种类型的重定位必须紧跟类型为IMAGE_REL_M32R_REFHI 的重定位项.此重定位项的SymbolTableIndex 域包含的是偏移而不是符号表索引
Public Const IMAGE_REL_M32R_SECTION             As Byte = &HC                       ' Section table index 包含重定位目标的节的16 位索引。用于支持调试信息
Public Const IMAGE_REL_M32R_SECREL32            As Byte = &HD                       ' 32 bit section relative reference重定位目标相对于它所在节开头的32 位偏移。用于支持调试信息和静态线程局部存储
Public Const IMAGE_REL_M32R_TOKEN               As Byte = &HE                       ' clr token CLR 记号。

Public Const IMAGE_REL_EBC_ABSOLUTE             As Byte = &H0                       ' No relocation required
Public Const IMAGE_REL_EBC_ADDR32NB             As Byte = &H1                       ' 32 bit address w/o image base
Public Const IMAGE_REL_EBC_REL32                As Byte = &H2                       ' 32-bit relative address from byte following reloc
Public Const IMAGE_REL_EBC_SECTION              As Byte = &H3                       ' Section table index
Public Const IMAGE_REL_EBC_SECREL               As Byte = &H4                       ' Offset within section

Public Const EMARCH_ENC_I17_IMM7B_INST_WORD_X   As Byte = 3                         ' Intel-IA64-Filler
Public Const EMARCH_ENC_I17_IMM7B_SIZE_X        As Byte = 7                         '  Intel-IA64-Filler
Public Const EMARCH_ENC_I17_IMM7B_INST_WORD_POS_X   As Byte = 4                     '  Intel-IA64-Filler
Public Const EMARCH_ENC_I17_IMM7B_VAL_POS_X     As Byte = 0                         '  Intel-IA64-Filler

Public Const EMARCH_ENC_I17_IMM9D_INST_WORD_X   As Byte = 3                         '  Intel-IA64-Filler
Public Const EMARCH_ENC_I17_IMM9D_SIZE_X        As Byte = 9                         '  Intel-IA64-Filler
Public Const EMARCH_ENC_I17_IMM9D_INST_WORD_POS_X   As Byte = 18                    '  Intel-IA64-Filler
Public Const EMARCH_ENC_I17_IMM9D_VAL_POS_X     As Byte = 7                         '  Intel-IA64-Filler

Public Const EMARCH_ENC_I17_IMM5C_INST_WORD_X   As Byte = 3                         '  Intel-IA64-Filler
Public Const EMARCH_ENC_I17_IMM5C_SIZE_X        As Byte = 5                         '  Intel-IA64-Filler
Public Const EMARCH_ENC_I17_IMM5C_INST_WORD_POS_X   As Byte = 13                    '  Intel-IA64-Filler
Public Const EMARCH_ENC_I17_IMM5C_VAL_POS_X     As Byte = 16                        '  Intel-IA64-Filler

Public Const EMARCH_ENC_I17_IC_INST_WORD_X      As Byte = 3                         '  Intel-IA64-Filler
Public Const EMARCH_ENC_I17_IC_SIZE_X           As Byte = 1                         '  Intel-IA64-Filler
Public Const EMARCH_ENC_I17_IC_INST_WORD_POS_X  As Byte = 12                        '  Intel-IA64-Filler
Public Const EMARCH_ENC_I17_IC_VAL_POS_X        As Byte = 21                        '  Intel-IA64-Filler

Public Const EMARCH_ENC_I17_IMM41a_INST_WORD_X  As Byte = 1                         '  Intel-IA64-Filler
Public Const EMARCH_ENC_I17_IMM41a_SIZE_X       As Byte = 10                        '  Intel-IA64-Filler
Public Const EMARCH_ENC_I17_IMM41a_INST_WORD_POS_X  As Byte = 14                    '  Intel-IA64-Filler
Public Const EMARCH_ENC_I17_IMM41a_VAL_POS_X    As Byte = 22                        '  Intel-IA64-Filler

Public Const EMARCH_ENC_I17_IMM41b_INST_WORD_X  As Byte = 1                         '  Intel-IA64-Filler
Public Const EMARCH_ENC_I17_IMM41b_SIZE_X       As Byte = 8                         '  Intel-IA64-Filler
Public Const EMARCH_ENC_I17_IMM41b_INST_WORD_POS_X  As Byte = 24                    '  Intel-IA64-Filler
Public Const EMARCH_ENC_I17_IMM41b_VAL_POS_X    As Byte = 32                        '  Intel-IA64-Filler

Public Const EMARCH_ENC_I17_IMM41c_INST_WORD_X  As Byte = 2                         '  Intel-IA64-Filler
Public Const EMARCH_ENC_I17_IMM41c_SIZE_X       As Byte = 23                        '  Intel-IA64-Filler
Public Const EMARCH_ENC_I17_IMM41c_INST_WORD_POS_X  As Byte = 0                     '  Intel-IA64-Filler
Public Const EMARCH_ENC_I17_IMM41c_VAL_POS_X    As Byte = 40                        '  Intel-IA64-Filler

Public Const EMARCH_ENC_I17_SIGN_INST_WORD_X    As Byte = 3                         '  Intel-IA64-Filler
Public Const EMARCH_ENC_I17_SIGN_SIZE_X         As Byte = 1                         '  Intel-IA64-Filler
Public Const EMARCH_ENC_I17_SIGN_INST_WORD_POS_X    As Byte = 27                    '  Intel-IA64-Filler
Public Const EMARCH_ENC_I17_SIGN_VAL_POS_X      As Byte = 63                        '  Intel-IA64-Filler

Public Const X3_OPCODE_INST_WORD_X              As Byte = 3                         '  Intel-IA64-Filler
Public Const X3_OPCODE_SIZE_X                   As Byte = 4                         '  Intel-IA64-Filler
Public Const X3_OPCODE_INST_WORD_POS_X          As Byte = 28                        '  Intel-IA64-Filler
Public Const X3_OPCODE_SIGN_VAL_POS_X           As Byte = 0                         '  Intel-IA64-Filler

Public Const X3_I_INST_WORD_X                   As Byte = 3                         '  Intel-IA64-Filler
Public Const X3_I_SIZE_X                        As Byte = 1                         '  Intel-IA64-Filler
Public Const X3_I_INST_WORD_POS_X               As Byte = 27                        '  Intel-IA64-Filler
Public Const X3_I_SIGN_VAL_POS_X                As Byte = 59                        '  Intel-IA64-Filler

Public Const X3_D_WH_INST_WORD_X                As Byte = 3                         '  Intel-IA64-Filler
Public Const X3_D_WH_SIZE_X                     As Byte = 3                         '  Intel-IA64-Filler
Public Const X3_D_WH_INST_WORD_POS_X            As Byte = 24                        '  Intel-IA64-Filler
Public Const X3_D_WH_SIGN_VAL_POS_X             As Byte = 0                         '  Intel-IA64-Filler

Public Const X3_IMM20_INST_WORD_X               As Byte = 3                         '  Intel-IA64-Filler
Public Const X3_IMM20_SIZE_X                    As Byte = 20                        '  Intel-IA64-Filler
Public Const X3_IMM20_INST_WORD_POS_X           As Byte = 4                         '  Intel-IA64-Filler
Public Const X3_IMM20_SIGN_VAL_POS_X            As Byte = 0                         '  Intel-IA64-Filler

Public Const X3_IMM39_1_INST_WORD_X             As Byte = 2                         '  Intel-IA64-Filler
Public Const X3_IMM39_1_SIZE_X                  As Byte = 23                        '  Intel-IA64-Filler
Public Const X3_IMM39_1_INST_WORD_POS_X         As Byte = 0                         '  Intel-IA64-Filler
Public Const X3_IMM39_1_SIGN_VAL_POS_X          As Byte = 36                        '  Intel-IA64-Filler

Public Const X3_IMM39_2_INST_WORD_X             As Byte = 1                         '  Intel-IA64-Filler
Public Const X3_IMM39_2_SIZE_X                  As Byte = 16                        '  Intel-IA64-Filler
Public Const X3_IMM39_2_INST_WORD_POS_X         As Byte = 16                        '  Intel-IA64-Filler
Public Const X3_IMM39_2_SIGN_VAL_POS_X          As Byte = 20                        '  Intel-IA64-Filler

Public Const X3_P_INST_WORD_X                   As Byte = 3                         '  Intel-IA64-Filler
Public Const X3_P_SIZE_X                        As Byte = 4                         '  Intel-IA64-Filler
Public Const X3_P_INST_WORD_POS_X               As Byte = 0                         '  Intel-IA64-Filler
Public Const X3_P_SIGN_VAL_POS_X                As Byte = 0                         '  Intel-IA64-Filler

Public Const X3_TMPLT_INST_WORD_X               As Byte = 0                         '  Intel-IA64-Filler
Public Const X3_TMPLT_SIZE_X                    As Byte = 4                         '  Intel-IA64-Filler
Public Const X3_TMPLT_INST_WORD_POS_X           As Byte = 0                         '  Intel-IA64-Filler
Public Const X3_TMPLT_SIGN_VAL_POS_X            As Byte = 0                         '  Intel-IA64-Filler

Public Const X3_BTYPE_QP_INST_WORD_X            As Byte = 2                         '  Intel-IA64-Filler
Public Const X3_BTYPE_QP_SIZE_X                 As Byte = 9                         '  Intel-IA64-Filler
Public Const X3_BTYPE_QP_INST_WORD_POS_X        As Byte = 23                        '  Intel-IA64-Filler
Public Const X3_BTYPE_QP_INST_VAL_POS_X         As Byte = 0                         '  Intel-IA64-Filler

Public Const X3_EMPTY_INST_WORD_X               As Byte = 1                         '  Intel-IA64-Filler
Public Const X3_EMPTY_SIZE_X                    As Byte = 2                         '  Intel-IA64-Filler
Public Const X3_EMPTY_INST_WORD_POS_X           As Byte = 14                        '  Intel-IA64-Filler
Public Const X3_EMPTY_INST_VAL_POS_X            As Byte = 0                         '  Intel-IA64-Filler

' Line number format.
Public Type IMAGE_LINENUMBER
    'Type
    SymbolTableIndex    As Long                 'Type_SymbolTableIndex Symbol table index of function name if Linenumber is 0.
'    VirtualAddress      As Long                 'Type_VirtualAddress Virtual address of line number.
    Linenumber          As Integer              'Line number.
End Type

Public Type IMAGE_BASE_RELOCATION
    VirtualAddress      As Long
    SizeOfBlock         As Long
    'WORD    TypeOffset[1];
End Type

' Based relocation types.
Public Const IMAGE_REL_BASED_ABSOLUTE           As Byte = 0
Public Const IMAGE_REL_BASED_HIGH               As Byte = 1
Public Const IMAGE_REL_BASED_LOW                As Byte = 2
Public Const IMAGE_REL_BASED_HIGHLOW            As Byte = 3
Public Const IMAGE_REL_BASED_HIGHADJ            As Byte = 4
Public Const IMAGE_REL_BASED_MACHINE_SPECIFIC_5 As Byte = 5
Public Const IMAGE_REL_BASED_RESERVED           As Byte = 6
Public Const IMAGE_REL_BASED_MACHINE_SPECIFIC_7 As Byte = 7
Public Const IMAGE_REL_BASED_MACHINE_SPECIFIC_8 As Byte = 8
Public Const IMAGE_REL_BASED_MACHINE_SPECIFIC_9 As Byte = 9
Public Const IMAGE_REL_BASED_DIR64              As Byte = 10

' Platform-specific based relocation types.
Public Const IMAGE_REL_BASED_IA64_IMM64         As Byte = 9

Public Const IMAGE_REL_BASED_MIPS_JMPADDR       As Byte = 5
Public Const IMAGE_REL_BASED_MIPS_JMPADDR16     As Byte = 9

Public Const IMAGE_REL_BASED_ARM_MOV32          As Byte = 5
Public Const IMAGE_REL_BASED_THUMB_MOV32        As Byte = 7


' Archive format.
Public Const IMAGE_ARCHIVE_START_SIZE           As Byte = 8
Public Const IMAGE_ARCHIVE_START                As String = "!<arch>\n"
Public Const IMAGE_ARCHIVE_END                  As String = "`\n"
Public Const IMAGE_ARCHIVE_PAD                  As String = "\n"
Public Const IMAGE_ARCHIVE_LINKER_MEMBER        As String = "/               "
Public Const IMAGE_ARCHIVE_LONGNAMES_MEMBER     As String = "//              "

Public Type IMAGE_ARCHIVE_MEMBER_HEADER
    vName(15)           As Byte                 'File member name - `/' terminated.
    Date(11)            As Byte                 'File member date - decimal.
    UserID(5)           As Byte                 'File member user id - decimal.
    GroupID(5)          As Byte                 'File member group id - decimal.
    Mode(7)             As Byte                 'File member mode - octal.
    Size(9)             As Byte                 'File member size - decimal.
    EndHeader(1)        As Byte                 'String to end header.
End Type

Public Const IMAGE_SIZEOF_ARCHIVE_MEMBER_HDR    As Integer = 60
'DLL support.

'Export Format
'导出目录表:一个只有一行的表（与调试目录不同）。它给出了其它各种导出表的位置和大小
Public Type IMAGE_EXPORT_DIRECTORY
    Characteristics     As Long                 '保留，必须为0。
    TimeDateStamp       As Long                 '导出数据被创建的日期和时间。
    MajorVersion        As Integer              '主版本号。用户可以自行设置主版本号和次版本号。
    MinorVersion        As Integer              '次版本号。
    Name                As Long                 '包含这个DLL 名称的ASCII 码字符串相对于映像基址的偏移地址
    Base                As Long                 '映像中导出符号的起始序数值。这个域指定了导出地址表的起始序数值.它通常被设置为1
    NumberOfFunctions   As Long                 '导出地址表中元素的数目。
    NumberOfNames       As Long                 '导出名称指针表中元素的数目。它同时也是导出序数表中元素的数目
    AddressOfFunctions  As Long                 'RVA from base of image,导出地址表相对于映像基址的偏移地址。
    AddressOfNames      As Long                 ' RVA from base of image,导出名称指针表相对于映像基址的偏移地址。它的大小由Number of Name Pointers 域给出。
    AddressOfNameOrdinals As Long               ' RVA from base of image,导出序数表相对于映像基址的偏移地址。
End Type

'Import Format
Public Type IMAGE_IMPORT_BY_NAME
    Hint                As Integer
    Name                As Byte
End Type

Public Type IMAGE_THUNK_DATA64
    'u1
    ForwarderString     As Big_Iint             'u1_ForwarderString  PBYTE
'    Function            As Big_Iint             'u1_Function  PDWORD
'    Ordinal             As Big_Iint             'u1_Ordinal
'    AddressOfData       As Big_Iint             'u1_AddressOfData  PIMAGE_IMPORT_BY_NAME
End Type


Public Type IMAGE_THUNK_DATA32
    'u1
    ForwarderString     As Long                 'u1_ForwarderString 址
'    Function            As Long                 'u1_Function  PDWORD
'    Ordinal             As Long                 'u1_Ordinal
'    AddressOfData       As Long                 'u1_AddressOfData  PIMAGE_IMPORT_BY_NAME
End Type

Public Type IMAGE_TLS_DIRECTORY64
    StartAddressOfRawData   As Big_Iint
    EndAddressOfRawData As Big_Iint
    AddressOfIndex      As Big_Iint             'PDWORD
    AddressOfCallBacks  As Big_Iint             'PIMAGE_TLS_CALLBACK
    SizeOfZeroFill      As Long
    Characteristics     As Long
End Type

Public Type IMAGE_TLS_DIRECTORY32
    StartAddressOfRawData   As Long
    EndAddressOfRawData As Long
    AddressOfIndex      As Long                 'PDWORD
    AddressOfCallBacks  As Long                 'PIMAGE_TLS_CALLBACK
    SizeOfZeroFill      As Long
    Characteristics     As Long
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


'New format import descriptors pointed to by DataDirectory[ IMAGE_DIRECTORY_ENTRY_BOUND_IMPORT ]
Public Type IMAGE_BOUND_IMPORT_DESCRIPTOR
    TimeDateStamp       As Long
    OffsetModuleName    As Integer
    NumberOfModuleForwarderRefs As Integer
    'Array of zero or more IMAGE_BOUND_FORWARDER_REF follows
End Type

Public Type IMAGE_BOUND_FORWARDER_REF
    TimeDateStamp       As Long
    OffsetModuleName    As Integer
    Reserved            As Integer
End Type

Public Type IMAGE_DELAYLOAD_DESCRIPTOR
    'Attributes
    AllAttributes       As Long                 'Attributes_AllAttributes
'    RvaBased            As Long                 'Attributes_RvaBased.Bits1  Delay load version 2
'    RvaBased            As Long                 'Attributes_RvaBased.Bits31
    DllNameRVA          As Long                 'RVA to the name of the target library (NULL-terminate ASCII string)
    ModuleHandleRVA     As Long                 'RVA to the HMODULE caching location (PHMODULE)
    ImportAddressTableRVA   As Long             'RVA to the start of the IAT (PIMAGE_THUNK_DATA)
    ImportNameTableRVA  As Long                 'RVA to the start of the name table (PIMAGE_THUNK_DATA::AddressOfData)
    BoundImportAddressTableRVA  As Long         'RVA to an optional bound IAT
    UnloadInformationTableRVA   As Long         'RVA to an optional unload info table
    TimeDateStamp          As Long              '0 if not bound,
                                                'Otherwise, date/time of the target DLL
End Type

' Resource Format.
' Resource directory consists of two counts, following by a variable length
' array of directory entries.  The first count is the number of entries at
' beginning of the array that have actual names associated with each entry.
' The entries are in ascending order, case insensitive strings.  The second
' count is the number of entries that immediately follow the named entries.
' This second count identifies the number of entries that have 16-bit integer
' Ids as their name.  These entries are also sorted in ascending order.
'
' This structure allows fast lookup by either name or number, but for any
' given resource entry only one form of lookup is supported, not both.
' This is consistant with the syntax of the .RC file and the .RES file.
Public Type IMAGE_RESOURCE_DIRECTORY
    Characteristics     As Long                 '理论上为资源的属性，不过事实上总是0
    TimeDateStamp       As Long                 '资源的产生时刻
    MajorVersion        As Integer              '理论上为资源的版本，不过事实上总是0
    MinorVersion        As Integer
    NumberOfNamedEntries    As Integer          '以名称命名的入口数量
    NumberOfIdEntries   As Integer              '以ID命名的入口数量
'    IMAGE_RESOURCE_DIRECTORY_ENTRY DirectoryEntries[];
End Type

Public Const IMAGE_RESOURCE_NAME_IS_STRING      As Long = &H80000000
Public Const IMAGE_RESOURCE_DATA_IS_DIRECTORY   As Long = &H80000000
'
' Each directory contains the 32-bit Name of the entry and an offset,
' relative to the beginning of the resource directory of the data associated
' with this directory entry.  If the name of the entry is an actual text
' string instead of an integer Id, then the high order bit of the name field
' is set to one and the low order 31-bits are an offset, relative to the
' beginning of the resource directory of the string, which is of type
' IMAGE_RESOURCE_DIRECTORY_STRING.  Otherwise the high bit is clear and the
' low-order 16-bits are the integer Id that identify this resource directory
' entry. If the directory entry is yet another resource directory (i.e. a
' subdirectory), then the high order bit of the offset field will be
' set to indicate this.  Otherwise the high bit is clear and the offset
' field points to a resource data entry.
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

' For resource directory entries that have actual string names, the Name
' field of the directory entry points to an object of the following type.
' All of these string objects are stored together after the last resource
' directory entry and before the first resource data object.  This minimizes
' the impact of these variable length objects on the alignment of the fixed
' size directory entry objects.

Public Type IMAGE_RESOURCE_DIRECTORY_STRING
    Length              As Integer
    NameString          As Byte
End Type


Public Type IMAGE_RESOURCE_DIR_STRING_U
    Length              As Integer              '字符串的长度
    NameString          As Integer              'UNICODE字符串，由于字符串是不定长的，所以这里只能用一个dw表示，实际上当长度为100的时候，这里的数据是NameString dw 100 dup (?)
End Type

' Each resource data entry describes a leaf node in the resource directory
' tree.  It contains an offset, relative to the beginning of the resource
' directory of the data for the resource, a size field that gives the number
' of bytes of data at that offset, a CodePage that should be used when
' decoding code point values within the resource data.  Typically for new
' applications the code page would be the unicode code page.

Public Type IMAGE_RESOURCE_DATA_ENTRY
    OffsetToData        As Long
    Size                As Long
    CodePage            As Long
    Reserved            As Long
End Type

'   Load Configuration Directory Entry
Public Type IMAGE_LOAD_CONFIG_DIRECTORY32
    Size                As Long
    TimeDateStamp       As Long
    MajorVersion        As Integer
    MinorVersion        As Integer
    GlobalFlagsClear    As Long
    GlobalFlagsSet      As Long
    CriticalSectionDefaultTimeout   As Long
    DeCommitFreeBlockThreshold  As Long
    DeCommitTotalFreeThreshold  As Long
    LockPrefixTable     As Long                 'VA
    MaximumAllocationSize   As Long
    VirtualMemoryThreshold  As Long
    ProcessHeapFlags    As Long
    ProcessAffinityMask As Long
    CSDVersion          As Integer
    Reserved1           As Integer
    EditList            As Long                 'VA
    SecurityCookie      As Long                 'VA
    SEHandlerTable      As Long                 'VA
End Type

Public Type IMAGE_LOAD_CONFIG_DIRECTORY64
    Size                As Long
    TimeDateStamp       As Long
    MajorVersion        As Integer
    MinorVersion        As Integer
    GlobalFlagsClear    As Long
    GlobalFlagsSet      As Long
    CriticalSectionDefaultTimeout   As Long
    DeCommitFreeBlockThreshold  As Big_Iint
    DeCommitTotalFreeThreshold  As Big_Iint
    LockPrefixTable     As Big_Iint             ' VA
    MaximumAllocationSize   As Big_Iint
    VirtualMemoryThreshold  As Big_Iint
    ProcessAffinityMask As Big_Iint
    ProcessHeapFlags    As Long
    CSDVersion          As Integer
    Reserved1           As Integer
    EditList            As Big_Iint             'VA
    SecurityCookie      As Big_Iint             ' VA
    SEHandlerTable      As Big_Iint             ' VA
    SEHandlerCount      As Big_Iint
End Type

Public Type IMAGE_CE_RUNTIME_FUNCTION_ENTRY
    FuncStart           As Long
    'DWORD
    PrologLen           As Long                 'DWORD_Bits8
'    FuncLen             As Long                 'DWORD_Bits22
'    ThirtyTwoBit        As Long                 'DWORD_Bits1
'    ExceptionFlag       As Long                 'DWORD_Bits1
End Type

Public Type IMAGE_ARM_RUNTIME_FUNCTION_ENTRY
    BeginAddress        As Long
    'DUMMYUNIONNAME
    UnwindData          As Long                 'DUMMYUNIONNAME_UnwindData
    'DUMMYUNIONNAME_DUMMYSTRUCTNAME
'    Flag                As Long                 'DUMMYUNIONNAME_DUMMYSTRUCTNAME_Flag_DWORD_Bits2
'    FunctionLength      As Long                 'DUMMYUNIONNAME_DUMMYSTRUCTNAME_FunctionLength_DWORD_Bits11
'    Ret                 As Long                 'DUMMYUNIONNAME_DUMMYSTRUCTNAME_Ret_DWORD_Bits2
'    H                   As Long                 'DUMMYUNIONNAME_DUMMYSTRUCTNAME_H_DWORD_Bits1
'    Reg                 As Long                 'DUMMYUNIONNAME_DUMMYSTRUCTNAME_Reg_DWORD_Bits3
'    R                   As Long                 'DUMMYUNIONNAME_DUMMYSTRUCTNAME_R_DWORD_Bits1
'    L                   As Long                 'DUMMYUNIONNAME_DUMMYSTRUCTNAME_L_DWORD_Bits1
'    C                   As Long                 'DUMMYUNIONNAME_DUMMYSTRUCTNAME_C_DWORD_Bits1
'    StackAdjust         As Long                 'DUMMYUNIONNAME_DUMMYSTRUCTNAME_StackAdjust_DWORD_Bits10
End Type

Public Type IMAGE_ALPHA64_RUNTIME_FUNCTION_ENTRY
    BeginAddress        As Big_Iint
    EndAddress          As Big_Iint
    ExceptionHandler    As Big_Iint
    HandlerData         As Big_Iint
    PrologEndAddress    As Big_Iint
End Type

Public Type IMAGE_ALPHA_RUNTIME_FUNCTION_ENTRY
    BeginAddress        As Long
    EndAddress          As Long
    ExceptionHandler    As Long
    HandlerData         As Long
    PrologEndAddress    As Long
End Type

Public Type IMAGE_RUNTIME_FUNCTION_ENTRY
    BeginAddress        As Long
    EndAddress          As Long
    'DUMMYUNIONNAME
    UnwindInfoAddress   As Long                 'DUMMYUNIONNAME_UnwindInfoAddress
'    UnwindData          As Long                 'DUMMYUNIONNAME_UnwindData
End Type

Public Type IMAGE_DEBUG_DIRECTORY
    Characteristics     As Long                 '保留，必须为0。
    TimeDateStamp       As Long                 '调试数据被创建的日期和时间。
    MajorVersion        As Integer              '调试数据格式的主版本号。
    MinorVersion        As Integer              '调试数据格式的次版本号。
    Type                As Long                 '调试信息的格式。这个域的存在使得可以支持多个调试器
    SizeOfData          As Long                 '调试数据（不包括调试目录本身）的大小。
    AddressOfRawData    As Long                 '当被加载时调试数据相对于映像基址的偏移地址。
    PointerToRawData    As Long                 '指向调试数据的文件指针。
End Type

Public Const IMAGE_DEBUG_TYPE_UNKNOWN           As Long = 0                         '未知值，所有工具均忽略此值。
Public Const IMAGE_DEBUG_TYPE_COFF              As Long = 1                         'COFF 调试信息（行号信息、符号表和字符串表）。文件头中也有相关域指向这种类型的调试信息
Public Const IMAGE_DEBUG_TYPE_CODEVIEW          As Long = 2                         'Visual C++调试信息。
Public Const IMAGE_DEBUG_TYPE_FPO               As Long = 3                         '帧指针省略（FPO）信息。这种信息告诉调试器如何解释非标准栈帧，这种帧将EBP 寄存器用于其它目的而不是作为帧指针
Public Const IMAGE_DEBUG_TYPE_MISC              As Long = 4                         'DBG 文件的位置。
Public Const IMAGE_DEBUG_TYPE_EXCEPTION         As Long = 5                         '.pdata 节的副本。
Public Const IMAGE_DEBUG_TYPE_FIXUP             As Long = 6                         '保留。
Public Const IMAGE_DEBUG_TYPE_OMAP_TO_SRC       As Long = 7                         '从经过代码重排后的映像中的RVA 到原映像中的RVA 的映射
Public Const IMAGE_DEBUG_TYPE_OMAP_FROM_SRC     As Long = 8                         '从原映像中的RVA 到经过代码重排后的映像中的RVA 的映射
Public Const IMAGE_DEBUG_TYPE_BORLAND           As Long = 9                         '保留，供Borland 公司使用。
Public Const IMAGE_DEBUG_TYPE_RESERVED10        As Long = 10                        '保留
Public Const IMAGE_DEBUG_TYPE_CLSID             As Long = 11                        '保留

Public Type IMAGE_COFF_SYMBOLS_HEADER
    NumberOfSymbols     As Long
    LvaToFirstSymbol    As Long
    NumberOfLinenumbers As Long
    LvaToFirstLinenumber    As Long
    RvaToFirstByteOfCode    As Long
    RvaToLastByteOfCode As Long
    RvaToFirstByteOfData    As Long
    RvaToLastByteOfData As Long
End Type

Public Const FRAME_FPO                          As Long = 0
Public Const FRAME_TRAP                         As Long = 1
Public Const FRAME_TSS                          As Long = 2
Public Const FRAME_NONFPO                       As Long = 3

Public Type FPO_DATA
    ulOffStart          As Long                 'offset 1st byte of function code  函数代码第一个字节的偏移
    cbProcSize          As Long                 '# bytes in function  函数代码所占的字节数
    cdwLocals           As Long                 '# bytes in locals/4  局部变量所占字节数除以4
    cdwParams           As Integer              '# bytes in params/4  参数所占字节数除以4
    cbProlog            As Integer              'WORD_bits8# bytes in prolog  函数prolog 代码所占字节数
'    cbRegs              As Integer              'WORD_bits3# regs saved    保存的寄存器数
'    fHasSEH             As Integer              'WORD_bits1 TRUE if SEH in func  如果函数中有SEH，此值为TRUE
'    fUseBP              As Integer              'WORD_bits1 TRUE if EBP has been allocated  如果EBP 寄存器已经被分配，此值为TRUE
'    reserved            As Integer              'WORD_bits1 reserved for future use  保留供将来使用
'    cbFrame             As Integer              'WORD_bits2 frame type帧类型
End Type


Public Const IMAGE_DEBUG_MISC_EXENAME           As Long = 1

Public Type IMAGE_DEBUG_MISC
    DataType            As Long                 'type of misc data, see defines
    Length              As Long                 'total length of record, rounded to four
                                                'byte multiple.
    Unicode             As Long                 'TRUE if data is unicode string
    Reserved(2)         As Byte
    Data                As Byte                  'Actual data
End Type

'Function table extracted from MIPS/ALPHA/IA64 images.  Does not contain
'information needed only for runtime support.  Just those fields for
'each entry needed by a debugger.

Public Type IMAGE_FUNCTION_ENTRY
    StartingAddress     As Long
    EndingAddress       As Long
    EndOfPrologue       As Long
End Type

Public Type IMAGE_FUNCTION_ENTRY64
    StartingAddress     As Big_Iint
    EndingAddress       As Big_Iint
    'DUMMYUNIONNAME
    EndOfPrologue       As Big_Iint             'DUMMYUNIONNAME_EndOfPrologue
'    UnwindInfoAddress   As Big_Iint             'DUMMYUNIONNAME_UnwindInfoAddress
End Type
'Debugging information can be stripped from an image file and placed
'in a separate .DBG file, whose file name part is the same as the
'image file name part (e.g. symbols for CMD.EXE could be stripped
'and placed in CMD.DBG).  This is indicated by the IMAGE_FILE_DEBUG_STRIPPED
'flag in the Characteristics field of the file header.  The beginning of
'the .DBG file contains the following structure which captures certain
'information from the image file.  This allows a debug to proceed even if
'the original image file is not accessable.  This header is followed by
'zero of more IMAGE_SECTION_HEADER structures, followed by zero or more
'IMAGE_DEBUG_DIRECTORY structures.  The latter structures and those in
'the image file contain file offsets relative to the beginning of the
'.DBG file.

'If symbols have been stripped from an image, the IMAGE_DEBUG_MISC structure
'is left in the image file, but not mapped.  This allows a debugger to
'compute the name of the .DBG file, from the name of the image in the
'IMAGE_DEBUG_MISC structure.

Public Type IMAGE_SEPARATE_DEBUG_HEADER
    Signature           As Integer
    Flags               As Integer
    Machine             As Integer
    Characteristics     As Integer
    TimeDateStamp       As Long
    Checksum            As Long
    ImageBase           As Long
    SizeOfImage         As Long
    NumberOfSections    As Long
    ExportedNamesSize   As Long
    DebugDirectorySize  As Long
    SectionAlignment    As Long
    Reserved(1)         As Long
End Type

Public Type NON_PAGED_DEBUG_INFO
    Signature           As Integer
    Flags               As Integer
    Size                As Long
    Machine             As Integer
    Characteristics     As Integer
    TimeDateStamp       As Long
    Checksum            As Long
    SizeOfImage         As Long
    ImageBase           As Big_Iint
    'DebugDirectorySize
    'IMAGE_DEBUG_DIRECTORY
End Type

Public Const IMAGE_SEPARATE_DEBUG_SIGNATURE     As Integer = &H4449                 'DI
Public Const NON_PAGED_DEBUG_SIGNATURE          As Integer = &H4E49                 'NI

Public Const IMAGE_SEPARATE_DEBUG_FLAGS_MASK    As Integer = &H8000
Public Const IMAGE_SEPARATE_DEBUG_MISMATCH      As Integer = &H8000                 'when DBG was updated, the
                                                                                    'old checksum didn't match.

' The .arch section is made up of headers, each describing an amask position/value
' pointing to an array of IMAGE_ARCHITECTURE_ENTRY's.  Each "array" (both the header
' and entry arrays) are terminiated by a quadword of 0xffffffffL.
' NOTE: There may be quadwords of 0 sprinkled around and must be skipped.

Public Type ImageArchitectureHeader
    AmaskValue          As Long                 'int_bits1,1 -> code section depends on mask bit
                                                '0 -> new instruction depends on mask bit
'    int                 As Long                 'int_bits7,MBZ
'    AmaskShift          As Long                 'int_bits8,Amask bit in question for this fixup
'    int                 As Long                 'int_bits16,MBZ
    FirstEntryRVA       As Long                 'RVA into .arch section to array of ARCHITECTURE_ENTRY's
End Type

Public Type ImageArchitectureEntry
    FixupInstRVA        As Long                 'RVA of instruction to fixup
    NewInst             As Long                 'fixup instruction (see alphaops.h)
End Type

'The following structure defines the new import object.  Note the values of the first two fields,
'which must be set as stated in order to differentiate old and new import members.
'Following this structure, the linker emits two null-terminated strings used to recreate the
'import at the time of use.  The first string is the import's name, the second is the dll's name.

Public Const IMPORT_OBJECT_HDR_SIG2             As Integer = &HFFFF

Public Type IMPORT_OBJECT_HEADER
    Sig1                As Integer              'Must be IMAGE_FILE_MACHINE_UNKNOWN
    Sig2                As Integer              'Must be IMPORT_OBJECT_HDR_SIG2.
    Version             As Integer
    Machine             As Integer
    TimeDateStamp       As Long                 'Time/date stamp
    SizeOfData          As Long                 'particularly useful for incremental links
    'DUMMYUNIONNAME
    Ordinal             As Integer              'DUMMYUNIONNAME_Ordinal,if grf & IMPORT_OBJECT_ORDINAL
'    Hint                As Integer             'DUMMYUNIONNAME_DUMMYUNIONNAME
    Type                As Integer              'WORD_Bit2,IMPORT_TYPE
    NameType            As Integer              'WORD_Bit3,IMPORT_NAME_TYPE
    Reserved            As Integer              'WORD_Bit11,Reserved. Must be zero.
End Type

Public Enum IMPORT_OBJECT_TYPE
    IMPORT_OBJECT_CODE = 0
    IMPORT_OBJECT_DATA = 1
    IMPORT_OBJECT_CONST = 2
End Enum

Public Enum IMPORT_OBJECT_NAME_TYPE
    IMPORT_OBJECT_ORDINAL = 0                   'Import by ordinal
    IMPORT_OBJECT_NAME = 1                      'Import name == public symbol name.
    IMPORT_OBJECT_NAME_NO_PREFIX = 2            'Import name == public symbol name skipping leading ?, @, or optionally _.
    IMPORT_OBJECT_NAME_UNDECORATE = 3           'Import name == public symbol name skipping leading ?, @, or optionally
                                                'and truncating at first @
End Enum

Public Enum ReplacesCorHdrNumericDefines
'COM+ Header entry point flags.32bits
    COMIMAGE_FLAGS_ILONLY = &H1
    COMIMAGE_FLAGS_32BITREQUIRED = &H2
    COMIMAGE_FLAGS_IL_LIBRARY = &H4
    COMIMAGE_FLAGS_STRONGNAMESIGNED = &H8
    COMIMAGE_FLAGS_NATIVE_ENTRYPOINT = &H10
    COMIMAGE_FLAGS_TRACKDEBUGDATA = &H10000

'Version flags for image.
    COR_VERSION_MAJOR_V2 = 2
    COR_VERSION_MAJOR = COR_VERSION_MAJOR_V2
    COR_VERSION_MINOR = 5
    COR_DELETED_NAME_LENGTH = 8
    COR_VTABLEGAP_NAME_LENGTH = 8

'Maximum size of a NativeType descriptor.8bits
    NATIVE_TYPE_MAX_CB = 1
    COR_ILMETHOD_SECT_SMALL_MAX_DATASIZE = &HFF

'#defines for the MIH FLAGS 16bits
    IMAGE_COR_MIH_METHODRVA = &H1
    IMAGE_COR_MIH_EHRVA = &H2
    IMAGE_COR_MIH_BASICBLOCK = &H8

'V-table constants 16bits
    COR_VTABLE_32BIT = &H1                              'V-table slots are 32-bits in size.
    COR_VTABLE_64BIT = &H2                              'V-table slots are 64-bits in size.
    COR_VTABLE_FROM_UNMANAGED = &H4                     'If set, transition from unmanaged.
    COR_VTABLE_FROM_UNMANAGED_RETAIN_APPDOMAIN = &H8    'If set, transition from unmanaged with keeping the current appdomain.
    COR_VTABLE_CALL_MOST_DERIVED = &H10                 'Call most derived method described by

'EATJ constants
    IMAGE_COR_EATJ_THUNK_SIZE = 32                    'Size of a jump thunk reserved range.

'Max name lengths
    '@todo: Change to unlimited name lengths.
    MAX_CLASS_NAME = 1024
    MAX_PACKAGE_NAME = 1024
End Enum

'CLR 2.0 header structure.
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

