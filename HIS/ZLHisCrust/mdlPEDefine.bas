Attribute VB_Name = "mdlPEDefine"
Option Explicit

' ----------------------------------------------------
'   MS-DOS 2.0 ����EXE �ļ�ͷ       |               |ӳ��ͷ����ַ
'-----------------------------------|               |
'       δʹ��                      |               |
'-----------------------------------|               |
'   OEM ��ʶ                        |               |
'   OEM ��Ϣ                        |               |
'   PE �ļ�ͷƫ��                   |               |MS-DOS 2.0 �ڣ�������MS-DOS ���ݣ�
'-----------------------------------|               |
'MS-DOS 2.0 ռλ���� �� �ض�λ��    |               |
'-----------------------------------|               |
'   δʹ��                          |               |
'---------------------------------------------------
'PE �ļ�ͷ����8 �ֽڱ߽���룩      |
'-----------------------------------
'   ��ͷ                            |
'-----------------------------------
'   ӳ��ҳ:                         |
'   ������Ϣ                        |
'   ������Ϣ                        |
'   ��ַ�ض�λ��Ϣ                  |
'   ��Դ��Ϣ                        |
'-----------------------------------

'    +-------------------+
'    | DOS-stub          |    --DOS-ͷ
'    +-------------------+
'    | file-header       |    --�ļ�ͷ
'    +-------------------+
'    | optional header   |    --��ѡͷ
'    |- - - - - - - - - -|
'    |                   |
'    | data directories  |    --����Ŀ¼
'    |                   |
'    +-------------------+
'    |                   |
'    | section headers   |     --��ͷ
'    |                   |
'    +-------------------+
'    |                   |
'    | section 1         |     --��1
'    |                   |
'    +-------------------+
'    |                   |
'    | section 2         |     --��2
'    |                   |
'    +-------------------+
'    |                   |
'    | ...               |
'    |                   |
'    +-------------------+
'    |                   |
'    | section n         |     --��n
'    |                   |
'    +-------------------+
'����       ����                                                        ����
'.bss       δ��ʼ�������ݣ����ɸ�ʽ��                                  IMAGE_SCN_CNT_UNINITIALIZED_DATA |IMAGE_SCN_MEM_READ |IMAGE_SCN_MEM_WRITE
'.cormeta   CLR Ԫ���ݣ�������Ŀ���ļ��а����йܴ���                    IMAGE_SCN_LNK_INFO
'.data      �ѳ�ʼ�������ݣ����ɸ�ʽ��                                  IMAGE_SCN_CNT_INITIALIZED_DATA |IMAGE_SCN_MEM_READ |IMAGE_SCN_MEM_WRITE
'.debug$F   ���ɵ�FPO ������Ϣ����������Ŀ���ļ���������x86 ƽ̨�����ѱ�������  IMAGE_SCN_CNT_INITIALIZED_DATA |IMAGE_SCN_MEM_READ |IMAGE_SCN_MEM_DISCARDABLE
'.debug$P   Ԥ����ĵ���������Ϣ����������Ŀ���ļ���                    IMAGE_SCN_CNT_INITIALIZED_DATA |IMAGE_SCN_MEM_READ |IMAGE_SCN_MEM_DISCARDABLE
'.debug$S   ���Է�����Ϣ����������Ŀ���ļ���                            IMAGE_SCN_CNT_INITIALIZED_DATA |IMAGE_SCN_MEM_READ |IMAGE_SCN_MEM_DISCARDABLE
'.debug$T   ����������Ϣ����������Ŀ���ļ���                            IMAGE_SCN_CNT_INITIALIZED_DATA |IMAGE_SCN_MEM_READ |IMAGE_SCN_MEM_DISCARDABLE
'.drectve   ������ѡ��                                                  IMAGE_SCN_LNK_INFO
'.edata     ������                                                      IMAGE_SCN_CNT_INITIALIZED_DATA |IMAGE_SCN_MEM_READ
'.idata     �����                                                      IMAGE_SCN_CNT_INITIALIZED_DATA |IMAGE_SCN_MEM_READ |IMAGE_SCN_MEM_WRITE
'.idlsym    ������ע���SEH����������ӳ���ļ�������������֧��IDL ����   IMAGE_SCN_LNK_INFO
'.pdata     �쳣��Ϣ                                                    IMAGE_SCN_CNT_INITIALIZED_DATA |IMAGE_SCN_MEM_READ
'.rdata     ֻ�����ѳ�ʼ������                                          IMAGE_SCN_CNT_INITIALIZED_DATA |IMAGE_SCN_MEM_READ
'.reloc     ӳ���ļ����ض�λ��Ϣ                                        IMAGE_SCN_CNT_INITIALIZED_DATA |IMAGE_SCN_MEM_READ |IMAGE_SCN_MEM_DISCARDABLE
'.rsrc      ��ԴĿ¼                                                    IMAGE_SCN_CNT_INITIALIZED_DATA |IMAGE_SCN_MEM_READ
'.sbss      ��GP ��ص�δ��ʼ�����ݣ����ɸ�ʽ��                         IMAGE_SCN_CNT_UNINITIALIZED_DATA |IMAGE_SCN_MEM_READ |IMAGE_SCN_MEM_WRITE |IMAGE _SCN_GPREL ����IMAGE_SCN_GPREL ��־������IA64 ƽ̨��������������ƽ̨���˱�־ֻ������Ŀ���ļ�����ӳ���ļ��г����������͵Ľ�ʱ��һ���������������־
'.sdata     ��GP ��ص��ѳ�ʼ�����ݣ����ɸ�ʽ��                         IMAGE_SCN_CNT_INITIALIZED_DATA |IMAGE_SCN_MEM_READ |IMAGE_SCN_MEM_WRITE |IMAGE _SCN_GPREL����IMAGE_SCN_GPREL ��־������IA64 ƽ̨��������������ƽ̨���˱�־ֻ������Ŀ���ļ�����ӳ���ļ��г����������͵Ľ�ʱ��һ���������������־
'.srdata    ��GP ��ص�ֻ�����ݣ����ɸ�ʽ��                             IMAGE_SCN_CNT_INITIALIZED_DATA |IMAGE_SCN_MEM_READ |IMAGE _SCN_GPREL����IMAGE_SCN_GPREL ��־������IA64 ƽ̨��������������ƽ̨���˱�־ֻ������Ŀ���ļ�����ӳ���ļ��г����������͵Ľ�ʱ��һ���������������־
'.sxdata    ��ע����쳣����������ݣ����ɸ�ʽ����������Ŀ���ļ���������x86 ƽ̨��  MAGE_SCN_LNK_INFO������а���Ŀ���ļ��еĴ������漰���������쳣��������ڷ��ű��е�����.��Щ���ſ�����IMAGE_SYM_UNDEFINED ���͵ķ��ţ�Ҳ�����Ƕ������Ǹ�ģ���еķ���
'.text      ��ִ�д��루���ɸ�ʽ��                                      IMAGE_SCN_CNT_CODE |IMAGE_SCN_MEM_EXECUTE |IIMAGE_SCN_MEM_READ
'.tls       �ֲ߳̾��洢����������Ŀ���ļ���                            IMAGE_SCN_CNT_INITIALIZED_DATA |IMAGE_SCN_MEM_READ |IMAGE_SCN_MEM_WRITE
'.tls$      �ֲ߳̾��洢����������Ŀ���ļ���                            IMAGE_SCN_CNT_INITIALIZED_DATA |IMAGE_SCN_MEM_READ |IMAGE_SCN_MEM_WRITE
'.vsdata    ��GP ��ص��ѳ�ʼ�����ݣ����ɸ�ʽ����������ARM��SH4 ��Thumb ƽ̨��  IMAGE_SCN_CNT_INITIALIZED_DATA |IMAGE_SCN_MEM_READ |IMAGE_SCN_MEM_WRITE
'.xdata     �쳣��Ϣ�����ɸ�ʽ��                                        IMAGE_SCN_CNT_INITIALIZED_DATA |IMAGE_SCN_MEM_READ
'*******************************************************************************************
'   MS-DOS 2.0 ����EXE �ļ�ͷ
'*******************************************************************************************
'PE_SIGNATURE(e_magic,ne_magic,e32_magic)
Public Const IMAGE_DOS_SIGNATURE                As Integer = &H5A4D                 'MZ
Public Const IMAGE_OS2_SIGNATURE                As Integer = &H454E                 'NE
Public Const IMAGE_OS2_SIGNATURE_LE             As Integer = &H454C                 'LE
Public Const IMAGE_NT_SIGNATURE                 As Long = &H4550                    'PE00

Public Type IMAGE_DOS_HEADER        'DOS .EXE header 64B
    e_magic             As Integer              'Magic number  �ֱ�ΪMZ,4Dh��5Ah
    e_cblp              As Integer              'Bytes on last page of file �ļ����һҳ�ֽ���
    e_cp                As Integer              'Pages in file          �ļ���ҳ��(512B/ҳ)
    e_crlc              As Integer              'Relocations            �ض�λ������
    e_cparhdr           As Integer              'Size of header in paragraphs   �ļ�ͷ�ܶ���(16B/��)
    e_minalloc          As Integer              'Minimum extra paragraphs needed
    e_maxalloc          As Integer              'Maximum extra paragraphs needed
    e_ss                As Integer              'Initial (relative) SS value SS:SP
    e_sp                As Integer              'Initial SP value   SS:SP
    e_csum              As Integer              'Checksum           У���
    e_ip                As Integer              'Initial IP value   CS:IP
    e_cs                As Integer              'Initial (relative) CS value    CS:IP
    e_lfarlc            As Integer              'File address of relocation table   �ض�λ��ƫ�Ƶ�ַ
    e_ovno              As Integer              'Overlay number
    e_res(3)            As Integer              'Reserved words
    e_oemid             As Integer              'OEM identifier (for e_oeminfo)
    e_oeminfo           As Integer              'OEM information; e_oemid specific
    e_res2(9)           As Integer              'Reserved words
    e_lfanew            As Long                 'File address of new exe header PEͷƫ��,ָ��PE�ļ�ͷ
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
'   �ļ�ͷ��3���֣�
'*******************************************************************************************
'-------NTͷ-------------
Public Const IMAGE_SIZEOF_FILE_HEADER       As Integer = 20
Public Type IMAGE_FILE_HEADER                   '20B
    Machine             As Integer              '��ʶĿ��������͵�����
    NumberOfSections    As Integer              '�ڵ���Ŀ���������˽ڱ�Ĵ�С�����ڱ�������ļ�ͷ
    TimeDateStamp       As Long                 '��UTC ʱ��1970 ��1 ��1 ��00:00 �����������һ��C ����ʱtime_t ���͵�ֵ���ĵ�32 λ����ָ���ļ���ʱ������
    PointerToSymbolTable    As Long             'COFF ���ű���ļ�ƫ�ơ����������COFF ���ű���ֵΪ0������ӳ���ļ���˵����ֵӦ��Ϊ0����Ϊ�Ѿ����޳�ʹ��COFF ������Ϣ��
    NumberOfSymbols     As Long                 '���ű��е�Ԫ����Ŀ�������ַ�����������ű����Կ����������ֵ����λ�ַ�����?����ӳ���ļ���˵����ֵӦ��Ϊ0����Ϊ�Ѿ����޳�ʹ��COFF������Ϣ��
    SizeOfOptionalHeader    As Integer          '��ѡ�ļ�ͷ�Ĵ�С����ִ���ļ���Ҫ��ѡ�ļ�ͷ��Ŀ���ļ�������Ҫ������Ŀ���ļ���˵����ֵӦ��Ϊ0
    Characteristics     As Integer              'ָʾ�ļ����Եı�־��
End Type

'IMAGE_FILE_MACHINE
Public Const IMAGE_FILE_MACHINE_UNKNOWN         As Integer = &H0                    '�������κ����ʹ�����
Public Const IMAGE_FILE_MACHINE_I386            As Integer = &H14C                  'Intel 386.Intel 386 ���̴�����������ݴ�����
Public Const IMAGE_FILE_MACHINE_R3000           As Integer = &H162                  'MIPS little-endian,  big-endian
Public Const IMAGE_FILE_MACHINE_R4000           As Integer = &H166                  'MIPS little-endian     MIPS Сβ������
Public Const IMAGE_FILE_MACHINE_R10000          As Integer = &H168                  'MIPS little-endian
Public Const IMAGE_FILE_MACHINE_WCEMIPSV2       As Integer = &H169                  'MIPS little-endian WCE v2  MIPS СβWCE v2 ������
Public Const IMAGE_FILE_MACHINE_ALPHA           As Integer = &H184                  'Alpha_AXP
Public Const IMAGE_FILE_MACHINE_SH3             As Integer = &H1A2                  'SH3 little-endian  Hitachi SH3 ������
Public Const IMAGE_FILE_MACHINE_SH3DSP          As Integer = &H1A3                  'Hitachi SH3 DSP ������
Public Const IMAGE_FILE_MACHINE_SH3E            As Integer = &H1A4                  'SH3E little-endian
Public Const IMAGE_FILE_MACHINE_SH4             As Integer = &H1A6                  'SH4 little-endian      Hitachi SH4 ������
Public Const IMAGE_FILE_MACHINE_SH5             As Integer = &H1A8                  'SH5        Hitachi SH5 ������
Public Const IMAGE_FILE_MACHINE_ARM             As Integer = &H1C0                  'ARM Little-Endian  ARM Сβ������
Public Const IMAGE_FILE_MACHINE_THUMB           As Integer = &H1C2                  'ARM Thumb/Thumb-2 Little-Endian    Thumb ������
Public Const IMAGE_FILE_MACHINE_ARMNT           As Integer = &H1C4                  'ARM Thumb-2 Little-Endian
Public Const IMAGE_FILE_MACHINE_AM33            As Integer = &H1D3                  'Matsushita AM33 ������
Public Const IMAGE_FILE_MACHINE_POWERPC         As Integer = &H1F0                  'IBM PowerPC Little-Endian  PowerPC Сβ������
Public Const IMAGE_FILE_MACHINE_POWERPCFP       As Integer = &H1F1                  '����������֧�ֵ�PowerPC ������
Public Const IMAGE_FILE_MACHINE_IA64            As Integer = &H200                  'Intel 64 Intel Itanium ����������
Public Const IMAGE_FILE_MACHINE_MIPS16          As Integer = &H266                  'MIPS   MIPS16 ������
Public Const IMAGE_FILE_MACHINE_ALPHA64         As Integer = &H284                  'ALPHA64
Public Const IMAGE_FILE_MACHINE_MIPSFPU         As Integer = &H366                  'MIPS   ��FPU ��MIPS ������
Public Const IMAGE_FILE_MACHINE_MIPSFPU16       As Integer = &H466                  'MIPS   ��FPU ��MIPS16 ������
Public Const IMAGE_FILE_MACHINE_AXP64           As Integer = IMAGE_FILE_MACHINE_ALPHA64
Public Const IMAGE_FILE_MACHINE_TRICORE         As Integer = &H520                  'Infineon
Public Const IMAGE_FILE_MACHINE_CEF             As Integer = &HCEF
Public Const IMAGE_FILE_MACHINE_EBC             As Integer = &HEBC                  'EFI Byte Code  EFI �ֽ��봦����
Public Const IMAGE_FILE_MACHINE_AMD64           As Integer = &H8664                 'AMD64 (K8) x64 ������
Public Const IMAGE_FILE_MACHINE_M32R            As Integer = &H9041                 'M32R little-endian  Mitsubishi M32R Сβ������
Public Const IMAGE_FILE_MACHINE_CEE             As Integer = &HC0EE

'IMAGE_FILE_MACHINE  Characteristics
Public Const IMAGE_FILE_RELOCS_STRIPPED         As Long = &H1                       ' ��������ӳ���ļ��������� Windows CE��Microsoft Windows NT. �����̲���ϵͳ.���������ļ���������ַ�ض�λ��Ϣ����˱��뱻���ص�����ѡ����ַ�ϡ��������ַ�����ã��������ᱨ��������Ĭ�ϻ��Ƴ���ִ�У�EXE���ļ��е��ض�λ��Ϣ.
Public Const IMAGE_FILE_EXECUTABLE_IMAGE        As Long = &H2                       ' File is executable  (i.e. no unresolved externel references).��������ӳ���ļ�����������ӳ���ļ��ǺϷ��ģ����Ա����С����δ���ô˱�־����������������������
Public Const IMAGE_FILE_LINE_NUMS_STRIPPED      As Long = &H4                       ' Line nunbers stripped from file.�к���Ϣ�Ѿ����Ƴ������޳�ʹ�ô˱�־����Ӧ��Ϊ0��
Public Const IMAGE_FILE_LOCAL_SYMS_STRIPPED     As Long = &H8                       ' Local symbols stripped from file.COFF ���ű����йؾֲ����ŵ����Ѿ����Ƴ������޳�ʹ�ô˱�־����Ӧ��Ϊ0��
Public Const IMAGE_FILE_AGGRESIVE_WS_TRIM       As Long = &H10                      ' Agressively trim working set �˱�־�Ѿ��������������ڵ�����������Windows 2000 �����̲���ϵͳ���޳�ʹ�ô˱�־����Ӧ��Ϊ0��
Public Const IMAGE_FILE_LARGE_ADDRESS_AWARE     As Long = &H20                      ' App can handle >2gb addresses Ӧ�ó�����Դ������2GB �ĵ�ַ��
Public Const IMAGE_FILE_BYTES_REVERSED_LO       As Long = &H80                      ' Bytes of machine word are reversed. Сβ�����ڴ��У����λ��LSB�������λ��MSB��ǰ�档���޳�ʹ�ô˱�־����Ӧ��Ϊ0.
Public Const IMAGE_FILE_32BIT_MACHINE           As Long = &H100                     ' 32 bit word machine.          �������ͻ���32 λ����ϵ�ṹ��
Public Const IMAGE_FILE_DEBUG_STRIPPED          As Long = &H200                     ' Debugging info stripped from file in .DBG file ������Ϣ�Ѿ��Ӵ�ӳ���ļ����Ƴ���
Public Const IMAGE_FILE_REMOVABLE_RUN_FROM_SWAP As Long = &H400                     ' If Image is on removable media, copy and run from the swap file. �����ӳ���ļ��ڿ��ƶ������ϣ���ȫ���������������Ƶ������ļ���
Public Const IMAGE_FILE_NET_RUN_FROM_SWAP       As Long = &H800                     ' If Image is on Net, copy and run from the swap file.�����ӳ���ļ�����������ϣ���ȫ���������������Ƶ������ļ���
Public Const IMAGE_FILE_SYSTEM                  As Long = &H1000                    ' System File.  ��ӳ���ļ���ϵͳ�ļ����������û�����
Public Const IMAGE_FILE_DLL                     As Long = &H2000                    ' File is a DLL.��ӳ���ļ��Ƕ�̬���ӿ⣨DLL�����������ļ��ܱ���Ϊ�ǿ�ִ���ļ����������ǲ�����ֱ�ӱ�����
Public Const IMAGE_FILE_UP_SYSTEM_ONLY          As Long = &H4000                    ' File should only be run on a UP machine   ���ļ�ֻ�������ڵ������������ϡ�
Public Const IMAGE_FILE_BYTES_REVERSED_HI       As Long = &H8000                    ' Bytes of machine word are reversed. ��β�����ڴ��У�MSB ��LSB ǰ�档���޳�ʹ�ô˱�־����Ӧ��Ϊ0��

'-------��ѡͷ-------------
'Ŀ¼��ʽ
Public Type IMAGE_DATA_DIRECTORY
    VirtualAddress      As Long                 '���ݿ��RVA
    Size                As Long                 '���ݿ��С
End Type
'Ŀ¼��Ŀ
Public Const IMAGE_NUMBEROF_DIRECTORY_ENTRIES   As Integer = 16
'ƫ�ƣ�PE32/PE32+��  ��С��PE32/PE32+��   �ļ�ͷ����        ����
'0                  28/24                 ��׼��            ��Щ������COFF ʵ�������壬���а���UNIX
'28/24              68/88                 Windows �ض���    ֧��Windows ���ԣ�������ϵͳ���ĸ�����
'96/112             Variable              ����Ŀ¼          ӳ���ļ��е���������絼���͵������ĵ�ַ/��С��ϣ����ǹ�����ϵͳʹ��.
Public Type IMAGE_OPTIONAL_HEADER32
    'Standard fields.��׼�ֶ�
    Magic               As Integer              '����޷�������ָ����ӳ���ļ���״̬����õ�������0x10B������������һ�������Ŀ�ִ���ļ���0x107 ��������һ��ROM ӳ��0x20B ��������һ��PE32 + ��ִ���ļ�?
    MajorLinkerVersion  As Byte                 '�����������汾��
    MinorLinkerVersion  As Byte                 '�������Ĵΰ汾��
    SizeOfCode          As Long                 '����ڣ�.text���Ĵ�С������ж������ڵĻ����������д���ڵĺ͡�
    SizeOfInitializedData   As Long             '�ѳ�ʼ�����ݽڵĴ�С������ж�����������ݽڵĻ�������������Щ���ݽڵĺ͡�
    SizeOfUninitializedData As Long             'δ��ʼ�����ݽڣ�.bss���Ĵ�С������ж��.bss �ڵĻ�������������Щ�ڵĺ͡�
    AddressOfEntryPoint As Long                 '����ִ���ļ������ؽ��ڴ�ʱ����ڵ������ӳ���ַ��ƫ�Ƶ�ַ.����һ�����ӳ����˵��������������ַ�������豸����������˵�����ǳ�ʼ�������ĵ�ַ����ڵ����DLL��˵�ǿ�ѡ�ġ������������ڵ�Ļ�����������Ϊ0.
    BaseOfCode          As Long                 '��ӳ�񱻼��ؽ��ڴ�ʱ����ڵĿ�ͷ�����ӳ���ַ��ƫ�Ƶ�ַ��
    BaseOfData          As Long                 '��ӳ�񱻼��ؽ��ڴ�ʱ���ݽڵĿ�ͷ�����ӳ���ַ��ƫ�Ƶ�ַ��PE32����
    ' NT additional fields.NT�����ֶ�
    ImageBase           As Long                 '�����ؽ��ڴ�ʱӳ��ĵ�һ���ֽڵ���ѡ��ַ.��������64K �ı���.DLLĬ����0x10000000��Windows CE EXEĬ����0x00010000.Windows NT?Windows 2000��Windows XP��Windows 95��Windows 98 ��Windows Me Ĭ����0x00400000��
    SectionAlignment    As Long                 '�����ؽ��ڴ�ʱ�ڵĶ���ֵ�����ֽڼƣ�����������ڻ����FileAlignment.Ĭ������Ӧϵͳ��ҳ���С
    FileAlignment       As Long                 '��������ӳ���ļ��Ľ��е�ԭʼ���ݵĶ������ӣ����ֽڼƣ�����Ӧ���ǽ���512 ��64K ֮���2 ���ݣ������������߽�ֵ����Ĭ����512�����SectionAlignment С����Ӧϵͳ��ҳ���С����ôFileAlignment ������SectionAlignment ƥ��.
    MajorOperatingSystemVersion As Integer      '�������ϵͳ�����汾�š�
    MinorOperatingSystemVersion As Integer      '�������ϵͳ�Ĵΰ汾��
    MajorImageVersion   As Integer              'ӳ������汾�š�
    MinorImageVersion   As Integer              'ӳ��Ĵΰ汾�š�
    MajorSubsystemVersion   As Integer          '��ϵͳ�����汾��
    MinorSubsystemVersion   As Integer          '��ϵͳ�Ĵΰ汾��
    Win32VersionValue   As Long                 '����������Ϊ0
    SizeOfImage         As Long                 '��ӳ�񱻼��ؽ��ڴ�ʱ�Ĵ�С�����ֽڼƣ����������е��ļ�ͷ����������SectionAlignment �ı���.
    SizeOfHeaders       As Long                 'MS-DOS ռλ����PE �ļ�ͷ�ͽ�ͷ���ܴ�С����������ΪFileAlignment�ı���.
    Checksum            As Long                 'ӳ���ļ���У��͡�����У��͵��㷨���ϲ�����IMAGEHLP.DLL �С����³����ڼ���ʱ��У����ȷ�����Ƿ�Ϸ�: ���е����������κ�������ʱ�����ص�DLL �Լ����ؽ��ؼ�Windows �����е�DLL.
    Subsystem           As Integer              '���д�ӳ���������ϵͳ
    DllCharacteristics  As Integer              'DLL����
    SizeOfStackReserve  As Long                 '�����Ķ�ջ��С��ֻ��SizeOfStackCommit ָ���Ĳ��ֱ��ύ�������ÿ�ο���һҳ��ֱ�����ﱣ���Ĵ�СΪֹ.
    SizeOfStackCommit   As Long                 '�ύ�Ķ�ջ��С��
    SizeOfHeapReserve   As Long                 '�����ľֲ��ѿռ��С��ֻ��SizeOfHeapCommit ָ���Ĳ��ֱ��ύ�������ÿ�ο���һҳ��ֱ�����ﱣ���Ĵ�СΪֹ.
    SizeOfHeapCommit    As Long                 '�ύ�ľֲ��ѿռ��С��
    LoaderFlags         As Long                 '����������Ϊ0��
    NumberOfRvaAndSizes As Long                 '��ѡ�ļ�ͷ���ಿ��������Ŀ¼��ĸ���.ÿ������Ŀ¼������һ�����λ�úʹ�С.
    DataDirectory(IMAGE_NUMBEROF_DIRECTORY_ENTRIES - 1)     As IMAGE_DATA_DIRECTORY
End Type

Public Type IMAGE_ROM_OPTIONAL_HEADER
    Magic               As Integer              '����޷�������ָ����ӳ���ļ���״̬����õ�������0x10B������������һ�������Ŀ�ִ���ļ���0x107 ��������һ��ROM ӳ��0x20B ��������һ��PE32 + ��ִ���ļ�?
    MajorLinkerVersion  As Byte                 '�����������汾��
    MinorLinkerVersion  As Byte                 '�������Ĵΰ汾��
    SizeOfCode          As Long                 '����ڣ�.text���Ĵ�С������ж������ڵĻ����������д���ڵĺ͡�
    SizeOfInitializedData   As Long             '�ѳ�ʼ�����ݽڵĴ�С������ж�����������ݽڵĻ�������������Щ���ݽڵĺ͡�
    SizeOfUninitializedData As Long             'δ��ʼ�����ݽڣ�.bss���Ĵ�С������ж��.bss �ڵĻ�������������Щ�ڵĺ͡�
    AddressOfEntryPoint As Long                 '����ִ���ļ������ؽ��ڴ�ʱ����ڵ������ӳ���ַ��ƫ�Ƶ�ַ.����һ�����ӳ����˵��������������ַ�������豸����������˵�����ǳ�ʼ�������ĵ�ַ����ڵ����DLL��˵�ǿ�ѡ�ġ������������ڵ�Ļ�����������Ϊ0.
    BaseOfCode          As Long                 '��ӳ�񱻼��ؽ��ڴ�ʱ����ڵĿ�ͷ�����ӳ���ַ��ƫ�Ƶ�ַ��
    BaseOfData          As Long                 '��ӳ�񱻼��ؽ��ڴ�ʱ���ݽڵĿ�ͷ�����ӳ���ַ��ƫ�Ƶ�ַ��PE32����
    BaseOfBss           As Long
    GprMask             As Long
    CprMask(3)          As Long
    GpValue             As Long
End Type
'64λ���ζ���
Private Type Big_Iint
    Low                 As Long
    High                As Long
End Type

Public Type IMAGE_OPTIONAL_HEADER64
    'PE32����
    Magic               As Integer              '����޷�������ָ����ӳ���ļ���״̬����õ�������0x10B������������һ�������Ŀ�ִ���ļ���0x107 ��������һ��ROM ӳ��0x20B ��������һ��PE32 + ��ִ���ļ�?
    MajorLinkerVersion  As Byte                 '�����������汾��
    MinorLinkerVersion  As Byte                 '�������Ĵΰ汾��
    SizeOfCode          As Long                 '����ڣ�.text���Ĵ�С������ж������ڵĻ����������д���ڵĺ͡�
    SizeOfInitializedData   As Long             '�ѳ�ʼ�����ݽڵĴ�С������ж�����������ݽڵĻ�������������Щ���ݽڵĺ͡�
    SizeOfUninitializedData As Long             'δ��ʼ�����ݽڣ�.bss���Ĵ�С������ж��.bss �ڵĻ�������������Щ�ڵĺ͡�
    AddressOfEntryPoint As Long                 '����ִ���ļ������ؽ��ڴ�ʱ����ڵ������ӳ���ַ��ƫ�Ƶ�ַ.����һ�����ӳ����˵��������������ַ�������豸����������˵�����ǳ�ʼ�������ĵ�ַ����ڵ����DLL��˵�ǿ�ѡ�ġ������������ڵ�Ļ�����������Ϊ0.
    BaseOfCode          As Long                 '��ӳ�񱻼��ؽ��ڴ�ʱ����ڵĿ�ͷ�����ӳ���ַ��ƫ�Ƶ�ַ��

    ImageBase           As Big_Iint             '�����ؽ��ڴ�ʱӳ��ĵ�һ���ֽڵ���ѡ��ַ.��������64K �ı���.DLLĬ����0x10000000��Windows CE EXEĬ����0x00010000.Windows NT?Windows 2000��Windows XP��Windows 95��Windows 98 ��Windows Me Ĭ����0x00400000��
    SectionAlignment    As Long                 '�����ؽ��ڴ�ʱ�ڵĶ���ֵ�����ֽڼƣ�����������ڻ����FileAlignment.Ĭ������Ӧϵͳ��ҳ���С
    FileAlignment       As Long                 '��������ӳ���ļ��Ľ��е�ԭʼ���ݵĶ������ӣ����ֽڼƣ�����Ӧ���ǽ���512 ��64K ֮���2 ���ݣ������������߽�ֵ����Ĭ����512�����SectionAlignment С����Ӧϵͳ��ҳ���С����ôFileAlignment ������SectionAlignment ƥ��.
    MajorOperatingSystemVersion As Integer      '�������ϵͳ�����汾�š�
    MinorOperatingSystemVersion As Integer      '�������ϵͳ�Ĵΰ汾��
    MajorImageVersion   As Integer              'ӳ������汾�š�
    MinorImageVersion   As Integer              'ӳ��Ĵΰ汾�š�
    MajorSubsystemVersion   As Integer          '��ϵͳ�����汾��
    MinorSubsystemVersion   As Integer          '��ϵͳ�Ĵΰ汾��
    Win32VersionValue   As Long                 '����������Ϊ0
    SizeOfImage         As Long                 '��ӳ�񱻼��ؽ��ڴ�ʱ�Ĵ�С�����ֽڼƣ����������е��ļ�ͷ����������SectionAlignment �ı���.
    SizeOfHeaders       As Long                 'MS-DOS ռλ����PE �ļ�ͷ�ͽ�ͷ���ܴ�С����������ΪFileAlignment�ı���.
    Checksum            As Long                 'ӳ���ļ���У��͡�����У��͵��㷨���ϲ�����IMAGEHLP.DLL �С����³����ڼ���ʱ��У����ȷ�����Ƿ�Ϸ�: ���е����������κ�������ʱ�����ص�DLL �Լ����ؽ��ؼ�Windows �����е�DLL.
    Subsystem           As Integer              '���д�ӳ���������ϵͳ
    DllCharacteristics  As Integer              'DLL����
    SizeOfStackReserve  As Big_Iint             '�����Ķ�ջ��С��ֻ��SizeOfStackCommit ָ���Ĳ��ֱ��ύ�������ÿ�ο���һҳ��ֱ�����ﱣ���Ĵ�СΪֹ.
    SizeOfStackCommit   As Big_Iint             '�ύ�Ķ�ջ��С��
    SizeOfHeapReserve   As Big_Iint             '�����ľֲ��ѿռ��С��ֻ��SizeOfHeapCommit ָ���Ĳ��ֱ��ύ�������ÿ�ο���һҳ��ֱ�����ﱣ���Ĵ�СΪֹ.
    SizeOfHeapCommit    As Big_Iint             '�ύ�ľֲ��ѿռ��С
    LoaderFlags         As Long                 '����������Ϊ0��
    NumberOfRvaAndSizes As Long                 '��ѡ�ļ�ͷ���ಿ��������Ŀ¼��ĸ���.ÿ������Ŀ¼������һ�����λ�úʹ�С.
    DataDirectory(IMAGE_NUMBEROF_DIRECTORY_ENTRIES - 1)     As IMAGE_DATA_DIRECTORY
End Type

'IMAGE_OPTIONAL_HEADER Magic
Public Const IMAGE_NT_OPTIONAL_HDR32_MAGIC      As Integer = &H10B                  '����һ��32λ�����ļ�
Public Const IMAGE_NT_OPTIONAL_HDR64_MAGIC      As Integer = &H20B                  '����һ��PE32+��ִ���ļ�
Public Const IMAGE_ROM_OPTIONAL_HDR_MAGIC       As Integer = &H107                  '����һ��ROM����

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
Public Const IMAGE_SUBSYSTEM_UNKNOWN            As Integer = 0                      ' Unknown subsystem. δ֪��ϵͳ
Public Const IMAGE_SUBSYSTEM_NATIVE             As Integer = 1                      ' Image doesn't require a subsystem.�豸���������Native Windows ����
Public Const IMAGE_SUBSYSTEM_WINDOWS_GUI        As Integer = 2                      ' Image runs in the Windows GUI subsystem.Windows ͼ���û����棨GUI����ϵͳ
Public Const IMAGE_SUBSYSTEM_WINDOWS_CUI        As Integer = 3                      ' Image runs in the Windows character subsystem.Windows �ַ�ģʽ��CUI����ϵͳ
Public Const IMAGE_SUBSYSTEM_OS2_CUI            As Integer = 5                      ' image runs in the OS/2 character subsystem.
Public Const IMAGE_SUBSYSTEM_POSIX_CUI          As Integer = 7                      ' image runs in the Posix character subsystem.Posix �ַ�ģʽ��ϵͳ
Public Const IMAGE_SUBSYSTEM_NATIVE_WINDOWS     As Integer = 8                      ' image is a native Win9x driver.
Public Const IMAGE_SUBSYSTEM_WINDOWS_CE_GUI     As Integer = 9                      ' Image runs in the Windows CE subsystem. Windows CE
Public Const IMAGE_SUBSYSTEM_EFI_APPLICATION    As Integer = 10                     ' ����չ�̼��ӿڣ�EFI��Ӧ�ó���
Public Const IMAGE_SUBSYSTEM_EFI_BOOT_SERVICE_DRIVER    As Integer = 11             ' �����������EFI ��������
Public Const IMAGE_SUBSYSTEM_EFI_RUNTIME_DRIVER As Integer = 12                     ' ������ʱ�����EFI ��������
Public Const IMAGE_SUBSYSTEM_EFI_ROM            As Integer = 13                     ' EFI ROM ӳ��
Public Const IMAGE_SUBSYSTEM_XBOX               As Integer = 14                     ' XBOX
Public Const IMAGE_SUBSYSTEM_WINDOWS_BOOT_APPLICATION   As Integer = 16
'IMAGE_OPTIONAL_HEADER DllCharacteristics Entries

'IMAGE_OPTIONAL_HEADER DllCharacteristics Entries
'Public Const IMAGE_LIBRARY_PROCESS_INIT            0x0001                           ' Reserved.����������Ϊ0��
'Public Const IMAGE_LIBRARY_PROCESS_TERM            0x0002                           ' Reserved.����������Ϊ0��
'Public Const IMAGE_LIBRARY_THREAD_INIT             0x0004                           ' Reserved.����������Ϊ0��
'Public Const IMAGE_LIBRARY_THREAD_TERM             0x0008                           ' Reserved.����������Ϊ0��
Public Const IMAGE_DLLCHARACTERISTICS_HIGH_ENTROPY_VA   As Integer = &H20           ' Image can handle a high entropy 64-bit virtual address space.
Public Const IMAGE_DLLCHARACTERISTICS_DYNAMIC_BASE  As Integer = &H40               ' DLL can move.DLL �����ڼ���ʱ���ض�λ��
Public Const IMAGE_DLLCHARACTERISTICS_FORCE_INTEGRITY   As Integer = &H80           ' Code Integrity Image  ǿ�ƽ��д���������У�顣
Public Const IMAGE_DLLCHARACTERISTICS_NX_COMPAT As Integer = &H100                  ' Image is NX compatible    ӳ�������NX��
Public Const IMAGE_DLLCHARACTERISTICS_NO_ISOLATION As Integer = &H200               ' Image understands isolation and doesn't want it   ���Ը��룬�����������ӳ��
Public Const IMAGE_DLLCHARACTERISTICS_NO_SEH    As Integer = &H400                  ' Image does not use SEH.  No SE handler may reside in this image   ��ʹ�ýṹ���쳣��SE�������ڴ�ӳ���в��ܵ���SE �������
Public Const IMAGE_DLLCHARACTERISTICS_NO_BIND   As Integer = &H800                  ' Do not bind this image.   ����ӳ��
Public Const IMAGE_DLLCHARACTERISTICS_APPCONTAINER As Integer = &H1000              ' Image should execute in an AppContainer
Public Const IMAGE_DLLCHARACTERISTICS_WDM_DRIVER   As Integer = &H2000              ' Driver uses WDM model WDM ��������
'                                                As Integer=&H4000                   ' Reserved.    ����������Ϊ0��
Public Const IMAGE_DLLCHARACTERISTICS_TERMINAL_SERVER_AWARE     As Integer = &H8000 '���������ն˷�������

'IMAGE_OPTIONAL_HEADER Directory Entries
Public Const IMAGE_DIRECTORY_ENTRY_EXPORT       As Integer = 0                      ' Export Directory  ������ĵ�ַ�ʹ�С
Public Const IMAGE_DIRECTORY_ENTRY_IMPORT       As Integer = 1                      ' Import Directory  �����ĵ�ַ�ʹ�С
Public Const IMAGE_DIRECTORY_ENTRY_RESOURCE     As Integer = 2                      ' Resource Directory ��Դ��ĵ�ַ�ʹ�С
Public Const IMAGE_DIRECTORY_ENTRY_EXCEPTION    As Integer = 3                      ' Exception Directory   �쳣��ĵ�ַ�ʹ�С
Public Const IMAGE_DIRECTORY_ENTRY_SECURITY     As Integer = 4                      ' Security Directory    ����֤���ĵ�ַ�ʹ�С Certificate Table ��ָ������֤�����Щ֤�鲢����Ϊӳ���һ���ֱ����ؽ��ڴ档��ʱ���ĵ�һ������һ���ļ�ָ�룬������ͨ����RVA
Public Const IMAGE_DIRECTORY_ENTRY_BASERELOC    As Integer = 5                      ' Base Relocation Table ��ַ�ض�λ��ĵ�ַ�ʹ�С
Public Const IMAGE_DIRECTORY_ENTRY_DEBUG        As Integer = 6                      ' Debug Directory       ����������ʼ��ַ�ʹ�С
'Public Const IMAGE_DIRECTORY_ENTRY_COPYRIGHT    As Integer = 7                      ' (X86 usage)          ����������ʼ��ַ�ʹ�С��
Public Const IMAGE_DIRECTORY_ENTRY_ARCHITECTURE As Integer = 7                      ' Architecture Specific Data ����������Ϊ0
Public Const IMAGE_DIRECTORY_ENTRY_GLOBALPTR    As Integer = 8                      ' RVA of GP             �����洢��ȫ��ָ��Ĵ����е�һ��ֵ��RVA������ṹ��Size �����Ϊ0
Public Const IMAGE_DIRECTORY_ENTRY_TLS          As Integer = 9                      ' TLS Directory         �ֲ߳̾��洢��TLS����ĵ�ַ�ʹ�С
Public Const IMAGE_DIRECTORY_ENTRY_LOAD_CONFIG  As Integer = 10                     ' Load Configuration Directory  �������ñ�ĵ�ַ�ʹ�С
Public Const IMAGE_DIRECTORY_ENTRY_BOUND_IMPORT As Integer = 11                     ' Bound Import Directory in headers �󶨵����ĵ�ַ�ʹ�С
Public Const IMAGE_DIRECTORY_ENTRY_IAT          As Integer = 12                     ' Import Address Table  �󶨵�����ұ�ĵ�ַ�ʹ�С
Public Const IMAGE_DIRECTORY_ENTRY_DELAY_IMPORT As Integer = 13                     ' Delay Load Import Descriptors �ӳٵ����������ĵ�ַ�ʹ�С
Public Const IMAGE_DIRECTORY_ENTRY_COM_DESCRIPTOR   As Integer = 14                 ' COM Runtime descriptor    CLR ����ʱͷ���ĵ�ַ�ʹ�С
Public Const DATA_DIRECTORY_OTHER               As Integer = 15                     ' COM Runtime descriptor    ����������Ϊ0

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
    vName(IMAGE_SIZEOF_SHORT_NAME - 1) As Byte  '����һ��8 �ֽڵ�UTF-8 ������ַ���������8 �ֽ�ʱ��NULL ��䡣�����������8 �ֽڣ��Ǿ�û������NULL �ַ���������Ƹ����Ļ������������һ��б�ܣ�/�����һ����ASCII ���ʾ��ʮ�����������ʮ��������ʾ�ַ������е�ƫ�ơ�����ӳ��ʹ���ַ�����Ҳ��֧�ֳ��ȳ���8 �ֽڵĽ���.���Ŀ���ļ����г������Ľ����Ҫ�����ڿ�ִ���ļ��У���ô��Ӧ�ĳ������ᱻ�ضϡ�

    Misc                As Long                 'PhysicalAddress or VirtualSize
    'PhysicalAddress
    'VirtualSize                                '�����ؽ��ڴ�ʱ����ڵ��ܴ�С�������ֵ��SizeOfRawData ����ô����Ĳ�����0 ��䡣�������Կ�ִ��ӳ���ǺϷ��ģ�����Ŀ���ļ���˵����Ӧ��Ϊ0��

    VirtualAddress      As Long                 '���ڿ�ִ��ӳ����˵��������ֵ������ڱ����ؽ��ڴ�֮�����ĵ�һ���ֽ������ӳ���ַ��ƫ�Ƶ�ַ������Ŀ���ļ���˵��������ֵ��û���ض�λ֮ǰ���һ���ֽڵĵ�ַ��Ϊ�˼������������Ӧ�ðѴ�ֵ����Ϊ0���������ֵ�Ǹ�����ֵ���������ض�λʱӦ�ô�ƫ�Ƶ�ַ�м�ȥ���ֵ
    SizeOfRawData       As Long                 '������Ŀ���ļ���˵���ڵĴ�С���ߣ�����ӳ���ļ���˵�������ļ����ѳ�ʼ�����ݵĴ�С�����ڿ�ִ��ӳ����˵���������ǿ�ѡ�ļ�ͷ��FileAlignment ��ı���.�����С��VirtualSize ���ֵ�����µĲ�����0 ��䡣����SizeOfRawData ��Ҫ�������룬����VirtualSize�򲢲����룬��˿��ܳ���SizeOfRawData �����VirtualSize ������.�����н�����δ��ʼ��������ʱ�������Ӧ��Ϊ0��
    PointerToRawData    As Long                 'ָ��COFF �ļ��нڵĵ�һ��ҳ����ļ�ָ�롣���ڿ�ִ��ӳ����˵���������ǿ�ѡ�ļ�ͷ��FileAlignment ��ı���������Ŀ���ļ���˵��Ҫ�����õ����ܣ���ֵӦ�ð�4 �ֽڱ߽���롣�����н�����δ��ʼ��������ʱ�������Ӧ��Ϊ0��
    PointerToRelocations    As Long             'ָ������ض�λ�ͷ���ļ�ָ�롣���ڿ�ִ���ļ�����û���ض�λ����ļ���˵����ֵӦ��Ϊ0��
    PointerToLinenumbers    As Long             'ָ������к��ͷ���ļ�ָ�롣���û��COFF�к���Ϣ�Ļ�����ֵӦ��Ϊ0������ӳ����˵����ֵӦ��Ϊ0����Ϊ�Ѿ����޳�ʹ��COFF ������Ϣ��
    NumberOfRelocations As Integer              '�����ض�λ��ĸ��������ڿ�ִ��ӳ����˵����ֵӦ��Ϊ0
    NumberOfLinenumbers As Integer              '�����к���ĸ���������ӳ����˵����ֵӦ��Ϊ0����Ϊ�Ѿ����޳�ʹ��COFF ������Ϣ�ˡ�
    Characteristics     As Long                 '�����������ı�־��
End Type

'IMAGE_SECTION_HEADER Section characteristics.
'Public Const IMAGE_SCN_TYPE_REG                 As Long=&H00000000                  ' Reserved.����������ʹ��
'Public Const IMAGE_SCN_TYPE_DSECT               As Long=&H00000001                  ' Reserved.����������ʹ��
'Public Const IMAGE_SCN_TYPE_NOLOAD              As Long=&H00000002                  ' Reserved.����������ʹ��
'Public Const IMAGE_SCN_TYPE_GROUP               As Long=&H00000004                  ' Reserved.����������ʹ��
Public Const IMAGE_SCN_TYPE_NO_PAD              As Long = &H8                       ' Reserved.������ڽ�β����һ���߽�֮�䲻����䡣�˱�־�����������Ѿ���IMAGE_SCN_ALIGN_1BYTES ��־ȡ��.�˱�־����Ŀ���ļ��Ϸ�
'Public Const IMAGE_SCN_TYPE_COPY                As Long=&H00000010                  ' Reserved.����������ʹ�á�

Public Const IMAGE_SCN_CNT_CODE                 As Long = &H20                      ' Section contains code.�˽ڰ�����ִ�д��롣
Public Const IMAGE_SCN_CNT_INITIALIZED_DATA     As Long = &H40                      ' Section contains initialized data.�˽ڰ����ѳ�ʼ�������ݡ�
Public Const IMAGE_SCN_CNT_UNINITIALIZED_DATA   As Long = &H80                      ' Section contains uninitialized data.�˽ڰ���δ��ʼ�������ݡ�

Public Const IMAGE_SCN_LNK_OTHER                As Long = &H100                     ' Reserved.����������ʹ�á�
Public Const IMAGE_SCN_LNK_INFO                 As Long = &H200                     ' Section contains comments or some other type of information.�˽ڰ���ע�ͻ���������Ϣ..drectve �ھ�����������?�˱�־����Ŀ���ļ��Ϸ�
'Public Const IMAGE_SCN_TYPE_OVER                As Long=&H00000400                  ' Reserved.����������ʹ�á�
Public Const IMAGE_SCN_LNK_REMOVE               As Long = &H800                     ' Section contents will not become part of image.�˽ڲ����Ϊ�����γɵ�ӳ���ļ���һ����.�˱�־����Ŀ���ļ��Ϸ�
Public Const IMAGE_SCN_LNK_COMDAT               As Long = &H1000                    ' Section contents comdat.�˽ڰ���COMDAT ���ݡ�
'Public Const                                    As Long=&H00002000                  ' Reserved.
'Public Const IMAGE_SCN_MEM_PROTECTED - Obsolete As Long=&H00004000
Public Const IMAGE_SCN_NO_DEFER_SPEC_EXC        As Long = &H4000                    ' Reset speculative exceptions handling bits in the TLB entries for this section.
Public Const IMAGE_SCN_GPREL                    As Long = &H8000                    ' Section content can be accessed relative to GP �˽ڰ���ͨ��ȫ��ָ�루GP�������õ�����
Public Const IMAGE_SCN_MEM_FARDATA              As Long = &H8000
'Public Const IMAGE_SCN_MEM_SYSHEAP  - Obsolete  As Long=&H00010000
Public Const IMAGE_SCN_MEM_PURGEABLE            As Long = &H20000                   '����������ʹ��
Public Const IMAGE_SCN_MEM_16BIT                As Long = &H20000                   '����������ʹ��
Public Const IMAGE_SCN_MEM_LOCKED               As Long = &H40000                   '����������ʹ��
Public Const IMAGE_SCN_MEM_PRELOAD              As Long = &H80000                   '����������ʹ��

Public Const IMAGE_SCN_ALIGN_1BYTES             As Long = &H100000                  '��1 �ֽڱ߽�������ݡ��˱�־����Ŀ���ļ��Ϸ�.
Public Const IMAGE_SCN_ALIGN_2BYTES             As Long = &H200000                  '��2�ֽڱ߽�������ݡ��˱�־����Ŀ���ļ��Ϸ�.
Public Const IMAGE_SCN_ALIGN_4BYTES             As Long = &H300000                  '��4�ֽڱ߽�������ݡ��˱�־����Ŀ���ļ��Ϸ�.
Public Const IMAGE_SCN_ALIGN_8BYTES             As Long = &H400000                  '��8�ֽڱ߽�������ݡ��˱�־����Ŀ���ļ��Ϸ�.
Public Const IMAGE_SCN_ALIGN_16BYTES            As Long = &H500000                  ' Default alignment if no others are specified.��16 �ֽڱ߽�������ݡ��˱�־����Ŀ���ļ��Ϸ�
Public Const IMAGE_SCN_ALIGN_32BYTES            As Long = &H600000                  '��32�ֽڱ߽�������ݡ��˱�־����Ŀ���ļ��Ϸ�.
Public Const IMAGE_SCN_ALIGN_64BYTES            As Long = &H700000                  '��64�ֽڱ߽�������ݡ��˱�־����Ŀ���ļ��Ϸ�.
Public Const IMAGE_SCN_ALIGN_128BYTES           As Long = &H800000                  '��128�ֽڱ߽�������ݡ��˱�־����Ŀ���ļ��Ϸ�.
Public Const IMAGE_SCN_ALIGN_256BYTES           As Long = &H900000                  '��256�ֽڱ߽�������ݡ��˱�־����Ŀ���ļ��Ϸ�.
Public Const IMAGE_SCN_ALIGN_512BYTES           As Long = &HA00000                  '��512�ֽڱ߽�������ݡ��˱�־����Ŀ���ļ��Ϸ�.
Public Const IMAGE_SCN_ALIGN_1024BYTES          As Long = &HB00000                  '��1024�ֽڱ߽�������ݡ��˱�־����Ŀ���ļ��Ϸ�.
Public Const IMAGE_SCN_ALIGN_2048BYTES          As Long = &HC00000                  '��2048�ֽڱ߽�������ݡ��˱�־����Ŀ���ļ��Ϸ�.
Public Const IMAGE_SCN_ALIGN_4096BYTES          As Long = &HD00000                  '��4096�ֽڱ߽�������ݡ��˱�־����Ŀ���ļ��Ϸ�.
Public Const IMAGE_SCN_ALIGN_8192BYTES          As Long = &HE00000                  '��8192�ֽڱ߽�������ݡ��˱�־����Ŀ���ļ��Ϸ�.
' Unused                                         As Long=&H00F00000
Public Const IMAGE_SCN_ALIGN_MASK               As Long = &HF00000                  '

Public Const IMAGE_SCN_LNK_NRELOC_OVFL          As Long = &H1000000                 ' Section contains extended relocations. �˽ڰ�����չ���ض�λ��Ϣ��
'                                       IMAGE_SCN_LNK_NRELOC_OVFL ��־���������ض�λ��ĸ��������˽�ͷ��Ϊÿ���ڱ�����16 λ���ܱ�ʾ�ķ�Χ?��������˴˱�־���ҽ�ͷ�е�NumberOfRelocations ���ֵ��0xffff����ôʵ�ʵ��ض�λ������������ڵ�һ����
'                                       ��λ���VirtualAddress ��32 λ���С����������IMAGE_SCN_LNK_NRELOC_OVFL��־�����е��ض�λ��ĸ�������0xffff�����ʾ�����˴���
Public Const IMAGE_SCN_MEM_DISCARDABLE          As Long = &H2000000                 ' Section can be discarded.�˽ڿ�������Ҫʱ��������
Public Const IMAGE_SCN_MEM_NOT_CACHED           As Long = &H4000000                 ' Section is not cachable.�˽ڲ��ܱ����档
Public Const IMAGE_SCN_MEM_NOT_PAGED            As Long = &H8000000                 ' Section is not pageable.�˽ڲ��ܱ�������ҳ���ļ��С�
Public Const IMAGE_SCN_MEM_SHARED               As Long = &H10000000                ' Section is shareable.�˽ڿ������ڴ��й���
Public Const IMAGE_SCN_MEM_EXECUTE              As Long = &H20000000                ' Section is executable.�˽ڿ�����Ϊ����ִ�С�
Public Const IMAGE_SCN_MEM_READ                 As Long = &H40000000                ' Section is readable.�˽ڿɶ���
Public Const IMAGE_SCN_MEM_WRITE                As Long = &H80000000                ' Section is writeable.�˽ڿ�д

'TLS Chaacteristic Flags
Public Const IMAGE_SCN_SCALE_INDEX              As Long = &H1                       'Tls index is scaled

'Symbol format.
Public Type IMAGE_SYMBOL
    ShortName(7)        As Byte                 '�������ƣ�����һ����������Ա��ɵĹ����塣������Ƶĳ��Ȳ�����8 ���ֽڣ���ô������һ��8 �ֽڳ������顣
    'Union
'    Short               As Long                 'if 0, use LongName
'    Long                As Long                 'offset into string table
    'Union
'    LongName(1)         As Long                 'PBYTE [2]
    Value               As Long                 '�������ص�ֵ��������������SectionNumber ��StorageClass ��������.��ͨ����ʾ���ض�λ�ĵ�ַ
    SectionNumber       As Integer              '��������������ǽڱ����������1 ��ʼ�������Ա�ʶ����˷��ŵĽ�
    vType               As Integer              'һ����ʾ���͵����֡�Microsoft �Ĺ��߽�������Ϊ0x20������Ǻ���������0x0��������Ǻ�����
    StorageClass        As Byte                 '����һ����ʾ�洢����ö������ֵ
    NumberOfAuxSymbols  As Byte                 '���ڱ���¼����ĸ������ű���ĸ�����
End Type
Public Const IMAGE_SIZEOF_SYMBOL                As Integer = 18

Public Type IMAGE_SYMBOL_EX
    ShortName(7)        As Byte                 '�������ƣ�����һ����������Ա��ɵĹ����塣������Ƶĳ��Ȳ�����8 ���ֽڣ���ô������һ��8 �ֽڳ������顣
    'Union
'    Short               As Long                 'if 0, use LongName
'    Long                As Long                 'offset into string table
    'Union
'    LongName(1)         As Long                 'PBYTE [2]
    Value               As Long                 '�������ص�ֵ��������������SectionNumber ��StorageClass ��������.��ͨ����ʾ���ض�λ�ĵ�ַ
    SectionNumber       As Long                 '��������������ǽڱ����������1 ��ʼ�������Ա�ʶ����˷��ŵĽ�
    vType               As Integer              'һ����ʾ���͵����֡�Microsoft �Ĺ��߽�������Ϊ0x20������Ǻ���������0x0��������Ǻ�����Type ��ռ2 ���ֽڣ����е�ÿһ���ֽڶ���ʾ������Ϣ����λ�ֽڣ�LSB����ʾ�򵥣��������������ͣ���λ�ֽڣ�MSB����ʾ�������ͣ�������ڣ�
                                                'MSB:�������ͣ��ޡ�ָ�롢����������,LSB:�������ͣ��������������ȡ�
    StorageClass        As Byte                 '����һ����ʾ�洢����ö������ֵ
    NumberOfAuxSymbols  As Byte                 '���ڱ���¼����ĸ������ű���ĸ�����
End Type
Public Const IMAGE_SIZEOF_SYMBOL_EX             As Integer = 20
' Section values.
' Symbols have a section number of the section in which they are defined. Otherwise, section numbers have the following meanings:
Public Const IMAGE_SYM_UNDEFINED                As Integer = 0                      ' Symbol is undefined or is common.��δΪ�˷��ż�¼����һ���ڡ������ֵ����������һ�������������ط����ⲿ���ţ�������ֵ�������һ����ͨ���ţ����С��Value �������
Public Const IMAGE_SYM_ABSOLUTE                 As Integer = -1                     ' Symbol is an absolute value.�˷����Ǹ����Է��ţ������ض�λ�������Ҳ��ǵ�ַ��
Public Const IMAGE_SYM_DEBUG                    As Integer = -2                     ' Symbol is a special debug item.�˷����ṩ��ͨ������Ϣ���ߵ�����Ϣ������������Ӧ��ĳһ���ڡ�Microsoft �Ĺ��߽�.file ��¼���洢���ΪFILE������Ϊ���ֵ��
Public Const IMAGE_SYM_SECTION_MAX              As Integer = &HFEFF                 ' Values 0xFF00-0xFFFF are special
Public Const IMAGE_SYM_SECTION_MAX_EX           As Integer = &HFFFF

' IMAGE_SYMBOL Type (fundamental) values.
Public Const IMAGE_SYM_TYPE_NULL                As Integer = &H0                    ' no type.������Ϣ�����ڣ�������δ֪�Ļ������͡�Microsoft �Ĺ���ʹ�����ֵ
Public Const IMAGE_SYM_TYPE_VOID                As Integer = &H1                    '���ǺϷ����ͣ�����void ָ��ͺ�����
Public Const IMAGE_SYM_TYPE_CHAR                As Integer = &H2                    ' type character.�ַ��������ŵ�1 ���ֽڣ���
Public Const IMAGE_SYM_TYPE_SHORT               As Integer = &H3                    ' type short integer.����Ϊ2 ���ֽڵĴ�����������
Public Const IMAGE_SYM_TYPE_INT                 As Integer = &H4                    '��Ȼ���������ͣ���Windows ��ͨ��Ϊ4 ���ֽڣ���
Public Const IMAGE_SYM_TYPE_LONG                As Integer = &H5                    '����Ϊ4 ���ֽڵĴ�����������
Public Const IMAGE_SYM_TYPE_FLOAT               As Integer = &H6                    '����Ϊ4 ���ֽڵĸ�������
Public Const IMAGE_SYM_TYPE_DOUBLE              As Integer = &H7                    '����Ϊ8 ���ֽڵĸ�������
Public Const IMAGE_SYM_TYPE_STRUCT              As Integer = &H8                    '�ṹ�塣
Public Const IMAGE_SYM_TYPE_UNION               As Integer = &H9                    '�����塣
Public Const IMAGE_SYM_TYPE_ENUM                As Integer = &HA                    ' enumeration.ö�����͡�
Public Const IMAGE_SYM_TYPE_MOE                 As Integer = &HB                    ' member of enumeration.ö�����ͳ�Ա������ֵ����
Public Const IMAGE_SYM_TYPE_BYTE                As Integer = &HC                    '�ֽڣ�����Ϊ1 ���ֽڵ��޷���������
Public Const IMAGE_SYM_TYPE_WORD                As Integer = &HD                    '�֣����������ֽڵ��޷���������
Public Const IMAGE_SYM_TYPE_UINT                As Integer = &HE                    '����Ϊ��Ȼ�ߴ���޷���������ͨ��Ϊ4 ���ֽڣ���
Public Const IMAGE_SYM_TYPE_DWORD               As Integer = &HF                    '����Ϊ4 ���ֽڵ��޷���������
Public Const IMAGE_SYM_TYPE_PCODE               As Integer = &H8000                 '
' Type (derived) values.
Public Const IMAGE_SYM_DTYPE_NULL               As Integer = 0                      ' no derived type.�ǵ������ͣ��˷����Ǽ򵥵ı���������
Public Const IMAGE_SYM_DTYPE_POINTER            As Integer = 1                      ' pointer.�˷�����ָ��������͵�ָ�롣
Public Const IMAGE_SYM_DTYPE_FUNCTION           As Integer = 2                      ' function.�˷����Ƿ��ػ������͵ĺ�����
Public Const IMAGE_SYM_DTYPE_ARRAY              As Integer = 3                      ' array.�˷������ɻ���������ɵ����顣

'IMAGE_SYMBOL Storage classes. ע��StorageClass ���ǳ���Ϊ1 ���ֽڵ��޷���������������������ֵΪ-1 �Ļ���ʵ����Ӧ�ñ�������������ȵ��޷�������Ҳ����0xFF�����ܴ�ͳ��COFF ��ʽʹ�����洢��𣬵���Microsoft �Ĺ���ʹ��Visual
'                               C++������Ϣ����ʾ�󲿷ַ�����Ϣ����ͨ����ʹ�����ִ洢���EXTERNAL��2����STATIC��3����FUNCTION��101����FILE��103����
Public Const IMAGE_SYM_CLASS_END_OF_FUNCTION    As Byte = &HFF                      '(0xFF)��ʾ������β��������ţ����ڵ��ԡ�
Public Const IMAGE_SYM_CLASS_NULL               As Byte = &H0                       'δ������洢���
Public Const IMAGE_SYM_CLASS_AUTOMATIC          As Byte = &H1                       '�Զ�����ջ��������Value ��ָ���˱�����ջ֡�е�ƫ��
Public Const IMAGE_SYM_CLASS_EXTERNAL           As Byte = &H2                       'Microsoft �Ĺ���ʹ�ô�ֵ����ʾ�ⲿ����.���SectionNumber ��Ϊ0��IMAGE_SYM_UNDEFINED������ôValue �������С�����SectionNumber ��Ϊ0����ôValue ��������е�ƫ��
Public Const IMAGE_SYM_CLASS_STATIC             As Byte = &H3                       '�����ڽ��е�ƫ�ơ����Value ��Ϊ0����ô�˷��ű�ʾ����
Public Const IMAGE_SYM_CLASS_REGISTER           As Byte = &H4                       '�Ĵ���������Value ������Ĵ�����š�
Public Const IMAGE_SYM_CLASS_EXTERNAL_DEF       As Byte = &H5                       '���ⲿ����ķ��š�
Public Const IMAGE_SYM_CLASS_LABEL              As Byte = &H6                       'ģ���ж���Ĵ����š�Value ������˷����ڽ��е�ƫ��
Public Const IMAGE_SYM_CLASS_UNDEFINED_LABEL    As Byte = &H7                       '���õ�δ����Ĵ����š�
Public Const IMAGE_SYM_CLASS_MEMBER_OF_STRUCT   As Byte = &H8                       '�ṹ���Ա��Value ��ָ���ǵڼ�����Ա��
Public Const IMAGE_SYM_CLASS_ARGUMENT           As Byte = &H9                       '��������ʽ�������βΣ���Value ��ָ���ǵڼ�������
Public Const IMAGE_SYM_CLASS_STRUCT_TAG         As Byte = &HA                       '�ṹ������
Public Const IMAGE_SYM_CLASS_MEMBER_OF_UNION    As Byte = &HB                       '�������Ա��Value ��ָ���ǵڼ�����Ա��
Public Const IMAGE_SYM_CLASS_UNION_TAG          As Byte = &HC                       '����������
Public Const IMAGE_SYM_CLASS_TYPE_DEFINITION    As Byte = &HD                       'Typedef �
Public Const IMAGE_SYM_CLASS_UNDEFINED_STATIC   As Byte = &HE                       '��̬����������
Public Const IMAGE_SYM_CLASS_ENUM_TAG           As Byte = &HF                       'ö����������
Public Const IMAGE_SYM_CLASS_MEMBER_OF_ENUM     As Byte = &H10                      'ö�����ͳ�Ա��Value ��ָ���ǵڼ�����Ա
Public Const IMAGE_SYM_CLASS_REGISTER_PARAM     As Byte = &H11                      '�Ĵ���������
Public Const IMAGE_SYM_CLASS_BIT_FIELD          As Byte = &H12                      'λ��Value ��ָ����λ���еĵڼ���λ��

Public Const IMAGE_SYM_CLASS_FAR_EXTERNAL       As Byte = &H44                      '

Public Const IMAGE_SYM_CLASS_BLOCK              As Byte = &H64                      '.bb��beginning of block���鿪ͷ)��.eb ��¼��end of block�����β����Value ���Ǵ���λ�ã�����һ�����ض�λ�ĵ�ַ
Public Const IMAGE_SYM_CLASS_FUNCTION           As Byte = &H65                      'Microsoft �Ĺ����ô�ֵ����ʾ���庯����Χ�ķ��ż�¼����Щ���ż�¼�ֱ��ǣ�.bf��begin function��������ͷ����.ef��endfunction��������β���Լ�.lf��lines in function�������е��У�������.lf ��¼��˵��Value �������Դ�����д˺�����ռ������������.ef ��¼��˵��Value ������˺�������Ĵ�С
Public Const IMAGE_SYM_CLASS_END_OF_STRUCT      As Byte = &H66                      '�ṹ��ĩβ��
Public Const IMAGE_SYM_CLASS_FILE               As Byte = &H67                      'Microsoft �Ĺ����Լ���ͳCOFF ��ʽ��ʹ�ô�ֵ����ʾԴ�ļ����ż�¼.���ַ��ű��¼������Ÿ����ļ����ĸ������ű��¼
'new
Public Const IMAGE_SYM_CLASS_SECTION            As Byte = &H68                      '�ڵĶ��壨Microsoft �Ĺ���ʹ��STATIC �洢�����棩
Public Const IMAGE_SYM_CLASS_WEAK_EXTERNAL      As Byte = &H69                      '���ⲿ���š�

Public Const IMAGE_SYM_CLASS_CLR_TOKEN          As Byte = &H6B                      '��ʾCLR �Ǻŵķ��š���������������Ǻŵ�ʮ������ֵ��ASCII ���ʾ?

'CLR �ǺŶ��壨��������Ŀ���ļ���
Public Type IMAGE_AUX_SYMBOL_TOKEN_DEF
    bAuxType            As Byte                 'IMAGE_AUX_SYMBOL_TYPE ����ΪIMAGE_AUX_SYMBOL_TYPE_TOKEN_DEF��1��
    bReserved           As Byte                 ' Must be 0
    SymbolTableIndex    As Long                 '��CLR �ǺŶ����漰��COFF �����ڷ��ű��е�������
    rgbReserved(11)     As Byte                 'Must be 0
End Type
'���һ�����ű��¼ӵ���������ԣ��洢���ΪEXTERNAL��2����Type ���ֵ��������һ��������0x20���Լ�SectionNumber ���ֵ����0�����ͱ�־�ź����Ŀ�ͷ.ע�����һ�����ű��¼SectionNumber ���ֵΪIMAGE_SYM_UNDEFINED��0������ô����������һ��������Ҳû����Ӧ�ĸ������ű��¼
Public Type IMAGE_AUX_SYMBOL
    'Sym
    TagIndex            As Long                 'struct, union, or enum tag index.��Ӧ��.bf��������ͷ����¼�ڷ��ű��е�������
    TotalSize           As Long                 'union_Misc ��������������ɵĿ�ִ�д���Ĵ�С������˺��������ɽڣ���ô���ݶ���ֵ�Ĳ�ͬ����ͷ�е�SizeOfRawData ����ܴ��ڻ���������
                                                '������ֵΪIMAGE_WEAK_EXTERN_SEARCH_NOLIBRARY����������ʱ���ڿ��в���sym1������ֵΪIMAGE_WEAK_EXTERN_SEARCH_LIBRARY����������ʱ�ڿ��в���sym1.������ֵΪIMAGE_WEAK_EXTERN_SEARCH_ALIAS������sym1 ��sym2 �ı�����
'    Linenumber          As Integer              'union_Misc,declaration line number
'    Size                As Integer              'union_Misc,size of struct, union, or enum
    Dimension(3)        As Integer              'Array ,Union_FcnAry,if ISARY, up to 4 dimen.
'    PointerToLinenumber As Long                 'Function,Union_FcnAry  if ISFCN, tag, or .bb����˺��������кż�¼����ô���ֵ��ʾ���ĵ�һ��COFF �кż�¼���ļ�ƫ�ƣ���������ڣ���ô���ֵΪ0
'    PointerToNextFunction   As Long             'Function,Union_FcnAry   if ISFCN, tag, or .bb��Ӧ����һ�������ķ��ű��¼�ڷ��ű��е�����������˺����Ƿ��ű��е����һ����������ô������ֵΪ0
    TvIndex             As Integer              'tv index
    'File
    vName(IMAGE_SIZEOF_SYMBOL - 1)  As Byte     '��ʾԴ�ļ�����ANSI �ַ��������Դ�ļ����ĳ���С����󳤶ȣ���NULL ��䡣
    'Section
    Length              As Long                 'section length�������ݵĴ�С�����ͷ��SizeOfRawData ��һ��
    NumberOfRelocations As Integer              'number of relocation entries �˽����ض�λ�����Ŀ��
    NumberOfLinenumbers As Integer              'number of line numbers �˽����к���Ϣ�����Ŀ��
    Checksum            As Long                 'checksum for communal �������ݵ�У��͡�ֻ�н�ͷ��������IMAGE_SCN_LNK_COMDAT ��־ʱ��ʹ�ô���
    Number              As Integer              'section number to associate with ��˽���صĽ��ڽڱ��е���������1 ��ʼ������COMDAT ��Selection ��Ϊ5 ʱ��ʹ�������
    Selection           As Byte                 'communal selection type ��ʾCOMDAT ѡ��ʽ�����֡������ֻ����COMDAT ��
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
Public Const IMAGE_COMDAT_SELECT_NODUPLICATES   As Byte = 1                         '����˷����Ѿ����������������������һ����multiply defined symbol�����Ŷ��ض��壩������
Public Const IMAGE_COMDAT_SELECT_ANY            As Byte = 2                         '����������Щ����ͬһ��COMDAT ���ŵĽ�����ѡһ�������ࣨδ��ѡ�У��Ľڶ����Ƴ���
Public Const IMAGE_COMDAT_SELECT_SAME_SIZE      As Byte = 3                         '�������Ӷ���������ŵĶ��������ѡһ�������������Щ�����С���ȣ�������������һ�������Ŷ��ض��塱����
Public Const IMAGE_COMDAT_SELECT_EXACT_MATCH    As Byte = 4                         '�������Ӷ���������ŵĶ��������ѡһ�������������Щ���岻�ϸ�һ�£�������������һ�������Ŷ��ض��塱����
Public Const IMAGE_COMDAT_SELECT_ASSOCIATIVE    As Byte = 5                         '���������ĳ����COMDAT �ڱ����ӵĻ����˽�ҲҪ�����ӡ�����ġ�����ĳ���������붨��
                                                                                    '�˽ڵķ��ű��¼��صĸ������ű��¼��Number �����.������ö�����Щ�ڶ�����ж�������ز��֣����������һ�����ж���������һ�����У���������Ϊһ������������ӻ����Ķ���ǳ�����?��˽ڹ��������
                                                                                    '������ĳ�����ڱ���Ҳ��COMDAT �ڲ�����������������COMDAT �ڹ�����Ҳ����˵�����������ĳ�����ڲ��ܽ�Selection ������ΪIMAGE_COMDAT_SELECT_ASSOCIATIVE����
Public Const IMAGE_COMDAT_SELECT_LARGEST        As Byte = 6                         '��������������ŵ����ж�����ѡȡ�������Ľ������ӡ�����������Ĳ�ֹһ������ô�����⼸����������ѡһ��
Public Const IMAGE_COMDAT_SELECT_NEWEST         As Byte = 7                         '

Public Const IMAGE_WEAK_EXTERN_SEARCH_NOLIBRARY As Byte = 1                         '
Public Const IMAGE_WEAK_EXTERN_SEARCH_LIBRARY   As Byte = 2                         '
Public Const IMAGE_WEAK_EXTERN_SEARCH_ALIAS     As Byte = 3                         '

Public Type IMAGE_RELOCATION
    'DUMMYUNIONNAME
    VirtualAddress      As Long                 'DUMMYUNIONNAME_VirtualAddress ��Ҫ�����ض�λ�Ĵ�������ݵĵ�ַ�����Ǵӽڿ�ͷ�����ƫ�ƣ����Ͻڵ�RVA/Offset ���ֵ��
'    RelocCount          As Long                 'DUMMYUNIONNAME_RelocCount,Set to the real count when IMAGE_SCN_LNK_NRELOC_OVFL is set
    SymbolTableIndex    As Long                 '���ű����������0 ��ʼ����������Ÿ����������ض�λ�ĵ�ַ��������ָ�����ŵĴ洢���Ϊ�ڣ���ô���ĵ�ַ���ǵ�һ������ͬ���Ľڵĵ�ַ
    Type                As Byte                 '�ض�λ���͡��Ϸ����ض�λ���������ڻ������͡�
End Type

' I386 relocation types.
Public Const IMAGE_REL_I386_ABSOLUTE            As Byte = &H0                       ' Reference is absolute, no relocation is necessary �ض�λ�����ԡ�
Public Const IMAGE_REL_I386_DIR16               As Byte = &H1                       ' Direct 16-bit reference to the symbols virtual address ��֧�֡�
Public Const IMAGE_REL_I386_REL16               As Byte = &H2                       ' PC-relative 16-bit reference to the symbols virtual address ��֧�֡�
Public Const IMAGE_REL_I386_DIR32               As Byte = &H6                       ' Direct 32-bit reference to the symbols virtual address �ض�λĿ���32 λVA��
Public Const IMAGE_REL_I386_DIR32NB             As Byte = &H7                       ' Direct 32-bit reference to the symbols virtual address, base not included �ض�λĿ���32 λRVA��
Public Const IMAGE_REL_I386_SEG12               As Byte = &H9                       ' Direct 16-bit reference to the segment-selector bits of a 32-bit virtual address ��֧�֡�
Public Const IMAGE_REL_I386_SECTION             As Byte = &HA                       '�����ض�λĿ��Ľڵ�16 λ����������֧�ֵ�����Ϣ
Public Const IMAGE_REL_I386_SECREL              As Byte = &HB                       '�ض�λĿ������������ڽڿ�ͷ��32 λƫ�ơ�����֧�ֵ�����Ϣ�;�̬�ֲ߳̾��洢
Public Const IMAGE_REL_I386_TOKEN               As Byte = &HC                       ' clr token CLR �Ǻš�
Public Const IMAGE_REL_I386_SECREL7             As Byte = &HD                       ' 7 bit offset from base of section containing target ������ض�λĿ�����ڽڻ���ַ��7 λƫ�ơ�
Public Const IMAGE_REL_I386_REL32               As Byte = &H14                      ' PC-relative 32-bit reference to the symbols virtual address �ض�λĿ���32 λ���ƫ�ơ�����֧��x86 ����Է�֧��CALL ָ��

' MIPS relocation types.
Public Const IMAGE_REL_MIPS_ABSOLUTE            As Byte = &H0                       ' Reference is absolute, no relocation is necessary �ض�λ������
Public Const IMAGE_REL_MIPS_REFHALF             As Byte = &H1                       '�ض�λĿ��32 λVA �ĸ�16 λ��
Public Const IMAGE_REL_MIPS_REFWORD             As Byte = &H2                       '�ض�λĿ���32 λVA��
Public Const IMAGE_REL_MIPS_JMPADDR             As Byte = &H3                       '�ض�λĿ��VA �ĵ�26 λ������֧��MIPS ƽ̨��J ��JAL ָ��
Public Const IMAGE_REL_MIPS_REFHI               As Byte = &H4                       '�ض�λĿ��32 λVA �ĸ�16 λ�������ڼ���һ��������ַ�������ָ�������еĵ�һ��ָ��.�����ض�λ���ͺ���������IMAGE_REL_MIPS_PAIR���͵��ض�λ������ߵ�SymbolTableIndex ���������һ��16 λƫ�ƣ��������������ƫ��Ҫ���ӵ��ض�λĿ��λ�õĸ�16 λ
Public Const IMAGE_REL_MIPS_REFLO               As Byte = &H5                       '�ض�λĿ��VA �ĵ�16 λ��
Public Const IMAGE_REL_MIPS_GPREL               As Byte = &H6                       '�ض�λĿ�������GP �Ĵ�����16 λƫ�ƣ�����������
Public Const IMAGE_REL_MIPS_LITERAL             As Byte = &H7                       '��IMAGE_REL_MIPS_GPREL ��ͬ��
Public Const IMAGE_REL_MIPS_SECTION             As Byte = &HA                       '�����ض�λĿ��Ľڵ�16 λ����������֧�ֵ�����Ϣ
Public Const IMAGE_REL_MIPS_SECREL              As Byte = &HB                       '�ض�λĿ������������ڽڿ�ͷ��32 λƫ�ơ�����֧�ֵ�����Ϣ�;�̬�ֲ߳̾��洢
Public Const IMAGE_REL_MIPS_SECRELLO            As Byte = &HC                       ' Low 16-bit section relative referemce (used for >32k TLS) �ض�λĿ������������ڽڿ�ͷ��32 λƫ�Ƶĵ�16 λ
Public Const IMAGE_REL_MIPS_SECRELHI            As Byte = &HD                       ' High 16-bit section relative reference (used for >32k TLS) �ض�λĿ������������ڽڿ�ͷ��32 λVA �ĸ�16 λ.�����ض�λ���ͺ���������IMAGE_REL_MIPS_PAIR ���͵��ض�λ������ߵ�SymbolTableIndex ���������һ��16 λƫ�ƣ��������������ƫ��Ҫ���ӵ��ض�λ����λ�õĸ�16 λ
Public Const IMAGE_REL_MIPS_TOKEN               As Byte = &HE                       ' clr token
Public Const IMAGE_REL_MIPS_JMPADDR16           As Byte = &H10                      '�ض�λĿ��VA �ĵ�26 λ������֧��MIPS16 ��JAL ָ��
Public Const IMAGE_REL_MIPS_REFWORDNB           As Byte = &H22                      '�ض�λĿ���32 λRVA��
Public Const IMAGE_REL_MIPS_PAIR                As Byte = &H25                      'ֻ�н���IMAGE_REL_MIPS_REFHI ��IMAGE_REL_MIPS_SECRELHI ���͵��ض�λʱ�����ض�λ���Ͳ��ǺϷ���.���ض�λ���SymbolTableIndex ���������ƫ�ƶ����Ƿ��ű�����

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
Public Const IMAGE_REL_PPC_ABSOLUTE             As Byte = &H0                       ' NOP �ض�λ�����ԡ�
Public Const IMAGE_REL_PPC_ADDR64               As Byte = &H1                       ' 64-bit address �ض�λĿ���64 λVA��
Public Const IMAGE_REL_PPC_ADDR32               As Byte = &H2                       ' 32-bit address �ض�λĿ���32 λVA��
Public Const IMAGE_REL_PPC_ADDR24               As Byte = &H3                       ' 26-bit address, shifted left 2 (branch absolute) �ض�λĿ��VA �ĵ�24 λ��ֻ�е��ض�λĿ������Ǿ��Է����ҿ��԰�������չ������ԭʼֵʱ���ǺϷ���
Public Const IMAGE_REL_PPC_ADDR16               As Byte = &H4                       ' 16-bit address �ض�λĿ��VA �ĵ�16 λ��
Public Const IMAGE_REL_PPC_ADDR14               As Byte = &H5                       ' 16-bit address, shifted left 2 (load doubleword) �ض�λĿ��VA �ĵ�14 λ��ֻ�е��ض�λĿ������Ǿ��Է����ҿ��԰�������չ������ԭʼֵʱ���ǺϷ���
Public Const IMAGE_REL_PPC_REL24                As Byte = &H6                       ' 26-bit PC-relative offset, shifted left 2 (branch relative) ����λ�������PC ��24 λƫ�ơ�
Public Const IMAGE_REL_PPC_REL14                As Byte = &H7                       ' 16-bit PC-relative offset, shifted left 2 (br cond relative) ����λ�������PC ��14 λƫ�ơ�
Public Const IMAGE_REL_PPC_TOCREL16             As Byte = &H8                       ' 16-bit offset from TOC base
Public Const IMAGE_REL_PPC_TOCREL14             As Byte = &H9                       ' 16-bit offset from TOC base, shifted left 2 (load doubleword)

Public Const IMAGE_REL_PPC_ADDR32NB             As Byte = &HA                       ' 32-bit addr w/o image base  �ض�λĿ���32 λRVA��
Public Const IMAGE_REL_PPC_SECREL               As Byte = &HB                       ' va of containing section (as in an image sectionhdr) �ض�λĿ������������ڽڿ�ͷ��32 λƫ�ơ�����֧�ֵ�����Ϣ�;�̬�ֲ߳̾��洢
Public Const IMAGE_REL_PPC_SECTION              As Byte = &HC                       ' sectionheader number�����ض�λĿ��Ľڵ�16 λ����������֧�ֵ�����Ϣ
Public Const IMAGE_REL_PPC_IFGLUE               As Byte = &HD                       ' substitute TOC restore instruction iff symbol is glue code
Public Const IMAGE_REL_PPC_SECREL16             As Byte = &HF                       ' va of containing section (limited to 16 bits) �ض�λĿ������������ڽڿ�ͷ��16 λƫ�ơ�����֧�ֵ�����Ϣ�;�̬�ֲ߳̾��洢
Public Const IMAGE_REL_PPC_IMGLUE               As Byte = &HE                       ' symbol is glue code; virtual address is TOC restore instruction
Public Const IMAGE_REL_PPC_REFHI                As Byte = &H10                      '�ض�λĿ��32 λVA �ĸ�16 λ�������ڼ���һ��������ַ�������ָ�������еĵ�һ��ָ��?�����ض�λ���ͺ���������IMAGE_REL_PPC_PAIR���͵��ض�λ������ߵ�SymbolTableIndex ���������һ��16 λƫ�ƣ��������������ƫ��Ҫ���ӵ��ض�λĿ��λ�õĸ�16 λ
Public Const IMAGE_REL_PPC_REFLO                As Byte = &H11                      '�ض�λĿ��VA �ĵ�16 λ��
Public Const IMAGE_REL_PPC_PAIR                 As Byte = &H12                      'ֻ�н���IMAGE_REL_PPC_REFHI ��IMAGE_REL_PPC_SECRELHI ���͵��ض�λʱ�����ض�λ���Ͳ��ǺϷ���.���ض�λ���SymbolTableIndex ���������ƫ�ƶ����Ƿ��ű������
Public Const IMAGE_REL_PPC_SECRELLO             As Byte = &H13                      ' Low 16-bit section relative reference (used for >32k TLS) �ض�λĿ������������ڽڿ�ͷ��32 λƫ�Ƶĵ�16 λ
Public Const IMAGE_REL_PPC_SECRELHI             As Byte = &H14                      ' High 16-bit section relative reference (used for >32k TLS) �ض�λĿ������������ڽڿ�ͷ��32 λƫ�Ƶĸ�16 λ
Public Const IMAGE_REL_PPC_GPREL                As Byte = &H15                      '�ض�λĿ�������GP �Ĵ�����16 λƫ�ƣ�������������
Public Const IMAGE_REL_PPC_TOKEN                As Byte = &H16                      ' clr token CLR �Ǻš�

Public Const IMAGE_REL_PPC_TYPEMASK             As Byte = &HFF                      ' mask to isolate above values in IMAGE_RELOCATION.Type

' Flag bits in IMAGE_RELOCATION.TYPE
Public Const IMAGE_REL_PPC_NEG                  As Integer = &H100                  ' subtract reloc value rather than adding it
Public Const IMAGE_REL_PPC_BRTAKEN              As Integer = &H200                  ' fix branch prediction bit to predict branch taken
Public Const IMAGE_REL_PPC_BRNTAKEN             As Integer = &H400                  ' fix branch prediction bit to predict branch not taken
Public Const IMAGE_REL_PPC_TOCDEFN              As Integer = &H800                  ' toc slot defined in file (or, data in toc)


' Hitachi SH3 relocation types.
Public Const IMAGE_REL_SH3_ABSOLUTE             As Byte = &H0                       ' No relocation �ض�λ�����ԡ�
Public Const IMAGE_REL_SH3_DIRECT16             As Byte = &H1                       ' 16 bit direct �԰����ض�λĿ�����VA ��16 λ��Ԫ�����á�
Public Const IMAGE_REL_SH3_DIRECT32             As Byte = &H2                       ' 32 bit direct �ض�λĿ����ŵ�32 λVA��
Public Const IMAGE_REL_SH3_DIRECT8              As Byte = &H3                       ' 8 bit direct, -128..255 �԰����ض�λĿ�����VA ��8 λ��Ԫ�����á�
Public Const IMAGE_REL_SH3_DIRECT8_WORD         As Byte = &H4                       ' 8 bit direct .W (0 ext.)�԰����ض�λĿ�����16 λ��ЧVA ��8 λָ�������
Public Const IMAGE_REL_SH3_DIRECT8_LONG         As Byte = &H5                       ' 8 bit direct .L (0 ext.)�԰����ض�λĿ�����32 λ��ЧVA ��8 λָ�������
Public Const IMAGE_REL_SH3_DIRECT4              As Byte = &H6                       ' 4 bit direct (0 ext.)�����4 λ�����ض�λĿ�����VA ��8 λ��Ԫ������
Public Const IMAGE_REL_SH3_DIRECT4_WORD         As Byte = &H7                       ' 4 bit direct .W (0 ext.)�����4 λ�����ض�λĿ�����16 λ��ЧVA ��8 λָ�������
Public Const IMAGE_REL_SH3_DIRECT4_LONG         As Byte = &H8                       ' 4 bit direct .L (0 ext.)�����4 λ�����ض�λĿ�����32 λ��ЧVA ��8 λָ�������
Public Const IMAGE_REL_SH3_PCREL8_WORD          As Byte = &H9                       ' 8 bit PC relative .W �԰����ض�λĿ�����16 λ��Ч���ƫ�Ƶ�8λָ�������
Public Const IMAGE_REL_SH3_PCREL8_LONG          As Byte = &HA                       ' 8 bit PC relative .L �԰����ض�λĿ�����32 λ��Ч���ƫ�Ƶ�8λָ�������
Public Const IMAGE_REL_SH3_PCREL12_WORD         As Byte = &HB                       ' 12 LSB PC relative .W �����12 λ�����ض�λĿ�����16 λ��Ч���ƫ�Ƶ�16 λָ�������
Public Const IMAGE_REL_SH3_STARTOF_SECTION      As Byte = &HC                       ' Start of EXE section �԰����ض�λĿ��������ڽ�VA ��32 λ��Ԫ������
Public Const IMAGE_REL_SH3_SIZEOF_SECTION       As Byte = &HD                       ' Size of EXE section �԰����ض�λĿ��������ڽڴ�С��32 λ��Ԫ������
Public Const IMAGE_REL_SH3_SECTION              As Byte = &HE                       ' Section table index �����ض�λĿ��Ľڵ�16 λ����������֧�ֵ�����Ϣ
Public Const IMAGE_REL_SH3_SECREL               As Byte = &HF                       ' Offset within section �ض�λĿ������������ڽڿ�ͷ��32 λƫ�ơ�����֧�ֵ�����Ϣ�;�̬�ֲ߳̾��洢
Public Const IMAGE_REL_SH3_DIRECT32_NB          As Byte = &H10                      ' 32 bit direct not based �ض�λĿ����ŵ�32 λRVA��
Public Const IMAGE_REL_SH3_GPREL4_LONG          As Byte = &H11                      ' GP-relative addressing    ��GP ��ء�
Public Const IMAGE_REL_SH3_TOKEN                As Byte = &H12                      ' clr token     CLR �Ǻš�
Public Const IMAGE_REL_SHM_PCRELPT              As Byte = &H13                      ' Offset from current �൱ǰָ���ƫ�ƣ����֣������û������IMAGE_REL_SHM_NOMODE ��־����ô����λȡ�����뵽��32 λ��ѡ��PTA ָ���PTB ָ�
                                                                                    '  instruction in longwords
                                                                                    '  if not NOMODE, insert the
                                                                                    '  inverse of the low bit at
                                                                                    '  bit 32 to select PTA/PTB
Public Const IMAGE_REL_SHM_REFLO                As Byte = &H14                      ' Low bits of 32-bit address  32 λ��ַ�ĵ�16 λ��
Public Const IMAGE_REL_SHM_REFHALF              As Byte = &H15                      ' High bits of 32-bit address 32 λ��ַ�ĸ�16 λ��
Public Const IMAGE_REL_SHM_RELLO                As Byte = &H16                      ' Low bits of relative reference ��Ե�ַ�ĵ�16 λ��
Public Const IMAGE_REL_SHM_RELHALF              As Byte = &H17                      ' High bits of relative reference ��Ե�ַ�ĸ�16 λ��
Public Const IMAGE_REL_SHM_PAIR                 As Byte = &H18                      ' offset operand for relocationֻ�н���IMAGE_REL_SHM_REFHALF��IMAGE_REL_SHM_RELLO ��IMAGE_REL_SHM_RELHALF ���͵��ض�λ��ʱ�����ض�λ���Ͳ��ǺϷ���.���ض�λ���SymbolTableIndex ���������ƫ�ƶ����Ƿ��ű������

Public Const IMAGE_REL_SH_NOMODE                As Integer = &H8000                 ' relocation ignores section mode �ض�λ���Խ�ģʽ��


Public Const IMAGE_REL_ARM_ABSOLUTE             As Byte = &H0                       ' No relocation required �ض�λ�����ԡ�
Public Const IMAGE_REL_ARM_ADDR32               As Byte = &H1                       ' 32 bit address �ض�λĿ���32 λVA��
Public Const IMAGE_REL_ARM_ADDR32NB             As Byte = &H2                       ' 32 bit address w/o image base �ض�λĿ���32 λRVA��
Public Const IMAGE_REL_ARM_BRANCH24             As Byte = &H3                       ' 24 bit offset << 2 & sign ext. �ض�λĿ���24 λ���ƫ�ơ�
Public Const IMAGE_REL_ARM_BRANCH11             As Byte = &H4                       ' Thumb: 2 11 bit offsets ���ӳ�����õ����á��������������16 λָ��
Public Const IMAGE_REL_ARM_TOKEN                As Byte = &H5                       ' clr token
Public Const IMAGE_REL_ARM_GPREL12              As Byte = &H6                       ' GP-relative addressing (ARM)
Public Const IMAGE_REL_ARM_GPREL7               As Byte = &H7                       ' GP-relative addressing (Thumb)
Public Const IMAGE_REL_ARM_BLX24                As Byte = &H8
Public Const IMAGE_REL_ARM_BLX11                As Byte = &H9
Public Const IMAGE_REL_ARM_SECTION              As Byte = &HE                       ' Section table index �����ض�λĿ��Ľڵ�16 λ����������֧�ֵ�����Ϣ
Public Const IMAGE_REL_ARM_SECREL               As Byte = &HF                       ' Offset within section�ض�λĿ������������ڽڿ�ͷ��32 λƫ�ơ�����֧�ֵ�����Ϣ�;�̬�ֲ߳̾��洢
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
Public Const IMAGE_REL_AMD64_ABSOLUTE           As Byte = &H0                       ' Reference is absolute, no relocation is necessary �ض�λ�����ԡ�
Public Const IMAGE_REL_AMD64_ADDR64             As Byte = &H1                       ' 64-bit address (VA). �ض�λĿ���64 λVA��
Public Const IMAGE_REL_AMD64_ADDR32             As Byte = &H2                       ' 32-bit address (VA). �ض�λĿ���32 λVA��
Public Const IMAGE_REL_AMD64_ADDR32NB           As Byte = &H3                       ' 32-bit address w/o image base (RVA). ������ӳ���ַ��32 λ��ַ��RVA����
Public Const IMAGE_REL_AMD64_REL32              As Byte = &H4                       ' 32-bit relative address from byte following reloc ������ض�λĿ���32 λ��ַ��
Public Const IMAGE_REL_AMD64_REL32_1            As Byte = &H5                       ' 32-bit relative address from byte distance 1 from reloc ����ھ��ض�λĿ��1 �ֽڴ���32 λ��ַ��
Public Const IMAGE_REL_AMD64_REL32_2            As Byte = &H6                       ' 32-bit relative address from byte distance 2 from reloc ����ھ��ض�λĿ��2 �ֽڴ���32 λ��ַ��
Public Const IMAGE_REL_AMD64_REL32_3            As Byte = &H7                       ' 32-bit relative address from byte distance 3 from reloc ����ھ��ض�λĿ��3 �ֽڴ���32 λ��ַ��
Public Const IMAGE_REL_AMD64_REL32_4            As Byte = &H8                       ' 32-bit relative address from byte distance 4 from reloc ����ھ��ض�λĿ��4 �ֽڴ���32 λ��ַ��
Public Const IMAGE_REL_AMD64_REL32_5            As Byte = &H9                       ' 32-bit relative address from byte distance 5 from reloc ����ھ��ض�λĿ��5 �ֽڴ���32 λ��ַ��
Public Const IMAGE_REL_AMD64_SECTION            As Byte = &HA                       ' Section index  �����ض�λĿ��Ľڵ�16 λ����������֧�ֵ�����Ϣ
Public Const IMAGE_REL_AMD64_SECREL             As Byte = &HB                       ' 32 bit offset from base of section containing target �ض�λĿ������������ڽڿ�ͷ��32 λƫ�ơ�����֧�ֵ�����Ϣ�;�̬�ֲ߳̾��洢?
Public Const IMAGE_REL_AMD64_SECREL7            As Byte = &HC                       ' 7 bit unsigned offset from base of section containing target ������ض�λĿ�����ڽڻ���ַ��7 λƫ�ƣ��޷���������
Public Const IMAGE_REL_AMD64_TOKEN              As Byte = &HD                       ' 32 bit metadata token    CLR �Ǻš�
Public Const IMAGE_REL_AMD64_SREL32             As Byte = &HE                       ' 32 bit signed span-dependent value emitted into object ����Ŀ���ļ��е�32 λ�������ֵ����������
Public Const IMAGE_REL_AMD64_PAIR               As Byte = &HF                       '��������ֵ�ɶԳ��֣����������ÿһ���������ֵ
Public Const IMAGE_REL_AMD64_SSPAN32            As Byte = &H10                      ' 32 bit signed span-dependent value applied at link time ����ʱӦ�õ�32 λ�������ֵ������������

' IA64 relocation types.
Public Const IMAGE_REL_IA64_ABSOLUTE            As Byte = &H0                       '�ض�λ�����ԡ�
Public Const IMAGE_REL_IA64_IMM14               As Byte = &H1                       '����ָ���ض�λ������Ը���IMAGE_REL_IA64_ADDEND ���͵��ض�λ������ߵ�ֵ�ڱ����뵽IMM14 ָ�����ָ����ָ�����֮ǰ���ӵ�Ŀ���ַ��.�����ض�λĿ������Ǿ��Է��ţ��������ӳ����뱻������
Public Const IMAGE_REL_IA64_IMM22               As Byte = &H2                       '����ָ���ض�λ������Ը���IMAGE_REL_IA64_ADDEND ���͵��ض�λ������ߵ�ֵ�ڱ����뵽IMM22 ָ�����ָ����ָ�����֮ǰ���ӵ�Ŀ���ַ��.�����ض�λĿ������Ǿ��Է��ţ��������ӳ����뱻������
Public Const IMAGE_REL_IA64_IMM64               As Byte = &H3                       '�����ض�λ���ָ��۱�ű���Ϊ1�������ض�λ������Ը���IMAGE_REL_IA64_ADDEND ���͵��ض�λ������ߵ�ֵ�ڱ��洢��IMM64 ָ���������ָ�����֮ǰ���ӵ�Ŀ���ַ��
Public Const IMAGE_REL_IA64_DIR32               As Byte = &H4                       '�ض�λĿ���32 λVA����֧��ʹ��/LARGEADDRESSAWARE:NO ������ѡ�����ɵ�ӳ��
Public Const IMAGE_REL_IA64_DIR64               As Byte = &H5                       '�ض�λĿ���64 λVA��
Public Const IMAGE_REL_IA64_PCREL21B            As Byte = &H6                       'ʹ�ð�16 λ�߽������ض�λĿ���25 λ���ƫ��������ָ����ƫ�Ƶĵ�4 λȫΪ0����˲�û�б��洢
Public Const IMAGE_REL_IA64_PCREL21M            As Byte = &H7                       'ʹ�ð�16 λ�߽������ض�λĿ���25 λ���ƫ��������ָ����ƫ�Ƶĵ�4 λȫΪ0����˲�û�б��洢
Public Const IMAGE_REL_IA64_PCREL21F            As Byte = &H8                       '�����ض�λĿ��ƫ�Ƶ�LSB ���ְ�������ָ��۱�ţ����ಿ�ְ�������ָ����ĵ�ַ��ʹ�ð�16 λ�߽������ض�λĿ���25 λ���ƫ��������ָ����ƫ�Ƶĵ�4 λȫΪ0����˲�û�б��洢
Public Const IMAGE_REL_IA64_GPREL22             As Byte = &H9                       '����ָ���ض�λ������Ը���IMAGE_REL_IA64_ADDEND ���͵��ض�λ����ߵ�ֵ���ӵ�Ŀ���ַ�ϣ��������GPREL22 ָ��������GP ��ƫ�Ʋ�Ӧ��
Public Const IMAGE_REL_IA64_LTOFF22             As Byte = &HA                       'ʹ���ض�λĿ����ŵĳ������������GP ��22λƫ��������ָ��?��������������ض�λ���Լ����ܸ�������IMAGE_REL_IA64_ADDEND ���͵��ض�λ�������������������
Public Const IMAGE_REL_IA64_SECTION             As Byte = &HB                       '�����ض�λĿ��Ľڵ�16 λ����������֧�ֵ�����Ϣ
Public Const IMAGE_REL_IA64_SECREL22            As Byte = &HC                       'ʹ���ض�λĿ������������ڽڿ�ͷ��22 λƫ��������ָ��.�������͵��ض�λ�������Խ�����IMAGE_REL_IA64_ADDEND ���͵��ض�λ����ߵ�Value ������ض�λĿ������������ڽڿ�ͷ��32 λƫ�ƣ��޷���������
Public Const IMAGE_REL_IA64_SECREL64I           As Byte = &HD                       '�����ض�λ���ָ��۱�ű���Ϊ1��ʹ���ض�λĿ������������ڽڿ�ͷ��64 λƫ��������ָ��.�������͵��ض�λ�������Խ�����IMAGE_REL_IA64_ADDEND ���͵��ض�λ����ߵ�Value ������ض�λĿ������������ڽڿ�ͷ��32 λƫ�ƣ��޷���������
Public Const IMAGE_REL_IA64_SECREL32            As Byte = &HE                       'ʹ���ض�λĿ������������ڽڿ�ͷ��32 λƫ�������������ݵĵ�ַ

Public Const IMAGE_REL_IA64_DIR32NB             As Byte = &H10                      'Ŀ���32 λRVA��
Public Const IMAGE_REL_IA64_SREL14              As Byte = &H11                      '���ڰ��������ض�λĿ��֮���14 λ������������������������������˵����һ��˵���򣬱����������Ѿ����������ֵ
Public Const IMAGE_REL_IA64_SREL22              As Byte = &H12                      '���ڰ��������ض�λĿ��֮���22 λ������������������������������˵����һ��˵���򣬱����������Ѿ����������ֵ
Public Const IMAGE_REL_IA64_SREL32              As Byte = &H13                      '���ڰ��������ض�λĿ��֮���32 λ������������������������������˵����һ��˵���򣬱����������Ѿ����������ֵ
Public Const IMAGE_REL_IA64_UREL32              As Byte = &H14                      '���ڰ��������ض�λĿ��֮���32 λ���������޷���������������������˵����һ��˵���򣬱����������Ѿ����������ֵ
Public Const IMAGE_REL_IA64_PCREL60X            As Byte = &H15                      ' This is always a BRL and never converted �����PC ��60 λ����������MLX ָ�����BRLָ��
Public Const IMAGE_REL_IA64_PCREL60B            As Byte = &H16                      ' If possible, convert to MBB bundle with NOP.B in slot 1 �����PC ��60 λ����������ض�λĿ��ƫ�Ʋ�����һ��25 λ�����ܱ�ʾ�ķ�Χ��������������ô����1 ��ָ�����ʹ��NOP.B ָ�2 ��ָ�����ʹ��25 λ�����4 λȫΪ0����������BR ָ�����ָ���ת����MBB ָ���
Public Const IMAGE_REL_IA64_PCREL60F            As Byte = &H17                      ' If possible, convert to MFB bundle with NOP.F in slot 1 �����PC ��60 λ����������ض�λĿ��ƫ�Ʋ�����һ��25 λ�����ܱ�ʾ�ķ�Χ��������������ô����1 ��ָ�����ʹ��NOP.F ָ�2 ��ָ�����ʹ��25 λ�����4 λȫΪ0����������BR ָ�����ָ���ת����MFB ָ���
Public Const IMAGE_REL_IA64_PCREL60I            As Byte = &H18                      ' If possible, convert to MIB bundle with NOP.I in slot 1 �����PC ��60 λ����������ض�λĿ��ƫ�Ʋ�����һ��25 λ�����ܱ�ʾ�ķ�Χ��������������ô����1 ��ָ�����ʹ��NOP.I ָ�2 ��ָ�����ʹ��25 λ�����4 λȫΪ0����������BR ָ�����ָ���ת����MIB ָ���
Public Const IMAGE_REL_IA64_PCREL60M            As Byte = &H19                      ' If possible, convert to MMB bundle with NOP.M in slot 1 �����PC ��60 λ����������ض�λĿ��ƫ�Ʋ�����һ��25 λ�����ܱ�ʾ�ķ�Χ��������������ô����1 ��ָ�����ʹ��NOP.M ָ�2 ��ָ�����ʹ��25 λ�����4 λȫΪ0����������BR ָ�����ָ���ת����MMB ָ���
Public Const IMAGE_REL_IA64_IMMGPREL64          As Byte = &H1A                      '�����GP ��64 λ������
Public Const IMAGE_REL_IA64_TOKEN               As Byte = &H1B                      ' clr token CLR �Ǻš�
Public Const IMAGE_REL_IA64_GPREL32             As Byte = &H1C                      '�����GP ��32 λ������
Public Const IMAGE_REL_IA64_ADDEND              As Byte = &H1F                      'ֻ�н����������͵��ض�λʱ�����ض�λ���Ͳ��ǺϷ���: IMAGE_REL_IA64_IMM14?IMAGE_REL_IA64_IMM22 IMAGE_REL_IA64_IMM64 IMAGE_REL_IA64_GPREL22 IMAGE_REL_IA64_LTOFF22 IMAGE_REL_IA64_LTOFF64 IMAGE_REL_IA64_SECREL22 IMAGE_REL_IA64_SECREL64I ��IMAGE_REL_IA64_SECREL32.����ֵ��Ӧ�õ�ָ����е�ָ���ϵļ������������������ݡ�

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


Public Const IMAGE_REL_M32R_ABSOLUTE            As Byte = &H0                       ' No relocation required  �ض�λ�����ԡ�
Public Const IMAGE_REL_M32R_ADDR32              As Byte = &H1                       ' 32 bit address �ض�λĿ���32 λVA��
Public Const IMAGE_REL_M32R_ADDR32NB            As Byte = &H2                       ' 32 bit address w/o image base �ض�λĿ���32 λRVA��
Public Const IMAGE_REL_M32R_ADDR24              As Byte = &H3                       ' 24 bit address �ض�λĿ���24 λVA��
Public Const IMAGE_REL_M32R_GPREL16             As Byte = &H4                       ' GP relative addressing �ض�λĿ�������GP �Ĵ�����16 λƫ�ơ�
Public Const IMAGE_REL_M32R_PCREL24             As Byte = &H5                       ' 24 bit offset << 2 & sign ext. �ض�λĿ������ڳ����������PC����24 λƫ�ƣ��Ѿ�����2 λ����������չ��
Public Const IMAGE_REL_M32R_PCREL16             As Byte = &H6                       ' 16 bit offset << 2 & sign ext. �ض�λĿ�������PC ��16 λƫ�ƣ��Ѿ�����2λ����������չ
Public Const IMAGE_REL_M32R_PCREL8              As Byte = &H7                       ' 8 bit offset << 2 & sign ext.  �ض�λĿ�������PC ��8 λƫ�ƣ��Ѿ�����2 λ����������չ
Public Const IMAGE_REL_M32R_REFHALF             As Byte = &H8                       ' 16 MSBs �ض�λĿ��VA ��16 λMSB��
Public Const IMAGE_REL_M32R_REFHI               As Byte = &H9                       ' 16 MSBs adj for LSB sign ext. �ض�λĿ��VA ��16 λMSB���Ѿ���LSB ������չ����.�����ڼ���һ��������32 λ��ַ�������ָ�������еĵ�һ��ָ��.�����ض�λ���ͺ���������IMAGE_REL_M32R_PAIR ���͵��ض�λ������ߵ�SymbolTableIndex ���������һ��16 λƫ�ƣ��������������ƫ��Ҫ���ӵ��ض�λ����λ�õĸ�16 λ
Public Const IMAGE_REL_M32R_REFLO               As Byte = &HA                       ' 16 LSBs �ض�λĿ��VA ��16 λLSB��
Public Const IMAGE_REL_M32R_PAIR                As Byte = &HB                       ' Link HI and LO �������͵��ض�λ�����������ΪIMAGE_REL_M32R_REFHI ���ض�λ��.���ض�λ���SymbolTableIndex ���������ƫ�ƶ����Ƿ��ű�����
Public Const IMAGE_REL_M32R_SECTION             As Byte = &HC                       ' Section table index �����ض�λĿ��Ľڵ�16 λ����������֧�ֵ�����Ϣ
Public Const IMAGE_REL_M32R_SECREL32            As Byte = &HD                       ' 32 bit section relative reference�ض�λĿ������������ڽڿ�ͷ��32 λƫ�ơ�����֧�ֵ�����Ϣ�;�̬�ֲ߳̾��洢
Public Const IMAGE_REL_M32R_TOKEN               As Byte = &HE                       ' clr token CLR �Ǻš�

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
'����Ŀ¼��:һ��ֻ��һ�еı������Ŀ¼��ͬ�������������������ֵ������λ�úʹ�С
Public Type IMAGE_EXPORT_DIRECTORY
    Characteristics     As Long                 '����������Ϊ0��
    TimeDateStamp       As Long                 '�������ݱ����������ں�ʱ�䡣
    MajorVersion        As Integer              '���汾�š��û����������������汾�źʹΰ汾�š�
    MinorVersion        As Integer              '�ΰ汾�š�
    Name                As Long                 '�������DLL ���Ƶ�ASCII ���ַ��������ӳ���ַ��ƫ�Ƶ�ַ
    Base                As Long                 'ӳ���е������ŵ���ʼ����ֵ�������ָ���˵�����ַ�����ʼ����ֵ.��ͨ��������Ϊ1
    NumberOfFunctions   As Long                 '������ַ����Ԫ�ص���Ŀ��
    NumberOfNames       As Long                 '��������ָ�����Ԫ�ص���Ŀ����ͬʱҲ�ǵ�����������Ԫ�ص���Ŀ
    AddressOfFunctions  As Long                 'RVA from base of image,������ַ�������ӳ���ַ��ƫ�Ƶ�ַ��
    AddressOfNames      As Long                 ' RVA from base of image,��������ָ��������ӳ���ַ��ƫ�Ƶ�ַ�����Ĵ�С��Number of Name Pointers �������
    AddressOfNameOrdinals As Long               ' RVA from base of image,���������������ӳ���ַ��ƫ�Ƶ�ַ��
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
    ForwarderString     As Long                 'u1_ForwarderString ַ
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
'    Characteristics     As Long                 'DUMMYUNIONNAME_Characteristics 0 for terminating null import descriptor ������ұ��RVA������������ÿһ��������ŵ����ƻ�����
    OriginalFirstThunk  As Long                 'DUMMYUNIONNAME_OriginalFirstThunk  RVA to original unbound IAT (PIMAGE_THUNK_DATA)
    TimeDateStamp       As Long                 ' 0 if not bound,
                                                ' -1 if bound, and real date\time stamp
                                                '   in IMAGE_DIRECTORY_ENTRY_BOUND_IMPORT (new BIND) �����һֱ������Ϊ0��ֱ��ӳ�񱻰󶨡���ӳ�񱻰�֮�����������Ϊ���DLL ������/ʱ�����
                                                'O.W. date/time stamp of DLL bound to (Old BIND)
    ForwarderChain      As Long                 '-1 if no forwarders ��һ��ת�����������
    Name                As Long                 ' ����DLL ���Ƶ�ASCII ���ַ��������ӳ���ַ��ƫ�Ƶ�ַ
    FirstThunk          As Long                 'RVA to IAT (if bound this IAT has actual addresses) �����ַ���RVA�������������뵼����ұ��������ȫһ����ֱ��ӳ�񱻰󶨡�
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
    Characteristics     As Long                 '������Ϊ��Դ�����ԣ�������ʵ������0
    TimeDateStamp       As Long                 '��Դ�Ĳ���ʱ��
    MajorVersion        As Integer              '������Ϊ��Դ�İ汾��������ʵ������0
    MinorVersion        As Integer
    NumberOfNamedEntries    As Integer          '�������������������
    NumberOfIdEntries   As Integer              '��ID�������������
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
    Name                As Long                 'DUMMYSTRUCTNAME_Name,Ŀ¼��������ַ���ָ���ID
'    NameOffset          As Long                 'DUMMYSTRUCTNAME_RvaBased.Bits31
'    NameIsString        As Long                 'DUMMYSTRUCTNAME_RvaBased.Bits1
'    Id                  As Integer              'DUMMYSTRUCTNAME_Id
    'DUMMYUNIONNAME2
    OffsetToData        As Long                 'DUMMYSTRUCTNAME2_OffsetToData,Ŀ¼��ָ��
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
    Length              As Integer              '�ַ����ĳ���
    NameString          As Integer              'UNICODE�ַ����������ַ����ǲ������ģ���������ֻ����һ��dw��ʾ��ʵ���ϵ�����Ϊ100��ʱ�������������NameString dw 100 dup (?)
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
    Characteristics     As Long                 '����������Ϊ0��
    TimeDateStamp       As Long                 '�������ݱ����������ں�ʱ�䡣
    MajorVersion        As Integer              '�������ݸ�ʽ�����汾�š�
    MinorVersion        As Integer              '�������ݸ�ʽ�Ĵΰ汾�š�
    Type                As Long                 '������Ϣ�ĸ�ʽ�������Ĵ���ʹ�ÿ���֧�ֶ��������
    SizeOfData          As Long                 '�������ݣ�����������Ŀ¼�����Ĵ�С��
    AddressOfRawData    As Long                 '��������ʱ�������������ӳ���ַ��ƫ�Ƶ�ַ��
    PointerToRawData    As Long                 'ָ��������ݵ��ļ�ָ�롣
End Type

Public Const IMAGE_DEBUG_TYPE_UNKNOWN           As Long = 0                         'δֵ֪�����й��߾����Դ�ֵ��
Public Const IMAGE_DEBUG_TYPE_COFF              As Long = 1                         'COFF ������Ϣ���к���Ϣ�����ű���ַ��������ļ�ͷ��Ҳ�������ָ���������͵ĵ�����Ϣ
Public Const IMAGE_DEBUG_TYPE_CODEVIEW          As Long = 2                         'Visual C++������Ϣ��
Public Const IMAGE_DEBUG_TYPE_FPO               As Long = 3                         'ָ֡��ʡ�ԣ�FPO����Ϣ��������Ϣ���ߵ�������ν��ͷǱ�׼ջ֡������֡��EBP �Ĵ�����������Ŀ�Ķ�������Ϊָ֡��
Public Const IMAGE_DEBUG_TYPE_MISC              As Long = 4                         'DBG �ļ���λ�á�
Public Const IMAGE_DEBUG_TYPE_EXCEPTION         As Long = 5                         '.pdata �ڵĸ�����
Public Const IMAGE_DEBUG_TYPE_FIXUP             As Long = 6                         '������
Public Const IMAGE_DEBUG_TYPE_OMAP_TO_SRC       As Long = 7                         '�Ӿ����������ź��ӳ���е�RVA ��ԭӳ���е�RVA ��ӳ��
Public Const IMAGE_DEBUG_TYPE_OMAP_FROM_SRC     As Long = 8                         '��ԭӳ���е�RVA �������������ź��ӳ���е�RVA ��ӳ��
Public Const IMAGE_DEBUG_TYPE_BORLAND           As Long = 9                         '��������Borland ��˾ʹ�á�
Public Const IMAGE_DEBUG_TYPE_RESERVED10        As Long = 10                        '����
Public Const IMAGE_DEBUG_TYPE_CLSID             As Long = 11                        '����

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
    ulOffStart          As Long                 'offset 1st byte of function code  ���������һ���ֽڵ�ƫ��
    cbProcSize          As Long                 '# bytes in function  ����������ռ���ֽ���
    cdwLocals           As Long                 '# bytes in locals/4  �ֲ�������ռ�ֽ�������4
    cdwParams           As Integer              '# bytes in params/4  ������ռ�ֽ�������4
    cbProlog            As Integer              'WORD_bits8# bytes in prolog  ����prolog ������ռ�ֽ���
'    cbRegs              As Integer              'WORD_bits3# regs saved    ����ļĴ�����
'    fHasSEH             As Integer              'WORD_bits1 TRUE if SEH in func  �����������SEH����ֵΪTRUE
'    fUseBP              As Integer              'WORD_bits1 TRUE if EBP has been allocated  ���EBP �Ĵ����Ѿ������䣬��ֵΪTRUE
'    reserved            As Integer              'WORD_bits1 reserved for future use  ����������ʹ��
'    cbFrame             As Integer              'WORD_bits2 frame type֡����
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
    Cb                  As Long                 'ͷ���ֽڴ�С
    MajorRuntimeVersion As Integer              'CLR��Ҫ���е���С�汾���屾��
    MinorRuntimeVersion As Integer              'CLR��Ҫ���е���С�汾�ΰ汾��

    'Symbol table and startup information
    MetaData            As IMAGE_DATA_DIRECTORY 'Rav��Ԫ���ݵĴ�С
    Flags               As Long                 '�����Ʊ��

    'If COMIMAGE_FLAGS_NATIVE_ENTRYPOINT is not set, EntryPointToken represents a managed entrypoint.
    'If COMIMAGE_FLAGS_NATIVE_ENTRYPOINT is set, EntryPointRVA represents an RVA to a native entrypoint.
    'DUMMYUNIONNAME
    EntryPointToken     As Long                 'DUMMYUNIONNAME_EntryPointToken
'    EntryPointRVA       As Long                 'DUMMYUNIONNAME_EntryPointRVAToken

    'Binding information
    Resources           As IMAGE_DATA_DIRECTORY 'Rav���й���Դ�Ĵ�С
    StrongNameSignature As IMAGE_DATA_DIRECTORY 'Rav���������pe�ļ��Ĺ�ϣ���ݵĴ�С

    'Regular fixup and binding information
    CodeManagerTable    As IMAGE_DATA_DIRECTORY 'Rva�ʹ�������Ĵ�С
    VTableFixups        As IMAGE_DATA_DIRECTORY 'Rav��һ���������������ɵ�������ֽڴ�С
    ExportAddressTableJumps As IMAGE_DATA_DIRECTORY 'Rav����jump thunk�ĵ�ַ��ɵ�����Ĵ�С
    'Precompiled image info (internal use only - set to zero)
    ManagedNativeHeader As IMAGE_DATA_DIRECTORY 'ΪԤ����������ģ�������Ϊ0
End Type

'.net Meta Data Structor
Public Type CLR_MetaDataVer
    Signature                                   As Long
    MajorVersion                                As Integer
    MinorVersion                                As Integer
    ExtraData                                   As Long
    Length                                      As Long
    VersionString(15)                           As Byte 'array[0..IMAGE_NUMBEROF_DIRECTORY_ENTRIES-1] of Char;   //.net�ַ���
    Flags                                       As Byte
    Pading                                      As Byte
    Streams                                     As Integer
End Type

