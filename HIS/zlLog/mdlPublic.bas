Attribute VB_Name = "mdlPublic"
Option Explicit


Private Const CP_UTF8                       As Long = 65001
Private Const WAIT_ABANDONED                As Long = &H80          'ָ���Ķ�����һ�����������������������ӵ���߳���ֹ֮ǰӵ�л��������߳����ͷŵġ�
                                                                    '������������Ȩ������������̣߳���������״̬������Ϊ���źš�
                                                                    '������������ڱ����־�״̬��Ϣ����ô��Ӧ�ü�����Ƿ����һ���ԡ�
Private Const WAIT_OBJECT_0                 As Long = &H0           'ָ�������״̬�����źŵ�?
Private Const WAIT_TIMEOUT                  As Long = &H102         '��ʱʱ�����������״̬��û���źŵġ�
Private Const WAIT_FAILED                   As Long = &HFFFFFFFF    '�������ʧ���ˡ�Ҫ�����չ������Ϣ�������GetLastError
Private Const MAX_PATH                      As Long = 260
Private Const TOKEN_QUERY                   As Long = &H8
Private Const PROCESS_QUERY_INFORMATION     As Long = &H400
Private Const PROCESS_VM_READ               As Long = &H10
Private Const WTS_CURRENT_SERVER_HANDLE     As Long = 0
Private Const WTS_CURRENT_SESSION           As Long = -1
Private Const IMAGE_SIZEOF_SHORT_NAME            As Integer = 8
Private Const IMAGE_NUMBEROF_DIRECTORY_ENTRIES   As Integer = 16

Public Const SM4_CRYPT_RANDOMIZE_KEY            As Long = 999                       'sm4�����㷨��Կ���������������
Public Const SM4_CRYPT_RANDOMIZE_IV             As Long = 666                       'sm4�����㷨��ʼ�������������������
Public Const IMAGE_DOS_SIGNATURE                As Integer = &H5A4D                 'MZ
Public Const IMAGE_OS2_SIGNATURE                As Integer = &H454E                 'NE
Public Const IMAGE_OS2_SIGNATURE_LE             As Integer = &H454C                 'LE
Public Const IMAGE_NT_SIGNATURE                 As Long = &H4550                    'PE00
Public Const IMAGE_NT_OPTIONAL_HDR32_MAGIC      As Integer = &H10B                  '����һ��32λ�����ļ�
Public Const IMAGE_NT_OPTIONAL_HDR64_MAGIC      As Integer = &H20B                  '����һ��PE32+��ִ���ļ�
Public Const IMAGE_ROM_OPTIONAL_HDR_MAGIC       As Integer = &H107                  '����һ��ROM����
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

Private Enum CrypeMode
    CM_Encrypt = 1   '����
    CM_Decrypt = 0   '����
End Enum

Public Type IMAGE_RESOURCE_DIR_STRING_U
    Length              As Integer              '�ַ����ĳ���
    NameString          As Integer              'UNICODE�ַ����������ַ����ǲ������ģ���������ֻ����һ��dw��ʾ��ʵ���ϵ�����Ϊ100��ʱ�������������NameString dw 100 dup (?)
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

Public Type IMAGE_DATA_DIRECTORY
    VirtualAddress      As Long                 '���ݿ��RVA
    Size                As Long                 '���ݿ��С
End Type

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

Public Type IMAGE_RESOURCE_DIRECTORY
    Characteristics     As Long                 '������Ϊ��Դ�����ԣ�������ʵ������0
    TimeDateStamp       As Long                 '��Դ�Ĳ���ʱ��
    MajorVersion        As Integer              '������Ϊ��Դ�İ汾��������ʵ������0
    MinorVersion        As Integer
    NumberOfNamedEntries    As Integer          '�������������������
    NumberOfIdEntries   As Integer              '��ID�������������
'    IMAGE_RESOURCE_DIRECTORY_ENTRY DirectoryEntries[];
End Type

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
    REG_SZ = 1 'Unicode���ս��ַ���
    REG_EXPAND_SZ = 2 'Unicode���ս��ַ���
    REG_BINARY = 3 '��������ֵ
    REG_DWORD = 4 '32-bit ����
    REG_DWORD_BIG_ENDIAN = 5
    REG_LINK = 6
    REG_MULTI_SZ = 7 ' ��������ֵ��
End Enum

Private Enum REGRoot
    HKEY_CLASSES_ROOT = &H80000000 '��¼Windows����ϵͳ�����������ļ��ĸ�ʽ�͹�����Ϣ����Ҫ��¼��ͬ�ļ����ļ�����׺����֮��Ӧ��Ӧ�ó��������Ӽ��ɷ�Ϊ���࣬һ�����Ѿ�ע��ĸ����ļ�����չ���������Ӽ�ǰ�涼��һ������������һ���Ǹ����ļ������й���Ϣ��
    HKEY_CURRENT_USER = &H80000001 '�˸��������˵�ǰ��¼�û����û������ļ���Ϣ����Щ��Ϣ��֤��ͬ���û���¼�����ʱ��ʹ���Լ��ĸ��Ի����ã������Լ������ǽֽ���Լ����ռ��䡢�Լ��İ�ȫ����Ȩ�޵ȡ�
    HKEY_LOCaL_MaCHINE = &H80000002 '�˸��������˵�ǰ��������������ݣ���������װ��Ӳ���Լ���������á���Щ��Ϣ��Ϊ���е��û���¼ϵͳ����ġ���������ע��������Ӵ�Ҳ������Ҫ�ĸ�����
    HKEY_USERS = &H80000003 '�˸�������Ĭ���û�����Ϣ��Default�Ӽ�����������ǰ��¼�û�����Ϣ��
    HKEY_PERFORMANCE_DATA = &H80000004 '��Windows NT/2000/XPע�������Ȼû��HKEY_DYN_DATA����������ȴ������һ����Ϊ��HKEY_ PERFOR MANCE_DATA����������ϵͳ�еĶ�̬��Ϣ���Ǵ���ڴ��Ӽ��С�ϵͳ�Դ���ע���༭���޷������˼�
    HKEY_CURRENT_CONFIG = &H80000005  '�˸���ʵ������HKEY_LOCAL_MACHINE�е�һ���֣����д�ŵ��Ǽ������ǰ���ã�����ʾ������ӡ���������������Ϣ�ȡ������Ӽ���HKEY_LOCAL_ MACHINE\ Config\0001��֧�µ�������ȫһ����
    HKEY_DYN_DATA = &H80000006 '�˸����б���ÿ��ϵͳ����ʱ��������ϵͳ���ú͵�ǰ������Ϣ���������ֻ������Windows 98�С�
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
      'һ����null��β���ַ����������û���¼ʱԶ������������еĳ�ʼ��������ơ�
      WTSApplicationName = 1
      'һ����null��β���ַ����������Ự�������е�Ӧ�ó�����ѷ������ơ�Windows Server 2008 R2��Windows 7��Windows Server 2008��Windows Vista:��֧�ִ�ֵ
      WTSWorkingDirectory = 2
      'һ����null��β���ַ���������������ʼ����ʱʹ�õ�Ĭ��Ŀ¼��
      WTSOEMId = 3
      '��ʹ�ô�ֵ��
      WTSSessionId = 4
      '�����Ự��ʶ����ULONGֵ��
      WTSUserName = 5
      '������Ự�������û�������null��β���ַ�����
      WTSWinStationName = 6
      'һ����null��β���ַ���������Զ���������Ự�����ơ�
        'ע�⣬����ָ���˸����͵����ƣ������������ش���վ���ơ��෴��������Զ���������Ự�����ơ�ÿ��Զ���������Ự����һ������ʽ����վ����������ڽ���ʽ����վΩһ֧�ֵĴ���վ�����ǡ�WinSta0�������ÿ���Ự�����Լ��ġ�WinSta0������վ��������йظ�����Ϣ����μ�����վ��
      WTSDomainName = 7
      'һ����null��β���ַ�����������¼�û�������������ơ�
      WTSConnectState = 8
      '�Ự�ĵ�ǰ����״̬���йظ�����Ϣ����μ�WTS_CONNECTSTATE_CLASS��
      WTSClientBuildNumber = 9
      '�����ͻ��˹����ŵ�ULONGֵ��
      WTSClientName = 10
      '�����ͻ������Ƶ���null��β���ַ�����
      WTSClientDirectory = 11
      'һ����null��β���ַ�����������װ�ͻ�����Ŀ¼��
      WTSClientProductId = 12
      'һ���ض��ڿͻ��˵Ĳ�Ʒ��ʶ����
      WTSClientHardwareId = 13
      '�����ض��ڿͻ�����Ӳ����ʶ����ULONGֵ����ѡ���Ϊ����ʹ�á�WTSQuerySessionInformation���Ƿ���0ֵ��
      WTSClientAddress = 14
      '�ͻ��˵��������ͺ������ַ���йظ�����Ϣ����μ�WTS_CLIENT_ADDRESS��IP��ַ��WTS_CLIENT_ADDRESS�ṹ��address��Ա��ʼƫ�������ֽ�?
      WTSClientDisplay = 15
      '�йؿͻ�����ʾ�ֱ��ʵ���Ϣ���йظ�����Ϣ����μ�WTS_CLIENT_DISPLAY��
      WTSClientProtocolType = 16
      'ָ���ỰЭ��������Ϣ��USHORTֵ����������ֵ֮һ��
        '0 ����̨�Ự?
        '1 ��ֵ��������������;?
        '2 RDPЭ��?
      WTSIdleTime = 17
      '��ֵ����FALSE�����������GetLastError����ȡ��չ�Ĵ�����Ϣ��GetLastError������ERROR_NOT_SUPPORTED��Windows Server 2008��Windows Vista:��ʹ�ô�ֵ��
      WTSLogonTime = 18
      '��ֵ����FALSE�����������GetLastError����ȡ��չ�Ĵ�����Ϣ��GetLastError������ERROR_NOT_SUPPORTED��Windows Server 2008��Windows Vista:��ʹ�ô�ֵ��
      WTSIncomingBytes = 19
      '��ֵ����FALSE�����������GetLastError����ȡ��չ�Ĵ�����Ϣ��GetLastError������ERROR_NOT_SUPPORTED��Windows Server 2008��Windows Vista:��ʹ�ô�ֵ��
      WTSOutgoingBytes = 20
      '��ֵ����FALSE�����������GetLastError����ȡ��չ�Ĵ�����Ϣ��GetLastError������ERROR_NOT_SUPPORTED��Windows Server 2008��Windows Vista:��ʹ�ô�ֵ��
      WTSIncomingFrames = 21
      '��ֵ����FALSE�����������GetLastError����ȡ��չ�Ĵ�����Ϣ��GetLastError������ERROR_NOT_SUPPORTED��Windows Server 2008��Windows Vista:��ʹ�ô�ֵ��
      WTSOutgoingFrames = 22
      '��ֵ����FALSE�����������GetLastError����ȡ��չ�Ĵ�����Ϣ��GetLastError������ERROR_NOT_SUPPORTED��Windows Server 2008��Windows Vista:��ʹ�ô�ֵ��
      WTSClientInfo = 24
      '�й�Զ����������(RDC)�ͻ�������Ϣ���йظ�����Ϣ����μ�WTSCLIENT��
      WTSSessionInfo = 25
      '�й�RD�Ự�����������ϵĿͻ����Ự����Ϣ���йظ�����Ϣ����μ�WTSINFO��
      WTSSessionInfoEx = 26
      '����RD�Ự�����������ϻỰ����չ��Ϣ���йظ�����Ϣ����μ�WTSINFOEX��Windows Server 2008��Windows Vista:��֧�ִ�ֵ��
      WTSConfigInfo = 27
      '�����й�RD�Ự����������������Ϣ��WTSCONFIGINFO�ṹ��Windows Server 2008��Windows Vista:��֧�ִ�ֵ��
      WTSValidationInfo = 28
      '��֧�ִ�ֵ��
      WTSSessionAddressV4 = 29
      '����������Ự��IPv4��ַ��WTS_SESSION_ADDRESS�ṹ������Ựû������IP��ַ��WTSQuerySessionInformation����������ERROR_NOT_SUPPORTED��Windows Server 2008��Windows Vista:��֧�ִ�ֵ��
      WTSIsRemoteSession = 30
      'ȷ����ǰ�Ự�Ƿ�ΪԶ�̻Ự��WTSQuerySessionInformation��������ֵTRUE����ʾ��ǰ�Ự��Զ�̻Ự������ֵFALSE��ʾ��ǰ�Ự�Ǳ��ػỰ�����ֵֻ�����ڱ��ػ��������WTSQuerySessionInformation������hServer�����������WTS_CURRENT_SERVER_HANDLE��Windows Server 2008��Windows Vista:��֧�ִ�ֵ��
End Enum

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

Public Type IMAGE_FILE_HEADER                   '20B
    Machine             As Integer              '��ʶĿ��������͵�����
    NumberOfSections    As Integer              '�ڵ���Ŀ���������˽ڱ�Ĵ�С�����ڱ�������ļ�ͷ
    TimeDateStamp       As Long                 '��UTC ʱ��1970 ��1 ��1 ��00:00 �����������һ��C ����ʱtime_t ���͵�ֵ���ĵ�32 λ����ָ���ļ���ʱ������
    PointerToSymbolTable    As Long             'COFF ���ű���ļ�ƫ�ơ����������COFF ���ű���ֵΪ0������ӳ���ļ���˵����ֵӦ��Ϊ0����Ϊ�Ѿ����޳�ʹ��COFF ������Ϣ��
    NumberOfSymbols     As Long                 '���ű��е�Ԫ����Ŀ�������ַ�����������ű����Կ����������ֵ����λ�ַ�����?����ӳ���ļ���˵����ֵӦ��Ϊ0����Ϊ�Ѿ����޳�ʹ��COFF������Ϣ��
    SizeOfOptionalHeader    As Integer          '��ѡ�ļ�ͷ�Ĵ�С����ִ���ļ���Ҫ��ѡ�ļ�ͷ��Ŀ���ļ�������Ҫ������Ŀ���ļ���˵����ֵӦ��Ϊ0
    Characteristics     As Integer              'ָʾ�ļ����Եı�־��
End Type

Private Type Big_Iint
    Low                 As Long
    High                As Long
End Type

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

'SM4����
'/**
' * \brief          SM4-ECB block encryption/decryption
' * \param mode     SM4_ENCRYPT or SM4_DECRYPT
' * \param length   length of the input data
' * \param input    input block
' * \param output   output block
' */
Private Declare Function sm4_crypt_ecb Lib "zlSm4.dll" (ByVal Mode As Long, ByVal Length As Long, key As Byte, in_put As Byte, out_put As Byte) As Long
'SM4�����������
'/**
' * \brief          SM4-CBC buffer encryption/decryption
' * \param mode     SM4_ENCRYPT or SM4_DECRYPT
' * \param length   length of the input data
' * \param iv       initialization vector (updated after use)
' * \param input    buffer holding the input data
' * \param output   buffer holding the output data
' */
Private Declare Function sm4_crypt_cbc Lib "zlSm4.dll" (ByVal Mode As Long, ByVal Length As Long, iv As Byte, key As Byte, in_put As Byte, out_put As Byte) As Long
'��ȡ�ַ����Ĺ�ϣ����
'/**
' * \brief          Output = SM3( input buffer )
' *
' * \param input    buffer holding the  data
' * \param ilen     length of the input data
' * \param output   SM3 checksum result
' */
Private Declare Sub sm3_hash Lib "zlSm4.dll" Alias "sm3" (in_put As Byte, ByVal Length As Long, out_put As Byte)
'��ȡ�ļ���sm��ϣ����
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
'HMAC����Կ��صĹ�ϣ������Ϣ��֤�룬HMAC�������ù�ϣ�㷨����һ����Կ��һ����ϢΪ���룬����һ����ϢժҪ��Ϊ�����
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
'��ȡZLSM4���޸İ汾
'1:ֻ֧��sm4_crypt_ecb,sm4_crypt_cbc
'2:����֧��sm3��sm3_file��sm3_hmac��sm_version
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
'���ܣ���ȡIP:Port/SID��Ϣ
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
                '��ͨ�û��£����أ�������\�û�����SYSTEM�û��·��ع�����\������
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
'���ܣ�ȥ���ַ�����\0�Ժ���ַ�
    Dim lngPos As Long
    
    lngPos = InStr(strInput, Chr(0))
    If lngPos > 0 Then
        TruncZero = Mid(strInput, 1, lngPos - 1)
    Else
        TruncZero = strInput
    End If
End Function

Public Function IsOLEDBConnection(ByVal cnMain As ADODB.Connection) As Boolean
'���ܣ��жϵ�ǰ�����Ƿ���OraOLEDB����
'����Provider���жϣ��������ַ�ʽ
'��ʽһ��'Provider=OraOLEDB.Oracle.1;Password=HIS;Persist Security Info=True;User ID=ZLHIS;Data Source="DYYY";Extended Properties="PLSQLRSet=1"
'��ʽ����
'.Provider = "OraOLEDB.Oracle"
'.Open "PLSQLRSet=1;Data Source=" & strServer & strPersist_Security_Info, strUserName, strPassWord
'�����ַ�ʽ�����Զ�����.Provider����
    'ʹ��Like����Ϊ���ܺ������Ӱ汾��OraOLEDB.Oracle.1
    If UCase(cnMain.Provider) Like "ORAOLEDB.ORACLE*" Then
        IsOLEDBConnection = True
    End If
End Function

Private Sub GetServerInfoByFile(ByVal strServer As String, ByRef setServiceName As String, strServerIp As String, ByRef strServerPort As String)
'����:����tnsname.ora�ļ���ȡ������IP���˿ڡ�ʵ����
'�������: strServer=������
'�������� setServiceName = ʵ����  strServerIp = ������IP   strServerPort = �������˿�

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
    strServer = UCase(strServer): strTxt = ConvertStr(strTxt) '��ʽ���ַ�
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
        '��ȡIP
        lngTmp = InStr(1, strTxt, "HOST=")
        strTmp = Mid(strTxt, lngTmp + Len("HOST="))
        strServerIp = Mid(strTmp, 1, InStr(1, strTmp, ")") - 1)
        
        '��ȡ�˿�
        lngTmp = InStr(1, strTxt, "PORT=")
        strTmp = Mid(strTxt, lngTmp + Len("PORT="))
        strServerPort = Mid(strTmp, 1, InStr(1, strTmp, ")") - 1)
        
        '��ȡ������
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
'���ܣ���ȡOracleHome·��
    Dim arrTmp  As Variant, arrSubKey   As Variant
    Dim strHome As String, strDefault   As String, strPath As String
    Dim i       As Integer
    Dim objPE   As New clsPEReader
    Dim blnRead As Boolean
    Dim objFSO As New FileSystemObject
    
    strHome = Environ("PATH")
    '1��PATH������û�У�����ϵͳ�Ļ�����������������߷�WInϵͳ������Ϊ�����ϵͳ��MAC��
    If strHome = "" Then Exit Function
    arrTmp = Split(strHome, ";")
    strHome = ""
    For i = LBound(arrTmp) To UBound(arrTmp)
    
        If UCase(arrTmp(i)) Like "*ORA*\BIN" Then
            '�ж�Oracle��OCI���������Ƿ����
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
    '2��Ѱ��TNS_ADMIN:ORACLE_HOME & "\network\ADMIN
    strHome = Environ("TNS_ADMIN")
    If strHome <> "" Then
        If InStr(UCase(strHome), "\NETWORK\ADMIN") > 0 Then
            '�ж�TNSNAME
            If Not objFSO.FileExists(strHome & "\tnsnames.ora") Then
                strHome = ""
            End If
            '��ȡORACLE_HOME,�ж�OCI
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
    '3��ORACLE_HOME��������
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
    
    '4��ע����ж�,��ȡ64λ��32Ŀ¼���Զ���λ��SOFTWARE\Wow6432Node\Oracle 2����ȡ32λ��32λĿ¼
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
    '4.2��ALL_Homes��ʽ,ֻ��ȡ��һ�����������ġ�
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
    '����:ȥ���ַ����Ŀո�\���з�,��ת��Ϊ��д
    
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
    '���ܣ��Ƿ���64λϵͳ
    '���أ�
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
'���ܣ���ȡע�����ָ��λ�õ�ֵ
'������strKey=ע����λ���硰HKEY_CURRENT_USER\Printers\DevModePerUser"
'          strValueName=������
'          strValue=����ֵ
'          strValueType=�������ͣ�Ĭ��Ϊ�ַ���
'           blnOneString = ��REG_EXPAND_SZ��REG_MULTI_SZ,REG_BINARY��Ч��-  True �������ص�һ�ַ������Ҳ����κδ���ֻȥ���ַ���β��
'���أ��Ƿ��ȡ�ɹ�
'˵������ǰֻ��REG_SZ, REG_EXPAND_SZ, REG_MULTI_SZ��REG_DWORD��REG_BINARYʵ���˶�ȡ��û�в�ѯ�������Զ����Ҽ���
    Dim hRootKey As REGRoot, strSubKey As String
    Dim lngReturn As Long
    Dim lngKey As Long, ruType As REGValueType
    Dim lngLength As Long, strBufVar() As String, lngBuf As Long, bytBuf() As Byte, strBuf As String
    Dim i As Long, strReturn As String, strTmp As String
    '������Ч��ע����λ,��ȡ��������
    If Not GetKeyValueInfo(strKey, strValueName, hRootKey, strSubKey, ruType) Then Exit Function
    '�򿪱���
    lngReturn = RegOpenKeyEx(hRootKey, strSubKey, 0, KEY_QUERY_VaLUE, lngKey)
    If lngReturn <> ERROR_SUCCESS Then
        Exit Function
    End If
    On Error GoTo ErrH
    Select Case ruType
        Case REG_SZ, REG_EXPAND_SZ, REG_MULTI_SZ '�ַ������Ͷ�ȡ
'            lngReturn = RegQueryValueEx(lngKey, strValueName, 0, ruType, 0, lngLength)
'            If lngReturn <> ERROR_SUCCESS Then Err.Clear '���ܳ��������������
            lngLength = 1024: strBuf = Space(lngLength)
            lngReturn = RegQueryValueEx_String(lngKey, strValueName, 0, ruType, strBuf, lngLength)
            If lngReturn <> ERROR_SUCCESS Then: RegCloseKey (lngKey): Exit Function
            Select Case ruType
                Case REG_SZ
                    varValue = mdlPublic.TruncZero(strBuf)
                Case REG_EXPAND_SZ ' ���价���ַ�������ѯ���������ͷ��ض���ֵ
                    If Not blnOneString Then
                        varValue = mdlPublic.TruncZero(ExpandEnvStr(mdlPublic.TruncZero(strBuf)))
                    Else
                        varValue = mdlPublic.TruncZero(strBuf)
                    End If
                Case REG_MULTI_SZ ' �����ַ���
                    If Not blnOneString Then
                        If Len(strBuf) <> 0 Then ' �������Ƿǿ��ַ��������Էָ
                            strBufVar = Split(Left$(strBuf, Len(strBuf) - 1), Chr$(0))
                        Else ' ���ǿ��ַ�����Ҫ����S(0) ���������
                            ReDim strBufVar(0) As String
                        End If
                        ' ��������ֵ������һ���ַ������飿��
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
            ' �����ַ�����ע�⣺Ҫ���ֽ��������ת����
            If blnOneString Then
                'ѭ�����ݣ����ֽ�ת��Ϊ16�����ַ���
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
'����:��ȡĳ�����������
'���أ�=��������
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
'���ܣ����ݼ�λ��ȡ����ֵ���ӽ�,�Լ�ֵ����
'������strKey=ע����λ���硰HKEY_CURRENT_USER\Printers\DevModePerUser"
'          strValueName=������
'���Σ�
'          hRootKey=����
'          strSubKey=�ӽ�
'          lngType=������
'���أ��Ƿ��ȡ�ɹ�
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
        'ʹ�ò�ѯ��ʽ�򿪣����м������Ͳ�ѯ
        lngReturn = RegOpenKeyEx(hRootKey, strSubKey, 0, KEY_QUERY_VaLUE, hKey)
        If lngReturn <> ERROR_SUCCESS Then
            Exit Function
        End If
        If strValueName <> "" Then
            lngReturn = RegQueryValueEx_ValueType(hKey, strValueName, ByVal 0&, lngType, ByVal strName, Len(strName))
            'SetRegKey����������ص�����Ϊ�ܴ��������ֵ���̶�,�������Ϊ0�����ݴ������������ж�
            If lngReturn = ERROR_BADKEY Then
                If lngType < REG_NONE Or lngType > REG_MULTI_SZ Then lngType = REG_NONE
            End If
            '�����ֶγ��������Ȳ��������Գ����˳�
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
'���ܣ����ַ����еĻ��������滻Ϊ����ֵ
'         strInput=���������������ַ���
'���أ���ʵ�ʵ�ֵ�滻�ַ����еĻ�����������ַ���
    '// �磺 %PATH% �򷵻� "c:\;c:\windows;"
    Dim lngLen As Long, strBuf As String, strOld As String
    strOld = strInput & "  " ' ��֪ΪʲôҪ�������ַ������򷵻�ֵ������������ַ���
    strBuf = "" '// ��֧��Windows 95
    '// get the length
    lngLen = ExpandEnvironmentStrings(strOld, strBuf, lngLen)
    '// չ���ַ���
    strBuf = String$(lngLen - 1, Chr$(0))
    lngLen = ExpandEnvironmentStrings(strOld, strBuf, LenB(strBuf))
    '// ���ػ�������
    ExpandEnvStr = mdlPublic.TruncZero(strBuf)
End Function

Public Function Decode(ParamArray arrPar() As Variant) As Variant
'���ܣ�ģ��Oracle��Decode����
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
'���ܣ�ȥ��TAB�ַ������߿ո񣬻س������ֻ�ɵ��ո�ָ���
'������strText=�����ַ�
'         blnCrlf=�Ƿ�ȥ�����з�
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
'����           Sm4DecryptEcb           SM4����
'����ֵ         String                  ���ܺ��ֵ
'����б�:
'������         ����                    ˵��
'strInput       String                  Ҫ���ܵ��ַ��������ַ�����Sm4EncryptEcb���ɵĽ����
'strKey         String(Optional)        ������ԿҲ���ǽ�����Կ��32λ��16�����ַ���������ͨ��HexStringToByte���أ�
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
        '��ǰ�ͻ��˵�ZLSM4��֧�ָð汾�ļ����ַ������ܣ��Ծɽ��ܣ���Ϊһ����˵���ܽ��ܳ���ͬ���ַ���
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
'����           Sm4EncryptCbc           SM4�������
'����ֵ         String                  ���ܺ��ֵ
'����б�:
'������         ����                    ˵��
'strInput       String                  Ҫ���ܵ��ַ���
'strKey         String(Optional)        ������Կ��32λ��16�����ַ���������ͨ��HexStringToByte���أ�
'strIv          String(Optional)        ���������Կ��32λ��16�����ַ���������ͨ��HexStringToByte���أ�
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
'���ܣ�ȥ���ַ�����\0�Ժ���ַ�,�������ù���,���Ե�������clsstring
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
    
    '�Ƚ��ַ����� Unicode ת��ϵͳ��ȱʡ��ҳ
    arrReturn = StrConv(strInput, vbFromUnicode)
    lngLenBef = UBound(arrReturn) + 1
    '�жϵõ�������ĳ��ȣ�������16�����������򲹿ո��:Chr(0)
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
'���ܣ���ȡ��������
'������
'˵����
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
                    '�Ƿ�������  ��Ϊ��֧�ֳ־��Բ���д����
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
            '�Ƿ�������  ��Ϊ��֧�ֳ־��Բ���д����
            Serialize = "{NotPersistable}"
            Err.Clear
        Else
            bytData = objBag.Contents
            Serialize = EncodeBase64(bytData())
        End If
    End If
End Function

'======================================================================================================================
'����           EncodeBase64            ����Base64���룬����Base64���ַ���
'����ֵ         String                  Base64������
'����б�:
'������         ����                    ˵��
'varInput       Variant                 ��Ҫ����Base64������ַ��������ֽ����飬�ַ�����ȡUTF-8���롣Byte()����ǰ������飬Ԫ�ظ�����3�ı��������һ�δ�������ʣ�µļ��ɡ�
'����˵����Base64�ǽ������ֽڣ�ÿ6λ�ָ�Ϊ�ĸ��ֽڴ����
'======================================================================================================================
Public Function EncodeBase64(varInput As Variant) As String
    Dim bytInput()  As Byte, lngInputLen    As Long
    Dim bytOut()    As Byte, lngOutLen      As Long
    Dim i           As Long, j              As Long, lngBit     As Long
    
    On Error GoTo ErrH
    
    If VarType(varInput) = vbString Then
        If Len(varInput) = 0 Then Exit Function
        'ԭʼ����,�Ƚ�ԭ����UTF-8�ķ�ʽ����
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
    '��8-bit�ֽ�����ת��Ϊ6-bit�ֽ�����
    For i = 0 To lngInputLen - 1
        If lngBit = 0 Then 'bytOut(J)δ��д��
            bytOut(j) = (bytInput(i) And &HFC) \ &H4
            j = j + 1
            bytOut(j) = (bytInput(i) And &H3) * &H10
            lngBit = 2 '234567 'NNNN01 'N:Next byte
        ElseIf lngBit = 2 Then 'bytOut(J)�ѱ�д����λ
            bytOut(j) = bytOut(j) Or ((bytInput(i) And &HF0) \ &H10)
            j = j + 1
            bytOut(j) = (bytInput(i) And &HF) * &H4
            lngBit = 4 '4567PP 'P:Prev byte 'NN0123 'N:Next byte
        ElseIf lngBit = 4 Then 'bytOut(J)�ѱ�д����λ
            bytOut(j) = bytOut(j) Or ((bytInput(i) And &HC0) / &H40)
            j = j + 1
            bytOut(j) = bytInput(i) And &H3F
            j = j + 1
            lngBit = 0 '67PPPP 'P:Prev byte '012345
        End If
    Next

    For i = 0 To lngOutLen - 1
        bytOut(i) = EncBase64Char(bytOut(i)) 'ת��ΪBase64�ַ�
    Next
    EncodeBase64 = StrConv(bytOut, vbUnicode) & String(2 - (lngInputLen - 1) Mod 3, "=") 'ԭ��ʣ�����ݲ���3���ֽ���Ҫ����
    Exit Function
ErrH:
    Err.Clear
    If 0 = 1 Then
        Resume
    End If
End Function

'======================================================================================================================
'����           StringToUTF8Bytes       ���ַ���ת��ΪUTF-8������ֽ�����
'����ֵ         Byte()                  16�����ַ���ת�����ֽ���
'����б�:
'������         ����                    ˵��
'strInput      String                  16�����ַ���
'======================================================================================================================
Public Function StringToUTF8Bytes(strInput As String) As Byte()
    Dim bytUTF8Bytes() As Byte
    Dim lngBytesRequired As Long
    
    '�ȼ��������ֽ���
    lngBytesRequired = WideCharToMultiByte(CP_UTF8, 0, ByVal StrPtr(strInput), Len(strInput), ByVal 0, 0, ByVal 0, ByVal 0)
     
    'Ȼ��ת��
    ReDim bytUTF8Bytes(lngBytesRequired - 1)
    WideCharToMultiByte CP_UTF8, 0, ByVal StrPtr(strInput), Len(strInput), bytUTF8Bytes(0), lngBytesRequired, ByVal 0, ByVal 0
    
    StringToUTF8Bytes = bytUTF8Bytes
End Function

'======================================================================================================================
'����           EncBase64Char           ��6-bit�ֽ�ת��ΪBase64�ַ�
'����ֵ         Byte                    �ַ���ֵ
'����б�:
'������         ����                    ˵��
'����˵����Base64�ǽ������ֽڣ�ÿ6λ�ָ�Ϊ�ĸ��ֽڴ����
'======================================================================================================================
Private Function EncBase64Char(ByVal bytValue As Byte) As Byte
    If bytValue < 26 Then '26����дӢ����ĸ
        EncBase64Char = bytValue + &H41
    ElseIf bytValue < 52 Then '26��СдӢ����ĸ
        EncBase64Char = bytValue + &H61 - 26
    ElseIf bytValue < 62 Then '10������
        EncBase64Char = bytValue + &H30 - 52
    ElseIf bytValue = 62 Then
        EncBase64Char = &H2B '+
    Else
        EncBase64Char = &H2F '/
    End If
End Function


