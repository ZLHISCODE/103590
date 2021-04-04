VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPEReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'PE�ṹ
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
Option Explicit
Private Type IMAGE_DOS_HEADER '64B
   e_magic(1)                                   As Byte     '�ֱ�ΪMZ,4Dh��5Ah
   e_cblp                                       As Integer  '�ļ����һҳ�ֽ���
   e_cp                                         As Integer  '�ļ���ҳ��(512B/ҳ)
   e_crlc                                       As Integer  '�ض�λ������
   e_cparhdr                                    As Integer  '�ļ�ͷ�ܶ���(16B/��)
   e_minalloc                                   As Integer
   e_maxalloc                                   As Integer
   e_ss                                         As Integer  'SS:SP
   e_sp                                         As Integer  'SS:SP
   e_csum                                       As Integer  'У���
   e_ip                                         As Integer  'CS:IP
   e_cs                                         As Integer  'CS:IP
   e_lfarlc                                     As Integer  '�ض�λ��ƫ�Ƶ�ַ
   e_ovno                                       As Integer
   e_res(3)                                     As Integer
   e_oemid                                      As Integer
   e_oeminfo                                    As Integer
   e_res2(9)                                    As Integer
   e_lfanew                                     As Long     'PEͷƫ��,ָ��PE�ļ�ͷ
End Type

Private Enum MACHINECONSTANTS
   Intel80386 = &H14C
   Intel80486 = &H14D
   IntelPentium = &H14E
   R3000High = &H160
   R3000Low = &H162
   R4000Low = &H166
   R10000Low = &H168
   DECAlpha = &H184
   PowerPc = &H1F0
End Enum

Private Type IMAGE_FILE_HEADER  '��20�ֽ�
   Machine                                      As Integer  'CPU������ƽ̨������Intel��Ϊ&H014C
   NumberOfSections                             As Integer  '���������ļ��Ľ�����
   TimeDateStamp                                As Long
   PointerToSymbolTable                         As Long     '���ű�ָ�룺ָ����ű���������
   NumberOfSymbols                              As Long     '���ű����������ű��з�������
   SizeOfOptionalHeader                         As Integer  '��ѡͷ��С����ѡ�ļ�ͷ�ṹ��С
   Characteristics                              As Integer  '�ļ���Ϣ���ļ���Ϣ��־����EXE��DLL
End Type

Private Type IMAGE_DATA_DIRECTORY
   VirtualAddress                               As Long     '���ݿ��RVA
   iSize                                        As Long     '���ݿ��С
End Type

Private Type IMAGE_OPTIONAL_HEADER '��224�ֽ�
   magic                                        As Integer  '
   MajorLinkerVersion                           As Byte
   MinorLinkerVersion                           As Byte
   SizeOfCode                                   As Long     '��������ֽ���
   SizeOfInitializedData                        As Long     '�ѳ�ʼ�����ݵ����ֽ���
   SizeOfUninitializedData                      As Long     'δ��ʼ�����ݵ����ֽ���
   AddressOfEntryPoint                          As Long     'PE�ļ�����ڵ�ַ
   BaseOfCode                                   As Long     '����ڵ���ʼRVA
   BaseOfData                                   As Long     '���ݽڵ���ʼRVA
   ImageBase                                    As Long     '����������������ַ�������س���ĵ�ַ
   SectionAlignment                             As Long     '�����ڴ��ڵĶ�������
   FileAlignment                                As Long     '�ļ��нڵĶ�������
   MajorOperatingSystemVersion                  As Integer
   MinorOperatingsystemVersion                  As Integer
   MajorImageVersion                            As Integer
   MinorImageVersion                            As Integer
   MajorSubsystemVersion                        As Integer
   MinorSubsystemVersion                        As Integer
   Reserved1                                    As Long
   SizeOfImage                                  As Long     '�ڴ����������ӳ��Ĵ�С
   SizeOfHeaders                                As Long     '����ͷ���Ͻڱ�������Ĵ�С
   CheckSum                                     As Long
   Subsystem                                    As Integer  '��ϵͳ��
   DllCharacteristics                           As Integer
   SizeOfStackReserve                           As Long
   SizeOfStackCommit                            As Long
   SizeOfHeapReserve                            As Long
   SizeOfHeapCommit                             As Long
   LoaderFlags                                  As Long
   NumberOfRvaAndSizes                          As Long     'Rva��Size����:����Ŀ¼������
   DataDirecoty(15)                             As IMAGE_DATA_DIRECTORY
End Type
'DataDirecoty�ĸ���Ŀ¼����;
Private Enum DATA_DIRECTORY
    DATA_DIRECTORY_EXPORT_TABLE = 0 '�������ĵ�ַ�ʹ�
    DATA_DIRECTORY_IMPORT_TABLE = 1 '������ĵ�ַ�ʹ�С
    DATA_DIRECTORY_RESOURCE_TABLE = 2 '��Դ���ĵ�ַ�ʹ�
    DATA_DIRECTORY_EXCEPTION_TABLE = 3 '�쳣���ĵ�ַ�ʹ�С
    DATA_DIRECTORY_CERTIFICATE_TABLE = 4 '����֤����ĵ�ַ�ʹ�С
    DATA_DIRECTORY_BASE_RELOCATION_TABLE = 5 '��ַ�ض�λ���ĵ�ַ�ʹ�С
    DATA_DIRECTORY_DEBUG_TABLE = 6 '����������ʼ��ַ�ʹ�С
    DATA_DIRECTORY_ARCHITECTURE = 7 '����������Ϊ0��
    DATA_DIRECTORY_GLOBAL_PTR = 8 '�����洢��ȫ��ָ��Ĵ����е�һ��ֵ��RVA������ṹ��Size �����Ϊ0.
    DATA_DIRECTORY_TLS_TABLE = 9 '�ֲ߳̾��洢��TLS�����ĵ�ַ�ʹ�С
    DATA_DIRECTORY_LOAD_CONFIG_TABLE = 10 '�������ñ��ĵ�ַ�ʹ�С
    DATA_DIRECTORY_BOUND_IMPORT = 11 '�󶨵�����ĵ�ַ�ʹ�С��
    DATA_DIRECTORY_IAT = 12 '�����ַ���ĵ�ַ�ʹ�С
    DATA_DIRECTORY_DELAY_IMPORT_DESCRIPTOR = 13 '�ӳٵ����������ĵ�ַ�ʹ�С
    DATA_DIRECTORY_CLR_RUNTIME_HEADER = 14 'CLR ����ʱͷ���ĵ�ַ�ʹ�С
    DATA_DIRECTORY_OTHER = 15 '����������Ϊ0��
End Enum

Private Type IMAGE_NT_HEADERS
   Signature(3)                                 As Byte     '�ֱ�ΪPE 0 0 50h, 45h, 00h, 00h
   FileHeader                                   As IMAGE_FILE_HEADER
   OptionalHeader                               As IMAGE_OPTIONAL_HEADER
End Type

Private Type IMAGE_SECTION_HEADER
   Name1(7)                                     As Byte     '�ڵ�����
   'Name As String * 8 '�ڵ�����
   PhysicalAddress                              As Long     '�������ݵ�ʵ���ֽڳ��ȣ���VirtualSize����ͬһ��ַ
   VirtualAddress                               As Long     '���ڴ��ڵ�RVA��ַ
   SizeOfRawData                                As Long     '�����ļ��ж�����ֽڳ���
   PointerToRawData                             As Long     '�ڻ����ļ���ƫ����
   PointerToRelocations                         As Long
   PointertoLinenumbers                         As Long
   NumberOfRelocations                          As Integer
   NumberOfLinenumbers                          As Integer
   Characteristics                              As Long     '�ڵ�����
End Type

Private Type IMAGE_RESOURCE_DIRECTORY
    Characteristics                             As Long     '������Ϊ��Դ�����ԣ�������ʵ������0
    TimeDateStamp                               As Long     '��Դ�Ĳ���ʱ��
    MajorVersion                                As Integer  '������Ϊ��Դ�İ汾��������ʵ������0
    MinorVersion                                As Integer  '         ?
    NumberOfNamedEntries                        As Integer  '�������������������
    NumberOfIdEntries                           As Integer  '��ID�������������
End Type

Private Type IMAGE_RESOURCE_DIRECTORY_ENTRY
    Name1                                       As Long     'Ŀ¼��������ַ���ָ���ID
    OffsetToData                                As Long     'Ŀ¼��ָ��
End Type

Private Type IMAGE_RESOURCE_DIR_STRING_U
    Length1                                     As Integer  '�ַ����ĳ���
    NameString                                  As Integer  'UNICODE�ַ����������ַ����ǲ������ģ���������ֻ����һ��dw��ʾ��ʵ���ϵ�����Ϊ100��ʱ�������������NameString dw 100 dup (?)
End Type
'CLR 2.0 header structure.
Private Type IMAGE_COR20_HEADER
    'Header versioning
    Cb                                          As Long     'ͷ���ֽڴ�С
    MajorRuntimeVersion                         As Integer  'CLR��Ҫ���е���С�汾���屾��
    MinorRuntimeVersion                         As Integer  'CLR��Ҫ���е���С�汾�ΰ汾��
    'Symbol table and startup information
    metaData                                    As IMAGE_DATA_DIRECTORY 'Rav��Ԫ���ݵĴ�С
    Flags                                       As Long     '�����Ʊ��
    EntryPointToken                             As Long
    Resources                                   As IMAGE_DATA_DIRECTORY 'Rav���й���Դ�Ĵ�С
    StrongNameSignature                         As IMAGE_DATA_DIRECTORY 'Rav���������pe�ļ��Ĺ�ϣ���ݵĴ�С
    'Regular fixup and binding information
    CodeManagerTable                            As IMAGE_DATA_DIRECTORY 'Rva�ʹ���������Ĵ�С
    VTableFixups                                As IMAGE_DATA_DIRECTORY 'Rav��һ���������������ɵ�������ֽڴ�С
    ExportAddressTableJumps                     As IMAGE_DATA_DIRECTORY 'Rav����jump thunk�ĵ�ַ��ɵ�����Ĵ�С
    'Precompiled image info (internal use only - set to zero)
    ManagedNativeHeader                         As IMAGE_DATA_DIRECTORY 'ΪԤ����������ģ�������Ϊ0
End Type
'.net Meta Data Structor
Private Type CLR_MetaDataVer
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
'IMAGE_FILE_HEADER��Characteristics��λ����
Private Const IMAGE_FILE_RELOCS_STRIPPED = 0 '��0λΪ1ʱ:�ļ��в������ض�λ��Ϣ
Private Const IMAGE_FILE_EXECUTABLE_IMAGE = 1 '��1λΪ1ʱ:�ļ��ǿ�ִ�е�
Private Const IMAGE_FILE_LINE_NUMS_STRIPPED = 2 '��2λΪ1ʱ:����������Ϣ
Private Const IMAGE_FILE_LOCAL_SYMS_STRIPPED = 3 '��3λΪ1ʱ:�����ڷ�����Ϣ
Private Const IMAGE_FILE_BYTES_REVERSED_LO = 7 '��7λΪ1ʱ:Сβ��ʽ
Private Const IMAGE_FILE_32BIT_MACHINE = 8 '��8λΪ1ʱ:ֻ��32λƽ̨����
Private Const IMAGE_FILE_DEBUG_STRIPPED = 9 '��9λΪ1ʱ:������������Ϣ
Private Const IMAGE_FILE_REMOVABLE_RUN_FROM_SWAP = 10 '��10λΪ1ʱ:���ܴӿ��ƶ�������
Private Const IMAGE_FILE_NET_RUN_FROM_SWAP = 11 '��11λΪ1ʱ:���ܴ���������
Private Const IMAGE_FILE_SYSTEM = 12 '��12λΪ1ʱ:ϵͳ�ļ�������ֱ������
Private Const IMAGE_FILE_DLL = 13 '��13λΪ1ʱ:DLL�ļ�
Private Const IMAGE_FILE_UP_SYSTEM_ONLY = 14 '��14λΪ1ʱ:�ļ������ڶദ����������
Private Const IMAGE_FILE_BYTES_REVERSED_HI = 15 '��15λΪ1ʱ:��β��ʽ

'���Ա���
Private mstrPEFile                              As String   'PE�ļ�·��
Private mblnIsDLL                               As Boolean  'PE�ļ��Ƿ���DLL
Private mblnIsActivexEXE                        As Boolean  'PE�Ƿ���ActivexEXE
Private mblnIsNetLibrary                        As Boolean  'PE�ļ��Ƿ���.NET DLL
Private mstrCLRVer                              As String   'PE�ļ���CLR�汾

Property Get PEFile() As String
'���ܣ���ȡPE�ļ�
    PEFile = mstrPEFile
End Property

Property Get IsDLL() As Boolean
'���ܣ�PE�ļ��Ƿ���DLL
    IsDLL = mblnIsDLL
End Property

Property Get IsActivexEXE() As Boolean
'���ܣ�PE�Ƿ���ActivexEXE
    IsActivexEXE = mblnIsActivexEXE
End Property

Property Get IsNetLibrary() As Boolean
'���ܣ�PE�ļ��Ƿ���.NET DLL
    IsNetLibrary = mblnIsNetLibrary
End Property

Property Get CLRVer() As String
'���ܣ�PE�ļ���CLR�汾
    CLRVer = mstrCLRVer
End Property

Public Function LoadPEFile(ByVal strPEFile As String, Optional ByRef strErr As String) As Boolean
    Dim idhDosHeader As IMAGE_DOS_HEADER 'DOSͷ
    Dim inhNTHeader As IMAGE_NT_HEADERS 'NTͷ
    Dim ishSections() As IMAGE_SECTION_HEADER '������
    Dim irdResDir As IMAGE_RESOURCE_DIRECTORY
    Dim irdeEntries() As IMAGE_RESOURCE_DIRECTORY_ENTRY '��Դ����
    Dim irdsuResDirString As IMAGE_RESOURCE_DIR_STRING_U
    Dim ichCLR As IMAGE_COR20_HEADER, cmdvMetaData As CLR_MetaDataVer
    Dim i As Integer, rsNum As Integer, j As Long
    Dim lngTmpAddr As Long, lngTmpAddr1 As Long, lngTmpAddr2 As Long
    Dim strTmp As String, intUnicodeChar As Integer
    
    On Error GoTo errH:
    mstrPEFile = "": mblnIsDLL = False
    mblnIsActivexEXE = False: mblnIsNetLibrary = False
    mstrCLRVer = ""
    If Not gobjFSO.FileExists(strPEFile) Then
        strErr = "δ�ҵ���Ҫע��/��ע����ļ�"
        Exit Function
    End If
    'PE�ļ���СС��1000�ֽھͲ�����ЧPE
    If gobjFSO.GetFile(strPEFile).Size < 1000 Then
        strErr = "�ļ���ʽ��Ч(С��1KB),������Ч��PE�ļ�"
        Exit Function
    End If
    Open strPEFile For Binary As #1
    Get #1, , idhDosHeader
    If idhDosHeader.e_magic(0) = Asc("M") And idhDosHeader.e_magic(1) = Asc("Z") Then
        Get #1, idhDosHeader.e_lfanew + 1, inhNTHeader
        If inhNTHeader.Signature(0) = Asc("P") And inhNTHeader.Signature(1) = Asc("E") Then
            ReDim ishSections(inhNTHeader.FileHeader.NumberOfSections - 1)
            Get #1, , ishSections
        Else
            strErr = "�ļ���ʽ��Ч(ȱʧPE�ؼ���),������Ч��PE�ļ�"
            Close #1
            Exit Function
        End If
    Else
        strErr = "�ļ���ʽ��Ч(ȱʧMZ�ؼ���),������Ч��PE�ļ�"
        Close #1
        Exit Function
    End If
    mstrPEFile = strPEFile
    '�ж��Ƿ���DLL
    If (inhNTHeader.FileHeader.Characteristics And (2 ^ IMAGE_FILE_DLL)) = 0 Then
        strErr = "������Ч��DLL Library�ļ�"
    Else
        mblnIsDLL = True
    End If
    '�ж��Ƿ���TypeLib��Ϣ
    If inhNTHeader.OptionalHeader.DataDirecoty(DATA_DIRECTORY_RESOURCE_TABLE).VirtualAddress = 0 And inhNTHeader.OptionalHeader.DataDirecoty(DATA_DIRECTORY_RESOURCE_TABLE).iSize = 0 Then
    Else
        rsNum = -1
        For i = 0 To inhNTHeader.FileHeader.NumberOfSections - 1
            '������ԴĿ¼���ڽڱ��
            '�ж���ԴĿ¼���ڵ������ַ����������һ���ڵķ�Χ
            If inhNTHeader.OptionalHeader.DataDirecoty(DATA_DIRECTORY_RESOURCE_TABLE).VirtualAddress >= ishSections(i).VirtualAddress And _
                inhNTHeader.OptionalHeader.DataDirecoty(DATA_DIRECTORY_RESOURCE_TABLE).VirtualAddress <= ishSections(i).VirtualAddress + ishSections(i).SizeOfRawData Then
                rsNum = i
            End If
        Next
'        If rsNum < 0 Then rsNum = 0
        If rsNum <> -1 Then
            '������ԴĿ¼��Խڵ��ļ�ƫ�Ƶ�ַ+������ԴĿ¼���ļ��е���Է��ʵ�ַ
            lngTmpAddr = inhNTHeader.OptionalHeader.DataDirecoty(DATA_DIRECTORY_RESOURCE_TABLE).VirtualAddress - ishSections(rsNum).VirtualAddress + ishSections(rsNum).PointerToRawData + 1
            Get #1, lngTmpAddr, irdResDir
            If irdResDir.NumberOfNamedEntries + irdResDir.NumberOfIdEntries <> 0 Then
                lngTmpAddr = inhNTHeader.OptionalHeader.DataDirecoty(DATA_DIRECTORY_RESOURCE_TABLE).VirtualAddress - ishSections(rsNum).VirtualAddress + ishSections(rsNum).PointerToRawData + 1
                lngTmpAddr1 = lngTmpAddr + Len(irdResDir)
                ReDim irdeEntries(irdResDir.NumberOfNamedEntries + irdResDir.NumberOfIdEntries - 1)
                Get #1, lngTmpAddr1, irdeEntries
                '��ȡ��Դ����
                For i = LBound(irdeEntries) To UBound(irdeEntries)
                    '������λΪ1����ʾ�Զ�����Դ��Ϊ0��ʾ��Դ���͵�id��
                    If irdeEntries(i).Name1 < 0 And irdeEntries(i).OffsetToData < 0 Then
                        irdeEntries(i).Name1 = irdeEntries(i).Name1 And &H7FFFFFFF
                        lngTmpAddr2 = lngTmpAddr + irdeEntries(i).Name1
                        Get #1, lngTmpAddr2, irdsuResDirString
                        lngTmpAddr2 = lngTmpAddr2 + Len(irdsuResDirString)
                        strTmp = ChrW(irdsuResDirString.NameString)
                        For j = 0 To irdsuResDirString.Length1 - 2
                          Get #1, lngTmpAddr2, intUnicodeChar
                          strTmp = strTmp + ChrW(intUnicodeChar)
                          lngTmpAddr2 = lngTmpAddr2 + Len(intUnicodeChar)
                        Next
                        If strTmp = "TYPELIB" Then
                            '����TYPELIB
                            mblnIsActivexEXE = True
                            Exit For
                        End If
                    End If
                Next
            End If
        End If
    End If
    '��ȡCLR�汾
    If inhNTHeader.OptionalHeader.DataDirecoty(DATA_DIRECTORY_CLR_RUNTIME_HEADER).VirtualAddress = 0 Then
    Else
        rsNum = -1
        For i = 0 To inhNTHeader.FileHeader.NumberOfSections - 1
            '������ԴĿ¼���ڽڱ��
            '�ж���ԴĿ¼���ڵ������ַ����������һ���ڵķ�Χ
            If inhNTHeader.OptionalHeader.DataDirecoty(DATA_DIRECTORY_CLR_RUNTIME_HEADER).VirtualAddress >= ishSections(i).VirtualAddress And _
                inhNTHeader.OptionalHeader.DataDirecoty(DATA_DIRECTORY_CLR_RUNTIME_HEADER).VirtualAddress <= ishSections(i).VirtualAddress + ishSections(i).SizeOfRawData Then
                rsNum = i
            End If
        Next
'        If rsNum < 0 Then rsNum = 0
        If rsNum <> -1 Then
            mblnIsNetLibrary = True
            '������ԴĿ¼��Խڵ��ļ�ƫ�Ƶ�ַ+������ԴĿ¼���ļ��е���Է��ʵ�ַ
            lngTmpAddr = inhNTHeader.OptionalHeader.DataDirecoty(DATA_DIRECTORY_CLR_RUNTIME_HEADER).VirtualAddress - ishSections(rsNum).VirtualAddress + ishSections(rsNum).PointerToRawData + 1
            Get #1, lngTmpAddr, ichCLR
            lngTmpAddr = ichCLR.metaData.VirtualAddress - ishSections(rsNum).VirtualAddress + ishSections(rsNum).PointerToRawData + 1
            Get #1, lngTmpAddr, cmdvMetaData
            mstrCLRVer = TruncZero(StrConv(cmdvMetaData.VersionString, vbUnicode))
        Else
            strErr = "δ�ܶ�ȡCLRͷ���ڵĽ���Ϣ����"
        End If
    End If
    Close #1
    '��DLL���ǿ�����ActivexEXE
    If Not mblnIsDLL And mblnIsActivexEXE Then
        strErr = ""
    End If
    LoadPEFile = mblnIsDLL Or mblnIsActivexEXE
'    gobjTrace.WriteInfo "LoadPEFile", "DLL", mblnIsDLL, "ActiveEXE", mblnIsActivexEXE, "NET", mblnIsNetLibrary, "CLRVer", mstrCLRVer, "ERROR", strErr
    Exit Function
errH:
    LoadPEFile = False
    Close #1
    err.Clear
    If False Then
        Resume
    End If
End Function
