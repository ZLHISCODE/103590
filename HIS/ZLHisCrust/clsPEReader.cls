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
'PE结构
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
Option Explicit
Private Type IMAGE_DOS_HEADER '64B
   e_magic(1)                                   As Byte     '分别为MZ,4Dh和5Ah
   e_cblp                                       As Integer  '文件最后一页字节数
   e_cp                                         As Integer  '文件总页数(512B/页)
   e_crlc                                       As Integer  '重定位项数量
   e_cparhdr                                    As Integer  '文件头总段数(16B/段)
   e_minalloc                                   As Integer
   e_maxalloc                                   As Integer
   e_ss                                         As Integer  'SS:SP
   e_sp                                         As Integer  'SS:SP
   e_csum                                       As Integer  '校验和
   e_ip                                         As Integer  'CS:IP
   e_cs                                         As Integer  'CS:IP
   e_lfarlc                                     As Integer  '重定位表偏移地址
   e_ovno                                       As Integer
   e_res(3)                                     As Integer
   e_oemid                                      As Integer
   e_oeminfo                                    As Integer
   e_res2(9)                                    As Integer
   e_lfanew                                     As Long     'PE头偏移,指向PE文件头
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
   IA64 = &H200
   AMD64 = &H8664
End Enum

Private Type IMAGE_FILE_HEADER  '共20字节
   Machine                                      As Integer  'CPU：运行平台，对于Intel，为&H014C
   NumberOfSections                             As Integer  '节数量：文件的节数量
   TimeDateStamp                                As Long
   PointerToSymbolTable                         As Long     '符号表指针：指向符号表，调试用
   NumberOfSymbols                              As Long     '符号表数量：符号表中符号数量
   SizeOfOptionalHeader                         As Integer  '可选头大小：可选文件头结构大小
   Characteristics                              As Integer  '文件信息：文件信息标志，如EXE或DLL
End Type

Private Type IMAGE_DATA_DIRECTORY
   VirtualAddress                               As Long     '数据块的RVA
   iSize                                        As Long     '数据块大小
End Type

Private Type IMAGE_OPTIONAL_HEADER '共224字节
   magic                                        As Integer  '
   MajorLinkerVersion                           As Byte
   MinorLinkerVersion                           As Byte
   SizeOfCode                                   As Long     '代码的总字节数
   SizeOfInitializedData                        As Long     '已初始化数据的总字节数
   SizeOfUninitializedData                      As Long     '未初始化数据的总字节数
   AddressOfEntryPoint                          As Long     'PE文件的入口地址
   BaseOfCode                                   As Long     '代码节的起始RVA
   BaseOfData                                   As Long     '数据节的起始RVA
   ImageBase                                    As Long     '整个程序的虚拟基地址，即加载程序的地址
   SectionAlignment                             As Long     '进入内存后节的对齐粒度
   FileAlignment                                As Long     '文件中节的对齐粒度
   MajorOperatingSystemVersion                  As Integer
   MinorOperatingsystemVersion                  As Integer
   MajorImageVersion                            As Integer
   MinorImageVersion                            As Integer
   MajorSubsystemVersion                        As Integer
   MinorSubsystemVersion                        As Integer
   Reserved1                                    As Long
   SizeOfImage                                  As Long     '内存中这个程序映像的大小
   SizeOfHeaders                                As Long     '所有头加上节表描述项的大小
   CheckSum                                     As Long
   Subsystem                                    As Integer  '子系统号
   DllCharacteristics                           As Integer
   SizeOfStackReserve                           As Long
   SizeOfStackCommit                            As Long
   SizeOfHeapReserve                            As Long
   SizeOfHeapCommit                             As Long
   LoaderFlags                                  As Long
   NumberOfRvaAndSizes                          As Long     'Rva与Size数量:数据目录项数量
   DataDirecoty(15)                             As IMAGE_DATA_DIRECTORY
End Type
'DataDirecoty的各个目录的用途
Private Enum DATA_DIRECTORY
    DATA_DIRECTORY_EXPORT_TABLE = 0 '导出表的地址和大
    DATA_DIRECTORY_IMPORT_TABLE = 1 '导入表的地址和大小
    DATA_DIRECTORY_RESOURCE_TABLE = 2 '资源表的地址和大
    DATA_DIRECTORY_EXCEPTION_TABLE = 3 '异常表的地址和大小
    DATA_DIRECTORY_CERTIFICATE_TABLE = 4 '属性证书表的地址和大小
    DATA_DIRECTORY_BASE_RELOCATION_TABLE = 5 '基址重定位表的地址和大小
    DATA_DIRECTORY_DEBUG_TABLE = 6 '调试数据起始地址和大小
    DATA_DIRECTORY_ARCHITECTURE = 7 '保留，必须为0。
    DATA_DIRECTORY_GLOBAL_PTR = 8 '将被存储在全局指针寄存器中的一个值的RVA。这个结构的Size 域必须为0.
    DATA_DIRECTORY_TLS_TABLE = 9 '线程局部存储（TLS）表的地址和大小
    DATA_DIRECTORY_LOAD_CONFIG_TABLE = 10 '加载配置表的地址和大小
    DATA_DIRECTORY_BOUND_IMPORT = 11 '绑定导入表的地址和大小。
    DATA_DIRECTORY_IAT = 12 '导入地址表的地址和大小
    DATA_DIRECTORY_DELAY_IMPORT_DESCRIPTOR = 13 '延迟导入描述符的地址和大小
    DATA_DIRECTORY_CLR_RUNTIME_HEADER = 14 'CLR 运行时头部的地址和大小
    DATA_DIRECTORY_OTHER = 15 '保留，必须为0。
End Enum

Private Type IMAGE_NT_HEADERS
   Signature(3)                                 As Byte     '分别为PE 0 0 50h, 45h, 00h, 00h
   FileHeader                                   As IMAGE_FILE_HEADER
   OptionalHeader                               As IMAGE_OPTIONAL_HEADER
End Type

Private Type IMAGE_SECTION_HEADER
   Name1(7)                                     As Byte     '节的名称
   'Name As String * 8 '节的名称
   PhysicalAddress                              As Long     '节区数据的实际字节长度，与VirtualSize共有同一地址
   VirtualAddress                               As Long     '如内存后节的RVA地址
   SizeOfRawData                                As Long     '节在文件中对齐后字节长度
   PointerToRawData                             As Long     '节基于文件的偏移量
   PointerToRelocations                         As Long
   PointertoLinenumbers                         As Long
   NumberOfRelocations                          As Integer
   NumberOfLinenumbers                          As Integer
   Characteristics                              As Long     '节的属性
End Type

Private Type IMAGE_RESOURCE_DIRECTORY
    Characteristics                             As Long     '理论上为资源的属性，不过事实上总是0
    TimeDateStamp                               As Long     '资源的产生时刻
    MajorVersion                                As Integer  '理论上为资源的版本，不过事实上总是0
    MinorVersion                                As Integer  '         ?
    NumberOfNamedEntries                        As Integer  '以名称命名的入口数量
    NumberOfIdEntries                           As Integer  '以ID命名的入口数量
End Type

Private Type IMAGE_RESOURCE_DIRECTORY_ENTRY
    Name1                                       As Long     '目录项的名称字符串指针或ID
    OffsetToData                                As Long     '目录项指针
End Type

Private Type IMAGE_RESOURCE_DIR_STRING_U
    Length1                                     As Integer  '字符串的长度
    NameString                                  As Integer  'UNICODE字符串，由于字符串是不定长的，所以这里只能用一个dw表示，实际上当长度为100的时候，这里的数据是NameString dw 100 dup (?)
End Type
'CLR 2.0 header structure.
Private Type IMAGE_COR20_HEADER
    'Header versioning
    Cb                                          As Long     '头的字节大小
    MajorRuntimeVersion                         As Integer  'CLR需要运行的最小版本主板本号
    MinorRuntimeVersion                         As Integer  'CLR需要运行的最小版本次版本号
    'Symbol table and startup information
    metaData                                    As IMAGE_DATA_DIRECTORY 'Rav和元数据的大小
    Flags                                       As Long     '二进制标记
    EntryPointToken                             As Long
    Resources                                   As IMAGE_DATA_DIRECTORY 'Rav和托管资源的大小
    StrongNameSignature                         As IMAGE_DATA_DIRECTORY 'Rav和用于这个pe文件的哈希数据的大小
    'Regular fixup and binding information
    CodeManagerTable                            As IMAGE_DATA_DIRECTORY 'Rva和代码管理表的大小
    VTableFixups                                As IMAGE_DATA_DIRECTORY 'Rav和一个由虚拟表修正组成的数组的字节大小
    ExportAddressTableJumps                     As IMAGE_DATA_DIRECTORY 'Rav和由jump thunk的地址组成的数组的大小
    'Precompiled image info (internal use only - set to zero)
    ManagedNativeHeader                         As IMAGE_DATA_DIRECTORY '为预编译而保留的，被设置为0
End Type
'.net Meta Data Structor
Private Type CLR_MetaDataVer
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
'IMAGE_FILE_HEADER：Characteristics各位含义
Private Const IMAGE_FILE_RELOCS_STRIPPED = 0 '第0位为1时:文件中不存在重定位信息
Private Const IMAGE_FILE_EXECUTABLE_IMAGE = 1 '第1位为1时:文件是可执行的
Private Const IMAGE_FILE_LINE_NUMS_STRIPPED = 2 '第2位为1时:不存在行信息
Private Const IMAGE_FILE_LOCAL_SYMS_STRIPPED = 3 '第3位为1时:不存在符号信息
Private Const IMAGE_FILE_BYTES_REVERSED_LO = 7 '第7位为1时:小尾方式
Private Const IMAGE_FILE_32BIT_MACHINE = 8 '第8位为1时:只在32位平台运行
Private Const IMAGE_FILE_DEBUG_STRIPPED = 9 '第9位为1时:不包含调试信息
Private Const IMAGE_FILE_REMOVABLE_RUN_FROM_SWAP = 10 '第10位为1时:不能从可移动盘运行
Private Const IMAGE_FILE_NET_RUN_FROM_SWAP = 11 '第11位为1时:不能从网络运行
Private Const IMAGE_FILE_SYSTEM = 12 '第12位为1时:系统文件。不能直接运行
Private Const IMAGE_FILE_DLL = 13 '第13位为1时:DLL文件
Private Const IMAGE_FILE_UP_SYSTEM_ONLY = 14 '第14位为1时:文件不能在多处理器上运行
Private Const IMAGE_FILE_BYTES_REVERSED_HI = 15 '第15位为1时:大尾方式

'属性变量
Private mstrPEFile                              As String   'PE文件路径
Private mblnIsDLL                               As Boolean  'PE文件是否是DLL
Private mblnIsActivexEXE                        As Boolean  'PE是否是ActivexEXE
Private mblnIsNetLibrary                        As Boolean  'PE文件是否是.NET DLL
Private mstrCLRVer                              As String   'PE文件的CLR版本

Property Get PEFile() As String
'功能：获取PE文件
    PEFile = mstrPEFile
End Property

Property Get IsDLL() As Boolean
'功能：PE文件是否是DLL
    IsDLL = mblnIsDLL
End Property

Property Get IsActivexEXE() As Boolean
'功能：PE是否是ActivexEXE
    IsActivexEXE = mblnIsActivexEXE
End Property

Property Get IsNetLibrary() As Boolean
'功能：PE文件是否是.NET DLL
    IsNetLibrary = mblnIsNetLibrary
End Property

Property Get CLRVer() As String
'功能：PE文件的CLR版本
    CLRVer = mstrCLRVer
End Property

Public Function LoadPEFile(ByVal strPEFile As String, Optional ByRef strErr As String) As Boolean
    Dim idhDosHeader As IMAGE_DOS_HEADER 'DOS头
    Dim inhNTHeader As IMAGE_NT_HEADERS 'NT头
    Dim ishSections() As IMAGE_SECTION_HEADER '段数组
    Dim irdResDir As IMAGE_RESOURCE_DIRECTORY
    Dim irdeEntries() As IMAGE_RESOURCE_DIRECTORY_ENTRY '资源数组
    Dim irdsuResDirString As IMAGE_RESOURCE_DIR_STRING_U
    Dim ichCLR As IMAGE_COR20_HEADER, cmdvMetaData As CLR_MetaDataVer
    Dim i As Integer, rsNum As Integer, j As Long
    Dim lngTmpAddr As Long, lngTmpAddr1 As Long, lngTmpAddr2 As Long
    Dim strTmp As String, intUnicodeChar As Integer
    
    On Error GoTo ErrH
    mstrPEFile = "": mblnIsDLL = False
    mblnIsActivexEXE = False: mblnIsNetLibrary = False
    mstrCLRVer = ""
    If Not gobjFSO.FileExists(strPEFile) Then
        strErr = "未找到需要注册/反注册的文件"
        Exit Function
    End If
    'PE文件大小小于1000字节就不是有效PE
    If gobjFSO.GetFile(strPEFile).Size < 1000 Then
        strErr = "文件格式无效(小于1KB),不是有效的PE文件"
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
            strErr = "文件格式无效(缺失PE关键字),不是有效的PE文件"
            Close #1
            Exit Function
        End If
    Else
        strErr = "文件格式无效(缺失MZ关键字),不是有效的PE文件"
        Close #1
        Exit Function
    End If
    mstrPEFile = strPEFile
    '判断是否是DLL
    If (inhNTHeader.FileHeader.Characteristics And (2 ^ IMAGE_FILE_DLL)) = 0 Then
        strErr = "不是有效的DLL Library文件"
    Else
        mblnIsDLL = True
    End If
    '判断是否有TypeLib信息
    If inhNTHeader.OptionalHeader.DataDirecoty(DATA_DIRECTORY_RESOURCE_TABLE).VirtualAddress = 0 And inhNTHeader.OptionalHeader.DataDirecoty(DATA_DIRECTORY_RESOURCE_TABLE).iSize = 0 Then
    Else
        rsNum = -1
        For i = 0 To inhNTHeader.FileHeader.NumberOfSections - 1
            '计算资源目录所在节编号
            '判断资源目录所在的虚拟地址最终落在那一个节的范围
            If inhNTHeader.OptionalHeader.DataDirecoty(DATA_DIRECTORY_RESOURCE_TABLE).VirtualAddress >= ishSections(i).VirtualAddress And _
                inhNTHeader.OptionalHeader.DataDirecoty(DATA_DIRECTORY_RESOURCE_TABLE).VirtualAddress <= CDbl(ishSections(i).VirtualAddress) + ishSections(i).SizeOfRawData Then
                rsNum = i
            End If
        Next
'        If rsNum < 0 Then rsNum = 0
        If rsNum <> -1 Then
            '计算资源目录相对节的文件偏移地址+计算资源目录在文件中的相对访问地址
            lngTmpAddr = inhNTHeader.OptionalHeader.DataDirecoty(DATA_DIRECTORY_RESOURCE_TABLE).VirtualAddress - ishSections(rsNum).VirtualAddress + ishSections(rsNum).PointerToRawData + 1
            Get #1, lngTmpAddr, irdResDir
            If irdResDir.NumberOfNamedEntries + irdResDir.NumberOfIdEntries <> 0 Then
                lngTmpAddr = inhNTHeader.OptionalHeader.DataDirecoty(DATA_DIRECTORY_RESOURCE_TABLE).VirtualAddress - ishSections(rsNum).VirtualAddress + ishSections(rsNum).PointerToRawData + 1
                lngTmpAddr1 = lngTmpAddr + Len(irdResDir)
                ReDim irdeEntries(irdResDir.NumberOfNamedEntries + irdResDir.NumberOfIdEntries - 1)
                Get #1, lngTmpAddr1, irdeEntries
                '读取资源名称
                For i = LBound(irdeEntries) To UBound(irdeEntries)
                    '如果最高位为1，表示自定义资源，为0表示资源类型的id号
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
                            '存在TYPELIB
                            mblnIsActivexEXE = True
                            Exit For
                        End If
                    End If
                Next
            End If
        End If
    End If
    '获取CLR版本
    If inhNTHeader.OptionalHeader.DataDirecoty(DATA_DIRECTORY_CLR_RUNTIME_HEADER).VirtualAddress = 0 Then
    Else
        rsNum = -1
        For i = 0 To inhNTHeader.FileHeader.NumberOfSections - 1
            '计算资源目录所在节编号
            '判断资源目录所在的虚拟地址最终落在那一个节的范围
            If inhNTHeader.OptionalHeader.DataDirecoty(DATA_DIRECTORY_CLR_RUNTIME_HEADER).VirtualAddress >= ishSections(i).VirtualAddress And _
                inhNTHeader.OptionalHeader.DataDirecoty(DATA_DIRECTORY_CLR_RUNTIME_HEADER).VirtualAddress <= CDbl(ishSections(i).VirtualAddress) + ishSections(i).SizeOfRawData Then
                rsNum = i
            End If
        Next
'        If rsNum < 0 Then rsNum = 0
        If rsNum <> -1 Then
            mblnIsNetLibrary = True
            '计算资源目录相对节的文件偏移地址+计算资源目录在文件中的相对访问地址
            lngTmpAddr = inhNTHeader.OptionalHeader.DataDirecoty(DATA_DIRECTORY_CLR_RUNTIME_HEADER).VirtualAddress - ishSections(rsNum).VirtualAddress + ishSections(rsNum).PointerToRawData + 1
            Get #1, lngTmpAddr, ichCLR
            lngTmpAddr = ichCLR.metaData.VirtualAddress - ishSections(rsNum).VirtualAddress + ishSections(rsNum).PointerToRawData + 1
            Get #1, lngTmpAddr, cmdvMetaData
            mstrCLRVer = TruncZero(StrConv(cmdvMetaData.VersionString, vbUnicode))
        Else
            strErr = "未能读取CLR头所在的节信息区域"
        End If
    End If
    Close #1
    '非DLL但是可以是ActivexEXE
    If Not mblnIsDLL And mblnIsActivexEXE Then
        strErr = ""
    End If
    LoadPEFile = mblnIsDLL Or mblnIsActivexEXE
    gobjTrace.WriteInfo "LoadPEFile", "DLL", mblnIsDLL, "ActiveEXE", mblnIsActivexEXE, "NET", mblnIsNetLibrary, "CLRVer", mstrCLRVer, "ERROR", strErr
    Exit Function
ErrH:
    gobjTrace.WriteInfo "LoadPEFile", "PE文件分析发生错误", Err.Description
    Err.Clear
End Function

Public Function Is64BitPE(ByVal strPEFile As String, Optional ByRef strErr As String) As Boolean
'功能：判断一个PE文件是否是64位PE
    Dim idhDosHeader As IMAGE_DOS_HEADER 'DOS头
    Dim inhNTHeader As IMAGE_NT_HEADERS 'NT头
    
    If Not gobjFSO.FileExists(strPEFile) Then
        strErr = "未找到需要注册/反注册的文件"
        Exit Function
    End If
    'PE文件大小小于1000字节就不是有效PE
    If gobjFSO.GetFile(strPEFile).Size < 1000 Then
        strErr = "文件格式无效(小于1KB),不是有效的PE文件"
        Exit Function
    End If
    On Error GoTo ErrH
    Open strPEFile For Binary As #1
    Get #1, , idhDosHeader
    If idhDosHeader.e_magic(0) = Asc("M") And idhDosHeader.e_magic(1) = Asc("Z") Then
        Get #1, idhDosHeader.e_lfanew + 1, inhNTHeader
        If inhNTHeader.Signature(0) = Asc("P") And inhNTHeader.Signature(1) = Asc("E") Then
            ReDim ishSections(inhNTHeader.FileHeader.NumberOfSections - 1)
            Get #1, , ishSections
        Else
            strErr = "文件格式无效(缺失PE关键字),不是有效的PE文件"
            Close #1
            Exit Function
        End If
    Else
        strErr = "文件格式无效(缺失MZ关键字),不是有效的PE文件"
        Close #1
        Exit Function
    End If
    Close #1
    Is64BitPE = inhNTHeader.FileHeader.Machine = IA64 Or inhNTHeader.FileHeader.Machine = AMD64
    Exit Function
ErrH:
    strErr = Err.Description
    Err.Clear
End Function

