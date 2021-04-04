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
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private marrishSections()   As IMAGE_SECTION_HEADER '段数组
Private midhDosHeader       As IMAGE_DOS_HEADER
Private minhNTHeader        As IMAGE_NT_HEADERS
    
Private mblnDLL         As Boolean
Private mblnEXE         As Boolean
Private mbln64BitPE     As Boolean
Private mblnHaveTypeLib As Boolean
Private mblnNetLib      As Boolean
Private mstrNetVersion  As String
Private mblnIsVBCom     As Boolean

Private mstrFile        As String
Private mstrErr         As String

Property Get CurFile() As String
    CurFile = mstrFile
End Property

Property Get CurErrDes() As String
    CurErrDes = mstrErr
End Property

Property Get IsCom(Optional ByVal strFile As String) As Boolean
    Dim strErr As String
    If strFile <> "" And mstrFile <> strFile Then
        Call LoadFile(strFile, strErr)
        mstrFile = strFile
        mstrErr = strErr
    End If
    IsCom = mblnDLL And mblnHaveTypeLib Or mblnEXE And mblnHaveTypeLib
End Property

Property Get IsHaveTypeLib(Optional ByVal strFile As String) As Boolean
    Dim strErr As String
    If strFile <> "" And mstrFile <> strFile Then
        Call LoadFile(strFile, strErr)
        mstrFile = strFile
        mstrErr = strErr
    End If
    IsHaveTypeLib = mblnHaveTypeLib
End Property

Property Get IsActivexDLL(Optional ByVal strFile As String) As Boolean
    Dim strErr As String
    If strFile <> "" And mstrFile <> strFile Then
        Call LoadFile(strFile, strErr)
        mstrFile = strFile
        mstrErr = strErr
    End If
    IsActivexDLL = mblnDLL And mblnHaveTypeLib
End Property

Property Get IsActivexEXE(Optional ByVal strFile As String) As Boolean
    Dim strErr As String
    If strFile <> "" And mstrFile <> strFile Then
        Call LoadFile(strFile, strErr)
        mstrFile = strFile
        mstrErr = strErr
    End If
    IsActivexEXE = mblnEXE And mblnHaveTypeLib
End Property

Property Get IsNetLibrary(Optional ByVal strFile As String) As Boolean
    Dim strErr As String
    If strFile <> "" And mstrFile <> strFile Then
        Call LoadFile(strFile, strErr)
        mstrFile = strFile
        mstrErr = strErr
    End If
    IsNetLibrary = mblnNetLib
End Property

Property Get CLRVer(Optional ByVal strFile As String) As String
    Dim strErr As String
    If strFile <> "" And mstrFile <> strFile Then
        Call LoadFile(strFile, strErr)
        mstrFile = strFile
        mstrErr = strErr
    End If
    CLRVer = mstrNetVersion
End Property

Property Get Is64bit(Optional ByVal strFile As String) As Boolean
    Dim strErr As String
    If strFile <> "" And mstrFile <> strFile Then
        Call LoadFile(strFile, strErr)
        mstrFile = strFile
        mstrErr = strErr
    End If
    Is64bit = mbln64BitPE
End Property

Property Get IsVBCom(Optional ByVal strFile As String) As Boolean
    Dim strErr As String
    If strFile <> "" And mstrFile <> strFile Then
        Call LoadFile(strFile, strErr)
        mstrFile = strFile
        mstrErr = strErr
    End If
    IsVBCom = mblnIsVBCom
End Property

Public Function AnalyseFile(ByVal strPEFile As String, Optional ByRef strErr As String) As Boolean
    AnalyseFile = LoadFile(strPEFile, strErr)
    mstrFile = strPEFile
    mstrErr = strErr
End Function

Private Function LoadFile(ByVal strPEFile As String, Optional ByRef strErr As String) As Boolean
'功能：读取PE文件
    Dim lngFileSize     As Long, intFileNum     As Integer
    Dim intNum          As Integer
    Dim lngAddress      As Long, lngTmpAddress  As Long, lngTmpAddress1 As Long
    Dim i               As Long, j              As Integer, k           As Integer
    Dim strTmp          As String
    Dim intUnicodeChar  As Integer, byteChar    As Byte
    Dim cllImportDLL    As Collection
    
    Dim irdResDir       As IMAGE_RESOURCE_DIRECTORY
    Dim irdResDir1      As IMAGE_RESOURCE_DIRECTORY
    Dim irdeEntries()   As IMAGE_RESOURCE_DIRECTORY_ENTRY
    Dim irdeEntries1()  As IMAGE_RESOURCE_DIRECTORY_ENTRY
    Dim ichCLR          As IMAGE_COR20_HEADER, cmdvMetaData As CLR_MetaDataVer
    Dim iidImport       As IMAGE_IMPORT_DESCRIPTOR
    Dim irdsuResUStr    As IMAGE_RESOURCE_DIR_STRING_U
    
    mblnDLL = False
    mblnEXE = False
    mbln64BitPE = False
    mblnHaveTypeLib = False
    mblnNetLib = False
    mstrNetVersion = ""
    mblnIsVBCom = False
    On Error Resume Next
    lngFileSize = FileLen(strPEFile)
    If Err.Number <> 0 Then
        strErr = Err.Description
        Err.Clear
        Exit Function
    End If
    If lngFileSize = 0 Then
        strErr = "不是有效的PE文件(无内容结构)"
        Exit Function
    End If

    On Error GoTo ErrH
    'Len(ntHead) + Len(dosHead)=316
    If lngFileSize < 316 Then
        strErr = "不是有效的PE文件,不具有NTheader和DOSHeader"
        Exit Function
    End If
    intFileNum = FreeFile
    Open strPEFile For Binary Access Read Shared As intFileNum
    'DOSHeader读取
    Get intFileNum, 1, midhDosHeader
    If midhDosHeader.e_magic <> IMAGE_DOS_SIGNATURE Then
        strErr = "文件格式无效(无DOS头签名),不是有效的PE文件"
        Exit Function
    End If
    'NTHeader读取
    Get intFileNum, midhDosHeader.e_lfanew + 1, minhNTHeader
    
    If minhNTHeader.Signature <> IMAGE_NT_SIGNATURE Then
        strErr = "文件格式无效(无PE签名),不是有效的PE文件"
        Close intFileNum
        Exit Function
    End If
    
    If minhNTHeader.OptionalHeader.Magic = IMAGE_NT_OPTIONAL_HDR64_MAGIC Then
        strErr = "文件是64位PE文件"
        mbln64BitPE = True
        Close intFileNum
        Exit Function
    ElseIf minhNTHeader.OptionalHeader.Magic <> IMAGE_NT_OPTIONAL_HDR32_MAGIC Then
        strErr = "文件不是32位PE文件"
        Close intFileNum
        Exit Function
    End If
    '读取段表
    ReDim marrishSections(minhNTHeader.FileHeader.NumberOfSections - 1)
    Get intFileNum, , marrishSections
    '判断是否是DLL，EXE
    If (minhNTHeader.FileHeader.Characteristics And IMAGE_FILE_DLL) = IMAGE_FILE_DLL Then
        mblnDLL = True
    ElseIf (minhNTHeader.FileHeader.Characteristics And IMAGE_FILE_EXECUTABLE_IMAGE) = IMAGE_FILE_EXECUTABLE_IMAGE Then
        mblnEXE = True
    End If
    
    '定位资源目录
    intNum = GetDataSectionsNo(IMAGE_DIRECTORY_ENTRY_RESOURCE, lngAddress)
    If intNum <> -1 Then
        Get intFileNum, lngAddress, irdResDir
        lngTmpAddress = lngAddress + Len(irdResDir)
        If irdResDir.NumberOfIdEntries > 0 Or irdResDir.NumberOfNamedEntries > 0 Then
            ReDim irdeEntries(irdResDir.NumberOfNamedEntries + irdResDir.NumberOfIdEntries - 1)
            Get intFileNum, lngTmpAddress, irdeEntries
            For i = LBound(irdeEntries) To UBound(irdeEntries)
                If irdeEntries(i).OffsetToData < 0 Then
                    If irdeEntries(i).Name < 0 Then
                        lngTmpAddress = lngAddress + (irdeEntries(i).Name And &H7FFFFFFF)
                        Get intFileNum, lngTmpAddress, irdsuResUStr
                        If irdsuResUStr.Length > 0 Then
                            lngTmpAddress = lngTmpAddress + Len(irdsuResUStr)
                            strTmp = ChrW(irdsuResUStr.NameString)
                            For j = 0 To irdsuResUStr.Length - 2
                              Get intFileNum, lngTmpAddress, intUnicodeChar
                              strTmp = strTmp & ChrW(intUnicodeChar)
                              lngTmpAddress = lngTmpAddress + Len(intUnicodeChar)
                            Next
                            If strTmp = "TYPELIB" Then
                                mblnHaveTypeLib = True
                                Exit For
                            End If
                        End If
                    End If
                End If
            Next
        End If
    End If
    '获取CLR版本
    intNum = GetDataSectionsNo(IMAGE_DIRECTORY_ENTRY_COM_DESCRIPTOR, lngAddress)
    If intNum <> -1 Then
        mblnNetLib = True
        Get intFileNum, lngAddress, ichCLR
        lngTmpAddress = ichCLR.MetaData.VirtualAddress - marrishSections(intNum).VirtualAddress + marrishSections(intNum).PointerToRawData + 1
        Get intFileNum, lngTmpAddress, cmdvMetaData
        mstrNetVersion = mdlPublic.TruncZero(StrConv(cmdvMetaData.VersionString, vbUnicode))
    End If
    '判断VB部件
    intNum = GetDataSectionsNo(IMAGE_DIRECTORY_ENTRY_IMPORT, lngAddress)
    If intNum <> -1 Then
        Set cllImportDLL = New Collection
        Do
            Get intFileNum, lngAddress, iidImport
            If iidImport.OriginalFirstThunk = 0 And iidImport.FirstThunk = 0 And iidImport.ForwarderChain = 0 And iidImport.Name = 0 And iidImport.TimeDateStamp = 0 Then Exit Do
            lngTmpAddress = iidImport.Name - marrishSections(intNum).VirtualAddress + marrishSections(intNum).PointerToRawData + 1
            strTmp = ""
            Do
              Get intFileNum, lngTmpAddress, byteChar
              If byteChar = 0 Then Exit Do
              strTmp = strTmp + Chr(byteChar)
              lngTmpAddress = lngTmpAddress + 1
            Loop
            cllImportDLL.Add strTmp
            lngAddress = lngAddress + Len(iidImport)
        Loop
    End If
    
    If Not cllImportDLL Is Nothing Then
        If cllImportDLL.Count = 1 Then
            mblnIsVBCom = UCase(cllImportDLL(1)) = "MSVBVM60.DLL"
        End If
    End If
    LoadFile = True
    Close intFileNum
    Exit Function
ErrH:
    strErr = Err.Description
    If 0 = 1 Then
        Resume
    End If
    Close intFileNum
End Function

Private Function GetDataSectionsNo(ByVal ideInput As Integer, ByRef lngAddress As Long) As Integer
'功能：获取数据类型的段号
'ideInput=数据类型
'返回：数据所在的节
    Dim intNum          As Integer
    Dim i               As Integer
    
    intNum = -1
    If minhNTHeader.OptionalHeader.DataDirectory(ideInput).VirtualAddress = 0 And minhNTHeader.OptionalHeader.DataDirectory(ideInput).Size = 0 Then
    Else
        For i = 0 To minhNTHeader.FileHeader.NumberOfSections - 1
            '目录所在的虚拟地址最终落在那一个节的范围
            If minhNTHeader.OptionalHeader.DataDirectory(ideInput).VirtualAddress >= marrishSections(i).VirtualAddress And _
                minhNTHeader.OptionalHeader.DataDirectory(ideInput).VirtualAddress - marrishSections(i).VirtualAddress <= marrishSections(i).SizeOfRawData Then
                intNum = i
                lngAddress = minhNTHeader.OptionalHeader.DataDirectory(ideInput).VirtualAddress - marrishSections(i).VirtualAddress + marrishSections(i).PointerToRawData + 1
                Exit For
            End If
        Next
    End If
    GetDataSectionsNo = intNum
End Function



