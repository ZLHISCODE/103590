Attribute VB_Name = "mdlSever"
Option Explicit

Private Declare Function GetFileVersionInfoSize Lib "version.dll" Alias "GetFileVersionInfoSizeA" (ByVal lptstrFilename As String, lpdwHandle As Long) As Long
Private Declare Function GetFileVersionInfo Lib "version.dll" Alias "GetFileVersionInfoA" (ByVal lptstrFilename As String, ByVal dwHandle As Long, ByVal dwLen As Long, lpData As Any) As Long
Private Declare Function VerQueryValue Lib "version.dll" Alias "VerQueryValueA" (ByVal pBlock As Long, ByVal lpSubBlock As String, lplpBuffer As Long, puLen As Long) As Long

Public Const FVN_Comments           As String = "Comments"          '注释
Public Const FVN_InternalName       As String = "InternalName"      '内部名称
Public Const FVN_ProductName        As String = "ProductName"       '产品名
Public Const FVN_CompanyName        As String = "CompanyName"       '公司名
Public Const FVN_ProductVersion     As String = "ProductVersion"    '产品版本
Public Const FVN_FileDescription    As String = "FileDescription"   '文件描述
Public Const FVN_OriginalFilename   As String = "OriginalFilename"  '原始文件名
Public Const FVN_FileVersion        As String = "FileVersion"       '文件版本
Public Const FVN_SpecialBuild       As String = "SpecialBuild"      '特殊编译号
Public Const FVN_PrivateBuild       As String = "PrivateBuild"      '私有编译号
Public Const FVN_LegalCopyright     As String = "LegalCopyright"    '合法版权
Public Const FVN_LegalTrademarks    As String = "LegalTrademarks"   '合法商标

Private Declare Function GetLastError Lib "kernel32" () As Long
Private Declare Function FormatMessage Lib "kernel32.dll" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long
Private Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000
Private Const FORMAT_MESSAGE_IGNORE_INSERTS = &H200
Private Declare Function WNetGetLastError Lib "mpr.dll" Alias "WNetGetLastErrorA" (lpError As Long, ByVal lpErrorBuf As String, ByVal nErrorBufSize As Long, ByVal lpNameBuf As String, ByVal nNameBufSize As Long) As Long
Private Const ERROR_EXTENDED_ERROR          As Long = 1208

'系统判断
Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Private Declare Function IsWow64Process Lib "kernel32" (ByVal hProc As Long, bWow64Process As Boolean) As Long

Public Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long

Public Enum LogTimeType                                 '日志时间类型
    LTT_None = 0                                        '不添加时间
    LTT_FullDate = 1                                    '全日期时间格式
    LTT_OnlyTime = 2                                    '只有时间
End Enum

Public Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Const NORMAL_PRIORITY_CLASS             As Long = &H20&
Private Const STARTF_USESTDHANDLES              As Long = &H100&
Private Const STARTF_USESHOWWINDOW              As Long = &H1
Private Const SW_HIDE                           As Integer = 0 '隐藏窗口，激活另一个窗口
Public Const INFINITE                           As Long = &HFFFF&
Private Declare Sub MDFile Lib "aamd532.dll" (ByVal f As String, ByVal R As String)
Private Const ALG_CLASS_HASH = 32768
Private Const ALG_TYPE_ANY = 0
Private Const ALG_SID_MD2 = 1
Private Const ALG_SID_MD4 = 2
Private Const ALG_SID_MD5 = 3
Private Const ALG_SID_SHA = 4
Private Enum HashAlgorithm
    HA_MD2 = ALG_CLASS_HASH Or ALG_TYPE_ANY Or ALG_SID_MD2
    HA_MD4 = ALG_CLASS_HASH Or ALG_TYPE_ANY Or ALG_SID_MD4
    HA_MD5 = ALG_CLASS_HASH Or ALG_TYPE_ANY Or ALG_SID_MD5
    HA_SHA = ALG_CLASS_HASH Or ALG_TYPE_ANY Or ALG_SID_SHA
End Enum
Private Type LARGE_INTEGER
    lowpart As Long
    highpart As Long
End Type

Public gobjFSO As New FileSystemObject
Public mstr7ZPath As String '7z地址初始化

Private mobjTrace           As TextStream               '跟踪对象
Public glngItemCout         As Long                     '日志条目长度

Private Function Byte2Hex(bytArray() As Byte, Optional ByVal lngStart As Long = 0, Optional ByVal lngEnd As Long = -1, Optional fReversed As Boolean = False) As String
    Dim i     As Long
    lngStart = IIf(lngStart < 0, 0, lngStart)
    lngEnd = IIf(lngEnd < 0, UBound(bytArray), lngEnd)
    
    If fReversed Then
        For i = lngEnd To lngStart Step -1
            Byte2Hex = Byte2Hex & Right$("00" & Hex(bytArray(i)), 2)
        Next
    Else
        For i = lngStart To lngEnd
            Byte2Hex = Byte2Hex & Right$("00" & Hex(bytArray(i)), 2)
        Next
    End If
End Function

Public Function GetVersionInfo(ByVal strFileName As String, ByVal strEntryName As String) As String
    Dim i               As Long
    Dim lngVerSize      As Long
    Dim bytVerBlock()   As Byte
    Dim strSubBlock  As String
    Dim bytTranslate()  As Byte, lngAdrTranslate    As Long, lngTranslateSize       As Long
    Dim bytBuffer()     As Byte, lngBuffer          As Long, lngAdrBuffer           As Long
    
    On Error GoTo errH
    If Not gobjFSO.FileExists(strFileName) Then Exit Function
    lngVerSize = GetFileVersionInfoSize(strFileName, 0&)
    If lngVerSize <= 0 Then Exit Function
    
    ReDim bytVerBlock(lngVerSize - 1)
    Call GetFileVersionInfo(strFileName, 0&, lngVerSize, bytVerBlock(0))
    
    VerQueryValue VarPtr(bytVerBlock(0)), "\\VarFileInfo\\Translation", lngAdrTranslate, lngTranslateSize
    ReDim bytTranslate(lngTranslateSize - 1)
    CopyMemory bytTranslate(0), ByVal lngAdrTranslate, lngTranslateSize
    For i = 1 To lngTranslateSize / (UBound(bytTranslate) + 1)
        strSubBlock = "\\StringFileInfo\\"
        strSubBlock = strSubBlock & Byte2Hex(bytTranslate(), 0, 1, True)
        strSubBlock = strSubBlock & Byte2Hex(bytTranslate(), 2, 3, True)
        strSubBlock = strSubBlock & "\\" & strEntryName
        
        VerQueryValue VarPtr(bytVerBlock(0)), strSubBlock, lngAdrBuffer, lngBuffer
        If lngAdrBuffer <> 0 And lngBuffer <> 0 Then
            ReDim bytBuffer(lngBuffer - 1)
            CopyMemory bytBuffer(0), ByVal lngAdrBuffer, lngBuffer
            ReDim Preserve bytBuffer(InStrB(bytBuffer, ChrB(0)) - 2)
            GetVersionInfo = StrConv(bytBuffer, vbUnicode)
        End If
    Next
    Exit Function
errH:
    err.Clear
End Function

Public Function GetLastDllErr(Optional ByVal lngErr As Long) As String
    Dim strReturn As String
    If lngErr = 0 Then
        lngErr = GetLastError
    End If
    If lngErr = ERROR_EXTENDED_ERROR Then
        GetLastDllErr = GetWNetErr(lngErr)
    Else
        strReturn = String$(256, 32)
        FormatMessage FORMAT_MESSAGE_FROM_SYSTEM Or FORMAT_MESSAGE_IGNORE_INSERTS, 0&, lngErr, 0&, strReturn, Len(strReturn), ByVal 0
        strReturn = Trim(strReturn)
        GetLastDllErr = Replace(Replace(strReturn, Chr(10), ""), Chr(13), "")
    End If
End Function

Private Function GetWNetErr(ByVal lngErr As Long) As String
    Dim strErr As String * 256
    Dim strName As String * 256
    Dim lngRet As Long
    lngRet = WNetGetLastError(lngErr, strErr, Len(strErr), strName, Len(strName))
    GetWNetErr = Replace(Replace("[" & TruncZero(strName) & "]" & TruncZero(strErr), Chr(10), ""), Chr(13), "")
End Function


'日志方法
Public Sub WriteLog(Optional ByVal strText As String, Optional ByVal lttAddTime As LogTimeType = LTT_None, Optional ByVal lngReturnLines As Long)
'参数:strText       =要写入的一行日志文本,如果为空表示写一行换行符
'     bytAddTime    >0时，在日志文本之前加入日志时间，1=日期和时间完整格式,2-仅时间,0-不加入
'     lngReturnLines=在日志文本之后写多少行换行符,0-不写换行符
    If Not mobjTrace Is Nothing Then
        '添加日志时间
        If lttAddTime <> LTT_None Then strText = LogTime(lttAddTime = LTT_OnlyTime) & strText
        
        '截取超长内容
        If Len(strText) > 500 Then strText = Mid(strText, 1, 500)
                
        '写日志文本
        If strText = "" Then
             mobjTrace.WriteBlankLines 1
        Else
            mobjTrace.WriteLine strText
            If lngReturnLines > 0 Then mobjTrace.WriteBlankLines lngReturnLines
        End If
    End If
End Sub

Private Function LogTime(Optional blnOnlyTime As Boolean) As String
    If blnOnlyTime Then
        LogTime = Format(Now, "HH:mm:ss")
    Else
        LogTime = Format(Now, "yyyy-MM-dd HH:mm:ss")
    End If
End Function

Public Function Init7Z() As Boolean
    Dim blnIs64Bits As Boolean
    Dim strSystemPath As String
    
    blnIs64Bits = Is64bit
    
    strSystemPath = gobjFSO.GetSpecialFolder(SystemFolder)
    
    If blnIs64Bits Then '64系统下32位程序应该放在C:\windows\SysWOW64
        strSystemPath = gobjFSO.GetParentFolderName(strSystemPath) & "\SysWOW64"
    End If
    
    Init7Z = False
    
'    mstr7ZPath = GetWinSystemPath & "\7z.dll"
    mstr7ZPath = strSystemPath & "\7z.dll"
    If gobjFSO.FileExists(mstr7ZPath) = False Then
        MsgBox "压缩文件7z.dll不存在,请手动放入系统目录下!", vbInformation + vbDefaultButton1, gstrSysName
        Exit Function
    End If
    
    mstr7ZPath = strSystemPath & "\7z.exe"
    If gobjFSO.FileExists(mstr7ZPath) = False Then
        MsgBox "压缩文件7z.exe不存在,请手动放入系统目录下!", vbInformation + vbDefaultButton1, gstrSysName
        Exit Function
    End If
    Init7Z = True
End Function

'SIZE是每次影射的文件大小 只能是2的N次方  如: 2^27=2的27次方=128M
Public Function FileMD5(ByVal szFilePath As String, Optional ByVal haCur As Long = HA_MD5, Optional ByVal Block_Size As Long = 32768) As String
    Dim lnghFile As Long, lnghMapFile As Long, lnglpBaseMap As Long
    Dim lnghCtx As Long, lngRet As Long, lnghHash As Long, lngLen As Long
    Dim i As Long, j As Long, lngPoint As Long
    Dim lintFI As LARGE_INTEGER, lintCurrent As LARGE_INTEGER, dblCurrentPoint As Double
    Dim lngTmp As Long, lngBlocks As Long, lngLastBlock As Long, Block() As Byte
    Dim lngSize As Long
    '创建文件指针
    On Error GoTo errH
    lngSize = 2 ^ 27
    lnghFile = CreateFileA(szFilePath, GENERIC_READ, FILE_SHARE_READ, ByVal 0&, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0)
    If lnghFile <> INVALID_HANDLE_VALUE Then
        lintFI.lowpart = GetFileSize(lnghFile, lintFI.highpart) '成功后 获取文件大小
        If lintFI.highpart > 0 Then lngBlocks = ((2 ^ 32 / lngSize) * lintFI.highpart) ' 高位   为1就是 2^32次字节  也就是4字节无符号长整型数值
        If lintFI.lowpart < 0 Then        '低位
            lngBlocks = lngBlocks + (2 ^ 31 / lngSize) '低位为负数 必然大于2^31次方  因为不大于2^31  VB可以正常显示
            lngTmp = LongToUnsigned(lintFI.lowpart) - 2 ^ 31 '转为无符号整型减掉2^31次 VB就能正常显示和运算了
            lngLastBlock = lngTmp \ lngSize
            lngBlocks = lngBlocks + lngLastBlock
            lngLastBlock = lngTmp - lngLastBlock * lngSize
        Else
            lngTmp = lintFI.lowpart \ lngSize
            lngBlocks = lngBlocks + lngTmp
            lngLastBlock = lintFI.lowpart - lngTmp * lngSize
        End If
        
        lnghMapFile = CreateFileMapping(lnghFile, ByVal 0&, PAGE_READONLY, lintFI.highpart, lintFI.lowpart, 0) '创建文件映射对象
        lngRet = CryptAcquireContextA(lnghCtx, vbNullString, vbNullString, PROV_RSA_FULL, 0)
        If err.LastDllError = &H80090016 Then lngRet = CryptAcquireContextA(lnghCtx, vbNullString, vbNullString, PROV_RSA_FULL, CRYPT_NEWKEYSET)
        lngRet = CryptCreateHash(lnghCtx, haCur, 0, 0, lnghHash)
        ReDim Block(Block_Size) As Byte
        
        For i = 1 To lngBlocks '成功后根据指定大小 开始影射文件到内存空间
            lnglpBaseMap = MapViewOfFile(lnghMapFile, FILE_MAP_READ, lintCurrent.highpart, lintCurrent.lowpart, lngSize)
            If lnglpBaseMap Then
                lngPoint = lnglpBaseMap
                For j = 1 To lngSize / Block_Size ' 2的N次方  必然除尽
                    
                    lngRet = CryptHashData(lnghHash, lngPoint, Block_Size, 0)
                    lngPoint = lngPoint + Block_Size
                Next
                UnmapViewOfFile (lnglpBaseMap)
            End If
            dblCurrentPoint = dblCurrentPoint + lngSize
            lintCurrent = Currency2LargeInteger(dblCurrentPoint / 10000@) '设置文件高低位
        Next
            
        If lngLastBlock > 0 Then '映射余数
            lnglpBaseMap = MapViewOfFile(lnghMapFile, FILE_MAP_READ, lintCurrent.highpart, lintCurrent.lowpart, lngLastBlock)
            If lnglpBaseMap Then
                lngPoint = lnglpBaseMap
                lngTmp = lngLastBlock \ Block_Size '不一定除尽 余数在FOR 循环完再次计算
                
                For j = 1 To lngTmp
                    lngRet = CryptHashData(lnghHash, lngPoint, Block_Size, 0)
                    lngPoint = lngPoint + Block_Size
                Next
                lngTmp = lngLastBlock - lngTmp * Block_Size
                lngRet = CryptHashData(lnghHash, lngPoint, lngTmp, 0)
                UnmapViewOfFile (lnglpBaseMap)
            End If
        End If
        Call CloseHandle(lnghMapFile)
        If lngRet Then
            lngRet = CryptGetHashParam(lnghHash, HP_HASHSIZE, lngLen, 4, 0)
            If lngRet Then
                ReDim hash(lngLen) As Byte
                lngRet = CryptGetHashParam(lnghHash, HP_HASHVAL, hash(0), lngLen, 0)
                If lngRet Then
                    For j = 0 To UBound(hash) - 1
                        FileMD5 = FileMD5 & Right$("0" & Hex$(hash(j)), 2)
                    Next
                End If
                CryptDestroyHash lnghHash
            End If
        End If
        CryptReleaseContext lnghCtx, 0
        CloseHandle (lnghFile)
        
        If FileMD5 = "" Then
            FileMD5 = MD5File(szFilePath)
        End If
    End If
    Exit Function
errH:
    If 0 = 1 Then
        Resume
    End If
End Function

Private Function Currency2LargeInteger(ByVal curDistance As Currency) As LARGE_INTEGER
    CopyMemory Currency2LargeInteger, curDistance, 8
End Function

Private Function MD5File(f As String) As String
    Dim R As String * 32
    R = Space(32)
    MDFile f, R
    MD5File = UCase(R)
End Function

Private Function LongToUnsigned(value As Long) As Double
    If value < 0 Then
        LongToUnsigned = value + 2 ^ 32
    Else
        LongToUnsigned = value
    End If
End Function

'密码加密程序
Public Function Cipher(ByVal strText As String) As String
    Const MIN_ASC = 32    '最小ASCII码
    Const MAX_ASC = 126 '最大ASCII码 字符
    Const NUM_ASC = MAX_ASC - MIN_ASC + 1
    Dim lngOffset As Long, intLen As Integer, intSeedLen As Integer
    Dim i As Integer, intChr As Integer
    Dim strDeText As String
    Dim strSeed As String
    
    If strText = "" Then Exit Function
    '获取随机种子
    '随机种子的随机数为999
    Rnd (-1)
    Randomize (999)
    strSeed = "456"
    intSeedLen = Len(strSeed)
    strDeText = Chr(intSeedLen + MIN_ASC)
    For i = 1 To intSeedLen
        intChr = Asc(Mid(strSeed, i, 1)) '取字母转变成ASCII码
        If intChr >= MIN_ASC And intChr <= MAX_ASC Then
            intChr = intChr - MIN_ASC
            lngOffset = Int((NUM_ASC + 1) * Rnd())
            intChr = ((intChr + lngOffset) Mod NUM_ASC)
            intChr = intChr + MIN_ASC
            strDeText = strDeText & Chr(intChr)
        End If
    Next
    Rnd (-1)
    Randomize (Val(strSeed))
    intLen = Len(strText)
    For i = 1 To intLen
        intChr = Asc(Mid(strText, i, 1)) '取字母转变成ASCII码
        If intChr >= MIN_ASC And intChr <= MAX_ASC Then
            intChr = intChr - MIN_ASC
            lngOffset = Int((NUM_ASC + 1) * Rnd())
            intChr = ((intChr + lngOffset) Mod NUM_ASC)
            intChr = intChr + MIN_ASC
            strDeText = strDeText & Chr(intChr)
        ElseIf intChr < 0 Then '非ASCII字符的处理,如中文，中文不加密
            strDeText = strDeText & Mid(strText, i, 1)
        End If
    Next
    Cipher = strDeText
End Function

Public Function Decipher(ByVal strText As String) As String
'密码解密程序
    Const MIN_ASC = 32    '最小ASCII码
    Const MAX_ASC = 126 '最大ASCII码 字符
    Const NUM_ASC = MAX_ASC - MIN_ASC + 1
    Dim lngOffset As Long, intLen As Integer, intSeedLen As Integer
    Dim intStart As Integer
    Dim i As Integer, intChr As Integer
    Dim strDeText As String
    
    If strText = "" Then Exit Function
    '随机种子长度
    intSeedLen = Asc(Mid(strText, 1, 1)) - MIN_ASC
    intLen = Len(strText)
    '采用旧的随机算法
    If intSeedLen > 0 And intSeedLen < intLen - 3 And intSeedLen < 5 Then
        '获取随机种子
        '随机种子的随机数为999
        Rnd (-1)
        Randomize (999)
        For i = 2 To 1 + intSeedLen
            intChr = Asc(Mid(strText, i, 1)) '取字母转变成ASCII码
            If intChr >= MIN_ASC And intChr <= MAX_ASC Then
                intChr = intChr - MIN_ASC
                lngOffset = Int((NUM_ASC + 1) * Rnd())
                intChr = ((intChr - lngOffset) Mod NUM_ASC)
                If intChr < 0 Then
                    intChr = intChr + NUM_ASC
                End If
                intChr = intChr + MIN_ASC
                strDeText = strDeText & Chr(intChr)
            End If
        Next
        If Not IsNumeric(strDeText) Then
            strDeText = "123"
            intStart = 1
        Else
            intStart = 2 + intSeedLen
        End If
    Else
        strDeText = "123"
        intStart = 1
    End If
        
    '内容解密的种子
    Rnd (-1)
    Randomize (Val(strDeText))
    strDeText = ""
    For i = intStart To intLen
        intChr = Asc(Mid(strText, i, 1)) '取字母转变成ASCII码
        If intChr >= MIN_ASC And intChr <= MAX_ASC Then
            intChr = intChr - MIN_ASC
            lngOffset = Int((NUM_ASC + 1) * Rnd())
            intChr = ((intChr - lngOffset) Mod NUM_ASC)
            If intChr < 0 Then
                intChr = intChr + NUM_ASC
            End If
            intChr = intChr + MIN_ASC
            strDeText = strDeText & Chr(intChr)
        Else '非ASCII字符的处理,如中文，中文不加密
            strDeText = strDeText & Mid(strText, i, 1)
        End If
    Next
    Decipher = strDeText
End Function

Public Function CheckAndAdjustFolder() As Collection
'功能：进行安装路径的修复
    Dim strSQL              As String, rsTmp        As ADODB.Recordset
    Dim strPath             As String, arrTmp       As Variant
    Dim i                   As Integer
    Dim strSystemPath As String
    Dim blnIs64Bits As Boolean
    Dim cllPaths As New Collection
    Dim strAppPath As String
    
    If gblnInIDE Then
        strAppPath = "C:\APPSOFT"
    Else
        strAppPath = App.Path
    End If
    
    On Error GoTo errH
    blnIs64Bits = Is64bit
    
    strSystemPath = gobjFSO.GetSpecialFolder(SystemFolder)
    
    If blnIs64Bits Then '64系统下32位程序应该放在C:\windows\SysWOW64
        strSystemPath = gobjFSO.GetParentFolderName(strSystemPath) & "\SysWOW64"
    End If
    
    strSQL = "Select Distinct Upper(安装路径) 安装路径 From zltools.Zlfilesupgrade union Select Distinct Upper(安装路径) 安装路径 From zltools.zlfiles"
    Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "获取路径文件夹")
    
    Do While Not rsTmp.EOF
        arrTmp = Split(rsTmp!安装路径 & "", "\")
'        arrTmp = Split("[APPSOFT] \ PACSLIST", "\")
        strPath = ""
        If UBound(arrTmp) <> -1 Then
            arrTmp(0) = Trim(arrTmp(0))
            If arrTmp(0) = "[APPSOFT]" Then
'                strPath = gstrAppPath
                strPath = strAppPath
            ElseIf arrTmp(0) = "[PUBLIC]" Then
                If Not gobjFSO.FolderExists(strAppPath & "\PUBLIC") Then
                    gobjFSO.CreateFolder (strAppPath & "\PUBLIC")
                End If
                strPath = strAppPath & "\PUBLIC"
            ElseIf arrTmp(0) = "[APPLY]" Then
                strPath = strAppPath & "\APPLY"
            ElseIf arrTmp(0) = "[OS:]" Then '系统盘
                strPath = Left(strSystemPath, 2)
            ElseIf arrTmp(0) = "[X:]" Then '当前安装盘
                strPath = Left(strAppPath, 2)
            ElseIf Not arrTmp(0) Like "[[]*[]]" Then
                cllPaths.Add rsTmp!安装路径 & "", "K_" & rsTmp!安装路径
            End If
            If strPath <> "" Then
                For i = 1 To UBound(arrTmp)
                    If arrTmp(i) <> "" Then
                        strPath = strPath & "\" & arrTmp(i)
                        If Not gobjFSO.FolderExists(strPath) Then
                            gobjFSO.CreateFolder (strPath)
                        End If
                    End If
                Next
                '缓存安装路径，优化转换速度。
                cllPaths.Add strPath, "K_" & rsTmp!安装路径
            End If
        End If
        rsTmp.MoveNext
    Loop
    '缓存基础安装路径，优化转换速度。
    On Error Resume Next
    cllPaths.Add strAppPath, "K_[APPSOFT]"
    cllPaths.Add strAppPath & "\PUBLIC", "K_[PUBLIC]"
    cllPaths.Add strAppPath & "\APPLY", "K_[APPLY]"
    cllPaths.Add Left(strSystemPath, 2), "K_[OS:]"
'    cllPaths.Add Left(strAppPath, 2), "K_[X:]"
    cllPaths.Add strSystemPath, "K_[SYSTEM]"
    cllPaths.Add strSystemPath, "K_[HELP]"
'    cllPaths.Add strSystemPath, "K_[APPSOFT]\APPLY"
'    If Not gobjFSO.FolderExists(gstrTempPath) Then
'        Call gobjFSO.CreateFolder(gstrTempPath)
'    End If
    If err.Number Then err.Clear
    Set CheckAndAdjustFolder = cllPaths
    Exit Function
errH:
    MsgBox err.Description, vbInformation, "中联软件"
'    Call RecordErrMsg(MT_InitEnv, "修复安装目录", err.Description)
    If 0 = 1 Then
        Resume
    End If
End Function

Public Sub InitTable(vsgInfo As VSFlexGrid, ByVal strHead As String)
    Dim arrHead As Variant, i As Long
    
    arrHead = Split(strHead, ";")
    With vsgInfo
        .Clear
        .FixedRows = 1
        .FixedCols = 0
        .Cols = .FixedCols + UBound(arrHead) + 1
        .Rows = .FixedRows + 1
        
        For i = 0 To UBound(arrHead)
            .TextMatrix(.FixedRows - 1, .FixedCols + i) = Split(arrHead(i), ",")(0)
        .ColKey(.FixedCols + i) = Split(arrHead(i), ",")(0)

            If UBound(Split(arrHead(i), ",")) > 0 Then
                .ColHidden(.FixedCols + i) = False
                .ColWidth(.FixedCols + i) = Val(Split(arrHead(i), ",")(1))
                .ColAlignment(.FixedCols + i) = Val(Split(arrHead(i), ",")(2))
            Else
                .ColHidden(.FixedCols + i) = True
                .ColWidth(.FixedCols + i) = 0
            End If
        Next
        .Cell(flexcpAlignment, 0, 0, .FixedRows - 1, .Cols - 1) = 4
    End With
End Sub

Public Function GetDealVersion(ByVal strFile As String) As String '处理后版本
    Dim objFile As New FileSystemObject
    Dim strVer As String, varVersion As Variant
    Dim strSPVer As String

    err = 0: On Error Resume Next
    '获取文件版本号
    
    strSPVer = GetVersionInfo(strFile, FVN_FileDescription)
    If IsVerSion(strSPVer) = False Then
        strVer = gobjFile.GetFileVersion(strFile)
        If err <> 0 Then
            err.Clear: err = 0
            GetDealVersion = ""
            Exit Function
        End If
        GetDealVersion = VersionCheck(gobjFile.GetFileName(strFile), strVer)
    Else
        GetDealVersion = strSPVer
    End If
End Function

Public Function VersionCheck(ByVal strFileName As String, ByVal strVersion As String) As String
    Dim strTemp As String
    Dim arrVersion() As String

    If strVersion = "" Or strFileName = "" Then VersionCheck = strVersion: Exit Function
    
    strTemp = UCase(Mid(Trim(strFileName), 1, 2))
    If strTemp = "ZL" Then
        arrVersion = Split(strVersion, ".")
        If UBound(arrVersion) = 3 Then
            strVersion = arrVersion(0) & "." & arrVersion(1) & "." & arrVersion(3)
        End If
    End If
    VersionCheck = strVersion
End Function
