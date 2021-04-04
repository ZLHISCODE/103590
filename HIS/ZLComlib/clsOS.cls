VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsOS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit
'##############################################################################
'#功能：操作系统相关或底层或特定功能API封装的公共方法
'命名规则：函数过程名称为通俗易懂的英文即可
'上级类：clsComLib
'##############################################################################
'注册表关键字根类型
Private Enum REGRoot
    HKEY_CLASSES_ROOT = &H80000000 '记录Windows操作系统中所有数据文件的格式和关联信息，主要记录不同文件的文件名后缀和与之对应的应用程序。其下子键可分为两类，一类是已经注册的各类文件的扩展名，这类子键前面都有一个“。”；另一类是各类文件类型有关信息。
    HKEY_CURRENT_USER = &H80000001 '此根键包含了当前登录用户的用户配置文件信息。这些信息保证不同的用户登录计算机时，使用自己的个性化设置，例如自己定义的墙纸、自己的收件箱、自己的安全访问权限等。
    HKEY_LOCaL_MaCHINE = &H80000002 '此根键包含了当前计算机的配置数据，包括所安装的硬件以及软件的设置。这些信息是为所有的用户登录系统服务的。它是整个注册表中最庞大也是最重要的根键！
    HKEY_USERS = &H80000003 '此根键包括默认用户的信息（Default子键）和所有以前登录用户的信息。
    HKEY_PERFORMANCE_DATA = &H80000004 '在Windows NT/2000/XP注册表中虽然没有HKEY_DYN_DATA键，但是它却隐藏了一个名为“HKEY_ PERFOR MANCE_DATA”键。所有系统中的动态信息都是存放在此子键中。系统自带的注册表编辑器无法看到此键
    HKEY_CURRENT_CONFIG = &H80000005  '此根键实际上是HKEY_LOCAL_MACHINE中的一部分，其中存放的是计算机当前设置，如显示器、打印机等外设的设置信息等。它的子键与HKEY_LOCAL_ MACHINE\ Config\0001分支下的数据完全一样。
    HKEY_DYN_DATA = &H80000006 '此根键中保存每次系统启动时，创建的系统配置和当前性能信息。这个根键只存在于Windows 98中。
End Enum

'注册表数据类型
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
'打开错误
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
'注册表访问权
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
' 扩充环境字符串。具体操作过程与命令行处理的所为差不多。也就是说，将由百分号封闭起来的环境变量名转换成那个变量的内容。比如，“%path%”会扩充成完整路径。
Private Declare Function ExpandEnvironmentStrings Lib "kernel32" Alias "ExpandEnvironmentStringsA" (ByVal lpSrc As String, ByVal lpDst As String, ByVal nSize As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal uloptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegEnumKey Lib "advapi32.dll" Alias "RegEnumKeyA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, ByVal cbName As Long) As Long
Private Declare Function RegQueryValueEx_ValueType Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegQueryValueEx_Long Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Long, lpcbData As Long) As Long
Private Declare Function RegQueryValueEx_String Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
Private Declare Function RegQueryValueEx_BINARY Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Byte, lpcbData As Long) As Long
Private Declare Function RegSetValueEx_String Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpData As String, ByVal lpcbData As Long) As Long
Private Declare Function RegSetValueEx_Long Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpValue As Long, ByVal cbData As Long) As Long
Private Declare Function RegSetValueEx_BINARY Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpValue As Byte, ByVal cbData As Long) As Long
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long

'暂停(Wait)
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
'电脑名称(ComputerName)
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
'是否是64位进程（Is64bit）
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
Private Declare Function IsWow64Process Lib "kernel32" (ByVal hProc As Long, bWow64Process As Long) As Long
'临时IP获取
Private Const MAX_IP = 5                                                    'To make a buffer... i dont think you have more than 5 ip on your pc..
Private Type IPINFO
    dwAddr As Long                                                          ' IP address
    dwIndex As Long                                                         ' interface index
    dwMask As Long                                                          ' subnet mask
    dwBCastAddr As Long                                                     ' broadcast address
    dwReasmSize  As Long                                                    ' assembly size
    unused1 As Integer                                                      ' not currently used
    unused2 As Integer                                                      '; not currently used
End Type
Private Type MIB_IPADDRTABLE
    dEntrys As Long                                                         'number of entries in the table
    mIPInfo(MAX_IP) As IPINFO                                               'array of IP address entries
End Type
Private Declare Function GetIpAddrTable Lib "IPHlpApi" (pIPAdrTable As Byte, pdwSize As Long, ByVal Sort As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
'读写ini 文件的API
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal LpApplicationName As String, ByVal LpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal LpApplicationName As String, ByVal LpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
'系统颜色转换（SysColor2RGB）
Private Declare Function OleTranslateColor Lib "OLEPRO32.DLL" (ByVal OLE_COLOR As Long, ByVal HPALETTE As Long, pccolorref As Long) As Long
'返回系统中可用的输入法个数及各输入法所在Layout,包括英文输入法。
Private Declare Function GetKeyboardLayoutList Lib "user32" (ByVal nBuff As Long, lpList As Long) As Long
'获取某个输入法的名称
Private Declare Function ImmGetDescription Lib "imm32.dll" Alias "ImmGetDescriptionA" (ByVal hkl As Long, ByVal lpsz As String, ByVal uBufLen As Long) As Long
'判断某个输入法是否中文输入法
Private Declare Function ImmIsIME Lib "imm32.dll" (ByVal hkl As Long) As Long
'切换到指定的输入法。
Private Declare Function ActivateKeyboardLayout Lib "user32" (ByVal hkl As Long, ByVal flags As Long) As Long

Public Sub Wait(ByVal lngMilliseconds As Long)
'功能：执行挂起一段时间
'lngMilliseconds=毫秒数，1000毫秒=1秒
    Call Sleep(lngMilliseconds)
End Sub

Public Function TempPath() As String
'功能:获取临时文件夹
    Dim strFileTemp As String
    Dim lngTemp As Long
    Dim objFile As New FileSystemObject
    '优先使用VB自带获取方法，失败这通过API
    strFileTemp = objFile.GetSpecialFolder(TemporaryFolder)
    If strFileTemp = "" Then
        strFileTemp = Space(256)
        lngTemp = GetTempPath(256, strFileTemp)
        TempPath = gobjComLib.zlStr.TruncZero(strFileTemp)
    Else 'C:\Users\ADMINI~1\AppData\Local\Temp\
        TempPath = strFileTemp & "\"
    End If
End Function

Public Function IP(Optional ByRef strErr As String) As String
    '功能：通过API获取临时IP
    
    Dim ret As Long, Tel As Long
    Dim bBytes() As Byte
    Dim TempList() As String
    Dim TempIP As String
    Dim Tempi As Long
    Dim Listing As MIB_IPADDRTABLE
    Dim L3 As String
    Dim strTmpErr As String, strALLErr As String
    
    strErr = ""
    On Error GoTo Errhand
    GetIpAddrTable ByVal 0&, ret, True
    If ret <= 0 Then Exit Function
    ReDim bBytes(0 To ret - 1) As Byte
    ReDim TempList(0 To ret - 1) As String
    'retrieve the data
    GetIpAddrTable bBytes(0), ret, False
    'Get the first 4 bytes to get the entry's.. ip installed
    CopyMemory Listing.dEntrys, bBytes(0), 4
    For Tel = 0 To Listing.dEntrys - 1
        'Copy whole structure to Listing..
        CopyMemory Listing.mIPInfo(Tel), bBytes(4 + (Tel * Len(Listing.mIPInfo(0)))), Len(Listing.mIPInfo(Tel))
        TempList(Tel) = ConvertAddressToString(Listing.mIPInfo(Tel).dwAddr, strTmpErr)
        If strTmpErr <> "" Then strALLErr = strALLErr & IIf(strALLErr = "", "", "|") & strTmpErr
    Next Tel
    'Sort Out The IP For WAN
        TempIP = TempList(0)
        For Tempi = 0 To Listing.dEntrys - 1
            L3 = Left(TempList(Tempi), 3)
            If L3 <> "169" And L3 <> "127" And L3 <> "192" Then
                TempIP = TempList(Tempi)
            End If
        Next Tempi
        IP = TempIP 'Return The TempIP
    Exit Function
    strErr = strALLErr
    '------------------------------------------------------------------------------------------------------------------
Errhand:
    strErr = strALLErr & IIf(strALLErr = "", "", "|") & Err.Description
    Err.Clear
End Function

Private Function ConvertAddressToString(longAddr As Long, Optional ByRef strErr As String) As String
    Dim myByte(3) As Byte
    Dim Cnt As Long
    
    strErr = ""
    On Error GoTo errH
    CopyMemory myByte(0), longAddr, 4
    For Cnt = 0 To 3
        ConvertAddressToString = ConvertAddressToString + CStr(myByte(Cnt)) + "."
    Next Cnt
    ConvertAddressToString = Left$(ConvertAddressToString, Len(ConvertAddressToString) - 1)
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errH:
    strErr = Err.Description
    Err.Clear
End Function

Public Function SystemImes() As Variant
'功能：将系统中文输入法名称返回到一个字符串数组中
'返回：如果不存在中文输入法,则返回空串
    Dim arrIme(99) As Long, arrName() As String
    Dim lngLen As Long, strName As String * 255
    Dim lngCount As Long, i As Long, j As Long
    
    lngCount = GetKeyboardLayoutList(UBound(arrIme) + 1, arrIme(0))
    For i = 0 To lngCount - 1
        If ImmIsIME(arrIme(i)) = 1 Then
            ReDim Preserve arrName(j)
            lngLen = ImmGetDescription(arrIme(i), strName, Len(strName))
            arrName(j) = Mid(strName, 1, InStr(strName, Chr(0)) - 1)
            j = j + 1
        End If
    Next
    SystemImes = IIf(j > 0, arrName, vbNullString)
End Function

Public Function SysColor2RGB(ByVal lngColor As Long) As Long
'功能：将VB的系统颜色转换为RGB色
    If lngColor < 0 Then
        Call OleTranslateColor(lngColor, 0, lngColor)
    End If
    SysColor2RGB = lngColor
End Function

Public Sub ShowChildWindow(ByVal lnghwnd As Long, Optional ByVal lngMainHwnd As Long, Optional ByVal blnMaximized As Boolean)
'功能：显示指定的窗体，以子窗体方式
'参数：lngHwnd=要作为子窗体显示的窗体的句柄
'      lngMainHwnd=父窗体句柄，不传时表明不以它的子窗体显示
'      blnMaximized=是否固定将窗体最大化处理；因为API设置为子窗体后，窗体原有的最大化特性将失效
'说明：该项函数主要用于在ZLBH中融合调用ZLHIS窗体显示
    Call gobjComLib.zlCommFun.ShowChildWindow(lnghwnd, lngMainHwnd, blnMaximized)
End Sub

Public Sub SetWindowsInTaskBar(ByVal lnghwnd As Long, ByVal blnShow As Boolean)
'功能：设置窗体是否在任务条上显示
    Call gobjComLib.zlCommFun.SetWindowsInTaskBar(lnghwnd, blnShow)
End Sub

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
    ExpandEnvStr = gobjComLib.zlStr.TruncZero(strBuf)
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
    
    On Error GoTo errH
    hRootKey = 0: strSubKey = "": lngType = 0
    lngPos = InStr(strKey, "\")
    If lngPos = 0 Then Exit Function
    strRoot = Mid(strKey, 1, lngPos - 1)
    strSubKey = Mid(strKey, lngPos + 1)
    
    hRootKey = gobjComLib.Decode(UCase(strRoot), "HKEY_CLASSES_ROOT", HKEY_CLASSES_ROOT, _
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
errH:
    If 0 = 1 Then
        Resume
    End If
    Err.Clear
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
    Dim lngLength As Long, varBufData As Variant, strBufVar() As String, lngBuf As Long, bytBuf() As Byte, strBuf As String
    Dim i As Long, strReturn As String, strTmp As String
    '不是有效的注册表键位,获取键名类型
    If Not GetKeyValueInfo(strKey, strValueName, hRootKey, strSubKey, ruType) Then Exit Function
    '打开变量
    lngReturn = RegOpenKeyEx(hRootKey, strSubKey, 0, KEY_QUERY_VaLUE, lngKey)
    If lngReturn <> ERROR_SUCCESS Then
        Exit Function
    End If
    On Error GoTo errH
    Select Case ruType
        Case REG_SZ, REG_EXPAND_SZ, REG_MULTI_SZ '字符串类型读取
'            lngReturn = RegQueryValueEx(lngKey, strValueName, 0, ruType, 0, lngLength)
'            If lngReturn <> ERROR_SUCCESS Then Err.Clear '可能出错，因此这样处理
            lngLength = 1024: strBuf = Space(lngLength)
            lngReturn = RegQueryValueEx_String(lngKey, strValueName, 0, ruType, strBuf, lngLength)
            If lngReturn <> ERROR_SUCCESS Then: RegCloseKey (lngKey): Exit Function
            Select Case ruType
                Case REG_SZ
                    varValue = gobjComLib.zlStr.TruncZero(strBuf)
                Case REG_EXPAND_SZ ' 扩充环境字符串，查询环境变量和返回定义值
                    If Not blnOneString Then
                        varValue = gobjComLib.zlStr.TruncZero(ExpandEnvStr(gobjComLib.zlStr.TruncZero(strBuf)))
                    Else
                        varValue = gobjComLib.zlStr.TruncZero(strBuf)
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
                        varValue = gobjComLib.zlStr.TruncZero(strBuf)
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
errH:
    If 0 = 1 Then
        Resume
    End If
End Function

Public Function SetRegValue(ByVal strKey As String, ByVal strValueName As String, varValue As Variant) As Boolean
'功能：设置注册表中指定位置的值
'参数：strKey=注册表键位，如“HKEY_CURRENT_USER\Printers\DevModePerUser"
'          strValueName=变量名
'          strValue=变量值
'          strValueType=变量类型，默认为字符串
'返回：是否设置成功
    Dim hRootKey As REGRoot, strSubKey As String
    Dim lngReturn As Long
    Dim lngKey As Long, ruType As REGValueType
    Dim lngLength As Long, varBufData As Variant, strBufVar() As String, lngBuf As Long, bytBuf() As Byte, strBuf As String
    Dim i As Long, lb As Long, ub As Long, strReturn As String, strTmp As String
    '不是有效的注册表键位,获取键名类型
    If Not GetKeyValueInfo(strKey, strValueName, hRootKey, strSubKey, ruType) Then Exit Function
    '打开变量
    lngReturn = RegOpenKeyEx(hRootKey, strSubKey, 0, KEY_SET_VaLUE, lngKey)
    If lngReturn <> ERROR_SUCCESS Then
        Exit Function
    End If
    '当无法获取数据类型时，根据传入的数据类型确定键值类型。
    If ruType = REG_NONE Then
        Select Case TypeName(varValue)
            Case "Byte()"
                ruType = REG_BINARY
            Case "String"
                ruType = REG_SZ
            Case "String()"
                ruType = REG_MULTI_SZ
            Case Else
                If IsNumeric(varValue) Then
                    ruType = REG_DWORD
                End If
        End Select
    End If
    On Error GoTo errH
    Select Case ruType
        Case REG_SZ, REG_EXPAND_SZ, REG_MULTI_SZ
            If ruType = REG_MULTI_SZ And varType(varValue) = vbArray + vbString Then 'string数组，则将数组合成字符串
                lngLength = UBound(varValue) - LBound(varValue) + 1
                For i = LBound(varValue) To UBound(varValue)
                    strBuf = strBuf & varValue(i) & Chr$(0)
                Next
                strBuf = gobjComLib.zlStr.TruncZero(strBuf)
                lngLength = gobjComLib.zlStr.ActualLen(strBuf)
            Else
                strBuf = gobjComLib.zlStr.TruncZero(varValue)
                lngLength = gobjComLib.zlStr.ActualLen(strBuf)
            End If
            lngReturn = RegSetValueEx_String(lngKey, strValueName, ByVal 0&, ruType, ByVal strBuf, lngLength)
            If lngReturn <> 0 Then RegCloseKey lngKey: Exit Function
        Case REG_DWORD
            lngBuf = Val(varValue): lngLength = Len(lngBuf)
            lngReturn = RegSetValueEx_Long(lngKey, strValueName, ByVal 0&, ruType, lngBuf, lngLength)
            If lngReturn <> 0 Then RegCloseKey lngKey: Exit Function
        Case REG_BINARY
            ' 1、varValue ＝ 字节数组，如 B()
            If varType(varValue) = vbArray + vbByte Then
                Dim binValue() As Byte, Length As Long
                bytBuf = varValue
                lngLength = UBound(bytBuf) - LBound(bytBuf) + 1
                lngReturn = RegSetValueEx_BINARY(lngKey, strValueName, 0, ruType, bytBuf(0), lngLength)
                If lngReturn <> 0 Then RegCloseKey lngKey: Exit Function
            ' 2、varValue ＝ 整型或长整型，如 520
            ElseIf varType(varValue) = vbLong Or varType(varValue) = vbInteger Then
                lngBuf = Val(varValue): lngLength = Len(lngBuf)
                lngReturn = RegSetValueEx_Long(lngKey, strValueName, 0, ruType, lngBuf, lngLength)
                If lngReturn <> 0 Then RegCloseKey lngKey: Exit Function
            ' 3、varValue ＝字符串，如 "BE 3E FF AB"
            ElseIf varType(varValue) = vbString Then
                ' 转化数据
                Dim ByteArray() As Byte
                Dim tmpArray() As String '//转换ASCII字符到16进制字节
                strTmp = varValue
                ' 以空格分割字符串
                strBufVar = Split(strTmp, " ")
                lb = LBound(strBufVar): ub = UBound(strBufVar)
                ' 为动态数组分配空间
                ReDim bytBuf(lb To ub)
                ' 循环转换
                For i = lb To ub - 1
                    bytBuf(i) = CByte(Val("&H" & Right$(strBufVar(i), 2)))
                Next i
                ' 注意：最后一个不知道字符串后面多了2个什么，要用 Left$(tmpArray(ub), 2)
                bytBuf(ub) = CByte(Val("&H" & Left$(strBufVar(ub), 2)))
                ' 将数据写入到注册表，注意：最后是 ub - lb + 1
                lngReturn = RegSetValueEx_BINARY(lngKey, strValueName, 0, ruType, bytBuf(0), ub - lb + 1)
                If lngReturn <> 0 Then RegCloseKey lngKey: Exit Function
            End If
    End Select
    RegCloseKey lngKey
    SetRegValue = True
    Exit Function
errH:
    If 0 = 1 Then
        Resume
    End If
End Function

Public Function DeleteRegValue(ByVal strKey As String, ByVal strValueName As String) As Boolean
'功能：删除注册表中指定位置的值
'参数：strKey=注册表键位，如“HKEY_CURRENT_USER\Printers\DevModePerUser"
'          strValueName=变量名
'返回：是否读取成功
    Dim lngLength As Long, lngReturn As Long
    Dim lngKey As Long, lngType As Long
    Dim hRootKey As REGRoot, strSubKey As String
    
    '不是有效的注册表键位
    If Not GetKeyValueInfo(strKey, strValueName, hRootKey, strSubKey, -1) Then Exit Function
    '打开键
    lngReturn = RegOpenKeyEx(hRootKey, strSubKey, 0, KEY_SET_VaLUE, lngKey)
    If lngReturn <> 0 Then
        Exit Function
    End If
    '删除键
    lngReturn = RegDeleteValue(lngKey, strValueName)
    If lngReturn = 0 Then
        DeleteRegValue = True
    End If
    '关闭键
    RegCloseKey lngKey
End Function

Public Sub PressKeyEx(ByVal KeyCode As Integer, Optional ByVal Shift As Integer)
'功能：向键盘发送一个键,类似SendKey
'参数：KeyCode=VirtualKey Codes，1-254，可以用vbKeyTab,vbKeyReturn,vbKeyF4,vbKeyA等
'      Shift=vbKeyControl,vbKeyShift,vbKeyMenu(ALT)，如果是组合，用Or运算
    Call gobjComLib.zlCommFun.PressKeyEx(KeyCode, Shift)
End Sub

Public Sub PressKey(bytKey As Byte)
'功能：向键盘发送一个键,类似SendKey
'参数：bytKey=VirtualKey Codes，1-254，可以用vbKeyTab,vbKeyReturn,vbKeyF4
    Call gobjComLib.zlCommFun.PressKey(bytKey)
End Sub

Public Function OpenImeByName(Optional strIme As String) As Boolean
'功能:按名称打开中文输入法,不指定名称时关闭中文输入法。支持部分名称。
    OpenImeByName = gobjComLib.zlCommFun.OpenIme(Trim(strIme) = "", Trim(strIme))
End Function

Public Function OpenIme(Optional blnOpen As Boolean = False) As Boolean
'功能:打开中文输入法，或关闭输入法
    OpenIme = gobjComLib.zlCommFun.OpenIme(blnOpen)
End Function

Public Function OpenDir(hwndOwner As Long, Optional strTitle As String, Optional ByVal strInitDir As String) As String
'功能：选择Windows的文件夹
    OpenDir = gobjComLib.zlCommFun.OpenDir(hwndOwner, strTitle, strInitDir)
End Function

Public Function IsInstallExcel() As Boolean
'功能：判断本机上装有EXCEL没有
'参数：
'返回：有则返回True
    Dim objTemp  As Object
    
    On Error GoTo errH
    Set objTemp = CreateObject("Excel.Application") '打开一个EXCEL程序
    Set objTemp = Nothing
    IsInstallExcel = True
    Exit Function
errH:
    Set objTemp = Nothing
    IsInstallExcel = False
    Err.Clear
End Function

Public Function IsDesinMode() As Boolean
'功能： 确定当前模式为设计模式
     Err = 0: On Error Resume Next
     Debug.Print 1 / 0
     If Err <> 0 Then
        IsDesinMode = True
     Else
        IsDesinMode = False
     End If
     Err.Clear: Err = 0
 End Function
 
Public Function IsWindowsNT() As Boolean
'功能：判断是否WindowNT操作系统
    IsWindowsNT = gobjComLib.zlCommFun.IsWindowsNT
End Function

Public Function IsWindows9X() As Boolean
'功能：判断是否Window95操作系统
    IsWindows9X = gobjComLib.zlCommFun.IsWindows9X
End Function

Public Function Is64bit() As Boolean
    '******************************************************************************************************************
    '功能：是否是64位系统
    '返回：
    '******************************************************************************************************************
    Dim handle As Long
    Dim lngFunc As Long
        
    lngFunc = 0
    handle = GetProcAddress(GetModuleHandle("kernel32"), "IsWow64Process")
    If handle > 0 Then
        IsWow64Process GetCurrentProcess(), lngFunc
    End If
    Is64bit = lngFunc <> 0
End Function

Public Function IniRead(strItem As String, strKey As String, strPath As String) As String
'功能：INi文件读
    Dim GetStr As String
    On Error GoTo errH

    GetStr = String(128, 0)
    GetPrivateProfileString strItem, strKey, "", GetStr, 256, strPath
    GetStr = Replace(GetStr, Chr(0), "")
    IniRead = GetStr
    Exit Function
errH:
    Err.Clear
    IniRead = ""
End Function

Public Function IniWrite(strItem As String, strKey As String, strVal As String, strPath As String) As Boolean
'功能：INi文件写
    On Error GoTo errH
    IniWrite = True
    WritePrivateProfileString strItem, strKey, strVal, strPath
    Exit Function
errH:
    Err.Clear
    IniWrite = False
End Function

Public Function Dec2Bin(bDec As Byte) As String
'功能：十进制转为二进制函数
'用法：String  Dec2Bin(Bdec as Byte)
'返回：  十进制的二进制 字符串(String)
'错误：  返回"0"

    Dim strBin As String

    On Error GoTo Err
    If bDec > 255 Then
        Dec2Bin = "-1"
        Exit Function
    End If
    strBin = ""
    '转为字符串
    While bDec > 0
        strBin = bDec Mod 2 & strBin
        bDec = Fix(bDec / 2)
    Wend
    '补零足8位
    If Len(strBin) < 9 Then
        While Len(strBin) < 8
            strBin = "0" & strBin
        Wend
    End If
    Dec2Bin = strBin
    Exit Function
Err:
    Err.Clear
   Dec2Bin = "0"
End Function

Public Function Bin2Dec(strBin As String) As Long
    '功能：二进制转为十进制函数
    '用法：Long  bin2dec(strBin as String)
    '返回：  二进制的十进制 长整数（Long）
    '错误：  返回-1
    Dim lDec As Long
    Dim lCount As Long
    Dim i As Long
    
    On Error GoTo Errhand
    lDec = 0
    If strBin = "" Then strBin = "0"
    lCount = Len(strBin)
    For i = 1 To lCount
        lDec = lDec + CInt(Left(strBin, 1)) * 2 ^ (Len(strBin) - 1)
        strBin = Right(strBin, Len(strBin) - 1)
        DoEvents
    Next
    Bin2Dec = lDec
    Exit Function
Errhand:
    Err.Clear
    Bin2Dec = -1
End Function

Public Function HIWORD(LongIn As Long) As Integer
'函 数 名：HIWORD
'输    入：LongIn(Long) - 32位值
'输    出：(Integer) - 32位值的高16位
'功能描述：取出32位值的高16位

   ' 取出32位值的高16位
     HIWORD = (LongIn And &HFFFF0000) \ &H10000
End Function

Public Function LOWORD(LongIn As Long) As Integer
'函 数 名：LOWORD
'输    入：LongIn(Long) - 32位值
'输    出：(Integer) - 32位值的低16位
'功能描述：取出32位值的低16位
   ' 取出32位值的低16位
     LOWORD = LongIn And &HFFFF&
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

Public Function ObjectHasProperty(obj As Object, ByVal strProName As String) As Boolean
'功能判断组件是否具备指定的属性
    On Error GoTo errH
    Dim objinfo As InterfaceInfo
    Dim meminfo As MemberInfo
    Dim strTmp As String
    On Error GoTo errH
    Set objinfo = InterfaceInfoFromObject(obj)
    For Each meminfo In objinfo.Members
        If meminfo.InvokeKind = INVOKE_PROPERTYGET Then
            If UCase(strProName) = UCase(meminfo.Name) Then
                ObjectHasProperty = True: Exit For
            End If
        End If
    Next
    ObjectHasProperty = True
    Exit Function
errH:
    Err.Clear
    ObjectHasProperty = False
End Function

Public Function ObjectHasMethod(obj As Object, ByVal strFunName As String) As Boolean
'警告：该判断动态调用了了方法，如果方法无必选参数，则方法会调用成功
'功能：判断组件是否具备指定的方法
'参数：obj=对象
'          strFunName=方法函数名称
'返回：True-存在方法，False-不存在方法
    Dim objinfo As InterfaceInfo
    Dim meminfo As MemberInfo
    Dim strTmp As String
    On Error GoTo errH
    Set objinfo = InterfaceInfoFromObject(obj)
    For Each meminfo In objinfo.Members
        If meminfo.InvokeKind = INVOKE_FUNC Then
            If UCase(strFunName) = UCase(meminfo.Name) Then
                ObjectHasMethod = True: Exit For
            End If
'            Debug.Print meminfo.Name
        End If
    Next
    Exit Function
errH:
    ObjectHasMethod = False
    Err.Clear
End Function

Public Function TaskbarHeight() As Integer
    '-----------------------------------------------------------------------------------------------------------
    '功能:获取任务栏高度
    '入参:
    '出参:
    '返回:成功,返回true,否则返回False
    '-----------------------------------------------------------------------------------------------------------
    Dim lRes As Long
    Dim vRect As RECT
    Err = 0: On Error GoTo Errhand:
    lRes = SystemParametersInfo(SPI_GETWORKAREA, 0, vRect, 0)
    TaskbarHeight = ((Screen.Height / Screen.TwipsPerPixelX) - vRect.Bottom) * Screen.TwipsPerPixelX
Errhand:
    Err.Clear
End Function


