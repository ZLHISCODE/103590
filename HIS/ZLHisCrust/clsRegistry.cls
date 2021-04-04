VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRegistry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

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
Public Enum REGValueType
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
    KEY_QUERY_VALUE = &H1
    KEY_SET_VALUE = &H2
    KEY_CREATE_SUB_KEY = &H4
    KEY_ENUMERATE_SUB_KEYS = &H8
    KEY_NOTIFY = &H10
    KEY_CREATE_LINK = &H20
    KEY_ALL_ACCESS = &H3F
    KEY_READ = &H20019
End Enum
Private Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type
Private Const REG_OPTION_NON_VOLATILE = 0        ' 当系统重新启动时，关键字被保留
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
Private Declare Function RegEnumValue Lib "advapi32" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
Private Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, lpSecurityAttributes As SECURITY_ATTRIBUTES, phkResult As Long, lpdwDisposition As Long) As Long

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
    ExpandEnvStr = TruncZero(strBuf)
End Function

Public Function GetAllSubKey(ByVal strKey As String, Optional ByVal lngPerCout As Long) As Variant
'功能:获取某项的所有子项
'参数：strKey-路径
'lngPerCout:每次枚举的数量，-1，只关闭句柄，不枚举。0-枚举所有的，>0枚举指定数量的子健
'返回：=子项数组
    Dim lngRet As Long, strName As String
    Dim strKeyName As String
    Dim strSubKey As Variant
    
    Dim lngBegin        As Long
    Static strLastKey   As Long
    Static lngIdx       As Long
    Static hRootKey     As Long
    Static lnghKey      As Long
    

    
    strSubKey = Array()
    strName = String(256, Chr(0))
    If UCase(strLastKey) <> UCase(strKey) Or lngPerCout = 0 Then
        '强制关闭句柄
        If lnghKey <> 0 And lngPerCout - 1 Then
            RegCloseKey lnghKey
            GetAllSubKey = strSubKey
            Exit Function
        End If
        lngIdx = 0
        If Not GetKeyValueInfo(strKey, "", hRootKey, strKeyName) Then Exit Function
        lngRet = RegOpenKey(hRootKey, strKeyName, lnghKey)
    End If
    lngBegin = lngIdx
    If lngRet = 0 Then
        Do
            lngRet = RegEnumKey(lnghKey, lngIdx, strName, Len(strName))
            If lngRet = 0 Then
                ReDim Preserve strSubKey(UBound(strSubKey) + 1)
                strSubKey(UBound(strSubKey)) = Left(strName, InStr(strName, Chr(0)) - 1)
                lngIdx = lngIdx + 1
                If lngIdx - lngBegin = lngPerCout Then
                    Exit Do
                End If
            End If
        Loop Until lngRet <> 0
    End If
    '枚举结束关闭句柄
    If lngRet <> 0 Then
        RegCloseKey lnghKey
        lnghKey = 0
    End If
    GetAllSubKey = strSubKey
End Function

Private Function GetKeyValueInfo(ByVal strKey As String, Optional ByVal strValueName As String, Optional ByRef hRootKey As REGRoot, Optional ByRef strSubKey As String, Optional ByRef lngType As Long, Optional ByVal blnCreateValue As Boolean) As Boolean
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
    
    hRootKey = Decode(UCase(strRoot), "HKEY_CLASSES_ROOT", HKEY_CLASSES_ROOT, _
                                                                         "HKEY_CURRENT_USER", HKEY_CURRENT_USER, _
                                                                         "HKEY_LOCAL_MACHINE", HKEY_LOCaL_MaCHINE, _
                                                                         "HKEY_USERS", HKEY_USERS, _
                                                                         "HKEY_PERFORMANCE_DATA", HKEY_PERFORMANCE_DATA, _
                                                                         "HKEY_CURRENT_CONFIG", HKEY_CURRENT_CONFIG, _
                                                                         "HKEY_DYN_DATA", HKEY_DYN_DATA, 0)
    If hRootKey = 0 Then Exit Function
    If lngType <> -1 Then
        '使用查询方式打开，进行键名类型查询
        lngReturn = RegOpenKeyEx(hRootKey, strSubKey, 0, KEY_QUERY_VALUE, hKey)
        If lngReturn <> ERROR_SUCCESS Then
            Exit Function
        End If

        lngReturn = RegQueryValueEx_ValueType(hKey, strValueName, ByVal 0&, lngType, ByVal strName, Len(strName))
        '可能字段超长，长度不够，所以出错不退出
        'If lngReturn <> ERROR_SUCCESS Then: RegCloseKey (hKey): Exit Function
'        If lngReturn <> ERROR_SUCCESS Then
'            '文件未找到
'            If Not blnCreateValue And lngReturn = 2 Then Exit Function
'        End If

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
    lngReturn = RegOpenKeyEx(hRootKey, strSubKey, 0, KEY_QUERY_VALUE, lngKey)
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
                    varValue = TruncZero(strBuf)
                Case REG_EXPAND_SZ ' 扩充环境字符串，查询环境变量和返回定义值
                    If Not blnOneString Then
                        varValue = TruncZero(ExpandEnvStr(TruncZero(strBuf)))
                    Else
                        varValue = TruncZero(strBuf)
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
                        varValue = TruncZero(strBuf)
                    End If
            End Select
        Case REG_DWORD
            lngReturn = RegQueryValueEx_Long(lngKey, strValueName, ByVal 0&, ruType, lngBuf, Len(lngBuf))
            If lngReturn <> ERROR_SUCCESS Then: RegCloseKey (lngKey): varValue = 0: Exit Function
            varValue = lngBuf
        Case REG_BINARY
            lngReturn = RegQueryValueEx_BINARY(lngKey, strValueName, 0, ruType, ByVal 0, lngLength)
            If lngReturn <> ERROR_SUCCESS Then
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
            If lngReturn <> ERROR_SUCCESS Then
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

Public Function SetRegValue(ByVal strKey As String, ByVal strValueName As String, varValue As Variant, Optional ByVal ruValueType As REGValueType = -1) As Boolean
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
    If Not GetKeyValueInfo(strKey, strValueName, hRootKey, strSubKey, ruType, ruValueType <> -1) Then
        Exit Function
    End If
    If ruValueType <> -1 Then ruType = ruValueType
    '打开变量
    lngReturn = RegOpenKeyEx(hRootKey, strSubKey, 0, KEY_SET_VALUE, lngKey)
    If lngReturn <> ERROR_SUCCESS Then
        Exit Function
    End If
    On Error GoTo ErrH
    Select Case ruType
        Case REG_SZ, REG_EXPAND_SZ, REG_MULTI_SZ
            If ruType = REG_MULTI_SZ And VarType(varValue) = vbArray + vbString Then 'string数组，则将数组合成字符串
                lngLength = UBound(varValue) - LBound(varValue) + 1
                For i = LBound(varValue) To UBound(varValue)
                    strBuf = strBuf & varValue(i) & Chr$(0)
                Next
                strBuf = TruncZero(strBuf)
                lngLength = ActualLen(strBuf)
            Else
                strBuf = TruncZero(varValue)
                lngLength = ActualLen(strBuf)
            End If
            lngReturn = RegSetValueEx_String(lngKey, strValueName, ByVal 0&, ruType, ByVal strBuf, lngLength)
            If lngReturn <> 0 Then RegCloseKey lngKey: Exit Function
        Case REG_DWORD
            lngBuf = Val(varValue): lngLength = Len(lngBuf)
            lngReturn = RegSetValueEx_Long(lngKey, strValueName, ByVal 0&, ruType, lngBuf, lngLength)
            If lngReturn <> 0 Then RegCloseKey lngKey: Exit Function
        Case REG_BINARY
            ' 1、varValue ＝ 字节数组，如 B()
            If VarType(varValue) = vbArray + vbByte Then
                Dim binValue() As Byte, Length As Long
                bytBuf = varValue
                lngLength = UBound(bytBuf) - LBound(bytBuf) + 1
                lngReturn = RegSetValueEx_BINARY(lngKey, strValueName, 0, ruType, bytBuf(0), lngLength)
                If lngReturn <> 0 Then RegCloseKey lngKey: Exit Function
            ' 2、varValue ＝ 整型或长整型，如 520
            ElseIf VarType(varValue) = vbLong Or VarType(varValue) = vbInteger Then
                lngBuf = Val(varValue): lngLength = Len(lngBuf)
                lngReturn = RegSetValueEx_Long(lngKey, strValueName, 0, ruType, lngBuf, lngLength)
                If lngReturn <> 0 Then RegCloseKey lngKey: Exit Function
            ' 3、varValue ＝字符串，如 "BE 3E FF AB"
            ElseIf VarType(varValue) = vbString Then
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
ErrH:
    If 0 = 1 Then
        Resume
    End If
End Function

Public Function GetAllKeyValue(ByVal strKey As String) As Variant
'功能:获取某项的所有键值
'返回：=键值数组，数组内容为，名称,键值


'功能:获取某项的所有子项
'参数：strKey-路径
'lngPerCout:每次枚举的数量，-1，只关闭句柄，不枚举。0-枚举所有的，>0枚举指定数量的子健
'返回：=子项数组
    Dim lngRet          As Long, hRootKey       As Long
    Dim strSubKey       As String
    Dim arrKeyValues    As Variant
    Dim lngIdx          As Long
    Dim lnghCurKey      As Long
    Dim strName         As String, strValue     As String
    
    arrKeyValues = Array()
    If Not GetKeyValueInfo(strKey, "", hRootKey, strSubKey) Then
        GetAllKeyValue = arrKeyValues
        Exit Function
    End If
    lngIdx = 0
    lngRet = RegOpenKeyEx(hRootKey, strSubKey, 0, KEY_READ, lnghCurKey)
    If lngRet = ERROR_SUCCESS Then
        Do
            strName = String(256, Chr(0)): strValue = String(256, Chr(0))
            lngRet = RegEnumValue(lnghCurKey, lngIdx, strName, Len(strName), ByVal 0&, REG_SZ, strValue, Len(strValue))
            If lngRet = ERROR_SUCCESS Then
                ReDim Preserve arrKeyValues(UBound(arrKeyValues) + 1)
                arrKeyValues(UBound(arrKeyValues)) = Left(strName, InStr(strName, Chr(0)) - 1)
                lngIdx = lngIdx + 1
            End If
        Loop Until lngRet <> ERROR_SUCCESS
        If lngRet <> ERROR_SUCCESS And lngRet <> ERROR_NO_MORE_ITEMS Then
            'ERR
        End If
        Call RegCloseKey(lnghCurKey)
    Else
        If lngRet <> ERROR_FILE_NOT_FOUND Then
            'Error
        End If
    End If
    GetAllKeyValue = arrKeyValues
End Function

Public Function CreateRegKey(ByVal strKey As String, ByVal strValueName As String, varValue As Variant, Optional ByVal ruValueType As REGValueType = REG_SZ) As Boolean
'功能：设置注册表中指定位置的值
'参数：strKeyParent=注册表键位，如“HKEY_CURRENT_USER\Printers\DevModePerUser"
'       strCreateKey=需要创建的子健
'          strValueName=变量名
'          strValue=变量值
'返回：是否设置成功
    Dim hRootKey As REGRoot, strSubKey As String
    Dim lngReturn As Long
    Dim lngKey As Long, ruType As REGValueType
    Dim lpAttr As SECURITY_ATTRIBUTES                   ' 注册表安全类型
    Dim hDepth As Long
    
    '不是有效的注册表键位,获取键名类型
    If Not GetKeyValueInfo(strKey, strValueName, hRootKey, strSubKey, ruType) Then
        '打开变量
        lpAttr.nLength = 50                                 ' 设置安全属性为缺省值...
        lpAttr.lpSecurityDescriptor = 0                     ' ...
        lpAttr.bInheritHandle = True                        ' ...
        
        lngReturn = RegCreateKeyEx(hRootKey, strSubKey, 0, REG_SZ, REG_OPTION_NON_VOLATILE, KEY_CREATE_SUB_KEY, lpAttr, lngKey, hDepth)
        If lngReturn <> ERROR_SUCCESS Then
            Exit Function
        End If
        lngReturn = RegCloseKey(lngKey)
    End If
        
    CreateRegKey = SetRegValue(strKey, strValueName, varValue, ruValueType)
End Function
