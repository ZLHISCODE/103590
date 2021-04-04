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
'@模块 clsRegistry-2019/6/26
'@编写 lshuo
'@功能
'
'@引用
'
'@备注
'    Dim Registry     As New clsRegistry
'    Dim arrKeys     As Variant, arrNames As Variant, arrValues As Variant
'    Dim blnSuc      As Boolean
'    Dim arrByt()    As Byte
'    Dim arrStr(2)   As String
'    If Registry.GetRegKeyValues("HKEY_CURRENT_USER\Software\VB and VBA Program Settings\ZLSOFT\私有模块\ZLHIS\界面设置\zl9AppTool\frmMessageManager\ListView", arrNames, arrValues) Then
'        Debug.Print DisPlayOneValue(arrNames)
'        Debug.Print DisPlayOneValue(arrValues)
'    End If
'    arrByt = "123456"
'    Debug.Print Registry.CreateRegKeyValue("HKEY_CURRENT_USER\Software\VB and VBA Program Settings\ZLSOFT\私有模块\ZLHIS\界面设置\zl9AppTool\frmMessageManager\ListView\测试1\测试2", "TEST", 11222)
'    Debug.Print Registry.CreateRegKeyValue("HKEY_CURRENT_USER\Software\VB and VBA Program Settings\ZLSOFT\私有模块\ZLHIS\界面设置\zl9AppTool\frmMessageManager\ListView\测试1\测试2", "TEST_Byt", arrByt)
'    Debug.Print Registry.CreateRegKeyValue("HKEY_CURRENT_USER\Software\VB and VBA Program Settings\ZLSOFT\私有模块\ZLHIS\界面设置\zl9AppTool\frmMessageManager\ListView\测试1\测试2", "TEST_Byt", "AC AD 00")
'    Debug.Print Registry.CreateRegKeyValue("HKEY_CURRENT_USER\Software\VB and VBA Program Settings\ZLSOFT\私有模块\ZLHIS\界面设置\zl9AppTool\frmMessageManager\ListView\测试1\测试2", "TEST_Str", "123")
'    arrStr(0) = "测试1"
'    arrStr(1) = "测试2"
'    arrStr(2) = "测试3"
'    Debug.Print Registry.CreateRegKeyValue("HKEY_CURRENT_USER\Software\VB and VBA Program Settings\ZLSOFT\私有模块\ZLHIS\界面设置\zl9AppTool\frmMessageManager\ListView\测试1\测试2", "TEST_StrArray", arrByt)
'    Debug.Print Registry.CreateRegKeyValue("HKEY_CURRENT_USER\Software\VB and VBA Program Settings\ZLSOFT\私有模块\ZLHIS\界面设置\zl9AppTool\frmMessageManager\ListView\测试1\测试2", "TEST_StrArray", 1)
'    Debug.Print Registry.CreateRegKeyValue("HKEY_CURRENT_USER\Software\VB and VBA Program Settings\ZLSOFT\私有模块\ZLHIS\界面设置\zl9AppTool\frmMessageManager\ListView\测试1\测试2", "TEST_StrArray", "测试")
'    arrStr(3) = "测试4"
'    Debug.Print Registry.CreateRegKeyValue("HKEY_CURRENT_USER\Software\VB and VBA Program Settings\ZLSOFT\私有模块\ZLHIS\界面设置\zl9AppTool\frmMessageManager\ListView\测试1\测试2", "TEST_StrArray", arrStr)
'    Debug.Print Registry.DeleteRegKeyValue("HKEY_CURRENT_USER\Software\VB and VBA Program Settings\ZLSOFT\私有模块\ZLHIS\界面设置\zl9AppTool\frmMessageManager\ListView", "lvwMain排序")
'    Debug.Print Registry.DeleteRegKey("HKEY_CURRENT_USER\Software\VB and VBA Program Settings\ZLSOFT\私有模块\ZLHIS\界面设置\zl9BaseItem")
'    Debug.Print Registry.GetRegValue("HKEY_CURRENT_USER\Software\VB and VBA Program Settings\ZLSOFT\私有模块\ZLHIS\界面设置\zl9AppTool\frmMessageManager\ListView", "lvwMain宽度")
'    arrKeys = Registry.GetRegSubKeys("HKEY_CURRENT_USER\Software\VB and VBA Program Settings\ZLSOFT\公共模块", blnSuc)
'    Debug.Print DisPlayOneValue(arrKeys)
Option Explicit
'---------------------------------------------------------------------------
'                0。API和常量声明
'---------------------------------------------------------------------------
'下表标识了各种注册表元素的大小限制?
'    键名    255个字符。键名包括注册表中键的绝对路径，总是从基键开始，例如HKEY_LOCAL_MACHINE。
'    值名    16,383个字符   windows 2000: 260个ANSI字符或16,383个Unicode字符。
'    可用内存(最新格式)1mb(标准格式)
'    一个注册表树可以有512层深?您可以通过一个注册表API调用一次创建多达32个级别?
'    长值(超过2048字节)应该存储在一个文件中，文件的位置应该存储在注册表中。这有助于注册表高效地执行。
'    文件位置可以是值的名称，也可以是值的数据。位置字符串中的每个反斜杠前必须有另一个反斜杠作为转义字符。例如，指定“C:\\mydir\\myfile”来存储字符串“C:\mydir\myfile”。如果文件位置在键名的255个字符限制之内，并且不包含反斜杠(键名中不允许使用反斜杠)，那么它也可以是键名。

'注册表根键句柄
Private Const HKEY_CLASSES_ROOT         As Long = &H80000000            '记录Windows操作系统中所有数据文件的格式和关联信息，主要记录不同文件的文件名后缀和与之对应的应用程序。其下子键可分为两类，一类是已经注册的各类文件的扩展名，这类子键前面都有一个“。”；另一类是各类文件类型有关信息。
Private Const HKEY_CURRENT_USER         As Long = &H80000001            '此根键包含了当前登录用户的用户配置文件信息。这些信息保证不同的用户登录计算机时，使用自己的个性化设置，例如自己定义的墙纸。自己的收件箱。自己的安全访问权限等。
Private Const HKEY_LOCAL_MACHINE        As Long = &H80000002            '此根键包含了当前计算机的配置数据，包括所安装的硬件以及软件的设置。这些信息是为所有的用户登录系统服务的。它是整个注册表中最庞大也是最重要的根键！
Private Const HKEY_USERS                As Long = &H80000003            '此根键包括默认用户的信息（Default子键）和所有以前登录用户的信息。
Private Const HKEY_PERFORMANCE_DATA     As Long = &H80000004            '在Windows NT/2000/XP注册表中虽然没有HKEY_DYN_DATA键，但是它却隐藏了一个名为“HKEY_ PERFOR MANCE_DATA”键。所有系统中的动态信息都是存放在此子键中。系统自带的注册表编辑器无法看到此键
Private Const HKEY_CURRENT_CONFIG       As Long = &H80000005            '此根键实际上是HKEY_LOCAL_MACHINE中的一部分，其中存放的是计算机当前设置，如显示器。打印机等外设的设置信息等。它的子键与HKEY_LOCAL_ MACHINE\ Config\0001分支下的数据完全一样。
Private Const HKEY_DYN_DATA             As Long = &H80000006            '此根键中保存每次系统启动时，创建的系统配置和当前性能信息。这个根键只存在于Windows 98中。
'注册表数据类型
Private Const REG_NONE                  As Long = 0                     ' No value type
Private Const REG_SZ                    As Long = 1                     ' 一个以null结尾的字符串。这将是Unicode或ANSI字符串，这取决于您是使用Unicode还是ANSI函数。
Private Const REG_EXPAND_SZ             As Long = 2                     '一个以null结尾的字符串，包含对环境变量的未展开引用(例如，“%PATH%”)。它将是Unicode或ANSI字符串，这取决于您是使用Unicode还是ANSI函数。要展开环境变量引用，请使用expandenvironmentstring函数。
Private Const REG_BINARY                As Long = 3                     '任何形式的二进制数据。
Private Const REG_DWORD                 As Long = 4                     '一个32位数字
Private Const REG_DWORD_LITTLE_ENDIAN   As Long = 4                     'little-endian格式的32位数字。Windows被设计为在little-endian计算机架构上运行。因此，这个值在Windows头文件中定义为REG_DWORD。
Private Const REG_DWORD_BIG_ENDIAN      As Long = 5                     '大端数字格式的32位数字。一些UNIX系统支持大端体系结构?
'   字节格式
'    在little-endian格式中，多字节值存储在内存中，从最低字节(“little end”)到最高字节。例如，值0x12345678以little-endian格式存储为(0x78 0x56 0x34 0x12)。
'    在big-endian格式中，多字节值存储在内存中，从最高字节(“big end”)到最低字节。例如，值0x12345678以大端字节格式存储为(0x12 0x34 0x56 0x78)。
Private Const REG_LINK                  As Long = 6                     '使用REG_OPTION_CREATE_LINK调用RegCreateKeyEx函数创建的符号链接的目标路径。
Private Const REG_MULTI_SZ              As Long = 7                     '由空字符串(\0)终止的以null结尾的字符串序列。
'   下面是一个例子:
'    String1 \ 0 string2\ 0 string3 \ 0 laststring \0 \ 0
'    第一个\0终止第一个字符串，第二个\0终止最后一个字符串，最后一个\0终止序列。注意，最后的终止符必须被分解到字符串的长度中。
'    如果数据具有REG_SZ、REG_MULTI_SZ或REG_EXPAND_SZ类型，则字符串可能没有存储正确的终止null字符。因此，当从注册表读取字符串时，必须确保字符串在使用前已正确终止;否则，它可能会覆盖缓冲区。(注意REG_MULTI_SZ字符串应该有两个终止null字符。)
'    在向注册表写入字符串时，必须指定字符串的长度，包括终止null字符(\0)。一个常见的错误是使用strlen函数来确定字符串的长度，但是忘记strlen只返回字符串中的字符数，不包括终止null。因此，字符串的长度应该计算如下:strlen(string) + 1
'    REG_MULTI_SZ字符串以长度为0的字符串结束。因此，不可能在序列中包含零长度的字符串。空序列的定义如下:\0。
Private Const REG_RESOURCE_LIST         As Long = 8                     '
Private Const REG_FULL_RESOURCE_DESCRIPTOR As Long = 9                  '
'注册表错误
Private Const ERROR_SUCCESS             As Long = &H0
Private Const ERROR_FILE_NOT_FOUND      As Long = &H2  'The system cannot find the file specified
Private Const ERROR_BADDB               As Long = &H3F1
Private Const ERROR_BADKEY              As Long = &H3F2
Private Const ERROR_CANTOPEN            As Long = &H3F3
Private Const ERROR_CANTREAD            As Long = &H3F4
Private Const ERROR_CANTWRITE           As Long = &H3F5
Private Const ERROR_OUTOFMEMORY         As Long = &HE
Private Const ERROR_INVALID_PARAMETER   As Long = &H57
Private Const ERROR_ACCESS_DENIED       As Long = &H5
Private Const ERROR_NO_MORE_ITEMS       As Long = &H103
Private Const ERROR_MORE_DATA           As Long = &HEA
'注册表访问权
Private Const KEY_QUERY_VALUE           As Long = &H1
Private Const KEY_SET_VALUE             As Long = &H2
Private Const KEY_CREATE_SUB_KEY        As Long = &H4
Private Const KEY_ENUMERATE_SUB_KEYS    As Long = &H8
Private Const KEY_NOTIFY                As Long = &H10
Private Const KEY_CREATE_LINK           As Long = &H20
Private Const KEY_ALL_ACCESS            As Long = &H3F
Private Const KEY_READ                  As Long = &H20019

Private Type SECURITY_ATTRIBUTES
    nLength                     As Long
    lpSecurityDescriptor        As Long
    bInheritHandle              As Long
End Type
Private Declare Function ExpandEnvironmentStrings Lib "kernel32" Alias "ExpandEnvironmentStringsW" (ByVal lpSrc As Long, ByVal lpDst As Long, ByVal nSize As Long) As Long
'@功能
'    展开环境变量字符串，并用为当前用户定义的值替换它们。
'    要为特定用户或系统指定环境块，请使用ExpandEnvironmentStringsForUser函数。
'@原型
'    DWORD ExpandEnvironmentStringsA(
'      LPCWSTR lpSrc,
'      LPSTR  lpDst,
'      DWORD nSize
'    );
'@参数
'lpSrc
'    以%variableName%的形式包含一个或多个环境变量字符串的缓冲区。对于每个这样的引用，%variableName%部分将替换为该环境变量的当前值。
'    在查找环境变量名时忽略大小写。如果没有找到名称，%variableName%部分将保持未展开。
'    注意，这个函数不支持Cmd.exe支持的所有特性。例如，它不支持%variableName:str1=str2%或%variableName:~offset,length%。
'lpDst
'    指向缓冲区的指针，该缓冲区接收在lpSrc缓冲区中展开环境变量字符串的结果。注意，这个缓冲区不能与lpSrc缓冲区相同。
'nSize
'    lpDst参数指向的缓冲区中可存储的最大字符数。当使用ANSI字符串时，缓冲区大小应该是字符串长度，加上终止null字符，再加上1。当使用Unicode字符串时，缓冲区大小应该是字符串长度加上终止null字符。
'@返回值
'    如果函数成功，返回值是存储在目标缓冲区中的tchar的数量，包括终止null字符。如果目标缓冲区太小，无法容纳展开的字符串，则返回值是所需的缓冲区大小(以字符为单位)。
'    如果函数失败，返回值为零。要获取扩展的错误信息，请调用GetLastError。
'@备注
'    lpSrc和lpDst缓冲区的大小限制为32K?
'    要将完全限定路径中的文件夹名替换为与其相关的环境变量字符串，请使用PathUnExpandEnvStrings函数。
'    要检索进程的环境变量列表，请使用getenvironmentstring函数。
'@Requirements
'    Minimum supported client    Windows 2000 Professional [desktop apps | UWP apps]
'    Minimum supported server    Windows 2000 Server [desktop apps | UWP apps]
'    Target Platform Windows
'    Header  processenv.h (include Windows.h)
'    Library kernel32.lib
'    dll kernel32.dll
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
'@功能
'    关闭指定注册表项的句柄。
'@原型
'    LSTATUS RegCloseKey(
'      hKey hKey
'    );
'@参数
'hKey
'    要关闭的打开键的句柄。句柄必须由RegCreateKeyEx。RegCreateKeyTransacted。RegOpenKeyEx。RegOpenKeyTransacted或RegConnectRegistry函数打开。
'@返回值
'    如果函数成功，返回值为ERROR_SUCCESS。
'    如果函数失败，返回值是在Winerror.h中定义的非零错误代码。您可以使用带有FORMAT_MESSAGE_FROM_SYSTEM标志的FormatMessage函数来获得错误的一般描述。
'@备注
'    指定键的句柄在关闭后不应使用，因为它将不再有效。键柄的开启时间不应超过必要的时间。
'    RegCloseKey函数不一定在返回之前向注册表写入信息;将缓存刷新到硬盘可能需要几秒钟的时间。如果应用程序必须显式地将注册表信息写入硬盘，则可以使用RegFlushKey函数。然而，RegFlushKey使用许多系统资源，只有在必要时才应该调用它。
'@Requirements
'    Minimum supported client    Windows 2000 Professional [desktop apps only]
'    Minimum supported server    Windows 2000 Server [desktop apps only]
'    Target Platform Windows
'    Header  winreg.h (include Windows.h)
'    Library advapi32.lib
'    dll advapi32.dll
Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As Long, phkResult As Long) As Long
'@功能
'    创建指定的注册表项。如果该键已经存在于注册表中，则函数将打开它。
'    注意，此函数仅用于兼容16位版本的Windows。应用程序应该使用RegCreateKeyEx函数。但是，备份或恢复系统状态(包括系统文件和注册表组)的应用程序应该使用卷影复制服务，而不是注册表函数。
'@原型
'    LSTATUS RegCreateKeyA(
'      HKEY   hKey,
'      LPCWSTR lpSubKey,
'      PHKEY phkResult
'    );
'@参数
'HKEY
'    打开注册表项的句柄。调用进程必须具有对键的KEY_CREATE_SUB_KEY访问权。有关更多信息，请参见注册表键安全性和访问权限。
'    键创建的访问是根据注册表键的安全描述符进行检查的，而不是在获取句柄时指定的访问掩码。因此，即使使用samDesired 传递 KEY_READ打开hKey，如果安全描述符允许，也可以在创建键的操作中使用它。
'    这个句柄由RegCreateKeyEx或RegOpenKeyEx函数返回，也可以是以下预定义键之一:
'        HKEY_CLASSES_ROOT
'        HKEY_CURRENT_CONFIG
'        HKEY_CURRENT_USER
'        HKEY_LOCAL_MACHINE
'        HKEY_USERS
'lpSubKey
'    此函数打开或创建的键的名称。此键必须是由hKey参数标识的键的子键。
'    有关键名的更多信息，请参见注册表结构。
'    如果hKey是预定义的键之一，lpSubKey可能为NULL。在这种情况下，phkResult接收传递给函数的相同的hKey句柄。
'phkResult
'    指向一个变量的指针，该变量接收打开或创建的键的句柄。如果键不是预定义的注册表键之一，则在使用句柄之后调用RegCloseKey函数。
'@返回值
'    如果函数成功，返回值为ERROR_SUCCESS。
'    如果函数失败，返回值是在Winerror.h中定义的非零错误代码。您可以使用带有FORMAT_MESSAGE_FROM_SYSTEM标志的FormatMessage函数来获得错误的一般描述。
'@备注
'    应用程序不能创建HKEY_USERS或HKEY_LOCAL_MACHINE的直接子健。应用程序可以在HKEY_USERS或HKEY_LOCAL_MACHINE树的较低层创建子键。
'    如果您的服务或应用程序模拟不同的用户，则不要将此函数与HKEY_CURRENT_USER一起使用。相反，调用RegOpenCurrentUser函数。
'    RegCreateKey函数在指定的路径中创建所有缺失的键。应用程序可以利用这种行为一次创建多个键。例如，一个应用程序可以通过为lpSubKey参数指定如下形式的字符串，同时创建一个四层深的子键，与前面的三个子键相同:
'    subkey1 \ subkey2 \ subkey3 \ subkey4
'    注意，如果路径中的现有键拼写错误，此行为将导致创建不需要的键。
'@Requirements
'    Minimum supported client    Windows 2000 Professional [desktop apps only]
'    Minimum supported server    Windows 2000 Server [desktop apps only]
'    Target Platform Windows
'    Header  winreg.h (include Windows.h)
'    Library advapi32.lib
'    dll advapi32.dll
Private Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExW" (ByVal hKey As Long, ByVal lpSubKey As Long, ByVal Reserved As Long, ByVal lpClass As Long, ByVal dwOptions As Long, ByVal samDesired As Long, lpSecurityAttributes As SECURITY_ATTRIBUTES, phkResult As Long, lpdwDisposition As Long) As Long
'@功能
'    创建指定的注册表项。如果键已经存在，函数将打开它。注意，键名不区分大小写。
'    要对键执行已处理的注册表操作，请调用regcreatekeytransacting函数。
'    备份或恢复系统状态(包括系统文件和注册表组)的应用程序应该使用卷影复制服务，而不是注册表函数。
'@原型
'    LSTATUS RegCreateKeyExA(
'      HKEY                        hKey,
'      LPCWSTR                      lpSubKey,
'      DWORD                       Reserved,
'      LPSTR                       lpClass,
'      DWORD                       dwOptions,
'      REGSAM                      samDesired,
'      const LPSECURITY_ATTRIBUTES lpSecurityAttributes,
'      PHKEY                       phkResult,
'      LPDWORD lpdwDisposition
'    );
'@参数
'hKey
'    打开注册表项的句柄。调用进程必须具有对键的KEY_CREATE_SUB_KEY访问权。有关更多信息，请参见注册表键安全性和访问权限。
'    键创建的访问是根据注册表键的安全描述符进行检查的，而不是在获取句柄时指定的访问掩码。因此，即使hKey是用samDesired of KEY_READ打开的，如果它的安全描述符允许，也可以在修改注册表的操作中使用它。
'    这个句柄由RegCreateKeyEx或RegOpenKeyEx函数返回，也可以是以下预定义键之一:
'        HKEY_CLASSES_ROOT
'        HKEY_CURRENT_CONFIG
'        HKEY_CURRENT_USER
'        HKEY_LOCAL_MACHINE
'        HKEY_USERS
'lpSubKey
'    此函数打开或创建的子键的名称。指定的子键必须是由hKey参数标识的键的子键;它可以在注册表树的最深处达到32层。有关键名的更多信息，请参见注册表结构。
'    如果lpSubKey是指向空字符串的指针，phkResult将接收hKey指定的键的新句柄。
'    此参数不能为空。
'Reserved
'    这个参数是保留的，必须为零。
'lpClass
'    此键的用户定义类类型。这个参数可以忽略。这个参数可以是NULL。
'dwOptions
'    此参数可以是以下值之一。
Private Const REG_OPTION_BACKUP_RESTORE As Long = &H4
'    如果设置了此标志，函数将忽略samDesired参数，并尝试打开键，以获得备份或恢复键所需的访问权限。如果调用线程启用了SE_BACKUP_NAME特权，则使用ACCESS_SYSTEM_SECURITY和KEY_READ访问权限打开键。如果调用线程启用了SE_RESTORE_NAME特权(从Windows Vista开始)，则使用ACCESS_SYSTEM_SECURITY。DELETE和KEY_WRITE访问权限打开键。如果启用了这两种特权，则键具有这两种特权的组合访问权限。有关更多信息，请参见使用特权运行。
Private Const REG_OPTION_CREATE_LINK    As Long = &H2
'    注: 注册表符号链接只能在绝对必要时用于应用程序兼容性。
'    这个键是一个符号链接。目标路径被分配给键的L“SymbolicLinkValue”值。目标路径必须是绝对注册表路径。
Private Const REG_OPTION_NON_VOLATILE   As Long = &H0
'    这个键不是易失的;这是默认值。信息存储在一个文件中，并在系统重新启动时保存。RegSaveKey函数保存非易失性的键。
Private Const REG_OPTION_VOLATILE       As Long = &H1
'    该函数创建的所有键都是易失性的。信息存储在内存中，在卸载相应的注册表单元格时不保存。对于HKEY_LOCAL_MACHINE，只有在系统启动完全关闭时才会发生这种情况。对于RegLoadKey函数加载的注册表项，当执行相应的RegUnLoadKey时，会发生这种情况。RegSaveKey函数不保存易失性键。对于已经存在的键，此标志将被忽略。
'    注意，对于用户选择的关机，快速启动关机是系统的默认行为。
'samDesired
'    指定要创建的键的访问权限的掩码。有关更多信息，请参见注册表键安全性和访问权限。
'lpSecurityAttributes
'    指向SECURITY_ATTRIBUTES结构的指针，该结构确定返回的句柄是否可以由子进程继承。如果lpSecurityAttributes为空，则不能继承句柄。
'    结构的lpSecurityDescriptor成员为新键指定一个安全描述符。如果lpSecurityAttributes为空，则键将获得默认的安全描述符。键的默认安全描述符中的acl是从其直接父键继承的。
'phkResult
'    指向一个变量的指针，该变量接收打开或创建的键的句柄。如果键不是预定义的注册表键之一，则在使用句柄之后调用RegCloseKey函数。
'lpdwDisposition
'    指向接受下列配置值之一的变量的指针。
Private Const REG_CREATED_NEW_KEY       As Long = &H1
'    键并不存在，而是被创建的。
Private Const REG_OPENED_EXISTING_KEY   As Long = &H2
'    键是存在的，只是打开了，没有改变。
'    如果lpdwdispose为空，则不返回任何配置信息。
'@返回值
'    如果函数成功，返回值为ERROR_SUCCESS。
'    如果函数失败，返回值是在Winerror.h中定义的非零错误代码。您可以使用带有FORMAT_MESSAGE_FROM_SYSTEM标志的FormatMessage函数来获得错误的一般描述。
'@备注
'    RegCreateKeyEx函数创建的键没有值。应用程序可以使用RegSetValueEx函数设置键值。
'    RegCreateKeyEx函数在指定的路径中创建所有丢失的键。应用程序可以利用这种行为一次创建多个键。例如，一个应用程序可以通过为lpSubKey参数指定如下形式的字符串，同时创建一个四层深的子键，与前面的三个子键相同:
'    subkey1 \ subkey2 \ subkey3 \ subkey4
'    注意，如果路径中的现有键拼写错误，此行为将导致创建不需要的键。
'    应用程序不能创建HKEY_USERS或HKEY_LOCAL_MACHINE的直接子键。应用程序可以在HKEY_USERS或HKEY_LOCAL_MACHINE树的较低层创建子键。
'    如果您的服务或应用程序模拟不同的用户，则不要将此函数与HKEY_CURRENT_USER一起使用。相反，调用RegOpenCurrentUser函数。
'    注意，访问某些注册表项的操作将被重定向。有关更多信息，请参见注册表虚拟化和注册表中的32位和64位应用程序数据。
'@Requirements
'    Minimum supported client    Windows 2000 Professional [desktop apps only]
'    Minimum supported server    Windows 2000 Server [desktop apps only]
'    Target Platform Windows
'    Header  winreg.h (include Windows.h)
'    Library advapi32.lib
'    dll advapi32.dll
Private Declare Function RegCreateKeyTransacted Lib "advapi32.dll" Alias "RegCreateKeyTransactedW" (ByVal hKey As Long, ByVal lpSubKey As Long, ByVal Reserved As Long, ByVal lpClass As Long, ByVal dwOptions As Long, ByVal samDesired As Long, lpSecurityAttributes As SECURITY_ATTRIBUTES, phkResult As Long, lpdwDisposition As Long, ByVal hTransaction As Long, pExtendedParemeter As Long) As Long
'@功能
'    创建指定的注册表项并将其与事务关联。如果键已经存在，函数将打开它。注意，键名不区分大小写。
'    备份或恢复系统状态(包括系统文件和注册表组)的应用程序应该使用卷影复制服务，而不是注册表函数。
'@原型
'    LSTATUS RegCreateKeyTransactedA(
'      HKEY                        hKey,
'      LPCWSTR                      lpSubKey,
'      DWORD                       Reserved,
'      LPSTR                       lpClass,
'      DWORD                       dwOptions,
'      REGSAM                      samDesired,
'      const LPSECURITY_ATTRIBUTES lpSecurityAttributes,
'      PHKEY                       phkResult,
'      LPDWORD                     lpdwDisposition,
'      HANDLE                      hTransaction,
'      PVOID pExtendedParemeter
'    );
'@参数
'hKey
'    打开注册表项的句柄。调用进程必须具有对键的KEY_CREATE_SUB_KEY访问权。有关更多信息，请参见注册表键安全性和访问权限。
'    键创建的访问是根据注册表键的安全描述符进行检查的，而不是在获取句柄时指定的访问掩码。因此，即使使用samDesired of KEY_READ打开hKey，如果安全描述符允许，也可以在创建键的操作中使用它。
'    此句柄由regcreatekeytransacting或regopenkeytransacting函数返回，也可以是以下预定义键之一:
'        HKEY_CLASSES_ROOT
'        HKEY_CURRENT_CONFIG
'        HKEY_CURRENT_USER
'        HKEY_LOCAL_MACHINE
'        HKEY_USERS
'lpSubKey
'    此函数打开或创建的子键的名称。指定的子键必须是由hKey参数标识的键的子键;它可以在注册表树的最深处达到32层。有关键名的更多信息，请参见注册表结构。
'    如果lpSubKey是指向空字符串的指针，phkResult将接收hKey指定的键的新句柄。
'    此参数不能为空。
'Reserved
'    这个参数是保留的，必须为零。
'lpClass
'    此键的用户定义类。这个参数可以忽略。这个参数可以是NULL。
'dwOptions
'    此参数可以是以下值之一。
'    REG_OPTION_BACKUP_RESTORE
'    如果设置了此标志，函数将忽略samDesired参数，并尝试打开键，以获得备份或恢复键所需的访问权限。如果调用线程启用了SE_BACKUP_NAME特权，则使用ACCESS_SYSTEM_SECURITY和KEY_READ访问权限打开键。如果调用线程启用了SE_RESTORE_NAME特权，则使用ACCESS_SYSTEM_SECURITY和KEY_WRITE访问权限打开键。如果启用了这两种特权，则键具有这两种特权的组合访问权限。有关更多信息，请参见使用特权运行。
'    REG_OPTION_NON_VOLATILE
'    这个键不是易失的;这是默认值。信息存储在一个文件中，并在系统重新启动时保存。RegSaveKey函数保存非易失性的键。
'    REG_OPTION_VOLATILE
'    该函数创建的所有键都是易失性的。信息存储在内存中，在卸载相应的注册表单元格时不保存。对于HKEY_LOCAL_MACHINE，这发生在系统关闭时。对于RegLoadKey函数加载的注册表项，当执行相应的RegUnLoadKey时，会发生这种情况。RegSaveKey函数不保存易失性键。对于已经存在的键，此标志将被忽略。
'samDesired
'    指定要创建的键的访问权限的掩码。有关更多信息，请参见注册表键安全性和访问权限。
'lpSecurityAttributes
'    指向SECURITY_ATTRIBUTES结构的指针，该结构确定返回的句柄是否可以由子进程继承。如果lpSecurityAttributes为空，则不能继承句柄。
'    结构的lpSecurityDescriptor成员为新键指定一个安全描述符。如果lpSecurityAttributes为空，则键将获得默认的安全描述符。键的默认安全描述符中的acl是从其直接父键继承的。
'phkResult
'    指向一个变量的指针，该变量接收打开或创建的键的句柄。如果键不是预定义的注册表键之一，则在使用句柄之后调用RegCloseKey函数。
'lpdwDisposition
'    指向接受下列配置值之一的变量的指针。
'    价值意义
'    REG_CREATED_NEW_KEY
'    键并不存在，而是被创建的。
'    REG_OPENED_EXISTING_KEY
'    键是存在的，只是打开了，没有改变。
'    如果lpdwdispose为空，则不返回任何配置信息。
'hTransaction
'    活动事务的句柄。这个句柄由CreateTransaction函数返回。
'pExtendedParemeter
'    保留此参数，且必须为空。
'@返回值
'    如果函数成功，返回值为ERROR_SUCCESS。
'    如果函数失败，返回值是在Winerror.h中定义的非零错误代码。您可以使用带有FORMAT_MESSAGE_FROM_SYSTEM标志的FormatMessage函数来获得错误的一般描述。
'@备注
'    当使用此函数创建键时，将处理对该键的后续操作。如果在提交事务之前对键执行非事务操作，则回滚事务。在提交或回滚事务之后，必须使用RegCreateKeyTransacted或RegOpenKeyTransacted(使用活动事务句柄)重新打开键，以处理其他操作。有关事务的更多信息，请参见内核事务管理器。
'    注意，此键的子键上的后续操作不会自动处理。因此，RegDeleteKeyEx不执行事务删除操作。相反，使用regdeletekeytransacting函数执行一个事务删除操作。
'    regcreatekeytransaction函数创建的键没有值。应用程序可以使用RegSetValueEx函数设置键值。
'    regcreatekeytransaction函数在指定的路径中创建所有丢失的键。应用程序可以利用这种行为一次创建多个键。例如，一个应用程序可以通过为lpSubKey参数指定如下形式的字符串，同时创建一个四层深的子键，与前面的三个子键相同:
'    subkey1 \ subkey2 \ subkey3 \ subkey4
'    注意，如果路径中的现有键拼写错误，此行为将导致创建不需要的键。
'    应用程序不能创建HKEY_USERS或HKEY_LOCAL_MACHINE的直接子键。应用程序可以在HKEY_USERS或HKEY_LOCAL_MACHINE树的较低层创建子键。
'@Requirements
'    Minimum supported client    Windows Vista [desktop apps only]
'    Minimum supported server    Windows Server 2008 [desktop apps only]
'    Target Platform Windows
'    Header  winreg.h (include Windows.h)
'    Library advapi32.lib
'    dll advapi32.dll
Private Declare Function RegConnectRegistry Lib "advapi32.dll" Alias "RegConnectRegistryW" (ByVal hKey As Long, ByVal lpSubKey As Long, phkResult As Long) As Long
'@功能
'    在另一台计算机上建立到预定义注册表项的连接。
'@原型
'    LSTATUS RegConnectRegistryA(
'      LPCWSTR lpMachineName,
'      HKEY   hKey,
'      PHKEY phkResult
'    );
'@参数
'lpMachineName
'    远程计算机的名称。字符串的形式如下:
'    \ computername
'    调用者必须能够访问远程计算机，否则该函数将失败。
'    如果该参数为NULL，则使用本地计算机名。
'hKey
'    预定义的注册表句柄。此参数可以是远程计算机上的下列预定义键之一。
'    HKEY_LOCAL_MACHINE HKEY_PERFORMANCE_DATA HKEY_USERS
'phkResult
'    指向一个变量的指针，该变量接收一个键句柄，该键句柄标识远程计算机上预定义的句柄。
'@返回值
'    如果函数成功，返回值为ERROR_SUCCESS。
'    如果函数失败，返回值是在Winerror.h中定义的非零错误代码。您可以使用带有FORMAT_MESSAGE_FROM_SYSTEM标志的FormatMessage函数来获得错误的一般描述。
'@备注
'    RegConnectRegistry要求远程注册中心服务在远程计算机上运行。默认情况下，将此服务配置为手动启动。要将远程注册中心服务配置为自动启动，请运行服务。并将服务的启动类型更改为Automatic。
'    Windows Server 2003和Windows XP/2000:远程注册表服务默认配置为自动启动。
'    当不再需要RegConnectRegistry返回的句柄时，应该通过调用RegCloseKey关闭它。
'    如果计算机连接到一个工作组，并且启用了“使用本地帐户身份验证为Guest的强制网络登录”策略，则该功能将失败。请注意，默认情况下，如果计算机连接到工作组，则启用此策略。
'    如果当前用户无法正确访问远程计算机，则对RegConnectRegistry的调用将失败。要连接到远程注册表，在调用RegConnectRegistry之前，使用LOGON32_LOGON_NEW_CREDENTIALS和ImpersonateLoggedOnUser调用LogonUser。
'    Windows 2000:一种可能的解决方案是使用一组不同的凭证建立一个到管理共享(如IPC$)的会话。要指定当前用户以外的凭据，请使用WNetAddConnection2函数连接到共享。访问完注册表后，取消连接。
'    Windows XP家庭版:不能使用此功能连接到运行Windows XP家庭版的远程计算机。这个函数可以处理本地计算机的名称，即使它运行的是Windows XP Home Edition，因为它绕过了身份验证层。
'@Requirements
'    Minimum supported client    Windows 2000 Professional [desktop apps only]
'    Minimum supported server    Windows 2000 Server [desktop apps only]
'    Target Platform Windows
'    Header  winreg.h (include Windows.h)
'    Library advapi32.lib
'    dll advapi32.dll
Private Declare Function RegCopyTree Lib "advapi32.dll" Alias "RegCopyTreeW" (ByVal hKeySrc As Long, ByVal lpSubKey As Long, ByVal hKeyDest As Long) As Long
'@功能
'    将指定的注册表项及其值和子项复制到指定的目标项。
'@原型
'    LSTATUS RegCopyTreeA(
'      HKEY   hKeySrc,
'      LPCWSTR lpSubKey,
'      hKey hKeyDest
'    );
'@参数
'hKeySrc
'    打开注册表项的句柄。键必须是用KEY_READ访问权限打开的。有关更多信息，请参见注册表键安全性和访问权限。
'    这个句柄由RegCreateKeyEx或RegOpenKeyEx函数返回，也可以是预定义的键之一。
'lpSubKey
'    键的名称。此键必须是由hKeySrc参数标识的键的子键。这个参数也可以是NULL。
'hKeyDest
'    目标键的句柄。调用进程必须具有对键的KEY_CREATE_SUB_KEY访问权。
'    这个句柄由RegCreateKeyEx或RegOpenKeyEx函数返回，也可以是预定义的键之一
'@返回值
'    如果函数成功，返回值为ERROR_SUCCESS。
'    如果函数失败，返回值是在Winerror.h中定义的非零错误代码。您可以使用带有FORMAT_MESSAGE_FROM_SYSTEM标志的FormatMessage函数来获得错误的一般描述。
'@备注
'    此函数还复制键的安全描述符。
'    要编译使用此函数的应用程序，请将_WIN32_WINNT定义为0x0600或更高。有关更多信息，请参见使用Windows头文件。
'@Requirements
'    Minimum supported client    Windows Vista [desktop apps only]
'    Minimum supported server    Windows Server 2008 [desktop apps only]
'    Target Platform Windows
'    Header  winreg.h (include Windows.h)
'    Library advapi32.lib
'    dll advapi32.dll
Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyW" (ByVal hKey As Long, ByVal lpSubKey As Long) As Long
'@功能
'  删除子键及其值。注意，键名不区分大小写。
'    64位Windows:在WOW64上，32位应用程序查看一个与64位应用程序查看的注册表树分离的注册表树。要使应用程序能够删除备用注册表视图中的条目，请使用RegDeleteKeyEx函数。
'@原型
'    LSTATUS RegDeleteKeyA(
'      HKEY   hKey,
'      LPCWSTR lpSubKey
'    );
'@参数
'hKey
'    打开注册表项的句柄。此键的访问权限不影响删除操作。有关访问权限的更多信息，请参见注册表项安全性和访问权限。
'    这个句柄由RegCreateKeyEx或RegOpenKeyEx函数返回，也可以是以下预定义键之一:
'        HKEY_CLASSES_ROOT
'        HKEY_CURRENT_CONFIG
'        HKEY_CURRENT_USER
'        HKEY_LOCAL_MACHINE
'        HKEY_USERS
'lpSubKey
'    要删除的键的名称。它必须是hKey标识的键的子键，但不能有子键。此参数不能为空。
'    函数打开具有删除访问权限的子键。
'    键名不区分大小写。
'    有关更多信息，请参见注册表元素大小限制。
'@返回值
'    如果函数成功，返回值为ERROR_SUCCESS。
'    如果函数失败，返回值是在Winerror.h中定义的非零错误代码。要获得错误的一般描述，可以使用带有FORMAT_MESSAGE_FROM_SYSTEM标志的FormatMessage函数。
'@备注
'    直到最后一个句柄关闭后，才会删除已删除的键。
'    要删除的子键必须没有子键。要删除键及其所有子键，需要枚举子键并分别删除它们。要递归删除键，请使用RegDeleteTree或SHDeleteKey函数。
'@Requirements
'    Minimum supported client    Windows 2000 Professional [desktop apps only]
'    Minimum supported server    Windows 2000 Server [desktop apps only]
'    Target Platform Windows
'    Header  winreg.h (include Windows.h)
'    Library advapi32.lib
'    dll advapi32.dll
Private Declare Function RegDeleteKeyEx Lib "advapi32.dll" Alias "RegDeleteKeyExW" (ByVal hKey As Long, ByVal lpSubKey As Long, ByVal samDesired As Long, ByVal Reserved As Long) As Long
'@功能
'    从注册表的特定于平台的指定视图中删除子键及其值。注意，键名不区分大小写。
'    要删除作为事务操作的子键，请调用RegDeleteKeyTransacted函数。
'@原型
'    LSTATUS RegDeleteKeyExA(
'      HKEY   hKey,
'      LPCWSTR lpSubKey,
'      REGSAM samDesired,
'      DWORD Reserved
'    );
'@参数
'hKey
'    打开注册表项的句柄。此键的访问权限不影响删除操作。有关访问权限的更多信息，请参见注册表项安全性和访问权限。
'    这个句柄由RegCreateKeyEx或RegOpenKeyEx函数返回，也可以是以下预定义键之一:
'        HKEY_CLASSES_ROOT
'        HKEY_CURRENT_CONFIG
'        HKEY_CURRENT_USER
'        HKEY_LOCAL_MACHINE
'        HKEY_USERS
'lpSubKey
'    要删除的键的名称。此键必须是由hKey参数的值指定的键的子键。
'    函数打开具有删除访问权限的子键。
'    键名不区分大小写。
'    此参数的值不能为空。
'samDesired
'    访问掩码指定注册表的特定于平台的视图。
Private Const KEY_WOW64_32KEY           As Long = &H200
'    从32位注册表视图中删除键。
Private Const KEY_WOW64_64KEY           As Long = &H100
'    从64位注册表视图中删除键。
'Reserved
'    这个参数是保留的，必须为零。
'@返回值
'    如果函数成功，返回值为ERROR_SUCCESS。
'    如果函数失败，返回值是在Winerror.h中定义的非零错误代码。您可以使用带有FORMAT_MESSAGE_FROM_SYSTEM标志的FormatMessage函数来获得错误的一般描述。
'@备注
'    直到最后一个句柄关闭后，才会删除已删除的键。
'    在WOW64上，32位应用程序查看一个与64位应用程序查看的注册表树分离的注册表树。此函数允许应用程序删除备用注册表视图中的条目。
'    要删除的子键必须没有子键。要删除键及其所有子键，需要枚举子键并分别删除它们。要递归删除键，请使用RegDeleteTree或SHDeleteKey函数。
'    如果函数成功，RegDeleteKeyEx将从注册表中删除指定的键。删除整个键，包括它的所有值。
'@Requirements
'    Minimum supported client    Windows Vista, Windows XP Professional x64 Edition [desktop apps only]
'    Minimum supported server    Windows Server 2008, Windows Server 2003 with SP1 [desktop apps only]
'    Target Platform Windows
'    Header  winreg.h (include Windows.h)
'    Library advapi32.lib
'    dll advapi32.dll
Private Declare Function RegDeleteKeyTransacted Lib "advapi32.dll" Alias "RegDeleteKeyTransactedW" (ByVal hKey As Long, ByVal lpSubKey As Long, ByVal samDesired As Long, ByVal Reserved As Long, ByVal hTransaction As Long, pExtendedParameter As Long) As Long
'@功能
'    从作为事务操作的注册表的特定于平台的指定视图中删除子键及其值。注意，键名不区分大小写。
'@原型
'    LSTATUS RegDeleteKeyTransactedA(
'      HKEY   hKey,
'      LPCWSTR lpSubKey,
'      REGSAM samDesired,
'      DWORD  Reserved,
'      HANDLE hTransaction,
'      PVOID pExtendedParameter
'    );
'@参数
'hKey
'    打开注册表项的句柄。此键的访问权限不影响删除操作。有关访问权限的更多信息，请参见注册表项安全性和访问权限。
'    这个句柄由RegCreateKeyEx。RegCreateKeyTransacted。RegOpenKeyEx或RegOpenKeyTransacted函数返回。它也可以是下列预定义键之一:
'        HKEY_CLASSES_ROOT
'        HKEY_CURRENT_CONFIG
'        HKEY_CURRENT_USER
'        HKEY_LOCAL_MACHINE
'        HKEY_USERS
'lpSubKey
'    要删除的键的名称。此键必须是由hKey参数的值指定的键的子键。
'    函数打开具有删除访问权限的子键。
'    键名不区分大小写。
'    此参数的值不能为空。
'samDesired
'    访问掩码指定注册表的特定于平台的视图。
'        KEY_WOW64_32KEY
'            从32位注册表视图中删除键。
'        KEY_WOW64_64KEY
'            从64位注册表视图中删除键。
'samDesired
'    这个参数是保留的，必须为零。
'hTransaction
'    活动事务的句柄。这个句柄由CreateTransaction函数返回。
'pExtendedParameter
'    保留此参数，且必须为空。
'@返回值
'    如果函数成功，返回值为ERROR_SUCCESS。
'    如果函数失败，返回值是在Winerror.h中定义的非零错误代码。您可以使用带有FORMAT_MESSAGE_FROM_SYSTEM标志的FormatMessage函数来获得错误的一般描述。
'@备注
'    直到最后一个句柄关闭后，才会删除已删除的键。
'    在WOW64上，32位应用程序查看一个与64位应用程序查看的注册表树分离的注册表树。此函数允许应用程序删除备用注册表视图中的条目。
'    要删除的子键必须没有子键。要删除键及其所有子键，需要枚举子键并分别删除它们。要递归删除键，请使用RegDeleteTree或SHDeleteKey函数。
'    如果函数成功，regdeletekeytransaction将从注册表中删除指定的键。删除整个键，包括它的所有值。要将整个树作为事务操作删除，请使用RegDeleteTree函数，并使用从RegCreateKeyTransacted或RegOpenKeyTransacted返回的句柄。
'@Requirements
'    Minimum supported client    Windows Vista [desktop apps only]
'    Minimum supported server    Windows Server 2008 [desktop apps only]
'    Target Platform Windows
'    Header  winreg.h (include Windows.h)
'    Library advapi32.lib
'    dll advapi32.dll
Private Declare Function RegDeleteTree Lib "advapi32.dll" Alias "RegDeleteTreeW" (ByVal hKey As Long, ByVal lpSubKey As Long) As Long
'@功能
'    递归地删除指定键的子键和值。
'@原型
'    LSTATUS RegDeleteTreeA(
'      HKEY   hKey,
'      LPCWSTR lpSubKey
'    );
'@参数
'hKey
'    打开注册表项的句柄。必须使用以下访问权限打开键:DELETE。KEY_ENUMERATE_SUB_KEYS和KEY_QUERY_VALUE。有关更多信息，请参见注册表键安全性和访问权限。
'    这个句柄由RegCreateKeyEx或RegOpenKeyEx函数返回，也可以是以下预定义键之一:
'        HKEY_CLASSES_ROOT
'        HKEY_CURRENT_CONFIG
'        HKEY_CURRENT_USER
'        HKEY_LOCAL_MACHINE
'        HKEY_USERS
'lpSubKey
'    键的名称。此键必须是由hKey参数标识的键的子键。如果该参数为NULL，则删除hKey的子键和值。
'@返回值
'    如果函数成功，返回值为ERROR_SUCCESS。
'    如果函数失败，返回值是在Winerror.h中定义的非零错误代码。您可以使用带有FORMAT_MESSAGE_FROM_SYSTEM标志的FormatMessage函数来获得错误的一般描述。
'@备注
'    如果键有值，则必须用KEY_SET_VALUE打开它，否则该函数将失败，ERROR_ACCESS_DENIED将被拒绝。
'    要编译使用此函数的应用程序，请将_WIN32_WINNT定义为0x0600或更高。有关更多信息，请参见使用Windows头文件
'@Requirements
'    Minimum supported client    Windows Vista [desktop apps only]
'    Minimum supported server    Windows Server 2008 [desktop apps only]
'    Target Platform Windows
'    Header  winreg.h (include Windows.h)
'    Library advapi32.lib
'    dll advapi32.dll
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueW" (ByVal hKey As Long, ByVal lpValueName As Long) As Long
'@功能
'    从指定的注册表项中移除命名值。注意，值名称不区分大小写。
'@原型
'    LSTATUS RegDeleteValueA(
'      HKEY   hKey,
'      LPCWSTR lpValueName
'    );
'@参数
'hKey
'    打开注册表项的句柄。键必须是用KEY_SET_VALUE访问权限打开的。有关更多信息，请参见注册表键安全性和访问权限。
'    这个句柄由RegCreateKeyEx。RegCreateKeyTransacted。RegOpenKeyEx或RegOpenKeyTransacted函数返回。它也可以是下列预定义键之一:
'        HKEY_CLASSES_ROOT
'        HKEY_CURRENT_CONFIG
'        HKEY_CURRENT_USER
'        HKEY_LOCAL_MACHINE
'        HKEY_USERS
'lpValueName
'    要删除的注册表值。如果该参数为NULL或空字符串，则删除RegSetValue函数设置的值。
'    有关更多信息，请参见注册表元素大小限制。
'@返回值
'    如果函数成功，返回值为ERROR_SUCCESS。
'    如果函数失败，返回值是在Winerror.h中定义的非零错误代码。您可以使用带有FORMAT_MESSAGE_FROM_SYSTEM标志的FormatMessage函数来获得错误的一般描述。
'@Requirements
'    Minimum supported client    Windows 2000 Professional [desktop apps only]
'    Minimum supported server    Windows 2000 Server [desktop apps only]
'    Target Platform Windows
'    Header  winreg.h (include Windows.h)
'    Library advapi32.lib
'    dll advapi32.dll

Private Declare Function RegEnumKey Lib "advapi32.dll" Alias "RegEnumKeyW" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As Long, ByVal cbName As Long) As Long
'@功能
'    枚举指定的开放注册表项的子键。函数在每次调用一个子键时检索该子键的名称。
'    注意，此函数仅用于兼容16位版本的Windows。应用程序应该使用RegEnumKeyEx函数。
'@原型
'    LSTATUS RegEnumKeyA(
'      HKEY  hKey,
'      DWORD dwIndex,
'      LPSTR lpName,
'      DWORD cchName
'    );
'@参数
'hKey
'    打开注册表项的句柄。键必须是用KEY_ENUMERATE_SUB_KEYS访问权限打开的。有关更多信息，请参见注册表键安全性和访问权限。
'    这个句柄由RegCreateKeyEx。RegCreateKeyTransacted。RegOpenKeyEx或RegOpenKeyTransacted函数返回。它也可以是下列预定义键之一:
'        HKEY_CLASSES_ROOT
'        HKEY_CURRENT_CONFIG
'        HKEY_CURRENT_USER
'        HKEY_LOCAL_MACHINE
'        HKEY_USERS
'dwIndex
'    要检索的hKey的子键的索引。对于第一次调用RegEnumKey函数，这个值应该为零，然后在后续调用中递增。
'    因为子键是无序的，所以任何新的子键都有一个任意的索引。这意味着函数可以以任何顺序返回子键。
'lpName
'    指向缓冲区的指针，该缓冲区接收子键的名称，包括终止空字符。这个函数只将子键的名称复制到缓冲区，而不是完整的键层次结构。
'    有关更多信息，请参见注册表元素大小限制。
'cchName
'    在TCHARs中，由lpName参数指向的缓冲区的大小。要确定所需的缓冲区大小，请使用RegQueryInfoKey函数来确定由hKey参数标识的键的最大子键的大小。
'@返回值
'    如果函数成功，返回值为ERROR_SUCCESS。
'    如果函数失败，返回值是一个系统错误代码。如果没有更多可用的子键，函数将返回ERROR_NO_MORE_ITEMS。
'    如果lpName缓冲区太小，无法接收键名，则函数返回ERROR_MORE_DATA。
'@备注
'    要枚举子键，应用程序首先应该调用RegEnumKey函数，并将dwIndex参数设置为零。然后，应用程序应该增加dwIndex参数并调用RegEnumKey函数，直到没有更多的子键(这意味着函数返回ERROR_NO_MORE_ITEMS)。
'    应用程序还可以在第一次调用函数时将dwIndex设置为最后一个键的索引，并递减索引，直到枚举出索引为0的子键为止。要检索最后一个子键的索引，请使用RegQueryInfoKey。
'    当应用程序使用RegEnumKey函数时，它不应该调用任何可能更改正在查询的键的注册函数。
'@Requirements
'    Minimum supported client    Windows 2000 Professional [desktop apps only]
'    Minimum supported server    Windows 2000 Server [desktop apps only]
'    Target Platform Windows
'    Header  winreg.h (include Windows.h)
'    Library advapi32.lib
'    dll advapi32.dll
Private Type FILETIME
    dwLowDateTime   As Long
    dwHighDateTime  As Long
End Type
'@功能
'    包含一个64位值，表示自1601年1月1日(UTC)以来100纳秒间隔的数量。
'@原型
'    typedef struct _FILETIME {
'      DWORD dwLowDateTime;
'      DWORD dwHighDateTime;
'    } FILETIME, *PFILETIME, *LPFILETIME;
'@成员
'dwLowDateTime
'    文件时间的低阶部分。
'dwHighDateTime
'    文件时间的高阶部分。
'@备注
'    要将FILETIME结构转换为易于显示给用户的时间，可以使用FileTimeToSystemTime函数。
'    不建议从FILETIME结构中添加或减去值来获得相对时间。相反，应该将文件时间的低阶和高阶部分复制到ULARGE_INTEGER结构中，对QuadPart成员执行64位算术，并将LowPart和HighPart成员复制到FILETIME结构中。
'    不要将指向FILETIME结构的指针强制转换为ULARGE_INTEGER*或_int64*值，因为这会导致64位Windows上的对齐错误。
'    不是所有的文件系统都能记录创建和最后访问时间，也不是所有的文件系统都能以相同的方式记录它们。例如，在NT FAT上，创建时间的分辨率为10毫秒，写时间的分辨率为2秒，访问时间的分辨率为1天(实际上是访问日期)。在NTFS上，访问时间的分辨率为1小时。因此，GetFileTime函数可能不会返回使用SetFileTime函数设置的相同文件时间信息。此外，FAT以本地时间记录磁盘上的时间。但是，NTFS在UTC中记录磁盘上的时间。有关更多信息，请参见文件时间。
'    使用FILETIME结构的函数可以允许零或正值之外的值，这些值通常由dwLowDateTime和dwHighDateTime成员指定。例如，SetFileTime函数使用0xFFFFFFFF指定应该保留文件之前的访问时间。有关更多信息，请参阅正在调用的函数的主题
'@Requirements
'    Minimum supported client    Windows 2000 Professional [desktop apps | UWP apps]
'    Minimum supported server    Windows 2000 Server [desktop apps | UWP apps]
'    Header  minwinbase.h (include Windows.h)
Private Declare Function RegEnumKeyEx Lib "advapi32.dll" Alias "RegEnumKeyExW" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As Long, lpcbName As Long, lpReserved As Long, ByVal lpClass As Long, lpcbClass As Long, lpftLastWriteTime As FILETIME) As Long
'@功能
'    枚举指定的开放注册表项的子键。该函数在每次调用一个子键时检索关于该子键的信息。
'@原型
'    LSTATUS RegEnumKeyExA(
'      HKEY      hKey,
'      DWORD     dwIndex,
'      LPSTR     lpName,
'      LPDWORD   lpcchName,
'      LPDWORD   lpReserved,
'      LPSTR     lpClass,
'      LPDWORD   lpcchClass,
'      PFILETIME lpftLastWriteTime
'    );
'@参数
'hKey
'    打开注册表项的句柄。键必须是用KEY_ENUMERATE_SUB_KEYS访问权限打开的。有关更多信息，请参见注册表键安全性和访问权限。
'    这个句柄由RegCreateKeyEx。RegCreateKeyTransacted。RegOpenKeyEx或RegOpenKeyTransacted函数返回。它也可以是下列预定义键之一:
'        HKEY_CLASSES_ROOT
'        HKEY_CURRENT_CONFIG
'        HKEY_CURRENT_USER
'        HKEY_LOCAL_MACHINE
'        HKEY_PERFORMANCE_DATA
'        HKEY_USERS
'dwIndex
'    要检索的子键的索引。对于第一次调用RegEnumKeyEx函数，该参数应该为零，然后在后续调用中递增。
'    因为子键是无序的，所以任何新的子键都有一个任意的索引。这意味着函数可以以任何顺序返回子键。
'lpName
'    指向缓冲区的指针，该缓冲区接收子键的名称，包括终止空字符。该函数只将子键的名称复制到缓冲区，而不是将整个键层次结构复制到缓冲区。
'    如果函数失败，则不会将任何信息复制到此缓冲区。
'    有关更多信息，请参见注册表元素大小限制。
'lpcchName
'    指向一个变量的指针，该变量指定由lpName参数指定的缓冲区大小，以字符为单位。这个大小应该包括终止空字符。如果函数成功，lpcName指向的变量包含存储在缓冲区中的字符数，不包括终止null字符。
'    要确定所需的缓冲区大小，请使用RegQueryInfoKey函数来确定由hKey参数标识的键的最大子键的大小。
'lpReserved
'    保留此参数，且必须为空。
'lpClass
'    指向缓冲区的指针，该缓冲区接收枚举子键的用户定义类。这个参数可以是NULL。
'lpcchClass
'    指向一个变量的指针，该变量指定由lpClass参数指定的缓冲区大小，以字符为单位。大小应该包括终止空字符。如果函数成功，lpcClass包含存储在缓冲区中的字符数，不包括终止null字符。只有当lpClass为空时，这个参数才可以为空。
'lpftLastWriteTime
'    一个指向FILETIME结构的指针，它接收上次写入枚举子键的时间。这个参数可以是NULL。
'@返回值
'    如果函数成功，返回值为ERROR_SUCCESS。
'    如果函数失败，返回值是一个系统错误代码。如果没有更多可用的子键，函数将返回ERROR_NO_MORE_ITEMS。
'    如果lpName缓冲区太小，无法接收键名，则函数返回ERROR_MORE_DATA。
'@备注
'    要枚举子键，应用程序首先应该调用RegEnumKeyEx函数，并将dwIndex参数设置为零。然后应用程序应该增加dwIndex参数并调用RegEnumKeyEx，直到没有更多的子键(这意味着函数返回ERROR_NO_MORE_ITEMS)。
'    应用程序还可以在第一次调用函数时将dwIndex设置为最后一个子键的索引，并递减索引，直到枚举出索引为0的子键为止。要检索最后一个子键的索引，可以使用RegQueryInfoKey函数。
'    当应用程序使用RegEnumKeyEx函数时，它不应该调用任何可能更改枚举的键的注册函数。
'    注意，访问某些注册表项的操作将被重定向。有关更多信息，请参见注册表虚拟化和注册表中的32位和64位应用程序数据。
'@Requirements
'    Minimum supported client    Windows 2000 Professional [desktop apps only]
'    Minimum supported server    Windows 2000 Server [desktop apps only]
'    Target Platform Windows
'    Header  winreg.h (include Windows.h)
'    Library advapi32.lib
'    dll advapi32.dll

Private Declare Function RegEnumValue Lib "advapi32" Alias "RegEnumValueW" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As Long, lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, ByVal lpData As Long, lpcbData As Long) As Long
'@功能
'    枚举指定的开放注册表项的值。函数在每次调用键时复制一个索引值名称和数据块
'@原型
'    LSTATUS RegEnumValueA(
'      HKEY    hKey,
'      DWORD   dwIndex,
'      LPSTR   lpValueName,
'      LPDWORD lpcchValueName,
'      LPDWORD lpReserved,
'      LPDWORD lpType,
'      LPBYTE  lpData,
'      LPDWORD lpcbData
'    );
'@参数
'hKey
'    打开注册表项的句柄。键必须是用KEY_QUERY_VALUE访问权限打开的。有关更多信息，请参见注册表键安全性和访问权限。
'    这个句柄由RegCreateKeyEx。RegCreateKeyTransacted。RegOpenKeyEx或RegOpenKeyTransacted函数返回。它也可以是下列预定义键之一:
'        HKEY_CLASSES_ROOT
'        HKEY_CURRENT_CONFIG
'        HKEY_CURRENT_USER
'        HKEY_LOCAL_MACHINE
'        HKEY_PERFORMANCE_DATA
'        HKEY_USERS
'dwIndex
'    要检索的值的索引。对于第一次调用RegEnumValue函数，该参数应该为零，然后在后续调用中递增。
'    因为值不是有序的，所以任何新值都有一个任意的索引。这意味着函数可以以任何顺序返回值。
'lpValueName
'    一个指向缓冲区的指针，该缓冲区以空结束字符串的形式接收值的名称。
'    此缓冲区必须足够大，以包含终止null字符。
'    有关更多信息，请参见注册表元素大小限制。
'lpcchValueName
'    指向一个变量的指针，该变量以字符为单位指定lpValueName参数所指向的缓冲区的大小。当函数返回时，变量接收存储在缓冲区中的字符数，不包括终止null字符。
'    注册表值名称限制为32,767字节。该函数的ANSI版本将此参数视为一个短值。因此，如果指定的值大于32,767字节，就会出现溢出，函数可能返回ERROR_MORE_DATA。
'lpReserved
'    保留此参数，且必须为空。
'lpType
'    指向一个变量的指针，该变量接收一个代码，该代码指示存储在指定值中的数据类型。有关可能的类型代码列表，请参见注册表值类型。如果不需要类型代码，则lpType参数可以为NULL。
'lpData
'    指向缓冲区的指针，该缓冲区接收值项的数据。如果不需要数据，此参数可以为NULL。
'    如果lpData为NULL, lpcbData为非NULL，则函数将数据的大小(以字节为单位)存储在lpcbData指向的变量中。这使应用程序能够确定为数据分配缓冲区的最佳方式。
'lpcbData
'    指向一个变量的指针，该变量指定lpData参数指向的缓冲区的大小，单位为字节。当函数返回时，变量接收存储在缓冲区中的字节数。
'    只有当lpData为空时，此参数才能为空。
'    如果数据具有REG_SZ。REG_MULTI_SZ或REG_EXPAND_SZ类型，则此大小包括任何一个或多个终止null字符。有关更多信息，请参见备注。
'    如果lpData指定的缓冲区不够大，无法容纳数据，则函数返回ERROR_MORE_DATA，并在lpcbData指向的变量中存储所需的缓冲区大小。在本例中，lpData的内容是未定义的。
'@返回值
'    如果函数成功，返回值为ERROR_SUCCESS。
'    如果函数失败，返回值是一个系统错误代码。如果没有更多可用的值，函数将返回ERROR_NO_MORE_ITEMS。
'    如果lpData缓冲区太小，无法接收该值，则函数返回ERROR_MORE_DATA。
'@备注
'    要枚举值，应用程序首先应该调用RegEnumValue函数，并将dwIndex参数设置为零。然后应用程序应该增加dwIndex并调用RegEnumValue函数，直到没有更多的值(直到函数返回ERROR_NO_MORE_ITEMS)。
'    应用程序还可以在第一次调用函数时将dwIndex设置为最后一个值的索引，并递减索引，直到枚举出索引为0的值为止。要检索最后一个值的索引，请使用RegQueryInfoKey函数。
'    在使用RegEnumValue时，应用程序不应该调用任何可能更改正在查询的键的注册表函数。
'    如果数据具有REG_SZ。REG_MULTI_SZ或REG_EXPAND_SZ类型，则字符串可能没有存储正确的空终止字符。因此，即使函数返回ERROR_SUCCESS，应用程序也应该确保字符串在使用之前已正确终止;否则，它可能会覆盖缓冲区。(注意REG_MULTI_SZ字符串应该有两个空终止字符。)
'    要确定名称和数据缓冲区的最大大小，请使用RegQueryInfoKey函数。
'@Requirements
'    Minimum supported client    Windows 2000 Professional [desktop apps only]
'    Minimum supported server    Windows 2000 Server [desktop apps only]
'    Target Platform Windows
'    Header  winreg.h (include Windows.h)
'    Library advapi32.lib
'    dll advapi32.dll
Private Declare Function RegFlushKey Lib "advapi32" (ByVal hKey As Long) As Long
'@功能
'    将指定的打开注册表项的所有属性写入注册表。
'@原型
'    LSTATUS RegFlushKey(
'      hKey hKey
'    );
'@参数
'hKey
'    打开注册表项的句柄。键必须是用KEY_QUERY_VALUE访问权限打开的。有关更多信息，请参见注册表键安全性和访问权限。
'    这个句柄由RegCreateKeyEx。RegCreateKeyTransacted。RegOpenKeyEx或RegOpenKeyTransacted函数返回。它也可以是下列预定义键之一:
'        HKEY_CLASSES_ROOT
'        HKEY_CURRENT_CONFIG
'        HKEY_CURRENT_USER
'        HKEY_LOCAL_MACHINE
'        HKEY_PERFORMANCE_DATA
'        HKEY_USERS
'@返回值
'    如果函数成功，返回值为ERROR_SUCCESS。
'    如果函数失败，返回值是在Winerror.h中定义的非零错误代码。您可以使用带有FORMAT_MESSAGE_FROM_SYSTEM标志的FormatMessage函数来获得错误的一般描述。
'@备注
'    调用RegFlushKey是一项昂贵的操作，它会显著地影响系统性能，因为它会消耗磁盘带宽，并阻塞registry hive中的所有进程对所有键的修改，这些进程正在被刷新，直到刷新操作完成。只有当应用程序必须保证注册表更改在修改后立即持久化到磁盘时，才应该显式地调用RegFlushKey。对键所做的所有修改对其他进程都是可见的，而不需要将它们刷新到磁盘。
'    或者，注册表有一个“延迟刷新”机制，定期将注册表修改刷新到磁盘。除了这个常规刷新操作之外，还会在系统关闭时将注册表更改刷新到磁盘。允许“延迟刷新”刷新注册表更改是管理注册表写入磁盘注册表存储的最有效方法。
'    只有当包含指定键的hive的所有数据都被写到磁盘上的注册表存储中时，RegFlushKey函数才会返回。
'    RegFlushKey函数为hive中自上次延迟刷新或系统启动以来修改过的其他键写出数据。
'    在RegFlushKey返回之后，使用RegCloseKey关闭注册表项的句柄。
'@Requirements
'    Minimum supported client    Windows 2000 Professional [desktop apps only]
'    Minimum supported server    Windows 2000 Server [desktop apps only]
'    Target Platform Windows
'    Header  winreg.h (include Windows.h)
'    Library advapi32.lib
'    dll advapi32.dll
Private Declare Function RegGetValue Lib "advapi32.dll" Alias "RegGetValueW" (ByVal hKey As Long, ByVal lpSubKey As Long, ByVal lpValue As Long, ByVal dwFlags As Long, pdwType As Long, lpData As Any, lpcbData As Long) As Long
'@功能
'    检索指定注册表值的类型和数据。
'@原型
'    LSTATUS RegGetValueA(
'      HKEY    hkey,
'      LPCWSTR  lpSubKey,
'      LPCWSTR  lpValue,
'      DWORD   dwFlags,
'      LPDWORD pdwType,
'      PVOID   pvData,
'      LPDWORD pcbData
'    );
'@参数
'hKey
'    打开注册表项的句柄。键必须是用KEY_QUERY_VALUE访问权限打开的。有关更多信息，请参见注册表键安全性和访问权限。
'    这个句柄由RegCreateKeyEx。RegCreateKeyTransacted。RegOpenKeyEx或RegOpenKeyTransacted函数返回。它也可以是下列预定义键之一:
'        HKEY_CLASSES_ROOT
'        HKEY_CURRENT_CONFIG
'        HKEY_CURRENT_USER
'        HKEY_LOCAL_MACHINE
'        HKEY_PERFORMANCE_DATA
'        HKEY_PERFORMANCE_NLSTEXT
'        HKEY_PERFORMANCE_TEXT
'        HKEY_USERS
'lpSubKey
'    注册表项的名称。此键必须是由hkey参数指定的键的子键。
'    键名不区分大小写。
'lpValue
'    注册表值的名称。
'    如果该参数为NULL或空字符串“”，则函数检索键的未命名值或默认值(如果有)的类型和数据。
'    有关更多信息，请参见注册表元素大小限制。
'    键不会自动具有未命名或默认值。未命名值可以是任何类型。
'dwFlags
'    限制要查询的值的数据类型的标志。如果值的数据类型不满足此条件，则函数将失败。此参数可以是以下值中的一个或多个。
Private Const RRF_RT_ANY                As Long = &HFFFF
'    任何类型的限制。
Private Const RRF_RT_DWORD              As Long = &H18
'    将类型限制为32位的RRF_RT_REG_BINARY | RRF_RT_REG_DWORD。
Private Const RRF_RT_QWORD              As Long = &H48
'    将类型限制为64位RRF_RT_REG_BINARY | RRF_RT_REG_DWORD。
Private Const RRF_RT_REG_BINARY         As Long = &H8
'    将类型限制为REG_BINARY。
Private Const RRF_RT_REG_DWORD          As Long = &H10
'    将类型限制为REG_DWORD。
Private Const RRF_RT_REG_EXPAND_SZ      As Long = &H4
'    将类型限制为REG_EXPAND_SZ。
Private Const RRF_RT_REG_MULTI_SZ       As Long = &H20
'    将类型限制为REG_MULTI_SZ。
Private Const RRF_RT_REG_NONE           As Long = &H1
'    将类型限制为REG_NONE。
Private Const RRF_RT_REG_QWORD          As Long = &H40
'    将类型限制为REG_QWORD。
Private Const RRF_RT_REG_SZ             As Long = &H2
'    将类型限制为REG_SZ。
'    此参数还可以包含以下一个或多个值。
Private Const RRF_NOEXPAND              As Long = &H10000000
'    如果值类型为REG_EXPAND_SZ，则不要自动展开环境字符串。
Private Const RRF_ZEROONFAILURE         As Long = &H20000000
'    如果pvData不为空，则在失败时将缓冲区的内容设置为零。
Private Const RRF_SUBKEY_WOW6464KEY     As Long = &H10000
'    如果lpSubKey不为空，则打开lpSubKey使用KEY_WOW64_64KEY访问权限指定的子键。有关这些访问权限的信息，请参见注册表项安全性和访问权限。
'    不能将RRF_SUBKEY_WOW6464KEY与RRF_SUBKEY_WOW6432KEY组合指定RRF_SUBKEY_WOW6464KEY。
Private Const RRF_SUBKEY_WOW6432KEY     As Long = &H20000
'    如果lpSubKey不为空，则打开lpSubKey使用KEY_WOW64_32KEY访问权限指定的子键。有关这些访问权限的信息，请参见注册表项安全性和访问权限。
'    不能将RRF_SUBKEY_WOW6432KEY与RRF_SUBKEY_WOW6464KEY组合指定。
'pdwType
'    指向一个变量的指针，该变量接收一个代码，该代码指示存储在指定值中的数据类型。有关可能的类型代码列表，请参见注册表值类型。如果类型不是必需的，则此参数可以为空。
'pvData
'    指向接收值数据的缓冲区的指针。如果不需要数据，此参数可以为NULL。
'    如果数据是字符串，函数将检查终止空字符。如果没有找到，如果缓冲区足够大，可以容纳额外的字符，则使用null终止符存储字符串。否则，函数将失败并返回ERROR_MORE_DATA。
'pcbData
'    指向变量的指针，该变量指定pvData参数指向的缓冲区的大小，单位为字节。当函数返回时，这个变量包含复制到pvData的数据的大小。
'    只有当pvData为空时，pcbData参数才能为空。
'    如果数据具有REG_SZ。REG_MULTI_SZ或REG_EXPAND_SZ类型，则此大小包括任何一个或多个终止null字符。有关更多信息，请参见备注。
'    如果pvData参数指定的缓冲区不够大，无法保存数据，则函数返回ERROR_MORE_DATA，并在pcbData指向的变量中存储所需的缓冲区大小。在本例中，pvData缓冲区的内容是未定义的。
'    如果pvData为空，而pcbData为非空，则函数返回ERROR_SUCCESS，并在pcbData指向的变量中存储数据的大小(以字节为单位)。这使应用程序能够确定为值的数据分配缓冲区的最佳方法。
'    如果hKey指定了HKEY_PERFORMANCE_DATA，并且pvData缓冲区不够大，无法包含所有返回的数据，则函数返回ERROR_MORE_DATA，并且通过pcbData参数返回的值未定义。这是因为性能数据的大小可以从一个调用更改到下一个调用。在这种情况下，必须增加缓冲区大小并再次调用RegGetValue，在pcbData参数中传递更新后的缓冲区大小。重复此操作，直到函数成功。您需要维护一个单独的变量来跟踪缓冲区大小，因为pcbData返回的值是不可预测的。
'@返回值
'    如果函数成功，返回值为ERROR_SUCCESS。
'    如果函数失败，返回值是一个系统错误代码。
'    如果pvData缓冲区太小，无法接收该值，则函数返回ERROR_MORE_DATA。
'    如果dwFlags指定RRF_SUBKEY_WOW6464KEY和RRF_SUBKEY_WOW6432KEY的组合，函数返回ERROR_INVALID_PARAMETER。
'@备注
'    应用程序通常调用RegEnumValue来确定值名称，然后调用RegGetValue来检索名称的数据。
'    如果数据有REG_SZ , REG_MULTI_SZ或REG_EXPAND_SZ类型, 并使用这个函数的ANSI版本(通过显式地调用RegGetValueA或不包括Windows.h文件之前定义UNICODE), 这个函数将存储UNICODE字符串转换为一个ANSI字符串复制到缓冲区之前由pvData指出。
'    当使用hkey设置为HKEY_PERFORMANCE_DATA句柄和指定对象的值字符串调用此函数时，返回的数据结构有时具有未请求的对象。不要惊讶;这是正常的行为。您应该始终期望遍历返回的数据结构来查找请求的对象。
'    注意，访问某些注册表项的操作将被重定向。有关更多信息，请参见注册表虚拟化和注册表中的32位和64位应用程序数据。
'    要编译使用此函数的应用程序，请将_WIN32_WINNT定义为0x0600或更高。有关更多信息，请参见使用Windows头文件。
'@Requirements
'    Minimum supported client    Windows Vista, Windows XP Professional x64 Edition [desktop apps only]
'    Minimum supported server    Windows Server 2008, Windows Server 2003 with SP1 [desktop apps only]
'    Target Platform Windows
'    Header  winreg.h (include Windows.h)
'    Library advapi32.lib
'    dll advapi32.dll
Private Declare Function RegOpenCurrentUser Lib "advapi32.dll" (ByVal samDesired As Long, phkResult As Long) As Long
'@功能
'    为当前线程正在模拟的用户检索HKEY_CURRENT_USER键的句柄。
'@原型
'    LSTATUS RegOpenCurrentUser(
'      REGSAM samDesired,
'      PHKEY phkResult
'    );
'@参数
'samDesired
'    指定键的所需访问权限的掩码。如果键的安全描述符不允许调用进程的请求访问，则该函数将失败。有关更多信息，请参见注册表键安全性和访问权限。
'phkResult
'    指向一个变量的指针，该变量接收打开的键的句柄。当不再需要返回句柄时，调用RegCloseKey函数关闭它。
'@返回值
'    如果函数成功，返回值为ERROR_SUCCESS。
'    如果函数失败，返回值是在Winerror.h中定义的非零错误代码。您可以使用带有FORMAT_MESSAGE_FROM_SYSTEM标志的FormatMessage函数来获得错误的一般描述。
'@备注
'    HKEY_CURRENT_USER键映射到HKEY_USERS键中当前用户分支的根。它为进程中的所有线程缓存。因此，当加载其他用户的概要文件时，此值不会更改。RegOpenCurrentUser使用线程的令牌访问适当的键，如果没有加载概要文件，则使用默认键。
'@Requirements
'    Minimum supported client    Windows 2000 Professional [desktop apps only]
'    Minimum supported server    Windows 2000 Server [desktop apps only]
'    Target Platform Windows
'    Header  winreg.h (include Windows.h)
'    Library advapi32.lib
'    dll advapi32.dll
Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyW" (ByVal hKey As Long, ByVal lpSubKey As Long, phkResult As Long) As Long
'@功能
'    打开指定的注册表项。
'    注意，此函数仅用于兼容16位版本的Windows。应用程序应该使用RegOpenKeyEx函数
'@原型
'    LSTATUS RegOpenKeyA(
'      HKEY   hKey,
'      LPCWSTR lpSubKey,
'      PHKEY phkResult
'    );
'@参数
'hKey
'    打开注册表项的句柄。这个句柄由RegCreateKeyEx或RegOpenKeyEx函数返回，也可以是以下预定义键之一:
'    HKEY_CLASSES_ROOT HKEY_CURRENT_CONFIG HKEY_CURRENT_USER HKEY_LOCAL_MACHINE HKEY_USERS
'lpSubKey
'    要打开的注册表项的名称。此键必须是由hKey参数标识的键的子键。
'    键名不区分大小写。
'    如果该参数为NULL或指向空字符串的指针，则函数返回传入的相同句柄。
'    有关更多信息，请参见注册表元素大小限制。
'phkResult
'    指向一个变量的指针，该变量接收打开的键的句柄。如果键不是预定义的注册表键之一，则在使用句柄之后调用RegCloseKey函数。
'@返回值
'    如果函数成功，返回值为ERROR_SUCCESS。
'    如果函数失败，返回值是在Winerror.h中定义的非零错误代码。您可以使用带有FORMAT_MESSAGE_FROM_SYSTEM标志的FormatMessage函数来获得错误的一般描述。
'@备注
'    RegOpenKey函数使用默认的安全访问掩码来打开键。如果打开键需要不同的访问权限，函数将失败，返回ERROR_ACCESS_DENIED。在这种情况下，应用程序应该使用RegOpenKeyEx函数来指定访问掩码。
'    如果键在数据库中不存在，则RegOpenKey不创建指定的键。
'    如果您的服务或应用程序模拟不同的用户，则不要将此函数与HKEY_CURRENT_USER一起使用。相反，调用RegOpenCurrentUser函数。
'@Requirements
'    Minimum supported client    Windows 2000 Professional [desktop apps only]
'    Minimum supported server    Windows 2000 Server [desktop apps only]
'    Target Platform Windows
'    Header  winreg.h (include Windows.h)
'    Library advapi32.lib
'    dll advapi32.dll
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExW" (ByVal hKey As Long, ByVal lpSubKey As Long, ByVal uloptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
'@功能
'    打开指定的注册表项。注意，键名不区分大小写。
'    要对键执行事务的注册表操作，请调用regopenkeytransacting函数。
'@原型
'    LSTATUS RegOpenKeyExA(
'      HKEY   hKey,
'      LPCWSTR lpSubKey,
'      DWORD  ulOptions,
'      REGSAM samDesired,
'      PHKEY phkResult
'    );
'@参数
'hKey
'    打开注册表项的句柄。这个句柄由RegCreateKeyEx或RegOpenKeyEx函数返回，也可以是以下预定义键之一:
'    HKEY_CLASSES_ROOT HKEY_CURRENT_CONFIG HKEY_CURRENT_USER HKEY_LOCAL_MACHINE HKEY_USERS
'lpSubKey
'    要打开的注册表子键的名称。
'    键名不区分大小写。
'    lpSubKey参数可以是指向空字符串的指针。如果lpSubKey是指向空字符串的指针，而hKey是HKEY_CLASSES_ROOT，那么phkResult将接收传递给函数的相同的hKey句柄。否则，phkResult将接收hKey指定的键的新句柄。
'    只有当hKey是预定义的键之一时，lpSubKey参数才能为空。如果lpSubKey为NULL, hKey为HKEY_CLASSES_ROOT, phkResult将接收hKey指定的键的新句柄。否则，phkResult将接收传递给函数的相同的hKey句柄。
'    有关更多信息，请参见注册表元素大小限制。
'uloptions
'    指定打开键时应用的选项。将此参数设置为零或如下:
Private Const REG_OPTION_OPEN_LINK      As Long = &H8
'    键是一个符号链接。注册表符号链接只应在绝对必要时使用。
'samDesired
'    指定要打开的键的所需访问权限的掩码。如果键的安全描述符不允许调用进程的请求访问，则该函数将失败。有关更多信息，请参见注册表键安全性和访问权限。
'phkResult
'    指向一个变量的指针，该变量接收打开的键的句柄。如果键不是预定义的注册表键之一，则在使用句柄之后调用RegCloseKey函数。
'@返回值
'    如果函数成功，返回值为ERROR_SUCCESS。
'    如果函数失败，返回值是在Winerror.h中定义的非零错误代码。您可以使用带有FORMAT_MESSAGE_FROM_SYSTEM标志的FormatMessage函数来获得错误的一般描述。
'@备注
'    与RegCreateKeyEx函数不同，如果注册表中没有指定的键，RegOpenKeyEx函数不会创建指定的键。
'    某些注册表操作根据键的安全描述符执行访问检查，而不是在获取键句柄时指定的访问掩码。例如，即使使用samDesired of KEY_READ打开键，如果键的安全描述符允许，也可以使用它创建注册表键。相反，RegSetValueEx函数特别要求使用KEY_SET_VALUE访问权限打开键。
'    如果您的服务或应用程序模拟不同的用户，则不要将此函数与HKEY_CURRENT_USER一起使用。相反，调用RegOpenCurrentUser函数。
'    注意，访问某些注册表项的操作将被重定向。有关更多信息，请参见注册表虚拟化和注册表中的32位和64位应用程序数据。
'@备注
'    与RegCreateKeyEx函数不同，如果注册表中没有指定的键，RegOpenKeyEx函数不会创建指定的键。
'    某些注册表操作根据键的安全描述符执行访问检查，而不是在获取键句柄时指定的访问掩码。例如，即使使用samDesired of KEY_READ打开键，如果键的安全描述符允许，也可以使用它创建注册表键。相反，RegSetValueEx函数特别要求使用KEY_SET_VALUE访问权限打开键。
'    如果您的服务或应用程序模拟不同的用户，则不要将此函数与HKEY_CURRENT_USER一起使用。相反，调用RegOpenCurrentUser函数。
'    注意，访问某些注册表项的操作将被重定向。有关更多信息，请参见注册表虚拟化和注册表中的32位和64位应用程序数据。
'@Requirements
'    Minimum supported client    Windows 2000 Professional [desktop apps only]
'    Minimum supported server    Windows 2000 Server [desktop apps only]
'    Target Platform Windows
'    Header  winreg.h (include Windows.h)
'    Library advapi32.lib
'    dll advapi32.dll
Private Declare Function RegQueryInfoKey Lib "advapi32.dll" Alias "RegQueryInfoKeyW" (ByVal hKey As Long, ByVal lpClass As Long, lpcbClass As Long, lpReserved As Long, lpcSubKeys As Long, lpcbMaxSubKeyLen As Long, lpcbMaxClassLen As Long, lpcValues As Long, lpcbMaxValueNameLen As Long, lpcbMaxValueLen As Long, lpcbSecurityDescriptor As Long, lpftLastWriteTime As FILETIME) As Long
'@功能
'    检索有关指定注册表项的信息。
'@原型
'    LSTATUS RegQueryInfoKeyA(
'      HKEY      hKey,
'      LPSTR     lpClass,
'      LPDWORD   lpcchClass,
'      LPDWORD   lpReserved,
'      LPDWORD   lpcSubKeys,
'      LPDWORD   lpcbMaxSubKeyLen,
'      LPDWORD   lpcbMaxClassLen,
'      LPDWORD   lpcValues,
'      LPDWORD   lpcbMaxValueNameLen,
'      LPDWORD   lpcbMaxValueLen,
'      LPDWORD   lpcbSecurityDescriptor,
'      PFILETIME lpftLastWriteTime
'    );
'@参数
'    hKey
'    打开注册表项的句柄。键必须是用KEY_QUERY_VALUE访问权限打开的。有关更多信息，请参见注册表键安全性和访问权限。
'    这个句柄由RegCreateKeyEx。RegCreateKeyTransacted。RegOpenKeyEx或RegOpenKeyTransacted函数返回。它也可以是下列预定义键之一:
'        HKEY_CLASSES_ROOT
'        HKEY_CURRENT_CONFIG
'        HKEY_CURRENT_USER
'        HKEY_LOCAL_MACHINE
'        HKEY_PERFORMANCE_DATA
'        HKEY_USERS
'lpClass
'    指向接收用户定义的键类的缓冲区的指针。这个参数可以是NULL。
'lpcchClass
'    指向一个变量的指针，该变量以字符为单位指定lpClass参数所指向的缓冲区的大小。
'    大小应该包括终止空字符。当函数返回时，这个变量包含存储在缓冲区中的类字符串的大小。返回的计数不包括终止空字符。如果缓冲区不够大，函数返回ERROR_MORE_DATA，变量包含字符串的大小，以字符为单位，不计算终止null字符。
'    如果lpClass为空，lpcClass可以为空。
'    如果lpClass参数是一个有效地址，但是lpcClass参数不是，例如，它是NULL，那么函数返回ERROR_INVALID_PARAMETER。
'lpReserved
'    保留此参数，且必须为空。
'lpClass
'    指向接收用户定义的键类的缓冲区的指针。lpcSubKeys
'    指向一个变量的指针，该变量接收指定键包含的子键数。这个参数可以是NULL。
'lpcbMaxSubKeyLen
'    一个指向一个变量的指针，该变量接收具有最长名称的键的子键的大小，以Unicode字符表示，不包括终止null字符。这个参数可以是NULL。
'lpcbMaxClassLen
'    一个指向变量的指针，该变量接收指定子键类的最长字符串的大小，以Unicode字符表示。返回的计数不包括终止空字符。这个参数可以是NULL。这个参数可以是NULL。
'lpcchClass
'    指向一个变量的指针，该变量以字符为单位指定lpClass参数所指向的缓冲区的大小。
'    大小应该包括终止空字符。当函数返回时，这个变量包含存储在缓冲区中的类字符串的大小。返回的计数不包括终止空字符。如果缓冲区不够大，函数返回ERROR_MORE_DATA，变量包含字符串的大小，以字符为单位，不计算终止null字符。
'    如果lpClass为空，lpcClass可以为空。
'    如果lpClass参数是一个有效地址，但是lpcClass参数不是，例如，它是NULL，那么函数返回ERROR_INVALID_PARAMETER。
'lpReserved
'    保留此参数，且必须为空。
'lpcValues
'    指向一个变量的指针，该变量接收与键关联的值的数量。这个参数可以是NULL。
'lpcbMaxValueNameLen
'    指向一个变量的指针，该变量接受键的最长值名称的大小(以Unicode字符表示)。大小不包括终止空字符。这个参数可以是NULL。
'lpcbMaxValueLen
'    指向一个变量的指针，该变量接收键值中最长数据组件的大小，单位为字节。这个参数可以是NULL。
'lpcbSecurityDescriptor
'    指向一个变量的指针，该变量接收键安全描述符的大小(以字节为单位)。这个参数可以是NULL。
'lpftLastWriteTime
'    一个指向FILETIME结构的指针，该结构接收最后一次写入时间。这个参数可以是NULL。
'    该函数设置FILETIME结构的成员，以指示键或其任何值条目最后一次被修改的时间。
'@返回值
'    如果函数成功，返回值为ERROR_SUCCESS。
'    如果函数失败，返回值是一个系统错误代码。
'    如果lpClass缓冲区太小，无法接收类名，则函数返回ERROR_MORE_DATA。
'@Requirements
'    Minimum supported client    Windows 2000 Professional [desktop apps only]
'    Minimum supported server    Windows 2000 Server [desktop apps only]
'    Target Platform Windows
'    Header  winreg.h (include Windows.h)
'    Library advapi32.lib
'    dll advapi32.dll
Private Declare Function RegQueryValue Lib "advapi32.dll" Alias "RegQueryValueW" (ByVal hKey As Long, ByVal lpSubKey As Long, ByVal lpValue As String, lpcbValue As Long) As Long
'@功能
'    检索与指定注册表项的默认值或未命名值关联的数据。数据必须是一个以null结尾的字符串。
'    注意，此函数仅用于兼容16位版本的Windows。应用程序应该使用RegQueryValueEx函数。
'@原型
'    LSTATUS RegQueryValueA(
'      HKEY   hKey,
'      LPCWSTR lpSubKey,
'      LPSTR  lpData,
'      PLONG lpcbData
'    );
'@参数
'hKey
'    打开注册表项的句柄。键必须是用KEY_QUERY_VALUE访问权限打开的。有关更多信息，请参见注册表键安全性和访问权限。
'    这个句柄由RegCreateKeyEx。RegCreateKeyTransacted。RegOpenKeyEx或RegOpenKeyTransacted函数返回。它也可以是下列预定义键之一:
'        HKEY_CLASSES_ROOT
'        HKEY_CURRENT_CONFIG
'        HKEY_CURRENT_USER
'        HKEY_LOCAL_MACHINE
'        HKEY_USERS
'lpSubKey
'    检索默认值的hKey参数的子键的名称。
'    键名不区分大小写。
'    如果该参数为NULL或指向空字符串，则函数检索hKey标识的键的默认值。
'    有关更多信息，请参见注册表元素大小限制。
'lpData
'    指向接收指定键的默认值的缓冲区的指针。
'    如果lpValue为NULL, lpcbValue为非NULL，则函数返回ERROR_SUCCESS，并在lpcbValue指向的变量中存储数据的大小(以字节为单位)。这使应用程序能够确定为值的数据分配缓冲区的最佳方法。
'lpcbData
'    指向一个变量的指针，该变量指定lpValue参数指向的缓冲区的大小，单位为字节。当函数返回时，这个变量包含复制到lpValue的数据的大小，包括任何终止null字符。
'    如果数据具有REG_SZ。REG_MULTI_SZ或REG_EXPAND_SZ类型，则此大小包括任何一个或多个终止null字符。有关更多信息，请参见备注。
'    如果指定的缓冲区lpValue不够大，无法容纳数据，则函数返回ERROR_MORE_DATA，并在lpcbValue指向的变量中存储所需的缓冲区大小。在本例中，lpValue缓冲区的内容是未定义的。
'@返回值
'    如果函数成功，返回值为ERROR_SUCCESS。
'    如果函数失败，返回值是一个系统错误代码。
'    如果lpValue缓冲区太小，无法接收该值，则函数返回ERROR_MORE_DATA。
'@备注
'    如果使用该函数的ANSI版本(通过显式调用RegQueryValueA或在包含Windows.h文件之前不定义UNICODE)，则该函数在将存储的UNICODE字符串复制到lpValue参数指定的缓冲区之前将其转换为ANSI字符串。
'    如果数据具有REG_SZ。REG_MULTI_SZ或REG_EXPAND_SZ类型，则字符串可能没有存储正确的空终止字符。因此，即使函数返回ERROR_SUCCESS，应用程序也应该确保字符串在使用之前已正确终止;否则，它可能会覆盖缓冲区。(注意REG_MULTI_SZ字符串应该有两个空终止字符。)
'    注意，访问某些注册表项的操作将被重定向。有关更多信息，请参见注册表虚拟化和注册表中的32位和64位应用程序数据。
'@Requirements
'    Minimum supported client    Windows 2000 Professional [desktop apps only]
'    Minimum supported server    Windows 2000 Server [desktop apps only]
'    Target Platform Windows
'    Header  winreg.h (include Windows.h)
'    Library advapi32.lib
'    dll advapi32.dll
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExW" (ByVal hKey As Long, ByVal lpValueName As Long, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
'@功能
'    检索与开放注册表项关联的指定值名称的类型和数据。
'    要确保返回的任何字符串值(REG_SZ。REG_MULTI_SZ和REG_EXPAND_SZ)以null结尾，请使用RegGetValue函数。
'@原型
'    LSTATUS RegQueryValueExA(
'      HKEY    hKey,
'      LPCWSTR  lpValueName,
'      LPDWORD lpReserved,
'      LPDWORD lpType,
'      LPBYTE  lpData,
'      LPDWORD lpcbData
'    );
'@参数
'hKey
'    打开注册表项的句柄。键必须是用KEY_QUERY_VALUE访问权限打开的。有关更多信息，请参见注册表键安全性和访问权限。
'    这个句柄由RegCreateKeyEx。RegCreateKeyTransacted。RegOpenKeyEx或RegOpenKeyTransacted函数返回。它也可以是下列预定义键之一:
'        HKEY_CLASSES_ROOT
'        HKEY_CURRENT_CONFIG
'        HKEY_CURRENT_USER
'        HKEY_LOCAL_MACHINE
'        HKEY_PERFORMANCE_DATA
'        HKEY_PERFORMANCE_NLSTEXT
'        HKEY_PERFORMANCE_TEXT
'        HKEY_USERS
'lpValueName
'    注册表值的名称。
'    如果lpValueName为NULL或空字符串，“”，函数将检索键的未命名值或默认值(如果有)的类型和数据。
'    如果lpValueName指定了一个不在注册表中的值，该函数将返回ERROR_FILE_NOT_FOUND。
'    键不会自动具有未命名或默认值。未命名值可以是任何类型。有关更多信息，请参见注册表元素大小限制。
'lpReserved
'    保留此参数，且必须为空。
'lpType
'    指向一个变量的指针，该变量接收一个代码，该代码指示存储在指定值中的数据类型。有关可能的类型代码列表，请参见注册表值类型。如果不需要类型代码，则lpType参数可以为NULL。
'lpData
'    指向接收值数据的缓冲区的指针。如果不需要数据，此参数可以为NULL。
'lpcbData
'    指向一个变量的指针，该变量指定lpData参数指向的缓冲区的大小，单位为字节。当函数返回时，这个变量包含复制到lpData的数据的大小。
'    只有当lpData为空时，lpcbData参数才能为空。
'    如果数据具有REG_SZ。REG_MULTI_SZ或REG_EXPAND_SZ类型，则此大小包括任何终止null字符，除非数据存储时没有这些字符。有关更多信息，请参见备注。
'    如果lpData参数指定的缓冲区不够大，无法保存数据，则函数返回ERROR_MORE_DATA，并在lpcbData指向的变量中存储所需的缓冲区大小。在本例中，lpData缓冲区的内容是未定义的。
'    如果lpData为NULL, lpcbData为非NULL，则函数返回ERROR_SUCCESS，并在lpcbData指向的变量中存储数据的大小(以字节为单位)。这使应用程序能够确定为值的数据分配缓冲区的最佳方法。
'    如果hKey指定HKEY_PERFORMANCE_DATA，并且lpData缓冲区不够大，无法包含所有返回的数据，那么RegQueryValueEx将返回ERROR_MORE_DATA，并且通过lpcbData参数返回的值是未定义的。这是因为性能数据的大小可以从一个调用更改到下一个调用。在这种情况下，必须增加缓冲区大小并再次调用RegQueryValueEx，在lpcbData参数中传递更新后的缓冲区大小。重复此操作，直到函数成功。您需要维护一个单独的变量来跟踪缓冲区大小，因为lpcbData返回的值是不可预测的。
'    如果lpValueName注册表值不存在，RegQueryValueEx将返回ERROR_FILE_NOT_FOUND，并且通过lpcbData参数返回的值是未定义的。
'@返回值
'    如果函数成功，返回值为ERROR_SUCCESS。
'    如果函数失败，返回值是一个系统错误代码。
'    如果lpData缓冲区太小而无法接收数据，则函数返回ERROR_MORE_DATA。
'    如果lpValueName注册表值不存在，则函数返回ERROR_FILE_NOT_FOUND。
'@备注
'    应用程序通常调用RegEnumValue来确定值名称，然后调用RegQueryValueEx来检索名称的数据。
'    如果数据具有REG_SZ。REG_MULTI_SZ或REG_EXPAND_SZ类型，则字符串可能没有存储正确的终止null字符。因此，即使函数返回ERROR_SUCCESS，应用程序也应该确保字符串在使用之前已正确终止;否则，它可能会覆盖缓冲区。(注意REG_MULTI_SZ字符串应该有两个终止null字符。)应用程序确保正确终止字符串的一种方法是使用RegGetValue，如果需要，它将添加终止null字符。
'    如果数据有REG_SZ , REG_MULTI_SZ或REG_EXPAND_SZ类型, 并使用这个函数的ANSI版本(通过显式地调用RegQueryValueExA或不包括Windows.h文件之前定义UNICODE), 这个函数将存储UNICODE字符串转换为一个ANSI字符串复制到缓冲区之前由lpData指出。
'    当使用hKey设置为HKEY_PERFORMANCE_DATA句柄和指定对象的值字符串调用RegQueryValueEx函数时，返回的数据结构有时具有未请求的对象。不要惊讶;这是正常的行为。在调用RegQueryValueEx函数时，您应该始终期望遍历返回的数据结构来查找请求的对象。
'    注意，访问某些注册表项的操作将被重定向。有关更多信息，请参见注册表虚拟化和注册表中的32位和64位应用程序数据。
'@Requirements
'    Minimum supported client    Windows 2000 Professional [desktop apps only]
'    Minimum supported server    Windows 2000 Server [desktop apps only]
'    Target Platform Windows
'    Header  winreg.h (include Windows.h)
'    Library advapi32.lib
'    dll advapi32.dll
Private Declare Function RegSetKeyValue Lib "advapi32.dll" Alias "RegSetKeyValueW" (ByVal hKey As Long, ByVal lpValueName As Long, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
'@功能
'    设置指定注册表项和子项中指定值的数据。
'@原型
'    LSTATUS RegSetKeyValueA(
'      HKEY    hKey,
'      LPCWSTR  lpSubKey,
'      LPCWSTR  lpValueName,
'      DWORD   dwType,
'      LPCVOID lpData,
'      DWORD cbData
'    );
'@参数
'hKey
'    打开注册表项的句柄。键必须是用KEY_SET_VALUE访问权限打开的。有关更多信息，请参见注册表键安全性和访问权限。
'    这个句柄由RegCreateKeyEx。RegCreateKeyTransacted。RegOpenKeyEx或RegOpenKeyTransacted函数返回。它也可以是下列预定义键之一:
'        HKEY_CLASSES_ROOT
'        HKEY_CURRENT_CONFIG
'        HKEY_CURRENT_USER
'        HKEY_LOCAL_MACHINE
'        HKEY_USERS
'lpSubKey
'    键的名称和由hKey标识的键的子键。如果该参数为NULL，则使用hKey值在键中创建该值，键将获得默认的安全描述符。
'lpValueName
'    要更新其数据的注册表值的名称。
'dwType
'    lpData参数指向的数据类型。有关可能的类型列表，请参见注册表值类型。
'lpData
'    要用指定值名称存储的数据。
'    对于基于字符串的类型，如REG_SZ，字符串必须以null结尾。对于REG_MULTI_SZ数据类型，字符串必须以两个空字符结束。
'cbData
'    lpData参数指向的信息的大小，以字节为单位。如果数据类型为REG_SZ。REG_EXPAND_SZ或REG_MULTI_SZ，则cbData必须包含终止null字符或多个字符的大小。
'@返回值
'    如果函数成功，返回值为ERROR_SUCCESS。
'    如果函数失败，返回值是在Winerror.h中定义的非零错误代码。您可以使用带有FORMAT_MESSAGE_FROM_SYSTEM标志的FormatMessage函数来获得错误的一般描述。
'@Requirements
'    Minimum supported client    Windows Vista [desktop apps only]
'    Minimum supported server    Windows Server 2008 [desktop apps only]
'    Target Platform Windows
'    Header  winreg.h (include Windows.h)
'    Library advapi32.lib
'    dll advapi32.dll
Private Declare Function RegSetValue Lib "advapi32.dll" Alias "RegSetValueW" (ByVal hKey As Long, ByVal lpSubKey As Long, ByVal dwType As Long, ByVal lpData As Long, ByVal cbData As Long) As Long
'@功能
'    设置指定注册表项的默认值或未命名值的数据。数据必须是文本字符串。
'    注意，此函数仅用于兼容16位版本的Windows。应用程序应该使用RegSetValueEx函数。
'@原型
'    LSTATUS RegSetValueA(
'      HKEY   hKey,
'      LPCWSTR lpSubKey,
'      DWORD  dwType,
'      LPCWSTR lpData,
'      DWORD cbData
'    );
'@参数
'hKey
'    打开注册表项的句柄。键必须是用KEY_SET_VALUE访问权限打开的。有关更多信息，请参见注册表键安全性和访问权限。
'    这个句柄由RegCreateKeyEx。RegCreateKeyTransacted。RegOpenKeyEx或RegOpenKeyTransacted函数返回。它也可以是下列预定义键之一:
'        HKEY_CLASSES_ROOT
'        HKEY_CURRENT_CONFIG
'        HKEY_CURRENT_USER
'        HKEY_LOCAL_MACHINE
'        HKEY_USERS
'lpSubKey
'    hKey参数的子键的名称。函数设置指定子键的默认值。如果lpSubKey不存在，函数将创建它。
'    键名不区分大小写。
'    如果该参数为NULL或指向空字符串，则函数将设置hKey标识的键的默认值。
'    有关更多信息，请参见注册表元素大小限制。
'dwType
'    要存储的信息的类型。该参数必须是REG_SZ类型。要存储其他数据类型，请使用RegSetValueEx函数。
'lpData
'    要存储的数据。此参数不能为空。
'cbData
'    忽略此参数。该函数根据lpData参数中的数据大小计算这个值。
'@返回值
'    如果函数成功，返回值为ERROR_SUCCESS。
'    如果函数失败，返回值是在Winerror.h中定义的非零错误代码。您可以使用带有FORMAT_MESSAGE_FROM_SYSTEM标志的FormatMessage函数来获得错误的一般描述。
'@备注
'    如果lpSubKey参数指定的键不存在，RegSetValue函数将创建它。
'    如果使用这个函数的ANSI版本(通过显式调用RegSetValueA或在包含Windows.h文件之前不定义UNICODE)， lpData参数必须是一个ANSI字符串。字符串在存储到注册表之前被转换为Unicode。
'@Requirements
'    Minimum supported client    Windows 2000 Professional [desktop apps only]
'    Minimum supported server    Windows 2000 Server [desktop apps only]
'    Target Platform Windows
'    Header  winreg.h (include Windows.h)
'    Library advapi32.lib
'    dll advapi32.dll

Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExW" (ByVal hKey As Long, ByVal lpValueName As Long, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
'@功能
'    设置注册表项下指定值的数据和类型。
'@原型
'    LSTATUS RegSetValueExA(
'      HKEY       hKey,
'      LPCWSTR     lpValueName,
'      DWORD      Reserved,
'      DWORD      dwType,
'      const BYTE *lpData,
'      DWORD cbData
'    );
'@参数
'hKey
'    打开注册表项的句柄。键必须是用KEY_SET_VALUE访问权限打开的。有关更多信息，请参见注册表键安全性和访问权限。
'    这个句柄由RegCreateKeyEx。RegCreateKeyTransacted。RegOpenKeyEx或RegOpenKeyTransacted函数返回。它也可以是下列预定义键之一:
'        HKEY_CLASSES_ROOT
'        HKEY_CURRENT_CONFIG
'        HKEY_CURRENT_USER
'        HKEY_LOCAL_MACHINE
'        HKEY_USERS
'        这个函数的Unicode版本支持以下额外的预定义键:
'        HKEY_PERFORMANCE_TEXT
'        HKEY_PERFORMANCE_NLSTEXT
'lpValueName
'    要设置的值的名称。如果键中还没有具有该名称的值，则函数将其添加到键中。
'    如果lpValueName为NULL或空字符串“”，函数将为键的未命名值或默认值设置类型和数据。
'    有关更多信息，请参见注册表元素大小限制。
'    注册表项没有默认值，但是它们可以有一个未命名的值，可以是任何类型的值。
'Reserved
'    这个参数是保留的，必须为零。
'dwType
'    lpData参数指向的数据类型。有关可能的类型列表，请参见注册表值类型。
'lpData
'    要存储的数据。
'    对于基于字符串的类型，如REG_SZ，字符串必须以null结尾。对于REG_MULTI_SZ数据类型，字符串必须以两个空字符结束。
'    注意，lpData表示null值是有效的，但是，如果是这种情况，则必须将cbData设置为'0'。
'cbData
'    lpData参数指向的信息的大小，以字节为单位。如果数据类型为REG_SZ。REG_EXPAND_SZ或REG_MULTI_SZ，则cbData必须包含终止null字符或多个字符的大小。
'@返回值
'    如果函数成功，返回值为ERROR_SUCCESS。
'    如果函数失败，返回值是在Winerror.h中定义的非零错误代码。您可以使用带有FORMAT_MESSAGE_FROM_SYSTEM标志的FormatMessage函数来获得错误的一般描述。
'@备注
'    值大小受可用内存的限制。然而，在注册表中存储较大的值会影响其性能。长值(超过2048字节)应该存储为文件，文件的位置存储在注册表中。
'    应用程序元素(如图标。位图和可执行文件)应该存储为文件，而不是放在注册表中。
'    如果dwType是REG_SZ。REG_MULTI_SZ或REG_EXPAND_SZ类型，并且使用该函数的ANSI版本(通过显式调用RegSetValueExA或在包含Windows.h文件之前不定义UNICODE)， lpData参数指向的数据必须是ANSI字符串。字符串在存储到注册表之前被转换为Unicode。
'    注意，访问某些注册表项的操作将被重定向。有关更多信息，请参见注册表虚拟化和注册表中的32位和64位应用程序数据。
'@Requirements
'    Minimum supported client    Windows 2000 Professional [desktop apps only]
'    Minimum supported server    Windows 2000 Server [desktop apps only]
'    Target Platform Windows
'    Header  winreg.h (include Windows.h)
'    Library advapi32.lib
'    dll advapi32.dll
Private Declare Sub RtlMoveMemory Lib "kernel32" (Destination As Any, Source As Any, ByVal Length As Long)
'说明：将内存块从一个位置移动到另一个位置
'Destination:指向移动目的地起始地址的指针。
'Source:指向要移动的内存块起始地址的指针。
'Length:内存块的大小以字节为单位移动。
'注意事项：这个函数定义为RtlMoveMemory函数。它的实现是内联的。有关更多信息，请参见WinBase。h和Winnt.h。源和目标块可能会重叠。
'           第一个参数，目的地，必须足够大，以容纳长度字节的源;否则，可能会出现缓冲区溢出。这可能导致拒绝服务攻击，如果有访问违反，或者在最坏的情况下，允许攻击者向您的进程注入可执行代码。如果目的地是一个基于堆栈的缓冲区，则尤其如此。要注意，最后一个参数，长度，是将字节复制到目的地的数量，而不是目的地的大小。

'---------------------------------------------------------------------------
'                1。常规变量
'---------------------------------------------------------------------------
Private mblnImpersonateLoggedOnUserOn   As Boolean      '是否开启用户模拟
Private mblnManualClose                 As Boolean      '是否手动关闭用户模拟


'---------------------------------------------------------------------------
'                2。属性变量与定义
'---------------------------------------------------------------------------

'---------------------------------------------------------------------------
'                3。公共方法
'---------------------------------------------------------------------------

'@方法    RegImpersonateLoggedOnUserOn
'   模拟调用线程的用户注册信息，主要是为了读取调用线程模拟用户的HKEY_CURRENT_USER注册表
'@返回值
'
'@参数:
'blnManualClose Boolean In
'   是否关闭手动模拟用户,非手动关闭，使用一次函数后就自动关闭
'@备注
'   非当前会话的操作系统用户，在读取注册表时需要调用ImpersonateLoggedOnUser来模拟指定的用户，或者直接读取HKEY_USERS
Public Sub RegImpersonateLoggedOnUserOn(ByVal blnManualClose As Boolean)
    mblnImpersonateLoggedOnUserOn = True
    mblnManualClose = blnManualClose
End Sub
'@方法    RegImpersonateLoggedOnUserOff
'   关闭用户模拟
'@返回值
'
'@参数:
'@备注
Public Sub RegImpersonateLoggedOnUserOff()
    mblnImpersonateLoggedOnUserOn = False
    mblnManualClose = False
End Sub

'@方法    CreateRegKeyValue
'   创建一个键并添加值
'@返回值  Boolean
'
'@参数:
'strKey String In
'   注册表键位，如“HKEY_CURRENT_USER\Printers\DevModePerUser"
'strValueName String In(Optional)
'   键的值名称
'varValue   Variant In
'   键的值。若该值不传，且strValueName="",则不添加键值。该数据必须传入正确的类型。若值不存在，则自动根据传入的类型自动转换。
'       String->REG_SZ,Long\Integer\Byte->REG_DWORD
'       String()->REG_MULTI_SZ
'       Byte()->REG_BINARY
'@备注
'
Public Function CreateRegKeyValue(ByVal strKey As String, Optional ByVal strValueName As String, Optional ByRef varValue As Variant) As Boolean
    Dim strSubKey       As String, hRootKey         As Long, blnMustCloseRootKey    As Boolean
    Dim hKey            As Long, lngDisposition     As Long, lngKeyValuType         As Long
    Dim lngError        As Long
    Dim lngsamDesired   As Long
    Dim saCur           As SECURITY_ATTRIBUTES
    Dim arrBytBuffer()  As Byte
    Dim blnOk           As Boolean
    
    On Error GoTo ErrH
    Call Logger.PushMethod("ZLHelperMain.clsRegistry.CreateRegKeyValue", strKey, strValueName, varValue)
    lngsamDesired = KEY_CREATE_SUB_KEY Or KEY_SET_VALUE Or KEY_QUERY_VALUE
    If GetRegKeyInfo(strKey, hRootKey, strSubKey, blnMustCloseRootKey, lngsamDesired) Then
        blnOk = True
        saCur.nLength = Len(saCur)
        saCur.lpSecurityDescriptor = 0
        saCur.bInheritHandle = True
        lngError = RegCreateKeyEx(hRootKey, StrPtr(strSubKey), 0, 0, 0, lngsamDesired, saCur, hKey, lngDisposition)
        If lngError <> ERROR_SUCCESS Then
            blnOk = False
            Logger.Error "RegCreateKeyEx", "Error", GetLastDllErr(lngError)
        Else
            '查询数据类型
            If lngDisposition = REG_OPENED_EXISTING_KEY Then
                Logger.DebugEx "RegCreateKeyEx", "EXISTING_KEY", True
                lngError = RegQueryValueEx(hKey, StrPtr(strValueName), 0, lngKeyValuType, ByVal 0, ByVal 0)
                If lngError <> ERROR_SUCCESS Then
                    If lngError <> ERROR_FILE_NOT_FOUND Then
                        blnOk = False
                        Logger.Error "RegQueryValueEx", "Error", GetLastDllErr(lngError)
                    Else
                        Logger.DebugEx "RegCreateKeyEx", "ValueNotExists", True
                    End If
                End If
            Else
                Logger.DebugEx "RegCreateKeyEx", "NEW_KEY", True
            End If
            If blnOk Then
                If Not (IsMissing(varValue) And strValueName = "") Then '添加一个值
                    '先转换，可能数据类型不支持，则返回False
                    If ConvertDataToByte(varValue, arrBytBuffer, lngKeyValuType) Then
                        If UboundEx(arrBytBuffer) > -1 Then
                            lngError = RegSetValueEx(hKey, StrPtr(strValueName), 0, lngKeyValuType, arrBytBuffer(0), UboundEx(arrBytBuffer) + 1)
                        Else
                            lngError = RegSetValueEx(hKey, StrPtr(strValueName), 0, lngKeyValuType, ByVal 0, 0)
                        End If
                        If lngError <> ERROR_SUCCESS Then
                            blnOk = False
                            Logger.Error "RegSetValueEx", "Error", GetLastDllErr(lngError)
                        End If
                    Else
                        blnOk = False
                    End If
                End If
            End If
            lngError = RegCloseKey(hKey)
            If lngError <> ERROR_SUCCESS Then
                blnOk = False
                Logger.Error "RegCloseKey", "Error", GetLastDllErr(lngError)
            Else
                hKey = 0
            End If
        End If
        'RootKey是线程模拟用户得到的，需要关闭
        If blnMustCloseRootKey Then
            lngError = RegCloseKey(hRootKey)
            If lngError <> ERROR_SUCCESS Then
                blnOk = False
                Logger.Error "RegCloseKey(ImpersonateLoggedOnUser)", "Error", GetLastDllErr(lngError)
            End If
        End If
    End If
    CreateRegKeyValue = blnOk
    Call Logger.PopMethod("ZLHelperMain.clsRegistry.CreateRegKeyValue", CreateRegKeyValue)
    Exit Function
ErrH:
    If Logger.ErrCenter("ZLHelperMain.clsRegistry.CreateRegKeyValue") = 1 Then
        Resume
    End If
    If hKey <> 0 Then
        lngError = RegCloseKey(hKey)
        If lngError <> ERROR_SUCCESS Then
            Logger.Error "RegCloseKey", "Error", GetLastDllErr(lngError)
        End If
    End If
    If blnMustCloseRootKey And hRootKey <> 0 Then
        lngError = RegCloseKey(hRootKey)
        If lngError <> ERROR_SUCCESS Then
            Logger.Error "RegCloseKey(ImpersonateLoggedOnUser)", "Error", GetLastDllErr(lngError)
        End If
    End If
    Call Logger.PopMethod("ZLHelperMain.clsRegistry.CreateRegKeyValue")
End Function



'@方法    SetRegValue
'   创建一个键并添加值
'@返回值  Boolean
'
'@参数:
'strKey String In
'   注册表键位，如“HKEY_CURRENT_USER\Printers\DevModePerUser"
'strValueName String In(Optional)
'   键的值名称
'varValue   Variant In
'   键的值。若该值不传，且strValueName="",则不添加键值。该数据必须传入正确的类型。若值不存在，则自动根据传入的类型自动转换。
'       String->REG_SZ,Long\Integer\Byte->REG_DWORD
'       String()->REG_MULTI_SZ
'       Byte()->REG_BINARY
'@备注
'
Public Function SetRegValue(ByVal strKey As String, Optional ByVal strValueName As String, Optional ByRef varValue As Variant, Optional blnAutoCreateKey As Boolean = True) As Boolean
    Dim strSubKey       As String, hRootKey         As Long, blnMustCloseRootKey    As Boolean
    Dim hKey            As Long, lngDisposition     As Long, lngKeyValuType         As Long
    Dim lngError        As Long
    Dim lngsamDesired   As Long
    Dim saCur           As SECURITY_ATTRIBUTES
    Dim arrBytBuffer()  As Byte
    Dim blnOk           As Boolean
    
    On Error GoTo ErrH
    Call Logger.PushMethod("ZLHelperMain.clsRegistry.SetRegValue", strKey, strValueName, varValue)
    If blnAutoCreateKey Then
        lngsamDesired = KEY_CREATE_SUB_KEY Or KEY_SET_VALUE Or KEY_QUERY_VALUE
    Else
        lngsamDesired = KEY_SET_VALUE Or KEY_QUERY_VALUE
    End If
    If GetRegKeyInfo(strKey, hRootKey, strSubKey, blnMustCloseRootKey, lngsamDesired) Then
        blnOk = True
        If blnAutoCreateKey Then
            saCur.nLength = Len(saCur)
            saCur.lpSecurityDescriptor = 0
            saCur.bInheritHandle = True
            lngError = RegCreateKeyEx(hRootKey, StrPtr(strSubKey), 0, 0, 0, lngsamDesired, saCur, hKey, lngDisposition)
        Else
            lngError = RegOpenKeyEx(hRootKey, StrPtr(strSubKey), 0, lngsamDesired, hKey)
            lngDisposition = REG_OPENED_EXISTING_KEY
        End If
        If lngError <> ERROR_SUCCESS Then
            blnOk = False
            If blnAutoCreateKey Then
                Logger.Error "RegCreateKeyEx", "Error", GetLastDllErr(lngError)
            Else
                Logger.Error "RegOpenKeyEx", "Error", GetLastDllErr(lngError)
            End If
        Else
            '查询数据类型
            If lngDisposition = REG_OPENED_EXISTING_KEY Then
                Logger.DebugEx "RegCreateKeyEx", "EXISTING_KEY", True
                lngError = RegQueryValueEx(hKey, StrPtr(strValueName), 0, lngKeyValuType, ByVal 0, ByVal 0)
                If lngError <> ERROR_SUCCESS Then
                    If lngError <> ERROR_FILE_NOT_FOUND Then
                        blnOk = False
                        Logger.Error "RegQueryValueEx", "Error", GetLastDllErr(lngError)
                    End If
                End If
            Else
                Logger.DebugEx "RegCreateKeyEx", "NEW_KEY", True
            End If
            If blnOk Then
                If Not (IsMissing(varValue) And strValueName = "") Then '添加一个值
                    '先转换，可能数据类型不支持，则返回False
                    If ConvertDataToByte(varValue, arrBytBuffer, lngKeyValuType) Then
                        If UboundEx(arrBytBuffer) > -1 Then
                            lngError = RegSetValueEx(hKey, StrPtr(strValueName), 0, lngKeyValuType, arrBytBuffer(0), UboundEx(arrBytBuffer) + 1)
                        Else
                            lngError = RegSetValueEx(hKey, StrPtr(strValueName), 0, lngKeyValuType, ByVal 0, 0)
                        End If
                        If lngError <> ERROR_SUCCESS Then
                            blnOk = False
                            Logger.Error "RegSetValueEx", "Error", GetLastDllErr(lngError)
                        End If
                    Else
                        blnOk = False
                    End If
                End If
            End If
            lngError = RegCloseKey(hKey)
            If lngError <> ERROR_SUCCESS Then
                blnOk = False
                Logger.Error "RegCloseKey", "Error", GetLastDllErr(lngError)
            Else
                hKey = 0
            End If
        End If
        'RootKey是线程模拟用户得到的，需要关闭
        If blnMustCloseRootKey Then
            lngError = RegCloseKey(hRootKey)
            If lngError <> ERROR_SUCCESS Then
                blnOk = False
                Logger.Error "RegCloseKey(ImpersonateLoggedOnUser)", "Error", GetLastDllErr(lngError)
            End If
        End If
    End If
    SetRegValue = blnOk
    Call Logger.PopMethod("ZLHelperMain.clsRegistry.SetRegValue", SetRegValue)
    Exit Function
ErrH:
    If Logger.ErrCenter("ZLHelperMain.clsRegistry.SetRegValue") = 1 Then
        Resume
    End If
    If hKey <> 0 Then
        lngError = RegCloseKey(hKey)
        If lngError <> ERROR_SUCCESS Then
            Logger.Error "RegCloseKey", "Error", GetLastDllErr(lngError)
        End If
    End If
    
    If blnMustCloseRootKey And hRootKey <> 0 Then
        lngError = RegCloseKey(hRootKey)
        If lngError <> ERROR_SUCCESS Then
            Logger.Error "RegCloseKey(ImpersonateLoggedOnUser)", "Error", GetLastDllErr(lngError)
        End If
    End If
    Call Logger.PopMethod("ZLHelperMain.clsRegistry.SetRegValue")
End Function

'@方法    GetRegValue
'   获取注册表键值
'@返回值  Variant
'
'@参数:
'strKey String In
'   注册表键位，如“HKEY_CURRENT_USER\Printers\DevModePerUser"
'strValueName String In
'   键的值名称
'varDefualtValue   Variant In(Optional)
'   读取不到或者读取错误失败时的缺省值
'blnSucess Boolean Out(Optional)
'   是否读取出错，当键值不存在，不算错误。
'@备注
'
Public Function GetRegValue(ByVal strKey As String, ByVal strValueName As String, Optional ByRef varDefualtValue As Variant, Optional ByRef blnSucess As Boolean) As Variant
    Dim strSubKey       As String, hRootKey         As Long, blnMustCloseRootKey    As Boolean
    Dim hKey            As Long, lngKeyValuType     As Long
    Dim lngError        As Long
    Dim lngsamDesired   As Long
    Dim arrBytBuffer()  As Byte, lngDataLenB        As Long
    Dim arrRet          As Variant
    Dim blnOk           As Boolean
    On Error GoTo ErrH
    Call Logger.PushMethod("ZLHelperMain.clsRegistry.GetRegValue", strKey, strValueName, varDefualtValue, blnSucess)
    blnSucess = False
    lngsamDesired = KEY_QUERY_VALUE
    If GetRegKeyInfo(strKey, hRootKey, strSubKey, blnMustCloseRootKey, lngsamDesired) Then
        lngError = RegOpenKeyEx(hRootKey, StrPtr(strSubKey), 0, lngsamDesired, hKey)
        If lngError <> ERROR_SUCCESS Then
            If lngError = ERROR_FILE_NOT_FOUND Then
                Logger.Warn "RegOpenKeyEx", "Error", GetLastDllErr(lngError)
            Else
                Logger.Error "RegOpenKeyEx", "Error", GetLastDllErr(lngError)
            End If
        Else
            lngError = RegQueryValueEx(hKey, StrPtr(strValueName), 0, lngKeyValuType, ByVal 0, lngDataLenB)
            If lngError <> ERROR_SUCCESS Then
                If lngError = ERROR_FILE_NOT_FOUND Then
                    blnSucess = True
                    blnOk = True
                    GetRegValue = varDefualtValue
                Else
                    Logger.Error "RegQueryValueEx(QueryLength)", "Error", GetLastDllErr(lngError)
                End If
            Else
                ReDim Preserve arrBytBuffer(lngDataLenB - 1)
                lngError = RegQueryValueEx(hKey, StrPtr(strValueName), 0, lngKeyValuType, ByVal VarPtr(arrBytBuffer(0)), lngDataLenB)
                If lngError <> ERROR_SUCCESS Then
                    Logger.Error "RegQueryValueEx(QueryValue)", "Error", GetLastDllErr(lngError)
                Else
                    If ConvertByteToNormalData(arrBytBuffer, lngKeyValuType, arrRet) Then
                        blnSucess = True
                        blnOk = True
                        GetRegValue = arrRet
                    End If
                End If
            End If
            lngError = RegCloseKey(hKey)
            If lngError <> ERROR_SUCCESS Then
                blnOk = False
                Logger.Error "RegCloseKey", "Error", GetLastDllErr(lngError)
            Else
                hKey = 0
            End If
        End If
        'RootKey是线程模拟用户得到的，需要关闭
        If blnMustCloseRootKey Then
            lngError = RegCloseKey(hRootKey)
            If lngError <> ERROR_SUCCESS Then
                blnOk = False
                Logger.Error "RegCloseKey(ImpersonateLoggedOnUser)", "Error", GetLastDllErr(lngError)
            End If
        End If
    End If
    If Not blnSucess Then
        GetRegValue = varDefualtValue
    Else
        If Not blnOk Then
            blnSucess = False
        End If
    End If

    Call Logger.PopMethod("ZLHelperMain.clsRegistry.GetRegValue", GetRegValue, blnSucess)
    Exit Function
ErrH:
    If Logger.ErrCenter("ZLHelperMain.clsRegistry.GetRegValue") = 1 Then
        Resume
    End If
    If hKey <> 0 Then
        lngError = RegCloseKey(hKey)
        If lngError <> ERROR_SUCCESS Then
            Logger.Error "RegCloseKey", "Error", GetLastDllErr(lngError)
        End If
    End If
    If blnMustCloseRootKey And hRootKey <> 0 Then
        lngError = RegCloseKey(hRootKey)
        If lngError <> ERROR_SUCCESS Then
            Logger.Error "RegCloseKey(ImpersonateLoggedOnUser)", "Error", GetLastDllErr(lngError)
        End If
    End If
    Call Logger.PopMethod("ZLHelperMain.clsRegistry.GetRegValue")
End Function


'@方法    DeleteRegKey
'   删除键值以及其下子健
'@返回值  Boolean
'
'@参数:
'strKey String In
'   注册表键位，如“HKEY_CURRENT_USER\Printers\DevModePerUser"
'blnRecursively Boolean In (Optional,Default=True)
'   是否递归删除其子健。否则只删除当前键，此时若有子健，则会出错，无法删除
'@备注
'
Public Function DeleteRegKey(ByVal strKey As String, Optional ByVal blnRecursively As Boolean = True) As Boolean
    Dim strSubKey       As String, hRootKey         As Long, blnMustCloseRootKey    As Boolean
    Dim lngError        As Long
    Dim lngsamDesired   As Long
    Dim blnSucess       As Boolean
    
    On Error GoTo ErrH
    Call Logger.PushMethod("ZLHelperMain.clsRegistry.DeleteRegKey", strKey, blnRecursively)
    If blnRecursively Then
        lngsamDesired = KEY_SET_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_QUERY_VALUE
    Else
        lngsamDesired = KEY_SET_VALUE Or KEY_QUERY_VALUE
    End If
    If GetRegKeyInfo(strKey, hRootKey, strSubKey, blnMustCloseRootKey, lngsamDesired) Then
        If Not blnRecursively Then
            If Not blnMustCloseRootKey Then
                lngError = RegOpenKeyEx(hRootKey, 0, 0, lngsamDesired, hRootKey)
                If lngError <> ERROR_SUCCESS Then
                    Logger.Error "RegOpenKeyEx", "Error", GetLastDllErr(lngError)
                Else
                    blnMustCloseRootKey = True
                    blnSucess = True
                End If
            End If
            If blnSucess Then
                lngError = RegDeleteKey(hRootKey, StrPtr(strSubKey))
                If lngError <> ERROR_SUCCESS Then
                    blnSucess = False
                    Logger.Error "RegDeleteKey", "Error", GetLastDllErr(lngError)
                End If
            End If
        Else
            If Not blnMustCloseRootKey Then
                lngError = RegOpenKeyEx(hRootKey, 0, 0, lngsamDesired, hRootKey)
                If lngError <> ERROR_SUCCESS Then
                    Logger.Error "RegOpenKeyEx", "Error", GetLastDllErr(lngError)
                Else
                    blnMustCloseRootKey = True
                    blnSucess = True
                End If
            End If
            blnSucess = DeleteRegKey_Sub(hRootKey, strSubKey, lngsamDesired)
        End If
        'RootKey是线程模拟用户得到的，需要关闭
        If blnMustCloseRootKey Then
            lngError = RegCloseKey(hRootKey)
            If lngError <> ERROR_SUCCESS Then
                blnSucess = False
                Logger.Error "RegCloseKey(ImpersonateLoggedOnUser)", "Error", GetLastDllErr(lngError)
            End If
        End If
    End If
    DeleteRegKey = blnSucess
    Call Logger.PopMethod("ZLHelperMain.clsRegistry.DeleteRegKey", DeleteRegKey, blnSucess)
    Exit Function
ErrH:
    If Logger.ErrCenter("ZLHelperMain.clsRegistry.DeleteRegKey") = 1 Then
        Resume
    End If
    If blnMustCloseRootKey And hRootKey <> 0 Then
        lngError = RegCloseKey(hRootKey)
        If lngError <> ERROR_SUCCESS Then
            Logger.Error "RegCloseKey(ImpersonateLoggedOnUser)", "Error", GetLastDllErr(lngError)
        End If
    End If
    Call Logger.PopMethod("ZLHelperMain.clsRegistry.DeleteRegKey")
End Function

'@方法    DeleteRegKeyValue
'   删除一个键的键值
'@返回值  Boolean
'
'@参数:
'strKey String In
'   注册表键位，如“HKEY_CURRENT_USER\Printers\DevModePerUser"
'strValueName   String In
'   键值名
'@备注
'
Public Function DeleteRegKeyValue(ByVal strKey As String, ByVal strValueName As String) As Boolean
    Dim strSubKey       As String, hRootKey         As Long, blnMustCloseRootKey    As Boolean
    Dim hKey            As Long
    Dim lngError        As Long
    Dim lngsamDesired   As Long
    Dim blnSucess       As Boolean
    
    On Error GoTo ErrH
    Call Logger.PushMethod("ZLHelperMain.clsRegistry.DeleteRegKeyValue", strKey, strValueName)
    lngsamDesired = KEY_SET_VALUE
    If GetRegKeyInfo(strKey, hRootKey, strSubKey, blnMustCloseRootKey, lngsamDesired) Then
        lngError = RegOpenKeyEx(hRootKey, StrPtr(strSubKey), 0, lngsamDesired, hKey)
        If lngError <> ERROR_SUCCESS Then
            Logger.Error "RegOpenKeyEx", "Error", GetLastDllErr(lngError)
        Else
            lngError = RegDeleteValue(hKey, StrPtr(strValueName))
            If lngError <> ERROR_SUCCESS Then
                Logger.Error "RegDeleteValue", "Error", GetLastDllErr(lngError)
            Else
                blnSucess = True
            End If
            lngError = RegCloseKey(hKey)
            If lngError <> ERROR_SUCCESS Then
                blnSucess = False
                Logger.Error "RegCloseKey", "Error", GetLastDllErr(lngError)
            End If
        End If
        'RootKey是线程模拟用户得到的，需要关闭
        If blnMustCloseRootKey Then
            lngError = RegCloseKey(hRootKey)
            If lngError <> ERROR_SUCCESS Then
                blnSucess = False
                Logger.Error "RegCloseKey(ImpersonateLoggedOnUser)", "Error", GetLastDllErr(lngError)
            End If
        End If
    End If
    DeleteRegKeyValue = blnSucess
    Call Logger.PopMethod("ZLHelperMain.clsRegistry.DeleteRegKeyValue", DeleteRegKeyValue, blnSucess)
    Exit Function
ErrH:
    If Logger.ErrCenter("ZLHelperMain.clsRegistry.DeleteRegKeyValue") = 1 Then
        Resume
    End If
    If blnMustCloseRootKey And hRootKey <> 0 Then
        lngError = RegCloseKey(hRootKey)
        If lngError <> ERROR_SUCCESS Then
            Logger.Error "RegCloseKey(ImpersonateLoggedOnUser)", "Error", GetLastDllErr(lngError)
        End If
    End If
    Call Logger.PopMethod("ZLHelperMain.clsRegistry.DeleteRegKeyValue")
End Function

'@方法    GetRegSubKeys
'   获取某一个键的直接子健
'@返回值  Variant
'
'@参数:
'strKey String In
'   注册表键位，如“HKEY_CURRENT_USER\Printers\DevModePerUser"
'blnSucess Boolean Out(Optional)
'   是否读取出错，如键不存在或者没有权限等错误
'@备注
'
Public Function GetRegSubKeys(ByVal strKey As String, Optional ByRef blnSucess As Boolean) As Variant
    Dim strSubKey       As String, hRootKey         As Long, blnMustCloseRootKey    As Boolean
    Dim lngError        As Long
    Dim lngsamDesired   As Long
    Dim arrSubKey       As Variant
    
    On Error GoTo ErrH
    Call Logger.PushMethod("ZLHelperMain.clsRegistry.GetRegSubKeys", strKey, blnSucess)
    lngsamDesired = KEY_ENUMERATE_SUB_KEYS Or KEY_QUERY_VALUE
    arrSubKey = Array()
    blnSucess = False
    If GetRegKeyInfo(strKey, hRootKey, strSubKey, blnMustCloseRootKey, lngsamDesired) Then
        GetRegSubKeys = GetRegSubKeys_Sub(hRootKey, strSubKey, lngsamDesired, blnSucess)
        'RootKey是线程模拟用户得到的，需要关闭
        If blnMustCloseRootKey Then
            lngError = RegCloseKey(hRootKey)
            If lngError <> ERROR_SUCCESS Then
                blnSucess = False
                Logger.Error "RegCloseKey(ImpersonateLoggedOnUser)", "Error", GetLastDllErr(lngError)
            Else
                hRootKey = 0
            End If
        End If
    Else
        GetRegSubKeys = arrSubKey
    End If
    Call Logger.PopMethod("ZLHelperMain.clsRegistry.GetRegSubKeys", blnSucess, UboundEx(arrSubKey))
    Exit Function
ErrH:
    If Logger.ErrCenter("ZLHelperMain.clsRegistry.GetRegSubKeys") = 1 Then
        Resume
    End If
    If blnMustCloseRootKey And hRootKey <> 0 Then
        lngError = RegCloseKey(hRootKey)
        If lngError <> ERROR_SUCCESS Then
            Logger.Error "RegCloseKey(ImpersonateLoggedOnUser)", "Error", GetLastDllErr(lngError)
        End If
    End If
    Call Logger.PopMethod("ZLHelperMain.clsRegistry.GetRegSubKeys")
End Function

'@方法    GetRegKeyValues
'   获取某一个键的所有键值
'@返回值  Boolean
'
'@参数:
'strKey String In
'   注册表键位，如“HKEY_CURRENT_USER\Printers\DevModePerUser"
'varName Variant Out
'   返回键值名称
'varValues Variant Out
'   返回键值，和名称按索引一一对应
'@备注
'
Public Function GetRegKeyValues(ByVal strKey As String, ByRef varNames As Variant, ByRef varValues As Variant) As Boolean
    Dim strSubKey       As String, hRootKey         As Long, blnMustCloseRootKey    As Boolean
    Dim hKey            As Long
    Dim lngError        As Long
    Dim lngsamDesired   As Long
    Dim ftTmp           As FILETIME
    Dim lngValuesCount  As Long, lngMaxValueLenB    As Long, lngMaxNameLenU         As Long
    Dim strName         As String, lngIdx           As Long, lngValueType           As Long, lngCurDataLenB         As Long, lngCurNameLenB     As Long
    Dim arrBytBuffer()  As Byte, varTmp             As Variant
    Dim blnSucess       As Boolean
    On Error GoTo ErrH
    Call Logger.PushMethod("ZLHelperMain.clsRegistry.GetRegKeyValues", strKey)
    lngsamDesired = KEY_QUERY_VALUE
    varNames = Array()
    varValues = Array()
    blnSucess = False
    If GetRegKeyInfo(strKey, hRootKey, strSubKey, blnMustCloseRootKey, lngsamDesired) Then
        lngError = RegOpenKeyEx(hRootKey, StrPtr(strSubKey), 0, lngsamDesired, hKey)
        If lngError <> ERROR_SUCCESS Then
            If lngError = ERROR_FILE_NOT_FOUND Then
                Logger.Warn "RegOpenKeyEx", "Error", GetLastDllErr(lngError)
            Else
                Logger.Error "RegOpenKeyEx", "Error", GetLastDllErr(lngError)
            End If
        Else
            lngError = RegQueryInfoKey(hKey, 0, ByVal 0, ByVal 0, ByVal 0, ByVal 0, ByVal 0, lngValuesCount, lngMaxNameLenU, lngMaxValueLenB, ByVal 0, ftTmp)
            If lngError <> ERROR_SUCCESS Then
                Logger.Error "RegQueryInfoKey", "Error", GetLastDllErr(lngError)
            Else

                If lngValuesCount <> 0 Then
                    ReDim Preserve arrBytBuffer(lngMaxValueLenB - 1)
                    ReDim Preserve varNames(lngValuesCount - 1)
                    ReDim Preserve varValues(lngValuesCount - 1)
                    Do
                        strName = String(lngMaxNameLenU, Chr(0))
                        varTmp = Empty
                        lngCurDataLenB = 0
                        lngCurNameLenB = (lngMaxNameLenU) * 2
                        lngValueType = REG_NONE
                        lngCurDataLenB = lngMaxValueLenB
                        lngError = RegEnumValue(hKey, lngIdx, StrPtr(strName), lngCurNameLenB, 0, lngValueType, ByVal VarPtr(arrBytBuffer(0)), lngCurDataLenB)
                        If lngError = ERROR_SUCCESS Then
                            varNames(lngIdx) = TruncZero(strName)
                            If ConvertByteToNormalData(arrBytBuffer, lngValueType, varTmp, lngCurDataLenB - 1) Then
                                varValues(lngIdx) = varTmp
                            End If
                        ElseIf lngError = ERROR_NO_MORE_ITEMS Then
                            blnSucess = True
                            Logger.DebugEx "RegEnumValue", "EnumEnd", True
                        Else
                            Logger.Error "RegEnumValue", "Error", GetLastDllErr(lngError), "Index", lngIdx
                        End If
                        lngIdx = lngIdx + 1
                    Loop Until lngError <> ERROR_SUCCESS
                Else
                    blnSucess = True
                    Logger.DebugEx "RegQueryInfoKey", "ValuesCount", lngValuesCount, "MaxNameLenU", lngMaxNameLenU, "MaxValueLenB", lngMaxValueLenB
                End If
            End If
            lngError = RegCloseKey(hKey)
            If lngError <> ERROR_SUCCESS Then
                blnSucess = False
                Logger.Error "RegCloseKey", "Error", GetLastDllErr(lngError)
            Else
                hKey = 0
            End If
        End If
        'RootKey是线程模拟用户得到的，需要关闭
        If blnMustCloseRootKey Then
            lngError = RegCloseKey(hRootKey)
            If lngError <> ERROR_SUCCESS Then
                blnSucess = False
                Logger.Error "RegCloseKey(ImpersonateLoggedOnUser)", "Error", GetLastDllErr(lngError)
            Else
                hRootKey = 0
            End If
        End If
    End If
    GetRegKeyValues = blnSucess
    Call Logger.PopMethod("ZLHelperMain.clsRegistry.GetRegKeyValues", GetRegKeyValues)
    Exit Function
ErrH:
    If Logger.ErrCenter("ZLHelperMain.clsRegistry.GetRegKeyValues") = 1 Then
        Resume
    End If
    If hKey <> 0 Then
        lngError = RegCloseKey(hKey)
        If lngError <> ERROR_SUCCESS Then
            blnSucess = False
            Logger.Error "RegCloseKey", "Error", GetLastDllErr(lngError)
        End If
    End If
    If blnMustCloseRootKey And hRootKey <> 0 Then
        lngError = RegCloseKey(hRootKey)
        If lngError <> ERROR_SUCCESS Then
            Logger.Error "RegCloseKey(ImpersonateLoggedOnUser)", "Error", GetLastDllErr(lngError)
        End If
    End If
    Call Logger.PopMethod("ZLHelperMain.clsRegistry.GetRegKeyValues")
End Function


'---------------------------------------------------------------------------
'                4。私有方法
'---------------------------------------------------------------------------

'@方法    GetRegKeyInfo
'   分隔键值
'@返回值  Boolean
'   是否分隔病处理成功
'@参数:
'strKey String In
'   注册表键位，如“HKEY_CURRENT_USER\Printers\DevModePerUser"
'hRootKey String Out
'   返回根键的句柄。当为HKEY_CURRENT_USER，且开启用户模拟时，自动调用模拟线程用户的HKEY_CURRENT_USER
'strSubKey String Out
'   去除根键的子健
'blnMustCoseRootKey Boolean Out
'   HKEY_CURRENT_USER且开启用户模拟时，需要手动关闭根键句柄
'@备注
'
Private Function GetRegKeyInfo(ByVal strKey As String, ByRef hRootKey As Long, ByRef strSubKey As String, ByRef blnMustCoseRootKey As Boolean, ByVal samDesired As Long) As Boolean
    Dim strRoot     As String, lngPos       As String
    Dim lngReturn   As Long
    
    On Error GoTo ErrH
    Call Logger.PushMethod("ZLHelperMain.clsRegistry.GetRegKeyInfo", strKey, hRootKey, strSubKey, blnMustCoseRootKey, samDesired)
    hRootKey = 0: strSubKey = "": blnMustCoseRootKey = False
    lngPos = InStr(UCase(strKey), "\")
    If lngPos <> 0 Then
        strRoot = Trim(Mid(strKey, 1, lngPos - 1))
        strSubKey = Trim(Mid(strKey, lngPos + 1))
    Else
        strRoot = Trim(strKey)
    End If
    
    hRootKey = DecodeEx(strRoot, "HKEY_CLASSES_ROOT", HKEY_CLASSES_ROOT, "HKEY_CURRENT_USER", HKEY_CURRENT_USER, _
                                "HKEY_LOCAL_MACHINE", HKEY_LOCAL_MACHINE, "HKEY_USERS", HKEY_USERS, _
                                "HKEY_PERFORMANCE_DATA", HKEY_PERFORMANCE_DATA, "HKEY_CURRENT_CONFIG", HKEY_CURRENT_CONFIG, _
                                "HKEY_DYN_DATA", HKEY_DYN_DATA, 0)
    If hRootKey <> 0 Then
        GetRegKeyInfo = True
        If hRootKey = HKEY_CURRENT_USER And mblnImpersonateLoggedOnUserOn Then
            hRootKey = 0
            lngReturn = RegOpenCurrentUser(samDesired, hRootKey)
            If lngReturn <> ERROR_SUCCESS Then
                GetRegKeyInfo = False
                Logger.Error "RegOpenCurrentUser", "Error", GetLastDllErr(lngReturn)
            Else
                blnMustCoseRootKey = True
                If Not mblnManualClose Then
                    mblnImpersonateLoggedOnUserOn = False
                End If
            End If
        End If
    End If
    Call Logger.PopMethod("ZLHelperMain.clsRegistry.GetRegKeyInfo", GetRegKeyInfo, hRootKey, strSubKey, blnMustCoseRootKey)
    Exit Function
ErrH:
    If Logger.ErrCenter("ZLHelperMain.clsRegistry.GetRegKeyInfo") = 1 Then
        Resume
    End If
    
    Call Logger.PopMethod("ZLHelperMain.clsRegistry.GetRegKeyInfo")
End Function

'@方法    ConvertDataToByte
'   将数据转换为字节数据，用户键值的存储使用
'@返回值  Boolean
'   是否转换成功
'@参数:
'varData Variant In
'   入参
'arrBytOut Byte() In/Out
'   接受数据的缓冲区
'lngRegValueType Long In/Out
'   要保存值的类型。或者返回（lngRegValueType=0时）
'lngStartPos Long In
'   缓冲区的写入位置
'@备注
'
'   键的值。若该值不传，且strValueName="",则不添加键值。该数据必须传入正确的类型。若值不存在，则自动根据传入的类型自动转换。
'       String->REG_SZ,Long\Integer\Byte->REG_DWORD
'       String()->REG_MULTI_SZ
'       Byte()->REG_BINARY
Private Function ConvertDataToByte(varData As Variant, ByRef arrBytOut() As Byte, ByRef lngRegValueType As Long) As Boolean
    Dim i               As Long
    Dim strTmp          As String
    Dim arrBytTmp()     As Byte
    On Error GoTo ErrH
    Call Logger.PushMethod("ZLHelperMain.clsRegistry.ConvertDataToByte", varData, arrBytOut, lngRegValueType)
    ConvertDataToByte = True
    If lngRegValueType = REG_NONE Then
        Select Case VarType(varData)
            Case vbString
                lngRegValueType = REG_SZ
                Call AddByteArrayToBuffer(arrBytOut, String2Byte(varData & "", True))
            Case vbLong, vbInteger, vbByte
                lngRegValueType = REG_DWORD
                Call AddByteArrayToBuffer(arrBytOut, Long2Byte(Val(varData & "")))
            Case vbString + vbArray
                lngRegValueType = REG_MULTI_SZ
                
                For i = LboundEx(varData) To UboundEx(varData)
                    Call AddByteArrayToBuffer(arrBytOut, String2Byte(varData(i) & "", True), UboundEx(arrBytOut) + 1)
                Next
                ReDim Preserve arrBytOut(UboundEx(arrBytOut) + 2)
            Case vbByte + vbArray
                lngRegValueType = REG_BINARY
                arrBytOut = varData
            Case Else
                ConvertDataToByte = False
                Logger.Error "ConvertDataToByte", "Error", "传入不支持的数据类型" & TypeName(varData) & ",无法确定键值类型"
        End Select
    Else
        '可能是缺失值，则自动指定一个缺省值
        If IsMissing(varData) And VarType(varData) = vbVariant Then
            Select Case lngRegValueType
                Case REG_SZ, REG_EXPAND_SZ, REG_MULTI_SZ
                    varData = ""
                Case REG_DWORD, REG_BINARY
                    varData = CLng(0)
            End Select
        End If
        Select Case lngRegValueType
            Case REG_SZ, REG_EXPAND_SZ
                Select Case VarType(varData)
                    Case vbString
                        Call AddByteArrayToBuffer(arrBytOut, String2Byte(varData & "", True))
                    Case vbLong, vbInteger, vbByte
                        Call AddByteArrayToBuffer(arrBytOut, String2Byte(varData & "", True))
                    Case vbString + vbArray
                        ConvertDataToByte = False
                        Logger.Error "ConvertDataToByte", "Error", "注册表键值类型" & IIf(lngRegValueType = REG_SZ, "REG_SZ", "REG_EXPAND_SZ") & "不支持数据类型String()"
                    Case vbByte + vbArray
                        arrBytOut = varData
                    Case Else
                        ConvertDataToByte = False
                        Logger.Error "ConvertDataToByte", "Error", "注册表键值类型" & IIf(lngRegValueType = REG_SZ, "REG_SZ", "REG_EXPAND_SZ") & "不支持数据类型" & TypeName(varData)
                End Select
            Case REG_DWORD
                Select Case VarType(varData)
                    Case vbString
                        Call AddByteArrayToBuffer(arrBytOut, Long2Byte(Val(varData & "")))
                    Case vbLong, vbInteger, vbByte
                        Call AddByteArrayToBuffer(arrBytOut, Long2Byte(CLng(varData)))
                    Case vbString + vbArray
                        ConvertDataToByte = False
                        Logger.Error "ConvertDataToByte", "Error", "注册表键值类型REG_DWORD不支持数据类型String()"
                    Case vbByte + vbArray
                        arrBytTmp = arrBytTmp
                        Call AddByteArrayToBuffer(arrBytOut, Long2Byte(Byte2Long(arrBytTmp)))
                    Case Else
                        ConvertDataToByte = False
                        Logger.Error "ConvertDataToByte", "Error", "注册表键值类型REG_DWORD不支持数据类型" & TypeName(varData)
                End Select
            Case REG_BINARY
                Select Case VarType(varData)
                    Case vbString   '支持16进制字符串。可以包含空格
                        Call AddByteArrayToBuffer(arrBytOut, HexStringToByte(Replace(varData, " ", "")))
                    Case vbLong, vbInteger, vbByte
                        Call AddByteArrayToBuffer(arrBytOut, Long2Byte(CLng(varData)))
                    Case vbString + vbArray
                        Logger.Error "ConvertDataToByte", "Error", "注册表键值类型REG_BINARY不支持数据类型String()"
                    Case vbByte + vbArray
                        arrBytOut = varData
                    Case Else
                        ConvertDataToByte = False
                        Logger.Error "ConvertDataToByte", "Error", "注册表键值类型REG_BINARY不支持数据类型" & TypeName(varData)
                End Select
            Case REG_MULTI_SZ
                Select Case VarType(varData)
                    Case vbString, vbLong, vbInteger, vbByte
                        strTmp = TruncZero(CStr(varData))
                        Call AddByteArrayToBuffer(arrBytOut, String2Byte(strTmp, True))
                        ReDim Preserve arrBytOut(UboundEx(arrBytOut) + 2)
                    Case vbByte + vbArray
                        arrBytTmp = varData
                        strTmp = arrBytTmp
                        strTmp = TruncZero(strTmp)
                        Call AddByteArrayToBuffer(arrBytOut, String2Byte(strTmp, True))
                        ReDim Preserve arrBytOut(UboundEx(arrBytOut) + 2)
                    Case vbString + vbArray
                        For i = LboundEx(varData) To UboundEx(varData)
                            Call AddByteArrayToBuffer(arrBytOut, String2Byte(varData(i) & "", True), UboundEx(arrBytOut) + 1)
                        Next
                        ReDim Preserve arrBytOut(UboundEx(arrBytOut) + 2)
                    Case Else
                        ConvertDataToByte = False
                        Logger.Error "ConvertDataToByte", "Error", "注册表键值类型REG_MULTI_SZ不支持数据类型" & TypeName(varData)
                End Select
            Case Else
                Logger.Error "ConvertDataToByte", "Error", "不支持注册表键值类型" & lngRegValueType & "的处理"
        End Select
    End If
    Call Logger.PopMethod("ZLHelperMain.clsRegistry.ConvertDataToByte", ConvertDataToByte, lngRegValueType)
    Exit Function
ErrH:
    If Logger.ErrCenter("ZLHelperMain.clsRegistry.ConvertDataToByte") = 1 Then
        Resume
    End If
    
    Call Logger.PopMethod("ZLHelperMain.clsRegistry.ConvertDataToByte")
End Function

'@方法    ConvertByteToNormalData
'   将接受的字节数据转换为指定的数据
'@返回值  Variant
'
'@参数:
'arrBytIn Byte() In
'   接受的数据
'lngRegValueType Long In
'   需要返回的类型
'lngDataByteUbound  Long In
'   数据缓存的上限
'@备注
'
Private Function ConvertByteToNormalData(arrBytIn() As Byte, ByVal lngRegValueType As Long, ByRef varRet As Variant, Optional ByVal lngDataByteUbound As Long = -1) As Boolean
    Dim i           As Long
    Dim strRet      As String
    Dim arrStrRet() As String
    Dim arrBytRet() As Byte
    Dim lngRet      As Long
    Dim arrTmp      As Variant
    
    On Error GoTo ErrH
    Call Logger.PushMethod("ZLHelperMain.clsRegistry.ConvertByteToNormalData", arrBytIn, lngRegValueType, varRet, lngDataByteUbound)
    If lngDataByteUbound = -1 Then
        lngDataByteUbound = UboundEx(arrBytIn)
    End If
    ConvertByteToNormalData = True
    Select Case lngRegValueType
        Case REG_SZ
            strRet = Byte2String(arrBytIn, , lngDataByteUbound)
            varRet = strRet
        Case REG_EXPAND_SZ
            strRet = ExpandEnvStr(Byte2String(arrBytIn, , lngDataByteUbound))
            varRet = strRet
        Case REG_DWORD
            lngRet = Byte2Long(arrBytIn)
            varRet = lngRet
        Case REG_BINARY
            If UboundEx(arrBytIn) = lngDataByteUbound Then
                varRet = arrBytIn
            Else
                ReDim Preserve arrBytRet(lngDataByteUbound)
                RtlMoveMemory ByVal VarPtr(arrBytRet(0)), ByVal VarPtr(arrBytIn(0)), lngDataByteUbound + 1
                varRet = arrBytRet
            End If
        Case REG_MULTI_SZ
            If UboundEx(arrBytIn) > -1 Then
                If UboundEx(arrBytIn) = lngDataByteUbound Then
                    strRet = arrBytIn
                Else
                    ReDim Preserve arrBytRet(lngDataByteUbound)
                    RtlMoveMemory ByVal VarPtr(arrBytRet(0)), ByVal VarPtr(arrBytIn(0)), lngDataByteUbound + 1
                    strRet = arrBytRet
                End If
                arrTmp = Split(Replace(strRet, Chr(0) & Chr(0), Chr(0)), Chr(0))
                If UboundEx(arrTmp) > 0 Then  '最后一个是chr(0)
                    ReDim Preserve arrStrRet(UboundEx(arrTmp) - 1)
                    For i = LboundEx(arrTmp) To UboundEx(arrTmp) - 1
                        arrStrRet(i) = arrTmp(i)
                    Next
                End If
            End If
            varRet = arrStrRet
        Case Else
            ConvertByteToNormalData = False
            Logger.Error "ConvertDataToByte", "Error", "不支持注册表键值类型" & lngRegValueType & "的处理"
    End Select
    Call Logger.PopMethod("ZLHelperMain.clsRegistry.ConvertByteToNormalData", ConvertByteToNormalData, varRet)
    Exit Function
ErrH:
    If Logger.ErrCenter("ZLHelperMain.clsRegistry.ConvertByteToNormalData") = 1 Then
        Resume
    End If
    
    Call Logger.PopMethod("ZLHelperMain.clsRegistry.ConvertByteToNormalData")
End Function
'@方法    Byte2Long
'   将Byte 转换为Longe
'@返回值  Byte()
'
'@参数:
'arrBytIn Byte() In
'   要转换的值
'lngStartPos Long In
'   起始位置
'lngDataByteUbound  Long In (Optional Default=-1)
'   数据缓存的上限,-1时取原始上限
'@备注
'   不足四字节，补足四字节
Private Function Byte2Long(ByRef arrBytIn() As Byte, Optional ByVal lngStartPos As Long, Optional ByVal lngDataByteUbound As Long = -1) As Long
    Dim arrByteRet(3)        As Byte
    On Error GoTo ErrH
    Call Logger.PushMethod("ZLHelperMain.clsRegistry.Byte2Long", arrBytIn, lngStartPos, lngDataByteUbound)
    If lngDataByteUbound = -1 Then
        lngDataByteUbound = UboundEx(arrBytIn)
    End If
    If lngStartPos <= lngDataByteUbound Then
        If lngStartPos + 3 <= lngDataByteUbound Then
            RtlMoveMemory arrByteRet(0), arrBytIn(lngStartPos), 4
        Else
            RtlMoveMemory arrByteRet(0), arrBytIn(lngStartPos), lngDataByteUbound - lngStartPos + 1
        End If
        RtlMoveMemory Byte2Long, arrByteRet(0), 4
    End If
    Call Logger.PopMethod("ZLHelperMain.clsRegistry.Byte2Long", Byte2Long)
    Exit Function
ErrH:
    If Logger.ErrCenter("ZLHelperMain.clsRegistry.Byte2Long") = 1 Then
        Resume
    End If
    
    Call Logger.PopMethod("ZLHelperMain.clsRegistry.Byte2Long")
End Function

'@方法    Long2Byte
'   将Long转换为Byte
'@返回值  Byte()
'
'@参数:
'lngValue Long In
'   要转换的值
'@备注
'
Private Function Long2Byte(ByVal lngValue As Long) As Byte()
    Dim arrByteRet(3)        As Byte
    On Error GoTo ErrH
    Call Logger.PushMethod("ZLHelperMain.clsRegistry.Long2Byte", lngValue)
    RtlMoveMemory arrByteRet(0), lngValue, 4
    Long2Byte = arrByteRet
    Call Logger.PopMethod("ZLHelperMain.clsRegistry.Long2Byte", Long2Byte)
    Exit Function
ErrH:
    If Logger.ErrCenter("ZLHelperMain.clsRegistry.Long2Byte") = 1 Then
        Resume
    End If
    
    Call Logger.PopMethod("ZLHelperMain.clsRegistry.Long2Byte")
End Function
'@方法    Byte2String
'   将Byte转换为String
'@返回值  String
'
'@参数:
'arrBytIn Byte() In
'   要转换的值
'lngStartPos Long In (Optional)
'   起始位置
'lngDataByteUbound  Long In (Optional Default=-1)
'   数据缓存的上限,-1时取原始上限
'@备注
'   支持BSTR类型,自动去掉\0字符
Private Function Byte2String(ByRef arrBytIn() As Byte, Optional ByVal lngStartPos As Long, Optional ByVal lngDataByteUbound As Long = -1) As String
    Dim arrTmp()    As Byte
    
    On Error GoTo ErrH
    Call Logger.PushMethod("ZLHelperMain.clsRegistry.Byte2String", arrBytIn, lngStartPos, lngDataByteUbound)
    If lngDataByteUbound = -1 Then
        lngDataByteUbound = UboundEx(arrBytIn)
    End If
    If lngDataByteUbound > lngStartPos Then
        If lngStartPos = LboundEx(arrBytIn) And UboundEx(arrBytIn) = lngDataByteUbound Then
            Byte2String = arrBytIn
        Else
            If lngDataByteUbound >= lngStartPos And lngStartPos >= LboundEx(arrBytIn) Then
                ReDim Preserve arrTmp(lngDataByteUbound - lngStartPos)
                RtlMoveMemory ByVal VarPtr(arrTmp(0)), ByVal VarPtr(arrBytIn(lngStartPos)), lngDataByteUbound - lngStartPos + 1
                Byte2String = arrTmp
            End If
        End If
    End If
    Byte2String = TruncZero(Byte2String)
    Call Logger.PopMethod("ZLHelperMain.clsRegistry.Byte2String", Byte2String)
    Exit Function
ErrH:
    If Logger.ErrCenter("ZLHelperMain.clsRegistry.Byte2String") = 1 Then
        Resume
    End If
    
    Call Logger.PopMethod("ZLHelperMain.clsRegistry.Byte2String")
End Function

'@方法    String2Byte
'   将String转换为Byte
'@返回值  Byte()
'
'@参数:
'strValue String In
'   要转换的值
'@备注
'   支持BSTR类型
Private Function String2Byte(ByRef strValue As String, Optional ByVal blnAddZero As Boolean) As Byte()
    Dim arrRet()        As Byte
    
    On Error GoTo ErrH
    Call Logger.PushMethod("ZLHelperMain.clsRegistry.String2Byte", strValue, blnAddZero)
    arrRet = strValue
    If blnAddZero Then
        ReDim Preserve arrRet(UboundEx(arrRet) + 2)
    End If
    String2Byte = arrRet
    Call Logger.PopMethod("ZLHelperMain.clsRegistry.String2Byte", String2Byte)
    Exit Function
ErrH:
    If Logger.ErrCenter("ZLHelperMain.clsRegistry.String2Byte") = 1 Then
        Resume
    End If
    
    Call Logger.PopMethod("ZLHelperMain.clsRegistry.String2Byte")
End Function

'@方法    AddByteArrayToBuffer
'   添加一个字节组到缓冲区
'@返回值
'
'@参数:
'arrBytOut Byte() In/Out
'   缓冲区
'arrBytIn Byte() In/Out
'   准备写入的数据
'lngStartPos    Long In
'   指定写入位置
'@备注
'
Private Sub AddByteArrayToBuffer(arrBytOut() As Byte, arrBytIn() As Byte, Optional ByVal lngStartPos As Long)
    On Error GoTo ErrH
    Call Logger.PushMethod("ZLHelperMain.clsRegistry.AddByteArrayToBuffer", arrBytOut, arrBytIn, lngStartPos)
    If UboundEx(arrBytOut) = -1 Then
        ReDim Preserve arrBytOut(UboundEx(arrBytIn))
    End If
    If lngStartPos + UboundEx(arrBytIn) > UboundEx(arrBytOut) Then
        ReDim Preserve arrBytOut(lngStartPos + UboundEx(arrBytIn))
    End If
    RtlMoveMemory arrBytOut(lngStartPos), arrBytIn(0), UboundEx(arrBytIn) + 1
    Call Logger.PopMethod("ZLHelperMain.clsRegistry.AddByteArrayToBuffer", arrBytOut)
    Exit Sub
ErrH:
    If Logger.ErrCenter("ZLHelperMain.clsRegistry.AddByteArrayToBuffer") = 1 Then
        Resume
    End If
    
    Call Logger.PopMethod("ZLHelperMain.clsRegistry.AddByteArrayToBuffer")
End Sub

'@方法    UboundEx
'   获取  Ubound
'@返回值  Long
'
'@参数:
'varArray Variant In
'   传入的数组
'@备注
'
Private Function UboundEx(varArray As Variant) As Long
    On Error GoTo ErrH
    UboundEx = UBound(varArray)
    Exit Function
ErrH:
    UboundEx = -1
End Function

'@方法    LboundEx
'   获取  Lbound
'@返回值  Long
'
'@参数:
'varArray Variant In
'   传入的数组
'@备注
'
Private Function LboundEx(varArray As Variant) As Long
    On Error GoTo ErrH
    LboundEx = LBound(varArray)
    Exit Function
ErrH:
    LboundEx = 0
End Function

'@方法    HexStringToByte
'   将16进制字符串转换为字节组
'@返回值  Byte()
'
'@参数:
'strInput String    In
'   16进制字符串,其中可能有空格
'lngRetBytLen String In
'   指定返回的字节组的长度,0-按原始长度返回，<>0返回指定的长度，不足补齐（补0），多了截取
'@备注
'
Public Function HexStringToByte(ByVal strInput As String, Optional ByVal lngRetBytLen As Long) As Byte()
    Dim arrReturn() As Byte
    Dim i           As Long
    Dim lngLen      As Long
    
    On Error GoTo ErrH
    Call Logger.PushMethod("ZLHelperMain.clsRegistry.HexStringToByte", strInput, lngRetBytLen)
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
    Call Logger.PopMethod("ZLHelperMain.clsRegistry.HexStringToByte", arrReturn)
    Exit Function
ErrH:
    If Logger.ErrCenter("ZLHelperMain.clsRegistry.HexStringToByte") = 1 Then
        Resume
    End If
    
    Call Logger.PopMethod("ZLHelperMain.clsRegistry.HexStringToByte")
End Function


'@方法    ExpandEnvStr
'   将字符串中的环境变量替换为常规值.如： %PATH% 则返回 "c:\;c:\windows;"
'@返回值  String
'
'@参数:
'strInput String In
'   包含环境变量的字符串
'
'@备注
'
Private Function ExpandEnvStr(ByVal strInput As String) As String
    Dim lngLen As Long, strBuf As String, strOld As String
    On Error GoTo ErrH
    Call Logger.PushMethod("ZLHelperMain.clsRegistry.ExpandEnvStr", strInput)
    strOld = strInput & "  " ' 不知为什么要加两个字符，否则返回值会少最后两个字符！
    strBuf = ""
    lngLen = ExpandEnvironmentStrings(StrPtr(strOld), StrPtr(strBuf), lngLen)
    If Err.LastDllError <> 0 Then
        Logger.DebugEx "ExpandEnvironmentStrings", "Error", GetLastDllErr(Err.LastDllError)
    End If
    '展开字符串
    strBuf = String$(lngLen + 2, Chr$(0))
    lngLen = ExpandEnvironmentStrings(StrPtr(strOld), StrPtr(strBuf), LenB(strBuf))
    If Err.LastDllError <> 0 Then
        Logger.Error "ExpandEnvironmentStrings", "Error", GetLastDllErr(Err.LastDllError)
    End If
    '返回环境变量
    ExpandEnvStr = TruncZero(strBuf)
    Call Logger.PopMethod("ZLHelperMain.clsRegistry.ExpandEnvStr", ExpandEnvStr)
    Exit Function
ErrH:
    If Logger.ErrCenter("ZLHelperMain.clsRegistry.ExpandEnvStr") = 1 Then
        Resume
    End If
    
    Call Logger.PopMethod("ZLHelperMain.clsRegistry.ExpandEnvStr")
End Function

'@方法    GetRegSubKeys_Sub
'   GetRegSubKeys,DeleteKey的子函数
'@返回值  Variant
'
'@参数:
'hKey Long In
'   枚举的键的句柄
'strSubKey Long In
'   枚举的子健
'lngsamDesired Long In
'   打开子健的权限常量
'blnSucess Boolean Out(Optional)
'   是否成功枚举
'@备注
'
Private Function GetRegSubKeys_Sub(ByVal hKey As Long, ByVal strSubKey As String, ByVal lngsamDesired As Long, Optional ByRef blnSucess As Boolean) As Variant
    Dim lngError        As Long, lngCurKey                  As Long
    Dim lngMaxSubKeyLenU   As Long, lngSubKeyCount          As Long
    Dim ftTmp           As FILETIME
    Dim arrSubKey       As Variant
    Dim strName         As String, lngIdx                   As Long
    Dim blnSucessTmp    As Boolean
    On Error GoTo ErrH
    Call Logger.PushMethod("ZLHelperMain.clsRegistry.GetRegSubKeys_Sub", hKey, strSubKey, lngsamDesired, blnSucess)
    arrSubKey = Array()
    blnSucessTmp = False
    lngError = RegOpenKeyEx(hKey, StrPtr(strSubKey), 0, lngsamDesired, lngCurKey)
    If lngError <> ERROR_SUCCESS Then
        If lngError = ERROR_FILE_NOT_FOUND Then
            Logger.Warn "RegOpenKeyEx", "Error", GetLastDllErr(lngError)
        Else
            Logger.Error "RegOpenKeyEx", "Error", GetLastDllErr(lngError)
        End If
    Else
        lngError = RegQueryInfoKey(lngCurKey, 0, ByVal 0, ByVal 0, lngSubKeyCount, lngMaxSubKeyLenU, ByVal 0, ByVal 0, ByVal 0, ByVal 0, ByVal 0, ftTmp)
        If lngError <> ERROR_SUCCESS Then
            Logger.Error "RegQueryInfoKey", "Error", GetLastDllErr(lngError)
        Else
            If lngSubKeyCount <> 0 Then
                lngMaxSubKeyLenU = lngMaxSubKeyLenU
                ReDim Preserve arrSubKey(lngSubKeyCount - 1)
                Do
                    strName = String(lngMaxSubKeyLenU, Chr(0))
                    lngError = RegEnumKey(lngCurKey, lngIdx, StrPtr(strName), Len(strName))
                    If lngError = ERROR_SUCCESS Then
                        arrSubKey(lngIdx) = Left(strName, InStr(strName, Chr(0)) - 1)
                    ElseIf lngError = ERROR_NO_MORE_ITEMS Then
                        blnSucessTmp = True
                        Logger.DebugEx "RegEnumKey", "EnumEnd", True
                    Else
                        Logger.Error "RegEnumKey", "Error", GetLastDllErr(lngError)
                    End If
                    lngIdx = lngIdx + 1
                Loop Until lngError <> ERROR_SUCCESS
            Else
                blnSucessTmp = True
                Logger.DebugEx "RegQueryInfoKey", "SubKeyCount", lngSubKeyCount, "MaxSubKeyLenU", lngMaxSubKeyLenU
            End If
        End If
        lngError = RegCloseKey(lngCurKey)
        If lngError <> ERROR_SUCCESS Then
            blnSucessTmp = False
            Logger.Error "RegCloseKey", "Error", GetLastDllErr(lngError)
        Else
            lngCurKey = 0
        End If
    End If
    blnSucess = blnSucessTmp
    GetRegSubKeys_Sub = arrSubKey
    Call Logger.PopMethod("ZLHelperMain.clsRegistry.GetRegSubKeys_Sub", blnSucess, arrSubKey)
    Exit Function
ErrH:
    If Logger.ErrCenter("ZLHelperMain.clsRegistry.GetRegSubKeys_Sub") = 1 Then
        Resume
    End If
    If lngCurKey <> 0 Then
        lngError = RegCloseKey(lngCurKey)
        If lngError <> ERROR_SUCCESS Then
            blnSucessTmp = False
            Logger.Error "RegCloseKey", "Error", GetLastDllErr(lngError)
        End If
    End If
    Call Logger.PopMethod("ZLHelperMain.clsRegistry.GetRegSubKeys_Sub")
End Function

'@方法    DeleteRegKey_Sub
'   DeleteKey的子函数，以便提供递归删除功能
'@返回值  Boolean
'
'@参数:
'hKey Long In
'   枚举的键的句柄
'strSubKey Long In
'   枚举的子健
'lngsamDesired Long In
'   打开子健的权限常量
'@备注
Private Function DeleteRegKey_Sub(ByVal hKey As Long, ByVal strSubKey As String, ByVal lngsamDesired As Long) As Boolean
    Dim arrSubKeys      As Variant
    Dim blnSucess       As Boolean, blnSucessTmp    As Boolean
    Dim lngError        As Long
    Dim i               As Long
    
    On Error GoTo ErrH
    Call Logger.PushMethod("ZLHelperMain.clsRegistry.DeleteRegKey_Sub", hKey, strSubKey, lngsamDesired)
    arrSubKeys = GetRegSubKeys_Sub(hKey, strSubKey, lngsamDesired, blnSucess)
    If blnSucess Then
        If TypeName(arrSubKeys) <> "Empty" Then
            If LenB(strSubKey) = 0 Then
                For i = LboundEx(arrSubKeys) To UboundEx(arrSubKeys)
                    blnSucessTmp = DeleteRegKey_Sub(hKey, arrSubKeys(i) & "", lngsamDesired)
                    If Not blnSucessTmp Then
                        blnSucess = False
                    End If
                Next
            Else
                For i = LboundEx(arrSubKeys) To UboundEx(arrSubKeys)
                    blnSucessTmp = DeleteRegKey_Sub(hKey, strSubKey & "\" & arrSubKeys(i), lngsamDesired)
                    If Not blnSucessTmp Then
                        blnSucess = False
                    End If
                Next
            End If
        End If
        lngError = RegDeleteKey(hKey, StrPtr(strSubKey))
        If lngError <> ERROR_SUCCESS Then
            blnSucess = False
            Logger.Error "RegDeleteKey", "Error", GetLastDllErr(lngError)
        End If
    End If
    DeleteRegKey_Sub = blnSucess
    Call Logger.PopMethod("ZLHelperMain.clsRegistry.DeleteRegKey_Sub", DeleteRegKey_Sub)
    Exit Function
ErrH:
    If Logger.ErrCenter("ZLHelperMain.clsRegistry.DeleteRegKey_Sub") = 1 Then
        Resume
    End If
    
    Call Logger.PopMethod("ZLHelperMain.clsRegistry.DeleteRegKey_Sub")
End Function
'---------------------------------------------------------------------------
'                5。对象方法与事件
'---------------------------------------------------------------------------

