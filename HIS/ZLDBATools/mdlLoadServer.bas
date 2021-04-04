Attribute VB_Name = "mdlLoadServer"
Option Explicit


Public Enum REGRoot
    HKEY_CLASSES_ROOT = &H80000000 '记录Windows操作系统中所有数据文件的格式和关联信息，主要记录不同文件的文件名后缀和与之对应的应用程序。其下子键可分为两类，一类是已经注册的各类文件的扩展名，这类子键前面都有一个“。”；另一类是各类文件类型有关信息。
    HKEY_CURRENT_USER = &H80000001 '此根键包含了当前登录用户的用户配置文件信息。这些信息保证不同的用户登录计算机时，使用自己的个性化设置，例如自己定义的墙纸、自己的收件箱、自己的安全访问权限等。
    HKEY_LOCAL_MACHINE = &H80000002 '此根键包含了当前计算机的配置数据，包括所安装的硬件以及软件的设置。这些信息是为所有的用户登录系统服务的。它是整个注册表中最庞大也是最重要的根键！
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

Public Const READ_CONTROL = &H20000
Public Const KEY_QUERY_VALUE = &H1
Public Const KEY_SET_VALUE = &H2
Public Const KEY_CREATE_SUB_KEY = &H4
Public Const KEY_ENUMERATE_SUB_KEYS = &H8
Public Const KEY_NOTIFY = &H10
Public Const KEY_CREATE_LINK = &H20
Public Const KEY_READ = KEY_QUERY_VALUE + KEY_ENUMERATE_SUB_KEYS + KEY_NOTIFY + READ_CONTROL
Public Const KEY_WRITE = KEY_SET_VALUE + KEY_CREATE_SUB_KEY + READ_CONTROL
Public Const KEY_EXECUTE = KEY_READ
Public Const KEY_ALL_ACCESS = KEY_QUERY_VALUE + KEY_SET_VALUE + _
                       KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS + _
                       KEY_NOTIFY + KEY_CREATE_LINK + READ_CONTROL

' 返回值...
Public Const ERROR_NONE = 0
Public Const ERROR_BADKEY = 2
Public Const ERROR_ACCESS_DENIED = 8
Public Const ERROR_SUCCESS = 0

Public gobjFile As New FileSystemObject

Public Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Public Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Public Declare Function GetCurrentProcess Lib "kernel32" () As Long
Public Declare Function IsWow64Process Lib "kernel32" (ByVal hProc As Long, bWow64Process As Boolean) As Long

Public Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Public Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long

Public Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Public Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Public Declare Function RegEnumKey Lib "advapi32.dll" Alias "RegEnumKeyA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, ByVal cbName As Long) As Long

Public Function TranPasswd(strOld As String) As String
    '------------------------------------------------
    '功能： 密码转换函数
    '参数：
    '   strOld：原密码
    '返回： 加密生成的密码
    '------------------------------------------------
    Dim iBit As Integer, strBit As String
    Dim strNew As String
    If Len(Trim(strOld)) = 0 Then TranPasswd = "": Exit Function
    strNew = ""
    For iBit = 1 To Len(Trim(strOld))
        strBit = UCase(Mid(Trim(strOld), iBit, 1))
        Select Case (iBit Mod 3)
        Case 1
            strNew = strNew & _
                Switch(strBit = "0", "W", strBit = "1", "I", strBit = "2", "N", strBit = "3", "T", strBit = "4", "E", strBit = "5", "R", strBit = "6", "P", strBit = "7", "L", strBit = "8", "U", strBit = "9", "M", _
                   strBit = "A", "H", strBit = "B", "T", strBit = "C", "I", strBit = "D", "O", strBit = "E", "K", strBit = "F", "V", strBit = "G", "A", strBit = "H", "N", strBit = "I", "F", strBit = "J", "J", _
                   strBit = "K", "B", strBit = "L", "U", strBit = "M", "Y", strBit = "N", "G", strBit = "O", "P", strBit = "P", "W", strBit = "Q", "R", strBit = "R", "M", strBit = "S", "E", strBit = "T", "S", _
                   strBit = "U", "T", strBit = "V", "Q", strBit = "W", "L", strBit = "X", "Z", strBit = "Y", "C", strBit = "Z", "X", True, strBit)
        Case 2
            strNew = strNew & _
                Switch(strBit = "0", "7", strBit = "1", "M", strBit = "2", "3", strBit = "3", "A", strBit = "4", "N", strBit = "5", "F", strBit = "6", "O", strBit = "7", "4", strBit = "8", "K", strBit = "9", "Y", _
                   strBit = "A", "6", strBit = "B", "J", strBit = "C", "H", strBit = "D", "9", strBit = "E", "G", strBit = "F", "E", strBit = "G", "Q", strBit = "H", "1", strBit = "I", "T", strBit = "J", "C", _
                   strBit = "K", "U", strBit = "L", "P", strBit = "M", "B", strBit = "N", "Z", strBit = "O", "0", strBit = "P", "V", strBit = "Q", "I", strBit = "R", "W", strBit = "S", "X", strBit = "T", "L", _
                   strBit = "U", "5", strBit = "V", "R", strBit = "W", "D", strBit = "X", "2", strBit = "Y", "S", strBit = "Z", "8", True, strBit)
        Case 0
            strNew = strNew & _
                Switch(strBit = "0", "6", strBit = "1", "J", strBit = "2", "H", strBit = "3", "9", strBit = "4", "G", strBit = "5", "E", strBit = "6", "Q", strBit = "7", "1", strBit = "8", "X", strBit = "9", "L", _
                   strBit = "A", "S", strBit = "B", "8", strBit = "C", "5", strBit = "D", "R", strBit = "E", "7", strBit = "F", "M", strBit = "G", "3", strBit = "H", "A", strBit = "I", "N", strBit = "J", "F", _
                   strBit = "K", "O", strBit = "L", "4", strBit = "M", "K", strBit = "N", "Y", strBit = "O", "D", strBit = "P", "2", strBit = "Q", "T", strBit = "R", "C", strBit = "S", "U", strBit = "T", "P", _
                   strBit = "U", "B", strBit = "V", "Z", strBit = "W", "0", strBit = "X", "V", strBit = "Y", "I", strBit = "Z", "W", True, strBit)
        End Select
    Next
    TranPasswd = strNew

End Function

Public Function OpenConnection(ByVal strServerName As String, ByVal strUserName As String, ByVal strUserPwd As String, ByRef cnOracle As ADODB.Connection) As Boolean
    '------------------------------------------------
    '功能： 打开指定的数据库
    '参数：
    '   strServerName：主机字符串
    '   strUserName：用户名
    '   strUserPwd：密码
    '返回： 数据库打开成功，返回true；失败，返回false
    '------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Dim strSQL As String
    Dim strError As String

    On Error Resume Next
    If Err <> 0 Then Err.Clear
    If cnOracle.State = adStateOpen Then cnOracle.Close
    With cnOracle
        .Provider = "MSDataShape"
        .Open "Driver={Microsoft ODBC for Oracle};Server=" & strServerName, strUserName, strUserPwd
    End With
            
    '保存错误信息
    If Err <> 0 Then strError = Err.Description

    If strError <> "" Then
        If InStr(strError, "自动化错误") > 0 Then
            If 1 = 1 Then
                strError = "msoracl32.dll"
            Else
                strError = "OraOLEDB.dll"
            End If
            strError = "无法创建连接对象，请检查数据访问部件(" & strError & ")是否正常安装并注册。"
        ElseIf InStr(strError, "ORA-12505") > 0 Then
            strError = "ORA-12505,监听程序当前无法识别连接描述符中所给出的 SID,请检查服务名中配置的实例名称。"
            
        ElseIf InStr(strError, "ORA-12170") > 0 Then
            strError = "ORA-12170,连接超时，请检查服务器名是否正确，网络是否可访问，以及是否被服务器防火墙阻止。"
            
        ElseIf InStr(strError, "ORA-12154") > 0 Then
            strError = "ORA-12154,无法分析服务器名，" & vbCrLf & "请检查本机的Oracle配置文件(tnsnames.ora)中是否存在当前使用的服务名。"
            
        ElseIf InStr(strError, "ORA-12541") > 0 Then
            strError = "ORA-12541,无法连接服务器，请检查服务器上的Oracle监听器服务是否启动。"
            
        ElseIf InStr(strError, "ORA-01033") > 0 Then
            strError = "ORA-01033,ORACLE正在初始化或在关闭，请稍候再试。"
            
        ElseIf InStr(strError, "ORA-01034") > 0 Then
            strError = "ORA-01034,ORACLE不可用，请检查数据库实例是否启动。"
            
        ElseIf InStr(strError, "ORA-02391") > 0 Then
            strError = "ORA-02391,用户" & strUserName & "已经登录，不允许重复登录(已达到系统所允许的最大登录数)。"
            
        ElseIf InStr(strError, "ORA-01017") > 0 Then
            strError = "ORA-01017,无效的用户名或密码，登录被拒绝。"
        
        ElseIf InStr(strError, "ORA-28000") > 0 Then
            strError = "ORA-28000,该用户已经被禁用，不允许登录。"
        End If
        
        OpenConnection = False
    Else
        OpenConnection = True
    End If
End Function


Public Function CheckIsDBA(ByRef connthis As Connection) As Boolean
    Dim rstmp As New ADODB.Recordset
    Dim strSQL As String

    strSQL = "Select 1 From User_Role_Privs Where granted_role='DBA'"
    On Error GoTo errh
    rstmp.Open strSQL, connthis, adOpenKeyset, adLockReadOnly

    CheckIsDBA = rstmp.RecordCount > 0
    Exit Function
errh:
    MsgBox Err.Description, vbExclamation
End Function
Public Sub AppendText(KeyAscii As Integer, ByRef cboServer As ComboBox, colServer As Collection)
'功能：向TextBox控件的Text追加内容，并根据当前Text的值在列表中检索可用的完整项目
'参数：KeyAscii    当前的按键
    Dim strTemp As String
    Dim strInput As String
    Dim lngIndex As Long, lngStart As Long
    Dim varItem As Variant
    
    '首先当前用户输入的字符
    If KeyAscii < 0 Or InStr("ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789.", UCase(Chr(KeyAscii))) > 0 Then
        '输入字符只能是数字、英文和汉字
        strInput = Chr(KeyAscii)
        KeyAscii = 0
    End If
    
    With cboServer
        '记录上次的插入点位置
        lngStart = .SelStart + IIf(strInput <> "", 1, 0)
        '接着得到用户击键完成后文本框中出现的内容
        strInput = Mid(.Text, 1, .SelStart) & strInput & Mid(.Text, .SelStart + .SelLength + 1)
    End With
    '根据假想的内容得到可能的列表项
    strTemp = ""
    For Each varItem In colServer
        If UCase(varItem(0)) Like UCase(strInput & "*") Then
            strTemp = varItem(0)
        End If
    Next
    If strTemp <> "" Then
        cboServer.Text = strTemp
        cboServer.SelStart = Len(strInput)
        cboServer.SelLength = 100
    Else
        cboServer.Text = strInput
        cboServer.SelStart = lngStart
    End If
End Sub


Public Function GetAllSubKey(ByVal KeyRoot As Long, KeyName As String) As Variant
'功能:获取某项的所有子项
'返回：=子项数组
    Dim lnghKey As Long, lngRet As Long, strName As String, lngIdx As Long
    Dim strSubKey As Variant
    strSubKey = Array()
    lngIdx = 0: strName = String(256, Chr(0))
    lngRet = RegOpenKey(KeyRoot, KeyName, lnghKey)
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


Public Function GetKeyValue(KeyRoot As Long, KeyName As String, SubKeyRef As String) As String
'功能：读注册表
    Dim i As Long                                           ' 循环计数器
    Dim rc As Long                                          ' 返回代码
    Dim hKey As Long                                        ' 处理打开的注册表关键字
    Dim hDepth As Long                                      '
    Dim sKeyVal As String
    Dim lKeyValType As Long                                 ' 注册表关键字数据类型
    Dim tmpVal As String                                    ' 注册表关键字的临时存储器
    Dim KeyValSize As Long                                  ' 注册表关键字变量尺寸
    
    ' 在 KeyRoot {HKEY_LOCAL_MACHINE...} 下打开注册表关键字
    '------------------------------------------------------------
    rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey) ' 打开注册表关键字
    
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' 处理错误...
    
    tmpVal = String$(1024, 0)                             ' 分配变量空间
    KeyValSize = 1024                                       ' 标记变量尺寸
    
    '------------------------------------------------------------
    ' 检索注册表关键字的值...
    '------------------------------------------------------------
    rc = RegQueryValueEx(hKey, SubKeyRef, 0, _
                         lKeyValType, tmpVal, KeyValSize)    ' 获得/创建关键字的值
                        
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' 错误处理
      
    tmpVal = Left$(tmpVal, InStr(tmpVal, Chr(0)) - 1)

    '------------------------------------------------------------
    ' 决定关键字值的转换类型...
    '------------------------------------------------------------
    Select Case lKeyValType                                  ' 搜索数据类型...
    Case REG_SZ, REG_EXPAND_SZ                              ' 字符串注册表关键字数据类型
        sKeyVal = tmpVal                                     ' 复制字符串的值
    Case REG_DWORD                                          ' 四字节注册表关键字数据类型
        For i = Len(tmpVal) To 1 Step -1                    ' 转换每一位
            sKeyVal = sKeyVal + Hex(Asc(Mid(tmpVal, i, 1)))   ' 一个字符一个字符地生成值。
        Next
        sKeyVal = Format$("&h" + sKeyVal)                     ' 转换四字节为字符串
    End Select
    
    GetKeyValue = sKeyVal                                   ' 返回值
    rc = RegCloseKey(hKey)                                  ' 关闭注册表关键字
    Exit Function                                           ' 退出
    
GetKeyError:    ' 错误发生过后进行清除...
    GetKeyValue = vbNullString                              ' 设置返回值为空字符串
    rc = RegCloseKey(hKey)                                  ' 关闭注册表关键字
End Function


Public Function ValEx(ByVal varInput As Variant) As Variant
'功能：由于Val只能以数字开头识别，ValEx以第一个数字进行识别
    Dim arrTmp As Variant, lngPos As Long
    If Val(varInput) = 0 Then
        varInput = varInput & ""
        If Trim(varInput) = "" Then ValEx = 0: Exit Function
        For lngPos = 1 To Len(varInput)
            If IsNumeric(Mid(varInput, lngPos, 1)) Then Exit For
        Next
        If lngPos = Len(varInput) + 1 Then
            ValEx = 0
        Else
            ValEx = Val(Mid(varInput, lngPos))
        End If
    Else
        ValEx = Val(varInput)
    End If
End Function



