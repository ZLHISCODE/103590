Attribute VB_Name = "mdlReg"
Option Explicit



Public Enum REGISTER
    注册信息
    私有模块
    私有全局
    公共模块
    公共全局
End Enum



Public Function UpdateKey(KeyRoot As Long, KeyName As String, SubKeyName As String, SubKeyValue As String) As Boolean
'功能：写注册表
    Dim rc As Long                                      ' 返回代码
    Dim hKey As Long                                    ' 处理一个注册表关键字
    Dim hDepth As Long                                  '
    Dim lpAttr As SECURITY_ATTRIBUTES                   ' 注册表安全类型
    
    lpAttr.nLength = 50                                 ' 设置安全属性为缺省值...
    lpAttr.lpSecurityDescriptor = 0                     ' ...
    lpAttr.bInheritHandle = True                        ' ...

    '------------------------------------------------------------
    '- 创建/打开注册表关键字...
    '------------------------------------------------------------
    rc = RegCreateKeyEx(KeyRoot, KeyName, _
                        0, REG_SZ, _
                        REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, lpAttr, _
                        hKey, hDepth)                   ' 创建/打开//KeyRoot//KeyName
    
    If (rc <> ERROR_SUCCESS) Then GoTo CreateKeyError   ' 错误处理...
    
    '------------------------------------------------------------
    '- 创建/修改关键字值...
    '------------------------------------------------------------
    If (SubKeyValue = "") Then SubKeyValue = " "        ' 要让RegSetValueEx() 工作需要输入一个空格...
    
    ' 创建/修改关键字值
    rc = RegSetValueEx(hKey, SubKeyName, _
                       0, REG_SZ, _
                       SubKeyValue, LenB(StrConv(SubKeyValue, vbFromUnicode)))
                       
    If (rc <> ERROR_SUCCESS) Then GoTo CreateKeyError   ' 错误处理
    '------------------------------------------------------------
    '- 关闭注册表关键字...
    '------------------------------------------------------------
    rc = RegCloseKey(hKey)                              ' 关闭关键字
    
    UpdateKey = True                                    ' 返回成功
    Exit Function                                       ' 退出
CreateKeyError:
    UpdateKey = False                                   ' 设置错误返回代码
    rc = RegCloseKey(hKey)                              ' 试图关闭关键字
End Function

'-------------------------------------------------------------------------------------------------
'sample usage - Debug.Print GetKeyValue(HKEY_CLASSES_ROOT, "COMCTL.ListviewCtrl.1\CLSID", "")
'-------------------------------------------------------------------------------------------------
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

Public Function ReadStartKey() As String
'功能：读取注册表中三个开始时间标志(之一有效即可)
    Dim strKey As String
    strKey = GetKeyValue(HKEY_CURRENT_USER, "SOFTWARE\VTCELUS6CS", "IXPHWP")  'FirstStart,1Start
    If strKey = "" Then strKey = GetKeyValue(HKEY_LOCAL_MACHINE, "SOFTWARE\EG5PZRELSML", "NXPHWP") 'SecondStart,2Start
    If strKey = "" Then strKey = GetKeyValue(HKEY_USERS, ".DEFAULT\SOFTWARE\S1NM9US6CS", "TXPHWP") 'ThirdStart,3Start
    If strKey <> "" Then ReadStartKey = CStr(CDate(strKey))
End Function

Public Function WriteStartKey() As Boolean
'功能:朝注册表中写三个开始时间标志
    Dim curDate As Date
    curDate = Format(Date, "yyyy-MM-dd")
    WriteStartKey = UpdateKey(HKEY_CURRENT_USER, "SOFTWARE\VTCELUS6CS", "IXPHWP", CCur(curDate)) 'FirstStart,1Start
    WriteStartKey = WriteStartKey And UpdateKey(HKEY_LOCAL_MACHINE, "SOFTWARE\EG5PZRELSML", "NXPHWP", CCur(curDate)) 'SecondStart,2Start
    WriteStartKey = WriteStartKey And UpdateKey(HKEY_USERS, ".DEFAULT\SOFTWARE\S1NM9US6CS", "TXPHWP", CCur(curDate)) 'ThirdStart,3Start
End Function

Public Function ReadValidKey() As String
'功能：读取注册表中三个过期标志(之一有效即可)
    Dim strKey As String
    strKey = GetKeyValue(HKEY_CURRENT_USER, "SOFTWARE\PZ7Q64F9", "IRSUTR") 'OneValid,1Valid
    If strKey = "" Then strKey = GetKeyValue(HKEY_LOCAL_MACHINE, "SOFTWARE\SDDQ64F9", "NRSUTR") 'TwoValid,2Valid
    If strKey = "" Then strKey = GetKeyValue(HKEY_USERS, ".DEFAULT\SOFTWARE\S1CKGZHPNO", "TRSUTR") 'ThreeValid,3Valid
    If strKey <> "" Then ReadValidKey = strKey
End Function

Public Function WriteValidKey() As Boolean
    '功能:朝注册表中写三个过期标志
    WriteValidKey = UpdateKey(HKEY_CURRENT_USER, "SOFTWARE\PZ7Q64F9", "IRSUTR", "Q64F9") 'OneValid,1Valid
    WriteValidKey = WriteStartKey And UpdateKey(HKEY_LOCAL_MACHINE, "SOFTWARE\SDDQ64F9", "NRSUTR", "Q64F9") 'TwoValid,2Valid
    WriteValidKey = WriteStartKey And UpdateKey(HKEY_USERS, ".DEFAULT\SOFTWARE\S1CKGZHPNO", "TRSUTR", "Q64F9") 'ThreeValid,3Valid
End Function

