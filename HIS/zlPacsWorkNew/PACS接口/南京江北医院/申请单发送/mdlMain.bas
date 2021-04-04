Attribute VB_Name = "mdlMain"
Option Explicit

Public gcnHIS As New ADODB.Connection
Public gcnPACS As New ADODB.Connection

Public gstrSysName As String                '系统名称
Public gstrDbUser As String                 '当前数据库用户
Public gstrHISUser As String                'HIS数据库用户名
Public gstrHISPassw As String               'HIS数据库密码
Public gstrHISsid As String                 'HIS数据库SID名
Public glngInterval As Long                 '监听时间间隔，单位秒
Public gstrPACSUser As String               'PACS数据库用户名
Public gstrPACSPassw As String              'PACS数据库密码
Public gstrPACSsid As String                'PACS数据库SID名
Public gstrPACSport As String                'PACS端口
Public gstrPACSIP As String                 'PACS数据库服务器的IP地址
Public gstrRegPath As String                '注册表路径




Private Declare Function InitCommonControls Lib "comctl32.dll" () As Long
Public Declare Function ShowWindow Lib "user32" (ByVal Hwnd As Long, ByVal nCmdShow As Long) As Long
Public Const SRCCOPY = &HCC0020
Public Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long


'---------------------------------------------------------------
'-注册表 API 声明...
'---------------------------------------------------------------
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long
Private Declare Function RegCreateKeyEx Lib "advapi32" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, ByRef lpSecurityAttributes As SECURITY_ATTRIBUTES, ByRef phkResult As Long, ByRef lpdwDisposition As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long
Public Declare Function SetActiveWindow Lib "user32" (ByVal Hwnd As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal Hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal Hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal Hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Const GWL_EXSTYLE = (-20)
Public Const WinStyle = &H40000
Public Const SWP_NOSIZE = &H1
Public Const SWP_SHOWWINDOW = &H40
Public Const HWND_TOPMOST = -1

'---------------------------------------------------------------
'- 注册表 Api 常数...
'---------------------------------------------------------------
' Reg Data Types...
Const REG_SZ = 1                         ' Unicode空终结字符串
Const REG_EXPAND_SZ = 2                  ' Unicode空终结字符串
Const REG_DWORD = 4                      ' 32-bit 数字

' 注册表创建类型值...
Const REG_OPTION_NON_VOLATILE = 0       ' 当系统重新启动时，关键字被保留

' 注册表关键字安全选项...
Const READ_CONTROL = &H20000
Const KEY_QUERY_VALUE = &H1
Const KEY_SET_VALUE = &H2
Const KEY_CREATE_SUB_KEY = &H4
Const KEY_ENUMERATE_SUB_KEYS = &H8
Const KEY_NOTIFY = &H10
Const KEY_CREATE_LINK = &H20
Const KEY_READ = KEY_QUERY_VALUE + KEY_ENUMERATE_SUB_KEYS + KEY_NOTIFY + READ_CONTROL
Const KEY_WRITE = KEY_SET_VALUE + KEY_CREATE_SUB_KEY + READ_CONTROL
Const KEY_EXECUTE = KEY_READ
Const KEY_ALL_ACCESS = KEY_QUERY_VALUE + KEY_SET_VALUE + KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS + KEY_NOTIFY + KEY_CREATE_LINK + READ_CONTROL
                     
' 注册表关键字根类型...
Public Const HKEY_LOCAL_MACHINE = &H80000002
Const HKEY_CLASSES_ROOT = &H80000000
Const HKEY_CURRENT_USER = &H80000001
Const HKEY_USERS = &H80000003
Const HKEY_PERFORMANCE_DATA = &H80000004

' 返回值...
Const ERROR_NONE = 0
Const ERROR_BADKEY = 2
Const ERROR_ACCESS_DENIED = 8
Const ERROR_SUCCESS = 0

'---------------------------------------------------------------
'- 注册表安全属性类型...
'---------------------------------------------------------------
Private Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Boolean
End Type


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


Public Function OraDataOpen() As Boolean
    '------------------------------------------------
    '功能： 打开指定的数据库
    '参数：
    '返回： 数据库打开成功，返回true；失败，返回false
    '------------------------------------------------
    Dim strSQL As String
    Dim strError As String
        
    On Error Resume Next
    Err = 0
    DoEvents
    With gcnHIS
        If .State = adStateOpen Then .Close
        .Provider = "MSDataShape"
        .Open "Driver={Microsoft ODBC for Oracle};Server=" & gstrHISsid, gstrHISUser, gstrHISPassw
        If Err <> 0 Then
            '保存错误信息
            strError = Err.Description
            If InStr(strError, "自动化错误") > 0 Then
                MsgBox "连接串无法创建，请检查数据访问部件是否正常安装。", vbInformation, gstrSysName
            ElseIf InStr(strError, "ORA-12154") > 0 Then
                MsgBox "无法分析服务器名，" & vbCrLf & "请检查在Oracle配置中是否存在该本地网络服务名（主机字符串）。", vbInformation, gstrSysName
            ElseIf InStr(strError, "ORA-12541") > 0 Then
                MsgBox "无法连接，请检查服务器上的Oracle监听器服务是否启动。", vbInformation, gstrSysName
            ElseIf InStr(strError, "ORA-01033") > 0 Then
                MsgBox "ORACLE正在初始化或在关闭，请稍候再试。", vbInformation, gstrSysName
            Else
                MsgBox "由于用户、口令或服务器指定错误，无法注册。 " + Err.Description, vbInformation, gstrSysName
            End If
            
            OraDataOpen = False
            Exit Function
        End If
    End With
    
    Err = 0
    On Error GoTo errHand
    
    gstrDbUser = UCase(gstrHISUser)
    SetDbUser gstrDbUser
    OraDataOpen = True
    
    
    '打开柯达 数据库连接
    '暂时使用HIS模拟
    With gcnPACS
        If .State = adStateOpen Then .Close
       ' .Provider = "MSDataShape"
       ' .Open "Driver={Microsoft ODBC for Oracle};Server=" & gstrPACSsid, gstrPACSUser, gstrPACSPassw
       'odbc
        .Open "Driver={SYBASE ASE ODBC Driver};NA=" & gstrPACSIP & "," & gstrPACSport & ";Uid=" & gstrPACSUser & ";Pwd=" & gstrPACSPassw & ";"
        'ole
       '.Open "Provider=Sybase.ASEOLEDBProvider.2;" & _
              "Server Name=" & gstrPACSIP & ";" & _
              "Server Port Address=" & gstrPACSport & ";" & _
              "Initial Catalog=" & gstrPACSsid & ";" & _
              "User ID=" & gstrPACSUser & ";" & _
              "Password=" & gstrPACSPassw & ";"
        If Err <> 0 Then
            '保存错误信息
            strError = Err.Description
            If InStr(strError, "自动化错误") > 0 Then
                MsgBox "连接串无法创建，请检查数据访问部件是否正常安装。", vbInformation, gstrSysName
            ElseIf InStr(strError, "ORA-12154") > 0 Then
                MsgBox "无法分析服务器名，" & vbCrLf & "请检查在Oracle配置中是否存在该本地网络服务名（主机字符串）。", vbInformation, gstrSysName
            ElseIf InStr(strError, "ORA-12541") > 0 Then
                MsgBox "无法连接，请检查服务器上的Oracle监听器服务是否启动。", vbInformation, gstrSysName
            ElseIf InStr(strError, "ORA-01033") > 0 Then
                MsgBox "ORACLE正在初始化或在关闭，请稍候再试。", vbInformation, gstrSysName
            Else
                MsgBox "由于用户、口令或服务器指定错误，无法注册。 " + Err.Description, vbInformation, gstrSysName
            End If
            
            OraDataOpen = False
            Exit Function
        End If
    End With
    
    '柯达使用的是SYBASE的数据库，需要通过ADODC部件产生连接字，然后替换到以下程序的“Open”中
'    With gcnPACS
'        If .State = adStateOpen Then .Close
'        .Open "Driver={Microsoft ODBC for Oracle};Server=" & gstrPACSsid, gstrPACSUser, gstrPACSPassw
'        If Err <> 0 Then
'            '保存错误信息
'            strError = Err.Description
'            MsgBox "连接SQL Server错误"
'            OraDataOpen = False
'            Exit Function
'        End If
'    End With


       
'这个是SQLServer2000数据库连接的例子，可以不用管
'    With gcnSQL2K
'        If .State = adStateOpen Then .Close
'
'        .Open "Provider=SQLOLEDB.1;Persist Security Info=False;Initial Catalog=yygl;Data Source=" & gstrHISIP, gstrUser, gstrPassw
'        If Err <> 0 Then
'            MsgBox "连接SQL Server错误"
'        End If
'    End With
    
    
    
    Exit Function
    
errHand:
    If ErrCenter() = 1 Then Resume
    OraDataOpen = False
    Err = 0
End Function


Public Sub Main()
    '用户注册
    gstrRegPath = "HIStoKodakPacs"
    
    '从注册表读取基本数据库连接参数
    gstrHISUser = GetSetting("ZLSOFT", gstrRegPath, "HIS用户名", "zlhis")
    gstrHISPassw = GetSetting("ZLSOFT", gstrRegPath, "HIS密码", "his")
    gstrHISsid = GetSetting("ZLSOFT", gstrRegPath, "HISsid", "")
    
    gstrPACSIP = GetSetting("ZLSOFT", gstrRegPath, "PACSIP地址", "172.16.9.13")
    gstrPACSUser = GetSetting("ZLSOFT", gstrRegPath, "PACS用户名", "zlhis")
    gstrPACSPassw = GetSetting("ZLSOFT", gstrRegPath, "PACS密码", "his123")
    gstrPACSsid = GetSetting("ZLSOFT", gstrRegPath, "PACSsid", "ris")
    gstrPACSport = GetSetting("ZLSOFT", gstrRegPath, "gstrPACSport", "4100")
    
    glngInterval = GetSetting("ZLSOFT", gstrRegPath, "监听间隔", "5")
    
    '基本数据库连接参数写入注册表
    SaveSetting "ZLSOFT", gstrRegPath, "HIS用户名", gstrHISUser
    SaveSetting "ZLSOFT", gstrRegPath, "HIS密码", gstrHISPassw
    SaveSetting "ZLSOFT", gstrRegPath, "HISsid", gstrHISsid
   
    
    SaveSetting "ZLSOFT", gstrRegPath, "监听间隔", glngInterval
    
    SaveSetting "ZLSOFT", gstrRegPath, "PACSIP地址", gstrPACSIP
    SaveSetting "ZLSOFT", gstrRegPath, "PACS用户名", gstrPACSUser
    SaveSetting "ZLSOFT", gstrRegPath, "PACS密码", gstrPACSPassw
    SaveSetting "ZLSOFT", gstrRegPath, "PACSsid", gstrPACSsid
    SaveSetting "ZLSOFT", gstrRegPath, "gstrPACSport", gstrPACSport
    
    OraDataOpen
    
    
    If gcnHIS.State <> adStateOpen Or gcnPACS.State <> adStateOpen Then
        Exit Sub
    End If
    
    frmSendOrder.Show 1
End Sub
