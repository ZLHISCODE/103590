Attribute VB_Name = "mdlMain"
Option Explicit
Public Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type
Public Type POINTAPI
        x As Long
        y As Long
End Type


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
Const KEY_ALL_ACCESS = KEY_QUERY_VALUE + KEY_SET_VALUE + _
                       KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS + _
                       KEY_NOTIFY + KEY_CREATE_LINK + READ_CONTROL
                     
' 注册表关键字根类型...
Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_USERS = &H80000003
Public Const HKEY_PERFORMANCE_DATA = &H80000004

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

Public Const WM_SYSCOMMAND = &H112
Public Const SC_MAXIMIZE = &HF030&
Public Const SC_RESTORE = &HF120&
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long

Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegEnumKey Lib "advapi32.dll" Alias "RegEnumKeyA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, ByVal cbName As Long) As Long

Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long
Private Declare Function RegCreateKeyEx Lib "advapi32" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, ByRef lpSecurityAttributes As SECURITY_ATTRIBUTES, ByRef phkResult As Long, ByRef lpdwDisposition As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long


Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function Htmlhelp Lib "hhctrl.ocx" Alias "HtmlHelpA" (ByVal hwndCaller As Long, ByVal pszFile As String, ByVal uCommand As Long, ByVal dwData As Any) As Long
Public Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Declare Function PathIsDirectory Lib "shlwapi.dll" Alias "PathIsDirectoryA" (ByVal pszPath As String) As Long
Public Declare Function ClientToScreen Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Public Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Public Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Public Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Public Declare Function GetCurrentProcess Lib "kernel32" () As Long
Public Declare Function IsWow64Process Lib "kernel32" (ByVal hProc As Long, bWow64Process As Boolean) As Long

Public gobjFile         As New FileSystemObject

Public gcnOracle As New ADODB.Connection    '公共数据库连接

Public gstrSysName As String                '系统名称
Public gstrUserName As String               '用户名
Public gstrPassword As String               '用户口令
Public gstrToolsPwd As String               '管理工具的密码
Public gstrServer As String                 '服务器名
Public gstrSQL    As String                 '通用的SQL语句变量
Public gblnDBA As Boolean                   '是否DBA
Public gblnOwner As Boolean                 '是否所有者
Public gdtStart As Long
Public gstrDBUser As String
Public gblnOK As Boolean
Public glngSessionID As Long

Public Function ShowHelp(SHwnd As Long, ByVal htmName As String) As Boolean
    '显示帮助窗体
    'SHwnd:传入窗口句柄(作为宿主窗口)
    'htmName:射映在CHM中的htm文件名称

    Dim path As String
    Dim strSave As String
    On Error GoTo ShowHelpErr
    
    ShowHelp = False
    strSave = String(200, Chr$(0))
    path = Left$(strSave, GetWindowsDirectory(strSave, Len(strSave))) + "\help\"
    If CBool(PathIsDirectory(path)) = False Then GoTo ShowHelpErr
    strSave = "zlSDK.CHM"
    path = Trim(path & strSave)
    If Trim(Dir(path)) = "" Then GoTo ShowHelpErr
    Call Htmlhelp(SHwnd, path, &H0, htmName & ".htm")
    ShowHelp = True
    Exit Function
ShowHelpErr:
    Err.Clear
End Function

Public Sub Main()
    frmUserLogin.Show 1

    If gcnOracle.State = adStateOpen Then
        frmMidMain.Show
        'Form1.Show
    End If
End Sub

Public Function OraDataOpen(ByVal strServerName As String, ByVal strUserName As String, ByVal strUserPwd As String) As Boolean
    '------------------------------------------------
    '功能： 打开指定的数据库
    '参数：
    '   strServerName：主机字符串
    '   strUserName：用户名
    '   strUserPwd：密码
    '返回： 数据库打开成功，返回true；失败，返回false
    '------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Dim strSql As String
    Dim strError As String
    
    On Error Resume Next
    If gcnOracle.State = adStateOpen Then gcnOracle.Close
    Set gcnOracle = Nothing
    With gcnOracle
        .CursorLocation = adUseClient
        .Provider = "OraOLEDB.Oracle.1;PLSQLRSet=1"
        .Open "PLSQLRSet=1;DistribTx=0;Persist Security Info=True;FetchSize=500;Data Source=" & strServerName, strUserName, strUserPwd

'        .Provider = "MSDataShape"
'        .Open "Driver={Microsoft ODBC for Oracle};BUFFERSIZE=65535;Server=" & strServerName, strUserName, strUserPwd
'        .CursorLocation = adUseClient
    End With
    If Err <> 0 Then
        '保存错误信息
        strError = Err.Description
        If InStr(strError, "自动化错误") > 0 Then
            strError = "无法创建连接对象，请检查数据访问部件(OraOLEDB.dll)是否正常安装并注册。"
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
        
        MsgBox strError, vbInformation, gstrSysName
        If gcnOracle.State = adStateClosed Then
          OraDataOpen = False
          Exit Function
        End If
    End If
    
    Err = 0
    With rsTemp
        strSql = "Select 1 From User_Role_Privs Where granted_role='DBA'"
        If .State = adStateOpen Then .Close
        .Open strSql, gcnOracle, adOpenKeyset
        gblnDBA = Not (.EOF Or .BOF)
    End With
    
    If Not gblnDBA Then
        OraDataOpen = False
        MsgBox "只有具有数据库DBA角色的用户才能使用本工具。", vbExclamation, gstrSysName
        Exit Function
    End If
    
    OraDataOpen = True
    gstrUserName = strUserName
    gstrPassword = strUserPwd
    gstrDBUser = UCase(strUserName)
    gstrServer = Trim(strServerName)
End Function

Public Sub SelAll(objTxt As Control)
'功能：对文本框的的文本选中
    If TypeName(objTxt) = "TextBox" Then
        If Trim(objTxt.Text) > 0 Then
            objTxt.SelStart = 0: objTxt.SelLength = Len(objTxt.Text)
        End If
    ElseIf TypeName(objTxt) = "MaskEdBox" Then
        If Not IsDate(objTxt.Text) Then
            objTxt.SelStart = 0: objTxt.SelLength = Len(objTxt.Text)
        Else
            objTxt.SelStart = 0: objTxt.SelLength = 10
        End If
    End If
End Sub

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



Public Function OpenSQLRecordByArray(ByVal strSql As String, ByVal strTitle As String, arrInput() As Variant) As ADODB.Recordset
'功能：通过Command对象打开带参数SQL的记录集
'参数：strSQL=条件中包含参数的SQL语句,参数形式为"[x]"
'             x>=1为自定义参数号,"[]"之间不能有空格
'             同一个参数可多处使用,程序自动换为ADO支持的"?"号形式
'             实际使用的参数号可不连续,但传入的参数值必须连续(如SQL组合时不一定要用到的参数)
'      arrInput=不定个数的参数值,按参数号顺序依次传入,必须是明确类型
'               因为使用绑定变量,对带"'"的字符参数,不需要使用"''"形式。
'      strTitle=用于SQLTest识别的调用窗体/模块标题
'返回：记录集，CursorLocation=adUseClient,LockType=adLockReadOnly,CursorType=adOpenStatic
'举例：
'SQL语句为="Select 姓名 From 病人信息 Where (病人ID=[3] Or 门诊号=[3] Or 姓名 Like [4]) And 性别=[5] And 登记时间 Between [1] And [2] And 险类 IN([6],[7])"
'调用方式为：Set rsPati=OpenSQLRecord(strSQL, Me.Caption, CDate(Format(rsMove!转出日期,"yyyy-MM-dd")),dtp时间.Value, lng病人ID, "张%", "男", 20, 21)
    Dim cmdData As New ADODB.Command
    Dim strPar As String, arrPar As Variant
    Dim lngLeft As Long, lngRight As Long
    Dim strSeq As String, intMax As Integer, i As Integer
    Dim strLog As String, varValue As Variant
    Dim strSQLTmp As String, arrstr As Variant
    Dim strTmp As String, strSQLtmp1 As String

    '检查如果使用了动态内存表，并且没有使用/*+ XXX*/等提示字时自动加上
    strSQLTmp = Trim(UCase(strSql))
    If Mid(Trim(Mid(strSQLTmp, 7)), 1, 2) <> "/*" And Mid(strSQLTmp, 1, 6) = "SELECT" Then
        arrstr = Split("F_STR2LIST,F_NUM2LIST,F_NUM2LIST2,F_STR2LIST2", ",")
        For i = 0 To UBound(arrstr)
            strSQLtmp1 = strSQLTmp
            Do While InStr(strSQLtmp1, arrstr(i)) > 0
                '判断前面是否用了IN 用了则不加Rule
                '先找到最近一个SELECT
                strTmp = Mid(strSQLtmp1, 1, InStr(strSQLtmp1, arrstr(i)) - 1)
                'strTmp = Replace(TrimEx(Mid(strTmp, 1, InStrRev(strTmp, "SELECT") - 1)), " ", "")
                If Len(strTmp) > 1 Then strTmp = Mid(strTmp, Len(strTmp) - 2)  '取后面3个字符
                
                If strTmp = "IN(" Then '属于in(select这种情况，则继续循环，看是否存在没有使用这种写法的其他动态内存函数
                   strSQLtmp1 = Mid(strSQLtmp1, InStr(strSQLtmp1, arrstr(i)) + Len(arrstr(i)))
                Else
                    Exit For
                End If
            Loop
        Next
        If i <= UBound(arrstr) Then
            strSql = "Select /*+ RULE*/" & Mid(Trim(strSql), 7)
        End If
    End If
    
    
    '分析自定的[x]参数
    lngLeft = InStr(1, strSql, "[")
    Do While lngLeft > 0
        lngRight = InStr(lngLeft + 1, strSql, "]")
        
        '可能是正常的"[编码]名称"
        strSeq = Mid(strSql, lngLeft + 1, lngRight - lngLeft - 1)
        If IsNumeric(strSeq) Then
            i = CInt(strSeq)
            strPar = strPar & "," & i
            If i > intMax Then intMax = i
        End If
        
        lngLeft = InStr(lngRight + 1, strSql, "[")
    Loop
    
    If UBound(arrInput) + 1 < intMax Then
        Err.Raise 9527, strTitle, "SQL语句绑定变量不全，调用来源：" & strTitle
    End If

    '替换为"?"参数
    strLog = strSql
    For i = 1 To intMax
        strSql = Replace(strSql, "[" & i & "]", "?")
        
        '产生用于SQL跟踪的语句
        varValue = arrInput(i - 1)
        Select Case TypeName(varValue)
        Case "Byte", "Integer", "Long", "Single", "Double", "Currency", "Decimal" '数字
            strLog = Replace(strLog, "[" & i & "]", varValue)
        Case "String" '字符
            strLog = Replace(strLog, "[" & i & "]", "'" & Replace(varValue, "'", "''") & "'")
        Case "Date" '日期
            strLog = Replace(strLog, "[" & i & "]", "To_Date('" & Format(varValue, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')")
        End Select
    Next

    '清除原有参数:不然不能重复执行
'    cmdData.CommandText = "" '不为空有时清除参数出错
'    Do While cmdData.Parameters.Count > 0
'        cmdData.Parameters.Delete 0
'    Loop
    
    '创建新的参数
    lngLeft = 0: lngRight = 0
    arrPar = Split(Mid(strPar, 2), ",")
    For i = 0 To UBound(arrPar)
        varValue = arrInput((arrPar(i) - 1))
        Select Case TypeName(varValue)
        Case "Byte", "Integer", "Long", "Single", "Double", "Currency", "Decimal" '数字
            cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i, adVarNumeric, adParamInput, 30, varValue)
        Case "String" '字符
            intMax = LenB(StrConv(varValue, vbFromUnicode))
            If intMax <= 2000 Then
                intMax = IIf(intMax <= 200, 200, 2000)
                cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i, adVarChar, adParamInput, intMax, varValue)
            Else
                If intMax < 4000 Then intMax = 4000
                cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i, adLongVarChar, adParamInput, intMax, varValue)
            End If
        Case "Date" '日期
            cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i, adDBTimeStamp, adParamInput, , varValue)
        Case "Variant()" '数组
            '这种方式可用于一些IN子句或Union语句
            '表示同一个参数的多个值,参数号不可与其它数组的参数号交叉,且要保证数组的值个数够用
            If arrPar(i) <> lngRight Then lngLeft = 0
            lngRight = arrPar(i)
            Select Case TypeName(varValue(lngLeft))
            Case "Byte", "Integer", "Long", "Single", "Double", "Currency" '数字
                cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i & "_" & lngLeft, adVarNumeric, adParamInput, 30, varValue(lngLeft))
                strLog = Replace(strLog, "[" & lngRight & "]", varValue(lngLeft), 1, 1)
            Case "String" '字符
                intMax = LenB(StrConv(varValue(lngLeft), vbFromUnicode))
                If intMax <= 2000 Then
                    intMax = IIf(intMax <= 200, 200, 2000)
                    cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i & "_" & lngLeft, adVarChar, adParamInput, intMax, varValue(lngLeft))
                Else
                    If intMax < 4000 Then intMax = 4000
                    cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i & "_" & lngLeft, adLongVarChar, adParamInput, intMax, varValue(lngLeft))
                End If
                
                strLog = Replace(strLog, "[" & lngRight & "]", "'" & Replace(varValue(lngLeft), "'", "''") & "'", 1, 1)
            Case "Date" '日期
                cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i & "_" & lngLeft, adDBTimeStamp, adParamInput, , varValue(lngLeft))
                strLog = Replace(strLog, "[" & lngRight & "]", "To_Date('" & Format(varValue(lngLeft), "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')", 1, 1)
            End Select
            lngLeft = lngLeft + 1 '该参数在数组中用到第几个值了
        End Select
    Next

    '执行返回记录集
    Set cmdData.ActiveConnection = gcnOracle '这句比较慢(这句执行1000次约0.5x秒)
 
    cmdData.CommandText = strSql
    
    
    Set OpenSQLRecordByArray = cmdData.Execute
    Set OpenSQLRecordByArray.ActiveConnection = Nothing
    
End Function


Public Function OpenSQLRecord(ByVal strSql As String, ByVal strTitle As String, ParamArray arrInput() As Variant) As ADODB.Recordset
'功能：通过Command对象打开带参数SQL的记录集
'参数：strSQL=条件中包含参数的SQL语句,参数形式为"[x]"
'             x>=1为自定义参数号,"[]"之间不能有空格
'             同一个参数可多处使用,程序自动换为ADO支持的"?"号形式
'             实际使用的参数号可不连续,但传入的参数值必须连续(如SQL组合时不一定要用到的参数)
'      arrInput=不定个数的参数值,按参数号顺序依次传入,必须是明确类型
'               因为使用绑定变量,对带"'"的字符参数,不需要使用"''"形式。
'      strTitle=用于SQLTest识别的调用窗体/模块标题
'返回：记录集，CursorLocation=adUseClient,LockType=adLockReadOnly,CursorType=adOpenStatic
'举例：
'SQL语句为="Select 姓名 From 病人信息 Where (病人ID=[3] Or 门诊号=[3] Or 姓名 Like [4]) And 性别=[5] And 登记时间 Between [1] And [2] And 险类 IN([6],[7])"
'调用方式为：Set rsPati=OpenSQLRecord(strSQL, Me.Caption, CDate(Format(rsMove!转出日期,"yyyy-MM-dd")),dtp时间.Value, lng病人ID, "张%", "男", 20, 21)
    Dim cmdData As New ADODB.Command
    Dim strPar As String, arrPar As Variant
    Dim lngLeft As Long, lngRight As Long
    Dim strSeq As String, intMax As Integer, i As Integer
    Dim strLog As String, varValue As Variant
    Dim strSQLTmp As String, arrstr As Variant
    Dim strTmp As String, strSQLtmp1 As String

    '检查如果使用了动态内存表，并且没有使用/*+ XXX*/等提示字时自动加上
    strSQLTmp = Trim(UCase(strSql))
    If Mid(Trim(Mid(strSQLTmp, 7)), 1, 2) <> "/*" And Mid(strSQLTmp, 1, 6) = "SELECT" Then
        arrstr = Split("F_STR2LIST,F_NUM2LIST,F_NUM2LIST2,F_STR2LIST2", ",")
        For i = 0 To UBound(arrstr)
            strSQLtmp1 = strSQLTmp
            Do While InStr(strSQLtmp1, arrstr(i)) > 0
                '判断前面是否用了IN 用了则不加Rule
                '先找到最近一个SELECT
                strTmp = Mid(strSQLtmp1, 1, InStr(strSQLtmp1, arrstr(i)) - 1)
                'strTmp = Replace(TrimEx(Mid(strTmp, 1, InStrRev(strTmp, "SELECT") - 1)), " ", "")
                If Len(strTmp) > 1 Then strTmp = Mid(strTmp, Len(strTmp) - 2)  '取后面3个字符
                
                If strTmp = "IN(" Then '属于in(select这种情况，则继续循环，看是否存在没有使用这种写法的其他动态内存函数
                   strSQLtmp1 = Mid(strSQLtmp1, InStr(strSQLtmp1, arrstr(i)) + Len(arrstr(i)))
                Else
                    Exit For
                End If
            Loop
        Next
        If i <= UBound(arrstr) Then
            strSql = "Select /*+ RULE*/" & Mid(Trim(strSql), 7)
        End If
    End If
    
    
    '分析自定的[x]参数
    lngLeft = InStr(1, strSql, "[")
    Do While lngLeft > 0
        lngRight = InStr(lngLeft + 1, strSql, "]")
        
        '可能是正常的"[编码]名称"
        strSeq = Mid(strSql, lngLeft + 1, lngRight - lngLeft - 1)
        If IsNumeric(strSeq) Then
            i = CInt(strSeq)
            strPar = strPar & "," & i
            If i > intMax Then intMax = i
        End If
        
        lngLeft = InStr(lngRight + 1, strSql, "[")
    Loop
    
    If UBound(arrInput) + 1 < intMax Then
        Err.Raise 9527, strTitle, "SQL语句绑定变量不全，调用来源：" & strTitle
    End If

    '替换为"?"参数
    strLog = strSql
    For i = 1 To intMax
        strSql = Replace(strSql, "[" & i & "]", "?")
        
        '产生用于SQL跟踪的语句
        varValue = arrInput(i - 1)
        Select Case TypeName(varValue)
        Case "Byte", "Integer", "Long", "Single", "Double", "Currency", "Decimal" '数字
            strLog = Replace(strLog, "[" & i & "]", varValue)
        Case "String" '字符
            strLog = Replace(strLog, "[" & i & "]", "'" & Replace(varValue, "'", "''") & "'")
        Case "Date" '日期
            strLog = Replace(strLog, "[" & i & "]", "To_Date('" & Format(varValue, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')")
        End Select
    Next

    '清除原有参数:不然不能重复执行
'    cmdData.CommandText = "" '不为空有时清除参数出错
'    Do While cmdData.Parameters.Count > 0
'        cmdData.Parameters.Delete 0
'    Loop
    
    '创建新的参数
    lngLeft = 0: lngRight = 0
    arrPar = Split(Mid(strPar, 2), ",")
    For i = 0 To UBound(arrPar)
        varValue = arrInput((arrPar(i) - 1))
        Select Case TypeName(varValue)
        Case "Byte", "Integer", "Long", "Single", "Double", "Currency", "Decimal" '数字
            cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i, adVarNumeric, adParamInput, 30, varValue)
        Case "String" '字符
            intMax = LenB(StrConv(varValue, vbFromUnicode))
            If intMax <= 2000 Then
                intMax = IIf(intMax <= 200, 200, 2000)
                cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i, adVarChar, adParamInput, intMax, varValue)
            Else
                If intMax < 4000 Then intMax = 4000
                cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i, adLongVarChar, adParamInput, intMax, varValue)
            End If
        Case "Date" '日期
            cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i, adDBTimeStamp, adParamInput, , varValue)
        Case "Variant()" '数组
            '这种方式可用于一些IN子句或Union语句
            '表示同一个参数的多个值,参数号不可与其它数组的参数号交叉,且要保证数组的值个数够用
            If arrPar(i) <> lngRight Then lngLeft = 0
            lngRight = arrPar(i)
            Select Case TypeName(varValue(lngLeft))
            Case "Byte", "Integer", "Long", "Single", "Double", "Currency" '数字
                cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i & "_" & lngLeft, adVarNumeric, adParamInput, 30, varValue(lngLeft))
                strLog = Replace(strLog, "[" & lngRight & "]", varValue(lngLeft), 1, 1)
            Case "String" '字符
                intMax = LenB(StrConv(varValue(lngLeft), vbFromUnicode))
                If intMax <= 2000 Then
                    intMax = IIf(intMax <= 200, 200, 2000)
                    cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i & "_" & lngLeft, adVarChar, adParamInput, intMax, varValue(lngLeft))
                Else
                    If intMax < 4000 Then intMax = 4000
                    cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i & "_" & lngLeft, adLongVarChar, adParamInput, intMax, varValue(lngLeft))
                End If
                
                strLog = Replace(strLog, "[" & lngRight & "]", "'" & Replace(varValue(lngLeft), "'", "''") & "'", 1, 1)
            Case "Date" '日期
                cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i & "_" & lngLeft, adDBTimeStamp, adParamInput, , varValue(lngLeft))
                strLog = Replace(strLog, "[" & lngRight & "]", "To_Date('" & Format(varValue(lngLeft), "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')", 1, 1)
            End Select
            lngLeft = lngLeft + 1 '该参数在数组中用到第几个值了
        End Select
    Next

    '执行返回记录集
    Set cmdData.ActiveConnection = gcnOracle '这句比较慢(这句执行1000次约0.5x秒)
 
    cmdData.CommandText = strSql
    
    
    Set OpenSQLRecord = cmdData.Execute
    Set OpenSQLRecord.ActiveConnection = Nothing
    
End Function

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

Public Function GetCurrentdate() As Date
    '-------------------------------------------------------------
    '功能：提取服务器上当前日期
    '参数：
    '返回：由于Oracle日期格式的问题，所以
    '-------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    
    Err = 0
    On Error GoTo errH
    With rsTemp
        .CursorLocation = adUseClient
        .Open "SELECT SYSDATE FROM DUAL", gcnOracle, adOpenKeyset
    End With
    GetCurrentdate = rsTemp.Fields(0).Value
    rsTemp.Close
    Exit Function
errH:
    GetCurrentdate = 0
    Err = 0
End Function


Public Function GetTimeString(ByVal datBegin As Date, ByVal datEnd As Date) As String
'功能：获取两个时间值差的格式字符串
'   datBegin=起始时间
'   datEnd=中止时间
    Dim intH As Integer, intM As Integer, intS As Integer
    Dim datTmp As Date

    intH = DateDiff("h", datBegin, datEnd)
    datTmp = DateAdd("h", intH, datBegin)
    intM = DateDiff("n", datTmp, datEnd)
    datTmp = DateAdd("n", intM, datTmp)
    intS = DateDiff("s", datTmp, datEnd)
    
    If intS < 0 Then
        intM = intM - 1
        intS = 60 + intS
    End If
    
    If intM < 0 Then
        intH = intH - 1
        intM = 60 + intM
    End If
    GetTimeString = IIf(intH <> 0, intH & "小时", "") & IIf(intM <> 0, intM & "分", "") & intS & "秒"
End Function



Public Sub SetNOParallel(ByVal cnThis As ADODB.Connection, ByVal bytType As Byte)
'功能：并行执行后会自动为表名索引加上并行度属性，如果不取消，会影响相关SQL的执行计划(全表扫描+并行查询，巨慢)
'参数：bytType：0=索引，1=表

    Dim strSql As String
    Dim rsTmp As ADODB.Recordset
    Dim cmdTmp As New ADODB.Command
        
    If bytType = 0 Then
        strSql = "Select Owner || '.' || Index_Name As Index_Name From DBA_Indexes Where Degree Not In ('1', '0')"
    Else
        strSql = "Select Owner || '.' || Table_Name As Table_Name From DBA_Tables Where Degree !=('         1')"
    End If
    Set rsTmp = New ADODB.Recordset
    rsTmp.Open strSql, cnThis, adOpenKeyset, adLockReadOnly
        
    Set cmdTmp.ActiveConnection = cnThis
    cmdTmp.CommandType = adCmdText
    
    '如果有索引组织表，会报错：ORA-25176: 不允许对主键使用存储说明
    On Error Resume Next
    
    While Not rsTmp.EOF
        If bytType = 0 Then
            strSql = "alter index " & rsTmp!Index_Name & " noparallel"
        Else
            strSql = "alter table " & rsTmp!table_name & " noparallel"
        End If
        cmdTmp.CommandText = strSql
        
        cmdTmp.Execute
        
        rsTmp.MoveNext
    Wend
End Sub


Public Sub InitTable(vsf As VSFlexGrid, strCol As String)
'功能: 初始化表头
    Dim arrHead As Variant
    Dim i As Long
    
    arrHead = Split(strCol, ";")
   
    With vsf
        .Redraw = flexRDNone
        .Clear
        .FixedRows = 1: .FixedCols = 0
        .Cols = UBound(arrHead) + 1
        .Rows = .FixedRows
        .Editable = flexEDNone
        .ExtendLastCol = True
        
        For i = 0 To UBound(arrHead)
            .TextMatrix(.FixedRows - 1, i) = Split(arrHead(i), ",")(0)
            .ColKey(i) = Split(arrHead(i), ",")(0)
            
            If UBound(Split(arrHead(i), ",")) > 0 Then
                .ColHidden(i) = False
                .ColWidth(i) = Val(Split(arrHead(i), ",")(1))
                .ColAlignment(i) = Val(Split(arrHead(i), ",")(2))
            Else
                .ColHidden(i) = True
                .ColWidth(i) = 0
            End If
        Next
        .Redraw = True
    End With
End Sub


Public Function OpenCursor(ByVal cnOwner As ADODB.Connection, _
                              ByVal strPackagesName As String, _
                              ParamArray varParValue() As Variant) As ADODB.Recordset
'-----------------------------------------
'功能：调用存储过程返回记录集
'入参：strPackagesName ，格式为 [所有者.]包.过程名
'-----------------------------------------
    Static cmdPackage As New ADODB.Command
    Dim parPackage As ADODB.Parameter
    Dim arrPar As Variant, i As Integer
    Dim varValue As Variant, intMax As Integer
    Dim intMaxArr As Integer  '记录参数个数
    Dim varOutPar As Variant
    On Error GoTo errHandle

    '清除原有参数:不然不能重复执行
   
    
    cmdPackage.CommandText = "" '不为空有时清除参数出错
    Do While cmdPackage.Parameters.Count > 0
        cmdPackage.Parameters.Delete 0
    Loop
    
    '------ IN 参数
    For i = 0 To UBound(varParValue)
        varValue = varParValue(i)
        Select Case TypeName(varValue)
            Case "Byte", "Integer", "Long", "Single", "Double", "Currency" '数字
                cmdPackage.Parameters.Append cmdPackage.CreateParameter("P" & i, adVarNumeric, adParamInput, 30, varValue)
            Case "String" '字符
                intMax = LenB(StrConv(varValue, vbFromUnicode))
                If intMax = 0 Or intMax < 10 Then intMax = 10
                cmdPackage.Parameters.Append cmdPackage.CreateParameter("P" & i, adVarChar, adParamInput, intMax, varValue)
            Case "Date" '日期
                cmdPackage.Parameters.Append cmdPackage.CreateParameter("P" & i, adDBTimeStamp, adParamInput, , varValue)
        End Select
    Next

    If cmdPackage.ActiveConnection Is Nothing Then
        If cnOwner Is Nothing Then
            Set cmdPackage.ActiveConnection = gcnOracle
        Else
            Set cmdPackage.ActiveConnection = cnOwner
        End If
    Else
        If Not cnOwner Is Nothing Then
            If cmdPackage.ActiveConnection.ConnectionString <> cnOwner.ConnectionString Then
                Set cmdPackage.ActiveConnection = cnOwner
            End If
        End If
    End If
    
    cmdPackage.CommandType = adCmdStoredProc
    cmdPackage.CommandText = strPackagesName
    cmdPackage.Properties("PLSQLRSet") = True
    Set OpenCursor = cmdPackage.Execute
    cmdPackage.Properties("PLSQLRSet") = False
    Exit Function
errHandle:
    If MsgBox(Err.Description, vbRetryCancel, gstrSysName) = vbRetry Then
        Resume
    End If

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


Public Sub ShowFlash(Optional strInfo As String, Optional sngPer As Single = -1, Optional frmParent As Object, Optional blnPer As Boolean)
'功能：显示或隐藏等待或进度窗体(strInfo)
'参数:strInfo=等待或进度提示信息
'     sngPer=进度
    Static blnShow As Boolean
    
    If sngPer > 1 Then sngPer = 1
    
    If strInfo = "" Then
        frmFlash.avi.Close
        Unload frmFlash
        blnShow = False
    Else
        If Not blnShow Then
            On Error Resume Next
            If sngPer = -1 Then
                '显示等待
                frmFlash.avi.Open GetSetting("ZLSOFT", "注册信息", "gstrAviPath", "") & "\" & "Findfile.avi"
                If Err.Number <> 0 Then
                    Err.Clear
                End If
                frmFlash.lbl.Caption = strInfo
                
                If frmParent Is Nothing Then
                    SetWindowPos frmFlash.hwnd, -1, (Screen.Width - frmFlash.Width) / 2 / 15, (Screen.Height - frmFlash.Height) / 2 / 15, 0, 0, 1
                    ShowWindow frmFlash.hwnd, 5
                Else
                    Err.Clear
                    frmFlash.Show , frmParent
                    If Err.Number <> 0 Then
                        Err.Clear
                        SetWindowPos frmFlash.hwnd, -1, (Screen.Width - frmFlash.Width) / 2 / 15, (Screen.Height - frmFlash.Height) / 2 / 15, 0, 0, 1
                        ShowWindow frmFlash.hwnd, 5
                    End If
                End If
                
                frmFlash.avi.Play
                frmFlash.Refresh
            Else
                '显示进度
                frmFlash.avi.Visible = False
                frmFlash.picDo.Visible = True
                frmFlash.lbl.Top = frmFlash.lbl.Top - frmFlash.lbl.Height / 2
                frmFlash.lbl.Left = frmFlash.picDo.Left
                frmFlash.lblPer.Top = frmFlash.lbl.Top
                frmFlash.lbl.Caption = strInfo
                frmFlash.lblDo.Caption = String(25 * sngPer, frmFlash.lblDo.Tag)
                If blnPer Then
                    If sngPer > 0 Then
                        frmFlash.lblPer.Caption = Int(sngPer * 100) & "%"
                    Else
                        frmFlash.lblPer.Caption = ""
                    End If
                    frmFlash.lblPer.Visible = True
                End If
                
                If frmParent Is Nothing Then
                    SetWindowPos frmFlash.hwnd, -1, (Screen.Width - frmFlash.Width) / 2 / 15, (Screen.Height - frmFlash.Height) / 2 / 15, 0, 0, 1
                    ShowWindow frmFlash.hwnd, 5
                Else
                    Err.Clear
                    frmFlash.Show , frmParent
                    If Err.Number <> 0 Then
                        Err.Clear
                        SetWindowPos frmFlash.hwnd, -1, (Screen.Width - frmFlash.Width) / 2 / 15, (Screen.Height - frmFlash.Height) / 2 / 15, 0, 0, 1
                        ShowWindow frmFlash.hwnd, 5
                    End If
                End If
                
                frmFlash.Refresh
            End If
            blnShow = True
        Else
            frmFlash.lbl.Caption = strInfo
            If sngPer >= 0 Then
                frmFlash.lblDo.Caption = String(25 * sngPer, frmFlash.lblDo.Tag)
                If sngPer > 0 Then
                    frmFlash.lblPer.Caption = Int(sngPer * 100) & "%"
                Else
                    frmFlash.lblPer.Caption = ""
                End If
            End If
            frmFlash.Refresh
        End If
    End If
End Sub

 Public Function Is64bit() As Boolean
    '******************************************************************************************************************
    '功能：是否是64位系统
    '返回：
    '******************************************************************************************************************
    Dim handle As Long
    Dim bolFunc As Boolean
        
    bolFunc = False
    handle = GetProcAddress(GetModuleHandle("kernel32"), "IsWow64Process")
    If handle > 0 Then
        IsWow64Process GetCurrentProcess(), bolFunc
    End If
    Is64bit = bolFunc
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


