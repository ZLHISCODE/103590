Attribute VB_Name = "mdlPublic"
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

Public gobjFile As New FileSystemObject
Public gstrFilePath As String
Public gcnOracle As adodb.Connection    '公共数据库连接
Public gstrDBUser As String
Public gblnOwner As Boolean
Public gcolSort As Collection
Public gfrmFind As New frmFind
Public gblnIsRac As Boolean
Public gintInstId As Integer
Public gblnZlhis As Boolean
Public gstrCompareExe As String
Public gstrLeft As String
Public gstrSysName As String                '系统名称
Public gstrUserName As String               '用户名
Public gstrPassword As String               '用户口令
Public gstrToolsPwd As String               '管理工具的密码
Public gstrServer As String                 '服务器名
Public gstrSQL    As String                 '通用的SQL语句变量
Public gblnDBA As Boolean                   '是否DBA
Public gdtStart As Long
Public gblnOK As Boolean
Public glngSessionID As Long

Public gblnHasZltables As Boolean '记录是否有zltable这张表

'********************************************************************
'CommandBar命令ID
Public Enum CommandBarIDCond
    conMenu_FilePopup = 1
    conMenu_EditPopup = 2
    conMenu_ViewPopup = 8
    conMenu_HelpPopup = 9
    
    '添加一个对比功能设置
    conMenu_ComparePopup = 3
    '文件菜单
    conMenu_File_Open = 101
    conMenu_File_CompareExe = 210
    conmenu_File_Logout = 108
    conMenu_File_Exit = 109
    
    '编辑菜单
    conMenu_Edit_Trace = 201
    conMenu_Edit_Trace_1 = 2011
    conMenu_Edit_Trace_4 = 2012
    conMenu_Edit_Trace_8 = 2013
    conMenu_Edit_Trace_12 = 2014
    conMenu_Edit_ChangeReg = 2015
    conMenu_Edit_TraceOff = 202
    conMenu_Edit_CompareLeft = 211
    conMenu_Edit_Compare = 212
    
    '查看菜单
    conMenu_View_Style = 801
    conMenu_View_Style_Report = 8011
    conMenu_View_Style_Table = 8012
    conMenu_View_Filter = 802
    conMenu_View_SQLPrev = 803
    conMenu_View_SQLNext = 804
    conMenu_View_Find = 805
    conMenu_View_FindNext = 806
    conMenu_View_Refresh = 809
    conMenu_View_Close = 810
    
    '帮助菜单
    conMenu_Help_About = 901
End Enum

'CommandBar固有常量定义
Public Const XTP_ID_WINDOW_LIST = 35000 '窗体列表
Public Const XTP_ID_TOOLBARLIST = 59392 '工具栏列表
Public Const ID_INDICATOR_CAPS = 59137 '状态栏（大写）
Public Const ID_INDICATOR_NUM = 59138 '状态栏（数字）
Public Const ID_INDICATOR_SCRL = 59139 '状态栏（滚动）

'CommandBar辅助热键
Public Const FSHIFT = 4
Public Const FCONTROL = 8
Public Const FALT = 16
'********************************************************************
Public Const CB_SETDROPPEDWIDTH As Long = &H160
Public Declare Sub InitCommonControls Lib "comctl32.dll" ()
Public Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function ShellAbout Lib "shell32.dll" Alias "ShellAboutA" (ByVal hWnd As Long, ByVal szApp As String, ByVal szOtherStuff As String, ByVal hIcon As Long) As Long

'-------------------------------------------------------------
Public Const Process_Query_Information = &H400
Public Const Still_Active = &H103
Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Public Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

'-------------------------------------------------------------
Public Const GWL_EXSTYLE = (-20)
Public Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
'-------------------------------------------------------------
Public Const EM_LINESCROLL = &HB6 'lngW=横向行数,lngL=纵向行数
Public Const EM_SCROLL = &HB5 '按滚动条几下
Public Const EM_GETFIRSTVISIBLELINE = &HCE 'lngR(>=0)
Public Const EM_GETLINECOUNT = &HBA 'lngR(>=1,包含自动折的行)
Public Const EM_LINELENGTH = &HC1 '第一行未折行前有效
Public Const EM_GETSEL = &HB0
Public Const EM_LINEFROMCHAR = &HC9
Public Const EM_LINEINDEX = &HBB
Public Const EM_SETSEL = &HB1

Public Const FR_DOWN = &H1
Public Const FR_WHOLEWORD = &H2
Public Const FR_MATCHCASE = &H4
Public Type CHARRANGE
    cpMin As Long
    cpMax As Long
End Type
Public Type FINDTEXT
    chrg As CHARRANGE
    lpstrText As String
End Type

Public Const WM_USER = &H400
Public Const EM_EXGETSEL = (WM_USER + 52)
Public Const EM_EXSETSEL = (WM_USER + 55)
Public Const EM_FINDTEXT = (WM_USER + 56)
Public Const EM_SETTARGETDEVICE = (WM_USER + 72)
'-------------------------------------------------------------
' Reg Data Types...
Const REG_SZ = 1                         ' Unicode空终结字符串
Const REG_EXPAND_SZ = 2                  ' Unicode空终结字符串
Const REG_DWORD = 4                      ' 32-bit 数字

' 注册表关键字安全选项...
Public Const READ_CONTROL = &H20000
Public Const KEY_QUERY_VALUE = &H1
Public Const KEY_SET_VALUE = &H2
Public Const KEY_CREATE_SUB_KEY = &H4
Public Const KEY_ENUMERATE_SUB_KEYS = &H8
Public Const KEY_NOTIFY = &H10
Public Const KEY_CREATE_LINK = &H20
Public Const KEY_ALL_ACCESS = KEY_QUERY_VALUE + KEY_SET_VALUE + KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS + KEY_NOTIFY + KEY_CREATE_LINK + READ_CONTROL
                     
' 注册表关键字根类型...
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_USERS = &H80000003

' 返回值...
Public Const ERROR_SUCCESS = 0
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long

Public Declare Function ClientToScreen Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Public Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Public Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Public Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Public Declare Function GetCurrentProcess Lib "kernel32" () As Long
Public Declare Function IsWow64Process Lib "kernel32" (ByVal hProc As Long, bWow64Process As Boolean) As Long


Public Sub Main()
    Dim strTmp As String
    Dim strServerName As String, strUserName As String, strUserPwd As String
    Dim intUserPosition As Integer, intPwdPosition As Integer, intServerPosition As Integer
    
    Call InitCommonControls
    
    gblnOwner = False
    strTmp = Command
    
    If strTmp = "" Then
        '用户注册
        frmUserLogin.Show 1
        If gcnOracle Is Nothing Then
            Set gcnOracle = New adodb.Connection
        End If
    Else
        '在管理工具中，通过Command命令进行登录
        intUserPosition = InStr(1, strTmp, "zlUserName=") + Len("zlUserName=")
        intPwdPosition = InStr(1, strTmp, "zlPassword=") + Len("zlPassword=")
        intServerPosition = InStr(1, strTmp, "zlServer=") + Len("zlServer=")
        
        strUserName = Mid(Left(strTmp, InStr(1, strTmp, "zlPassword=") - 1), intUserPosition)
        strUserPwd = Mid(Left(strTmp, InStr(1, strTmp, "zlServer=") - 1), intPwdPosition)
        strServerName = Mid(strTmp, intServerPosition)
        gstrDBUser = UCase(strUserName)
        
        If Not OraDataOpen(strServerName, strUserName, strUserPwd) Then
            Exit Sub
        End If
    End If

    frmMain.Show
End Sub

Public Function OraDataOpen(ByVal strServerName As String, ByVal strUserName As String, ByVal strUserPwd As String, Optional bln转换 As Boolean) As Boolean
    Dim rstmp As adodb.Recordset
    Dim strSql As String, i As Integer
    
    On Error Resume Next
    
    If gcnOracle Is Nothing Then
        Set gcnOracle = New adodb.Connection
    End If
    If gcnOracle.State = adStateOpen Then gcnOracle.Close
    With gcnOracle
        .CursorLocation = adUseClient
        .Provider = "MSDataShape"
        .Open "Driver={Microsoft ODBC for Oracle};Server=" & strServerName, strUserName, strUserPwd
    End With
    If Err <> 0 Then
        MsgBox "连接失败！（请确保用户名与密码输入无误）", vbInformation, App.Title
        Err.Clear: Exit Function
    End If

    With rstmp
        strSql = "Select 1 From User_Role_Privs Where Granted_Role = 'DBA'"
        If .State = adStateOpen Then .Close
        .Open strSql, gcnOracle, adOpenKeyset
        gblnDBA = Not (.EOF Or .BOF)
    End With

    '功能：检查是否为RAC环境
    Err.Clear
    strSql = "Select 1 from gv$active_instances"
    Set rstmp = OpenSQLRecord(strSql, "CheckRAC")
    gblnIsRac = rstmp.RecordCount > 0
    
    If gblnIsRac Then
        strSql = "Select UserENV('instance') Inst_ID From dual"
        Set rstmp = OpenSQLRecord(strSql, "CheckRAC")
        gintInstId = Val("" & rstmp!INST_ID)
    End If
    
    If Err.Number > 0 Then Exit Function
    
    gstrUserName = strUserName: gstrPassword = strUserPwd: gstrServer = strServerName
    OraDataOpen = True
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

Public Function Decode(ParamArray arrPar() As Variant) As Variant
'功能：模拟Oracle的Decode函数
    Dim varValue As Variant, i As Integer
    
    i = 1
    varValue = arrPar(0)
    Do While i <= UBound(arrPar)
        If i = UBound(arrPar) Then
            Decode = arrPar(i): Exit Function
        ElseIf varValue = arrPar(i) Then
            Decode = arrPar(i + 1): Exit Function
        Else
            i = i + 2
        End If
    Loop
End Function

Public Function Nvl(ByVal varValue As Variant, Optional DefaultValue As Variant = "") As Variant
'功能：相当于Oracle的NVL，将Null值改成另外一个预设值
    Nvl = IIf(IsNull(varValue), DefaultValue, varValue)
End Function

Public Function GetShortName(ByVal strFile As String) As String
    Dim strShort As String, lngLen As Long
    
    GetShortName = strFile
    
    If InStr(strFile, " ") > 0 Then
        If gobjFile.FileExists(strFile) Then
            GetShortName = gobjFile.GetFile(strFile).ShortPath
        ElseIf gobjFile.FolderExists(strFile) Then
            GetShortName = gobjFile.GetFolder(strFile).ShortPath
        Else
            strShort = Space(255)
            lngLen = GetShortPathName(strFile, strShort, 255)
            GetShortName = Left(strShort, lngLen)
        End If
    End If
End Function

Public Sub CboAppendText(cboControl As Object, KeyAscii As Integer)
'功能：对ComboBox实现输入过程中自动完成的功能
'说明：在Combox.KeyPress事件中调用
    Dim strInput As String
    Dim lngIndex As Long
    Const CB_FINDSTRING = &H14C
    
    If cboControl.Style <> 0 Then Exit Sub
    If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyTab Then Exit Sub
    strInput = Chr(KeyAscii): KeyAscii = 0

    With cboControl
        '接着得到用户击键完成后文本框中出现的内容
        strInput = Mid(.Text, 1, .SelStart) & strInput

        '根据假想的内容得到可能的列表项
        lngIndex = SendMessage(cboControl.hWnd, CB_FINDSTRING, -1, ByVal strInput)
        If lngIndex >= 0 Then
            .ListIndex = lngIndex
            '.Text = .List(lngIndex)
            
            .SelStart = Len(strInput)
            .SelLength = Len(.Text) - Len(strInput)
        Else
            .Text = strInput
            .SelStart = Len(strInput)
        End If
    End With
End Sub


Public Function OpenSQLRecord(ByVal strSql As String, ByVal strTitle As String, ParamArray arrInput() As Variant) As adodb.Recordset
    Dim arrPars() As Variant, i As Long
    arrPars = arrInput
    Set OpenSQLRecord = OpenSQLRecordByArray(gcnOracle, strSql, strTitle, arrPars)
End Function

Public Function OpenSQLRecordByArray(ByVal cnOracle As adodb.Connection, ByVal strSql As String, ByVal strTitle As String, arrInput() As Variant) As adodb.Recordset
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
    Dim cmdData As New adodb.Command
    Dim strPar As String, arrPar As Variant
    Dim lngLeft As Long, lngRight As Long
    Dim strSeq As String, intMax As Integer, i As Integer
    Dim strLog As String, varValue As Variant
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
    'If cmdData.ActiveConnection Is Nothing Then
        Set cmdData.ActiveConnection = cnOracle '这句比较慢(这句执行1000次约0.5x秒)
    'End If

    cmdData.CommandText = strSql
    
    Set OpenSQLRecordByArray = cmdData.Execute
    Set OpenSQLRecordByArray.ActiveConnection = Nothing

End Function



Public Function CheckZlhis() As Boolean
    Dim strSql As String, rstmp As adodb.Recordset
    
    On Error GoTo errh
    
    strSql = "Select 1 From dba_tables Where table_name = 'ZLSYSTEMS'"
    Set rstmp = OpenSQLRecord(strSql, "CheckZlhis")
    
    CheckZlhis = rstmp.RecordCount > 0
    Exit Function
errh:
    MsgBox "获取ZLHIS参数失败。"
    CheckZlhis = False
End Function

Public Sub CheckSqlPlan(vsfPlanTbl As VSFlexGrid, ByVal intOptCol As Integer, ByVal intObjCol As Integer, _
                                            rsBigtbl As adodb.Recordset, rsBigIdx As adodb.Recordset, rsLowIdx As adodb.Recordset)
'功能:检查VSF表格中的执行计划
'         1.大表全表扫描zltables+zlbigtable+zlbaktables，
'         2.中型表全表扫描(如果有统计信息，User_tab_statistics:num_rows>3000(药品目录一般是这个值以上) AND num_rows<100 0000百万以内)
'         3.大表上引用基础表(非大表)的外键上的索引
'         4.大表和中型表索引全扫描（inex full scan，INDEX FAST FULL SCAN）
'         5.大表和中型表跳跃式索引扫描（INDEX SKIP SCAN）
'参数:
'vsfPlanTbl - 执行计划表格
'intOptCol - 操作列,如:Index full scan ,intObjCol - 操作涉及的对象列,如: 病人医嘱记录_IX_ID
'rsBigtbl,rsBigIdx,rsLowIdx -涉及的表/索引
    
    Dim strOperation As String, strObject As String
    Dim strTmp() As String, i As Integer, j As Integer
    Dim blnTmp As Boolean
    
    On Error GoTo errh
    With vsfPlanTbl
        If .Redraw = flexRDNone Then Exit Sub
        
        '遍历表格,获取对象
        For i = .FixedRows To .Rows - .FixedRows
            If intOptCol <> intObjCol Then
                '执行计划的操作和对象不在一列中,直接获取
                strOperation = TrimEx(.TextMatrix(i, intOptCol))
                strObject = TrimEx(.TextMatrix(i, intObjCol))
            Else
                '涉及情况:TABLE ACCESS FULL/INDEX FAST FULL SCAN/INDEX FULL SCAN/INDEX SKIP SCAN/INDEX RANGE SCAN
                strTmp = Split("TABLE ACCESS FULL/INDEX FULL SCAN/INDEX SKIP SCAN/INDEX RANGE SCAN/INDEX FAST FULL SCAN", "/")
                
                For j = 0 To UBound(strTmp)
                    If InStr(1, TrimEx(.TextMatrix(i, intOptCol)), strTmp(j)) > 0 Then
                        strOperation = strTmp(j)
                        strObject = Split(Trim(Replace(TrimEx(.TextMatrix(i, intOptCol)), strTmp(j), "")), " ")(0)
                        Exit For
                    End If
                Next
            End If
            
            If strOperation <> "" And strObject <> "" Then
                If strOperation = "TABLE ACCESS FULL" Then '获取全表扫描
                    blnTmp = CheckRs(rsBigtbl, "表名 = '" & strObject & "'") Or gcnOracle = ""
                ElseIf InStr(1, "INDEX FULL SCAN/INDEX SKIP SCAN/INDEX FAST FULL SCAN", strOperation) > 0 Then '索引全扫描\索引跳扫描
                    blnTmp = CheckRs(rsBigIdx, "索引名 = '" & strObject & "'") Or gcnOracle = ""
                ElseIf strOperation = "INDEX RANGE SCAN" And gcnOracle <> "" Then '索引范围扫描:低效索引
                    blnTmp = CheckRs(rsLowIdx, "约束名= '" & GetFkByIdx(strObject) & "'")
                End If
            End If
                
            If blnTmp Then .Cell(flexcpBackColor, i, 0, i, .Cols - 1) = &HF0F0FF
            strOperation = "": strObject = ""
            blnTmp = False
        Next

    End With
    Exit Sub
errh:
    MsgBox Err.Description
    If 0 = 1 Then
        Resume
    End If
End Sub


Public Sub GetMidTabSize(ByRef lngMinSize As Long, ByRef lngMaxSize As Long)
    '功能:获取中型表大小
    
    Dim strSql As String, rstmp As adodb.Recordset
    
    lngMinSize = 3000: lngMaxSize = 1000000
    
    On Error GoTo errh
    strSql = "Select A.参数名,Nvl(A.参数值,A.缺省值) As 参数值 " & _
                 "From zlParameters A " & _
                 "Where A.参数名 = '检查中型表' And a.系统 is null And a.模块 is null"
    Set rstmp = OpenSQLRecord(strSql, "GetMidTabSize")
    
    If rstmp.EOF Then Exit Sub
    lngMinSize = Split(rstmp!参数值, ",")(0)
    lngMaxSize = Split(rstmp!参数值, ",")(1)
    
    Exit Sub
errh:
    MsgBox Err.Description
    If 0 = 1 Then
        Resume
    End If
End Sub

Public Function GetCheckObj(ByVal intMod As Integer, Optional ByVal lngMinSize As Long, Optional ByVal lngMaxSize As Long) As adodb.Recordset
'功能:获取涉及性能问题的表/索引对象,返回一个记录集
'参数intMod: 1-表,2-索引,3-低效索引
'lngMinSize,lngMaxSize - 判定中型表的区间,不传则默认为3000-1000000

    Dim strSql As String
    
    On Error GoTo errh
    
    If gblnHasZltables Then
        strSql = "Union Select Distinct 表名 From Zltables Where 分类 In ('B1', 'B2', 'B3', 'C1', 'C2', 'C3')"
    Else
        strSql = "Union Select Distinct 表名 From Zlbigtables" & vbNewLine & _
                        "Union" & vbNewLine & _
                        "Select Distinct 表名 From zlBakTables"
    End If
    
    Select Case intMod
        Case 1
            strSql = "Select distinct  Table_Name 表名" & vbNewLine & _
                            "From Dba_Tab_Statistics" & vbNewLine & _
                            "Where Num_Rows Between " & IIf(lngMinSize = 0, 3000, lngMinSize) & " And " & IIf(lngMaxSize = 0, 1000000, lngMaxSize) & vbNewLine & _
                            strSql
                            
        Case 2
            strSql = "Select distinct Index_Name 索引名" & vbNewLine & _
                            "From Dba_Indexes" & vbNewLine & _
                            "Where Table_Name In" & vbNewLine & _
                            " ( Select Table_Name 表名 From Dba_Tab_Statistics Where Num_Rows Between " & IIf(lngMinSize = 0, 3000, lngMinSize) & " And " & IIf(lngMaxSize = 0, 1000000, lngMaxSize) & vbNewLine & _
                            strSql & ")"

        Case 3
            strSql = "Select distinct  a.Constraint_Name 约束名" & vbNewLine & _
                            "From Dba_Constraints A, Dba_Indexes B" & vbNewLine & _
                            "Where a.Constraint_Type = 'R' And b.uniqueness='UNIQUE' And a.r_Constraint_Name = b.Index_Name And a.r_Owner = b.Owner And" & vbNewLine & _
                            "      b.Table_Name Not In" & vbNewLine & _
                            "      (Select Distinct 表名 From Zlbigtables" & vbNewLine & _
                            "       Union Select Distinct 表名 From zlBakTables" & vbNewLine & _
                            IIf(gblnHasZltables, "Union Select Distinct 表名 From Zltables Where 分类 In ('B1', 'B2', 'B3', 'C1', 'C2', 'C3')", "") & vbNewLine & _
                            "       )"

    End Select
    
    Set GetCheckObj = OpenSQLRecord(strSql, "GetCheckObj")
    Exit Function
errh:
    Set GetCheckObj = Nothing
    MsgBox Err.Description
End Function


Public Function CheckRs(rsData As adodb.Recordset, ByVal strFilter As String) As Boolean
'功能:对传入的记录集添加过滤,如果有匹配项则返回True
    
    If rsData Is Nothing Then Exit Function
    rsData.Filter = strFilter
    CheckRs = Not rsData.EOF
    rsData.Filter = 0
End Function

Public Function TrimEx(ByVal strText As String, Optional ByVal blnCrlf As Boolean) As String
'功能：去掉TAB字符，两边空格，回车，最后只由单空格分隔。
'说明：主要是RunSQLFile的子函数
    Dim i As Long
    
    If blnCrlf Then
        strText = Replace(strText, vbCrLf, " ")
        strText = Replace(strText, vbCr, " ")
        strText = Replace(strText, vbLf, " ")
    End If
    strText = Trim(Replace(strText, vbTab, " "))
    
    i = 5
    Do While i > 1
        strText = Replace(strText, String(i, " "), " ")
        If InStr(strText, String(i, " ")) = 0 Then i = i - 1
    Loop
    TrimEx = strText
End Function

Public Function CheckTblExist(ByVal strTableName As String) As Boolean
    '功能：根据表名判断表是否存在
    '参数：strTableName - 要查询的表名
    Dim strSql As String, rsData As adodb.Recordset
    
    On Error GoTo errh
    strSql = "select 1 from dba_all_tables where table_name =[1] "
    Set rsData = OpenSQLRecord(strSql, "CheckTblExist", strTableName)
    CheckTblExist = (rsData.RecordCount > 0)
    
    Exit Function
errh:
    MsgBox Err.Description
End Function

Public Function GetFkByIdx(ByVal strIdxName As String) As String
'功能:根据传入的索引返回对应的外键约束名称
    
    Dim strSql As String, rsData As adodb.Recordset
    
    On Error GoTo errh:
    
    strSql = "Select Distinct a.Constraint_Name" & vbNewLine & _
                    "From Dba_Cons_Columns A, Dba_Ind_Columns B" & vbNewLine & _
                    "Where a.Table_Name = b.Table_Name And a.Column_Name = b.Column_Name And a.Position = b.Column_Position And" & vbNewLine & _
                    "      b.Index_Name = [1]"

    Set rsData = OpenSQLRecord(strSql, "GetFkByIdx", strIdxName)
        
    If Not rsData.EOF Then
        GetFkByIdx = rsData!Constraint_Name & ""
    End If
    Exit Function
errh:
    GetFkByIdx = ""
    MsgBox Err.Description
    If 0 = 1 Then
        Resume
    End If
End Function
