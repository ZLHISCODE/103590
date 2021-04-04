Attribute VB_Name = "mdlMain"
Option Explicit

Public gobjPacsSrv As Object                '导航台
Public SplashObj As New frmSplash
Public gcnOracle As New ADODB.Connection    '公共数据库连接

Public gstrSysName As String                '系统名称
Public gstrVersion As String                '系统版本
Public gstrAviPath As String                'AVI文件的存放目录

Public gstrUserFlag As String               '当前用户标志(两位表示)，第1位：是否DBA；第2位：系统所有者

Public gstrDbUser As String                 '当前数据库用户
Public glngUserId As Long                   '当前用户id
Public gstrUserCode As String               '当前用户编码
Public gstrUserName As String               '当前用户姓名
Public gstrUserAbbr As String               '当前用户简码

Public glngDeptId As Long                   '当前用户部门id
Public gstrDeptCode As String               '当前用户部门编码
Public gstrDeptName As String               '当前用户部门名称

Public gstrStation As String                '本工作站名称
Public gstrMenuSys As String                '系统菜单

Public strTimePolicy As String              '时间策略  "time,month,29,23:12:00,0,1"
'Public strCapacityPolicy As String          '容量策略  "capacity,30,0,1"
'Public strPercentPolicy As String           '容量百分比策略  "percent,10,1,1"
Public beginDay As Date                     '进入系统的时间
Public bAutoArchive As String
Public Const cAllStorageDevice = 1111       '选择全部存储设备的标识

Public ZlCommFun As New zl9comlib.clsCommFun
Public gobjFile As New FileSystemObject

'-----------------------------------------
'发行码、注册码、发行码解析串、注册码解析串
Public gstrRegCode As String
Public gstrPublish As String
Public gstrParseRegCode As String
Public gstrParsePublish As String
'-----------------------------------------
Public gstrSystems As String

Private Declare Function InitCommonControls Lib "comctl32.dll" () As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Const SRCCOPY = &HCC0020
Public Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

'---------------------------------------------------------------
'-注册表 API 声明...
'---------------------------------------------------------------
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long
Private Declare Function RegCreateKeyEx Lib "advapi32" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, ByRef lpSecurityAttributes As SECURITY_ATTRIBUTES, ByRef phkResult As Long, ByRef lpdwDisposition As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long
Public Declare Function SetActiveWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
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


'---------------------------------------------------------------
'   授权、菜单、试用版本
'---------------------------------------------------------------
Public Sub Main()
    Dim lngReturn As Long
    Dim StrUnitName As String
    Dim BlnShowFlash As Boolean
    Dim strCode As String, IntCount As Integer, StrStyle As String
    Dim rsMenu As ADODB.Recordset, StrHaveSys As String
    
    '为实现XP风格，在显示窗体前必须执行该函数
    Call InitCommonControls

    
    BlnShowFlash = False
    Load SplashObj
    '由注册表中获取用户注册相关信息,如果用户单位名称不为空,则显示闪现窗体
    StrUnitName = GetSetting("ZLSOFT", "注册信息", "单位名称", "")
    gstrSysName = GetSetting("ZLSOFT", "注册信息", "提示", "")
    If StrUnitName <> "" Then
        With SplashObj
            '有两处需要处理
            Call ApplyOEM_Picture(.ImgIndicate, "Picture")
            Call ApplyOEM_Picture(.imgPic, "PictureB")
            .Show
            .lblGrant = StrUnitName
            StrUnitName = GetSetting("ZLSOFT", "注册信息", "开发商", "")
            If Trim(StrUnitName) = "" Then
                .Label3.Visible = False
                .lbl开发商.Visible = False
            Else
                .lbl开发商.Caption = ""
                For IntCount = 0 To UBound(Split(StrUnitName, ";"))
                    .lbl开发商.Caption = .lbl开发商.Caption & Split(StrUnitName, ";")(IntCount) & vbCrLf
                Next
            End If
            .LblProductName = GetSetting("ZLSOFT", "注册信息", "产品全称", "")
            .lbl技术支持商 = GetSetting("ZLSOFT", "注册信息", "技术支持商", "")
        End With
        
        BlnShowFlash = True
        DoEvents
    End If
    
    gstrStation = Space(200)
    lngReturn = GetComputerName(gstrStation, 200)
    gstrStation = Trim(gstrStation)
    If Len(gstrStation) > 1 Then
        gstrStation = Left(gstrStation, Len(gstrStation) - 1)
    Else
        gstrStation = "..."
    End If
    
    '用户注册
    frmUserLogin.Show 1
    If gcnOracle.State <> adStateOpen Then
        Unload frmUserLogin
        Unload SplashObj
        Exit Sub
    End If
    
    
    '初始化公共部件
    InitCommon gcnOracle
    If RegCheck = False Then
        Unload SplashObj
        Exit Sub
    End If
    
    '如果发行码无效（为空或为"-"），则退出
    gstrParsePublish = zlRegInfo("产品简名")
    gstrParseRegCode = zlRegInfo("单位名称", , -1)
    
    gstrSysName = gstrParsePublish & "软件"
    SaveSetting "ZLSOFT", "注册信息", "提示", gstrSysName
    SaveSetting "ZLSOFT", "注册信息", UCase("gstrSysName"), gstrSysName
    gstrVersion = App.Major & "." & App.Minor & "." & App.Revision
    SaveSetting "ZLSOFT", "注册信息", UCase("gstrVersion"), gstrVersion
    gstrAviPath = App.Path & "\附加文件"
    SaveSetting "ZLSOFT", "注册信息", UCase("gstrAviPath"), gstrAviPath
    
    With SplashObj
        If BlnShowFlash = False Then
            .lblGrant = gstrParseRegCode
            .lbl技术支持商.Caption = zlRegInfo("技术支持商", , -1)
            .LblProductName = zlRegInfo("产品标题")
            
            strCode = zlRegInfo("产品开发商", , -1)
            .lbl开发商.Caption = ""
            For IntCount = 0 To UBound(Split(strCode, ";"))
                .lbl开发商.Caption = .lbl开发商.Caption & Split(strCode, ";")(IntCount) & vbCrLf
            Next
            Call ApplyOEM_Picture(.ImgIndicate, "Picture")
            .Show
            BlnShowFlash = True
        End If
        DoEvents
    End With
    
    '将用户注册相关信息写入注册表,供下次启动时显示
    SaveSetting "ZLSOFT", "注册信息", "单位名称", gstrParseRegCode
    SaveSetting "ZLSOFT", "注册信息", "产品全称", zlRegInfo("产品标题")
    SaveSetting "ZLSOFT", "注册信息", "产品名称", zlRegInfo("产品简名")
    SaveSetting "ZLSOFT", "注册信息", "技术支持商", zlRegInfo("技术支持商", , -1)
    SaveSetting "ZLSOFT", "注册信息", "开发商", zlRegInfo("产品开发商", , -1)
    SaveSetting "ZLSOFT", "注册信息", "WEB支持商简名", zlRegInfo("支持商简名")
    SaveSetting "ZLSOFT", "注册信息", "WEB支持EMAIL", zlRegInfo("支持商MAIL")
    SaveSetting "ZLSOFT", "注册信息", "WEB支持URL", zlRegInfo("支持商URL")

    On Error Resume Next
    Err = 0
    Unload SplashObj
    frmMain.Show
End Sub

Public Function MenuGranted() As ADODB.Recordset
    '-------------------------------------------------------------
    '功能：分析授权使用并安装的部件，进而产生授权使用的菜单集合
    '参数：注册码
    '-------------------------------------------------------------
    Dim strSQL As String
    Dim rsTemp As New ADODB.Recordset
    Dim strCodes As String
    Dim strObjs As String
    Dim IntCount As Integer
    Dim gstrSystems1 As String
    Dim gstrMenuSys As String
    Dim BlnOnlySys As Boolean '只有报表系统
    
    strCodes = zlRegFunctions(gstrRegCode) & " Or 系统 Is NULL "
    BlnOnlySys = (gstrSystems = "")
    If BlnOnlySys Then
        gstrSystems = " 系统 Is NULL"
        gstrSystems1 = " G.系统 Is NULL"
        strCodes = "(" & strCodes & ") Or " & gstrSystems
    Else
        gstrSystems1 = Replace(gstrSystems, "系统", "G.系统")
        strCodes = "(" & strCodes & ") And " & gstrSystems
        
    End If
    
    '--分析权限菜单--
    With rsTemp
        gstrMenuSys = IIf(Command() = "", "缺省", Command())
        strObjs = "'ZL9DEMAND'"
        
        strSQL = "select Distinct M.层次,M.编号,nvl(M.上级,0) as 上级,M.标题,M.短标题,M.快键,M.说明,nvl(M.模块,0) as 模块,M.系统,nvl(M.图标,0) 图标,upper(rtrim(P.部件)) 部件,Decode(upper(rtrim(P.部件)),'ZL9REPORT',1,0) " & _
                " from (select level as 层次,ID 编号,上级ID 上级,标题,短标题,快键,说明,模块,Nvl(系统,0) 系统,图标" & _
                "         from zlMenus " & IIf(BlnOnlySys, "", " where " & gstrSystems) & _
                "         start with 组别='" & gstrMenuSys & "'" & " and 上级ID is null" & _
                "         connect by prior ID=上级ID and 组别='" & gstrMenuSys & "'" & ") M," & _
                "     (select distinct ID 编号" & _
                "         from zlMenus" & _
                "         start with 模块 in (" & _
                "             select 序号" & _
                "             from zlPrograms" & _
                "             where 序号 in (select distinct 序号 from zlProgFuncs where (" & strCodes & "))" & _
                "             and " & gstrSystems & _
                "                  and (exists(select 1 from session_roles where role='DBA')" & _
                "                       or 系统 in (select 编号 from zlSystems where upper(所有者)=user " & _
                "                                   Union" & _
                "                                   Select 0 From dual" & _
                "                                   Where Exists(select 1 from zlSystems where upper(所有者)=User))" & _
                "                       or 序号 in (select 序号 from zlRoleGrant G,session_roles S where G.角色=S.Role" & " And " & gstrSystems1 & ")) " & _
                "                  and upper(rtrim(部件)) in (" & strObjs & ")" & _
                "             and " & gstrSystems & _
                "       ) and 组别='" & gstrMenuSys & "' and " & gstrSystems & _
                "         connect by prior 上级ID=ID and 组别='" & gstrMenuSys & "' and " & gstrSystems & ") G,"
        strSQL = strSQL & "" & _
                "     (Select Nvl(系统,0) 系统,序号,部件 From zlPrograms " & _
                "             Where 序号 in (select distinct 序号 from zlProgFuncs where (" & strCodes & "))" & _
                "             and " & gstrSystems & _
                "                  and (exists(select 1 from session_roles where role='DBA')" & _
                "                       or 系统 in (select 编号 from zlSystems where upper(所有者)=user " & _
                "                                   Union" & _
                "                                   Select 0 From dual" & _
                "                                   Where Exists(select 1 from zlSystems where upper(所有者)=User))" & _
                "                       or 序号 in (select 序号 from zlRoleGrant G,session_roles S where G.角色=S.Role" & " And " & gstrSystems1 & ")) " & _
                "                  and upper(rtrim(部件)) in (" & strObjs & ")" & _
                "             and " & gstrSystems & _
                ")P" & _
                " where M.编号=G.编号 and M.模块=P.序号(+) And M.系统=P.系统(+) and M.模块=110 " & _
                " order by M.层次,Decode(upper(rtrim(P.部件)),'ZL9REPORT',1,0),M.编号"

        If .State = adStateOpen Then .Close
        .Open strSQL, gcnOracle, adOpenKeyset
    End With
    
    Set MenuGranted = rsTemp
End Function

Public Function OraDataOpen(ByVal strServerName As String, ByVal strUserName As String, ByVal strUserPwd As String) As Boolean
    '------------------------------------------------
    '功能： 打开指定的数据库
    '参数：
    '   strServerName：主机字符串
    '   strUserName：用户名
    '   strUserPwd：密码
    '返回： 数据库打开成功，返回true；失败，返回false
    '------------------------------------------------
    Dim strSQL As String
    Dim strError As String
    
    On Error Resume Next
    Err = 0
    DoEvents
    With gcnOracle
        If .State = adStateOpen Then .Close
        .Provider = "MSDataShape"
        .Open "Driver={Microsoft ODBC for Oracle};Server=" & strServerName, strUserName, strUserPwd
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
                MsgBox "由于用户、口令或服务器指定错误，无法注册。", vbInformation, gstrSysName
            End If
            
            OraDataOpen = False
            Exit Function
        End If
    End With
    
    Err = 0
    On Error GoTo ErrHand
    
    gstrDbUser = UCase(strUserName)
    SetDbUser gstrDbUser
    
    OraDataOpen = True
    Exit Function
    
ErrHand:
    If ErrCenter() = 1 Then Resume
    OraDataOpen = False
    Err = 0
End Function

Public Function OraDataClose() As Boolean
    '------------------------------------------------
    '功能： 关闭数据库
    '参数：
    '返回： 关闭数据库，返回True；失败，返回False
    '------------------------------------------------
    Err = 0
    On Error Resume Next
    gcnOracle.Close
    OraDataClose = True
    Err = 0

End Function

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

Public Function UpdatePassword(ByVal strUserName As String, ByVal strPasswd As String) As Boolean
    '-------------------------------------------------------------
    '功能：按人员ID，修改其密码
    '参数：CurrUser
    '      当前用户集
    '返回：如果成功则退回True，否则返回False
    '-------------------------------------------------------------
    Err = 0
    On Error GoTo ErrorHand
    
    DoEvents
    gcnOracle.Execute "alter user " & strUserName & " identified by " & strPasswd
    UpdatePassword = True
    Exit Function
    
ErrorHand:
    If ErrCenter() = 1 Then Resume
    UpdatePassword = False

End Function

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

Public Function GetUserInfo(ByVal strSystems As String)
    Dim rsTmp As New ADODB.Recordset, rsUser As New ADODB.Recordset
    Dim strSQL As String, i As Integer
    '读用户信息赋予公共，便于其他程序使用
    
    With rsTmp
        If .State = adStateOpen Then .Close
        strSQL = "Select S.*" & _
                " From zlSystems S,(Select Distinct owner From All_Tables Where Table_Name='部门表') D" & _
                " Where Upper(S.所有者)=D.Owner And S.编号 In (" & strSystems & ") Order by S.编号"
        .Open strSQL, gcnOracle, adOpenKeyset
        If Not .EOF Then
            '因为可能该用户具有多个系统的身份，所以循环取身份
            glngUserId = 0 '当前用户id
            gstrUserCode = "" '当前用户编码
            gstrUserName = "" '当前用户姓名
            gstrUserAbbr = "" '当前用户简码
            glngDeptId = 0 '当前用户部门id
            gstrDeptCode = "" '当前用户
            gstrDeptName = "" '当前用户
            
            For i = 1 To .RecordCount
                strSQL = "Select R.*,D.编码 as 部门编码,D.名称 as 部门名称,P.编号,P.姓名,P.简码" & _
                        " From " & !所有者 & ".上机人员表 U," & !所有者 & ".人员表 P," & !所有者 & ".部门表 D," & !所有者 & ".部门人员 R" & _
                        " Where U.人员ID = P.ID And R.部门ID = D.ID And P.ID=R.人员ID and U.用户名=USER and R.缺省=1"
                Set rsUser = New ADODB.Recordset
                rsUser.CursorLocation = adUseClient
                rsUser.Open strSQL, gcnOracle, adOpenKeyset
                If Not rsUser.EOF Then
                    glngUserId = rsUser!人员ID '当前用户id
                    gstrUserCode = rsUser!编号 '当前用户编码
                    gstrUserName = IIf(IsNull(rsUser!姓名), "", rsUser!姓名) '当前用户姓名
                    gstrUserAbbr = IIf(IsNull(rsUser!简码), "", rsUser!简码) '当前用户简码
                    glngDeptId = rsUser!部门id '当前用户部门id
                    gstrDeptCode = rsUser!部门编码 '当前用户
                    gstrDeptName = rsUser!部门名称 '当前用户
                    Exit For
                End If
                DoEvents
                .MoveNext
            Next
        End If
        .Close
    End With
End Function

Public Function OpenSQLRecord(ByVal strSQL As String, ByVal strTitle As String, ParamArray arrInput() As Variant) As ADODB.Recordset
'功能：通过Command对象打开带参数SQL的记录集
'参数：strSQL=条件中包含参数的SQL语句,参数形式为"[x]"
'             x>=1为自定义参数号,"[]"之间不能有空格
'             同一个参数可多处使用,程序自动换为ADO支持的"?"号形式
'             实际使用的参数号可不连续,但传入的参数值必须连续(如SQL组合时不一定要用到的参数)
'      arrInput=不定个数的参数值,按参数号顺序依次传入,必须是明确类型
'      strTitle=用于SQLTest识别的调用窗体/模块标题
'返回：记录集，CursorLocation=adUseClient,LockType=adLockReadOnly,CursorType=adOpenStatic
'举例：
'SQL语句为="Select 姓名 From 病人信息 Where (病人ID=[3] Or 门诊号=[3] Or 姓名 Like [4]) And 性别=[5] And 登记时间 Between [1] And [2] And 险类 IN([6],[7])"
'调用方式为：Set rsPati=OpenSQLRecord(strSQL, Me.Caption, CDate(Format(rsMove!转出日期,"yyyy-MM-dd")),dtp时间.Value, lng病人ID, "张%", "男", 20, 21)
    Static cmdData As New ADODB.Command
    Dim strPar As String, arrPar As Variant
    Dim lngLeft As Long, lngRight As Long
    Dim strSeq As String, intMax As Integer, i As Integer
    Dim strLog As String, varValue As Variant
    
    '分析自定的[x]参数
    lngLeft = InStr(1, strSQL, "[")
    Do While lngLeft > 0
        lngRight = InStr(lngLeft + 1, strSQL, "]")
        
        '可能是正常的"[编码]名称"
        strSeq = Mid(strSQL, lngLeft + 1, lngRight - lngLeft - 1)
        If IsNumeric(strSeq) Then
            i = CInt(strSeq)
            strPar = strPar & "," & i
            If i > intMax Then intMax = i
        End If
        
        lngLeft = InStr(lngRight + 1, strSQL, "[")
    Loop

    '替换为"?"参数
    strLog = strSQL
    For i = 1 To intMax
        strSQL = Replace(strSQL, "[" & i & "]", "?")
        
        '产生用于SQL跟踪的语句
        varValue = arrInput(i - 1)
        Select Case TypeName(varValue)
        Case "Byte", "Integer", "Long", "Single", "Double", "Currency" '数字
            strLog = Replace(strLog, "[" & i & "]", varValue)
        Case "String" '字符
            strLog = Replace(strLog, "[" & i & "]", "'" & Replace(varValue, "'", "''") & "'")
        Case "Date" '日期
            strLog = Replace(strLog, "[" & i & "]", "To_Date('" & Format(varValue, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')")
        Case "Variant" '不明确类型
            strLog = Replace(strLog, "[" & i & "]", "?")
        End Select
    Next

    '清除原有参数:不然不能重复执行
    cmdData.CommandText = "" '不为空有时清除参数出错
    Do While cmdData.Parameters.Count > 0
        cmdData.Parameters.Delete 0
    Loop
    
    '创建新的参数
    arrPar = Split(Mid(strPar, 2), ",")
    For i = 0 To UBound(arrPar)
        varValue = arrInput((arrPar(i) - 1))
        Select Case TypeName(varValue)
        Case "Byte", "Integer", "Long", "Single", "Double", "Currency" '数字
            cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i, adVarNumeric, adParamInput, 30, varValue)
        Case "Date" '日期
            cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i, adDBTimeStamp, adParamInput, , varValue)
        Case "String" '字符
            cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i, adVarChar, adParamInput, 500, varValue)
        Case "Variant" '不明确类型
        End Select
    Next

    '执行返回记录集
    If cmdData.ActiveConnection Is Nothing Then
        Set cmdData.ActiveConnection = gcnOracle '这句比较慢
    End If
    cmdData.CommandText = strSQL
    
    Call SQLTest(App.ProductName, strTitle, strLog)
    Set OpenSQLRecord = cmdData.Execute
    Call SQLTest
End Function




