Attribute VB_Name = "mdlMain"
Option Explicit

Public ZlBrowerDll As Object                '导航台
Public SplashObj As New frmSplash
'Public gcnOracle As New ADODB.Connection    '公共数据库连接

Public gstrSysName As String                '系统名称
Public gstrVersion As String                '系统版本
Public gstrAviPath As String                'AVI文件的存放目录
Public gstrUserFlag As String               '当前用户标志(两位表示)，第1位：是否DBA；第2位：系统所有者

Public gstrDbUser As String                 '当前数据库用户
Public glngUserId As Long                   '当前用户id
Public gstrUserCode As String               '当前用户编码
Public gstrUserName As String               '当前用户姓名
Public gstrUserAbbr As String               '当前用户简码
Public gstrServerName As String
Public glngDeptId As Long                   '当前用户部门id
Public gstrDeptCode As String               '当前用户部门编码
Public gstrDeptName As String               '当前用户部门名称

Public gstrStation As String                '本工作站名称
Public gstrMenuSys As String                '系统菜单
Public gstrCommand As String
Public gstrSystems As String

Public gobjFile As New FileSystemObject

Private Declare Function InitCommonControls Lib "comctl32.dll" () As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Public Const SRCCOPY = &HCC0020
Public Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, _
        ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
'---------------------------------------------------------------
'-注册表 API 声明...
'---------------------------------------------------------------
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long
Private Declare Function RegCreateKeyEx Lib "advapi32" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, ByRef lpSecurityAttributes As SECURITY_ATTRIBUTES, ByRef phkResult As Long, ByRef lpdwDisposition As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long
Public Declare Function SetActiveWindow Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

'切换到指定的输入法。
Public Declare Function ActivateKeyboardLayout Lib "user32" (ByVal hkl As Long, ByVal flags As Long) As Long
'返回系统中可用的输入法个数及各输入法所在Layout,包括英文输入法。
Public Declare Function GetKeyboardLayoutList Lib "user32" (ByVal nBuff As Long, lpList As Long) As Long
'获取某个输入法的名称
Public Declare Function ImmGetDescription Lib "imm32.dll" Alias "ImmGetDescriptionA" (ByVal hkl As Long, ByVal lpsz As String, ByVal uBufLen As Long) As Long
'判断某个输入法是否中文输入法
Public Declare Function ImmIsIME Lib "imm32.dll" (ByVal hkl As Long) As Long



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
Const KEY_ALL_ACCESS = KEY_QUERY_VALUE + KEY_SET_VALUE + _
                       KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS + _
                       KEY_NOTIFY + KEY_CREATE_LINK + READ_CONTROL
                     
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

Public Enum REGISTER
    注册信息
    私有模块
    私有全局
    公共模块
    公共全局
End Enum

'---------------------------------------------------------------
'启动时间，用以判断闪现屏幕的等待时间
'---------------------------------------------------------------
Public gdtStart As Long

'---------------------------------------------------------------
'   授权、菜单、试用版本
'---------------------------------------------------------------

'----------------------------------------------------------------------------------------------------------------------------------------------------------------
'功能:设置相关进程处理的API声明:2008-10-30 11:34:11:刘兴宏
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessID As Long) As Long
Private Declare Function CreateToolhelpSnapshot Lib "kernel32" Alias "CreateToolhelp32Snapshot" (ByVal lFlags As Long, ByVal lProcessId As Long) As Long
Private Declare Function ProcessFirst Lib "kernel32" Alias "Process32First" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function ProcessNext Lib "kernel32" Alias "Process32Next" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Public Declare Function EnumWindows Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Public Declare Function IsWindowVisible Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long
Public Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Type PROCESSENTRY32
      lSize             As Long
      lUsage            As Long
      lProcessId        As Long
      lDefaultHeapId    As Long
      lModuleId         As Long
      lThreads          As Long
      lParentProcessId  As Long
      lPriClassBase     As Long
      lFlags            As Long
      sExeFile          As String * 1024
End Type
Private Const PROCESS_TERMINATE = &H1
Public gcll_His_PId As Collection        '存储相关的进程信息:array(进程名称,PID,窗口个数),"K"+进程数


Public zlCommFun As New zlMipClientComLib.clsCommFun
Public zlDataBase As New zlMipClientComLib.clsDatabase
Public zlComLib As New zlMipClientComLib.clsComLib
Public zlControl As New zlMipClientComLib.clsControl
Private mobjComLib As Object

Public gclsMsgSystem As New clsBusiness
Public gclsMsgOracle As New zlDataOracle.clsDataOracle

#Const SYS_TRYUSE = "正式" '正式/试用

Public Sub Main()
    Dim lngReturn As Long
    Dim StrUnitName As String
    Dim BlnShowFlash As Boolean
    Dim strCode As String, intCount As Integer, strStyle As String
    Dim strTitle As String                  '产品标题
    Dim strTag As String                    '旗舰版标志
    Dim rsMenu As ADODB.Recordset
    
'''    Call SetTNSNameFile 20110112-ZQ
    
     '为实现XP风格，在显示窗体前必须执行该函数
    Call InitCommonControls

    BlnShowFlash = False
    If InStr(Command(), "=") <= 0 Then Load SplashObj
    
    '由注册表中获取用户注册相关信息,如果用户单位名称不为空,则显示闪现窗体
    StrUnitName = GetSetting("ZLSOFT", "注册信息", "单位名称", "")
    gstrSysName = GetSetting("ZLSOFT", "注册信息", "提示", "")
    If StrUnitName <> "" And StrUnitName <> "-" Then
        gdtStart = Timer
        With SplashObj
            '有两处需要处理
            Call zlComLib.ApplyOEM_Picture(.ImgIndicate, "Picture")
            Call zlComLib.ApplyOEM_Picture(.imgPic, "PictureB")
            If InStr(Command(), "=") <= 0 Then .Show
            .lblGrant = StrUnitName
            StrUnitName = GetSetting("ZLSOFT", "注册信息", "开发商", "")
            If Trim(StrUnitName) = "" Then
                .Label3.Visible = False
                .lbl开发商.Visible = False
            Else
                .lbl开发商.Caption = ""
                For intCount = 0 To UBound(Split(StrUnitName, ";"))
                    .lbl开发商.Caption = .lbl开发商.Caption & Split(StrUnitName, ";")(intCount) & vbCrLf
                Next
            End If
            .LblProductName = GetSetting("ZLSOFT", "注册信息", "产品全称", "")
            If Len(.LblProductName) > 10 Then
                .LblProductName.FontSize = 15.75 '三号
            Else
                .LblProductName.FontSize = 21.75 '二号
            End If
            .lbl技术支持商 = GetSetting("ZLSOFT", "注册信息", "技术支持商", "")
            .lbltag = GetSetting("ZLSOFT", "注册信息", "产品系列", "")
            
        End With
        Do
            If (Timer - gdtStart) > 3 Then Exit Do
            DoEvents
        Loop
        
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
    
    Call zlKillHISPID
    
    '用户注册
    If InStr(Command(), "=") > 0 Then
        Call frmUserLogin.Docmd(Command())
    Else
        frmUserLogin.Show 1
    End If
        
    If gclsMsgOracle.DatabaseState <> adStateOpen Then
        Unload frmUserLogin
        Unload SplashObj
        Exit Sub
    End If
    
    '写入本次启动程序的EXE文件名
    Call SaveSetting("ZLSOFT", "公共全局", "执行文件", App.EXEName & ".exe")
    
    '初始化公共部件
    Call zlComLib.InitCommon(gclsMsgOracle.DatabaseConnection)
    
    Set mobjComLib = CreateObject("zl9Comlib.clsComlib")
    If Not (mobjComLib Is Nothing) Then
        Call mobjComLib.InitCommon(gclsMsgOracle.DatabaseConnection)
        If mobjComLib.RegCheck = False Then
            Unload SplashObj
            Exit Sub
        End If
    End If
    gstrSysName = mobjComLib.zlRegInfo("产品简名") & "软件"
    
'    SaveSetting "ZLSOFT", "注册信息", "提示", gstrSysName
'    SaveSetting "ZLSOFT", "注册信息", UCase("gstrSysName"), gstrSysName
'    gstrVersion = App.Major & "." & App.Minor & "." & App.Revision
'    SaveSetting "ZLSOFT", "注册信息", UCase("gstrVersion"), gstrVersion
'    gstrAviPath = App.Path & "\附加文件"
'    SaveSetting "ZLSOFT", "注册信息", UCase("gstrAviPath"), gstrAviPath
'
'    strTag = ""
'    strTitle = zlComLib.zlRegInfo("产品标题")
'    If strTitle <> "" Then
'        If InStr(strTitle, "-") > 0 Then
'            If Split(strTitle, "-")(1) = "Ultimate" Then
'                strTag = "旗舰版"
'            ElseIf Split(strTitle, "-")(1) = "Professional" Then
'                strTag = "专业版"
'            End If
'        End If
'    End If
'    strTitle = Split(strTitle, "-")(0)
'    With SplashObj
'        If BlnShowFlash = False Then
'            .lblGrant = zlComLib.zlRegInfo("单位名称", , -1)
'            .lbl技术支持商.Caption = zlComLib.zlRegInfo("技术支持商", , -1)
'
'            .LblProductName = strTitle
'            .lbltag = strTag
'            strCode = zlComLib.zlRegInfo("产品开发商", , -1)
'            .lbl开发商.Caption = ""
'            For IntCount = 0 To UBound(Split(strCode, ";"))
'                .lbl开发商.Caption = .lbl开发商.Caption & Split(strCode, ";")(IntCount) & vbCrLf
'            Next
'            Call zlComLib.ApplyOEM_Picture(.ImgIndicate, "Picture")
'            If InStr(Command(), "=") <= 0 Then .Show
'            BlnShowFlash = True
'        End If
'        DoEvents
'    End With
    
'    '将用户注册相关信息写入注册表,供下次启动时显示
'    SaveSetting "ZLSOFT", "注册信息", "单位名称", zlComLib.zlRegInfo("单位名称", , -1)
'    SaveSetting "ZLSOFT", "注册信息", "产品全称", strTitle
'    SaveSetting "ZLSOFT", "注册信息", "产品名称", zlComLib.zlRegInfo("产品简名")
'    SaveSetting "ZLSOFT", "注册信息", "技术支持商", zlComLib.zlRegInfo("技术支持商", , -1)
'    SaveSetting "ZLSOFT", "注册信息", "开发商", zlComLib.zlRegInfo("产品开发商", , -1)
'    SaveSetting "ZLSOFT", "注册信息", "WEB支持商简名", zlComLib.zlRegInfo("支持商简名")
'    SaveSetting "ZLSOFT", "注册信息", "WEB支持EMAIL", zlComLib.zlRegInfo("支持商MAIL")
'    SaveSetting "ZLSOFT", "注册信息", "WEB支持URL", zlComLib.zlRegInfo("支持商URL")
'    SaveSetting "ZLSOFT", "注册信息", "产品系列", strTag
        
    Unload SplashObj
    
    SplashObj.Tag = gstrSystems
            
            
    '检测是否安装了消息集成平台ZLHIS客户端，如果没有安装，则必须弹出安装程序
    '------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    Dim strSQL As String

    strSQL = "Select 行号,内容 From zlRegInfo Where 项目='消息集成平台客户端'"
    Set rsTemp = gclsMsgOracle.OpenSQLRecord(strSQL, gstrSysName)
    If rsTemp.BOF = True Then
        '无记录表示未安装
        If frmAppCreate.ShowDialog Then
            MsgBox "消息集成平台客户端已经安装成功，可以使用！", vbInformation, gstrSysName
        Else
            MsgBox "消息集成平台客户端已经取消安装，点击确定后将自动退出！", vbInformation, gstrSysName
            Unload frmUserLogin
            Unload SplashObj
            Exit Sub
        End If
    End If
    
    '------------------------------------------------------------------------------------------------------------------
    Call frmNativateStart.SetEnvironment(gstrSysName, gstrVersion, gstrAviPath, _
                          gstrUserFlag, gstrDbUser, glngUserId, _
                          gstrUserCode, gstrUserName, gstrUserAbbr, _
                          glngDeptId, gstrDeptCode, gstrDeptName, _
                          gstrStation, gstrMenuSys, CStr(Command()))
    Call frmNativateStart.InitBrower(SplashObj, gclsMsgOracle.DatabaseConnection, rsMenu)
        
End Sub

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
    Call gclsMsgOracle.UpdateUserPassword(strUserName, strPasswd)
'    gcnOracle.Execute "alter user " & strUserName & " identified by " & strPasswd
    UpdatePassword = True
    Exit Function
    
ErrorHand:
    If zlComLib.ErrCenter() = 1 Then Resume
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

Public Function GetUserInfo(ByVal strSystems As String)
    Dim rsTmp As New ADODB.Recordset, rsUser As New ADODB.Recordset
    Dim strSQL As String, i As Integer
    '读用户信息赋予公共，便于其他程序使用
'
'    With rsTmp
'        If .State = adStateOpen Then .Close
'        strSQL = "Select S.*" & _
'                " From zlSystems S,(Select Distinct owner From All_Tables Where Table_Name='部门表') D" & _
'                " Where Upper(S.所有者)=D.Owner And S.编号 In (" & strSystems & ") Order by S.编号"
'        .Open strSQL, gcnOracle, adOpenKeyset
        
    Set rsTmp = gclsMsgSystem.GetSystemInfo(strSystems)
    With rsTmp
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
                
                
                Set rsUser = gclsMsgSystem.GetUserInfo(!所有者)
                
'                strSQL = "Select R.*,D.编码 as 部门编码,D.名称 as 部门名称,P.编号,P.姓名,P.简码" & _
'                        " From " & !所有者 & ".上机人员表 U," & !所有者 & ".人员表 P," & !所有者 & ".部门表 D," & !所有者 & ".部门人员 R" & _
'                        " Where U.人员ID = P.ID And R.部门ID = D.ID And P.ID=R.人员ID and U.用户名=USER And (P.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or P.撤档时间 Is Null) and R.缺省=1"
'                Set rsUser = New ADODB.Recordset
'                rsUser.CursorLocation = adUseClient
'                rsUser.Open strSQL, gcnOracle, adOpenKeyset
                Set rsUser.ActiveConnection = Nothing
                If Not rsUser.EOF Then
                    glngUserId = rsUser!人员ID '当前用户id
                    gstrUserCode = rsUser!编号 '当前用户编码
                    gstrUserName = IIf(IsNull(rsUser!姓名), "", rsUser!姓名) '当前用户姓名
                    gstrUserAbbr = IIf(IsNull(rsUser!简码), "", rsUser!简码) '当前用户简码
                    glngDeptId = rsUser!部门ID '当前用户部门id
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

Private Function RunningInIDE() As Boolean
    '--检测是否源代码环境
    RunningInIDE = (App.EXEName = "zl9WizardMain")
End Function

'**********************************************************************************************************************
'功能:以下处理相关进程的函数
'编制:刘兴洪
'日期:2008-10-30 11:38:58
Public Function zlKillHISPID() As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '功能:杀死所有HIS启动程序的移常进程(杀的条件是:所有ZLHIS+.exe的进程中无任何窗口)
    '入参:
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2008-10-30 11:06:16
    '-----------------------------------------------------------------------------------------------------------
    Dim lngProcess As Long, i As Long
    
    zlKillHISPID = False
    Err = 0: On Error GoTo errHand:
    '第一步:需要处理相关的ZLHIS的相关进程
    Set gcll_His_PId = New Collection
    If zlHISPidToCollect(gcll_His_PId) = False Then zlKillHISPID = True: Exit Function  '如果存在相关的错误，就直接返回了
    If gcll_His_PId Is Nothing Then zlKillHISPID = True: Exit Function
    If gcll_His_PId.Count = 0 Then zlKillHISPID = True: Exit Function
    
    '第二步:需要处理相关ZLHIS的相关进程的相关窗口个数,这样才好判断出相关的进程是否存在异常,出现异常的，就得杀掉
    Call EnumWindows(AddressOf EnumWindowsProc, 0&)
    For i = 1 To gcll_His_PId.Count
        If Val(gcll_His_PId(i)(2)) <= 1 Then
            '肯定窗口数小于1或零,那么肯定有异常，需要杀死他
            If Val(gcll_His_PId(i)(1)) <> 0 Then
                '可能未成功，暂无处理此种情况
                Call TerminatePID(Val(gcll_His_PId(i)(1)))
            End If
        End If
    Next
    zlKillHISPID = True
errHand:
End Function

Private Function EnumWindowsProc(ByVal hWnd As Long, ByVal lParam As Long) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '功能:获取所有窗口符合HIS的进程的窗口
    '入参:
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2008-10-30 10:26:02
    '-----------------------------------------------------------------------------------------------------------
    Dim strTittle As String, lngPID As Long, strName As String
    Dim lngCount As Long
    
    If GetParent(hWnd) = 0 Then
        '读取 hWnd 的视窗标题
        strTittle = String(80, 0)
        Call GetWindowText(hWnd, strTittle, 80)
        strTittle = Left(strTittle, InStr(strTittle, Chr(0)) - 1)
        If Trim(strTittle) <> "" Then
            Call GetWindowThreadProcessId(hWnd, lngPID)
            If IsWindowVisible(hWnd) Then
                Err = 0: On Error Resume Next
                strName = gcll_His_PId("K" & lngPID)(0)
                If Err = 0 Then
                    lngCount = Val(gcll_His_PId("K" & lngPID)(2)) + 1
                    gcll_His_PId.Remove "K" & lngPID
                    gcll_His_PId.Add Array(strName, lngPID, lngCount), "K" & lngPID
                End If
                Err.Clear: On Error GoTo 0
            End If
        End If
    End If
    EnumWindowsProc = True ' 表示继续列举 hWnd
    Exit Function
End Function

Private Function TerminatePID(ByVal lngPID As Long) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '功能:结束指定的进程
    '入参:
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2008-10-30 11:06:16
    '-----------------------------------------------------------------------------------------------------------
    Dim lngProcess As Long
    TerminatePID = False
    
    Err = 0: On Error GoTo errHand:
    lngProcess = OpenProcess(PROCESS_TERMINATE, 0&, lngPID)
    Call TerminateProcess(lngProcess, 1&)
    
    TerminatePID = True
errHand:
End Function

Private Function zlHISPidToCollect(ByRef cll_His_Pid As Collection) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '功能:获取ZLHIS的进程给相关的集合(gcll_HIS_Pid)
    '入参:
    '出参:cll_His_Pid-将符合HIS.exe的程序，装载该集合中
    '返回:
    '编制:刘兴洪
    '日期:2008-10-30 10:07:38
    '-----------------------------------------------------------------------------------------------------------
    Dim strExeName  As String, lngSnapShot As Long, lngProcess As Long, lngCount  As Long
    Dim strCurExeName As String, lngCurPid As Long
    Dim uProcess   As PROCESSENTRY32
    Const TH32CS_SNAPPROCESS = &H2
    
    Err = 0: On Error GoTo errHand:
    strCurExeName = "*" & UCase(App.EXEName) & "*"
    
    lngCurPid = GetCurrentProcessId '获取当前应用程序进程
    lngSnapShot = CreateToolhelpSnapshot(TH32CS_SNAPPROCESS, 0&)
    If lngSnapShot <> 0 Then
        uProcess.lSize = Len(uProcess)
        lngProcess = ProcessFirst(lngSnapShot, uProcess)
        lngCount = 0
        Do While lngProcess
            '不等于当前进程的才处理
            If lngCurPid <> uProcess.lProcessId Then
                strExeName = UCase(Left(uProcess.sExeFile, InStr(1, uProcess.sExeFile, vbNullChar) - 1))
                If strExeName Like strCurExeName Then
                    cll_His_Pid.Add Array(strExeName, uProcess.lProcessId, 0), "K" & uProcess.lProcessId
                End If
            End If
            lngProcess = ProcessNext(lngSnapShot, uProcess)
        Loop
        CloseHandle (lngSnapShot)
    End If
    zlHISPidToCollect = True
    Exit Function
errHand:
End Function

''''**********************************************************************************************************************
'''
'''Private Function SetTNSNameFile() As Boolean
'''    Dim strFile As String, intFile As Integer
'''    Dim arrData() As Byte
'''
'''    On Error GoTo errHandle
'''
'''    strFile = GetKeyValue(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\VisualStudio\6.0\Setup\Microsoft Visual Basic", "ProductDir")
'''    If strFile <> "" Then
'''        If gobjFile.FolderExists(strFile) Then
'''            SetTNSNameFile = True
'''            Exit Function
'''        End If
'''    End If
'''
'''    strFile = GetKeyValue(HKEY_LOCAL_MACHINE, "SOFTWARE\ORACLE", "ORACLE_HOME")
'''    If Not gobjFile.FolderExists(strFile) Then '10G
'''        strFile = GetKeyValue(HKEY_LOCAL_MACHINE, "SOFTWARE\ORACLE", "ORA_CRS_HOME")
'''    End If
'''    If Not gobjFile.FolderExists(strFile) Then '10Gr2
'''        strFile = GetKeyValue(HKEY_LOCAL_MACHINE, "SOFTWARE\ORACLE\KEY_OraDb10g_home1", "ORACLE_HOME")
'''    End If
'''    If Not gobjFile.FolderExists(strFile) Then '10Gr2
'''        strFile = GetKeyValue(HKEY_LOCAL_MACHINE, "SOFTWARE\ORACLE\KEY_OraDb10g_home2", "ORACLE_HOME")
'''    End If
'''    If Not gobjFile.FolderExists(strFile) Then '10G 企业版
'''        strFile = GetKeyValue(HKEY_LOCAL_MACHINE, "SOFTWARE\ORACLE\KEY_OraClient10g_home1", "ORACLE_HOME")
'''    End If
'''    strFile = strFile & "\network\admin\tnsnames.ora"
'''    If Not gobjFile.FileExists(strFile) Then Exit Function
'''    gobjFile.DeleteFile strFile
'''
'''    arrData = LoadResData(101, "CUSTOM")
'''    intFile = FreeFile
'''
'''    Open strFile For Binary As intFile
'''    Put intFile, , arrData()
'''    Close intFile
'''
'''    SetTNSNameFile = True
'''
'''    Exit Function
'''errHandle:
'''    'MsgBox "错误：" & Err.Number & vbCrLf & vbTab & Err.Description, vbExclamation, App.Title
'''End Function

Public Function SetRegister(ByVal enmRegister As REGISTER, ByVal strSection As String, ByVal strKey As String, ByVal strKeyValue As String) As Boolean
    '******************************************************************************************************************
    '功能： 将指定的信息保存在注册表中
    '参数： enmRegister-注册类型
    '       strSection-注册表目录
    '       strKey-键名
    '       strKeyValue-键值
    '返回：
    '******************************************************************************************************************
    On Error GoTo errHand
    
    Select Case enmRegister
    Case 注册信息
        
        Call SaveSetting("ZLSOFT", "注册信息\" & strSection, strKey, strKeyValue)
        
    Case 私有模块

        Call SaveSetting("ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\" & strSection, strKey, strKeyValue)
        
    Case 私有全局

        Call SaveSetting("ZLSOFT", "私有全局\" & gstrDbUser & "\" & strSection, strKey, strKeyValue)
        
    Case 公共模块

        Call SaveSetting("ZLSOFT", "公共模块\" & App.ProductName & "\" & strSection, strKey, strKeyValue)
        
    Case 公共全局
        
        Call SaveSetting("ZLSOFT", "公共全局\" & strSection, strKey, strKeyValue)
        
    End Select
    
    SetRegister = True
    
errHand:
    
End Function

Public Function GetRegister(ByVal enmRegister As REGISTER, ByVal strSection As String, ByVal strKey As String, ByVal strDefKeyValue As String) As String
    '******************************************************************************************************************
    '功能： 将指定的注册信息读取出来
    '参数： enmRegister-注册类型
    '       strSection-注册表目录
    '       strKey-键名
    '       strDefKeyValue-缺省键值
    '返回： strKeyValue-键值
    '******************************************************************************************************************

    Dim strValue As String
    
    On Error GoTo errHand
    
    Select Case enmRegister
    Case 注册信息
        
        strValue = GetSetting("ZLSOFT", "注册信息\" & strSection, strKey, strDefKeyValue)
        
    Case 私有模块

        strValue = GetSetting("ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\" & strSection, strKey, strDefKeyValue)
        
    Case 私有全局

        strValue = GetSetting("ZLSOFT", "私有全局\" & gstrDbUser & "\" & strSection, strKey, strDefKeyValue)
        
    Case 公共模块

        strValue = GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\" & strSection, strKey, strDefKeyValue)
        
    Case 公共全局
        
        strValue = GetSetting("ZLSOFT", "公共全局\" & strSection, strKey, strDefKeyValue)
        
    End Select
    
    GetRegister = strValue
    
errHand:
End Function
