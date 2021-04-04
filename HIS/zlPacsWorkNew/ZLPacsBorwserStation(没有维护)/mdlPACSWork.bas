Attribute VB_Name = "mdlPACSWork"
Option Explicit
Public SplashObj As New frmSplash
Public gstrStation As String                '本工作站名称
Public gstrSystems As String
Public gstr单位名称 As String
'-----------------------------------------
'发行码、注册码、发行码解析串、注册码解析串
Public gstrRegCode As String
Public gstrPublish As String
Public gstrParseRegCode As String
Public gstrParsePublish As String
'-----------------------------------------



Public gcnOracle As New ADODB.Connection        '公共数据库连接，特别注意：不能设置为新的实例
Public gstrPrivs As String                   '当前用户具有的当前模块的功能
Public gstrSysName As String                '系统名称
Public gstrVersion As String                '系统版本
Public gstrAviPath As String                'AVI文件的存放目录
Public glngModul As Long
Public glngSys As Long
Public gstrIme As String                    '是否自动开启输入法
Public gstrDbUser As String                 '当前数据库用户
Public glngUserId As Long                   '当前用户id
Public gstrUserCode As String               '当前用户编码
Public gstrUserName As String               '当前用户姓名
Public gstrUserAbbr As String               '当前用户简码

Public glngDeptId As Long                   '当前用户部门id
Public gstrDeptCode As String               '当前用户部门编码
Public gstrDeptName As String               '当前用户部门名称


Public gstrUnitName As String '用户单位名称
Public gfrmMain As Object

Public gstrSQL As String
Public glngTXTProc As Long
Public gbln加班加价 As Boolean
Public grsDuty As ADODB.Recordset '存放医生职务
Public grsSysPars As ADODB.Recordset
Public gbytCardNOLen As Byte

'系统参数
Public gbytDec As Byte '费用金额的小数点位数
Public gstrDec As String '按小数位数计算的格式化串,如"0.0000"

Public gobjKernel As New zlCISKernel.clsCISKernel '医嘱对像
Public gobjRichEPR As New zlRichEPR.cRichEPR

Public gbytCardLen As Byte '就诊卡号长度
Public gblnCardHide As Boolean '就诊卡号密文显示
Public gstrCardMask As String  '就诊卡允许的字母前缀:AA|BB|CC...
Public gint挂号天数 As Integer '挂号单有效天数

'列表颜色配置
Public gdblColor已登记 As Double
Public gdblColor已报到 As Double
Public gdblColor已检查 As Double
Public gdblColor已报告 As Double
Public gdblColor已完成 As Double
Public gdblColor已审核 As Double
Public gdblColor处理中 As Double
Public gdblColor报告中 As Double
Public gdblColor审核中 As Double
Public gdblColor已拒绝 As Double


Public gConnectedShardDir() As String   '已经连接过的共享目录的设备号数组

'---------------------------设备数量控制，注册-------------------------------
Public Const LOGIN_TYPE_视频设备 As String = "影像视频设备数量"
Public Const LOGIN_TYPE_胶片打印机 As String = "影像胶片打印机数量"
Public Const LOGIN_TYPE_DICOM设备 As String = "影像DICOM设备数量"
Public gint视频设备数量 As Integer
Public gint胶片打印机数量 As Integer
Public gintDICOM设备数量 As Integer


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
Public Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Boolean
End Type




Public mrsDeptParas As ADODB.Recordset '本科参数表缓存
'-----------------------------------------------------------
Public Type TYPE_USER_INFO
    ID As Long
    部门ID As Long
    编号 As String
    姓名 As String
    简码 As String
    用户名 As String
End Type
Public UserInfo As TYPE_USER_INFO

'读取网卡的多个IP
Private Const WS_VERSION_REQD = &H101
Private Const WS_VERSION_MAJOR = WS_VERSION_REQD \ &H100 And &HFF&
Private Const WS_VERSION_MINOR = WS_VERSION_REQD And &HFF&
Private Const MIN_SOCKETS_REQD = 1
Private Const SOCKET_ERROR = -1
Private Const WSADescription_Len = 256
Private Const WSASYS_Status_Len = 128

Private Type HOSTENT
    hName As Long
    hAliases As Long
    hAddrType As Integer
    hLength As Integer
    hAddrList As Long
End Type

Private Type WSADATA
    wversion As Integer
    wHighVersion As Integer
    szDescription(0 To WSADescription_Len) As Byte
    szSystemStatus(0 To WSASYS_Status_Len) As Byte
    iMaxSockets As Integer
    iMaxUdpDg As Integer
    lpszVendorInfo As Long
End Type

Private Declare Function InitCommonControls Lib "comctl32.dll" () As Long
Private Declare Function WSAGetLastError Lib "WSOCK32.DLL" () As Long
Private Declare Function WSAStartup Lib "WSOCK32.DLL" (ByVal wVersionRequired As Integer, lpWSAData As WSADATA) As Long
Private Declare Function WSACleanup Lib "WSOCK32.DLL" () As Long

Private Declare Function gethostname Lib "WSOCK32.DLL" (ByVal hostname$, ByVal HostLen As Long) As Long
Private Declare Function gethostbyname Lib "WSOCK32.DLL" (ByVal hostname$) As Long
Private Declare Sub RtlMoveMemory Lib "kernel32" (hpvDest As Any, ByVal hpvSource&, ByVal cbCopy&)

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


Public Sub Main()
    Dim lngReturn As Long
    Dim StrUnitName As String
    Dim BlnShowFlash As Boolean
    Dim strCode As String, IntCount As Integer, StrStyle As String
    Dim rsMenu As ADODB.Recordset, StrHaveSys As String
    
    
    If App.PrevInstance Then
        MsgBox "影像接收服务已经启动，不能再次运行。", vbInformation, "警告"
        Exit Sub
    End If
    
    
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
        
        Set gcnOracle = Nothing
        Exit Sub
    End If
    
    '初始化公共部件
    InitCommon gcnOracle
    If RegCheck = False Then
        Unload SplashObj
        
        Set gcnOracle = Nothing
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

    gstrSystems = " (系统 =100 Or 系统 Is NULL)"
    glngSys = 100
    
    '-------------------------------------------------------------
    '创建同义词
    '-------------------------------------------------------------
    zlDatabase.ExecuteProcedure "Zl_Createsynonyms(" & glngSys & ")", "创建同义词"
    

    Unload SplashObj
    
    CodeMan 1290
End Sub


Public Sub CodeMan(ByVal lngModul As Long)
    '------------------------------------------------
    '功能： 根据主程序指定功能，调用执行相关程序
    '参数：
    '   lngModul:需要执行的功能序号
    '返回：
    '------------------------------------------------
    Dim rsUser As ADODB.Recordset
    
    gstrAviPath = GetSetting(appName:="ZLSOFT", Section:="注册信息", Key:=UCase("gstrAviPath"), Default:="")
    gstrSysName = GetSetting(appName:="ZLSOFT", Section:="注册信息", Key:=UCase("gstrSysName"), Default:="")
    gstrVersion = GetSetting(appName:="ZLSOFT", Section:="注册信息", Key:=UCase("gstrVersion"), Default:="")
    gstr单位名称 = GetUnitName()
    
    '提取用户的信息
    Set rsUser = zlDatabase.GetUserInfo
    If rsUser.RecordCount <> 0 Then
        glngUserId = Nvl(rsUser!ID)
        gstrUserCode = Nvl(rsUser!编号)
        gstrUserName = Nvl(rsUser!姓名)
        gstrUserAbbr = Nvl(rsUser!简码)
        glngDeptId = Nvl(rsUser!部门ID)
        gstrDeptCode = Nvl(rsUser!部门码)
        gstrDeptName = Nvl(rsUser!部门名)
    Else
        glngUserId = 0
        gstrUserCode = ""
        gstrUserName = ""
        gstrUserAbbr = ""
        glngDeptId = 0
        gstrDeptCode = ""
        gstrDeptName = ""
    End If
    
    gstrPrivs = GetPrivFunc(glngSys, lngModul)
    '-------------------------------------------------
    
    Select Case lngModul
        Case 1290
            frmBrowserStation.Show
    End Select
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
    err = 0
    DoEvents
    With gcnOracle
        If .State = adStateOpen Then .Close
        .Provider = "MSDataShape"
        .Open "Driver={Microsoft ODBC for Oracle};Server=" & strServerName, strUserName, strUserPwd
        If err <> 0 Then
            '保存错误信息
            strError = err.Description
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
    
    err = 0
    On Error GoTo errHand
    
    gstrDbUser = UCase(strUserName)
    SetDbUser gstrDbUser
    OraDataOpen = True
    Exit Function
    
errHand:
    If ErrCenter() = 1 Then Resume
    OraDataOpen = False
    err = 0
End Function

Public Function OraDataClose() As Boolean
    '------------------------------------------------
    '功能： 关闭数据库
    '参数：
    '返回： 关闭数据库，返回True；失败，返回False
    '------------------------------------------------
    err = 0
    On Error Resume Next
    gcnOracle.Close
    OraDataClose = True
    err = 0

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


Public Function UpdatePassword(ByVal strUserName As String, ByVal strPasswd As String) As Boolean
    '-------------------------------------------------------------
    '功能：按人员ID，修改其密码
    '参数：CurrUser
    '      当前用户集
    '返回：如果成功则退回True，否则返回False
    '-------------------------------------------------------------
    err = 0
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


Public Sub ReadStudyListColor(ByVal lngDeptID As Long)

  gdblColor报告中 = GetStudyListColor(lngDeptID, "报告中")
  If gdblColor报告中 < 0 Then
    gdblColor报告中 = ColorConstants.vbWhite
  End If
  
  gdblColor处理中 = GetStudyListColor(lngDeptID, "处理中")
  If gdblColor处理中 < 0 Then
    gdblColor处理中 = ColorConstants.vbWhite
  End If
  
  gdblColor审核中 = GetStudyListColor(lngDeptID, "审核中")
  If gdblColor审核中 < 0 Then
    gdblColor审核中 = ColorConstants.vbWhite
  End If
  
  gdblColor已报到 = GetStudyListColor(lngDeptID, "已报到")
  If gdblColor已报到 < 0 Then
    gdblColor已报到 = ColorConstants.vbWhite
  End If
  
  gdblColor已登记 = GetStudyListColor(lngDeptID, "已登记")
  If gdblColor已登记 < 0 Then
    gdblColor已登记 = ColorConstants.vbWhite
  End If
  
  gdblColor已检查 = GetStudyListColor(lngDeptID, "已检查")
  If gdblColor已检查 < 0 Then
    gdblColor已检查 = ColorConstants.vbWhite
  End If
  
  gdblColor已审核 = GetStudyListColor(lngDeptID, "已审核")
  If gdblColor已审核 < 0 Then
    gdblColor已审核 = ColorConstants.vbWhite
  End If
  
  gdblColor已完成 = GetStudyListColor(lngDeptID, "已完成")
  If gdblColor已完成 < 0 Then
    gdblColor已完成 = ColorConstants.vbGreen
  End If
  
  gdblColor已报告 = GetStudyListColor(lngDeptID, "已报告")
  If gdblColor已报告 < 0 Then
    gdblColor已报告 = ColorConstants.vbWhite
  End If
  
  gdblColor已拒绝 = GetStudyListColor(lngDeptID, "已拒绝")
  If gdblColor已拒绝 < 0 Then
    gdblColor已拒绝 = ColorConstants.vbYellow
  End If
End Sub


Public Function GetUserInfo() As Boolean
'功能：获取登陆用户信息
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    Set rsTmp = zlDatabase.GetUserInfo
    
    UserInfo.用户名 = gstrDbUser
    UserInfo.姓名 = gstrDbUser
    If Not rsTmp.EOF Then
        UserInfo.ID = rsTmp!ID
        UserInfo.编号 = rsTmp!编号
        UserInfo.部门ID = IIf(IsNull(rsTmp!部门ID), 0, rsTmp!部门ID)
        UserInfo.简码 = IIf(IsNull(rsTmp!简码), "", rsTmp!简码)
        UserInfo.姓名 = IIf(IsNull(rsTmp!姓名), "", rsTmp!姓名)
        UserInfo.用户名 = IIf(IsNull(rsTmp!用户名), "", rsTmp!用户名)
        GetUserInfo = True
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetUser科室IDs(Optional ByVal bln病区 As Boolean) As String
'功能：获取操作员所属的科室(本身所在科室+所属病区包含的科室),可能有多个
'参数：是否取所属病区下的科室
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    
    strSQL = "Select 部门ID From 部门人员 Where 人员ID=[1]"
    If bln病区 Then
        strSQL = strSQL & " Union" & _
            " Select Distinct B.科室ID From 部门人员 A,床位状况记录 B" & _
            " Where A.部门ID=B.病区ID And A.人员ID=[1]"
    End If
    
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISWork", UserInfo.ID)
    For i = 1 To rsTmp.RecordCount
        GetUser科室IDs = GetUser科室IDs & "," & rsTmp!部门ID
        rsTmp.MoveNext
    Next
    GetUser科室IDs = Mid(GetUser科室IDs, 2)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

'取得检查列表指定的配置颜色
Public Function GetStudyListColor(ByVal lngDeptID As Long, ByVal strParameterName As String) As Double
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim lngTemp As Long
             
    On Error GoTo err
        
    strSQL = "select ID ,科室ID,参数名,参数值 from 影像流程参数 where 科室ID = [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "取得检查列表颜色", lngDeptID)
        
    GetStudyListColor = -1
    
    While Not rsTemp.EOF
        If rsTemp!参数名 = strParameterName Then
          GetStudyListColor = Val(rsTemp!参数值)
          Exit Function
        End If
        rsTemp.MoveNext
    Wend
    
    Exit Function
err:
    If ErrCenter() = 1 Then Resume Next
    Call SaveErrLog
End Function

Public Function getID_TO_名称(ByVal lngID As Long, ByVal strDict As String) As String
Dim rsTemp As ADODB.Recordset
    On Error GoTo errHandle
    gstrSQL = "select 名称 FROM " & strDict & " WHERE ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "通过编码提取ID", lngID)
    If Not rsTemp.EOF Then
        getID_TO_名称 = rsTemp!名称
    End If
    Exit Function
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Sub RemoveCheckImages(ByVal lng医嘱ID As Long, ByVal lng发送号 As Long)
    '删除指定医嘱的检查影像
    Dim strSQL As String, rsTmp As New ADODB.Recordset
    
    Dim Inte As New clsFtp
    Dim strDeviceNO As String
    On Error GoTo ProcError
    '先删除图像
    strSQL = "select a.IP地址, a.FTP目录, a.FTP用户名, a.FTP密码, a.医嘱ID, a.发送号, a.检查UID, a.位置, a.接收日期 ,a.设备号 ,c.图像UID" & _
             " from (select IP地址, FTP目录, FTP用户名, FTP密码, 医嘱ID, 发送号, 检查UID, 位置一 as 位置, 接收日期, a.设备号 " & _
             "       from 影像设备目录 a, 影像检查记录 b " & _
             "       Where a.设备号 = B.位置一 " & _
             "       Union All " & _
             "       select IP地址, FTP目录, FTP用户名, FTP密码, 医嘱ID, 发送号, 检查UID, 位置二 as 位置, 接收日期, a.设备号" & _
             "       from 影像设备目录 a, 影像检查记录 b " & _
             "       Where a.设备号 = B.位置二 " & _
             "       Union All " & _
             "       select IP地址, FTP目录, FTP用户名, FTP密码, 医嘱ID, 发送号, 检查UID, 位置三 as 位置, 接收日期, a.设备号 " & _
             "       from 影像设备目录 a, 影像检查记录 b " & _
             "       Where a.设备号 = B.位置三 " & _
             "       ) a , 影像检查序列 b , 影像检查图象 c " & _
             " Where a.检查uid = B.检查uid " & _
             " and b.序列uid = c.序列uid " & _
             " and a.医嘱ID = [1] And 发送号 = [2] "
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "查图", lng医嘱ID, lng发送号)
    Do Until rsTmp.EOF
        If strDeviceNO <> Nvl(rsTmp("设备号")) Then
            strDeviceNO = Nvl(rsTmp("设备号"))
            Inte.FuncFtpConnect Nvl(rsTmp("IP地址")), Nvl(rsTmp("FTP用户名")), Nvl(rsTmp("FTP密码"))
        End If
        Inte.FuncDelFile IIf(IsNull(rsTmp("FTP目录")), "", rsTmp("FTP目录") & "/") & Format(rsTmp("接收日期"), "YYYYMMDD") & "/" & rsTmp("检查UID"), rsTmp("图像UID")
        rsTmp.MoveNext
    Loop
    strDeviceNO = ""
    Inte.FuncFtpDisConnect
    '删除目录
    strSQL = "select IP地址,FTP目录,FTP用户名,FTP密码,医嘱ID,发送号,检查UID,设备号,位置,接收日期 from " & _
             "      (select IP地址,FTP目录,FTP用户名,FTP密码,医嘱ID,发送号,检查UID,a.设备号,位置一 as 位置,接收日期 from 影像设备目录 a , 影像检查记录 b " & _
             "      Where a.设备号 = B.位置一 " & _
             "      Union All " & _
             "      select IP地址,FTP目录,FTP用户名,FTP密码,医嘱ID,发送号,检查UID,a.设备号,位置二 as 位置,接收日期 from 影像设备目录 a , 影像检查记录 b " & _
             "      Where a.设备号 = B.位置二 " & _
             "      Union All " & _
             "      select IP地址,FTP目录,FTP用户名,FTP密码,医嘱ID,发送号,检查UID,a.设备号,位置三 as 位置,接收日期 from 影像设备目录 a , 影像检查记录 b " & _
             "      where a.设备号 = b.位置三 ) a " & _
             " Where a.医嘱ID = [1] And 发送号 = [2] "
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "查目录", lng医嘱ID, lng发送号)
    Do Until rsTmp.EOF
        If strDeviceNO <> Nvl(rsTmp("设备号")) Then
            strDeviceNO = Nvl(rsTmp("设备号"))
            Inte.FuncFtpConnect Nvl(rsTmp("IP地址")), Nvl(rsTmp("FTP用户名")), Nvl(rsTmp("FTP密码"))
        End If
        Inte.FuncFtpDelDir IIf(IsNull(rsTmp("FTP目录")), "", rsTmp("FTP目录")), Format(rsTmp("接收日期"), "YYYYMMDD") & "/" & rsTmp("检查UID")
        rsTmp.MoveNext
    Loop
    Inte.FuncFtpDisConnect
    Exit Sub
ProcError:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
Public Function MovedByDate(ByVal vDate As Date) As Boolean
'功能：判断指定日期之前的是否可能已经执行了数据转出
'参数：vDate=时间点或时间段的开始时间
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    strSQL = "Select 上次日期 From zlDataMove Where 系统=[1] And 组号=1 And 上次日期 is Not Null"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISWork", glngSys)
    If Not rsTmp.EOF Then
        '上次日期没有时点,"<"判断与转出过程中一致
        If vDate < rsTmp!上次日期 Then
            MovedByDate = True
        End If
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Public Function GetFullNO(ByVal strNO As String, ByVal intNUM As Integer) As String
'功能：由用户输入的部份单号，返回全部的单号。
'参数：intNum=项目序号,为0时固定按年产生
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, intType As Integer
    Dim curDate As Date
    
    If Len(strNO) >= 8 Then
        GetFullNO = Right(strNO, 8)
        Exit Function
    ElseIf Len(strNO) = 7 Then
        GetFullNO = PreFixNO & strNO
        Exit Function
    ElseIf intNUM = 0 Then
        GetFullNO = PreFixNO & Format(Right(strNO, 7), "0000000")
        Exit Function
    End If
    GetFullNO = strNO
    
    strSQL = "Select 编号规则,Sysdate as 日期 From 号码控制表 Where 项目序号=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlExpense", intNUM)
    If Not rsTmp.EOF Then
        intType = Nvl(rsTmp!编号规则, 0)
        curDate = rsTmp!日期
    End If

    If intType = 1 Then
        '按日编号
        strSQL = Format(CDate("1992-" & Format(rsTmp!日期, "MM-dd")) - CDate("1992-01-01") + 1, "000")
        GetFullNO = PreFixNO & strSQL & Format(Right(strNO, 4), "0000")
    Else
        '按年编号
        GetFullNO = PreFixNO & Format(Right(strNO, 7), "0000000")
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
Public Function InitSysPar() As Boolean
'初始化全局参数
    Dim strValue As String
    On Error Resume Next
    
    gstrSysName = GetSetting("ZLSOFT", "注册信息", "gstrSysName", "")
    strValue = zlDatabase.GetPara("输入法")
    gstrIme = IIf(strValue = "", "不自动开启", strValue)
    
    strValue = zlDatabase.GetPara(20, glngSys, , "7|7|7|7|7")
    gbytCardNOLen = Val(Split(strValue, "|")(4)) '就诊卡号长度
    
        '费用金额小数点位数
    gbytDec = Val(zlDatabase.GetPara(9, glngSys, , 2))
    gstrDec = "0." & String(gbytDec, "0")
    
    If Not GetUserInfo Then
        MsgBox "当前用户未设置对应的人员信息,请与系统管理员联系,先到用户授权管理中设置！", vbInformation, gstrSysName
        Exit Function
    End If

    gstrUnitName = GetUnitName
    
    InitSysPar = True
End Function
Public Function MergeImageFiles(ByVal strCurrUID As String, ByVal strNewUID As String, _
    Optional ByVal strReceiveDate As String = "", Optional ByVal strMoveFiles As String = "") As Boolean
'------------------------------------------------
'功能：将一个检查的影像文件转移到另外检查中去，支持影像关联和取消关联
'参数： strCurrUID －－源检查UID
'       strNewUID －－转移后新的目的检查UID
'       strReceiveDate －－ 接收日期，用来创建图像存储路径，当strNewUID不在数据库中时，才需要使用本参数
'       strMoveFiles －－ 需要移动的文件名列表，使用"|"分隔文件名，如果没有，则转移源检查UID指向的目录下的所有图像
'返回：True--成功；False－失败
'------------------------------------------------
    Dim objSrcFtp As New clsFtp, objDestFtp As New clsFtp
    Dim strSrcPath As String, strDestPath As String
    Dim rsTmp As New ADODB.Recordset, strSQL As String, strTmpFile As String
    Dim aFiles() As String, i As Integer, objFile As New Scripting.FileSystemObject
    Dim strFTPUser As String, strFTPPassw As String, strFTPHost As String, strFTPRoot As String
    Dim lngResult As Long       '记录FTP操作的结果
        
    '如果新检查UID＝旧检查UID，则认为合并完成，并直接退出
    If strCurrUID = strNewUID Then
        MergeImageFiles = True
        Exit Function
    End If
    
    On Error GoTo errH

    '根据移动的方向不同，源图有可能在“影像临时记录”或者“影像检查记录”中
    '关联时从临时记录搬移到正常记录，取消关联时从正常记录搬移到临时记录
    
    strSQL = "Select D.FTP用户名 As FtpUser,D.FTP密码 As FtpPwd," & _
        "D.IP地址 As Host," & _
        "'/'||D.Ftp目录||'/' As Root,Decode(C.接收日期,Null,'',to_Char(C.接收日期,'YYYYMMDD')||'/')" & _
        "||C.检查UID As URL " & _
        "From 影像检查记录 C,影像设备目录 D " & _
        "Where Decode(C.位置一,Null,C.位置二,C.位置一)=D.设备号(+)" & _
        "And C.检查UID= [1] Union All " & _
        "Select D.FTP用户名 As FtpUser,D.FTP密码 As FtpPwd," & _
        "D.IP地址 As Host," & _
        "'/'||D.Ftp目录||'/' As Root,Decode(C.接收日期,Null,'',to_Char(C.接收日期,'YYYYMMDD')||'/')" & _
        "||C.检查UID As URL " & _
        "From 影像临时记录 C,影像设备目录 D " & _
        "Where Decode(C.位置一,Null,C.位置二,C.位置一)=D.设备号(+)" & _
        "And C.检查UID= [1]"
    '在数据库中查询旧检查UID
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "ZLPACSWork", strCurrUID)
    '当前检查UID在数据库中不存在，则退出本程序
    If rsTmp.EOF Then
        Exit Function
    End If
    
    '存储并创建FTP连接设置
    strFTPHost = Nvl(rsTmp("Host"))
    strFTPPassw = Nvl(rsTmp("FtpPwd"))
    strFTPRoot = Nvl(rsTmp("Root"))
    strFTPUser = Nvl(rsTmp("FtpUser"))
    strSrcPath = Nvl(rsTmp("Root")) & Nvl(rsTmp("URL"))
    lngResult = objSrcFtp.FuncFtpConnect(strFTPHost, strFTPUser, strFTPPassw)
    If lngResult = 0 Then Exit Function     'FTP连接失败，退出程序
    
    '在数据库中查询新检查UID，初始化目标Ftp,如果目标UID不存在，创建一个新路径
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "ZLPACSWork", strNewUID)
    If rsTmp.EOF Then
    '从正常图像转成临时图像的时候，目的检查UID暂时不会出现在数据库中，此时直接使用原有的FTP连接
    '在向数据库中转移记录的时候，还会使用原来的FTP连接
        If strReceiveDate <> "" Then
                objDestFtp.FuncFtpConnect strFTPHost, strFTPUser, strFTPPassw
                strDestPath = strFTPRoot & Format(strReceiveDate, "YYYYMMDD") & "/" & strNewUID
                '创建FTP目录
                objDestFtp.FuncFtpMkDir strFTPRoot, Format(strReceiveDate, "YYYYMMDD") & "/" & strNewUID
        Else
            Exit Function
        End If
    Else
        objDestFtp.FuncFtpConnect Nvl(rsTmp("Host")), Nvl(rsTmp("FtpUser")), Nvl(rsTmp("FtpPwd"))
        strDestPath = Nvl(rsTmp("Root")) & Nvl(rsTmp("URL"))
    End If
    
    '提取需要移动的文件名
    If strMoveFiles <> "" Then
        aFiles = Split(strMoveFiles, "|")
    Else
        aFiles = Split(objSrcFtp.FuncDirFiles(strSrcPath), "|")
    End If
    
    '先转移图像
    For i = 0 To UBound(aFiles)
        strTmpFile = App.Path & "\TmpImage\" & aFiles(i)
        lngResult = objSrcFtp.FuncDownloadFile(strSrcPath, strTmpFile, aFiles(i))
        If lngResult <> 0 Then Exit Function
        lngResult = objDestFtp.FuncUploadFile(strDestPath, strTmpFile, aFiles(i))
        If lngResult <> 0 Then Exit Function
    Next i
    
    '转移图像成功后，在删除临时图像和原有FTP的图像和目录，清场操作出现错误可以不处理
    On Error Resume Next
    For i = 0 To UBound(aFiles)
        strTmpFile = App.Path & "\TmpImage\" & aFiles(i)
        Kill strTmpFile
        Call objSrcFtp.FuncDelFile(strSrcPath, aFiles(i))
    Next i
    Call objSrcFtp.FuncFtpDelDir(Replace(strSrcPath, strCurrUID, ""), strCurrUID)
    
    objSrcFtp.FuncFtpDisConnect
    objDestFtp.FuncFtpDisConnect
    MergeImageFiles = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Sub MkLocalDir(ByVal strDir As String)
'------------------------------------------------
'功能：创建本地目录
'参数： strDir－－本地目录
'返回：无
'------------------------------------------------
    Dim objFile As New Scripting.FileSystemObject
    Dim aNestDirs() As String, i As Integer
    Dim strPath As String
    On Error Resume Next
    
    '读取全部需要创建的目录信息
    ReDim Preserve aNestDirs(0)
    aNestDirs(0) = strDir
    
    strPath = objFile.GetParentFolderName(strDir)
    Do While Len(strPath) > 0
        ReDim Preserve aNestDirs(UBound(aNestDirs) + 1)
        aNestDirs(UBound(aNestDirs)) = strPath
        strPath = objFile.GetParentFolderName(strPath)
    Loop
    '创建全部目录
    For i = UBound(aNestDirs) To 0 Step -1
        MkDir aNestDirs(i)
    Next
End Sub


Public Sub ClearCacheFolder(ByVal strCacheFolder As String)
'------------------------------------------------
'功能：当指定目录的大小达到一定百分比时，清空该目录
'参数： strCacheFolder--需要检查是否清空的目录
'返回：无
'------------------------------------------------
    Dim objFile As New Scripting.FileSystemObject
    Dim objCurFolder As Scripting.Folder, objCurFile As Scripting.File, objFiles As Scripting.Files
    Dim strDriver As String
    
    On Error Resume Next
    strDriver = objFile.GetDriveName(strCacheFolder)
    Set objCurFolder = objFile.GetFolder(strCacheFolder)
    If objCurFolder.Size / objFile.GetDrive(strDriver).FreeSpace > 0.2 Then
        objCurFolder.Delete True
    End If
End Sub

Public Function GetTrayHeight() As Long
    '------------------------------------------------------------------------------------------------------------------
    '功能:获取任务栏的高度
    '------------------------------------------------------------------------------------------------------------------
    Dim lngHwd As Long
    Dim objRect As RECT
    
    On Error Resume Next
    
    lngHwd = FindWindow("shell_traywnd", "")
    Call GetWindowRect(lngHwd, objRect)

    GetTrayHeight = Screen.TwipsPerPixelX * (objRect.Bottom - objRect.Top)
    
    If GetTrayHeight < 0 Then GetTrayHeight = 0
    
End Function

Public Sub ResizeRegion(ByVal ImageCount As Integer, ByVal RegionWidth As Long, _
    ByVal RegionHeight As Long, Rows As Integer, Cols As Integer)
'-----------------------------------------------------------------------------
'功能：根据输入的图像数量，图像区域的宽度和高度，返回最佳的图像排列行数和列数
'参数： ImageCount－－图像数量
'       RegionWidth--图像显示区域的宽度
'       RegionHeight--图像显示区域的高度
'       Rows－－[返回]最佳行数
'       Cols－－[返回]最佳列数
'返回：返回最佳行数Rows，最佳列数Cols
'-----------------------------------------------------------------------------
    Dim iCols As Integer, iRows As Integer
    Dim iBase As Integer, blnDoLoop As Integer
    
    On Error GoTo err
    iCols = CInt(Sqr(ImageCount * RegionWidth / RegionHeight))
    iRows = CInt(Sqr(ImageCount * RegionHeight / RegionWidth))

    If iRows < 1 Then iRows = 1
    If iCols < 1 Then iCols = 1
    Do While iRows * iCols > ImageCount
        If RegionWidth / RegionHeight > 1 Then
            iCols = iCols - 1
        Else
            iRows = iRows - 1
        End If
    Loop
    
    If iRows < 1 Then iRows = 1
    If iCols < 1 Then iCols = 1
    Do While iRows * iCols < ImageCount
        If RegionWidth / RegionHeight > 1 Then
            iCols = iCols + 1
        Else
            iRows = iRows + 1
        End If
    Loop
    Rows = iRows: Cols = iCols
    
    If ImageCount <> 0 Then
        If Rows * Cols > ImageCount Then
            iBase = 6
            blnDoLoop = True
            
            While blnDoLoop
                iBase = iBase - 1
                
                If ImageCount Mod iBase = 0 Then
                    blnDoLoop = False
                End If
            Wend
        

            If RegionWidth > RegionHeight Then
                If ImageCount / iBase > iBase Then
                    Cols = ImageCount / iBase
                    Rows = iBase
                Else
                    Rows = ImageCount / iBase
                    Cols = iBase
                End If
            Else
                If ImageCount / iBase > iBase Then
                    Cols = iBase
                    Rows = ImageCount / iBase
                Else
                    Rows = iBase
                    Cols = ImageCount / iBase
                End If
            End If
        End If
    End If
err:
End Sub


Public Function funGetStudyUID(ByVal strOldStudyUID As String) As String
'-----------------------------------------------------------------------------
'功能:查询数据库，判断当前图像的检查UID是否已经存在于正常表和临时表中，
'     如果存在，则在检查UID后面增加后缀，不存在则直接返回输入的检查UID
'修改人:黄捷
'修改日期:2007-1-27
'-----------------------------------------------------------------------------
    '
    Dim rsMatch As New ADODB.Recordset
    
    funGetStudyUID = strOldStudyUID
    gstrSQL = "select 检查UID from 影像检查记录 where 检查UID = [1]" & _
              " Union All Select 检查UID from 影像临时记录 where 检查UID = [1]"
    Set rsMatch = zlDatabase.OpenSQLRecord(gstrSQL, "PACS图像保存", strOldStudyUID)
    If Not rsMatch.EOF Then
        '创建一个新的检查UID
        gstrSQL = "Select 影像检查UID序号_ID.Nextval From Dual"
        Set rsMatch = zlDatabase.OpenSQLRecord(gstrSQL, "PACS图像保存")
        If Len(strOldStudyUID) <= 55 Then
            funGetStudyUID = strOldStudyUID & ".A" & rsMatch(0)
        Else
            funGetStudyUID = Left(strOldStudyUID, 55) & ".A" & rsMatch(0)
        End If
    End If
End Function


Public Function GetImageAttribute(objAttr As DicomAttributes, ByVal AttrName As String) As Variant
'-----------------------------------------------------------------------------
'功能:提取DICOM属性集中的指定属性值
'修改人:黄捷
'修改日期:2007-2-6
'-----------------------------------------------------------------------------
    Dim AttrTag() As String
    
    GetImageAttribute = ""
    AttrTag = Split(AttrName, ":")
    If objAttr("&h" & AttrTag(0), "&h" & AttrTag(1)).Exists Then
        GetImageAttribute = Nvl(objAttr("&h" & AttrTag(0), "&h" & AttrTag(1)).Value)
    End If
End Function

'Public Function funRelateSeries(lng医嘱ID As Long, lng发送号 As Long)
''-----------------------------------------------------------------------------
''功能:关联图像，移动FTP图像到新的位置，修改数据库记录，从临时表转到正式表中
''参数： lng医嘱ID －－医嘱ID
''       lng发送号 －－ 发送号
''返回：无
''-----------------------------------------------------------------------------
'    Dim blnCancel As Boolean, rsTmp As ADODB.Recordset
'    Dim rsStudyUID As ADODB.Recordset
'    Dim strFilter As String
'    Dim strModality As String
'
'    On Error GoTo errHandle
'
'    gstrSQL = "Select 影像类别 From 影像检查记录 Where 医嘱ID= [1] And 发送号 = [2]"
'    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "关联图像，提取影像类被", lng医嘱ID, lng发送号)
'    strModality = Nvl(rsTmp!影像类别)
'
'    gstrSQL = "Select 0 as 选择, A.检查UID As ID,Nvl(A.姓名,' ') As 姓名,Nvl(A.英文名,' ') As 英文名," & _
'            "Nvl(A.检查号,0) As 检查号,Nvl(A.性别,' ') As 性别,Nvl(A.年龄,' ') As 年龄," & _
'            "Nvl(A.检查设备,' ') As 检查设备,to_char(Nvl(A.接收日期,Sysdate),'YYYY-MM-DD hh24:mi:ss') As 检查时间," & _
'            "to_char(A.出生日期,'YYYY-MM-DD') As 出生日期," & _
'            "Nvl(A.身高,0) As 身高,Nvl(A.体重,0) As 体重, c.序列描述,a.影像类别 " & _
'            " From 影像临时记录 a," & _
'            "(Select x.序列描述,x.检查uid, row_number() over(partition by 检查UID order by 检查UID) As  rank from 影像临时序列 x) c " & _
'            " Where a.检查UID = c.检查UID And c.rank = 1"
'
'    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "关联图像", lng医嘱ID, lng发送号)
'
'    frmSelectMuli.ShowSelect rsTmp, "ID,900,0,1;影像类别,900,0,1;姓名,800,0,1;英文名,800,0,1;检查号,900,0,1;" _
'            & "性别,600,0,1;年龄,600,0,1;序列描述,1200,0,1;检查设备,900,0,1;检查时间,1200,0,1;出生日期,1200,0,1;" _
'            & "身高,500,0,1;体重,500,0,1", 0, 0, 14000, 10000, "关联图像", , , strModality
'
'    If frmSelectMuli.mblnOK = True And frmSelectMuli.strFilter <> "ID=-1" Then
'
'        If MsgBox("是否确认选择的影像是当前检查的？", vbDefaultButton2 + vbQuestion + vbYesNo, gstrSysName) = vbNo Then Exit Function
'        strFilter = frmSelectMuli.strFilter
'        rsTmp.Filter = strFilter
'        '如果有选中的临时纪录，则处理每一临时纪录的关联
'        While Not rsTmp.EOF
'            '移动Ftp上的影像文件,移动成功，才更新数据库
'            gstrSQL = "Select 检查UID From 影像检查记录 Where 医嘱ID=[1] And 发送号=[2]"
'            Set rsStudyUID = zlDatabase.OpenSQLRecord(gstrSQL, "关联影像", lng医嘱ID, lng发送号)
'            If Not rsStudyUID.EOF Then
'                If Len(Trim(Nvl(rsStudyUID(0)))) > 0 Then
'                    If MergeImageFiles(rsTmp("ID"), rsStudyUID(0)) = False Then
'                        MsgBox "文件转移失败，不能关联影像。" & vbCrLf & vbCrLf & "可能是网络连接有问题，请检查。"
'                        Exit Function
'                    End If
'                End If
'            End If
'
'            gstrSQL = "ZL_影像检查_SET(" & lng医嘱ID & "," & lng发送号 & ",'" & _
'                rsTmp("ID") & "')"
'            zlDatabase.ExecuteProcedure gstrSQL, "关联影像"
'
'            rsTmp.MoveNext
'        Wend
'    End If
'
'    Exit Function
'errHandle:
'    If ErrCenter() = 1 Then Resume
'    Call SaveErrLog
'End Function
Public Function SetDeptPara(ByVal lngDeptID As Long, ByVal varPara As String, ByVal strValue As String) As Boolean
'功能：设置指定的参数值
'参数：lngDept=科室ID
'      varPara=参数名
'      strValue=参数名值
'返回：设置是否成功
    Dim strSQL As String
    
    On Error GoTo errH
        
    strSQL = "ZL_影像流程参数_UPDATE(" & lngDeptID & ",'" & varPara & "','" & strValue & "')"
    Call zlDatabase.ExecuteProcedure(strSQL, "SetPara")
    
    '设置成功后清除缓存
    Set mrsDeptParas = Nothing
    
    SetDeptPara = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
End Function
Public Function GetDeptPara(ByVal lngDeptID As Long, ByVal varPara As String, Optional ByVal strDefault As String, Optional ByVal blnNotCache As Boolean) As String
'功能：读取指定的参数值
'参数：lngDept=科室ID
'      varPara=参数名
'      strDefault=当数据库中没有该参数时使用的缺省值(注意不是为空时)
'      blnNotCache=是否不从缓存中读取
'返回：参数值，字符串形式
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, blnNew As Boolean
    
    On Error GoTo errH
    
    If blnNotCache Then
        Set rsTmp = New ADODB.Recordset
        strSQL = "Select 参数值 from 影像流程参数 where 科室ID = [1] and 参数名=[2]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "读取参数", lngDeptID, varPara)
        
        If Not rsTmp.EOF Then
            GetDeptPara = Nvl(rsTmp!参数值)
        Else
            GetDeptPara = strDefault
        End If
    Else
        '第一次加载参数缓存
        If mrsDeptParas Is Nothing Then
            blnNew = True
        ElseIf mrsDeptParas.State = 0 Then
            blnNew = True
        End If
        If blnNew Then
            strSQL = "Select 参数值,参数名,科室ID from 影像流程参数"
            Set mrsDeptParas = New ADODB.Recordset
            Set mrsDeptParas = zlDatabase.OpenSQLRecord(strSQL, "读取参数")
        End If
        
        '根据缓存读取参数值
        mrsDeptParas.Filter = "参数名='" & CStr(varPara) & "' AND 科室ID=" & lngDeptID
        If Not mrsDeptParas.EOF Then
            GetDeptPara = Nvl(mrsDeptParas!参数值)
        Else
            GetDeptPara = strDefault
        End If
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
End Function

Public Function GetIsValidOfStorageDevice(ByVal lngDeptID As Long) As Boolean
'初始化科室级参数
    Dim rsTmp As New ADODB.Recordset
    Dim strSaveDeviceId As String
    
    On Error GoTo DBError
    
    '读取并检测存储设备号
    strSaveDeviceId = GetDeptPara(lngDeptID, "存储设备号")
    
    gstrSQL = "Select 设备号,设备名 From 影像设备目录 Where 类型=1 and 设备号=[1] and NVL(状态,0)=1"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "获取存储设备信息", strSaveDeviceId)
    
    
    GetIsValidOfStorageDevice = Not rsTmp.EOF
    
    Exit Function
DBError:
    GetIsValidOfStorageDevice = False
    
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function

Public Sub subCancelSeriesRelate(lngAdviceNo As Long, lngSendNO As Long, strSeriesNo As String)
'-----------------------------------------------------------------------------
'功能:取消序列图象的关联，移动FTP图像到新的位置，修改数据库记录，从正式表移动到临时表中
'参数： lngAdviceNo －－医嘱ID
'       lngSendNO －－ 发送号
'       strSeriesNo －－序列UID
'返回：无
'-----------------------------------------------------------------------------
    
    Dim mcnFTP As New clsFtp
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim strCachePath As String
    Dim strCacheFileName As String
    Dim objFile As New Scripting.FileSystemObject
    Dim imgs As New DicomImages
    Dim img As New DicomImage
    Dim strNewStudyUID As String    '新生成的检查UID
    Dim strOldStudyUID As String    '图象里面原来的检查UID
    Dim strDBStudyUID As String     '数据库中保存的检查UID，跟图象存储路径相关
    Dim strMoveFiles As String  '存储需要移动的图象文件名，使用“|”分隔
    Dim blnNoImage As Boolean   '1没有图象，直接读取数据库信息。0有图象，使用图象信息
    Dim lngResult As Long    '记录FTP返回结果
    
    '图像中的病人基本信息
    Dim strModality As String
    Dim strPatientID As String
    Dim strPatientName As String
    Dim strSex As String
    Dim strAge As String
    Dim strDateOfBirth As String
    Dim strManufacturer As String
    Dim strReceiveDateTime As String
    
    
    On Error GoTo DBError
    
    '查找序列中第一个图像的 病人ID，英文名，性别，年龄，出生日期，检查UID，检查设备，接收时间
    strCachePath = App.Path & "\TmpImage\"
    strSQL = "Select A.图像号,D.FTP用户名 As User1,D.FTP密码 As Pwd1,a.图像UID, " & _
        "D.IP地址 As Host1,c.检查uid," & _
        "'/'||D.Ftp目录||'/' As Root1,Decode(C.接收日期,Null,'',to_Char(C.接收日期,'YYYYMMDD')||'/')" & _
        "||C.检查UID||'/'||A.图像UID As URL,d.设备号 as 设备号1, " & _
        "E.FTP用户名 As User2,E.FTP密码 As Pwd2," & "E.IP地址 As Host2," & _
        "'/'||E.Ftp目录||'/' As Root2,e.设备号 as 设备号2 " & _
        "From 影像检查图象 A,影像检查序列 B,影像检查记录 C,影像设备目录 D,影像设备目录 E " & _
        "Where A.序列UID=B.序列UID And B.检查UID=C.检查UID And C.位置一=D.设备号(+) And C.位置二=E.设备号(+) " & _
        "And A.序列UID= [1] Order By A.图像号"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "查数据", strSeriesNo)
    
    If Not rsTmp.EOF Then   '序列中存在图象
        strDBStudyUID = Nvl(rsTmp("检查uid"))
        '新建本地目录
        strCacheFileName = strCachePath & rsTmp("URL")
        MkLocalDir objFile.GetParentFolderName(strCacheFileName)
        
        '下载图象
        If rsTmp("设备号1") <> "" And mcnFTP.FuncFtpConnect(Nvl(rsTmp("Host1")), Nvl(rsTmp("User1")), Nvl(rsTmp("Pwd1"))) <> 0 Then
            mcnFTP.FuncDownloadFile objFile.GetParentFolderName(Nvl(rsTmp("Root1")) & rsTmp("URL")), strCacheFileName, objFile.GetFileName(rsTmp("URL"))
            mcnFTP.FuncFtpDisConnect
        ElseIf rsTmp("设备号2") <> "" And mcnFTP.FuncFtpConnect(Nvl(rsTmp("Host2")), Nvl(rsTmp("User2")), Nvl(rsTmp("Pwd2"))) <> 0 Then
            mcnFTP.FuncDownloadFile objFile.GetParentFolderName(Nvl(rsTmp("Root2")) & rsTmp("URL")), strCacheFileName, objFile.GetFileName(rsTmp("URL"))
            mcnFTP.FuncFtpDisConnect
        Else
            'FTP连接错误，提示并退出本次取消关联操作
            MsgBox "FTP连接错误，不能取消关联。" & vbCrLf & vbCrLf & "可能是网络连接出现问题。"
            Exit Sub
        End If
                    
        '读取图象信息
        If Dir(strCacheFileName) <> vbNullString Then
            Set img = imgs.ReadFile(strCacheFileName)
            '使用变量将图象基本信息读取出来
            strOldStudyUID = img.StudyUID
            strModality = GetImageAttribute(img.Attributes, ATTR_影像类别)
            strPatientID = img.PatientID
            strPatientName = img.Name
            strSex = img.Sex
            If IsDate(img.DateOfBirthAsDate) Then
                If img.Attributes(&H10, &H1010).Exists And Not IsNull(img.Attributes(&H10, &H1010)) Then
                    strAge = img.Attributes(&H10, &H1010).Value
                Else
                    strAge = CStr(Year(Date) - Year(img.DateOfBirthAsDate))
                End If
                        
                If img.DateOfBirthAsDate <> "0:00:00" Then
                    strDateOfBirth = Format(img.DateOfBirthAsDate, "YYYY-MM-DD")
                Else
                    strDateOfBirth = ""
                End If
            Else
                strAge = "": strDateOfBirth = ""
            End If
            strManufacturer = GetImageAttribute(img.Attributes, ATTR_检查设备)
            strReceiveDateTime = GetImageAttribute(img.Attributes, ATTR_检查日期) & " " & _
                        Format(GetImageAttribute(img.Attributes, ATTR_检查时间), "HH:MM")
            '删除临时图象
            Set img = Nothing
            imgs.Remove (1)
            On Error Resume Next
            objFile.DeleteFile strCacheFileName
            On Error GoTo 0
        Else
            '如果第一个图象下载不正确，读取数据库信息，这种情况存在吗？
            blnNoImage = True
        End If
    Else
        '序列中没有图象，不需要取消关联，应该不会存在这种情况
        Exit Sub
    End If
    
    '对于没有图象信息可读取，或者图像重要信息读取不完整的，直接读取数据库中的信息
    If blnNoImage = True Or Trim(strReceiveDateTime) = "" Then
        strSQL = "select a.影像类别,a.检查号,a.姓名,a.英文名,a.性别,a.年龄,a.出生日期,a.检查uid," & _
                " a.检查设备,a.接收日期 from 影像检查记录 a where a.医嘱id =[1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "提取数据", lngAdviceNo)
        If Not rsTmp.EOF Then
            If blnNoImage = True Then
                strOldStudyUID = Nvl(rsTmp("检查uid"))
                strDBStudyUID = Nvl(rsTmp("检查uid"))
                strPatientID = Nvl(rsTmp("检查号"))
                strPatientName = Nvl(rsTmp("英文名"))
                strSex = Nvl(rsTmp("性别"))
                strAge = Nvl(rsTmp("年龄"))
                strDateOfBirth = Nvl(rsTmp("出生日期"), "")
                strManufacturer = Nvl(rsTmp("检查设备"))
            End If
            strModality = Nvl(rsTmp("影像类别"))
            strReceiveDateTime = Nvl(rsTmp("接收日期"))
        End If
    End If
    '组织图象文件名称串
    strSQL = "select 图像UID from 影像检查序列 a,影像检查图象 b where a.序列UID =[1] and a.序列UID = b.序列UID"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "提取数据", strSeriesNo)
    If Not rsTmp.EOF Then
        strMoveFiles = rsTmp(0)
        rsTmp.MoveNext
        While Not rsTmp.EOF
            strMoveFiles = strMoveFiles & "|" & rsTmp(0)
            rsTmp.MoveNext
        Wend
    End If
    
    '如果检查UID跟数据库中现存的检查UID相同，则创建新的检查UID，且修改图像FTP路径
    strNewStudyUID = funGetStudyUID(strOldStudyUID)
    If strNewStudyUID <> strDBStudyUID Then
        If MergeImageFiles(strDBStudyUID, strNewStudyUID, Format(strReceiveDateTime, "YYYY-MM-DD"), strMoveFiles) = False Then
            MsgBox "图像转移不成功，不能取消关联。"
            Exit Sub
        End If
    End If
    
    '修改数据库，正常记录转成临时记录
    strSQL = "ZL_影像检查_PhotoCancel(" & lngAdviceNo & "," & lngSendNO & ",'" & strNewStudyUID & "','" & _
              strSeriesNo & "','" & strModality & "'," & Val(strPatientID) & ",'" & _
              strPatientName & "','" & strSex & "','" & strAge & "'," & _
              IIf(Len(strDateOfBirth) = 0, "null", "to_date('" & strDateOfBirth & "','YYYY-MM-DD')") & _
              ",'" & strManufacturer & "',to_date('" & strReceiveDateTime & "','YYYY-MM-DD HH24:MI:SS'))"
              
    zlDatabase.ExecuteProcedure strSQL, "取消关联"
    
    Exit Sub
DBError:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub


Public Sub GetAllImages(dcmViewer As DicomViewer, blnMoved As Boolean, intSearchType As Integer, _
    Optional lngAdviceID As Long, Optional strSeriesUID As String, Optional intGetImgNum As Integer = 0, _
    Optional intShowImgNum As Integer = 0, Optional blnTempDB As Boolean = False, _
    Optional strStudyUID As String = "", Optional strImageUID As String = "")
'------------------------------------------------
'功能：删除dcmViewer中的图像后，将读取的图像文件放入dcmViewer中
'参数： dcmViewer－－打开图像的DicomViewer
'       blnMoved －－ 是否被转储了
'       intSearchType －－检索类型,只对正式表查询有效  1－按照医嘱ID和发送号查，2－按照序列UID查，3 - 按照图像UID查
'       lngAdviceID －－ 医嘱ID
'       strSeriesUID －－ 序列UID
'       intGetImgNum －－本次读取的图像数量
'       intShowImgNum －－本次显示的图像数量
'       blnTempDB - - 是否从临时表中提取图像
'       strStudyUID - - 检查UID,只有从临时表查找的时候，才使用这个参数
'       strImageUID - - 图像UID，只有从正式表查找的时候，才使用这个参数
'返回：无，直接修改dcmViewer中显示的图像
'------------------------------------------------
    
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
    Dim curImage As DicomImage, i As Integer
    Dim iCols As Integer, iRows As Integer
    Dim objFile As New Scripting.FileSystemObject, strTmpFile As String
    Dim Inet1 As New clsFtp
    Dim Inet2 As New clsFtp
    Dim strCachePath As String
    Dim iCurrentIndex As Integer
    
    On Error GoTo DBError
    If blnTempDB = False Then       '从正式图像库中查找图像
        strSQL = "Select /*+RULE*/ A.图像号,D.FTP用户名 As User1,D.FTP密码 As Pwd1," & _
            "D.IP地址 As Host1,'/'||D.Ftp目录||'/' As Root1," & _
            "Decode(C.接收日期,Null,'',to_Char(C.接收日期,'YYYYMMDD')||'/')" & _
            "||C.检查UID||'/'||A.图像UID As URL,d.设备号 as 设备号1, " & _
            "E.FTP用户名 As User2,E.FTP密码 As Pwd2," & _
            "E.IP地址 As Host2,'/'||E.Ftp目录||'/' As Root2," & _
            "e.设备号 as 设备号2,C.检查UID,B.序列UID " & _
            "From 影像检查图象 A,影像检查序列 B,影像检查记录 C,影像设备目录 D,影像设备目录 E " & _
            "Where A.序列UID=B.序列UID And B.检查UID=C.检查UID And C.位置一=D.设备号(+) And C.位置二=E.设备号(+) "
        If blnMoved Then
            strSQL = Replace(strSQL, "影像检查图象", "H影像检查图象")
            strSQL = Replace(strSQL, "影像检查序列", "H影像检查序列")
            strSQL = Replace(strSQL, "影像检查记录", "H影像检查记录")
        End If
        If intShowImgNum <> 0 Then
            strSQL = strSQL & " And Rownum<=[2] "
        End If
        
        If intSearchType = 1 Then       '1－按照医嘱ID和发送号查
            strSQL = strSQL & "And C.医嘱ID=[1] Order By A.图像号"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "读取图像", lngAdviceID, intGetImgNum)
        ElseIf intSearchType = 2 Then   '2－按照序列UID查
            strSQL = strSQL & "And A.序列UID= [1] Order By A.图像号"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "读取图像", strSeriesUID, intGetImgNum)
        ElseIf intSearchType = 3 Then   '3 - 按照图像UID查
            strSQL = strSQL & "And A.图像UID = [1] "
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "读取图像", strImageUID, intGetImgNum)
        End If
        
    Else                '从临时表中查找图像
        
        strSQL = "Select /*+RULE*/ c.图像号,d.FTP用户名 As User1, d.FTP密码 As Pwd1, d.Ip地址 As Host1," _
                & "'/' || d.Ftp目录 || '/' As Root1," _
                & " Decode(a.接收日期, Null, '', To_Char(a.接收日期, 'YYYYMMDD') || '/') || a.检查uid || '/' || c.图像uid As URL," _
                & " d.设备号 As 设备号1,a.检查UID,b.序列UID,d.FTP用户名 As User2, d.FTP密码 As Pwd2, " _
                & " d.Ip地址 As Host2, '/' || d.Ftp目录 || '/' As Root2, " _
                & " d.设备号 As 设备号2 " _
                & " From 影像临时记录 a,影像临时序列 b,影像临时图象 c ,影像设备目录 d " _
                & " Where a.检查UID=b.检查UID And b.序列UID = c.序列UID And a.位置一 = d.设备号 "
                
        If strStudyUID <> "" Then   '按照检查uid查找
            strSQL = strSQL & "And a.检查UID=[1] and Rownum<=[2] Order By c.图像号  "
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "读取图像", strStudyUID, CLng(6))
        Else        '按照序列UID查找
            strSQL = strSQL & "And b.序列UID=[1] and Rownum<=[2] Order By c.图像号  "
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "读取图像", strSeriesUID, CLng(6))
        End If
    End If
    
        dcmViewer.Images.Clear
        If rsTmp.RecordCount > 0 Then
            If intShowImgNum = 0 Or intShowImgNum >= rsTmp.RecordCount Then
                ResizeRegion rsTmp.RecordCount, dcmViewer.Width, dcmViewer.Height, iRows, iCols
            Else
                ResizeRegion intShowImgNum, dcmViewer.Width, dcmViewer.Height, iRows, iCols
            End If
            dcmViewer.MultiColumns = iCols
            dcmViewer.MultiRows = iRows
            
            '创建本地目录
            strCachePath = App.Path & "\TmpImage\"
            MkLocalDir strCachePath & objFile.GetParentFolderName(Nvl(rsTmp("URL")))
            
            Do While Not rsTmp.EOF
                If Dir(strCachePath & Nvl(rsTmp("URL"))) = vbNullString Then
                    '本地缓存图像不存在，则读取FTP图像
                    strTmpFile = strCachePath & Nvl(rsTmp("URL"))
                    '建立FTP连接
                    If Nvl(rsTmp("设备号1")) <> vbNullString And Inet1.hConnection = 0 Then
                        If Inet1.FuncFtpConnect(Nvl(rsTmp("Host1")), Nvl(rsTmp("User1")), Nvl(rsTmp("Pwd1"))) = 0 Then
                            If Nvl(rsTmp("设备号2")) <> vbNullString Then
                                If Inet2.FuncFtpConnect(Nvl(rsTmp("Host2")), Nvl(rsTmp("User2")), Nvl(rsTmp("Pwd2"))) = 0 Then
                                    MsgBox "FTP不能正常连接，请检查网络设置。"
                                    Exit Sub
                                End If
                            Else
                                MsgBox "FTP不能正常连接，请检查网络设置。"
                                Exit Sub
                            End If
                        End If
                    End If
                    If Inet1.FuncDownloadFile(objFile.GetParentFolderName(Nvl(rsTmp("Root1")) & rsTmp("URL")), strTmpFile, objFile.GetFileName(rsTmp("URL"))) <> 0 Then
                        '从设备号1提取图像失败，则从设备号2提取图像
                        If Nvl(rsTmp("设备号2")) <> vbNullString Then
                            If Inet2.hConnection = 0 Then Inet2.FuncFtpConnect Nvl(rsTmp("Host2")), Nvl(rsTmp("User2")), Nvl(rsTmp("Pwd2"))
                            Call Inet2.FuncDownloadFile(objFile.GetParentFolderName(Nvl(rsTmp("Root2")) & rsTmp("URL")), strTmpFile, objFile.GetFileName(rsTmp("URL")))
                        End If
                    End If
                End If
                If Dir(strCachePath & Nvl(rsTmp("URL"))) <> vbNullString Then
                    Set curImage = dcmViewer.Images.ReadFile(strCachePath & Nvl(rsTmp("URL")))
                    With curImage
                        .BorderStyle = 6
                        .BorderWidth = 1
                        .BorderColour = vbWhite
                    End With
                    
                    '取消自动剪影,因为DicomObjects控件本身对处理剪影有BUG，存在（0028，6100）时，会自动对图像进行剪影，
                    '导致晋煤的DSA图像不能正常显示
                    '虽然设置图像的mask=0 ,可以取消剪影，但是每次图像被添加到新的Dicomimages之后，自动又将mask设置成1了，
                    '这样在程序中无法很好的控制，因此直接去掉（0028，6100）这个属性。
                    If Not IsNull(curImage.Attributes(&H28, &H6100).Value) Then
                        curImage.Attributes.Remove &H28, &H6100
                    End If
                End If
                
                rsTmp.MoveNext
            Loop
            If dcmViewer.Images.Count > 0 Then
                dcmViewer.CurrentIndex = 1
                dcmViewer.Images(1).BorderColour = vbRed
            End If
        Else
            dcmViewer.MultiColumns = 1
            dcmViewer.MultiRows = 1
        End If
    Inet1.FuncFtpDisConnect
    Inet2.FuncFtpDisConnect
    Exit Sub
DBError:
    If ErrCenter() = 1 Then
        Resume
    End If
    
    Call SaveErrLog
End Sub

Public Function funGetStorageDevice(strSaveDeviceId As String, ByRef strDirURL As String, ByRef strIp As String, _
        ByRef strUser As String, ByRef strPwd As String) As Boolean
'------------------------------------------------
'功能：从数据库中读取制定存储设备ID的FTP访问参数
'参数： strSaveDeviceID －－存储设备ID
'       strDirURL－－[OUT] FTP目录
'       strIp －－[OUT] IP地址
'       strUser －－ [OUT]用户名
'       strPwd －－[OUT]用户名
'返回：True－－获取成功，False－－获取失败
'-----------------------------------------------
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    
    '检查存储设备是否存在
    strSQL = "Select '/'||Decode(Ftp目录,Null,'',Ftp目录||'/') As URL,FTP用户名,FTP密码,IP地址 " & _
        "From 影像设备目录 " & "Where 设备号=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, App.ProductName, strSaveDeviceId)
     '没有存储设备时退出
    If rsTemp.EOF = True Then
        MsgBox "没有找到存储设备,请重新选择存储设备!", vbInformation, gstrSysName
        funGetStorageDevice = False
        Exit Function
    End If
    strDirURL = Nvl(rsTemp("URL"))
    strIp = Nvl(rsTemp("IP地址"))
    strUser = Nvl(rsTemp("FTP用户名"))
    strPwd = Nvl(rsTemp("FTP密码"))
    funGetStorageDevice = True
End Function

Public Function OpenViewer(ByRef objPacsCore As Object, lngAdviceID As Long, _
        blnAddImage As Boolean, objParent As Object, Optional ByVal strSerials As String = "", _
        Optional ByVal blnMoved As Boolean = False, Optional ByVal blnLocalizerBackward As Boolean = False, _
        Optional ByVal intImageInterval As Integer = 0, Optional ByVal strImageString As String = "") As Boolean
'------------------------------------------------
'功能：根据传入的医嘱ID和发送号，打开objPacsCore指向的观片站
'参数： objPacsCore －－观片站对象
'       lngAdviceID －－医嘱ID
'       blnAddImage--True 在原有图像基础上增加当前图像；False删除原有图像，打开当前图像
'       objParent -- 父窗体
'       strSerials－－可选，序列UID名称串，用逗号分隔，如果不输入，则选择全部序列
'       blnMoved－－可选，是否被转储
'       blnLocalizerBackward--可选，定位像后置,跟strImageString互斥
'       intImageInterval ---可选，打开图像的间隔，比如5，表示每5个图打开一个图,跟strImageString互斥
'       strImageString --- 可选，每个序列中需要打开的图象号组合，跟intImageInterval和blnLocalizerBackward互斥，
'                           以strImageString为主
'                           规则是“序列UID1|1-3;5-27;33-100+序列UID2|全部”,全部表示打开全部图象
'返回：图像文件名串数组
'------------------------------------------------
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
    Dim strFTPHost As String
    Dim strSDPath As String, strSDUser As String, strSDPwd As String
    Dim strDeviceNO As String
    Dim i As Integer
    Dim blnConnectDS As Boolean         '是否连接当前的共享目录
    
    On Error GoTo DBError
    strFTPHost = ""
           
    '查找需要打开的所有图象信息
    strSQL = "Select /*+RULE*/ D.IP地址 As Host1,d.设备号 as 设备号1," & _
        "Decode(C.接收日期,Null,'',to_Char(C.接收日期,'YYYYMMDD')||'/')" & _
        "||C.检查UID||'/' As Path,E.IP地址 As Host2,e.设备号 as 设备号2, " & _
        "D.共享目录 AS 共享目录1, E.共享目录 AS 共享目录2,D.共享目录用户名 as 共享目录用户名1, " & _
        "E.共享目录用户名 AS 共享目录用户名2,D.共享目录密码 AS 共享目录密码1,E.共享目录密码 AS 共享目录密码2 " & _
        "From 影像检查记录 C,影像设备目录 D,影像设备目录 E " & _
        "Where C.位置一=D.设备号(+) And C.位置二=E.设备号(+) And C.医嘱ID=[1] "
    
    '如果有转储标志，则读取转储的历史表
    If blnMoved Then
        strSQL = Replace(strSQL, "影像检查图象", "H影像检查图象")
        strSQL = Replace(strSQL, "影像检查序列", "H影像检查序列")
        strSQL = Replace(strSQL, "影像检查记录", "H影像检查记录")
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "获取共享目录信息", lngAdviceID)
    
    If rsTmp.RecordCount > 0 Then
        '创建本地的缓存目录，需要在调用观片站之前先创建这个目录，观片站中只是下载，不创建本地缓存目录
        MkLocalDir App.Path & "\TmpImage\" & rsTmp("Path")
        
        '读取FTP参数，包括用户名，密码，IP地址等
        If rsTmp("设备号1") <> "" Then
            strDeviceNO = rsTmp("设备号1")
            strFTPHost = rsTmp("Host1")
            strSDPath = Nvl(rsTmp("共享目录1"))
            strSDUser = Nvl(rsTmp("共享目录用户名1"))
            strSDPwd = Nvl(rsTmp("共享目录密码1"))
        ElseIf Nvl(rsTmp("设备号2")) <> "" Then
            strDeviceNO = rsTmp("设备号2")
            strFTPHost = rsTmp("Host2")
            strSDPath = Nvl(rsTmp("共享目录2"))
            strSDUser = Nvl(rsTmp("共享目录用户名2"))
            strSDPwd = Nvl(rsTmp("共享目录密码2"))
        End If
        
        '判断共享目录是否已经连接，如果没有连接，则进行连接
        blnConnectDS = True
        For i = 1 To UBound(gConnectedShardDir)
            If gConnectedShardDir(i) = strDeviceNO Then
                blnConnectDS = False
                Exit For
            End If
        Next i
        If blnConnectDS = True And strSDPath <> "" Then
            If funcConnectShardDir("\\" & strFTPHost & "\" & strSDPath, strSDUser, strSDPwd) = 0 Then
                ReDim Preserve gConnectedShardDir(UBound(gConnectedShardDir) + 1) As String
                gConnectedShardDir(UBound(gConnectedShardDir)) = strDeviceNO
            End If
        End If
        
        If objPacsCore Is Nothing Then
            Exit Function
        Else
            objPacsCore.CallOpenViewer strImageString, lngAdviceID, objParent, gcnOracle, blnMoved, blnAddImage, intImageInterval
        End If
    Else    '没有查找到图象记录，则关闭原来已经打开的观片窗口
        If Not objPacsCore Is Nothing Then objPacsCore.Closefrom
    End If
    
    OpenViewer = True
    Exit Function
DBError:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function CheckChargeState(ByVal lng医嘱ID As Long, ByVal lng来源 As Long) As Integer
'判断当前的医嘱是否收费
'一条医嘱会有多部位的子医嘱

    Dim rsTemp As New ADODB.Recordset
    Dim strTable As String
    
    CheckChargeState = 0
    
    '住院病人查住院费用记录，门诊、外诊等病人查门诊费用记录
    If lng来源 = 2 Then
        strTable = "住院费用记录"
    Else
        strTable = "门诊费用记录"
    End If
    
    gstrSQL = "Select A.医嘱id, B.记录状态" & vbNewLine & _
                "From 病人医嘱发送 A, " & strTable & " B" & vbNewLine & _
                "Where A.医嘱id = [1] And A.NO = B.NO And A.记录性质 = B.记录性质"
                
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "是否收费", lng医嘱ID)
    
    '缴费的几种情况：1 没记录 ；2 有记录全部为记录状态=1 ；3 有记录，且部份记录状态<>1，表示有退费或有部份未缴
    '只有2这种情况算已缴费（全缴）
    If rsTemp.BOF Then Exit Function
    Do Until rsTemp.EOF
        If Nvl(rsTemp!记录状态, 0) <> 1 Then Exit Function
        rsTemp.MoveNext
    Loop
    CheckChargeState = 1
End Function
Public Function CheckConcurrentReport(ByVal lngOrderID As Long, Optional blnSilence As Boolean = False) As Boolean
'功能：检查当前病人是否有医生正在操作报告
Dim rsTemp As ADODB.Recordset
    
    On Error GoTo errHandle
    CheckConcurrentReport = True
    gstrSQL = "Select 报告操作 From 影像检查记录 Where 医嘱ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "读取记录", lngOrderID)
    
    If Not rsTemp Is Nothing Then
        If Not rsTemp.EOF Then
            If Nvl(rsTemp!报告操作) <> "" And Nvl(rsTemp!报告操作) <> UserInfo.姓名 Then
                If blnSilence = False Then
                    MsgBox "当前病人的报告正在被 " & Nvl(rsTemp!报告操作) & " 操作，请稍后再试。", vbInformation, gstrSysName
                End If
                CheckConcurrentReport = False
            End If
        End If
    End If
    Exit Function
    
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
Public Sub UpdateReporter(ByVal lngOrderID As Long, ByVal Reporter As String)
    On Error GoTo errHandle
    
    gstrSQL = "ZL_影像报告操作_Update(" & lngOrderID & ",'" & Reporter & "')"
    zlDatabase.ExecuteProcedure gstrSQL, "更新操作者"
    Exit Sub
    
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
Public Function bln存在未审划价单(ByVal lng医嘱ID As Long, ByVal lng来源 As Long) As Boolean
    Dim rsTemp As ADODB.Recordset
    Dim strFeeTable As String
    
    '住院病人查住院费用记录，门诊、外诊等病人查门诊费用记录
    If lng来源 = 2 Then
        strFeeTable = "住院费用记录"
    Else
        strFeeTable = "门诊费用记录"
    End If

    On Error GoTo errHandle
    gstrSQL = "Select /*+ RULE */ A.ID" & vbNewLine & _
            "From " & strFeeTable & " A" & vbNewLine & _
            "Where A.医嘱序号 + 0 In (Select ID From 病人医嘱记录 Where ID = [1] Or 相关id = [1]) And (A.记录性质, A.NO) In" & vbNewLine & _
            "      (Select 记录性质, NO" & vbNewLine & _
            "       From 病人医嘱附费" & vbNewLine & _
            "       Where 医嘱id In (Select ID From 病人医嘱记录 Where ID = [1] Or 相关id = [1])" & vbNewLine & _
            "       Union All" & vbNewLine & _
            "       Select 记录性质, NO" & vbNewLine & _
            "       From 病人医嘱发送" & vbNewLine & _
            "       Where 医嘱id In (Select ID From 病人医嘱记录 Where ID = [1] Or 相关id = [1])" & vbNewLine & _
            "       ) And A.记帐费用 = 1 And A.记录状态 = 0"

    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取未审划价单", lng医嘱ID)
    If rsTemp.EOF Then
        Exit Function
    Else
        bln存在未审划价单 = True
    End If
    Exit Function
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
Public Function bln病人在院(ByVal lng病人ID As Long, ByVal lng主页ID As Long) As Boolean
Dim rsTemp As ADODB.Recordset
    On Error GoTo errHandle
    gstrSQL = "SELECT to_char(出院日期,'YYYY-MM-DD HH24:MI:SS') as 出院日期 from 病案主页 where 病人ID=[1] AND 主页ID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取出院时间", lng病人ID, lng主页ID)
    If rsTemp.EOF Then
        Exit Function
    Else
        If Nvl(rsTemp!出院日期) = "" Then
            bln病人在院 = True
        End If
    End If
    Exit Function
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function GetFullPY(strIn As String) As String
'------------------------------------------------
'功能：把传入的字符串中包含的中文转换成拼音，英文字母和数字不做处理
'参数： strIn －－传入的字符串
'返回：把汉字转换成拼音后的字符串
'------------------------------------------------
    Dim i As Integer
    Dim strChar As String
    
    strIn = Trim(strIn)
    For i = 1 To Len(strIn)
        strChar = Mid(strIn, i, 1)
        If Asc(strChar) < 0 Then
            GetFullPY = GetFullPY & UCase(Replace(zlCommFun.mGetFullPY(strChar), vbCrLf, "")) & " "
        Else
            GetFullPY = GetFullPY & strChar
        End If
    Next i
    GetFullPY = Trim(GetFullPY)
End Function

Public Function GetRptImages(ByRef RptViewer As DicomViewer, ByVal lngOrderID As Long, ByVal blnMoved As Boolean)
'------------------------------------------------
'功能：获取报告图像到本地，并刷新显示
'参数： RptViewer －－显示图像的控件
'       lngOrderID -- 医嘱ID
'       blnMoved -- 是否转储
'返回：无，直接往RptViewer 中加入图像
'------------------------------------------------
    Dim objFileSystem As New Scripting.FileSystemObject
    Dim strSQL As String, rsTemp As New ADODB.Recordset
    Dim aryFiles() As String    '报告图像数组
    Dim strFiles As String      '按分号分隔的成功下载的文件
    Dim aryRptFileName() As String    '报告文件名数组
    
    Dim cFtpNet As New cFTP
    Dim strVirtualPath As String
    Dim strLocalPath As String
    Dim IntCount As Integer
    Dim curImage As DicomImage
    
    '先清空RptViewer 中的图像
    RptViewer.Images.Clear
    
    '检查本地缓存图像的根目录是否存在，如果不存在则创建本地根目录，如果创建失败，则直接退出程序
    If objFileSystem.FolderExists(App.Path & "\TmpImage\") = False Then objFileSystem.CreateFolder App.Path & "\TmpImage\"
    If objFileSystem.FolderExists(App.Path & "\TmpImage\") = False Then GetRptImages = False: Exit Function
    
    '从数据库读取图像来源信息
    err = 0: On Error Resume Next
    strSQL = "Select To_Char(L.接收日期, 'yyyymmdd') As 子目录, L.检查uid, L.报告图象, A1.Ftp目录 As Root1, A1.Ip地址 As Ip1," & vbNewLine & _
            "       A1.FTP用户名 As Usr1, A1.FTP密码 As Pwd1, A2.Ftp目录 As Root2, A2.Ip地址 As Ip2, A2.FTP用户名 As Usr2, A2.FTP密码 As Pwd2" & vbNewLine & _
            "From 影像检查记录 L, 影像设备目录 A1, 影像设备目录 A2" & vbNewLine & _
            "Where L.位置一 = A1.设备号(+) And L.位置二 = A2.设备号(+) And L.医嘱id = [1]"
    If blnMoved = True Then
        strSQL = Replace(strSQL, "影像检查记录", "H影像检查记录")
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "提取报告图像", lngOrderID)
    If rsTemp.RecordCount <= 0 Then GetRptImages = False: Exit Function
    aryFiles = Split("" & rsTemp!报告图象, ";")
    aryRptFileName = Split("" & rsTemp!报告图象, ";")
    If UBound(aryFiles) < 0 Then GetRptImages = False: Exit Function
        
    '检查本机缓存中本次检查的目录是否存在，如果不存在则创建本地存储目录，如果创建失败，则退出程序
    err = 0: On Error Resume Next
    strLocalPath = App.Path & "\TmpImage\" & rsTemp!子目录
    If objFileSystem.FolderExists(strLocalPath) = False Then objFileSystem.CreateFolder strLocalPath
    If objFileSystem.FolderExists(strLocalPath) = False Then GetRptImages = False: Exit Function
    strLocalPath = strLocalPath & "\" & rsTemp!检查uid
    If objFileSystem.FolderExists(strLocalPath) = False Then objFileSystem.CreateFolder strLocalPath
    If objFileSystem.FolderExists(strLocalPath) = False Then GetRptImages = False: Exit Function
        
    strFiles = ""
    '检查本地缓存图像是否存在，如果存在，则不从FTP下载，而直接读取本机缓存图像
    For IntCount = 0 To UBound(aryFiles)
        '如果文件存在，则不需要下载，设置标记
        If Dir(strLocalPath & "\" & Trim(aryFiles(IntCount))) <> "" Then
            strFiles = strFiles & ";" & strLocalPath & "\" & Trim(aryFiles(IntCount))
            aryFiles(IntCount) = ""
        End If
    Next IntCount
    If strFiles <> "" Then strFiles = Mid(strFiles, 2)
    
    
    '如果本次存在的文件数量跟需要打开的文件数量不一致，则从FTP下载本机不存在的图像
    If UBound(Split(strFiles, ";")) <> UBound(aryFiles) Then
        '首先从设备1下载图像
        If "" & rsTemp!Ip1 <> "" Then
            If cFtpNet.FuncFtpConnect("" & rsTemp!Ip1, "" & rsTemp!Usr1, "" & rsTemp!pwd1) <> 0 Then
                strVirtualPath = rsTemp!Root1 & "/" & rsTemp!子目录 & "/" & rsTemp!检查uid
                For IntCount = 0 To UBound(aryFiles)
                    If aryFiles(IntCount) <> "" Then
                        If cFtpNet.FuncDownloadFile(strVirtualPath, strLocalPath & "\" & Trim(aryFiles(IntCount)), Trim(aryFiles(IntCount))) = 0 Then
                            If Dir(strLocalPath & "\" & Trim(aryFiles(IntCount))) <> "" Then
                                strFiles = strFiles & ";" & strLocalPath & "\" & Trim(aryFiles(IntCount))
                                aryFiles(IntCount) = ""
                            End If
                        End If
                    End If
                Next IntCount
            End If
            cFtpNet.FuncFtpDisConnect
        End If
        
        '如果设备1下载图像不完整，再从设备2下载图像
        If strFiles <> "" Then strFiles = Mid(strFiles, 2)
        If UBound(Split(strFiles, ";")) <> UBound(aryFiles) And "" & rsTemp!Ip2 <> "" Then
            If cFtpNet.FuncFtpConnect("" & rsTemp!Ip2, "" & rsTemp!Usr2, "" & rsTemp!pwd2) <> 0 Then
                strVirtualPath = rsTemp!Root2 & "/" & rsTemp!子目录 & "/" & rsTemp!检查uid
                For IntCount = 0 To UBound(aryFiles)
                    If aryFiles(IntCount) <> "" Then
                        If cFtpNet.FuncDownloadFile(strVirtualPath, strLocalPath & "\" & Trim(aryFiles(IntCount)), Trim(aryFiles(IntCount))) = 0 Then
                            If Dir(strLocalPath & "\" & Trim(aryFiles(IntCount))) <> "" Then
                                strFiles = strFiles & ";" & strLocalPath & "\" & Trim(aryFiles(IntCount))
                            End If
                        End If
                    End If
                Next IntCount
            End If
            cFtpNet.FuncFtpDisConnect
        End If
        If strFiles <> "" Then
            If Left(strFiles, 1) = ";" Then strFiles = Mid(strFiles, 2)
        End If
    End If
    
    '将获得的文件装入
    Dim iRows As Integer, iCols As Integer
    aryFiles = Split(strFiles, ";")
    With RptViewer
        For IntCount = 0 To UBound(aryFiles)
            Set curImage = New DicomImage
            curImage.FileImport aryFiles(IntCount), "JPG"
            curImage.BorderWidth = 3: curImage.BorderColour = vbWhite
            curImage.Tag = aryRptFileName(IntCount)
            .Images.Add curImage
        Next
        If UBound(aryFiles) >= 0 Then
            .CurrentIndex = 1
            .Images(.CurrentIndex).BorderColour = vbBlue
        End If
        If .Images.Count > 0 Then
            ResizeRegion .Images.Count, .Width, .Height, iRows, iCols
            .MultiColumns = iCols: .MultiRows = iRows
        Else
            '暂无内容
        End If
    End With
    
    GetRptImages = True: Exit Function

errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

'Public Sub PromptResult(lngOrderID As Long, lngModul As Long, frmParent As Form)
'    Dim strResult As String
'
'    strResult = frmResult.zlGetResult(frmParent, lngModul, lngOrderID)    '提示输入阴阳性和影像质量
'    If strResult = "" Then Exit Sub
'
'    If Split(strResult, "-")(0) = 1 Then    '阴阳性
'        gstrSQL = "ZL_影像检查_结果(" & lngOrderID & ",1)"
'    Else
'        gstrSQL = "ZL_影像检查_结果(" & lngOrderID & ",0)"
'    End If
'    zlDatabase.ExecuteProcedure gstrSQL, "标记阴阳性"
'
'    If lngModul = 1290 Then         '影像医技站才记录影像质量
'        If Split(strResult, "-")(1) = 1 Then    '影像质量
'            gstrSQL = "Zl_影像质量_Update(" & lngOrderID & ",'甲')"
'        Else
'            gstrSQL = "Zl_影像质量_Update(" & lngOrderID & ",'乙')"
'        End If
'        zlDatabase.ExecuteProcedure gstrSQL, "影像质量"
'    End If
'End Sub
'Public Function BillingWarn(frmParent As Object, ByVal strPrivs As String, _
'    rsWarn As ADODB.Recordset, ByVal str姓名 As String, ByVal cur剩余款额 As Currency, _
'    ByVal cur当日金额 As Currency, ByVal cur记帐金额 As Currency, ByVal cur担保金额 As Currency, _
'    ByVal str收费类别 As String, ByVal str类别名称 As String, str已报类别 As String, _
'    intWarn As Integer, Optional ByVal bln划价 As Boolean) As Integer
''功能:对病人记帐进行报警提示
''参数:rsWarn=包含报警参数设置的记录集(该病人病区,并区分好了医保)
''     str收费类别=当前要检查的类别,用于分类报警
''     str类别名称=类别名称,用于提示
''     bln划价=生成划价费用时的报警，类似具有强制记帐权限时的处理
''     intWarn=是否显示询问性的提示,-1=要显示,0=缺省为否,1-缺省为是
''返回:str已报类别="CDE":具体在本次报警的一组类别,"-"为所有类别。该返回用于处理重复报警
''     intWarn=本次询问性提示中的选择结果,0=为否,1-为是
''     0;没有报警,继续
''     1:报警提示后用户选择继续
''     2:报警提示后用户选择中断
''     3:报警提示必须中断
''     4:强制记帐报警,继续
'    Dim bln已报警 As Boolean, byt标志 As Byte
'    Dim byt方式 As Byte, byt已报方式 As Byte
'    Dim arrtmp As Variant, vMsg As VbMsgBoxResult
'    Dim str担保 As String, i As Long
'
'    BillingWarn = 0
'
'    '报警参数检查:NULL是没有设置,0是设置了的
'    If rsWarn.State = 0 Then Exit Function
'    If rsWarn.EOF Then Exit Function
'    If IsNull(rsWarn!报警值) Then Exit Function
'
'    '对应类别定位有效报警设置
'    If Not IsNull(rsWarn!报警标志1) Then
'        If rsWarn!报警标志1 = "-" Or InStr(rsWarn!报警标志1, str收费类别) > 0 Then byt标志 = 1
'        If rsWarn!报警标志1 = "-" Then str类别名称 = "" '所有类别时,不必提示具体的类别
'    End If
'    If byt标志 = 0 And Not IsNull(rsWarn!报警标志2) Then
'        If rsWarn!报警标志2 = "-" Or InStr(rsWarn!报警标志2, str收费类别) > 0 Then byt标志 = 2
'        If rsWarn!报警标志2 = "-" Then str类别名称 = "" '所有类别时,不必提示具体的类别
'    End If
'    If byt标志 = 0 And Not IsNull(rsWarn!报警标志3) Then
'        If rsWarn!报警标志3 = "-" Or InStr(rsWarn!报警标志3, str收费类别) > 0 Then byt标志 = 3
'        If rsWarn!报警标志3 = "-" Then str类别名称 = "" '所有类别时,不必提示具体的类别
'    End If
'    If byt标志 = 0 Then Exit Function '无有效设置
'
'    '报警标志2实际上是两种判断①②,其它只有一种判断①
'    '这种处理的前提是一种类别只能属于一种报警方式(报警参数设置时)
'    '示例："-" 或 ",ABC,567,DEF"
'    '报警标志2示例："-①" 或 ",ABC②,567①,DEF①"
'    bln已报警 = InStr(str已报类别, str收费类别) > 0 Or str已报类别 Like "-*"
'
'    If bln已报警 Then '当intWarn = -1时,也可强行再报警
'        If byt标志 = 2 Then
'            If str已报类别 Like "-*" Then
'                byt已报方式 = IIf(Right(str已报类别, 1) = "②", 2, 1)
'            Else
'                arrtmp = Split(str已报类别, ",")
'                For i = 0 To UBound(arrtmp)
'                    If InStr(arrtmp(i), str收费类别) > 0 Then
'                        byt已报方式 = IIf(Right(arrtmp(i), 1) = "②", 2, 1)
'                        'Exit For '取消说明见住院记帐模块
'                    End If
'                Next
'            End If
'        Else
'            Exit Function
'        End If
'    End If
'
'    If str类别名称 <> "" Then str类别名称 = """" & str类别名称 & """费用"
'    str担保 = IIf(cur担保金额 = 0, "", "(含担保额:" & Format(cur担保金额, "0.00") & ")")
'    cur剩余款额 = cur剩余款额 + cur担保金额 - cur记帐金额
'    cur当日金额 = cur当日金额 + cur记帐金额
'
'    '---------------------------------------------------------------------
'    If rsWarn!报警方法 = 1 Then  '累计费用报警(低于)
'        Select Case byt标志
'            Case 1 '低于报警值(包括预交款耗尽)提示询问记帐
'                If cur剩余款额 < rsWarn!报警值 Then
'                    If Not (InStr(";" & strPrivs & ";", ";强制记帐;") > 0 Or bln划价) Then
'                        If intWarn = -1 Then
'                            vMsg = frmMsgBox.ShowMsgBox(str姓名 & " 当前剩余款" & str担保 & ":" & Format(cur剩余款额, "0.00") & ",低于" & str类别名称 & "报警值:" & Format(rsWarn!报警值, "0.00") & ",允许该病人记帐吗？", frmParent)
'                            If vMsg = vbNo Or vMsg = vbCancel Then
'                                If vMsg = vbCancel Then intWarn = 0
'                                BillingWarn = 2
'                            ElseIf vMsg = vbYes Or vMsg = vbIgnore Then
'                                If vMsg = vbIgnore Then intWarn = 1
'                                BillingWarn = 1
'                            End If
'                        Else
'                            If intWarn = 0 Then
'                                BillingWarn = 2
'                            ElseIf intWarn = 1 Then
'                                BillingWarn = 1
'                            End If
'                        End If
'                    Else
'                        If intWarn = -1 Then
'                            vMsg = frmMsgBox.ShowMsgBox(IIf(bln划价, "", "强制记帐") & "提醒:" & vbCrLf & vbCrLf & str姓名 & " 当前剩余款" & str担保 & ":" & Format(cur剩余款额, "0.00") & " 低于" & str类别名称 & "报警值:" & Format(rsWarn!报警值, "0.00") & "。", frmParent, True)
'                            If vMsg = vbIgnore Then intWarn = 1
'                        End If
'                        BillingWarn = 4
'                    End If
'                End If
'            Case 2 '低于报警值提示询问记帐,预交款耗尽时禁止记帐
'                If Not bln已报警 Then
'                    If cur剩余款额 < 0 Then
'                        byt方式 = 2
'                        If Not (InStr(";" & strPrivs & ";", ";强制记帐;") > 0 Or bln划价) Then
'                            If intWarn = -1 Then
'                                vMsg = frmMsgBox.ShowMsgBox(str姓名 & " 当前剩余款" & str担保 & "已经耗尽," & str类别名称 & "禁止记帐。", frmParent, True)
'                                If vMsg = vbIgnore Then intWarn = 1
'                            End If
'                            BillingWarn = 3
'                        Else
'                            If intWarn = -1 Then
'                                vMsg = frmMsgBox.ShowMsgBox(str类别名称 & IIf(bln划价, "", "强制记帐") & "提醒:" & vbCrLf & vbCrLf & str姓名 & " 当前剩余款" & str担保 & "已经耗尽。", frmParent, True)
'                                If vMsg = vbIgnore Then intWarn = 1
'                            End If
'                            BillingWarn = 4
'                        End If
'                    ElseIf cur剩余款额 < rsWarn!报警值 Then
'                        byt方式 = 1
'                        If Not (InStr(";" & strPrivs & ";", ";强制记帐;") > 0 Or bln划价) Then
'                            If intWarn = -1 Then
'                                vMsg = frmMsgBox.ShowMsgBox(str姓名 & " 当前剩余款" & str担保 & ":" & Format(cur剩余款额, "0.00") & ",低于" & str类别名称 & "报警值:" & Format(rsWarn!报警值, "0.00") & ",允许该病人记帐吗？", frmParent)
'                                If vMsg = vbNo Or vMsg = vbCancel Then
'                                    If vMsg = vbCancel Then intWarn = 0
'                                    BillingWarn = 2
'                                ElseIf vMsg = vbYes Or vMsg = vbIgnore Then
'                                    If vMsg = vbIgnore Then intWarn = 1
'                                    BillingWarn = 1
'                                End If
'                            Else
'                                If intWarn = 0 Then
'                                    BillingWarn = 2
'                                ElseIf intWarn = 1 Then
'                                    BillingWarn = 1
'                                End If
'                            End If
'                        Else
'                            If intWarn = -1 Then
'                                vMsg = frmMsgBox.ShowMsgBox(IIf(bln划价, "", "强制记帐") & "提醒:" & vbCrLf & vbCrLf & str姓名 & " 当前剩余款" & str担保 & ":" & Format(cur剩余款额, "0.00") & ",低于" & str类别名称 & "报警值:" & Format(rsWarn!报警值, "0.00") & "。", frmParent, True)
'                                If vMsg = vbIgnore Then intWarn = 1
'                            End If
'                            BillingWarn = 4
'                        End If
'                    End If
'                Else
'                    '上次已报警并选择继续或强制继续
'                    If byt已报方式 = 1 Then
'                        '上次低于报警值并选择继续或强制继续,不再处理低于的情况,但还需要判断预交款是否耗尽
'                        If cur剩余款额 < 0 Then
'                            byt方式 = 2
'                            If Not (InStr(";" & strPrivs & ";", ";强制记帐;") > 0 Or bln划价) Then
'                                If intWarn = -1 Then
'                                    vMsg = frmMsgBox.ShowMsgBox(str姓名 & " 当前剩余款" & str担保 & "已经耗尽," & str类别名称 & "禁止记帐。", frmParent, True)
'                                    If vMsg = vbIgnore Then intWarn = 1
'                                End If
'                                BillingWarn = 3
'                            Else
'                                If intWarn = -1 Then
'                                    vMsg = frmMsgBox.ShowMsgBox(str类别名称 & IIf(bln划价, "", "强制记帐") & "提醒:" & vbCrLf & vbCrLf & str姓名 & " 当前剩余款" & str担保 & "已经耗尽。", frmParent, True)
'                                    If vMsg = vbIgnore Then intWarn = 1
'                                End If
'                                BillingWarn = 4
'                            End If
'                        End If
'                    ElseIf byt已报方式 = 2 Then
'                        '上次预交款已经耗尽并强制继续,不再处理
'                        Exit Function
'                    End If
'                End If
'            Case 3 '低于报警值禁止记帐
'                If cur剩余款额 < rsWarn!报警值 Then
'                    If Not (InStr(";" & strPrivs & ";", ";强制记帐;") > 0 Or bln划价) Then
'                        If intWarn = -1 Then
'                            vMsg = frmMsgBox.ShowMsgBox(str姓名 & " 当前剩余款" & str担保 & ":" & Format(cur剩余款额, "0.00") & ",低于" & str类别名称 & "报警值:" & Format(rsWarn!报警值, "0.00") & ",禁止记帐。", frmParent, True)
'                            If vMsg = vbIgnore Then intWarn = 1
'                        End If
'                        BillingWarn = 3
'                    Else
'                        If intWarn = -1 Then
'                            vMsg = frmMsgBox.ShowMsgBox(IIf(bln划价, "", "强制记帐") & "提醒:" & vbCrLf & vbCrLf & str姓名 & " 当前剩余款" & str担保 & ":" & Format(cur剩余款额, "0.00") & ",低于" & str类别名称 & "报警值:" & Format(rsWarn!报警值, "0.00") & "。", frmParent, True)
'                            If vMsg = vbIgnore Then intWarn = 1
'                        End If
'                        BillingWarn = 4
'                    End If
'                End If
'        End Select
'    ElseIf rsWarn!报警方法 = 2 Then  '每日费用报警(高于)
'        Select Case byt标志
'            Case 1 '高于报警值提示询问记帐
'                If cur当日金额 > rsWarn!报警值 Then
'                    If Not (InStr(";" & strPrivs & ";", ";强制记帐;") > 0 Or bln划价) Then
'                        If intWarn = -1 Then
'                            vMsg = frmMsgBox.ShowMsgBox(str姓名 & " 当日费用:" & Format(cur当日金额, gstrDec) & ",高于" & str类别名称 & "报警值:" & Format(rsWarn!报警值, "0.00") & ",允许该病人记帐吗？", frmParent)
'                            If vMsg = vbNo Or vMsg = vbCancel Then
'                                If vMsg = vbCancel Then intWarn = 0
'                                BillingWarn = 2
'                            ElseIf vMsg = vbYes Or vMsg = vbIgnore Then
'                                If vMsg = vbIgnore Then intWarn = 1
'                                BillingWarn = 1
'                            End If
'                        Else
'                            If intWarn = 0 Then
'                                BillingWarn = 2
'                            ElseIf intWarn = 1 Then
'                                BillingWarn = 1
'                            End If
'                        End If
'                    Else
'                        If intWarn = -1 Then
'                            vMsg = frmMsgBox.ShowMsgBox(IIf(bln划价, "", "强制记帐") & "提醒:" & vbCrLf & vbCrLf & str姓名 & " 当日费用:" & Format(cur当日金额, gstrDec) & ",高于" & str类别名称 & "报警值:" & Format(rsWarn!报警值, "0.00") & "。", frmParent, True)
'                            If vMsg = vbIgnore Then intWarn = 1
'                        End If
'                        BillingWarn = 4
'                    End If
'                End If
'            Case 3 '高于报警值禁止记帐
'                If cur当日金额 > rsWarn!报警值 Then
'                    If Not (InStr(";" & strPrivs & ";", ";强制记帐;") > 0 Or bln划价) Then
'                        If intWarn = -1 Then
'                            vMsg = frmMsgBox.ShowMsgBox(str姓名 & " 当日费用:" & Format(cur当日金额, gstrDec) & ",高于" & str类别名称 & "报警值:" & Format(rsWarn!报警值, "0.00") & ",禁止记帐。", frmParent, True)
'                            If vMsg = vbIgnore Then intWarn = 1
'                        End If
'                        BillingWarn = 3
'                    Else
'                        If intWarn = -1 Then
'                            vMsg = frmMsgBox.ShowMsgBox(IIf(bln划价, "", "强制记帐") & "提醒:" & vbCrLf & vbCrLf & str姓名 & " 当日费用:" & Format(cur当日金额, gstrDec) & ",高于" & str类别名称 & "报警值:" & Format(rsWarn!报警值, "0.00") & "。", frmParent, True)
'                            If vMsg = vbIgnore Then intWarn = 1
'                        End If
'                        BillingWarn = 4
'                    End If
'                End If
'        End Select
'    End If
'
'    '对于继续类的操作,返回已报警类别
'    If BillingWarn = 1 Or BillingWarn = 4 Then
'        If byt标志 = 1 Then
'            If rsWarn!报警标志1 = "-" Then
'                str已报类别 = "-"
'            Else
'                str已报类别 = str已报类别 & "," & rsWarn!报警标志1
'            End If
'        ElseIf byt标志 = 2 Then
'            If rsWarn!报警标志2 = "-" Then
'                str已报类别 = "-"
'            Else
'                str已报类别 = str已报类别 & "," & rsWarn!报警标志2
'            End If
'            '附加标注以判断已报警的具体方式
'            str已报类别 = str已报类别 & IIf(byt方式 = 2, "②", "①")
'        ElseIf byt标志 = 3 Then
'            If rsWarn!报警标志3 = "-" Then
'                str已报类别 = "-"
'            Else
'                str已报类别 = str已报类别 & "," & rsWarn!报警标志3
'            End If
'        End If
'    End If
'End Function

'Public Function FinishBillingWarn(ByVal frmParent As Object, ByVal strPrivs As String, ByVal lng病人ID As Long, _
'    ByVal lng主页ID As Long, ByVal lng病区ID As Long, ByVal cur金额 As Currency, ByVal str类别 As String, ByVal str类别名 As String) As Boolean
''功能：当执行完成有自动审核的费用时，对病人费用进行记帐报警。
''参数：str类别="CDE..."，报警金额涉及到的收费类别
''      str类别名="检查,检验,..."，对应的类别名用于提示
'    Dim rsPati As ADODB.Recordset
'    Dim rsWarn As ADODB.Recordset
'    Dim strWarn As String, intWarn As Integer
'    Dim strSQL As String, intR As Integer, i As Long
'    Dim cur当日 As Currency
'
'    On Error GoTo errH
'
'    If lng主页ID <> 0 Then
'        '住院病人报警
'        strSQL = _
'            " Select 病人ID,预交余额,费用余额,0 as 预结费用 From 病人余额 Where 性质=1 And 病人ID=[1]" & _
'            " Union ALL" & _
'            " Select A.病人ID,0,0,Sum(金额) From 保险模拟结算 A,病案主页 B" & _
'            " Where A.病人ID=B.病人ID And A.主页ID=B.主页ID And B.险类 Is Not Null And A.病人ID=[1] And A.主页ID=[2] Group by A.病人ID"
'        strSQL = "Select 病人ID,Nvl(Sum(预交余额),0)-Nvl(Sum(费用余额),0)+Nvl(Sum(预结费用),0) as 剩余款 From (" & strSQL & ") Group by 病人ID"
'
'        strSQL = "Select A.姓名,zl_PatiWarnScheme(A.病人ID,B.主页ID) as 适用病人,C.剩余款," & _
'            " Decode(A.担保额,Null,Null,zl_PatientSurety(A.病人ID,B.主页ID)) as 担保额" & _
'            " From 病人信息 A,病案主页 B,(" & strSQL & ") C" & _
'            " Where A.病人ID=B.病人ID And A.病人ID=C.病人ID(+)" & _
'            " And A.病人ID=[1] And B.主页ID=[2]"
'        Set rsPati = zlDatabase.OpenSQLRecord(strSQL, "FinishBillingWarn", lng病人ID, lng主页ID)
'    Else
'        '其他按门诊报警
'        strSQL = "Select 病人ID,预交余额,费用余额 From 病人余额 Where 性质=1 And 病人ID=[1]"
'        strSQL = "Select A.姓名,zl_PatiWarnScheme(A.病人ID) as 适用病人,A.担保额," & _
'            " Nvl(B.预交余额,0)-Nvl(B.费用余额,0)+Nvl(E.帐户余额,0) as 剩余款" & _
'            " From 病人信息 A,(" & strSQL & ") B,医保病人关联表 D,医保病人档案 E" & _
'            " Where A.病人ID=B.病人ID(+) And A.病人id = D.病人id(+) And A.险类=D.险类(+)" & _
'            " And D.险类=E.险类(+) And D.中心=E.中心(+) And D.医保号=E.医保号(+) And D.标志(+)=1" & _
'            " And A.病人ID=[1]"
'        Set rsPati = zlDatabase.OpenSQLRecord(strSQL, "FinishBillingWarn", lng病人ID)
'    End If
'
'    intWarn = -1 '记帐报警时缺省要提示
'    '执行报警:门诊病人病区ID=0
'    strSQL = "Select Nvl(报警方法,1) as 报警方法,报警值,报警标志1,报警标志2,报警标志3 From 记帐报警线 Where Nvl(病区ID,0)=[1] And 适用病人=[2]"
'    Set rsWarn = zlDatabase.OpenSQLRecord(strSQL, "FinishBillingWarn", lng病区ID, CStr(Nvl(rsPati!适用病人)))
'    If Not rsWarn.EOF Then
'        If rsWarn!报警方法 = 2 Then cur当日 = GetPatiDayMoney(lng病人ID)
'        str类别名 = Mid(str类别名, 2)
'        For i = 1 To Len(str类别)
'            intR = BillingWarn(frmParent, strPrivs, rsWarn, Nvl(rsPati!姓名), Nvl(rsPati!剩余款, 0), cur当日, cur金额, Nvl(rsPati!担保额, 0), Mid(str类别, i, 1), Split(str类别名, ",")(i - 1), strWarn, intWarn)
'            If InStr(",2,3,", intR) > 0 Then Exit Function
'        Next
'    End If
'
'    FinishBillingWarn = True
'    Exit Function
'errH:
'    If ErrCenter() = 1 Then Resume
'    Call SaveErrLog
'End Function

Public Function GetAdviceMoney(ByVal lngAdviceID As Long, ByVal lng来源 As Long, str类别 As String, str类别名 As String) As Currency
'功能：根据指定的医嘱ID串，获取医嘱对应未审核的记帐费用合计
'参数：lngAdviceID,strSendNo
'返回：str类别,str类别名=用于报警提示
'说明：当系统参数为执行后审核费用时才返回。
    Dim rsTmp As New ADODB.Recordset
    Dim curMoney As Currency
    Dim strFeeTable As String
    
    str类别 = "": str类别名 = ""
    
    On Error GoTo errH
    
    '需要根据系统参数判断，81号参数是"执行后自动审核划价单"
    If zlDatabase.GetPara(81, glngSys) <> "1" Then Exit Function
    
    '住院病人查住院费用记录，门诊、外诊等病人查门诊费用记录
    If lng来源 = 2 Then
        strFeeTable = "住院费用记录"
    Else
        strFeeTable = "门诊费用记录"
    End If
    
    gstrSQL = "Select /*+ RULE */" & vbNewLine & _
                " B.编码, B.名称, Sum(A.实收金额) As 金额" & vbNewLine & _
                "From " & strFeeTable & " A, 收费项目类别 B" & vbNewLine & _
                "Where A.医嘱序号 + 0 In (Select ID From 病人医嘱记录 Where ID = [1] Or 相关id = [1]) And" & vbNewLine & _
                "      (A.记录性质, A.NO) In" & vbNewLine & _
                "      ( Select 记录性质, NO" & vbNewLine & _
                "        From 病人医嘱附费" & vbNewLine & _
                "        Where 医嘱id In (Select ID From 病人医嘱记录 Where ID = [1] Or 相关id = [1])" & vbNewLine & _
                "        Union All" & vbNewLine & _
                "        Select 记录性质, NO" & vbNewLine & _
                "        From 病人医嘱发送" & vbNewLine & _
                "        Where 医嘱id In (Select ID From 病人医嘱记录 Where ID = [1] Or 相关id = [1] )" & vbNewLine & _
                "       ) And A.记帐费用 = 1 And A.记录状态 = 0 And A.收费类别 = B.编码 " & vbNewLine & _
                "Group By B.编码, B.名称"
                
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "GetAdviceMoney", lngAdviceID)
    
    curMoney = 0
    Do While Not rsTmp.EOF
        curMoney = curMoney + Nvl(rsTmp!金额, 0)
        str类别 = str类别 & rsTmp!编码
        str类别名 = str类别名 & "," & rsTmp!名称
        rsTmp.MoveNext
    Loop
    
    str类别名 = Mid(str类别名, 2)
    GetAdviceMoney = curMoney
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
Public Function GetPatiDayMoney(lng病人ID As Long) As Currency
'功能：获取指定病人当天发生的费用总额
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select zl_PatiDayCharge([1]) as 金额 From Dual"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", lng病人ID)
    If Not rsTmp.EOF Then GetPatiDayMoney = Nvl(rsTmp!金额, 0)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function funcConnectShardDir(strShareRemoteDir As String, strUserName As String, strPassWord As String) As Long
    '创建网络资源
    Dim NetR As NETRESOURCE
    Dim lngResult As Long
    
    NetR.dwType = RESOURCETYPE_ANY
    NetR.lpLocalName = vbNullString
    NetR.lpRemoteName = strShareRemoteDir
    NetR.lpProvider = vbNullString
    lngResult = WNetAddConnection2(NetR, strPassWord, strUserName, 0)
    
    If lngResult <> 0 Then
        MsgBox "网络连接失败，请检查网络设置是否正确！"
    End If
    funcConnectShardDir = lngResult
End Function

Public Function bln费用未审核(ByVal lng病人ID As Long, ByVal lng主页ID As Long, ByVal lng医嘱ID As Long, ByVal lng来源 As Long) As Boolean
'判断病人是否已出院或为门诊病人，且有记账费用未审核
'需要根据系统参数判断，81号参数是"执行后自动审核划价单"
    
    bln费用未审核 = False
    
    If zlDatabase.GetPara(81, glngSys) = 1 Then
        If Not bln病人在院(lng病人ID, lng主页ID) And bln存在未审划价单(lng医嘱ID, lng来源) Then
            bln费用未审核 = True
        End If
    End If
End Function

Public Function AssembleImage(AssembleViewer As DicomImages, ByVal intRows As Integer, ByVal intCols As Integer, _
    ByVal lngHeight As Long, ByVal lngWidth As Long) As DicomImage

'组合viewer中的显示的所有图像成一个图像

    Dim Image As New DicomImage '新图像
    Dim imgs As New DicomImages '临时存储屏幕采集的图像集
    Dim intWidth As Integer     '新图像的宽度
    Dim intHeight As Integer    '新图像的高度
    Dim Simg As New DicomImage
    Dim sZoom As Single
    Dim intImgRectWidth As Integer  '单张图像可占用的区域宽度
    Dim intImgRectHeight As Integer '单张图像可占用的区域高度
    Dim i As Integer
    Dim intMaxWidth As Integer      '拼接后图像的最大宽度
    Dim intMaxHeight As Integer     '拼接后图像的最大高度
    Dim intBorder As Integer        '图像之间的边距
    Dim intOffsetX As Integer       '拼接时X方向的位移
    Dim intOffsetY As Integer       '拼接时Y方向的位移
    Dim lngWhiteX As Long           '将图象底色改成白色的X宽度
    Dim lngWhiteY As Long           '将图象底色改成白色的Y高度
    
    If AssembleViewer.Count <= 0 Then
        '返回一个黑图**************
        Exit Function
    End If

    On Error GoTo err
    '计算新图像的宽度和高度

    '新图像的宽度和高度不能够大于intMaxWidth×intMaxHeight（宽度×高度）
    intMaxWidth = 3073
    intMaxHeight = 3073
    intBorder = 10

    intImgRectWidth = 0
    intImgRectHeight = 0

    '估算新图像的宽度和高度

    '使用原图像的宽度和高度和，并用Viewer的比例来修正。

    '估算图像的新宽高
    For i = 1 To AssembleViewer.Count
        If intImgRectWidth < AssembleViewer(i).SizeX Then intImgRectWidth = AssembleViewer(i).SizeX
        If intImgRectHeight < AssembleViewer(i).SizeY Then intImgRectHeight = AssembleViewer(i).SizeY
    Next i
    
    '计算横向和纵向图像数量
    intWidth = intImgRectWidth * intCols
    intHeight = intImgRectHeight * intRows
    
    '修正图像的宽高，不能大于最大值
    '如果大于intMaxWidth×intMaxHeight则，按照图像总长宽比，使用小于等于intMaxWidth×intMaxHeight作为新宽高,
    If intWidth > intMaxWidth Or intHeight > intMaxHeight Then
        If intHeight / intWidth > intMaxHeight / intMaxWidth Then
            intWidth = intWidth / intHeight * intMaxHeight
            intHeight = intMaxHeight
        Else
            intHeight = intHeight / intWidth * intMaxWidth
            intWidth = intMaxWidth
        End If
    End If
    
    '采集图像
    '将图像采集到临时图像集
    For i = 1 To AssembleViewer.Count
        '计算缩放比例 hj修改,解决多图合并时，放大的图象无法真正放大的问题
        sZoom = intImgRectHeight / AssembleViewer(i).SizeY
        If sZoom > intImgRectWidth / AssembleViewer(i).SizeX Then
            sZoom = intImgRectWidth / AssembleViewer(i).SizeX
        End If
        
        AssembleViewer(i).StretchToFit = False
        AssembleViewer(i).Zoom = sZoom
        '采集图像
        Set Simg = AssembleViewer(i).PrinterImage(8, 3, True, sZoom, 0, AssembleViewer(i).SizeX, 0, AssembleViewer(i).SizeY)
        imgs.Add Simg
    Next i

    '精确计算新图像的宽度和高度
    intImgRectWidth = 0
    intImgRectHeight = 0

    For i = 1 To imgs.Count
        If intImgRectWidth < imgs(i).SizeX Then intImgRectWidth = imgs(i).SizeX
        If intImgRectHeight < imgs(i).SizeY Then intImgRectHeight = imgs(i).SizeY
        imgs(i).Attributes.Add &H8, &H16, "doSOP_SecondaryCapture"
    Next i
    intImgRectWidth = intImgRectWidth + intBorder
    intImgRectHeight = intImgRectHeight + intBorder
    intWidth = intImgRectWidth * intCols
    intHeight = intImgRectHeight * intRows

    '创建新图像
    Image.Name = "print"
    Image.PatientID = "print001"
    
    Image.Attributes.Add &H8, &H16, doSOP_SecondaryCapture
    Image.Attributes.Add &H28, &H2, 3 ' samples/pixel
    Image.Attributes.Add &H28, &H4, "RGB" ' photometric interpreation  'CT都是MONOCHROME2,CR都是MONOCHROME1？
    Image.Attributes.Add &H28, &H10, intHeight  'x,Rows
    Image.Attributes.Add &H28, &H11, intWidth 'Y,Columns
    Image.Attributes.Add &H28, &H100, 8 'bits allocated
    Image.Attributes.Add &H28, &H101, 8 ' bits stored
    Image.Attributes.Add &H28, &H102, 7 ' high bit
    ReDim pix(intWidth * 3, intHeight * 3) As Byte
    For lngWhiteX = 0 To intWidth * 3
        For lngWhiteY = 0 To intHeight * 3
            pix(lngWhiteX, lngWhiteY) = 255
        Next lngWhiteY
    Next lngWhiteX
    Image.Attributes.Add &H7FE0, &H10, pix

    '拼接新图像
    For i = 1 To imgs.Count
        '计算图像内位移
        intOffsetX = (intImgRectWidth - imgs(i).SizeX - intBorder) / 2
        intOffsetY = (intImgRectHeight - imgs(i).SizeY - intBorder) / 2
        Image.Blt imgs(i), 0, 0, ((i - 1) Mod intCols) * intImgRectWidth + intOffsetX, ((i - 1) \ intCols) * intImgRectHeight + intOffsetY, imgs(i).SizeX, imgs(i).SizeY, 1, 1, 1, False
    Next i

    Set AssembleImage = Image
    Exit Function
err:
End Function

Public Function FunLogIn(frmParent As Form, str类型 As String) As String
'功能：对程序进行注册，如果注册成功，则返回注册时间
'参数： frmParent ---父窗体
'       str类型 ---'在注册码中使用的类型名称
'返回值：注册成功注册日期；注册失败返回空

    Dim intNUM As Integer
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim strIP地址 As String         '需要注册的IP地址
    
    On Error GoTo err
    
    strIP地址 = funGetOneIP
    
    '从注册码中提取授权的数量，-1--无限制；0--禁止；X（X>0）--按照数量控制
    intNUM = gint视频设备数量
    
    'intNUM >0 ,则调用过程注册程序
    If intNUM > 0 Then  '按数量限制
        strSQL = "Zl_影像操作记录_Update('" & strIP地址 & "','" & str类型 & "'," & intNUM & ")"
        zlDatabase.ExecuteProcedure strSQL, "注册" & str类型
        '检查注册是否成功
        strSQL = "Select 启动时间,IP地址 from 影像操作记录 where  类型=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "提取启动时间", str类型)
        
        If rsTemp.RecordCount <= intNUM Then
            rsTemp.Filter = "IP地址='" & strIP地址 & "'"
            If rsTemp.RecordCount = 1 Then  '注册成功
                FunLogIn = rsTemp!启动时间
                Exit Function
            End If
        End If
    ElseIf intNUM = -1 Then     '无限制
        FunLogIn = Now
        Exit Function
    Else    '=0，或者其他值，禁止，不做任何处理，后面有提示
    
    End If
    
    '注册失败，可能是两个原因：
    '1、注册的数量超过了许可的数量，无法注册IP地址
    '2、直接通过SQL向表中添加了IP地址，导致表中的记录总数量超过了许可的数量
    Call MsgBoxD(frmParent, "打开的" & str类型 & "超过您购买的总数量（" & intNUM & "）。请向软件供应商联系。", vbOKOnly, gstrSysName)
    FunLogIn = ""
    
    Exit Function
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function FunLogOut(frmParent As Form, str类型 As String, str启动时间 As String) As Boolean
'功能：退出程序的时候，检查程序是否合法注册过，避免有人通过触发器等手段定时删除“影像操作记录”表中的记录。
'参数： frmParent ---父窗体
'       str类型 ---'在注册码中使用的类型名称
'       str启动时间 --- 注册工作站时返回的时间
'返回值：合法注册True；非法启动的False
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim strIP地址 As String         '需要注册的IP地址
    Dim intNUM As Integer
    
    On Error GoTo err
    strIP地址 = funGetOneIP
    
    '启动时间为空，表示注册失败，没有正常启动，因此退出的时候不再检测数据库
    If str启动时间 = "" Then
        FunLogOut = True
        Exit Function
    End If
    
    '从注册码中提取授权的数量，-1--无限制；0--禁止；X（X>0）--按照数量控制
    intNUM = gint视频设备数量
    
    If intNUM > 0 Then '按照数量控制
        strSQL = "Select 启动时间 from 影像操作记录 where IP地址=[1] and 类型=[2]"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "提取启动时间", strIP地址, str类型)
        If rsTemp.EOF = False Then
            FunLogOut = True
        Else
            '对比启动时间和数据库的时间，如果不是同一天，说明是前一天开启程序后注册信息被删除了，
            '这种情况认为是合法注册
            strSQL = "Select sysdate from dual"
            Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "提取数据库时间")
            If Format(rsTemp!sysdate, "yyyy-mm-dd") <> Format(str启动时间, "yyyy-mm-dd") Then
                FunLogOut = True
            Else
                FunLogOut = False
            End If
        End If
    ElseIf intNUM = -1 Then     '无限制
        FunLogOut = True
    Else    '=0，或者其他值，禁止
    
    End If
    If FunLogOut = False Then
        Call MsgBoxD(frmParent, "打开的" & str类型 & "超过您购买的总数量（" & intNUM & "）。请向软件供应商联系。", vbOKOnly, gstrSysName)
    End If
    Exit Function
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function getLicenseCount(strLicenseName As String) As Integer
'读取授权的数量
'参数： strLicenseName --- 授权名称
    Dim strLiceseCount As String
    
    On Error GoTo err
    
    strLiceseCount = zl9comlib.zlRegInfo(strLicenseName)
    If strLiceseCount = "" Then '无限制
        getLicenseCount = -1
    ElseIf Val(strLiceseCount) > 0 Then '按照数量限制
        getLicenseCount = Val(strLiceseCount)
    Else '禁止
        getLicenseCount = 0
    End If
    
    Exit Function
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function getStudyState(ByVal lngOrderID As Long, Optional ByRef lngSendNO As Long, _
        Optional ByRef str创建人 As String, Optional ByRef str签名 As String, Optional ByRef str保存人 As String, _
        Optional ByRef bln保存结果阳性 As Boolean) As Integer
'检查报告的签名情况，确定本次检查进行的程度。
'参数： lngOrderID [IN] --- 医嘱id
'       lngSendNo [OUT] --- 返回，发送号
'       str创建人 [OUT] --- 返回，报告的创建人
'       str签名   [OUT] --- 返回，报告的最后签名
'       str保存人 [OUT] --- 返回，报告的最后保存人
'       bln保存结果阳性[OUT]--- 返回，结果阳性是否已经输入,True-已输入，False-未输入
'返回值：1--已登记；2--已报到；3--已检查；4--已报告；5--已审核；6--已完成（本过程不存在这个返回值）
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim rsLevel As ADODB.Recordset
    
    On Error GoTo err
    
    strSQL = "Select d.医嘱id As 影像医嘱ID,e.医嘱id As 报告医嘱ID,c.发送号,d.检查uid, " _
             & " e.病历id,e.创建人, e.保存人, e.签名级别, e.完成时间, e.最后版本,c.结果阳性 " _
             & " From 病人医嘱发送 c, 影像检查记录 d, " _
             & " (Select a.医嘱id,a.病历id,b.创建人, b.保存人, b.签名级别, b.完成时间, b.最后版本 " _
             & "  From 病人医嘱报告 a, 电子病历记录 b Where a.医嘱id = [1] And a.病历id = b.Id) e " _
             & " Where c.医嘱id = [1] And d.医嘱id(+) = c.医嘱id And e.医嘱id(+) = c.医嘱id "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "提取是否签名", CLng(lngOrderID))
    
    '如果查询没有结果，就退出
    If rsTemp.EOF = True Then Exit Function
    
    lngSendNO = rsTemp!发送号
    str创建人 = Nvl(rsTemp!创建人)
    str保存人 = Nvl(rsTemp!保存人)
    bln保存结果阳性 = Not IsNull(rsTemp!结果阳性)
    
    '如果影像医嘱ID为空，则过程=1,已登记
    '如果报告医嘱ID为空，且 检查UID为空，则过程=2，已报到
    '如果报告医嘱ID为空，检查UID不为空，则过程=3，已检查
    '其他检查签名和报告完成情况，确定过程为2,3,4，5，已报到,已检查,已报告，已审核
    
    If Nvl(rsTemp!影像医嘱ID) = "" Then     '过程=1,已登记
        getStudyState = 1
    ElseIf Nvl(rsTemp!报告医嘱ID) = "" And Nvl(rsTemp!检查uid) = "" Then    '过程=2，已报到
        getStudyState = 2
    ElseIf Nvl(rsTemp!报告医嘱ID) = "" And Nvl(rsTemp!检查uid) <> "" Then    '过程=3，已检查
        getStudyState = 3
    Else    '检查签名和报告完成情况,确定过程为2,3,4，5，已报到,已检查,已报告，已审核
        If Nvl(rsTemp!完成时间) = "" And rsTemp!最后版本 = 1 Then
            '未签名保存 或最后一次医师退签，执行过程有图像为已检查，无图像为已报到
            getStudyState = IIf(Nvl(rsTemp!检查uid) = "", 2, 3)
        Else
            '判断当前报告的签名情况，如果“电子病历内容”中有大于1的签名，则属于已审核。
            If rsTemp!签名级别 > 1 Then '已审核
                getStudyState = 5
            ElseIf rsTemp!签名级别 = 0 And rsTemp!最后版本 > 1 Then
                '回退出现的状态，可能已审核，需要检查“电子病历内容”中最大的签名级别
                strSQL = "Select 要素表示 As 签名级别,内容文本 as 签名  From 电子病历内容 Where 文件ID=[1] " _
                        & " And 对象类型= 8 And 开始版 = [2] order by 签名级别 desc "
                Set rsLevel = zlDatabase.OpenSQLRecord(strSQL, "提取签名级别", CLng(rsTemp!病历Id), CLng(rsTemp!最后版本 - 1))
                
                If rsLevel.EOF = False Then
                    If rsLevel!签名级别 > 1 Then
                        getStudyState = 5
                        str签名 = Split(Nvl(rsLevel!签名), ";")(0)
                    Else
                        getStudyState = 4
                    End If
                Else
                    getStudyState = 4
                End If
            Else
                getStudyState = 4
            End If
        End If
    End If
    
    Exit Function
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function funGetOneIP() As String
'读取当前计算机的首个IP地址
    Dim strIP地址 As String
    
    On Error Resume Next
    
    strIP地址 = funcGetLocalIP
    If strIP地址 = "" Then
        funGetOneIP = "127.0.0.1"
    ElseIf InStr(strIP地址, ",") <> 0 Then
        funGetOneIP = Split(strIP地址, ",")(0)
    Else
        funGetOneIP = strIP地址
    End If
End Function

Private Function funcGetLocalIP() As String
'返回当前计算机的IP地址串，用逗号分隔
    Dim hostname As String * 256
    Dim hostent_addr As Long
    Dim host As HOSTENT
    Dim hostip_addr As Long
    Dim temp_ip_address() As Byte
    Dim i As Integer
    Dim ip_address As String
    Dim strLocalIPs As String

    '启动Socket
    Call SocketsInitialize

    If gethostname(hostname, 256) = SOCKET_ERROR Then
        MsgBox "Windows Sockets error " & Str(WSAGetLastError())
        Exit Function
    Else
        hostname = Trim$(hostname)
    End If

    hostent_addr = gethostbyname(hostname)

    If hostent_addr = 0 Then
        MsgBox "Winsock.dll is not responding."
        Exit Function
    End If

    RtlMoveMemory host, hostent_addr, LenB(host)
    RtlMoveMemory hostip_addr, host.hAddrList, 4

    ''''''''''''''''get all of the IP address if machine is  multi-homed

    Do
        ReDim temp_ip_address(1 To host.hLength)
        RtlMoveMemory temp_ip_address(1), hostip_addr, host.hLength

        For i = 1 To host.hLength
            ip_address = ip_address & temp_ip_address(i) & "."
        Next
        ip_address = Mid$(ip_address, 1, Len(ip_address) - 1)

        strLocalIPs = IIf(strLocalIPs = "", ip_address, strLocalIPs & "," & ip_address)

        ip_address = ""
        host.hAddrList = host.hAddrList + LenB(host.hAddrList)
        RtlMoveMemory hostip_addr, host.hAddrList, 4
     Loop While (hostip_addr <> 0)

    '清除Socket
    Call SocketsCleanup
    
    funcGetLocalIP = strLocalIPs
End Function

Private Sub SocketsInitialize()
    Dim WSAD As WSADATA
    Dim iReturn As Integer
    Dim sLowByte As String, sHighByte As String, sMsg As String

    iReturn = WSAStartup(WS_VERSION_REQD, WSAD)

    If iReturn <> 0 Then
        MsgBox "Winsock.dll is not responding."
        Exit Sub
    End If

    If lobyte(WSAD.wversion) < WS_VERSION_MAJOR Or (lobyte(WSAD.wversion) = _
        WS_VERSION_MAJOR And hibyte(WSAD.wversion) < WS_VERSION_MINOR) Then

        sHighByte = Trim$(Str$(hibyte(WSAD.wversion)))
        sLowByte = Trim$(Str$(lobyte(WSAD.wversion)))
        sMsg = "Windows Sockets version " & sLowByte & "." & sHighByte
        sMsg = sMsg & " is not supported by winsock.dll "
        MsgBox sMsg
        Exit Sub
    End If

    ''''''''''''''''iMaxSockets is not used in winsock 2. So the following check is only
    ''''''''''''''''necessary for winsock 1. If winsock 2 is requested,
    ''''''''''''''''the following check can be skipped.

    If WSAD.iMaxSockets < MIN_SOCKETS_REQD Then
        sMsg = "This application requires a minimum of "
        sMsg = sMsg & Trim$(Str$(MIN_SOCKETS_REQD)) & " supported sockets."
        MsgBox sMsg
        Exit Sub
    End If
End Sub

Private Sub SocketsCleanup()
Dim lReturn As Long

    lReturn = WSACleanup()

    If lReturn <> 0 Then
        MsgBox "Socket error " & Trim$(Str$(lReturn)) & " occurred in Cleanup "
        Exit Sub
    End If
End Sub

Private Function hibyte(ByVal wParam As Integer)
    hibyte = wParam \ &H100 And &HFF&
End Function

Private Function lobyte(ByVal wParam As Integer)
    lobyte = wParam And &HFF&
End Function
