Attribute VB_Name = "mdlDefine"
Option Explicit

'-----------------------------------------
'全局常量、全局变量、API函数等定义模块
'-----------------------------------------
Public Declare Function RegEnumKeyEx Lib "advapi32.dll" Alias "RegEnumKeyExA" (ByVal hKey As Long, _
    ByVal dwIndex As Long, ByVal lpName As String, lpcbName As Long, ByVal lpReserved As Long, _
    ByVal lpClass As String, lpcbClass As Long, lpftLastWriteTime As FILETIME) As Long
Public Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Public Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Public Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Public Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long
Public Declare Function GetKeyboardLayoutList Lib "user32" (ByVal nBuff As Long, lpList As Long) As Long
Public Declare Function ImmIsIME Lib "imm32.dll" (ByVal hkl As Long) As Long
Public Declare Function ImmGetDescription Lib "imm32.dll" Alias "ImmGetDescriptionA" (ByVal hkl As Long, ByVal lpsz As String, ByVal uBufLen As Long) As Long
Public Declare Function ActivateKeyboardLayout Lib "user32" (ByVal hkl As Long, ByVal flags As Long) As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long

Public Const GSTR_MSG As String = "药品自动化设备管理工具"
Public Const GSTR_CONFIG_FILE As String = "zlDrugMachine.cfg"
Public Const GLNG_SYSTEM As Long = 100
Public Const GLNG_MODULE As Long = 9010

Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_USERS = &H80000003
Public Const HKEY_PERFORMANCE_DATA = &H80000004
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
Public Const ERROR_NONE = 0
Public Const ERROR_BADKEY = 2
Public Const ERROR_ACCESS_DENIED = 8
Public Const ERROR_SUCCESS = 0
Public Const REG_SZ = 1                         ' Unicode空终结字符串
Public Const REG_EXPAND_SZ = 2                  ' Unicode空终结字符串
Public Const REG_DWORD = 4                      ' 32-bit 数字

Public gobjRegister As Object
Public gobjComLib As Object
Public gobjZLPrint As Object
Public gobjEncrypt As Object
Public gcnOracle As ADODB.Connection
Public gobjFile As FileSystemObject
Public gstrUser As String
Public gobjXML As clsXML
Public gstrSQL As String

Public Type TYPE_USER_INFO
    ID As Long
    部门ID As Long
    编号 As String
    姓名 As String
    简码 As String
    用户名 As String
End Type
Public UserInfo As TYPE_USER_INFO

Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public Type TYPE_PARAMS
    定时周期 As Integer
    有效天数 As Integer
    显示最大行数 As Integer
    输出日志 As Boolean
    详细日志 As Boolean
    保存日志天数 As Integer
End Type

Public Enum enuEditState
    查看 = 0
    新增 = 1
    修改 = 2
End Enum

Public Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

'支持的接口类型
Public Const GSTR_TYPE As String = "1-韦乐海茨|2-TOSHO|3-蝶和|4-关拉尼|5-YUYAMA|6-高园|7-苏州厚宏"

Public Enum enuMenus
    文件 = 10
        打印设置 = 1001
        打印预览 = 1002
        打印 = 1003
        输出Excel = 1004
        参数设置 = 1005
        退出 = 1099
    编辑 = 20
        新增 = 2001
        修改 = 2002
        删除 = 2003
    操作 = 30
        启用 = 3011
        停用 = 3012
        设备接口管理 = 3021
        基础数据传送 = 3022
        显示 = 3031
        隐藏 = 3032
    查看 = 80
        工具栏 = 8001
            标准按钮 = 800101
            文本标签 = 800102
            大图标 = 800103
        状态栏 = 8002
        刷新 = 8010
    帮助 = 90
        帮助主题 = 9001
        WEB上的中联 = 9002
            中联主页 = 900201
            中联论坛 = 900203
            发送反馈 = 900202
        关于 = 9091
End Enum
