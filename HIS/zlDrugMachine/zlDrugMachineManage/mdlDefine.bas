Attribute VB_Name = "mdlDefine"
Option Explicit

'-----------------------------------------
'ȫ�ֳ�����ȫ�ֱ�����API�����ȶ���ģ��
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

Public Const GSTR_MSG As String = "ҩƷ�Զ����豸������"
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
Public Const REG_SZ = 1                         ' Unicode���ս��ַ���
Public Const REG_EXPAND_SZ = 2                  ' Unicode���ս��ַ���
Public Const REG_DWORD = 4                      ' 32-bit ����

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
    ����ID As Long
    ��� As String
    ���� As String
    ���� As String
    �û��� As String
End Type
Public UserInfo As TYPE_USER_INFO

Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public Type TYPE_PARAMS
    ��ʱ���� As Integer
    ��Ч���� As Integer
    ��ʾ������� As Integer
    �����־ As Boolean
    ��ϸ��־ As Boolean
    ������־���� As Integer
End Type

Public Enum enuEditState
    �鿴 = 0
    ���� = 1
    �޸� = 2
End Enum

Public Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

'֧�ֵĽӿ�����
Public Const GSTR_TYPE As String = "1-Τ�ֺ���|2-TOSHO|3-����|4-������|5-YUYAMA|6-��԰|7-���ݺ��"

Public Enum enuMenus
    �ļ� = 10
        ��ӡ���� = 1001
        ��ӡԤ�� = 1002
        ��ӡ = 1003
        ���Excel = 1004
        �������� = 1005
        �˳� = 1099
    �༭ = 20
        ���� = 2001
        �޸� = 2002
        ɾ�� = 2003
    ���� = 30
        ���� = 3011
        ͣ�� = 3012
        �豸�ӿڹ��� = 3021
        �������ݴ��� = 3022
        ��ʾ = 3031
        ���� = 3032
    �鿴 = 80
        ������ = 8001
            ��׼��ť = 800101
            �ı���ǩ = 800102
            ��ͼ�� = 800103
        ״̬�� = 8002
        ˢ�� = 8010
    ���� = 90
        �������� = 9001
        WEB�ϵ����� = 9002
            ������ҳ = 900201
            ������̳ = 900203
            ���ͷ��� = 900202
        ���� = 9091
End Enum
