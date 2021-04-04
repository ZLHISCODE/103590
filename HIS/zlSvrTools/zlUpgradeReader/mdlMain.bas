Attribute VB_Name = "mdlMain"
Option Explicit
Public gstrDBUser As String
Public gcnOracle As ADODB.Connection
Public gstrSysname As String '程序名称

Public gstrSystems As String '系统名称
Public gstr用户单位名称 As String '已登录时不为空

Public mclsAppTool As New zl9AppTool.clsAppTool
Public rsMenu As ADODB.Recordset
Public rsMenuPEIS As ADODB.Recordset

'-------------------------------------------------------------
Public Declare Sub InitCommonControls Lib "comctl32.dll" ()
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long


Public Const WinStyle = &H40000

'---读写INI文件的API声明
#If Win32 Then
   Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Integer, ByVal lpFileName As String) As Integer
   Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal Appname As String, ByVal KeyName As Any, ByVal NewString As Any, ByVal Filename As String) As Integer
#Else
   Private Declare Function GetPrivateProfileString Lib "Kernel" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Integer, ByVal lpFileName As String) As Integer
   Private Declare Function WritePrivateProfileString Lib "Kernel" (ByVal Appname As String, ByVal KeyName As Any, ByVal NewString As Any, ByVal Filename As String) As Integer
#End If
'----------------------
Public gobjRegister As Object               '注册授权部件zlRegister
Public Enum 工具清单
    导航功能清单 = 10
    字典管理工具 = 11
    消息收发工具 = 12
    系统选项设置 = 13
    EXCEL报表工具 = 14
    本地参数管理 = 15
End Enum

Public Sub Main()
    Dim objLogin As Object
    
    gstrDBUser = ""
    gstrSysname = "升级说明阅读器"
    gstr用户单位名称 = ""
    On Error Resume Next
    If objLogin Is Nothing Then
        Set objLogin = CreateObject("ZLLogin.clsLogin")
    End If
    If objLogin Is Nothing Then
        Set gcnOracle = New ADODB.Connection
    Else
        Set gcnOracle = objLogin.Login(0, CStr(Command()))
        If gcnOracle Is Nothing And Not objLogin.IsCancel Then
            Exit Sub
        ElseIf gcnOracle Is Nothing Then '取消退出，以非登陆模式进入
            Set gcnOracle = New ADODB.Connection
        End If
    End If
    
    If gcnOracle.State = adStateOpen Then
        gstrSystems = objLogin.Systems
        gstrDBUser = objLogin.DBUser
        Set rsMenu = MenuGranted(objLogin.MenuGroup)
        Set rsMenuPEIS = MenuGranted("PEIS")
        
        If rsMenu.EOF Then
            MsgBox "您没有操作任何系统的权限,程序被迫退出！", vbInformation, gstrSysname
            Exit Sub
        End If
        gstr用户单位名称 = zlRegInfo("单位名称", , -1)
        Call frmMain.Show_me(1) '0- 未登录方式 1－已登录方式
    Else
        Call frmMain.Show_me(0) '0- 未登录方式 1－已登录方式
    End If
End Sub

Private Function MenuGranted(ByVal strMenuGroup As String) As ADODB.Recordset
    '-------------------------------------------------------------
    '功能：分析授权使用并安装的部件，进而产生授权使用的菜单集合
    '参数：注册码
    '-------------------------------------------------------------
    Dim ArrCommand
    Dim StrSQL As String
    Dim rsTemp As New ADODB.Recordset
    Dim strCodes As String
    Dim strObjs As String
    Dim IntCount As Integer
    Dim strSystems As String
    Dim gstrMenuSys As String
    Dim BlnOnlySys As Boolean '只有报表系统
    Dim strSYS As String
    
    BlnOnlySys = (gstrSystems = "REPORT")
    If BlnOnlySys Then
        strSystems = " '0'"
    Else
        strSystems = Replace(gstrSystems, "','", ",")
    End If
    
    '--分析权限菜单--
    With rsTemp
        If strMenuGroup <> "" Then gstrMenuSys = strMenuGroup
        strObjs = GetSetting("ZLSOFT", "注册信息", "本机部件", "")
        If strObjs = "" Then strObjs = "'Zl9Common'"
        strObjs = Replace(strObjs, "','", ",")

        StrSQL = "SELECT 层次, Id AS 编号, Nvl(上级id, 0) AS 上级, 标题, Decode(Nvl(短标题,'空'),'空',标题,短标题) As 短标题, 快键, 说明, Nvl(模块, 0) AS 模块, Nvl(系统, 0) AS 系统, " & _
                 "        Nvl(图标, 0) AS 图标, nvl(部件,'0') as 部件, Decode(Upper(Rtrim(部件)), 'ZL9REPORT', 1, 0) AS 报表 " & _
                 " FROM TABLE(CAST(Zltools.f_Reg_Menu('" & gstrMenuSys & "', " & strSystems & ", " & strObjs & ") As " & _
                 " Zltools.t_Menu_Rowset)) " & _
                 " ORDER BY 层次, Id"

        If .State = adStateOpen Then .Close
        .Open StrSQL, gcnOracle, adOpenKeyset
    End With
    
    Set MenuGranted = rsTemp
    
End Function

Public Sub WriteToIni(ByVal Filename As String, ByVal Section As String, ByVal Key As String, ByVal Value As String)
''写INI文件
    Dim buff As String * 128
    buff = Trim(Value) + Chr(0)
    WritePrivateProfileString Section, Key, buff, Filename

End Sub

Public Function ReadFromIni(ByVal Filename As String, ByVal Section As String, ByVal Key As String) As String
''读INI文件
    Dim i As Long
    Dim buff As String * 128
    GetPrivateProfileString Section, Key, "", buff, 128, Filename
    i = InStr(buff, Chr(0))
    ReadFromIni = Trim(Left(buff, i - 1))
End Function
