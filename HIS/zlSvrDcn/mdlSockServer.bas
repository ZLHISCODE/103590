Attribute VB_Name = "mdlSockServer"
Option Explicit

'**************************
'       OEM代号
'
'爱生    B0AEC9FA
'医业    D2BDD2B5
'托普    CDD0C6D5
'中软    D6D0C8ED
'金康泰  BDF0BFB5CCA9
'医院    D2BDD4BA
'**************************

Public Type POINTAPI
        x As Long
        Y As Long
End Type
'---------------------------------------------------------------
'- 注册表安全属性类型...
'---------------------------------------------------------------
Private Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Boolean
End Type

Public Const GWL_WNDPROC = -4
Private Const BIF_RETURNONLYFSDIRS = 1
Private Const BIF_DONTGOBELOWDOMAIN = 2
Private Const MAX_PATH = 260
Public Const SWP_NOSIZE = &H1
Public Const SWP_SHOWWINDOW = &H40
Public Const HWND_TOPMOST = -1
Public Const GWL_EXSTYLE = (-20)
Public Const WinStyle = &H40000
Public Const BF_LEFT = &H1
Public Const BF_TOP = &H2
Public Const BF_RIGHT = &H4
Public Const BF_BOTTOM = &H8
Public Const BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)
Public Const BDR_RAISEDOUTER = &H1
Public Const BDR_SUNKENINNER = &H8
Public Const BDR_SUNKENOUTER = &H2 '浅凹下
Public Const BDR_RAISEDINNER = &H4 '浅凸起
Public Const EDGE_RAISED = (BDR_RAISEDOUTER Or BDR_RAISEDINNER) '深凸起
Public Const EDGE_SUNKEN = (BDR_SUNKENOUTER Or BDR_SUNKENINNER) '深凹下
Public Const EDGE_ETCHED = (BDR_SUNKENOUTER Or BDR_RAISEDINNER) 'Frame边线样式
Public Const EDGE_BUMP = (BDR_RAISEDOUTER Or BDR_SUNKENINNER) '反Frame边线样式



Public Declare Function ClientToScreen Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long

Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Public Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    uId As Long
    uFlags As Long
    ucallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type

Public Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
Public Const NIM_ADD = &H0
Public Const NIM_MODIFY = &H1
Public Const NIM_DELETE = &H2
Public Const WM_MOUSEMOVE = &H200
Public Const NIF_MESSAGE = &H1
Public Const NIF_ICON = &H2
Public Const NIF_TIP = &H4

Public gstrProductName As String
Public gstrSysName As String                '系统名称
Public gstrUserName As String               '用户名
Public gstrUserPwd As String                  '密码
Public gstrServer As String                 '服务器名
Public gstrSQL    As String                 '通用的SQL语句变量

Public gcnOracle As ADODB.Connection     '公共数据库连接
Public gcnZltools As ADODB.Connection     'zltools连接对象,用于修改


Public Sub Main()
    Dim objLogin As Object
    
    '为实现XP风格，在显示窗体前必须执行该函数
    
    If App.PrevInstance Then
        MsgBox " 数据变动通知服务已经启动！ ", vbOKOnly, "提示"
        Exit Sub
    End If
    On Error Resume Next
    If objLogin Is Nothing Then
        Set objLogin = CreateObject("ZLLogin.clsLogin")
    End If
    If objLogin Is Nothing Then
        MsgBox "创建ZLLogin部件对象失败,请检查文件是否存在并且正确注册。"
        Exit Sub
    Else
        Set gcnOracle = objLogin.Login(2, CStr(Command()), , True)
        If gcnOracle Is Nothing Then
            Exit Sub
        ElseIf gcnOracle.State <> adStateOpen Then
            Exit Sub
        End If
    End If
    
    If Not IsDBA Then
        MsgBox "当前工具要求使用DBA登录。"
        Exit Sub
    End If
    gstrServer = objLogin.ServerName
    gstrUserName = objLogin.InputUser
    
    If IsDesinMode Then '编译环境 直接取HIS
        gstrUserPwd = "HIS"
    Else
        gstrUserPwd = GetDBPassword
    End If

    gstrSysName = GetSetting("ZLSOFT", "注册信息", "产品名称", "") & "软件"
    gstrProductName = GetSetting("ZLSOFT", "注册信息", "产品名称", "")
    frmMain.Show
End Sub

Private Function IsDBA() As Boolean
    Dim strSql As String, rsTemp As ADODB.Recordset
    
    On Error GoTo errH
    strSql = "SELECT 1 FROM SESSION_ROLES WHERE ROLE='DBA'"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "DBA判定")
    IsDBA = Not rsTemp.EOF
    
    Exit Function
errH:
    ErrCenter
End Function
Public Function IsDesinMode() As Boolean
'功能： 确定当前模式为设计模式
     Err = 0: On Error Resume Next
     Debug.Print 1 / 0
     If Err <> 0 Then
        IsDesinMode = True
     Else
        IsDesinMode = False
     End If
     Err.Clear: Err = 0
 End Function
 
 
Private Function GetDBPassword() As String
    '获取数据库密码
    Dim objRegister  As Object
    
    On Error Resume Next
    Set objRegister = CreateObject("zlRegister.clsRegister")
    If objRegister Is Nothing Then
        MsgBox "创建zlRegister部件对象失败,请检查文件是否存在并且正确注册。", vbExclamation, gstrSysName
        Exit Function
    End If
    Call SaveSetting("ZLSOFT", "公共全局", "升级程序", UCase("zlHisCrust.exe")) '用于ZLRegister中特殊判断
    GetDBPassword = objRegister.GetPassword(App.hInstance)
End Function

Public Function NVL(ByVal varValue As Variant, Optional DefaultValue As Variant = "") As Variant
'功能：相当于Oracle的NVL，将Null值改成另外一个预设值
    NVL = IIf(IsNull(varValue), DefaultValue, varValue)
End Function


Public Function ReplaceAll(vTar As String, vFind As String, vRep As String) As String
    Dim intPos As Long
    
    ReplaceAll = vTar
    intPos = InStr(ReplaceAll, vFind)
    
    While intPos > 0
        ReplaceAll = Replace(ReplaceAll, vFind, vRep)
        intPos = InStr(ReplaceAll, vFind)
    Wend
End Function


