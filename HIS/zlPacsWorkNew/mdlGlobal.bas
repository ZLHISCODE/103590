Attribute VB_Name = "mdlGlobal"
Option Explicit
 

Private mstrInstitution As String
Private mstrSysRootPath As String
Private mobjDcmGlobal As DicomGlobal

Private hHook As Long
Private hFormhWnd As Long

'全局属性方法

'注册的单位名称
Property Get RegInstitution() As String
    If Len(mstrInstitution) <= 0 Then
        mstrInstitution = zlRegInfo("单位名称")
        
        If Len(mstrInstitution) <= 0 Then mstrInstitution = "未注册"
    End If
    
    RegInstitution = mstrInstitution
End Property

'系统路径
Property Get SysRootPath() As String
    If Len(mstrSysRootPath) <= 0 Then mstrSysRootPath = GetAppRootPath
    
    SysRootPath = mstrSysRootPath
End Property

Property Let SysRootPath(value As String)
    mstrSysRootPath = value
End Property


'获取私有注册表路径
Public Function GetPrivateRegPath(ByVal strItemName As String) As String
    GetPrivateRegPath = "私有模块\" & UserInfo.用户名 & "\" & App.EXEName & "\界面设置\" & strItemName
End Function

'获取公共注册表路径部分
Public Function GetPublicRegPath(ByVal strItemName As String) As String
    GetPublicRegPath = "公共模块\" & App.EXEName & "\" & strItemName
End Function



'创建UID
Public Function CreateUID() As String
    If mobjDcmGlobal Is Nothing Then
        Set mobjDcmGlobal = New DicomGlobal
        mobjDcmGlobal.RegString("UIDRoot") = "1"
    End If
    
    CreateUID = mobjDcmGlobal.NewUID
End Function


'检查进程是否运行，exeName 参数是要检查的进程 exe 名字，比如 VB6.EXE
Public Function CheckExeIsRun(ByVal strExeName As String) As Boolean
    Dim objWMIService As Object
    Dim colProcessList As Object
    
On Error Resume Next

    CheckExeIsRun = False
    
    Set objWMIService = VBA.GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
    Set colProcessList = objWMIService.ExecQuery("Select name from Win32_Process Where Name='" & strExeName & "'")
    
    CheckExeIsRun = IIf(colProcessList.Count > 0, True, False)
    
    Set colProcessList = Nothing
    Set objWMIService = Nothing
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''
''为了处理双屏时对话框的正确显示位置，用API函数改写了一下MsgBox函数
'''''''''''''''''''''''''''''''''''''''''''''''''
Public Function MsgBoxD(objParent As Object, ByVal strPrompt As String, Optional ByVal dwStyle As VbMsgBoxStyle = MB_OK, Optional strTitle As String = "") As Long

    Dim lngHwnd As Long
 
    If objParent Is Nothing Then
        lngHwnd = GetActiveWindow
    Else
        lngHwnd = objParent.hwnd
    End If

    If lngHwnd = GetDesktopWindow Or lngHwnd = 0 Then
        lngHwnd = GetForegroundWindow
    End If
 

    MsgBoxD = MsgboxH(lngHwnd, strPrompt, dwStyle, strTitle)

End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''
'替代VB中的Msgbox函数
'''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function MsgboxH(hwnd As Long, strPrompt As String, Optional ByVal dwStyle As VbMsgBoxStyle = MB_OK, Optional strTitle As String) As Long

    Dim hInstance As Long
    Dim hThreadId As Long

    hInstance = App.hInstance
    hThreadId = App.ThreadID

    If dwStyle = 0 Then dwStyle = vbOKOnly
    If Len(strTitle) = 0 Then strTitle = App.EXEName

    '将当前窗口的句柄付给变量
    hFormhWnd = hwnd

    '设置钩子
    hHook = SetWindowsHookEx(WH_CBT, AddressOf BoxPro, hInstance, hThreadId)
    
    '调用MessageBox API
    MsgboxH = MessageBox(hwnd, strPrompt, strTitle, dwStyle)
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''
'HOOK处理
'''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function BoxPro(ByVal nCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

    '变量声明
    Dim rc As RECT
    Dim rcFrm As RECT

    Dim newLeft As Long
    Dim newTop As Long
    Dim dlgWidth As Long
    Dim dlgHeight As Long
    Dim scrWidth As Long
    Dim scrHeight As Long
    Dim frmLeft As Long
    Dim frmTop As Long
    Dim frmWidth As Long
    Dim frmHeight As Long
    Dim hwndMsgBox As Long

    '当MessageBox出现时，将Msgbox对话框居中与所在的窗口
    If nCode = HCBT_ACTIVATE Then
    
        '消息为HCBT_ACTIVATE时，参数wParam包含的是MessageBox的句柄
        hwndMsgBox = wParam
        
        '得到MessageBox对话框的Rect
        Call GetWindowRect(hwndMsgBox, rc)
        Call GetWindowRect(hFormhWnd, rcFrm)
        
        If rcFrm.Right = 0 Or rcFrm.Bottom = 0 Then
            Call GetWindowRect(GetDesktopWindow, rcFrm)
        End If
        
        '使MessageBox居中
        frmLeft = rcFrm.Left
        frmTop = rcFrm.Top
        frmWidth = rcFrm.Right - rcFrm.Left
        frmHeight = rcFrm.Bottom - rcFrm.Top
        dlgWidth = rc.Right - rc.Left
        dlgHeight = rc.Bottom - rc.Top
    
        scrWidth = Screen.Width \ Screen.TwipsPerPixelX
        scrHeight = Screen.Height \ Screen.TwipsPerPixelY
    
        newLeft = frmLeft + ((frmWidth - dlgWidth) \ 2)
        newTop = frmTop + ((frmHeight - dlgHeight) \ 2)
        
'        '修改确定按钮的文字
'        Call SetDlgItemText(hwndMsgBox, IDOK, "这是确定按钮")
        SetWindowPos hwndMsgBox, -1, rcFrm.Left, rcFrm.Top, dlgWidth, dlgHeight, 3 '将窗口置顶
        SetForegroundWindow hwndMsgBox
        
        'Msgbox居中
        Call MoveWindow(hwndMsgBox, newLeft, newTop, dlgWidth, dlgHeight, True)
        
        '卸载钩子
        UnhookWindowsHookEx hHook
    End If
    
    BoxPro = False
End Function


Public Function MainForm() As Object
    Dim objForm As Object
    
    Set MainForm = Nothing
    
    If Forms.Count <= 0 Then Exit Function
    
    For Each objForm In Forms
        If InStr(objForm.Name, "PacsMain") > 0 Then
            Set MainForm = objForm
            Exit Function
        End If
    Next
    
    Set MainForm = Forms(0)
End Function
