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

Public Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long

Public Type NOTIFYICONDATA
    cbSize As Long
    hWnd As Long
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
Public gstrServer As String                 '服务器名
Public gstrSQL    As String                 '通用的SQL语句变量

Public gcnOracle As ADODB.Connection     '公共数据库连接

Public Sub Main()
    Dim objLogin As Object
    '为实现XP风格，在显示窗体前必须执行该函数
    
    If App.PrevInstance Then
        MsgBox " 自动提醒服务已经启动！ ", vbOKOnly, "自动提醒"
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
    gstrServer = objLogin.ServerName
    gstrUserName = objLogin.InputUser
    gstrSysName = GetSetting("ZLSOFT", "注册信息", "产品名称", "") & "软件"
    gstrProductName = GetSetting("ZLSOFT", "注册信息", "产品名称", "")
    frmMain.Show
End Sub

Public Function NVL(ByVal varValue As Variant, Optional DefaultValue As Variant = "") As Variant
'功能：相当于Oracle的NVL，将Null值改成另外一个预设值
    NVL = IIf(IsNull(varValue), DefaultValue, varValue)
End Function

Public Sub ApplyOEM_Picture(objPicture As Object, ByVal str属性 As String)
'针对各种图标应用OEM策略
    Dim strOEM As String
    Dim blnCorp As Boolean
    On Error Resume Next
    
    If gstrProductName <> "-" Then
        '处理状态栏图标的OEM策略
        If gstrProductName <> "中联" Then
            If Right(str属性, 1) = "B" Then
                '表示产品图片
                blnCorp = False
                str属性 = Mid(str属性, 1, Len(str属性) - 1)
            Else
                '表示公司徽标
                blnCorp = True
            End If
            
            strOEM = GetOEM(gstrProductName, blnCorp)
            If str属性 = "Picture" Then
                Set objPicture.Picture = LoadCustomPicture(strOEM)
            ElseIf str属性 = "Icon" Then
                Set objPicture.Icon = LoadCustomPicture(strOEM)
            End If
            
            If Err <> 0 Then
                Err.Clear
            End If
        End If
    End If
End Sub

Public Function GetOEM(ByVal strAsk As String, Optional ByVal blnCorp As Boolean = True) As String
    '-------------------------------------------------------------
    '功能：返回每个字线的ASCII码
    '参数：
    '返回：
    '-------------------------------------------------------------
    Dim intBit As Integer, iCount As Integer, blnCan As Boolean
    Dim strCode As String
    
    'OEM图片有两种类型 ，一是指公司徽标，另一个是产品标识
    strCode = IIf(blnCorp = True, "OEM_", "PIC_")
    For intBit = 1 To Len(strAsk)
        '取每个字的ASCII码
        strCode = strCode & Hex(Asc(Mid(strAsk, intBit, 1)))
    Next
    GetOEM = strCode
End Function

Public Sub ApplyOEM(objStatus As Object)
'针对状态栏应用OEM策略
    Dim strOEM As String
    On Error Resume Next
    
    If gstrProductName <> "-" Then
        objStatus.Panels(1).Text = gstrProductName & "软件"
        '处理状态栏图标的OEM策略
        If gstrProductName = "中联" Then
            Set objStatus.Panels(1).Picture = LoadCustomPicture("Logo")
        Else
            strOEM = GetOEM(gstrProductName)
            Set objStatus.Panels(1).Picture = LoadCustomPicture(strOEM)
            If Err <> 0 Then
                Err.Clear
                Set objStatus.Panels(1).Picture = LoadCustomPicture("Logo")
            End If
        End If
        objStatus.Panels(1).ToolTipText = ""
        objStatus.Height = 360
    End If
End Sub

Public Function LoadCustomPicture(strID As String) As StdPicture
'功能:将资源文件中的指定资源生成磁盘文件
'参数:ID=资源号,strExt=要生成文件的扩展名(如BMP)
'返回:生成文件名
    Dim arrData() As Byte
    Dim intFile As Integer
    Dim strFile As String * 255, strR As String
    
    arrData = LoadResData(strID, "CUSTOM")
    intFile = FreeFile
    
    GetTempPath 255, strFile
    strR = Trim(Left(strFile, InStr(strFile, Chr(0)) - 1)) & CLng(Timer * 100) & ".pic"

    Open strR For Binary As intFile
    Put intFile, , arrData()
    Close intFile
    Set LoadCustomPicture = VB.LoadPicture(strR)
    Kill strR
End Function

Public Sub AddIcon(ByVal lngHwnd As Long, ByVal stdIcon As StdPicture, Optional ByVal strTip As String = "")
    
    '功能：在任务栏上增加一个图标
    
    Dim t As NOTIFYICONDATA
    
    On Error Resume Next
    
    t.cbSize = Len(t)
    t.hWnd = lngHwnd   '事件发生的载体，为了不与其它鼠标事件相冲突，所以单独放一个控件
    t.uId = 1&
    t.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
    t.ucallbackMessage = WM_MOUSEMOVE
    t.hIcon = stdIcon
    t.szTip = strTip & Chr$(0)

    Shell_NotifyIcon NIM_ADD, t
    
End Sub

Public Sub RemoveIcon(ByVal lngHwnd As Long)
    
    '功能：从任务栏上删除图标
    
    Dim t As NOTIFYICONDATA
    
    On Error Resume Next
    
    t.cbSize = Len(t)
    t.hWnd = lngHwnd   '事件发生的载体
    t.uId = 1&
    
    Shell_NotifyIcon NIM_DELETE, t
    
End Sub

Public Function ReplaceAll(vTar As String, vFind As String, vRep As String) As String
    Dim intPos As Long
    
    ReplaceAll = vTar
    intPos = InStr(ReplaceAll, vFind)
    
    While intPos > 0
        ReplaceAll = Replace(ReplaceAll, vFind, vRep)
        intPos = InStr(ReplaceAll, vFind)
    Wend
End Function


