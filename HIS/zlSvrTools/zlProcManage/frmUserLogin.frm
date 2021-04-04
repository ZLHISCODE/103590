VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmUserLogin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "工具登录"
   ClientHeight    =   2595
   ClientLeft      =   30
   ClientTop       =   330
   ClientWidth     =   4470
   Icon            =   "frmUserLogin.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2595
   ScaleWidth      =   4470
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton cmdSet 
      Caption         =   "配置服务器"
      Height          =   345
      Left            =   150
      TabIndex        =   11
      ToolTipText     =   "启动Oracle主机字符串配置程序"
      Top             =   2115
      Width           =   1095
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "…"
      Height          =   300
      Left            =   3720
      TabIndex        =   7
      TabStop         =   0   'False
      ToolTipText     =   "选择存在的服务器列表"
      Top             =   1455
      Width           =   300
   End
   Begin VB.TextBox txt数据库 
      Height          =   300
      IMEMode         =   2  'OFF
      Left            =   1920
      MaxLength       =   20
      TabIndex        =   6
      Top             =   1455
      Width           =   1785
   End
   Begin VB.TextBox txt密码 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   1920
      MaxLength       =   20
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   1050
      Width           =   2115
   End
   Begin VB.TextBox txt用户 
      Height          =   300
      IMEMode         =   2  'OFF
      Left            =   1920
      MaxLength       =   30
      TabIndex        =   2
      Top             =   645
      Width           =   2115
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   3135
      TabIndex        =   9
      Top             =   2115
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Default         =   -1  'True
      Height          =   350
      Left            =   1875
      TabIndex        =   8
      Top             =   2115
      Width           =   1100
   End
   Begin VB.Frame Frame1 
      Height          =   120
      Left            =   -150
      TabIndex        =   10
      Top             =   1860
      Width           =   4965
   End
   Begin MSComDlg.CommonDialog cdgFile 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lblNote 
      Caption         =   "    只有具有数据库DBA角色或相关系统的所有者才能使用本工具。"
      Height          =   375
      Left            =   990
      TabIndex        =   0
      Top             =   105
      Width           =   3195
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "口令"
      Height          =   180
      Left            =   1485
      TabIndex        =   3
      Top             =   1110
      Width           =   360
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "用户名"
      Height          =   180
      Left            =   1305
      TabIndex        =   1
      Top             =   705
      Width           =   540
   End
   Begin VB.Label lblDataBase 
      AutoSize        =   -1  'True
      Caption         =   "服务器"
      Height          =   180
      Left            =   1305
      TabIndex        =   5
      Top             =   1515
      Width           =   540
   End
   Begin VB.Image imgFlag 
      Height          =   720
      Left            =   180
      Picture         =   "frmUserLogin.frx":1CFA
      Top             =   105
      Width           =   720
   End
End
Attribute VB_Name = "frmUserLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim intTimes As Integer
Dim strNote As String
Dim strUserName As String
Dim strServerName As String
Dim strPassword As String

Dim mcolServer As New Collection

Private Sub cmdOK_Click()
    
    intTimes = intTimes + 1
    
    '------检验用户是否oracle合法用户----------------
    strUserName = Trim(txt用户.Text)
    strPassword = Trim(txt密码.Text)
    strServerName = Trim(txt数据库.Text)
    
    '有效字符串效验
    If Len(Trim(txt用户)) = 0 Then
        strNote = "请输入用户名。"
        txt用户.SetFocus
        GoTo InputError
    End If
    
    If Len(strUserName) <> 1 Then
        If Mid(strUserName, 1, 1) = "/" Or Mid(strUserName, 1, 1) = "@" Or Mid(strUserName, Len(strUserName) - 1, 1) = "/" Or Mid(strUserName, Len(strUserName) - 1, 1) = "@" Then
            txt用户.SetFocus
            strNote = "用户名错误。"
            Exit Sub
        End If
    End If
    If Trim(strPassword) <> "" And Len(strPassword) <> 1 Then
        If Mid(strPassword, Len(strPassword) - 1, 1) = "/" Or Mid(strPassword, Len(strPassword) - 1, 1) = "@" Or Mid(strPassword, 1, 1) = "/" Or Mid(strPassword, 1, 1) = "@" Then
            txt密码.SetFocus
            strNote = "口令错误。"
            GoTo InputError
        End If
    End If
    If Trim(strServerName) <> "" Then
        If Mid(strServerName, Len(strServerName) - 1, 1) = "/" Or Mid(strServerName, Len(strServerName) - 1, 1) = "@" Or Mid(strServerName, 1, 1) = "/" Or Mid(strServerName, 1, 1) = "@" Then
            strNote = "主机连接串错误。"
            txt数据库.SetFocus
            GoTo InputError
        End If
    End If
    
    '分离字符串
    Dim intPos As Integer
    intPos = InStr(1, strUserName, "@")
    If intPos > 0 Then
        strServerName = Mid(strUserName, intPos + 1)
        strUserName = Mid(strUserName, 1, intPos - 1)
    End If
    
    intPos = InStr(1, strUserName, "/")
    If intPos > 0 Then
        strPassword = Mid(strUserName, intPos + 1)
        strUserName = Mid(strUserName, 1, intPos - 1)
    End If
    
    intPos = InStr(1, strPassword, "@")
    If intPos > 0 Then
        strServerName = Mid(strPassword, intPos + 1)
        strPassword = Mid(strPassword, 1, intPos - 1)
    End If
    
    If Len(Trim(strPassword)) = 0 Then
        strNote = "未输入密码，不能注册。"
        txt密码.SetFocus
        GoTo InputError
    End If
        
    strUserName = UCase(strUserName)
    If Not OraDataOpen(strServerName, strUserName, strPassword) Then
        If Me.Visible = False Then Me.Visible = True
        txt密码.Text = ""
        Exit Sub
    End If
    '修改注册表
    SaveSetting "ZLSOFT", "注册信息\登陆信息", "MANAGER", strUserName
    SaveSetting "ZLSOFT", "注册信息\登陆信息", "SERVER", strServerName
    
    Unload Me
    Exit Sub
InputError:
    If intTimes > 3 Then
        MsgBox "超过三次注册失败，系统将自动退出。", vbExclamation, gstrSysName
        cmdCancel_Click
    Else
        If strNote <> "" Then
            MsgBox strNote, vbExclamation, gstrSysName
        End If
        Exit Sub
    End If

End Sub

Private Function OraDataOpen(ByVal strServerName As String, ByVal strUserName As String, ByVal strPassword As String) As Boolean
'功能： 打开指定的数据库连接，如果是普通用户，则使用管理员帐号重新打开连接
'参数：
'   strServerName：主机字符串
'   strUserName：用户名
'   strUserPwd：密码
'返回： 数据库打开成功，返回true；失败，返回false
    Dim blnLogin As Boolean, blnTransPassword As Boolean
    Dim strError As String
    Dim objRegister35 As Object
    
    Set objRegister35 = CreateObject("zlRegister.clsRegister")
    blnTransPassword = Not (strUserName = "SYS" Or strUserName = "SYSTEM")
    If Not objRegister35 Is Nothing Then '检测当前环境
       Set gcnOracle = objRegister35.GetConnection(strServerName, strUserName, strPassword, blnTransPassword, OraOLEDB, strError)
       Set gcnOldOra = objRegister35.GetConnection(strServerName, strUserName, strPassword, blnTransPassword, MSODBC, strError)
    Else
        '支持strServerName = "192.168.2.13:1521/dyyy"这种格式
        Set gcnOracle = gobjRegister.GetConnection(strServerName, strUserName, strPassword, blnTransPassword, OraOLEDB, strError)
        Set gcnOldOra = gobjRegister.GetConnection(strServerName, strUserName, strPassword, blnTransPassword, MSODBC, strError)
    End If
    
    If gcnOracle.State = adStateClosed Then
        If InStr(strError, "ORA-00604") > 0 Then
            If InStr(strError, "ORA-20002") > 0 Then
                strError = "当前用户不能使用该应用登录数据库，请联系管理员。"
            Else
                strError = "当前用户被禁止登录数据库，请联系管理员。"
            End If
        End If
        MsgBox strError, vbInformation, gstrSysName
        OraDataOpen = False
        Exit Function
    End If

    OraDataOpen = True
    gstrUserName = strUserName
End Function

Private Sub cmdCancel_Click()
    Set gcnOracle = Nothing
    Unload Me
End Sub


Private Sub cmdSelect_Click()
    Dim strServer As String
    Dim p As POINTAPI
    
    p.x = txt数据库.Left / Screen.TwipsPerPixelX
    p.y = (cmdSelect.Top + cmdSelect.Height) / Screen.TwipsPerPixelY
    ClientToScreen Me.hwnd, p
    
    strServer = frmServerSelect.GetServer(mcolServer, p.x * Screen.TwipsPerPixelX, p.y * Screen.TwipsPerPixelY, txt数据库.Text)
    If strServer <> "" Then
        txt数据库.Text = strServer
        txt数据库.SetFocus
    End If
End Sub

Private Sub Form_Activate()
    Dim LngStyle As Long
    
    '设置当前窗口在任务栏显示
    LngStyle = GetWindowLong(Me.hwnd, GWL_EXSTYLE)
    LngStyle = LngStyle Or WinStyle
    Call SetWindowLong(Me.hwnd, GWL_EXSTYLE, LngStyle)
    
    ShowWindow Me.hwnd, 0 '先隐藏
    ShowWindow Me.hwnd, 1 '再显示
        
    If Len(txt用户) <> 0 Then
        txt密码.SetFocus
    End If
End Sub

Private Sub Form_Load()
    Dim strFileInfo As String
    txt用户.Text = GetSetting("ZLSOFT", "注册信息\登陆信息", "MANAGER", "")
    txt数据库.Text = GetSetting("ZLSOFT", "注册信息\登陆信息", "SERVER", "")
    intTimes = 0
    
    '不加这一句的话，由于已显示frmSplash窗体，在开启输入法的情况下，启动源程序，不会显示登录窗口，VB只能异常终止退出
    SetActiveWindow Me.hwnd
    
    Set mcolServer = LoadServer(strFileInfo)
    txt数据库.ToolTipText = strFileInfo
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = 0 Then
        Set gcnOracle = Nothing
    End If
End Sub

Private Sub txt数据库_KeyPress(KeyAscii As Integer)
    If KeyAscii <> vbKeyReturn Then
        '回车键另行处理
        If KeyAscii <> vbKeyBack Then
            Call AppendText(KeyAscii)
        End If
    End If
End Sub

Private Sub txt用户_GotFocus()
    If Me.ActiveControl Is txt用户 Then
        SelAll txt用户
        OpenIme False
    End If
End Sub

Private Sub TXT密码_GotFocus()
    SelAll txt密码
End Sub

Private Sub txt数据库_GotFocus()
    If Me.ActiveControl Is txt数据库 Then
        SelAll txt数据库
        OpenIme False
    End If
End Sub

Private Sub cmdSet_Click()
    Dim strPath As String   'Oracle安装目录
    Dim strCommond As String, strError As String
    
    strPath = GetOracleHomePath(strError)
    If strPath = "" Then
        MsgBox "本机的Oracle是否正常安装，请检查。" & vbCrLf & strError, vbInformation, "提示"
        Exit Sub
    End If
    
    '执行Oracle 8 的Net Easy配置的程序
    strCommond = strPath & "\BIN\N8SW.EXE"
    If ExecuteCommand(strCommond) = True Then
        '已经成功
        Exit Sub
    End If
    
    '执行Oracle 8i,9i,10g,11g的Net Easy配置的程序
    strCommond = strPath & "\BIN\launch.exe """ & strPath & "\network\tools"" " & strPath & "\network\tools\netca.cl"
    If ExecuteCommand(strCommond) = True Then
        '已经成功
        Exit Sub
    End If
    
End Sub

Private Function GetOracleHomePath(ByVal strError As String) As String
'功能：获取OracleHome路径
    Dim strPath As String
    Dim strServer As String, strComputer As String, strSID As String
    Dim arrTmp As Variant
    Dim rsOraHome As ADODB.Recordset
    Dim intVersion As Integer, intTimes As Integer, intServer As Integer
    Dim i As Long, blnRead As Boolean

    Set rsOraHome = New ADODB.Recordset
    With rsOraHome
        .Fields.Append "Name", adVarChar, 256 'Name
        .Fields.Append "VerSion", adInteger  '版本
        .Fields.Append "Times", adInteger '第几次安装
        .Fields.Append "Server", adInteger '1-服务器,2-客户端
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
        
        '1:读取64位下32目录会自动定位到SOFTWARE\Wow6432Node\Oracle 2：读取32位下32位目录
        arrTmp = GetAllSubKey(HKEY_LOCAL_MACHINE, "SOFTWARE\Oracle")
        
        If TypeName(arrTmp) = "Empty" Then
            If Is64bit Then
                strError = "没有找到注册表项HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\Oracle！"
            Else
                strError = "没有找到注册表项HKEY_LOCAL_MACHINE\SOFTWARE\Oracle！"
            End If
        Else
            For i = LBound(arrTmp) To UBound(arrTmp)
                If UCase(arrTmp(i)) Like "KEY_ORA*HOME*" Then
                    intVersion = 0: intTimes = 0:  intServer = 1
                    If GetOraInfoByRegKey(arrTmp(i), intVersion, intTimes, intServer) Then
                        .AddNew Array("Name", "VerSion", "Times", "Server"), Array("\" & arrTmp(i), intVersion, intTimes, intServer)
                        .Update
                    End If
                End If
            Next
            If UBound(arrTmp) <> -1 Then ''顶级目录可能有Oracle_Home信息，默认读取这个
                .AddNew Array("Name", "VerSion", "Times", "Server"), Array("", 0, 0, 1): .Update
            End If
            .Sort = "VerSion Desc,Times Desc,Server"    '高版本优先
            Do While Not .EOF
                strPath = ""
                blnRead = Not GetRegValue("HKEY_LOCAL_MACHINE\SOFTWARE\Oracle" & !Name, "ORACLE_HOME", strPath)
                blnRead = blnRead Or strPath = "" And !Name & "" = ""
                If blnRead Then
                    Call GetRegValue("HKEY_LOCAL_MACHINE\SOFTWARE\Oracle", "ORA_CRS_HOME", strPath)
                End If
                If strPath <> "" Then
                    GetOracleHomePath = strPath
                    Exit Function
                End If
                
                .MoveNext
            Loop
        End If
    End With
End Function

Private Function GetOraInfoByRegKey(ByVal strOraHome As String, ByRef intVer As Integer, ByRef intTimes As Integer, ByRef intServer As Integer) As Boolean
'功能:通过OracleHome键获取Oracle信息
    Dim arrTmp As Variant
    Dim i As Long, blnRetrun As Boolean
    'KEY_OraDb11g_home1_32bit
    'Key_Ora*版本Home_32Bit
    'Key_Ora*版本_Home*
    arrTmp = Split(UCase(strOraHome), "_")
    For i = 1 To UBound(arrTmp)
        If arrTmp(i) Like "HOME*" Then
            intTimes = ValEx(arrTmp(2))
            blnRetrun = True
        ElseIf arrTmp(i) Like "*HOME*" Then
            intTimes = Val(Mid(arrTmp(1), InStr(UCase(arrTmp(1)), "HOME") + 4))
            blnRetrun = True
        End If
        If arrTmp(i) Like "ORADB*" Then
            intVer = ValEx(Mid(arrTmp(1), 6))
            intServer = 1
            blnRetrun = True
        ElseIf arrTmp(i) Like "ORACLIENT*" Then
            intVer = ValEx(Mid(arrTmp(1), 10))
            intServer = 2
            blnRetrun = True
        ElseIf arrTmp(i) Like "*CLIENT*" Then
            intServer = 2
            intVer = ValEx(arrTmp(i))
            blnRetrun = True
        End If
    Next
    GetOraInfoByRegKey = blnRetrun
End Function

Private Function ExecuteCommand(ByVal strCommand As String) As Boolean
'功能：执行指定命令
    Dim lngShell As Long, lngProcess As Long
    
    On Error Resume Next
    lngShell = Shell(strCommand, vbNormalFocus)
    
    If Err <> 0 Then
        Exit Function
    End If
    
    ExecuteCommand = True
End Function

Private Sub AppendText(KeyAscii As Integer)
'功能：向TextBox控件的Text追加内容，并根据当前Text的值在列表中检索可用的完整项目
'参数：KeyAscii    当前的按键
    Dim strTemp As String
    Dim strInput As String
    Dim lngIndex As Long, lngStart As Long
    Dim varItem As Variant
    
    '首先当前用户输入的字符
    If KeyAscii < 0 Or InStr("ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789.", UCase(Chr(KeyAscii))) > 0 Then
        '输入字符只能是数字、英文和汉字
        strInput = Chr(KeyAscii)
        KeyAscii = 0
    End If
    
    With txt数据库
        '记录上次的插入点位置
        lngStart = .SelStart + IIf(strInput <> "", 1, 0)
        '接着得到用户击键完成后文本框中出现的内容
        strInput = Mid(.Text, 1, .SelStart) & strInput & Mid(.Text, .SelStart + .SelLength + 1)
    End With
    '根据假想的内容得到可能的列表项
    strTemp = ""
    For Each varItem In mcolServer
        If UCase(varItem(0)) Like UCase(strInput & "*") Then
            strTemp = varItem(0)
        End If
    Next
    If strTemp <> "" Then
        txt数据库.Text = strTemp
        txt数据库.SelStart = Len(strInput)
        txt数据库.SelLength = 100
    Else
        txt数据库.Text = strInput
        txt数据库.SelStart = lngStart
    End If

End Sub

