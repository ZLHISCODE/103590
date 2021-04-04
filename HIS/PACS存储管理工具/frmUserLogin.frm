VERSION 5.00
Begin VB.Form frmUserLogin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "操作员登录"
   ClientHeight    =   2205
   ClientLeft      =   30
   ClientTop       =   330
   ClientWidth     =   4170
   Icon            =   "frmUserLogin.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2205
   ScaleWidth      =   4170
   StartUpPosition =   2  '屏幕中心
   Begin VB.ComboBox cmb数据库 
      Height          =   300
      Left            =   1950
      Sorted          =   -1  'True
      TabIndex        =   5
      Top             =   1050
      Width           =   1920
   End
   Begin VB.Frame Frame1 
      Height          =   120
      Left            =   -360
      TabIndex        =   9
      Top             =   1455
      Width           =   5025
   End
   Begin VB.CommandButton cmd修改密码 
      Caption         =   "修改密码(&M)"
      Height          =   350
      Left            =   180
      Style           =   1  'Graphical
      TabIndex        =   8
      TabStop         =   0   'False
      ToolTipText     =   "单击此处修改密码"
      Top             =   1710
      Width           =   1335
   End
   Begin VB.CommandButton CMD放弃 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   2865
      TabIndex        =   7
      Top             =   1710
      Width           =   1100
   End
   Begin VB.CommandButton CMD确认 
      Caption         =   "确定(&O)"
      Default         =   -1  'True
      Height          =   350
      Left            =   1755
      TabIndex        =   6
      Top             =   1710
      Width           =   1100
   End
   Begin VB.TextBox TXT密码 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   1950
      MaxLength       =   20
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   630
      Width           =   1920
   End
   Begin VB.TextBox txt用户 
      Height          =   300
      Left            =   1950
      TabIndex        =   1
      Top             =   195
      Width           =   1920
   End
   Begin VB.Image imgFlag 
      Height          =   720
      Left            =   315
      Picture         =   "frmUserLogin.frx":1CFA
      Top             =   240
      Width           =   720
   End
   Begin VB.Label Lbl服务器 
      AutoSize        =   -1  'True
      Caption         =   "服务器"
      Height          =   180
      Left            =   1320
      TabIndex        =   4
      Top             =   1110
      Width           =   540
   End
   Begin VB.Label Lbl口令 
      AutoSize        =   -1  'True
      Caption         =   "密码"
      Height          =   180
      Left            =   1500
      TabIndex        =   2
      Top             =   690
      Width           =   360
   End
   Begin VB.Label Lbl用户名 
      AutoSize        =   -1  'True
      Caption         =   "用户名"
      Height          =   180
      Left            =   1320
      TabIndex        =   0
      Top             =   255
      Width           =   540
   End
End
Attribute VB_Name = "frmUserLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'命令行格式：
'zlhis90.exe 菜单
'zlhis90.exe 用户名/密码        此种情况不需要进行密码转换
'zlhis90.exe 用户名 密码
'zlhis90.exe 用户名 密码 菜单
Public mblnChangePass As Boolean
Private mblnShowChangePassFrm As Boolean
Private mblnFirst As Boolean  '为True表示已经正常显示出
Private mintTimes As Integer  '登录重试次数
Private mbln转换 As Boolean     '表示传入的密码是否为数据库密码，是否不需要再转换
Private mcolServer As New Collection  '保存服务器串列表

Private Sub CMD确认_Click()
    Dim strNote As String
    Dim strUserName As String
    Dim strServerName As String
    Dim strPassword As String
    On Error GoTo InputError
    SetConState False
    mintTimes = mintTimes + 1
    
    '------检验用户是否oracle合法用户----------------
    strUserName = Trim(txt用户.Text)
    If mblnChangePass = False Then
        strPassword = Trim(TXT密码.Text)
    Else
        strPassword = Trim(FrmChangePass.TXT原密码.Text)
    End If
    strServerName = Trim(cmb数据库.Text)
    
    '有效字符串效验
    If Len(Trim(txt用户)) = 0 Then
        strNote = "请输入用户名"
        txt用户.SetFocus
        GoTo InputError
    End If
    
    If Len(strUserName) <> 1 Then
        If Mid(strUserName, 1, 1) = "/" Or Mid(strUserName, 1, 1) = "@" Or Mid(strUserName, Len(strUserName) - 1, 1) = "/" Or Mid(strUserName, Len(strUserName) - 1, 1) = "@" Then
            txt用户.SetFocus
            strNote = "用户名错误"
            SetConState
            Exit Sub
        End If
    End If
    If Trim(strPassword) <> "" And Len(strPassword) <> 1 Then
        If Mid(strPassword, Len(strPassword) - 1, 1) = "/" Or Mid(strPassword, Len(strPassword) - 1, 1) = "@" Or Mid(strPassword, 1, 1) = "/" Or Mid(strPassword, 1, 1) = "@" Then
            If TXT密码.Enabled Then TXT密码.SetFocus
            strNote = "密码错误"
            GoTo InputError
        End If
    End If
    If Trim(strServerName) <> "" Then
        If Mid(strServerName, Len(strServerName) - 1, 1) = "/" Or Mid(strServerName, Len(strServerName) - 1, 1) = "@" Or Mid(strServerName, 1, 1) = "/" Or Mid(strServerName, 1, 1) = "@" Then
            strNote = "主机连接串错误"
            cmb数据库.SetFocus
            GoTo InputError
        End If
    End If
    
    '分离字符串
    Dim intPos As Integer
    intPos = InStr(strUserName, "@")
    If intPos > 0 Then
        strServerName = Mid(strUserName, intPos + 1)
        strUserName = Mid(strUserName, 1, intPos - 1)
    End If
    
    intPos = InStr(strUserName, "/")
    If intPos > 0 Then
        strPassword = Mid(strUserName, intPos + 1)
        strUserName = Mid(strUserName, 1, intPos - 1)
    End If
    
    intPos = InStr(strPassword, "@")
    If intPos > 0 Then
        strServerName = Mid(strPassword, intPos + 1)
        strPassword = Mid(strPassword, 1, intPos - 1)
    End If
    
    
    If Len(Trim(strPassword)) = 0 Then
        strNote = "请输入密码"
        GoTo InputError
    End If
    
    If Not OraDataOpen(strServerName, strUserName, IIf(UCase(strUserName) = "SYS" Or UCase(strUserName) = "SYSTEM", strPassword, IIf(mbln转换, TranPasswd(strPassword), strPassword))) Then
        TXT密码.Text = ""
        If TXT密码.Enabled Then TXT密码.SetFocus
        SetConState
        Exit Sub
    End If
    
    If mblnChangePass = False And strUserName = strPassword Then
        MsgBox "登录用户名和密码相同，不符合系统安全要求，请您立即修改密码。", vbInformation
        cmd修改密码_Click
        If mblnChangePass = False Then
            Unload FrmChangePass
            CMD放弃_Click
        Else
            cmb数据库.Enabled = False
            SetConState
            Call CMD确认_Click
        End If
        Exit Sub
    End If
    
    If UCase(strServerName) = "RBO" Then
        SetRunWithRBO
    End If
    
    '启动SQL Trace
    '-----------------------------------------------
    strNote = SetSQLTrace(strServerName)
    If strNote <> "" Then
        MsgBox "已启动SQL Trace功能!" & vbCrLf & "跟踪结果文件:" & strNote & vbCrLf & _
                "存放在Oracle服务器udump目录下,超过10M后将停止写入.", vbInformation, "提示"
    End If
    
    '-----------------------------------------------
    '更改密码处理
    '
    If Not TXT密码.Enabled Then
        If Trim(FrmChangePass.TXT原密码.Text) <> Trim(FrmChangePass.TXT新密码.Text) Then
            
            '保存新密码
            If UpdatePassword(strUserName, TranPasswd(Trim(FrmChangePass.TXT新密码.Text))) Then
                MsgBox "密码修改成功", vbInformation + vbOKOnly, "提示"
            Else
                SetConState
                Exit Sub
            End If
        End If
    End If
    
    '修改注册表
    SaveSetting "ZLSOFT", "注册信息\登陆信息", "USER", strUserName
    SaveSetting "ZLSOFT", "注册信息\登陆信息", "SERVER", strServerName
    
    '创建快捷方式用
    SaveSetting "ZLSOFT", "公共全局", "程序路径", App.Path & "\" & App.EXEName & ".exe"
    
    If mblnShowChangePassFrm Then Unload FrmChangePass
    Unload Me
    Exit Sub
InputError:
    If mintTimes > 3 Then
        MsgBox "超过三次登录失败，系统将自动退出", vbExclamation, gstrSysName
        CMD放弃_Click
    Else
        If strNote <> "" Then
            MsgBox strNote, vbExclamation, gstrSysName
        End If
        SetConState
        Exit Sub
    End If

End Sub

Private Sub SetRunWithRBO()
'功能：当前会话以RBO优化器模式运行SQL语句
    Dim strSQL As String
    
    strSQL = "alter session set optimizer_mode=rule"
    On Error Resume Next
    gcnOracle.Execute strSQL
    If Err.Number = 0 Then
        MsgBox "已设置当前会话以RBO优化器模式运行！", vbInformation, gstrSysName
    End If
End Sub

Private Function SetSQLTrace(ByVal strServerName As String) As String
'功能:调用100046事件启动SQL Trace功能
'返回:Trc文件名
    Dim strSQL As String, strLevel As String, strFile As String
    
    strServerName = UCase(strServerName)
    
    If strServerName Like "SQLTRACE*" Then
        On Error Resume Next
        strSQL = "alter session set timed_statistics=true"
        gcnOracle.Execute strSQL
        strSQL = "alter session set max_dump_file_size=10M"
        gcnOracle.Execute strSQL
        Err.Clear
        
        '下面这一条语句在8.1.7及以后才支持
        strFile = "ZL_" & gstrDeptName & "_" & gstrUserName
        strSQL = "alter session set tracefile_identifier='" & strFile & "'"
        gcnOracle.Execute strSQL
        If Err.Number <> 0 Then strFile = "*.trc": Err.Clear
        
        strLevel = "1"
        If Replace(strServerName, "SQLTRACE", "") = "4" Then
            strLevel = "4"
        ElseIf Replace(strServerName, "SQLTRACE", "") = "8" Then
            strLevel = "8"
        ElseIf Replace(strServerName, "SQLTRACE", "") = "12" Then
            strLevel = "12"
        End If
        strSQL = "alter session set events '10046 trace name context forever ,level " & strLevel & "'"
        gcnOracle.Execute strSQL
        If Err.Number = 0 Then SetSQLTrace = strFile
    End If
End Function

Private Sub cmb数据库_Change()
    Call ClearComponent
End Sub

Private Sub cmb数据库_Click()
    Call ClearComponent
End Sub

Private Sub CMD放弃_Click()
    Set gcnOracle = Nothing
    Unload Me
End Sub

Private Sub cmd修改密码_Click()
    mblnShowChangePassFrm = True
    With FrmChangePass
        .Show 1, Me
        If mblnChangePass Then
            txt用户.Enabled = False
            TXT密码.Enabled = False
            cmb数据库.Enabled = False
            If CMD确认.Enabled Then CMD确认.SetFocus
        Else
            TXT密码.Enabled = True
            TXT密码.SetFocus
        End If
    End With
End Sub

Private Sub Form_Activate()
    Dim LngStyle As Long
    If mblnFirst = False Then
        LngStyle = GetWindowLong(Me.hwnd, GWL_EXSTYLE)
        LngStyle = LngStyle Or WinStyle
        Call SetWindowLong(Me.hwnd, GWL_EXSTYLE, LngStyle)
        If InStr(Command(), "=") <= 0 Then
            ShowWindow Me.hwnd, 0 '先隐藏
            ShowWindow Me.hwnd, 1 '再显示
        End If
'
'        Call SetWindowPos(Me.hwnd, HWND_TOPMOST, Me.Left / 15, Me.Top / 15, Me.Height / 15, Me.Width / 15, SWP_NOSIZE + SWP_SHOWWINDOW)

        If Trim(txt用户.Text) = "" Then
            CMD确认.Default = False
            txt用户.SetFocus
        Else
            If TXT密码.Enabled Then
                TXT密码.SetFocus
            Else
                CMD确认.SetFocus
            End If
        End If
        mblnFirst = True
    
        If Trim(txt用户.Text) <> "" And Trim(TXT密码.Text) <> "" Then Call CMD确认_Click
    End If
    If InStr(Command(), "=") > 0 Then Me.Hide
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If Me.ActiveControl.Name = "TXT密码" Then
            Call CMD确认_Click
        Else
            SendKeys "{Tab}"
        End If
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Dim ArrCommand
    Dim i As Integer
    Call LoadServer
    On Error GoTo errH
    mbln转换 = True
    mblnFirst = False
    mintTimes = 1
    txt用户.Text = GetSetting(appName:="ZLSOFT", Section:="注册信息\登陆信息", Key:="USER", Default:="")
    cmb数据库.Text = GetSetting(appName:="ZLSOFT", Section:="注册信息\登陆信息", Key:="SERVER", Default:="")
    Call ApplyOEM_Picture(Me, "Icon")
    mblnChangePass = False
    mblnShowChangePassFrm = False
    
    If InStr(Command(), "=") > 0 Then Me.Hide
    '如果命令行参数中有用户名及密码，则填充并执行
    If Command() <> "" Then
        
        ArrCommand = Split(Command(), " ")
        
        If UBound(ArrCommand) >= 1 Then
            If InStr(ArrCommand(0), "=") <= 0 Then
                Me.txt用户.Text = ArrCommand(0)
                Me.TXT密码.Text = ArrCommand(1)
            End If
        ElseIf UBound(ArrCommand) = 0 Then
            '如果含有/，表示同时输入了用户名与密码，而且密码不需要进行转换
            If InStr(1, ArrCommand(0), "/") <> 0 Then
                Me.txt用户.Text = Split(ArrCommand(0), "/")(0)
                Me.TXT密码.Text = Split(ArrCommand(0), "/")(1)
                mbln转换 = False
            End If
        End If
    End If
    Exit Sub
errH:
    If CStr(Command()) <> "" Then MsgBox CStr(Erl()) & "行出现错误，请手动登录！" & vbNewLine & Err.Description, vbQuestion
End Sub

Private Sub GetFocus(ByVal TxtBox As TextBox)
    With TxtBox
        .SelStart = 0
        .SelLength = LenB(StrConv(.Text, vbFromUnicode))
    End With
End Sub

Private Sub cmb数据库_KeyPress(KeyAscii As Integer)
    If KeyAscii <> vbKeyReturn Then
        '回车键另行处理
        If KeyAscii <> vbKeyBack Then
            Call AppendText(KeyAscii)
        End If
    End If
End Sub

Private Sub txt用户_Change()
    If Not mblnFirst Then Exit Sub
    CMD确认.Default = False
End Sub

Private Sub txt用户_GotFocus()
    GetFocus txt用户
End Sub

Private Sub TXT密码_GotFocus()
    GetFocus TXT密码
End Sub

Private Sub cmb数据库_GotFocus()
    With cmb数据库
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub SetConState(Optional ByVal BlnState As Boolean = True)
    CMD放弃.Enabled = BlnState
    cmd修改密码.Enabled = BlnState
    CMD确认.Enabled = BlnState
End Sub

Private Sub LoadServer()
'功能：读出本地的服务器列表
    Dim strPath As String, strFile As String, lngFile As Integer
    Dim strLine As String, lngPos As Long
    Dim strServer As String, strComputer As String, strSID As String
    
    cmb数据库.Clear
    
    strPath = GetKeyValue(HKEY_LOCAL_MACHINE, "SOFTWARE\ORACLE", "ORACLE_HOME")
    If Not gobjFile.FolderExists(strPath) Then '10G
        strPath = GetKeyValue(HKEY_LOCAL_MACHINE, "SOFTWARE\ORACLE", "ORA_CRS_HOME")
    End If
    If Not gobjFile.FolderExists(strPath) Then '10Gr2
        strPath = GetKeyValue(HKEY_LOCAL_MACHINE, "SOFTWARE\ORACLE\KEY_OraDb10g_home1", "ORACLE_HOME")
    End If
    If Not gobjFile.FolderExists(strPath) Then '10Gr2
        strPath = GetKeyValue(HKEY_LOCAL_MACHINE, "SOFTWARE\ORACLE\KEY_OraDb10g_home2", "ORACLE_HOME")
    End If
    If Not gobjFile.FolderExists(strPath) Then    '10G 企业版
        strPath = GetKeyValue(HKEY_LOCAL_MACHINE, "SOFTWARE\ORACLE\KEY_OraClient10g_home1", "ORACLE_HOME")
    End If
    
    strFile = strPath & "\network\ADMIN\tnsnames.ora" 'Oracle 8i以上
    If Not gobjFile.FileExists(strFile) Then
        strFile = strPath & "\NET80\ADMIN\tnsnames.ora" 'Oracle 8
        If Not gobjFile.FileExists(strFile) Then Exit Sub
    End If
    
    lngFile = FreeFile()
    Open strFile For Input Access Read As lngFile
    
    Set mcolServer = Nothing
    Do Until EOF(lngFile)
        Input #lngFile, strLine
        
        strLine = Trim(strLine)
        If strLine <> "" And Left(strLine, 1) <> "#" Then
            '非注释行或空行
            If InStr(strLine, "(") = 0 And InStr(strLine, ")") = 0 Then
                '该行的内容就是服务器名了，把所有内容都初始化
                strServer = Trim(Mid(strLine, 1, InStr(strLine, "=") - 1))
                strComputer = ""
                strSID = ""
            ElseIf InStr(strLine, "(ADDRESS") > 0 Then
                '该行的内容是主机名
                If InStr(strLine, "PROTOCOL = TCP") > 0 And InStr(strLine, "PORT = 1521") > 0 Then
                    '符合我们的程序要求
                    strComputer = Mid(strLine, InStr(strLine, "HOST =") + Len("HOST ="))
                    strComputer = Trim(Mid(strComputer, 1, InStr(strComputer, ")") - 1))
                End If
            Else
                lngPos = InStr(strLine, "(SID")
                If lngPos = 0 Then
                    lngPos = InStr(strLine, "(SERVICE_NAME")
                End If
                
                If lngPos > 0 Then
                    '该行的内容是实例名
                    strSID = Mid(strLine, InStr(lngPos, strLine, "=") + 1)
                    strSID = Trim(Mid(strSID, 1, InStr(strSID, ")") - 1))
                    
                    If strServer <> "" And strComputer <> "" And strSID <> "" Then
                        '已经得到所有需要的内容
                        mcolServer.Add Array(strServer, strComputer, strSID)
                        cmb数据库.AddItem strServer
                    End If
                End If
            End If
        End If
    Loop
End Sub

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
    
    With cmb数据库
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
        cmb数据库.Text = strTemp
        cmb数据库.SelStart = Len(strInput)
        cmb数据库.SelLength = 100
    Else
        cmb数据库.Text = strInput
        cmb数据库.SelStart = lngStart
    End If

End Sub

Private Sub ClearComponent()
'功能：--清空注册表[本机部件]--因为不同的数据库可能使用的系统和版本不同
    If mblnFirst = True Then '启动时对控件的赋值不考虑在内
        SaveSetting "ZLSOFT", "注册信息", "本机部件", ""
    End If
End Sub

Public Sub Docmd(ByVal strCmd As String)
    Dim ArrCommand
    Dim i As Integer
    ArrCommand = Split(strCmd, " ")
    If InStr(ArrCommand(0), "=") > 0 Then
        '第三方部件调用导航台登录的格式
        For i = LBound(ArrCommand) To UBound(ArrCommand)
            If UCase(ArrCommand(i)) Like "USER=*" Then
                Me.txt用户.Text = Split(ArrCommand(i), "=")(1)
            ElseIf UCase(ArrCommand(i)) Like "PASS=*" Then
                Me.TXT密码.Text = Split(ArrCommand(i), "=")(1)
            ElseIf UCase(ArrCommand(i)) Like "SERVER=*" Then
                Me.cmb数据库.Text = Split(ArrCommand(i), "=")(1)
            ElseIf UCase(ArrCommand(i)) Like "ONLYONE=*" Then
                If Split(ArrCommand(i), "=")(1) = "1" Then
                    If App.PrevInstance = True Then
                        MsgBox "不能重复运行这个程序！"
                        End
                    End If
                End If
            End If
        Next
        If Trim(txt用户.Text) <> "" And Trim(TXT密码.Text) <> "" Then Call CMD确认_Click
    End If
End Sub

