VERSION 5.00
Begin VB.Form frmUserLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "用户登录"
   ClientHeight    =   2175
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4140
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmUserLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2175
   ScaleWidth      =   4140
   StartUpPosition =   2  '屏幕中心
   Begin VB.TextBox txtPwd 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   2040
      PasswordChar    =   "*"
      TabIndex        =   3
      Text            =   "aqa"
      Top             =   600
      Width           =   1920
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   -120
      TabIndex        =   7
      Top             =   1440
      Width           =   4335
      Begin VB.CommandButton cmdCancel 
         Caption         =   "取 消(&C)"
         Height          =   350
         Left            =   2400
         TabIndex        =   6
         Top             =   240
         Width           =   1100
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "确 定(&O)"
         Height          =   350
         Left            =   840
         TabIndex        =   5
         Top             =   240
         Width           =   1100
      End
   End
   Begin VB.ComboBox cboServerName 
      Height          =   315
      Left            =   2040
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   1080
      Width           =   1920
   End
   Begin VB.TextBox txtUserName 
      Enabled         =   0   'False
      Height          =   300
      Left            =   2040
      TabIndex        =   2
      Text            =   "ZLHIS"
      Top             =   120
      Width           =   1920
   End
   Begin VB.Label lblPwd 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "密码"
      Height          =   195
      Left            =   1320
      TabIndex        =   8
      Top             =   653
      Width           =   420
   End
   Begin VB.Label lblServerName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "服务器"
      Height          =   195
      Left            =   1320
      TabIndex        =   1
      Top             =   1140
      Width           =   540
   End
   Begin VB.Label lblUserName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "用户名"
      Height          =   195
      Left            =   1320
      TabIndex        =   0
      Top             =   173
      Width           =   540
   End
   Begin VB.Image imgFlag 
      Height          =   720
      Left            =   315
      Picture         =   "frmUserLogin.frx":058A
      Top             =   240
      Width           =   720
   End
End
Attribute VB_Name = "frmUserLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mblnFirst As Boolean  '为True表示已经正常显示出
Private mintTimes As Integer  '登录重试次数
Private mbln转换 As Boolean     '表示传入的密码是否为数据库密码，是否不需要再转换
Private mcolServer As New Collection  '保存服务器串列表


Private Sub cboServerName_Change()
On Error GoTo errHandle

    Call ClearComponent
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub cboServerName_Click()
On Error GoTo errHandle

    Call ClearComponent

Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub cmdCancel_Click()
On Error GoTo errHandle

 Set gcnOracle = Nothing
    Unload Me

Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub cmdOK_Click()
On Error GoTo errHandle

    Dim strNote As String
    Dim strUserName As String
    Dim strServerName As String
    Dim strPassword As String
    On Error GoTo InputError
    
    '------检验用户是否oracle合法用户----------------
    strUserName = Trim(txtUserName.Text)
    strPassword = Trim(txtPwd.Text)

    strServerName = Trim(cboServerName.Text)
    
    '有效字符串效验
    If Len(Trim(txtUserName)) = 0 Then
        strNote = "请输入用户名"
        txtUserName.SetFocus
        GoTo InputError
    End If
    
    If Len(strUserName) <> 1 Then
        If Mid(strUserName, 1, 1) = "/" Or Mid(strUserName, 1, 1) = "@" Or Mid(strUserName, Len(strUserName) - 1, 1) = "/" Or Mid(strUserName, Len(strUserName) - 1, 1) = "@" Then
            txtUserName.SetFocus
            strNote = "用户名错误"
            Exit Sub
        End If
    End If
    If Trim(strPassword) <> "" And Len(strPassword) <> 1 Then
        If Mid(strPassword, Len(strPassword) - 1, 1) = "/" Or Mid(strPassword, Len(strPassword) - 1, 1) = "@" Or Mid(strPassword, 1, 1) = "/" Or Mid(strPassword, 1, 1) = "@" Then
            If txtPwd.Enabled Then txtPwd.SetFocus
            strNote = "密码错误"
            GoTo InputError
        End If
    End If
    If Trim(strServerName) <> "" Then
        If Mid(strServerName, Len(strServerName) - 1, 1) = "/" Or Mid(strServerName, Len(strServerName) - 1, 1) = "@" Or Mid(strServerName, 1, 1) = "/" Or Mid(strServerName, 1, 1) = "@" Then
            strNote = "主机连接串错误"
            cboServerName.SetFocus
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
        txtPwd.Text = ""
        If txtPwd.Enabled Then txtPwd.SetFocus

        Exit Sub
    End If
      
    
    '修改注册表
    SaveSetting "ZLSOFT", "注册信息\登陆信息", "SERVER", strServerName
    
    '创建快捷方式用
    SaveSetting "ZLSOFT", "公共全局", "程序路径", App.Path & "\" & App.EXEName & ".exe"
    
    Unload Me
    
    Call frmMain.Show

InputError:
    If mintTimes > 3 Then
        MsgBox "超过三次登录失败，系统将自动退出", vbExclamation, gstrSysName
        cmdCancel_Click
    Else
        If strNote <> "" Then
            MsgBox strNote, vbExclamation, gstrSysName
        End If
        Exit Sub
    End If

Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Dim ArrCommand
    Dim i As Integer
    Call LoadServer
    mbln转换 = True
    On Error GoTo errH
    mblnFirst = False
    mintTimes = 1
    cboServerName.Text = GetSetting(appName:="ZLSOFT", Section:="注册信息\登陆信息", Key:="SERVER", Default:="")
    
    If InStr(Command(), "=") > 0 Then Me.Hide
    '如果命令行参数中有用户名及密码，则填充并执行
    If Command() <> "" Then
        
        ArrCommand = Split(Command(), " ")
        
        If UBound(ArrCommand) >= 1 Then
            If InStr(ArrCommand(0), "=") <= 0 Then
                Me.txtUserName.Text = ArrCommand(0)
                Me.txtPwd.Text = ArrCommand(1)
            End If
        ElseIf UBound(ArrCommand) = 0 Then
            '如果含有/，表示同时输入了用户名与密码，而且密码不需要进行转换
            If InStr(1, ArrCommand(0), "/") <> 0 Then
                Me.txtUserName.Text = Split(ArrCommand(0), "/")(0)
                Me.txtPwd.Text = Split(ArrCommand(0), "/")(1)
                mbln转换 = False
            End If
        End If
    End If
    Exit Sub
errH:
    If CStr(Command()) <> "" Then MsgBox CStr(Erl()) & "行出现错误，请手动登录！" & vbNewLine & Err.Description, vbQuestion
End Sub


Private Sub LoadServer()
'功能：读出本地的服务器列表
    Dim strPath As String, strFile As String, lngFile As Integer
    Dim strLine As String, lngPos As Long
    Dim strServer As String, strComputer As String, strSID As String
    
    cboServerName.Clear
    
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
                        cboServerName.AddItem strServer
                    End If
                End If
            End If
        End If
    Loop
End Sub

Private Sub ClearComponent()
'功能：--清空注册表[本机部件]--因为不同的数据库可能使用的系统和版本不同
    If mblnFirst = True Then '启动时对控件的赋值不考虑在内
        SaveSetting "ZLSOFT", "注册信息", "本机部件", ""
    End If
End Sub
