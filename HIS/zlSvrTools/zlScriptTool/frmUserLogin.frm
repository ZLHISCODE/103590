VERSION 5.00
Begin VB.Form frmUserLogin 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "操作员登录"
   ClientHeight    =   2250
   ClientLeft      =   30
   ClientTop       =   330
   ClientWidth     =   4065
   Icon            =   "frmUserLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2250
   ScaleWidth      =   4065
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton cmdSelect 
      Caption         =   "…"
      Height          =   300
      Left            =   3360
      TabIndex        =   6
      TabStop         =   0   'False
      ToolTipText     =   "选择存在的服务器列表"
      Top             =   960
      Width           =   255
   End
   Begin VB.TextBox txt数据库 
      Height          =   300
      Left            =   2040
      MaxLength       =   20
      TabIndex        =   5
      Top             =   960
      Width           =   1335
   End
   Begin VB.TextBox txt密码 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   2055
      MaxLength       =   20
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   600
      Width           =   1515
   End
   Begin VB.TextBox txt用户 
      Height          =   300
      Left            =   2055
      MaxLength       =   30
      TabIndex        =   1
      Top             =   195
      Width           =   1515
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   2520
      TabIndex        =   8
      Top             =   1680
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Default         =   -1  'True
      Height          =   350
      Left            =   600
      TabIndex        =   7
      Top             =   1680
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Height          =   50
      Left            =   0
      TabIndex        =   9
      Top             =   1440
      Width           =   4725
   End
   Begin VB.Image imgFlag 
      Height          =   720
      Left            =   240
      Picture         =   "frmUserLogin.frx":1CFA
      Top             =   360
      Width           =   720
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "口  令"
      Height          =   180
      Left            =   1440
      TabIndex        =   2
      Top             =   660
      Width           =   540
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "用户名"
      Height          =   180
      Left            =   1440
      TabIndex        =   0
      Top             =   255
      Width           =   540
   End
   Begin VB.Label lblDataBase 
      AutoSize        =   -1  'True
      Caption         =   "服务器"
      Height          =   180
      Left            =   1440
      TabIndex        =   4
      Top             =   1065
      Width           =   540
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
Dim mcolServer As New Collection


Private Sub cmdOK_Click()

    intTimes = intTimes + 1
    
    '------检验用户是否oracle合法用户----------------
    gstrUserName = Trim(txt用户.Text)
    gstrPassword = Trim(txt密码.Text)
    gstrServer = Trim(txt数据库.Text)
    
    '有效字符串效验
    If Len(Trim(txt用户)) = 0 Then
        strNote = "请输入用户名。"
        txt用户.SetFocus
        GoTo InputError
    End If
    
    If Len(gstrUserName) <> 1 Then
        If Mid(gstrUserName, 1, 1) = "/" Or Mid(gstrUserName, 1, 1) = "@" Or Mid(gstrUserName, Len(gstrUserName) - 1, 1) = "/" Or Mid(gstrUserName, Len(gstrUserName) - 1, 1) = "@" Then
            txt用户.SetFocus
            strNote = "用户名错误。"
            Exit Sub
        End If
    End If
    If Trim(gstrPassword) <> "" And Len(gstrPassword) <> 1 Then
        If Mid(gstrPassword, Len(gstrPassword) - 1, 1) = "/" Or Mid(gstrPassword, Len(gstrPassword) - 1, 1) = "@" Or Mid(gstrPassword, 1, 1) = "/" Or Mid(gstrPassword, 1, 1) = "@" Then
            txt密码.SetFocus
            strNote = "口令错误。"
            GoTo InputError
        End If
    End If
    If Trim(gstrServer) <> "" Then
        If Mid(gstrServer, Len(gstrServer) - 1, 1) = "/" Or Mid(gstrServer, Len(gstrServer) - 1, 1) = "@" Or Mid(gstrServer, 1, 1) = "/" Or Mid(gstrServer, 1, 1) = "@" Then
            strNote = "主机连接串错误。"
            txt数据库.SetFocus
            GoTo InputError
        End If
    End If
    
    '分离字符串
    Dim intPos As Integer
    intPos = InStr(1, gstrUserName, "@", vbTextCompare)
    If intPos > 0 Then
        gstrServer = Mid(gstrUserName, intPos + 1)
        gstrUserName = Mid(gstrUserName, 1, intPos - 1)
    End If
    
    intPos = InStr(1, gstrUserName, "/", vbTextCompare)
    If intPos > 0 Then
        gstrPassword = Mid(gstrUserName, intPos + 1)
        gstrUserName = Mid(gstrUserName, 1, intPos - 1)
    End If
    
    intPos = InStr(1, gstrPassword, "@", vbTextCompare)
    If intPos > 0 Then
        gstrServer = Mid(gstrPassword, intPos + 1)
        gstrPassword = Mid(gstrPassword, 1, intPos - 1)
    End If
    
    If Len(Trim(gstrPassword)) = 0 Then
        strNote = "未输入密码，不能注册。"
        txt密码.SetFocus
        GoTo InputError
    End If
    
    gstrUserName = UCase(gstrUserName)
    If gstrUserName <> "SYSTEM" And gstrUserName <> "SYS" Then
        gstrPassword = TranPasswd(gstrPassword)
  
    End If
    
    If Not OraDataOpen(gstrServer, gstrUserName, gstrPassword, gcnOracle) Then
        txt密码.Text = ""
        Exit Sub
    End If
    
    '初始化公共部件
''    Call InitCommon(gcnOracle)
    
    Set gmobjCommon = CreateObject("ZL9ComLib.clsComLib")
'    Set gmobjCommon = New zl9ComLib.clsComLib
    Call gmobjCommon.InitCommon(gcnOracle)
  
'    If Not gmobjCommon.RegCheck Then
'        MsgBox "ZL9ComLib注册失败!", vbExclamation, "提示!"
'        Exit Sub
'    End If
    
    Call SaveSetting("ZLSOFT", "注册信息\登陆信息", "主服务器名", txt数据库.Text)
    
    Unload Me
    
    '显示主窗体
'    Me.Hide
    frmMain.Show
    
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

Private Sub cmdCancel_Click()
    On Error Resume Next
    Set gcnOracle = Nothing
    Unload Me
End Sub

Private Sub cmdSelect_Click()
    Dim strServer As String
    Dim p As POINTAPI
    
    p.X = txt数据库.Left / Screen.TwipsPerPixelX
    p.Y = (cmdSelect.Top + cmdSelect.Height) / Screen.TwipsPerPixelY
    ClientToScreen Me.Hwnd, p
    
    strServer = frmServerSelect.GetServer(mcolServer, p.X * Screen.TwipsPerPixelX, p.Y * Screen.TwipsPerPixelY, txt数据库.Text)
    If strServer <> "" Then
        txt数据库.Text = strServer
        txt数据库.SetFocus
    End If
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    Dim lngStyle As Long
    lngStyle = GetWindowLong(Hwnd, GWL_EXSTYLE)
    lngStyle = lngStyle Or WinStyle
    Call SetWindowLong(Hwnd, GWL_EXSTYLE, lngStyle)
    
    ShowWindow Me.Hwnd, 0 '先隐藏
    ShowWindow Me.Hwnd, 1 '再显示
    
    If Len(txt用户) <> 0 Then
        txt密码.SetFocus
    End If
End Sub

Private Sub Form_Load()
    txt用户.Text = "zlhis" ''GetSetting("ZLSOFT", "注册信息\登陆信息", "MANAGER", "")
    txt数据库.Text = GetSetting("ZLSOFT", "注册信息\登陆信息", "主服务器名", "")
    intTimes = 0
    
    Call LoadServer
   ' Call ApplyOEM_Picture(Me, "Icon")
    
 
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
    Me.txt用户.SelStart = 0: Me.txt用户.SelLength = 100
End Sub

Private Sub txt密码_GotFocus()
    Me.txt密码.SelStart = 0: Me.txt密码.SelLength = 100
End Sub

Private Sub txt数据库_GotFocus()
    Me.txt数据库.SelStart = 0: Me.txt数据库.SelLength = 100
End Sub

Public Sub LoadServer()
'功能：读出本地的服务器列表
    Dim objSys As New Scripting.FileSystemObject
    Dim txtStream As Scripting.TextStream
    Dim strPath As String, strFile As String
    Dim strLine As String, lngPos As Long
    Dim strServer As String, strComputer As String, strSID As String
    
    strPath = GetKeyValue(HKEY_LOCAL_MACHINE, "SOFTWARE\ORACLE", "ORACLE_HOME")
    
    '读取10gOracleHome目录
    If strPath = "" Then
       strPath = GetKeyValue(HKEY_LOCAL_MACHINE, "SOFTWARE\ORACLE\KEY_OraDb10g_home1", "ORACLE_HOME")
    End If
    
    '首先试验Oracle 8i的配置文件在否
    strFile = strPath & "\network\ADMIN\tnsnames.ora"
    If objFso.FileExists(strFile) = False Then
        '再试验Oracle 8的配置文件在否
        strFile = strPath & "\NET80\ADMIN\tnsnames.ora"
        If objSys.FileExists(strFile) = False Then
            Exit Sub
        End If
    End If
    
    On Error Resume Next
    
    Set mcolServer = Nothing
    Set txtStream = objSys.OpenTextFile(strFile)
    Do Until txtStream.AtEndOfStream
        strLine = Trim(txtStream.ReadLine)
        
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

