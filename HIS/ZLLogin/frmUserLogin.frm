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
   Begin VB.ComboBox cboServer 
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
   Begin VB.CommandButton cmdModify 
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
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   2865
      TabIndex        =   7
      Top             =   1710
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Default         =   -1  'True
      Height          =   350
      Left            =   1755
      TabIndex        =   6
      Top             =   1710
      Width           =   1100
   End
   Begin VB.TextBox txtPassWord 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   1950
      MaxLength       =   20
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   630
      Width           =   1920
   End
   Begin VB.TextBox txtUser 
      Height          =   300
      Left            =   1950
      TabIndex        =   1
      Top             =   195
      Width           =   1920
   End
   Begin VB.CommandButton cmdSet 
      Caption         =   "配置服务器"
      Height          =   350
      Left            =   180
      TabIndex        =   10
      ToolTipText     =   "启动Oracle主机字符串配置程序"
      Top             =   1710
      Width           =   1335
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

'zlhis.exe 菜单
'zlhis.exe 用户名/密码        此种情况不需要进行密码转换
'zlhis.exe 用户名 密码
'zlhis.exe 用户名 密码 菜单
Private mblnFirst As Boolean  '为True表示已经正常显示出
Private mintTimes As Integer  '登录重试次数
Private mbln转换 As Boolean     '表示传入的密码是否为数据库密码，是否不需要再转换
Private mcolServer As New Collection  '保存服务器串列表
Private mblnAccess As Boolean  '为True外部调用ZLHIS成功
Private mblnUAAddUser As Boolean

Private mobjHttp As New XMLHTTP
Private mstrPostData As String
Private mstr断言 As String
Private mstrUserURL As String
Private mstrSamlAssertion As String
Private mstrError As String
Private mblnZLUA As Boolean
Private mstrAppID As String
Private mstrZLUAUser As String
Private mblnOK          As Boolean
Private Declare Function SetActiveWindow Lib "user32" (ByVal hwnd As Long) As Long

Private Sub cmdOK_Click()
    Dim strNote             As String
    Dim strUserName         As String
    Dim strServerName       As String
    Dim strPassword         As String
    Dim blnTransPassword    As Boolean
    Dim strError            As String
    Dim strSQL              As String
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo errH
    SetConState False
    If Not CheckInput(strUserName, strPassword, strServerName) Then
        SetConState
        Exit Sub
    End If
    mintTimes = mintTimes + 1
    
    If UCase(strUserName) = "SYS" Or UCase(strUserName) = "SYSTEM" Then
        blnTransPassword = False
    Else
        blnTransPassword = mbln转换
    End If
    
    Set gcnOracle = gobjRegister.GetConnection(strServerName, strUserName, strPassword, blnTransPassword, , strError)
    'ora-28002:密码还有多少天过期，不会返回，因此，必须CheckPwdExpiry来辅助提示密码过期
    If gcnOracle.State = adStateClosed Then
        If InStr(strError, "ORA-00604") > 0 Or InStr(strError, "ORA-04088") > 0 Then
            If InStr(strError, "ORA-20002") > 0 Then
                strError = "当前用户不能使用该应用登录数据库，请联系管理员。"
            Else
                strError = "当前用户被禁止登录数据库，请联系管理员。"
            End If
        End If
        If InStr(strError, "ORA-28001") > 0 Then
            strError = "密码已经过期。请联系管理员重置密码！"
        End If
        MsgBox strError, vbInformation, gstrSysName
        txtPassWord.Text = ""
        mblnAccess = False
        If mblnZLUA = True Then mblnUAAddUser = True
        txtPassWord.SetFocus
        SetConState
        Exit Sub
    Else
        gclsLogin.DBUser = UCase(strUserName)
        If strUserName = strPassword Then
            MsgBox "登录用户名和密码相同，不符合系统安全要求，请您立即修改密码。", vbInformation, gstrSysName
            If gintCallType = 0 Then '现实修改按钮
                cmdModify_Click
                SetConState
            End If
            Exit Sub
        End If
        '检查密码复杂度是否符合要求
        If Not CheckPWDComplex(gcnOracle, strPassword) Then
            If gintCallType = 0 Then '现实修改按钮
                cmdModify_Click
                SetConState
            End If
            Exit Sub
        End If
        
        '是否密码过期提醒
        If CheckPwdExpiry = True Then
            If gintCallType = 0 Then '现实修改按钮
                cmdModify_Click
                SetConState
            End If
            Exit Sub
        End If
    End If
    
    strSQL = "Select 1 From 上机人员表 a, 人员表 b Where a.人员id = b.Id And b.撤档时间 < Sysdate And a.用户名 = [1]"
    Set rsTemp = OpenSQLRecord(strSQL, "帐号到期检查", UCase(strUserName))
    If rsTemp.RecordCount > 0 Then
        MsgBox "该账户对应的人员已撤档，登录失败！"
        txtPassWord.Text = ""
        SetConState
        Exit Sub
    End If
    '启动SQL Trace
    '-----------------------------------------------
    strNote = SetSQLTrace(strServerName)
    If strNote <> "" Then
        MsgBox "已启动SQL Trace功能!" & vbCrLf & "跟踪结果文件:" & strNote & vbCrLf & _
                "存放在Oracle服务器udump目录下,超过100M后将停止写入.", vbInformation, "提示"
    End If
    If UCase(strServerName) = "RBO" Then
        SetRunWithRBO
    End If
    '接口调用，放到Trace启动之后
    '-----------------------------------------------
    '1.中联单点登录添加ZLUA账户
    If mblnUAAddUser = True And mstrUserURL <> "" Then
        mstr断言 = SoapEnvelope("AddUserAppInfo", mstrZLUAUser, mstrAppID, txtUser.Text & "/" & txtPassWord.Text & "@" & cboServer.Text, mstrSamlAssertion)
        Call PostData(mstrUserURL, "AddUserAppInfo", mstr断言, 5)
        mblnUAAddUser = False
    End If
    
    '2.新版病历、自动升级程序、导航台，需要的用户名及密码(用户输入的密码，zlbrw部件中会使用)
    gclsLogin.InputUser = strUserName
    gclsLogin.InputPwd = strPassword
    gclsLogin.ServerName = strServerName
    gclsLogin.IsTransPwd = blnTransPassword
    '修改注册表
    SaveSetting "ZLSOFT", "注册信息\登陆信息", "USER", strUserName
    SaveSetting "ZLSOFT", "注册信息\登陆信息", "SERVER", strServerName
    
    mblnAccess = True
    mblnOK = True
    Unload Me
    Exit Sub
errH:
    If mintTimes > 3 Then
        MsgBox "超过三次登录失败，系统将自动退出", vbInformation, gstrSysName
        cmdCancel_Click
    Else
        MsgBox Err.Description, vbInformation, gstrSysName
        SetConState
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
    Dim rsTmp As ADODB.Recordset
    
    strServerName = UCase(strServerName)
    
    If strServerName Like "SQLTRACE*" Then
        On Error Resume Next
        strSQL = "alter session set timed_statistics=true"
        gcnOracle.Execute strSQL
        strSQL = "alter session set max_dump_file_size='100M'"
        gcnOracle.Execute strSQL
        Err.Clear
        
        '下面这一条语句在8.1.7及以后才支持
        strFile = "ZL_" & gclsLogin.DBUser
        strSQL = "alter session set tracefile_identifier='" & strFile & "'"
        gcnOracle.Execute strSQL
        If Err.Number <> 0 Then strFile = "*.trc": Err.Clear
        
        strLevel = "12"
        If Replace(strServerName, "SQLTRACE", "") = "4" Then
            strLevel = "4"
        ElseIf Replace(strServerName, "SQLTRACE", "") = "8" Then
            strLevel = "8"
        ElseIf Replace(strServerName, "SQLTRACE", "") = "12" Then
            strLevel = "12"
        End If
        strSQL = "alter session set events '10046 trace name context forever ,level " & strLevel & "'"
        gcnOracle.Execute strSQL
        If Err.Number = 0 Then
            SetSQLTrace = strFile
            
            strSQL = "Select 1 From zlreginfo Where 项目='TRACE文件'"
            Set rsTmp = gcnOracle.Execute(strSQL)
            
            If rsTmp.RecordCount > 0 Then
                strSQL = "Update zlreginfo Set 内容 ='TRACE文件' Where 项目='" & strFile & ".trc'"
            Else
                strSQL = "Insert Into zlreginfo (项目,内容) Values ('TRACE文件','" & strFile & ".trc')"
            End If
            gcnOracle.Execute strSQL

        End If
    End If
End Function

Private Sub cboServer_Change()
    Call ClearComponent
End Sub

Private Sub cboServer_Click()
    Call ClearComponent
End Sub

Private Sub cmdCancel_Click()
    Set gobjRegister = Nothing
    gclsLogin.IsCancel = True
    '密码不符合规则，修改密码点取消，此时gcnOracle不为nothing
    If Not gcnOracle Is Nothing Then
        If gcnOracle.State = adStateOpen Then
            gcnOracle.Close
        End If
    End If
    Unload Me
End Sub

Private Sub cmdModify_Click()
    Dim strUserName As String
    Dim strPassword As String
    Dim strServerName As String
    Dim strNote As String
    
    On Error GoTo InputError
    '------检验用户是否oracle合法用户----------------
    strUserName = Trim(txtUser.Text)
    strPassword = Trim(txtPassWord.Text)
    strServerName = Trim(cboServer.Text)
    
    '有效字符串效验
    If Len(Trim(txtUser.Text)) = 0 Then
        strNote = "请输入用户名"
        txtUser.SetFocus
        GoTo InputError
    End If
    
    If Len(strUserName) <> 1 Then
        If Mid(strUserName, 1, 1) = "/" Or Mid(strUserName, 1, 1) = "@" Or Mid(strUserName, Len(strUserName) - 1, 1) = "/" Or Mid(strUserName, Len(strUserName) - 1, 1) = "@" Then
            txtUser.SetFocus
            strNote = "用户名错误"
            SetConState
            Exit Sub
        End If
    End If
    
    If Trim(strPassword) <> "" And Len(strPassword) <> 1 Then
        If Mid(strPassword, Len(strPassword) - 1, 1) = "/" Or Mid(strPassword, Len(strPassword) - 1, 1) = "@" Or Mid(strPassword, 1, 1) = "/" Or Mid(strPassword, 1, 1) = "@" Then
            If txtPassWord.Enabled Then txtPassWord.SetFocus
            strNote = "密码错误"
            GoTo InputError
        End If
    End If
    If Trim(strServerName) <> "" Then
        If Mid(strServerName, Len(strServerName) - 1, 1) = "/" Or Mid(strServerName, Len(strServerName) - 1, 1) = "@" Or Mid(strServerName, 1, 1) = "/" Or Mid(strServerName, 1, 1) = "@" Then
            strNote = "主机连接串错误"
            cboServer.SetFocus
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
    
    If FrmChangePass.ShowMe(Me, strUserName, strPassword, strServerName, mbln转换) Then
        txtPassWord.Text = strPassword
        cboServer.Text = strServerName
        If cmdOK.Enabled Then Call cmdOK_Click
    Else
        txtPassWord.SetFocus
    End If
    Exit Sub
InputError:
    If strNote <> "" Then
        MsgBox strNote, vbInformation, gstrSysName
    Else
        MsgBox Err.Description, vbInformation, gstrSysName
    End If
End Sub

Private Sub cmdSet_Click()
    Dim strPath As String   'Oracle安装目录
    Dim strCommond As String, strError As String
    
    strPath = cmdSet.Tag
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

Private Sub Form_Activate()
    Dim LngStyle As Long
    
    If mblnFirst = False Then
        
        If InStr(gstrCommand, "=") <= 0 And InStr(gstrCommand, "&") <= 0 Then
            '设置当前窗口在任务栏显示
            LngStyle = GetWindowLong(Me.hwnd, GWL_EXSTYLE)
            LngStyle = LngStyle Or WinStyle
            Call SetWindowLong(Me.hwnd, GWL_EXSTYLE, LngStyle)
            
            ShowWindow Me.hwnd, 0 '先隐藏
            ShowWindow Me.hwnd, 1 '再显示
        
            If Trim(txtUser.Text) = "" Then
                cmdOK.Default = False
                txtUser.SetFocus
            Else
                txtPassWord.SetFocus
            End If
        End If
        
        mblnFirst = True
        If Trim(txtUser.Text) <> "" And Trim(txtPassWord.Text) <> "" Then Call cmdOK_Click
    End If
    If InStr(gstrCommand, "=") > 0 And InStr(gstrCommand, "&") = 0 Then Me.Hide
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If Me.ActiveControl.Name = "txtPassWord" Then
            Call cmdOK_Click
        Else
            SendKeys "{Tab}"
        End If
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If Chr(KeyAscii) = "'" Then
        KeyAscii = 0
    End If
End Sub

Private Sub Form_Load()
    Dim ArrCommand
    Dim i As Integer
    
    Call InitFaceType
    Call LoadServer
    
    On Error GoTo errH
    txtUser.Text = GetSetting(appName:="ZLSOFT", Section:="注册信息\登陆信息", Key:="USER", Default:="")
    cboServer.Text = GetSetting(appName:="ZLSOFT", Section:="注册信息\登陆信息", Key:="SERVER", Default:="")
    
    Call ApplyOEM_Picture(Me, "Icon")
    
    If InStr(gstrCommand, "=") > 0 And InStr(gstrCommand, "&") = 0 Then
        Me.Hide
    Else
        '不加这一句的话，由于已显示frmSplash窗体，在开启输入法的情况下，启动源程序，不会显示登录窗口，VB只能异常终止退出
        SetActiveWindow Me.hwnd
    End If
        
    '如果命令行参数中有用户名及密码，则填充并执行
    If gstrCommand <> "" And InStr(gstrCommand, "&") = 0 Then
        ArrCommand = Split(gstrCommand, " ")
        If UBound(ArrCommand) >= 1 Then
            If InStr(ArrCommand(0), "=") <= 0 Then
                Me.txtUser.Text = ArrCommand(0)
                Me.txtPassWord.Text = ArrCommand(1)
            End If
        ElseIf UBound(ArrCommand) = 0 Then
            '如果含有/，表示同时输入了用户名与密码，而且密码不需要进行转换
            If InStr(1, ArrCommand(0), "/") <> 0 And InStr(1, ArrCommand(0), ",") = 0 Then
                Me.txtUser.Text = Split(ArrCommand(0), "/")(0)
                Me.txtPassWord.Text = Split(ArrCommand(0), "/")(1)
                mbln转换 = False
            End If
        End If
    End If
    Exit Sub
errH:
    If CStr(gstrCommand) <> "" Then MsgBox CStr(Erl()) & "行出现错误，请手动登录！" & vbNewLine & Err.Description, vbQuestion
End Sub

Private Sub GetFocus(ByVal TxtBox As TextBox)
    With TxtBox
        If Trim(TxtBox.Text) = "" Then Exit Sub
        .SelStart = 0
        .SelLength = Len(TxtBox.Text)
    End With
End Sub

Private Sub cboServer_KeyPress(KeyAscii As Integer)
    If KeyAscii <> vbKeyReturn Then
        '回车键另行处理
        If KeyAscii <> vbKeyBack Then
            Call AppendText(KeyAscii)
        End If
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    '密码不符合规则，修改密码点X，此时gcnOracle不为nothing
    If Not mblnOK Then
        If Not gcnOracle Is Nothing Then
            If gcnOracle.State = adStateOpen Then
                gcnOracle.Close
            End If
        End If
    End If
    Set mobjHttp = Nothing
    Set mcolServer = Nothing
End Sub

Private Sub txtUser_Change()
    If Not mblnFirst Then Exit Sub
    cmdOK.Default = False
End Sub

Private Sub txtUser_GotFocus()
    If Me.ActiveControl Is txtUser Then
        OS.OpenIme (False)
        GetFocus txtUser
    End If
End Sub

Private Sub txtPassWord_GotFocus()
    GetFocus txtPassWord
End Sub

Private Sub cboServer_GotFocus()
    If Me.ActiveControl Is cboServer Then
        OS.OpenIme (False)
        If Trim(cboServer.Text) <> "" Then
            With cboServer
                .SelStart = 0
                .SelLength = Len(.Text)
            End With
        End If
    End If
End Sub

Private Sub SetConState(Optional ByVal BlnState As Boolean = True)
    cmdCancel.Enabled = BlnState
    cmdModify.Enabled = BlnState
    cmdOK.Enabled = BlnState
End Sub

Private Sub LoadServer()
'功能：读出本地的服务器列表
    Dim strPath As String, strFile As String, lngFile As Integer
    Dim strLine As String, lngPos As Long
    Dim strServer As String, strComputer As String, strSID As String
    Dim arrTmp As Variant
    Dim rsOraHome As ADODB.Recordset
    Dim intVersion As Integer, intTimes As Integer, intServer As Integer
    Dim i As Long, blnRead As Boolean

    cboServer.Clear
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
        arrTmp = OS.GetAllSubKey("HKEY_LOCAL_MACHINE\SOFTWARE\Oracle")
        If TypeName(arrTmp) = "Empty" Then
            If OS.Is64bit Then
                cboServer.ToolTipText = "没有找到注册表项HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\Oracle！"
            Else
                cboServer.ToolTipText = "没有找到注册表项HKEY_LOCAL_MACHINE\SOFTWARE\Oracle！"
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
            .Sort = "VerSion Desc,Times Desc,Server"
            Do While Not .EOF
                strPath = ""
                blnRead = Not OS.GetRegValue("HKEY_LOCAL_MACHINE\SOFTWARE\Oracle" & !Name, "ORACLE_HOME", strPath)
                blnRead = blnRead Or strPath = "" And !Name & "" = ""
                If blnRead Then
                    Call OS.GetRegValue("HKEY_LOCAL_MACHINE\SOFTWARE\Oracle", "ORA_CRS_HOME", strPath)
                End If
                If strPath <> "" Then
                    cmdSet.Tag = strPath '缓存OracleHome路径
                    strFile = strPath & "\network\ADMIN\tnsnames.ora" 'Oracle 8i以上
                    If Dir(strFile) <> "" Then Exit Do
                    strFile = strPath & "\NET80\ADMIN\tnsnames.ora" 'Oracle 8
                    If Dir(strFile) <> "" Then Exit Do
                End If
                strFile = ""
                .MoveNext
            Loop
        End If
    End With
    If strFile = "" Then Exit Sub
    
    cboServer.ToolTipText = "服务器列表来源:" & strFile
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
                If InStr(strLine, "PROTOCOL = TCP") > 0 And InStr(strLine, "PORT = ") > 0 Then
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
                        cboServer.AddItem strServer
                    End If
                End If
            End If
        End If
    Loop
    Close #lngFile
End Sub
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
    
    With cboServer
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
        cboServer.Text = strTemp
        cboServer.SelStart = Len(strInput)
        cboServer.SelLength = 100
    Else
        cboServer.Text = strInput
        cboServer.SelStart = lngStart
    End If

End Sub

Private Sub ClearComponent()
'功能：--清空注册表[本机部件]--因为不同的数据库可能使用的系统和版本不同
    If mblnFirst = True Then '启动时对控件的赋值不考虑在内
        SaveSetting "ZLSOFT", "注册信息", "本机部件", ""
    End If
End Sub

Private Function ReadINIToRec(ByVal strFile As String) As ADODB.Recordset
'功能：将指定INI配置文件的内容读取到记录集中
'返回：Nothing或包含"项目,内容"的记录集,其中同一项目可能有多行内容
    Dim rsTmp As New ADODB.Recordset
    Dim objINI As TextStream
    
    Dim strItem As String, strText As String
    Dim strLine As String
            
    rsTmp.Fields.Append "项目", adVarChar, 200
    rsTmp.Fields.Append "内容", adVarChar, 200
    rsTmp.CursorLocation = adUseClient
    rsTmp.LockType = adLockOptimistic
    rsTmp.CursorType = adOpenStatic
    rsTmp.Open
    
    Set objINI = gobjFile.OpenTextFile(strFile, ForReading)
    Do While Not objINI.AtEndOfStream
        strLine = Replace(objINI.ReadLine, vbTab, " ")
        strItem = Trim(Mid(strLine, InStr(strLine, "[") + 1, InStr(strLine, "]") - InStr(strLine, "[") - 1))
        strText = Trim(Mid(strLine, InStr(strLine, "]") + 1))
        If strItem <> "" And strText <> "" Then
            rsTmp.AddNew
            rsTmp!项目 = strItem
            rsTmp!内容 = strText
            rsTmp.Update
        End If
    Loop
    
    objINI.Close
    
    If rsTmp.RecordCount > 0 Then rsTmp.MoveFirst
    
    Set ReadINIToRec = rsTmp
End Function


Private Function SoapEnvelope(ByVal strMethod As String, ByVal parm1 As String, ByVal parm2 As String, ByVal parm3 As String, ByVal samlAssertion As String) As String
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim strEnvelope As String
    
    SoapEnvelope = strEnvelope

    On Error GoTo errHand
    
    strEnvelope = ""
    
    strEnvelope = strEnvelope & "<soapenv:Envelope xmlns:soapenv=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:Item=""http://tempuri.org/"">"
    
    If samlAssertion <> "" Then
        strEnvelope = strEnvelope & "<soapenv:Header>"
        strEnvelope = strEnvelope & "<wsse:Security xmlns:wsu=""http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-utility-1.0.xsd"" xmlns:wsse=""http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-secext-1.0.xsd"">"
        strEnvelope = strEnvelope & samlAssertion
        strEnvelope = strEnvelope & "</wsse:Security>"
        strEnvelope = strEnvelope & "</soapenv:Header>"
    End If
    
    strEnvelope = strEnvelope & "<soapenv:Body>"
    strEnvelope = strEnvelope & "<Item:" & strMethod & ">"
    Select Case strMethod
    Case "GetSAMLResponseByArtifact"
        strEnvelope = strEnvelope & "<Item:artifact>" & parm1 & "</Item:artifact>"
    Case "AddUserAppInfo"
        strEnvelope = strEnvelope & "<Item:account>" & parm1 & "</Item:account>"
        strEnvelope = strEnvelope & "<Item:appID>" & parm2 & "</Item:appID>"
        strEnvelope = strEnvelope & "<Item:appInfo>" & parm3 & "</Item:appInfo>"
    End Select
    strEnvelope = strEnvelope & "</Item:" & strMethod & ">"
    strEnvelope = strEnvelope & "</soapenv:Body>"
    strEnvelope = strEnvelope & "</soapenv:Envelope>"
    
    
    SoapEnvelope = strEnvelope
   
    Exit Function
errHand:
    
End Function

Private Function PostData(ByVal strPostURL As String, _
                        ByVal strMethod As String, _
                        ByVal strPostContent As String, _
                        Optional ByVal intSendWaitTime As Integer = 30) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim lngWaitTimeOut As Long
    Dim oXmlDoc As Object
    Dim strPostCookie As String
    
    On Error GoTo errHand
        
    If UCase(Left(strPostURL, 4)) <> "HTTP" Then strPostURL = "http://" & strPostURL
    strPostCookie = "ASPSESSIONIDAQACTAQB=HKFHJOPDOMAIKGMPGBJJDKLJ;"
    
    strPostCookie = Replace(strPostCookie, Chr(32), "%20")
    With mobjHttp
        Call .Open("POST", strPostURL, True)
        Select Case strMethod
        Case "GetSAMLResponseByArtifact"
            Call .setRequestHeader("SOAPAction", "http://tempuri.org/ISSOService/GetSAMLResponseByArtifact")
        Case "AddUserAppInfo"
            Call .setRequestHeader("SOAPAction", "http://tempuri.org/IAccountService/AddUserAppInfo")
        End Select
        Call .setRequestHeader("Content-Length", LenB(strPostContent))
        Call .setRequestHeader("Content-Type", "text/xml; charset=utf-8")
        Call .send(strPostContent)
    End With
    lngWaitTimeOut = 0
'    lngSecondNumber = 30 '超时多少秒
    Do
        DoEvents
        Call Wait(10)
        lngWaitTimeOut = lngWaitTimeOut + 1
    Loop Until (mobjHttp.readyState = 4 Or lngWaitTimeOut >= 100 * intSendWaitTime)
    
    If mobjHttp.readyState = 4 Then
        Set oXmlDoc = CreateObject("MSXML2.DOMDocument")

        oXmlDoc.Load mobjHttp.ResponseXML
        If oXmlDoc.xml = "" Then
            mstrError = mobjHttp.responseText
            PostData = False
        Else
            mstrPostData = oXmlDoc.xml
            PostData = True
        End If
    Else
        mstrError = mobjHttp.responseText
        PostData = False
    End If
    Exit Function
    
errHand:
    mstrError = Err.Description
End Function


Private Sub Wait(tt)
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim t, t1, t2, i
    t = tt
    If t > 10 Then
        t1 = Int(t / 10)
        t2 = t - t1 * 10
        For i = 1 To t1
            Call OS.Wait(10)
            DoEvents
        Next i
        If t2 > 0 Then Call OS.Wait(t2)
    Else
        If t > 0 Then Call OS.Wait(t)
    End If
End Sub

Private Sub ClearValues()
    '清理变量
    mblnFirst = False
    mintTimes = 1
    mbln转换 = True
    mblnAccess = False
    mblnUAAddUser = False
    
    mstrPostData = ""
    mstr断言 = ""
    mstrUserURL = ""
    mstrSamlAssertion = ""
    mstrError = ""
    mblnZLUA = False
    mstrAppID = ""
    mstrZLUAUser = ""
    mblnOK = False
End Sub

Public Function ShowMe() As Boolean
    '清理变量
    Call ClearValues
    Me.Show vbModal
End Function

Public Function Docmd(ByVal strCmd As String) As Boolean
    Dim ArrCommand
    Dim ArrCommandPortal
    Dim objSoap As Object
    Dim objDoc As Object
    Dim rsIni As ADODB.Recordset
    Dim strIp As String
    Dim strList As String
    Dim strResult As String
    Dim i As Integer
    Dim strPortURL As String
    Dim ResponseXML As Object
    Dim ResponseNode As Object
    Dim strArtifact助诊符 As String
    Dim strStatus As String
    Dim arrSamlAssertion() As String
    Dim strSoapPost As String
    Dim strErr As String
    Dim strAppStart As String
    On Error GoTo errHand
    '清理变量
    Call ClearValues
    'ZLUA登录
    strAppStart = gobjFile.GetParentFolderName(App.Path)
    If Len(strCmd) > 0 And InStr(strCmd, ",") = 0 And InStr(gstrCommand, "&") > 0 Then
        
        If Not gobjFile.FileExists(strAppStart & "\" & "ZLUA.ini") Then
            MsgBox "未找到" & strAppStart & "\" & "ZLUA.ini，无法读取配置文件", vbInformation + vbOKOnly, "提示"
            GoTo errHand
        End If
        Set rsIni = ReadINIToRec(strAppStart & "\" & "ZLUA.ini")
        rsIni.Filter = ""
        rsIni.Filter = "项目='PortURL'"
        strPortURL = rsIni("内容").Value
        rsIni.Filter = ""
        rsIni.Filter = "项目='UserURL'"
        mstrUserURL = rsIni("内容").Value
        rsIni.Filter = "项目='AppID'"
        mstrAppID = rsIni("内容").Value
        
        strArtifact助诊符 = Split(gstrCommand, "&")(0)
        
        If Trim(strPortURL) = "" Then
            MsgBox "请配置单点登录服务地址", vbInformation + vbOKOnly, "提示"
        ElseIf (Trim(mstrUserURL) = "") Then
            MsgBox "请配置账户服务地址", vbInformation + vbOKOnly, "提示"
        Else
            '采用httprequest方式-----------------
            mstr断言 = SoapEnvelope("GetSAMLResponseByArtifact", strArtifact助诊符, "", "", "")
            Call PostData(strPortURL, "GetSAMLResponseByArtifact", mstr断言, 5)
            strSoapPost = mstrPostData
            strSoapPost = Replace(strSoapPost, "&gt;", ">")
            strSoapPost = Replace(strSoapPost, "&lt;", "<")
            
            '-------------
            '解析XML文本内容并判断是否返回正确验证结果
            If strSoapPost <> "" Then
                Set objDoc = CreateObject("MSXML2.DOMDocument")
                Call objDoc.loadXML(strSoapPost)
                Set ResponseXML = objDoc.documentElement
                Set ResponseNode = ResponseXML.selectSingleNode(".//samlp:StatusCode")
                strStatus = ResponseNode.Attributes(0).Text
                If strStatus <> "" Then
                    Select Case strStatus
                    Case "urn:oasis:names:tc:SAML:2.0:status:Success"
                        '令牌请求成功
                        '获取登录信息:用户名/密码/服务器
                        Set ResponseNode = ResponseXML.selectSingleNode(".//saml:AttributeValue")
                        If ResponseNode Is Nothing Then
                            strStatus = ""
                        Else
                            strStatus = ResponseNode.Text
                        End If
                        
                        '获取ZLUA账户名
                        Set ResponseNode = ResponseXML.selectSingleNode(".//saml:NameID")
                        mstrZLUAUser = ResponseNode.Text
                        
                        Set ResponseNode = ResponseXML.selectSingleNode(".//saml:Assertion")
                        mstrSamlAssertion = ResponseNode.xml
                        '如果信息为空，则显示登录信息框，并调用接口上传信息以便下次成功获取
                        mblnZLUA = True
                        If Trim(strStatus) = "" Then
                            mblnUAAddUser = True
                            '--测试添加ZLUA用户账户
                        Else
                            If InStr(strStatus, "/") > 0 And InStr(strStatus, "@") > 0 And InStr(strStatus, "/") < InStr(strStatus, "@") Then
                               Me.txtUser.Text = Mid(strStatus, 1, InStr(strStatus, "/") - 1)
                               Me.txtPassWord.Text = Mid(strStatus, InStr(strStatus, "/") + 1, InStr(strStatus, "@") - InStr(strStatus, "/") - 1)
                               Me.cboServer.Text = Mid(strStatus, InStr(strStatus, "@") + 1)
                            End If
                            If Trim(txtUser.Text) <> "" And Trim(txtPassWord.Text) <> "" Then cmdOK_Click
                        End If
                    Case Else
                        '令牌请求失败，重新获取三级错误信息
                        Set ResponseNode = ResponseXML.selectSingleNode(".//samlp:StatusMessage")
                        strStatus = ResponseNode.Text
                        strErr = "错误信息：" & strStatus
                        GoTo errHand
                    End Select
                End If
            End If
            
        End If
    End If

    '单点登录
    ReDim ArrCommandPortal(0)
    If InStr(strCmd, ",") > 0 Then
        If objSoap Is Nothing Then
            Set objSoap = CreateObject("MSSOAP.SoapClient30")
        End If
        
        If Err.Number <> 0 Then
            Screen.MousePointer = 0
            Err.Clear
            MsgBox "无法创建SOAP对象！", vbOKOnly + vbInformation, "提示"
            Set objSoap = Nothing
            GoTo errHand
        End If
        If Not gobjFile.FileExists(strAppStart & "\" & "Portal.ini") Then
            MsgBox "未找到 " & strAppStart & "\" & "Portal.ini 路径", vbInformation + vbOKOnly, "提示"
            GoTo errHand
        End If
        Set rsIni = ReadINIToRec(strAppStart & "\" & "Portal.ini")
        rsIni.Filter = ""
        rsIni.Filter = "项目='IP'"
        strIp = rsIni("内容").Value
        rsIni.Filter = ""
        rsIni.Filter = "项目='List'"
        strList = rsIni("内容").Value
        '以前丢失，10.35.10新增
        ArrCommandPortal = Split(strCmd, ",")
    End If
    
    ArrCommand = Split(strCmd, " ")
    
    If UBound(ArrCommandPortal) > 0 Then
        Call objSoap.MSSoapInit("http://" & strIp & "/" & strList & "?wsdl")
        strResult = objSoap.getZLSSORet(ArrCommandPortal(0), ArrCommandPortal(1))
        If strResult <> "" And InStr(strResult, "/") > 0 And InStr(strResult, "@") > 0 And InStr(strResult, "/") < InStr(strResult, "@") Then
           Me.txtUser.Text = Mid(strResult, 1, InStr(strResult, "/") - 1)
           Me.txtPassWord.Text = Mid(strResult, InStr(strResult, "/") + 1, InStr(strResult, "@") - InStr(strResult, "/") - 1)
           Me.cboServer.Text = Mid(strResult, InStr(strResult, "@") + 1)
        End If
        mbln转换 = True
        If Trim(txtUser.Text) <> "" And Trim(txtPassWord.Text) <> "" Then cmdOK_Click
    ElseIf InStr(ArrCommand(0), "=") > 0 And InStr(ArrCommand(0), "&") = 0 Then
        '第三方部件调用导航台登录的格式
        For i = LBound(ArrCommand) To UBound(ArrCommand)
            If UCase(ArrCommand(i)) Like "USER=*" Then
                Me.txtUser.Text = Split(ArrCommand(i), "=")(1)
            ElseIf UCase(ArrCommand(i)) Like "PASS=*" Then
                Me.txtPassWord.Text = Split(ArrCommand(i), "=")(1)
            ElseIf UCase(ArrCommand(i)) Like "SERVER=*" Then
                Me.cboServer.Text = Split(ArrCommand(i), "=")(1)
            ElseIf UCase(ArrCommand(i)) Like "ONLYONE=*" Then
                If Split(ArrCommand(i), "=")(1) = "1" Then
                    If App.PrevInstance = True Then
                        MsgBox "不能重复运行这个程序！"
                        gblnExitApp = True
                        Exit Function
                    End If
                End If
            End If
        Next
        If Trim(txtUser.Text) <> "" And Trim(txtPassWord.Text) <> "" Then Call cmdOK_Click
    End If
    Docmd = mblnAccess
    Set objSoap = Nothing
    Exit Function
errHand:
    If strErr <> "" Then
        MsgBox strErr, vbInformation + vbOKOnly, "提示"
        strErr = ""
    Else
        If Err.Number <> 0 Then
            MsgBox Err.Description, vbInformation + vbOKOnly, "提示"
        End If
    End If
    Set objSoap = Nothing
    Err.Clear
End Function

Private Function GetXMLVersion() As String
    
    Dim varXMLVersion As Variant
    Dim strXMLVer As String
    Dim intLoop As Integer
    Dim objXML As Object
    
    On Error GoTo errHand
        
    varXMLVersion = Split("6.0,4.0", ",")
    
    On Error Resume Next
    For intLoop = 0 To UBound(varXMLVersion)
        Err = 0
        Set objXML = CreateObject("MSXML2.DOMDocument." & varXMLVersion(intLoop))
        If Err = 0 Then
            strXMLVer = varXMLVersion(intLoop)
            Exit For
        End If
    Next
    On Error GoTo errHand
    
    If strXMLVer = "" Then
        MsgBox "创建MSXML2.DOMDocument对象失败"
        Exit Function
    End If
    
    GetXMLVersion = strXMLVer
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
    MsgBox Err.Description
End Function

Private Sub txtUser_LostFocus()
    Call UpdateUser
End Sub

Private Sub txtUser_Validate(Cancel As Boolean)
    Call UpdateUser
End Sub

Private Sub UpdateUser()
On Error GoTo errH
    If IsNumeric(txtUser.Text) Then
        txtUser.Text = "U" & txtUser.Text
    End If
    Exit Sub
errH:
    MsgBox Err.Description, vbCritical, gstrSysName
    Err.Clear
End Sub

Private Function CheckInput(ByRef strUserName As String, ByRef strPassword As String, ByRef strServerName As String) As Boolean
'功能:检查用户，密码，服务器的输入值
    '分离字符串
    Dim intPos As Integer, strNote As String
    
    On Error GoTo InputError
    '------检验用户是否oracle合法用户----------------
    strUserName = Trim(txtUser.Text)
    strPassword = Trim(txtPassWord.Text)
    strServerName = Trim(cboServer.Text)
    
    '有效字符串效验
    If Len(Trim(txtUser.Text)) = 0 Then
        strNote = "请输入用户名"
        txtUser.SetFocus
        GoTo InputError
    End If
    
    If Len(strUserName) <> 1 Then
        If Mid(strUserName, 1, 1) = "/" Or Mid(strUserName, 1, 1) = "@" Or Mid(strUserName, Len(strUserName) - 1, 1) = "/" Or Mid(strUserName, Len(strUserName) - 1, 1) = "@" Then
            txtUser.SetFocus
            strNote = "用户名错误"
            SetConState
            Exit Function
        End If
    End If
    
    If Trim(strPassword) <> "" And Len(strPassword) <> 1 Then
        If Mid(strPassword, Len(strPassword) - 1, 1) = "/" Or Mid(strPassword, Len(strPassword) - 1, 1) = "@" Or Mid(strPassword, 1, 1) = "/" Or Mid(strPassword, 1, 1) = "@" Then
            If txtPassWord.Enabled Then txtPassWord.SetFocus
            strNote = "密码错误"
            GoTo InputError
        End If
    End If
    If Trim(strServerName) <> "" Then
        If Mid(strServerName, Len(strServerName) - 1, 1) = "/" Or Mid(strServerName, Len(strServerName) - 1, 1) = "@" Or Mid(strServerName, 1, 1) = "/" Or Mid(strServerName, 1, 1) = "@" Then
            strNote = "主机连接串错误"
            cboServer.SetFocus
            GoTo InputError
        End If
    End If
    
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
    CheckInput = True
    Exit Function
InputError:
    If strNote <> "" Then
        MsgBox strNote, vbExclamation, gstrSysName
    End If
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

Private Sub InitFaceType()
    cmdModify.Enabled = gintCallType = 0
    cmdModify.Visible = gintCallType = 0
    cmdSet.Enabled = gintCallType = 1
    cmdSet.Visible = gintCallType = 1
End Sub

Private Function CheckPwdExpiry() As Boolean
    Dim strSQL As String
    Dim rsData As ADODB.Recordset
    Dim dtExpiryDate As Date
    Dim dtNow As Date
    Dim intDiff As Integer
    
    strSQL = "Select EXPIRY_DATE From User_Users Where UserName=User"
    Set rsData = OpenSQLRecord(strSQL, "检查密码期效")
    
    If rsData.BOF = False Then
        If IsNull(rsData("EXPIRY_DATE").Value) = True Then
            CheckPwdExpiry = False
            Exit Function
        End If
        dtExpiryDate = Format(rsData("EXPIRY_DATE").Value, "YYYY-MM-DD HH:MM:SS")
        '判断过期日期与当前日期相差天数
        dtNow = Format(Currentdate, "YYYY-MM-DD HH:MM:SS")
       
        intDiff = DateDiff("d", dtNow, dtExpiryDate)
        
        If intDiff > 7 Then
            CheckPwdExpiry = False
            Exit Function
        End If
        
        If intDiff > 3 And intDiff <= 7 Then
            '提示修改密码
            If MsgBox("密码有效期还有" & intDiff & "天,是否立即修改密码?", vbQuestion + vbYesNo, "密码过期提醒") = vbYes Then
                CheckPwdExpiry = True
            Else
                CheckPwdExpiry = False
                Exit Function
            End If
        ElseIf intDiff <= 3 Then
            CheckPwdExpiry = True
            MsgBox "密码有效期还有" & intDiff & "天，请您立即修改密码。", vbInformation
        Else
            CheckPwdExpiry = False
            Exit Function
        End If
    End If
End Function

