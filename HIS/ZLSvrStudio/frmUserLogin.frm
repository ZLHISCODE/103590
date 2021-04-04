VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmUserLogin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "管理工具登录"
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
      Height          =   350
      Left            =   150
      TabIndex        =   10
      ToolTipText     =   "启动Oracle主机字符串配置程序"
      Top             =   2115
      Width           =   1100
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
      TabIndex        =   11
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
Private mstrCommand As String
Private mbln转换 As Boolean
Private mblnAccess As Boolean

Dim mcolServer As New Collection

Public Sub ShowMe(ByVal strCommand As String)
    mstrCommand = strCommand
    mbln转换 = True
    Me.Show vbModal
End Sub

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
        If glngSysNo <> -1 Then Me.Visible = False
        mblnAccess = False
        txt密码.Text = ""
        Exit Sub
    End If
    
    '修改注册表
    SaveSetting "ZLSOFT", "注册信息\登陆信息", "MANAGER", strUserName
    SaveSetting "ZLSOFT", "注册信息\登陆信息", "SERVER", strServerName
    mblnAccess = True
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

Private Function GetRegFunFile() As String
    With cdgFile
        .DialogTitle = "选择注册函数文件，重新创建相关函数"
        .Filter = IIf(GetOracleVersion(True, True) > 11, "注册函数文件(ZLREGIST12C.PLB)|ZLREGIST12C.PLB", "注册函数文件(ZLREGIST.PLB)|ZLREGIST.PLB")
        .flags = &H200000 Or &H4 Or &H2 Or &H800 Or &H4000
        .CancelError = False
        .ShowOpen
        
        GetRegFunFile = .Filename
    End With
End Function

Private Function RegCheckAndGetUnit() As String
'功能：验证系统注册授权的正确性，并返回单位名称
    Dim strUnit As String, strRegFunFile As String
    Dim strRegErr As String, strPassword As String, strError As String, strSQL As String
    Dim cnTools As ADODB.Connection
    Dim rstmp As ADODB.Recordset, blnLoginAgain As Boolean
    
    strRegErr = gobjRegister.zlRegCheck(False)
    If strRegErr <> "" Then
        Me.Visible = False
        If strRegErr Like "*恢复正确的注册函数！*" Then
            If GetOracleVersion(True, True) > 11 Then
                strRegErr = Replace(strRegErr, "ZLREGIST.PLB", "ZLREGIST12C.PLB")
            End If
            If MsgBox(strRegErr & vbCrLf & "现在要重新创建注册函数吗？", vbYesNo + vbDefaultButton2, "提示") = vbNo Then
                End
            Else
                '先用缺省密码ZLTOOLS执行连接
                strPassword = "ZLTOOLS"
                blnLoginAgain = True
openconn:       Set cnTools = gobjRegister.GetConnection(strServerName, "ZLTOOLS", strPassword, False, OraOLEDB, strError, False)
                If strError <> "" Then
                    If blnLoginAgain Then
                        strError = ""
                        strPassword = "ZLSOFT"
                        Set cnTools = gobjRegister.GetConnection(strServerName, "ZLTOOLS", strPassword, False, OraOLEDB, strError, False)
                        blnLoginAgain = False
                    End If
                End If
                
                If strError <> "" Then
                    strPassword = InputBox("注册函数验证失败，即将重新创建注册函数(" & IIf(GetOracleVersion(True, True) > 11, "ZLREGIST12C.PLB", "ZLREGIST.PLB") & ")。" & vbCrLf & "需要以ZLTOOLS用户登录执行，请输入该用户的密码", "提示")
                    If strPassword = "" Then
                        End
                    Else
                        strError = ""
                        GoTo openconn
                    End If
                End If
                
                On Error GoTo errH
                '1.检查注册函数所需的表结构是否需要修正
                strSQL = "Select Table_Name" & vbNewLine & _
                        "From User_Tab_Columns" & vbNewLine & _
                        "Where Table_Name In ('ZLREGFILE', 'ZLREGAUDIT') And Column_Name = '项目' And Data_Length <> 20"
                Set rstmp = gclsBase.OpenSQLRecord(cnTools, strSQL, "检查数据结构")
                If rstmp.RecordCount > 0 Then
                    
                    '2.如果需要修正，则弹出选择，需断掉所有ZLHIS客户端
                    If MsgBox("检测到注册相关数据结构需要修正，要求断开所有ZLHIS客户端并禁用应用系统账户。" & vbCrLf & "你确定现在立即结束所有ZLHIS客户端并禁用应用系统账户吗？" & vbCrLf & _
                            "注：请到系统升迁管理界面进行用户的解锁、客户端的启用。", vbQuestion + vbOKCancel + vbDefaultButton2, "提示") = vbCancel Then
                        End
                    Else
                        '断开所有ZLHIS客户端连接才能修改两张临时表的结构
                        Call LockAppUser
                        Call KillSessions
                    End If
                    
                    rstmp.Filter = "Table_Name='ZLREGFILE'"
                    If rstmp.RecordCount > 0 Then
                        strSQL = "Alter Table zlRegFile Modify 项目 Varchar2(20)"
                        cnTools.Execute strSQL
                    End If
                    
                    rstmp.Filter = "Table_Name='ZLREGAUDIT'"
                    If rstmp.RecordCount > 0 Then
                        strSQL = "Alter Table ZLREGAUDIT Modify 项目 Varchar2(20)"
                        cnTools.Execute strSQL
                    End If
                    
                    strSQL = "Drop Type t_Reg_Rowset Force"
                    cnTools.Execute strSQL
                    strSQL = "Drop Type t_Reg_Record Force"
                    cnTools.Execute strSQL
                    strSQL = "Create Or Replace Type t_Reg_Record  As Object(Item Varchar2(20), Prog number(18), Text Varchar2(1000))"
                    cnTools.Execute strSQL
                    strSQL = "Create Or Replace Type t_Reg_Rowset As Table Of t_Reg_Record"
                    cnTools.Execute strSQL
                    
                    
                    On Error Resume Next
                    strSQL = "Grant Execute on t_Reg_Record to Public"
                    cnTools.Execute strSQL
                    If err.Number <> 0 Then
                        MsgBox "执行包授权时失败，错误描述：" & vbCrLf & err.Description & vbCrLf _
                            & "不要点击确定，先进行相关处理后，手工运行以下脚本：" & vbCrLf & strSQL, vbExclamation, "提示"
                        err.Clear
                    End If
                    
                    '德阳医院存在ZLHIS的包T_DB_ROLEUSER（BH溶合相关的）引用了该对象，导致授权失败
                    'ORA-04045: 在重新编译/重新验证 ZLHIS.T_DB_ROLEUSER 时出错
                    'ORA -1031: 权限不足
                    strSQL = "Grant Execute on t_Reg_Rowset to Public"
                    cnTools.Execute strSQL
                    If err.Number <> 0 Then
                        MsgBox "执行包授权时失败，错误描述：" & vbCrLf & err.Description & vbCrLf _
                            & "不要点击确定，先进行相关处理后，手工运行以下脚本：" & vbCrLf & strSQL, vbExclamation, "提示"
                        err.Clear
                    End If
                    On Error GoTo errH
                End If
                
                
                '3.执行注册文件
                If Not gblnInIDE Then '增加多环境支持
                    strRegFunFile = App.Path & "\TOOLS\" & IIf(GetOracleVersion(True, True) > 11, "ZLREGIST12C.PLB", "zlRegist.plb")
                Else
                    strRegFunFile = "C:\APPSOFT\TOOLS\" & IIf(GetOracleVersion(True, True) > 11, "ZLREGIST12C.PLB", "zlRegist.plb")
                End If
                If gobjFile.FileExists(strRegFunFile) = False Then
                    strRegFunFile = GetRegFunFile
                    If strRegFunFile = "" Then
                        End
                    End If
                End If
                
                If RunRegistFile(Me, cnTools, strPassword, strServerName, strRegFunFile) Then
                    MsgBox "注册函数创建完成，请重新进行注册！" & vbCrLf & "注册函数来源：" & vbCrLf & strRegFunFile, vbInformation
                Else
                    End
                End If
            End If
        Else
            MsgBox "注册验证失败，请重新注册！" & vbCrLf & strRegErr, vbInformation, "提醒"
        End If
        
        If Not frmReg.ReReg Then
            End
        End If
    End If
    strUnit = gobjRegister.zlRegInfo("单位名称", False, 0)
    If strUnit = "" Then
        MsgBox "未能读取到单位名称，请检查注册码及注册函数，或者重新注册!", vbExclamation, "提醒"
        If Not frmReg.ReReg Then
            End
        End If
    End If
    RegCheckAndGetUnit = strUnit
    Exit Function
    
errH:
    MsgBox err.Description & vbCrLf & "最近一次执行的SQL：" & strSQL, vbExclamation, "提示"
    End
End Function

Private Function OraDataOpen(ByVal strServerName As String, ByVal strUserName As String, ByVal strPassword As String) As Boolean
'功能： 打开指定的数据库连接，如果是普通用户，则使用管理员帐号重新打开连接
'参数：
'   strServerName：主机字符串
'   strUserName：用户名
'   strUserPwd：密码
'返回： 数据库打开成功，返回true；失败，返回false

    Dim rsTemp As New ADODB.Recordset
    Dim strSQL As String
    Dim strDest() As Byte
    Dim StrJiemi() As Byte
    Dim StrJiami() As Byte
    Dim blnGrantMgr As Boolean '授权的工具所有者
    Dim strPwdTxt As String, strRegErr As String, strUnit As String
    Dim blnLogin As Boolean, blnTransPassword As Boolean
    Dim strHaveProg As String
    Dim strError As String
    
    '支持strServerName = "192.168.2.13:1521/dyyy"这种格式
    
    gstrLoginPwd = strPassword
    
    If UCase(strUserName) = "SYS" Or UCase(strUserName) = "SYSTEM" Then
        blnTransPassword = False
    Else
        blnTransPassword = mbln转换
    End If
    Set gcnOracle = gobjRegister.GetConnection(strServerName, strUserName, strPassword, blnTransPassword, OraOLEDB, strError)
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
    Call SetSQLTrace(strServerName, strUserName, gcnOracle)
    
    
    Call gobjRegister.zlRegInit(gcnOracle)
    
    On Error Resume Next
    strSQL = "SELECT 1 FROM ZLTOOLS.ZLSYSTEMS WHERE 所有者=USER"
    Set rsTemp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "系统所有者判定")
    If err.Number <> 0 Then
        gblnCreate = False
        gblnOwner = False
        err.Clear
    Else
        gblnCreate = True
        gblnOwner = Not rsTemp.EOF
    End If
    
    strSQL = "SELECT 1 FROM SESSION_ROLES WHERE ROLE='DBA'"
    Set rsTemp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "DBA判定")
    gblnDBA = Not rsTemp.EOF
        gblnRac = CheckRAC(gintInstID)
    If err.Number <> 0 Then err.Clear
    
    If Not (gblnDBA) And Not (gblnCreate) Then
        OraDataOpen = False
        MsgBox "首次运行，必须是DBA注册，以便创建管理工具！", vbExclamation, gstrSysName
        Exit Function
    End If
    
    
    '普通用户登录管理工具时，以系统所有者进行实质性连接
    If gblnCreate Then
        strUnit = RegCheckAndGetUnit
        If strUnit = "" Then End

        gstrHaveProg = "": blnGrantMgr = False: blnLogin = False
        gstrLoginUserName = strUserName
        gstrLoginUserPwd = gobjRegister.GetPassword
        
        If Not gblnDBA And Not gblnOwner Then
            '检查是否有管理工具的权限
            strSQL = "select 功能 from zltools.Zlmgrgrant Where 用户名='" & gstrLoginUserName & "'"
            Set rsTemp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "管理工具授权用户")
            If rsTemp.RecordCount > 0 Then
                gstrHaveProg = rsTemp!功能 & ""
                If gstrHaveProg <> "" Then
                    ReDim Preserve strDest(0): ReDim Preserve StrJiemi(0)
                    Call Func16CodeToByte(gstrHaveProg, strDest)
                    Call DES_Decode(strDest, StrJiemi, strUnit)
                    gstrHaveProg = Replace(StrConv(StrJiemi, vbUnicode), Chr(0), "")
                    
                    '对权限字符串进行初始化操作
                    gstrHaveProg = GetProgFuncs(gstrHaveProg, True)
                    
                    blnGrantMgr = True
                    
                    '判断是否为单系统登录
                    If glngSysNo <> -1 Then
                        If InStr(gstrHaveProg, "0401") Then
                            strHaveProg = "0401"
                        End If
                        If InStr(gstrHaveProg, "0402") Then
                            strHaveProg = IIf(strHaveProg = "", "", strHaveProg & ",") & "0402"
                        End If
                        gstrHaveProg = strHaveProg
                        If gstrHaveProg = "" Then
                            blnGrantMgr = False
                        End If
                    End If
                    
                End If
            End If
            If Not blnGrantMgr Then
                OraDataOpen = False
                MsgBox "您没有管理工具的使用权限，请联系管理员。", vbExclamation, gstrSysName
                Exit Function
            ElseIf gstrHaveProg = "" Then
                OraDataOpen = False
                MsgBox "您的管理工具的使用权限丢失，请联系管理员重新授权。", vbExclamation, gstrSysName
                Exit Function
            End If
            
            '使用系统管理员登录
            If err.Number <> 0 Then err.Clear
            strUserName = "": strPassword = ""
            strSQL = "Select Max(Decode(项目,'管理员',内容,'')) AS 管理员 ,Max(Decode(项目,'验证码',内容,'')) AS 验证码 From zltools.zlRegInfo where 项目='管理员' Or 项目='验证码'"
            Set rsTemp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "授权登录信息")
            If rsTemp!管理员 & "" <> "" And rsTemp!验证码 & "" <> "" Then
                strUserName = rsTemp!管理员 & ""
                ReDim Preserve strDest(0): ReDim Preserve StrJiemi(0)
                Call Func16CodeToByte(rsTemp!验证码 & "", strDest)
                Call DES_Decode(strDest, StrJiemi, strUnit)
                strPassword = Replace(StrConv(StrJiemi, vbUnicode), Chr(0), "")
                
                '重新打开数据库链接(存储的是数据库密码，所以不需要转换)
                Set gcnOracle = gobjRegister.GetConnection(strServerName, strUserName, strPassword, False, OraOLEDB)
                blnLogin = gcnOracle.State = adStateOpen
                
                If blnLogin Then
                    Call SetSQLTrace(strServerName, strUserName, gcnOracle)
                    '重新认证会话
                    Call gobjRegister.zlRegInit(gcnOracle)
                    strRegErr = gobjRegister.zlRegCheck(False)
                    If strRegErr <> "" Then
                        MsgBox strRegErr, vbQuestion, "提醒"
                        If Not frmReg.ReReg Then
                            End
                        End If
                    End If
                End If
            End If
            
            '不能使用管理员登录，要求重新输入管理员帐号密码
            If Not blnLogin Then
                MsgBox "管理员授权信息丢失，请验证管理员账户！", vbExclamation, gstrSysName
                If Not frmUserCheckLogin.ShowLogin(UCT_SysOwner, gcnOracle, strUserName, strServerName) Then Exit Function
                strPassword = gobjRegister.GetPassword
                Call SetSQLTrace(strServerName, strUserName, gcnOracle)
                '重新认证会话
                Call gobjRegister.zlRegInit(gcnOracle)
                strRegErr = gobjRegister.zlRegCheck(False)
                If strRegErr <> "" Then
                    MsgBox strRegErr, vbQuestion, "提醒"
                    If Not frmReg.ReReg Then
                        End
                    End If
                End If
                '未授权程序不更新管理员信息
                If Not strPassword Like "未授权的程序:*" Then
                    '更新管理员账户信息
                    strSQL = "Delete zltools.zlRegInfo where 项目='管理员' Or 项目='验证码'"
                    gcnOracle.Execute strSQL
                    strSQL = "Insert into zltools.zlRegInfo(项目,内容) values('管理员','" & strUserName & "')"
                    gcnOracle.Execute strSQL
                    
                    strPwdTxt = ""
                    ReDim Preserve StrJiami(0)
                    Call DES_Encode(StrConv(strPassword, vbFromUnicode), StrJiami, strUnit)
                    strPwdTxt = FuncByteTo16Code(StrJiami)
                    strSQL = "Insert into zltools.zlRegInfo(项目,内容) values('验证码','" & strPwdTxt & "')"
                    gcnOracle.Execute strSQL
                End If
            End If
            
            strSQL = "SELECT 1 FROM SESSION_ROLES WHERE ROLE='DBA'"
            Set rsTemp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "DBA判定")
            gblnDBA = Not rsTemp.EOF
            
            gblnOwner = True
        Else
            strPassword = gobjRegister.GetPassword
             '未授权程序不更新管理员信息
            If Not strPassword Like "未授权的程序:*" Then
                strSQL = "Select 1 From zltools.zlRegInfo where 项目='管理员'"
                Set rsTemp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "管理工具授权模式")
                If rsTemp.RecordCount > 0 Then
                    strSQL = "Update zltools.zlRegInfo Set 内容='" & strUserName & "' Where 项目='管理员' And 内容<>'" & strUserName & "'"
                    gcnOracle.Execute strSQL
                    '验证码
                    strPwdTxt = ""
                    ReDim Preserve StrJiami(0)
                    Call DES_Encode(StrConv(strPassword, vbFromUnicode), StrJiami, strUnit)
                    strPwdTxt = FuncByteTo16Code(StrJiami)
                    strSQL = "Select 1 From zlRegInfo where 项目='验证码'"
                    Set rsTemp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "验证码判断")
                    If rsTemp.RecordCount > 0 Then
                        strSQL = "Update zlRegInfo Set 内容='" & strPwdTxt & "' Where 项目='验证码'"
                    Else
                        strSQL = "Insert into zlRegInfo(项目,内容) values('验证码','" & strPwdTxt & "')"
                    End If
                    gcnOracle.Execute strSQL
                End If
            End If
            '若为单系统登录，则只赋予角色授权与用户授权两个模块的权限
            If glngSysNo <> -1 Then
                gstrHaveProg = "0401,0402"
            End If
        End If
    End If

    OraDataOpen = True
    gstrUserName = strUserName
    gstrPassword = gobjRegister.GetPassword
    gstrServer = Trim(strServerName)
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
    If Trim(txt用户.Text) <> "" And Trim(txt密码.Text) <> "" Then Call cmdOK_Click
End Sub

Private Sub Form_Load()
    Dim strFileInfo As String
    Dim ArrCommand() As String
    
    On Error GoTo errH
    txt用户.Text = GetSetting("ZLSOFT", "注册信息\登陆信息", "MANAGER", "")
    txt数据库.Text = GetSetting("ZLSOFT", "注册信息\登陆信息", "SERVER", "")
    intTimes = 0
    
    Set mcolServer = LoadServer(strFileInfo)
    txt数据库.ToolTipText = strFileInfo
    Call ApplyOEM_Picture(Me, "Icon")

    If Val(Me.Tag) = 1 Then
        Me.Hide
    Else
        '不加这一句的话，由于已显示frmSplash窗体，在开启输入法的情况下，启动源程序，不会显示登录窗口，VB只能异常终止退出
        SetActiveWindow Me.hwnd
    End If
    
    '如果含有/，表示同时输入了用户名与密码，而且密码不需要进行转换
    If mstrCommand <> "" Then
        ArrCommand = Split(mstrCommand, " ")
        If InStr(1, ArrCommand(0), "/") <> 0 And InStr(1, ArrCommand(0), ",") = 0 Then
            Me.txt用户.Text = Split(ArrCommand(0), "/")(0)
            Me.txt密码.Text = Split(ArrCommand(0), "/")(1)
            mbln转换 = False
        End If
    End If
    Exit Sub
errH:
    MsgBox err.Description, vbInformation, gstrSysName
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
                blnRead = Not GetRegValue("HKEY_LOCAL_MACHINE\SOFTWARE\Oracle" & !name, "ORACLE_HOME", strPath)
                blnRead = blnRead Or strPath = "" And !name & "" = ""
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
    
    If err <> 0 Then
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

Public Function Docmd(ByVal strCmd As String, ByRef blnAnalysis As Boolean) As Boolean
    '功能：Shell命令方式登录管理工具
    '参数
    'strCmd：Shell命令参数
    '     blnAnalysis：标记以第一种方式解析是否成功
    '     blnAnalysis=True，表示strCmd解析成功
    '     blnAnalysis=False，表示strCmd解析失败
    '如果命令行参数中有用户名及密码，则填充并执行
    Dim ArrCommand() As String
    Dim strUser As String, strPasswd As String, strServer As String
    Dim i As Long
    
    mblnAccess = False
    mbln转换 = True
    mstrCommand = strCmd
    ArrCommand = Split(strCmd, " ")
    If InStr(ArrCommand(0), "=") > 0 Then
        For i = LBound(ArrCommand) To UBound(ArrCommand)
            If UCase(ArrCommand(i)) Like "USER=*" Then
                strUser = Split(ArrCommand(i), "=")(1)
            ElseIf UCase(ArrCommand(i)) Like "PASS=*" Then
                strPasswd = Split(ArrCommand(i), "=")(1)
            ElseIf UCase(ArrCommand(i)) Like "SERVER=*" Then
                strServer = Split(ArrCommand(i), "=")(1)
            End If
        Next
    End If
    
    If strUser <> "" And strPasswd <> "" And strServer <> "" Then
        '表示是以第一种Shell命令格式登录
        Me.Tag = 1
        blnAnalysis = True
        Me.txt用户.Text = strUser
        Me.txt密码.Text = strPasswd
        Me.txt数据库.Text = strServer
        Call cmdOK_Click
    End If
    Docmd = mblnAccess
End Function
