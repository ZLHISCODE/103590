VERSION 5.00
Begin VB.Form frmLabAuditingLand 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "审核人登陆"
   ClientHeight    =   3240
   ClientLeft      =   30
   ClientTop       =   330
   ClientWidth     =   4800
   DrawMode        =   1  'Blackness
   Icon            =   "frmLabAuditingLand.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmLabAuditingLand.frx":000C
   ScaleHeight     =   3240
   ScaleWidth      =   4800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.CheckBox chkClueTo 
      Caption         =   "审核时不需要提示:"
      Height          =   195
      Left            =   300
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   2460
      Width           =   1830
   End
   Begin VB.ComboBox cboHour 
      Height          =   300
      ItemData        =   "frmLabAuditingLand.frx":E14E
      Left            =   2160
      List            =   "frmLabAuditingLand.frx":E19A
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   2130
      Width           =   795
   End
   Begin VB.Frame Frame1 
      Height          =   30
      Left            =   -30
      TabIndex        =   7
      Top             =   1995
      Width           =   4980
   End
   Begin VB.CheckBox chk有限时间 
      Caption         =   "权限有效期(小时):"
      Height          =   195
      Left            =   300
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   2183
      Width           =   1830
   End
   Begin VB.CommandButton cmd放弃 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   3345
      TabIndex        =   6
      Top             =   2730
      Width           =   1100
   End
   Begin VB.CommandButton cmd确认 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   2235
      TabIndex        =   4
      Top             =   2730
      Width           =   1100
   End
   Begin VB.TextBox txt密码 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   1695
      MaxLength       =   20
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   1500
      Width           =   2550
   End
   Begin VB.TextBox txt用户 
      Height          =   300
      Left            =   1695
      MaxLength       =   30
      TabIndex        =   1
      Top             =   1065
      Width           =   2550
   End
   Begin VB.Label lbl密码 
      AutoSize        =   -1  'True
      Caption         =   "密  码(&P)"
      Height          =   180
      Left            =   600
      TabIndex        =   2
      Top             =   1560
      Width           =   810
   End
   Begin VB.Label lbl用户名 
      AutoSize        =   -1  'True
      Caption         =   "审核人(&U)"
      Height          =   180
      Left            =   600
      TabIndex        =   0
      Top             =   1125
      Width           =   810
   End
End
Attribute VB_Name = "frmLabAuditingLand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mintTimes As Integer                        '登录重试次数
Private mintAuditing As Integer                     '是否有审核权限 0=没有权限 1=有权限 -1至-24=有效时间
Private mstrVerifyMan As String                     '检验人
Private mblnCancel As Boolean                       '是否是取消
Private mstrLogID As String                         '登陆名
Private Sub chk修改密码_Click()
    Me.cboHour.ListIndex = 0
End Sub

Private Sub chk有限时间_Click()
    If chk有限时间.Value = 1 Then
        Me.cboHour.Enabled = True
    Else
        Me.cboHour.Enabled = False
    End If
End Sub

Private Sub cmd确认_Click()
    Dim strNote As String
    Dim strUserName As String
    Dim strServerName As String
    Dim strPassword As String
    Dim strsql As String
    Dim rsTmp As New ADODB.Recordset
    Dim blGood As Boolean                                           '有"审核标本"权限
    
    zlDatabase.SetPara "是否有具有审核权限", 0, 100, 1208
    mstrLogID = ""
    mintTimes = mintTimes + 1
    '------检验用户是否oracle合法用户----------------
    strUserName = Trim(txt用户.Text)
    strPassword = Trim(txt密码.Text)
    
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
            Exit Sub
        End If
    End If
    If Trim(strPassword) <> "" And Len(strPassword) <> 1 Then
        If Mid(strPassword, Len(strPassword) - 1, 1) = "/" Or Mid(strPassword, Len(strPassword) - 1, 1) = "@" Or Mid(strPassword, 1, 1) = "/" Or Mid(strPassword, 1, 1) = "@" Then
            If txt密码.Enabled Then txt密码.SetFocus
            strNote = "密码错误"
            GoTo InputError
        End If
    End If
    
    '分离字符串
    Dim intPos As Integer
    intPos = InStr(1, strUserName, "@", vbTextCompare)
    If intPos > 0 Then
        strServerName = Mid(strUserName, intPos + 1)
        strUserName = Mid(strUserName, 1, intPos - 1)
    End If
    
    intPos = InStr(1, strUserName, "/", vbTextCompare)
    If intPos > 0 Then
        strPassword = Mid(strUserName, intPos + 1)
        strUserName = Mid(strUserName, 1, intPos - 1)
    End If
    
    intPos = InStr(1, strPassword, "@", vbTextCompare)
    If intPos > 0 Then
        strServerName = Mid(strPassword, intPos + 1)
        strPassword = Mid(strPassword, 1, intPos - 1)
    End If
    
    If Len(Trim(strPassword)) = 0 Then
        strNote = "请输入密码"
                    If txt密码.Enabled Then txt密码.SetFocus
        GoTo InputError
    End If
    
    strServerName = GetSetting(appName:="ZLSOFT", Section:="注册信息\登陆信息", Key:="SERVER", Default:="")
    
    If Not OraDataOpen(strServerName, UCase(strUserName), strPassword) Then
        txt密码.Text = ""
        If txt密码.Enabled Then txt密码.SetFocus
        Exit Sub
    End If
    
    blGood = True

'    strSQL = "select user from dual "
'    Call zldatabase.OpenRecordset(rsTmp, strSQL, gstrSysName)
    
    If UCase(strUserName) = UserInfo.用户名 Then
        MsgBox "您登陆的用户和现在的用户为同一用户,请重新登陆!", vbInformation, gstrSysName
        Me.txt密码 = ""
        Me.txt用户 = ""
        Me.txt用户.SetFocus
        Exit Sub
    End If
        
    
'    strSQL = "select 所有者 from zlsystems where 编号 =100 and 所有者 = [1] "
'    Set rsTmp = zldatabase.OpenSQLRecord(strSQL, gstrSysName, strUserName)
'
'    If rsTmp.EOF = True Then
'        strSQL = "Select 功能 From dba_role_privs A, zlRoleGrant B " & vbCrLf & _
'                " Where Granted_Role = B.角色 And grantee = [1] And Granted_Role Like 'ZL_%' " & vbCrLf & _
'                " And 系统 = [2] And 序号 = [3] and 功能 = [4] "
'        Set rsTmp = zldatabase.OpenSQLRecord(strSQL, gstrSysName, strUserName, glngSys, glngModul, "审核标本")
'
'        If rsTmp.EOF = False Then
'            Do Until rsTmp.EOF
'                If rsTmp("功能") = "审核标本" Then
'                    blGood = False
'                    Exit Do
'                End If
'                rsTmp.MoveNext
'            Loop
'        End If
'    End If
    
    If blGood = False Then
        MsgBox "你所登陆的用户没有<审核>权限!", vbInformation, gstrSysName
        Me.txt用户.SetFocus
        Exit Sub
    End If
    
    If Me.chk有限时间.Value = 0 Then
        mintAuditing = 1
    Else
        mintAuditing = -CInt(Me.cboHour.Text)
    End If
    
    strsql = "select b.姓名 from 上机人员表 a ,人员表 b where 用户名 = [1] and a.人员id = b.id " & vbCrLf & _
             " And (b.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or b.撤档时间 Is Null) "
    Set rsTmp = zlDatabase.OpenSQLRecord(strsql, gstrSysName, UCase(strUserName))
    
    If rsTmp.EOF = False Then
        mstrLogID = UCase$(strUserName)
        zlDatabase.SetPara "审核人", rsTmp("姓名"), 100, 1208
    Else
        MsgBox "用户已停用!不能在进行审核.", vbInformation, gstrSysName
        mintAuditing = 0
    End If
    
    zlDatabase.SetPara "是否有具有审核权限", mintAuditing, 100, 1208
    zlDatabase.SetPara "审核时不需要提示", chkClueTo.Value, 100, 1208
    
    Unload Me
    Exit Sub

InputError:
    If mintTimes > 3 Then
        MsgBox "超过三次登录失败，自动退出", vbExclamation, gstrSysName
        Call cmd放弃_Click
    Else
        If strNote <> "" Then
            MsgBox strNote, vbExclamation, gstrSysName
        End If
        Exit Sub
    End If

End Sub

Private Sub cmd放弃_Click()
    mblnCancel = True
    Unload Me
End Sub

Private Sub Form_Activate()
'    Me.txt用户.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = vbKeyReturn Then
'        If Me.ActiveControl.Name = "TXT密码" Then
'            Call cmd确认_Click
'        Else
'            SendKeys "{Tab}"
'        End If
'    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    mintTimes = 1
    Me.chk有限时间.Value = zlDatabase.GetPara("frmLabAuditingLand_时限", 100, 1208, 0)
    Me.cboHour.ListIndex = zlDatabase.GetPara("frmLabAuditingLand_时间", 100, 1208, 0)
    If Me.chk有限时间.Value = 0 Then
        Me.cboHour.Enabled = False
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    zlDatabase.SetPara "frmLabAuditingLand_时限", Me.chk有限时间.Value, 100, 1208
    zlDatabase.SetPara "frmLabAuditingLand_时间", Me.cboHour.ListIndex, 100, 1208
End Sub

Private Sub txt密码_GotFocus()
    Me.txt密码.SelStart = 0: Me.txt密码.SelLength = Len(Me.txt密码.Text)
End Sub

Private Sub txt密码_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub txt用户_GotFocus()
    Me.txt用户.SelStart = 0: Me.txt用户.SelLength = Len(Me.txt用户.Text)
End Sub

Public Function TranPasswd(strOld As String) As String
    '------------------------------------------------
    '功能： 密码转换函数
    '参数：
    '   strOld：原密码
    '返回： 加密生成的密码
    '------------------------------------------------
    Dim iBit As Integer, strBit As String
    Dim strNew As String
    If Len(Trim(strOld)) = 0 Then TranPasswd = "": Exit Function
    strNew = ""
    For iBit = 1 To Len(Trim(strOld))
        strBit = UCase(Mid(Trim(strOld), iBit, 1))
        Select Case (iBit Mod 3)
        Case 1
            strNew = strNew & _
                Switch(strBit = "0", "W", strBit = "1", "I", strBit = "2", "N", strBit = "3", "T", strBit = "4", "E", strBit = "5", "R", strBit = "6", "P", strBit = "7", "L", strBit = "8", "U", strBit = "9", "M", _
                   strBit = "A", "H", strBit = "B", "T", strBit = "C", "I", strBit = "D", "O", strBit = "E", "K", strBit = "F", "V", strBit = "G", "A", strBit = "H", "N", strBit = "I", "F", strBit = "J", "J", _
                   strBit = "K", "B", strBit = "L", "U", strBit = "M", "Y", strBit = "N", "G", strBit = "O", "P", strBit = "P", "W", strBit = "Q", "R", strBit = "R", "M", strBit = "S", "E", strBit = "T", "S", _
                   strBit = "U", "T", strBit = "V", "Q", strBit = "W", "L", strBit = "X", "Z", strBit = "Y", "C", strBit = "Z", "X", True, strBit)
        Case 2
            strNew = strNew & _
                Switch(strBit = "0", "7", strBit = "1", "M", strBit = "2", "3", strBit = "3", "A", strBit = "4", "N", strBit = "5", "F", strBit = "6", "O", strBit = "7", "4", strBit = "8", "K", strBit = "9", "Y", _
                   strBit = "A", "6", strBit = "B", "J", strBit = "C", "H", strBit = "D", "9", strBit = "E", "G", strBit = "F", "E", strBit = "G", "Q", strBit = "H", "1", strBit = "I", "T", strBit = "J", "C", _
                   strBit = "K", "U", strBit = "L", "P", strBit = "M", "B", strBit = "N", "Z", strBit = "O", "0", strBit = "P", "V", strBit = "Q", "I", strBit = "R", "W", strBit = "S", "X", strBit = "T", "L", _
                   strBit = "U", "5", strBit = "V", "R", strBit = "W", "D", strBit = "X", "2", strBit = "Y", "S", strBit = "Z", "8", True, strBit)
        Case 0
            strNew = strNew & _
                Switch(strBit = "0", "6", strBit = "1", "J", strBit = "2", "H", strBit = "3", "9", strBit = "4", "G", strBit = "5", "E", strBit = "6", "Q", strBit = "7", "1", strBit = "8", "X", strBit = "9", "L", _
                   strBit = "A", "S", strBit = "B", "8", strBit = "C", "5", strBit = "D", "R", strBit = "E", "7", strBit = "F", "M", strBit = "G", "3", strBit = "H", "A", strBit = "I", "N", strBit = "J", "F", _
                   strBit = "K", "O", strBit = "L", "4", strBit = "M", "K", strBit = "N", "Y", strBit = "O", "D", strBit = "P", "2", strBit = "Q", "T", strBit = "R", "C", strBit = "S", "U", strBit = "T", "P", _
                   strBit = "U", "B", strBit = "V", "Z", strBit = "W", "0", strBit = "X", "V", strBit = "Y", "I", strBit = "Z", "W", True, strBit)
        End Select
    Next
    TranPasswd = strNew

End Function

Private Function OraDataOpen(ByVal strServerName As String, ByVal strUserName As String, ByVal strUserPwd As String) As Boolean
    '------------------------------------------------
    '功能： 打开指定的数据库
    '参数：
    '   strServerName：主机字符串
    '   strUserName：用户名
    '   strUserPwd：密码
    '返回： 数据库打开成功，返回true；失败，返回false
    '------------------------------------------------
    Dim strsql As String
    Dim strError As String
    Dim cnOracle As New ADODB.Connection
    Dim objzlRegister As Object
    
    
    '兼容性设置,如果有zlRegister则采用新的注册方式,否则采用老的注册方式
    On Error GoTo errhandOld
    If objzlRegister Is Nothing Then Set objzlRegister = CreateObject("zlRegister.clsRegister")
    If Not objzlRegister.LoginValidate(strServerName, strUserName, strUserPwd, strError) Then
        If strError <> "" Then
            MsgBox strError, vbInformation, "登陆"
        End If
        Exit Function
    End If
    OraDataOpen = True
    Exit Function
errhandOld:


    On Error Resume Next
    Err = 0
    DoEvents
    With cnOracle
        If .State = adStateOpen Then .Close
        .Provider = "MSDataShape"
        .Open "Driver={Microsoft ODBC for Oracle};Server=" & strServerName, strUserName, IIf(UCase(strUserName) = "SYS" Or UCase(strUserName) = "SYSTEM", strUserPwd, TranPasswd(strUserPwd))
        If Err <> 0 Then
            '保存错误信息
            strError = Err.Description
            If InStr(strError, "自动化错误") > 0 Then
                MsgBox "连接串无法创建，请检查数据访问部件是否正常安装。", vbInformation, gstrSysName
            ElseIf InStr(strError, "ORA-12154") > 0 Then
                MsgBox "无法分析服务器名，" & vbCrLf & "请检查在Oracle配置中是否存在该本地网络服务名（主机字符串）。", vbInformation, gstrSysName
            ElseIf InStr(strError, "ORA-12541") > 0 Then
                MsgBox "无法连接，请检查服务器上的Oracle监听器服务是否启动。", vbInformation, gstrSysName
            ElseIf InStr(strError, "ORA-01033") > 0 Then
                MsgBox "ORACLE正在初始化或在关闭，请稍候再试。", vbInformation, gstrSysName
            ElseIf InStr(strError, "ORA-01034") > 0 Then
                MsgBox "ORACLE不可用，请检查服务或数据库实例是否启动。", vbInformation, gstrSysName
            ElseIf InStr(strError, "ORA-02391") > 0 Then
                MsgBox "用户" & UCase(strUserName) & "已经登录，不允许重复登录(已达到系统所允许的最大登录数)。", vbExclamation, gstrSysName
            ElseIf InStr(strError, "ORA-01017") > 0 Then
                MsgBox "由于用户、口令或服务器指定错误，无法注册。", vbInformation, gstrSysName
            Else
                MsgBox strError, vbInformation, gstrSysName
            End If
            
            OraDataOpen = False
            Exit Function
        End If
    End With
    
    Err = 0
    On Error GoTo ErrHand
    cnOracle.Close
    Set cnOracle = Nothing
        
    OraDataOpen = True
    Exit Function
    
ErrHand:
    If ErrCenter() = 1 Then Resume
    OraDataOpen = False
    Err = 0
End Function

Private Sub txt用户_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub
Public Sub ShowMe(objfrm As Object, strReportMan As String, blnCancel As Boolean, strLogID As String)
    '功能               显示审核人登陆窗体
    '参数               Objfrm 父窗体对象
    '                   strReportMan 报告人
    mstrVerifyMan = strReportMan
    mblnCancel = False
    Me.Show vbModal, objfrm
    blnCancel = mblnCancel
    strLogID = mstrLogID
End Sub


