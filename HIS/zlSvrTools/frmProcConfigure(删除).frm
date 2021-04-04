VERSION 5.00
Begin VB.Form frmProcConfigure 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "配置"
   ClientHeight    =   3348
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   4992
   Icon            =   "frmProcConfigure.frx":0000
   LinkTopic       =   "连接配置"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3348
   ScaleWidth      =   4992
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.TextBox txtPort 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   2190
      MaxLength       =   20
      TabIndex        =   4
      Top             =   2175
      Width           =   2625
   End
   Begin VB.TextBox txtServerIP 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   2190
      MaxLength       =   20
      TabIndex        =   2
      Top             =   1410
      Width           =   2625
   End
   Begin VB.TextBox txtSID 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   2190
      MaxLength       =   20
      TabIndex        =   3
      Top             =   1800
      Width           =   2625
   End
   Begin VB.TextBox txtUserName 
      Height          =   300
      Left            =   2190
      TabIndex        =   0
      Top             =   615
      Width           =   2625
   End
   Begin VB.TextBox txtUserPwd 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   2190
      MaxLength       =   20
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1005
      Width           =   2625
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Default         =   -1  'True
      Height          =   350
      Left            =   2145
      TabIndex        =   8
      Top             =   2880
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   3330
      TabIndex        =   7
      Top             =   2880
      Width           =   1100
   End
   Begin VB.CommandButton cmdTest 
      Caption         =   "测试连接(&T)"
      Height          =   350
      Left            =   105
      Style           =   1  'Graphical
      TabIndex        =   6
      TabStop         =   0   'False
      ToolTipText     =   "单击此处测试连接"
      Top             =   2865
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Height          =   120
      Left            =   -75
      TabIndex        =   5
      Top             =   2610
      Width           =   5310
   End
   Begin VB.Image imgFlag 
      Height          =   576
      Left            =   228
      Picture         =   "frmProcConfigure.frx":000C
      Top             =   612
      Width           =   576
   End
   Begin VB.Label lblPort 
      AutoSize        =   -1  'True
      Caption         =   "端口号"
      Height          =   180
      Left            =   1545
      TabIndex        =   14
      Top             =   2205
      Width           =   540
   End
   Begin VB.Label LblIP 
      AutoSize        =   -1  'True
      Caption         =   "数据服务器IP"
      Height          =   180
      Left            =   1005
      TabIndex        =   13
      Top             =   1470
      Width           =   1080
   End
   Begin VB.Label lblSID 
      AutoSize        =   -1  'True
      Caption         =   "数据库实例"
      Height          =   180
      Left            =   1185
      TabIndex        =   12
      Top             =   1860
      Width           =   900
   End
   Begin VB.Label lblMarks 
      BackStyle       =   0  'Transparent
      Caption         =   "配置页面展示所连接的数据库，请填写对应数据库的用户名、密码、IP以及实例名"
      Height          =   390
      Left            =   240
      TabIndex        =   11
      Top             =   150
      Width           =   4590
   End
   Begin VB.Label lblUserName 
      AutoSize        =   -1  'True
      Caption         =   "连接用户"
      Height          =   180
      Left            =   1365
      TabIndex        =   10
      Top             =   675
      Width           =   720
   End
   Begin VB.Label lblPwd 
      AutoSize        =   -1  'True
      Caption         =   "用户密码"
      Height          =   180
      Left            =   1365
      TabIndex        =   9
      Top             =   1065
      Width           =   720
   End
End
Attribute VB_Name = "frmProcConfigure"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mstrUserName As String
Private mstrUserPwd As String
Private mstrServerIP As String
Private mstrSID As String
Private mstrPort As String
Private mstrConnection As String
Private mobjMain As Object
Private mblnOk As Boolean
Private mblnLocal As Boolean
Private mclsCiph As clsCipher
Private mcnThis As ADODB.Connection

Public Event AfterConn(ByVal cnOracle As ADODB.Connection)

Public Function ShowConfigure(ByVal objMain As Object) As Boolean
    Set mobjMain = objMain
    Me.Show 1, mobjMain
    ShowConfigure = mblnOk
End Function

Private Function OraDataOpen(ByVal strServerIP As String, ByVal strSID As String, ByVal strUserName As String, ByVal strPassword As String, ByVal strPort As String, ByRef strErr As String) As Boolean
'功能： 打开指定的数据库
    Dim strServer As String, blnTransPassword As Boolean

    strServer = "(DESCRIPTION=(ADDRESS_LIST=(ADDRESS=(PROTOCOL=TCP)(HOST=" & strServerIP & ")(PORT = " & strPort & ")))(CONNECT_DATA=(SERVICE_NAME=" & strSID & ")))"
    blnTransPassword = Not (UCase(strUserName) = "SYS" Or UCase(strUserName) = "SYSTEM")
    Set mcnThis = gobjRegister.GetConnection(strServer, strUserName, strPassword, blnTransPassword, OraOLEDB, strErr, False)
    If mcnThis.State = adStateClosed Then
        OraDataOpen = False
        Exit Function
    End If
    
    OraDataOpen = True
End Function

Private Sub cmdCancel_Click()
    mblnOk = False
    Unload Me
End Sub

Private Sub cmdOK_Click()

    Dim rsSQL As ADODB.Recordset
    Dim clsCiph As New clsCipher
    Dim lngRow As Long

    On Error GoTo errHand
    
    If TestConnect = False Then Exit Sub
    
    Call gclsBase.SQLRecord(rsSQL)

    mstrConnection = mstrUserName & "/" & mstrUserPwd & "/" & mstrSID & "/" & mstrServerIP & "/" & mstrPort
    mstrConnection = clsCiph.Cipher("zLw09OewKKO1`;owEWO-=,./w[]wwqq3##=``44314325", mstrConnection)
    
    '对字符中&符号的处理
    mstrConnection = Replace(mstrConnection, "&", "' || chr(38) || '")
    gstrSQL = "Zl_Zlprocedureconnect_Update('连接配置','" & mstrConnection & "')"
    Call gclsBase.SQLRecordAdd(rsSQL, gstrSQL)
    
    If SQLRecordExecute(rsSQL) Then
        mblnOk = True
        RaiseEvent AfterConn(mcnThis)
        Unload Me
    End If
    Set clsCiph = Nothing
    Exit Sub
errHand:
    MsgBox err.Description, vbCritical, Me.Caption
End Sub

Private Function TestConnect() As Boolean

    Dim strUserName As String
    Dim strServerIP As String
    Dim strPassword As String
    Dim strSID As String
    Dim strPort As String
    Dim strNote As String
    Dim strPwdTmp As String
    
    On Error GoTo InputError
    
    '------检验用户是否oracle合法用户----------------
    strUserName = Trim(txtUserName.Text)
    strPassword = Trim(txtUserPwd.Text)
    strServerIP = Trim(txtServerIP.Text)
    strSID = Trim(txtSID.Text)
    strPort = Trim(txtPort.Text)
    
    '有效字符串效验
    If Len(Trim(txtUserName.Text)) = 0 Then
        strNote = "请输入用户名"
        txtUserName.SetFocus
        GoTo InputError
    End If
    
    If Len(Trim(txtUserPwd.Text)) = 0 Then
        strNote = "请输入密码"
        txtUserName.SetFocus
        GoTo InputError
    End If
    
    If Len(Trim(txtSID.Text)) = 0 Then
        strNote = "请输入数据库实例"
        txtSID.SetFocus
        GoTo InputError
    End If
    
    If Len(Trim(txtServerIP.Text)) = 0 Then
        strNote = "请输入数据库IP"
        txtServerIP.SetFocus
        GoTo InputError
    End If
    
    If Len(Trim(txtPort.Text)) = 0 Then
        strNote = "请输入端口号"
        txtServerIP.SetFocus
        GoTo InputError
    End If
    
    If Len(strUserName) <> 1 Then
        If Mid(strUserName, 1, 1) = "/" Or Mid(strUserName, 1, 1) = "@" Or Mid(strUserName, Len(strUserName) - 1, 1) = "/" Or Mid(strUserName, Len(strUserName) - 1, 1) = "@" Then
            txtUserName.SetFocus
            strNote = "用户名错误"
            GoTo InputError
        End If
    End If
    If Trim(strPassword) <> "" And Len(strPassword) <> 1 Then
        If Mid(strPassword, Len(strPassword) - 1, 1) = "/" Or Mid(strPassword, Len(strPassword) - 1, 1) = "@" Or Mid(strPassword, 1, 1) = "/" Or Mid(strPassword, 1, 1) = "@" Then
            txtUserPwd.SetFocus
            strNote = "密码错误"
            GoTo InputError
        End If
    End If
    
    strPwdTmp = strPassword
    strUserName = UCase(strUserName)
    
    If OraDataOpen(strServerIP, strSID, strUserName, strPassword, strPort, strNote) Then
        mstrUserName = strUserName
        mstrUserPwd = strPwdTmp
        mstrServerIP = strServerIP
        mstrSID = strSID
        mstrPort = strPort
        TestConnect = True
    Else
        GoTo InputError
    End If
    
    Exit Function
InputError:
    If strNote <> "" Then
        MsgBox strNote, vbInformation, gstrSysName
    End If
End Function

Private Sub cmdTest_Click()
    If TestConnect = True Then
        MsgBox "测试连接成功，可以正常访问该数据库。", vbInformation + vbOKOnly, gstrSysName
    End If
End Sub

Private Sub Form_Activate()

    On Error GoTo errHand
    Dim strConnection As String
    Dim strCon() As String
    
    Set mclsCiph = New clsCipher
    strConnection = gclsBase.GetOraConn("连接配置")
    If strConnection <> "" Then
        strConnection = mclsCiph.Decipher("zLw09OewKKO1`;owEWO-=,./w[]wwqq3##=``44314325", strConnection)
        strCon = Split(strConnection, "/")
        txtUserName.Text = strCon(0)
        txtUserPwd.Text = strCon(1)
        txtSID.Text = strCon(2)
        txtServerIP.Text = strCon(3)
        txtPort.Text = strCon(4)
    End If
    Set mclsCiph = Nothing
    txtUserPwd.SetFocus
    Exit Sub
errHand:
    MsgBox err.Description, vbCritical, Me.Caption
End Sub

Private Sub Form_Load()
    txtUserName.Text = GetSetting(appName:="ZLSOFT", Section:="注册信息\登陆信息", Key:="USER", Default:="")
    txtPort.Text = "1521"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not (mclsCiph Is Nothing) Then
        Set mclsCiph = Nothing
    End If
End Sub

Private Sub txtPort_GotFocus()
    gclsBase.TxtSelAll txtPort
End Sub

Private Sub txtServerIP_GotFocus()
    gclsBase.TxtSelAll txtServerIP
End Sub

Private Sub txtSID_GotFocus()
    gclsBase.TxtSelAll txtSID
End Sub

Private Sub txtUserName_GotFocus()
    gclsBase.TxtSelAll txtUserName
End Sub

Private Sub txtUserPwd_GotFocus()
    gclsBase.TxtSelAll txtUserPwd
End Sub

