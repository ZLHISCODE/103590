VERSION 5.00
Begin VB.Form frmCheckUserIdentify 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "用户验证"
   ClientHeight    =   3270
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5070
   Icon            =   "frmCheckUserIdentify.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3270
   ScaleWidth      =   5070
   StartUpPosition =   1  '所有者中心
   Begin VB.TextBox txtRemarks 
      Height          =   840
      Left            =   1260
      MultiLine       =   -1  'True
      TabIndex        =   3
      ToolTipText     =   "该备注最多可输入128个汉字或256个字符"
      Top             =   1710
      Width           =   3495
   End
   Begin VB.Frame fraSplit 
      Height          =   120
      Left            =   0
      TabIndex        =   8
      Top             =   2565
      Width           =   6000
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   3645
      TabIndex        =   5
      Top             =   2820
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Default         =   -1  'True
      Height          =   350
      Left            =   2268
      TabIndex        =   4
      Top             =   2820
      Width           =   1100
   End
   Begin VB.TextBox txtUser 
      Height          =   300
      Left            =   1260
      MaxLength       =   30
      TabIndex        =   1
      Top             =   900
      Width           =   3495
   End
   Begin VB.TextBox txtPWD 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   1260
      MaxLength       =   20
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   1308
      Width           =   3495
   End
   Begin VB.Label lblRemarks 
      AutoSize        =   -1  'True
      Caption         =   "操作说明"
      Height          =   180
      Left            =   480
      TabIndex        =   9
      Top             =   1780
      Width           =   720
   End
   Begin VB.Image imgFlag 
      Height          =   720
      Left            =   210
      Picture         =   "frmCheckUserIdentify.frx":1CFA
      Top             =   120
      Width           =   720
   End
   Begin VB.Label lblPWD 
      AutoSize        =   -1  'True
      Caption         =   "口令"
      Height          =   180
      Left            =   840
      TabIndex        =   7
      Top             =   1368
      Width           =   360
   End
   Begin VB.Label lblUser 
      AutoSize        =   -1  'True
      Caption         =   "用户名"
      Height          =   180
      Left            =   660
      TabIndex        =   6
      Top             =   960
      Width           =   540
   End
   Begin VB.Label lblNote 
      AutoSize        =   -1  'True
      Caption         =   "    请输入应用系统的所有者用户进行验证。"
      Height          =   180
      Left            =   915
      TabIndex        =   0
      Top             =   390
      Width           =   3600
   End
End
Attribute VB_Name = "frmCheckUserIdentify"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnOK As Boolean
Private mstrRemarks As String
Private mblnNormal As Boolean   '表示当前登录用户为普通用户
Private mintTimes As Integer
Private mlngSysNo As Long

Public Function ShowMe(ByVal objParent As Object, ByVal lngSysNo As Long, Optional ByRef strRemarks As String) As Boolean
'功能：验证用户身份
'参数：
'      objParent = 父窗体
'      strUser=验证的用户
'      strRemarks=备注,用于执行重要操作验证身份时输入备注
'说明：验证当前用户是否为系统所有者，若是，则仅展示出操作说明输入框即可，若不是，则还需展示用户名、密码输入框，验证管理员身份
    mlngSysNo = lngSysNo
    mstrRemarks = strRemarks
    Me.Show vbModal, objParent
    strRemarks = mstrRemarks
    mstrRemarks = ""
    ShowMe = mblnOK
    mblnOK = False
End Function

Private Sub cmdCancel_Click()
    mblnOK = False
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim strNote As String, strRemarks As String
    Dim strUser As String, strPwd As String, strServer As String
    Dim intPos As Integer
    
    SetConState False
    mintTimes = mintTimes + 1
    
    strRemarks = Trim(txtRemarks.Text)
    If mblnNormal Then
        '------检验用户是否oracle合法用户----------------
        strUser = Trim(txtUser.Text)
        strPwd = Trim(txtPWD.Text)
        strServer = gobjRegister.GetServerName
    
        '有效字符串效验
        If Len(Trim(txtUser.Text)) = 0 Then
            strNote = "请输入用户名。"
            txtUser.SetFocus
            GoTo InputError
        End If
        
        If Len(strUser) <> 1 Then
            If Mid(strUser, 1, 1) = "/" Or Mid(strUser, 1, 1) = "@" Or Mid(strUser, Len(strUser) - 1, 1) = "/" Or Mid(strUser, Len(strUser) - 1, 1) = "@" Then
                txtUser.SetFocus
                strNote = "用户名错误。"
                Exit Sub
            End If
        End If
        If Trim(strPwd) <> "" And Len(strPwd) <> 1 Then
            If Mid(strPwd, Len(strPwd) - 1, 1) = "/" Or Mid(strPwd, Len(strPwd) - 1, 1) = "@" Or Mid(strPwd, 1, 1) = "/" Or Mid(strPwd, 1, 1) = "@" Then
                txtPWD.SetFocus
                strNote = "口令错误。"
                GoTo InputError
            End If
        End If
    
        '分离字符串
        intPos = InStr(1, strUser, "@", vbTextCompare)
        If intPos > 0 Then
            strServer = Mid(strUser, intPos + 1)
            strUser = Mid(strUser, 1, intPos - 1)
        End If
        
        intPos = InStr(1, strUser, "/", vbTextCompare)
        If intPos > 0 Then
            strPwd = Mid(strUser, intPos + 1)
            strUser = Mid(strUser, 1, intPos - 1)
        End If
        
        intPos = InStr(1, strPwd, "@", vbTextCompare)
        If intPos > 0 Then
            strServer = Mid(strPwd, intPos + 1)
            strPwd = Mid(strPwd, 1, intPos - 1)
        End If
        
        If Len(Trim(strPwd)) = 0 And mblnNormal Then
            strNote = "请输入密码"
            txtPWD.SetFocus
            GoTo InputError
        End If
        strUser = UCase(strUser)
        
        If Not OracleOpen(strServer, strUser, strPwd, strNote) Then
            txtPWD.Text = ""
            If txtPWD.Enabled Then txtPWD.SetFocus
            SetConState
            If strNote <> "" Then GoTo InputError
            Exit Sub
        End If
    End If
    
    If strRemarks = "" Then
        strNote = "请输入备注"
        txtRemarks.SetFocus
        GoTo InputError
    ElseIf strRemarks <> "" Then
        If gobjComLib.zlCommFun.StrIsValid(txtRemarks.Text, 256) = False Then
            txtRemarks.SetFocus
            SetConState
            Exit Sub
        End If
    End If
    mstrRemarks = strRemarks
    mblnOK = True
    Unload Me
    Exit Sub
InputError:
    If mintTimes >= 3 Then
        MsgBox "超过三次登录失败，系统将自动退出。", vbExclamation, gstrSysName
        cmdCancel_Click
    Else
        If strNote <> "" Then
            MsgBox strNote, vbExclamation, gstrSysName
        End If
        SetConState
        Exit Sub
    End If
End Sub

Private Sub SetConState(Optional ByVal BlnState As Boolean = True)
    cmdOK.Enabled = BlnState
    cmdCancel.Enabled = BlnState
End Sub

Private Function OracleOpen(ByVal strServerName As String, ByVal strUserName As String, ByVal strPassWord As String, ByRef strError As String) As Boolean
'功能： 打开指定的数据库
    Dim blnOwner As Boolean, blnTransPassword As Boolean
    Dim cnOracle As ADODB.Connection
    
    strError = ""
    If UCase(strUserName) = "SYS" Or UCase(strUserName) = "SYSTEM" Then
        blnTransPassword = False
    Else
        blnTransPassword = True
    End If
    Set cnOracle = gobjRegister.GetConnection(strServerName, strUserName, strPassWord, blnTransPassword, , strError)
    If cnOracle.State = adStateClosed Then
        OracleOpen = False
        Exit Function
    End If
    Set cnOracle = Nothing
    OracleOpen = True
End Function

Private Sub Form_Load()
    Call CheckUser
    If Not mblnNormal Then  '系统所有者用户登录
        Me.Height = 2865
        Me.Width = 4965
        txtUser.Visible = False
        txtPWD.Visible = False
        If mstrRemarks <> "" Then
            Me.Caption = mstrRemarks
        Else
            Me.Caption = "操作说明"
        End If
        lblNote.Caption = "请输入操作说明："
        imgFlag.Visible = False
        lblRemarks.Visible = False
        lblNote.Left = 150
        lblNote.Top = 100
        txtRemarks.Left = 150
        txtRemarks.Top = lblNote.Top + lblNote.Height + 100
        txtRemarks.Width = 4560
        txtRemarks.Height = 1440
        fraSplit.Top = txtRemarks.Top + txtRemarks.Height
        cmdOK.Top = fraSplit.Top + fraSplit.Height + 50
        cmdCancel.Left = 3590
        cmdCancel.Top = cmdOK.Top
    End If
End Sub

Private Sub CheckUser()
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim blnOwner As Boolean, blnDBA As Boolean

    On Error GoTo errH
    mblnNormal = False
    strSQL = "SELECT 1 FROM ZLSYSTEMS WHERE 所有者=USER"
    Set rsTemp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, "系统所有者判定")
    blnOwner = Not rsTemp.EOF
    
    strSQL = "SELECT 1 FROM SESSION_ROLES WHERE ROLE='DBA'"
    Set rsTemp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, "DBA判定")
    blnDBA = Not rsTemp.EOF
    
    If Not blnOwner And Not blnDBA Then
        mblnNormal = True
        strSQL = "Select 所有者 From zlSystems Where 编号 =[1]"
        Set rsTemp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, "获取系统所有者", mlngSysNo)
        If rsTemp.RecordCount > 0 Then
            txtUser.Text = rsTemp!所有者
            txtUser.Enabled = False
        End If
    End If
    Exit Sub
errH:
    MsgBox Err.Description, vbInformation, gstrSysName
End Sub

Private Sub txtRemarks_KeyPress(KeyAscii As Integer)
    If gobjComLib.zlCommFun.ActualLen(txtRemarks.Text) >= 256 And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub
