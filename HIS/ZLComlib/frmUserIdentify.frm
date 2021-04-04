VERSION 5.00
Begin VB.Form frmUserIdentify 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "用户验证"
   ClientHeight    =   2040
   ClientLeft      =   30
   ClientTop       =   330
   ClientWidth     =   4170
   Icon            =   "frmUserIdentify.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2040
   ScaleWidth      =   4170
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.Frame Frame1 
      Height          =   120
      Left            =   -360
      TabIndex        =   6
      Top             =   1335
      Width           =   5025
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   2865
      TabIndex        =   3
      Top             =   1590
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   1755
      TabIndex        =   2
      Top             =   1590
      Width           =   1100
   End
   Begin VB.TextBox txtPass 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   1950
      MaxLength       =   20
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   960
      Width           =   1920
   End
   Begin VB.TextBox txtUser 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   1950
      TabIndex        =   0
      Top             =   555
      Width           =   1920
   End
   Begin VB.Label lblNote 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "身份验证，请输入用户名与密码"
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   1335
      TabIndex        =   7
      Top             =   105
      Width           =   2520
   End
   Begin VB.Image imgFlag 
      Height          =   720
      Left            =   315
      Picture         =   "frmUserIdentify.frx":000C
      Top             =   240
      Width           =   720
   End
   Begin VB.Label lblPass 
      AutoSize        =   -1  'True
      Caption         =   "密码"
      Height          =   180
      Left            =   1500
      TabIndex        =   5
      Top             =   1020
      Width           =   360
   End
   Begin VB.Label lblUser 
      AutoSize        =   -1  'True
      Caption         =   "用户名"
      Height          =   180
      Left            =   1320
      TabIndex        =   4
      Top             =   615
      Width           =   540
   End
End
Attribute VB_Name = "frmUserIdentify"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrNote As String
Private mlngSys As Long
Private mlngProgID As Long
Private mstrFunc As String

Private mcnNew As ADODB.Connection
Private mcnNewOLEDB As ADODB.Connection
Private mstrServer As String
Private mstrUserName As String
Private mblnOK As Boolean
Private mblnDefaultPreUser As Boolean
Private mblnDBUser As Boolean
Private mstrDBUser  As String

Public Function ShowMe(frmParent As Object, ByVal strNote As String, ByVal lngSys As Long, ByVal lngProgId As Long, ByVal strFunc As String, Optional cnNew As ADODB.Connection, Optional ByVal blnDefaultPreUser As Boolean, Optional ByVal blnDBUser As Boolean, Optional ByRef strDBUser As String, Optional cnNewOLEDB As ADODB.Connection) As String
'参数：strNote=提示信息(简短)
'      lngProgID=程序序号
'      strFunc=授权功能,blnDBUser为True时为指定用户
'      cnNew=要返回的连接,blnDBUse=false时，,必须传入非Nothing的对象,并且需要由调用程序关闭连接；如果是当前登录用户,返回Nothing
'            blnDBUse=true时，传入空对象，返回打开的连接对象
'      blnDefaultPreUser-缺省显示上次登录人
'      blnDBUser=用数据库用户直接验证登录，并返回该用户创建的连接与输入的用户名，此时参数lngProgId，strFunc，blnDefaultPreUser
'      strDBUser=返回输入的输入的数据库用户
'      cnNewOLEDB=需要获取的OLEDB连接，和CNNEW都是同一用户，但是连接不同。当该参数不是Nothing时，才返回 OLEDB连接，否则不返回
'返回：成功返回人员姓名
'      strDBUser=输入的数据库用户
    mstrNote = strNote
    mlngSys = lngSys
    mlngProgID = lngProgId
    mblnDefaultPreUser = blnDefaultPreUser
    mblnDBUser = blnDBUser
    mstrDBUser = strDBUser
    mstrFunc = ""
    mstrUserName = ""
    Set mcnNewOLEDB = cnNewOLEDB
    If mblnDBUser Then
        mstrUserName = strFunc
    Else
        mstrFunc = strFunc
    End If
    
    Me.Show 1, frmParent
    If mblnOK Then
        ShowMe = mstrUserName
        If blnDBUser Then
            If Not mcnNew Is Nothing Then
                Set cnNew = mcnNew
                Set cnNewOLEDB = mcnNewOLEDB
            End If
        Else
            If Not cnNew Is Nothing Then
                Set cnNew = mcnNew
                Set cnNewOLEDB = mcnNewOLEDB
            ElseIf Not mcnNew Is Nothing Then
                mcnNew.Close
                Set mcnNew = Nothing
                If Not mcnNewOLEDB Is Nothing Then mcnNewOLEDB.Close
                Set mcnNewOLEDB = Nothing
            End If
        End If
        strDBUser = mstrDBUser
    Else
        Set cnNew = Nothing
        Set cnNewOLEDB = Nothing
    End If
End Function

Private Sub cmdOK_Click()
    Dim strUser As String
    Dim strPass As String
    
    strUser = Trim(txtUser.Text)
    strPass = Trim(txtPass.Text)
    
    '有效字符串效验
    If strUser = "" Then
        MsgBox "请输入用户名。", vbInformation, gstrSysName
        txtUser.SetFocus: Exit Sub
    End If
    If InStr(strUser, "/") > 0 Or InStr(strUser, "@") > 0 Then
        MsgBox "输入了无效的用户名，请重新输入。", vbInformation, gstrSysName
        txtUser.SetFocus: Exit Sub
    End If
    If strPass = "" Then
        MsgBox "请输入密码。", vbInformation, gstrSysName
        txtPass.SetFocus: Exit Sub
    End If
    If InStr(strPass, "/") > 0 Or InStr(strPass, "@") > 0 Then
        MsgBox "输入了无效的密码，请重新输入。", vbInformation, gstrSysName
        txtPass.Text = "": txtPass.SetFocus: Exit Sub
    End If

    If Not OpenOracle(strUser, strPass) Then Exit Sub
    mstrDBUser = UCase(strUser)
    Call SaveSetting("ZLSOFT", "公共模块\" & App.ProductName, "IdentifyUser", txtUser.Text)
    mblnOK = True
    Unload Me
End Sub

Private Sub cmdCancel_Click()
    mstrUserName = ""
    Unload Me
End Sub

Private Sub Form_Activate()
    If Trim(txtUser.Text) <> "" Then txtPass.SetFocus
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Then
        KeyAscii = 0
    ElseIf KeyAscii = 13 Then
        KeyAscii = 0
        If Me.ActiveControl Is txtPass Then
            Call cmdOK_Click
        Else
            Call gobjComLib.zlCommFun.PressKey(vbKeyTab)
        End If
    End If
End Sub

Private Sub Form_Load()
    mblnOK = False
   
    Set mcnNew = Nothing
    mstrServer = GetSetting("ZLSOFT", "注册信息\登陆信息", "Server", "")
    
    If mblnDBUser Then
        txtUser.Text = mstrUserName
        txtUser.Enabled = False
    Else
        If mblnDefaultPreUser Then
            txtUser.Text = GetSetting("ZLSOFT", "公共模块\" & App.ProductName, "IdentifyUser", "")
        End If
    End If
    
    If mstrNote <> "" Then lblNote.Caption = mstrNote
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mblnOK = False Then mstrUserName = ""
End Sub

Private Sub txtUser_GotFocus()
    Call gobjComLib.zlControl.TxtSelAll(txtUser)
End Sub

Private Sub txtPass_GotFocus()
    Call gobjComLib.zlControl.TxtSelAll(txtPass)
End Sub

Private Sub SetEnabled(ByVal blnEnabled As Boolean)
    cmdCancel.Enabled = blnEnabled
    cmdOK.Enabled = blnEnabled
    Screen.MousePointer = IIf(Not blnEnabled, 11, 0)
End Sub

Private Function OpenOracle(ByVal strUser As String, ByVal strPass As String) As Boolean
'功能：验证用户,并返回用户名和连接
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    Dim strError As String
    Dim strUserName As String
    
    Call SetEnabled(False)
    strUser = UCase(strUser)
    
    On Error GoTo errh
    
    '检查用户名
    strSQL = "Select UserName From All_Users Where UserName=[1]"
    Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strUser)
    If rsTmp.EOF Then
        MsgBox "该用户不存在。", vbInformation, gstrSysName
        Call SetEnabled(True)
        txtPass.Text = "": txtUser.SetFocus
        Exit Function
    End If
    
    '检查连接
    Set mcnNew = gobjRegister.GetConnection(mstrServer, strUser, strPass, Not mblnDBUser, , , False)
    If mcnNew.State = adStateClosed Then
        Call SetEnabled(True)
        txtPass.Text = "": txtPass.SetFocus
        Set mcnNew = Nothing: Exit Function
    End If
    If Not mcnNewOLEDB Is Nothing Then
        Set mcnNewOLEDB = gobjRegister.GetConnection(mstrServer, strUser, strPass, Not mblnDBUser, OraOLEDB, , False)
    End If
    If mblnDBUser Then
        mstrUserName = strUser
    Else
        '检查上机用户
        strSQL = "Select B.姓名 From 上机人员表 A,人员表 B Where (B.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or B.撤档时间 Is Null) And A.人员ID=B.ID And A.用户名=[1]"
        Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strUser)
        If rsTmp.EOF Then
            MsgBox "该用户未设置对应的人员信息。", vbInformation, gstrSysName
            Call SetEnabled(True)
            txtPass.Text = "": txtUser.SetFocus
            Exit Function
        End If
        strUserName = rsTmp!姓名
        
        '检查权限
        If mstrFunc <> "" Then
            If gobjComLib.SystemOwner(mlngSys) <> strUser Then
                strSQL = _
                    " Select 1 From zlUserRoles A,zlRoleGrant B " & _
                    " Where A.角色=B.角色 And B.系统=[1] And B.序号=[2] And B.功能=[3] And A.用户 = [4]"
                Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngSys, mlngProgID, mstrFunc, strUser)
                If rsTmp.EOF Then
                    MsgBox "该用户没有权限进行操作。", vbInformation, gstrSysName
                    Call SetEnabled(True)
                    txtPass.Text = "": txtUser.SetFocus
                    Exit Function
                End If
            End If
        End If
        
        '如果是当前用户则不需要使用单独的连接
        If strUser = UCase(gstrDBUser) Then
            mcnNew.Close: Set mcnNew = Nothing
        End If
        mstrUserName = strUserName
    End If
    
    Call SetEnabled(True)
    OpenOracle = True
    Exit Function
errh:
    If gobjComLib.ErrCenter() = 1 Then Resume
    Call gobjComLib.SaveErrLog
End Function

