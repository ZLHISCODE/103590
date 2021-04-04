VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmCheckUser 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "#"
   ClientHeight    =   2100
   ClientLeft      =   30
   ClientTop       =   330
   ClientWidth     =   4440
   Icon            =   "frmCheckUser.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2100
   ScaleWidth      =   4440
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.PictureBox picVisible 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   15
      Left            =   0
      ScaleHeight     =   15
      ScaleWidth      =   15
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   0
      Width           =   15
   End
   Begin MSComCtl2.DTPicker dtpDate 
      Height          =   300
      Left            =   1950
      TabIndex        =   5
      Top             =   1050
      Width           =   2115
      _ExtentX        =   3731
      _ExtentY        =   529
      _Version        =   393216
      CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
      Format          =   248446979
      CurrentDate     =   43074
   End
   Begin VB.Frame Frame1 
      Height          =   120
      Left            =   -360
      TabIndex        =   8
      Top             =   1380
      Width           =   5025
   End
   Begin VB.CommandButton CMD放弃 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   3015
      TabIndex        =   7
      Top             =   1650
      Width           =   1100
   End
   Begin VB.CommandButton CMD确认 
      Caption         =   "确定(&O)"
      Default         =   -1  'True
      Height          =   350
      Left            =   1905
      TabIndex        =   6
      Top             =   1650
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
      Width           =   2115
   End
   Begin VB.TextBox txt用户 
      Height          =   300
      Left            =   1950
      TabIndex        =   1
      Top             =   195
      Width           =   2115
   End
   Begin VB.Image ImgAudit 
      Height          =   810
      Left            =   315
      Picture         =   "frmCheckUser.frx":000C
      Stretch         =   -1  'True
      Top             =   210
      Width           =   720
   End
   Begin VB.Label lblDate 
      AutoSize        =   -1  'True
      Caption         =   "核对时间"
      Height          =   180
      Left            =   1125
      TabIndex        =   4
      Top             =   1095
      Width           =   720
   End
   Begin VB.Image imgFlag 
      Height          =   720
      Left            =   330
      Picture         =   "frmCheckUser.frx":0396
      Top             =   240
      Width           =   720
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
Attribute VB_Name = "frmCheckUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mintTimes As Integer
Dim mstrCaption As String
Dim mstrUser As String

'时间控件变量
Private mblnShowTime As Boolean
Private mstrShowTitle As String
Private mstrExeTitle As String
Private mstrExeTime As String
Private mstrCurDate As String
Private mstrFormat As String
Private mstrModeName As String

Public Function IsValidUser(ByVal strModeName As String, ByVal strTitle As String, Optional ByVal blnShowTime As Boolean = False, Optional ByVal strShowTitle As String = "", Optional ByVal strExeTitle As String = "", _
    Optional ByVal strExeTime As String, Optional ByVal strCurDate As String = "", Optional ByVal strFormat As String = "yyyy-MM-dd HH:mm") As String
    '返回数据格式：登录用户名;姓名，以及日期
    '参数：
    '         strModeName--调用模块名称
    '         blnShowTime---是否显示日期选择，该参数为TRUE时后面的参数才有效
    '         strShowTitle-日期标题；当需要限定输入的日期不能小于某个时间点，可传入strExeTime和strExeTitle。
    '         strCurDate-缺省显示的日期，为空则默认为当前时间。
    '         strFormat--日期显示的格式
    mstrUser = ""
    mintTimes = 1
    mstrModeName = strModeName
    mstrCaption = strTitle
    mblnShowTime = blnShowTime
    mstrShowTitle = strShowTitle
    mstrExeTitle = strExeTitle
    mstrExeTime = strExeTime
    mstrCurDate = strCurDate
    mstrFormat = strFormat
    If mstrFormat = "" Then mstrFormat = "yyyy-MM-dd HH:mm"
    Me.Show 1
    IsValidUser = mstrUser
End Function

Private Sub CMD确认_Click()
    Dim strSQL As String
    Dim strNote As String
    Dim strUserName As String
    Dim strServerName As String
    Dim strPassword As String
    Dim rsUser As New ADODB.Recordset
    Dim strCurDate As String
    On Error GoTo InputError
    
    '------检验用户是否oracle合法用户----------------
    strUserName = Trim(UCase(txt用户.Text))
    strPassword = Trim(TXT密码.Text)
    strServerName = GetSetting("ZLSOFT", "注册信息\登陆信息", "SERVER", "")
    
    '有效字符串效验
    If Len(strUserName) = 0 Then
        strNote = "请输入用户名"
        If txt用户.Enabled And txt用户.Visible Then txt用户.SetFocus
        GoTo InputError
    End If
    
    If Len(strPassword) = 0 Then
        strNote = "请输入密码"
        If TXT密码.Enabled And TXT密码.Visible Then TXT密码.SetFocus
        GoTo InputError
    End If
    
    If mblnShowTime = True Then
        If IsDate(mstrExeTime) Then
            If Format(dtpDate.Value, mstrFormat) < Format(mstrExeTime, mstrFormat) Then
                strNote = mstrShowTitle & "不能小于" & mstrExeTitle & "【" & Format(mstrExeTime, mstrFormat) & "】"
                If dtpDate.Enabled And dtpDate.Visible Then dtpDate.SetFocus
                GoTo InputError
            End If
        End If
        strCurDate = Format(gobjDatabase.Currentdate, mstrFormat)
        If Format(dtpDate.Value, mstrFormat) > strCurDate Then
            MsgBox mstrShowTitle & "不能大于当前时间【" & strCurDate & "】", vbInformation, gstrSysName
            If dtpDate.Enabled And dtpDate.Visible Then dtpDate.SetFocus
            GoTo InputError
        End If
        strCurDate = Format(dtpDate.Value, mstrFormat)
    End If
    
    SetConState False
    mintTimes = mintTimes + 1
     '用户登录验证
    If GetObjectRegister = False Then Exit Sub
    strServerName = gobjRegister.GetServerName
    If gobjRegister.LoginValidate(strServerName, strUserName, strPassword, strNote) = False Then
        TXT密码.Text = ""
        If TXT密码.Enabled Then TXT密码.SetFocus
        SetConState
        GoTo InputError
    End If
        
    '修改注册表
    Call SaveSetting("ZLSOFT", "公共模块\" & App.ProductName & "\" & mstrModeName & "\" & mstrCaption, "用户名", strUserName)
    strSQL = " Select A.姓名 From 人员表 A,上机人员表 B Where A.ID=B.人员ID And B.用户名=[1] "
    Set rsUser = gobjDatabase.OpenSQLRecord(strSQL, "提取登录用户姓名", strUserName)
    mstrUser = strUserName & "'" & rsUser!姓名
    If mblnShowTime = True Then mstrUser = mstrUser & "'" & strCurDate
    
    Unload Me
    Exit Sub
InputError:
    If mintTimes > 3 Then
        MsgBox "超过三次登录失败，程序退出！", vbExclamation, gstrSysName
        CMD放弃_Click
    Else
        If strNote <> "" Then
            MsgBox strNote, vbExclamation, gstrSysName
        End If
        SetConState
        Exit Sub
    End If

End Sub

Private Sub CMD放弃_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
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
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Me.ActiveControl.name = "dtpDate" And KeyCode = vbKeyReturn Then picVisible.SetFocus: Exit Sub
    If KeyCode = vbKeyReturn Then
        If Me.ActiveControl.name = "TXT密码" Then
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
    Dim lngTop As Long
    Me.Caption = mstrCaption
    txt用户.Text = GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\" & mstrModeName & "\" & mstrCaption, "用户名", "")
    
    If mblnShowTime = False Then
        lblDate.Visible = False
        dtpDate.Visible = False
        dtpDate.Enabled = False
        lngTop = TXT密码.Top + TXT密码.Height + 150
    Else
        lblDate.Visible = True
        dtpDate.Visible = True
        dtpDate.Enabled = True
        dtpDate.CustomFormat = mstrFormat
        lblDate.Caption = mstrShowTitle
        lngTop = dtpDate.Top + dtpDate.Height + 150
        If IsDate(mstrCurDate) Then
            dtpDate.Value = Format(mstrCurDate, mstrFormat)
        Else
            dtpDate.Value = Format(gobjDatabase.Currentdate, mstrFormat)
        End If
    End If
    Frame1.Top = lngTop
    CMD确认.Top = Frame1.Top + Frame1.Height + 150
    CMD放弃.Top = CMD确认.Top
    
    Me.Height = CMD放弃.Top + CMD放弃.Height + 535
    
    '屏蔽三方程序获取密码内容
    HookDefend TXT密码.hWnd
End Sub

Private Sub GetFocus(ByVal TxtBox As TextBox)
    With TxtBox
        .SelStart = 0
        .SelLength = LenB(StrConv(.Text, vbFromUnicode))
    End With
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Me.ActiveControl.name = "dtpDate" Then picVisible.SetFocus
End Sub

Private Sub txt用户_Change()
    CMD确认.Default = False
End Sub

Private Sub TXT密码_GotFocus()
    GetFocus TXT密码
End Sub

Private Sub SetConState(Optional ByVal BlnState As Boolean = True)
    CMD放弃.Enabled = BlnState
    CMD确认.Enabled = BlnState
End Sub
