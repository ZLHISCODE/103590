VERSION 5.00
Begin VB.Form frm登录中心 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "登录中心"
   ClientHeight    =   2175
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4245
   Icon            =   "frm登录中心.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2175
   ScaleWidth      =   4245
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton cmdModify 
      Caption         =   "修改密码(M)"
      Height          =   375
      Left            =   90
      TabIndex        =   11
      Top             =   1650
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.Frame fra 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   975
      Index           =   26
      Left            =   1020
      TabIndex        =   13
      Top             =   270
      Visible         =   0   'False
      Width           =   2865
      Begin VB.TextBox TXT密码 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   26
         Left            =   1080
         MaxLength       =   20
         PasswordChar    =   "*"
         TabIndex        =   7
         Top             =   525
         Width           =   1620
      End
      Begin VB.TextBox txt用户 
         Height          =   300
         Index           =   26
         Left            =   1080
         TabIndex        =   5
         Top             =   90
         Width           =   1620
      End
      Begin VB.Label Lbl口令 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "口令(&P)"
         Height          =   180
         Index           =   26
         Left            =   330
         TabIndex        =   6
         Top             =   585
         Width           =   630
      End
      Begin VB.Label Lbl用户名 
         AutoSize        =   -1  'True
         Caption         =   "工号(&U)"
         Height          =   180
         Index           =   26
         Left            =   330
         TabIndex        =   4
         Top             =   150
         Width           =   630
      End
   End
   Begin VB.Frame Frame1 
      Height          =   30
      Left            =   -240
      TabIndex        =   8
      Top             =   1440
      Width           =   4785
   End
   Begin VB.CommandButton CDM确认 
      Caption         =   "确定(&O)"
      Default         =   -1  'True
      Height          =   350
      Left            =   1710
      TabIndex        =   9
      Top             =   1665
      Width           =   1100
   End
   Begin VB.CommandButton CMD放弃 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   2910
      TabIndex        =   10
      Top             =   1665
      Width           =   1100
   End
   Begin VB.Frame fra 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   975
      Index           =   43
      Left            =   1020
      TabIndex        =   12
      Top             =   270
      Visible         =   0   'False
      Width           =   2865
      Begin VB.TextBox txt用户 
         Height          =   300
         Index           =   43
         Left            =   1080
         TabIndex        =   1
         Top             =   90
         Width           =   1620
      End
      Begin VB.TextBox TXT密码 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   43
         Left            =   1080
         MaxLength       =   20
         PasswordChar    =   "*"
         TabIndex        =   3
         Top             =   525
         Width           =   1620
      End
      Begin VB.Label Lbl用户名 
         AutoSize        =   -1  'True
         Caption         =   "工号(&U)"
         Height          =   180
         Index           =   43
         Left            =   330
         TabIndex        =   0
         Top             =   150
         Width           =   630
      End
      Begin VB.Label Lbl口令 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "口令(&P)"
         Height          =   180
         Index           =   43
         Left            =   330
         TabIndex        =   2
         Top             =   585
         Width           =   630
      End
   End
   Begin VB.Image imgFlag 
      Height          =   720
      Left            =   180
      Picture         =   "frm登录中心.frx":000C
      Top             =   300
      Width           =   720
   End
End
Attribute VB_Name = "frm登录中心"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mint险类 As Integer
Private mbln修改密码 As Boolean
Private mblnLogin As Boolean
Private mstr新密码 As String

Public Function LoginCenter(ByVal int险类 As Integer, Optional ByVal bln修改密码 As Boolean = False) As Boolean
    '登录中心
    On Error Resume Next
    mblnLogin = False
    mint险类 = int险类
    mbln修改密码 = bln修改密码
    
    Me.Show 1
    LoginCenter = mblnLogin
End Function

Private Sub CDM确认_Click()
    If Trim(txt用户(mint险类).Text) = "" Then
        MsgBox "请输入操作员工号！", vbInformation, gstrSysName
        txt用户(mint险类).SetFocus
        Exit Sub
    End If
    
    Call Login
End Sub

Private Sub cmdModify_Click()
    With frm修改密码
        mstr新密码 = .ChangePassword(txt密码(mint险类).Text)
    End With
End Sub

Private Sub CMD放弃_Click()
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    fra(mint险类).Visible = True
    cmdModify.Visible = mbln修改密码
    
    '自动登录
    Call AutoLogin
End Sub

Private Sub AutoLogin()
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    
    Select Case mint险类
    Case TYPE_沈阳市, TYPE_乐山
        If mint险类 = TYPE_沈阳市 Then
            gstrSQL = "Select 序号,参数值 From 保险参数 Where 险类=[1] And 序号 In (4,5) Order by 序号"
        Else
            gstrSQL = "Select 序号,参数值 From 保险参数 Where 险类=[2] And 序号 In (2,3) Order by 序号"
        End If
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "读取登录参数", mint险类)
        If rsTemp.EOF Then Exit Sub
        
        Do While Not rsTemp.EOF
            If rsTemp.AbsolutePosition = 1 Then
                txt用户(mint险类) = Nvl(rsTemp!参数值)
            Else
                txt密码(mint险类) = Nvl(rsTemp!参数值)
            End If
            rsTemp.MoveNext
        Loop
    End Select
    
    If Trim(txt用户(mint险类)) <> "" And Trim(txt密码(mint险类)) <> "" Then Call Login
    Exit Sub
errHand:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub Login()
    On Error GoTo errHand
    
    Select Case mint险类
    Case TYPE_沈阳市
        If Not 调用接口_准备_沈阳市(Function_沈阳市.登录中心) Then Exit Sub
        '填写入口参数
        Call CZ_DataPut(glngInterface_沈阳市, 1, "staff_id", Trim(txt用户(mint险类).Text))
        Call CZ_DataPut(glngInterface_沈阳市, 1, "staff_pwd", txt密码(mint险类).Text)
        '运行
        If Not 调用接口_执行_沈阳市 Then Exit Sub
        '登录成功，保存操作员工号
        gCominfo_沈阳市.操作员工号 = Trim(txt用户(mint险类).Text)
    Case TYPE_乐山
        Dim strUserName As TStringOfChar, strPassWord As TStringOfChar
        strUserName.Data = Trim(txt用户(mint险类).Text)
        strPassWord.Data = txt密码(mint险类).Text
        gbytReturn_乐山 = LS_UserLogin(strUserName, strPassWord)
        If GetErrInfo_乐山 Then Exit Sub
    End Select
    
    If mstr新密码 <> "" Then
        '修改密码
        Select Case mint险类
        Case TYPE_乐山
            Dim strNewPwd As TStringOfChar
            strNewPwd.Data = mstr新密码
            gbytReturn_乐山 = LS_ChangePwd(strPassWord, strNewPwd)
            If gbytReturn_乐山 <> 0 Then MsgBox "密码修改失败！", vbInformation, gstrSysName
        End Select
    End If
    
    mblnLogin = True
    Unload Me
    Exit Sub
errHand:
    If ErrCenter = 1 Then Resume
End Sub
