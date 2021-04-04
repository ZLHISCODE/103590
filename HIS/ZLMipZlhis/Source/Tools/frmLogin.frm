VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "登录"
   ClientHeight    =   2265
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5715
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2265
   ScaleWidth      =   5715
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton CMD确认 
      Caption         =   "确定(&O)"
      Default         =   -1  'True
      Height          =   350
      Left            =   4470
      TabIndex        =   2
      Top             =   195
      Width           =   1100
   End
   Begin VB.CommandButton CMD放弃 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   4470
      TabIndex        =   1
      Top             =   720
      Width           =   1100
   End
   Begin VB.Frame Frame2 
      Caption         =   "ZLHIS产品"
      Height          =   1710
      Left            =   105
      TabIndex        =   0
      Top             =   165
      Width           =   4245
      Begin VB.TextBox txt用户 
         Height          =   300
         Left            =   855
         TabIndex        =   5
         Text            =   "ZLHIS"
         Top             =   315
         Width           =   2850
      End
      Begin VB.TextBox TXT密码 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   855
         MaxLength       =   20
         PasswordChar    =   "*"
         TabIndex        =   4
         Text            =   "aqa"
         Top             =   750
         Width           =   2850
      End
      Begin VB.ComboBox cmb数据库 
         Height          =   300
         Left            =   855
         Sorted          =   -1  'True
         TabIndex        =   3
         Text            =   "2.133.ORCL"
         Top             =   1170
         Width           =   2850
      End
      Begin VB.Label Lbl用户名 
         AutoSize        =   -1  'True
         Caption         =   "用户名"
         Height          =   180
         Left            =   135
         TabIndex        =   8
         Top             =   375
         Width           =   540
      End
      Begin VB.Label Lbl口令 
         AutoSize        =   -1  'True
         Caption         =   "密码"
         Height          =   180
         Left            =   135
         TabIndex        =   7
         Top             =   810
         Width           =   360
      End
      Begin VB.Label Lbl服务器 
         AutoSize        =   -1  'True
         Caption         =   "服务器"
         Height          =   180
         Left            =   135
         TabIndex        =   6
         Top             =   1230
         Width           =   540
      End
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mblnOK As Boolean

Public Function ShowLogin() As Boolean
    mblnOK = False
    
    Me.Show 1
    ShowLogin = mblnOK
End Function

Private Sub CMD放弃_Click()
    Unload Me
End Sub

Private Sub CMD确认_Click()
    Dim blnTransPassword As Boolean
    
    blnTransPassword = Not (UCase(txt用户.Text) = "SYS" Or UCase(txt用户.Text) = "SYSTEM")
    Set gcnOracle = gobjRegister.GetConnection(cmb数据库.Text, txt用户.Text, TXT密码.Text, blnTransPassword)
    If gcnOracle.State = adStateClosed Then
        TXT密码.Text = ""
        If TXT密码.Enabled Then TXT密码.SetFocus
        Exit Sub
    Else
        gstrDbUser = UCase(txt用户.Text)
    End If

    mblnOK = True
    Unload Me
End Sub

Private Sub Form_Load()
    txt用户.Text = GetSetting(App.ProductName, "登录", "用户", "")
    cmb数据库.Text = GetSetting(App.ProductName, "登录", "服务器", "")
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveSetting(App.ProductName, "登录", "用户", txt用户.Text)
    Call SaveSetting(App.ProductName, "登录", "服务器", cmb数据库.Text)
    
End Sub
