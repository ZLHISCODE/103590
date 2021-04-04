VERSION 5.00
Begin VB.Form frmModifyPassWord_重庆渝北 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "设置"
   ClientHeight    =   1995
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5175
   Icon            =   "frmModifyPassWord_重庆渝北.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1995
   ScaleWidth      =   5175
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame Fra密码 
      Caption         =   "更改密码"
      Height          =   1455
      Left            =   210
      TabIndex        =   2
      Top             =   270
      Width           =   3165
      Begin VB.TextBox TXT原密码 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1110
         MaxLength       =   10
         PasswordChar    =   "*"
         TabIndex        =   5
         Top             =   270
         Width           =   1590
      End
      Begin VB.TextBox TXT新密码 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1110
         MaxLength       =   10
         PasswordChar    =   "*"
         TabIndex        =   4
         Top             =   645
         Width           =   1590
      End
      Begin VB.TextBox TXT确认密码 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1110
         MaxLength       =   10
         PasswordChar    =   "*"
         TabIndex        =   3
         Top             =   1005
         Width           =   1590
      End
      Begin VB.Label Lbl密码验证 
         AutoSize        =   -1  'True
         Caption         =   "密码验证"
         Height          =   180
         Left            =   270
         TabIndex        =   8
         Top             =   1065
         Width           =   720
      End
      Begin VB.Label Lbl新密码 
         AutoSize        =   -1  'True
         Caption         =   "新密码"
         Height          =   180
         Left            =   450
         TabIndex        =   7
         Top             =   705
         Width           =   540
      End
      Begin VB.Label Lbl旧密码 
         AutoSize        =   -1  'True
         Caption         =   "旧密码"
         Height          =   180
         Left            =   450
         TabIndex        =   6
         Top             =   330
         Width           =   540
      End
   End
   Begin VB.CommandButton CMD放弃 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   3870
      TabIndex        =   1
      Top             =   705
      Width           =   1230
   End
   Begin VB.CommandButton CDM确认 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   3870
      TabIndex        =   0
      Top             =   255
      Width           =   1230
   End
End
Attribute VB_Name = "frmModifyPassWord_重庆渝北"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnChange As Boolean

Private Declare Function SetActiveWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Const GWL_EXSTYLE = (-20)
Private Const WinStyle = &H40000
Private Const SWP_NOSIZE = &H1
Private Const SWP_SHOWWINDOW = &H40
Private Const HWND_TOPMOST = -1
Private Sub CDM确认_Click()
    If Trim(TXT原密码) = "" Then
        MsgBox "请输入旧密码！", vbInformation, gstrSysName
        TXT原密码.SetFocus
        Exit Sub
    End If
    If Trim(TXT新密码) = "" Then
        MsgBox "请输入新密码！", vbInformation, gstrSysName
        TXT新密码.SetFocus
        Exit Sub
    End If
    If Trim(TXT确认密码) = "" Then
        MsgBox "请输入密码验证！", vbInformation, gstrSysName
        TXT确认密码.SetFocus
        Exit Sub
    End If
    If TXT新密码.Text <> TXT确认密码.Text Then
        MsgBox "新密码输入错误，请重新输入！", vbInformation, gstrSysName
        TXT新密码.SetFocus
        Exit Sub
    End If
    If 修改密码_重庆渝北(TXT原密码, TXT新密码) = True Then
        g病人身份_重庆渝北.密码 = Trim(TXT新密码.Text)
        mblnChange = True
        Unload Me
        Exit Sub
    End If
End Sub

Private Sub CMD放弃_Click()
    TXT确认密码 = ""
    TXT新密码 = ""
    TXT原密码 = ""
    mblnChange = False
    Unload Me
End Sub

Private Sub Form_Activate()
    Call SetWindowPos(Me.hwnd, HWND_TOPMOST, Me.Left / 15, Me.Top / 15, Me.Height / 15, Me.Width / 15, SWP_NOSIZE + SWP_SHOWWINDOW)
    If TXT原密码 = "" Then
        TXT原密码.SetFocus
    Else
        TXT新密码.SetFocus
    End If
End Sub

Private Sub TXT确认密码_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{Tab}", 1
End Sub

Private Sub txt新密码_GotFocus()
    GetFocus TXT新密码
End Sub

Private Sub TXT新密码_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{Tab}", 1
End Sub

Private Sub txt原密码_GotFocus()
    GetFocus TXT原密码
End Sub

Private Sub TXT确认密码_GotFocus()
    GetFocus TXT确认密码
End Sub

Private Sub GetFocus(ByVal TxtBox As TextBox)
    With TxtBox
        .SelStart = 0
        .SelLength = LenB(StrConv(.Text, vbFromUnicode))
    End With
End Sub

Private Sub txt原密码_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{Tab}", 1
End Sub

Public Function ShowEdit(ByVal frmMain As Object) As Boolean
    '入口
    mblnChange = False
    Me.Show , frmMain
    ShowEdit = mblnChange
End Function
