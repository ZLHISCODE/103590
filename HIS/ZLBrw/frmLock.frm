VERSION 5.00
Begin VB.Form frmLock 
   BackColor       =   &H80000004&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "解除窗口锁定"
   ClientHeight    =   1905
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4230
   Icon            =   "frmLock.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1905
   ScaleWidth      =   4230
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton cmdShutDown 
      Caption         =   "关闭导航台(&C)"
      Height          =   350
      Left            =   960
      TabIndex        =   4
      Top             =   1320
      Visible         =   0   'False
      Width           =   1530
   End
   Begin VB.TextBox txtPwd 
      Height          =   350
      IMEMode         =   3  'DISABLE
      Left            =   960
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   870
      Width           =   2940
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "解锁(&U)"
      Default         =   -1  'True
      Height          =   350
      Left            =   2800
      TabIndex        =   1
      Top             =   1320
      Width           =   1100
   End
   Begin VB.Label lblUser 
      AutoSize        =   -1  'True
      Caption         =   "当前操作员："
      Height          =   180
      Left            =   960
      TabIndex        =   3
      Top             =   240
      Width           =   1080
   End
   Begin VB.Image imgFlag 
      Height          =   720
      Left            =   120
      Picture         =   "frmLock.frx":6852
      Top             =   120
      Width           =   720
   End
   Begin VB.Label lblNote 
      AutoSize        =   -1  'True
      Caption         =   "请输入登录密码解锁"
      Height          =   180
      Left            =   960
      TabIndex        =   2
      Top             =   600
      Width           =   1620
   End
End
Attribute VB_Name = "frmLock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Private Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Private Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Private Declare Function DrawMenuBar Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function SetActiveWindow Lib "user32" (ByVal hwnd As Long) As Long

Private Const MF_BYPOSITION = &H400&
Private Const MF_REMOVE = &H1000&
Private mobjRegister As Object

Private Sub cmdOK_Click()
    Dim strError As String
    If txtPwd.Text = "" Then
        MsgBox "请输入登录密码！", vbInformation, gstrSysName
        Exit Sub
    End If
    On Error Resume Next
    If mobjRegister Is Nothing Then
        Set mobjRegister = CreateObject("zlRegister.clsRegister")
        If mobjRegister Is Nothing Then
            If Err.Number <> 0 Then Err.Clear
            MsgBox "zlRegister对象创建失败！信息：" & Err.Description, vbInformation, gstrSysName
            Exit Sub
        End If
    End If
    If Err.Number <> 0 Then Err.Clear
    If mobjRegister.LoginValidate(gobjRelogin.ServerName, gobjRelogin.InputUser, Trim(txtPwd.Text), strError) Then
        '隐藏界面
        Call LockProg(False)
        Unload Me
    Else
        MsgBox "解锁失败！信息：" & strError, vbInformation, gstrSysName
        txtPwd.Text = ""
        Call txtPwd.SetFocus
    End If
End Sub

Private Sub cmdShutDown_Click()
    Call LockProg(False)
    Unload Me
    Unload frmBrower
End Sub

Private Sub Form_Activate()
'    Call SetActiveWindow(Me.hwnd)
    Call txtPwd.SetFocus
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Then
        KeyAscii = 0
    End If
End Sub

Private Sub Form_Load()
    Me.Caption = gstrUserName & "-解除锁定"
    lblUser.Caption = "当前操作员：" & gstrUserName
    
    If gblnShutDown Then
        Me.Width = 4320
        txtPwd.Width = 2940
        cmdOK.Left = 2800
        cmdShutDown.Visible = True
    Else
        Me.Width = 3816
        txtPwd.Width = 2412
        cmdOK.Left = 2280
        cmdShutDown.Visible = False
    End If
    Call DisableX
    Call SetActiveWindow(Me.hwnd)
'    Call txtPwd.SetFocus
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Cancel = gblnLock
End Sub

Private Sub txtPwd_GotFocus()
    zlControl.TxtSelAll txtPwd
End Sub

Private Sub DisableX()
     Dim hMenu As Long
     Dim nCount As Long
     hMenu = GetSystemMenu(Me.hwnd, 0)
     nCount = GetMenuItemCount(hMenu)
     Call RemoveMenu(hMenu, nCount - 1, MF_REMOVE Or MF_BYPOSITION)
     Call RemoveMenu(hMenu, nCount - 2, MF_REMOVE Or MF_BYPOSITION)
End Sub
