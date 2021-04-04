VERSION 5.00
Begin VB.Form frmInterfaceUser 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ZLInterface用户创建"
   ClientHeight    =   2205
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4170
   Icon            =   "frmInterfaceUser.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2205
   ScaleWidth      =   4170
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.TextBox txtUser 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   1560
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   1
      Text            =   "ZLInterface"
      Top             =   120
      Width           =   2295
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   1440
      TabIndex        =   6
      Top             =   1680
      Width           =   1230
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   2760
      TabIndex        =   7
      Top             =   1680
      Width           =   1230
   End
   Begin VB.TextBox txtComfirmPwd 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   1560
      MaxLength       =   15
      PasswordChar    =   "*"
      TabIndex        =   5
      Top             =   1200
      Width           =   2295
   End
   Begin VB.TextBox txtNewPWD 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   1560
      MaxLength       =   15
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   640
      Width           =   2295
   End
   Begin VB.Image imgFlag 
      Height          =   720
      Left            =   120
      Picture         =   "frmInterfaceUser.frx":6633E
      Top             =   120
      Width           =   720
   End
   Begin VB.Label lblUser 
      AutoSize        =   -1  'True
      Caption         =   "用户"
      Height          =   180
      Left            =   1080
      TabIndex        =   0
      Top             =   180
      Width           =   360
   End
   Begin VB.Label lblNewPwd 
      AutoSize        =   -1  'True
      Caption         =   "密码"
      Height          =   180
      Left            =   1080
      TabIndex        =   2
      Top             =   700
      Width           =   360
   End
   Begin VB.Label lblComfirmPwd 
      AutoSize        =   -1  'True
      Caption         =   "密码验证"
      Height          =   180
      Left            =   720
      TabIndex        =   4
      Top             =   1220
      Width           =   720
   End
End
Attribute VB_Name = "frmInterfaceUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrOperate     As String '修改密码，创建用户
Private mblnOK          As Boolean
'===========================================================================
'==公共接口
'===========================================================================
Public Function ShowMe(ByVal strOperate As String) As Boolean
    mstrOperate = strOperate
    mblnOK = False
    Me.Show vbModal, frmMDIMain
    ShowMe = mblnOK
End Function

'===========================================================================
'==事件
'===========================================================================
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim strError  As String
    
    On Error GoTo errH
    If Trim(txtNewPWD.Text) = "" Then
        MsgBox "请输入密码！", vbInformation, gstrSysName
        txtNewPWD.SetFocus
        Exit Sub
    End If
    
    If Len(Trim(txtNewPWD.Text)) < 8 Then
        MsgBox "请输入至少8位的密码！", vbInformation, gstrSysName
        txtNewPWD.SetFocus
        Exit Sub
    End If
    
    If Trim(txtComfirmPwd.Text) = "" Then
        MsgBox "请输入密码验证！", vbInformation, gstrSysName
        txtComfirmPwd.SetFocus
        Exit Sub
    End If
    If txtNewPWD.Text <> txtComfirmPwd.Text Then
        MsgBox "两次输入的密码不一致，请重新输入！", vbInformation, gstrSysName
        txtComfirmPwd.SetFocus
        Exit Sub
    End If
    If Not RepairGeneralAccount(gcnOracle, "ZLINTERFACE", Trim(txtNewPWD.Text), strError) Then
        MsgBox mstrOperate & "失败。信息：" & strError, vbInformation, gstrSysName
    End If
    mblnOK = True
    Unload Me
    Exit Sub
errH:
    MsgBox mstrOperate & "失败。信息：" & err.Description, vbInformation, gstrSysName
    err.Clear
End Sub

Private Sub Form_Activate()
    txtNewPWD.SetFocus
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Then
        KeyAscii = 0
    ElseIf KeyAscii = vbKeyReturn Then
        KeyAscii = 0: PressKey vbKeyTab
    End If
End Sub

Private Sub Form_Load()
    If mstrOperate = "创建用户" Then
        Me.Caption = "ZLInterface用户创建"
    Else
        Me.Caption = "ZLInterface用户密码重置"
    End If
    
    HookDefend txtNewPWD.hwnd
End Sub

