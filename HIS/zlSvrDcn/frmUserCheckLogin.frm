VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmUserCheckLogin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "用户验证"
   ClientHeight    =   2460
   ClientLeft      =   30
   ClientTop       =   330
   ClientWidth     =   4755
   Icon            =   "frmUserCheckLogin.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2460
   ScaleWidth      =   4755
   StartUpPosition =   2  '屏幕中心
   Begin VB.PictureBox pctServer 
      BorderStyle     =   0  'None
      Height          =   2415
      Left            =   0
      ScaleHeight     =   2415
      ScaleWidth      =   4815
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   0
      Width           =   4815
      Begin VB.TextBox txtIp 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   225
         IMEMode         =   3  'DISABLE
         Index           =   0
         Left            =   1680
         MaxLength       =   3
         TabIndex        =   4
         Tag             =   "IP地址"
         Top             =   1005
         Width           =   450
      End
      Begin VB.CommandButton cmdSerCancel 
         Cancel          =   -1  'True
         Caption         =   "取消(&C)"
         Height          =   345
         Left            =   3405
         TabIndex        =   23
         Top             =   2040
         Width           =   1095
      End
      Begin VB.CommandButton cmdSerOK 
         Caption         =   "确定(&O)"
         Default         =   -1  'True
         Height          =   345
         Left            =   2160
         TabIndex        =   13
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Frame fraSplit 
         Height          =   120
         Index           =   1
         Left            =   0
         TabIndex        =   22
         Top             =   1800
         Width           =   5000
      End
      Begin VB.TextBox txtSerPort 
         Height          =   285
         Left            =   1560
         MaxLength       =   5
         TabIndex        =   9
         Tag             =   "21"
         Top             =   1395
         Width           =   735
      End
      Begin VB.TextBox txtIp 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   225
         IMEMode         =   3  'DISABLE
         Index           =   2
         Left            =   3000
         MaxLength       =   3
         TabIndex        =   7
         Tag             =   "IP地址"
         Top             =   1005
         Width           =   435
      End
      Begin VB.TextBox txtIp 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   225
         IMEMode         =   3  'DISABLE
         Index           =   1
         Left            =   2280
         MaxLength       =   3
         TabIndex        =   6
         Tag             =   "IP地址"
         Top             =   1005
         Width           =   450
      End
      Begin VB.TextBox txtIp 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   225
         IMEMode         =   3  'DISABLE
         Index           =   3
         Left            =   3600
         MaxLength       =   3
         TabIndex        =   8
         Tag             =   "IP地址"
         Top             =   1005
         Width           =   435
      End
      Begin VB.TextBox txtSer 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   3240
         MaxLength       =   20
         TabIndex        =   11
         Top             =   1380
         Width           =   900
      End
      Begin MSComCtl2.UpDown udSerPort 
         Height          =   285
         Left            =   2280
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   1395
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   503
         _Version        =   393216
         Value           =   9999
         AutoBuddy       =   -1  'True
         BuddyControl    =   "txtSerPort"
         BuddyDispid     =   196614
         OrigLeft        =   2040
         OrigTop         =   795
         OrigRight       =   2295
         OrigBottom      =   1065
         Max             =   99999
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin VB.TextBox txtIpSet 
         Enabled         =   0   'False
         Height          =   300
         Left            =   1560
         TabIndex        =   19
         TabStop         =   0   'False
         Text            =   "      ．      ．     ．         "
         Top             =   960
         Width           =   2595
      End
      Begin VB.Label lblPort 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "端口号"
         Height          =   195
         Left            =   960
         TabIndex        =   21
         Top             =   1440
         Width           =   540
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "IP"
         Height          =   180
         Left            =   1350
         TabIndex        =   20
         Top             =   1005
         Width           =   180
      End
      Begin VB.Label lblSer 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "实例名"
         Height          =   180
         Left            =   2640
         TabIndex        =   18
         Top             =   1440
         Width           =   795
      End
      Begin VB.Label lblNote 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "    开启数据变动通知需要验证服务器IP、端口、实例名等信息，请输入后点击确认。"
         Height          =   360
         Index           =   1
         Left            =   1200
         TabIndex        =   17
         Top             =   360
         Width           =   3555
         WordWrap        =   -1  'True
      End
      Begin VB.Image imgFlag 
         Height          =   720
         Index           =   1
         Left            =   360
         Picture         =   "frmUserCheckLogin.frx":1CFA
         Top             =   120
         Width           =   720
      End
   End
   Begin VB.PictureBox pctZltools 
      BorderStyle     =   0  'None
      Height          =   2655
      Left            =   960
      ScaleHeight     =   2655
      ScaleWidth      =   4095
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   480
      Width           =   4095
      Begin VB.TextBox txtPWD 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1710
         MaxLength       =   20
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   1425
         Width           =   1500
      End
      Begin VB.TextBox txtUser 
         Enabled         =   0   'False
         ForeColor       =   &H80000011&
         Height          =   300
         Left            =   1710
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   24
         Text            =   "ZLTOOLS"
         Top             =   1020
         Width           =   1500
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "取消(&C)"
         Height          =   345
         Left            =   3525
         TabIndex        =   3
         Top             =   2040
         Width           =   1095
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "确定(&O)"
         Height          =   345
         Left            =   2280
         TabIndex        =   2
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Frame fraSplit 
         Height          =   120
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Top             =   1800
         Width           =   5000
      End
      Begin VB.Label lblNote 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "    调整自动作业需要ZLTOOLS登录验证，请输入密码后点击确认。"
         Height          =   360
         Index           =   0
         Left            =   1140
         TabIndex        =   15
         Top             =   360
         Width           =   3555
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblPWD 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "密码"
         Height          =   180
         Left            =   1275
         TabIndex        =   14
         Top             =   1485
         Width           =   360
      End
      Begin VB.Label lblUser 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "用户名"
         Height          =   180
         Left            =   1095
         TabIndex        =   12
         Top             =   1080
         Width           =   540
      End
      Begin VB.Image imgFlag 
         Height          =   720
         Index           =   0
         Left            =   210
         Picture         =   "frmUserCheckLogin.frx":2384
         Top             =   240
         Width           =   720
      End
   End
End
Attribute VB_Name = "frmUserCheckLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnLogin As Boolean
Private mblnServer As Boolean

Private mstrIp As String
Private mlngPort As String
Private mstrServer As String

Public Function GetZltoolsByLogin() As Boolean
    pctZltools.Visible = True
    pctZltools.Enabled = True
    pctServer.Visible = False
    pctServer.Enabled = False
    
    Me.Show
    GetZltoolsByLogin = mblnLogin
End Function

Public Function GetSerInfo(strIp As String, lngPort As String, strServer As String) As Boolean
    '功能:获取服务器信息
    Dim i As Integer
    pctZltools.Visible = False
    pctZltools.Enabled = False
    pctServer.Visible = True
    pctServer.Enabled = True
    
    If strIp <> "" Then
        For i = 0 To 3
            txtIp(i).Text = Split(strIp, ".")(i)
        Next
    End If
    
    txtSerPort.Text = Val(lngPort)
    txtSer.Text = strServer
    
    Me.Show 1
    
    strIp = mstrIp: lngPort = mlngPort: strServer = mstrServer
    GetSerInfo = mblnServer
End Function


Private Sub cmdCancel_Click()
    mblnLogin = False
    Unload Me
End Sub

Private Sub cmdOK_Click()

   On Error GoTo errH
    
    With gcnZltools
        .Provider = "OraOLEDB.Oracle"
        .Open "PLSQLRSet=1;Data Source=" & gstrServer, "ZLTOOLS", txtPWD.Text
      
        If .State = adStateOpen Then
            mblnLogin = True
        End If
    End With
    
    Exit Sub
errH:
    ErrCenter
End Sub

Private Sub cmdSerCancel_Click()
    mblnServer = False
    Unload Me
End Sub

Private Sub cmdSerOK_Click()
    
    '将信息保存至注册表
    mstrIp = txtIp(0).Text & "." & txtIp(1).Text & "." & txtIp(2).Text & "." & txtIp(3).Text
    mlngPort = Val(txtSerPort.Text)
    mstrServer = txtSer.Text
    
    SaveSetting "ZLSOFT\公共模块", "zlSvrNotice", "IP", mstrIp
    SaveSetting "ZLSOFT\公共模块", "zlSvrNotice", "PORT", mlngPort
    SaveSetting "ZLSOFT\公共模块", "zlSvrNotice", "Server", mstrServer
    
    mblnServer = True
    Unload Me
End Sub

Private Sub Form_Load()
    pctServer.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
    pctZltools.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
End Sub


Private Sub txtIp_GotFocus(Index As Integer)
    txtIp(Index).SelLength = Len(txtIp(Index).Text)
End Sub

Private Sub txtIp_KeyPress(Index As Integer, KeyAscii As Integer)
    If Chr(KeyAscii) = "." Then
        PressKey vbKeyTab
    End If
    
    OnlyIntCK KeyAscii
End Sub

Private Sub txtSer_KeyPress(KeyAscii As Integer)
    OnlyStrCK KeyAscii, 3, 22
End Sub

Private Sub txtSerPort_GotFocus()
    txtSerPort.SelStart = Len(txtSerPort.Text)
End Sub

Private Sub txtSer_GotFocus()
    txtSer.SelStart = Len(txtSer.Text)
End Sub

Private Sub txtSerPort_KeyPress(KeyAscii As Integer)
    OnlyIntCK KeyAscii
End Sub
