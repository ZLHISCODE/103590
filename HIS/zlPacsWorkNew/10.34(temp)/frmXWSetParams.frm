VERSION 5.00
Begin VB.Form frmXWSetParams 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "PACS参数设置"
   ClientHeight    =   5535
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8850
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   12
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmXWSetParams.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5535
   ScaleWidth      =   8850
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.TextBox txtImageShare 
      Height          =   375
      Left            =   6495
      TabIndex        =   29
      Text            =   "DCMSHARE"
      Top             =   3540
      Width           =   2145
   End
   Begin VB.TextBox txtWebServerIP 
      Height          =   375
      Left            =   6120
      TabIndex        =   27
      Top             =   2955
      Width           =   2520
   End
   Begin VB.CommandButton Command1 
      Caption         =   "测试(&T)"
      Height          =   400
      Left            =   7575
      TabIndex        =   26
      Top             =   2400
      Width           =   1000
   End
   Begin VB.CheckBox chkLog 
      Caption         =   "记录日志"
      Height          =   255
      Left            =   4680
      TabIndex        =   25
      Top             =   4275
      Width           =   3975
   End
   Begin VB.Frame Frame4 
      Caption         =   "PACS 用户设置"
      Height          =   5175
      Left            =   240
      TabIndex        =   9
      Top             =   120
      Width           =   4335
      Begin VB.Frame Frame6 
         Caption         =   "光盘刻录"
         Height          =   1455
         Left            =   240
         TabIndex        =   20
         Top             =   3480
         Width           =   3975
         Begin VB.TextBox txtDVDBurnPswd 
            Height          =   375
            Left            =   1080
            TabIndex        =   22
            Top             =   840
            Width           =   2655
         End
         Begin VB.TextBox txtDVDBurnUser 
            Height          =   375
            Left            =   1080
            TabIndex        =   21
            Top             =   360
            Width           =   2655
         End
         Begin VB.Label Label10 
            Caption         =   "密码"
            Height          =   255
            Left            =   240
            TabIndex        =   24
            Top             =   900
            Width           =   855
         End
         Begin VB.Label Label9 
            Caption         =   "用户名"
            Height          =   255
            Left            =   240
            TabIndex        =   23
            Top             =   420
            Width           =   855
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "发送图像"
         Height          =   1455
         Left            =   240
         TabIndex        =   15
         Top             =   1920
         Width           =   3975
         Begin VB.TextBox txtSendImageUser 
            Height          =   375
            Left            =   1080
            TabIndex        =   17
            Top             =   360
            Width           =   2655
         End
         Begin VB.TextBox txtSendImagePswd 
            Height          =   375
            Left            =   1080
            TabIndex        =   16
            Top             =   840
            Width           =   2655
         End
         Begin VB.Label Label8 
            Caption         =   "用户名"
            Height          =   255
            Left            =   240
            TabIndex        =   19
            Top             =   420
            Width           =   855
         End
         Begin VB.Label Label7 
            Caption         =   "密码"
            Height          =   255
            Left            =   240
            TabIndex        =   18
            Top             =   900
            Width           =   855
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "删除图像"
         Height          =   1455
         Left            =   240
         TabIndex        =   10
         Top             =   360
         Width           =   3975
         Begin VB.TextBox txtDelImagePswd 
            Height          =   375
            Left            =   1080
            TabIndex        =   12
            Top             =   840
            Width           =   2655
         End
         Begin VB.TextBox txtDelImageUser 
            Height          =   375
            Left            =   1080
            TabIndex        =   11
            Top             =   360
            Width           =   2655
         End
         Begin VB.Label Label4 
            Caption         =   "密码"
            Height          =   255
            Left            =   240
            TabIndex        =   14
            Top             =   900
            Width           =   855
         End
         Begin VB.Label Label5 
            Caption         =   "用户名"
            Height          =   255
            Left            =   240
            TabIndex        =   13
            Top             =   420
            Width           =   855
         End
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "PACS 数据库服务器"
      Height          =   2055
      Left            =   4680
      TabIndex        =   2
      Top             =   240
      Width           =   3975
      Begin VB.TextBox txtDBServerPswd 
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   1080
         PasswordChar    =   "*"
         TabIndex        =   8
         Top             =   1440
         Width           =   2655
      End
      Begin VB.TextBox txtDBServerUser 
         Height          =   375
         Left            =   1080
         TabIndex        =   6
         Top             =   960
         Width           =   2655
      End
      Begin VB.TextBox txtDBServerIP 
         Height          =   375
         Left            =   1080
         TabIndex        =   4
         Top             =   480
         Width           =   2655
      End
      Begin VB.Label Label3 
         Caption         =   "密码"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   1500
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "用户名"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   1020
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "服务名"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   540
         Width           =   855
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(&C)"
      Height          =   400
      Left            =   7200
      TabIndex        =   1
      Top             =   4920
      Width           =   1000
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   400
      Left            =   5280
      TabIndex        =   0
      Top             =   4920
      Width           =   1000
   End
   Begin VB.Label Label11 
      Caption         =   "历史图像共享目录"
      Height          =   255
      Left            =   4680
      TabIndex        =   30
      Top             =   3600
      Width           =   1725
   End
   Begin VB.Label Label6 
      Caption         =   "WEB服务器IP"
      Height          =   255
      Left            =   4680
      TabIndex        =   28
      Top             =   3015
      Width           =   1380
   End
End
Attribute VB_Name = "frmXWSetParams"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Function zlShowMe(frmParent As Form) As Long
'------------------------------------------------
'功能：打开新网PACS的参数设置窗口
'返回：
'------------------------------------------------
    On Error GoTo err
    
    Call fillParams
    Me.Show 1, frmParent
    
    Exit Function
err:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Call SaveParams
    Unload Me
End Sub

Private Sub fillParams()
'------------------------------------------------
'功能：填充新网PACS的参数
'返回：
'------------------------------------------------
    On Error GoTo err
    
    '从中联ORACLE 模块参数中获取新网的数据库服务器IP地址，用户名和密码
    txtDBServerIP = zlDatabase.GetPara("XW数据库服务器IP", glngSys, G_LNG_XWPACSVIEW_MODULE, "")
    txtDBServerUser = zlDatabase.GetPara("XW数据库服务器用户名", glngSys, G_LNG_XWPACSVIEW_MODULE, "")
    txtDBServerPswd = zlDatabase.GetPara("XW数据库服务器密码", glngSys, G_LNG_XWPACSVIEW_MODULE, "")
    
    txtWebServerIP = zlDatabase.GetPara("XWWEB服务器IP", glngSys, G_LNG_XWPACSVIEW_MODULE, "")
    
    txtDelImageUser = zlDatabase.GetPara("XW删除图像用户名", glngSys, G_LNG_XWPACSVIEW_MODULE, "")
    txtDelImagePswd = zlDatabase.GetPara("XW删除图像密码", glngSys, G_LNG_XWPACSVIEW_MODULE, "")
    
    txtSendImageUser = zlDatabase.GetPara("XW发送图像用户名", glngSys, G_LNG_XWPACSVIEW_MODULE, "")
    txtSendImagePswd = zlDatabase.GetPara("XW发送图像密码", glngSys, G_LNG_XWPACSVIEW_MODULE, "")
    
    txtDVDBurnUser = zlDatabase.GetPara("XW光盘刻录用户名", glngSys, G_LNG_XWPACSVIEW_MODULE, "")
    txtDVDBurnPswd = zlDatabase.GetPara("XW光盘刻录密码", glngSys, G_LNG_XWPACSVIEW_MODULE, "")
    
    txtImageShare = zlDatabase.GetPara("XW历史图像共享目录", glngSys, G_LNG_XWPACSVIEW_MODULE, "DCMSHARE")
    
    chkLog.value = IIf(Val(zlDatabase.GetPara("XW记录接口日志", glngSys, G_LNG_XWPACSVIEW_MODULE, "0")) = 1, 1, 0)
    
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub SaveParams()
'------------------------------------------------
'功能：保存新网PACS的参数
'返回：
'------------------------------------------------
    On Error GoTo err
    
    '将新网PACS的参数设置保存到中联ORACLE 模块参数中
    Call zlDatabase.SetPara("XW数据库服务器IP", txtDBServerIP.Text, glngSys, G_LNG_XWPACSVIEW_MODULE)
    Call zlDatabase.SetPara("XW数据库服务器用户名", txtDBServerUser.Text, glngSys, G_LNG_XWPACSVIEW_MODULE)
    Call zlDatabase.SetPara("XW数据库服务器密码", txtDBServerPswd.Text, glngSys, G_LNG_XWPACSVIEW_MODULE)
    
    Call zlDatabase.SetPara("XWWEB服务器IP", txtWebServerIP.Text, glngSys, G_LNG_XWPACSVIEW_MODULE)
    
    Call zlDatabase.SetPara("XW删除图像用户名", txtDelImageUser.Text, glngSys, G_LNG_XWPACSVIEW_MODULE)
    Call zlDatabase.SetPara("XW删除图像密码", txtDelImagePswd.Text, glngSys, G_LNG_XWPACSVIEW_MODULE)
    
    Call zlDatabase.SetPara("XW发送图像用户名", txtSendImageUser.Text, glngSys, G_LNG_XWPACSVIEW_MODULE)
    Call zlDatabase.SetPara("XW发送图像密码", txtSendImagePswd.Text, glngSys, G_LNG_XWPACSVIEW_MODULE)
    
    Call zlDatabase.SetPara("XW光盘刻录用户名", txtDVDBurnUser.Text, glngSys, G_LNG_XWPACSVIEW_MODULE)
    Call zlDatabase.SetPara("XW光盘刻录密码", txtDVDBurnPswd.Text, glngSys, G_LNG_XWPACSVIEW_MODULE)
    
    Call zlDatabase.SetPara("XW历史图像共享目录", txtImageShare.Text, glngSys, G_LNG_XWPACSVIEW_MODULE)
    
    Call zlDatabase.SetPara("XW记录接口日志", chkLog.value, glngSys, G_LNG_XWPACSVIEW_MODULE)
    
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub Command1_Click()
On Error GoTo errHandle
    Call XWTestDBConnection(txtDBServerIP.Text, txtDBServerUser.Text, txtDBServerPswd.Text)
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

