VERSION 5.00
Begin VB.Form frmParaSZCA 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "参数设置"
   ClientHeight    =   4170
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8715
   Icon            =   "frmParaSZCA.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4170
   ScaleWidth      =   8715
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.PictureBox picPara 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3495
      Left            =   0
      ScaleHeight     =   3495
      ScaleWidth      =   8685
      TabIndex        =   3
      Top             =   0
      Width           =   8685
      Begin VB.Frame fraPara 
         BackColor       =   &H8000000E&
         Caption         =   "时间戳服务器配置"
         Height          =   1080
         Index           =   1
         Left            =   120
         TabIndex        =   7
         Top             =   1800
         Width           =   8415
         Begin VB.TextBox txtPara 
            Height          =   360
            Index           =   3
            Left            =   5520
            MaxLength       =   8
            TabIndex        =   9
            Top             =   390
            Width           =   1935
         End
         Begin VB.TextBox txtPara 
            Height          =   360
            Index           =   2
            Left            =   1320
            MaxLength       =   16
            TabIndex        =   8
            Top             =   390
            Width           =   1935
         End
         Begin VB.Label lblName 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000E&
            Caption         =   "服务器端口"
            Height          =   180
            Index           =   3
            Left            =   4440
            TabIndex        =   11
            Top             =   480
            Width           =   900
         End
         Begin VB.Label lblName 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000E&
            Caption         =   "服务器IP"
            Height          =   180
            Index           =   2
            Left            =   240
            TabIndex        =   10
            Top             =   480
            Width           =   720
         End
      End
      Begin VB.OptionButton optVer 
         BackColor       =   &H80000005&
         Caption         =   "老版"
         Height          =   255
         Index           =   0
         Left            =   960
         TabIndex        =   6
         Top             =   233
         Width           =   855
      End
      Begin VB.OptionButton optVer 
         BackColor       =   &H80000005&
         Caption         =   "新版"
         Height          =   255
         Index           =   1
         Left            =   2640
         TabIndex        =   5
         Top             =   233
         Width           =   735
      End
      Begin VB.TextBox txtPara 
         Height          =   360
         Index           =   0
         Left            =   120
         TabIndex        =   4
         Top             =   960
         Width           =   8415
      End
      Begin VB.Label lblNote 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "示例：http://202.103.144.98:7006/SZCAJavaCAS/services/szcaCAValidate"
         ForeColor       =   &H00FF0000&
         Height          =   180
         Left            =   1440
         TabIndex        =   14
         Top             =   720
         Width           =   6120
      End
      Begin VB.Label lblName 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "版本"
         Height          =   180
         Index           =   1
         Left            =   120
         TabIndex        =   13
         Top             =   270
         Width           =   360
      End
      Begin VB.Label lblName 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "签名服务(WSDL)"
         Height          =   180
         Index           =   0
         Left            =   120
         TabIndex        =   12
         Top             =   720
         Width           =   1260
      End
   End
   Begin VB.PictureBox picBottom 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H00EFF0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   0
      ScaleHeight     =   615
      ScaleWidth      =   8715
      TabIndex        =   0
      Top             =   3555
      Width           =   8715
      Begin VB.CommandButton cmdOK 
         BackColor       =   &H8000000E&
         Caption         =   "确定(&O)"
         Height          =   360
         Left            =   6240
         TabIndex        =   2
         Top             =   120
         Width           =   1100
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "取消(&C)"
         Height          =   360
         Left            =   7440
         TabIndex        =   1
         Top             =   120
         Width           =   1100
      End
   End
End
Attribute VB_Name = "frmParaSZCA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
     
    With gudtPara
        .strSignURL = txtPara(0).Text
        .strTSIP = txtPara(2).Text
        .strTSPort = txtPara(3).Text
        If optVer(0).Value = True Then
            .bytSignVersion = 0
        Else
            .bytSignVersion = 1
        End If
    End With
    Call SZCA_SetParaStr
    Unload Me
End Sub

Private Sub Form_Load()
    Call SZCA_GetPara
    With gudtPara
        txtPara(0).Text = .strSignURL
        txtPara(2).Text = .strTSIP
        txtPara(3).Text = .strTSPort
        optVer(0).Value = (.bytSignVersion = 0)
        optVer(1).Value = (.bytSignVersion = 1)
    End With
End Sub

