VERSION 5.00
Begin VB.Form frmParaSDCA 
   BackColor       =   &H8000000E&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "参数设置"
   ClientHeight    =   4020
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4905
   Icon            =   "frmParaSDCA.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4020
   ScaleWidth      =   4905
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox picPara 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3195
      Left            =   120
      ScaleHeight     =   3195
      ScaleWidth      =   4695
      TabIndex        =   3
      Top             =   120
      Width           =   4695
      Begin VB.OptionButton opt 
         BackColor       =   &H80000005&
         Caption         =   "版本1.0"
         Height          =   255
         Index           =   0
         Left            =   600
         TabIndex        =   15
         Top             =   90
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.OptionButton opt 
         BackColor       =   &H80000005&
         Caption         =   "版本2.0"
         Height          =   255
         Index           =   1
         Left            =   3000
         TabIndex        =   14
         Top             =   90
         Width           =   1095
      End
      Begin VB.Frame fraPara 
         BackColor       =   &H8000000E&
         Caption         =   "时间戳服务器"
         Height          =   1080
         Index           =   1
         Left            =   0
         TabIndex        =   9
         Top             =   1920
         Width           =   4695
         Begin VB.TextBox txt 
            Height          =   360
            Index           =   3
            Left            =   3480
            MaxLength       =   8
            TabIndex        =   11
            Top             =   390
            Width           =   975
         End
         Begin VB.TextBox txt 
            Height          =   360
            Index           =   2
            Left            =   750
            MaxLength       =   16
            TabIndex        =   10
            Top             =   390
            Width           =   2055
         End
         Begin VB.Label lblName 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000E&
            Caption         =   "端口"
            Height          =   180
            Index           =   3
            Left            =   3000
            TabIndex        =   13
            Top             =   480
            Width           =   360
         End
         Begin VB.Label lblName 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000E&
            Caption         =   "地址"
            Height          =   180
            Index           =   2
            Left            =   240
            TabIndex        =   12
            Top             =   480
            Width           =   360
         End
      End
      Begin VB.Frame fraPara 
         BackColor       =   &H8000000E&
         Caption         =   "验签服务器"
         Height          =   1080
         Index           =   0
         Left            =   0
         TabIndex        =   4
         Top             =   480
         Width           =   4695
         Begin VB.TextBox txt 
            Height          =   360
            Index           =   0
            Left            =   750
            MaxLength       =   16
            TabIndex        =   6
            Top             =   390
            Width           =   2055
         End
         Begin VB.TextBox txt 
            Height          =   360
            Index           =   1
            Left            =   3480
            MaxLength       =   8
            TabIndex        =   5
            Top             =   390
            Width           =   975
         End
         Begin VB.Label lblName 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000E&
            Caption         =   "地址"
            Height          =   180
            Index           =   1
            Left            =   240
            TabIndex        =   8
            Top             =   480
            Width           =   360
         End
         Begin VB.Label lblName 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000E&
            Caption         =   "端口"
            Height          =   180
            Index           =   0
            Left            =   3000
            TabIndex        =   7
            Top             =   480
            Width           =   360
         End
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "版本"
         Height          =   180
         Left            =   0
         TabIndex        =   16
         Top             =   120
         Width           =   360
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
      ScaleWidth      =   4905
      TabIndex        =   0
      Top             =   3405
      Width           =   4905
      Begin VB.CommandButton cmdPara 
         Caption         =   "取消(&C)"
         Height          =   360
         Index           =   1
         Left            =   3600
         TabIndex        =   2
         Top             =   120
         Width           =   1100
      End
      Begin VB.CommandButton cmdPara 
         BackColor       =   &H8000000E&
         Caption         =   "确定(&O)"
         Height          =   360
         Index           =   0
         Left            =   2400
         TabIndex        =   1
         Top             =   120
         Width           =   1100
      End
   End
End
Attribute VB_Name = "frmParaSDCA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdPara_Click(Index As Integer)
    With gudtPara
        .strSignIP = Trim(txt(0).Text)
        .strSignPort = Trim(txt(1).Text)
        .strTSIP = Trim(txt(2).Text)
        .strTSPort = Trim(txt(3).Text)
        If opt(0).Value Then
            .bytSignVersion = 0
        Else
            .bytSignVersion = 1
        End If
    End With
    Call SDCA_SetPara
    Unload Me
End Sub

Private Sub Form_Load()
    Call SDCA_GetPara
    With gudtPara
        txt(0).Text = .strSignIP
        txt(1).Text = .strSignPort
        txt(2).Text = .strTSIP
        txt(3).Text = .strTSPort
        opt(0).Value = (.bytSignVersion = 0)
        opt(1).Value = (.bytSignVersion = 1)
    End With
End Sub

 
