VERSION 5.00
Begin VB.Form frmXINANCA 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "参数设置"
   ClientHeight    =   2400
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4830
   Icon            =   "frmXINANCA.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2400
   ScaleWidth      =   4830
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox picMain 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1815
      Left            =   0
      ScaleHeight     =   1815
      ScaleWidth      =   4815
      TabIndex        =   2
      Top             =   0
      Width           =   4815
      Begin VB.TextBox txtTSPort 
         Height          =   360
         Left            =   1440
         TabIndex        =   8
         Top             =   750
         Width           =   2625
      End
      Begin VB.TextBox txtTSIP 
         Height          =   360
         Left            =   1440
         TabIndex        =   3
         Top             =   270
         Width           =   2625
      End
      Begin VB.Label lblPort 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "时间戳端口号"
         ForeColor       =   &H00C00000&
         Height          =   180
         Left            =   240
         TabIndex        =   5
         Top             =   840
         Width           =   1080
      End
      Begin VB.Label lblUrl 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "时间戳IP"
         ForeColor       =   &H00C00000&
         Height          =   180
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Width           =   720
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
      ScaleWidth      =   4830
      TabIndex        =   0
      Top             =   1785
      Width           =   4830
      Begin VB.CommandButton cmdPara 
         BackColor       =   &H8000000E&
         Caption         =   "确定(&O)"
         Height          =   360
         Index           =   0
         Left            =   2400
         TabIndex        =   7
         Top             =   120
         Width           =   1100
      End
      Begin VB.CommandButton cmdPara 
         Caption         =   "取消(&C)"
         Height          =   360
         Index           =   1
         Left            =   3600
         TabIndex        =   6
         Top             =   120
         Width           =   1100
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "取消(&C)"
         Height          =   360
         Left            =   5625
         TabIndex        =   1
         Top             =   150
         Width           =   1100
      End
   End
End
Attribute VB_Name = "frmXINANCA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum CMD_ENUM
    CMD_OK = 0
    CMD_CANCEL = 1
End Enum

Private Sub cmdPara_Click(Index As Integer)
    If Index = CMD_OK Then
        gudtPara.strTSIP = Trim(txtTSIP.Text)
        gudtPara.strTSPort = Trim(txtTSPort.Text)
        Call UpdateThirdPara(CON_PAR_信安, 1, "时间戳IP", gudtPara.strTSIP, "时间戳服务IP地址")
        Call UpdateThirdPara(CON_PAR_信安, 2, "时间戳端口", gudtPara.strTSPort, "时间戳服务端口号")
    End If
    Unload Me
End Sub

Private Sub Form_Load()
    Call XINANCA_GetPara
    txtTSIP.Text = gudtPara.strTSIP
    txtTSPort.Text = gudtPara.strTSPort
End Sub
