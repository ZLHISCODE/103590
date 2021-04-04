VERSION 5.00
Begin VB.Form frmParaXibuCA 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "参数设置"
   ClientHeight    =   3330
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8310
   Icon            =   "frmParaXibuCA.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3330
   ScaleWidth      =   8310
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox picBottom 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H00EFF0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   0
      ScaleHeight     =   615
      ScaleWidth      =   8310
      TabIndex        =   5
      Top             =   2715
      Width           =   8310
      Begin VB.CommandButton cmdPara 
         Caption         =   "取消(&C)"
         Height          =   360
         Index           =   1
         Left            =   7080
         TabIndex        =   7
         Top             =   120
         Width           =   1100
      End
      Begin VB.CommandButton cmdPara 
         BackColor       =   &H8000000E&
         Caption         =   "确定(&O)"
         Height          =   360
         Index           =   0
         Left            =   5880
         TabIndex        =   6
         Top             =   120
         Width           =   1100
      End
   End
   Begin VB.PictureBox picMain 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2655
      Left            =   0
      ScaleHeight     =   2655
      ScaleWidth      =   8295
      TabIndex        =   0
      Top             =   0
      Width           =   8295
      Begin VB.TextBox txtURL 
         Height          =   360
         Left            =   120
         TabIndex        =   8
         Top             =   600
         Width           =   8025
      End
      Begin VB.TextBox txtTSIP 
         Height          =   360
         Left            =   960
         TabIndex        =   2
         Top             =   1830
         Width           =   2865
      End
      Begin VB.TextBox txtTSPort 
         Height          =   360
         Left            =   5280
         TabIndex        =   1
         Top             =   1830
         Width           =   2865
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "示例:http://113.204.104.142:8082/SignatureServer/services/SignatureService?wsdl"
         ForeColor       =   &H00000000&
         Height          =   180
         Left            =   120
         TabIndex        =   10
         Top             =   1080
         Width           =   7110
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "验签服务(WSDL)"
         ForeColor       =   &H00000000&
         Height          =   180
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   1260
      End
      Begin VB.Label lblUrl 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "时间戳IP"
         ForeColor       =   &H00000000&
         Height          =   180
         Left            =   120
         TabIndex        =   4
         Top             =   1920
         Width           =   720
      End
      Begin VB.Label lblPort 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "时间戳端口号"
         ForeColor       =   &H00000000&
         Height          =   180
         Left            =   4080
         TabIndex        =   3
         Top             =   1920
         Width           =   1080
      End
   End
End
Attribute VB_Name = "frmParaXibuCA"
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
        gudtPara.strSignURL = Trim(txtUrl.Text)
        gudtPara.strTSIP = Trim(txtTSIP.Text)
        gudtPara.strTSPort = Trim(txtTSPort.Text)
        Call UpdateThirdPara(CON_PAR_西部, 1, "签名服务WSDL", gudtPara.strSignURL, "签名服务WSDL")
        Call UpdateThirdPara(CON_PAR_西部, 2, "时间戳IP", gudtPara.strTSIP, "时间戳服务IP地址")
        Call UpdateThirdPara(CON_PAR_西部, 3, "时间戳端口", gudtPara.strTSPort, "时间戳服务端口号")
    End If
    Unload Me
End Sub

Private Sub Form_Load()
    
    Call XiBuCA_GetPara
    txtUrl.Text = gudtPara.strSignURL
    txtTSIP.Text = gudtPara.strTSIP
    txtTSPort.Text = gudtPara.strTSPort
End Sub

