VERSION 5.00
Begin VB.Form frmParaHNCA 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "参数设置"
   ClientHeight    =   3990
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7065
   Icon            =   "frmParaHNCA.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3990
   ScaleWidth      =   7065
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox picPara 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3375
      Index           =   2
      Left            =   0
      ScaleHeight     =   3375
      ScaleWidth      =   7035
      TabIndex        =   3
      Top             =   0
      Width           =   7035
      Begin VB.OptionButton opt 
         BackColor       =   &H80000005&
         Caption         =   "RSA"
         Height          =   255
         Index           =   0
         Left            =   5040
         TabIndex        =   14
         Top             =   1425
         Value           =   -1  'True
         Width           =   735
      End
      Begin VB.OptionButton opt 
         BackColor       =   &H80000005&
         Caption         =   "SM2"
         Height          =   255
         Index           =   1
         Left            =   6000
         TabIndex        =   13
         Top             =   1425
         Width           =   735
      End
      Begin VB.TextBox txtSSLPort 
         Height          =   360
         Left            =   1410
         TabIndex        =   8
         Top             =   2685
         Width           =   2145
      End
      Begin VB.TextBox txtUrl 
         Height          =   360
         Left            =   1410
         TabIndex        =   7
         Top             =   765
         Width           =   5265
      End
      Begin VB.CheckBox chkTS 
         BackColor       =   &H8000000E&
         Caption         =   "启用时间戳"
         Height          =   375
         Left            =   1410
         TabIndex        =   6
         Top             =   1365
         Width           =   1455
      End
      Begin VB.TextBox txtTSIP 
         Height          =   360
         Left            =   1410
         TabIndex        =   5
         Top             =   2070
         Width           =   2145
      End
      Begin VB.TextBox txtTSPort 
         Height          =   360
         Left            =   5010
         TabIndex        =   4
         Top             =   2070
         Width           =   1665
      End
      Begin VB.Label lblUrl 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "签名服务器URL"
         Height          =   180
         Left            =   120
         TabIndex        =   16
         Top             =   855
         Width           =   1170
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "签名算法"
         Height          =   180
         Left            =   4080
         TabIndex        =   15
         Top             =   1455
         Width           =   720
      End
      Begin VB.Label lblNote 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   $"frmParaHNCA.frx":000C
         Height          =   360
         Left            =   120
         TabIndex        =   12
         Top             =   270
         Width           =   6570
      End
      Begin VB.Label lblPenUrl 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "SSL端口"
         Height          =   180
         Left            =   660
         TabIndex        =   11
         Top             =   2775
         Width           =   630
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "时间戳IP"
         Height          =   180
         Left            =   570
         TabIndex        =   10
         Top             =   2160
         Width           =   720
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "时间戳端口"
         Height          =   180
         Left            =   3900
         TabIndex        =   9
         Top             =   2160
         Width           =   900
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
      ScaleWidth      =   7065
      TabIndex        =   0
      Top             =   3375
      Width           =   7065
      Begin VB.CommandButton cmdCancel 
         Caption         =   "取消(&C)"
         Height          =   360
         Left            =   5865
         TabIndex        =   2
         Top             =   150
         Width           =   1100
      End
      Begin VB.CommandButton cmdOK 
         BackColor       =   &H8000000E&
         Caption         =   "确定(&O)"
         Height          =   360
         Left            =   4665
         TabIndex        =   1
         Top             =   150
         Width           =   1100
      End
   End
End
Attribute VB_Name = "frmParaHNCA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim objCA As New clsHNCA
    gudtPara.strSignURL = Trim(txtUrl.Text)
    gudtPara.strTSIP = Trim(txtTSIP.Text)
    gudtPara.strTSPort = Trim(txtTSPort.Text)
    gudtPara.strSSLPort = Trim(txtSSLPort.Text)
    gudtPara.blnISTS = IIf(chkTS.Value = vbChecked, True, False)
    gudtPara.bytSignVersion = IIf(opt(0).Value, 0, 1)
     
    Call objCA.HNCA_SetParaStr
    Unload Me
End Sub

Private Sub Form_Load()
    Dim objCA As New clsHNCA
    
    Call objCA.HNCA_GetPara
    txtUrl.Text = gudtPara.strSignURL
    txtTSIP.Text = gudtPara.strTSIP
    txtTSPort.Text = gudtPara.strTSPort
    txtSSLPort.Text = gudtPara.strSSLPort
    chkTS.Value = IIf(gudtPara.blnISTS, vbChecked, vbUnchecked)
    opt(0).Value = gudtPara.bytSignVersion = 0
    opt(1).Value = gudtPara.bytSignVersion = 1
End Sub
