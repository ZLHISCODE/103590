VERSION 5.00
Begin VB.Form frmParaXJCA 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��������"
   ClientHeight    =   3105
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6825
   Icon            =   "frmParaXJCA.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3105
   ScaleWidth      =   6825
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.PictureBox picPara 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2535
      Index           =   2
      Left            =   0
      ScaleHeight     =   2535
      ScaleWidth      =   6795
      TabIndex        =   3
      Top             =   0
      Width           =   6795
      Begin VB.ComboBox cboKey 
         Height          =   300
         Left            =   1208
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   1920
         Width           =   2055
      End
      Begin VB.TextBox txtUrl 
         Height          =   360
         Left            =   1208
         TabIndex        =   5
         Top             =   750
         Width           =   5025
      End
      Begin VB.TextBox txtAppID 
         Height          =   360
         Left            =   1208
         TabIndex        =   4
         Top             =   1350
         Width           =   5025
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "KEY����"
         Height          =   180
         Index           =   1
         Left            =   428
         TabIndex        =   10
         Top             =   1980
         Width           =   630
      End
      Begin VB.Label lblNote 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "������URLʾ��:http://124.117.245.71:48080/webServices/ssoService"
         Height          =   180
         Left            =   120
         TabIndex        =   8
         Top             =   270
         Width           =   5760
      End
      Begin VB.Label lblUrl 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "������URL"
         Height          =   180
         Left            =   248
         TabIndex        =   7
         Top             =   840
         Width           =   810
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "Ӧ��ID"
         Height          =   180
         Left            =   518
         TabIndex        =   6
         Top             =   1440
         Width           =   540
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
      ScaleWidth      =   6825
      TabIndex        =   0
      Top             =   2490
      Width           =   6825
      Begin VB.CommandButton cmdCancel 
         Caption         =   "ȡ��(&C)"
         Height          =   360
         Left            =   5625
         TabIndex        =   2
         Top             =   150
         Width           =   1100
      End
      Begin VB.CommandButton cmdOK 
         BackColor       =   &H8000000E&
         Caption         =   "ȷ��(&O)"
         Height          =   360
         Left            =   4425
         TabIndex        =   1
         Top             =   150
         Width           =   1100
      End
   End
End
Attribute VB_Name = "frmParaXJCA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    gudtPara.strSignURL = Trim(txtUrl.Text)
    gudtPara.strOption = Trim(txtAppID.Text)
    gudtPara.intKeyType = cboKey.ListIndex
     
    Call XJCA_SetParaStr
    Unload Me
End Sub

Private Sub Form_Load()
    Call XJCA_GetPara
    txtUrl.Text = gudtPara.strSignURL
    txtAppID.Text = gudtPara.strOption
    cboKey.AddItem "��̩"
    cboKey.AddItem "����"
    cboKey.ListIndex = gudtPara.intKeyType
End Sub
