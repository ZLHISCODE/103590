VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmShowMessage 
   BorderStyle     =   0  'None
   Caption         =   "提示消息"
   ClientHeight    =   3990
   ClientLeft      =   0
   ClientTop       =   -15
   ClientWidth     =   6420
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3990
   ScaleWidth      =   6420
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.Timer tmrMain 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   0
      Top             =   0
   End
   Begin zl9NewQuery.ctlButton ctlOK 
      Height          =   720
      Left            =   2430
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   2880
      Width           =   1560
      _ExtentX        =   2752
      _ExtentY        =   1270
      Caption         =   "确定"
      AutoSize        =   0   'False
      ButtonHeight    =   600
   End
   Begin MSComctlLib.ImageList ilsImage 
      Left            =   0
      Top             =   480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShowMessage.frx":0000
            Key             =   "up"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShowMessage.frx":039A
            Key             =   "down"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShowMessage.frx":0734
            Key             =   "reset"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShowMessage.frx":0ACE
            Key             =   "ok"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShowMessage.frx":7330
            Key             =   "close"
         EndProperty
      EndProperty
   End
   Begin VB.Label lblMessage 
      BackStyle       =   0  'Transparent
      Caption         =   "您已欠费！"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   855
      Left            =   2280
      TabIndex        =   3
      Top             =   1320
      Width           =   3255
   End
   Begin VB.Image Imgbak 
      Height          =   2130
      Left            =   480
      Picture         =   "frmShowMessage.frx":DB92
      Stretch         =   -1  'True
      Top             =   840
      Width           =   1605
   End
   Begin VB.Label lblTime 
      BackStyle       =   0  'Transparent
      Caption         =   "(5)"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4200
      TabIndex        =   2
      Top             =   600
      Width           =   375
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "提示消息"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   615
      Left            =   2070
      TabIndex        =   1
      Top             =   360
      Width           =   2295
   End
End
Attribute VB_Name = "frmShowMessage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mintLoop As Integer

Private Sub ctlOK_CommandClick()
    tmrMain.Enabled = False
    Unload Me
End Sub

Private Sub Form_Load()
    Me.ctlOK.Picture = ilsImage.ListImages("ok")
    If Dir(App.Path & "\图形\提示信息确认窗体左面背景.pic") <> "" Then
        Imgbak.Picture = LoadPicture(App.Path & "\图形\提示信息确认窗体左面背景.pic")
    End If
End Sub

Private Sub Form_Paint()
    Call DrawColorToColor(Me, Me.BackColor, &HFFC0C1, , True)
End Sub

Private Sub tmrMain_Timer()
    mintLoop = mintLoop - 1
    Me.lblTime.Caption = "(" & mintLoop & ")"
    If mintLoop = 0 Then
        tmrMain.Enabled = False
        Unload Me
    End If
End Sub

Public Sub ShowMe(frmParent As Object, strMessage As String)
    tmrMain.Enabled = True
    mintLoop = 5
    Me.lblTime.Caption = "(" & mintLoop & ")"
    Me.lblMessage.Caption = strMessage
    Me.Show vbModal, frmParent
End Sub
