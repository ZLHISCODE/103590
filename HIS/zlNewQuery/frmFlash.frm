VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmFlash 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   0  'None
   ClientHeight    =   870
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5070
   ControlBox      =   0   'False
   Icon            =   "frmFlash.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   870
   ScaleWidth      =   5070
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin MSComCtl2.Animation avi 
      Height          =   675
      Left            =   225
      TabIndex        =   1
      Top             =   75
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   1191
      _Version        =   393216
      FullWidth       =   50
      FullHeight      =   45
   End
   Begin VB.Shape shp 
      BorderColor     =   &H00FF8080&
      BorderWidth     =   2
      Height          =   810
      Left            =   15
      Top             =   15
      Width           =   5010
   End
   Begin VB.Label lblFile 
      BackStyle       =   0  'Transparent
      Caption         =   "<AVI文件不存在>"
      Height          =   210
      Left            =   1125
      TabIndex        =   2
      Top             =   555
      Visible         =   0   'False
      Width           =   1440
   End
   Begin VB.Label lbl提示 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "正在装入数据，请稍候..."
      Height          =   180
      Left            =   1185
      TabIndex        =   0
      Top             =   270
      Width           =   2070
   End
End
Attribute VB_Name = "frmFlash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
    On Error Resume Next
    shp.Left = 15
    shp.Top = 15
    shp.Width = Me.Width - 30
    shp.Height = Me.Height - 30
    
End Sub
