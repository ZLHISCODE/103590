VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmFlash 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   810
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   5775
   ControlBox      =   0   'False
   Icon            =   "frmFlash.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   810
   ScaleWidth      =   5775
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin MSComCtl2.Animation avi 
      Height          =   675
      Left            =   135
      TabIndex        =   1
      Top             =   75
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   1191
      _Version        =   393216
      FullWidth       =   50
      FullHeight      =   45
   End
   Begin VB.Label lblFile 
      Caption         =   "<AVI文件不存在>"
      ForeColor       =   &H00C0C0C0&
      Height          =   210
      Left            =   1005
      TabIndex        =   2
      Top             =   555
      Visible         =   0   'False
      Width           =   1440
   End
   Begin VB.Label lbl提示 
      Caption         =   "正在装入数据，请稍候..."
      Height          =   180
      Left            =   1005
      TabIndex        =   0
      Top             =   270
      Width           =   4575
   End
End
Attribute VB_Name = "frmFlash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

