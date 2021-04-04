VERSION 5.00
Begin VB.Form frmWait 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "请等待……"
   ClientHeight    =   1650
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4725
   ControlBox      =   0   'False
   LinkTopic       =   "frmWait"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   11  'Hourglass
   ScaleHeight     =   110
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   315
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   5250
      Top             =   165
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   75
      Left            =   -45
      Picture         =   "frmWait.frx":0000
      ScaleHeight     =   5
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   335
      TabIndex        =   0
      Top             =   1005
      Width           =   5025
   End
   Begin VB.Label lbl内容 
      Caption         =   "#"
      Height          =   180
      Left            =   330
      TabIndex        =   4
      Top             =   1290
      Width           =   4140
   End
   Begin VB.Label lblSoft 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "服务器管理工具"
      ForeColor       =   &H000000C0&
      Height          =   180
      Left            =   2220
      TabIndex        =   3
      Top             =   600
      Width           =   1260
   End
   Begin VB.Label lblCompany 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "中联软件"
      BeginProperty Font 
         Name            =   "楷体_GB2312"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   300
      Left            =   1095
      TabIndex        =   2
      Top             =   210
      Width           =   1260
   End
   Begin VB.Label lblBack 
      BackColor       =   &H8000000B&
      Height          =   1125
      Left            =   -30
      TabIndex        =   1
      Top             =   1065
      Width           =   5895
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   225
      Picture         =   "frmWait.frx":096C
      Stretch         =   -1  'True
      Top             =   120
      Width           =   480
   End
End
Attribute VB_Name = "frmWait"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    lblCompany = gstrProductName & "软件"
    Call ApplyOEM_Picture(Image1, "Picture")
End Sub

Private Sub Timer1_Timer()
    Static i As Long
    i = i + 20
    If i >= Picture1.ScaleWidth Then i = 1
    
    Picture1.PaintPicture Picture1.Picture, i, 0, Picture1.ScaleWidth - i, Picture1.ScaleHeight, 0, 0, Picture1.ScaleWidth - i, Picture1.ScaleHeight
    Picture1.PaintPicture Picture1.Picture, 0, 0, i, Picture1.ScaleHeight, Picture1.ScaleWidth - i, 0, i, Picture1.ScaleHeight
End Sub

Public Sub BeginWait(ByVal str内容 As String)
    lbl内容.Caption = str内容
    frmWait.Show , frmMDIMain
    frmWait.Show
    DoEvents
End Sub

Public Sub EndWait()
    Unload frmWait
    Set frmWait = Nothing
End Sub
