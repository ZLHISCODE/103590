VERSION 5.00
Begin VB.Form frmTips 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1935
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2265
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   129
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   151
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.Timer tmrControl 
      Enabled         =   0   'False
      Left            =   960
      Top             =   720
   End
   Begin VB.Image imgArrow 
      Height          =   120
      Index           =   3
      Left            =   720
      Picture         =   "frmTips.frx":0000
      Top             =   1440
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.Image imgArrow 
      Height          =   120
      Index           =   2
      Left            =   480
      Picture         =   "frmTips.frx":04AA
      Top             =   1440
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.Image imgArrow 
      Height          =   120
      Index           =   1
      Left            =   240
      Picture         =   "frmTips.frx":0954
      Top             =   1440
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.Image imgArrow 
      Height          =   120
      Index           =   0
      Left            =   0
      Picture         =   "frmTips.frx":0DFE
      Top             =   1440
      Visible         =   0   'False
      Width           =   120
   End
End
Attribute VB_Name = "frmTips"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Click()
    Me.Hide
End Sub

Private Sub Form_Load()
    '鼠标穿透效果
    Dim Ret As Long
    Ret = GetWindowLong(Me.hWnd, GWL_EXSTYLE)
    Ret = Ret Or WS_EX_LAYERED Or WS_EX_Transparent
    SetWindowLong Me.hWnd, GWL_EXSTYLE, Ret
    '窗体透明
    SetTransparentForm Me, 255
    
    tmrControl.Interval = 1000
    tmrControl.Enabled = False
End Sub

Private Sub Form_Terminate()
    Set frmTips = Nothing
End Sub

Private Sub tmrControl_Timer()
    On Error Resume Next
    m_WndStopoverTimeVal = m_WndStopoverTimeVal + 1
    If m_WndStopoverTimeVal = 4 Then
       m_WndStopoverTimeVal = 0
       tmrControl.Enabled = False
       frmTips.ZOrder 0
       frmTips.Hide
    End If
End Sub

