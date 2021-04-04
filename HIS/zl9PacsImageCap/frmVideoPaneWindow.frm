VERSION 5.00
Begin VB.Form frmVideoPaneWindow 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3030
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4560
   ForeColor       =   &H00808080&
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer TimerRePaint 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   345
      Top             =   2040
   End
   Begin VB.PictureBox picBackImg 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1920
      Left            =   0
      Picture         =   "frmVideoPaneWindow.frx":0000
      ScaleHeight     =   1920
      ScaleWidth      =   1920
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1920
   End
End
Attribute VB_Name = "frmVideoPaneWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub DrawBackground()
'绘制背景图像
    Dim i As Integer
    Dim Count As Integer
    Dim wordRect As RECT
    
    Count = 2
    wordRect.Bottom = 45
    wordRect.Right = 200

    If Me.picBackImg.Height * 3 >= Me.Height Then Count = 1

    Call Me.Cls
    
    For i = 0 To Count
        Call Me.PaintPicture(Me.picBackImg.Picture, _
            Round(Me.Width / (i + 1)) - Me.picBackImg.Width + 200, _
            Round((Me.Height / 3) * (i + 1) - Me.picBackImg.Height), _
            Me.picBackImg.Width, Me.picBackImg.Height)

        If gobjVideo Is Nothing Then
            wordRect.Left = Me.ScaleX(Round(Me.Width / (i + 1)) - Me.picBackImg.Width, vbTwips, vbPixels) + 35
        Else
            wordRect.Left = Me.ScaleX(Round(Me.Width / (i + 1)) - Me.picBackImg.Width, vbTwips, vbPixels) + 17
        End If
        
        wordRect.Top = Me.ScaleY(Round((Me.Height / 3) * (i + 1) - Me.picBackImg.Height), vbTwips, vbPixels) - 30

        wordRect.Right = wordRect.Left + 200
        wordRect.Bottom = wordRect.Top + 90

        If gobjVideo Is Nothing Then
            Call DrawText(Me.hdc, "视频未被注册" & vbCrLf & "已禁用视频源", 27, wordRect, 0)
        Else
            Call DrawText(Me.hdc, "视频已被其他窗口占用" & vbCrLf & " 双击鼠标可恢复显示", 44, wordRect, 0)
        End If
    Next i
End Sub

Private Sub Form_Paint()
    TimerRePaint.Enabled = True
End Sub

Private Sub TimerRePaint_Timer()
On Error Resume Next
    TimerRePaint.Enabled = False
    
    Call DrawBackground
err.Clear
End Sub
