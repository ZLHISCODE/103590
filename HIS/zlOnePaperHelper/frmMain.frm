VERSION 5.00
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1305
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1440
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   1305
   ScaleWidth      =   1440
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'ÆÁÄ»ÖÐÐÄ
   Begin VB.Timer TimerShow 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   465
      Top             =   330
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlngTime As Long

Private Sub Form_Load()
    Call GetWindowThreadProcessId(Me.hwnd, glngPid)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mlngTime = 0
End Sub

Private Sub TimerShow_Timer()
    If mlngTime < 5 Then
        gblnFinded = False
        EnumChildWindows GetDesktopWindow, AddressOf EnumChildProc, ByVal 0
        If gblnFinded Then
            TimerShow.Enabled = False
            mlngTime = 0
        Else
            mlngTime = mlngTime + 1
        End If
        gblnFinded = False
    Else
        TimerShow.Enabled = False
        mlngTime = 0
    End If
End Sub
