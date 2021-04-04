VERSION 5.00
Begin VB.Form frmShowHisForms 
   BorderStyle     =   0  'None
   Caption         =   "中联显示HIS窗口NEW"
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   Icon            =   "frmShowHisForms.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer TimerShow 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   2760
      Top             =   1200
   End
   Begin VB.Timer TimerCaption 
      Interval        =   5000
      Left            =   0
      Top             =   0
   End
End
Attribute VB_Name = "frmShowHisForms"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private mlngTime As Long

Public Sub ShowMe(blnShow As Boolean)

On Error GoTo ErrorHand
    
    If blnShow Then Call Me.Show
    
    Me.Caption = HIS_CAPTION
    Exit Sub
ErrorHand:
    If errHandle("zlSoftCISInterface.frmShowHisForms.ShowMe", "显示窗口出现错误") = 1 Then Resume
End Sub

Private Sub Form_Load()
    On Error Resume Next
    
    TimerCaption.Interval = 30000    '30秒钟
    '挂上截获消息的hook
    plngPreWndProc = Hook(Me.hWnd)
    
    Call GetWindowThreadProcessId(Me.hWnd, glngPid)
    
    
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    '卸载hook
    Unhook Me.hWnd, plngPreWndProc
End Sub

Private Sub TimerCaption_Timer()
    Dim lngWinHandle As Long
    
    On Error GoTo err
    
    If Me.Caption <> HIS_CAPTION Then
        Me.Caption = HIS_CAPTION
    End If
    
    '查找消息循环主窗体
    lngWinHandle = FindWindow(vbNullString, HIS_CAPTION)
    If lngWinHandle = 0 Then
        '如果窗口句柄为0 ，则强制退出当前程序
        Call CloseAllForms
    End If
    Exit Sub
err:
   
End Sub


Private Sub TimerShow_Timer()
    If mlngTime > 5 Then
        EnumChildWindows GetDesktopWindow, AddressOf EnumChildProc, ByVal 0
        mlngTime = 0
        TimerShow.Enabled = False
    Else
        mlngTime = mlngTime + 1
    End If
End Sub
