VERSION 5.00
Begin VB.Form frmViewImage 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   90
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   90
   Icon            =   "frmViewImage.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   90
   ScaleWidth      =   90
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer TimerCaption 
      Interval        =   5000
      Left            =   240
      Top             =   480
   End
End
Attribute VB_Name = "frmViewImage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub ShowMe(blnShow As Boolean)

On Error GoTo ErrorHand
    
    If blnShow Then Call Me.Show
    
    Me.Caption = HIS_CAPTION
    Exit Sub
ErrorHand:
    If errHandle("zlSoftViewImage.frmViewImage.ShowMe", "显示窗口出现错误") = 1 Then Resume
End Sub

Private Sub Form_Load()
    On Error Resume Next
    
    TimerCaption.Interval = 30000    '30秒钟
    '挂上截获消息的hook
    glngPreWndProc = Hook(Me.hWnd)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    '卸载hook
    Unhook Me.hWnd, glngPreWndProc
End Sub

Private Sub TimerCaption_Timer()
    Dim lngWinHandle As Long
    
    On Error GoTo err
    
    If Me.Caption <> HIS_CAPTION Then
        Call WriteCommLog("zlSoftViewImage.frmViewImage.TimerCaption_Timer", "窗体标题发生改变", "新标题为：" & Me.Caption, ltDebug)
        Me.Caption = HIS_CAPTION
    End If
    
    '查找消息循环主窗体
    lngWinHandle = FindWindow(vbNullString, HIS_CAPTION)
    If lngWinHandle = 0 Then
        '如果窗口句柄为0 ，则强制退出当前程序
        Call WriteCommLog("zlSoftViewImage.frmViewImage.TimerCaption_Timer", "自我查找窗口句柄=0", "退出程序", ltError)
        Call CloseAllForms
    End If
    Exit Sub
err:
    Call WriteCommLog("zlSoftViewImage.frmViewImage.TimerCaption_Timer", "发生错误，窗口标题为：" & Me.Caption, err.Description, ltError)
End Sub
