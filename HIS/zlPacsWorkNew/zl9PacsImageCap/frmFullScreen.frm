VERSION 5.00
Begin VB.Form frmFullScreen 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "全屏显示"
   ClientHeight    =   6510
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8595
   LinkTopic       =   "Form1"
   ScaleHeight     =   434
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   573
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
End
Attribute VB_Name = "frmFullScreen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mCaptureObj As clsVfwCapture
Private mSourceWindow As PictureBox


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  
  '当产生以下按键事件时，则退出全屏显示
  If Shift = 2 Or Shift = 4 Then
    Call Me.Hide   ' Call Unload(Me) 'Shift = 1 Or
    
    Call QuitFullScreen
  End If
  
  If KeyCode = 27 Or KeyCode = 91 Then
    Call Me.Hide   ' Call Unload(Me)
    
    Call QuitFullScreen
  End If
  
End Sub


'退出全屏显示
Private Sub QuitFullScreen()
   Call mCaptureObj.StopPreview
   Call mCaptureObj.StartPreview(mSourceWindow.hWnd)
   Call mCaptureObj.UpdateCaptureWindowPos(mSourceWindow.ScaleWidth, mSourceWindow.ScaleHeight)
End Sub


'进入全屏显示
Private Sub EnterFullScreen()
  If mCaptureObj.hWnd = 0 Then
    Call mCaptureObj.StartPreview(Me.hWnd)
    Call mCaptureObj.UpdateCaptureWindowPos(Me.ScaleWidth, Me.ScaleHeight)
  End If
End Sub


'全屏显示
Public Sub ShowFullScreen(ByRef captureObj As clsVfwCapture, ByRef parameter As clsVfwParameterCfg, _
  ByRef owner As Object, ByRef sourceWindow As PictureBox, ByVal monitorIndex As Integer)
    
  '保存当前视频显示对象的引用，必需每次都对该变量赋值，因为嵌入式采集窗口和浮动窗口没有使用相同的采集对象
  Set mCaptureObj = captureObj
  
  Set mSourceWindow = sourceWindow
    
  '设置全屏显示模式（拉伸或者按比例缩放）
  If mCaptureObj.CaptureParameterInf.VideoShowWay = swAutoFitCut Or mCaptureObj.CaptureParameterInf.VideoShowWay = swNormal _
    Or mCaptureObj.CaptureParameterInf.VideoShowWay = swWindowAutoFit Then
    mCaptureObj.CaptureParameterInf.VideoShowWay = swFit
  ElseIf (mCaptureObj.CaptureParameterInf.VideoShowWay = swStretch) Then
    mCaptureObj.CaptureParameterInf.VideoShowWay = swStretch
  End If
        
  '取得屏幕的相对位置
  Me.Left = ScaleX(gmonitors(monitorIndex + 1).monitorInf.rcMonitor.Left, vbPixels, vbTwips)
  Me.Top = ScaleY(gmonitors(monitorIndex + 1).monitorInf.rcMonitor.Top, vbPixels, vbTwips)
        
  '取得屏幕的大小
  Me.Width = ScaleX(gmonitors(monitorIndex + 1).monitorInf.rcMonitor.Right - gmonitors(monitorIndex + 1).monitorInf.rcMonitor.Left, vbPixels, vbTwips)  'Screen.Width
  Me.Height = ScaleY(gmonitors(monitorIndex + 1).monitorInf.rcMonitor.Bottom - gmonitors(monitorIndex + 1).monitorInf.rcMonitor.Top, vbPixels, vbTwips) 'Screen.Height
        
  Call Me.Show(0, owner)
  
  SetWindowPos Me.hWnd, -1, Me.CurrentX, Me.CurrentY, Me.ScaleWidth, Me.ScaleHeight, 3 '将窗口置顶
  
  '进入全屏显示
  EnterFullScreen
End Sub


Private Sub Form_Unload(Cancel As Integer)
  Call QuitFullScreen
End Sub
