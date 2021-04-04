VERSION 5.00
Begin VB.Form frmFullScreen 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "ȫ����ʾ"
   ClientHeight    =   6510
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8595
   LinkTopic       =   "Form1"
   ScaleHeight     =   434
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   573
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
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
  
  '���������°����¼�ʱ�����˳�ȫ����ʾ
  If Shift = 2 Or Shift = 4 Then
    Call Me.Hide   ' Call Unload(Me) 'Shift = 1 Or
    
    Call QuitFullScreen
  End If
  
  If KeyCode = 27 Or KeyCode = 91 Then
    Call Me.Hide   ' Call Unload(Me)
    
    Call QuitFullScreen
  End If
  
End Sub


'�˳�ȫ����ʾ
Private Sub QuitFullScreen()
   Call mCaptureObj.StopPreview
   Call mCaptureObj.StartPreview(mSourceWindow.hWnd)
   Call mCaptureObj.UpdateCaptureWindowPos(mSourceWindow.ScaleWidth, mSourceWindow.ScaleHeight)
End Sub


'����ȫ����ʾ
Private Sub EnterFullScreen()
  If mCaptureObj.hWnd = 0 Then
    Call mCaptureObj.StartPreview(Me.hWnd)
    Call mCaptureObj.UpdateCaptureWindowPos(Me.ScaleWidth, Me.ScaleHeight)
  End If
End Sub


'ȫ����ʾ
Public Sub ShowFullScreen(ByRef captureObj As clsVfwCapture, ByRef parameter As clsVfwParameterCfg, _
  ByRef owner As Object, ByRef sourceWindow As PictureBox, ByVal monitorIndex As Integer)
    
  '���浱ǰ��Ƶ��ʾ��������ã�����ÿ�ζ��Ըñ�����ֵ����ΪǶ��ʽ�ɼ����ں͸�������û��ʹ����ͬ�Ĳɼ�����
  Set mCaptureObj = captureObj
  
  Set mSourceWindow = sourceWindow
    
  '����ȫ����ʾģʽ��������߰��������ţ�
  If mCaptureObj.CaptureParameterInf.VideoShowWay = swAutoFitCut Or mCaptureObj.CaptureParameterInf.VideoShowWay = swNormal _
    Or mCaptureObj.CaptureParameterInf.VideoShowWay = swWindowAutoFit Then
    mCaptureObj.CaptureParameterInf.VideoShowWay = swFit
  ElseIf (mCaptureObj.CaptureParameterInf.VideoShowWay = swStretch) Then
    mCaptureObj.CaptureParameterInf.VideoShowWay = swStretch
  End If
        
  'ȡ����Ļ�����λ��
  Me.Left = ScaleX(gmonitors(monitorIndex + 1).monitorInf.rcMonitor.Left, vbPixels, vbTwips)
  Me.Top = ScaleY(gmonitors(monitorIndex + 1).monitorInf.rcMonitor.Top, vbPixels, vbTwips)
        
  'ȡ����Ļ�Ĵ�С
  Me.Width = ScaleX(gmonitors(monitorIndex + 1).monitorInf.rcMonitor.Right - gmonitors(monitorIndex + 1).monitorInf.rcMonitor.Left, vbPixels, vbTwips)  'Screen.Width
  Me.Height = ScaleY(gmonitors(monitorIndex + 1).monitorInf.rcMonitor.Bottom - gmonitors(monitorIndex + 1).monitorInf.rcMonitor.Top, vbPixels, vbTwips) 'Screen.Height
        
  Call Me.Show(0, owner)
  
  SetWindowPos Me.hWnd, -1, Me.CurrentX, Me.CurrentY, Me.ScaleWidth, Me.ScaleHeight, 3 '�������ö�
  
  '����ȫ����ʾ
  EnterFullScreen
End Sub


Private Sub Form_Unload(Cancel As Integer)
  Call QuitFullScreen
End Sub
