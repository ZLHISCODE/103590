VERSION 5.00
Begin VB.Form frmShowHisForms 
   BorderStyle     =   0  'None
   Caption         =   "������ʾHIS����NEW"
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
   StartUpPosition =   3  '����ȱʡ
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
    If errHandle("zlSoftCISInterface.frmShowHisForms.ShowMe", "��ʾ���ڳ��ִ���") = 1 Then Resume
End Sub

Private Sub Form_Load()
    On Error Resume Next
    
    TimerCaption.Interval = 30000    '30����
    '���Ͻػ���Ϣ��hook
    plngPreWndProc = Hook(Me.hWnd)
    
    Call GetWindowThreadProcessId(Me.hWnd, glngPid)
    
    
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'ж��hook
    Unhook Me.hWnd, plngPreWndProc
End Sub

Private Sub TimerCaption_Timer()
    Dim lngWinHandle As Long
    
    On Error GoTo err
    
    If Me.Caption <> HIS_CAPTION Then
        Me.Caption = HIS_CAPTION
    End If
    
    '������Ϣѭ��������
    lngWinHandle = FindWindow(vbNullString, HIS_CAPTION)
    If lngWinHandle = 0 Then
        '������ھ��Ϊ0 ����ǿ���˳���ǰ����
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
