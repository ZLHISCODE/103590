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
   StartUpPosition =   3  '����ȱʡ
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
    If errHandle("zlSoftViewImage.frmViewImage.ShowMe", "��ʾ���ڳ��ִ���") = 1 Then Resume
End Sub

Private Sub Form_Load()
    On Error Resume Next
    
    TimerCaption.Interval = 30000    '30����
    '���Ͻػ���Ϣ��hook
    glngPreWndProc = Hook(Me.hWnd)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'ж��hook
    Unhook Me.hWnd, glngPreWndProc
End Sub

Private Sub TimerCaption_Timer()
    Dim lngWinHandle As Long
    
    On Error GoTo err
    
    If Me.Caption <> HIS_CAPTION Then
        Call WriteCommLog("zlSoftViewImage.frmViewImage.TimerCaption_Timer", "������ⷢ���ı�", "�±���Ϊ��" & Me.Caption, ltDebug)
        Me.Caption = HIS_CAPTION
    End If
    
    '������Ϣѭ��������
    lngWinHandle = FindWindow(vbNullString, HIS_CAPTION)
    If lngWinHandle = 0 Then
        '������ھ��Ϊ0 ����ǿ���˳���ǰ����
        Call WriteCommLog("zlSoftViewImage.frmViewImage.TimerCaption_Timer", "���Ҳ��Ҵ��ھ��=0", "�˳�����", ltError)
        Call CloseAllForms
    End If
    Exit Sub
err:
    Call WriteCommLog("zlSoftViewImage.frmViewImage.TimerCaption_Timer", "�������󣬴��ڱ���Ϊ��" & Me.Caption, err.Description, ltError)
End Sub
