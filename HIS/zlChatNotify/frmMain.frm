VERSION 5.00
Begin VB.Form frmMain 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "����"
   ClientHeight    =   840
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1395
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   840
   ScaleWidth      =   1395
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.Timer tmrIcon 
      Enabled         =   0   'False
      Interval        =   400
      Left            =   840
      Top             =   0
   End
   Begin VB.PictureBox picMsg 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   0
      Picture         =   "frmMain.frx":6852
      ScaleHeight     =   330
      ScaleWidth      =   375
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   405
   End
   Begin VB.PictureBox PicNo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   480
      Picture         =   "frmMain.frx":D0A4
      ScaleHeight     =   330
      ScaleWidth      =   255
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   285
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mnfIconData As NOTIFYICONDATA
Private mblnIconShow As Boolean             '����ͼ����˸״̬


Public Sub SetNotifyIcon(ByVal intType As Integer, Optional ByVal strMsg As String)
    'intType 0-��ʼ��  1-��Ϣ 2-��˸ 3-��ԭ
    'strMsg
    On Error Resume Next
    '����Ĵ�����Խ�ͼ����ӵ�ϵͳͼ��
    If intType = 0 And mnfIconData.hWnd <> 0 Then Call Shell_NotifyIcon(NIM_DELETE, mnfIconData)
    mnfIconData.hWnd = Me.hWnd
    mnfIconData.uID = picMsg.Picture '����ȷ��ʹ���ĸ�ͼ��
    mnfIconData.uFlags = NIF_ICON Or NIF_MESSAGE Or NIF_TIP
    mnfIconData.uCallbackMessage = WM_MOUSEMOVE
    mnfIconData.hIcon = IIf(intType = 2, PicNo.Picture.Handle, picMsg.Picture.Handle)
    If strMsg = "" Then strMsg = gstrSysName & vbCrLf & "��ǰ�û���" & gstrUser
    mnfIconData.szTip = strMsg & vbNullChar  '�����ǽ�����Ƶ�ͼ����ʱ������ʾ������
    mnfIconData.cbSize = Len(mnfIconData)
    Call Shell_NotifyIcon(IIf(intType = 0, NIM_ADD, NIM_MODIFY), mnfIconData)
End Sub

Public Function SetIcon(ByVal bytFunc As Byte) As Boolean
'����:bytFunc =1 ������˸;=2 �ر���˸
    If bytFunc = 1 Then
        tmrIcon.Enabled = True
    Else
        tmrIcon.Enabled = False
        Call SetNotifyIcon(1)
    End If
End Function

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lngMsg As Long
    If Not tmrIcon.Enabled Then Exit Sub
    lngMsg = X / Screen.TwipsPerPixelX
     
    If lngMsg = WM_LBUTTONDBLCLK Then
        Call frmChatList.ShowMe(1)
        Exit Sub
    ElseIf lngMsg = WM_RBUTTONUP Then '����Ҽ�
        Exit Sub
    End If
     
    '��δ����Ϣ�Ҵ���δ��ʾʱ,��ʾδ���嵥
    If Not grsList Is Nothing Then
        grsList.Filter = ""
        If grsList.RecordCount > 0 Then
            If Not gblnShow Then
                Call frmChatList.ShowMe(0)
            End If
        End If
    End If

    
End Sub
 
Private Sub Form_Unload(Cancel As Integer)
    Call Shell_NotifyIcon(NIM_DELETE, mnfIconData)
End Sub

Private Sub tmrIcon_Timer()
    Call SetNotifyIcon(IIf(mblnIconShow, 1, 2), "��ǰ�û���" & gstrUser)
    mblnIconShow = Not mblnIconShow
End Sub
