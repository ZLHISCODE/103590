VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTrayIcon 
   Caption         =   "��̨�ļ�����"
   ClientHeight    =   1455
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   3180
   Icon            =   "frmTrayIcon.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1455
   ScaleWidth      =   3180
   StartUpPosition =   1  '����������
   Begin MSComctlLib.ImageList imgIcon 
      Left            =   1200
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTrayIcon.frx":076A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTrayIcon.frx":0EE4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTrayIcon.frx":165E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTrayIcon.frx":1DD8
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTrayIcon.frx":2552
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTrayIcon.frx":2CCC
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTrayIcon.frx":3446
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Timer tmrState 
      Interval        =   500
      Left            =   360
      Top             =   360
   End
   Begin VB.Menu mnuRight 
      Caption         =   "�Ҽ��˵�"
      Begin VB.Menu mnuDir 
         Caption         =   "��Ŀ¼"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuShow 
         Caption         =   "����鿴"
         Shortcut        =   ^S
         Visible         =   0   'False
      End
      Begin VB.Menu mnuQuit 
         Caption         =   "�˳�����"
         Shortcut        =   ^Q
      End
   End
End
Attribute VB_Name = "frmTrayIcon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents mobjIcon As clsTaskIcon
Attribute mobjIcon.VB_VarHelpID = -1
Private mobjFileSys As FileSystemObject
 
Private mstrComputerName As String

Private Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long

 
Private Sub Form_Load()
On Error GoTo ErrorHand
    '������ͼ��
    Set mobjIcon = New clsTaskIcon
    mobjIcon.frmHwnd = Me.hWnd ' hwnd
    mobjIcon.message = "��̨�ļ�����"
    
    Call ResetTrayIcon
    
    mobjIcon.AddIcon
    
    Set mobjFileSys = New FileSystemObject
     
    Exit Sub
ErrorHand:
    MsgBox Err.Description, vbExclamation, "��ʾ"
    Err.Clear
End Sub
 

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrorHand
    mobjIcon.MouseState X
    Exit Sub
ErrorHand:
    MsgBox Err.Description, vbExclamation, "��ʾ"
    Err.Clear
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrorHand

    tmrState.Enabled = False
 
    '�������ͼ��
    mobjIcon.Icon = 0
    mobjIcon.DelIcon
    
    Set mobjIcon = Nothing
    
    Set mobjFileSys = Nothing
    
'    Call mnuQuit_Click
    
    Exit Sub
ErrorHand:
    MsgBox Err.Description, vbExclamation, "��ʾ"
    Err.Clear
End Sub

Private Sub mnuDir_Click()
On Error GoTo errHandle
    ShellExecute 0, "open", gstrCmdPath & "\Failed", "", "", 1
Exit Sub
errHandle:
   MsgBox Err.Description, vbExclamation, "��ʾ"
End Sub

'��ȫ�˳�
Private Sub mnuQuit_Click()
    Dim objForm As Form
On Error GoTo ErrorHand

    If MsgBox("ȷ���˳���̨�ļ�������", vbYesNo, "��ʾ") = vbNo Then Exit Sub
    
    'ж�����д���
    For Each objForm In Forms
        Unload objForm
    Next
    
    Exit Sub
ErrorHand:
    MsgBox Err.Description, vbExclamation, "��ʾ"
    Err.Clear
End Sub

'�����ô���
Private Sub mnuSetup_Click()
On Error GoTo ErrorHand
    Call frmTransView.zlShowMe(gstrCmdPath, Me)
    Exit Sub
ErrorHand:
    MsgBox Err.Description, vbExclamation, "��ʾ"
    Err.Clear
End Sub

'�����ô���
Private Sub mobjIcon_MouseLeftDBClick()
On Error GoTo ErrorHand
    Call mnuSetup_Click
    Exit Sub
ErrorHand:
    MsgBox Err.Description, vbExclamation, "��ʾ"
    Err.Clear
End Sub

'��ʾ�Ҽ��˵�
Private Sub mobjIcon_MouseRightUp()
On Error GoTo ErrorHand
    SetForegroundWindow Me.hWnd
    
    PopupMenu mnuRight
    Exit Sub
ErrorHand:
    MsgBox Err.Description, vbExclamation, "��ʾ"
    Err.Clear
End Sub
 

Private Sub tmrState_Timer()
    Dim strFailedPath As String
On Error GoTo errHandle
    
    strFailedPath = Replace(gstrCmdPath & "\Failed\", "\\", "\")
    If mobjFileSys.GetFolder(strFailedPath).Files.Count > 0 Then
        gblnIsFailed = True
    Else
        gblnIsFailed = False
    End If
    
    Call ResetTrayIcon
Exit Sub
errHandle:
End Sub


Public Sub ResetTrayIcon()
    '0-���棬1-�ϴ���2-����
    Dim lngImgIndex As Long

On Error GoTo errHandle

    If gblnSingle Then
        If gblnIsFailed Then
            lngImgIndex = 6
        Else
            lngImgIndex = 5
        End If
    Else
        If gblnWorking Then
            If Val(mobjIcon.Tag) = 0 Then
                lngImgIndex = 2
                mobjIcon.Tag = 1
            Else
                lngImgIndex = 3
                mobjIcon.Tag = 0
            End If
        Else
            If Val(mobjIcon.Tag) = 0 Then
                lngImgIndex = 7
                mobjIcon.Tag = 1
            Else
                If gblnIsFailed Then
                    lngImgIndex = 4 'ʧ��ͼ��
                Else
                    lngImgIndex = 1 '����ͼ��
                End If
                
                mobjIcon.Tag = 0
            End If
        End If
    End If

    Set Me.Icon = imgIcon.ListImages(lngImgIndex).Picture
    mobjIcon.Icon = Me.Icon.Handle
Exit Sub
errHandle:
End Sub


Public Sub ShowMessage(ByVal strMsg As String, ByVal lngMsgType As Long)
    Call mobjIcon.ShowMsg(strMsg, lngMsgType)
End Sub
