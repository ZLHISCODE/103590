VERSION 5.00
Begin VB.Form frmModifyPassWord_�����山 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "����"
   ClientHeight    =   1995
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5175
   Icon            =   "frmModifyPassWord_�����山.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1995
   ScaleWidth      =   5175
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.Frame Fra���� 
      Caption         =   "��������"
      Height          =   1455
      Left            =   210
      TabIndex        =   2
      Top             =   270
      Width           =   3165
      Begin VB.TextBox TXTԭ���� 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1110
         MaxLength       =   10
         PasswordChar    =   "*"
         TabIndex        =   5
         Top             =   270
         Width           =   1590
      End
      Begin VB.TextBox TXT������ 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1110
         MaxLength       =   10
         PasswordChar    =   "*"
         TabIndex        =   4
         Top             =   645
         Width           =   1590
      End
      Begin VB.TextBox TXTȷ������ 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1110
         MaxLength       =   10
         PasswordChar    =   "*"
         TabIndex        =   3
         Top             =   1005
         Width           =   1590
      End
      Begin VB.Label Lbl������֤ 
         AutoSize        =   -1  'True
         Caption         =   "������֤"
         Height          =   180
         Left            =   270
         TabIndex        =   8
         Top             =   1065
         Width           =   720
      End
      Begin VB.Label Lbl������ 
         AutoSize        =   -1  'True
         Caption         =   "������"
         Height          =   180
         Left            =   450
         TabIndex        =   7
         Top             =   705
         Width           =   540
      End
      Begin VB.Label Lbl������ 
         AutoSize        =   -1  'True
         Caption         =   "������"
         Height          =   180
         Left            =   450
         TabIndex        =   6
         Top             =   330
         Width           =   540
      End
   End
   Begin VB.CommandButton CMD���� 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   3870
      TabIndex        =   1
      Top             =   705
      Width           =   1230
   End
   Begin VB.CommandButton CDMȷ�� 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   3870
      TabIndex        =   0
      Top             =   255
      Width           =   1230
   End
End
Attribute VB_Name = "frmModifyPassWord_�����山"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnChange As Boolean

Private Declare Function SetActiveWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Const GWL_EXSTYLE = (-20)
Private Const WinStyle = &H40000
Private Const SWP_NOSIZE = &H1
Private Const SWP_SHOWWINDOW = &H40
Private Const HWND_TOPMOST = -1
Private Sub CDMȷ��_Click()
    If Trim(TXTԭ����) = "" Then
        MsgBox "����������룡", vbInformation, gstrSysName
        TXTԭ����.SetFocus
        Exit Sub
    End If
    If Trim(TXT������) = "" Then
        MsgBox "�����������룡", vbInformation, gstrSysName
        TXT������.SetFocus
        Exit Sub
    End If
    If Trim(TXTȷ������) = "" Then
        MsgBox "������������֤��", vbInformation, gstrSysName
        TXTȷ������.SetFocus
        Exit Sub
    End If
    If TXT������.Text <> TXTȷ������.Text Then
        MsgBox "����������������������룡", vbInformation, gstrSysName
        TXT������.SetFocus
        Exit Sub
    End If
    If �޸�����_�����山(TXTԭ����, TXT������) = True Then
        g�������_�����山.���� = Trim(TXT������.Text)
        mblnChange = True
        Unload Me
        Exit Sub
    End If
End Sub

Private Sub CMD����_Click()
    TXTȷ������ = ""
    TXT������ = ""
    TXTԭ���� = ""
    mblnChange = False
    Unload Me
End Sub

Private Sub Form_Activate()
    Call SetWindowPos(Me.hwnd, HWND_TOPMOST, Me.Left / 15, Me.Top / 15, Me.Height / 15, Me.Width / 15, SWP_NOSIZE + SWP_SHOWWINDOW)
    If TXTԭ���� = "" Then
        TXTԭ����.SetFocus
    Else
        TXT������.SetFocus
    End If
End Sub

Private Sub TXTȷ������_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{Tab}", 1
End Sub

Private Sub txt������_GotFocus()
    GetFocus TXT������
End Sub

Private Sub TXT������_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{Tab}", 1
End Sub

Private Sub txtԭ����_GotFocus()
    GetFocus TXTԭ����
End Sub

Private Sub TXTȷ������_GotFocus()
    GetFocus TXTȷ������
End Sub

Private Sub GetFocus(ByVal TxtBox As TextBox)
    With TxtBox
        .SelStart = 0
        .SelLength = LenB(StrConv(.Text, vbFromUnicode))
    End With
End Sub

Private Sub txtԭ����_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{Tab}", 1
End Sub

Public Function ShowEdit(ByVal frmMain As Object) As Boolean
    '���
    mblnChange = False
    Me.Show , frmMain
    ShowEdit = mblnChange
End Function
