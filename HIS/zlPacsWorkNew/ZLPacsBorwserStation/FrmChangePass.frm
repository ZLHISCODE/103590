VERSION 5.00
Begin VB.Form FrmChangePass 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "�޸�����"
   ClientHeight    =   1725
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4860
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1725
   ScaleWidth      =   4860
   StartUpPosition =   1  '����������
   Begin VB.CommandButton CDMȷ�� 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   3510
      TabIndex        =   3
      Top             =   240
      Width           =   1230
   End
   Begin VB.CommandButton CMD���� 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   3510
      TabIndex        =   4
      Top             =   690
      Width           =   1230
   End
   Begin VB.Frame Fra���� 
      Caption         =   "��������"
      Height          =   1455
      Left            =   150
      TabIndex        =   5
      Top             =   150
      Width           =   3165
      Begin VB.TextBox TXTȷ������ 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1110
         MaxLength       =   10
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   1005
         Width           =   1590
      End
      Begin VB.TextBox TXT������ 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1110
         MaxLength       =   10
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   645
         Width           =   1590
      End
      Begin VB.TextBox TXTԭ���� 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1110
         MaxLength       =   10
         PasswordChar    =   "*"
         TabIndex        =   0
         Top             =   270
         Width           =   1590
      End
      Begin VB.Label Lbl������ 
         AutoSize        =   -1  'True
         Caption         =   "������"
         Height          =   180
         Left            =   450
         TabIndex        =   8
         Top             =   330
         Width           =   540
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
      Begin VB.Label Lbl������֤ 
         AutoSize        =   -1  'True
         Caption         =   "������֤"
         Height          =   180
         Left            =   270
         TabIndex        =   6
         Top             =   1065
         Width           =   720
      End
   End
End
Attribute VB_Name = "FrmChangePass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public StrInputPass As String

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
    
    frmUserLogin.mblnChangePass = True
    Me.Hide
End Sub

Private Sub CMD����_Click()
    TXTȷ������ = ""
    TXT������ = ""
    TXTԭ���� = ""
    
    frmUserLogin.mblnChangePass = False
    Me.Hide
End Sub

Private Sub Form_Activate()
    Call SetWindowPos(Me.Hwnd, HWND_TOPMOST, Me.Left / 15, Me.Top / 15, Me.Height / 15, Me.Width / 15, SWP_NOSIZE + SWP_SHOWWINDOW)
    If Trim(frmUserLogin.TXT����) <> "" Then TXTԭ���� = Trim(frmUserLogin.TXT����)
    If TXTԭ���� = "" Then
        TXTԭ����.SetFocus
    Else
        TXT������.SetFocus
    End If
End Sub

Private Sub TXTȷ������_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{Tab}", 1
End Sub

Private Sub TXT������_GotFocus()
    GetFocus TXT������
End Sub

Private Sub TXT������_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{Tab}", 1
End Sub

Private Sub TXTԭ����_GotFocus()
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

Private Sub TXTԭ����_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{Tab}", 1
End Sub
