VERSION 5.00
Begin VB.Form frmDeptGo 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��λ����"
   ClientHeight    =   1725
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4605
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1725
   ScaleWidth      =   4605
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   3360
      TabIndex        =   7
      Top             =   720
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   3360
      TabIndex        =   6
      Top             =   240
      Width           =   1100
   End
   Begin VB.Frame fraPati 
      Height          =   1440
      Left            =   240
      TabIndex        =   8
      Top             =   120
      Width           =   2955
      Begin VB.TextBox txtסԺ�� 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1245
         MaxLength       =   18
         TabIndex        =   3
         Top             =   615
         Width           =   1275
      End
      Begin VB.TextBox txt���� 
         Height          =   300
         Left            =   1245
         MaxLength       =   15
         TabIndex        =   5
         Top             =   990
         Width           =   1275
      End
      Begin VB.TextBox txt���� 
         Height          =   300
         Left            =   1245
         MaxLength       =   10
         TabIndex        =   1
         Top             =   240
         Width           =   1275
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "סԺ��(&2)"
         Height          =   180
         Left            =   375
         TabIndex        =   2
         Top             =   675
         Width           =   810
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����(&3)"
         Height          =   180
         Left            =   555
         TabIndex        =   4
         Top             =   1050
         Width           =   630
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����(&1)"
         Height          =   180
         Left            =   555
         TabIndex        =   0
         Top             =   300
         Width           =   630
      End
   End
End
Attribute VB_Name = "frmDeptGo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'Ҫ���������

Private Sub cmdCancel_Click()
    gblnOK = False
    Hide
End Sub



Private Sub cmdOK_Click()
    If txtסԺ��.Text = "" And txt����.Text = "" And txt����.Text = "" Then
        MsgBox "�������趨һ��������", vbInformation, gstrSysName
        txt����.SetFocus: Exit Sub
    End If
    gblnOK = True
    Hide
End Sub

Private Sub Form_Activate()
   txt����.SetFocus
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Then KeyAscii = 0
    If KeyAscii = 13 Then KeyAscii = 0: Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub Form_Load()
    gblnOK = False
End Sub

Private Sub txtסԺ��_GotFocus()
    zlControl.TxtSelAll txtסԺ��
End Sub

Private Sub txt����_GotFocus()
    zlControl.TxtSelAll txt����
End Sub

Private Sub txt����_GotFocus()
    zlControl.TxtSelAll txt����
End Sub

Private Sub txtסԺ��_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub
